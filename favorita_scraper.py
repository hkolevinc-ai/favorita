#!/usr/bin/env python3
"""Scrape selected Favorita categories into the supplied Temu upload template."""

from __future__ import annotations

import argparse
import copy
import html
import json
import logging
import re
import sys
import time
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen

from lxml import html as lxml_html
from openpyxl import load_workbook


BASE_URL = "https://www.favorita.bg/"
CATEGORY_URLS = (
    urljoin(BASE_URL, "fc-barcelona.html"),
    urljoin(BASE_URL, "liverpool-fc.html"),
    urljoin(BASE_URL, "real-madrid-cf.html"),
)
USER_AGENT = "Mozilla/5.0 (compatible; FavoritaTemuScraper/1.0)"
TEMPLATE_SHEET = "Template"
FIRST_DATA_ROW = 5
MAX_DATA_ROW = 2998
MAX_IMAGES = 10

# Columns in the supplied Temu template (1-based).
CATEGORY_COL = 5
CATEGORY_NAME_COL = 6
NAME_COL = 12
GOODS_SKU_COL = 13
SKU_COL = 14
ACTION_COL = 15
BRAND_COL = 18
TRADEMARK_COL = 19
DESCRIPTION_COL = 20
DETAIL_IMAGE_COL = 27
VARIATION_THEME_COL = 1008
SIZE_FAMILY_COL = 1009
SIZE_COL = 1011
COLOR_COL = 1012
MODEL_COL = 1024
SKU_IMAGE_COL = 1093
QTY_COL = 1104
BASE_PRICE_COL = 1105
REFERENCE_LINK_COL = 1106
LIST_PRICE_COL = 1107
WEIGHT_COL = 1109
LENGTH_COL = 1110
WIDTH_COL = 1111
HEIGHT_COL = 1112
SKU_TYPE_COL = 1113
INDIVIDUALLY_PACKED_COL = 1114
PACK_QTY_COL = 1115
PACK_UNIT_COL = 1116
SHIPPING_TEMPLATE_COL = 1124
COUNTRY_COL = 1128
PRODUCT_IDENTIFICATION_COL = 1181
MANUFACTURER_COL = 1182

# The attached template offers this constrained Temu category set. The rules
# below choose the closest valid category available in that file.
CATEGORY_MAP = {
    "БУТИЛКА / ХАЛБА": 34090,
    "ДЕТСКА ТЕНИСКА": 34988,
    "ДЕТСКИ АНЦУГ": 34988,
    "ДЕТСКИ ФУТБОЛЕН ЕКИП": 34988,
    "МЪЖКА ТЕНИСКА": 30469,
    "МЪЖКИ АНЦУГ": 30469,
    "МЪЖКИ СУИЧЪР": 30469,
    "Блуза/Суичър": 30469,
    "НЕСЕСЕР": 31752,
    "ПЛЕТЕНА ШАПКА": 34859,
    "ПОРТМОНЕ / ЧАНТА": 30718,
    "Раница": 31021,
    "РАНИЦА / МЕШКА": 31021,
    "САК": 29158,
    "ФУТБОЛНА ТОПКА": 36731,
    "Хавлия": 11812,
    "ЧАСОВНИК": 31748,
    "ЧАША": 10585,
    "ЧОРАПИ": 34988,
    "ШАЛ": 30253,
    "ШАПКА ИДИОТКА": 34859,
    "ШАПКА С КОЗИРКА": 34859,
    "ЯКЕ": 30469,
    "ВРАТАРСКИ РЪКАВИЦИ": 36854,
    "Зимни ръкавици": 36854,
    "КАЛЕНДАР": 12150,
    "КУТИЯ ЗА ХРАНА": 31752,
    "КЪСИ ПАНТАЛОНИ / БАНСКИ": 30469,
    "КОЛАН": 30625,
    "АКСЕСОАРИ": 30625,
    "БЕБЕШКИ ЕКИП": 34988,
}


@dataclass(frozen=True)
class Variant:
    variant_id: str
    label: str
    price: Decimal
    in_stock: bool
    code: str


@dataclass(frozen=True)
class Product:
    url: str
    team: str
    source_category: str
    name: str
    description: str
    code: str
    images: tuple[str, ...]
    variants: tuple[Variant, ...]
    base_price: Decimal
    in_stock: bool


def clean_text(value: str | None) -> str:
    if not value:
        return ""
    value = html.unescape(value).replace("\xa0", " ")
    return re.sub(r"[ \t]+", " ", re.sub(r"\n[ \t]+", "\n", value)).strip()


def slug_token(value: str) -> str:
    value = clean_text(value).upper()
    value = re.sub(r"[^0-9A-ZА-Я]+", "-", value, flags=re.IGNORECASE)
    return value.strip("-")[:70]


class Fetcher:
    def __init__(self, retries: int = 4, delay: float = 0.35, timeout: int = 35):
        self.retries = retries
        self.delay = delay
        self.timeout = timeout

    def get(self, url: str) -> str:
        last_error = None
        for attempt in range(1, self.retries + 1):
            try:
                req = Request(url, headers={"User-Agent": USER_AGENT, "Accept-Language": "bg,en;q=0.8"})
                with urlopen(req, timeout=self.timeout) as response:
                    charset = response.headers.get_content_charset() or "utf-8"
                    body = response.read().decode(charset, errors="replace")
                time.sleep(self.delay)
                return body
            except (HTTPError, URLError, TimeoutError, OSError) as exc:
                last_error = exc
                logging.warning("Fetch failed (%s/%s): %s", attempt, self.retries, url)
                time.sleep(min(2 ** attempt, 8))
        raise RuntimeError(f"Unable to download {url}: {last_error}")


def first_text(doc, xpath: str) -> str:
    values = doc.xpath(xpath)
    if not values:
        return ""
    value = values[0]
    if hasattr(value, "text_content"):
        value = value.text_content()
    return clean_text(str(value))


def discover_product_urls(fetcher: Fetcher, category_url: str) -> list[str]:
    found: list[str] = []
    page = 1
    while True:
        url = category_url if page == 1 else f"{category_url}?page={page}"
        doc = lxml_html.fromstring(fetcher.get(url))
        links = [urljoin(BASE_URL, x) for x in doc.xpath('//a[contains(@class,"gs-item-title")]/@href')]
        links = list(dict.fromkeys(links))
        if not links:
            break
        for link in links:
            if link not in found:
                found.append(link)
        next_pages = [int(x) for x in doc.xpath('//a[contains(@class,"gs-paging-link")]/@data-page') if str(x).isdigit()]
        if not next_pages or page >= max(next_pages):
            break
        page += 1
    return found


def parse_variant_json(page_source: str) -> dict:
    match = re.search(r"pub\.product\.variant\.init\('(.*?)',\s*'orderFrm'", page_source, re.S)
    if not match or match.group(1) in ("", "{}"):
        return {}
    encoded = match.group(1)
    try:
        # JavaScript string escaping around JSON is compatible with this targeted unescape.
        decoded = bytes(encoded, "utf-8").decode("unicode_escape")
        return json.loads(decoded)
    except (UnicodeDecodeError, json.JSONDecodeError):
        # Most pages contain only ASCII JSON plus \u escapes; direct parsing covers them.
        return json.loads(encoded)


def parse_price(text: str) -> Decimal:
    match = re.search(r"([0-9]+(?:[.,][0-9]+)?)", clean_text(text))
    if not match:
        raise ValueError(f"Price not found in {text!r}")
    return Decimal(match.group(1).replace(",", ".")).quantize(Decimal("0.01"))


def parse_product(fetcher: Fetcher, url: str) -> Product:
    source = fetcher.get(url)
    doc = lxml_html.fromstring(source)
    name = first_text(doc, '//div[contains(@class,"gs-item-title")]/h1')
    code = first_text(doc, '//strong[@data-code]')
    price_text = first_text(doc, '//strong[@data-var-price]')
    base_price = parse_price(price_text)
    desc_nodes = doc.xpath('//div[contains(@class,"gs-item-txt") and contains(@class,"gs-rtf")]')
    description = clean_text(desc_nodes[0].text_content()) if desc_nodes else first_text(doc, '//meta[@name="description"]/@content')
    # Remove shop-service boilerplate from the end when present.
    description = re.split(r"Допълнителна информация:|Всички поръчки се изпращат|Всички продукти се изпращат", description, maxsplit=1)[0].strip()
    images = []
    for path in doc.xpath('//div[@id="gs-gallery"]//img/@data-zoom-image | //div[@id="gs-gallery"]//img/@src'):
        full = urljoin(BASE_URL, path)
        if full not in images:
            images.append(full)
    team = ""
    source_category = ""
    for row in doc.xpath('//div[contains(@class,"gs-tab-row")]'):
        label = first_text(row, './/strong')
        value = first_text(row, './/span')
        if label.startswith("Отбор"):
            team = value
        elif label.startswith("Категория"):
            source_category = value
    variants = []
    for vid, raw in parse_variant_json(source).items():
        label = clean_text(raw.get("optionsName"))
        price = Decimal(str(raw.get("price", base_price))).quantize(Decimal("0.01"))
        variants.append(Variant(
            variant_id=str(raw.get("Id", vid)),
            label=label,
            price=price,
            in_stock=bool(raw.get("inStock")),
            code=clean_text(raw.get("code")) or code,
        ))
    if variants:
        in_stock = any(v.in_stock for v in variants)
    else:
        availability = first_text(doc, '//meta[@property="og:availability"]/@content').lower()
        button_exists = bool(doc.xpath('//*[@id="addToCartBtn"]'))
        sold_out = bool(doc.xpath('//*[contains(@class,"out-of-stock") or contains(@class,"not-available")]'))
        in_stock = availability in ("instock", "in stock") or (button_exists and not sold_out)
    return Product(
        url=url,
        team=team,
        source_category=source_category,
        name=name,
        description=description,
        code=code,
        images=tuple(images[:MAX_IMAGES]),
        variants=tuple(variants),
        base_price=base_price,
        in_stock=in_stock,
    )


def choose_temu_category(product: Product) -> int:
    title = product.name.lower()
    cat = product.source_category
    if "портф" in title or "портмоне" in title:
        return 30656
    if "подаръчна торб" in title:
        return 17328
    if "чанта за обув" in title:
        return 31111
    if "раница" in title:
        return 31021
    if "сак" in title:
        return 29158
    if "чанта" in title and "обяд" not in title and "термо" not in title:
        return 30718
    if "термо чаша" in title:
        return 12703
    if any(x in title for x in ("чаша", "халба")):
        return 10585
    if "бутилка" in title:
        return 34090
    if "ключодърж" in title:
        return 30625
    if "картичка" in title:
        return 1111
    if "бадж" in title or "значк" in title:
        return 1432
    if "молив" in title:
        return 1291
    if "химикал" in title:
        return 1180
    if "шал" in title:
        return 30253
    if "топка" in title:
        return 36731
    if "вратарски ръкавиц" in title:
        return 36854
    if "хавли" in title:
        return 11812
    if "чадър" in title:
        return 31015
    if "пижам" in title:
        return 30827
    if "часовник" in title:
        return 31748
    if "шапка" in title:
        return 34859
    if any(x in title for x in ("детск", "yamal", "salah", "bellingham", "mbapp")) and any(
        x in title for x in ("екип", "тениска", "анцуг", "худи", "ветровка", "яке", "бански", "чорап")
    ):
        return 34988
    return CATEGORY_MAP.get(cat, 31752)


def extract_color(name: str, description: str) -> str:
    text = f"{name} {description}".lower()
    colors = (
        ("Черна", ("черна", "черен", "black")), ("Бяла", ("бяла", "бял", "white")),
        ("Червена", ("червена", "червен", "red")), ("Синя", ("синя", "син", "blue")),
        ("Тъмносиня", ("тъмносиня", "тъмносин")), ("Зелена", ("зелена", "зелен")),
        ("Сива", ("сива", "сив")), ("Жълта", ("жълта", "жълт")),
    )
    for label, needles in colors:
        if any(x in text for x in needles):
            return label
    return ""


def team_brand(team: str, name: str) -> tuple[str, str]:
    text = f"{team} {name}".lower()
    if "barcelona" in text:
        return "FC BARCELONA", "FC Barcelona"
    if "liverpool" in text:
        return "LIVERPOOL FC", "Liverpool FC"
    return "REAL MADRID", "Real Madrid"


def team_country(team: str, name: str) -> str:
    return "United Kingdom" if "liverpool" in f"{team} {name}".lower() else "Spain"


def infer_dimensions(description: str) -> tuple[int, int, int, int]:
    match = re.search(r"(?:Размер(?:и)?|Dimensions?)\s*:?\s*(\d+(?:[.,]\d+)?)\s*[xх×]\s*(\d+(?:[.,]\d+)?)(?:\s*[xх×]\s*(\d+(?:[.,]\d+)?))?\s*см", description, re.I)
    if match:
        vals = [max(1, int(Decimal(x.replace(",", ".")).quantize(Decimal("1")))) for x in match.groups() if x]
        while len(vals) < 3:
            vals.append(5)
        return 500, vals[0], vals[1], vals[2]
    return 500, 20, 20, 20


def copy_row_format(ws, source_row: int, target_row: int) -> None:
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for col in range(1, ws.max_column + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy.copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy.copy(src.alignment)
        if src.protection:
            dst.protection = copy.copy(src.protection)


def clear_template_rows(ws) -> None:
    for row in ws.iter_rows(min_row=FIRST_DATA_ROW, max_row=MAX_DATA_ROW, max_col=ws.max_column):
        for cell in row:
            cell.value = None
            if cell.comment:
                cell.comment = None


def formula_for_category_name(row: int) -> str:
    return '=IFERROR(VLOOKUP(INDIRECT(ADDRESS(ROW(), COLUMN()-1)), \'Category Name\'!$A:$B, 2, FALSE), "Please select the [category] first. After selecting, the full category path will be displayed here. Please do not modify the functions in the cells")'


def product_rows(products: Iterable[Product]):
    for product in products:
        if product.variants:
            for variant in product.variants:
                if variant.in_stock:
                    yield product, variant
        elif product.in_stock:
            yield product, None


def write_output(template: Path, output: Path, products: list[Product]) -> int:
    wb = load_workbook(template)
    ws = wb[TEMPLATE_SHEET]
    clear_template_rows(ws)
    rows = list(product_rows(products))
    if len(rows) > MAX_DATA_ROW - FIRST_DATA_ROW + 1:
        raise RuntimeError(f"Too many output rows ({len(rows)}) for the template")

    for offset, (product, variant) in enumerate(rows):
        row = FIRST_DATA_ROW + offset
        if row != FIRST_DATA_ROW:
            copy_row_format(ws, FIRST_DATA_ROW, row)
        category_id = choose_temu_category(product)
        price = variant.price if variant else product.base_price
        price = price.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        list_price = (price * 2).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        brand, trademark = team_brand(product.team, product.name)
        variant_label = variant.label if variant else ""
        parent_sku = product.code or f"FAV-{slug_token(product.name)}"
        child_sku = parent_sku if not variant_label else f"{parent_sku}-{slug_token(variant_label)}"
        weight, length, width, height = infer_dimensions(product.description)

        ws.cell(row, CATEGORY_COL, category_id)
        ws.cell(row, CATEGORY_NAME_COL, formula_for_category_name(row))
        ws.cell(row, NAME_COL, product.name[:500])
        ws.cell(row, GOODS_SKU_COL, parent_sku[:80])
        ws.cell(row, SKU_COL, child_sku[:80])
        ws.cell(row, ACTION_COL, "Add")
        ws.cell(row, BRAND_COL, brand)
        ws.cell(row, TRADEMARK_COL, trademark)
        ws.cell(row, DESCRIPTION_COL, product.description[:5000])
        for idx, image in enumerate(product.images[:MAX_IMAGES]):
            ws.cell(row, DETAIL_IMAGE_COL + idx, image)
            ws.cell(row, SKU_IMAGE_COL + idx, image)
        ws.cell(row, VARIATION_THEME_COL, "Size" if variant_label else "Model")
        ws.cell(row, SIZE_FAMILY_COL, "101 - Custom size" if variant_label else None)
        ws.cell(row, SIZE_COL, variant_label or None)
        ws.cell(row, COLOR_COL, extract_color(product.name, product.description) or None)
        ws.cell(row, MODEL_COL, product.team or brand)
        ws.cell(row, QTY_COL, 10)
        ws.cell(row, BASE_PRICE_COL, float(price))
        ws.cell(row, REFERENCE_LINK_COL, product.url)
        ws.cell(row, LIST_PRICE_COL, float(list_price))
        ws.cell(row, WEIGHT_COL, weight)
        ws.cell(row, LENGTH_COL, length)
        ws.cell(row, WIDTH_COL, width)
        ws.cell(row, HEIGHT_COL, height)
        ws.cell(row, SKU_TYPE_COL, "Single set")
        ws.cell(row, INDIVIDUALLY_PACKED_COL, "Yes")
        ws.cell(row, PACK_QTY_COL, 1)
        ws.cell(row, PACK_UNIT_COL, "piece")
        ws.cell(row, SHIPPING_TEMPLATE_COL, "Доставка")
        ws.cell(row, COUNTRY_COL, team_country(product.team, product.name))
        ws.cell(row, PRODUCT_IDENTIFICATION_COL, child_sku[:80])
        ws.cell(row, MANUFACTURER_COL, trademark)

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return len(rows)


def run(template: Path, output: Path, limit: int = 0) -> None:
    fetcher = Fetcher()
    all_urls: list[str] = []
    for category_url in CATEGORY_URLS:
        links = discover_product_urls(fetcher, category_url)
        logging.info("Found %d products in %s", len(links), category_url)
        all_urls.extend(links)
    all_urls = list(dict.fromkeys(all_urls))
    if limit:
        all_urls = all_urls[:limit]
    products = []
    for idx, url in enumerate(all_urls, 1):
        try:
            product = parse_product(fetcher, url)
            products.append(product)
            logging.info("[%d/%d] %s", idx, len(all_urls), product.name)
        except Exception as exc:
            logging.exception("Skipped product %s: %s", url, exc)
    if not products:
        raise RuntimeError("No products were successfully scraped")
    rows = write_output(template, output, products)
    logging.info("Saved %d in-stock SKU rows to %s", rows, output)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", type=Path, default=Path(__file__).with_name("Temu_upload.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("output") / "favorita_temu_upload.xlsx")
    parser.add_argument("--limit", type=int, default=0, help="Testing only: scrape at most N products")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    try:
        run(args.template, args.output, args.limit)
        return 0
    except Exception as exc:
        logging.exception("Scraper failed: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
