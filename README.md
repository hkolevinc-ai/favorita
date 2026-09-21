# Favorita → Temu scraper

Scrapes only these Favorita categories:

- `https://www.favorita.bg/fc-barcelona.html`
- `https://www.favorita.bg/liverpool-fc.html`
- `https://www.favorita.bg/real-madrid-cf.html`

The scraper follows every category page, opens every product, keeps only products and variants marked as available by Favorita, and creates one Temu SKU row per available size/variant. Variants whose label says that they are sold out are excluded even if the source JSON contains an inconsistent availability flag.

Pricing rules:

- `Base Price - EUR` = current Favorita price
- `List Price - EUR` = `Base Price × 2`
- `Quantity` = 10

Additional validation and mapping rules:

- `Contribution Goods` is unique for every product, while all variants of that product share it.
- `Contribution SKU` is unique for every available variant.
- Zero variant prices are repaired from the matching positive sibling price or the current product price.
- Variation themes, size families, colors, required category attributes and size-chart fields are populated according to the supplied Temu template.
- Package weight and dimensions use values from the product description when available, otherwise category-specific estimates.
- `EU Responsible person` is intentionally left blank for the merchant to complete.
- The run fails instead of producing an artifact if it detects duplicate SKU codes, merged parent products, invalid variation themes, missing images or invalid prices.

The `Template` product area is cleared before every run, so no previous product rows remain.

## Performance

Version 1.2 downloads up to 8 product pages in parallel and avoids creating or
copying millions of empty Excel cells. A full run is expected to finish in
roughly 15–30 minutes under normal GitHub Actions network conditions. The
workflow stops after 45 minutes instead of continuing indefinitely.

## GitHub Actions

1. Create a new GitHub repository.
2. Upload the complete contents of this folder, including `.github/workflows/scrape.yml`.
3. Open the repository's **Actions** tab.
4. Select **Favorita to Temu scraper** and click **Run workflow**.
5. Download the `favorita-temu-results` artifact after the run finishes.

The artifact contains `favorita_temu_upload.xlsx`.

## Local run

```bash
python -m pip install -r requirements.txt
python favorita_scraper.py
```

The result is written to `output/favorita_temu_upload.xlsx`.

If Favorita temporarily limits requests, reduce concurrency with
`python favorita_scraper.py --workers 8`.
