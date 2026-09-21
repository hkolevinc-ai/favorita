# Favorita → Temu scraper

Scrapes only these Favorita categories:

- `https://www.favorita.bg/fc-barcelona.html`
- `https://www.favorita.bg/liverpool-fc.html`
- `https://www.favorita.bg/real-madrid-cf.html`

The scraper follows every category page, opens every product, keeps only products and variants marked as available by Favorita, and creates one Temu SKU row per available size/variant.

Pricing rules:

- `Base Price - EUR` = current Favorita price
- `List Price - EUR` = `Base Price × 2`
- `Quantity` = 10

The `Template` product area is cleared before every run, so no previous product rows remain.

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
