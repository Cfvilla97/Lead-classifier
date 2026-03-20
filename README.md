# Lead Classifier

A self-serve tool for Digital Sales teams to classify ALG leads across EU markets.

## What it does

Upload three files and get a classified Excel report with 5 tabs:

| Label | Meaning |
|---|---|
| ✅ Qualified / Convert | Clean lead ready for sales |
| 🔴 Duplicate | Already exists in Salesforce CRM |
| 🟡 Business Closed | Permanently or temporarily closed on Google Maps |
| 🟠 Wrong Target Group | Not a food/drink business |
| ⚫ Invalid Data | Not found on Google / no Apify match |

## Supported markets

| Code | Country | Phone prefix |
|---|---|---|
| TR | Turkey | +90 |
| NO | Norway | +47 |
| SE | Sweden | +46 |
| CZ | Czech Republic | +420 |
| HU | Hungary | +36 |
| AT | Austria | +43 |

## How to use

### Tab 1 — Classify leads
1. Select your market from the dropdown
2. Upload your **ALG leads file** (.xlsx or .csv)
3. Upload your **Apify Google Maps results** (.csv)
4. Upload your **Salesforce CRM export** (.csv)
5. Click **Run classification**
6. Download the Excel report

### Tab 2 — Generate Apify URLs
If you don't have Apify results yet:
1. Upload your leads file
2. Download the CSV with generated Google Maps URLs
3. Paste URLs into Apify → Google Maps Extractor
4. Run Apify and download the output
5. Come back to Tab 1

## Matching logic

**Duplicate detection (vs CRM):**
- Phone match — normalised phone number found in CRM (exact)
- Name match — exact name match after normalising market-specific characters

**Google Maps validation (vs Apify):**
- Phone ✓ — lead phone matches Apify phone
- Name score — token overlap + sequence similarity (threshold ≥ 0.5)
- Address ✓ — postal code or street tokens overlap
- Confirmed if: phone matches OR name score ≥ 0.5

**Match Confidence score (0–1):**
- Phone hit: +0.4
- Name score: ×0.4
- Address hit: +0.2

## Deploy to Streamlit Community Cloud

1. Fork or push this repo to your GitHub account
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app**
4. Select your repo, branch `main`, file `app.py`
5. Click **Deploy**

The app will be live at `https://your-app-name.streamlit.app`

## Local development

```bash
pip install -r requirements.txt
streamlit run app.py
```

## File format notes

The tool auto-detects column names — no specific column order required.
It looks for these column name variants across markets:

| Field | Accepted column names |
|---|---|
| Business name | Company / Account, Account Name, Name |
| Phone | Phone, phone_number, Telefon |
| Street | Street, Address, Adresse |
| City | City, By, Stad, Město |
| GRID | GRID, GRID__c |
| Google URL | GOOGLE URL, Google URL, url |
| Coordinates | Coordinates (Latitude/Longitude), Latitude, lat |
