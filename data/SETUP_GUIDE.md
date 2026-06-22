# Amul SAP PO Converter — Setup & Deployment Guide

## WHAT THIS APP DOES
Upload e-commerce platform PO files (Zepto today; Blinkit, Swiggy
Instamart, BigBasket, Reliance Mart configured but waiting on a real
sample file from you) → the app maps platform SKUs to Amul SAP codes,
converts ordered units into SAP box/case quantities (rounding up and
flagging any remainder), and builds a single monthly projection table
your manager can use to raise POs.

---

## PART 1 — TEST IT ON YOUR OWN COMPUTER FIRST

1. Install Python 3.11: https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe
   (tick "Add python.exe to PATH" during install)

2. Open the `amul_sap_converter` folder → click the address bar →
   type `cmd` → press Enter

3. Install requirements (one time):
   ```
   pip install -r requirements.txt
   ```

4. Run the app:
   ```
   streamlit run app.py
   ```

5. Browser opens at http://localhost:8501 — upload your Zepto PO file
   and confirm the output looks right before going further.

---

## PART 2 — HOST IT FOR YOUR MANAGER (FREE, Streamlit Community Cloud)

This gives your manager a permanent web link — no installs, no
commands, opens in any browser like a normal website.

### Step 1: Create a free GitHub account
Go to https://github.com/signup if you don't have one already.

### Step 2: Create a new repository
1. Click the "+" icon top-right → "New repository"
2. Name it `amul-sap-converter`
3. Set it to **Private** (recommended, since it holds your SKU data)
4. Click "Create repository"

### Step 3: Upload your project files
1. On your new repo page, click "uploading an existing file"
2. Drag in ALL files from the `amul_sap_converter` folder:
   - `app.py`
   - `mapping_engine.py`
   - `platform_readers.py`
   - `etl_engine.py`
   - `requirements.txt`
   - the `data` folder (containing `Amul_Article_Master.xlsx`)
3. Click "Commit changes"

### Step 4: Deploy on Streamlit Community Cloud
1. Go to https://streamlit.io/cloud and sign up free using your GitHub account
2. Click "New app"
3. Select your `amul-sap-converter` repository
4. Main file path: `app.py`
5. Click "Deploy"

Wait 2–3 minutes. You'll get a permanent link like:
```
https://amul-sap-converter.streamlit.app
```

**Share this link with your manager** — that's the entire website. They
just open it in a browser, upload files, and download results, every day.

### Updating the app later
Any time you change a file (e.g. update `app.py`), just upload the new
version to the same GitHub repo — Streamlit Cloud auto-redeploys within
a minute or two. No need to repeat the deploy steps.

---

## PART 3 — HOW THE MAPPING WORKS

- The master file lives at `data/Amul_Article_Master.xlsx`
- It has one row per SAP product, with a column per platform holding
  that platform's SKU/code for the same product
- The app reads the **SAP Product Description** to automatically work
  out the case-pack size (e.g. "30x200 Ml TP" → 30 units per box)
- When platform order quantity doesn't divide evenly into the case
  pack, the app rounds UP to the next box and highlights it in yellow
  in both the on-screen table and the downloaded Excel

## UPDATING THE MAPPING TABLE
Two ways, both inside the website itself (sidebar → "Manage SKU Mapping"):
1. **Add New Mapping** — a form to add a brand-new SAP product with its
   platform SKU codes
2. **View / Search Mapping** — search the existing 959+ products to
   verify or check current values

For any new platform beyond the 5 currently registered (Zepto,
Blinkit, Swiggy Instamart, BigBasket, Reliance Mart), share one real
sample PO file and the column names will be added to
`platform_readers.py`.

---

## CURRENT STATUS / WHAT'S NEXT

✅ Zepto — fully wired and tested against your real PO file
   (29.6% of SKUs in your sample matched the current master file —
   the rest are products not yet in the master mapping; add them via
   "Manage SKU Mapping" as you confirm their SAP codes)

⏳ Blinkit, Swiggy Instamart, BigBasket, Reliance Mart — column
   mapping is pre-configured using assumed column names. Share one
   real sample export from each and these will be corrected to match
   exactly.
