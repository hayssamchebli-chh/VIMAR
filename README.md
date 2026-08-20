# Vimar Datasheet Pack Builder

Paste Vimar codes (or upload an Excel list), download each product's datasheet
from vimar.com and merge everything into one PDF pack — same layout as the
Lightning datasheet pack: a clickable table of contents, then a cover page
carrying the item's Type before each datasheet.

## Run it

```bash
cd C:\Users\hayss\local\vimar-datasheets
pip install -r requirements.txt
python -m playwright install chromium
python -m streamlit run app.py
```

## Codes

The `VIMA` prefix is optional and dots are kept, so all of these work:

```
VIMA-19755.2
VIMA-20582
VIMA-20755.3.B
VIMA-21860
```

Excel files use three columns — **Type**, **Code**, **Description**. The Type is
written on the cover page and used as the table-of-contents entry.

## How it gets past Vimar's human check

Vimar's site sits behind an Imperva/Incapsula firewall that asks visitors to
confirm they are human before opening the catalogue. Plain scripting cannot
pass it, so the app does the honest thing: **you** pass the check once in a
real browser window, and the app then downloads through that same verified
session.

**Step 1 in the app:**

1. Press **Open Vimar browser** — a Chrome window opens on a Vimar product page.
2. Complete the human check in that window if it appears.
3. Press **Check verification**. Once it says *Verified*, the window can stay
   open all day and every download goes through it.

Then paste your codes and press *Download and merge* as normal.

Two details that make it work:

- The PDF is fetched **by the page itself**, because only the page carries the
  verified browser context — Vimar answers anything else with 403.
- Discontinued items are not in the main catalogue, so each code is tried in
  the **product** catalogue and then the **obsolete** one (20582, for example,
  is only in the obsolete catalogue).

If the check reappears mid-run the app reloads and waits for it to clear, and
only then asks you to look at the window.

### Download layers, in order

1. A PDF already on disk (`manual_datasheets/` or your Downloads folder)
2. **The verified browser window** — the one that actually works
3. Vimar's `download-pdf` endpoint through a plain session
4. A fresh Playwright browser

The results table reports which layer each file came from.

### If a code still fails

Failed codes are listed with a **direct link to each product page**. Open the
link, save the data sheet anywhere in your Downloads folder — **no renaming** —
and press *Download and merge* again.

The app finds saved PDFs two ways, so Vimar's own cryptic filenames
(`B_C12005_EE+E_EN.27602.pdf`) work untouched:

- the filename contains the code, **or**
- the code is printed **inside** the PDF (it reads the first pages)

The folder it watches is set by **"Also look for saved PDFs in"** (defaults to
your Downloads folder); `manual_datasheets/` is always searched too. The
*Already saved* metric shows how many of your codes it can already find before
you press anything.

## Files

| File | Purpose |
| --- | --- |
| `app.py` | Streamlit interface |
| `vimar.py` | Code parsing and the four download layers |
| `merge.py` | Cover pages, clickable contents, PDF merge |
| `manual_datasheets/` | Drop manually saved PDFs here |
| `item_type_template.pdf`, `toc_logo.png` | Harb Electric branding |
