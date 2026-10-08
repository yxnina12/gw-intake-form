# GW Material Intake Form — Spare Parts

A small web app that turns a GW Material Intake Form request file into the
plant-specific SAP upload files, following the exact rules we use today.

**There is no AI in this app.** Every rule is written out in `gw_core.py`:
field mapping by SAP code, per-plant fixed values, unit conversion, country-code
truncation, EAN text formatting, the US20 copy, the NL10 Sales-Org override,
the contamination scan, external-link stripping, and the file naming.

Rules version: **2026-09-28**.

---

## What it does

| Step | Who does it |
|---|---|
| Read the request file, map fields by SAP code | app |
| Apply each plant's fixed values | app |
| Convert inch/LB → CM/KG per row where the target plant needs it | app |
| Truncate Country of Origin to 2 letters, force EAN/commodity code to text | app |
| Blank out Volume (untrustworthy in both source layouts) | app |
| Produce the US20 Sales-Org copy, override NL10 Sales Org to UK10 | app |
| Scan every output for contaminated / non-numeric cells | app |
| Strip the template's garbage external links | app |
| Name the files `M.D Plant-SalesOrg-version.xlsx` | app |
| **Enter the prices** | you (paste from the price e-mail) |
| **Pick the rows and the plants** | you |
| Save the downloads into `Documents\N_2025\<YY.M>\` | you |

Supported plants: **US30, CA10, NL10, UK11, IT20** (plus UK10 and UK12 if ever
needed). IT10 was retired on 2026-09-28 — Italy output is IT20 only.

Both source layouts work: the bulk **Intake form** sheet (SAP codes in row 3,
data from row 10) and the flat **Sheet1** extend layout (plain headers in row 1).

Spare Parts (Z002) only for now. Finished Goods (Z001) is a different template
and is not wired up yet.

---

## Deploying it (all in a browser — nothing to install)

### 1. Put the files in the GitHub repo

Open `github.com/yxnina12/gw-intake-form` in a browser.

If the repo already has old files, delete them first: click a file → the trash
icon → **Commit changes**. The old code is out of date and should not be left
next to the new code.

Then **Add file → Upload files**, drag in:

```
app.py
gw_core.py
requirements.txt
README.md
.gitignore
Material Intake Form-Spare Parts.xlsx
```

All six go in the top level of the repo — no folders needed. Then
**Commit changes**.

(The app also accepts the template in a `templates/` subfolder if you ever
prefer to tidy it away; it looks in both places.)

*Why:* Streamlit Cloud builds the app straight from this repo. `requirements.txt`
tells it to install `streamlit`, `openpyxl` and `pandas`; `.gitignore` stops
request files or generated outputs from ever being committed by accident.

### 2. Make sure the repo is private

Repo → **Settings** → scroll to **Danger Zone** → **Change repository
visibility** → Private.

*Why:* the blank template lives in the repo. It holds no material data, but
private costs nothing and keeps it off the public internet.

### 3. Deploy

Go to `share.streamlit.io` and sign in with GitHub.

- If the old app is still listed, it will redeploy by itself within a minute or
  two of the commit. Nothing else to do.
- If not: **Create app** → **Deploy from GitHub** → pick
  `yxnina12/gw-intake-form` → branch `main` → **Main file path**
  `app.py` → **Deploy**.

First build takes 1–3 minutes while it installs the three libraries.

### 4. Lock it down

In the app → **⋮ menu → Settings → Sharing** → choose
**"Only specific people can view this app"** and add your e-mail address.

*Why:* without this, anyone with the link can open it. Note two limits of the
free tier: you can only have **one** private app at a time, and **people you
invite can invite others**.

### 5. Use it

Open the URL, sign in, and work down the page: upload → pick rows → pick
plants → paste prices → **Generate files** → download.

---

## Privacy

Files are read into memory, processed, and handed back as downloads. The app
writes nothing to disk, keeps no database, and logs no file contents. Closing
the browser tab discards everything.

That said, the file does travel to Streamlit's servers (US-hosted) while it is
being processed. If that is ever a problem, the same code runs locally with
`streamlit run app.py` on any machine that has Python.

---

## Changing the rules later

Everything lives in `PLANT_CONFIG` at the top of `gw_core.py`, one block per
plant, with a dated comment on every value that was changed. Edit it in the
GitHub web editor, commit, and Streamlit redeploys in about a minute.

Keep `gw_core.py` and the `gw-intake-form` skill in sync. **If they ever
disagree, the skill is the source of truth and this file gets fixed.**

### Known-good regression test

The engine was checked against six files that were really uploaded to SAP
(the 2026-09-28 batch): it reproduces all six exactly, with differences only in
the four values that were deliberately changed on 2026-09-28 (US30 planned
delivery time 77→87, CA10 GR processing time 2→4, CA10 Sales Unit →PCA,
UK11 Reorder Point →1). Re-run that comparison after any change to
`PLANT_CONFIG`.
