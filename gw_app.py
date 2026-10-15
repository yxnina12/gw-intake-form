"""
GW Material Intake Form — Spare Parts (Z002)
Deterministic generator. No AI. Nothing is stored: files live in memory for
the duration of your browser session and disappear when you close the tab.
"""

import os
import re
from datetime import date

import pandas as pd
import streamlit as st

import gw_core as G

TEMPLATE_NAME = "Material Intake Form-Spare Parts.xlsx"
# The template can sit either next to app.py or in a templates/ subfolder —
# whichever is easier to upload. First one found wins.
TEMPLATE_CANDIDATES = [TEMPLATE_NAME, f"templates/{TEMPLATE_NAME}"]

st.set_page_config(page_title="GW Intake Form — Spare Parts", page_icon="🔧",
                   layout="wide")


@st.cache_data(show_spinner=False)
def load_template() -> bytes:
    for path in TEMPLATE_CANDIDATES:
        if os.path.exists(path):
            with open(path, "rb") as f:
                return f.read()
    raise FileNotFoundError(
        f"Template not found. Put '{TEMPLATE_NAME}' next to app.py "
        f"(or in a templates/ folder) in the repo.")


def ss(key, default):
    if key not in st.session_state:
        st.session_state[key] = default
    return st.session_state[key]


# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.header("Settings")
    out_date = st.date_input("Batch date", value=date.today(),
                             help="Used for the file name (M.D ...) and for "
                                  "L10 Cost Date inside the file.")
    version = st.number_input("Version", min_value=1, value=1, step=1,
                              help="Bump this if you already saved a file with "
                                   "the same date + plant + sales org today.")
    st.divider()
    st.caption(f"Rules version **{G.CONFIG_VERSION}**")
    st.caption("Template: built-in corrected Spare Parts template "
               "(tabs CA10 / US / NL10 / UK10 / UK11 / UK12 / IT20).")
    st.caption("Save outputs to `Documents\\N_2025\\<YY.M>\\`")
    with st.expander("Privacy"):
        st.write("Files are processed in memory only. Nothing is written to "
                 "disk, logged, or stored. Closing the tab discards everything.")

st.title("GW Material Intake Form — Spare Parts")

# ---------------------------------------------------------------- step 1
st.subheader("1 · Source files")
ups = st.file_uploader(
    "Upload the request file(s) (.xlsx)", type=["xlsx"],
    accept_multiple_files=True, key="src_upload",
    help="You can upload more than one. Rows from every file are merged into "
         "the same output sheet per plant.")

if not ups:
    st.info("Upload a request file to start. Both layouts are supported: the "
            "bulk **Intake form** sheet (SAP codes in row 3, data from row 10) "
            "and the flat **Sheet1** extend layout (plain headers in row 1). "
            "Several files can be uploaded at once — their rows end up in one "
            "sheet per plant.")
    st.stop()


# Everything below is cached on the file bytes (plus the row window). Without
# this, every widget click re-parsed the whole workbook and the app looked
# frozen; and the window keeps a 5,000-row request file as cheap as a small one
# when only the last few rows are wanted (both confirmed 2026-10-15).
@st.cache_data(show_spinner=False, max_entries=10)
def peek(file_bytes: bytes):
    return G.peek_source(file_bytes)


@st.cache_data(show_spinner="Reading those rows...", max_entries=10)
def parse_source(file_bytes: bytes, start: int | None, end: int | None):
    info = G.peek_source(file_bytes)
    if info["fmt"] == "bulk":
        sheet_name, rows = G.read_bulk_source(file_bytes, start or 10, end)
    else:
        sheet_name, rows = G.read_flat_source(file_bytes)
    return info["fmt"], sheet_name, rows


# ---------------------------------------------------------------- step 2
st.subheader("2 · Rows to process")

selected = []
for n, up in enumerate(ups):
    src_bytes = up.getvalue()
    try:
        info = peek(src_bytes)
    except Exception as e:  # noqa: BLE001
        st.error(f"**{up.name}** — could not read that file: {e}")
        st.stop()

    with st.container(border=True):
        st.markdown(f"**{up.name}** · `{info['fmt']}` layout · "
                    f"sheet `{info['sheet_name']}`")
        if info["fmt"] == "bulk":
            lo, hi = info["first_data_row"], info["last_row"]
            st.caption(f"The sheet runs to row **{hi}**. Narrow this to the "
                       "rows you actually need — only those are read. These "
                       "are the real Excel row numbers, so you can type them "
                       "straight off the request e-mail.")
            c1, c2 = st.columns(2)
            start = int(c1.number_input("From Excel row", min_value=lo,
                                        max_value=hi, value=lo,
                                        key=f"row_from_{n}"))
            end = int(c2.number_input("To Excel row", min_value=lo,
                                      max_value=hi, value=hi,
                                      key=f"row_to_{n}"))
            if end < start:
                st.error("'To' row is before 'From' row.")
                st.stop()
        else:
            start, end = None, None
            st.caption("Flat extend file — small, so every row is read.")

        try:
            _fmt, _sheet, rows = parse_source(src_bytes, start, end)
        except Exception as e:  # noqa: BLE001
            st.error(f"**{up.name}** — could not read that file: {e}")
            st.stop()

        if not rows:
            st.error(f"No data rows between {start} and {end}. The sheet "
                     f"claims to run to row {info['last_row']}, but trailing "
                     "rows are often empty formatting — try a lower 'To' row.")
            st.stop()

        st.caption(f"{len(rows)} row(s) from this file.")

    for r in rows:
        r = dict(r)
        r["_src"] = up.name
        selected.append(r)

if not selected:
    st.error("No rows selected.")
    st.stop()


def row_label(r):
    d = r.get("_date")
    try:
        d = d.strftime("%Y-%m-%d")
    except Exception:  # noqa: BLE001
        d = str(d) if d else ""
    return d


multi = len(ups) > 1
preview = pd.DataFrame([{
    **({"Source file": r["_src"]} if multi else {}),
    "Excel row": r["_row"],
    "Date": row_label(r),
    "Material": G.clean_str(r.get(G.SAP_MATNR)),
    "Description": G.clean_str(r.get(G.SAP_DESC)),
    "Plant (source text)": G.clean_str(r.get(G.SAP_PLANT)),
    "Unit": f'{G.clean_str(r.get(G.SAP_LEN_UNIT))} / {G.clean_str(r.get(G.SAP_WT_UNIT))}',
} for r in selected])

st.dataframe(preview, use_container_width=True, hide_index=True)

# duplicate material codes (kept, not dropped — user's call, confirmed 2026-10-15)
codes = [G.clean_str(r.get(G.SAP_MATNR)) for r in selected]
dupes = sorted({c for c in codes if c and codes.count(c) > 1})
if dupes:
    where = {c: [f"{r['_src']} row {r['_row']}" for r in selected
                 if G.clean_str(r.get(G.SAP_MATNR)) == c] for c in dupes}
    st.warning("Duplicate material code(s) — **both rows are kept**, check "
               "before generating: " +
               "; ".join(f"**{c}** ({', '.join(v)})" for c, v in where.items()))
blank_codes = [f"{r['_src']} row {r['_row']}" for r in selected
               if not G.clean_str(r.get(G.SAP_MATNR))]
if blank_codes:
    st.warning(f"Rows with no material code: {', '.join(blank_codes)}")

st.caption(f"**{len(selected)} row(s) selected"
           + (f" from {len(ups)} files." if multi else ".") + "**")

# ---------------------------------------------------------------- step 3
st.subheader("3 · Plants")
cols = st.columns(len(G.PLANT_ORDER))
chosen = []
for i, p in enumerate(G.PLANT_ORDER):
    label = f"{p} ·EU" if p in G.EU_DEFAULT else p
    # No `value=` here on purpose: session_state (set by the quick-select
    # callbacks) is the single source of truth, and passing both logs a
    # "created with a default value but also had its value set" warning.
    if cols[i].checkbox(label, key=f"plant_{p}"):
        chosen.append(p)

# Must be on_click callbacks: a callback runs BEFORE the script re-executes,
# so the checkbox widgets do not exist yet and writing their session_state keys
# is legal. Setting them inline after the checkboxes are drawn raises
# StreamlitWidgetAlreadyInstantiatedError on Streamlit >= 1.6x (seen 2026-10-15).
def _set_plants(targets):
    for q in G.PLANT_ORDER:
        st.session_state[f"plant_{q}"] = q in targets


bq1, bq2 = st.columns([1, 6])
bq1.button("EU default (3)", on_click=_set_plants,
           args=(set(G.EU_DEFAULT),), key="btn_eu")
bq2.button("US + CA", on_click=_set_plants,
           args=({"US30", "CA10"},), key="btn_usca")

if not chosen:
    st.info("Pick at least one plant.")
    st.stop()

plan = G.planned_outputs(chosen)
st.caption("Will produce: " + " · ".join(
    f"**{G.PLANT_CONFIG[p]['tab']}-{vk}**" + (" (Sales-Org copy)" if cp else "")
    for p, vk, cp in plan))
if "NL10" in chosen:
    st.caption("NL10 Sales Org is overridden from SE10 to **UK10** automatically.")

# ---------------------------------------------------------------- step 4
st.subheader("4 · Prices")
needed = []
for p in chosen:
    col = G.PRICE_COLUMN_FOR_PLANT[p]
    if col not in needed:
        needed.append(col)
price_cols = needed + ["26-L10_USD"]

st.caption("From the price e-mail. Only these columns matter — ignore "
           "L8HKSR, MS状态, 来源码 and 备注. "
           "**L8 → Standard Price/MAP** (`MBEW-VERPR`) · "
           "**26-L10_USD → L10 Cost** (`MBEW-ZPLP3`), same value for every plant.")

descs = [G.clean_str(r.get(G.SAP_DESC)) for r in selected]
base = pd.DataFrame({"Material": codes, "Description": descs})
for c in price_cols:
    base[c] = pd.NA

# Pre-fill from the source file's own price columns (confirmed 2026-10-15).
# L10 Cost is currency-neutral so it always pre-fills. The L8 price only
# pre-fills when every chosen plant shares ONE price column -- with US+CA the
# source carries a single Standard Price/MAP that cannot be both, and silently
# copying it into the CA column would ship a wrong price.
prefilled = set()
one_l8 = needed[0] if len(needed) == 1 else None
for i, r in enumerate(selected):
    v = r.get(G.SAP_L10_COST)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        base.at[i, "26-L10_USD"] = float(v)
        prefilled.add(i)
    if one_l8:
        v = r.get(G.SAP_STD_PRICE)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            base.at[i, one_l8] = float(v)
            prefilled.add(i)

if prefilled:
    msg = (f"Pre-filled {len(prefilled)} row(s) **from the source file's own "
           "price columns**. Paste the e-mail prices over them — the e-mail "
           "wins. Where the two disagree, check which is current.")
    if one_l8 is None and len(needed) > 1:
        msg += (f" Only L10 Cost was pre-filled: the source has a single "
                f"Standard Price/MAP and this batch needs {len(needed)} "
                "different L8 columns, so those are left blank on purpose.")
    st.info(msg)

paste = st.text_area(
    "Paste rows copied from the price e-mail (tab-separated). First column is "
    "the material code; the remaining numbers fill the price columns below, "
    "in the order shown.",
    height=90, key="paste_box",
    placeholder="R0218124-00\t" + "\t".join(f"<{c}>" for c in price_cols))


def parse_paste(text):
    parsed = {}
    for line in text.strip().splitlines():
        parts = [q.strip() for q in re.split(r"\t|\s{2,}|,", line.strip())
                 if q.strip()]
        if not parts:
            continue
        nums = []
        for q in parts[1:]:
            try:
                nums.append(float(re.sub(r"[^\d.\-]", "", q)))
            except ValueError:
                pass
        parsed[parts[0]] = nums
    return parsed


# Keep whatever is already in the editor, then apply a NEW paste on top.
# Auto-applying on change (rather than behind a button) is deliberate: the
# button was easy to miss and silently produced price-less files (2026-10-15).
if "price_df" in st.session_state:
    prev = st.session_state["price_df"]
    if list(prev["Material"]) == list(base["Material"]) and \
       all(c in prev.columns for c in price_cols):
        base = prev[["Material", "Description"] + price_cols].copy()

applied = st.session_state.get("paste_applied")
if paste.strip() and paste != applied:
    parsed = parse_paste(paste)
    hit = 0
    for i, m in enumerate(base["Material"]):
        nums = parsed.get(m)
        if not nums:
            continue
        hit += 1
        for j, c in enumerate(price_cols):
            if j < len(nums):
                base.at[i, c] = nums[j]
    st.session_state["paste_applied"] = paste
    st.session_state["price_df"] = base
    unmatched = sorted(set(parsed) - set(base["Material"]))
    if hit:
        st.success(f"Filled {hit} of {len(base)} materials from the pasted text."
                   + (f" Not in this batch: {', '.join(unmatched)}"
                      if unmatched else ""))
    else:
        st.warning("None of the pasted material codes match this batch. "
                   f"Pasted: {', '.join(sorted(parsed)[:5])} … | "
                   f"expected: {', '.join(sorted(set(base['Material']))[:5])}")

edited = st.data_editor(
    base, use_container_width=True, hide_index=True, num_rows="fixed",
    disabled=["Material", "Description"],
    column_config={c: st.column_config.NumberColumn(c, format="%.2f")
                   for c in price_cols},
    key="price_editor")
st.session_state["price_df"] = edited

missing = []
for _, r in edited.iterrows():
    gaps = [c for c in price_cols if pd.isna(r.get(c))]
    if gaps:
        missing.append((str(r["Material"]), gaps))
if missing:
    st.error("Missing price(s) — SAP needs both, these rows would upload "
             "blank:\n" + "\n".join(f"- **{m}** → {', '.join(g)}"
                                     for m, g in missing))
else:
    st.success(f"All {len(edited)} material(s) have every price.")

# ---------------------------------------------------------------- step 5
st.subheader("5 · Generate")

if st.button("Generate files", type="primary"):
    tpl = load_template()

    def price_map(plant):
        col = G.PRICE_COLUMN_FOR_PLANT[plant]
        out = {}
        for _, r in edited.iterrows():
            m = str(r["Material"])
            std = r.get(col)
            l10 = r.get("26-L10_USD")
            std = None if pd.isna(std) else float(std)
            l10 = None if pd.isna(l10) else float(l10)
            out[m] = (std, l10)
        return out

    results, all_warn, all_info, price_warn = [], [], [], []
    try:
        for p, vk, is_copy in plan:
            if is_copy:
                base_bytes = next(b for (pp, vv, cc, b, _n) in results
                                  if pp == p and not cc)
                data = G.make_vkorg_copy(base_bytes, G.PLANT_CONFIG[p]["tab"], vk)
            else:
                override = vk if vk != G.PLANT_CONFIG[p]["fixed"][G.SAP_VKORG] else None
                data = G.build_output(tpl, p, selected, price_map(p),
                                      out_date, vkorg=override)
                w, i = G.scan_output(data, p)
                all_warn += [(p, x) for x in w]
                all_info += [(p, x) for x in i]
                for m, s, l in G.check_missing_prices(selected, price_map(p)):
                    gaps = []
                    if s is None:
                        gaps.append("Standard Price/MAP")
                    if l is None:
                        gaps.append("L10 Cost")
                    price_warn.append((p, m, " and ".join(gaps)))
            name = G.output_filename(out_date, p, vk, int(version))
            results.append((p, vk, is_copy, data, name))
    except Exception as e:  # noqa: BLE001
        st.error(f"Generation failed: {e}")
        st.stop()

    st.session_state["results"] = [(n, d) for (_p, _v, _c, d, n) in results]
    st.session_state["warn"] = all_warn
    st.session_state["info"] = all_info
    st.session_state["price_warn"] = price_warn

if "results" in st.session_state:
    warn = st.session_state.get("warn", [])
    info = st.session_state.get("info", [])
    pwarn = st.session_state.get("price_warn", [])

    if pwarn:
        by_mat = {}
        for p, m, gaps in pwarn:
            by_mat.setdefault((m, gaps), []).append(p)
        st.error("These materials have no price — SAP rejects a blank price, "
                 "so fix them in step 4 and generate again:\n" +
                 "\n".join(f"- **{m}** is missing {gaps} "
                           f"({'/'.join(ps)})"
                           for (m, gaps), ps in by_mat.items()))

    if warn:
        st.warning(f"{len(warn)} column(s) look odd — check before uploading:")
        for p, w in warn:
            kind = w.get("kind")
            lead = f"- **{p}** · `{w['sap']}` ({w['header']}, col {w['col']})"
            if kind:
                st.write(f"{lead} — **{kind}**: {w['outliers'][:5]}")
            else:
                st.write(f"{lead} — most rows are `{w['consensus']}`, "
                         f"but: {w['outliers'][:5]}")
        st.caption("Each entry shows (Excel row in the output file, value). "
                   "Fix the source file and regenerate, or clear the cell by "
                   "hand after downloading.")

    if not warn and not pwarn:
        st.success("Contamination scan clean — no unexpected column varies "
                   "between rows, and every material has both prices.")

    with st.expander(f"Columns that vary by material (expected) — {len(info)}"):
        for p, w in info:
            st.write(f"- **{p}** · `{w['sap']}` ({w['header']})")

    st.divider()
    st.write("### Download")
    dl = st.columns(min(4, len(st.session_state["results"])))
    for i, (name, data) in enumerate(st.session_state["results"]):
        dl[i % len(dl)].download_button(
            name, data=data, file_name=name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{i}")
    st.caption("Save them into `Documents\\N_2025\\"
               f"{str(out_date.year)[2:]}.{out_date.month}\\`. "
               "Keep the US30 file and its US20 copy together.")
