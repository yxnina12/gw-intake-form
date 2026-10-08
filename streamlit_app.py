"""
GW Material Intake Form — Spare Parts (Z002)
Deterministic generator. No AI. Nothing is stored: files live in memory for
the duration of your browser session and disappear when you close the tab.
"""

import re
from datetime import date

import pandas as pd
import streamlit as st

import gw_core as G

TEMPLATE_PATH = "templates/Material Intake Form-Spare Parts.xlsx"

st.set_page_config(page_title="GW Intake Form — Spare Parts", page_icon="🔧",
                   layout="wide")


@st.cache_data(show_spinner=False)
def load_template() -> bytes:
    with open(TEMPLATE_PATH, "rb") as f:
        return f.read()


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
st.subheader("1 · Source file")
up = st.file_uploader("Upload the request file (.xlsx)", type=["xlsx"],
                      key="src_upload")

if not up:
    st.info("Upload a request file to start. Both layouts are supported: the "
            "bulk **Intake form** sheet (SAP codes in row 3, data from row 10) "
            "and the flat **Sheet1** extend layout (plain headers in row 1).")
    st.stop()

src_bytes = up.getvalue()
try:
    fmt = G.detect_format(src_bytes)
    if fmt == "bulk":
        sheet_name, all_rows = G.read_bulk_source(src_bytes)
    else:
        sheet_name, all_rows = G.read_flat_source(src_bytes)
except Exception as e:  # noqa: BLE001
    st.error(f"Could not read that file: {e}")
    st.stop()

if not all_rows:
    st.error("No data rows found in that file.")
    st.stop()

st.success(f"Detected **{fmt}** layout · sheet `{sheet_name}` · "
           f"{len(all_rows)} data rows")

# ---------------------------------------------------------------- step 2
st.subheader("2 · Rows to process")


def row_label(r):
    d = r.get("_date")
    try:
        d = d.strftime("%Y-%m-%d")
    except Exception:  # noqa: BLE001
        d = str(d) if d else ""
    return d


preview = pd.DataFrame([{
    "Excel row": r["_row"],
    "Date": row_label(r),
    "Material": G.clean_str(r.get(G.SAP_MATNR)),
    "Description": G.clean_str(r.get(G.SAP_DESC)),
    "Plant (source text)": G.clean_str(r.get(G.SAP_PLANT)),
    "Unit": f'{G.clean_str(r.get(G.SAP_LEN_UNIT))} / {G.clean_str(r.get(G.SAP_WT_UNIT))}',
} for r in all_rows])

if fmt == "bulk":
    lo, hi = all_rows[0]["_row"], all_rows[-1]["_row"]
    c1, c2, c3 = st.columns([1, 1, 2])
    start = c1.number_input("From Excel row", min_value=lo, max_value=hi,
                            value=max(lo, hi - 9))
    end = c2.number_input("To Excel row", min_value=lo, max_value=hi, value=hi)
    c3.caption(f"The file's data rows run from {lo} to {hi}. "
               "Row numbers are the real Excel row numbers, so you can read "
               "them straight off the request e-mail.")
    if end < start:
        st.error("'To' row is before 'From' row.")
        st.stop()
    selected = [r for r in all_rows if start <= r["_row"] <= end]
    st.dataframe(preview[(preview["Excel row"] >= start) &
                         (preview["Excel row"] <= end)],
                 use_container_width=True, hide_index=True)
else:
    st.caption("Flat extend files are small — all rows are processed.")
    selected = all_rows
    st.dataframe(preview, use_container_width=True, hide_index=True)

if not selected:
    st.error("No rows in that range.")
    st.stop()

# duplicate material codes
codes = [G.clean_str(r.get(G.SAP_MATNR)) for r in selected]
dupes = sorted({c for c in codes if c and codes.count(c) > 1})
if dupes:
    st.warning(f"Duplicate material codes in the selection: {', '.join(dupes)} "
               "— check the source file before generating.")
blank_codes = [r["_row"] for r in selected if not G.clean_str(r.get(G.SAP_MATNR))]
if blank_codes:
    st.warning(f"Rows with no material code: {blank_codes}")

st.caption(f"**{len(selected)} row(s) selected.**")

# ---------------------------------------------------------------- step 3
st.subheader("3 · Plants")
cols = st.columns(len(G.PLANT_ORDER))
chosen = []
for i, p in enumerate(G.PLANT_ORDER):
    default = False
    label = p
    if p in G.EU_DEFAULT:
        label = f"{p} ·EU"
    if cols[i].checkbox(label, value=default, key=f"plant_{p}"):
        chosen.append(p)

bq1, bq2 = st.columns([1, 6])
if bq1.button("EU default (3)"):
    for p in G.PLANT_ORDER:
        st.session_state[f"plant_{p}"] = p in G.EU_DEFAULT
    st.rerun()
if bq2.button("US + CA"):
    for p in G.PLANT_ORDER:
        st.session_state[f"plant_{p}"] = p in ("US30", "CA10")
    st.rerun()

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

st.caption("Type or paste the values from the price e-mail. "
           "Only these columns matter — ignore L8HKSR, MS状态, 来源码 and 备注.")

base = pd.DataFrame({
    "Material": codes,
    "Description": [G.clean_str(r.get(G.SAP_DESC)) for r in selected],
})
for c in price_cols:
    base[c] = pd.NA

paste = st.text_area(
    "Optional: paste rows copied from Excel (tab-separated). "
    "First column must be the material code; the remaining numbers are matched "
    "to the price columns below, in order.",
    height=90, key="paste_box", placeholder="R0218071-00\t0.51\t0.59\t0.51")

if st.button("Fill table from pasted text") and paste.strip():
    parsed = {}
    for line in paste.strip().splitlines():
        parts = re.split(r"\t|\s{2,}|,", line.strip())
        parts = [p.strip() for p in parts if p.strip()]
        if not parts:
            continue
        mat = parts[0]
        nums = []
        for p in parts[1:]:
            try:
                nums.append(float(re.sub(r"[^\d.\-]", "", p)))
            except ValueError:
                pass
        parsed[mat] = nums
    for i, m in enumerate(base["Material"]):
        nums = parsed.get(m)
        if nums:
            for j, c in enumerate(price_cols):
                if j < len(nums):
                    base.at[i, c] = nums[j]
    st.session_state["price_df"] = base
    st.success(f"Matched {sum(1 for m in base['Material'] if m in parsed)} "
               f"of {len(base)} material codes.")

if "price_df" in st.session_state:
    prev = st.session_state["price_df"]
    if list(prev["Material"]) == list(base["Material"]) and \
       all(c in prev.columns for c in price_cols):
        base = prev[["Material", "Description"] + price_cols]

edited = st.data_editor(
    base, use_container_width=True, hide_index=True, num_rows="fixed",
    disabled=["Material", "Description"],
    column_config={c: st.column_config.NumberColumn(c, format="%.4f")
                   for c in price_cols},
    key="price_editor")
st.session_state["price_df"] = edited

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

    results, all_warn, all_info = [], [], []
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
                miss = G.check_missing_prices(selected, price_map(p))
                for m, s, l in miss:
                    all_warn.append((p, {"sap": "price", "header": "missing price",
                                         "consensus": None,
                                         "outliers": [(m, f"std={s} l10={l}")],
                                         "col": "-"}))
            name = G.output_filename(out_date, p, vk, int(version))
            results.append((p, vk, is_copy, data, name))
    except Exception as e:  # noqa: BLE001
        st.error(f"Generation failed: {e}")
        st.stop()

    st.session_state["results"] = [(n, d) for (_p, _v, _c, d, n) in results]
    st.session_state["warn"] = all_warn
    st.session_state["info"] = all_info

if "results" in st.session_state:
    warn = st.session_state.get("warn", [])
    info = st.session_state.get("info", [])

    if warn:
        st.error(f"{len(warn)} thing(s) to check before you upload to SAP:")
        for p, w in warn:
            kind = w.get("kind")
            lead = f"- **{p}** · col {w['col']} · `{w['sap']}` ({w['header']})"
            if kind:
                st.write(f"{lead} — **{kind}**: {w['outliers'][:5]}")
            else:
                st.write(f"{lead} — most rows are `{w['consensus']}`, "
                         f"but: {w['outliers'][:5]}")
        st.caption("Each entry shows (Excel row in the output file, value). "
                   "Fix the source file and regenerate, or clear the cell by "
                   "hand after downloading.")
    else:
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
