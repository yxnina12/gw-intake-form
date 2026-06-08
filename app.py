"""
GW Intake Form Processor – Streamlit Web App
"""

import io
import sys
import tempfile
import zipfile
from datetime import date
from pathlib import Path

import streamlit as st

# ── Make sure the scripts folder is on the Python path ───────────────────────
sys.path.insert(0, str(Path(__file__).parent / "scripts"))
from process_intake import process, PLANT_CONFIG  # noqa: E402

# ── Constants ─────────────────────────────────────────────────────────────────
SUPPORTED_PLANTS = ["US30", "CA10", "NL10", "UK11"]

PLANT_LABELS = {
    "US30": "US30 — United States",
    "CA10": "CA10 — Canada",
    "NL10": "NL10 — Netherlands",
    "UK11": "UK11 — United Kingdom",
}

# Imperial source plants (unit_of_length == '"')
IMPERIAL_PLANTS = {p for p, cfg in PLANT_CONFIG.items() if cfg["unit_of_length"] == '"'}


def needs_conversion(source_plant: str, target_plant: str) -> bool:
    """True when source is imperial and target expects metric."""
    return source_plant in IMPERIAL_PLANTS and target_plant not in IMPERIAL_PLANTS


# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="GW Intake Form Processor",
    page_icon="📋",
    layout="centered",
)

st.title("📋 GW Intake Form Processor")
st.caption("Spare Parts · Material intake automation")
st.divider()

# ── Step 1: Upload files ──────────────────────────────────────────────────────
st.subheader("Step 1 — Upload files")

col1, col2 = st.columns(2)
with col1:
    request_file = st.file_uploader(
        "Request file *",
        type=["xlsx"],
        help="GW-Form-XXXX Material Intake Form_YYYYMMDD.xlsx  (Includes Intake form + Sheet1 price table)",
    )
with col2:
    template_file = st.file_uploader(
        "Template file *",
        type=["xlsx"],
        help="Material Intake Form-Spare Parts.xlsx",
    )

st.divider()

# ── Step 2: Select plants ─────────────────────────────────────────────────────
st.subheader("Step 2 — Select plants")

col_src, col_tgt = st.columns(2)

with col_src:
    source_plant = st.selectbox(
        "Source plant (request file)",
        options=SUPPORTED_PLANTS,
        format_func=lambda p: PLANT_LABELS[p],
    )

with col_tgt:
    st.markdown("**Target plant(s) (output)**")
    target_plants = []
    for plant in SUPPORTED_PLANTS:
        checked = st.checkbox(PLANT_LABELS[plant], key=f"tgt_{plant}")
        if checked:
            target_plants.append(plant)

# Auto-detect conversion info
if target_plants:
    convert_targets = [p for p in target_plants if needs_conversion(source_plant, p)]
    no_convert_targets = [p for p in target_plants if not needs_conversion(source_plant, p)]

    info_parts = []
    if no_convert_targets:
        info_parts.append(f"**{', '.join(no_convert_targets)}** — no unit conversion")
    if convert_targets:
        info_parts.append(f"**{', '.join(convert_targets)}** — imperial → metric conversion")

    st.info("Auto-detected: " + " · ".join(info_parts))

st.divider()

# ── Step 3: Generate ──────────────────────────────────────────────────────────
st.subheader("Step 3 — Generate & Download")

generate_clicked = st.button(
    "⚙️  Generate",
    type="primary",
    disabled=(request_file is None or template_file is None or len(target_plants) == 0),
)

if generate_clicked:
    today = date.today().strftime("%Y%m%d")
    results = []   # list of (plant, filename, bytes_or_None, error_or_None)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)

        # Save uploaded files to temp dir
        req_path = tmp / "request.xlsx"
        tmpl_path = tmp / "template.xlsx"
        req_path.write_bytes(request_file.getvalue())
        tmpl_path.write_bytes(template_file.getvalue())

        progress = st.progress(0, text="Processing…")

        for i, plant in enumerate(target_plants):
            out_filename = f"GW-IntakeForm-{plant}-{today}.xlsx"
            out_path = tmp / out_filename
            convert = needs_conversion(source_plant, plant)

            try:
                rows = process(
                    request_path=str(req_path),
                    template_path=str(tmpl_path),
                    plant=plant,
                    output_path=str(out_path),
                    convert_units=convert,
                )
                file_bytes = out_path.read_bytes()
                results.append((plant, out_filename, file_bytes, rows, None))
            except Exception as e:
                results.append((plant, out_filename, None, 0, str(e)))

            progress.progress((i + 1) / len(target_plants), text=f"Processed {plant}")

        progress.empty()

    # ── Show results ─────────────────────────────────────────────────────────
    st.markdown("#### Results")

    success_count = sum(1 for r in results if r[4] is None)

    for plant, filename, file_bytes, rows, error in results:
        if error:
            st.error(f"❌ **{plant}** — {error}")
        else:
            col_info, col_dl = st.columns([3, 1])
            with col_info:
                st.success(f"✅ **{filename}** — {rows} data row(s)")
            with col_dl:
                st.download_button(
                    label="⬇ Download",
                    data=file_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{plant}",
                )

    # Offer a ZIP if multiple files were generated successfully
    success_files = [(fn, fb) for (_, fn, fb, _, err) in results if err is None]
    if len(success_files) > 1:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fn, fb in success_files:
                zf.writestr(fn, fb)
        st.download_button(
            label=f"⬇ Download all ({len(success_files)} files) as ZIP",
            data=zip_buf.getvalue(),
            file_name=f"GW-IntakeForm-{today}.zip",
            mime="application/zip",
            key="dl_zip",
        )

# ── Footer ─────────────────────────────────────────────────────────────────────
st.divider()
st.caption("Supported plants: US30 · CA10 · NL10 · UK11")
