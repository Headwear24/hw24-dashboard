"""
HW24 Sales Dashboard Generator — Streamlit App
Upload the raw data file, set the week config, generate all 6 dashboards.
"""

import streamlit as st
import pandas as pd
import io, zipfile, sys, os, importlib, types, tempfile

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="HW24 Dashboard Generator",
    page_icon="📊",
    layout="wide",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .block-container { padding-top: 2rem; }
  .stButton > button { width: 100%; }
  div[data-testid="stSidebarContent"] { padding-top: 1rem; }
  .success-box {
    background: #e6f4ea; border: 1px solid #4caf50;
    border-radius: 8px; padding: 12px 16px;
    color: #1b5e20; font-size: 14px; margin: 8px 0;
  }
  .info-step {
    display: flex; gap: 10px; align-items: flex-start;
    margin-bottom: 10px; font-size: 14px;
  }
  .step-num {
    background: #1F3864; color: white; border-radius: 50%;
    width: 24px; height: 24px; display: flex;
    align-items: center; justify-content: center;
    font-size: 12px; font-weight: 600; flex-shrink: 0; margin-top: 1px;
  }
</style>
""", unsafe_allow_html=True)

# ── Sidebar — configuration ────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Report settings")
    st.markdown("---")

    month = st.selectbox("Month", [
        "June 2026", "July 2026", "August 2026",
        "September 2026", "October 2026", "November 2026",
        "December 2026", "January 2027", "February 2027",
    ])

    week = st.selectbox("Current week", ["Week 1", "Week 2", "Week 3", "Week 4"])

    days_elapsed = st.number_input(
        "Working days elapsed in current week",
        min_value=1, max_value=5, value=4,
        help="Mon=1, Tue=2, Wed=3, Thu=4, Fri=5"
    )

    # Calculate totals
    week_num = int(week.split()[1])
    total_days = (week_num - 1) * 5 + days_elapsed
    total_days_in_month = 20
    pct = days_elapsed / 5 * 100
    pro_rate = total_days / total_days_in_month

    st.markdown(f"""
    **Pro-rating summary**
    - W{week_num} target = **{pct:.0f}%** of weekly ({days_elapsed}/5 days)
    - MTD = **{total_days}/{total_days_in_month}** working days
    - Pro-rated factor = **{pro_rate:.2f}**
    """)

    st.markdown("---")
    st.markdown("**Target methodology**")
    st.markdown("All targets = AFS26 × 1.19")

# ── Main area ──────────────────────────────────────────────────────────────────
st.markdown("# 📊 HW24 Sales Dashboard Generator")
st.markdown(f"**{month}** &nbsp;·&nbsp; {week} &nbsp;·&nbsp; {days_elapsed} days elapsed &nbsp;·&nbsp; {total_days}/{total_days_in_month} MTD working days")
st.markdown("---")

# Steps guide
col_steps, col_upload = st.columns([1, 1.4])

with col_steps:
    st.markdown("#### How it works")
    for num, txt in [
        ("1", "Export the raw data file from your system"),
        ("2", "Set the week and days elapsed in the left panel"),
        ("3", "Upload the file on the right"),
        ("4", "Click Generate — takes about 30 seconds"),
        ("5", "Download all 6 dashboards or as a ZIP"),
    ]:
        st.markdown(f"""
        <div class="info-step">
          <div class="step-num">{num}</div>
          <div>{txt}</div>
        </div>
        """, unsafe_allow_html=True)

with col_upload:
    st.markdown("#### Upload data file")
    uploaded_file = st.file_uploader(
        "Drop your .xlsx file here",
        type=["xlsx"],
        help="File must contain a sheet named 'AFS 26 Data'"
    )

    if uploaded_file:
        st.markdown(f"""
        <div class="success-box">
          ✅ <strong>{uploaded_file.name}</strong> — ready to process
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

# ── Generate button ────────────────────────────────────────────────────────────
if uploaded_file is None:
    st.info("👆 Upload your data file above to get started")
    st.stop()

generate_btn = st.button("🚀 Generate all 6 dashboards", type="primary", use_container_width=True)

if not generate_btn:
    st.stop()

# ── Run generation ─────────────────────────────────────────────────────────────
progress = st.progress(0)
status   = st.empty()

try:
    # Save uploaded file to temp location
    status.markdown("**Reading data file...**")
    progress.progress(5)

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name

    # Load dashboard_generator dynamically with updated config
    status.markdown("**Loading generator...**")
    progress.progress(10)

    import importlib.util, types
    spec = importlib.util.spec_from_file_location("gen", "dashboard_generator.py")
    gen = importlib.util.module_from_spec(spec)

    # Patch the file path and WD_PER before exec
    original_source = open("dashboard_generator.py").read()
    patched_source  = original_source \
        .replace(
            "'/mnt/user-data/uploads/Raw_data_uplaod_file_upto_10th_June26.xlsx'",
            f"'{tmp_path}'"
        ).replace(
            "UPLIFT=1.19; WD_PER=9/20",
            f"UPLIFT=1.19; WD_PER={total_days}/{total_days_in_month}"
        ).replace(
            "w2_tgt=round(weekly_tgt*4/5)",
            f"w2_tgt=round(weekly_tgt*{days_elapsed}/5)"
        )

    # Write patched version to temp file and import
    with tempfile.NamedTemporaryFile(suffix=".py", delete=False, mode="w") as ptmp:
        ptmp.write(patched_source)
        patched_path = ptmp.name

    spec2 = importlib.util.spec_from_file_location("gen_patched", patched_path)
    gen2  = importlib.util.module_from_spec(spec2)

    status.markdown("**Extracting regional metrics...**")
    progress.progress(20)
    spec2.loader.exec_module(gen2)

    regions_cfg = [
        ('Gauteng',       '#185FA5'),
        ('KZN',           '#ED7D31'),
        ('Western Cape',  '#375623'),
        ('Eastern Cape',  '#843C0C'),
        ('International', '#444444'),
    ]

    generated = {}
    for i, (region, col) in enumerate(regions_cfg):
        pct_done = 25 + i * 13
        status.markdown(f"**Building {region} dashboard...**")
        progress.progress(pct_done)
        d = gen2.calc(region)
        d['col'] = col
        html = gen2.build_html(d)
        generated[region] = html

    progress.progress(92)
    status.markdown("**Finalising dashboards...**")

    os.unlink(tmp_path)
    os.unlink(patched_path)

    progress.progress(100)
    status.markdown("✅ **All dashboards generated!**")

except Exception as e:
    st.error(f"❌ Generation failed: {e}")
    import traceback
    st.code(traceback.format_exc())
    st.stop()

# ── Download section ───────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("### 📥 Download dashboards")

REG_COLOURS = {
    'Gauteng': '#185FA5', 'KZN': '#ED7D31', 'Western Cape': '#375623',
    'Eastern Cape': '#843C0C', 'International': '#444444',
}

cols = st.columns(3)
for i, (region, html) in enumerate(generated.items()):
    safe_name = region.replace(' ', '_')
    filename  = f"HW24_{safe_name}_Dashboard.html"
    col = cols[i % 3]
    colour = REG_COLOURS.get(region, '#1F3864')
    with col:
        st.markdown(f"<div style='border-left:4px solid {colour};padding-left:10px;margin-bottom:4px'><strong>{region}</strong></div>", unsafe_allow_html=True)
        st.download_button(
            label=f"⬇ Download {region}",
            data=html.encode("utf-8"),
            file_name=filename,
            mime="text/html",
            key=f"dl_{region}",
            use_container_width=True,
        )

st.markdown("---")

# ZIP all
zip_buf = io.BytesIO()
with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
    for region, html in generated.items():
        safe = region.replace(' ', '_')
        zf.writestr(f"HW24_{safe}_Dashboard.html", html.encode("utf-8"))
zip_buf.seek(0)

month_safe = month.replace(' ', '_')
st.download_button(
    label="📦 Download ALL 5 regional dashboards as ZIP",
    data=zip_buf,
    file_name=f"HW24_Regional_Dashboards_{month_safe}.zip",
    mime="application/zip",
    use_container_width=True,
    type="primary",
)

st.markdown(f"""
<br>
<div style='text-align:center;color:#888;font-size:12px'>
  HW24 Group Sales &nbsp;·&nbsp; Generated {month} &nbsp;·&nbsp; {week} ({days_elapsed} days elapsed)
</div>
""", unsafe_allow_html=True)
