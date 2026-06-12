"""
HW24 Sales Dashboard Generator — Streamlit App
Generates all 6 dashboards (1 national + 5 regional) from a single data upload.
"""
import streamlit as st
import pandas as pd
import io, zipfile, sys, os, tempfile, importlib
 
st.set_page_config(page_title="HW24 Dashboard Generator", page_icon="📊", layout="wide")
 
st.markdown("""
<style>
.block-container{padding-top:2rem}
.stButton>button{width:100%}
.success-box{background:#e6f4ea;border:1px solid #4caf50;border-radius:8px;
  padding:12px 16px;color:#1b5e20;font-size:14px;margin:8px 0}
.info-step{display:flex;gap:10px;align-items:flex-start;margin-bottom:10px;font-size:14px}
.step-num{background:#1F3864;color:white;border-radius:50%;width:24px;height:24px;
  display:flex;align-items:center;justify-content:center;font-size:12px;
  font-weight:600;flex-shrink:0;margin-top:1px}
</style>
""", unsafe_allow_html=True)
 
# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Report settings")
    st.markdown("---")
    month = st.selectbox("Month", [
        "June 2026","July 2026","August 2026","September 2026",
        "October 2026","November 2026","December 2026",
        "January 2027","February 2027",
    ])
    week = st.selectbox("Current week", ["Week 1","Week 2","Week 3","Week 4"])
    days_elapsed = st.number_input(
        "Working days elapsed in current week",
        min_value=1, max_value=5, value=4,
        help="Mon=1  Tue=2  Wed=3  Thu=4  Fri=5"
    )
    week_num   = int(week.split()[1])
    total_days = (week_num - 1) * 5 + days_elapsed
    wd_total   = 20
    wd_per     = total_days / wd_total
    w2_frac    = days_elapsed / 5
 
    st.markdown(f"""
    **Pro-rating summary**
    - W{week_num} target = **{w2_frac*100:.0f}%** of weekly ({days_elapsed}/5 days)
    - MTD = **{total_days}/{wd_total}** working days
    - Pro-rated factor = **{wd_per:.2f}**
    """)
    st.markdown("---")
    st.markdown("**Target methodology:** All targets = AFS26 × 1.19")
 
# ── Main ───────────────────────────────────────────────────────────────────────
st.markdown("# 📊 HW24 Sales Dashboard Generator")
st.markdown(f"**{month}** · {week} · {days_elapsed} days elapsed · {total_days}/{wd_total} MTD working days")
st.markdown("---")
 
col_steps, col_upload = st.columns([1, 1.4])
with col_steps:
    st.markdown("#### How it works")
    for num, txt in [
        ("1","Export raw data from your system"),
        ("2","Set week and days elapsed in the left panel"),
        ("3","Upload the .xlsx file on the right"),
        ("4","Click Generate — takes about 45 seconds"),
        ("5","Download all 6 dashboards or as a single ZIP"),
    ]:
        st.markdown(
            f'<div class="info-step"><div class="step-num">{num}</div><div>{txt}</div></div>',
            unsafe_allow_html=True)
 
with col_upload:
    st.markdown("#### Upload data file")
    uploaded_file = st.file_uploader(
        "Drop your .xlsx file here", type=["xlsx"],
        help="Must contain a sheet named 'AFS 26 Data'")
    if uploaded_file:
        st.markdown(
            f'<div class="success-box">✅ <strong>{uploaded_file.name}</strong> — ready to process</div>',
            unsafe_allow_html=True)
 
st.markdown("---")
if uploaded_file is None:
    st.info("👆 Upload your data file above to get started")
    st.stop()
 
if not st.button("🚀 Generate all 6 dashboards", type="primary", use_container_width=True):
    st.stop()
 
# ── Generate ───────────────────────────────────────────────────────────────────
progress = st.progress(0)
status   = st.empty()
 
try:
    status.markdown("**Saving uploaded file...**")
    progress.progress(5)
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name
 
    status.markdown("**Loading generator...**")
    progress.progress(10)
 
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import dashboard_generator as gen
    importlib.reload(gen)
 
    status.markdown("**Reading data file...**")
    progress.progress(20)
    gen.initialise(tmp_path, wd_per=wd_per, w2_fraction=w2_frac)
 
    # ── National dashboard ─────────────────────────────────────────────────────
    import datetime
    today_str = datetime.date.today().strftime("%-d %b %Y")
    status.markdown("**Building National dashboard...**")
    progress.progress(30)
    national_html = gen.generate_national(
        report_date=today_str,
        week_label=week,
        days_elapsed=int(days_elapsed)
    )
 
    # ── Regional dashboards ────────────────────────────────────────────────────
    regions_cfg = [
        ('Gauteng',       '#185FA5'),
        ('KZN',           '#ED7D31'),
        ('Western Cape',  '#375623'),
        ('Eastern Cape',  '#843C0C'),
        ('International', '#444444'),
    ]
    generated = {'National': national_html}
 
    for i, (region, col) in enumerate(regions_cfg):
        pct_done = 40 + i * 12
        status.markdown(f"**Building {region} dashboard...**")
        progress.progress(pct_done)
        d = gen.calc(region)
        d['col'] = col
        html = gen.build_html(d)
        generated[region] = html
 
    os.unlink(tmp_path)
    progress.progress(100)
    status.markdown("✅ **All 6 dashboards generated!**")
 
except Exception as e:
    st.error(f"❌ Generation failed: {e}")
    import traceback
    st.code(traceback.format_exc())
    st.stop()
 
# ── Downloads ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("### 📥 Download dashboards")
 
COLOURS = {
    'National':      '#1F3864',
    'Gauteng':       '#185FA5',
    'KZN':           '#ED7D31',
    'Western Cape':  '#375623',
    'Eastern Cape':  '#843C0C',
    'International': '#444444',
}
 
cols = st.columns(3)
for i, (name, html) in enumerate(generated.items()):
    col    = cols[i % 3]
    colour = COLOURS.get(name, '#1F3864')
    safe   = name.replace(' ', '_')
    with col:
        st.markdown(
            f"<div style='border-left:4px solid {colour};padding-left:10px;margin-bottom:4px'>"
            f"<strong>{name}</strong></div>",
            unsafe_allow_html=True)
        st.download_button(
            label=f"⬇ Download {name}",
            data=html.encode("utf-8"),
            file_name=f"HW24_{safe}_Dashboard.html",
            mime="text/html",
            key=f"dl_{name}",
            use_container_width=True,
        )
 
st.markdown("---")
zip_buf = io.BytesIO()
with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
    for name, html in generated.items():
        zf.writestr(f"HW24_{name.replace(' ','_')}_Dashboard.html", html.encode("utf-8"))
zip_buf.seek(0)
 
st.download_button(
    label="📦 Download ALL 6 dashboards as ZIP",
    data=zip_buf,
    file_name=f"HW24_All_Dashboards_{month.replace(' ','_')}.zip",
    mime="application/zip",
    use_container_width=True,
    type="primary",
)
 
st.markdown(
    f"<br><div style='text-align:center;color:#888;font-size:12px'>"
    f"HW24 Group Sales · {month} · {week} · {days_elapsed} days elapsed</div>",
    unsafe_allow_html=True)
