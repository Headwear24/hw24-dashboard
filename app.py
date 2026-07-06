"""
HW24 Sales Dashboard Generator — Streamlit App
Generates all 6 dashboards (1 national + 5 regional) from a single data upload.
 
The dashboard generator is fully DATA-DRIVEN: it detects the reporting month,
the weeks present, the working days elapsed, and whether the month has 4 or 5
weeks directly from the uploaded file. This app therefore does not ask the user
to set those manually — the sidebar reflects what the generator derived after
the file is read.
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
.derived{background:#eef3fa;border:1px solid #b7cce8;border-radius:8px;
  padding:10px 14px;font-size:13px;color:#1F3864;margin:8px 0}
</style>
""", unsafe_allow_html=True)
 
# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Report settings")
    st.markdown("---")
    st.markdown(
        "<div class='derived'>📈 <strong>Auto-detected from your data</strong><br>"
        "The generator reads the reporting month, the weeks present, the working "
        "days elapsed, and whether the month has 4 or 5 weeks directly from the "
        "uploaded file. Nothing needs to be set here.</div>",
        unsafe_allow_html=True)
    # Placeholder that gets filled in after the file is read (see below).
    sidebar_summary = st.empty()
    st.markdown("---")
    st.markdown("**Target methodology:** All targets = AFS26 × 1.19")
 
# ── Main ───────────────────────────────────────────────────────────────────────
st.markdown("# 📊 HW24 Sales Dashboard Generator")
st.markdown("Reporting period, weeks and pro-rating are detected automatically from the data.")
st.markdown("---")
 
col_steps, col_upload = st.columns([1, 1.4])
with col_steps:
    st.markdown("#### How it works")
    for num, txt in [
        ("1","Export raw data from your system"),
        ("2","Upload the .xlsx file on the right"),
        ("3","Click Generate — takes about 45 seconds"),
        ("4","Review the auto-detected period in the sidebar"),
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
    # Fully data-driven: initialise derives month, weeks, days elapsed and 4/5-week count.
    gen.initialise(tmp_path)
 
    # ── Show what the generator detected (sidebar reflects the data) ────────────
    total_weeks = getattr(gen, "TOTAL_WEEKS", 4)
    cur_week    = getattr(gen, "CUR_WEEK_LABEL", "Week 1")
    cur_days    = getattr(gen, "CUR_WEEK_DAYS", 5)
    weeks_seen  = getattr(gen, "WEEKS_PRESENT", [])
    report_date = getattr(gen, "REPORT_DATE", "")
    cur_month   = getattr(gen, "CUR_MONTH", "")
    n_full      = max(len(weeks_seen) - 1, 0)
    wd_elapsed  = n_full * 5 + cur_days
    wd_total    = total_weeks * 5
    pro_factor  = wd_elapsed / wd_total if wd_total else 0
 
    sidebar_summary.markdown(f"""
**Detected reporting period**
- Report date: **{report_date}**
- Month: **{cur_month}** — a **{total_weeks}-week** month
- Weeks with data: **{len(weeks_seen)}** ({', '.join(weeks_seen) if weeks_seen else '—'})
- Current week: **{cur_week}** ({cur_days}/5 days)
- MTD = **{wd_elapsed}/{wd_total}** working days
- Pro-rated factor = **{pro_factor:.2f}**
""")
 
    # ── National dashboard ─────────────────────────────────────────────────────
    status.markdown("**Building National dashboard...**")
    progress.progress(30)
    national_html = gen.generate_national()   # report date & weeks derived from data
 
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
    status.markdown(f"✅ **All 6 dashboards generated** — {cur_month} ({total_weeks}-week month), report date {report_date}")
 
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
    file_name=f"HW24_All_Dashboards_{cur_month}.zip",
    mime="application/zip",
    use_container_width=True,
    type="primary",
)
 
st.markdown(
    f"<br><div style='text-align:center;color:#888;font-size:12px'>"
    f"HW24 Group Sales · {cur_month} · {cur_week} · {cur_days} days in current week · "
    f"{total_weeks}-week month</div>",
    unsafe_allow_html=True)
