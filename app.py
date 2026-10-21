import os, uuid
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
from datetime import date, datetime
import pandas as pd
import streamlit as st
import qrcode
import plotly.express as px
import plotly.graph_objects as go
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader

st.set_page_config(page_title="Tool Station Pro", page_icon="🧰", layout="wide")
ROOT=Path(__file__).resolve().parent; FILE=ROOT/"tools.xlsx"; PHOTO=ROOT/"photos"; QR=ROOT/"qr_labels"
PHOTO.mkdir(exist_ok=True); QR.mkdir(exist_ok=True)
SHEETS={
"Tools":["tool_id","make","model","category","power_source","voltage","battery_platform","serial_number","purchase_date","purchase_price","vendor","condition","storage_location","location_area","location_cabinet","location_shelf","location_bin","warranty_expiration","photo_file","notes"],
"Maintenance":["maintenance_id","tool_id","service_date","service_type","description","cost","next_service_date","performed_by","notes"],
"Accessories":["accessory_id","name","type","make","model","battery_platform","quantity","compatible_tools","storage_location","location_area","location_cabinet","location_shelf","location_bin","condition","notes"],
"Dropdowns":["option_type","option_value"]}
POWER=["Corded electric","Battery","Gas","Pneumatic","Manual","Other"]
CONDITIONS=["New","Excellent","Good","Fair","Needs repair","Retired"]
SERVICES=["Inspection","Cleaning","Lubrication","Repair","Battery replacement","Blade / bit replacement","Calibration","Other"]
TYPES=["Battery","Charger","Blade","Bit","Attachment","Case","Consumable","Other"]
NUMERIC_COLUMNS={"Tools":["purchase_price"],"Maintenance":["cost"]}
OPTION_DEFAULTS={
    "Make / Brand":["Bosch","Craftsman","DeWalt","Festool","Makita","Metabo HPT","Milwaukee","Other","Ridgid","Ryobi"],
    "Category":["Air Compressor","Circular Saw","Drill","Grinder","Hammer Drill","Impact Driver","Jigsaw","Miter Saw","Nail Gun","Oscillating Tool","Planer","Router","Sander","Table Saw","Other"],
    "Voltage":["12V","18V","20V","24V","36V","40V","60V","120V","120V / 240V","Corded / N/A","Other"],
    "Battery platform":["DeWalt 20V MAX","Milwaukee M18","Milwaukee M12","Makita LXT 18V","Makita XGT 40V","Bosch 18V","Ryobi ONE+ 18V","Ridgid 18V","Craftsman V20","Other"],
    "Vendor":["Home Depot","Lowe’s","Amazon","Ace Hardware","Harbor Freight","Northern Tool","Walmart","Grainger","Other"]
}
def uid(): return str(uuid.uuid4())[:8].upper()
def empty(cols): return pd.DataFrame(columns=cols)
def normalize_frame(sheet, frame):
    """Keep ordinary fields as strings and designated numeric fields as floats/NaN."""
    d=frame.copy()
    for c in SHEETS[sheet]:
        if c not in d: d[c]=""
    for c in SHEETS[sheet]:
        if c in NUMERIC_COLUMNS.get(sheet, []):
            d[c]=pd.to_numeric(d[c].replace(r"^\s*$", pd.NA, regex=True), errors="coerce")
        else:
            d[c]=d[c].map(lambda v: "" if pd.isna(v) else str(v))
    return d[SHEETS[sheet]]
def ensure():
    old={}
    if FILE.exists():
        try:
            with pd.ExcelFile(FILE) as x:
                old={s:pd.read_excel(x,sheet_name=s) for s in x.sheet_names}
        except Exception: pass
    with pd.ExcelWriter(FILE,engine="openpyxl") as w:
        for s,cols in SHEETS.items():
            d=normalize_frame(s,old.get(s,empty(cols)))
            if s=="Dropdowns" and d.empty:
                d=pd.DataFrame([{"option_type":kind,"option_value":value} for kind,values in OPTION_DEFAULTS.items() for value in values],columns=cols)
            if s=="Dropdowns" and not (d.option_type=="Vendor").any():
                d=pd.concat([d,pd.DataFrame([{"option_type":"Vendor","option_value":v} for v in OPTION_DEFAULTS["Vendor"]])],ignore_index=True)
            d.to_excel(w,index=False,sheet_name=s)
def load(s):
    if not FILE.exists(): ensure()
    d=pd.read_excel(FILE,sheet_name=s)
    return normalize_frame(s,d)
def save(frames):
    tmp=FILE.with_suffix(".tmp.xlsx")
    with pd.ExcelWriter(tmp,engine="openpyxl") as w:
        for s,cols in SHEETS.items():
            d=normalize_frame(s,frames.get(s,load(s)))
            # NaN in numeric columns writes as an empty Excel cell, not an empty string.
            d.to_excel(w,index=False,sheet_name=s)
    os.replace(tmp,FILE)
def date_value(v):
    """Convert an Excel-stored date to a date_input value; blank/invalid values stay empty."""
    if v is None or not str(v).strip(): return None
    try: return pd.to_datetime(v, errors="raise").date()
    except (ValueError, TypeError): return None
def date_text(v):
    return v.isoformat() if v else ""
def valdate(v,label="Date",required=False):
    """Return an error message for invalid ISO dates; allow blank optional dates."""
    text=str(v or "").strip()
    if not text:
        return f"{label} is required." if required else ""
    try:
        datetime.strptime(text, "%Y-%m-%d")
        return ""
    except (ValueError, TypeError):
        return f"{label} must be a valid date in YYYY-MM-DD format."
def money(v,label="Amount"):
    if v is None or not str(v).strip(): return None,""
    try: return float(str(v).replace("$","").replace(",","")),""
    except ValueError: return None,f"{label} must be numeric."
def loc(area,cabinet,shelf,bin_name,fallback=""):
    return " / ".join(str(x).strip() for x in [area,cabinet,shelf,bin_name] if str(x).strip()) or fallback
def showphoto(p,width=170):
    if p and (ROOT/p).exists(): st.image(str(ROOT/p),width=width)
    else: st.caption("No photo")
def savephoto(upload,tid):
    if not upload: return ""
    ext=Path(upload.name).suffix.lower()
    if ext not in [".jpg",".jpeg",".png",".webp"]: raise ValueError("Use JPG, PNG, or WEBP.")
    name=tid+ext; (PHOTO/name).write_bytes(upload.getvalue()); return "photos/"+name
def complete_backup_zip():
    """Create a GitHub-ready snapshot of the running application and its local data."""
    buffer=BytesIO()
    with ZipFile(buffer,"w",compression=ZIP_DEFLATED,compresslevel=6) as archive:
        # Use the workbook as stored on disk; it includes every inventory sheet.
        for filename in ("app.py","requirements.txt","README.txt","tools.xlsx","tool_qr_labels.pdf"):
            path=ROOT/filename
            if path.is_file() and not path.is_symlink():
                archive.write(path,arcname=filename)
        for dirname in ("photos","qr_labels"):
            folder=ROOT/dirname
            archive.writestr(dirname+"/.keep", "")
            if folder.is_dir():
                for path in sorted(folder.rglob("*")):
                    if path.is_file() and not path.is_symlink() and path.name != ".keep":
                        archive.write(path,arcname=path.relative_to(ROOT).as_posix())
        archive.writestr("RESTORE_INSTRUCTIONS.txt", """TOOL INVENTORY PRO - RESTORE INSTRUCTIONS

1. Extract this ZIP to a local folder and confirm app.py, requirements.txt,
   tools.xlsx, photos/ and qr_labels/ are present.
2. In your GitHub repository, upload these extracted files while preserving
   the photos/ and qr_labels/ paths, then commit your changes.
   You can also use git add/commit/push from your local clone.
3. Streamlit Community Cloud will redeploy from the updated repository.
4. Verify your tools, maintenance, accessories, photos, and QR labels.

WARNING: This backup contains only data still present in the running app.
Streamlit Community Cloud local files are ephemeral; create and save backups
regularly. Photos and inventory may be sensitive: use a PRIVATE repository
and appropriate access controls. Avoid committing credentials or secrets.
""")
    return buffer.getvalue()

def toolname(r): return f"{r['tool_id']} — {r['make']} {r['model']}"
def dropdown_options(kind, current_values=()):
    vals=[]
    if "dropdowns" in globals() and not dropdowns.empty:
        vals.extend(dropdowns.loc[dropdowns.option_type==kind,"option_value"].astype(str).tolist())
    if not vals:
        vals.extend(OPTION_DEFAULTS.get(kind, []))
    vals.extend(str(v).strip() for v in current_values if str(v).strip())
    return [""]+sorted({v.strip() for v in vals if v.strip()},key=str.casefold)
def location(r): return loc(r.get("location_area",""),r.get("location_cabinet",""),r.get("location_shelf",""),r.get("location_bin",""),r.get("storage_location",""))
ensure(); tools=load("Tools"); maint=load("Maintenance"); acc=load("Accessories"); dropdowns=load("Dropdowns")
st.title("🧰 Tool Station Pro")
st.caption("Excel backend • photos • QR labels • maintenance • accessories • storage locations • reports")
today=pd.Timestamp.today().normalize()
purchase=pd.to_numeric(tools.purchase_price,errors="coerce").sum()
service=pd.to_numeric(maint.cost,errors="coerce").sum()
# Dashboard theme: pink highlights, clean light cards, and dark sidebar.
st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&display=swap');
:root { --accent-pink: #FF417D; }
[data-testid="stAppViewContainer"] { background: #f6f8fc; }
[data-testid="stSidebar"] { background: #202c39; }
[data-testid="stSidebar"] { min-width: 300px !important; width: 300px !important; }
/* Remove Streamlit's default inset so navigation begins at the sidebar edge. */
[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] { padding-left: 0 !important; padding-right: 0 !important; }
[data-testid="stSidebar"] [data-testid="stSidebarContent"] { padding-left: 0 !important; padding-right: 0 !important; }
[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] > div { padding-left: 0 !important; padding-right: 0 !important; }
/* Flush sidebar navigation to the left, without changing icon/text sizing or gaps. */
[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] > div,
[data-testid="stSidebar"] [data-testid="stVerticalBlock"],
[data-testid="stSidebar"] [data-testid="stElementContainer"]:has([data-testid="stButton"]) {
    padding-left: 0 !important;
    margin-left: 0 !important;
}
[data-testid="stSidebar"] [data-testid="stButton"] > button {
    margin-left: 0 !important;
    padding-left: 0 !important;
}
/* Align the entire sidebar navigation to the actual left content edge. */
[data-testid="stSidebar"] [data-testid="stSidebarUserContent"],
[data-testid="stSidebar"] [data-testid="stSidebarContent"],
[data-testid="stSidebar"] [data-testid="stVerticalBlock"],
[data-testid="stSidebar"] [data-testid="stElementContainer"],
[data-testid="stSidebar"] [data-testid="stButton"],
[data-testid="stSidebar"] [data-testid="stButton"] > button {
    margin-inline-start: 0 !important;
    padding-inline-start: 0 !important;
    text-indent: 0 !important;
}
[data-testid="stSidebar"] [data-testid="stButton"] > button {
    justify-content: flex-start !important;
    text-align: left !important;
    direction: ltr !important;
}
[data-testid="stSidebar"] [data-testid="stButton"] > button > div,
[data-testid="stSidebar"] [data-testid="stButton"] > button > span {
    margin-inline-start: 0 !important;
    text-align: left !important;
}
/* Compact, vertically stacked navigation, with a fixed icon column. */
/* Streamlit adds its own vertical block gaps; override those as well. */
[data-testid="stSidebar"] [data-testid="stVerticalBlock"]:has([data-testid="stButton"]) { gap: 0 !important; row-gap: 0 !important; }
[data-testid="stSidebar"] [data-testid="stElementContainer"]:has([data-testid="stButton"]) { margin: 0 !important; padding: 0 !important; }
[data-testid="stSidebar"] [data-testid="stButton"] { margin: 0 0 5px !important; width: 100%; padding: 0 !important; }
[data-testid="stSidebar"] [data-testid="stButton"] > button {
    display: flex !important;
    flex-direction: row !important;
    justify-content: flex-start !important;
    align-items: center !important;
    flex-wrap: nowrap !important;
    width: 100% !important;
    min-height: 36px !important;
    padding: 4px 0 !important;
    gap: 25px !important;
    background: transparent !important;
    border: 0 !important;
    border-radius: 8px !important;
    box-shadow: none !important;
    text-align: left !important;
    color: #fff !important;
    font-family: "Montserrat", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
    font-size: 16px !important;
    font-weight: 500 !important;
}
[data-testid="stSidebar"] [data-testid="stButton"] > button:hover { background: rgba(255,255,255,.055) !important; }
[data-testid="stSidebar"] [data-testid="stButton"] > button:focus,
[data-testid="stSidebar"] [data-testid="stButton"] > button:focus-visible,
[data-testid="stSidebar"] [data-testid="stButton"] > button:active { background: transparent !important; box-shadow: none !important; outline: none !important; }
[data-testid="stSidebar"] [data-testid="stButton"] > button p {
    margin: 0 !important;
    padding: 0 !important;
    font-family: "Montserrat", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
    font-size: 16px !important;
    font-weight: 500 !important;
    line-height: 1.35 !important;
    text-align: left !important;
    white-space: nowrap !important;
    color: white !important;
}
[data-testid="stSidebar"] [data-testid="stButton"] > button [data-testid="stIconMaterial"],
[data-testid="stSidebar"] [data-testid="stButton"] > button span.material-symbols-rounded,
[data-testid="stSidebar"] [data-testid="stButton"] > button span.material-symbols-outlined {
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    font-size: 18px !important;
    width: 18px !important;
    min-width: 18px !important;
    flex: 0 0 18px !important;
    margin: 0 !important;
    color: #fff !important;
}
[data-testid="stSidebar"] [data-testid="stButton"] > button svg {
    width: 18px !important;
    height: 18px !important;
    min-width: 18px !important;
    flex: 0 0 18px !important;
    margin: 0 !important;
    color: #fff !important;
}
/* Force Streamlit's inner button content wrapper to use left alignment, too. */
[data-testid="stSidebar"] [data-testid="stButton"] > button {
    justify-content: flex-start !important;
    text-align: left !important;
}
[data-testid="stSidebar"] [data-testid="stButton"] > button > div,
[data-testid="stSidebar"] [data-testid="stButton"] > button [data-testid="stMarkdownContainer"] {
    display: flex !important;
    flex-direction: row !important;
    justify-content: flex-start !important;
    align-items: center !important;
    text-align: left !important;
    margin-left: 0 !important;
    margin-right: auto !important;
    width: auto !important;
    flex-grow: 0 !important;
}
[data-testid="stSidebar"] [data-testid="stButton"] > button p {
    text-align: left !important;
    margin-left: 0 !important;
    margin-right: auto !important;
}
[data-testid="stSidebar"] * { color: #f3f6fa; }
[data-testid="stSidebar"] [data-testid="stMetricLabel"] p,
[data-testid="stSidebar"] [data-testid="stMetricLabel"] { color: #FF417D !important; }
[data-testid="stSidebar"] [data-testid="stMetricValue"] * { color: #fff !important; }
[data-testid="stSidebar"] [data-testid="stMetric"] { border-bottom: 1px solid rgba(255,255,255,.10); padding: .4rem 0 .65rem; }
.dashboard-card { background: white; border: 1px solid #e0e6ee; border-radius: 11px; padding: 18px 20px; min-height: 118px; box-shadow: 0 2px 9px rgba(25,43,65,.035); }
/* Match overview charts and tables to the white, borderless metric cards. */
div[class*="st-key-overview_"] {
    background: #ffffff !important;
    border: none !important;
    border-radius: 11px !important;
    box-shadow: 0 2px 9px rgba(25,43,65,.035) !important;
    padding: 14px 18px !important;
}
div[class*="st-key-overview_"] [data-testid="stVerticalBlockBorderWrapper"] {
    border: none !important;
    box-shadow: none !important;
    background: transparent !important;
}

.dashboard-card .card-label { color: #FF417D; font-weight: 700; font-size: 1rem; }
.dashboard-card .card-value { color: #152238; font-size: 1.9rem; font-weight: 750; margin-top: .4rem; }
.dashboard-card .card-help { color: #66748a; font-size: .82rem; margin-top: .2rem; }
.dashboard-section { color: #142138; font-size: 1.18rem; font-weight: 700; margin: .6rem 0 .4rem; }
[data-testid="stVerticalBlockBorderWrapper"] > div { border-radius: 11px; }

/* Unified white card design for every non-dashboard page and individual tool. */
[data-testid="stMain"] div[class*="st-key-page_"] {
    background: #ffffff !important;
    border: 0 !important;
    border-radius: 14px !important;
    box-shadow: 0 3px 16px rgba(25,43,65,.065) !important;
    padding: 24px 28px !important;
    margin-bottom: 20px !important;
}
/* Nested forms, browse results, tables and expanders use the same quiet card language. */
[data-testid="stMain"] [data-testid="stForm"] {
    background: #ffffff !important;
    border: 1px solid #edf1f6 !important;
    border-radius: 12px !important;
    box-shadow: 0 2px 9px rgba(25,43,65,.035) !important;
    padding: 18px 20px !important;
}
[data-testid="stMain"] [data-testid="stVerticalBlockBorderWrapper"]:has(> [data-testid="stVerticalBlock"]),
[data-testid="stMain"] [data-testid="stExpander"] {
    border-color: #edf1f6 !important;
    border-radius: 12px !important;
}
[data-testid="stMain"] [data-testid="stDataFrame"],
[data-testid="stMain"] [data-testid="stPlotlyChart"] {
    border-radius: 11px !important;
    overflow: hidden;
}
[data-testid="stMain"] h2, [data-testid="stMain"] h3 {
    color: #152238 !important;
    font-family: "Montserrat", sans-serif !important;
    font-weight: 650 !important;
}
@media (max-width: 768px) {
    [data-testid="stMain"] div[class*="st-key-page_"] { padding: 16px !important; }
}
</style>""", unsafe_allow_html=True)
wd=pd.to_datetime(tools.warranty_expiration,errors="coerce")

# Sidebar navigation mirrors the visual dashboard preview while retaining the
# existing application screens and data-editing logic.
MENU = [
    ("Overview", "Dashboard", ":material/home:"),
    ("Browse Tools", "Browse tools", ":material/handyman:"),
    ("Add New Tool", "Add tool", ":material/add:"),
    ("Accessories", "Batteries & accessories", ":material/inventory_2:"),
    ("Maintenance", "Maintenance & warranty", ":material/calendar_month:"),
    ("Warranties", "Maintenance & warranty", ":material/verified_user:"),
    ("Batteries", "Batteries & accessories", ":material/battery_full:"),
    ("Locations", "Browse tools", ":material/location_on:"),
    ("QR Labels", "QR labels", ":material/qr_code_2:"),
    ("Reports", "Reports / backup", ":material/bar_chart:"),
    ("Backup / Restore", "Reports / backup", ":material/cloud_download:"),
    ("Settings", "Dropdown settings", ":material/settings:"),
    ("Edit / Delete", "Edit / delete", ":material/edit_note:"),
]
# Keep the same session-state key so tool-page back navigation continues to work.
MENU_NAMES = [name for name, _, _ in MENU]
if "sidebar_menu" not in st.session_state or st.session_state.sidebar_menu not in MENU_NAMES:
    st.session_state.sidebar_menu = "Overview"
with st.sidebar:
    for menu_name, _, menu_icon in MENU:
        if st.button(menu_name, key=f"sidebar_nav_{menu_name}", icon=menu_icon, use_container_width=True):
            st.session_state.sidebar_menu = menu_name
            st.rerun()
menu_choice = st.session_state.sidebar_menu
active_section = {name: section for name, section, _ in MENU}[menu_choice]

def dashboard_card(label,value,help_text,icon):
    import html
    st.markdown(f'<div class="dashboard-card"><div class="card-label">{icon} &nbsp; {html.escape(label)}</div>'
                f'<div class="card-value">{html.escape(str(value))}</div>'
                f'<div class="card-help">{html.escape(help_text)}</div></div>', unsafe_allow_html=True)

def chart_style(fig, height=340):
    fig.update_layout(height=height, margin=dict(l=12,r=12,t=10,b=12),
                      paper_bgcolor="white",plot_bgcolor="white",
                      font=dict(color="#25364d",size=12),
                      legend=dict(orientation="v",yanchor="middle",y=.5,xanchor="left",x=1.01))
    return fig

def render_tool_editor(tid):
    """Edit one uniquely identified tool, including all existing fields."""
    global tools, maint
    matching=tools[tools.tool_id.astype(str)==str(tid)]
    if matching.empty:
        st.warning("Tool not found. It may have been deleted.")
        return
    r=matching.iloc[0].to_dict()
    st.subheader(f"{r['make']} {r['model']}")
    st.caption(f"Tool ID: {tid}")
    showphoto(r["photo_file"],200)
    # Initialize each tool's calendar state once, before the date widgets exist.
    # Date reset actions sit together at the right edge, immediately above the edit form.
    purchase_key=f"purchase_edit_{tid}"
    warranty_key=f"warranty_edit_{tid}"
    if purchase_key not in st.session_state:
        st.session_state[purchase_key]=date_value(r["purchase_date"])
    if warranty_key not in st.session_state:
        st.session_state[warranty_key]=date_value(r["warranty_expiration"])
    _, clear_purchase_col, clear_warranty_col = st.columns([5, 1.45, 1.4], gap="small")
    with clear_purchase_col:
        if st.button("Clear purchase date", key=f"clear_edit_purchase_{tid}", use_container_width=True):
            st.session_state[purchase_key]=None
            st.rerun()
    with clear_warranty_col:
        if st.button("Clear warranty date", key=f"clear_edit_warranty_{tid}", use_container_width=True):
            st.session_state[warranty_key]=None
            st.rerun()
    with st.form(f"edittool_{tid}"):
        a,b,c=st.columns(3)
        with a:
            makes=dropdown_options("Make / Brand",tools.make.tolist()); cats=dropdown_options("Category",tools.category.tolist()); volts=dropdown_options("Voltage",tools.voltage.tolist())
            make=st.selectbox("Make / Brand *",makes,index=makes.index(r["make"]) if r["make"] in makes else 0)
            model=st.text_input("Model *",value=r["model"]); category=st.selectbox("Category",cats,index=cats.index(r["category"]) if r["category"] in cats else 0)
            power=st.selectbox("Power source",POWER,index=POWER.index(r["power_source"]) if r["power_source"] in POWER else 0); voltage=st.selectbox("Voltage",volts,index=volts.index(r["voltage"]) if r["voltage"] in volts else 0)
            platforms=dropdown_options("Battery platform",tools.battery_platform.tolist()); platform=st.selectbox("Battery platform",platforms,index=platforms.index(r["battery_platform"]) if r["battery_platform"] in platforms else 0)
        with b:
            serial=st.text_input("Serial number",value=r["serial_number"]); condition=st.selectbox("Condition",CONDITIONS,index=CONDITIONS.index(r["condition"]) if r["condition"] in CONDITIONS else 2)
            price_raw=st.text_input("Purchase price",value="" if pd.isna(r["purchase_price"]) else str(r["purchase_price"]))
            vendors=dropdown_options("Vendor",tools.vendor.tolist()); vendor=st.selectbox("Vendor",vendors,index=vendors.index(r["vendor"]) if r["vendor"] in vendors else 0)
            pdate=st.date_input("Purchase date",value=None,format="YYYY-MM-DD",key=purchase_key)
            st.caption(f"Purchase date: {date_text(pdate) or 'Not set'}")
            warranty=st.date_input("Warranty expiration",value=None,format="YYYY-MM-DD",key=warranty_key)
            st.caption(f"Warranty expiration: {date_text(warranty) or 'Not set'}")
        with c:
            area=st.text_input("Storage area / room",value=r["location_area"]); cabinet=st.text_input("Cabinet / rack",value=r["location_cabinet"]); shelf=st.text_input("Shelf",value=r["location_shelf"]); bin_name=st.text_input("Drawer / bin",value=r["location_bin"])
            otherloc=st.text_input("Other location details",value=r["storage_location"]); notes=st.text_area("Notes",value=r["notes"]); upload=st.file_uploader("Add/replace photo",type=["jpg","jpeg","png","webp"],key=f"editphoto_{tid}"); remove=st.checkbox("Remove photo")
        submit=st.form_submit_button("Save changes",type="primary")
    if submit:
        err="Make / Brand and Model are required." if not make.strip() or not model.strip() else ""
        amount,pe=money(price_raw,"Purchase price")
        if not err: err=pe
        if err: st.error(err)
        else:
            pfile=r["photo_file"]
            try:
                if upload: pfile=savephoto(upload,tid)
                elif remove and pfile:
                    p=ROOT/pfile
                    if p.exists(): p.unlink()
                    pfile=""
            except ValueError as e: st.error(str(e)); st.stop()
            storage=loc(area,cabinet,shelf,bin_name,otherloc.strip())
            vals={"make":make.strip(),"model":model.strip(),"category":category.strip(),"power_source":power,"voltage":voltage.strip(),"battery_platform":platform.strip(),"serial_number":serial.strip(),"purchase_date":date_text(pdate),"purchase_price":amount,"vendor":vendor.strip(),"condition":condition,"storage_location":storage,"location_area":area.strip(),"location_cabinet":cabinet.strip(),"location_shelf":shelf.strip(),"location_bin":bin_name.strip(),"warranty_expiration":date_text(warranty),"photo_file":pfile,"notes":notes.strip()}
            idx=tools.index[tools.tool_id==tid][0]
            for k,v in vals.items(): tools.at[idx,k]=v
            save({"Tools":tools,"Maintenance":maint,"Accessories":acc})
            st.session_state.pop(purchase_key, None)
            st.session_state.pop(warranty_key, None)
            st.success("Updated")
            st.rerun()
    with st.expander("Danger zone"):
        confirm=st.checkbox("Delete this tool and its maintenance history",key=f"deltool_{tid}")
        if st.button("Delete tool",disabled=not confirm):
            p=ROOT/r["photo_file"]
            if r["photo_file"] and p.exists(): p.unlink()
            tools=tools[tools.tool_id!=tid].reset_index(drop=True); maint=maint[maint.tool_id!=tid].reset_index(drop=True)
            save({"Tools":tools,"Maintenance":maint,"Accessories":acc}); st.rerun()

selected_tool_id=st.query_params.get("tool")
if selected_tool_id:
    if st.button("← Back to tool browser"):
        st.session_state.sidebar_menu="Browse Tools"
        del st.query_params["tool"]
        st.rerun()
    with st.container(key="page_individual_tool"):
        render_tool_editor(selected_tool_id)
    st.stop()


if active_section == 'Dashboard':
    st.markdown("# Overview")
    st.caption("Summary of your tool inventory, value, and key information.")
    expired_count=int((wd<today).sum())
    cards=st.columns(5,gap="small")
    with cards[0]: dashboard_card("Tools",len(tools),"Total tools","🛠")
    with cards[1]: dashboard_card("Accessories",len(acc),"Accessory records","⬡")
    with cards[2]: dashboard_card("Purchase Value",f"${purchase:,.2f}","Total purchase price","＄")
    with cards[3]: dashboard_card("Service Costs",f"${service:,.2f}","Total maintenance","▥")
    with cards[4]: dashboard_card("Expired Warranties",expired_count,"Tools out of warranty","◇")
    st.write("")
    left,right=st.columns(2,gap="medium")
    with left:
        with st.container(border=True, key="overview_category"):
            st.markdown('<div class="dashboard-section">Tools by Category</div>',unsafe_allow_html=True)
            if tools.empty:
                st.info("Add tools to see category distribution.")
            else:
                categories=tools.category.fillna("").replace("","Uncategorized").value_counts().reset_index()
                categories.columns=["Category","Count"]
                fig=px.pie(categories,names="Category",values="Count",hole=0,
                           color_discrete_sequence=["#1886cf","#ff871b","#24b04b","#ef333c","#9263c9","#8b5848","#ed6cbd","#19a5a5"])
                fig.update_traces(textposition="inside",textinfo="percent",hovertemplate="%{label}: %{value}<extra></extra>")
                st.plotly_chart(chart_style(fig),use_container_width=True,config={"displayModeBar":False})
    with right:
        with st.container(border=True, key="overview_brand"):
            st.markdown('<div class="dashboard-section">Tools by Brand</div>',unsafe_allow_html=True)
            if tools.empty:
                st.info("Add tools to see brand distribution.")
            else:
                brands=tools.make.fillna("").replace("","Unspecified").value_counts().head(10).sort_values().reset_index()
                brands.columns=["Brand","Count"]
                fig=px.bar(brands,x="Count",y="Brand",orientation="h",text="Count",
                           color="Brand",color_discrete_sequence=["#1c7bd0","#ff881a","#27bd54","#f44748","#9c5cd2","#00a18c","#e8b122","#57b4e5","#ed6cbd","#757f90"])
                fig.update_traces(textposition="outside",cliponaxis=False)
                fig.update_layout(showlegend=False,xaxis=dict(showgrid=True,gridcolor="#edf0f4",title=None),
                                  yaxis=dict(title=None),bargap=.24)
                st.plotly_chart(chart_style(fig),use_container_width=True,config={"displayModeBar":False})
    with st.container(border=True, key="overview_recent"):
        st.markdown('<div class="dashboard-section">Recent Tools</div>',unsafe_allow_html=True)
        if tools.empty:
            st.info("No tools added yet.")
        else:
            recent=tools.iloc[::-1].head(8).copy()
            recent["Tool Name"]=(recent.make.fillna("")+" "+recent.model.fillna("")).str.strip()
            recent=recent.rename(columns={"make":"Brand","category":"Category","purchase_date":"Purchase Date","vendor":"Vendor"})
            recent["Tool Price"]=pd.to_numeric(recent["purchase_price"],errors="coerce")
            st.dataframe(recent[["Tool Name","Brand","Category","Purchase Date","Tool Price","Vendor"]],
                         column_config={"Tool Price": st.column_config.NumberColumn("Tool Price",format="$%.2f")},
                         use_container_width=True,hide_index=True)
    if not tools.empty:
        x,y=st.columns(2,gap="medium")
        with x:
            with st.container(border=True, key="overview_condition"):
                st.markdown('<div class="dashboard-section">Tools by Condition</div>',unsafe_allow_html=True)
                conditions=tools.condition.fillna("").replace("","Unspecified").value_counts().reset_index()
                conditions.columns=["Condition","Count"]
                fig=px.bar(conditions,x="Condition",y="Count",color_discrete_sequence=["#FF417D"])
                fig.update_layout(xaxis_title=None,yaxis_title=None,showlegend=False)
                st.plotly_chart(chart_style(fig,290),use_container_width=True,config={"displayModeBar":False})
        with y:
            with st.container(border=True, key="overview_value_category"):
                st.markdown('<div class="dashboard-section">Purchase Value by Category</div>',unsafe_allow_html=True)
                pv=tools.copy()
                pv["purchase_price"]=pd.to_numeric(pv.purchase_price,errors="coerce").fillna(0)
                pv["category"]=pv.category.fillna("").replace("","Uncategorized")
                amounts=pv.groupby("category",as_index=False).purchase_price.sum().sort_values("purchase_price",ascending=False).head(8)
                fig=px.bar(amounts,x="category",y="purchase_price",color_discrete_sequence=["#1886cf"])
                fig.update_layout(xaxis_title=None,yaxis_title=None,showlegend=False)
                st.plotly_chart(chart_style(fig,290),use_container_width=True,config={"displayModeBar":False})
        warr=tools.copy()
        warr["Days to expiry"]=(pd.to_datetime(warr.warranty_expiration,errors="coerce")-today).dt.days
        alerts=warr[warr["Days to expiry"].notna()&(warr["Days to expiry"]<=60)]
        with st.expander(f"Expired or soon-expiring warranties ({len(alerts)})"):
            if alerts.empty: st.success("No warranties expiring within 60 days.")
            else: st.dataframe(alerts[["tool_id","make","model","warranty_expiration","Days to expiry"]],use_container_width=True,hide_index=True)
    if not maint.empty:
        md=maint.copy()
        md["cost"]=pd.to_numeric(md.cost,errors="coerce").fillna(0)
        md["month"]=pd.to_datetime(md.service_date,errors="coerce").dt.to_period("M").astype(str)
        monthly=md[md.month!="NaT"].groupby("month",as_index=False).cost.sum()
        if not monthly.empty:
            with st.container(border=True, key="overview_maintenance"):
                st.markdown('<div class="dashboard-section">Monthly Maintenance Spend</div>',unsafe_allow_html=True)
                fig=px.line(monthly,x="month",y="cost",markers=True,color_discrete_sequence=["#FF417D"])
                fig.update_layout(xaxis_title=None,yaxis_title=None)
                st.plotly_chart(chart_style(fig,290),use_container_width=True,config={"displayModeBar":False})
if active_section == 'Browse tools':
    with st.container(key="page_browse_tools"):
        st.subheader("Browse / search tools")
        q=st.text_input("Search make, model, serial, platform or location")
        a,b,c=st.columns(3)
        with a: cats=["All"]+sorted(x for x in tools.category.unique() if x); cat=st.selectbox("Category",cats)
        with b: conds=["All"]+sorted(x for x in tools.condition.unique() if x); cond=st.selectbox("Condition",conds)
        with c: powers=["All"]+sorted(x for x in tools.power_source.unique() if x); power=st.selectbox("Power source",powers)
        view=tools.copy()
        if q: view=view[view.astype(str).apply(lambda col:col.str.contains(q,case=False,na=False)).any(axis=1)]
        if cat!="All": view=view[view.category==cat]
        if cond!="All": view=view[view.condition==cond]
        if power!="All": view=view[view.power_source==power]
        view=view.sort_values(by=["make", "model", "tool_id"], key=lambda col: col.fillna("").astype(str).str.casefold(), kind="stable")
        if view.empty: st.info("No matching tools.")
        for _,r in view.iterrows():
            with st.container(border=True):
                a,b=st.columns([1,4])
                with a: showphoto(r.photo_file)
                with b:
                    if st.button(f"{r.make} {r.model}", key=f"browse_tool_{r.tool_id}", type="tertiary", help="Open this tool’s individual edit page"):
                        st.query_params["tool"]=str(r.tool_id)
                        st.rerun()
                    st.caption(f"Tool ID: {r.tool_id}")
                    st.write(f"Category: {r.category or '—'} | Power: {r.power_source or '—'} | Voltage: {r.voltage or '—'}")
                    st.write(f"Battery platform: {r.battery_platform or '—'} | Condition: {r.condition or '—'}")
                    st.write(f"Location: {location(r) or '—'} | Serial: {r.serial_number or '—'}")
                    st.write(f"Purchase: {r.purchase_date or '—'} | Price: {r.purchase_price or '—'} | Warranty: {r.warranty_expiration or '—'}")
                    if r.notes: st.caption(r.notes)
        st.caption(f"{len(view)} of {len(tools)} tools")
if active_section == 'Add tool':
    with st.container(key="page_add_tool"):
        st.subheader("Add a tool")
        # Compact right-aligned date actions immediately above the tool details form.
        _, clear_purchase_col, clear_warranty_col = st.columns([5, 1.45, 1.4], gap="small")
        with clear_purchase_col:
            if st.button("Clear purchase date", key="clear_add_purchase_date", use_container_width=True):
                st.session_state["add_tool_purchase_date"] = None
                st.rerun()
        with clear_warranty_col:
            if st.button("Clear warranty date", key="clear_add_warranty_date", use_container_width=True):
                st.session_state["warranty_add"] = None
                st.rerun()
        with st.form("addtool",clear_on_submit=True):
            a,b,c=st.columns(3)
            with a:
                make=st.selectbox("Make / Brand *",dropdown_options("Make / Brand",tools.make.tolist()))
                model=st.text_input("Model *"); category=st.selectbox("Category",dropdown_options("Category",tools.category.tolist()))
                power=st.selectbox("Power source",POWER); voltage=st.selectbox("Voltage",dropdown_options("Voltage",tools.voltage.tolist())); platform=st.selectbox("Battery platform",dropdown_options("Battery platform",tools.battery_platform.tolist()))
            with b:
                serial=st.text_input("Serial number"); condition=st.selectbox("Condition",CONDITIONS,index=CONDITIONS.index("New"))
                price_raw=st.text_input("Purchase price")
                vendor=st.selectbox("Vendor",dropdown_options("Vendor",tools.vendor.tolist()))
                pdate=st.date_input("Purchase date",value=None,format="YYYY-MM-DD",key="add_tool_purchase_date")
                st.caption(f"Purchase date: {date_text(pdate) or 'Not set'}")
                warranty=st.date_input("Warranty expiration",value=None,format="YYYY-MM-DD",key="warranty_add")
                st.caption(f"Warranty expiration: {date_text(warranty) or 'Not set'}")
            with c:
                area=st.text_input("Storage area / room"); cabinet=st.text_input("Cabinet / rack"); shelf=st.text_input("Shelf"); bin_name=st.text_input("Drawer / bin")
                otherloc=st.text_input("Other location details"); notes=st.text_area("Notes"); upload=st.file_uploader("Tool photo",type=["jpg","jpeg","png","webp"])
            submit=st.form_submit_button("Save tool",type="primary")
        if submit:
            err="Make / Brand and Model are required." if not make.strip() or not model.strip() else ""
            amount,pe=money(price_raw,"Purchase price")
            if not err: err=pe
            if err: st.error(err)
            else:
                tid=uid()
                try: pfile=savephoto(upload,tid)
                except ValueError as e: st.error(str(e)); st.stop()
                storage=loc(area,cabinet,shelf,bin_name,otherloc.strip())
                row=dict(zip(SHEETS["Tools"],[tid,make.strip(),model.strip(),category.strip(),power,voltage.strip(),platform.strip(),serial.strip(),date_text(pdate),amount,vendor.strip(),condition,storage,area.strip(),cabinet.strip(),shelf.strip(),bin_name.strip(),date_text(warranty),pfile,notes.strip()]))
                tools=pd.concat([tools,pd.DataFrame([row])],ignore_index=True); save({"Tools":tools,"Maintenance":maint,"Accessories":acc}); st.success(f"Saved tool ID {tid}"); st.rerun()
if active_section == 'Edit / delete':
    with st.container(key="page_edit_delete"):
        st.subheader("Edit or delete")
        if tools.empty:
            st.info("Add a tool first.")
        else:
            sorted_tools=tools.sort_values(by=["make", "model", "tool_id"], key=lambda col: col.fillna("").astype(str).str.casefold(), kind="stable")
            choices=sorted_tools.apply(toolname,axis=1).tolist()
            chosen=st.selectbox("Tool",choices)
            render_tool_editor(chosen.split(" — ",1)[0])
if active_section == 'Maintenance & warranty':
    with st.container(key="page_maintenance_warranty"):
        st.subheader("Maintenance & warranty")
        if tools.empty: st.info("Add tools before logging service.")
        else:
            with st.form("service",clear_on_submit=True):
                chosen=st.selectbox("Tool",tools.apply(toolname,axis=1).tolist()); sd=st.text_input("Service date YYYY-MM-DD",value=date.today().isoformat())
                typ=st.selectbox("Service type",SERVICES); desc=st.text_input("Work performed"); costraw=st.text_input("Service cost")
                nextd=st.text_input("Next service due YYYY-MM-DD"); by=st.text_input("Performed by"); notes=st.text_area("Notes")
                submit=st.form_submit_button("Add maintenance record",type="primary")
            if submit:
                err=valdate(sd.strip(),"Service date",True)
                if not err: err=valdate(nextd.strip(),"Next service date")
                amount,pe=money(costraw,"Service cost")
                if not err: err=pe
                if err: st.error(err)
                else:
                    tid=chosen.split(" — ",1)[0]; row=dict(zip(SHEETS["Maintenance"],[uid(),tid,sd.strip(),typ,desc.strip(),amount,nextd.strip(),by.strip(),notes.strip()]))
                    maint=pd.concat([maint,pd.DataFrame([row])],ignore_index=True); save({"Tools":tools,"Maintenance":maint,"Accessories":acc}); st.success("Service record saved"); st.rerun()
        st.markdown("**Warranty dates**")
        if not tools.empty:
            w=tools[tools.warranty_expiration.astype(str).str.strip()!=""].copy()
            if not w.empty:
                w["Days remaining"]=(pd.to_datetime(w.warranty_expiration,errors="coerce")-today).dt.days
                st.dataframe(w[["tool_id","make","model","warranty_expiration","Days remaining"]],use_container_width=True,hide_index=True)
        st.markdown("**Maintenance due / due within 30 days**")
        if not maint.empty:
            d=maint[maint.next_service_date.astype(str).str.strip()!=""].copy()
            if not d.empty:
                d["parsed"]=pd.to_datetime(d.next_service_date,errors="coerce"); d=d[d.parsed<=today+pd.Timedelta(days=30)]
                if not d.empty: st.dataframe(d.merge(tools[["tool_id","make","model"]],on="tool_id",how="left"),use_container_width=True,hide_index=True)
        st.markdown("**Service history**"); st.dataframe(maint.merge(tools[["tool_id","make","model"]],on="tool_id",how="left") if not maint.empty else maint,use_container_width=True,hide_index=True)
if active_section == 'Batteries & accessories':
    with st.container(key="page_batteries_accessories"):
        st.subheader("Batteries & accessories")
        with st.form("addacc",clear_on_submit=True):
            a,b,c=st.columns(3)
            with a: name=st.text_input("Name *"); typ=st.selectbox("Type",TYPES); make=st.selectbox("Make / Brand",dropdown_options("Make / Brand",acc.make.tolist()))
            with b: model=st.text_input("Model"); platform=st.selectbox("Battery platform",dropdown_options("Battery platform",acc.battery_platform.tolist())); qty=st.number_input("Quantity",min_value=0,value=1,step=1)
            with c:
                compatible=st.text_input("Compatible tool IDs / models"); area=st.text_input("Storage area / room"); cabinet=st.text_input("Cabinet / rack")
                shelf=st.text_input("Shelf"); bin_name=st.text_input("Drawer / bin"); condition=st.selectbox("Condition",CONDITIONS,index=2)
            notes=st.text_area("Notes"); submit=st.form_submit_button("Add accessory",type="primary")
        if submit:
            if not name.strip(): st.error("Name is required.")
            else:
                storage=loc(area,cabinet,shelf,bin_name); row=dict(zip(SHEETS["Accessories"],[uid(),name.strip(),typ,make.strip(),model.strip(),platform.strip(),int(qty),compatible.strip(),storage,area.strip(),cabinet.strip(),shelf.strip(),bin_name.strip(),condition,notes.strip()]))
                acc=pd.concat([acc,pd.DataFrame([row])],ignore_index=True); save({"Tools":tools,"Maintenance":maint,"Accessories":acc}); st.success("Accessory saved"); st.rerun()
        q=st.text_input("Search accessories")
        av=acc.copy()
        if q: av=av[av.astype(str).apply(lambda col:col.str.contains(q,case=False,na=False)).any(axis=1)]
        st.dataframe(av,use_container_width=True,hide_index=True)
        if not acc.empty:
            with st.expander("Edit/delete accessory"):
                choices=acc.apply(lambda r:f"{r.accessory_id} — {r.name}",axis=1).tolist(); ch=st.selectbox("Accessory",choices); aid=ch.split(" — ",1)[0]; r=acc[acc.accessory_id==aid].iloc[0].to_dict()
                with st.form("editacc"):
                    name=st.text_input("Name *",value=r["name"]); typ=st.selectbox("Type",TYPES,index=TYPES.index(r["type"]) if r["type"] in TYPES else 0)
                    makes=dropdown_options("Make / Brand",acc.make.tolist()); make=st.selectbox("Make / Brand",makes,index=makes.index(r["make"]) if r["make"] in makes else 0); model=st.text_input("Model",value=r["model"]); platforms=dropdown_options("Battery platform",acc.battery_platform.tolist()); platform=st.selectbox("Battery platform",platforms,index=platforms.index(r["battery_platform"]) if r["battery_platform"] in platforms else 0)
                    qty=st.number_input("Quantity",min_value=0,value=int(float(r["quantity"] or 0)),step=1); compatible=st.text_input("Compatible tools",value=r["compatible_tools"])
                    area=st.text_input("Storage area",value=r["location_area"]); cabinet=st.text_input("Cabinet",value=r["location_cabinet"]); shelf=st.text_input("Shelf",value=r["location_shelf"]); bin_name=st.text_input("Drawer/bin",value=r["location_bin"])
                    condition=st.selectbox("Condition",CONDITIONS,index=CONDITIONS.index(r["condition"]) if r["condition"] in CONDITIONS else 2); notes=st.text_area("Notes",value=r["notes"])
                    submit=st.form_submit_button("Save accessory")
                if submit:
                    if not name.strip(): st.error("Name is required.")
                    else:
                        storage=loc(area,cabinet,shelf,bin_name); vals={"name":name.strip(),"type":typ,"make":make.strip(),"model":model.strip(),"battery_platform":platform.strip(),"quantity":int(qty),"compatible_tools":compatible.strip(),"storage_location":storage,"location_area":area.strip(),"location_cabinet":cabinet.strip(),"location_shelf":shelf.strip(),"location_bin":bin_name.strip(),"condition":condition,"notes":notes.strip()}
                        idx=acc.index[acc.accessory_id==aid][0]
                        for k,v in vals.items(): acc.at[idx,k]=v
                        save({"Tools":tools,"Maintenance":maint,"Accessories":acc}); st.rerun()
                if st.button("Delete accessory",key="deleteacc"):
                    acc=acc[acc.accessory_id!=aid].reset_index(drop=True); save({"Tools":tools,"Maintenance":maint,"Accessories":acc}); st.rerun()
if active_section == 'QR labels':
    with st.container(key="page_qr_labels"):
        st.subheader("Printable QR labels — Avery 5160 / 8160")
        st.caption("30 labels per US Letter sheet (3 columns × 10 rows; 2⅝ × 1 inch). Each label has a centered inset border around its QR code and details. Print at Actual size / 100%, with no Fit to page scaling.")
        st.write("QR codes encode the inventory Tool ID. Scan the code, then search that ID in Browse tools; the QR does not open a remote website.")
        if tools.empty: st.info("Add tools first.")
        else:
            ids=st.multiselect("Select tool IDs",tools.tool_id.tolist(),default=tools.tool_id.tolist())
            if st.button("Generate printable QR label PDF",type="primary"):
                if not ids: st.warning("Select at least one tool.")
                else:
                    # Avery 5160 / 8160: US Letter, 3 columns x 10 rows,
                    # 2-5/8 x 1 inch labels, 3/16 inch left margin, 1/2 inch top.
                    path=ROOT/"tool_qr_labels.pdf"
                    pdf=canvas.Canvas(str(path), pagesize=letter)
                    pw, ph=letter
                    label_w, label_h=2.625*72, 1.0*72
                    left_margin, top_margin=0.1875*72, 0.5*72
                    col_pitch=2.75*72  # 1/8 inch gap between columns
                    row_pitch=1.0*72   # No vertical gap
                    for n, tid in enumerate(ids):
                        if n and n % 30 == 0:
                            pdf.showPage()
                        p=n % 30
                        col, row=p % 3, p // 3
                        x=left_margin+col*col_pitch
                        y=ph-top_margin-(row+1)*row_pitch
                        r=tools[tools.tool_id==tid].iloc[0]
                        qr=qrcode.make("TOOL-ID:"+tid)
                        qp=QR/(tid+".png")
                        qr.save(qp)
                        # Center an inset border and all contents within each Avery label.
                        # Inset keeps the border clear of the sheet's die-cut edges.
                        inset_x, inset_y = 5, 5
                        box_x, box_y = x + inset_x, y + inset_y
                        box_w, box_h = label_w - 2 * inset_x, label_h - 2 * inset_y
                        pdf.setStrokeColorRGB(0.20, 0.20, 0.20)
                        pdf.setLineWidth(0.65)
                        pdf.roundRect(box_x, box_y, box_w, box_h, 3, stroke=1, fill=0)
                        qr_size = 45
                        qr_x = box_x + 6
                        qr_y = box_y + (box_h - qr_size) / 2
                        pdf.drawImage(ImageReader(str(qp)), qr_x, qr_y,
                                      width=qr_size, height=qr_size, mask="auto")
                        text_x = qr_x + qr_size + 8
                        text_right = box_x + box_w - 6
                        max_text_w = max(10, text_right - text_x)
                        def fit_label_text(value, font, size):
                            value = str(value or "")
                            if pdf.stringWidth(value, font, size) <= max_text_w:
                                return value
                            while value and pdf.stringWidth(value + "…", font, size) > max_text_w:
                                value = value[:-1]
                            return value + "…" if value else ""
                        pdf.setFillColorRGB(0, 0, 0)
                        pdf.setFont("Helvetica-Bold", 7.5)
                        pdf.drawString(text_x, box_y + 42, fit_label_text(f"{r['make']} {r['model']}", "Helvetica-Bold", 7.5))
                        pdf.setFont("Helvetica", 7)
                        pdf.drawString(text_x, box_y + 28, fit_label_text("ID: " + tid, "Helvetica", 7))
                        pdf.drawString(text_x, box_y + 14, fit_label_text(location(r) or "Location not set", "Helvetica", 7))
                    pdf.save(); st.session_state["qr_ready"]=True; st.success("PDF generated.")
            if st.session_state.get("qr_ready") and (ROOT/"tool_qr_labels.pdf").exists():
                with open(ROOT/"tool_qr_labels.pdf","rb") as f: st.download_button("Download QR label PDF",f.read(),file_name="tool_qr_labels.pdf",mime="application/pdf")
if active_section == 'Reports / backup':
    with st.container(key="page_reports_backup"):
        st.subheader("Reports & backup")
        st.write("Download a ZIP snapshot of this app, its current Excel data, photos, and QR labels. Extract the ZIP and upload its contents to your GitHub repository to restore your deployment.")
        try:
            backup_data=complete_backup_zip()
            st.download_button(
                "⬇️ Download Complete Application Backup (ZIP)",
                data=backup_data,
                file_name="tool_inventory_complete_backup_"+date.today().isoformat()+".zip",
                mime="application/zip",
                type="primary",
                help="Includes app.py, requirements, Excel workbook, photos, QR images, and any generated QR-label PDF.",
            )
        except OSError as exc:
            st.error(f"Could not prepare the full backup: {exc}")
        st.caption("Back up regularly. Streamlit Cloud's local files are temporary; missing files cannot be recovered from a new backup.")
        with open(FILE,"rb") as f: st.download_button("Download full Excel workbook",f.read(),file_name="tool_inventory.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        q=st.text_input("Filter tool CSV by text")
        report=tools.copy()
        if q: report=report[report.astype(str).apply(lambda col:col.str.contains(q,case=False,na=False)).any(axis=1)]
        st.download_button("Download tool report (CSV)",report.to_csv(index=False).encode("utf-8-sig"),file_name="tool_report.csv",mime="text/csv")
        if not maint.empty:
            joined=maint.merge(tools[["tool_id","make","model","category"]],on="tool_id",how="left")
            st.dataframe(joined,use_container_width=True,hide_index=True)
            st.download_button("Download maintenance report (CSV)",joined.to_csv(index=False).encode("utf-8-sig"),file_name="maintenance_report.csv",mime="text/csv")
        if not acc.empty: st.download_button("Download accessories report (CSV)",acc.to_csv(index=False).encode("utf-8-sig"),file_name="accessories_report.csv",mime="text/csv")
        st.info("The ZIP backup includes app files, tools.xlsx, photos, QR images, and any generated label PDF. Keep a copy outside Streamlit Cloud.")

if active_section == 'Dropdown settings':
    with st.container(key="page_dropdown_settings"):
        st.subheader("Manage dropdown selections")
        st.write("Edit the backend lists used by the Make / Brand, Category, Voltage, Battery platform, and Vendor dropdowns. Enter one option per line. Existing inventory values remain selectable even if they are not in these lists.")
        def options_text(kind):
            if dropdowns.empty: return "\n".join(OPTION_DEFAULTS[kind])
            vals=dropdowns.loc[dropdowns.option_type==kind,"option_value"].astype(str).tolist()
            return "\n".join(vals if vals else OPTION_DEFAULTS[kind])
        with st.form("dropdown_settings"):
            make_text=st.text_area("Make / Brand options",value=options_text("Make / Brand"),height=180,help="Example: DeWalt, Milwaukee, Makita — one per line.")
            category_text=st.text_area("Category options",value=options_text("Category"),height=180)
            voltage_text=st.text_area("Voltage options",value=options_text("Voltage"),height=140,help="Example: 12V, 18V, 20V, Corded / N/A.")
            platform_text=st.text_area("Battery platform options",value=options_text("Battery platform"),height=160,help="Example: DeWalt 20V MAX, Milwaukee M18, Makita LXT 18V.")
            vendor_text=st.text_area("Vendor options",value=options_text("Vendor"),height=160,help="One vendor per line, such as Home Depot, Lowe’s, or Amazon.")
            save_options=st.form_submit_button("Save dropdown options",type="primary")
        if save_options:
            rows=[]
            for kind,raw in [("Make / Brand",make_text),("Category",category_text),("Voltage",voltage_text),("Battery platform",platform_text),("Vendor",vendor_text)]:
                values=list(dict.fromkeys(line.strip() for line in raw.splitlines() if line.strip()))
                if not values:
                    st.error(f"Keep at least one option for {kind}.")
                    st.stop()
                rows.extend({"option_type":kind,"option_value":value} for value in values)
            dropdowns=pd.DataFrame(rows,columns=SHEETS["Dropdowns"])
            save({"Tools":tools,"Maintenance":maint,"Accessories":acc,"Dropdowns":dropdowns})
            st.success("Dropdown options saved to the Excel backend (Dropdowns worksheet).")
            st.rerun()
        st.caption("Backend storage: tools.xlsx → Dropdowns worksheet. You can also edit that worksheet directly, keeping the option_type and option_value columns intact.")


# Consistent footer across all navigation pages.
st.markdown(
    """
    <style>
      .tool-station-footer {
        margin-top: 3rem; padding: 1.25rem 0 0.6rem;
        border-top: 1px solid #e7ebf2; text-align: center;
        color: #718096; font-family: 'Montserrat', sans-serif;
        font-size: 0.82rem; letter-spacing: 0.015em;
      }
      .tool-station-footer strong { color: #344256; font-weight: 600; }
      .tool-station-footer .heart { color: #FF417D; }
    </style>
    <div class="tool-station-footer">
      <span class="heart">✦</span> Created with AI by <strong>Fuad Khandaker</strong>
      <span class="heart">✦</span>
    </div>
    """,
    unsafe_allow_html=True,
)
