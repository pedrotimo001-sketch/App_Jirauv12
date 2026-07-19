from pathlib import Path
import base64

RED = "#C7192D"
RED_DARK = "#8E1020"
RED_BRIGHT = "#E12B3F"
WHITE = "#FFFFFF"
INK = "#17202A"
MUTED = "#697386"
BG = "#F5F7FA"
BORDER = "#E5E9F0"


def svg_data_uri(path: Path) -> str:
    data = path.read_bytes()
    return "data:image/svg+xml;base64," + base64.b64encode(data).decode("ascii")


def premium_css() -> str:
    return r"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
:root{
  --jirau-red:#C7192D;--jirau-red-dark:#8E1020;--jirau-red-bright:#E12B3F;
  --jirau-white:#FFFFFF;--jirau-ink:#17202A;--jirau-muted:#697386;
  --jirau-bg:#F5F7FA;--jirau-border:#E5E9F0;--jirau-shadow:0 12px 30px rgba(23,32,42,.08);
}
html,body,[class*="css"]{font-family:'Inter',sans-serif!important}
.stApp{background:
 radial-gradient(circle at 100% 0%,rgba(199,25,45,.07),transparent 28rem),
 linear-gradient(180deg,#FAFBFC 0%,var(--jirau-bg) 100%);color:var(--jirau-ink)}
.block-container{padding-top:1rem;padding-bottom:4rem;max-width:1600px}
#MainMenu,footer{visibility:hidden}
header[data-testid="stHeader"]{background:rgba(255,255,255,.72);backdrop-filter:blur(14px);border-bottom:1px solid rgba(229,233,240,.8)}
section[data-testid="stSidebar"]{background:linear-gradient(180deg,#A91125 0%,#710B19 100%);border-right:1px solid rgba(255,255,255,.12);box-shadow:12px 0 40px rgba(63,5,15,.18)}
section[data-testid="stSidebar"]>div{padding-top:1rem}
section[data-testid="stSidebar"] *{color:#fff}
section[data-testid="stSidebar"] [data-testid="stSidebarNav"]{display:none}
section[data-testid="stSidebar"] div[role="radiogroup"] label{background:rgba(255,255,255,.055);border:1px solid rgba(255,255,255,.10);border-radius:12px;padding:10px 12px;margin-bottom:7px;transition:.22s ease}
section[data-testid="stSidebar"] div[role="radiogroup"] label:hover{background:rgba(255,255,255,.13);transform:translateX(3px)}
section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked){background:#fff;border-color:#fff;box-shadow:0 8px 20px rgba(40,0,8,.2)}
section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) p{color:var(--jirau-red-dark)!important;font-weight:800}
section[data-testid="stSidebar"] [data-testid="stProgress"]>div>div{background:#fff!important}
.jirau-brand-box{padding:14px 12px 18px;text-align:center;border-bottom:1px solid rgba(255,255,255,.15);margin-bottom:14px}
.jirau-brand-box img{width:150px;max-height:72px;object-fit:contain;filter:drop-shadow(0 5px 16px rgba(0,0,0,.14))}
.jirau-brand-box .brand-sub{font-size:11px;text-transform:uppercase;letter-spacing:.14em;opacity:.82;margin-top:8px;font-weight:700}
.jirau-version{display:inline-flex;padding:5px 10px;border-radius:999px;background:rgba(255,255,255,.12);font-size:10px;letter-spacing:.08em;font-weight:800;margin-top:9px}
.jirau-header{position:relative;overflow:hidden;background:linear-gradient(115deg,#fff 0%,#fff 66%,#FFF3F5 100%);border:1px solid var(--jirau-border);border-radius:22px;padding:22px 26px;margin:2px 0 26px;box-shadow:var(--jirau-shadow)}
.jirau-header:after{content:"";position:absolute;right:-60px;top:-100px;width:280px;height:280px;border-radius:50%;background:radial-gradient(circle,rgba(199,25,45,.16),rgba(199,25,45,0) 68%)}
.jirau-header-inner{display:flex;align-items:center;gap:18px;position:relative;z-index:2}
.jirau-header-logo{width:124px;height:62px;display:flex;align-items:center;justify-content:center;padding-right:18px;border-right:1px solid var(--jirau-border)}
.jirau-header-logo img{max-width:112px;max-height:54px}
.jirau-eyebrow{font-size:11px;color:var(--jirau-red);font-weight:800;letter-spacing:.16em;text-transform:uppercase;margin-bottom:5px}
.jirau-title{font-size:24px;line-height:1.1;font-weight:800;color:var(--jirau-ink)}
.jirau-subtitle{font-size:13px;color:var(--jirau-muted);margin-top:6px}
.premium-hero{background:linear-gradient(128deg,#A91125,#C7192D 60%,#E12B3F);color:#fff;border-radius:24px;padding:30px;box-shadow:0 18px 45px rgba(142,16,32,.25);margin-bottom:24px;position:relative;overflow:hidden}
.premium-hero:after{content:"";position:absolute;right:-70px;top:-120px;width:330px;height:330px;border:1px solid rgba(255,255,255,.2);border-radius:50%;box-shadow:0 0 0 40px rgba(255,255,255,.04),0 0 0 80px rgba(255,255,255,.03)}
.premium-hero h1{color:#fff!important;margin:0 0 8px;font-size:30px}.premium-hero p{margin:0;opacity:.9;max-width:760px}
.card,.premium-card{border:1px solid var(--jirau-border);border-radius:18px;padding:20px;background:rgba(255,255,255,.96);box-shadow:0 8px 22px rgba(23,32,42,.055);min-height:140px;transition:.25s ease;position:relative;overflow:hidden}
.card:before,.premium-card:before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:linear-gradient(var(--jirau-red),var(--jirau-red-bright))}
.card:hover,.premium-card:hover{transform:translateY(-3px);box-shadow:0 14px 32px rgba(23,32,42,.10)}
div[data-testid="stMetric"]{background:#fff;border:1px solid var(--jirau-border);border-radius:17px;padding:18px;box-shadow:0 7px 20px rgba(23,32,42,.055);transition:.22s ease}
div[data-testid="stMetric"]:hover{transform:translateY(-2px);box-shadow:0 12px 26px rgba(23,32,42,.09)}
div[data-testid="stMetric"] label{color:var(--jirau-muted)!important;font-weight:700!important}
div[data-testid="stMetric"] [data-testid="stMetricValue"]{color:var(--jirau-ink);font-weight:800}
div[data-testid="stForm"]{background:#fff;border:1px solid var(--jirau-border);border-radius:18px;padding:20px;box-shadow:0 8px 22px rgba(23,32,42,.055)}
.stButton>button,.stDownloadButton>button,button[kind="primary"]{border-radius:11px!important;min-height:43px;font-weight:800!important;border:1px solid var(--jirau-red)!important;background:linear-gradient(135deg,var(--jirau-red-dark),var(--jirau-red))!important;color:#fff!important;box-shadow:0 7px 16px rgba(199,25,45,.20)!important;transition:.2s ease!important}
.stButton>button:hover,.stDownloadButton>button:hover{transform:translateY(-1px);border-color:var(--jirau-red-bright)!important;box-shadow:0 10px 22px rgba(199,25,45,.28)!important}
.stButton>button:disabled{opacity:.45;box-shadow:none!important}
[data-baseweb="input"]>div,[data-baseweb="select"]>div,textarea{border-radius:11px!important;border-color:var(--jirau-border)!important;background:#fff!important}
[data-baseweb="input"]>div:focus-within,[data-baseweb="select"]>div:focus-within{border-color:var(--jirau-red)!important;box-shadow:0 0 0 3px rgba(199,25,45,.10)!important}
.stTabs [data-baseweb="tab-list"]{gap:8px;background:#fff;border:1px solid var(--jirau-border);border-radius:14px;padding:6px;box-shadow:0 5px 18px rgba(23,32,42,.04)}
.stTabs [data-baseweb="tab"]{border-radius:9px;padding:9px 16px;font-weight:700;color:var(--jirau-muted)}
.stTabs [aria-selected="true"]{background:#FFF0F2!important;color:var(--jirau-red-dark)!important}
.stTabs [data-baseweb="tab-highlight"]{background:var(--jirau-red)!important}
[data-testid="stDataFrame"], [data-testid="stTable"]{border:1px solid var(--jirau-border);border-radius:16px;overflow:hidden;box-shadow:0 8px 22px rgba(23,32,42,.05);background:#fff}
h1,h2,h3,h4{color:var(--jirau-ink);letter-spacing:-.02em}h1{font-weight:800!important}h2,h3{font-weight:750!important}
hr{border-color:var(--jirau-border)!important}
[data-testid="stAlert"]{border-radius:14px;border-width:1px}
.jirau-footer{margin-top:34px;padding:16px 20px;border-top:1px solid var(--jirau-border);color:var(--jirau-muted);font-size:11px;text-align:center;letter-spacing:.03em}
@media(max-width:900px){.jirau-header-logo{display:none}.jirau-header{padding:18px}.jirau-title{font-size:20px}.premium-hero{padding:22px}}
</style>
"""
