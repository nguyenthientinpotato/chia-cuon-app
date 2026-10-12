"""
App chia cuon mang - giao dien Streamlit.
Chay:  streamlit run app.py   (hoac double-click run_app.bat / run_app.command)
"""
from __future__ import annotations

import csv
import html
import json
from pathlib import Path

import pandas as pd
import streamlit as st

import export_excel
import slitter as s

APP_DIR = Path(__file__).parent
PRODUCT_CACHE = APP_DIR / "products_cache.csv"

# Bang mau phan loai (fixed order, khong lap vong). Item thu 9 tro di -> xam + nhan.
SERIES = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
OVERFLOW = "#8a8987"
STATUS_VI = {"OK": "✅ Đạt", "VUOT DUNG SAI": "⚠️ Vượt dung sai", "THIEU": "❌ Thiếu",
             "DU (VUA DU)": "⚠️ Dư (đơn vừa đủ)", "STOCK": "Stock", "-": "Không dùng"}

# initial_sidebar_state=<px>: thanh ben mac dinh hep (van keo rong duoc)
st.set_page_config(page_title="Chia cuộn màng", page_icon="🧻", layout="wide", initial_sidebar_state=300)


# ----------------------------------------------------------------------------
# Du lieu mac dinh & state
# ----------------------------------------------------------------------------
ORDER_COLS = {"id": "Mã đơn", "width_mm": "Khổ (mm)", "kg": "KL cần (kg)",
              "roll_length_m": "Số m/cuộn", "over_tol_pct": "Dung sai dư (%)", "exact": "Chia vừa đủ"}
STOCK_COLS = {"priority": "Ưu tiên", "width_mm": "Khổ (mm)", "roll_length_m": "Số m/cuộn",
              "max_kg": "KL tối đa (kg)"}
# Moi dong = 1 lo NVL cap (KL tong + so cuon). "Số m/cuộn" tu tinh (khong nhap).
# Dinh luong dung chung cho moi lo (o 'Định lượng' canh o chon san pham).
MASTER_COLS = {"id": "Mã cuộn", "width_mm": "Khổ NL (mm)",
               "available_kg": "KL cấp (kg)", "rolls": "Số cuộn", "length_m": "Số m/cuộn",
               "min_trim_mm": "Biên min (mm)", "max_trim_mm": "Biên max (mm)"}
MC = MASTER_COLS
TRIM_DEFAULT = (10.0, 60.0)
DEFAULT_GSM = 119.14

DEFAULTS = {
    "product": None, "max_strips": 12, "min_strip": 40.0, "max_strip": 1250.0,
    "multiples": [1, 2, 3, 4], "reslit_trim": 0.0, "len_tol": 5.0, "allow_join": True,
    "time_limit": 30, "mip_solver": "SCIP",
    **{f"w_{k}": v for k, v in vars(s.Weights()).items()},
}
DEFAULT_ORDERS = pd.DataFrame([
    {"Mã đơn": "KH-280", "Khổ (mm)": 280.0, "KL cần (kg)": 178.0, "Số m/cuộn": 2000.0, "Dung sai dư (%)": 10.0, "Chia vừa đủ": False},
    {"Mã đơn": "KH-222", "Khổ (mm)": 222.0, "KL cần (kg)": 14.0, "Số m/cuộn": 2000.0, "Dung sai dư (%)": 10.0, "Chia vừa đủ": False},
    {"Mã đơn": "KH-238", "Khổ (mm)": 238.0, "KL cần (kg)": 575.0, "Số m/cuộn": 2000.0, "Dung sai dư (%)": 10.0, "Chia vừa đủ": False},
])
DEFAULT_STOCK = pd.DataFrame(columns=list(STOCK_COLS.values())).astype(float)


def _num(x) -> float | None:
    """So tu o bang (NaN / trong / chu -> None)."""
    try:
        x = float(x)
    except (TypeError, ValueError):
        return None
    return None if pd.isna(x) else x


def _cells(col: pd.Series) -> list:
    """Gia tri 1 cot de so sanh, khong phu thuoc dtype (NaN / None / '' -> None)."""
    return [None if pd.isna(x) or x == "" else (float(x) if isinstance(x, (int, float)) else str(x))
            for x in col.tolist()]


def with_meters(df: pd.DataFrame, gsm: float | None = None) -> pd.DataFrame:
    """Tu dien 'Số m/cuộn' = KL / so cuon / (kho x dinh luong chung), 'Số cuộn' trong = 1
    va 'Mã cuộn' con trong (NL1, NL2...)."""
    g = gsm or st.session_state.get("gsm") or DEFAULT_GSM
    out = df.reset_index(drop=True).copy()
    if MC["rolls"] not in out:  # bang cu chua co cot so cuon
        out.insert(list(MC).index("rolls"), MC["rolls"], 1.0)
    for c in (MC["width_mm"], MC["available_kg"], MC["rolls"], MC["min_trim_mm"], MC["max_trim_mm"]):
        out[c] = pd.to_numeric(out[c], errors="coerce").astype(float)
    out[MC["rolls"]] = out[MC["rolls"]].where(out[MC["rolls"]] >= 1, 1.0).round(0)
    kg, w, n = out[MC["available_kg"]], out[MC["width_mm"]], out[MC["rolls"]]
    out[MC["length_m"]] = (kg / n * 1e6 / (w * g)).where((kg > 0) & (w > 0)).round(0)
    ids = [str(x).strip() if pd.notna(x) and str(x).strip() else "" for x in out[MC["id"]]]
    n = 0
    for i, x in enumerate(ids):
        if not x:
            n += 1
            while f"NL{n}" in ids:
                n += 1
            ids[i] = f"NL{n}"
    out[MC["id"]] = ids
    return out


DEFAULT_MASTERS = with_meters(pd.DataFrame([{
    MC["id"]: "NL1", MC["width_mm"]: 1030.0, MC["available_kg"]: 0.0,
    MC["rolls"]: 1, MC["length_m"]: None, MC["min_trim_mm"]: TRIM_DEFAULT[0], MC["max_trim_mm"]: TRIM_DEFAULT[1]}]),
    DEFAULT_GSM)

st.session_state.setdefault("gsm", DEFAULT_GSM)
for k, v in DEFAULTS.items():
    st.session_state.setdefault(k, v)
st.session_state.setdefault("masters_df", DEFAULT_MASTERS)
st.session_state.setdefault("orders_df", DEFAULT_ORDERS)
st.session_state.setdefault("stock_df", DEFAULT_STOCK)
st.session_state.setdefault("editor_ver", 0)
st.session_state.setdefault("masters_ver", 0)


@st.cache_data(show_spinner=False)
def load_product_cache(mtime: float) -> dict:
    if not PRODUCT_CACHE.exists():
        return {}
    with open(PRODUCT_CACHE, encoding="utf-8-sig") as f:
        return {r["ten_san_pham"]: float(r["dinh_luong_gsm"]) for r in csv.DictReader(f)}


def products() -> dict:
    return load_product_cache(PRODUCT_CACHE.stat().st_mtime if PRODUCT_CACHE.exists() else 0.0)


def on_product_change():
    """Chon san pham -> dien o dinh luong (dung chung cho moi lo NVL); so m/cuon tu tinh lai."""
    ss = st.session_state
    name = ss.product
    if name and name in products():
        ss.gsm = products()[name]


def master_list() -> list:
    """Cac dong hop le cua bang NVL -> list dict theo schema 'masters' cua slitter."""
    ss = st.session_state
    g = ss.gsm
    out = []
    for _, r in ss.get("masters_now", ss.masters_df).iterrows():
        w = _num(r[MC["width_mm"]])
        if not w or not g or w <= 0 or g <= 0:
            continue
        tmin, tmax = _num(r[MC["min_trim_mm"]]), _num(r[MC["max_trim_mm"]])
        out.append({"id": r[MC["id"]], "width_mm": w, "gsm": g,
                    "available_kg": _num(r[MC["available_kg"]]) or None,
                    "rolls": int(_num(r[MC["rolls"]]) or 1),
                    "min_trim_mm": TRIM_DEFAULT[0] if tmin is None else tmin,
                    "max_trim_mm": TRIM_DEFAULT[1] if tmax is None else tmax})
    return out


def ref_gsm() -> float:
    """Dinh luong chung (dung cho goi y KL 1 cuon con)."""
    return st.session_state.gsm or DEFAULT_GSM


def build_config() -> dict:
    ss = st.session_state
    orders = []
    for _, r in ss.orders_now.iterrows():
        if pd.isna(r["Khổ (mm)"]) or pd.isna(r["KL cần (kg)"]) or not r["KL cần (kg)"]:
            continue
        orders.append({"id": str(r["Mã đơn"]) if pd.notna(r["Mã đơn"]) else "",
                       "width_mm": float(r["Khổ (mm)"]), "kg": float(r["KL cần (kg)"]),
                       "roll_length_m": float(r["Số m/cuộn"]) if pd.notna(r["Số m/cuộn"]) else 2000.0,
                       "over_tol_pct": float(r["Dung sai dư (%)"]) if pd.notna(r["Dung sai dư (%)"]) else 0.0,
                       "exact": bool(r.get("Chia vừa đủ")) if pd.notna(r.get("Chia vừa đủ")) else False})
    stock_df = ss.stock_now.dropna(subset=["Khổ (mm)", "Số m/cuộn", "KL tối đa (kg)"])
    stock_df = stock_df.sort_values("Ưu tiên", na_position="last", kind="stable")
    stock = [{"width_mm": float(r["Khổ (mm)"]), "roll_length_m": float(r["Số m/cuộn"]),
              "max_kg": float(r["KL tối đa (kg)"])} for _, r in stock_df.iterrows()
             if r["KL tối đa (kg)"] > 0]
    masters = master_list()
    return {
        "master": {"product": ss.product, "gsm": ss.gsm},
        "masters": masters,
        "machine": {"max_strips": int(ss.max_strips), "min_strip_mm": ss.min_strip,
                    "max_strip_mm": ss.max_strip, "multiples": sorted(ss.multiples) or [1],
                    "reslit_trim_mm": ss.reslit_trim, "length_tol_pct": ss.len_tol,
                    "allow_join": ss.allow_join},
        "orders": orders, "stock": stock,
        "weights": {k: ss[f"w_{k}"] for k in vars(s.Weights())},
        "solver": {"time_limit_s": ss.time_limit, "mip_solver": ss.mip_solver},
    }


def apply_config(cfg: dict) -> None:
    """Nap 1 kich ban JSON vao giao dien (goi TRUOC khi ve widget)."""
    ss = st.session_state
    m, mc = cfg.get("master", {}), cfg.get("machine", {})
    ss.product = m.get("product") if m.get("product") in products() else None
    # Kich ban cu chi co 1 'master' (bien phe nam trong 'machine') -> thanh bang 1 dong
    rows = cfg.get("masters") or ([m] if m.get("width_mm") else [])
    gsm = m.get("gsm") or next((x["gsm"] for x in rows if x.get("gsm")), None) \
        or products().get(ss.product) or ss.gsm
    ss.gsm = float(gsm)
    trim = (mc.get("min_trim_mm", TRIM_DEFAULT[0]), mc.get("max_trim_mm", TRIM_DEFAULT[1]))
    ss.masters_df = with_meters(pd.DataFrame(
        [{MC["id"]: x.get("id"), MC["width_mm"]: x.get("width_mm"),
          MC["available_kg"]: x.get("available_kg") or 0.0, MC["rolls"]: x.get("rolls") or 1,
          MC["length_m"]: None,
          MC["min_trim_mm"]: trim[0] if x.get("min_trim_mm") is None else x["min_trim_mm"],
          MC["max_trim_mm"]: trim[1] if x.get("max_trim_mm") is None else x["max_trim_mm"]} for x in rows],
        columns=list(MC.values())) if rows else DEFAULT_MASTERS)
    ss.pop("masters_now", None)
    for key, src in [("min_strip", "min_strip_mm"),
                     ("max_strip", "max_strip_mm"), ("reslit_trim", "reslit_trim_mm"), ("len_tol", "length_tol_pct")]:
        if src in mc:
            ss[key] = float(mc[src])
    if "max_strips" in mc:
        ss.max_strips = int(mc["max_strips"])
    if "multiples" in mc:
        ss.multiples = [int(x) for x in mc["multiples"] if int(x) in (1, 2, 3, 4)]
    if "allow_join" in mc:
        ss.allow_join = bool(mc["allow_join"])
    for k, v in cfg.get("weights", {}).items():
        if f"w_{k}" in ss:
            ss[f"w_{k}"] = float(v)
    sv = cfg.get("solver", {})
    ss.time_limit = int(sv.get("time_limit_s", ss.time_limit))
    ss.mip_solver = sv.get("mip_solver", ss.mip_solver)
    ss.orders_df = pd.DataFrame([{ORDER_COLS[k]: o.get(k, False if k == "exact" else None) for k in ORDER_COLS}
                                 for o in cfg.get("orders", [])], columns=list(ORDER_COLS.values()))
    ss.stock_df = pd.DataFrame([{"Ưu tiên": i, "Khổ (mm)": x["width_mm"], "Số m/cuộn": x["roll_length_m"],
                                 "KL tối đa (kg)": x["max_kg"]}
                                for i, x in enumerate(cfg.get("stock", []), start=1)],
                               columns=list(STOCK_COLS.values()))
    ss.editor_ver += 1
    ss.pop("result", None)


# ----------------------------------------------------------------------------
# So do dao (SVG)
# ----------------------------------------------------------------------------
def item_color(idx: int) -> str:
    return SERIES[idx] if idx < len(SERIES) else OVERFLOW


def cut_diagram(p: s.Problem, pt: s.Pattern) -> str:
    """1 thanh ngang = kho NL. Moi cuon con 1 khoi; bien phe gach cheo 2 ben."""
    VW, H, BAR_Y, BAR_H = 1000.0, 92, 18, 46
    master_w = s.master_of(p, pt).width_mm
    scale = VW / master_w
    product_w = s.product_width(p, pt)
    trim = master_w - product_w
    parts = []
    x = trim / 2 * scale
    defs = ('<defs><pattern id="hatch" width="6" height="6" patternUnits="userSpaceOnUse" '
            'patternTransform="rotate(45)"><line x1="0" y1="0" x2="0" y2="6" stroke="#8a8987" '
            'stroke-width="1.5"/></pattern></defs>')
    if trim > 0:
        for tx in (0, VW - trim / 2 * scale):
            parts.append(f'<rect x="{tx:.1f}" y="{BAR_Y}" width="{trim / 2 * scale:.1f}" height="{BAR_H}" '
                         f'fill="url(#hatch)" opacity="0.7"><title>Biên phế {trim / 2:g} mm</title></rect>')
    for md, c in pt.counts:
        it = p.items[md.item]
        col = item_color(md.item)
        sw = it.width_mm * md.mult
        for _ in range(c):
            w = sw * scale
            tip = (f"{'STOCK' if it.is_stock else ', '.join(it.order_ids)} · khổ {it.width_mm:g} mm"
                   + (f" ×{md.mult} (xẻ lại sau)" if md.mult > 1 else "")
                   + (" · nối 2 đoạn" if md.join == 2 else ""))
            dash = ' stroke-dasharray="5 3"' if it.is_stock else ""
            parts.append(
                f'<g><title>{html.escape(tip)}</title>'
                f'<rect x="{x + 1:.1f}" y="{BAR_Y}" width="{max(w - 2, 1):.1f}" height="{BAR_H}" rx="3" '
                f'fill="{col}" fill-opacity="0.18" stroke="{col}" stroke-width="1.5"{dash}/>'
                f'<rect x="{x + 1:.1f}" y="{BAR_Y}" width="{max(w - 2, 1):.1f}" height="5" rx="2" fill="{col}"/>')
            for j in range(1, md.mult):  # duong xe lai cua kho doi/ba/tu
                lx = x + w * j / md.mult
                parts.append(f'<line x1="{lx:.1f}" y1="{BAR_Y + 8}" x2="{lx:.1f}" y2="{BAR_Y + BAR_H - 3}" '
                             f'stroke="{col}" stroke-width="1" stroke-dasharray="3 3"/>')
            if w >= 34:
                label = f"{it.width_mm:g}" if md.mult == 1 else f"{md.mult}×{it.width_mm:g}"
                if md.join == 2:
                    label += " ⛓" if w < 70 else " · nối"
                parts.append(f'<text x="{x + w / 2:.1f}" y="{BAR_Y + BAR_H / 2 + 7}" text-anchor="middle" '
                             f'font-size="14" fill="currentColor">{label}</text>')
            parts.append("</g>")
            x += w
    axis = (f'<text x="0" y="12" font-size="12" fill="currentColor" opacity="0.65">0</text>'
            f'<text x="{VW}" y="12" font-size="12" text-anchor="end" fill="currentColor" opacity="0.65">'
            f'{master_w:g} mm</text>'
            f'<text x="{VW / 2}" y="{BAR_Y + BAR_H + 22}" font-size="12" text-anchor="middle" '
            f'fill="currentColor" opacity="0.65">dùng {product_w:g} mm · biên phế {trim:g} mm '
            f'({trim / master_w:.1%})</text>')
    return (f'<svg viewBox="-4 0 {VW + 8} {H}" width="100%" role="img" '
            f'aria-label="Sơ đồ dao {html.escape(s.formula_str(p, pt))}" '
            f'style="max-width:1100px;display:block;font-family:inherit">{defs}{"".join(parts)}{axis}</svg>')


def hint_table(rows: list) -> str:
    """Bang uoc tinh so cuon gon: chu nho, vua be rong cot. Don < 1 cuon (chac chan du) to vang."""
    cols = [("Mã đơn", "Mã đơn", "{}", "left"), ("Khổ (mm)", "Khổ", "{:g}", "right"),
            ("Số m/cuộn", "m/cuộn", "{:g}", "right"), ("KL cần (kg)", "KL cần", "{:,.1f}", "right"),
            ("KL 1 cuộn (kg)", "KL/cuộn", "{:,.1f}", "right"), ("Cần (cuộn)", "Cần (cuộn)", "{:,.2f}", "right")]
    cell = "padding:3px 8px;border-bottom:1px solid rgba(128,128,128,0.2);"
    head = "".join(f'<th style="{cell}text-align:{a};font-weight:600;white-space:nowrap">{lab}</th>'
                   for _, lab, _, a in cols)
    body = []
    for r in rows:
        bg = "background:rgba(237,161,0,0.15);" if r["Cần (cuộn)"] < 1 else ""
        tds = []
        for key, _, fmt, a in cols:
            tds.append(f'<td style="{cell}text-align:{a};white-space:nowrap">{html.escape(fmt.format(r[key]))}</td>')
        tip = ' title="Nhỏ hơn 1 cuộn, chắc chắn dư"' if bg else ""
        body.append(f'<tr style="{bg}"{tip}>{"".join(tds)}</tr>')
    return (f'<table style="width:100%;border-collapse:collapse;font-size:0.78rem;line-height:1.35;'
            f'margin:-0.5rem 0 1rem">'
            f'<thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table>')


def legend(p: s.Problem) -> str:
    chips = []
    for i, it in enumerate(p.items):
        name = f"Stock #{it.priority}" if it.is_stock else ", ".join(it.order_ids)
        chips.append(f'<span style="display:inline-flex;align-items:center;gap:6px;margin:0 16px 6px 0">'
                     f'<span style="width:12px;height:12px;border-radius:3px;background:{item_color(i)};'
                     f'display:inline-block"></span>{html.escape(name)} · {it.width_mm:g} mm × '
                     f'{it.roll_length_m:g} m</span>')
    return f'<div style="font-size:0.9rem">{"".join(chips)}</div>'


# ----------------------------------------------------------------------------
# Sidebar: thong so
# ----------------------------------------------------------------------------
with st.sidebar:
    st.header("Thông số")
    up_cfg = st.file_uploader("Mở kịch bản đã lưu (.json)", type=["json"], key="cfg_upload")
    if up_cfg is not None and st.session_state.get("cfg_loaded") != up_cfg.file_id:
        try:
            apply_config(json.loads(up_cfg.getvalue().decode("utf-8")))
            st.session_state.cfg_loaded = up_cfg.file_id
            st.toast("Đã nạp kịch bản")
        except (ValueError, KeyError, TypeError) as e:
            st.error(f"File kịch bản không đọc được: {e}")

    st.subheader("Máy chia")
    c1, c2 = st.columns(2)
    c1.number_input("Số cuộn con tối đa / bộ", min_value=1, max_value=60, step=1, key="max_strips",
                    help="Số dao = số cuộn con + 1")
    c2.number_input("Dung sai số mét (±%)", min_value=0.0, max_value=20.0, step=0.5, format="%.1f", key="len_tol")
    c1.number_input("Khổ con nhỏ nhất (mm)", min_value=1.0, step=5.0, format="%.0f", key="min_strip")
    c2.number_input("Khổ con lớn nhất (mm)", min_value=1.0, step=10.0, format="%.0f", key="max_strip")
    st.multiselect("Cho phép khổ đơn / đôi / ba / tư", options=[1, 2, 3, 4], key="multiples",
                   format_func=lambda m: {1: "Đơn", 2: "Đôi (×2)", 3: "Ba (×3)", 4: "Tư (×4)"}[m])
    st.number_input("Phế khi xẻ lại khổ đôi/ba/tư (mm)", min_value=0.0, step=1.0, format="%.1f", key="reslit_trim")
    st.toggle("Cho phép nối 2 cuộn thành 1", key="allow_join")

    with st.expander("Nâng cao: ưu tiên & solver"):
        st.caption("Chi phí quy về 'kg tương đương'. Số càng lớn, solver càng tránh.")
        labels = {"trim": "1 kg biên phế", "over_in_tol": "1 kg dư trong dung sai",
                  "over_beyond_tol": "1 kg dư vượt dung sai", "stock_base": "1 kg stock ưu tiên #1",
                  "stock_step": "Mỗi bậc ưu tiên stock sau +", "setup_per_pattern": "Mỗi công thức (đổi dao)",
                  "reslit_per_roll": "Mỗi cuộn phải xẻ lại", "join_per_roll": "Mỗi cuộn phải nối",
                  "length_dev": "Hệ số phạt lệch số mét", "over_exact": "1 kg dư ở đơn 'Chia vừa đủ'"}
        for k, lab in labels.items():
            st.number_input(lab, min_value=0.0, step=0.1, key=f"w_{k}")
        st.slider("Thời gian giải tối đa (giây)", 5, 300, key="time_limit")
        st.selectbox("Solver", ["SCIP", "HIGHS", "CBC"], key="mip_solver")

    with st.expander("Cập nhật danh mục sản phẩm"):
        st.caption(f"Hiện có {len(products()):,} sản phẩm có định lượng.")
        up_prod = st.file_uploader("File Excel có sheet 'HS sản phẩm'", type=["xlsx", "xlsm"], key="prod_upload")
        if up_prod is not None and st.button("Nạp danh mục"):
            try:
                rows = s.load_products(up_prod)
                with open(PRODUCT_CACHE, "w", newline="", encoding="utf-8-sig") as f:
                    w = csv.writer(f)
                    w.writerow(["ten_san_pham", "dinh_luong_gsm"])
                    w.writerows(rows)
                load_product_cache.clear()
                st.success(f"Đã nạp {len(rows):,} sản phẩm")
            except Exception as e:  # noqa: BLE001 - bao loi file cho Planner
                st.error(f"Không đọc được file: {e}")


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
st.title("Chia cuộn màng")
st.caption("Tìm phương án chia cuộn khổ lớn thành cuộn khổ nhỏ: đủ hàng, ít phế, ít dư nhất.")

tab_plan, tab_list = st.tabs(["Tính phương án", "Liệt kê công thức (kiểu VBA cũ)"])

with tab_plan:
    st.subheader("Nguyên vật liệu")
    prods = products()
    c_prod, c_gsm = st.columns([4, 1])
    c_prod.selectbox("Sản phẩm (gõ để tìm)", options=list(prods.keys()), index=None, key="product",
                     placeholder="Chọn để tự lấy định lượng…", on_change=on_product_change,
                     help="Danh mục lấy từ sheet 'HS sản phẩm'. Chọn sản phẩm sẽ tự điền định lượng. "
                          "Cập nhật danh mục ở mục cuối thanh bên.")
    c_gsm.number_input("Định lượng (g/m²)", min_value=0.1, step=1.0, format="%.2f", key="gsm",
                       help="Tự lấy theo sản phẩm đã chọn, dùng chung cho mọi lô NVL. Sửa tay được nếu cần.")
    ss = st.session_state
    edited = st.data_editor(
        ss.masters_df, key=f"masters_{ss.editor_ver}_{ss.masters_ver}",
        num_rows="dynamic", width="stretch", hide_index=True, disabled=[MC["length_m"]],
        column_config={
            MC["id"]: st.column_config.TextColumn(help="Tên / mã lô cuộn NVL. Để trống sẽ tự đặt NL1, NL2…"),
            MC["width_mm"]: st.column_config.NumberColumn(min_value=40, max_value=3000, format="%g",
                                                          default=DEFAULT_MASTERS[MC["width_mm"]][0]),
            MC["available_kg"]: st.column_config.NumberColumn(
                min_value=0, format="%.0f", default=0,
                help="Tổng KL NVL cấp của dòng này (kg) · 0 = không giới hạn"),
            MC["rolls"]: st.column_config.NumberColumn(
                min_value=1, step=1, format="%d", default=1,
                help="Số cuộn ứng với KL cấp (vd 3 cuộn tổng 1.500 kg). Dùng để tính số m/cuộn."),
            MC["length_m"]: st.column_config.NumberColumn(
                format="%.0f", help="Tự tính (tương đối): KL ÷ số cuộn ÷ (khổ × định lượng). "
                                    "Trống = không giới hạn."),
            MC["min_trim_mm"]: st.column_config.NumberColumn(
                min_value=0, format="%g", default=TRIM_DEFAULT[0],
                help="Biên phế tối thiểu (mm), tổng 2 bên, máy bắt buộc xén"),
            MC["max_trim_mm"]: st.column_config.NumberColumn(
                min_value=0, format="%g", default=TRIM_DEFAULT[1],
                help="Biên phế tối đa (mm) = Khổ NL − NL sử dụng (ô D1 cũ)"),
        })
    filled = with_meters(edited)
    if any(_cells(filled[c]) != _cells(edited[c]) for c in (MC["length_m"], MC["id"], MC["rolls"])):
        ss.masters_df = filled  # ve lai bang de hien so m / ma cuon / so cuon vua tu dien
        ss.masters_ver += 1
        st.rerun()
    ss.masters_now = filled
    st.caption("Mỗi dòng = 1 lô NVL cấp vào (KL tổng + số cuộn). Chỉ 1 dòng với KL 0 = như cũ "
               "(NVL không giới hạn). Nhiều dòng: solver tự chọn lô nào chạy lệnh nào, không vượt KL từng lô.")

    left, right = st.columns([3, 2], gap="large")
    with left:
        st.subheader("Đơn hàng")
        st.session_state.orders_now = st.data_editor(
            st.session_state.orders_df, key=f"orders_{st.session_state.editor_ver}",
            num_rows="dynamic", width="stretch", hide_index=True,
            column_config={
                "Mã đơn": st.column_config.TextColumn(required=False),
                "Khổ (mm)": st.column_config.NumberColumn(min_value=1, max_value=3000, format="%g"),
                "KL cần (kg)": st.column_config.NumberColumn(min_value=0, format="%.1f"),
                "Số m/cuộn": st.column_config.NumberColumn(min_value=1, format="%g", default=2000),
                "Dung sai dư (%)": st.column_config.NumberColumn(min_value=0, max_value=100, format="%g",
                                                                  default=10),
                "Chia vừa đủ": st.column_config.CheckboxColumn(
                    default=False, help="Ưu tiên tối đa KHÔNG dư cho đơn này (bỏ qua dung sai): "
                                        "phần thừa được đẩy sang stock / đơn khác / biên phế."),
            })
        # Goi y nhanh: 1 cuon nang bao nhieu, can may cuon
        hints = []
        for _, r in st.session_state.orders_now.iterrows():
            if pd.notna(r["Khổ (mm)"]) and pd.notna(r["Số m/cuộn"]) and pd.notna(r["KL cần (kg)"]):
                kg_roll = s.kg_strip(r["Khổ (mm)"], r["Số m/cuộn"], ref_gsm())
                if kg_roll > 0:
                    n = r["KL cần (kg)"] / kg_roll
                    hints.append({"Mã đơn": r["Mã đơn"] if pd.notna(r["Mã đơn"]) and r["Mã đơn"] else "",
                                  "Khổ (mm)": r["Khổ (mm)"], "Số m/cuộn": r["Số m/cuộn"],
                                  "KL cần (kg)": r["KL cần (kg)"], "KL 1 cuộn (kg)": kg_roll,
                                  "Cần (cuộn)": n})
        if hints:
            st.caption(f"Ước tính số cuộn cần (theo định lượng {ref_gsm():g} g/m²). "
                       "Dòng tô vàng: nhỏ hơn 1 cuộn, chắc chắn dư.")
        if hints:
            st.markdown(hint_table(hints), unsafe_allow_html=True)

    with right:
        st.subheader("Stock (tùy chọn)")
        st.session_state.stock_now = st.data_editor(
            st.session_state.stock_df, key=f"stock_{st.session_state.editor_ver}",
            num_rows="dynamic", width="stretch", hide_index=True,
            column_config={
                "Ưu tiên": st.column_config.NumberColumn(min_value=1, step=1, format="%d",
                                                         help="1 = quan trọng nhất"),
                "Khổ (mm)": st.column_config.NumberColumn(min_value=1, format="%g"),
                "Số m/cuộn": st.column_config.NumberColumn(min_value=1, format="%g", default=2000),
                "KL tối đa (kg)": st.column_config.NumberColumn(min_value=0, format="%.0f"),
            })
        st.caption("Stock chỉ được dùng khi lấp được biên phế. Phương án tốt nhất là không cần stock.")

    run = st.button("Tính phương án", type="primary")
    if run:
        cfg = build_config()
        if not cfg["masters"]:
            st.warning("Nhập ít nhất 1 cuộn NVL có khổ và định lượng (thanh bên trái).")
        elif not cfg["orders"]:
            st.warning("Nhập ít nhất 1 đơn hàng có khổ và KL cần.")
        else:
            try:
                with st.spinner(f"Đang tối ưu… (tối đa {st.session_state.time_limit} giây)"):
                    sol = s.solve(s.problem_from_config(cfg))
                st.session_state.result = (sol, cfg)
            except (RuntimeError, ValueError) as e:
                st.session_state.pop("result", None)
                st.error(str(e))

    if "result" in st.session_state:
        sol, cfg = st.session_state.result
        p = sol.problem
        t = s.totals(sol)
        summ = s.summary_rows(sol)
        st.divider()
        if sol.status == "OPTIMAL" and sol.method.startswith("liet ke"):
            st.success("Phương án **tối ưu** (đã chứng minh, xét toàn bộ công thức khả thi).")
        else:
            st.info(f"Phương án **tốt nhất tìm được** trong {p.time_limit_s:g} giây "
                    f"({sol.method}). Tăng thời gian giải ở mục Nâng cao nếu muốn tìm thêm.")

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("NVL cần", f"{t['NVL (kg)']:,.1f} kg", help="Tổng KL cuộn khổ lớn đưa vào máy")
        m2.metric(f"Phế biên · {t['% phe']:.2f}% NVL", f"{t['Phe (kg)']:,.1f} kg")
        m3.metric(f"Dư so với đơn · {t['Du don (kg)'] / t['Nhu cau (kg)'] * 100:.1f}% nhu cầu",
                  f"{t['Du don (kg)']:,.1f} kg")
        m4.metric("Stock", f"{t['Stock (kg)']:,.1f} kg")
        m5.metric("Lệnh / bộ", f"{t['So cong thuc']} / {t['Tong so bo']}")

        bad = [r for r in summ if r["Trang thai"] in ("VUOT DUNG SAI", "THIEU", "DU (VUA DU)")]
        for r in bad:
            it = next(i for i in p.items if i.key == r["Item"])
            one = s.kg_strip(it.width_mm, it.roll_length_m, p.gsm)
            why = (f"đơn chỉ cần {it.kg:g} kg, nhỏ hơn 1 cuộn (~{one:,.0f} kg)" if it.kg < one
                   else ("không có cách ghép khớp hơn trong ràng buộc biên phế / số dao hiện tại"
                         if sol.status == "OPTIMAL" and sol.method.startswith("liet ke")
                         else "chưa tìm được cách ghép khớp hơn; thử tăng thời gian giải hoặc nới biên phế"))
            if it.exact:
                st.warning(f"**{r['Loai']}** ({it.width_mm:g} mm, chia vừa đủ) vẫn dư {r['Du (kg)']:,.1f} kg "
                           f"= {r['% du']}%: {why}" + ("" if any(i.is_stock for i in p.items)
                                                     else "; thêm stock để có chỗ đẩy phần thừa") + ".")
            else:
                st.warning(f"**{r['Loai']}** ({it.width_mm:g} mm) dư {r['Du (kg)']:,.1f} kg = {r['% du']}% "
                           f"> dung sai {r['Dung sai %']}%: {why}.")

        multi = len(p.masters) > 1
        if multi:
            st.subheader("Sử dụng NVL theo cuộn")
            st.dataframe(pd.DataFrame([{
                "Cuộn": r["Cuon NL"], "Khổ (mm)": r["Kho NL"], "g/m²": r["Dinh luong"],
                "Cấp (kg)": r["Cap (kg)"] if isinstance(r["Cap (kg)"], float) else None,
                "Dùng (kg)": r["Dung (kg)"],
                "Còn lại (kg)": r["Con lai (kg)"] if r["Con lai (kg)"] != "" else None,
                "Số cuộn": r["So cuon"], "Số m/cuộn": r["So m/cuon"] or None, "Số m dùng": r["So m dung"],
                "Biên phế (mm)": r["Bien phe (mm)"],
            } for r in s.master_rows(sol)]).style.format(precision=1, na_rep="không giới hạn"),
                hide_index=True, width="stretch")

        st.subheader("Sơ đồ dao")
        st.markdown(legend(p), unsafe_allow_html=True)
        plan = s.plan_rows(sol)
        for n, (pt, sets) in enumerate(sol.runs, start=1):
            rows = [r for r in plan if r["Lenh"] == n]
            master_kg = rows[0]["SL NL can (kg)"]
            nl = f" · cuộn **{rows[0]['Cuon NL']}** ({rows[0]['Kho NL']:g} mm)" if multi else ""
            st.markdown(f"**Lệnh {n}**{nl} · `{s.formula_str(p, pt)}` · **{sets} bộ × {pt.length_m:,} m** · "
                        f"NVL {master_kg:,.1f} kg · phế {rows[0]['Phe (kg)']:,.1f} kg")
            st.markdown(cut_diagram(p, pt), unsafe_allow_html=True)
            with st.expander(f"Chi tiết lệnh {n}"):
                st.dataframe(pd.DataFrame([{
                    "Khổ trên máy (mm)": r["Kho cuon con"], "SD": r["SD"], "Loại": r["Loai"],
                    "Cho đơn": r["Don"], "Cuộn trên máy": r["SC tren may"],
                    "Cuộn thành phẩm": r["SC thanh pham"], "Nối cuộn": r["Noi cuon"], "Số kg": r["So kg"],
                } for r in rows]), hide_index=True, width="stretch")

        st.subheader("Đáp ứng theo đơn")
        df = pd.DataFrame([{
            "Đơn": r["Loai"] if not r["Loai"].startswith("stock") else r["Loai"].replace("stock", "Stock"),
            "Khổ × số m": f"{it.width_mm:g} mm × {it.roll_length_m:g} m",
            "Cần (kg)": r["Can (kg)"] or None, "Ra (kg)": r["Ra (kg)"],
            "Dư (kg)": r["Du (kg)"], "% dư": r["% du"] if r["% du"] != "" else None,
            "Dung sai": ("vừa đủ" if it.exact else f"{r['Dung sai %']:g}%")
                        if isinstance(r["Dung sai %"], (int, float)) else r["Dung sai %"],
            "Số cuộn": r["So cuon"], "Trạng thái": STATUS_VI.get(r["Trang thai"], r["Trang thai"]),
        } for r, it in zip(summ, p.items)])

        def _hl(row):
            bad_row = row["Trạng thái"] in (STATUS_VI["VUOT DUNG SAI"], STATUS_VI["THIEU"], STATUS_VI["DU (VUA DU)"])
            return ["background-color: rgba(227,73,72,0.15)" if bad_row else "" for _ in row]

        st.dataframe(df.style.apply(_hl, axis=1).format(precision=1, na_rep=""), hide_index=True,
                     width="stretch")

        d1, d2, _ = st.columns([1, 1, 3])
        d1.download_button("⬇️ Xuất Excel lệnh chia", data=export_excel.to_excel_bytes(sol, cfg),
                           file_name="lenh_chia.xlsx", type="primary",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2.download_button("💾 Lưu kịch bản (.json)",
                           data=json.dumps(cfg, ensure_ascii=False, indent=2).encode("utf-8"),
                           file_name="kich_ban_chia_cuon.json", mime="application/json")

with tab_list:
    st.caption("Liệt kê công thức hợp lệ cho đơn đang nhập, sắp theo biên phế tăng dần "
               "(giống bảng FormulaFastCal cũ). Dùng khi Planner muốn tự chọn công thức. "
               f"Quá {s.LIST_LIMIT:,} công thức thì chỉ giữ {s.LIST_LIMIT:,} công thức ít phế nhất. "
               "Công thức có trong phương án đã tính (tab 'Tính phương án') được tô màu theo lệnh chia.")
    top = st.number_input("Số công thức hiển thị", 10, 2000, 100, step=10)
    if st.button("Liệt kê"):
        try:
            cfg = build_config()
            if not cfg["masters"] or not cfg["orders"]:
                raise ValueError("Nhập ít nhất 1 cuộn NVL và 1 đơn hàng.")
            p = s.problem_from_config(cfg)
            with st.spinner("Đang liệt kê công thức…"):
                pats, complete = s.pattern_listing(p)
            uniq = s.sort_pattern_rows([s.pattern_row(p, pt) for pt in pats])

            # Cong thuc cua phuong an da tinh -> (cuon NL, ten CT): [so lenh]
            plan, plan_rows = {}, {}
            if "result" in st.session_state:
                sol = st.session_state.result[0]
                for n, (pt, _) in enumerate(sol.runs, start=1):
                    r = s.pattern_row(sol.problem, pt)
                    plan.setdefault((r["Cuon NL"], r["Ten CT"]), []).append(n)
                    plan_rows.setdefault((r["Cuon NL"], r["Ten CT"]), r)
            rank = {(r["Cuon NL"], r["Ten CT"]): i for i, r in enumerate(uniq, start=1)}
            shown = list(uniq[:top])
            shown_keys = {(r["Cuon NL"], r["Ten CT"]) for r in shown}
            # CT cua phuong an luon hien, ke ca khi nam ngoai top / ngoai gioi han liet ke
            for key in plan:
                if key not in shown_keys:
                    shown.append(next((r for r in uniq if (r["Cuon NL"], r["Ten CT"]) == key), plan_rows[key]))
            shown = s.sort_pattern_rows(shown)

            if complete:
                st.write(f"{len(uniq):,} công thức khả thi")
            else:
                st.info(f"Quá nhiều công thức: đã giữ {len(uniq):,} công thức ít phế nhất "
                        f"(biên phế ≤ {uniq[-1]['PL (mm)']:g} mm). Công thức nhiều phế hơn bị bỏ qua.")
            if plan:
                st.caption("Tô màu = công thức trong phương án đã tính gần nhất · "
                           + " · ".join(f"Lệnh {', '.join(map(str, v))}: `{k[1]}`"
                                        + ("" if k in rank else " (ngoài danh sách liệt kê)")
                                        for k, v in sorted(plan.items(), key=lambda kv: kv[1][0])))
            multi = len(cfg["masters"]) > 1
            df = pd.DataFrame([{
                "Lệnh chia": ", ".join(f"Lệnh {n}" for n in plan.get((r["Cuon NL"], r["Ten CT"]), [])),
                "Hạng": rank.get((r["Cuon NL"], r["Ten CT"])),
                **({"Cuộn NL": f"{r['Cuon NL']} ({r['Kho NL']:g} mm)"} if multi else {}),
                "Tên CT": r["Ten CT"], "Khổ SD (mm)": r["Kho SD"],
                "PL (mm)": r["PL (mm)"], "%PL": r["%PL"],
            } for r in shown])
            lenh_of = [plan.get((r["Cuon NL"], r["Ten CT"]), [None])[0] for r in shown]

            def _hl_plan(row):
                n = lenh_of[row.name]
                return [f"background-color: {item_color(n - 1)}33; font-weight: 600" if n else ""
                        for _ in row]

            st.dataframe(df.style.apply(_hl_plan, axis=1).format(precision=2, na_rep=""),
                         hide_index=True, width="stretch",
                         column_config={"Hạng": st.column_config.NumberColumn(
                                            format="%d", help="Thứ hạng theo biên phế trong danh sách liệt kê"),
                                        "Khổ SD (mm)": st.column_config.NumberColumn(format="%g"),
                                        "PL (mm)": st.column_config.NumberColumn(format="%g")})
        except (RuntimeError, ValueError) as e:
            st.error(str(e))
