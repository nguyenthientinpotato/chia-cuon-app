"""
slitter.py - Toi uu chia cuon mang (slitting / 1D cutting stock).

Thay the CTCC2.xlsm (VBA Find_ver1 / Find_ver2). Khac biet chinh:
  * VBA chi liet ke 1 cong thuc (pattern) roi Planner tu quyet so bo.
  * Ban nay giai CA BAI: chon nhom cong thuc + so bo moi cong thuc sao cho
    du nhu cau voi tong (bien phe + du + stock) nho nhat.

Mo hinh (tom tat):
  - 1 "bo" (set) = 1 lan cuon NVL dai L met, xen thanh cac cuon con.
  - Pattern p = (L, {cuon con}). Bien nguyen k_p = so bo chay pattern p.
  - KL ra cua 1 cuon con rong w, dai L:  w/1000 * L * gsm/1000  (kg)
  - Rang buoc: du nhu cau (cung), du <= dung sai (mem, phat nang),
    stock <= KL toi da (cung), bien phe trong [min_trim, max_trim],
    nhieu cuon NVL cap (khac kho / gsm / bien phe): moi CT chay tren 1 cuon,
    tong KL chay tren 1 cuon <= KL cap cua cuon do,
    so cuon con <= max_strips, cuon con trong [min_strip, max_strip],
    do dai bo L nam trong +/- length_tol cua chuan tung don
    (hoac 2*L neu cho phep noi 2 cuon).
  - Muc tieu (quy ve "kg tuong duong"): phe*1.0 + du_trong_dung_sai*0.3
    + du_vuot_dung_sai*3.0 + stock*(0.5..0.95 theo uu tien)
    + phat setup/pattern, xen lai kho doi, noi cuon, lech chieu dai.

Chi phu thuoc: ortools (GLOP cho LP, SCIP cho bai nguyen) + numpy. openpyxl chi can neu tra gsm tu file Excel.
"""
from __future__ import annotations

import csv
import json
import math
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

SCALE_W = 10      # do rong luu theo 0.1 mm (so nguyen)


# --------------------------------------------------------------------------
# Input
# --------------------------------------------------------------------------
@dataclass
class Machine:
    min_trim_mm: float = 0.0        # bien phe bat buoc toi thieu (tong 2 ben)
    max_trim_mm: float = 60.0       # = Kho NL - NL su dung (o D1 cu)
    max_strips: int = 30            # so cuon con toi da / bo (so dao = strips + 1)
    min_strip_mm: float = 40.0
    max_strip_mm: float = 1250.0
    multiples: list = field(default_factory=lambda: [1, 2, 3, 4])  # kho don/doi/ba/tu
    reslit_trim_mm: float = 0.0     # phe them khi xen lai 1 cuon kho doi/ba/tu
    length_tol_pct: float = 5.0     # cuon dai/ngan hon chuan toi da +/- %
    allow_join: bool = True         # cho phep noi 2 cuon cung kho thanh 1


@dataclass
class Weights:
    """Chi phi tuong duong 'kg phe'. Chinh o day de doi uu tien."""
    trim: float = 1.0               # 1 kg bien phe
    over_in_tol: float = 0.3        # 1 kg du nhung trong dung sai khach nhan
    over_beyond_tol: float = 3.0    # 1 kg du vuot dung sai (thanh ton kho ngoai y muon)
    stock_base: float = 0.5         # 1 kg stock uu tien #1
    stock_step: float = 0.1         # moi bac uu tien sau +0.1 (toi da 0.95)
    setup_per_pattern: float = 5.0  # moi cong thuc khac nhau (doi dao) ~ 5 kg
    reslit_per_roll: float = 2.0    # moi cuon kho doi/ba/tu phai xen lai
    join_per_roll: float = 20.0     # moi cuon thanh pham phai noi (rat han che)
    length_dev: float = 1.0         # phat = length_dev * |lech %| * kg cuon do
    over_exact: float = 50.0        # 1 kg du o don 'Chia vua du' (bo qua dung sai, day sang stock/phe)


@dataclass
class Item:
    """1 dong nhu cau (order) hoac 1 dong stock."""
    key: str
    width_mm: float
    roll_length_m: float
    kg: float = 0.0                 # order: KL can; stock: 0
    over_tol_pct: float = 0.0       # order: dung sai du
    max_kg: float = 0.0             # stock: KL stock toi da
    is_stock: bool = False
    priority: int = 0               # stock: 1 = quan trong nhat
    order_ids: list = field(default_factory=list)
    exact: bool = False             # order: 'Chia vua du' - uu tien tuyet doi khong du

    @property
    def cap_kg(self) -> float:
        return self.max_kg if self.is_stock else self.kg * (1 + self.over_tol_pct / 100)


@dataclass
class Master:
    """1 cuon NVL cap vao may. Moi cuon co the khac kho / dinh luong / KL / bien phe."""
    width_mm: float
    gsm: float
    available_kg: Optional[float] = None    # None = khong gioi han
    min_trim_mm: Optional[float] = None     # None = lay theo Machine
    max_trim_mm: Optional[float] = None
    label: str = ""
    rolls: int = 1                          # so cuon ung voi available_kg (de tinh so m/cuon)

    @property
    def length_m(self) -> Optional[float]:
        """So met tuong doi cua 1 cuon (KL cap / so cuon / (kho x gsm)); None neu khong gioi han."""
        if not self.available_kg:
            return None
        return self.available_kg / max(self.rolls, 1) / kg_strip(self.width_mm, 1, self.gsm)


@dataclass
class Problem:
    master_width_mm: float        # cuon NVL dau tien (giu de tuong thich cu)
    gsm: float
    items: list
    machine: Machine
    weights: Weights
    available_kg: Optional[float] = None
    time_limit_s: float = 30.0
    max_patterns: int = 300_000   # tran cho che do --list
    enum_limit: int = 5_000       # vuot muc nay -> column generation
    mip_solver: str = "SCIP"      # SCIP | HIGHS | CBC
    length_grid_pct: float = 1.0   # buoc luoi do dai bo trong vung dung sai
    masters: list = field(default_factory=list)   # [Master]; rong -> 1 cuon tu master_width_mm/gsm

    def __post_init__(self):
        if not self.masters:
            self.masters = [Master(self.master_width_mm, self.gsm, self.available_kg, label="NL1")]

    def trim_range(self, mi: int) -> tuple:
        """(bien phe toi thieu, toi da) cua cuon NVL thu mi."""
        m, mc = self.masters[mi], self.machine
        return (mc.min_trim_mm if m.min_trim_mm is None else m.min_trim_mm,
                mc.max_trim_mm if m.max_trim_mm is None else m.max_trim_mm)


def load_products(xlsx_path) -> list:
    """Doc danh sach (ten san pham, dinh luong g/m2) tu sheet 'HS san pham'. Bo dong gsm = 0."""
    import openpyxl  # chi import khi can
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb["HS sản phẩm"]
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) > 7 and row[1] and isinstance(row[7], (int, float)) and row[7] > 0:
            out.append((str(row[1]).strip(), float(row[7])))
    wb.close()
    return out


def lookup_gsm(xlsx_path: str, product_name: str) -> float:
    """Tra Dinh luong (g/m2) tu sheet 'HS san pham' (cot B ten, cot H dinh luong)."""
    for name, gsm in load_products(xlsx_path):
        if name == product_name:
            return gsm
    raise ValueError(f"Khong tim thay san pham: {product_name}")


def load_problem(path: str) -> Problem:
    cfg = json.loads(Path(path).read_text(encoding="utf-8"))
    return problem_from_config(cfg, base_dir=Path(path).parent)


def problem_from_config(cfg: dict, base_dir=None) -> Problem:
    """Tao Problem tu dict (cung schema voi file JSON). Dung chung cho CLI va app."""
    m = cfg.get("master", {})
    gsm = m.get("gsm")
    if gsm is None and m.get("product_file"):
        gsm = lookup_gsm(str(Path(base_dir or ".") / m["product_file"]), m["product"])
    # 'masters': danh sach cuon NVL (moi cuon kho/gsm/KL/bien phe rieng). Khong co -> 1 cuon tu 'master'.
    masters = []
    for i, r in enumerate(cfg.get("masters") or [m], start=1):
        avail = r.get("available_kg")
        masters.append(Master(
            width_mm=float(r["width_mm"]), gsm=float(r.get("gsm") or gsm),
            available_kg=float(avail) if avail else None,
            min_trim_mm=None if r.get("min_trim_mm") is None else float(r["min_trim_mm"]),
            max_trim_mm=None if r.get("max_trim_mm") is None else float(r["max_trim_mm"]),
            label=str(r.get("id") or f"NL{i}"), rolls=int(r.get("rolls") or 1)))
    machine = Machine(**cfg.get("machine", {}))
    weights = Weights(**cfg.get("weights", {}))

    # Gop cac don cung (kho, so met, vua du) thanh 1 item -> giam doi xung, de giai hon.
    # Don 'vua du' tach rieng de phan du (neu co) duoc day sang don thuong cung kho.
    groups: dict = {}
    for o in cfg["orders"]:
        k = (float(o["width_mm"]), float(o["roll_length_m"]), bool(o.get("exact", False)))
        g = groups.setdefault(k, {"kg": 0.0, "cap": 0.0, "ids": []})
        tol = float(o.get("over_tol_pct", 0.0) or 0.0)
        g["kg"] += float(o["kg"])
        g["cap"] += float(o["kg"]) * (1 + tol / 100)
        g["ids"].append(str(o.get("id") or f"{k[0]:g}mm"))
    items = []
    for (w, L, exact), g in groups.items():
        tol = (g["cap"] / g["kg"] - 1) * 100 if g["kg"] > 0 else 0.0
        items.append(Item(key=f"{w:g}x{L:g}m" + (" VUA DU" if exact else ""), width_mm=w,
                          roll_length_m=L, kg=g["kg"], over_tol_pct=tol, order_ids=g["ids"],
                          exact=exact))
    for rank, s in enumerate(cfg.get("stock", []), start=1):
        items.append(Item(key=f"STOCK{rank}:{float(s['width_mm']):g}x{float(s['roll_length_m']):g}m",
                          width_mm=float(s["width_mm"]), roll_length_m=float(s["roll_length_m"]),
                          max_kg=float(s["max_kg"]), is_stock=True, priority=rank))
    solver = cfg.get("solver", {})
    return Problem(master_width_mm=masters[0].width_mm, gsm=masters[0].gsm, items=items,
                   machine=machine, weights=weights, available_kg=masters[0].available_kg,
                   time_limit_s=float(solver.get("time_limit_s", 30.0)),
                   max_patterns=int(solver.get("max_patterns", 300_000)),
                   enum_limit=int(solver.get("enum_limit", 5_000)),
                   mip_solver=solver.get("mip_solver", "SCIP"),
                   length_grid_pct=float(solver.get("length_grid_pct", 1.0)), masters=masters)


def validate(p: Problem) -> None:
    mc = p.machine
    errs = []
    if not p.masters:
        errs.append("chua co cuon NVL nao")
    for mi, m in enumerate(p.masters):
        lo, hi = p.trim_range(mi)
        if m.width_mm <= 0 or m.gsm <= 0:
            errs.append(f"{m.label}: kho NL va dinh luong phai > 0")
        if lo > hi:
            errs.append(f"{m.label}: bien phe toi thieu > toi da")
    for it in p.items:
        if not (mc.min_strip_mm <= it.width_mm <= mc.max_strip_mm):
            errs.append(f"{it.key}: kho {it.width_mm} ngoai [{mc.min_strip_mm}, {mc.max_strip_mm}]")
        if all(it.width_mm > m.width_mm - p.trim_range(mi)[0] for mi, m in enumerate(p.masters)):
            errs.append(f"{it.key}: kho lon hon kho NL tru bien")
        if not it.is_stock and it.kg <= 0:
            errs.append(f"{it.key}: KL can phai > 0")
    if errs:
        raise ValueError("Input loi:\n  - " + "\n  - ".join(errs))


# --------------------------------------------------------------------------
# Pattern generation
# --------------------------------------------------------------------------
@dataclass(frozen=True)
class Mode:
    item: int        # index trong p.items
    mult: int        # 1 = kho don, 2 = kho doi, ...
    join: int        # 1 = binh thuong, 2 = noi 2 doan thanh 1 cuon
    strip_w10: int   # do rong cuon con tren may (0.1 mm)


@dataclass
class Pattern:
    length_m: int
    counts: tuple    # ((mode, so cuon con), ...)
    used_w10: int
    master: int = 0  # index trong p.masters


def _accepts(item: Item, L: float, tol_pct: float, allow_join: bool) -> Optional[int]:
    lo, hi = item.roll_length_m * (1 - tol_pct / 100), item.roll_length_m * (1 + tol_pct / 100)
    if lo - 1e-6 <= L <= hi + 1e-6:
        return 1
    if allow_join and lo - 1e-6 <= 2 * L <= hi + 1e-6:
        return 2
    return None


def candidate_lengths(p: Problem) -> list:
    """Cac do dai bo L can thu.

    Voi moi don: luoi diem trong vung [chuan -tol%, chuan +tol%] (buoc length_grid_pct),
    tuong tu cho 1/2 chuan neu cho noi cuon, cong diem giua vung giao giua 2 don
    (de ghep chung 1 bo cac don khac so met). Do dai ngan hon giup giam du khi KL
    can khong chia het cho KL 1 cuon.
    """
    mc = p.machine
    t = mc.length_tol_pct / 100
    step = max(p.length_grid_pct, 0.1) / 100
    wins = []
    for it in p.items:
        L = it.roll_length_m
        wins.append((L * (1 - t), L * (1 + t), L))
        if mc.allow_join:
            wins.append((L * (1 - t) / 2, L * (1 + t) / 2, L / 2))
    cands = set()
    for lo, hi, L in wins:
        n = int(round(2 * t / step))
        cands.update(round(lo + (hi - lo) * i / n) for i in range(n + 1)) if n else cands.add(round(L))
        cands.add(round(L))
    # Don 'vua du': do dai bo de N cuon con ra dung KL can (N*KL 1 cuon = kg), theo gsm tung cuon NL
    for it in p.items:
        if it.is_stock or not it.exact:
            continue
        for gsm in {m.gsm for m in p.masters}:
            per_m = kg_strip(it.width_mm, 1, gsm)
            for j in ((1, 2) if mc.allow_join else (1,)):
                lo, hi = it.roll_length_m * (1 - t) / j, it.roll_length_m * (1 + t) / j
                for N in range(max(1, math.ceil(it.kg / (per_m * hi * j))),
                               math.floor(it.kg / (per_m * lo * j)) + 1):
                    cands.add(round(it.kg / (per_m * N * j)))
    for i in range(len(wins)):
        for j in range(i + 1, len(wins)):
            lo, hi = max(wins[i][0], wins[j][0]), min(wins[i][1], wins[j][1])
            if lo <= hi:
                cands.add(round((lo + hi) / 2))
    return sorted(L for L in cands
                  if any(_accepts(it, L, mc.length_tol_pct, mc.allow_join) for it in p.items))


def modes_for_length(p: Problem, L: int) -> list:
    mc = p.machine
    out = []
    for idx, it in enumerate(p.items):
        j = _accepts(it, L, mc.length_tol_pct, mc.allow_join)
        if j is None:
            continue
        for m in sorted(set(mc.multiples)):
            sw = it.width_mm * m + (mc.reslit_trim_mm if m > 1 else 0)
            if mc.min_strip_mm <= sw <= mc.max_strip_mm:
                out.append(Mode(idx, m, j, int(round(sw * SCALE_W))))
    return sorted(out, key=lambda md: -md.strip_w10)


class TooManyPatterns(RuntimeError):
    pass


def enumerate_patterns(p: Problem, cap: Optional[int] = None) -> list:
    """Liet ke TOAN BO cong thuc kha thi (nhu VBA, nhung khong bo sot)."""
    cap = cap or p.max_patterns
    mc = p.machine
    patterns: list = []
    lengths = candidate_lengths(p)

    for mi, L in ((mi, L) for mi in range(len(p.masters)) for L in lengths):
        W10 = int(round(p.masters[mi].width_mm * SCALE_W))
        tmin, tmax = p.trim_range(mi)
        lo = W10 - int(round(tmax * SCALE_W))
        hi = W10 - int(round(tmin * SCALE_W))
        modes = modes_for_length(p, L)
        n = len(modes)
        counts = [0] * n

        def dfs(i: int, used: int, strips: int) -> None:
            if lo <= used <= hi and strips > 0:
                patterns.append(Pattern(L, tuple((modes[k], counts[k]) for k in range(n) if counts[k]),
                                        used, mi))
                if len(patterns) > cap:
                    raise TooManyPatterns(
                        f"Qua nhieu cong thuc (> {cap}). Giam 'multiples', "
                        "tang min_trim/giam max_trim, hoac giam max_strips.")
            if strips >= mc.max_strips:
                return
            for k in range(i, n):
                w = modes[k].strip_w10
                if used + w > hi:
                    continue
                if used + w * (mc.max_strips - strips) < lo:
                    break  # modes sap giam dan: cac kho sau con nho hon, khong the lap toi bien toi da
                counts[k] += 1
                dfs(k, used + w, strips + 1)
                counts[k] -= 1

        dfs(0, 0, 0)
    return patterns


LIST_LIMIT = 20_000   # so CT khac nhau toi da cho bang liet ke (bang chi hien <= 2000; ~2 s ca lon)


def pattern_listing(p: Problem, limit: int = LIST_LIMIT) -> tuple:
    """`limit` CT khac ten it phe nhat (theo % bien phe) -> (patterns, day_du).

    Nhanh-can 1 luot: giu heap `limit` CT tot nhat; khi day, nguong bien phe toi da tu siet
    lai bang CT te nhat trong heap -> cat nhanh som. day_du=False neu da bo bot CT.
    """
    import heapq
    validate(p)
    mc = p.machine
    heap: list = []        # (-%phe, seq, key, pattern); heap[0] = CT nhieu phe nhat dang giu
    keys: set = set()
    sigs: set = set()
    state = {"seq": 0, "full_cut": False}
    lengths = candidate_lengths(p)

    for mi, L in ((mi, L) for mi in range(len(p.masters)) for L in lengths):
        W10 = int(round(p.masters[mi].width_mm * SCALE_W))
        tmin, tmax = p.trim_range(mi)
        lo0 = W10 - int(round(tmax * SCALE_W))
        hi = W10 - int(round(tmin * SCALE_W))
        modes = modes_for_length(p, L)
        n = len(modes)
        counts = [0] * n
        wm = [(p.items[md.item].width_mm, md.mult) for md in modes]   # khoa CT khong phu thuoc so met
        pw = [int(round(w * m * SCALE_W)) for w, m in wm]   # kho thanh pham (khong tinh phe xen lai)
        sig = (mi, tuple(wm), tuple(md.strip_w10 for md in modes))
        if not modes or hi <= 0 or sig in sigs:   # cung bo kho voi muc so met truoc -> CT y het
            continue
        sigs.add(sig)
        # reach[k][r]: bitset cac tong kho (don vi `unit`) xep duoc tu modes[k:] voi <= r cuon con.
        # Dung de bo ngay nhanh khong the cham vung bien phe hop le.
        unit = math.gcd(*[md.strip_w10 for md in modes])
        wu = [md.strip_w10 // unit for md in modes]
        K, cap_bits = mc.max_strips, (1 << (hi // unit + 1)) - 1
        reach = [[1] * (K + 1) for _ in range(n + 1)]
        for k in range(n - 1, -1, -1):
            for r in range(1, K + 1):
                reach[k][r] = (reach[k + 1][r] | (reach[k][r - 1] << wu[k])) & cap_bits

        def lower() -> int:
            """KL su dung toi thieu (0.1 mm) de con lot vao top `limit`."""
            if len(heap) < limit:
                return lo0
            return max(lo0, W10 - math.floor(-heap[0][0] * W10))

        lo = [lower()]

        def dfs(i: int, used: int, strips: int) -> None:
            if lo[0] <= used <= hi and strips > 0:
                key = (mi, tuple((wm[k], counts[k]) for k in range(n) if counts[k]))
                if key not in keys:
                    pt = Pattern(L, tuple((modes[k], counts[k]) for k in range(n) if counts[k]), used, mi)
                    # % phe that = kho NL - kho thanh pham (gom ca phe xen lai) -> giong cot PL.
                    # Nguong cat nhanh theo `used` van an toan vi phe that >= W10 - used.
                    ratio = (W10 - sum(pw[k] * counts[k] for k in range(n) if counts[k])) / W10
                    if len(heap) < limit:
                        heapq.heappush(heap, (-ratio, state["seq"], key, pt))
                        keys.add(key)
                    elif ratio < -heap[0][0]:
                        heapq.heapreplace(heap, (-ratio, state["seq"], key, pt))
                        keys.add(key)
                        state["full_cut"] = True
                    else:
                        state["full_cut"] = True
                    state["seq"] += 1
                    lo[0] = lower()
            if strips >= mc.max_strips:
                return
            for k in range(i, n):
                w = modes[k].strip_w10
                if used + w > hi:
                    continue
                if used + w * (mc.max_strips - strips) < lo[0]:
                    if lo[0] > lo0:
                        state["full_cut"] = True   # cat bo CT nhieu phe hon nguong top `limit`
                    break
                nu = used + w
                a = max(0, -(-(lo[0] - nu) // unit))
                b = (hi - nu) // unit
                if not (reach[k][K - strips - 1] >> a) & ((1 << (b - a + 1)) - 1):
                    if lo[0] > lo0:
                        state["full_cut"] = True
                    continue
                counts[k] += 1
                dfs(k, used + w, strips + 1)
                counts[k] -= 1

        dfs(0, 0, 0)
    return [e[3] for e in heap], not state["full_cut"]


# --------------------------------------------------------------------------
# Optimisation
# --------------------------------------------------------------------------
def kg_strip(width_mm: float, length_m: float, gsm: float) -> float:
    """KL 1 cuon con: rong(mm) x dai(m) x gsm(g/m2) -> kg."""
    return width_mm / 1000 * length_m * gsm / 1000


def master_of(p: Problem, pt: Pattern) -> Master:
    return p.masters[pt.master]


def product_width(p: Problem, pt: Pattern) -> float:
    """Tong kho thanh pham cua 1 bo (khong tinh phe xen lai)."""
    return sum(p.items[md.item].width_mm * md.mult * c for md, c in pt.counts)


def strip_extra_cost(p: Problem, md: Mode, L: float, gsm: Optional[float] = None) -> float:
    """Phat cho 1 cuon con tren may (ngoai phe): xen lai, noi cuon, lech chieu dai."""
    wt = p.weights
    it = p.items[md.item]
    kg = kg_strip(it.width_mm * md.mult, L, gsm or p.gsm)
    cost = 0.0
    if md.mult > 1:
        cost += wt.reslit_per_roll
    if md.join == 2:
        cost += wt.join_per_roll * md.mult / 2
    cost += wt.length_dev * abs(L * md.join - it.roll_length_m) / it.roll_length_m * kg
    return cost


def stock_weight(p: Problem, it: Item) -> float:
    return min(p.weights.stock_base + p.weights.stock_step * (it.priority - 1), 0.95)


def over_weights(p: Problem, it: Item) -> tuple:
    """(phat 1 kg du trong dung sai, phat 1 kg du vuot dung sai) cua 1 don."""
    wt = p.weights
    if it.exact:
        return wt.over_exact, max(wt.over_exact, wt.over_beyond_tol)
    return wt.over_in_tol, wt.over_beyond_tol


def pattern_cost(p: Problem, pt: Pattern) -> float:
    """Chi phi 1 bo (chua tinh du/stock): phe + phat tung cuon con."""
    m = master_of(p, pt)
    cost = p.weights.trim * kg_strip(m.width_mm - product_width(p, pt), pt.length_m, m.gsm)
    return cost + sum(strip_extra_cost(p, md, pt.length_m, m.gsm) * c for md, c in pt.counts)


def column_generation(p: Problem, verbose: bool = False, max_iter: int = 300,
                      per_length: int = 5) -> list:
    """Sinh cong thuc 'co loi' thay vi liet ke het (khi so cong thuc qua lon).

    Gilmore-Gomory: giai LP noi long -> lay gia tri doi ngau (dual) cua tung don
    -> bai knapsack (DP numpy) tim cong thuc co reduced cost am -> them vao -> lap.
    """
    import numpy as np
    from ortools.linear_solver import pywraplp

    mc, wt = p.machine, p.weights
    lengths = candidate_lengths(p)
    modes_by_L = {L: modes_for_length(p, L) for L in lengths}
    keys = [(mi, L) for mi in range(len(p.masters)) for L in lengths]
    pool: dict = {}

    def price(mi, L, duals, mu, topk, thresh):
        modes = modes_by_L[L]
        if not modes:
            return []
        W, g = p.masters[mi].width_mm, p.masters[mi].gsm
        W10 = int(round(W * SCALE_W))
        tmin, tmax = p.trim_range(mi)
        lo10 = W10 - int(round(tmax * SCALE_W))
        hi10 = W10 - int(round(tmin * SCALE_W))
        if hi10 <= 0:
            return []
        unit = math.gcd(*[md.strip_w10 for md in modes])
        N = hi10 // unit
        K = mc.max_strips
        v = np.empty(len(modes))
        for m, md in enumerate(modes):
            it = p.items[md.item]
            kg = kg_strip(it.width_mm * md.mult, L, g)
            val = -wt.trim * kg + strip_extra_cost(p, md, L, g) - duals[md.item] * kg
            if it.is_stock:
                val += stock_weight(p, it) * kg
            v[m] = val
        const = wt.trim * kg_strip(W, L, g) - mu[mi] * kg_strip(W, L, g)
        INF = 1e18
        dp = np.full((K + 1, N + 1), INF)
        arg = np.full((K + 1, N + 1), -1, dtype=np.int32)
        dp[0, 0] = 0.0
        for k in range(1, K + 1):
            for m, md in enumerate(modes):
                s = md.strip_w10 // unit
                if s > N:
                    continue
                cand = dp[k - 1, :N + 1 - s] + v[m]
                better = cand < dp[k, s:]
                dp[k, s:] = np.where(better, cand, dp[k, s:])
                arg[k, s:] = np.where(better, m, arg[k, s:])
        x_lo = -(-lo10 // unit)
        cells = [(dp[k, x] + const, k, x) for k in range(1, K + 1)
                 for x in range(x_lo, N + 1) if dp[k, x] < INF / 2]
        cells.sort()
        out = []
        for rc, k, x in cells:
            if rc >= thresh or len(out) >= topk:
                break
            cnt = [0] * len(modes)
            kk, xx = k, x
            while kk > 0:
                m = arg[kk, xx]
                cnt[m] += 1
                xx -= modes[m].strip_w10 // unit
                kk -= 1
            pt = Pattern(L, tuple((modes[i], c) for i, c in enumerate(cnt) if c),
                         sum(modes[i].strip_w10 * c for i, c in enumerate(cnt)), mi)
            key = (mi, L, pt.counts)
            if key not in pool and all(key != (o.master, o.length_m, o.counts) for o in out):
                out.append(pt)
        return out

    def solve_lp():
        lp = pywraplp.Solver.CreateSolver("GLOP")
        cols = list(pool.values())
        k = [lp.NumVar(0, lp.infinity(), "") for _ in cols]
        obj = lp.Objective()
        rows, extra = {}, []
        for idx, it in enumerate(p.items):
            if it.is_stock:
                rows[idx] = lp.Constraint(-lp.infinity(), it.max_kg)
            else:
                r = lp.Constraint(it.kg, it.kg)
                oi = lp.NumVar(0, it.kg * it.over_tol_pct / 100, "")
                oo = lp.NumVar(0, lp.infinity(), "")
                sh = lp.NumVar(0, lp.infinity(), "")  # bien gia: dam bao LP luon kha thi
                r.SetCoefficient(oi, -1)
                r.SetCoefficient(oo, -1)
                r.SetCoefficient(sh, 1)
                w_in, w_out = over_weights(p, it)
                obj.SetCoefficient(oi, w_in)
                obj.SetCoefficient(oo, w_out)
                obj.SetCoefficient(sh, 1000.0)
                rows[idx] = r
                extra += [oi, oo, sh]
        avail = {mi: lp.Constraint(-lp.infinity(), m.available_kg)
                 for mi, m in enumerate(p.masters) if m.available_kg}
        for var, pt in zip(k, cols):
            c = pattern_cost(p, pt)
            m = master_of(p, pt)
            for md, n in pt.counts:
                it = p.items[md.item]
                kg = kg_strip(it.width_mm * md.mult, pt.length_m, m.gsm) * n
                rows[md.item].SetCoefficient(var, rows[md.item].GetCoefficient(var) + kg)
                if it.is_stock:
                    c += stock_weight(p, it) * kg
            obj.SetCoefficient(var, c)
            if pt.master in avail:
                avail[pt.master].SetCoefficient(var, kg_strip(m.width_mm, pt.length_m, m.gsm))
        obj.SetMinimization()
        lp.Solve()
        duals = {i: r.dual_value() for i, r in rows.items()}
        mu = {mi: avail[mi].dual_value() if mi in avail else 0.0 for mi in range(len(p.masters))}
        return obj.Value(), duals, mu

    for it_no in range(max_iter):
        val, duals, mu = solve_lp()
        added = 0
        for mi, L in keys:
            for pt in price(mi, L, duals, mu, per_length, -1e-6):
                pool[(mi, L, pt.counts)] = pt
                added += 1
        if verbose:
            print(f"  CG iter {it_no}: LP = {val:.2f}, +{added} cong thuc (tong {len(pool)})")
        if not added:
            break
    # LP hoi tu = can duoi (lower bound) cua bai toan that -> dung de bao % cach toi uu
    lp_bound = val if not added else None
    # Them cong thuc 'gan toi uu' de bai nguyen co nhieu lua chon hon
    val, duals, mu = solve_lp()
    for mi, L in keys:
        for pt in price(mi, L, duals, mu, 20, 0.05 * max(1.0, abs(val)) / max(1, len(keys))):
            pool[(mi, L, pt.counts)] = pt
    return list(pool.values()), lp_bound


def plan_cost(p: Problem, runs: list) -> float:
    """Muc tieu cua 1 ke hoach, tinh doc lap voi solver."""
    wt = p.weights
    cost = 0.0
    prod = [0.0] * len(p.items)
    for pt, sets in runs:
        cost += pattern_cost(p, pt) * sets + wt.setup_per_pattern
        for md, c in pt.counts:
            it = p.items[md.item]
            prod[md.item] += kg_strip(it.width_mm * md.mult, pt.length_m, master_of(p, pt).gsm) * c * sets
    for it, kg in zip(p.items, prod):
        if it.is_stock:
            cost += stock_weight(p, it) * kg
        else:
            over, tol = max(0.0, kg - it.kg), it.kg * it.over_tol_pct / 100
            w_in, w_out = over_weights(p, it)
            cost += w_in * min(over, tol) + w_out * max(0.0, over - tol)
    return cost


@dataclass
class Solution:
    status: str
    objective: float
    runs: list           # [(Pattern, sets)]
    problem: Problem
    n_patterns: int
    method: str = ""
    bound: Optional[float] = None   # can duoi cua muc tieu -> gap

    @property
    def gap_pct(self) -> Optional[float]:
        if self.bound is None or self.objective <= 1e-9:
            return None
        return max(0.0, (self.objective - self.bound) / self.objective * 100)


def solve(p: Problem, verbose: bool = False) -> Solution:
    """Chon cong thuc (liet ke het neu it, column generation neu nhieu) roi giai bai nguyen."""
    from ortools.linear_solver import pywraplp

    validate(p)
    lp_bound = None
    try:
        pats = enumerate_patterns(p, cap=p.enum_limit)
        method = f"liet ke toan bo {len(pats)} CT"
    except TooManyPatterns:
        if verbose:
            print(f"> {p.enum_limit} cong thuc -> chuyen sang column generation")
        pats, lp_bound = column_generation(p, verbose=verbose)
        method = "column generation"
    if not pats:
        raise RuntimeError("Khong co cong thuc nao thoa bien phe / so dao. Noi rong max_trim_mm?")

    wt = p.weights
    for idx, it in enumerate(p.items):
        if not it.is_stock and not any(md.item == idx for pt in pats for md, _ in pt.counts):
            trims = ", ".join(f"{m.label} {m.width_mm:g} mm: [{p.trim_range(mi)[0]:g}, {p.trim_range(mi)[1]:g}]"
                              for mi, m in enumerate(p.masters))
            raise RuntimeError(
                f"{it.key}: khong co cong thuc nao chua kho nay. Kho nay khong ghep duoc voi "
                f"kho khac cung so met (+/-{p.machine.length_tol_pct:g}%), va tu no khong lap du "
                f"kho NL trong bien phe ({trims} mm). "
                "Thu: them stock cung so met, noi rong max_trim, bat multiples/allow_join.")

    mip = pywraplp.Solver.CreateSolver(p.mip_solver)
    if mip is None:
        raise RuntimeError(f"Solver {p.mip_solver} khong co. Dung SCIP / HIGHS / CBC.")
    mip.SetTimeLimit(int(p.time_limit_s * 1000))
    if verbose:
        mip.EnableOutput()

    def ub(pt: Pattern) -> int:
        """Gioi han tren so bo cua 1 CT (de rang buoc setup y_p chat hon)."""
        best = 1
        for md, c in pt.counts:
            it = p.items[md.item]
            per_set = kg_strip(it.width_mm * md.mult, pt.length_m, master_of(p, pt).gsm) * c
            best = max(best, math.ceil(it.cap_kg / per_set) + 1)
        return best

    k = [mip.IntVar(0, ub(pt), f"k{i}") for i, pt in enumerate(pats)]
    y = [mip.BoolVar(f"y{i}") for i in range(len(pats))]
    obj = []
    for i, pt in enumerate(pats):
        mip.Add(k[i] <= ub(pt) * y[i])
        obj += [pattern_cost(p, pt) * k[i], wt.setup_per_pattern * y[i]]

    for idx, it in enumerate(p.items):
        terms, join_terms = [], []
        for i, pt in enumerate(pats):
            for md, c in pt.counts:
                if md.item == idx:
                    terms.append(kg_strip(it.width_mm * md.mult, pt.length_m, master_of(p, pt).gsm) * c * k[i])
                    if md.join == 2:
                        join_terms.append(c * md.mult * k[i])
        prod = mip.Sum(terms) if terms else 0
        if it.is_stock:
            if terms:
                mip.Add(prod <= it.max_kg)
                obj.append(stock_weight(p, it) * prod)
        else:
            over_in = mip.NumVar(0, it.kg * it.over_tol_pct / 100, f"oi{idx}")
            over_out = mip.NumVar(0, mip.infinity(), f"oo{idx}")
            mip.Add(prod == it.kg + over_in + over_out)
            w_in, w_out = over_weights(p, it)
            obj += [w_in * over_in, w_out * over_out]
        if join_terms:  # so doan noi phai chan (2 doan = 1 cuon)
            z = mip.IntVar(0, mip.infinity(), f"z{idx}")
            mip.Add(mip.Sum(join_terms) == 2 * z)

    for mi, m in enumerate(p.masters):  # KL tung cuon NVL cap (0/None = khong gioi han)
        if m.available_kg:
            mip.Add(mip.Sum(kg_strip(m.width_mm, pt.length_m, m.gsm) * k[i]
                            for i, pt in enumerate(pats) if pt.master == mi) <= m.available_kg)

    mip.Minimize(mip.Sum(obj))
    st = mip.Solve()
    names = {pywraplp.Solver.OPTIMAL: "OPTIMAL", pywraplp.Solver.FEASIBLE: "FEASIBLE",
             pywraplp.Solver.INFEASIBLE: "INFEASIBLE"}
    name = names.get(st, f"STATUS_{st}")
    if st not in (pywraplp.Solver.OPTIMAL, pywraplp.Solver.FEASIBLE):
        raise RuntimeError(f"Khong tim duoc loi giai: {name}. Kiem tra KL NVL cap / bien phe.")
    runs = [(pt, int(round(k[i].solution_value()))) for i, pt in enumerate(pats)
            if k[i].solution_value() > 0.5]
    runs.sort(key=lambda r: -r[1])
    # Tinh lai muc tieu tu ke hoach thuc (solver co the bat y_p=1 thua khi het gio)
    objective = plan_cost(p, runs)
    if lp_bound is not None:
        bound = lp_bound                       # CG: can duoi toan cuc
    elif st == pywraplp.Solver.OPTIMAL:
        bound = objective                      # liet ke het + OPTIMAL = toi uu that
    else:
        bound = mip.Objective().BestBound()
    if bound is not None and bound < 0:
        bound = None
    return Solution(name, objective, runs, p, len(pats), method, bound)


# --------------------------------------------------------------------------
# Reporting
# --------------------------------------------------------------------------
def formula_str(p: Problem, pt: Pattern) -> str:
    parts = []
    for md, c in pt.counts:
        w = p.items[md.item].width_mm
        tag = f"{w:g}" if md.mult == 1 else f"({md.mult}x{w:g})"
        parts.append(f"{c}x{tag}")
    return " + ".join(parts)


def plan_rows(sol: Solution) -> list:
    """1 dong / cuon con trong moi lenh - giong block A:I cu tren sheet FastCal."""
    p = sol.problem
    rows = []
    for n, (pt, sets) in enumerate(sol.runs, start=1):
        m = master_of(p, pt)
        trim_w = m.width_mm - product_width(p, pt)
        master_kg = kg_strip(m.width_mm, pt.length_m, m.gsm) * sets
        for md, c in pt.counts:
            it = p.items[md.item]
            strip_w = it.width_mm * md.mult
            rolls_machine = c * sets
            rolls_final = rolls_machine * md.mult / md.join
            rows.append({
                "Lenh": n, "Cong thuc": formula_str(p, pt), "Cuon NL": m.label,
                "Kho NL": m.width_mm, "So m/bo": pt.length_m, "So bo": sets,
                "SL NL can (kg)": round(master_kg, 1),
                "Kho cuon con": strip_w, "SD": c,
                "Loai": {1: "don", 2: "doi", 3: "ba", 4: "tu"}.get(md.mult, f"x{md.mult}"),
                "Cho": it.key, "Don": ",".join(it.order_ids) if it.order_ids else "STOCK",
                "SC tren may": rolls_machine,
                "SC thanh pham": rolls_final,
                "Noi cuon": "co" if md.join == 2 else "",
                "So kg": round(kg_strip(strip_w, pt.length_m, m.gsm) * rolls_machine, 1),
                "Bien phe (mm)": round(trim_w, 1),
                "Phe (kg)": round(kg_strip(trim_w, pt.length_m, m.gsm) * sets, 1),
            })
    return rows


def master_rows(sol: Solution) -> list:
    """Muc su dung tung cuon NVL: cap bao nhieu, dung bao nhieu, con lai."""
    p = sol.problem
    used = [0.0] * len(p.masters)
    meters = [0.0] * len(p.masters)
    for pt, sets in sol.runs:
        m = master_of(p, pt)
        used[pt.master] += kg_strip(m.width_mm, pt.length_m, m.gsm) * sets
        meters[pt.master] += pt.length_m * sets
    out = []
    for mi, m in enumerate(p.masters):
        tmin, tmax = p.trim_range(mi)
        out.append({"Cuon NL": m.label, "Kho NL": m.width_mm, "Dinh luong": m.gsm,
                    "Cap (kg)": round(m.available_kg, 1) if m.available_kg else "khong gioi han",
                    "So cuon": m.rolls, "So m/cuon": round(m.length_m) if m.length_m else "",
                    "Bien phe (mm)": f"{tmin:g}-{tmax:g}",
                    "Dung (kg)": round(used[mi], 1), "So m dung": round(meters[mi]),
                    "Con lai (kg)": round(m.available_kg - used[mi], 1) if m.available_kg else ""})
    return out


def summary_rows(sol: Solution) -> list:
    p = sol.problem
    prod = {i: 0.0 for i in range(len(p.items))}
    rolls = {i: 0.0 for i in range(len(p.items))}
    for pt, sets in sol.runs:
        for md, c in pt.counts:
            it = p.items[md.item]
            prod[md.item] += kg_strip(it.width_mm * md.mult, pt.length_m, master_of(p, pt).gsm) * c * sets
            rolls[md.item] += c * sets * md.mult / md.join
    out = []
    for i, it in enumerate(p.items):
        if it.is_stock:
            out.append({"Item": it.key, "Loai": f"stock #{it.priority}", "Can (kg)": 0,
                        "Ra (kg)": round(prod[i], 1), "Du (kg)": round(prod[i], 1),
                        "% du": "", "Dung sai %": f"max {it.max_kg:g} kg",
                        "So cuon": round(rolls[i], 1), "Trang thai": "STOCK" if prod[i] else "-"})
        else:
            over = prod[i] - it.kg
            pct = over / it.kg * 100
            status = "OK" if pct <= it.over_tol_pct + 1e-6 else "VUOT DUNG SAI"
            if over < -0.05:
                status = "THIEU"
            if it.exact and status == "OK" and over > 0.05:
                status = "DU (VUA DU)"
            out.append({"Item": it.key, "Loai": ",".join(it.order_ids), "Can (kg)": round(it.kg, 1),
                        "Ra (kg)": round(prod[i], 1), "Du (kg)": round(over, 1),
                        "% du": round(pct, 1), "Dung sai %": round(it.over_tol_pct, 1),
                        "So cuon": round(rolls[i], 1), "Trang thai": status})
    return out


def totals(sol: Solution) -> dict:
    p = sol.problem
    master = trim = 0.0
    for pt, s in sol.runs:
        m = master_of(p, pt)
        master += kg_strip(m.width_mm, pt.length_m, m.gsm) * s
        trim += kg_strip(m.width_mm - product_width(p, pt), pt.length_m, m.gsm) * s
    summ = summary_rows(sol)
    over = sum(r["Du (kg)"] for r in summ if not r["Loai"].startswith("stock"))
    stock = sum(r["Ra (kg)"] for r in summ if r["Loai"].startswith("stock"))
    need = sum(it.kg for it in p.items if not it.is_stock)
    return {"NVL (kg)": round(master, 1), "Nhu cau (kg)": round(need, 1),
            "Phe (kg)": round(trim, 1), "% phe": round(trim / master * 100, 2) if master else 0,
            "Du don (kg)": round(over, 1), "Stock (kg)": round(stock, 1),
            "So cong thuc": len(sol.runs), "Tong so bo": sum(s for _, s in sol.runs)}


def _table(rows: list, cols: list) -> str:
    if not rows:
        return "(trong)"
    widths = [max(len(str(c)), *(len(str(r[c])) for r in rows)) for c in cols]
    line = " | ".join(str(c).ljust(w) for c, w in zip(cols, widths))
    sep = "-+-".join("-" * w for w in widths)
    body = "\n".join(" | ".join(str(r[c]).ljust(w) for c, w in zip(cols, widths)) for r in rows)
    return f"{line}\n{sep}\n{body}"


def report(sol: Solution) -> str:
    p = sol.problem
    t = totals(sol)
    if sol.status == "OPTIMAL" and sol.method.startswith("liet ke"):
        gap_txt = "TOI UU (da chung minh)"
    else:
        gap_txt = f"tot nhat tim duoc trong {p.time_limit_s:g}s - chua chung minh toi uu"
    nl = ", ".join(f"{m.label} {m.width_mm:g} mm / gsm {m.gsm:g}" for m in p.masters)
    out = [f"== KET QUA ({sol.status} | {sol.method} | {sol.n_patterns} CT da xet | {gap_txt}) ==",
           f"{nl} | " + " | ".join(f"{k}: {v}" for k, v in t.items()), ""]
    if len(p.masters) > 1:
        out += ["== CUON NVL ==", _table(master_rows(sol), list(master_rows(sol)[0].keys())), ""]
    for n, (pt, sets) in enumerate(sol.runs, start=1):
        rows = [r for r in plan_rows(sol) if r["Lenh"] == n]
        r0 = rows[0]
        nl_tag = f"[{r0['Cuon NL']} {r0['Kho NL']:g} mm] " if len(p.masters) > 1 else ""
        out.append(f"[Lenh {n}] {nl_tag}{r0['Cong thuc']}  | {sets} bo x {pt.length_m} m | "
                   f"NL {r0['SL NL can (kg)']} kg | bien phe {r0['Bien phe (mm)']} mm = {r0['Phe (kg)']} kg")
        out.append(_table(rows, ["Kho cuon con", "SD", "Loai", "Cho", "SC tren may",
                                 "SC thanh pham", "Noi cuon", "So kg"]))
        out.append("")
    out.append("== DAP UNG THEO DON ==")
    out.append(_table(summary_rows(sol), ["Item", "Loai", "Can (kg)", "Ra (kg)", "Du (kg)",
                                          "% du", "Dung sai %", "So cuon", "Trang thai"]))
    return "\n".join(out)


def write_csv(rows: list, path: str) -> None:
    if not rows:
        return
    with open(path, "w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig: Excel doc dung
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)


def pattern_row(p: Problem, pt: Pattern) -> dict:
    m = master_of(p, pt)
    product_w = product_width(p, pt)
    trim = m.width_mm - product_w
    return {"Cuon NL": m.label, "Kho NL": m.width_mm, "Ten CT": formula_str(p, pt),
            "So m": pt.length_m, "Kho SD": round(product_w, 1), "PL (mm)": round(trim, 1),
            "%PL": round(trim / m.width_mm * 100, 2)}


def sort_pattern_rows(rows: list) -> list:
    return sorted(rows, key=lambda r: (r["PL (mm)"] / r["Kho NL"], r["So m"], r["Cuon NL"], r["Ten CT"]))


def pattern_rows(p: Problem, limit: int = LIST_LIMIT) -> list:
    """Moi cong thuc kha thi (khac ten) 1 dong, bien phe tang dan (giong bang VBA cu).
    Qua `limit` CT thi chi giu cac CT it phe nhat (xem pattern_listing)."""
    pats, _ = pattern_listing(p, limit)
    return sort_pattern_rows([pattern_row(p, pt) for pt in pats])


def list_patterns(p: Problem, top: int = 30) -> str:
    """Che do giong VBA cu: liet ke cong thuc theo % bien phe tang dan."""
    rows = pattern_rows(p)
    return f"{len(rows)} cong thuc kha thi, top {top}:\n" + _table(rows[:top], list(rows[0].keys()))
