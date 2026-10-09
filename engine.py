# -*- coding: utf-8 -*-
"""确定性排程内核封装 —— 排程智能体第 1 层。

设计原则:
- 不改动原排程脚本 schedule_heat_treatment_strict_compact.py,直接以函数级 API 调用;
- 新增 validate(硬约束校验)与 kpis(排程质量指标);
- 智能体只能通过这里暴露的动作改变排程,硬约束永远由原内核裁决。
"""
from __future__ import annotations

import importlib
import json
import re
import sys
from pathlib import Path

PROJECT = Path(r"C:\Users\mblon\Desktop\project")
CORE_MODULE = "schedule_heat_treatment_strict_compact"

if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))
core = importlib.import_module(CORE_MODULE)

TIME_KEYS = ("_front_start", "_front_last_in", "_front_end",
             "_temper_start", "_temper_last_in", "_temper_end")

# 双炉中间等待(前炉出炉→回火进炉)超过该分钟数,记入"堵炉风险"提示(可调)
TRANSFER_WAIT_LIMIT_MIN = 8 * 60
TRANSFER_RISK_TOP = 5


def transfer_wait_of(x: dict, front_steps=None) -> tuple[float, float] | None:
    """双炉订单"前炉出炉→回火进炉"的中间等待,返回 (分钟, 步)。

    同一根管子前炉在炉时间 = 前炉步数 × 步进周期;若回火进炉时间晚于
    前炉出炉时间,差值为管子在两炉之间的等待(输送/缓冲段占用)。
    """
    if x.get("_front_start") is None or x.get("_temper_start") is None:
        return None
    step = x["_speed"] * core.STEP_UNIT_MIN
    hold = (front_steps if front_steps is not None else core.FRONT_STEPS) * step
    wait_min = x["_temper_start"] - (x["_front_start"] + hold)
    return wait_min, (wait_min / step if step else 0.0)

# ---------------------------------------------------------------- 原因可读化
CAT_FILLS = {
    "连续生产": "C6EFCE",   # 绿
    "首单": "DDEBF7",       # 浅蓝
    "必要换规": "FFEB9C",   # 黄
    "大批量恢复": "E4DFEC",  # 紫
    "补前炉": "D9E1F2",     # 蓝灰
    "只回火衔接": "FCE4D6",  # 橙
    "人工锁定": "D9F2F2",   # 青
    "暂缓/剔除": "FFC7CE",  # 红
}

CAT_DESC = {
    "连续生产": "与前一单无缝连续,空格≤20格",
    "首单": "序列启动首单",
    "必要换规": "跨温区接续,正火/回火分别等待N步(1格=1步进周期),原因列含换规规则编号",
    "大批量恢复": "同温区累计产量≥50吨,批量恢复排产",
    "补前炉": "只回火期间补排一单仅前炉订单",
    "只回火衔接": "仅回火订单的衔接段",
    "人工锁定": "人工锁定顺序,请关注校验结果",
    "暂缓/剔除": "暂缓/剔除主序列",
}

KEEP_DETAIL = {"补前炉", "只回火衔接", "人工锁定", "暂缓/剔除"}


def _cat_of(s: str) -> str:
    for key, cat in (("人工锁定", "人工锁定"), ("补前炉", "补前炉"), ("只回火", "只回火衔接"),
                     ("同温区大批量", "大批量恢复"), ("必要换规", "必要换规"), ("首单", "首单")):
        if key in s:
            return cat
    return "连续生产"


def _rule_ref(detail: str) -> str:
    """从换规说明中提取规则编号,如 '第4条，ΔC=10℃' -> '第4条'。"""
    nums = re.findall(r"第(\d+)条", detail)
    return "第" + "/".join(dict.fromkeys(nums)) + "条" if nums else ""


def format_main_reason(x: dict, front_blank=None, temper_blank=None) -> tuple[str, str]:
    """把主序列订单原因改写成【分类】+关键数字的短格式。

    front_blank/temper_blank:core.blank_count 的返回值 (数值, 说明),
    供换规行显示"正火等X步/回火等Y步";规则编号与中间等待等细节由调用方写入备注。
    """
    s = core.text(x.get("_start_reason", ""))
    cat = _cat_of(s)
    fb_val = front_blank[0] if front_blank else None
    tb_val = temper_blank[0] if temper_blank else None
    fb_ref = core.text(front_blank[1]) if front_blank else ""
    tb_ref = core.text(temper_blank[1]) if temper_blank else ""
    parts = []
    if cat == "首单":
        pass  # 只写【首单】,哪座炉先开看时间列即可
    elif cat == "补前炉":
        m1 = re.search(r"前侧(\d+(?:\.\d+)?)格", s)
        m2 = re.search(r"后侧(\d+(?:\.\d+)?)格", s)
        if m1 and m2:
            parts.append(f"前侧{m1.group(1)}格/后侧{m2.group(1)}格")
        elif "未找到" in s:
            parts.append("无≤20格搭配")
    elif cat == "大批量恢复":
        m = re.search(r"(\d+(?:\.\d+)?)吨", s)
        if m:
            parts.append(f"{float(m.group(1)):g}吨")
    if cat == "必要换规":
        # 换规行只留核心:两炉分别等多少步(规则编号/中间等待写入备注)
        proc = core.text(x.get("_process"))
        fname = "正火" if "正火" in proc else ("淬火" if "淬火" in proc else "前炉")
        if x.get("_needs_front") and fb_val:
            parts.append(f"{fname}等{float(fb_val):g}步")
        if x.get("_needs_temper") and tb_val:
            parts.append(f"回火等{float(tb_val):g}步")
    else:
        # 统一规则:正常空格(≤20格)不显示,只有超过20格的异常才写数字
        if cat != "连续生产":
            if fb_val and float(fb_val) > core.LIMIT_BLANKS:
                parts.append(f"前炉{float(fb_val):g}格")
            if tb_val and float(tb_val) > core.LIMIT_BLANKS:
                parts.append(f"回火{float(tb_val):g}格")
    tag = f"【{cat}】"
    return tag + "/".join(parts) if parts else tag, cat


def format_defer_reason(x: dict) -> tuple[str, str]:
    """暂缓订单原因压缩为【暂缓】+一句短因(兼容全角冒号)。"""
    s = core.text(x.get("_defer_reason", "")).replace(":", ":")
    for prefix in ("不能接入当前集中生产路径:", "暂缓:"):
        if s.startswith(prefix):
            s = s[len(prefix):]
    m = re.search(r"小批温区，(\d+)单/([\d.]+)吨", s)
    if m:
        short = f"小批温区{m.group(1)}单/{float(m.group(2)):g}吨,前后超20格"
    else:
        short = re.split(r"[;:]", s)[0].strip()[:18] or "不能接续"
    return "【暂缓】" + short, "暂缓/剔除"


def _update_state(state: dict, x: dict) -> None:
    """与原脚本 build_sequence 一致的状态推进逻辑。"""
    if x["_needs_front"]:
        state["front_prev"], state["front"] = state.get("front"), x
    if x["_needs_temper"]:
        state["temper"] = x
        state["front_inserted"] = 0
    elif x["_needs_front"]:
        state["front_inserted"] = state.get("front_inserted", 0) + 1


def build_sequence(orders, rules, specials, pinned=(), boosted=()):
    """与原脚本一致的贪心构造,支持"锁定前缀"与"急催加权"。"""
    pinned = [core.norm(str(c)) for c in pinned]
    boosted = {core.norm(str(c)) for c in boosted}
    pool = core.seed_order_with_ortools(orders)
    pool.sort(key=lambda x: (x["_temper"] if x["_needs_temper"] else 10_000,
                             x["_front"] or 10_000, x["_wall"], x["_speed"]))
    main, deferred = [], []
    state = {"front": None, "front_prev": None, "temper": None, "front_inserted": 0}

    # 1) 锁定前缀:按给定顺序放置;即使违规也放置,违规由 validate 报告
    for c in pinned:
        x = next((y for y in pool if core.norm(str(y["_contract"])) == c), None)
        if x is None:
            continue
        ok, reason = core.edge_ok(state, x, rules, specials)
        x["_start_reason"] = "人工锁定顺序,约束满足" if ok else f"人工锁定顺序,存在违规:{reason}"
        x["_pinned"] = True
        main.append(x)
        pool.remove(x)
        _update_state(state, x)

    # 2) 其余订单按原脚本规则贪心接续
    while pool:
        candidates = []
        for x in pool:
            ok, reason = core.edge_ok(state, x, rules, specials)
            if not ok:
                continue
            fb = core.blank_count(state.get("front"), x, "前炉", rules, specials)[0] if x["_needs_front"] else 0
            tb = core.blank_count(state.get("temper"), x, "回火炉", rules, specials)[0] if x["_needs_temper"] else 0
            t = x["_temper"] if x["_needs_temper"] else (state["temper"]["_temper"] if state.get("temper") else 9999)
            front_only = 1 if (x["_needs_front"] and not x["_needs_temper"]) else 0
            boost = 0 if core.norm(str(x["_contract"])) in boosted else 1
            candidates.append((boost, front_only, t, tb, fb,
                               abs((x.get("_front") or 0) - (state.get("front") or {}).get("_front", x.get("_front") or 0)),
                               x["_wall"], x, reason))
        if not candidates:
            for x in pool:
                ok, reason = core.edge_ok(state, x, rules, specials)
                x["_defer_reason"] = f"不能接入当前集中生产路径:{reason}"
                deferred.append(x)
            break
        *_, chosen, reason = min(candidates, key=lambda v: v[:6])
        chosen["_start_reason"] = reason
        main.append(chosen)
        pool.remove(chosen)
        _update_state(state, chosen)
    return main, deferred


def shift_all_times(main, minutes):
    """整线顺延(模拟开工前检修/停机)。"""
    for x in main:
        for k in TIME_KEYS:
            if x.get(k) is not None:
                x[k] = x[k] + minutes


def validate(main, rules, specials, front_steps=None) -> list[dict]:
    """重放全部硬约束,返回问题清单。

    类别说明:
    - "违规":原引擎 edge_ok 明确拒绝的情况(通常来自恢复大批量/补前炉等修复阶段);
    - "必要换规":相邻空格>20格,但属于引擎允许的跨温区接续,仅提示关注;
    - "堵炉风险":双炉订单前炉出炉后等待过久才进回火炉(输送段滞留),仅提示关注。
    """
    issues, seen = [], set()
    risks = []
    state = {"front": None, "front_prev": None, "temper": None, "front_inserted": 0}
    for i, x in enumerate(main):
        ok, reason = core.edge_ok(state, x, rules, specials)
        if not ok:
            issues.append({"位置": i + 1, "订单编号": x["_contract"], "类别": "违规", "问题": reason})
        if i > 0:
            mx, detail = core.edge_blank_summary(main[i - 1], x, rules, specials)
            if mx > core.LIMIT_BLANKS and ok:
                issues.append({"位置": i + 1, "订单编号": x["_contract"], "类别": "必要换规",
                               "问题": f"相邻空格{mx:g}格>20,属引擎允许的跨温区接续:{detail}"})
        w = transfer_wait_of(x, front_steps)
        if w and w[0] > TRANSFER_WAIT_LIMIT_MIN:
            risks.append((w, i + 1, x))
        _update_state(state, x)
    for w, i, x in sorted(risks, key=lambda t: -t[0][0])[:TRANSFER_RISK_TOP]:
        issues.append({"位置": i, "订单编号": x["_contract"], "类别": "堵炉风险",
                       "问题": f"前炉出炉后约等{float(w[0]):g}分钟({float(w[1]):g}步)才进回火炉,输送段滞留过长"})
    unique = []
    for it in issues:
        k = (it["位置"], it["问题"])
        if k not in seen:
            seen.add(k)
            unique.append(it)
    return unique


def kpis(main, deferred, rules, specials, front_steps=None) -> dict:
    """排程质量指标。"""
    blanks = []
    for i in range(1, len(main)):
        mx, _ = core.edge_blank_summary(main[i - 1], main[i], rules, specials)
        blanks.append(mx)
    issues = validate(main, rules, specials, front_steps)
    ends = [x.get(k) for x in main for k in ("_front_end", "_temper_end") if x.get(k) is not None]
    makespan = max(ends) if ends else 0.0
    waits = [w for x in main if (w := transfer_wait_of(x, front_steps)) is not None and w[0] > 0]
    return {
        "订单总数": len(main) + len(deferred),
        "正常生产单数": len(main),
        "暂缓/剔除单数": len(deferred),
        "正常生产吨位": round(sum(float(x.get("_plan_tons", 0) or 0) for x in main), 1),
        "暂缓吨位": round(sum(float(x.get("_plan_tons", 0) or 0) for x in deferred), 1),
        "相邻空格总数": round(sum(blanks), 1),
        "平均相邻空格": round(sum(blanks) / len(blanks), 2) if blanks else 0.0,
        "超20格相邻边数": len([b for b in blanks if b > core.LIMIT_BLANKS]),
        "最大相邻空格": max(blanks) if blanks else 0.0,
        "硬约束违规数": sum(1 for i in issues if i.get("类别") == "违规"),
        "必要换规边数": sum(1 for i in issues if i.get("类别") == "必要换规"),
        "温度回摆次数": sum(1 for i in issues if "回摆" in i["问题"] and i.get("类别") == "违规"),
        "双炉中间最大等待": core.clock(max(w[0] for w in waits)) if waits else core.clock(0),
        "双炉中间平均等待分钟": round(sum(w[0] for w in waits) / len(waits), 1) if waits else 0.0,
        "堵炉风险单数": sum(1 for w in waits if w[0] > TRANSFER_WAIT_LIMIT_MIN),
        "急催订单数": sum(1 for x in main if core.text(x.get("_urgent", ""))),
        "急催暂缓数": sum(1 for x in deferred if core.text(x.get("_urgent", ""))),
        "完工时间": core.clock(makespan),
        "完工分钟": round(makespan, 1),
    }


def diff_kpis(base: dict, new: dict) -> dict:
    """数值指标差异(基准 -> 当前)。"""
    out = {}
    for k, v in base.items():
        w = new.get(k)
        if isinstance(v, (int, float)) and isinstance(w, (int, float)) and v != w:
            out[k] = {"基准": v, "当前": w, "变化": round(w - v, 2)}
    return out


def explain(main, deferred, contract) -> dict:
    """解释订单的排产原因、时间与工模具。"""
    n = core.norm(str(contract))
    for x in main:
        if core.norm(str(x["_contract"])) == n:
            return {
                "订单编号": x["_contract"], "阶段": "正常生产", "序号": x.get("_seq"),
                "主体厂": core.text(x.get("_factory")), "品种": core.text(x.get("_variety")),
                "牌号": core.text(x.get("_brand")), "钢级": core.text(x.get("_steel")),
                "外径": x.get("_outer"), "壁厚": x.get("_wall"), "数量": x.get("_qty"),
                "计划产量": x.get("_plan_tons"), "热处理方式": core.text(x.get("_process")),
                "前炉温度": x.get("_front"), "回火温度": x.get("_temper"),
                "步进周期": x.get("_speed"), "布料方式": core.text(x.get("_loading")),
                "本单开始原因": x.get("_start_reason"),
                "前炉首支进炉": core.clock(x.get("_front_start")),
                "回火首支进炉": core.clock(x.get("_temper_start")),
                "喷嘴规格": core.text(x.get("_nozzle")), "除鳞环/挡水板规格": core.text(x.get("_descale")),
                "急催": core.text(x.get("_urgent")), "备注": core.text(x.get("_note")),
            }
    for x in deferred:
        if core.norm(str(x["_contract"])) == n:
            return {"订单编号": x["_contract"], "阶段": "暂缓/剔除主序列",
                    "原因": core.text(x.get("_defer_reason", "不能接入当前集中生产路径")),
                    "前炉温度": x.get("_front"), "回火温度": x.get("_temper"),
                    "壁厚": x.get("_wall"), "数量": x.get("_qty"), "计划产量": x.get("_plan_tons"),
                    "急催": core.text(x.get("_urgent")), "备注": core.text(x.get("_note"))}
    return {"error": f"未找到订单:{contract}"}


def _style_overview(path: Path, main, deferred) -> None:
    """导出后美化:新增带颜色的"排程状态"列 + 图例 sheet。"""
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    cat_map = {}
    for x in main:
        key = (core.norm(str(x["_contract"])), str(x.get("_seq") or ""))
        reason = x.get("_start_reason") or ""
        cat = x.get("_reason_cat", "") or _cat_of(core.text(reason))
        tag = reason.split("】")[0] + "】" if "】" in core.text(reason) else core.text(reason)[:8]
        cat_map[key] = (tag, cat)
    for i, x in enumerate(deferred):
        key = (core.norm(str(x["_contract"])), str(len(main) + i + 1))
        reason = x.get("_defer_reason") or ""
        cat = x.get("_reason_cat", "") or _cat_of(core.text(reason))
        tag = reason.split("】")[0] + "】" if "】" in core.text(reason) else core.text(reason)[:8]
        cat_map[key] = (tag, cat)

    wb = load_workbook(path)
    ws = wb["排产总览"]
    headers = [c.value for c in ws[1]]
    try:
        ccol = headers.index("订单编号") + 1
    except ValueError:
        return

    # 1) 在"订单编号"右侧插入"排程状态"列
    ws.insert_cols(ccol + 1)
    cell = ws.cell(row=1, column=ccol + 1, value="排程状态")
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.font = Font(color="FFFFFF", bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(bottom=Side(style="thin", color="D9E2F3"))
    ws.column_dimensions[get_column_letter(ccol + 1)].width = 12

    # 2) 行级着色:排程状态列填标签,原因列按分类填色
    headers = [c.value for c in ws[1]]
    rcol = headers.index("本单开始原因") + 1
    for row in ws.iter_rows(min_row=2):
        cid = row[ccol - 1].value
        if cid is None or str(cid).strip() == "":
            continue
        hit = cat_map.get((core.norm(str(cid)), str(row[0].value or "")))
        if not hit:
            continue
        tag, cat = hit
        fill = PatternFill("solid", fgColor=CAT_FILLS.get(cat, "F2F2F2"))
        st = row[ccol]          # 新插入的排程状态列(0-based 下标 = ccol)
        st.value = tag
        st.fill = fill
        st.alignment = Alignment(horizontal="center", vertical="center")
        rc = row[rcol - 1]      # 本单开始原因列
        rc.fill = fill

    # 3) 原始订单数据 sheet 也按同样分类给原因列着色
    if "原始订单数据" in wb.sheetnames:
        ws2 = wb["原始订单数据"]
        h2 = [c.value for c in ws2[1]]
        if "本单开始原因" in h2:
            c2 = h2.index("订单编号") + 1
            r2 = h2.index("本单开始原因") + 1
            c0 = h2.index("排产序号") + 1 if "排产序号" in h2 else None
            for row in ws2.iter_rows(min_row=2):
                cid = row[c2 - 1].value
                if cid is None or str(cid).strip() == "":
                    continue
                key = (core.norm(str(cid)), str(row[c0 - 1].value or "") if c0 else "")
                hit = cat_map.get(key)
                if hit:
                    row[r2 - 1].fill = PatternFill("solid", fgColor=CAT_FILLS.get(hit[1], "F2F2F2"))

    # 4) 图例 sheet
    if "图例" in wb.sheetnames:
        del wb["图例"]
    lg = wb.create_sheet("图例")
    lg.append(["排程状态", "含义"])
    for cat, color in CAT_FILLS.items():
        lg.append([cat, CAT_DESC.get(cat, "")])
        lg.cell(row=lg.max_row, column=1).fill = PatternFill("solid", fgColor=color)
    for c in lg[1]:
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.font = Font(color="FFFFFF", bold=True)
    lg.column_dimensions["A"].width = 14
    lg.column_dimensions["B"].width = 46
    lg.freeze_panes = "A2"
    wb.save(path)


class Schedule:
    """一次排程结果:序列 + KPI + 违规清单 + 导出。"""

    def __init__(self, engine, main, deferred):
        self.engine = engine
        self.main, self.deferred = main, deferred
        self.kpi = kpis(main, deferred, engine.rules, engine.specials, engine.front_steps)
        self.issues = validate(main, engine.rules, engine.specials, engine.front_steps)

    def export(self, path: Path) -> Path:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        core.write_excel(path, self.main, self.deferred, self.engine.source_headers)
        _style_overview(path, self.main, self.deferred)
        return path


class Engine:
    """读取数据/规则/工模具,运行整条排程管线,支持改动重排。"""

    def __init__(self, input_dir=None, rule_dir=None, tooling_dir=None):
        self.input_dir = Path(input_dir) if input_dir else core.DEFAULT_INPUT
        self.rule_dir = Path(rule_dir) if rule_dir else core.DEFAULT_RULE
        self.tooling_dir = Path(tooling_dir) if tooling_dir else core.DEFAULT_TOOLING
        self.loaded = False

    def load(self) -> int:
        self.orders, self.source_headers = core.read_inputs(self.input_dir)
        self.rules, self.front_steps, self.temper_steps, self.specials = core.read_rules(self.rule_dir)
        core.FRONT_STEPS, core.TEMPER_STEPS = self.front_steps, self.temper_steps
        self.nozzles, self.descales = core.read_tooling(self.tooling_dir)
        self.loaded = True
        return len(self.orders)

    def run(self, overrides=None) -> tuple[list, list]:
        """按改动项重跑管线;不改动已加载的原始订单。"""
        if not self.loaded:
            self.load()
        o = overrides or {}
        orders = [dict(x) for x in self.orders]

        remove = {core.norm(str(c)) for c in o.get("remove_contracts", [])}
        if remove:
            orders = [x for x in orders if core.norm(str(x["_contract"])) not in remove]

        temper_map = {core.norm(str(c)): float(t) for c, t in (o.get("temper_map") or {}).items()}
        for x in orders:
            t = temper_map.get(core.norm(str(x["_contract"])))
            if t is not None:
                x["_temper"] = t
                x["_needs_temper"] = True
                if "回火" not in str(x.get("_process", "")):
                    x["_process"] = f"{core.text(x.get('_process'))} + 回火(智能体调整)"

        main, deferred = build_sequence(orders, self.rules, self.specials,
                                        pinned=o.get("pinned", []), boosted=o.get("boost_contracts", []))
        main, deferred = core.restore_large_platforms(main, deferred)
        main, deferred = core.insert_front_order_near_temper_only(main, deferred, self.rules, self.specials)
        main, deferred = core.defer_small_islands(main, deferred, self.rules, self.specials)
        main = core.annotate_reasons(main, self.rules, self.specials)
        for x in main:
            if x.get("_pinned"):
                x["_start_reason"] = "人工锁定顺序:" + x.get("_start_reason", "")
        main = core.schedule_times(main, self.rules, self.specials, self.front_steps, self.temper_steps)
        delay = float(o.get("start_delay_minutes") or 0)
        if delay:
            shift_all_times(main, delay)
        core.apply_tooling(main + deferred, self.nozzles, self.descales)
        # 原因改写为一眼可读的短格式,并打上分类(供 Excel 着色)
        pf = pt = None
        last_batch = None  # 大批量恢复组:同组仅首单显示吨位
        for x in main:
            fb = core.blank_count(pf, x, "前炉", self.rules, self.specials) if x["_needs_front"] else (0, "无")
            tb = core.blank_count(pt, x, "回火炉", self.rules, self.specials) if x["_needs_temper"] else (0, "无")
            w = transfer_wait_of(x, self.front_steps)
            orig = core.text(x.get("_start_reason"))
            text, cat = format_main_reason(x, fb, tb)
            if x.get("_platform_tons"):
                bk = (x.get("_platform_tons"), core.platform_key(x))
                if (cat == "大批量恢复" and bk == last_batch
                        and float(fb[0]) <= core.LIMIT_BLANKS and float(tb[0]) <= core.LIMIT_BLANKS):
                    cat, text = "连续生产", "【连续生产】"
                last_batch = bk
            else:
                last_batch = None
            x["_start_reason"] = text
            x["_reason_cat"] = cat
            if cat == "必要换规":
                det = []
                if x["_needs_front"] and _rule_ref(core.text(fb[1])):
                    det.append(f"前炉{_rule_ref(core.text(fb[1]))}")
                if x["_needs_temper"] and _rule_ref(core.text(tb[1])):
                    det.append(f"回火{_rule_ref(core.text(tb[1]))}")
                if w and w[1] > 0:
                    det.append(f"中间等{float(w[1]):.1f}步")
                if det:
                    x["_note"] = core.append_note(x.get("_note", ""), "换规详情:" + ",".join(det))
            if cat in KEEP_DETAIL and orig and orig != text:
                x["_note"] = core.append_note(x.get("_note", ""), f"详情:{orig}")
            if x["_needs_front"]:
                pf = x
            if x["_needs_temper"]:
                pt = x
        for x in deferred:
            orig = core.text(x.get("_defer_reason"))
            text, cat = format_defer_reason(x)
            x["_defer_reason"] = text
            x["_reason_cat"] = cat
            if orig and orig != text:
                x["_note"] = core.append_note(x.get("_note", ""), f"详情:{orig}")
        return main, deferred


if __name__ == "__main__":
    eng = Engine()
    n = eng.load()
    main, deferred = eng.run({})
    sched = Schedule(eng, main, deferred)
    print(f"载入订单 {n} 单;正常生产 {len(main)} 单,暂缓 {len(deferred)} 单")
    print("示例订单:", [x["_contract"] for x in main[:5]])
    print("KPI:", json.dumps(sched.kpi, ensure_ascii=False, indent=2))
    print("违规:", json.dumps(sched.issues, ensure_ascii=False, indent=2))
    p = sched.export(Path(__file__).resolve().parent / "out" / "热处理排程_默认基准.xlsx")
    print("已导出:", p)
