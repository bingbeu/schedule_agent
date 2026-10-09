# -*- coding: utf-8 -*-
"""排程动作 API —— 排程智能体第 2 层。

把"自然语言意图"翻译成这里的动作;每个动作返回结构化结果,
智能体据此判断改动是改善还是恶化。
"""
from __future__ import annotations

import re
from pathlib import Path

from engine import Schedule, core, diff_kpis, explain as _explain_order


def resolve_position(position: str, others: list[str]) -> int:
    """解析位置描述:first/last/数字/before:订单编号/after:订单编号。"""
    p = str(position).strip()
    pl = p.lower()
    if pl in {"first", "最前", "首位", "第一"}:
        return 0
    if pl in {"last", "最后", "末尾", "末位"}:
        return len(others)
    if pl.startswith("before:"):
        return others.index(core.norm(p.split(":", 1)[1]))
    if pl.startswith("after:"):
        return others.index(core.norm(p.split(":", 1)[1])) + 1
    if re.fullmatch(r"\d+", p):
        return max(0, min(int(p) - 1, len(others)))
    raise ValueError(f"无法识别的position:{position}(可用 first/last/数字/before:订单编号/after:订单编号)")


def _issue_counts(s) -> tuple[int, int]:
    """返回 (硬约束违规数, 必要换规边数)。"""
    v = sum(1 for i in s.issues if i.get("类别") == "违规")
    c = sum(1 for i in s.issues if i.get("类别") == "必要换规")
    return v, c


class SchedulerAgent:
    """会话状态 + 全部排程动作(即智能体的工具)。"""

    def __init__(self, engine, out_dir=None):
        self.engine = engine
        self.out_dir = Path(out_dir) if out_dir else Path(__file__).resolve().parent / "out"
        self.active: dict = {}
        self.baseline: Schedule | None = None
        self.current: Schedule | None = None
        self.previous: Schedule | None = None

    # ---------------- 会话管理 ----------------
    def reset(self):
        """重新读 Excel 并生成默认排程(基线)。"""
        self.active = {}
        self.engine.load()
        self.baseline = Schedule(self.engine, *self.engine.run({}))
        self.current, self.previous = self.baseline, None
        v, c = _issue_counts(self.current)
        return {"说明": f"已重新读取数据并生成默认排程:正常生产{len(self.current.main)}单,暂缓{len(self.current.deferred)}单",
                "当前KPI": self.current.kpi, "硬约束违规数": v, "必要换规边数": c}

    def reset_changes(self):
        """撤销全部改动,回到默认排程。"""
        self.active = {}
        return self._report("已撤销全部改动,回到默认排程", self._rerun())

    def _rerun(self) -> Schedule:
        s = Schedule(self.engine, *self.engine.run(self.active))
        self.previous, self.current = self.current, s
        return s

    def _report(self, note, s=None):
        s = s or self.current
        d = diff_kpis(self.previous.kpi, s.kpi) if self.previous else {}
        v, c = _issue_counts(s)
        return {"说明": note, "当前KPI": s.kpi, "相对上一步变化": d,
                "硬约束违规数": v, "必要换规边数": c, "问题明细": s.issues[:5]}

    # ---------------- 查询类工具 ----------------
    def kpi(self):
        v, c = _issue_counts(self.current)
        return {"当前KPI": self.current.kpi,
                "相对默认排程变化": diff_kpis(self.baseline.kpi, self.current.kpi),
                "硬约束违规数": v, "必要换规边数": c}

    def validate(self):
        v, c = _issue_counts(self.current)
        return {"硬约束违规数": v, "必要换规边数": c, "明细": self.current.issues}

    def compare(self):
        return {"默认KPI": self.baseline.kpi, "当前KPI": self.current.kpi,
                "相对默认变化": diff_kpis(self.baseline.kpi, self.current.kpi)}

    def status(self):
        v, _ = _issue_counts(self.current)
        return {"生效改动": {k: (sorted(v) if isinstance(v, set) else v) for k, v in self.active.items()},
                "正常生产": len(self.current.main), "暂缓": len(self.current.deferred),
                "硬约束违规数": v}

    def list_orders(self, filter=""):
        f = core.norm(str(filter))
        rows = []
        for x in self.engine.orders:
            hay = core.norm(" ".join(core.text(v) for v in (
                x["_contract"], x.get("_factory"), x.get("_variety"),
                x.get("_brand"), x.get("_steel"), x.get("_urgent"))))
            if f and f not in hay:
                continue
            rows.append({"订单编号": x["_contract"], "主体厂": core.text(x.get("_factory")),
                         "品种": core.text(x.get("_variety")), "牌号": core.text(x.get("_brand")),
                         "钢级": core.text(x.get("_steel")), "外径": x.get("_outer"),
                         "壁厚": x.get("_wall"), "前炉温度": x.get("_front"),
                         "回火温度": x.get("_temper"), "数量": x.get("_qty"),
                         "计划产量": x.get("_plan_tons"), "急催": core.text(x.get("_urgent")),
                         "热处理方式": core.text(x.get("_process"))})
        return {"命中": len(rows), "订单": rows[:40]}

    def explain(self, contract):
        return _explain_order(self.current.main, self.current.deferred, contract)

    # ---------------- 改动类工具 ----------------
    def move(self, contract, position):
        n = core.norm(str(contract))
        order = [core.norm(str(x["_contract"])) for x in self.current.main]
        if n not in order:
            return {"error": f"订单 {contract} 不在主序列,无法移动;可先 list_orders 查编号"}
        others = [c for c in order if c != n]
        idx = resolve_position(position, others)
        if idx == len(others):
            self.active.pop("pinned", None)
            s = self._rerun()
            return self._report(f"已解除 {contract} 的位置约束(移到末尾=交回引擎自由优化)", s)
        order.remove(n)
        order.insert(idx, n)
        self.active["pinned"] = order[: idx + 1]
        s = self._rerun()
        pos_txt = "最前" if idx == 0 else f"第{idx + 1}位"
        return self._report(f"已把 {contract} 移到{pos_txt},其余订单由引擎重新优化", s)

    def pin(self, contracts):
        pinned, rest = [], [core.norm(str(x["_contract"])) for x in self.current.main]
        for c in contracts:
            n = core.norm(str(c))
            if n in rest:
                rest.remove(n)
                pinned.append(n)
        if not pinned:
            return {"error": "没有找到可锁定的订单编号"}
        self.active["pinned"] = pinned
        s = self._rerun()
        return self._report(f"已锁定 {len(pinned)} 单为序列最前:{'、'.join(contracts)}", s)

    def boost(self, contracts=None):
        if contracts is None:
            contracts = [x["_contract"] for x in self.engine.orders if core.text(x.get("_urgent", ""))]
        norms = {core.norm(str(c)) for c in contracts}
        if not norms:
            return {"error": "没有可加优先的订单(数据中无急催标记,或参数为空)"}
        self.active.setdefault("boost_contracts", set()).update(norms)
        s = self._rerun()
        return self._report(f"已给 {len(norms)} 单加急催优先权重", s)

    def unboost(self, contracts=None):
        cur = self.active.get("boost_contracts", set())
        cur = set() if contracts is None else cur - {core.norm(str(c)) for c in contracts}
        if cur:
            self.active["boost_contracts"] = cur
        else:
            self.active.pop("boost_contracts", None)
        s = self._rerun()
        return self._report("已取消加急优先权重", s)

    def temper_change(self, contract, temper):
        self.active.setdefault("temper_map", {})[core.norm(str(contract))] = float(temper)
        s = self._rerun()
        return self._report(f"已把 {contract} 回火温度改为 {temper}℃ 并重排", s)

    def delay(self, minutes):
        self.active["start_delay_minutes"] = float(minutes)
        s = self._rerun()
        return self._report(f"已整线顺延 {minutes} 分钟(模拟开工前检修/停机)", s)

    def remove_orders(self, contracts):
        self.active.setdefault("remove_contracts", set()).update(core.norm(str(c)) for c in contracts)
        s = self._rerun()
        return self._report(f"已移除 {len(contracts)} 单(完工/取消)并重排", s)

    def restore_orders(self, contracts):
        cur = self.active.get("remove_contracts", set()) - {core.norm(str(c)) for c in contracts}
        if cur:
            self.active["remove_contracts"] = cur
        else:
            self.active.pop("remove_contracts", None)
        s = self._rerun()
        return self._report(f"已恢复 {len(contracts)} 单并重排", s)

    def export(self, path=None):
        path = Path(path) if path else self.out_dir / "热处理排程_智能体版.xlsx"
        p = self.current.export(path)
        return {"文件": str(p), "正常生产": len(self.current.main), "暂缓": len(self.current.deferred)}
