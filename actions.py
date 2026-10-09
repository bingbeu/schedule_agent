# -*- coding: utf-8 -*-
"""排程动作 API —— 排程智能体第 2 层。

把"自然语言意图"翻译成这里的动作;每个动作返回结构化结果,
智能体据此判断改动是改善还是恶化。
"""
from __future__ import annotations

import re
import copy
import json
import math
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
        idx = int(p) - 1
        if not 0 <= idx <= len(others):
            raise ValueError("位置超出当前主序列范围")
        return idx
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
        self.version = 0
        self.history = []
        self.baseline: Schedule | None = None
        self.current: Schedule | None = None
        self.previous: Schedule | None = None

    # 所有动作先在副本生成候选，成功且可行才提交。
    def reset(self):
        old = copy.deepcopy(self.engine.__dict__)
        try:
            self.engine.load()
            result = Schedule(self.engine, *self.engine.run({}))
            self._require_feasible(result)
        except Exception:
            self.engine.__dict__.clear()
            self.engine.__dict__.update(old)
            raise
        self.active = {}
        self.baseline = self.current = result
        self.previous = None
        self.history = []
        self.version += 1
        return self._report("已重新加载输入并生成通过校验的默认方案")

    @staticmethod
    def _require_feasible(result):
        problems = [i for i in result.issues if i.get("类别") == "违规"]
        if problems:
            raise ValueError("候选违反硬约束，原方案保留:" + json.dumps(problems, ensure_ascii=False))

    def _commit(self, changes, note):
        proposed = copy.deepcopy(self.active)
        changes(proposed)
        old = copy.deepcopy(self.engine.__dict__)
        try:
            result = Schedule(self.engine, *self.engine.run(proposed))
            self._require_feasible(result)
        except Exception:
            self.engine.__dict__.clear()
            self.engine.__dict__.update(old)
            raise
        self.history.append(copy.deepcopy(self.active))
        self.active = proposed
        self.previous, self.current = self.current, result
        self.version += 1
        return self._report(note)

    def _report(self, note):
        s = self.current
        v, c = _issue_counts(s)
        return {"说明": note, "版本": self.version, "当前KPI": s.kpi,
                "相对上一步变化": diff_kpis(self.previous.kpi,s.kpi) if self.previous else {},
                "硬约束违规数": v, "必要换规边数": c, "问题明细": s.issues,
                "后处理记录": self.engine.diagnostics, "模型限制": self.engine.warnings}

    def reset_changes(self):
        return self._commit(lambda o:o.clear(), "已撤销全部改动")

    def undo(self):
        if not self.history:
            raise ValueError("没有可撤销的上一步")
        previous = copy.deepcopy(self.history[-1])
        def change(o):
            o.clear(); o.update(previous)
        result = self._commit(change,"已撤销上一步")
        self.history = self.history[:-2]
        return result

    def _refs(self, contracts):
        if not isinstance(contracts,list) or not contracts or not all(isinstance(c,str) and c.strip() for c in contracts):
            raise ValueError("任务编号必须是非空字符串列表")
        result = []
        for c in contracts:
            c = core.norm(c)
            hits = [x for x in self.engine.orders if core.order_id(x)==c]
            if not hits:
                hits = [x for x in self.engine.orders if core.norm(x["_contract"])==c]
            if not hits:
                raise ValueError(f"任务不存在:{c}")
            if len(hits)>1:
                raise ValueError(f"合同号{c}对应多条任务，请指定任务ID:"+"、".join(core.order_id(x) for x in hits))
            result.append(core.order_id(hits[0]))
        if len(result)!=len(set(result)):
            raise ValueError("参数中任务重复")
        return result

    def kpi(self):
        return {"当前KPI": self.current.kpi, "相对默认排程变化":diff_kpis(self.baseline.kpi,self.current.kpi),
                "硬约束违规数":self.current.kpi["硬约束违规数"], "版本":self.version}

    def validate(self):
        v,c = _issue_counts(self.current)
        return {"硬约束违规数":v,"必要换规边数":c,"明细":self.current.issues,
                "模型限制":self.engine.warnings,"后处理记录":self.engine.diagnostics}

    def compare(self):
        return {"默认KPI":self.baseline.kpi,"当前KPI":self.current.kpi,
                "相对默认变化":diff_kpis(self.baseline.kpi,self.current.kpi)}

    def status(self):
        return {"生效改动":copy.deepcopy(self.active),"版本":self.version,
                "正常生产":len(self.current.main),"暂缓":len(self.current.deferred),
                "硬约束违规数":self.current.kpi["硬约束违规数"]}

    def list_orders(self, filter="", offset=0, limit=40):
        if not isinstance(offset,int) or isinstance(offset,bool) or offset<0 or not isinstance(limit,int) or isinstance(limit,bool) or not 1<=limit<=200:
            raise ValueError("offset需为非负整数，limit需为1到200的整数")
        f=core.norm(str(filter))
        states={core.order_id(x):("生产",x) for x in self.current.main}
        states.update({core.order_id(x):("暂缓",x) for x in self.current.deferred})
        rows=[]
        for original in self.engine.orders:
            uid=core.order_id(original)
            stage,x=states.get(uid,("已移除",original))
            hay=core.norm(" ".join(core.text(x.get(k)) for k in ("_uid","_contract","_factory","_variety","_brand","_steel","_urgent")))
            if f and f not in hay:continue
            row={"任务ID":uid,"阶段":stage,"来源文件":x.get("_source_file"),"来源行":x.get("_source_row")}
            for label,key in (("订单编号","contract"),("主体厂","factory"),("品种","variety"),("牌号","brand"),("钢级","steel"),("外径","outer"),("壁厚","wall"),("前炉温度","front"),("回火温度","temper"),("数量","qty"),("计划产量","plan_tons"),("急催","urgent"),("热处理方式","process")):
                row[label]=core.text(x.get("_"+key)) if key in {"contract","factory","variety","brand","steel","urgent","process"} else x.get("_"+key)
            row["输入待确认"]=x.get("_input_errors",[])
            rows.append(row)
        return {"命中":len(rows),"订单":rows[offset:offset+limit],"下一页offset":offset+limit if offset+limit<len(rows) else None}

    def explain(self, contract):
        uid=self._refs([contract])[0]
        result=_explain_order(self.current.main,self.current.deferred,uid)
        result["任务ID"]=uid
        return result

    def move(self, contract, position):
        uid=self._refs([contract])[0]
        order=[core.order_id(x) for x in self.current.main]
        if uid not in order:raise ValueError("目标任务不在当前主序列")
        others=[c for c in order if c!=uid]
        if str(position).lower().startswith(("before:","after:")):
            direction,ref=str(position).split(":",1)
            position=direction+":"+self._refs([ref])[0]
        idx=resolve_position(position,others)
        def change(o):
            pinned=o.get("pinned",[])
            if uid in pinned or idx<len(pinned):
                raise ValueError("移动与已有锁定前缀冲突，请先解除锁定")
            o.setdefault("positions",{})[uid]=idx
        return self._commit(change,f"已将任务{uid}约束到第{idx+1}位；其余任务重新排程")

    def pin(self, contracts):
        ids=self._refs(contracts)
        if not set(ids)<=set(core.order_id(x) for x in self.current.main):raise ValueError("只能锁定当前主序列任务")
        return self._commit(lambda o:o.update(pinned=ids),f"已锁定前缀:{'、'.join(ids)}")

    def unpin(self):
        def change(o):o.pop("pinned",None);o.pop("positions",None)
        return self._commit(change,"已解除人工位置约束")

    def boost(self, contracts=None):
        ids=self._refs(contracts) if contracts is not None else [core.order_id(x) for x in self.engine.orders if core.text(x.get("_urgent")) and core.order_id(x) not in self.active.get("remove_contracts",[])]
        if not ids:raise ValueError("没有可加急的任务")
        def change(o):o["boost_contracts"]=sorted(set(o.get("boost_contracts",[]))|set(ids))
        return self._commit(change,f"已给{len(ids)}条任务加急")

    def unboost(self, contracts=None):
        ids=self._refs(contracts) if contracts is not None else None
        def change(o):o["boost_contracts"]=sorted(set(o.get("boost_contracts",[]))-set(ids)) if ids is not None else []
        return self._commit(change,"已取消加急优先")

    @staticmethod
    def _number(value, name, allow_zero=False):
        if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value) or value<0 or (value==0 and not allow_zero):
            raise ValueError(f"{name}必须为{'非负' if allow_zero else '正'}有限数")
        return float(value)

    def temper_change(self, contract, temper):
        uid=self._refs([contract])[0]
        temperature=self._number(temper,"回火温度")
        row=next(x for x in self.engine.orders if core.order_id(x)==uid)
        if not row["_needs_temper"]:raise ValueError("该任务无回火工序，修改温度不能新增工序")
        def change(o):o.setdefault("temper_map",{})[uid]=temperature
        return self._commit(change,f"已将{uid}回火温度设为{temperature:g}℃；工艺允许温度范围仍需现场确认")

    def delay(self, minutes):
        minutes=self._number(minutes,"开工顺延分钟",True)
        return self._commit(lambda o:o.update(start_delay_minutes=minutes),f"开工顺延设为{minutes:g}分钟(替换原设置)")

    def remove_orders(self, contracts):
        ids=self._refs(contracts)
        if set(ids)&set(self.active.get("remove_contracts",[])):raise ValueError("部分任务已经移除")
        if set(ids)&(set(self.active.get("pinned",[]))|set(self.active.get("positions",{}))):raise ValueError("请先解除相关人工位置约束")
        def change(o):o["remove_contracts"]=sorted(set(o.get("remove_contracts",[]))|set(ids))
        return self._commit(change,f"已移除{len(ids)}条任务")

    def restore_orders(self, contracts):
        ids=self._refs(contracts)
        if not set(ids)<=set(self.active.get("remove_contracts",[])):raise ValueError("部分任务未被移除，不能恢复")
        def change(o):o["remove_contracts"]=sorted(set(o.get("remove_contracts",[]))-set(ids))
        return self._commit(change,f"已恢复{len(ids)}条任务")

    def export(self):
        p=self.current.export(self.out_dir/"热处理排程_智能体版.xlsx")
        return {"文件":str(p),"版本":self.version,"正常生产":len(self.current.main),"暂缓":len(self.current.deferred),"模型限制":self.engine.warnings}
