# -*- coding: utf-8 -*-
"""对话式排程顾问 —— 排程智能体第 3 层(入口)。

两种模式:
1. LLM 模式:设置环境变量 DEEPSEEK_API_KEY 后自动启用,走 DeepSeek 工具调用;
2. 离线模式:未设置密钥时使用内置中文指令解析,同样演示完整动作链。

运行:
    python agent_cli.py                 # 默认读取原脚本 data/rule/工模具 目录
    python agent_cli.py --no-llm        # 强制离线模式
"""
from __future__ import annotations

import argparse
import json
import math
import os
import re
import urllib.error
import urllib.request
from pathlib import Path

from actions import SchedulerAgent
from engine import Engine
from llm_config import API_KEY_ENV, BASE_URL, MODEL as DEFAULT_MODEL

SYSTEM_PROMPT = """你是某钢管厂630热处理产线的排程顾问智能体。你只能通过工具操作排程,绝不能自己心算顺序或时间。

工作原则:
1. 硬约束(相邻空格≤20格、回火温度回摆、前炉温度回摆)由排程引擎校验;候选必须满足硬约束才会提交。错误时原方案保留，不要声称动作成功。合法跨温区换规可以超过20格。
2. 排程质量看KPI:相邻空格总数、超20格边数、暂缓单数、完工时间;对比"相对默认排程变化"。
3. 不知道任务时先 list_orders;合同号可能对应多任务，必须用工具返回的任务ID消歧，不要猜测。再用 move/pin/boost/temper_change/delay/remove_orders。
4. 用户问"为什么这样排"用 explain。
5. 用户要文件时用 export。
6. 温度范围、实际设备容量和节拍联动没有确认时不要宣称方案可直接生产。
7. 简体中文回答,简洁,引用工具返回的具体数字,不编造。"""

HELP_TEXT = """离线模式指令示例(LLM 模式直接说人话即可):
  KPI / 指标              查看当前排程指标
  校验 / 违规             校验硬约束
  有哪些订单 / 列出       列出订单目录
  解释 <订单号>           解释该单为何这样排
  把 <订单号> 提到最前    移动订单并重排
  把 <订单号> 放在 <订单号> 前面
  锁定 <订单号> <订单号>  锁定序列最前相对顺序
  急催                   给所有急催单加优先权重
  取消加急               撤销急催权重
  延迟 2 小时             整线顺延(模拟开工前检修)
  改 <订单号> 回火温度 620
  移除 <订单号>           完工/取消移出
  恢复订单 <订单号>
  撤销改动               回到默认排程
  撤销上一步             撤销最近一次成功动作
  解除锁定               解除人工位置约束
  对比                   当前 vs 默认 KPI
  导出                   输出 Excel
  重排 / 重置            重新读 Excel 重排
  帮助                   显示本帮助"""

TOOL_SPECS = [
    {"name": "list_orders", "description": "列出订单目录(编号/主体厂/牌号/钢级/外径/壁厚/温度/急催),支持关键字过滤。",
     "parameters": {"type": "object",
                    "properties": {"filter": {"type": "string", "description": "关键字,留空列出全部"}},
                    "required": ["filter"]}},
    {"name": "kpi", "description": "查看当前排程KPI及相对默认排程的变化。",
     "parameters": {"type": "object", "properties": {}, "required": []}},
    {"name": "validate", "description": "校验硬约束(20格上限/温度回摆),返回违规清单。",
     "parameters": {"type": "object", "properties": {}, "required": []}},
    {"name": "explain", "description": "解释某订单的排产原因、时间与工模具。",
     "parameters": {"type": "object",
                    "properties": {"contract": {"type": "string", "description": "订单编号"}},
                    "required": ["contract"]}},
    {"name": "move", "description": "把订单移到指定位置并重排。position:first/last/第N位数字/before:订单编号/after:订单编号。",
     "parameters": {"type": "object",
                    "properties": {"contract": {"type": "string"},
                                   "position": {"type": "string", "description": "first 或 last 或 数字 或 before:订单编号 或 after:订单编号"}},
                    "required": ["contract", "position"]}},
    {"name": "pin", "description": "锁定若干订单为序列最前的固定相对顺序,其余由引擎优化。",
     "parameters": {"type": "object",
                    "properties": {"contracts": {"type": "array", "items": {"type": "string"},
                                                 "description": "订单编号列表,按给定顺序锁定"}},
                    "required": ["contracts"]}},
    {"name": "boost", "description": "给订单加急催优先权重(不传参数则给所有带急催标记的订单)。",
     "parameters": {"type": "object",
                    "properties": {"contracts": {"type": "array", "items": {"type": "string"}}},
                    "required": []}},
    {"name": "unboost", "description": "取消加急优先权重(不传参数则全部取消)。",
     "parameters": {"type": "object",
                    "properties": {"contracts": {"type": "array", "items": {"type": "string"}}},
                    "required": []}},
    {"name": "temper_change", "description": "修改订单回火温度并重排(如工艺变更)。",
     "parameters": {"type": "object",
                    "properties": {"contract": {"type": "string"},
                                   "temper": {"type": "number", "description": "新回火温度℃"}},
                    "required": ["contract", "temper"]}},
    {"name": "delay", "description": "整线顺延N分钟(模拟开工前检修/停机)。",
     "parameters": {"type": "object",
                    "properties": {"minutes": {"type": "number"}},
                    "required": ["minutes"]}},
    {"name": "remove_orders", "description": "移除订单(完工/取消)并重排。",
     "parameters": {"type": "object",
                    "properties": {"contracts": {"type": "array", "items": {"type": "string"}}},
                    "required": ["contracts"]}},
    {"name": "restore_orders", "description": "把已移除的订单放回并重排。",
     "parameters": {"type": "object",
                    "properties": {"contracts": {"type": "array", "items": {"type": "string"}}},
                    "required": ["contracts"]}},
    {"name": "reset_changes", "description": "撤销全部改动,回到默认排程(不重新读Excel)。",
     "parameters": {"type": "object", "properties": {}, "required": []}},
    {"name": "reset", "description": "重新读取Excel数据并生成默认排程。",
     "parameters": {"type": "object", "properties": {}, "required": []}},
    {"name": "compare", "description": "对比当前排程与默认排程KPI。",
     "parameters": {"type": "object", "properties": {}, "required": []}},
    {"name": "export", "description": "导出当前排程Excel(排产总览/原始订单/钢管级时间轴)。",
     "parameters": {"type": "object", "properties": {}, "required": []}},
    {"name": "status", "description": "查看当前生效改动与订单概况。",
     "parameters": {"type": "object", "properties": {}, "required": []}},
]

for spec in TOOL_SPECS:
    spec["parameters"]["additionalProperties"] = False
TOOL_SPECS[0]["parameters"]["properties"].update(offset={"type":"integer"},limit={"type":"integer"})
for name,description in (("undo","撤销上一步成功改动"),("unpin","解除全部人工位置约束")):
    TOOL_SPECS.append({"name":name,"description":description,"parameters":{"type":"object","properties":{},"required":[],"additionalProperties":False}})


def wrap_tools():
    tools = []
    for t in TOOL_SPECS:
        spec = {"name": t["name"], "description": t["description"], "parameters": t["parameters"]}
        if not spec["parameters"].get("required"):
            spec["parameters"] = {k: v for k, v in spec["parameters"].items() if k != "required"}
        tools.append({"type": "function", "function": spec})
    return tools


def dispatch(agent, name, kwargs):
    try:
        specs={t["name"]:t for t in TOOL_SPECS}
        if name not in specs:
            raise ValueError("不允许的工具")
        if not isinstance(kwargs,dict):
            raise ValueError("工具参数必须为JSON对象")
        schema=specs[name]["parameters"]
        unknown=set(kwargs)-set(schema.get("properties",{}))
        if unknown:raise ValueError("未知参数:"+",".join(sorted(unknown)))
        missing=set(schema.get("required",[]))-set(kwargs)
        if missing:raise ValueError("缺少参数:"+",".join(sorted(missing)))
        for key,value in kwargs.items():
            typ=schema["properties"][key]["type"]
            valid={"string":isinstance(value,str),"number":isinstance(value,(int,float)) and not isinstance(value,bool) and math.isfinite(value),"integer":isinstance(value,int) and not isinstance(value,bool),"array":isinstance(value,list) and all(isinstance(v,str) for v in value)}.get(typ,False)
            if not valid:raise ValueError(f"参数{key}类型不正确")
        return getattr(agent,name)(**kwargs)
    except Exception as e:
        return {"error":f"工具{name}执行失败:{e}","版本":agent.version}


def split_ids(s):
    return [x for x in re.split(r"[\s,，、;；]+",s.strip()) if x]


def parse_offline(text):
    t=text.strip()
    if not t:return (None,{}),"请输入指令。"
    if re.search(r"不要|别(?:移|删|改|排)|取消(?:删除|移除|修改)",t):
        return (None,{}),"识别到否定或取消操作，本次未执行。撤销上一步请使用‘撤销上一步’。"
    if re.search(r"然后|接着|再导出|同时|之后",t):
        return (None,{}),"离线模式一次只接受一条指令，请分开发送；本次未执行任何动作。"
    literals={"帮助":None,"help":None,"示例":None,"怎么用":None,
              "KPI":"kpi","kpi":"kpi","指标":"kpi","情况":"kpi",
              "状态":"status","校验":"validate","违规":"validate","检查":"validate",
              "对比":"compare","导出":"export","输出":"export","生成excel":"export",
              "急催":"boost","加急":"boost","取消加急":"unboost","取消急催":"unboost",
              "解除锁定":"unpin","撤销上一步":"undo","撤销改动":"reset_changes",
              "恢复默认":"reset_changes","回到默认":"reset_changes","重置":"reset","重排":"reset"}
    if t in literals:
        tool=literals[t]
        return (tool,{}),HELP_TEXT if tool is None else ""
    if t in {"有哪些订单","列出","订单列表","所有订单"}:return ("list_orders",{"filter":""}),""
    m=re.fullmatch(r"(?:查|查询|列出)\s+(.+)",t)
    if m:return ("list_orders",{"filter":m.group(1)}),""
    m=re.fullmatch(r"(?:解释|说明)\s+(.+)",t)
    if m:return ("explain",{"contract":m.group(1)}),""
    m=re.fullmatch(r"(?:把|将)?\s*(.+?)\s*(?:提到|移到|排到)\s*(最前|最后|末尾|第\s*\d+\s*位)",t)
    if m:
        target=m.group(2)
        position="first" if target=="最前" else "last" if target in {"最后","末尾"} else re.search(r"\d+",target).group()
        return ("move",{"contract":m.group(1).strip(),"position":position}),""
    m=re.fullmatch(r"(?:把|将)?\s*(.+?)\s*放在\s*(.+?)\s*(前面|后面)",t)
    if m:return ("move",{"contract":m.group(1).strip(),"position":("before:" if m.group(3)=="前面" else "after:")+m.group(2).strip()}),""
    m=re.fullmatch(r"(?:改|修改)\s+(.+?)\s+回火温度\s+([+-]?\d+(?:\.\d+)?)",t)
    if not m:m=re.fullmatch(r"(?:把|将)?\s*(.+?)\s*的?回火温度(?:改为|改|调为|设为|降到|升到)\s*([+-]?\d+(?:\.\d+)?)",t)
    if m:return ("temper_change",{"contract":m.group(1).strip(),"temper":float(m.group(2))}),""
    m=re.fullmatch(r"(?:延迟|顺延)\s*([+-]?\d+(?:\.\d+)?)\s*(小时|分钟|min|h)?",t)
    if m:return ("delay",{"minutes":float(m.group(1))*(60 if m.group(2) in {"小时","h"} else 1)}),""
    m=re.fullmatch(r"(锁定|移除|删除|恢复订单|加急|取消加急)\s+(.+)",t)
    if m:return ({"锁定":"pin","移除":"remove_orders","删除":"remove_orders","恢复订单":"restore_orders","加急":"boost","取消加急":"unboost"}[m.group(1)],{"contracts":split_ids(m.group(2))}),""
    return (None,{}),"未识别完整指令，本次未执行。输入‘帮助’查看模板。"


def call_llm(messages, tools, api_key, model):
    payload = {"model": model, "messages": messages, "tools": tools, "tool_choice": "auto"}
    req = urllib.request.Request(
        BASE_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
    )
    with urllib.request.urlopen(req, timeout=45) as resp:
        return json.loads(resp.read().decode("utf-8"))


def llm_preflight(api_key, model):
    """启动前连通性检查(1 token 的 ping,20 秒超时)。"""
    try:
        payload = {"model": model, "messages": [{"role": "user", "content": "ping"}], "max_tokens": 1}
        req = urllib.request.Request(
            BASE_URL,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
        )
        with urllib.request.urlopen(req, timeout=20) as resp:
            resp.read()
        return True, "连接正常"
    except urllib.error.HTTPError as e:
        try:
            detail = e.read().decode("utf-8", "replace")[:300]
        except Exception:  # noqa: BLE001
            detail = e.reason
        return False, f"HTTP {e.code}: {detail}"
    except Exception as e:  # noqa: BLE001
        return False, f"{type(e).__name__}: {e}"


def llm_chat(agent, api_key, model):
    msgs = [{"role": "system", "content": SYSTEM_PROMPT}]
    print("LLM 模式已启用。直接说需求即可,例如:")
    print("  · 现在排程的整体情况怎么样")
    print("  · 把订单 2326030045001 提到最前,看看对整体有什么影响")
    print("  · 解释订单 1126040218013 为什么这样排")
    print("  · 给急催订单加优先,然后和默认排程对比")
    print("  · 导出排程表")
    print("输入 exit 退出。\n")
    while True:
        try:
            user = input("你 > ").strip()
        except EOFError:
            break
        if user.lower() in {"exit", "quit", "退出"}:
            break
        if not user:
            continue
        msgs.append({"role": "user", "content": user})
        final = ""
        for _ in range(6):
            try:
                body = call_llm(msgs, wrap_tools(), api_key, model)
            except urllib.error.HTTPError as e:
                try:
                    detail = e.read().decode("utf-8", "replace")[:300]
                except Exception:  # noqa: BLE001
                    detail = e.reason
                print(f"助手 > [API错误 HTTP {e.code}] {detail}")
                final = ""
                break
            except Exception as e:  # noqa: BLE001
                print("助手 > [LLM调用失败]", type(e).__name__, e)
                final = ""
                break
            msg = body["choices"][0]["message"]
            msgs.append(msg)
            if not msg.get("tool_calls"):
                final = msg.get("content") or ""
                break
            for tc in msg["tool_calls"]:
                fn = tc["function"]
                try:
                    kwargs = json.loads(fn.get("arguments") or "{}")
                except Exception:  # noqa: BLE001
                    kwargs = None
                print(f"  [工具] {fn['name']} {json.dumps(kwargs, ensure_ascii=False)}")
                res = dispatch(agent, fn["name"], kwargs)
                msgs.append({"role": "tool", "tool_call_id": tc["id"],
                             "content": json.dumps(res, ensure_ascii=False, default=str)})
        else:
            final = "已达到最大工具轮数,请查看上面工具结果。"
        if final:
            print("助手 >", final)


def offline_chat(agent):
    print("离线模式(未设置 DEEPSEEK_API_KEY,使用内置中文指令解析)。输入 exit 退出,输入'帮助'看示例。\n")
    while True:
        try:
            user = input("你 > ").strip()
        except EOFError:
            break
        if user.lower() in {"exit", "quit", "退出"}:
            break
        if not user:
            continue
        (tool, kwargs), hint = parse_offline(user)
        if tool is None:
            print("助手 >", hint)
            continue
        res = dispatch(agent, tool, kwargs)
        print("助手 >", json.dumps(res, ensure_ascii=False, indent=2, default=str))


def main():
    ap = argparse.ArgumentParser(description="热处理排程智能体 PoC(对话式排程顾问)")
    ap.add_argument("--input", type=Path, default=None, help="订单数据目录,默认用原脚本 data 目录")
    ap.add_argument("--rule", type=Path, default=None, help="换规规则目录,默认用原脚本 rule 目录")
    ap.add_argument("--tooling", type=Path, default=None, help="工模具目录,默认用原脚本工模具目录")
    ap.add_argument("--combine-inputs", action="store_true", help="明确合并多个排产文件")
    ap.add_argument("--out", type=Path, default=None, help="Excel 输出目录,默认本目录 out/")
    ap.add_argument("--model", default=DEFAULT_MODEL, help=f"模型名(默认取 llm_config.py:{DEFAULT_MODEL})")
    ap.add_argument("--no-llm", action="store_true", help="强制离线模式")
    args = ap.parse_args()

    engine = Engine(args.input, args.rule, args.tooling, args.combine_inputs)
    agent = SchedulerAgent(engine, args.out)
    print("正在加载数据并生成默认排程 ...")
    print(json.dumps(agent.reset(), ensure_ascii=False, indent=2))

    api_key = os.environ.get(API_KEY_ENV, "").strip() if API_KEY_ENV else "local"
    if api_key and not args.no_llm:
        print(f"\n[预检] 正在检查 LLM 接口连通性({BASE_URL}) ...")
        ok, diag = llm_preflight(api_key, args.model)
        if not ok:
            print(f"[警告] LLM 预检失败:{diag}")
            print("[提示] 常见原因:密钥无效/余额不足/网络不通/模型名不存在。")
            print("[提示] 可先用离线模式体验:python agent_cli.py --no-llm")
            try:
                ans = input("是否切换到离线模式继续?[y/n,默认y]").strip().lower()
            except EOFError:
                ans = ""
            if ans not in {"n", "no", "否", "不"}:
                print("已切换到离线模式。\n")
                offline_chat(agent)
                return
            print("保持 LLM 模式继续(出错时会显示具体原因)。")
        else:
            print("[预检] 连接正常,进入 LLM 模式。")
        llm_chat(agent, api_key, args.model)
    else:
        if not args.no_llm:
            print(f"\n[提示] 未检测到 {API_KEY_ENV},进入离线模式;设置该环境变量即可启用大模型对话"
                  f"(服务商/模型可在 llm_config.py 修改)。")
        offline_chat(agent)


if __name__ == "__main__":
    main()

