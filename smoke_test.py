# -*- coding: utf-8 -*-
"""脚本化冒烟测试:不经过 stdin,直接驱动动作 API 与离线解析器。"""
from __future__ import annotations

import json

from actions import SchedulerAgent
from agent_cli import dispatch, parse_offline
from engine import Engine


def show(label, res):
    print(f"\n===== {label} =====")
    print(json.dumps(res, ensure_ascii=False, indent=2, default=str)[:2500])


def main():
    engine = Engine()
    agent = SchedulerAgent(engine)
    show("reset(基线排程)", agent.reset())
    first = [x["_contract"] for x in agent.current.main[:3]]
    print("\n基线前3单:", first)
    cmds = [
        "KPI",
        "校验",
        f"解释 {first[0]}",
        f"把 {first[1]} 提到最前",
        "KPI",
        "对比",
        "延迟 2 小时",
        "KPI",
        "撤销改动",
        "急催",
        "对比",
        "列出",
        "导出",
    ]
    for c in cmds:
        (tool, kwargs), hint = parse_offline(c)
        if tool is None:
            show(f"指令: {c} -> 未识别", {"提示": hint})
            continue
        res = dispatch(agent, tool, kwargs)
        show(f"指令: {c} -> {tool} {json.dumps(kwargs, ensure_ascii=False)}", res)
        if tool == "move":
            moved_head = agent.current.main[0]["_contract"]
            print(">>> move 后首单:", moved_head)

    # ---- 结果断言 ----
    print("\n===== 断言 =====")
    print("移动目标单:", first[1], "| move 后首单:", moved_head, "| 移动是否改变序列:", moved_head == first[1])
    if moved_head != first[1]:
        raise SystemExit("FAIL: move 未把订单移到最前")
    from pathlib import Path
    out = Path("out") / "热处理排程_智能体版.xlsx"
    print("导出文件存在:", out.exists(), out)
    if not out.exists():
        raise SystemExit("FAIL: 导出文件不存在")
    print("全部断言通过")


if __name__ == "__main__":
    main()
