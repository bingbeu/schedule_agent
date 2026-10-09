# -*- coding: utf-8 -*-
"""一键自检:验证 llm_config.py 里的模型名与密钥是否可用。

用法(在 schedule_agent 目录):
    $env:DEEPSEEK_API_KEY = "sk-..."
    python check_model.py

不需要改 llm_config.py 时也可以临时测别的模型:
    python check_model.py --model deepseek-reasoner
"""
from __future__ import annotations

import argparse
import os

from agent_cli import llm_preflight
from llm_config import API_KEY_ENV, BASE_URL, MODEL


def main():
    ap = argparse.ArgumentParser(description="LLM 模型/密钥一键自检")
    ap.add_argument("--model", default=MODEL, help=f"要检测的模型名(默认 llm_config.py:{MODEL})")
    args = ap.parse_args()

    print(f"接口地址: {BASE_URL}")
    print(f"检测模型: {args.model}")
    key = os.environ.get(API_KEY_ENV, "").strip() if API_KEY_ENV else "local"
    if not key:
        print(f"⚠ 未设置 {API_KEY_ENV} 环境变量,无法完成验证。")
        print("  请先设置:$env:DEEPSEEK_API_KEY = \"sk-...\"(换成你实际用的变量名)")
        raise SystemExit(2)
    print("正在发送最小连通性请求(是否计费以服务商为准)...")
    ok, diag = llm_preflight(key, args.model)
    if ok:
        print("✅ 预检通过:密钥与模型名均可用,重启 webui.py / agent_cli.py 即可。")
    else:
        print(f"❌ 预检失败:{diag}")
        print("  常见原因:模型名填错(报400)/密钥无效(报401)/余额不足(报402)/网络不通。")


if __name__ == "__main__":
    main()

