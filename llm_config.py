# -*- coding: utf-8 -*-
"""LLM 配置 —— 更换大模型/服务商只改这一个文件。

所有"OpenAI 兼容"接口都可以直接替换(改 BASE_URL + MODEL + API_KEY_ENV 三行):

- DeepSeek:      https://api.deepseek.com/chat/completions
                 MODEL: deepseek-flash(当前默认,Flash 版;官方别名,不带版本号)
                        deepseek-v4-pro(Pro 版)
                        deepseek-chat / deepseek-reasoner(旧别名,是否可用视账号而定)
                 密钥环境变量: DEEPSEEK_API_KEY
- 月之暗面 Kimi:  https://api.moonshot.cn/v1/chat/completions
                 MODEL: kimi-k2-turbo / moonshot-v1-8k          密钥环境变量: MOONSHOT_API_KEY
- 阿里通义千问:   https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions
                 MODEL: qwen-plus / qwen-max                    密钥环境变量: DASHSCOPE_API_KEY
- 智谱 GLM:      https://open.bigmodel.cn/api/paas/v4/chat/completions
                 MODEL: glm-4.5 / glm-4-flash                   密钥环境变量: ZHIPUAI_API_KEY
- 本地 Ollama:   http://127.0.0.1:11434/v1/chat/completions
                 MODEL: qwen2.5:14b / llama3.1 等               API_KEY_ENV 留空 ""

注意:
1. 所选模型必须支持 function calling(工具调用),否则智能体无法操作排程工具;
   DeepSeek/Kimi/通义/智谱的对话模型均支持,Ollama 建议用 qwen2.5、llama3.1 及以上。
2. 改完本文件后重启 python webui.py / python agent_cli.py 生效。
3. 不确定模型名是否正确时,先跑 python check_model.py 一键自检(或直接看启动预检);
   预检会发出最小请求；是否计费以服务商为准。
"""

BASE_URL = "https://api.deepseek.com/chat/completions"
MODEL = "deepseek-flash"
API_KEY_ENV = "DEEPSEEK_API_KEY"   # 密钥从哪个环境变量读取;本地模型(如 Ollama)可设为 ""

