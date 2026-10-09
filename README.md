# 热处理排程智能体 PoC(对话式排程顾问)

在 `C:\Users\mblon\Desktop\project\schedule_heat_treatment_strict_compact.py`(原排程内核,一行未改)
之上搭建的三层智能体原型。核心原则:**智能体只做"提议与解释",硬约束(20格上限、温度回摆、换规规则)
和时间计算永远由确定性引擎裁决。**

```
┌──────────────────────────────────────────────────────────────┐
│ 第3层 agent_cli.py  对话入口                                  │
│   LLM 模式:DeepSeek 工具调用(设置 DEEPSEEK_API_KEY 即启用)    │
│   离线模式:内置中文指令解析(无密钥也能演示完整动作链)          │
└───────────────────────────┬──────────────────────────────────┘
┌───────────────────────────▼──────────────────────────────────┐
│ 第2层 actions.py  动作 API(即智能体的"工具")                  │
│   reset / kpi / validate / compare / status / list_orders     │
│   explain / move / pin / boost / unboost / temper_change      │
│   delay / remove_orders / restore_orders / reset_changes      │
│   export                                                      │
└───────────────────────────┬──────────────────────────────────┘
┌───────────────────────────▼──────────────────────────────────┐
│ 第1层 engine.py  确定性内核封装                               │
│   原排程脚本(函数级 API 调用,不复制不修改)                    │
│   + validate 校验器(违规 vs 必要换规 分流)                    │
│   + kpis 质量指标(空格/超20格/暂缓/急催/完工时间)             │
└──────────────────────────────────────────────────────────────┘
```

## 文件

| 文件 | 作用 |
|---|---|
| `engine.py` | 内核封装:Engine(读取/重排)、Schedule(结果+KPI+导出)、validate、kpis、diff_kpis、explain |
| `actions.py` | SchedulerAgent:会话状态 + 全部排程动作,每个动作返回结构化结果(含"相对上一步变化") |
| `agent_cli.py` | 对话入口:LLM 工具调用(OpenAI 兼容接口)+ 离线中文指令解析,工具 Schema 见 TOOL_SPECS |
| `llm_config.py` | **LLM 配置:换模型/换服务商只改这一个文件** |
| `smoke_test.py` | 脚本化冒烟测试(真实数据全流程 + 断言) |
| `out/` | 导出的 Excel(默认基准 / 智能体调整版) |

## 快速开始

```powershell
cd C:\Users\mblon\Desktop\working\deepseek\schedule_agent

# 离线模式(默认,无需任何密钥)
python agent_cli.py

# 启用大模型对话(需要 DeepSeek API Key;不要把真实密钥写进任何文件)
  # 每次开新终端窗口时设置一次即可
python agent_cli.py

# 冒烟测试(真实数据)
python smoke_test.py
```

LLM 模式启动时会先做一次连通性预检(1 token,20 秒超时):密钥无效/余额不足/网络不通都会
直接显示原因,并询问是否切换到离线模式,不会再出现"输入后没反应"的情况。

## 网页版(推荐日常使用)

零依赖,单文件 `webui.py`,浏览器里对话 + 看 KPI + 导出 Excel:

```powershell
python webui.py                 # 然后浏览器打开 http://127.0.0.1:8787
python webui.py --port 9000     # 换端口
```

- 左侧对话:LLM 模式说人话;离线模式用模板指令(左下角有快捷指令按钮)
- 右侧面板:排程 KPI(含相对默认的变化,改善绿/恶化红)、硬约束违规清单、必要换规清单、订单速查
- 顶部按钮:重新加载数据 / 导出 Excel
- 有 DEEPSEEK_API_KEY 自动启用大模型;密钥无效/网络不通自动退回离线模式(启动日志说明原因)
- 默认只监听 127.0.0.1(仅本机可访问),**不要**用 `--host 0.0.0.0` 暴露到公网

接口(便于以后接别的系统):`GET /api/state`、`POST /api/chat`、`POST /api/action`、
`GET /api/orders?filter=`、`GET /api/export`、`POST /api/reset`

## 更换大模型 / 服务商

**只改 `llm_config.py` 一个文件**(CLI 与网页共用),改完重启生效:

```python
BASE_URL    = "https://api.deepseek.com/chat/completions"   # 接口地址
MODEL       = "deepseek-chat"                                # 模型名
API_KEY_ENV = "DEEPSEEK_API_KEY"                             # 密钥环境变量名
```

凡是 OpenAI 兼容接口都能直接换,常用例子(文件注释里也有):

| 服务商 | BASE_URL | MODEL 例子 | 密钥环境变量 |
|---|---|---|---|
| DeepSeek | https://api.deepseek.com/chat/completions | deepseek-flash / deepseek-v4-pro | DEEPSEEK_API_KEY |
| 月之暗面 Kimi | https://api.moonshot.cn/v1/chat/completions | kimi-k2-turbo | MOONSHOT_API_KEY |
| 阿里通义千问 | https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions | qwen-plus / qwen-max | DASHSCOPE_API_KEY |
| 智谱 GLM | https://open.bigmodel.cn/api/paas/v4/chat/completions | glm-4.5 | ZHIPUAI_API_KEY |
| 本地 Ollama | http://127.0.0.1:11434/v1/chat/completions | qwen2.5:14b | 留空 "" |

两点注意:
1. 所选模型必须**支持 function calling(工具调用)**,否则智能体无法操作排程工具;
   Ollama 建议 qwen2.5 / llama3.1 及以上,且需能联网下载模型。
2. 不改文件时也可以临时指定:启动命令加 `--model <模型名>`(如 `python webui.py --model deepseek-reasoner`)。

说明:PoC 默认读取原脚本的数据目录(`C:\Users\mblon\Desktop\project\data` / `rule` / `工模具`),
导出 Excel 到本目录 `out\`,不触碰原项目任何文件。可用 `--input/--rule/--tooling/--out` 覆盖。

## 对话示例

LLM 模式下直接说人话(在 `你 >` 后面输入,例如):

```
现在排程的整体情况怎么样
把订单 2326030045001 提到最前,看看对整体有什么影响
解释订单 1126040218013 为什么这样排
给急催订单加优先,然后和默认排程对比
延迟 2 小时(模拟开工前检修)
导出排程表
```

离线模式(未设置密钥)用下面的模板指令:

```
你 > 急催                                # 给69个急催单加优先权重并重排
你 > KPI                                 # 看效果:超20格边数 17→22,完工时间 -193分钟
你 > 对比                                # 与默认排程逐项对比(这就是智能体要报告的取舍)
你 > 解释 2326030045001                  # 该单为什么这样排、几点进炉、用什么喷嘴
你 > 把 1126040215003 提到最前           # 人工插单,引擎重排其余订单并校验
你 > 延迟 2 小时                         # 模拟开工前检修,整线顺延
你 > 改 1126040215003 回火温度 620       # 工艺变更重排
你 > 移除 1126010208032                  # 完工/取消移出
你 > 撤销改动                            # 回到默认排程
你 > 导出                                # 输出 Excel(排产总览/原始订单/钢管级时间轴)
```

## 与真实数据跑出来的基线(150单)

| 指标 | 默认排程 |
|---|---|
| 正常生产 / 暂缓 | 135 / 15 单 |
| 相邻空格总数 / 平均 | 3185 / 23.77 格 |
| 超20格相邻边 | 17(其中硬约束违规 6、引擎允许的跨温区接续 11) |
| 急催单 | 排入 64、暂缓 5 |
| 完工时间 | 第7天 20:17 |

有意思的实测:给急催单加优先后,完工时间提前了约 3.2 小时,但超 20 格边数从 17 升到 22、
暂缓单 +2 —— 这正是智能体存在的价值:**每个改动都有代价,由它量化并报告给你**。

## 设计要点

1. **原内核不动**:`engine.py` 通过 `sys.path` 直接 `import` 原脚本,零复制、零漂移;你以后
   更新原脚本的规则,智能体自动跟随。
2. **验证器分流**:`validate` 区分"违规"(原引擎 edge_ok 明确拒绝,主要来自恢复大批量等修复阶段)
   与"必要换规"(超20格但属于引擎允许的跨温区接续),避免把合法排程误报为错误。
3. **改动可回滚**:每次动作都基于"基线重跑"(overrides 机制),`撤销改动` 一键回到默认排程;
   每个动作返回"相对上一步变化",LLM 据此判断改好了还是改坏了。
4. **锁定即约束**:move/pin 实现为"锁定前缀 + 贪心重排其余",锁定本身不绕过校验——
   锁定造成违规时 validate 会如实报告。
5. **输出可读性**:导出 Excel 新增带颜色的"排程状态"列(绿=连续、黄=必要换规、红=暂缓等),
   "本单开始原因"只写分类标签:连续生产/首单只写标签,大批量恢复仅组内首单写吨位
   (后续标连续生产),换规只写"正火等X步/回火等Y步"(规则编号、中间等待等细节在备注列),
   一眼可扫;颜色含义见"图例"工作表。

## 当前限制(诚实声明)

- **What-if 能力**:延迟仅支持"开工前"整线顺延(时间轴整体平移);支持移除/恢复订单、
  回火温度调整、移动/锁定/急催加权。中途检修(第3天前炉停2小时)需要给 schedule_times
  加窗口模型,是下一步。
- **优化内核**:贪心 + OR-Tools 初排仍是原脚本的策略;若要进一步压空格总数/换规次数,
  可把 build_sequence 换成 CP-SAT(见路线图)。
- **LLM 模式**需要 `DEEPSEEK_API_KEY`;离线模式只能识别模板化指令,不会自由发挥。

## 下一步路线

1. 把 `validate` 违规项作为 LLM 的硬反馈,做"自动修复循环"(LLM 提议 → 引擎裁决 → 违规回喂)。
2. 检修窗口模型:把"前炉第N天停M小时"建模进 schedule_times。
3. 优化内核升级:CP-SAT 多目标(空格+换规+急催+完工),保留现有管线作对比基线。
4. Web 界面:把动作 API 暴露成 HTTP,做网页版排程台(表格 + 甘特图 + 对话窗)。
