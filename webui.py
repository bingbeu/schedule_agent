# -*- coding: utf-8 -*-
"""热处理排程智能体 · 本地网页版(零依赖,仅用 Python 标准库)。

运行:
    python webui.py                # 浏览器打开 http://127.0.0.1:8787
    python webui.py --port 9000    # 换端口
    python webui.py --no-llm       # 强制离线指令模式

有 DEEPSEEK_API_KEY 时自动启用大模型对话;密钥无效/网络不通时自动退回离线指令模式。
"""
from __future__ import annotations

import argparse
import json
import os
import socket
import threading
import time
import urllib.error
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

from actions import SchedulerAgent
from agent_cli import SYSTEM_PROMPT, call_llm, dispatch, llm_preflight, parse_offline, wrap_tools
from engine import Engine, diff_kpis
from llm_config import API_KEY_ENV, MODEL as DEFAULT_MODEL

MAX_LLM_ROUNDS = 6
ALLOWED_ACTIONS = {
    "kpi", "validate", "compare", "status", "list_orders", "explain",
    "move", "pin", "boost", "unboost", "temper_change", "delay",
    "remove_orders", "restore_orders", "reset_changes", "export",
}

engine = Engine()
agent = SchedulerAgent(engine)
LOCK = threading.Lock()
TURN_LOCK = threading.Lock()  # 同一时间只处理一条对话,防止连点导致状态交错
API_KEY = ""
MODEL = DEFAULT_MODEL
OUT_DIR = Path(__file__).resolve().parent / "out"


# ---------------------------------------------------------------- 业务逻辑
def sanitize(obj):
    """递归清洗,保证输出是合法 JSON(去掉 NaN/Inf)。"""
    if isinstance(obj, dict):
        return {k: sanitize(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [sanitize(v) for v in obj]
    if isinstance(obj, float) and (obj != obj or obj in (float("inf"), float("-inf"))):
        return None
    return obj


def _json_safe(o):
    """JSON 兜底:其余未知对象转字符串。"""
    return str(o)


def summarize_offline(res) -> str:
    """把离线动作结果压缩成一句人话。"""
    if not isinstance(res, dict):
        return str(res)
    if "error" in res:
        return f"⚠️ {res['error']}"
    # explain 结果
    if "订单编号" in res and ("本单开始原因" in res or "原因" in res):
        stage = res.get("阶段", "")
        if "暂缓" in stage:
            return f"{res['订单编号']} 当前{stage}:{res.get('原因', '')}"
        return (f"{res['订单编号']}({stage},第{res.get('序号', '?')}单):前炉{res.get('前炉温度')}℃/"
                f"回火{res.get('回火温度')}℃,{res.get('本单开始原因', '')};"
                f"前炉首支{res.get('前炉首支进炉', '')},回火首支{res.get('回火首支进炉', '')},"
                f"喷嘴{res.get('喷嘴规格', '')}")
    # list_orders 结果
    if "订单" in res and "命中" in res:
        ids = "、".join(str(o.get("订单编号", "")) for o in (res.get("订单") or [])[:5])
        return f"命中 {res['命中']} 单,例如:{ids}" + (" …" if res["命中"] > 5 else "")
    # validate 结果
    if "硬约束违规数" in res and "明细" in res:
        det = res.get("明细") or []
        first = f"首项:位置{det[0]['位置']} {det[0]['订单编号']} {det[0]['问题']}" if det else "无问题"
        return f"硬约束违规 {res['硬约束违规数']} 项,必要换规 {res.get('必要换规边数', 0)} 项;{first}"
    # compare 结果
    if "默认KPI" in res and "当前KPI" in res:
        b, n = res["默认KPI"], res["当前KPI"]
        d = res.get("相对默认变化") or {}
        items = "、".join(f"{k}:{v['基准']}→{v['当前']}" for k, v in list(d.items())[:6])
        return (f"默认:正常{b.get('正常生产单数')}单/暂缓{b.get('暂缓/剔除单数')}单,"
                f"空格{b.get('相邻空格总数')},完工{b.get('完工时间')};"
                f"当前:正常{n.get('正常生产单数')}单/暂缓{n.get('暂缓/剔除单数')}单,"
                f"空格{n.get('相邻空格总数')},完工{n.get('完工时间')}"
                + (f";变化:{items}" if items else ";与默认一致"))
    # export 结果
    if "文件" in res:
        return f"已导出:{res['文件']}(正常{res.get('正常生产')}单/暂缓{res.get('暂缓')}单)"
    # status 结果
    if "生效改动" in res:
        ov = res.get("生效改动") or {}
        txt = "、".join(f"{k}={v}" for k, v in list(ov.items())[:6])
        return (f"当前改动:{txt or '无'};正常生产{res.get('正常生产')}单,"
                f"暂缓{res.get('暂缓')}单,硬约束违规{res.get('硬约束违规数')}项")
    # 通用:说明 + KPI
    parts = []
    if "说明" in res:
        parts.append(res["说明"])
    k = res.get("当前KPI")
    if k:
        parts.append(f"正常{k.get('正常生产单数')}单/暂缓{k.get('暂缓/剔除单数')}单,"
                     f"相邻空格总数{k.get('相邻空格总数')},硬约束违规{k.get('硬约束违规数')},"
                     f"完工{k.get('完工时间')}")
    d = res.get("相对默认排程变化") or res.get("相对上一步变化")
    if d:
        items = "、".join(f"{kk}:{vv['基准']}→{vv['当前']}" for kk, vv in list(d.items())[:6])
        parts.append(f"变化:{items}")
    out = ";".join(p for p in parts if p)
    return out or "已执行,详见工具记录。"


def run_turn(text: str) -> tuple[str, list]:
    """执行一轮对话,返回 (助手文字, 工具步骤列表)。"""
    text = (text or "").strip()
    if not text:
        return "请输入内容。", []
    steps = []

    if API_KEY:
        msgs = [{"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": text}]
        final = ""
        for _ in range(MAX_LLM_ROUNDS):
            try:
                body = call_llm(msgs, wrap_tools(), API_KEY, MODEL)
            except urllib.error.HTTPError as e:
                try:
                    detail = e.read().decode("utf-8", "replace")[:300]
                except Exception:  # noqa: BLE001
                    detail = e.reason
                return f"⚠️ API错误 HTTP {e.code}: {detail}", steps
            except Exception as e:  # noqa: BLE001
                return f"⚠️ LLM调用失败:{type(e).__name__}: {e}", steps
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
                    kwargs = {}
                with LOCK:
                    res = dispatch(agent, fn["name"], kwargs)
                steps.append({"name": fn["name"], "args": kwargs, "result": res})
                msgs.append({"role": "tool", "tool_call_id": tc["id"],
                             "content": json.dumps(res, ensure_ascii=False, default=_json_safe)})
        else:
            final = "已达到最大工具轮数,请查看工具执行记录。"
        return final or "(无文字回复)", steps

    # 离线指令模式
    (tool, kwargs), hint = parse_offline(text)
    if tool is None:
        return hint, []
    with LOCK:
        res = dispatch(agent, tool, kwargs)
    steps.append({"name": tool, "args": kwargs, "result": res})
    return summarize_offline(res), steps


def state_payload() -> dict:
    """当前排程状态(供前端渲染)。"""
    with LOCK:
        cur, base = agent.current, agent.baseline
        v = sum(1 for i in cur.issues if i.get("类别") == "违规")
        c = sum(1 for i in cur.issues if i.get("类别") == "必要换规")
        return {
            "mode": "LLM" if API_KEY else "离线指令",
            "main": len(cur.main), "deferred": len(cur.deferred),
            "kpi": cur.kpi, "baseline_kpi": base.kpi,
            "diff": diff_kpis(base.kpi, cur.kpi),
            "violations": [i for i in cur.issues if i.get("类别") == "违规"],
            "crosses": [i for i in cur.issues if i.get("类别") == "必要换规"],
            "risks": [i for i in cur.issues if i.get("类别") == "堵炉风险"],
            "violation_count": v, "cross_count": c,
            "risk_count": sum(1 for i in cur.issues if i.get("类别") == "堵炉风险"),
            "overrides": {k: (sorted(x) if isinstance(x, set) else x) for k, x in agent.active.items()},
        }


# ---------------------------------------------------------------- HTTP 服务
class Handler(BaseHTTPRequestHandler):
    server_version = "ScheduleAgent/1.0"

    def _json(self, obj, code=200):
        data = json.dumps(sanitize(obj), ensure_ascii=False, default=_json_safe).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _html(self, body, code=200):
        data = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        try:
            self._route_get()
        except Exception as e:  # noqa: BLE001
            try:
                self._json({"error": f"服务器内部错误:{type(e).__name__}: {e}"}, 500)
            except Exception:  # noqa: BLE001
                pass

    def do_POST(self):
        try:
            self._route_post()
        except Exception as e:  # noqa: BLE001
            try:
                self._json({"error": f"服务器内部错误:{type(e).__name__}: {e}"}, 500)
            except Exception:  # noqa: BLE001
                pass

    def _route_get(self):
        u = urlparse(self.path)
        if u.path == "/":
            self._html(INDEX_HTML)
            return
        if u.path == "/api/state":
            self._json(state_payload())
            return
        if u.path == "/api/orders":
            q = parse_qs(u.query)
            filt = (q.get("filter") or [""])[0]
            with LOCK:
                res = agent.list_orders(filt)
            self._json(res)
            return
        if u.path == "/api/export":
            path = OUT_DIR / f"热处理排程_网页版_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
            try:
                with LOCK:
                    p = agent.current.export(path)
                data = p.read_bytes()
            except Exception as e:  # noqa: BLE001
                self._json({"error": str(e)}, 500)
                return
            ascii_name = f"schedule_agent_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition",
                             f'attachment; filename="{ascii_name}"; filename*=UTF-8\'\'{quote(p.name)}')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        self._json({"error": "not found"}, 404)

    def _route_post(self):
        u = urlparse(self.path)
        length = int(self.headers.get("Content-Length") or 0)
        try:
            payload = json.loads(self.rfile.read(length).decode("utf-8")) if length else {}
        except Exception:  # noqa: BLE001
            payload = {}
        if u.path == "/api/chat":
            if not TURN_LOCK.acquire(blocking=False):
                self._json({"final": "⏳ 上一条请求还在处理中,请等它出结果后再发送。",
                            "steps": [], "state": state_payload()})
                return
            try:
                final, steps = run_turn(str(payload.get("text", "")))
            finally:
                TURN_LOCK.release()
            self._json({"final": final, "steps": steps, "state": state_payload()})
            return
        if u.path == "/api/action":
            name = str(payload.get("name", ""))
            kwargs = payload.get("kwargs") or {}
            if name not in ALLOWED_ACTIONS:
                self._json({"error": f"未知动作:{name}"}, 400)
                return
            if not isinstance(kwargs, dict):
                kwargs = {}
            with LOCK:
                res = dispatch(agent, name, kwargs)
            self._json({"result": res, "state": state_payload()})
            return
        if u.path == "/api/reset":
            with LOCK:
                res = agent.reset()
            self._json({"result": res, "state": state_payload()})
            return
        self._json({"error": "not found"}, 404)

    def log_message(self, fmt, *args):  # 静默访问日志
        pass


# ---------------------------------------------------------------- 前端页面
INDEX_HTML = """<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>热处理排程智能体 · 630厂</title>
<style>
:root{--bg:#0f172a;--panel:#1e293b;--panel2:#273449;--line:#334155;--fg:#e2e8f0;
--muted:#94a3b8;--accent:#38bdf8;--ok:#34d399;--warn:#fbbf24;--bad:#f87171;--user:#0c4a6e}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font-family:"Segoe UI","Microsoft YaHei",sans-serif;height:100vh;display:flex;flex-direction:column}
header{display:flex;align-items:center;justify-content:space-between;padding:10px 16px;background:var(--panel);border-bottom:1px solid var(--line)}
.brand{font-size:17px;font-weight:700}
.badge{display:inline-block;margin-left:8px;padding:2px 10px;border-radius:10px;font-size:12px;font-weight:600;background:#334155}
.badge.llm{background:#0e7490}
.headbtns button,button{background:var(--panel2);color:var(--fg);border:1px solid var(--line);border-radius:8px;padding:6px 12px;cursor:pointer;font-size:13px}
button:hover{border-color:var(--accent)}
button.primary{background:#0e7490;border-color:#0e7490}
main{flex:1;display:grid;grid-template-columns:1fr 400px;gap:12px;padding:12px;min-height:0}
#chatPane{display:flex;flex-direction:column;background:var(--panel);border:1px solid var(--line);border-radius:12px;min-height:0}
#messages{flex:1;overflow-y:auto;padding:14px;display:flex;flex-direction:column;gap:10px}
.msg{max-width:78%;padding:10px 13px;border-radius:12px;white-space:pre-wrap;word-break:break-word;font-size:14px;line-height:1.55}
.msg.user{align-self:flex-end;background:var(--user);border:1px solid #155e75}
.msg.bot{align-self:flex-start;background:#16213b;border:1px solid var(--line)}
.msg.err{align-self:flex-start;background:#3b1621;border:1px solid #7f1d1d}
.step{margin-top:8px;border:1px dashed var(--line);border-radius:8px;font-size:12px;background:#0b1526}
.step summary{cursor:pointer;padding:5px 9px;color:var(--accent);font-family:Consolas,monospace}
.step pre{margin:0;padding:6px 9px;max-height:180px;overflow:auto;color:var(--muted);white-space:pre-wrap;font-family:Consolas,monospace}
#quick{display:flex;flex-wrap:wrap;gap:6px;padding:8px 14px;border-top:1px solid var(--line)}
.chip{background:var(--panel2);border:1px solid var(--line);color:var(--muted);border-radius:14px;padding:4px 10px;font-size:12px;cursor:pointer}
.chip:hover{color:var(--accent);border-color:var(--accent)}
#chatForm{display:flex;gap:8px;padding:12px;border-top:1px solid var(--line)}
#chatInput{flex:1;background:#0b1526;border:1px solid var(--line);border-radius:10px;color:var(--fg);padding:10px 12px;font-size:14px}
#chatInput:focus{outline:none;border-color:var(--accent)}
#sidePane{overflow-y:auto;display:flex;flex-direction:column;gap:10px}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:12px}
.panel h3{margin:0 0 8px;font-size:14px;color:var(--accent);display:flex;justify-content:space-between;align-items:center}
table{width:100%;border-collapse:collapse;font-size:12.5px}
td,th{padding:4px 6px;border-bottom:1px solid #26344d;text-align:left}
th{color:var(--muted);font-weight:600}
.delta{color:var(--muted);font-size:11px}
.delta.up{color:var(--bad)}.delta.down{color:var(--ok)}
.issue{border-left:3px solid var(--bad);background:#241621;padding:6px 8px;border-radius:6px;font-size:12px;margin-bottom:6px}
.issue.cross{border-left-color:var(--warn);background:#242014}
.issue.risk{border-left-color:var(--accent);background:#16213b}
.issue b{color:#fca5a5}.issue.cross b{color:#fcd34d}.issue.risk b{color:#7dd3fc}
.row{display:flex;gap:6px;margin-bottom:8px}
.row input{flex:1;background:#0b1526;border:1px solid var(--line);border-radius:8px;color:var(--fg);padding:6px 9px;font-size:12.5px}
#orderList{max-height:300px;overflow:auto;font-size:12px}
.ord{border-bottom:1px solid #26344d;padding:5px 2px}
.ord .id{color:var(--accent);font-family:Consolas,monospace}
.ord .meta{color:var(--muted);font-size:11px}
#busy{display:none;position:fixed;top:56px;left:50%;transform:translateX(-50%);background:#0e7490;border:1px solid #38bdf8;border-radius:20px;padding:7px 18px;font-size:13px;z-index:9}
#modeHint{font-size:12px;color:var(--muted);padding:0 14px 8px;line-height:1.6}
@media(max-width:900px){main{grid-template-columns:1fr}#sidePane{display:none}}
</style>
</head>
<body>
<header>
  <div class="brand">🔥 热处理排程智能体 · 630厂 <span id="modeBadge" class="badge">...</span></div>
  <div class="headbtns">
    <button id="btnReset">重新加载数据</button>
    <button id="btnExport" class="primary">导出 Excel</button>
  </div>
</header>
<main>
  <section id="chatPane">
    <div id="messages"></div>
    <div id="modeHint"></div>
    <div id="quick"></div>
    <form id="chatForm" autocomplete="off">
      <input id="chatInput" placeholder="直接说需求,例如:把订单 2326030045001 提到最前 / 现在排程整体情况怎么样 / 给急催订单加优先并对比 / 延迟 2 小时 / 导出排程表">
      <button type="submit" class="primary">发送</button>
    </form>
  </section>
  <aside id="sidePane">
    <div class="panel">
      <h3>排程 KPI <button id="btnRefresh">刷新</button></h3>
      <table id="kpiTable"></table>
    </div>
    <div class="panel">
      <h3>硬约束违规(<span id="violCount">0</span>)</h3>
      <div id="violList"></div>
      <h3 style="margin-top:10px">必要换规(<span id="crossCount">0</span>,引擎允许)</h3>
      <div id="crossList"></div>
      <h3 style="margin-top:10px">堵炉风险提示(<span id="riskCount">0</span>)</h3>
      <div id="riskList"></div>
    </div>
    <div class="panel">
      <h3>订单速查</h3>
      <div class="row"><input id="orderFilter" placeholder="编号/牌号/主体厂关键字"><button id="btnOrders">查询</button></div>
      <div id="orderList"></div>
    </div>
  </aside>
</main>
<div id="busy">处理中…</div>
<script>
const $ = id => document.getElementById(id);
let state = null;
let busy = false;
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const GOOD_DOWN = new Set(['相邻空格总数','平均相邻空格','超20格相邻边数','最大相邻空格','硬约束违规数','温度回摆次数','暂缓/剔除单数','暂缓吨位','完工分钟','急催暂缓数']);

function setBusy(b, txt){
  const el = $('busy');
  el.textContent = txt || '处理中…';
  el.style.display = b ? 'block' : 'none';
  $('chatForm').querySelector('button').disabled = b;
}

function addMsg(role, text, steps){
  const div = document.createElement('div');
  div.className = 'msg ' + (role === 'user' ? 'user' : (text && text.startsWith('⚠️') ? 'err' : 'bot'));
  div.textContent = text || '';
  if (steps && steps.length){
    steps.forEach(st => {
      const d = document.createElement('details');
      d.className = 'step';
      const args = Object.keys(st.args || {}).length ? ' ' + JSON.stringify(st.args) : '';
      d.innerHTML = '<summary>🔧 ' + esc(st.name) + esc(args) + '</summary><pre>' + esc(JSON.stringify(st.result, null, 2)) + '</pre>';
      div.appendChild(d);
    });
  }
  $('messages').appendChild(div);
  $('messages').scrollTop = $('messages').scrollHeight;
}

function renderState(s){
  state = s;
  $('modeBadge').textContent = s.mode + '模式';
  $('modeBadge').className = 'badge' + (s.mode === 'LLM' ? ' llm' : '');
  $('modeHint').innerHTML = s.mode === 'LLM'
    ? '当前为大模型模式:直接说人话,智能体会自己调用排程工具并汇报结果。'
    : '当前为<strong>离线指令模式</strong>(未设置 DEEPSEEK_API_KEY)。可输入模板指令:KPI / 校验 / 对比 / 急催 / 解释 &lt;订单号&gt; / 把 &lt;订单号&gt; 提到最前 / 延迟 2 小时 / 撤销改动 / 导出。';
  const kpi = s.kpi, diff = s.diff || {};
  let html = '<tr><th>指标</th><th>当前</th><th>相对默认</th></tr>';
  for (const [k, v] of Object.entries(kpi)){
    const d = diff[k];
    let delta = '';
    if (d){
      const dir = d.变化 > 0 ? 'up' : (d.变化 < 0 ? 'down' : '');
      const good = GOOD_DOWN.has(k) ? (d.变化 < 0) : (d.变化 > 0);
      const cls = d.变化 === 0 ? '' : (good ? 'down' : 'up');
      delta = '<span class="delta ' + cls + '">(' + (d.变化 > 0 ? '+' : '') + d.变化 + ')</span>';
    }
    html += '<tr><td>' + esc(k) + '</td><td>' + esc(v) + '</td><td>' + delta + '</td></tr>';
  }
  $('kpiTable').innerHTML = html;
  $('violCount').textContent = s.violation_count;
  $('violList').innerHTML = (s.violations || []).slice(0, 8).map(i =>
    '<div class="issue"><b>#' + i.位置 + ' ' + esc(i.订单编号) + '</b><br>' + esc(i.问题) + '</div>').join('') || '<div style="color:var(--ok)">✅ 无违规</div>';
  $('crossCount').textContent = s.cross_count;
  $('crossList').innerHTML = (s.crosses || []).slice(0, 5).map(i =>
    '<div class="issue cross"><b>#' + i.位置 + ' ' + esc(i.订单编号) + '</b><br>' + esc(i.问题) + '</div>').join('') || '<div style="color:var(--muted)">—</div>';
  $('riskCount').textContent = s.risk_count;
  $('riskList').innerHTML = (s.risks || []).slice(0, 5).map(i =>
    '<div class="issue risk"><b>#' + i.位置 + ' ' + esc(i.订单编号) + '</b><br>' + esc(i.问题) + '</div>').join('') || '<div style="color:var(--muted)">—</div>';
}

async function refreshState(){
  try{
    const r = await fetch('/api/state');
    renderState(await r.json());
  }catch(e){ console.error(e); }
}

async function sendChat(text){
  if (busy) return;
  if (!text || !text.trim()) return;
  busy = true;
  addMsg('user', text);
  setBusy(true, '处理中…(大模型模式可能需几十秒)');
  try{
    const r = await fetch('/api/chat', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({text})});
    const data = await r.json();
    addMsg('bot', data.final, data.steps);
    renderState(data.state);
  }catch(e){
    addMsg('bot', '⚠️ 请求失败:' + e, []);
  }finally{
    busy = false;
    setBusy(false);
  }
}

$('chatForm').addEventListener('submit', e => { e.preventDefault(); sendChat($('chatInput').value); $('chatInput').value=''; });
$('btnRefresh').onclick = refreshState;
$('btnExport').onclick = () => window.location.href = '/api/export';
$('btnReset').onclick = async () => {
  setBusy(true, '重新加载数据并生成默认排程…');
  try{
    const r = await fetch('/api/reset', {method:'POST'});
    const data = await r.json();
    addMsg('bot', data.result.说明 || '已重置', []);
    renderState(data.state);
  }finally{ setBusy(false); }
};
$('btnOrders').onclick = $('orderFilter').onkeydown = async ev => {
  if (ev && ev.key !== 'Enter') return;
  const f = $('orderFilter').value.trim();
  const r = await fetch('/api/orders?filter=' + encodeURIComponent(f));
  const data = await r.json();
  $('orderList').innerHTML = (data.订单 || []).map(o =>
    '<div class="ord"><span class="id">' + esc(o.订单编号) + '</span> ' +
    '<span class="meta">' + esc(o.主体厂) + ' | ' + esc(o.牌号) + ' | Φ' + esc(o.外径) + '×' + esc(o.壁厚) +
    ' | 前炉' + esc(o.前炉温度) + '℃/回火' + esc(o.回火温度) + '℃ | ' + esc(o.数量) + '支' +
    (o.急催 ? ' | 🔴急催' : '') + '</span></div>').join('') || '<div style="color:var(--muted)">无匹配(' + (data.命中 || 0) + ')</div>';
};
const QUICK = ['现在排程整体情况怎么样','KPI','校验','对比','急催','撤销改动','列出订单'];
$('quick').innerHTML = QUICK.map(q => '<span class="chip">' + q + '</span>').join('');
document.querySelectorAll('#quick .chip').forEach(el => el.onclick = () => sendChat(el.textContent));

refreshState();
$('orderFilter').dispatchEvent(new KeyboardEvent('keydown', {key:'Enter'}));
</script>
</body>
</html>
"""


# ---------------------------------------------------------------- 入口
def _port_in_use(host: str, port: int) -> bool:
    """探测端口是否已有服务在监听(Windows 上双进程可同绑端口,必须主动探测)。"""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(1.0)
    try:
        s.connect((host, port))
        return True
    except OSError:
        return False
    finally:
        s.close()


def main():
    global engine, agent, API_KEY, MODEL, OUT_DIR
    import sys
    try:
        sys.stdout.reconfigure(line_buffering=True)
    except Exception:  # noqa: BLE001
        pass
    ap = argparse.ArgumentParser(description="热处理排程智能体 · 本地网页版")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8787)
    ap.add_argument("--input", type=Path, default=None, help="订单数据目录(默认原脚本 data)")
    ap.add_argument("--rule", type=Path, default=None, help="换规规则目录(默认原脚本 rule)")
    ap.add_argument("--tooling", type=Path, default=None, help="工模具目录(默认原脚本工模具)")
    ap.add_argument("--out", type=Path, default=None, help="Excel 输出目录(默认 out/)")
    ap.add_argument("--model", default=DEFAULT_MODEL, help=f"模型名(默认取 llm_config.py:{DEFAULT_MODEL})")
    ap.add_argument("--no-llm", action="store_true", help="强制离线指令模式")
    args = ap.parse_args()

    MODEL = args.model
    OUT_DIR = Path(args.out) if args.out else Path(__file__).resolve().parent / "out"
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    engine = Engine(args.input, args.rule, args.tooling)
    agent = SchedulerAgent(engine, OUT_DIR)
    API_KEY = "" if args.no_llm else os.environ.get(API_KEY_ENV, "").strip()

    print("正在加载数据并生成默认排程 ...")
    agent.reset()
    st = state_payload()
    print(f"载入 {st['kpi']['订单总数']} 单:正常生产 {st['main']} 单,暂缓 {st['deferred']} 单;"
          f"完工 {st['kpi']['完工时间']}")

    if API_KEY:
        print("[预检] 正在检查 DeepSeek API 连通性 ...")
        ok, diag = llm_preflight(API_KEY, MODEL)
        if ok:
            print("[预检] 连接正常,网页将以大模型模式对话。")
        else:
            print(f"[警告] LLM 预检失败:{diag}")
            print("[提示] 网页自动退回离线指令模式;修复密钥后重启即可。")
            API_KEY = ""
    else:
        print(f"[提示] 未设置 {API_KEY_ENV},网页为离线指令模式;"
              f"设置该环境变量后重启即启用大模型对话(服务商/模型可在 llm_config.py 修改)。")

    if _port_in_use(args.host, args.port):
        print(f"❌ 端口 {args.port} 已被占用:很可能有一个旧的 webui 实例还在运行,")
        print("   浏览器实际连到的是那个旧实例,页面显示的模式也会是旧实例的。")
        print("   处理办法:关闭旧终端窗口(或在其终端按 Ctrl+C),再重新运行本命令;")
        print(f"   或者换一个端口:python webui.py --port 9000")
        raise SystemExit(1)

    print(f"[模式] 本次启动:{'大模型对话(LLM)' if API_KEY else '离线指令'};页面右上角徽标应显示相同模式。")
    srv = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"\n✅ 已启动:http://{args.host}:{args.port}")
    print("   在浏览器打开上面地址即可使用;按 Ctrl+C 停止。")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")


if __name__ == "__main__":
    main()
