#!/usr/bin/env python3
"""分箱单生成工具 — Web GUI (零依赖, 浏览器访问)"""

import os
import sys
import json
import shutil
import threading
import queue
import webbrowser
import email.parser
import email.policy
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, quote

# PyInstaller 支持: 打包后资源在 _MEIPASS, 输出路径在可执行文件所在目录
if getattr(sys, "frozen", False):
    BUNDLE_DIR = sys._MEIPASS
    WORK_DIR = os.path.dirname(sys.executable)
else:
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))
    WORK_DIR = BUNDLE_DIR

sys.path.insert(0, BUNDLE_DIR)

import generate_packing_list as gpl

PORT = 8090

# ── Shared state ─────────────────────────────────────────
progress_queue = queue.Queue()
progress_log = []
progress_state = {"current": 0, "total": 0}
generation_done = threading.Event()
generation_error_msg = None
output_path = os.path.join(WORK_DIR, "分箱单_生成结果.xlsx")

# 模板始终使用默认路径, 其他文件由用户上传
uploaded_files = {"zd": None, "sp": None}  # key → filename on disk

TPL_PATH = os.path.join(BUNDLE_DIR, "template.xlsx")
TPL_EXISTS = os.path.isfile(TPL_PATH)

# ── HTML page ─────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>分箱单生成工具</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
  background: #f0f2f5; color: #333; display: flex; justify-content: center; padding: 40px 16px;
}
.card {
  background: #fff; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,.08);
  max-width: 680px; width: 100%; padding: 32px 28px;
}
h1 { font-size: 22px; font-weight: 700; text-align: center; margin-bottom: 28px; color: #1a1a1a; }

.file-row {
  display: flex; align-items: center; gap: 10px;
  padding: 10px 0; border-bottom: 1px solid #f0f0f0;
}
.file-row:last-of-type { border-bottom: none; }
.file-label { min-width: 72px; font-size: 14px; color: #555; }
.file-name {
  flex: 1; font-size: 13px; color: #999; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.file-name.exists { color: #333; font-weight: 500; }
.file-status { font-size: 13px; min-width: 42px; text-align: right; }
.file-status.ok { color: #52c41a; }
.file-status.missing { color: #ff4d4f; }
.upload-btn {
  padding: 5px 12px; font-size: 13px; background: #f5f5f5; border: 1px solid #d9d9d9;
  border-radius: 6px; cursor: pointer; color: #555; white-space: nowrap;
}
.upload-btn:hover { border-color: #4a90d9; color: #4a90d9; }
.upload-btn.primary {
  background: #4a90d9; color: #fff; border-color: #4a90d9;
}
.upload-btn.primary:hover { background: #357abd; }
.required-hint { font-size: 12px; color: #ff4d4f; min-width: 48px; text-align: right; }

.btn-row { text-align: center; margin: 24px 0 16px; }
#btn-generate {
  padding: 10px 48px; font-size: 16px; font-weight: 600;
  background: #4a90d9; color: #fff; border: none; border-radius: 8px; cursor: pointer;
}
#btn-generate:hover { background: #357abd; }
#btn-generate:disabled { background: #a0c4e8; cursor: not-allowed; }

.progress-wrap { margin-bottom: 8px; }
.progress-bar {
  width: 100%; height: 8px; background: #e8e8e8; border-radius: 4px; overflow: hidden;
}
.progress-fill {
  height: 100%; width: 0%; background: #4a90d9; border-radius: 4px; transition: width .3s;
}
.progress-text { text-align: center; font-size: 13px; color: #888; margin-top: 4px; }

.status-text { text-align: center; font-size: 14px; color: #666; margin-bottom: 12px; }

.log-area {
  background: #fafafa; border: 1px solid #eee; border-radius: 8px;
  height: 200px; overflow-y: auto; padding: 10px 14px;
  font-family: "SF Mono", "Menlo", "Consolas", monospace; font-size: 12px;
  line-height: 1.7; color: #555;
}

.download-wrap { text-align: center; margin-top: 16px; display: none; }
#btn-download {
  padding: 8px 24px; font-size: 14px; background: #52c41a; color: #fff;
  border: none; border-radius: 6px; cursor: pointer; text-decoration: none; display: inline-block;
}
#btn-download:hover { background: #45a614; }
</style>
</head>
<body>
<div class="card">
  <h1>分箱单生成工具</h1>

  <div id="file-rows"></div>

  <div class="btn-row">
    <button id="btn-generate" onclick="startGenerate()">生成分箱单</button>
  </div>

  <div class="progress-wrap">
    <div class="progress-bar"><div class="progress-fill" id="progress-fill"></div></div>
    <div class="progress-text" id="progress-text">就绪</div>
  </div>

  <div class="status-text" id="status-text"></div>

  <div class="log-area" id="log-area"></div>

  <div class="download-wrap" id="download-wrap">
    <a id="btn-download" href="/api/download" download>下载生成结果</a>
  </div>
</div>

<script>
const FILE_TYPES = [
  { key: 'tpl', label: '模板文件', defaultName: 'template.xlsx', accept: '.xlsx', optional: false },
  { key: 'zd',  label: '总单文件', defaultName: '总单.xlsx',   accept: '.xlsx', optional: false },
  { key: 'sp',  label: '备件清单', defaultName: '附件备件清单.xls', accept: '.xls,.xlsx', optional: false },
];

let pollTimer = null;
let fileStates = {}; // key → { name, exists }

function checkAllReady() {
  let allReady = true;
  FILE_TYPES.forEach(ft => {
    const info = fileStates[ft.key] || {};
    if (!info.exists) allReady = false;
  });
  document.getElementById('btn-generate').disabled = !allReady;
}

// ── Build file rows ──
function buildRows(files) {
  fileStates = files;
  const container = document.getElementById('file-rows');
  container.innerHTML = '';
  FILE_TYPES.forEach(ft => {
    const info = files[ft.key] || { name: ft.defaultName, exists: false };
    const isReady = info.exists;
    const div = document.createElement('div');
    div.className = 'file-row';
    div.innerHTML =
      `<span class="file-label">${ft.label}</span>
       <span class="file-name ${isReady ? 'exists' : ''}" id="fn-${ft.key}">${info.name}</span>
       ${isReady
         ? '<span class="file-status ok" id="fs-' + ft.key + '">✓</span>'
         : '<span class="required-hint" id="fs-' + ft.key + '">请上传</span>'}
       <button class="upload-btn ${isReady ? '' : 'primary'}" id="ub-${ft.key}"
               onclick="uploadFile('${ft.key}', '${ft.accept}')">${isReady ? '更换' : '选择文件'}</button>
       <input type="file" id="input-${ft.key}" accept="${ft.accept}" style="display:none"
              onchange="doUpload('${ft.key}', this)">`;
    container.appendChild(div);
  });
  checkAllReady();
}

// ── Init ──
function init() {
  fetch('/api/files').then(r => r.json()).then(files => {
    buildRows(files);
  });
}

// ── Upload ──
function uploadFile(key, accept) {
  document.getElementById('input-' + key).click();
}

async function doUpload(key, input) {
  const file = input.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append('file', file);
  fd.append('key', key);

  const btn = document.getElementById('ub-' + key);
  btn.textContent = '上传中...';
  btn.disabled = true;

  try {
    const resp = await fetch('/api/upload', { method: 'POST', body: fd });
    const data = await resp.json();
    if (data.ok) {
      fileStates[key] = { name: data.name, exists: true };
      document.getElementById('fn-' + key).textContent = data.name;
      document.getElementById('fn-' + key).classList.add('exists');
      document.getElementById('fs-' + key).textContent = '✓';
      document.getElementById('fs-' + key).className = 'file-status ok';
      btn.textContent = '更换';
      btn.classList.remove('primary');
      checkAllReady();
    }
  } catch(e) {
    alert('上传失败: ' + e);
  }
  btn.disabled = false;
  input.value = '';
}

// ── Generate ──
async function startGenerate() {
  document.getElementById('btn-generate').disabled = true;
  document.getElementById('log-area').innerHTML = '';
  document.getElementById('progress-fill').style.width = '0%';
  document.getElementById('progress-text').textContent = '正在启动...';
  document.getElementById('status-text').textContent = '';
  document.getElementById('download-wrap').style.display = 'none';

  try {
    const resp = await fetch('/api/generate', { method: 'POST' });
    const data = await resp.json();
    if (!data.ok) {
      alert(data.error || '启动失败');
      document.getElementById('btn-generate').disabled = false;
      checkAllReady();
      return;
    }
    pollTimer = setInterval(pollProgress, 400);
  } catch(e) {
    document.getElementById('status-text').textContent = '启动失败: ' + e;
    document.getElementById('btn-generate').disabled = false;
    checkAllReady();
  }
}

async function pollProgress() {
  try {
    const resp = await fetch('/api/progress');
    const data = await resp.json();

    if (data.total > 0) {
      const pct = Math.round((data.current / data.total) * 100);
      document.getElementById('progress-fill').style.width = pct + '%';
      document.getElementById('progress-text').textContent = data.current + '/' + data.total + ' 页';
    }

    document.getElementById('status-text').textContent = data.status || '';

    const logEl = document.getElementById('log-area');
    if (data.messages && data.messages.length > 0) {
      data.messages.forEach(m => {
        const div = document.createElement('div');
        div.textContent = m;
        logEl.appendChild(div);
      });
      logEl.scrollTop = logEl.scrollHeight;
    }

    if (data.done) {
      clearInterval(pollTimer);
      pollTimer = null;
      document.getElementById('progress-fill').style.width = '100%';
      if (data.error) {
        document.getElementById('status-text').textContent = '失败: ' + data.error;
      } else {
        document.getElementById('progress-text').textContent = '完成';
        document.getElementById('status-text').textContent = '生成完成!';
        document.getElementById('download-wrap').style.display = 'block';
      }
      document.getElementById('btn-generate').disabled = false;
    }
  } catch(e) {}
}

init();
</script>
</body>
</html>"""

# ── Request Handler ──────────────────────────────────────
class Handler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        pass  # suppress stderr log

    def _send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)

    def _send_html(self, html, status=200):
        body = html.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)

    def _parse_multipart(self):
        content_type = self.headers.get("Content-Type", "")
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length)
        msg = email.parser.BytesParser(policy=email.policy.default).parsebytes(
            b"Content-Type: " + content_type.encode() + b"\r\n\r\n" + body
        )
        fields = {}
        for part in msg.iter_parts():
            name = part.get_param("name", header="Content-Disposition")
            if part.get_filename():
                fields[name] = {
                    "filename": part.get_filename(),
                    "data": part.get_payload(decode=True),
                }
            else:
                fields[name] = part.get_payload(decode=True).decode("utf-8")
        return fields

    def do_GET(self):
        path = urlparse(self.path).path

        if path == "/" or path == "/index.html":
            self._send_html(HTML)

        elif path == "/api/files":
            self._send_json({
                "tpl": {"name": "template.xlsx", "exists": TPL_EXISTS},
                "zd":  {"name": uploaded_files.get("zd") or "总单.xlsx",
                        "exists": uploaded_files["zd"] is not None},
                "sp":  {"name": uploaded_files.get("sp") or "附件备件清单.xls",
                        "exists": uploaded_files["sp"] is not None},
            })

        elif path == "/api/progress":
            msgs = []
            while not progress_queue.empty():
                try:
                    msgs.append(progress_queue.get_nowait())
                except queue.Empty:
                    break
            progress_log.extend(msgs)
            resp = {
                "done": generation_done.is_set(),
                "current": progress_state["current"],
                "total": progress_state["total"],
                "status": progress_log[-1] if progress_log else "",
                "messages": msgs,
            }
            if generation_error_msg:
                resp["error"] = generation_error_msg
            self._send_json(resp)

        elif path == "/api/download":
            out = os.path.join(WORK_DIR, "分箱单_生成结果.xlsx")
            if not os.path.isfile(out):
                self._send_json({"error": "文件不存在"}, 404)
                return
            with open(out, "rb") as f:
                data = f.read()
            self.send_response(200)
            filename = "分箱单_生成结果.xlsx"
            encoded = quote(filename, safe="")
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition",
                             f"attachment; filename*=UTF-8''{encoded}")
            self.send_header("Content-Length", len(data))
            self.end_headers()
            self.wfile.write(data)

        else:
            self._send_json({"error": "not found"}, 404)

    def do_POST(self):
        path = urlparse(self.path).path

        if path == "/api/upload":
            try:
                fields = self._parse_multipart()
                key = fields.get("key", "")
                file_info = fields.get("file")
                if not file_info:
                    self._send_json({"ok": False, "error": "未收到文件"}, 400)
                    return
                filename = os.path.basename(file_info["filename"])
                dest = os.path.join(WORK_DIR, filename)
                with open(dest, "wb") as f:
                    f.write(file_info["data"])
                if key in uploaded_files:
                    uploaded_files[key] = filename
                self._send_json({"ok": True, "name": filename, "key": key})
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)}, 500)

        elif path == "/api/generate":
            # Validate: template must exist, zd and sp must be uploaded
            if not TPL_EXISTS:
                self._send_json({"ok": False, "error": "模板文件(template.xlsx)不存在，请上传"}, 400)
                return
            zd_name = uploaded_files.get("zd")
            sp_name = uploaded_files.get("sp")
            if not zd_name:
                self._send_json({"ok": False, "error": "请先上传总单文件"}, 400)
                return
            if not sp_name:
                self._send_json({"ok": False, "error": "请先上传备件清单"}, 400)
                return

            zd_path = os.path.join(WORK_DIR, zd_name)
            sp_path = os.path.join(WORK_DIR, sp_name)
            if not os.path.isfile(zd_path):
                self._send_json({"ok": False, "error": f"总单文件不存在: {zd_name}"}, 400)
                return
            if not os.path.isfile(sp_path):
                self._send_json({"ok": False, "error": f"备件清单不存在: {sp_name}"}, 400)
                return

            # Reset state
            global progress_log, generation_error_msg
            progress_log = []
            while not progress_queue.empty():
                try: progress_queue.get_nowait()
                except queue.Empty: break
            generation_done.clear()
            generation_error_msg = None
            progress_state["current"] = 0
            progress_state["total"] = 0

            self._send_json({"ok": True})

            def run():
                global generation_error_msg
                def on_progress(msg, cur, tot):
                    progress_queue.put(msg)
                    if tot > 0:
                        progress_state["current"] = cur
                        progress_state["total"] = tot

                try:
                    gpl.generate(
                        template_path=TPL_PATH,
                        zongdan_path=zd_path,
                        spare_path=sp_path,
                        output_path=os.path.join(WORK_DIR, "分箱单_生成结果.xlsx"),
                        progress_callback=on_progress,
                    )
                except Exception as e:
                    generation_error_msg = str(e)
                finally:
                    generation_done.set()

            threading.Thread(target=run, daemon=True).start()

        elif path == "/api/shutdown":
            self._send_json({"ok": True})
            threading.Thread(target=self.server.shutdown, daemon=True).start()

        else:
            self._send_json({"error": "not found"}, 404)


if __name__ == "__main__":
    server = HTTPServer(("127.0.0.1", PORT), Handler)
    server.allow_reuse_address = True
    url = f"http://localhost:{PORT}"
    print(f"分箱单生成工具已启动: {url}")
    print("按 Ctrl+C 停止服务器")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
        server.server_close()
