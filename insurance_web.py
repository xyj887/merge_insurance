# -*- coding: utf-8 -*-
"""
投保数据汇总工具 - 桌面版（本地 Web 服务 + 原生窗口）

启动方式：
  python3 insurance_web.py

技术模型：
  1. 启动 Flask 服务，只监听 127.0.0.1（外部机器访问不到，数据不出本机）
  2. 用 pywebview 开一个原生窗口加载该地址（无浏览器地址栏/标签页）
  3. 页面上传 Excel → 后端汇总 → 浏览器下载结果

后续若要做成内网服务，去掉 webview 部分，把 Flask 应用部署到内网服务器即可，
前端页面无需改动。
"""

import os
import socket
import sys
import tempfile
import threading
import traceback
import uuid
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from merge_insurance import build_result_df, write_excel

APP_TITLE = "投保数据汇总工具"
MAX_UPLOAD_MB = 200


def resource_path(relative):
    """获取资源绝对路径。PyInstaller 打包后资源在 sys._MEIPASS 临时目录中。"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


def find_free_port():
    """找一个空闲端口，避免端口冲突导致启动失败。"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


app = Flask(
    __name__,
    template_folder=resource_path("templates"),
    static_folder=resource_path("static"),
)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.json.ensure_ascii = False  # 日志中文原样返回，不被转义成 \uXXXX

# 上传/输出文件放在临时目录，关闭程序后可被系统回收
WORK_DIR = Path(tempfile.gettempdir()) / "insurance_tool"
WORK_DIR.mkdir(parents=True, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html", title=APP_TITLE, max_mb=MAX_UPLOAD_MB)


@app.route("/api/process", methods=["POST"])
def api_process():
    """接收上传的 Excel，返回汇总结果与下载地址"""
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "没有收到文件"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"ok": False, "error": "文件名为空"}), 400

    suffix = Path(file.filename).suffix.lower()
    if suffix not in (".xlsx", ".xls"):
        return jsonify({"ok": False, "error": f"不支持的文件类型：{suffix}，请上传 xlsx 或 xls"}), 400

    # 用 uuid 前缀避免多用户/多次上传重名冲突
    token = uuid.uuid4().hex[:8]
    safe_name = secure_filename(file.filename) or f"input{suffix}"
    src_path = WORK_DIR / f"{token}_{safe_name}"
    file.save(str(src_path))

    try:
        logs = []
        result_df = build_result_df(str(src_path), log=logs.append)

        if len(result_df) == 0:
            return jsonify({"ok": False, "error": "汇总结果为空，请检查文件内容是否正确"}), 400

        out_name = f"{Path(file.filename).stem}_汇总.xlsx"
        out_path = WORK_DIR / f"{token}_{out_name}"
        write_excel(result_df, str(out_path))

        return jsonify({
            "ok": True,
            "rows": len(result_df),
            "days": int(result_df["保障天数"].sum()),
            "logs": logs,
            # 只传 token，文件名由服务端读取，避免中文文件名走 URL 传参出错
            "download_url": f"/download?token={token}",
            "filename": out_name,
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e).replace("\n", " ")}), 500


@app.route("/download")
def download():
    """下载处理结果。只用 token 定位，文件名从磁盘读取，规避中文 URL 传参问题。"""
    token = request.args.get("token", "").strip()
    # token 只允许十六进制字符，防止路径穿越
    if not token or not all(c in "0123456789abcdef" for c in token):
        return "无效的请求", 400

    matches = [p for p in WORK_DIR.glob(f"{token}_*.xlsx")]
    if not matches:
        return "文件不存在或已过期，请重新处理", 404

    path = matches[0]
    original_name = path.name[len(token) + 1:]  # 去掉 "{token}_" 前缀还原文件名
    return send_file(str(path), as_attachment=True, download_name=original_name)


@app.route("/api/quit", methods=["POST"])
def api_quit():
    """页面请求退出程序"""
    threading.Timer(0.6, _force_exit).start()
    return jsonify({"ok": True})


def _force_exit():
    os._exit(0)


def start_flask(port):
    """Flask 服务跑在后台线程，主线程留给窗口。关闭 werkzeug 的请求日志避免干扰。"""
    import logging
    logging.getLogger("werkzeug").setLevel(logging.ERROR)
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True, use_reloader=False)


def main():
    port = find_free_port()
    threading.Thread(target=start_flask, args=(port,), daemon=True).start()

    # 等待 Flask 起来
    import time
    for _ in range(50):
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.2):
                break
        except OSError:
            time.sleep(0.1)

    url = f"http://127.0.0.1:{port}"

    try:
        import webview
        window = webview.create_window(
            APP_TITLE, url=url, width=1000, height=720,
            resizable=True, min_size=(800, 600),
        )
        webview.start()
    except ImportError:
        # 未装 pywebview 时退化为打开系统默认浏览器
        print(f"未安装 pywebview，已改用系统默认浏览器打开：{url}")
        import webbrowser
        webbrowser.open(url)
        print("按 Ctrl+C 退出")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass


if __name__ == "__main__":
    main()
