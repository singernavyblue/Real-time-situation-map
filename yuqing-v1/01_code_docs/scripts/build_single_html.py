# -*- coding: utf-8 -*-
"""把大屏打包成自包含的单一 HTML 文件（内联 echarts/china/data）"""
import os, re, sys
sys.stdout.reconfigure(encoding="utf-8")

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
def _resolve_base():
    env = os.environ.get("DASH_BASE")
    if env:
        return env
    pkg = os.path.abspath(os.path.join(_SCRIPT_DIR, "..", "dashboard"))
    if os.path.isdir(pkg):
        return pkg
    return os.path.abspath(os.path.join(_SCRIPT_DIR, "..", "民族相关网络舆情态势"))

BASE = _resolve_base()
SRC = os.path.join(BASE, "index.html")
OUT = os.path.join(BASE, "民族舆情态势感知大屏_单文件版.html")
print("大屏目录:", BASE)

with open(SRC, encoding="utf-8") as f:
    html = f.read()

def inline(asset_path, placeholder):
    global html
    with open(os.path.join(BASE, asset_path), encoding="utf-8") as f:
        content = f.read()
    # 防止内联内容提前闭合 <script> 标签
    content = content.replace("</script", "<\\/script").replace("<!--", "<\\!--")
    if placeholder not in html:
        print(f"[!] 占位符未找到：{placeholder}，跳过")
        return
    html = html.replace(placeholder, "<script>\n" + content + "\n</script>")
    print(f"[+] 已内联 {asset_path}（{len(content)//1024} KB）")

inline("assets/echarts.min.js", '<script src="assets/echarts.min.js"></script>')
inline("assets/china.js", '<script src="assets/china.js"></script>')
inline("assets/data.js", '<script src="assets/data.js"></script>')

with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)

size_kb = os.path.getsize(OUT) // 1024
print(f"[+] 已生成：{OUT}（{size_kb} KB）")
print("[+] 残留外部脚本引用：", len(re.findall(r'<script src=', html)))
