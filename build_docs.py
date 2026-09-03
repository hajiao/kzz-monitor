from pathlib import Path

import markdown

source = Path("KzzMonitor详细操作手册.md").read_text(encoding="utf-8")
body = markdown.markdown(source, extensions=["tables", "fenced_code", "toc", "sane_lists"])
css = r"""
:root{color-scheme:light}*{box-sizing:border-box}
body{margin:0;background:#f3f6fa;color:#17202a;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei","PingFang SC",sans-serif;line-height:1.72}
.page{max-width:1050px;margin:28px auto;background:#fff;padding:56px 72px;box-shadow:0 8px 35px #18324a22;border-radius:12px}
h1{color:#174f84;border-bottom:3px solid #2f80c9;padding-bottom:16px}h2{color:#1e659f;border-bottom:1px solid #d8e4ef;padding-bottom:8px;margin-top:38px}h3{color:#2974ab;margin-top:27px}
table{border-collapse:collapse;width:100%;margin:16px 0;font-size:.94rem}th,td{border:1px solid #cbd9e6;padding:9px 12px;vertical-align:top}th{background:#eaf3fb;color:#174f84;text-align:left}tr:nth-child(even) td{background:#f8fbfe}
code{font-family:Consolas,Menlo,monospace;background:#eef3f7;padding:2px 5px;border-radius:4px}pre{background:#152230;color:#e8f1f8;padding:16px;border-radius:8px;overflow:auto}pre code{background:transparent;padding:0}
blockquote{margin:18px 0;padding:12px 18px;border-left:5px solid #e49b26;background:#fff7e7}a{color:#1769aa}
@media print{body{background:#fff}.page{max-width:none;margin:0;padding:12mm 13mm;box-shadow:none;border-radius:0}table,pre,blockquote{break-inside:avoid}@page{size:A4;margin:10mm}}
@media(max-width:700px){.page{margin:0;padding:24px 18px;border-radius:0}table{display:block;overflow-x:auto}}
"""
html = f'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>KzzMonitor 详细操作手册</title><style>{css}</style></head><body><main class="page">{body}</main></body></html>'''
Path("KzzMonitor详细操作手册.html").write_text(html, encoding="utf-8")
print("Generated KzzMonitor详细操作手册.html")
