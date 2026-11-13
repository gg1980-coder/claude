"""Render email.md to a self-contained HTML email (charts embedded) for preview / pasting into Outlook."""
import base64, re, markdown
md = open('email.md', encoding='utf-8').read()
charts = {'Chart': 'chart_1700_vs_history.png'}
def img(m):
    key = 'Chart'
    b64 = base64.b64encode(open(charts[key], 'rb').read()).decode()
    return f'<img src="data:image/png;base64,{b64}" alt="{m.group(1)}" style="width:100%;max-width:900px;border:1px solid #e4e3df;margin:6px 0">'
md = re.sub(r'^\[(Chart \d[^\]]*)\]$', lambda m: f'@@IMG@@{m.group(1)}@@', md, flags=re.M)
html = markdown.markdown(md, extensions=['tables'])
html = re.sub(r'<p>@@IMG@@(.*?)@@</p>', img, html)
css = """body{font-family:Calibri,Carlito,Arial,sans-serif;font-size:11pt;color:#000;margin:12px;line-height:1.3;background:#fff}
p{margin:0 0 11pt}ul,ol{margin:0 0 11pt;padding-left:28px}li{margin:0}
table{border-collapse:collapse;margin:0 0 11pt;font-size:11pt}th,td{border:1px solid #000;padding:2px 6px;text-align:left;vertical-align:top}
th{font-weight:bold;background:#fff;color:#000}"""
open('email.html', 'w', encoding='utf-8').write(f'<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>US soybean exports email</title><style>{css}</style></head><body>{html}</body></html>')
print('ok', len(html))
