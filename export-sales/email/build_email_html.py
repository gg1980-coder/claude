"""Render email.md to a self-contained HTML email (charts embedded) for preview / pasting into Outlook."""
import base64, re, markdown
md = open('email.md', encoding='utf-8').read()
charts = {'Chart 1': '1_bridge.png', 'Chart 2': '2_milestones.png', 'Chart 4': '4_mental_map.png', 'Chart 6': '6_who_ships_when.png'}
def img(m):
    key = m.group(1).split(':')[0].strip()
    b64 = base64.b64encode(open(charts[key], 'rb').read()).decode()
    return f'<img src="data:image/png;base64,{b64}" alt="{m.group(1)}" style="width:100%;max-width:900px;border:1px solid #e4e3df;margin:6px 0">'
md = re.sub(r'^\[(Chart \d[^\]]*)\]$', lambda m: f'@@IMG@@{m.group(1)}@@', md, flags=re.M)
html = markdown.markdown(md, extensions=['tables'])
html = re.sub(r'<p>@@IMG@@(.*?)@@</p>', img, html)
css = """body{font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#0b0b0b;max-width:920px;margin:16px auto;padding:0 16px;line-height:1.45}
table{border-collapse:collapse;margin:8px 0 14px;font-size:13px}th{background:#1F4E78;color:#fff;padding:6px 8px;text-align:left}
td{border-bottom:1px solid #e4e3df;padding:5px 8px;vertical-align:top}hr{border:0;border-top:1px solid #e4e3df;margin:18px 0}
strong{color:#0b0b0b}th strong{color:#fff}h1,h2,h3{color:#1F4E78}"""
open('email.html', 'w', encoding='utf-8').write(f'<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>US soybean exports vs USDA</title><style>{css}</style></head><body>{html}</body></html>')
print('ok', len(html))
