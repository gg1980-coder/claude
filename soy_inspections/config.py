# Inputs for the weekly build. Update CY2026 (and CY2025 if it changes) each week.
import os
SCR=os.environ.get('SOY_DATA','/tmp/claude-0/-home-user-claude/49d7a1ba-dec0-5229-9dfb-47d4b528b708/scratchpad/hist/')
CY2025=os.environ.get('SOY_CY2025',SCR+'CY2025.csv')
CY2026=os.environ.get('SOY_CY2026','/root/.claude/uploads/49d7a1ba-dec0-5229-9dfb-47d4b528b708/02cca70a-CY2026_7.csv')
HIST={y:SCR+f'CY{y}.csv' for y in range(2015,2025)}
PREV=os.environ.get('SOY_PREV','/home/user/claude/Soybean_Inspections_Sep_Nov_2025_vs_2026_China_vs_ROW_upd_W2Nov.xlsx')
OUT=os.environ.get('SOY_OUT','/home/user/claude/Soybean_Inspections_Sep_Nov_2025_vs_2026_China_vs_ROW_upd_W3Nov.xlsx')
DATA_NOTE=('CY2015–CY2025 from fgisonline.ams.usda.gov/ExportGrainReport (CY2025 identical to the user-supplied file); CY2026 user-supplied.')
# Periods: (tab name, short label, 2025 Thursdays, 2026 Thursdays, description)
PERIODS=[('September','Sep',['20250904','20250911','20250918','20250925'],['20260903','20260910','20260917','20260924'],
          'weeks ending Thu 4/11/18/25 Sep 2025 vs Thu 3/10/17/24 Sep 2026'),
         ('Week 1 October','W1 Oct',['20251002'],['20261001'],'week ending Thu 2 Oct 2025 (26 Sep–2 Oct) vs week ending Thu 1 Oct 2026 (25 Sep–1 Oct)'),
         ('Week 2 October','W2 Oct',['20251009'],['20261008'],'week ending Thu 9 Oct 2025 (3–9 Oct) vs week ending Thu 8 Oct 2026 (2–8 Oct)'),
         ('Week 3 October','W3 Oct',['20251016'],['20261015'],'week ending Thu 16 Oct 2025 (10–16 Oct) vs week ending Thu 15 Oct 2026 (9–15 Oct)'),
         ('Week 4 October','W4 Oct',['20251023'],['20261022'],'week ending Thu 23 Oct 2025 (17–23 Oct) vs week ending Thu 22 Oct 2026 (16–22 Oct)'),
         ('Week 5 October','W5 Oct',['20251030'],['20261029'],'week ending Thu 30 Oct 2025 (24–30 Oct) vs week ending Thu 29 Oct 2026 (23–29 Oct)'),
         ('Week 1 November','W1 Nov',['20251106'],['20261105'],'week ending Thu 6 Nov 2025 (31 Oct–6 Nov) vs week ending Thu 5 Nov 2026 (30 Oct–5 Nov)'),
         ('Week 2 November','W2 Nov',['20251113'],['20261112'],'week ending Thu 13 Nov 2025 (7–13 Nov) vs week ending Thu 12 Nov 2026 (6–12 Nov)'),
         ('Week 3 November','W3 Nov',['20251120'],['20261119'],'week ending Thu 20 Nov 2025 (14–20 Nov) vs week ending Thu 19 Nov 2026 (13–19 Nov)')]
SEASON_LABEL='Season to date (Sep–Nov)'
COUNTRY_TITLE='season to date (Sep–Oct + Nov wk 1–3)'
EARLIER_REVISIONS=('Earlier revisions: 16 Nov update Wk5 Oct +490 / Wk1 Nov +17,994 MT (RoW); 9 Nov update Wk4 Oct +196 / Wk5 Oct +67,145 MT (RoW); 2 Nov Wk1–Wk4 Oct +31,295 MT (RoW); 26 Oct Wk2 +3,209 / Wk3 +66,336 (China +52,036); '
                   '19 Oct Wk2 +12,830; 13 Oct Sep +759, Wk1 +61,014 MT.')
