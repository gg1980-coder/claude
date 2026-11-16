"""Weekly US soybean inspections workbook: China vs RoW, regions, history and pace.
Usage: python3 -I build.py   (inputs/outputs in config.py, overridable via SOY_* env vars)"""
import os
HERE=os.path.dirname(os.path.abspath(__file__)); NS={}
for f in ['config.py','s1_base.py','s2_country.py','s3_summary.py','s4_weekly.py','s5_hist.py','s6_progress.py','s7_bench.py']:
    exec(compile(open(os.path.join(HERE,f)).read(),f,'exec'),NS)
NS['wb'].save(NS['OUT']); print(NS['OUT'])
