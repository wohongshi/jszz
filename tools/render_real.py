#!/usr/bin/env python3
"""把 U 盘真实课件转 PNG：每份取 封面/前部/中部/后部 各 1 页，留证分析。"""
import json, subprocess
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / 'analysis/realppt/png'
OUT.mkdir(parents=True, exist_ok=True)
jobs = json.load(open(OUT.parent / 'jobs.json'))
tmp = Path('/tmp/realppt'); tmp.mkdir(exist_ok=True)

def soffice_pdf(src, tag):
    od = tmp / tag; od.mkdir(exist_ok=True)
    for attempt in range(2):
        r = subprocess.run(['timeout', '280', 'soffice', '--headless',
            f'-env:UserInstallation=file:///tmp/lo_{tag}_{attempt}',
            '--convert-to', 'pdf', '--outdir', str(od), src],
            capture_output=True, text=True)
        pdfs = list(od.glob('*.pdf'))
        if pdfs:
            return pdfs[0]
        print('RETRY', tag, r.stdout[-160:], r.stderr[-160:], flush=True)
    print('FAIL', src, flush=True)
    return None

for subj, files in jobs.items():
    for i, src in enumerate(files):
        tag = f"{subj}{i}"
        if list(OUT.glob(f'{tag}-*.png')):
            print('skip', tag, flush=True)
            continue
        pdf = soffice_pdf(src, tag)
        if not pdf:
            continue
        n = int(subprocess.run(['pdfinfo', str(pdf)], capture_output=True, text=True)
                .stdout.split('Pages:')[1].split()[0])
        pages = sorted(set([1, 2, max(3, n // 2), max(4, n - 1), n]))
        for pg in pages:
            subprocess.run(['pdftoppm', '-r', '70', '-png', '-f', str(pg), '-l', str(pg),
                            str(pdf), str(OUT / f'{tag}')], capture_output=True)
        print(f'{tag} {Path(src).name[:30]} {n}页 -> {pages}', flush=True)
print('DONE')
