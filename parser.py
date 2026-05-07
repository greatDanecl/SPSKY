"""
FDS Parser — lee todos los xlsx en /data y genera index.html con datos embebidos.
Uso: python parser.py
"""
import pandas as pd, re, json, numpy as np, os, glob, sys
from datetime import timedelta
from collections import defaultdict
from pathlib import Path

# ── CONFIGURACIÓN ────────────────────────────────────────────────────────────
DATA_DIR    = Path(__file__).parent / "data"
OUTPUT_HTML = Path(__file__).parent / "index.html"

MONTH_MAP = {
    'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06',
    'Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12',
    'Ene':'01','Feb':'02','Mar':'03','Abr':'04','May':'05','Jun':'06',
    'Jul':'07','Ago':'08','Sep':'09','Oct':'10','Nov':'11','Dic':'12',
}
PERIOD_LABELS = {
    '2025-10':'Oct 2025','2026-01':'Ene 2026','2026-02':'Feb 2026',
    '2026-03':'Mar 2026','2026-04':'Abr 2026','2026-05':'May 2026',
    '2026-06':'Jun 2026','2026-07':'Jul 2026','2026-08':'Ago 2026',
    '2026-09':'Sep 2026','2026-10':'Oct 2026','2026-11':'Nov 2026',
    '2026-12':'Dic 2026',
}

# ── HELPERS ──────────────────────────────────────────────────────────────────
def parse_td(val):
    if val is None or (isinstance(val, float) and np.isnan(val)): return 0.0
    if isinstance(val, timedelta): return round(val.total_seconds()/3600, 2)
    s = str(val)
    m = re.match(r'(\d+) days? (\d+):(\d+)', s)
    if m: return round(int(m.group(1))*24 + int(m.group(2)) + int(m.group(3))/60, 2)
    m2 = re.match(r'(\d+):(\d+)', s)
    if m2: return round(int(m2.group(1)) + int(m2.group(2))/60, 2)
    return 0.0

def detect_period(df):
    """Extrae el período del encabezado del archivo (fila 1 o 8)."""
    # Try row 1, col 2
    try:
        cell = str(df.iloc[1, 2])
        m = re.search(r'(\d{2})(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|'
                      r'Ene|Abr|Ago|Dic)(\d{2})', cell, re.I)
        if m:
            mon = m.group(2).capitalize()
            yr  = '20' + m.group(3)
            return f"{yr}-{MONTH_MAP.get(mon,'00')}"
    except: pass
    # Try row 8 first date cell
    try:
        for col in range(1, df.shape[1]):
            cell = str(df.iloc[8, col])
            m = re.match(r'(\d{2})(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|'
                         r'Ene|Abr|Ago|Dic)', cell, re.I)
            if m:
                mon = m.group(2).capitalize()
                yr_cell = str(df.iloc[1, col]) if col < df.shape[1] else ''
                yr_m = re.search(r'20(\d{2})', yr_cell)
                yr = '20' + yr_m.group(1) if yr_m else '2026'
                return f"{yr}-{MONTH_MAP.get(mon,'00')}"
    except: pass
    return None

def classify_role_from_filename(fname):
    """Intenta detectar si es programado o efectuado desde el nombre del archivo."""
    fl = fname.lower()
    if any(w in fl for w in ['prog','plan','mando','master','scheduled','sched']): 
        return 'programmed'
    if any(w in fl for w in ['actual','efect','horas','flown','real']):            
        return 'actual'
    return None  # indeterminado — se detecta por contenido

def detect_role_from_sheet(sheet_name):
    sl = sheet_name.lower()
    if 'hora' in sl or 'actual' in sl or 'efect' in sl: return 'actual'
    if 'pdc' in sl or 'plan' in sl or 'prog' in sl:     return 'programmed'
    return None

def parse_sheet(df, period, role):
    pilots = []
    if len(df) < 10 or df.shape[1] < 2: return pilots
    r9 = str(df.iloc[9, 0]).strip()
    abcd = bool(re.match(r'^[A-H]$', r9))
    i = 9
    while i < len(df):
        c0 = str(df.iloc[i, 0]).strip()
        if abcd:
            if c0 != 'A': i += 1; continue
            code    = str(df.iloc[i,   1]).strip()
            fname   = str(df.iloc[i+1, 1]).strip() if i+1 < len(df) else ''
            lname   = str(df.iloc[i+2, 1]).strip() if i+2 < len(df) else ''
            rut_pos = str(df.iloc[i+3, 1]).strip() if i+3 < len(df) else ''
            base    = str(df.iloc[i+4, 1]).strip() if i+4 < len(df) else ''
            cred_h  = parse_td(df.iloc[i+5, 1] if i+5 < len(df) else None)
            blk_h   = parse_td(df.iloc[i+6, 1] if i+6 < len(df) else None)
            duty_h  = parse_td(df.iloc[i+7, 1] if i+7 < len(df) else None)
            sched   = [str(v).strip() for v in df.iloc[i, 2:].tolist()]
            i += 8
        else:
            if not re.match(r'^[A-Z]{4,5}$', c0): i += 1; continue
            code    = c0
            fname   = str(df.iloc[i+1, 0]).strip() if i+1 < len(df) else ''
            lname   = str(df.iloc[i+2, 0]).strip() if i+2 < len(df) else ''
            rut_pos = str(df.iloc[i+3, 0]).strip() if i+3 < len(df) else ''
            base    = str(df.iloc[i+4, 0]).strip() if i+4 < len(df) else ''
            cred_h  = parse_td(df.iloc[i+6, 1] if i+6 < len(df) else None)
            blk_h   = parse_td(df.iloc[i+7, 1] if i+7 < len(df) else None)
            duty_h  = parse_td(df.iloc[i+8, 1] if i+8 < len(df) else None)
            sched   = [str(v).strip() for v in df.iloc[i, 2:].tolist()]
            i += 9

        if not re.match(r'^[A-Z]{4,5}$', code): continue
        pos = rut_pos.split(' - ')[-1].strip() if ' - ' in rut_pos else ''
        if not pos or pos in ['nan','NaT','']: continue
        name = f'{fname} {lname}'.strip()
        if not name or any(t in name.upper() for t in ['TEST','PRUEBA','NAN']): continue

        pg = 'Otro'
        if pos in ['CP','CPN']:                    pg = 'Capitán'
        elif pos in ['FO','FON']:                  pg = 'Primer Oficial'
        elif pos in ['INS','INST','IOA','C15M']:   pg = 'Instructor'

        vac  = sum(1 for s in sched if any(w in s.upper() for w in ['VACAC','VACAO','VACAP','VACAS','VACAI','VACAOP']))
        med  = sum(1 for s in sched if any(w in s.upper() for w in ['LM','LICM','LICMED']))
        lib  = sum(1 for s in sched if s in ['LIBRE','FINDE'])
        sim  = sum(1 for s in sched if 'SIM' in s.upper())
        total_days = len([s for s in sched if s not in ['nan','NaT','','None']])
        excl = ((vac + med) / max(total_days, 1)) > 0.35 or blk_h < 5

        pilots.append({
            'period':period, 'role':role, 'code':code, 'name':name,
            'pos':pos, 'pos_group':pg, 'base':base,
            'block_h':blk_h, 'duty_h':duty_h, 'credits_h':cred_h,
            'libre_days':lib, 'vac_days':vac, 'med_days':med, 'sim_days':sim,
            'exclude_from_avg':excl,
        })
    return pilots

# ── PROCESO PRINCIPAL ────────────────────────────────────────────────────────
def build_dataset():
    xlsx_files = sorted(glob.glob(str(DATA_DIR / "*.xlsx")))
    if not xlsx_files:
        print(f"ERROR: No se encontraron archivos .xlsx en {DATA_DIR}")
        sys.exit(1)

    print(f"Procesando {len(xlsx_files)} archivos en {DATA_DIR}...\n")

    all_records = []
    for fpath in xlsx_files:
        fname = os.path.basename(fpath)
        role_from_name = classify_role_from_filename(fname)
        try:
            xl = pd.ExcelFile(fpath)
        except Exception as e:
            print(f"  ✗ {fname}: error al abrir — {e}")
            continue

        for sheet_name in xl.sheet_names:
            try:
                df = pd.read_excel(fpath, sheet_name=sheet_name, header=None)
                period = detect_period(df)
                if not period:
                    print(f"  ⚠ {fname}/{sheet_name}: no se detectó período, omitiendo")
                    continue
                role = (role_from_name
                        or detect_role_from_sheet(sheet_name)
                        or ('actual' if sheet_name.lower() != 'pdc report' else 'programmed'))
                recs = parse_sheet(df, period, role)
                all_records.extend(recs)
                lbl = PERIOD_LABELS.get(period, period)
                print(f"  ✓ {fname}/{sheet_name}: {lbl} · {role:12s} · {len(recs)} pilotos")
            except Exception as e:
                print(f"  ✗ {fname}/{sheet_name}: {e}")

    # Merge programmed + actual per pilot+period
    summary = {}
    for r in all_records:
        key = (r['period'], r['code'])
        if key not in summary:
            summary[key] = {
                'period':r['period'], 'code':r['code'], 'name':r['name'],
                'pos':r['pos'], 'pos_group':r['pos_group'], 'base':r['base'],
                'libre_days':r['libre_days'], 'vac_days':r['vac_days'],
                'med_days':r['med_days'], 'sim_days':r['sim_days'],
                'exclude_from_avg':r['exclude_from_avg'],
                'block_h_programmed':None, 'duty_h_programmed':None, 'credits_h_programmed':None,
                'block_h_actual':None,      'duty_h_actual':None,     'credits_h_actual':None,
            }
        rk = r['role']
        summary[key][f'block_h_{rk}']   = r['block_h']
        summary[key][f'duty_h_{rk}']    = r['duty_h']
        summary[key][f'credits_h_{rk}'] = r['credits_h']
        if r['exclude_from_avg']:
            summary[key]['exclude_from_avg'] = True

    records = list(summary.values())
    periods = sorted(set(r['period'] for r in records))

    # Summary
    from collections import Counter
    groups = Counter(r['pos_group'] for r in records)
    print(f"\n{'─'*50}")
    print(f"Total registros: {len(records)} · Períodos: {[PERIOD_LABELS.get(p,p) for p in periods]}")
    for g, n in sorted(groups.items()):
        names = len(set(r['name'] for r in records if r['pos_group'] == g))
        print(f"  {g}: {names} pilotos únicos")

    return records, periods

# ── GENERAR HTML ─────────────────────────────────────────────────────────────
def generate_html(records, periods):
    period_labels_js = {p: PERIOD_LABELS.get(p, p) for p in periods}
    data_json    = json.dumps(records,        ensure_ascii=False, default=str)
    periods_json = json.dumps(periods,        ensure_ascii=False)
    labels_json  = json.dumps(period_labels_js, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>FDS · Productividad de Tripulación</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600&family=DM+Sans:wght@300;400;500&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
:root{{
  --sand-50:#FAF7F2;--sand-100:#F3EDE2;--sand-200:#E8DCC8;--sand-300:#D4C4A8;
  --sand-400:#BFA882;--sand-500:#A68B5B;--sand-600:#8A7048;--sand-700:#6B5535;
  --earth-800:#4A3B25;--earth-900:#2E2416;
  --clay:#C4856A;--clay-dim:rgba(196,133,106,0.12);
  --rust:#B5603A;--sage:#7A9E7E;--sage-dim:rgba(122,158,126,0.12);
  --dusk:#8B7BA8;--dusk-dim:rgba(139,123,168,0.12);
  --warm-red:#C4534A;--warm-red-dim:rgba(196,83,74,0.10);
  --bg:#F7F3ED;--surface:#FDFAF6;--surface2:#F3EDE2;
  --border:#E2D8C8;--border2:#D4C4A8;
  --text:#2E2416;--text2:#6B5535;--muted:#A68B5B;--dim:#BFA882;
  --r:10px;--r2:16px;
  --shadow:0 1px 3px rgba(46,36,22,.06),0 4px 16px rgba(46,36,22,.04);
  --shadow2:0 2px 8px rgba(46,36,22,.08),0 8px 32px rgba(46,36,22,.06);
  --font:'DM Sans',sans-serif;--display:'Playfair Display',serif;--mono:'DM Mono',monospace;
}}
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
html{{font-size:14px}}
body{{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;line-height:1.5}}
.shell{{display:grid;grid-template-columns:230px 1fr;min-height:100vh}}
.sidebar{{background:var(--earth-800);display:flex;flex-direction:column;position:sticky;top:0;height:100vh;overflow:hidden}}
.sidebar-top{{padding:26px 22px 18px;border-bottom:1px solid rgba(255,255,255,.07)}}
.brand{{display:flex;align-items:center;gap:10px;margin-bottom:3px}}
.brand-icon{{width:30px;height:30px;background:var(--clay);border-radius:7px;display:flex;align-items:center;justify-content:center}}
.brand-icon svg{{width:15px;height:15px;stroke:white}}
.brand-name{{font-family:var(--display);font-size:15px;color:var(--sand-100);letter-spacing:.02em}}
.brand-sub{{font-size:9px;color:rgba(255,255,255,.3);letter-spacing:.09em;text-transform:uppercase;margin-left:40px;font-family:var(--mono)}}
.filters{{padding:18px 22px;display:flex;flex-direction:column;gap:13px;border-bottom:1px solid rgba(255,255,255,.07)}}
.f-block{{display:flex;flex-direction:column;gap:5px}}
.f-label{{font-size:9px;color:rgba(255,255,255,.35);text-transform:uppercase;letter-spacing:.1em;font-family:var(--mono)}}
.f-select{{appearance:none;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);border-radius:8px;color:var(--sand-100);font-family:var(--font);font-size:12px;padding:8px 28px 8px 10px;cursor:pointer;outline:none;transition:all .15s;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%23BFA882' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 9px center}}
.f-select:focus,.f-select:hover{{border-color:var(--clay);background-color:rgba(255,255,255,.1)}}
.f-select option{{background:var(--earth-800);color:var(--sand-100)}}
.sidebar-nav{{padding:14px 12px;flex:1}}
.nav-item{{display:flex;align-items:center;gap:9px;padding:8px 10px;border-radius:7px;font-size:12px;color:rgba(255,255,255,.4);cursor:pointer;transition:all .15s;margin-bottom:2px}}
.nav-item:hover,.nav-item.active{{color:var(--sand-100);background:rgba(196,133,106,.16)}}
.nav-item svg{{width:14px;height:14px;flex-shrink:0}}
.sidebar-footer{{padding:14px 22px;border-top:1px solid rgba(255,255,255,.07)}}
.pilot-badge{{display:flex;align-items:center;gap:10px}}
.pilot-avatar{{width:34px;height:34px;border-radius:50%;background:var(--clay);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:500;color:white;flex-shrink:0;font-family:var(--mono)}}
.pilot-name-s{{font-size:11px;font-weight:500;color:var(--sand-100);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.pilot-pos-s{{font-size:10px;color:rgba(255,255,255,.35);font-family:var(--mono)}}
.main{{display:flex;flex-direction:column;min-height:100vh}}
.topbar{{background:var(--surface);border-bottom:1px solid var(--border);padding:14px 30px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10}}
.page-title{{font-family:var(--display);font-size:17px;color:var(--text)}}
.page-title span{{color:var(--clay)}}
.page-sub{{font-size:11px;color:var(--muted);margin-top:1px;font-family:var(--mono)}}
.topbar-right{{display:flex;align-items:center;gap:8px}}
.pill{{display:flex;align-items:center;gap:5px;padding:5px 11px;border-radius:20px;font-size:11px;font-family:var(--mono);border:1px solid var(--border);background:var(--surface2);color:var(--text2)}}
.dot{{width:6px;height:6px;border-radius:50%;background:var(--sage)}}
.content{{padding:22px 30px;display:flex;flex-direction:column;gap:16px;flex:1}}
.kpi-grid{{display:grid;grid-template-columns:repeat(5,1fr);gap:11px}}
.kpi{{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:14px 16px;box-shadow:var(--shadow);position:relative;overflow:hidden;transition:box-shadow .2s,transform .15s}}
.kpi:hover{{box-shadow:var(--shadow2);transform:translateY(-1px)}}
.kpi::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:var(--r2) var(--r2) 0 0}}
.kpi.k-clay::before{{background:var(--clay)}}.kpi.k-sage::before{{background:var(--sage)}}
.kpi.k-dusk::before{{background:var(--dusk)}}.kpi.k-sand::before{{background:var(--sand-400)}}
.kpi.k-rust::before{{background:var(--rust)}}
.kpi-label{{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;margin-bottom:6px;font-family:var(--mono)}}
.kpi-val{{font-size:25px;font-weight:600;color:var(--text);font-family:var(--mono);letter-spacing:-.03em;line-height:1}}
.kpi-unit{{font-size:12px;font-weight:400;color:var(--muted);margin-left:2px}}
.kpi-footer{{display:flex;align-items:center;justify-content:space-between;margin-top:7px}}
.kpi-vs{{font-size:10px;color:var(--muted)}}.kpi-vs b{{color:var(--text2);font-weight:500}}
.delta{{font-size:10px;font-family:var(--mono);padding:2px 6px;border-radius:4px}}
.d-up{{background:var(--sage-dim);color:var(--sage)}}.d-down{{background:var(--warm-red-dim);color:var(--warm-red)}}
.d-neu{{background:var(--sand-100);color:var(--muted)}}.d-warn{{background:rgba(181,96,58,.1);color:var(--rust)}}
.card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--r2);padding:18px 20px;box-shadow:var(--shadow)}}
.card-head{{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:16px}}
.card-title{{font-size:13px;font-weight:500;color:var(--text)}}.card-sub{{font-size:10px;color:var(--muted);margin-top:2px;font-family:var(--mono)}}
.legend{{display:flex;gap:12px;align-items:center;font-size:10px;color:var(--muted);font-family:var(--mono);flex-wrap:wrap}}
.leg{{display:flex;align-items:center;gap:5px}}
.charts-row{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
.chart-wrap{{position:relative;height:220px}}
.comp-table{{width:100%;border-collapse:collapse;font-size:12px}}
.comp-table th{{text-align:left;padding:7px 10px;font-size:10px;font-weight:500;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);border-bottom:1px solid var(--border);font-family:var(--mono);background:var(--sand-50)}}
.comp-table td{{padding:8px 10px;border-bottom:1px solid rgba(226,216,200,.5)}}
.comp-table tr:last-child td{{border-bottom:none}}.comp-table tr:hover td{{background:var(--sand-50)}}
.winner{{display:inline-flex;align-items:center;font-size:10px;padding:2px 7px;border-radius:4px;font-family:var(--mono);font-weight:500}}
.w-prog{{background:var(--dusk-dim);color:var(--dusk)}}.w-act{{background:var(--sage-dim);color:var(--sage)}}.w-eq{{background:var(--sand-100);color:var(--muted)}}
.bottom-row{{display:grid;grid-template-columns:1fr 300px;gap:14px}}
.prog-list{{display:flex;flex-direction:column;gap:13px}}
.prog-head{{display:flex;justify-content:space-between;font-size:11px;margin-bottom:4px}}
.prog-lbl{{color:var(--text2)}}.prog-num{{font-family:var(--mono);font-size:11px}}
.prog-track{{height:5px;background:var(--sand-200);border-radius:3px;overflow:hidden}}
.prog-fill{{height:100%;border-radius:3px}}
.prog-note{{font-size:9px;color:var(--dim);margin-top:2px;font-family:var(--mono)}}
.alert-list{{display:flex;flex-direction:column;gap:7px}}
.alert{{display:flex;align-items:flex-start;gap:9px;padding:9px 11px;border-radius:8px;border:1px solid}}
.alert.ok{{background:var(--sage-dim);border-color:rgba(122,158,126,.25)}}
.alert.warn{{background:rgba(181,96,58,.07);border-color:rgba(181,96,58,.2)}}
.alert.danger{{background:var(--warm-red-dim);border-color:rgba(196,83,74,.2)}}
.alert-dot{{width:7px;height:7px;border-radius:50%;flex-shrink:0;margin-top:3px}}
.alert.ok .alert-dot{{background:var(--sage)}}.alert.warn .alert-dot{{background:var(--rust)}}.alert.danger .alert-dot{{background:var(--warm-red)}}
.alert-title{{font-size:11px;font-weight:500;color:var(--text)}}.alert-desc{{font-size:10px;color:var(--muted);margin-top:1px;font-family:var(--mono)}}
.excl-note{{display:flex;align-items:center;gap:6px;padding:8px 11px;border-radius:7px;background:var(--sand-100);border:1px solid var(--border2);font-size:10px;color:var(--muted);margin-top:10px}}
.excl-note svg{{width:12px;height:12px;flex-shrink:0}}
::-webkit-scrollbar{{width:4px}}::-webkit-scrollbar-thumb{{background:var(--sand-300);border-radius:2px}}
@keyframes fadeUp{{from{{opacity:0;transform:translateY(7px)}}to{{opacity:1;transform:none}}}}
.kpi,.card{{animation:fadeUp .28s ease both}}
.kpi:nth-child(1){{animation-delay:.04s}}.kpi:nth-child(2){{animation-delay:.08s}}
.kpi:nth-child(3){{animation-delay:.12s}}.kpi:nth-child(4){{animation-delay:.16s}}
.kpi:nth-child(5){{animation-delay:.20s}}
</style>
</head>
<body>
<div class="shell">
<div class="sidebar">
  <div class="sidebar-top">
    <div class="brand">
      <div class="brand-icon"><svg fill="none" stroke="currentColor" stroke-width="1.8" viewBox="0 0 24 24"><path d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8"/></svg></div>
      <span class="brand-name">FDS Portal</span>
    </div>
    <div class="brand-sub">Flight Data System</div>
  </div>
  <div class="filters">
    <div class="f-block">
      <div class="f-label">Cargo</div>
      <select class="f-select" id="selGroup">
        <option value="">— Seleccionar cargo —</option>
        <option value="Capitán">Capitán</option>
        <option value="Primer Oficial">Primer Oficial</option>
        <option value="Instructor">Instructor</option>
      </select>
    </div>
    <div class="f-block">
      <div class="f-label">Tripulante</div>
      <select class="f-select" id="selPilot" disabled>
        <option value="">— Seleccione un cargo primero —</option>
      </select>
    </div>
  </div>
  <nav class="sidebar-nav">
    <div class="nav-item active"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>Resumen</div>
    <div class="nav-item"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>Horas bloque</div>
    <div class="nav-item"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>Horas deber</div>
    <div class="nav-item"><svg fill="none" stroke="currentColor" stroke-width="1.6" viewBox="0 0 24 24"><path d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z"/></svg>DAN 121</div>
  </nav>
  <div class="sidebar-footer">
    <div class="pilot-badge">
      <div class="pilot-avatar" id="sideAvatar">—</div>
      <div><div class="pilot-name-s" id="sideName">Sin selección</div><div class="pilot-pos-s" id="sidePos">—</div></div>
    </div>
  </div>
</div>
<div class="main">
  <div class="topbar">
    <div>
      <div class="page-title" id="pageTitle">Seleccione un <span>tripulante</span></div>
      <div class="page-sub" id="pageSub">FDS Portal · Productividad de tripulación</div>
    </div>
    <div class="topbar-right">
      <div class="pill"><span class="dot"></span>Sistema activo</div>
      <div class="pill" id="periodPill">Cargando períodos...</div>
    </div>
  </div>
  <div class="content" id="mainContent">
    <div id="placeholder" style="display:flex;flex-direction:column;align-items:center;justify-content:center;flex:1;gap:14px;color:var(--dim);padding:60px 0;">
      <svg width="44" height="44" fill="none" stroke="currentColor" stroke-width="1.2" viewBox="0 0 24 24" style="stroke:var(--border2)"><path d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8"/></svg>
      <div style="font-family:var(--display);font-size:18px;color:var(--sand-400)">FDS · Flight Data System</div>
      <div style="font-size:12px;text-align:center;max-width:300px;line-height:1.7;color:var(--muted)">Seleccione un cargo y un tripulante para visualizar sus indicadores de productividad.</div>
      <div style="font-size:10px;font-family:var(--mono);color:var(--dim);margin-top:4px" id="periodsHint"></div>
    </div>
    <div id="dashboard" style="display:none;flex-direction:column;gap:16px;">
      <div class="kpi-grid" id="kpiRow"></div>
      <div class="card">
        <div class="card-head">
          <div><div class="card-title">Horas Bloque · Evolución mensual</div><div class="card-sub" id="chartSub">Piloto vs. promedio del cargo (meses activos)</div></div>
          <div class="legend">
            <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="var(--clay)" stroke-width="2.5"/><circle cx="9" cy="4" r="3" fill="var(--clay)"/></svg><span>Piloto</span></div>
            <div class="leg"><svg width="18" height="8"><line x1="0" y1="4" x2="18" y2="4" stroke="var(--sand-400)" stroke-width="1.5" stroke-dasharray="4 3"/><circle cx="9" cy="4" r="2.5" fill="var(--sand-400)"/></svg><span>Prom. cargo</span></div>
            <div class="leg"><svg width="14" height="12"><polygon points="7,1 13,11 1,11" fill="none" stroke="var(--rust)" stroke-width="1.5"/></svg><span style="color:var(--rust)">Excluido prom.</span></div>
          </div>
        </div>
        <div class="chart-wrap"><canvas id="blockChart"></canvas></div>
        <div class="excl-note" id="exclNote" style="display:none">
          <svg fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
          <span id="exclText"></span>
        </div>
      </div>
      <div class="charts-row">
        <div class="card">
          <div class="card-head">
            <div><div class="card-title">Rol Programado vs. Efectuado</div><div class="card-sub">Horas bloque por período</div></div>
            <div class="legend">
              <div class="leg"><span style="width:10px;height:10px;border-radius:2px;background:var(--dusk);display:inline-block"></span><span>Programado</span></div>
              <div class="leg"><span style="width:10px;height:10px;border-radius:2px;background:var(--clay);display:inline-block"></span><span>Efectuado</span></div>
            </div>
          </div>
          <div class="chart-wrap"><canvas id="compareChart"></canvas></div>
        </div>
        <div class="card">
          <div class="card-head"><div class="card-title">Comparativo por Período</div><div class="card-sub" style="font-size:10px;color:var(--muted);font-family:var(--mono)">Programado vs. efectuado · Δ horas</div></div>
          <div id="compTableWrap"></div>
        </div>
      </div>
      <div class="bottom-row">
        <div class="card"><div class="card-head"><div class="card-title">Acumulado & Proyección</div><div class="card-sub">Basado en meses activos del piloto</div></div><div class="prog-list" id="progList"></div></div>
        <div class="card"><div class="card-head"><div class="card-title">Cumplimiento DAN 121</div><div class="card-sub">Último período disponible</div></div><div class="alert-list" id="alertList"></div></div>
      </div>
    </div>
  </div>
</div>
</div>
<script>
const RAW={data_json};
const PERIODS={periods_json};
const PERIOD_LABELS={labels_json};

document.getElementById('periodPill').textContent=Object.values(PERIOD_LABELS).join(' · ');
document.getElementById('periodsHint').textContent='Períodos: '+Object.values(PERIOD_LABELS).join(' · ');

let blockChartInst=null,compareChartInst=null;
const selGroup=document.getElementById('selGroup');
const selPilot=document.getElementById('selPilot');

selGroup.addEventListener('change',()=>{{
  const g=selGroup.value;
  const names=[...new Set(RAW.filter(r=>r.pos_group===g).map(r=>r.name))].sort((a,b)=>a.localeCompare(b,'es'));
  selPilot.innerHTML='<option value="">— Seleccionar tripulante —</option>';
  names.forEach(n=>{{const o=document.createElement('option');o.value=o.textContent=n;selPilot.appendChild(o);}});
  selPilot.disabled=false;
  document.getElementById('placeholder').style.display='flex';
  document.getElementById('dashboard').style.display='none';
}});

selPilot.addEventListener('change',()=>{{if(selPilot.value)render(selPilot.value,selGroup.value);}});

function fmt(v,d=1){{if(v==null||v===0)return'—';return(+v).toFixed(d);}}
function avg(arr){{const v=arr.filter(x=>x!=null&&x>0);return v.length?v.reduce((a,b)=>a+b,0)/v.length:0;}}
function dc(d){{return d>2?'d-up':d<-2?'d-down':'d-neu';}}
function ds(d){{return(d>=0?'+':'')+d.toFixed(1)+'%';}}
function makeGrad(ctx,ca,c1,c2){{if(!ca)return'transparent';const g=ctx.createLinearGradient(0,ca.top,0,ca.bottom);g.addColorStop(0,c1);g.addColorStop(1,c2);return g;}}

function render(pilotName,group){{
  document.getElementById('placeholder').style.display='none';
  const dash=document.getElementById('dashboard');
  dash.style.display='flex';
  const pr=RAW.filter(r=>r.name===pilotName);
  const gr=RAW.filter(r=>r.pos_group===group);
  const latest=pr.filter(r=>r.block_h_actual>0).sort((a,b)=>b.period.localeCompare(a.period))[0]||pr[0];
  const lp=latest?.period||PERIODS[PERIODS.length-1];
  const init=pilotName.split(' ').filter((_,i)=>i<2).map(w=>w[0]).join('');
  document.getElementById('sideAvatar').textContent=init;
  document.getElementById('sideName').textContent=pilotName.split(' ').slice(0,2).join(' ');
  document.getElementById('sidePos').textContent=(latest?.pos||group)+' · '+(latest?.base||'');
  document.getElementById('pageTitle').innerHTML='<span>'+pilotName.split(' ').slice(0,2).join(' ')+'</span> · Productividad';
  document.getElementById('pageSub').textContent=(latest?.pos_group||group)+' · '+(latest?.base||'')+' · '+Object.values(PERIOD_LABELS).join(' · ');
  const ga=gr.filter(r=>r.period===lp&&!r.exclude_from_avg&&r.block_h_actual>0);
  const ab=avg(ga.map(r=>r.block_h_actual)),ad=avg(ga.map(r=>r.duty_h_actual)),al=avg(ga.map(r=>r.libre_days));
  const mb=latest?.block_h_actual||0,md=latest?.duty_h_actual||0,ml=latest?.libre_days||0;
  const bd=ab>0?(mb-ab)/ab*100:0,dd=ad>0?(md-ad)/ad*100:0;
  const actP=pr.filter(r=>!r.exclude_from_avg&&r.block_h_actual>0);
  const accB=actP.reduce((s,r)=>s+(r.block_h_actual||0),0);
  const accProg=pr.reduce((s,r)=>s+(r.block_h_programmed||0),0);
  const accAct=pr.reduce((s,r)=>s+(r.block_h_actual||0),0);
  const pva=accProg>0?(accAct-accProg)/accProg*100:0;
  document.getElementById('kpiRow').innerHTML=`
    <div class="kpi k-clay"><div class="kpi-label">Bloque · ${{PERIOD_LABELS[lp]||lp}}</div><div class="kpi-val">${{fmt(mb)}}<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">Prom.: <b>${{fmt(ab)}}h</b></span><span class="delta ${{dc(bd)}}">${{ds(bd)}}</span></div></div>
    <div class="kpi k-sand"><div class="kpi-label">Deber · ${{PERIOD_LABELS[lp]||lp}}</div><div class="kpi-val">${{fmt(md)}}<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">Prom.: <b>${{fmt(ad)}}h</b></span><span class="delta ${{dc(dd)}}">${{ds(dd)}}</span></div></div>
    <div class="kpi k-sage"><div class="kpi-label">Días libres · ${{PERIOD_LABELS[lp]||lp}}</div><div class="kpi-val">${{ml}}<span class="kpi-unit">d</span></div><div class="kpi-footer"><span class="kpi-vs">Prom.: <b>${{fmt(al,0)}}d</b></span><span class="delta ${{dc(ml-al)}}">${{ml-al>=0?'+':''}}&nbsp;${{(ml-al).toFixed(0)}}d</span></div></div>
    <div class="kpi k-dusk"><div class="kpi-label">Bloque acum.</div><div class="kpi-val">${{fmt(accB,0)}}<span class="kpi-unit">h</span></div><div class="kpi-footer"><span class="kpi-vs">${{actP.length}} meses activos</span><span class="delta d-neu">/${{PERIODS.length}}m</span></div></div>
    <div class="kpi k-rust"><div class="kpi-label">Prog. vs Efectuado</div><div class="kpi-val">${{fmt(Math.abs(pva),1)}}<span class="kpi-unit">%</span></div><div class="kpi-footer"><span class="kpi-vs">P:<b>${{fmt(accProg,0)}}h</b> E:<b>${{fmt(accAct,0)}}h</b></span><span class="delta ${{pva>=0?'d-up':'d-down'}}">${{pva>=0?'▲':'▼'}} efectuado</span></div></div>`;
  // line chart
  const excl=pr.filter(r=>r.exclude_from_avg).map(r=>r.period);
  const pData=PERIODS.map(p=>{{const r=pr.find(x=>x.period===p);return r?(r.block_h_actual||0):null;}});
  const gData=PERIODS.map(p=>{{const a=gr.filter(r=>r.period===p&&!r.exclude_from_avg&&r.block_h_actual>0);return a.length?avg(a.map(r=>r.block_h_actual)):null;}});
  const bc=document.getElementById('blockChart').getContext('2d');
  if(blockChartInst)blockChartInst.destroy();
  blockChartInst=new Chart(bc,{{type:'line',data:{{labels:PERIODS.map(p=>PERIOD_LABELS[p]||p),datasets:[
    {{label:'Piloto',data:pData,borderColor:'#C4856A',backgroundColor(c){{return makeGrad(bc,c.chart.chartArea,'rgba(196,133,106,.15)','rgba(196,133,106,.01)');}},borderWidth:2.5,pointRadius(c){{return excl.includes(PERIODS[c.dataIndex])?6:4;}},pointStyle(c){{return excl.includes(PERIODS[c.dataIndex])?'triangle':'circle';}},pointBackgroundColor(c){{return excl.includes(PERIODS[c.dataIndex])?'#B5603A':'#C4856A';}},pointBorderColor(c){{return excl.includes(PERIODS[c.dataIndex])?'#B5603A':'#C4856A';}},pointHoverRadius:7,tension:.35,fill:true,spanGaps:true,order:1}},
    {{label:'Prom. cargo',data:gData,borderColor:'#BFA882',borderWidth:1.5,borderDash:[5,4],pointBackgroundColor:'#BFA882',pointRadius:3,pointHoverRadius:5,tension:.35,fill:false,spanGaps:false,order:2}}
  ]}},options:{{responsive:true,maintainAspectRatio:false,interaction:{{mode:'index',intersect:false}},plugins:{{legend:{{display:false}},tooltip:{{backgroundColor:'#FAF7F2',borderColor:'#E2D8C8',borderWidth:1,titleColor:'#2E2416',bodyColor:'#A68B5B',padding:11,titleFont:{{family:"'DM Sans',sans-serif",size:12,weight:500}},bodyFont:{{family:"'DM Mono',monospace",size:11}},callbacks:{{title(i){{const p=PERIODS[i[0].dataIndex];return(PERIOD_LABELS[p]||p)+(excl.includes(p)?' · ⚠ excluido del prom.':'');}},label(i){{if(i.raw==null)return null;return'  '+i.dataset.label+': '+i.raw.toFixed(1)+'h';}},afterBody(i){{const p=PERIODS[i[0].dataIndex];const my=pData[i[0].dataIndex],av=gData[i[0].dataIndex];if(av==null||my==null)return['  (mes excluido del cálculo)'];const d=my-av;return['  vs promedio: '+(d>=0?'+':'')+d.toFixed(1)+'h'];}}}}}}}},scales:{{x:{{grid:{{color:'rgba(226,216,200,.6)',drawBorder:false}},ticks:{{color:'#A68B5B',font:{{size:11,family:"'DM Mono',monospace"}}}},border:{{display:false}}}},y:{{min:0,grid:{{color:'rgba(226,216,200,.6)',drawBorder:false}},ticks:{{color:'#A68B5B',font:{{size:11,family:"'DM Mono',monospace"}},callback:v=>v+'h'}},border:{{display:false}}}}}}}}}}});
  const en=document.getElementById('exclNote');
  const ep=excl.map(p=>PERIOD_LABELS[p]||p).filter(Boolean);
  if(ep.length){{en.style.display='flex';document.getElementById('exclText').textContent='Meses excluidos del promedio comparativo (ausencias prolongadas): '+ep.join(', ')+'. Los datos del piloto se muestran igualmente en el gráfico.';}}
  else en.style.display='none';
  // bar chart
  const prog=PERIODS.map(p=>{{const r=pr.find(x=>x.period===p);return r?(r.block_h_programmed||0):0;}});
  const act=PERIODS.map(p=>{{const r=pr.find(x=>x.period===p);return r?(r.block_h_actual||0):0;}});
  const cc=document.getElementById('compareChart').getContext('2d');
  if(compareChartInst)compareChartInst.destroy();
  compareChartInst=new Chart(cc,{{type:'bar',data:{{labels:PERIODS.map(p=>PERIOD_LABELS[p]||p),datasets:[
    {{label:'Programado',data:prog,backgroundColor:'rgba(139,123,168,.55)',borderColor:'#8B7BA8',borderWidth:1,borderRadius:5,borderSkipped:false}},
    {{label:'Efectuado',data:act,backgroundColor:'rgba(196,133,106,.55)',borderColor:'#C4856A',borderWidth:1,borderRadius:5,borderSkipped:false}}
  ]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{backgroundColor:'#FAF7F2',borderColor:'#E2D8C8',borderWidth:1,titleColor:'#2E2416',bodyColor:'#A68B5B',padding:11,bodyFont:{{family:"'DM Mono',monospace",size:11}},callbacks:{{label(i){{return'  '+i.dataset.label+': '+i.raw.toFixed(1)+'h';}},afterBody(i){{const idx=i[0].dataIndex;const d=act[idx]-prog[idx];if(prog[idx]===0&&act[idx]===0)return['  Sin datos'];const w=d>.5?'▲ Efectuado mayor':d<-.5?'▲ Programado mayor':'≈ Similares';return['  Diferencia: '+(d>=0?'+':'')+d.toFixed(1)+'h  '+w];}}}}}}}},scales:{{x:{{grid:{{display:false}},ticks:{{color:'#A68B5B',font:{{size:11,family:"'DM Mono',monospace"}}}},border:{{display:false}}}},y:{{min:0,grid:{{color:'rgba(226,216,200,.6)',drawBorder:false}},ticks:{{color:'#A68B5B',font:{{size:11,family:"'DM Mono',monospace"}},callback:v=>v+'h'}},border:{{display:false}}}}}}}}}}});
  // comparison table
  let tbl='<table class="comp-table"><thead><tr><th>Período</th><th>Programado</th><th>Efectuado</th><th>Δ</th><th>Mayor</th></tr></thead><tbody>';
  PERIODS.forEach(p=>{{
    const r=pr.find(x=>x.period===p);if(!r)return;
    const pg=r.block_h_programmed||0,ac=r.block_h_actual||0,d=ac-pg;
    const ex=r.exclude_from_avg?'<span style="color:var(--rust);font-size:9px"> ✱</span>':'';
    const w=pg===0&&ac===0?'<span class="winner w-eq">Sin datos</span>':pg===0?'<span class="winner w-act">Solo efectuado</span>':d>.5?'<span class="winner w-act">▲ Efectuado</span>':d<-.5?'<span class="winner w-prog">▲ Programado</span>':'<span class="winner w-eq">≈ Iguales</span>';
    const ds=pg>0?((d>=0?'+':'')+d.toFixed(1)+'h'):'—';
    tbl+=`<tr><td style="font-family:var(--mono);font-size:11px;color:var(--text2)">${{PERIOD_LABELS[p]||p}}${{ex}}</td><td style="font-family:var(--mono)">${{pg>0?pg.toFixed(1)+'h':'—'}}</td><td style="font-family:var(--mono)">${{ac>0?ac.toFixed(1)+'h':'—'}}</td><td style="font-family:var(--mono);color:${{d>=0?'var(--sage)':'var(--warm-red)'}}">${{ds}}</td><td>${{w}}</td></tr>`;
  }});
  if(excl.length)tbl+='<tr><td colspan="5" style="font-size:9px;color:var(--muted);font-family:var(--mono);padding:6px 10px">✱ Excluido del promedio comparativo</td></tr>';
  tbl+='</tbody></table>';
  document.getElementById('compTableWrap').innerHTML=tbl;
  // progress
  const pct1=Math.min(accB/1000*100,100);
  const avgM=actP.length?accB/actP.length:0;
  const proj=avgM*12,pctP=Math.min(proj/1000*100,100);
  const totL=pr.reduce((s,r)=>s+(r.libre_days||0),0);
  const avgL=pr.length?totL/pr.length:0;
  document.getElementById('progList').innerHTML=`
    <div><div class="prog-head"><span class="prog-lbl">Horas bloque acumuladas</span><span class="prog-num" style="color:var(--clay)">${{accB.toFixed(0)}}h</span></div><div class="prog-track"><div class="prog-fill" style="width:${{pct1}}%;background:var(--clay)"></div></div><div class="prog-note">Límite DAN 121: 1.000h/año · ${{(100-pct1).toFixed(1)}}% disponible</div></div>
    <div><div class="prog-head"><span class="prog-lbl">Proyección a 12 meses</span><span class="prog-num" style="color:var(--sand-500)">~${{proj.toFixed(0)}}h est.</span></div><div class="prog-track"><div class="prog-fill" style="width:${{pctP}}%;background:linear-gradient(90deg,var(--clay),var(--sand-400))"></div></div><div class="prog-note">Prom. ${{avgM.toFixed(1)}}h/mes en meses activos</div></div>
    <div><div class="prog-head"><span class="prog-lbl">Descanso promedio</span><span class="prog-num" style="color:var(--sage)">${{avgL.toFixed(1)}} d/mes</span></div><div class="prog-track"><div class="prog-fill" style="width:${{Math.min(avgL/20*100,100)}}%;background:var(--sage)"></div></div><div class="prog-note">Mínimo reglamentario DAN 121: 8 días/mes</div></div>
    <div><div class="prog-head"><span class="prog-lbl">Meses activos</span><span class="prog-num">${{actP.length}} / ${{PERIODS.length}}</span></div><div style="display:flex;gap:3px;margin-top:4px"><div style="height:5px;border-radius:2px 0 0 2px;background:var(--sage);flex:${{actP.length}}"></div><div style="height:5px;border-radius:0 2px 2px 0;background:var(--rust);opacity:.5;flex:${{Math.max(PERIODS.length-actP.length,0)}}"></div></div><div class="prog-note">${{excl.length?excl.map(p=>PERIOD_LABELS[p]||p).join(', ')+' excluidos':'Sin ausencias prolongadas'}}</div></div>`;
  // alerts
  function al(t,title,desc){{return`<div class="alert ${{t}}"><div class="alert-dot"></div><div><div class="alert-title">${{title}}</div><div class="alert-desc">${{desc}}</div></div></div>`;}}
  let alerts='';
  alerts+=al(mb>100?'danger':mb>85?'warn':'ok',`Bloque mensual · ${{fmt(mb)}}h`,mb>100?'Supera límite DAN 121 de 100h/mes':mb>85?'Cercano al límite de 100h/mes':'Dentro del límite (100h/mes)');
  alerts+=al(accB>900?'danger':accB>750?'warn':'ok',`Bloque acumulado · ${{accB.toFixed(0)}}h`,accB>900?'Muy cerca del límite anual de 1.000h':accB>750?'Supera el 75% del límite anual':'Sin riesgo límite anual ('+(1000-accB).toFixed(0)+'h disponibles)');
  alerts+=al(ml<8?'danger':ml<10?'warn':'ok',`Días libres · ${{ml}}d`,ml<8?'Bajo el mínimo reglamentario (8d/mes)':ml<10?'Dentro del mínimo, bajo el promedio del cargo':'Descanso adecuado según DAN 121');
  alerts+=al(md>130?'danger':md>105?'warn':'ok',`Horas deber · ${{fmt(md)}}h`,md>130?'Horas deber muy elevadas, revisar FDPs':md>105?'Sobre promedio del cargo':'Dentro de rango normal');
  alerts+='<div style="margin-top:6px;padding:9px 11px;background:var(--sand-100);border-radius:7px;font-size:10px;color:var(--muted);line-height:1.5;font-family:var(--mono)">Alertas indicativas. El cálculo oficial de FDP y límites es responsabilidad de Operaciones.</div>';
  document.getElementById('alertList').innerHTML=alerts;
}}
</script>
</body>
</html>"""
    return html

if __name__ == '__main__':
    records, periods = build_dataset()
    html = generate_html(records, periods)
    OUTPUT_HTML.write_text(html, encoding='utf-8')
    print(f"\n✓ Dashboard generado: {OUTPUT_HTML}")
    print(f"  Tamaño: {len(html)/1024:.0f} KB")
