# coding: utf-8
"""
update_vars.py — Atualiza TODAS as variáveis JS do dashboard Motiva.
Regras completas em REGRAS_DASHBOARD.md
"""
import pandas as pd
import re, os, sys, json, shutil, subprocess, zipfile
from datetime import datetime

# ──────────────────────────────────────────────────────────────
# 0. Setup
# ──────────────────────────────────────────────────────────────
def find_mnt():
    for s in (os.listdir('/sessions') if os.path.isdir('/sessions') else []):
        p = f'/sessions/{s}/mnt'
        if os.path.isdir(p): return p
    return None

MNT = find_mnt()
if not MNT:
    print("ERRO: pasta de sessao nao encontrada"); sys.exit(1)

WORKSPACE   = f"{MNT}/KIM+ BI - Motiva"
CONFIG_FILE = f"{WORKSPACE}/Automacao/github_config.txt"
GIT_DIR     = f"/tmp/motiva-update-{datetime.now().strftime('%H%M%S')}"

def find_excel():
    for folder in [WORKSPACE, f"{MNT}/uploads"]:
        if not os.path.isdir(folder): continue
        cands = sorted(
            [os.path.join(folder, f) for f in os.listdir(folder)
             if 'Motiva' in f and f.endswith('.xlsx')],
            key=os.path.getmtime, reverse=True)
        for path in cands:
            try: zipfile.ZipFile(path, 'r').close(); return path
            except: continue
    return None

EXCEL = find_excel()
if not EXCEL:
    print("ERRO: Excel nao encontrado"); sys.exit(1)

print(f"Excel : {EXCEL}")
print(f"Workspace: {WORKSPACE}")

MESES = {'01':'Jan','02':'Fev','03':'Mar','04':'Abr','05':'Mai','06':'Jun',
         '07':'Jul','08':'Ago','09':'Set','10':'Out','11':'Nov','12':'Dez'}

def mlabel(m):
    yr, mo = m.split('_'); return f"{MESES[mo]}/{yr[2:]}"

def replace_const(html, name, vals, strings=False):
    if strings:
        vs = ','.join(f"'{v}'" for v in vals)
    else:
        vs = ','.join(str(int(v)) if float(v) == int(float(v)) else str(round(v,2)) for v in vals)
    pat = rf'(const\s+{re.escape(name)}\s*=\s*\[)[^\]]*(\])'
    new, n = re.subn(pat, rf'\g<1>{vs}\g<2>', html)
    print(f"  {'OK' if n else 'XX'} {name}: {len(vals)} valores")
    return new

# ──────────────────────────────────────────────────────────────
# 1. Acompanhamento
# ──────────────────────────────────────────────────────────────
print("\n[1/6] Acompanhamento...")
df_ac = pd.read_excel(EXCEL, sheet_name='Acompanhamento', header=None, engine='openpyxl')
months_row = df_ac.iloc[1, 7:]
m_info = [(i+7, v) for i, v in enumerate(months_row) if isinstance(v, str) and '_' in str(v)]
m_labels = [v for _, v in m_info]
m_idxs   = [i for i, _ in m_info]

def get_row(r0):
    row = df_ac.iloc[r0]
    return [round(float(row.iloc[m_idxs[j]])) if pd.notna(row.iloc[m_idxs[j]]) else 0
            for j in range(len(m_labels))]

gmv_vt      = get_row(10)
gmv_pat     = get_row(14)
gmv_total   = get_row(18)
rec_pat     = get_row(36)
financ_rec  = get_row(31)   # Financiamento — exibido separado, NÃO entra em vtTxSucesso
via2_rec    = get_row(29)   # 2° Via receitas

LABELS = [mlabel(m) for m in m_labels]
last_complete = next((m for m in reversed(m_labels) if get_row(14)[m_labels.index(m)] > 0), m_labels[-1])
print(f"  {len(m_labels)} meses: {LABELS[0]} → {LABELS[-1]}  (último completo: {mlabel(last_complete)})")

# ──────────────────────────────────────────────────────────────
# 2. PCC — Pedido de Carga Completo
# ──────────────────────────────────────────────────────────────
print("\n[2/6] PCC...")
df_pcc = pd.read_excel(EXCEL, sheet_name='Pedido de carga Completo',
    usecols=['mes_pagamento','taxa_administrativa','taxa_aproveitamento',  # COL R para Tx Sucesso
             'taxa_entrega','nota_fiscal','Vendedor','Comissão','Codigo_bilhetagem'],
    engine='openpyxl')
df_pcc['mes_pagamento'] = df_pcc['mes_pagamento'].astype(str)
for c in ['taxa_administrativa','taxa_aproveitamento','taxa_entrega','nota_fiscal','Comissão']:
    df_pcc[c] = pd.to_numeric(df_pcc[c], errors='coerce').fillna(0)
vt_tax = df_pcc.groupby('mes_pagamento')[
    ['taxa_administrativa','taxa_aproveitamento','taxa_entrega','nota_fiscal']].sum()

# Cartões processados (contagem de bilhetagens por mês)
cart_pcc = df_pcc.groupby('mes_pagamento')['Codigo_bilhetagem'].count()
cart_mes = [int(cart_pcc.loc[m]) if m in cart_pcc.index else 0 for m in m_labels]

# Comissões PCC por vendedor
df_vc = df_pcc[df_pcc['Vendedor'].notna()].copy()
df_vc['Vendedor'] = df_vc['Vendedor'].astype(str).str.strip()
df_vc = df_vc[~df_vc['Vendedor'].isin(['0','nan',''])]
comm_vt = df_vc.pivot_table(index='mes_pagamento', columns='Vendedor',
                             values='Comissão', aggfunc='sum', fill_value=0)
comm_vt.columns = [c.strip() for c in comm_vt.columns]

# ──────────────────────────────────────────────────────────────
# 3. PE — Pedido Externo (todos são VT)
# ──────────────────────────────────────────────────────────────
print("\n[3/6] PE...")
df_pe_full = pd.read_excel(EXCEL, sheet_name='Pedido externo', engine='openpyxl', header=0)
df_pe_full.columns = [str(c).strip() for c in df_pe_full.columns]
# Coluna E (Mês) como referência de mês
df_pe_full['mes_str'] = pd.to_datetime(df_pe_full['Mês'], errors='coerce').apply(
    lambda d: f'{d.year}_{d.month:02d}' if pd.notna(d) else None)
for c in ['Comissão','Taxa Adm','Tx de Sucesso','Valor do Boleto Motiva',
          'Valor Recarga (I)','Valor Recarga (F)','Projeção de Saldo','NF']:
    if c in df_pe_full.columns:
        df_pe_full[c] = pd.to_numeric(df_pe_full[c], errors='coerce').fillna(0)
df_pe_full['Vendedor'] = df_pe_full['Vendedor'].astype(str).str.strip()
df_pe_full = df_pe_full.dropna(subset=['mes_str'])

# Comissões PE: TODOS os pedidos (com e sem SK têm comissão real)
comm_pat = df_pe_full.pivot_table(index='mes_str', columns='Vendedor',
                                   values='Comissão', aggfunc='sum', fill_value=0)

# Taxas PE: TODOS os pedidos (são todos VT, suas taxas não estão no PCC)
pe_tax = df_pe_full.groupby('mes_str')[['Taxa Adm','Tx de Sucesso']].sum()

def get_tax(col): return [round(float(vt_tax.loc[m, col])) if m in vt_tax.index else 0 for m in m_labels]
def get_pe_tax(col): return [round(float(pe_tax.loc[m, col])) if m in pe_tax.index else 0 for m in m_labels]

pe_adm = get_pe_tax('Taxa Adm')
pe_suc = get_pe_tax('Tx de Sucesso')

# vtTxAdm = PCC col P + PE col N (TODOS pedidos PE)
vt_adm = [get_tax('taxa_administrativa')[i] + pe_adm[i] for i in range(len(m_labels))]

# vtTxSucesso = PCC col R (taxa_aproveitamento) + PE col O — SEM subtrair financiamento
# (Col R já exclui financiamento por natureza; verificado: soma bate com recVT para Jan-Abr/26)
vt_sucesso = [get_tax('taxa_aproveitamento')[i] + pe_suc[i] for i in range(len(m_labels))]

nf_pcc    = get_tax('nota_fiscal')
vt_entrega = via2_rec  # 2°Via = Acompanhamento linha 29

# recVT = vtTxAdm + vtTxSucesso + vtTxEntrega (calculado aqui, substitui via replace_const)
rec_vt = [vt_adm[i] + vt_sucesso[i] + vt_entrega[i] for i in range(len(m_labels))]

def get_comm(seller):
    vals = []
    for m in m_labels:
        v1 = float(comm_vt.loc[m, seller]) if (m in comm_vt.index and seller in comm_vt.columns) else 0.0
        v2 = float(comm_pat.loc[m, seller]) if (m in comm_pat.index and seller in comm_pat.columns) else 0.0
        vals.append(round(v1 + v2, 2))
    return vals

comm_karine   = get_comm('Karine')
comm_carol    = get_comm('Carol')
comm_andre    = get_comm('André')
comm_lysandra = get_comm('Lysandra')
comm_vanessa  = get_comm('Vanessa')
comm_pedro    = get_comm('Pedro')
comm_total    = [round(sum(x), 2) for x in zip(comm_karine, comm_carol, comm_andre,
                                                comm_lysandra, comm_vanessa, comm_pedro)]

# Verificação Abr/26
li = m_labels.index('2026_04') if '2026_04' in m_labels else -1
if li >= 0:
    soma = rec_vt[li]
    print(f"  Abr/26 vtTxAdm={vt_adm[li]:,.0f} + vtTxSuc={vt_sucesso[li]:,.0f} + vtTxEnt={vt_entrega[li]:,.0f} = recVT={soma:,.0f}")
    print(f"  Cartões Abr/26: {cart_mes[li]:,}")

# ──────────────────────────────────────────────────────────────
# 4. Clientes + Segmentação
# ──────────────────────────────────────────────────────────────
print("\n[4/6] Clientes e segmentacao...")
df_cli = pd.read_excel(EXCEL, sheet_name='Clientes', engine='openpyxl', usecols=['mês cadastro'])
df_cli['mes_str'] = pd.to_datetime(df_cli['mês cadastro'], errors='coerce').apply(
    lambda d: f'{d.year}_{d.month:02d}' if pd.notna(d) else None)
new_per_month = df_cli.groupby('mes_str').size().to_dict()
base = sum(v for k, v in new_per_month.items() if k < m_labels[0])
qtd_clients, qtd_new = [], []
running = base
for m in m_labels:
    running += new_per_month.get(m, 0)
    qtd_clients.append(running)
    qtd_new.append(new_per_month.get(m, 0))

df_seg_hdr = pd.read_excel(EXCEL, sheet_name='GMV por Cliente', engine='openpyxl', nrows=1, header=None)
raw_hdrs   = list(df_seg_hdr.iloc[0, 7:])
df_seg     = pd.read_excel(EXCEL, sheet_name='GMV por Cliente', engine='openpyxl', skiprows=1)
df_seg.columns = ['cod','empresa','cnpj','vinculo','dt_cadastro','segmentacao','maior_gmv'] + \
                 [f'gmv_{i}' for i in range(len(df_seg.columns)-7)]
df_seg['segmentacao'] = df_seg['segmentacao'].astype(str).str.strip()
seg_col_map = {str(h): f"gmv_{i}" for i, h in enumerate(raw_hdrs) if h}
seg_col = seg_col_map.get(last_complete)
seg_out = {}
if seg_col and seg_col in df_seg.columns:
    df_seg[seg_col] = pd.to_numeric(df_seg[seg_col], errors='coerce').fillna(0)
    active = df_seg[df_seg[seg_col] > 0]
    SEG_MAP = {'Diamante':'Diamante','Ouro':'Ouro','Prata':'Prata',
               'Bronze - P':'Bronze-P','Bronze - PP':'Bronze-PP'}
    for seg, cnt in active.groupby('segmentacao').size().items():
        k = SEG_MAP.get(seg)
        if k: seg_out[k] = int(cnt)
print(f"  Segmentacao ({mlabel(last_complete)}): {seg_out}")

SEG_IDS = {'Diamante':0,'Ouro':1,'Prata':2,'Bronze - P':3,'Bronze - PP':4,'Prata ':2}
def fmt_cnpj(v):
    s = str(v).replace('.0','').strip().zfill(14)
    try: return f"{s[:2]}.{s[2:5]}.{s[5:8]}/{s[8:12]}-{s[12:14]}"
    except: return str(v)
def fmt_date(v):
    try: d = pd.to_datetime(v); return f"{d.month:02d}/{str(d.year)[2:]}"
    except: return ''

seg_month_cols = [seg_col_map.get(m) for m in m_labels]
for col in seg_month_cols:
    if col and col in df_seg.columns:
        df_seg[col] = pd.to_numeric(df_seg[col], errors='coerce').fillna(0)
cd_parts = []
for _, row in df_seg.iterrows():
    if not row.get('empresa') or str(row['empresa']).strip() in ('', 'nan'): continue
    gmv_vals = [round(float(row[col]), 1) if (col and col in df_seg.columns) else 0
                for col in seg_month_cols]
    if sum(gmv_vals) == 0: continue
    cnpj  = fmt_cnpj(row.get('cnpj', ''))
    emp   = str(row['empresa']).strip().replace('"', '\\"')
    seg_i = SEG_IDS.get(str(row.get('segmentacao','')).strip(), 5)
    cad   = fmt_date(row.get('dt_cadastro'))
    cd_parts.append(f'["{cnpj}","{emp}",{seg_i},"{cad}",0,{",".join(str(v) for v in gmv_vals)},0]')
CD_JS = 'const _CD=[\n' + ',\n'.join(cd_parts) + '\n];'
print(f"  _CD: {len(cd_parts)} clientes")

# ──────────────────────────────────────────────────────────────
# 5. VD — Dados por empresa/vendedor (Aba Comissões)
# ──────────────────────────────────────────────────────────────
print("\n[5/6] VD (comissoes por empresa)...")
SELLERS_VD = ['Karine','André','Carol','Lysandra','Vanessa','Pedro']
df_pcc_vd = pd.read_excel(EXCEL, sheet_name='Pedido de carga Completo', engine='openpyxl',
    usecols=['mes_pagamento','CNPJ','nome_fantasia','total_pedido','nota_fiscal','Vendedor','Comissão'])
df_pcc_vd['mes_pagamento'] = df_pcc_vd['mes_pagamento'].astype(str)
for c in ['total_pedido','nota_fiscal','Comissão']:
    df_pcc_vd[c] = pd.to_numeric(df_pcc_vd[c], errors='coerce').fillna(0)
df_pcc_vd['Vendedor'] = df_pcc_vd['Vendedor'].astype(str).str.strip()
df_pcc_vd = df_pcc_vd[df_pcc_vd['Vendedor'].isin(SELLERS_VD)]
df_pcc_vd['CNPJ'] = df_pcc_vd['CNPJ'].astype(str).str.strip()
df_pcc_vd['nome_fantasia'] = df_pcc_vd['nome_fantasia'].astype(str).str.strip()

df_pe_vd = df_pe_full[df_pe_full['Vendedor'].isin(SELLERS_VD)].copy()
df_pe_vd['CNPJ'] = df_pe_vd['CNPJ'].astype(str).str.strip()
df_pe_vd['Cliente'] = df_pe_vd['Cliente'].astype(str).str.strip()

m_idx_map_vd = {m: i for i, m in enumerate(m_labels)}
NM = len(m_labels)

def build_seller_vd(seller):
    entries = {}
    for _, row in df_pcc_vd[df_pcc_vd['Vendedor']==seller].iterrows():
        m = row['mes_pagamento']
        if m not in m_idx_map_vd: continue
        key = (str(row['nome_fantasia']), str(row['CNPJ']))
        if key not in entries: entries[key] = {'gmv':[0.]*NM,'nf':[0.]*NM,'comm':[0.]*NM}
        i = m_idx_map_vd[m]
        entries[key]['gmv'][i] += float(row['total_pedido'])
        entries[key]['nf'][i]  += float(row['nota_fiscal'])
        entries[key]['comm'][i]+= float(row['Comissão'])
    for _, row in df_pe_vd[df_pe_vd['Vendedor']==seller].iterrows():
        m = row['mes_str']
        if not m or m not in m_idx_map_vd: continue
        key = (str(row['Cliente']), str(row['CNPJ']))
        if key not in entries: entries[key] = {'gmv':[0.]*NM,'nf':[0.]*NM,'comm':[0.]*NM}
        i = m_idx_map_vd[m]
        is_sem_sk = 'SEM SK' in str(row.get('BaseSK','')).upper()
        if is_sem_sk:
            entries[key]['gmv'][i] += float(row.get('Valor do Boleto Motiva', 0))
        entries[key]['nf'][i]  += float(row.get('NF', 0))
        entries[key]['comm'][i]+= float(row['Comissão'])
    result = []
    for (nome, cnpj), vals in entries.items():
        if sum(vals['comm']) > 0 or sum(vals['gmv']) > 0:
            result.append([nome, cnpj,
                           [round(v,1) for v in vals['gmv']],
                           [round(v,2) for v in vals['nf']],
                           [round(v,2) for v in vals['comm']]])
    result.sort(key=lambda x: -sum(x[4]))
    return result

def arr_to_js(arr):
    return '['+','.join(str(int(v)) if v==int(v) else str(round(v,2)) for v in arr)+']'
def build_vd_js(d):
    parts = []
    for s, entries in d.items():
        inner = ','.join(
            f'["{e[0].replace(chr(34),chr(92)+chr(34))}","{e[1]}",{arr_to_js(e[2])},{arr_to_js(e[3])},{arr_to_js(e[4])}]'
            for e in entries)
        parts.append(f'"{s}":[{inner}]')
    return '{' + ','.join(parts) + '}'

vd_dict = {s: build_seller_vd(s) for s in SELLERS_VD}
vd_js_str = 'const VD=' + build_vd_js(vd_dict) + ';'
for s, v in vd_dict.items():
    print(f"  {s}: {len(v)} empresas, comm={sum(sum(e[4]) for e in v):.2f}")

# ──────────────────────────────────────────────────────────────
# 6. _CC — Controle Cliente (PCC + PE por empresa/mês)
# ──────────────────────────────────────────────────────────────
print("\n[6/6] _CC (Controle Cliente)...")
# PCC
df_pcc_cc = pd.read_excel(EXCEL, sheet_name='Pedido de carga Completo', engine='openpyxl',
    usecols=['mes_pagamento','CNPJ','nome_fantasia','valor_recarga','projecao_saldo',
             'Carga Realizada','taxa_aproveitamento','Codigo_bilhetagem'])
df_pcc_cc['mes_pagamento'] = df_pcc_cc['mes_pagamento'].astype(str)
df_pcc_cc['CNPJ'] = df_pcc_cc['CNPJ'].astype(str).str.strip()
df_pcc_cc['nome_fantasia'] = df_pcc_cc['nome_fantasia'].astype(str).str.strip()
for c in ['valor_recarga','projecao_saldo','Carga Realizada','taxa_aproveitamento']:
    df_pcc_cc[c] = pd.to_numeric(df_pcc_cc[c], errors='coerce').fillna(0)
pcc_cc = df_pcc_cc.groupby(['mes_pagamento','CNPJ','nome_fantasia']).agg(
    compra=('valor_recarga','sum'), projecao=('projecao_saldo','sum'),
    carga=('Carga Realizada','sum'), tx_aprov=('taxa_aproveitamento','sum'),
    cartoes=('Codigo_bilhetagem','count')).reset_index()
pcc_cc['economia'] = pcc_cc['projecao'] - pcc_cc['tx_aprov']
pcc_cc['origem'] = 0

# PE (coluna E = Mês)
df_pe_cc = df_pe_full.copy()
df_pe_cc['CNPJ'] = df_pe_cc['CNPJ'].astype(str).str.strip()
df_pe_cc['Cliente'] = df_pe_cc['Cliente'].astype(str).str.strip()
pe_cc = df_pe_cc.groupby(['mes_str','CNPJ','Cliente']).agg(
    compra=('Valor Recarga (I)','sum'), projecao=('Projeção de Saldo','sum'),
    carga=('Valor Recarga (F)','sum'), tx_suc=('Tx de Sucesso','sum'),
    cartoes=('CNPJ','count')).reset_index()
pe_cc.rename(columns={'mes_str':'mes_pagamento','Cliente':'nome_fantasia'}, inplace=True)
pe_cc['economia'] = pe_cc['projecao'] - pe_cc['tx_suc']
pe_cc['origem'] = 1

all_cc = pd.concat([
    pcc_cc[['mes_pagamento','CNPJ','nome_fantasia','compra','projecao','carga','economia','cartoes','origem']],
    pe_cc[['mes_pagamento','CNPJ','nome_fantasia','compra','projecao','carga','economia','cartoes','origem']]
], ignore_index=True)
all_cc = all_cc[all_cc['mes_pagamento'] >= '2025_01']

cc_meses_uniq = sorted(all_cc['mes_pagamento'].unique())
cc_meses_labels = [mlabel(m) for m in cc_meses_uniq]
cc_idx_map = {m: i for i, m in enumerate(cc_meses_uniq)}

cc_rows_js = []
for _, row in all_cc.iterrows():
    m = row['mes_pagamento']
    if m not in cc_idx_map: continue
    mi = cc_idx_map[m]
    cnpj = str(row['CNPJ']).strip().replace('"','\\"')
    emp  = str(row['nome_fantasia']).strip().replace('"','\\"')
    cc_rows_js.append(
        f'["{cnpj}","{emp}",{mi},{round(float(row["compra"]),2)},'
        f'{round(float(row["projecao"]),2)},{round(float(row["carga"]),2)},'
        f'{round(float(row["economia"]),2)},{int(row["cartoes"])},{int(row["origem"])}]')
CC_JS = 'var _CC_MESES=' + json.dumps(cc_meses_labels, ensure_ascii=False) + ';\nvar _CC=[\n' + \
        ',\n'.join(cc_rows_js) + '\n];'
print(f"  _CC: {len(cc_rows_js)} entradas, meses: {cc_meses_labels[0]} → {cc_meses_labels[-1]}")

# ──────────────────────────────────────────────────────────────
# 7. Aplicar no HTML (clonar do git, modificar, publicar)
# ──────────────────────────────────────────────────────────────
print("\n[HTML] Clonando repositorio...")
cfg = {}
if os.path.exists(CONFIG_FILE):
    for line in open(CONFIG_FILE, encoding='utf-8-sig'):
        if '=' in line:
            k, v = line.strip().split('=', 1); cfg[k.strip()] = v.strip()
REMOTE_URL = cfg.get('REMOTE_URL','')
if not REMOTE_URL:
    print("ERRO: REMOTE_URL nao encontrado"); sys.exit(1)

subprocess.run(['git','clone', REMOTE_URL, GIT_DIR], capture_output=True)
DASHBOARD = f"{GIT_DIR}/dashboard_motiva_v2.html"
INDEX     = f"{GIT_DIR}/index.html"
if not os.path.exists(DASHBOARD):
    print(f"ERRO: {DASHBOARD} nao encontrado"); sys.exit(1)

with open(DASHBOARD, 'r', encoding='utf-8') as f:
    html = f.read()
orig_kb = len(html)//1024

# Substituir todas as variáveis
html = replace_const(html, 'months',       LABELS,       strings=True)
html = replace_const(html, 'gmvVT',        gmv_vt)
html = replace_const(html, 'gmvPAT',       gmv_pat)
html = replace_const(html, 'gmvTotal',     gmv_total)
html = replace_const(html, 'recVT',        rec_vt)    # = vtTxAdm + vtTxSucesso + vtTxEntrega
html = replace_const(html, 'recPAT',       rec_pat)
html = replace_const(html, 'vtTxAdm',      vt_adm)
html = replace_const(html, 'vtTxSucesso',  vt_sucesso)
html = replace_const(html, 'vtTxEntrega',  vt_entrega)
html = replace_const(html, 'financiamento',financ_rec)
html = replace_const(html, 'nfPedCarga',   nf_pcc)
html = replace_const(html, 'cartoesMes',   cart_mes)  # novo: cartões processados
html = replace_const(html, 'commTotal',    comm_total)
html = replace_const(html, 'commKarine',   comm_karine)
html = replace_const(html, 'commAndre',    comm_andre)
html = replace_const(html, 'commCarol',    comm_carol)
html = replace_const(html, 'commLysandra', comm_lysandra)
html = replace_const(html, 'commVanessa',  comm_vanessa)
html = replace_const(html, 'commPedro',    comm_pedro)
html = replace_const(html, 'qtdClients',   qtd_clients)
html = replace_const(html, 'qtdNew',       qtd_new)

if seg_out:
    def seg_key(k): return '"'+k+'"' if '-' in k else k
    seg_js = 'const segData = {'+','.join(f"{seg_key(k)}:{v}" for k,v in seg_out.items())+'};'
    html, n = re.subn(r'const segData\s*=\s*\{[^}]*\};', seg_js, html)
    print(f"  {'OK' if n else 'XX'} segData: {seg_out}")

html, n = re.subn(r'const _CD=\[.*?\];', CD_JS, html, flags=re.DOTALL)
print(f"  {'OK' if n else 'XX'} _CD: {len(cd_parts)} clientes")

html, n = re.subn(r'const VD=\{.*?\};', vd_js_str, html, flags=re.DOTALL)
print(f"  {'OK' if n else 'XX'} VD: {sum(len(v) for v in vd_dict.values())} entradas")

html, n = re.subn(r'var _CC_MESES=.*?var _CC=\[\n[\s\S]*?\n\];', CC_JS, html)
print(f"  {'OK' if n else 'XX'} _CC: {len(cc_rows_js)} entradas")

last_label = LABELS[-1]
html = re.sub(r'[^ ]* em andamento', f'{last_label} em andamento', html)
badge = 'Dashboard Financeiro \xb7 ' + LABELS[0] + ' – ' + last_label
html = re.sub(r'Dashboard Financeiro.*?–.*?(?=<)', badge, html)

with open(DASHBOARD, 'w', encoding='utf-8') as f: f.write(html)
shutil.copy2(DASHBOARD, INDEX)
shutil.copy2(DASHBOARD, f"{WORKSPACE}/dashboard_motiva_v2.html")
print(f"  HTML: {orig_kb}KB → {len(html)//1024}KB")

# Push GitHub
print("\nPublicando no GitHub...")
data_hoje = datetime.now().strftime('%d/%m/%Y %H:%M')
def git(*args):
    r = subprocess.run(['git','-C',GIT_DIR]+list(args), capture_output=True, text=True)
    if r.returncode != 0 and r.stderr.strip(): print(f"  git {args[0]}: {r.stderr.strip()[:80]}")
    return r

git('config','user.email','vitor.leite@usekim.com.br')
git('config','user.name','Vitor')
git('add','dashboard_motiva_v2.html','index.html')
commit = git('commit','-m', f'Auto-update {data_hoje} - {last_label}')
if 'nothing to commit' in commit.stdout:
    print("  Sem alteracoes.")
else:
    push = git('push','origin','main')
    if push.returncode == 0:
        print(f"  OK → https://vitorabl.github.io/motiva-dashboard/")
    else:
        print(f"  Push falhou: {push.stderr.strip()[:100]}")

# Resumo final
sep = "=" * 50
last_i = m_labels.index(last_complete)
print(sep)
print(f"  Motiva Dashboard — {data_hoje}")
print(f"  Periodo: {LABELS[0]} a {LABELS[-1]}")
print(f"  Último mês completo: {mlabel(last_complete)}")
if last_i >= 0:
    print(f"  GMV VT:    R$ {gmv_vt[last_i]:>12,.0f}")
    print(f"  GMV PAT:   R$ {gmv_pat[last_i]:>12,.0f}")
    print(f"  recVT:     R$ {rec_vt[last_i]:>12,.0f}")
    print(f"  Cartões:   {cart_mes[last_i]:>13,}")
    print(f"  Karine:    R$ {comm_karine[last_i]:>10,.2f}")
print(sep)
