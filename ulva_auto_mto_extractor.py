#!/usr/bin/env python
"""
ULVA Auto-MTO Estimator – Custom Full Feature Version

Features:
  - Recognises from PDF:
    • Straight pipes
    • 45° & 90° elbows
    • Equal & Unequal Tees
    • End caps (valve & flange)
    • Collars for small branches (<250 mm)
    • Clamp covers

Calculations:
  • Straight M² = (circumference + 0.05 m lap) × length
  • Circumferential joints at each metre & at fittings: include in bond length
  • Bead length per fitting (heel + throat or full circumference)
  • ULVAShield £ per m²
  • ULVASeal tubes = ceil(total bead length / 6)
  • ULVABond tins = ceil((total bond length × 0.1 m strip) / 2)

Output:
  - Separate Excel sheets: Straights, Elbow45, Elbow90, EqualTee, UnequalTee, EndCap, Collar, ClampCover, Summary
  - Sorted by DN ascending
  - Summary tab with all totals and costs

Usage:
  python ulva_auto_mto_extractor.py [--thickness MM]
  Drop PDFs into pdf_in/ and run. Excel saved to mto_out/.
"""
import sys, re, math, datetime, argparse
from math import pi, ceil
from pathlib import Path
import pandas as pd
import pdfplumber

# Defaults
DEFAULT_THK = 20     # mm insulation thickness
MIN_THK = 5
MAX_THK = 300
LAP = 0.05           # m longitudinal lap (50 mm)
TUBE_COVER_M = 6     # m per ULVASeal tube
BOND_STRIP_W = 0.1   # m (100 mm strip)
DN_COLLAR = 250      # DN threshold for collar

# Rates (£)
RATE_SHIELD = 36.74  # per m²
RATE_SEAL = 12.50    # per tube
RATE_BOND = 9.50     # per m² on strip area
RATE_CLAMP = 24.00   # per clamp cover

# OD lookup (mm)
OD = {15:21.3,20:26.9,25:33.7,32:42.4,40:48.3,50:60.3,65:76.1,
      80:88.9,90:101.6,100:114.3,125:141.3,150:168.3,200:219.1,
      250:273.0,300:323.9,350:355.6,400:406.4,450:457.0,500:508.0,600:610.0}

# CLI
p = argparse.ArgumentParser()
p.add_argument('-t','--thickness',type=int,default=DEFAULT_THK,
               help=f'Insulation thickness (mm) {MIN_THK}-{MAX_THK}')
args = p.parse_args()
if args.thickness < MIN_THK or args.thickness > MAX_THK:
    sys.exit(f'Error: thickness must be {MIN_THK}-{MAX_THK} mm')
INS_THK = args.thickness

# Helpers
def circ_m(dn):
    od = OD.get(dn, dn) + 2*INS_THK
    return pi*od/1000

cut_rx = re.compile(r"<\d+>\s+(\d{2,5})\s+(\d{2,3})")
parse_cuts = lambda txt: [(int(l),int(dn)) for l,dn in cut_rx.findall(txt)]

def parse_fittings(txt):
    items=[]
    for ln in txt.splitlines():
        l=ln.lower()
        if '45' in l and 'elbow' in l:
            dn=int(re.search(r"(\d{2,3})",ln).group(1)); items.append(('Elbow45',dn))
        elif '90' in l and 'elbow' in l:
            dn=int(re.search(r"(\d{2,3})",ln).group(1)); items.append(('Elbow90',dn))
        elif ' tee' in l:
            dns=[int(n) for n in re.findall(r"(\d{2,3})",ln)]
            if len(dns)>=2:
                typ='EqualTee' if dns[0]==dns[1] else 'UnequalTee'
                items.append((typ,dns[0],dns[1]))
        elif 'flange' in l or 'valve' in l:
            dn=int(re.search(r"(\d{2,3})",ln).group(1)); items.append(('EndCap',dn))
        elif ('weldolet' in l or 'threadolet' in l):
            dn=int(re.search(r"(\d{2,3})",ln).group(1))
            if dn < DN_COLLAR: items.append(('Collar',dn))
        elif 'clamp' in l:
            dn=int(re.search(r"(\d{2,3})",ln).group(1)); items.append(('ClampCover',dn))
    return items

# Process PDF
def process_pdf(path):
    txt='\n'.join(p.extract_text() or '' for p in pdfplumber.open(path).pages)
    # Straights
    straights=[]; total_clad=0; bond_len=0; bead_sum=0
    for Lmm,dn in parse_cuts(txt):
        L=ceil(Lmm/1000)
        C=circ_m(dn)
        clad_m2=(C+LAP)*L
        bead=L + 2*C
        straights.append({'PDF':path.name,'DN':dn,'Length_m':L,'Circ_m':round(C,3),
                          'Clad_m2':round(clad_m2,3),'Bead_m':round(bead,3),
                          'Shield_£':round(clad_m2*RATE_SHIELD,2)})
        total_clad+=clad_m2; bond_len+=L + C; bead_sum+=bead
    # Fittings
    fits={k:[] for k in ['Elbow45','Elbow90','EqualTee','UnequalTee','EndCap','Collar','ClampCover']}
    for item in parse_fittings(txt):
        key=item[0]
        if key.startswith('Elbow'):
            angle=int(key[5:]); dn=item[1]
            arc=(angle/360)*2*pi*(OD.get(dn,dn)+2*INS_THK)/1000
            fits[key].append({'PDF':path.name,'DN':dn,'Bead_m':round(arc,3)})
            bead_sum+=arc; bond_len+=arc*2
        elif key in ['EqualTee','UnequalTee']:
            hdr,br=item[1],item[2]
            bead=2*circ_m(hdr)+circ_m(br)
            fits[key].append({'PDF':path.name,'DN_main':hdr,'DN_branch':br,'Bead_m':round(bead,3)})
            bead_sum+=bead; bond_len+=bead
        elif key in ['EndCap','Collar']:
            dn=item[1]; b=circ_m(dn)
            fits[key].append({'PDF':path.name,'DN':dn,'Bead_m':round(b,3)})
            bead_sum+=b; bond_len+=b
        elif key=='ClampCover':
            dn=item[1]; fits[key].append({'PDF':path.name,'DN':dn,'Clamp_£':RATE_CLAMP})
            bond_len+=dn*0  # no bond length
    # Totals
    total_bead=bead_sum
    tubes=ceil(total_bead/TUBE_COVER_M)
    bond_area=(bond_len*BOND_STRIP_W)
    tins=ceil(bond_area/2)
    summary={'Clad_m2':round(total_clad,2),'Bead_m':round(total_bead,2),
             'Tubes':tubes,'Bond_tins':tins}
    return straights,fits,summary

# Main
def main():
    inp,out=Path('pdf_in'),Path('mto_out'); inp.mkdir(exist_ok=True); out.mkdir(exist_ok=True)
    pdfs=list(inp.glob('*.pdf'))+list(inp.glob('*.PDF'))
    if not pdfs:
        print(f'Drop PDFs into {inp}'); sys.exit(1)
    all_str=[]; all_fits={}; all_sum={'Clad_m2':0,'Bead_m':0,'Tubes':0,'Bond_tins':0}
    for pdf in pdfs:
        s,fits,sm=process_pdf(pdf)
        all_str+=s
        for k,v in fits.items(): all_fits.setdefault(k,[]).extend(v)
        for k in all_sum: all_sum[k]+=sm[k]
    ts=datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    out_file=out/f'Auto_MTO_{ts}.xlsx'
    with pd.ExcelWriter(out_file,engine='xlsxwriter') as w:
        pd.DataFrame(all_str).sort_values('DN').to_excel(w,'Straights',index=False)
        for sheet,rows in sorted(all_fits.items()):
            key=sheet
            df=pd.DataFrame(rows)
            sort_col=df.columns[1]
            df.sort_values(sort_col).to_excel(w,key,index=False)
        pd.DataFrame([all_sum]).to_excel(w,'Summary',index=False)
        # Autofit
        for sheet in w.sheets:
            df=pd.read_excel(out_file,sheet_name=sheet)
            ws=w.sheets[sheet]
            for i,col in enumerate(df.columns):
                ws.set_column(i,i,max(df[col].astype(str).map(len).max(),len(col))+2)
    print(f'Excel saved → {out_file}')

if __name__=='__main__':
    main()
