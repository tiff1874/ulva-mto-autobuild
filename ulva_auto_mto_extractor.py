import sys, re, math, datetime
from math import pi, ceil
from pathlib import Path
import pandas as pd
try:
    import pdfplumber
except ImportError:
    sys.exit("pdfplumber missing – pip install pdfplumber")
try:
    import xlsxwriter
except ImportError:
    sys.exit("xlsxwriter missing – pip install xlsxwriter")

INS_THK = 20           # default mm insulation
TUBE_LEN = 6           # metres of bead per ULVASeal tube
BOND_STRIP = 0.10      # 100 mm bond width
ULVASHIELD_RATE = 36.74  # £/m²
ULVASEAL_RATE = 12.50    # £/tube
ULVABOND_RATE = 9.50     # £/litre
CLAMP_COVER_RATE = 24.00 # £ per clamp cover (example rate)

OD = {15:21.3,20:26.9,25:33.7,32:42.4,40:48.3,50:60.3,65:76.1,80:88.9,90:101.6,
      100:114.3,125:141.3,150:168.3,200:219.1,250:273.0,300:323.9,350:355.6,
      400:406.4,450:457.0,500:508.0,600:610.0}

def ins_od(dn:int, thk:int)->float: return OD.get(dn,dn)+2*thk

def circ_m(dn:int, thk:int)->float: return pi*ins_od(dn, thk)/1000

def parse_cuts(txt:str):
    return [(int(l),int(dn)) for l,dn in re.findall(r"<\d+>\s+(\d{2,5})\s+(\d{2,3})",txt)]

def parse_fittings(txt:str):
    out=[]
    for ln in txt.splitlines():
        l=ln.lower()
        if " elbow" in l:
            ang = 90 if "90" in l else 45 if "45" in l else 90
            m=re.search(r"(\d{2,3})",ln)
            if m: out.append({"elbow":ang,"dn":int(m.group(1))})
        elif " tee" in l:
            nums=[int(n) for n in re.findall(r"(\d{2,3})",ln)]
            if nums: out.append({"tee":nums})
        elif "reducer" in l:
            nums=[int(n) for n in re.findall(r"(\d{2,3})",ln)]
            if len(nums)>=2: out.append({"reducer":nums})
        elif "flange" in l or "valve" in l:
            m=re.search(r"(\d{2,3})",ln)
            if m: out.append({"cap":int(m.group(1))})
        elif "weldolet" in l or "threadolet" in l:
            m=re.search(r"(\d{2,3})",ln)
            if m: out.append({"collar":int(m.group(1))})
        elif "clamp" in l:
            m=re.search(r"(\d{2,3})",ln)
            if m: out.append({"clamp":int(m.group(1))})
    return out

def process_pdf(path:Path, thk:int):
    with pdfplumber.open(path) as doc:
        txt="\n".join(p.extract_text() or "" for p in doc.pages)
    straights=[]; elbows=[]; tees=[]; reducers=[]; caps=[]; collars=[]; clamps=[]; bead=0; area_total=0
    for l,dn in parse_cuts(txt):
        m=math.ceil(l/1000); circ=circ_m(dn, thk); area=circ*m; b=m+2*circ
        straights.append({"PDF":path.name,"DN":dn,"Rounded_m":m,"Circ_m":round(circ,3),"M2":round(area,3),"Bead_m":round(b,3),"ULVAShield_£":round(area*ULVASHIELD_RATE,2)})
        bead+=b; area_total+=area
    for f in parse_fittings(txt):
        if "elbow" in f:
            arc=(f["elbow"]/360)*pi*ins_od(f["dn"], thk)/1000; b=arc*2; bead+=b
            elbows.append({"PDF":path.name,"Component":f"Elbow {f['elbow']}","DN":f["dn"],"Bead_m":round(b,3)})
        elif "tee" in f:
            hdr=f["tee"][0]; br=f["tee"][1] if len(f["tee"])>1 else hdr
            b=2*circ_m(hdr, thk)+circ_m(br, thk); bead+=b
            tees.append({"PDF":path.name,"Component":"Tee","DN":hdr,"Branch_DN":br,"Bead_m":round(b,3)})
        elif "reducer" in f:
            big,small=f["reducer"][:2]; b=circ_m(big, thk)+circ_m(small, thk); bead+=b
            reducers.append({"PDF":path.name,"Component":"Reducer","DN_big":big,"DN_small":small,"Bead_m":round(b,3)})
        elif "cap" in f:
            b=circ_m(f["cap"], thk); bead+=b
            caps.append({"PDF":path.name,"DN":f["cap"],"Bead_m":round(b,3)})
        elif "collar" in f:
            b=circ_m(f["collar"], thk); bead+=b
            collars.append({"PDF":path.name,"DN":f["collar"],"Bead_m":round(b,3)})
        elif "clamp" in f:
            clamps.append({"PDF":path.name,"DN":f["clamp"],"Clamp_£":CLAMP_COVER_RATE})
    return straights, elbows, tees, reducers, caps, collars, clamps, bead, area_total

def collect_inputs(args):
    if len(args)==1:
        return Path("pdf_in"), Path("mto_out"), INS_THK
    thk = int(args[1]) if len(args)>1 and args[1].isdigit() else INS_THK
    return Path("pdf_in"), Path("mto_out"), thk

def main(argv):
    inputs, out_dir, thk = collect_inputs(argv)
    inputs.mkdir(exist_ok=True)
    out_dir.mkdir(exist_ok=True)
    pdfs=list(inputs.glob("*.pdf"))+list(inputs.glob("*.PDF"))
    if not pdfs: sys.exit(f"Drop PDFs into {inputs} and run again.")

    st=elb=tee=red=cap=col=cl=[]; total_bead=0; total_m2=0
    for pdf in pdfs:
        s,e,t,r,c,k,m,b,a=process_pdf(pdf, thk)
        st+=s; elb+=e; tee+=t; red+=r; cap+=c; col+=k; cl+=m
        total_bead+=b; total_m2+=a

    seal_tubes=ceil(total_bead/TUBE_LEN)
    bond_area=total_bead*BOND_STRIP
    bond_tins=ceil(bond_area)
    df_s=pd.DataFrame(st)
    df_el=pd.DataFrame(elb)
    df_t=pd.DataFrame(tee)
    df_r=pd.DataFrame(red)
    df_c=pd.DataFrame(cap)
    df_k=pd.DataFrame(col)
    df_cl=pd.DataFrame(cl)
    df_seal=pd.DataFrame({"Total_Bead_m":[round(total_bead,2)],"ULVASeal_Tubes":[seal_tubes],"ULVASeal_£":[round(seal_tubes*ULVASEAL_RATE,2)]})
    df_bond=pd.DataFrame({"Bond_Area_m2":[round(bond_area,2)],"ULVABond_Tins":[bond_tins],"ULVABond_£":[round(bond_tins*ULVABOND_RATE,2)]})
    df_total=pd.DataFrame({"Straight_M2_Total":[round(total_m2,2)],"ULVAShield_£":[round(total_m2*ULVASHIELD_RATE,2)]})

    ts=datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    out=out_dir/f"Auto_MTO_{ts}.xlsx"
    with pd.ExcelWriter(out,engine="xlsxwriter") as w:
        for name,df in [
            ("Straights",df_s),("Elbows",df_el),("Tees",df_t),
            ("Reducers",df_r),("EndCaps",df_c),("Collars",df_k),
            ("ClampCovers",df_cl),("ULVASeal",df_seal),("ULVABond",df_bond),("Totals",df_total)
        ]:
            df.to_excel(w,name,index=False)
            ws=w.sheets[name]
            for i,col in enumerate(df.columns):
                width = max(len(str(col)), max(df[col].astype(str).map(len).max(), default=0)) + 2
                ws.set_column(i,i,width)
    print(f"✅ Excel → {out}")

if __name__=="__main__":
    main(sys.argv)
