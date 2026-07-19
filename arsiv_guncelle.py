#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ARŞİV REFERANSINI KUR/GÜNCELLE
Geçmiş beyanname listesinden (Ozel Rapor Kalem Bazinda.xlsx) eşleştirme motorunu kurar.
Arşiv güncellenince tekrar çalıştır.
KULLANIM: python3 arsiv_guncelle.py --arsiv "veri/Ozel Rapor Kalem Bazinda.xlsx"
"""
import pandas as pd, re, pickle, os, argparse, warnings
warnings.filterwarnings('ignore')
BASE=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
def norm(s):
    if pd.isna(s): return None
    s=re.sub(r'\s+',' ',str(s).upper().strip());return s or None
def main():
    ap=argparse.ArgumentParser();ap.add_argument('--arsiv',required=True);a=ap.parse_args()
    print("Arşiv okunuyor...")
    df=pd.read_excel(a.arsiv,engine='openpyxl')
    print(f"  {len(df):,} kalem")
    d=pd.DataFrame({'gtip':df['Gtip'].astype(str).str.strip().replace({'nan':None,'':None}),
        'tanim':df['Ticari Tanım'].map(norm),'gonderen':df['Gönderen'].map(norm) if 'Gönderen' in df else None,
        'mense_kod':df['Menşe Kodu'].astype(str).str.strip().replace({'nan':None,'':None}) if 'Menşe Kodu' in df else None,
        'kdv':df['KDV Oran'].astype(str).str.strip().replace({'nan':None,'':None}) if 'KDV Oran' in df else None,
        'olcu':df['Ölçü Birimi'].astype(str).str.strip().replace({'nan':None,'':None}) if 'Ölçü Birimi' in df else None,
        }).dropna(subset=['gtip','tanim'])
    def build(keys):
        out={}
        for kv,grp in d.dropna(subset=keys).groupby(keys):
            vc=grp['gtip'].value_counts();top=vc.index[0]
            k=kv if isinstance(kv,tuple) else (kv,)
            sub=grp[grp['gtip']==top]
            out[k]={'gtip':top,'confidence':round(vc.iloc[0]/vc.sum(),2),'n':int(vc.sum()),
                'mense_kod':sub['mense_kod'].mode().iloc[0] if 'mense_kod' in sub and not sub['mense_kod'].mode().empty else None,
                'kdv':sub['kdv'].mode().iloc[0] if 'kdv' in sub and not sub['kdv'].mode().empty else None,
                'olcu':sub['olcu'].mode().iloc[0] if 'olcu' in sub and not sub['olcu'].mode().empty else None}
        return out
    words={}
    for t in d['tanim'].dropna().unique():
        for w in t.split():
            if len(w)>=3: words.setdefault(w,set()).add(t)
    ref={'lut_combo':build(['gonderen','tanim']) if d['gonderen'].notna().any() else {},
         'lut_tanim':build(['tanim']),'tanim_words':words,
         'stats':{'kalem':len(d),'gtip':d['gtip'].nunique(),'tanim':d['tanim'].nunique()}}
    os.makedirs(os.path.join(BASE,'veri'),exist_ok=True)
    pickle.dump(ref,open(os.path.join(BASE,'veri','reference.pkl'),'wb'))
    print(f"  reference.pkl kuruldu: {len(ref['lut_tanim']):,} tanım anahtarı, {len(words):,} kelime")
if __name__=='__main__':main()
