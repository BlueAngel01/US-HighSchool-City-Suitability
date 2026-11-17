# -*- coding: utf-8 -*-
"""Download official 21st CCLC awardee/site source files (no invented counts).
Run: py -m pip install requests pandas openpyxl beautifulsoup4 lxml
     py R1_21CCLC_official_download.py
"""
import re, json, time
from pathlib import Path
from urllib.parse import urljoin
import requests
import pandas as pd
from bs4 import BeautifulSoup

OUT=Path('R1_21CCLC_official'); OUT.mkdir(exist_ok=True)
S=requests.Session(); S.headers.update({'User-Agent':'Mozilla/5.0 (compatible; ResearchDataDownloader/1.0)'})
SOURCES=[
 ('NY','https://www.nysed.gov/student-support-services/21stcclc-award-recipients', ['round 8','21st','participant','award']),
 ('CA','https://www.cde.ca.gov/fg/fo/r27/cclcemassetsresults24.asp',['2024','2025','assets','funding']),
]
log=[]
def record(region,kind,url,status,filename='',note=''):
    log.append(dict(region=region,type=kind,url=url,status=status,file=filename,note=note))

def save_binary(region,url,kind):
    try:
        r=S.get(url,timeout=65); r.raise_for_status()
        if len(r.content)<500: raise ValueError('File unexpectedly small')
        if b'<html' in r.content[:200].lower(): raise ValueError('HTML returned instead of binary file')
        ext='.xlsx' if '.xlsx' in url.lower() else '.xls' if '.xls' in url.lower() else '.pdf' if '.pdf' in url.lower() else '.bin'
        name=f'{region}_{len(log)+1:03d}{ext}'
        (OUT/name).write_bytes(r.content)
        record(region,kind,url,'downloaded',name)
    except Exception as e: record(region,kind,url,'failed',note=str(e))

def scrape_official(region,page,terms):
    try:
        r=S.get(page,timeout=45); r.raise_for_status()
        soup=BeautifulSoup(r.text,'html.parser')
        links=[]
        for a in soup.find_all('a',href=True):
            u=urljoin(page,a['href']); label=a.get_text(' ',strip=True)
            if (u.lower().split('?')[0].endswith(('.xlsx','.xls','.csv','.pdf')) and
                (any(t in (label+' '+u).lower() for t in terms))):
                links.append((u,label))
        links=list(dict((u,(u,label)) for u,label in links).values())
        if not links: record(region,'page',page,'no_matching_downloads',note='Open page manually; may be blocked or use JavaScript')
        for url,label in links:
            save_binary(region,url,'award_list: '+label[:90]); time.sleep(.3)
    except Exception as e: record(region,'page',page,'failed',note=str(e))

def arcgis_wisconsin():
    # Verified publicly listed 2024-25 awardees FeatureServer layer 10.
    root='https://services8.arcgis.com/o4NJgD3NfeHnWy06/ArcGIS/rest/services/CCLC_Awardees_2024_25/FeatureServer/10'
    try:
        info=S.get(root,params={'f':'json'},timeout=45);info.raise_for_status(); meta=info.json()
        if 'error' in meta: raise ValueError(str(meta['error']))
        oid=meta.get('objectIdField') or next((f['name'] for f in meta.get('fields',[]) if f.get('type')=='esriFieldTypeOID'),'OBJECTID')
        q=S.get(root+'/query',params={'where':'1=1','outFields':'*','returnGeometry':'false','f':'json','returnIdsOnly':'true'},timeout=60)
        q.raise_for_status();ids=q.json().get('objectIds',[])
        if not ids: raise ValueError('No object IDs returned')
        features=[]
        for start in range(0,len(ids),200):
            batch=ids[start:start+200]
            rr=S.post(root+'/query',data={'objectIds':','.join(map(str,batch)),'outFields':'*','returnGeometry':'false','f':'json'},timeout=90)
            rr.raise_for_status(); data=rr.json()
            if 'error' in data: raise ValueError(str(data['error']))
            features += [f.get('attributes',{}) for f in data.get('features',[])]
        df=pd.DataFrame(features).drop_duplicates(subset=[oid] if oid in pd.DataFrame(features).columns else None)
        fn='WI_21CCLC_2024_25_awardee_sites.xlsx';df.to_excel(OUT/fn,index=False)
        record('WI','ArcGIS site records',root,'downloaded',fn,f'{len(df)} records; awardee sites not activity programs')
    except Exception as e:record('WI','ArcGIS site records',root,'failed',note=str(e))

if __name__=='__main__':
    for region,page,terms in SOURCES: scrape_official(region,page,terms)
    arcgis_wisconsin()
    pd.DataFrame(log).to_excel(OUT/'DOWNLOAD_LOG.xlsx',index=False)
    print(pd.DataFrame(log).to_string(index=False,max_colwidth=70))
    print('Files saved to:',OUT.resolve())
    print('CAUTION: Awardees/sites are NOT unique STEM/arts/sports/volunteer programs; do not calculate final R1 from them directly.')
