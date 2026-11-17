#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""HiMCM R2 fast 10-city OSM query. Reuses old cache, one Overpass query per city.
Install: py -m pip install requests pyshp shapely pandas openpyxl
Run: py -u .\R2_OSM_10_Cities_FAST.py
"""
import io, json, time, zipfile
from pathlib import Path
from datetime import datetime, timezone
import requests, shapefile, pandas as pd
from shapely.geometry import shape, Point

OUT=Path('R2_OSM_10_Cities_Output'); OUT.mkdir(exist_ok=True)
CACHE=OUT/'cache'; CACHE.mkdir(exist_ok=True)
CITIES=[('Boston','25'),('New York','36'),('Philadelphia','42'),('Cleveland','39'),('Kansas City','29'),('Jacksonville','12'),('Birmingham','01'),('Houston','48'),('Phoenix','04'),('San Diego','06')]
CATS=['pitch','basketball','pool','sports_centre','community_centre']
TAGS={'pitch':'["leisure"="pitch"]','basketball':'["sport"="basketball"]','pool':'["leisure"="swimming_pool"]','sports_centre':'["leisure"="sports_centre"]','community_centre':'["amenity"="community_centre"]'}
SERVERS=['https://overpass.kumi.systems/api/interpreter','https://overpass-api.de/api/interpreter','https://overpass.nchc.org.tw/api/interpreter']
S=requests.Session(); S.headers.update({'User-Agent':'HiMCM-R2-academic-research/1.1'})

def boundary(city,state):
    p=CACHE/f'tl_2025_{state}_place.zip'
    if not p.exists():
        url=f'https://www2.census.gov/geo/tiger/TIGER2025/PLACE/tl_2025_{state}_place.zip'
        print('  Census download:',url,flush=True)
        r=S.get(url,timeout=(10,35));r.raise_for_status();p.write_bytes(r.content)
    with zipfile.ZipFile(p) as z:
        def b(ext):return io.BytesIO(z.read(next(n for n in z.namelist() if n.lower().endswith(ext))))
        sf=shapefile.Reader(shp=b('.shp'),shx=b('.shx'),dbf=b('.dbf'))
        matches=[(r.record.as_dict(),shape(r.shape.__geo_interface__)) for r in sf.iterShapeRecords() if r.record.as_dict().get('NAME')==city]
    if len(matches)!=1:raise ValueError(f'Census {city}/{state}: {len(matches)} matches')
    return matches[0]

def query(b):
    w,s,e,n=b;bb=f'({s:.6f},{w:.6f},{n:.6f},{e:.6f})'
    return '[out:json][timeout:35];('+''.join('nwr'+TAGS[c]+bb+';' for c in CATS)+');out center;'

def download(city,b):
    combined=CACHE/(city.replace(' ','_')+'_ALL_FAST.json')
    if combined.exists():return json.loads(combined.read_text(encoding='utf-8'))
    old=[CACHE/(city.replace(' ','_')+'_'+c+'.json') for c in CATS]
    if all(p.exists() for p in old):
        print('  Reusing 5 existing category caches',flush=True)
        merged={}
        for p in old:
            for el in json.loads(p.read_text(encoding='utf-8'))['elements']:
                merged[(el['type'],el['id'])]=el
        data={'elements':list(merged.values())}
        combined.write_text(json.dumps(data,ensure_ascii=False),encoding='utf-8')
        return data
    q=query(b)
    for attempt in range(2):
        server=SERVERS[attempt]
        print(f'  Overpass request {attempt+1}/2 (max 45 seconds): {server}',flush=True)
        try:
            r=S.post(server,data={'data':q},timeout=(10,45));r.raise_for_status()
            data=r.json()
            if not isinstance(data.get('elements'),list):raise ValueError('No elements array')
            combined.write_text(json.dumps(data,ensure_ascii=False),encoding='utf-8')
            print(f'  Received {len(data["elements"])} OSM objects',flush=True)
            time.sleep(1)
            return data
        except Exception as ex:
            print('  Server error:',repr(ex),flush=True)
    raise RuntimeError('Two attempts failed; city skipped for later retry')

def classify(t):
    if t.get('sport')=='basketball':return 'basketball'
    if t.get('leisure')=='pitch':return 'pitch_other'
    if t.get('leisure')=='swimming_pool':return 'pool'
    if t.get('leisure')=='sports_centre':return 'sports_centre'
    if t.get('amenity')=='community_centre':return 'community_centre'
    return 'other'

def access(t):
    a=str(t.get('access','')).lower()
    if a in ('private','no','customers','permit','members'):return 'restricted'
    if a in ('yes','public','permissive','designated'):return 'explicit_access'
    return 'unknown_access'

def save(rows,sums,errors):
    audit=pd.DataFrame(rows);counts=pd.DataFrame(sums)
    audit.to_csv(OUT/'R2_OSM_10_Cities_Audit.csv',index=False,encoding='utf-8-sig')
    counts.to_csv(OUT/'R2_OSM_10_Cities_Counts.csv',index=False,encoding='utf-8-sig')
    with pd.ExcelWriter(OUT/'R2_OSM_10_Cities.xlsx',engine='openpyxl') as x:
        counts.to_excel(x,'City_Counts',index=False)
        audit.to_excel(x,'OSM_Audit',index=False)
        pd.DataFrame(errors,columns=['city','error']).to_excel(x,'Errors',index=False)
        pd.DataFrame([
            ['Boundary','Census TIGER/Line 2025 PLACE, not MSA'],
            ['Source','OpenStreetMap Overpass API; © OpenStreetMap contributors, ODbL'],
            ['Counting','OSM objects by exclusive class; NOT confirmed public facilities'],
            ['Spatial selection','OSM node or Overpass out center in city polygon; boundary crossing features may be misclassified'],
            ['Access','Unknown access not treated as confirmed public'],
            ['Population','Not joined; use Census PLACE population before per capita comparison'],
            ['Completeness','OSM coverage varies; compare with TPL and run sensitivity analysis']
        ],columns=['item','note']).to_excel(x,'Methodology',index=False)

def main():
    rows=[];sums=[];errors=[];stamp=datetime.now(timezone.utc).isoformat()
    for idx,(city,state) in enumerate(CITIES,1):
        print(f'\n[{idx}/10] {city}',flush=True)
        try:
            props,poly=boundary(city,state)
            data=download(city,poly.bounds)
            unique={(e['type'],e['id']):e for e in data['elements']}
            count={c:0 for c in ('pitch_other','basketball','pool','sports_centre','community_centre')}
            in_count=0;restricted=0;unknown=0
            for (kind,oid),el in unique.items():
                lat=el.get('lat',el.get('center',{}).get('lat'))
                lon=el.get('lon',el.get('center',{}).get('lon'))
                if lat is None or lon is None:continue
                if not poly.covers(Point(lon,lat)):continue
                tags=el.get('tags',{});group=classify(tags);a=access(tags)
                if group not in count:continue
                in_count+=1;count[group]+=1;restricted+=a=='restricted';unknown+=a=='unknown_access'
                rows.append({'city':city,'state_fips':state,'place_geoid':props.get('GEOID',''),
                             'osm_id':f'{kind}/{oid}','latitude':lat,'longitude':lon,
                             'category_exclusive':group,'sport':tags.get('sport',''),
                             'name':tags.get('name',''),'leisure':tags.get('leisure',''),
                             'amenity':tags.get('amenity',''),'access':tags.get('access',''),
                             'access_status':a,'fee':tags.get('fee',''),
                             'operator':tags.get('operator',''),'ownership':tags.get('ownership',''),
                             'tags_json':json.dumps(tags,ensure_ascii=False)})
            sums.append({'city':city,'state_fips':state,'place_geoid':props.get('GEOID',''),
                         'extraction_utc':stamp,'unique_osm_objects_in_place':in_count,
                         'restricted_access_objects':restricted,'unknown_access_objects':unknown,**count})
            print('  SAVED:',in_count,count,flush=True)
        except Exception as ex:
            errors.append((city,str(ex)));print('  SKIPPED:',ex,flush=True)
        save(rows,sums,errors)
        print(f'  Checkpoint written: {len(sums)}/10 successful cities',flush=True)
    print('\nDone. Results:',OUT.resolve(),flush=True)
    if errors:print('Cities to retry:',[e[0] for e in errors],flush=True)

if __name__=='__main__':main()
