#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""HiMCM R2: uniform OSM extraction for 10 US incorporated places.
Install: py -m pip install requests pyshp shapely pandas openpyxl
Run:     py R2_OSM_10_Cities_Unified.py
Data:    Census TIGER/Line 2025 Places + OSM Overpass, extraction timestamp logged.
"""
import io, json, time, zipfile, csv
from pathlib import Path
from datetime import datetime, timezone
import requests, shapefile, pandas as pd
from shapely.geometry import shape, Point
from shapely.ops import unary_union

OUT = Path('R2_OSM_10_Cities_Output'); OUT.mkdir(exist_ok=True)
CACHE = OUT / 'cache'; CACHE.mkdir(exist_ok=True)
OVERPASS = ['https://overpass.kumi.systems/api/interpreter',
            'https://overpass-api.de/api/interpreter',
            'https://overpass.nchc.org.tw/api/interpreter']
# Census STATEFP + exact NAME (Jacksonville is consolidated city; verify TIGER boundary scope)
CITIES = [
 ('Boston','25','Boston'), ('New York','36','New York'),
 ('Philadelphia','42','Philadelphia'), ('Cleveland','39','Cleveland'),
 ('Kansas City','29','Kansas City'), ('Jacksonville','12','Jacksonville'),
 ('Birmingham','01','Birmingham'), ('Houston','48','Houston'),
 ('Phoenix','04','Phoenix'), ('San Diego','06','San Diego')]
HEADERS = {'User-Agent':'HiMCM-R2-academic-research/1.0 (noncommercial; rate-limited)'}
SESSION = requests.Session(); SESSION.headers.update(HEADERS)


def get_census_boundary(state, name):
    path = CACHE / f'tl_2025_{state}_place.zip'
    if not path.exists():
        url = f'https://www2.census.gov/geo/tiger/TIGER2025/PLACE/tl_2025_{state}_place.zip'
        print('Download Census:', url, flush=True)
        r=SESSION.get(url,timeout=90); r.raise_for_status(); path.write_bytes(r.content)
    with zipfile.ZipFile(path) as z:
        with shapefile.Reader(shp=io.BytesIO(z.read(next(x for x in z.namelist() if x.endswith('.shp')))),
                              shx=io.BytesIO(z.read(next(x for x in z.namelist() if x.endswith('.shx')))),
                              dbf=io.BytesIO(z.read(next(x for x in z.namelist() if x.endswith('.dbf'))))) as sf:
            matches=[]
            for rec in sf.iterShapeRecords():
                if rec.record.as_dict().get('NAME') == name:
                    matches.append((rec.record.as_dict(),shape(rec.shape.__geo_interface__)))
    if len(matches)!=1:
        raise ValueError(f'{name} state={state}: found {len(matches)} boundaries; inspect Census NAME')
    props,geom=matches[0]
    return props,geom


def overpass_query(bbox, category):
    south,west,north,east=bbox
    bb=f'({south:.6f},{west:.6f},{north:.6f},{east:.6f})'
    tags={
      'pitch': '["leisure"="pitch"]',
      'basketball': '["sport"="basketball"]',
      'pool': '["leisure"="swimming_pool"]',
      'sports_centre': '["leisure"="sports_centre"]',
      'community_centre': '["amenity"="community_centre"]',
    }
    return f'[out:json][timeout:240];(nwr{tags[category]}{bb};);out center;'


def fetch(city,cat,bbox):
    path=CACHE / (city.replace(' ','_')+'_'+cat+'.json')
    if path.exists():
        return json.loads(path.read_text(encoding='utf-8'))
    query=overpass_query(bbox,cat)
    last=None
    for attempt in range(9):
        endpoint=OVERPASS[attempt%len(OVERPASS)]
        try:
            response=SESSION.post(endpoint,data={'data':query},timeout=300)
            if response.status_code in (429,502,503,504):
                raise RuntimeError(f'HTTP {response.status_code}')
            response.raise_for_status()
            data=response.json()
            if 'elements' not in data: raise RuntimeError('No elements in response')
            path.write_text(json.dumps(data,ensure_ascii=False),encoding='utf-8')
            time.sleep(2.5)  # rate limit
            return data
        except Exception as e:
            last=e; delay=min(10*(attempt+1),75)
            print(f'  retry {attempt+1}/9: {e}; wait {delay}s',flush=True)
            time.sleep(delay)
    raise RuntimeError(f'{city} {cat} Overpass failed: {last}')


def classify(tags):
    # EXCLUSIVE classification: basketball pitch counted only in basketball
    if tags.get('sport')=='basketball': return 'basketball'
    if tags.get('leisure')=='pitch': return 'pitch_other'
    if tags.get('leisure')=='swimming_pool': return 'pool'
    if tags.get('leisure')=='sports_centre': return 'sports_centre'
    if tags.get('amenity')=='community_centre': return 'community_centre'
    return 'other'


def access_status(tags):
    a=str(tags.get('access','')).lower()
    if a in ('private','no','customers','permit','members'): return 'restricted'
    if a in ('yes','public','permissive','designated'): return 'explicit_access'
    return 'unknown_access'


def main():
    stamp=datetime.now(timezone.utc).isoformat()
    allrows=[]; summaries=[]; errors=[]
    categories=['pitch','basketball','pool','sports_centre','community_centre']
    for city,state,place in CITIES:
        print(f'\n[{city}] Census boundary...',flush=True)
        try:
            props,polygon=get_census_boundary(state,place)
            west,south,east,north=polygon.bounds
            bbox=(south,west,north,east)
            unique={}
            for cat in categories:
                print('  querying',cat,flush=True)
                data=fetch(city,cat,bbox)
                for item in data['elements']:
                    key=f"{item['type']}/{item['id']}"
                    if key not in unique: unique[key]=item
            counts={c:0 for c in ['pitch_other','basketball','pool','sports_centre','community_centre']}
            for key,item in unique.items():
                # Overpass 'center' for ways/relations is bounding-box center, NOT guaranteed inside polygon.
                lat=item.get('lat',item.get('center',{}).get('lat'))
                lon=item.get('lon',item.get('center',{}).get('lon'))
                if lat is None or lon is None: continue
                inside=polygon.covers(Point(lon,lat))
                tags=item.get('tags',{})
                group=classify(tags); access=access_status(tags)
                if inside and group in counts: counts[group]+=1
                allrows.append({
                    'city':city,'state_fips':state,'place_geoid':props.get('GEOID',''),
                    'osm_id':key,'latitude':lat,'longitude':lon,'within_place_by_center':inside,
                    'category_exclusive':group,'sport':tags.get('sport',''),
                    'name':tags.get('name',''),'leisure':tags.get('leisure',''),
                    'amenity':tags.get('amenity',''),'access':tags.get('access',''),
                    'access_status':access,'fee':tags.get('fee',''),
                    'operator':tags.get('operator',''),'ownership':tags.get('ownership',''),
                    'indoor':tags.get('indoor',''),'source_query':'bbox + center point-in-polygon',
                    'tags_json':json.dumps(tags,ensure_ascii=False)})
            cityrows=[r for r in allrows if r['city']==city and r['within_place_by_center']]
            row={'city':city,'state_fips':state,'place_geoid':props.get('GEOID',''),
                 'extraction_utc':stamp,'boundary':'Census TIGER/Line 2025 PLACE',
                 'unique_osm_objects_in_place':len(cityrows),
                 'restricted_access_objects':sum(r['access_status']=='restricted' for r in cityrows),
                 'unknown_access_objects':sum(r['access_status']=='unknown_access' for r in cityrows),
                 **counts}
            summaries.append(row)
            print('  within city:',len(cityrows),counts,flush=True)
        except Exception as e:
            errors.append({'city':city,'error':str(e)})
            print('  ERROR:',e,flush=True)
    df=pd.DataFrame(allrows); summary=pd.DataFrame(summaries)
    df.to_csv(OUT/'R2_OSM_10_Cities_Audit.csv',index=False,encoding='utf-8-sig')
    summary.to_csv(OUT/'R2_OSM_10_Cities_Counts.csv',index=False,encoding='utf-8-sig')
    with pd.ExcelWriter(OUT/'R2_OSM_10_Cities.xlsx',engine='openpyxl') as writer:
        summary.to_excel(writer,sheet_name='City_Counts',index=False)
        df.to_excel(writer,sheet_name='OSM_Audit',index=False)
        pd.DataFrame([
          ['Scope','Census TIGER/Line 2025 PLACE city boundaries; NOT MSA'],
          ['OSM extraction',stamp],
          ['OSM data license','ODbL; attribution: © OpenStreetMap contributors'],
          ['Count unit','OSM object (node/way/relation), not validated physical facility'],
          ['Geometry caveat','Ways/relations tested using Overpass out center (bbox center); may misclassify boundary-crossing features'],
          ['Basketball','sport=basketball assigned exclusively to basketball; not also pitch'],
          ['Pool','May include private/residential pools; access_status must be audited'],
          ['Public access','unknown_access does NOT mean public'],
          ['Population','Not supplied; join consistent Census place population before per-capita rates'],
          ['Comparison','TPL city-level data are external validation, not directly interchangeable'],
          ['Completeness','OSM completeness varies by city; sensitivity tests required'],
        ],columns=['item','note']).to_excel(writer,sheet_name='Methodology',index=False)
        pd.DataFrame(errors).to_excel(writer,sheet_name='Errors',index=False)
    print('\nFinished. Output:',OUT.resolve())
    if errors: print('WARNING failed cities:',errors)

if __name__=='__main__': main()
