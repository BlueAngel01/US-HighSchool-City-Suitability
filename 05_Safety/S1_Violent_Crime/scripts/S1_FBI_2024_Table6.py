"""FBI 2024 Table 6 extraction with audit trail. Does not invent MSA rates."""
import argparse, re, pathlib, requests, pdfplumber, pandas as pd
URL='https://www.justfacts.com/document/crime_united_states_2024_fbi_full.pdf'
TARGETS=['New York','Boston','Pittsburgh','Chicago','Minneapolis','Omaha','Charlotte','Birmingham','Los Angeles','San Jose','Seattle','Denver','Naperville','Plano']

def main():
 p=argparse.ArgumentParser()
 p.add_argument('--pdf',help='Local FBI 2024 report PDF (preferred: downloaded from FBI CDE)')
 p.add_argument('--download-mirror',action='store_true',help='Download publicly hosted copy of FBI report from third-party mirror')
 p.add_argument('--out',default='S1_FBI_2024')
 args=p.parse_args()
 out=pathlib.Path(args.out);out.mkdir(exist_ok=True,parents=True)
 if args.pdf: path=pathlib.Path(args.pdf)
 elif args.download_mirror:
  path=out/'FBI_CIUS_2024_mirror.pdf'
  if not path.exists():
   r=requests.get(URL,timeout=120,headers={'User-Agent':'Mozilla/5.0'});r.raise_for_status()
   if not r.content.startswith(b'%PDF'):raise RuntimeError('Downloaded content is not a PDF')
   path.write_bytes(r.content)
 else: p.error('Supply --pdf PATH for official download, or --download-mirror')
 rows=[]; active=False; found=False
 with pdfplumber.open(path) as pdf:
  for idx,page in enumerate(pdf.pages,1):
   txt=page.extract_text(layout=False) or ''
   if re.search(r'Table\s*6\s*[-–:]\s*Crime in the United States|Table\s*6\s*[-–:]\s*Crime',txt,re.I):active=True;found=True
   if active and re.search(r'Table\s*7\s*[-–:]|Table\s*7\s*Offense',txt,re.I):break
   if active:
    for line in txt.splitlines():
     rows.append({'pdf_page':idx,'raw_line':line,'target_matches':'; '.join(t for t in TARGETS if t.lower() in line.lower())})
 if not found:
  print('Table 6 heading not detected; extracting pages containing MSA names for manual review.')
  with pdfplumber.open(path) as pdf:
   for idx,page in enumerate(pdf.pages,1):
    for line in (page.extract_text() or '').splitlines():
     if any(t.lower() in line.lower() for t in TARGETS):rows.append({'pdf_page':idx,'raw_line':line,'target_matches':'; '.join(t for t in TARGETS if t.lower() in line.lower())})
 df=pd.DataFrame(rows,columns=['pdf_page','raw_line','target_matches'])
 matches=df[df.target_matches.ne('')].copy()
 df.to_csv(out/'S1_Table6_raw_lines.csv',index=False,encoding='utf-8-sig')
 matches.to_csv(out/'S1_MSA_candidate_lines.csv',index=False,encoding='utf-8-sig')
 with pd.ExcelWriter(out/'S1_Table6_audit.xlsx',engine='openpyxl') as w:
  matches.to_excel(w,sheet_name='MSA_candidate_lines',index=False)
  df.to_excel(w,'Table6_raw_lines',index=False)
  pd.DataFrame({'target_city':TARGETS,'MSA_name_confirmed':['']*len(TARGETS),'violent_crimes_2024':['']*len(TARGETS),'population_covered':['']*len(TARGETS),'rate_per_100k':['']*len(TARGETS),'coverage_note':['']*len(TARGETS)}).to_excel(w,'Review_template',index=False)
 print('Done:',out.resolve());print('Candidate lines:',len(matches));print('IMPORTANT: candidate lines are NOT verified MSA totals. Review Table 6 headers, footnotes, and MSA boundaries before calculating S1.')
if __name__=='__main__':main()
