import argparse
from pathlib import Path
import pandas as pd

MISSING = {'', '*', 'PS', 'NA', 'N/A', '-', 'NULL', 'SUPPRESSED'}

def load(path):
    path = Path(path)
    if path.suffix.lower() in ('.xlsx', '.xls'):
        return pd.read_excel(path, dtype=str)
    for encoding in ('utf-8-sig', 'utf-16', 'cp1252', 'latin1'):
        try:
            return pd.read_csv(path, dtype=str, encoding=encoding, sep=None, engine='python', on_bad_lines='warn')
        except (UnicodeError, pd.errors.ParserError):
            continue
    raise ValueError('Cannot read input file')

def num(series):
    return pd.to_numeric(series.astype('string').str.strip().str.replace(',', '', regex=False).str.replace('%', '', regex=False), errors='coerce')

def main():
    p = argparse.ArgumentParser(description='E1 2023-24 public high school ACGR extraction (no fabricated values)')
    p.add_argument('--input', required=True, help='Official ACGR CSV/TXT/XLSX')
    p.add_argument('--output', default='E1_ACGR_2024.xlsx')
    p.add_argument('--inspect', action='store_true', help='Only show columns and sample records')
    p.add_argument('--mapping', help='CSV/XLSX with school_id,msa; use with --id-col')
    p.add_argument('--id-col', help='School identifier column in official data')
    p.add_argument('--graduates-col', help='Graduates count column')
    p.add_argument('--cohort-col', help='Adjusted cohort count column')
    p.add_argument('--rate-col', help='Optional published ACGR rate column')
    p.add_argument('--year-col', help='Optional school year column')
    p.add_argument('--year-value', default='2023-24')
    p.add_argument('--subgroup-col', help='Optional subgroup column')
    p.add_argument('--subgroup-value', help='ALL')
    p.add_argument('--level-col', help='Optional school/district/state level column')
    p.add_argument('--level-value', help='School')
    p.add_argument('--msa-col', help='MSA column already present in source')
    args = p.parse_args()
    df = load(args.input)
    print('Rows:', len(df), '\nColumns:')
    for col in df.columns: print(' ', repr(col))
    print('\nFirst 3 rows:\n', df.head(3).to_string(index=False, max_cols=20))
    if args.inspect:
        print('\nInspect only: no statistics computed.')
        return
    required = {'graduates-col': args.graduates_col, 'cohort-col': args.cohort_col}
    missing = [key for key, value in required.items() if not value]
    if missing: p.error('Specify exact source columns: ' + ', '.join('--' + m for m in missing))
    for name in [args.graduates_col, args.cohort_col, args.rate_col, args.year_col, args.subgroup_col, args.level_col, args.msa_col, args.id_col]:
        if name and name not in df.columns: p.error(f'Column not found: {name}')
    if args.year_col:
        df = df[df[args.year_col].astype(str).str.strip() == args.year_value].copy()
    if args.subgroup_col:
        df = df[df[args.subgroup_col].astype(str).str.strip() == args.subgroup_value].copy()
    if args.level_col:
        df = df[df[args.level_col].astype(str).str.strip() == args.level_value].copy()
    if not args.subgroup_col or not args.level_col:
        print('WARNING: Verify source has ONLY all-students school-level rows; otherwise double counting is possible.')
    df['graduates_n'] = num(df[args.graduates_col])
    df['cohort_n'] = num(df[args.cohort_col])
    df['valid'] = (df['cohort_n'] > 0) & (df['graduates_n'] >= 0) & (df['graduates_n'] <= df['cohort_n'])
    df['calculated_acgr_pct'] = (100 * df['graduates_n'] / df['cohort_n']).where(df['valid'])
    if args.rate_col:
        df['published_acgr_pct'] = num(df[args.rate_col])
        df['rate_difference_pp'] = df['calculated_acgr_pct'] - df['published_acgr_pct']
    if args.mapping:
        if not args.id_col: p.error('--id-col required with --mapping')
        mapping = load(args.mapping)
        if not {'school_id', 'msa'}.issubset(mapping.columns): p.error('Mapping must contain school_id, msa')
        mapping = mapping[['school_id','msa']].copy()
        mapping['school_id'] = mapping['school_id'].astype(str).str.strip()
        if mapping['school_id'].duplicated().any(): p.error('Duplicate school_id in mapping; fix before aggregation')
        df['_join_id'] = df[args.id_col].astype(str).str.strip()
        df = df.merge(mapping, left_on='_join_id', right_on='school_id', how='left', validate='many_to_one')
        msa_col = 'msa'
    elif args.msa_col:
        msa_col = args.msa_col
    else:
        msa_col = None
    summary = pd.DataFrame()
    if msa_col:
        records = []
        for msa, g in df[df[msa_col].notna() & (df[msa_col].astype(str).str.strip() != '')].groupby(msa_col):
            good = g[g['valid']]
            cohort = good['cohort_n'].sum()
            grads = good['graduates_n'].sum()
            records.append({'MSA':msa, 'source_rows':len(g), 'valid_rows':len(good), 'invalid_or_suppressed_rows':len(g)-len(good), 'adjusted_cohort_known':cohort, 'graduates_known':grads, 'E1_ACGR_pct_known_rows_only':100*grads/cohort if cohort>0 else None, 'complete_rows':len(g)==len(good), 'interpretation':'PARTIAL - suppressed/missing' if len(g)!=len(good) else 'Complete among mapped rows'})
        summary = pd.DataFrame(records)
        if not summary.empty: summary = summary.sort_values('MSA')
    notes = pd.DataFrame({'note':[
        'Source: ' + str(Path(args.input).resolve()),
        'School year requested: ' + args.year_value,
        'Use only ALL STUDENTS + FOUR-YEAR ACGR + SCHOOL level; check source codebook manually.',
        'Never average school graduation percentages. E1 = sum(graduates)/sum(adjusted cohort)*100.',
        'If suppressed/invalid rows exist, calculated result is PARTIAL, not an official MSA-wide ACGR.',
        'Mapping must include every relevant public high school in each MSA; unmatched schools are excluded.',
        'MSA boundaries and school geography must be validated; do not assume district name equals MSA.',
        'If published rates are rounded or count definitions differ, differences require review.',
    ]})
    out = Path(args.output)
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='School_records', index=False)
        summary.to_excel(writer, sheet_name='MSA_summary', index=False)
        notes.to_excel(writer, sheet_name='Method_notes', index=False)
    print('\nSaved:', out.resolve())
    print('Valid rows:', int(df['valid'].sum()), '/', len(df))
    if msa_col: print('Unmapped rows:', int(df[msa_col].isna().sum()))

if __name__ == '__main__':
    main()
