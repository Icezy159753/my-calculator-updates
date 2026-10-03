import importlib.util
import io
import contextlib
import json
from pathlib import Path
from types import SimpleNamespace

import numpy as np
import pandas as pd
import pyreadstat
import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('brandsense', ROOT / 'All_Programs/123_Program_Run_Brandsence2026.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
cls = module.SpssProcessorApp
proxy = SimpleNamespace(_qc_use_model_safeguards=False,
    _rotated_factor_loadings=cls._rotated_factor_loadings,
    _factor_mapping_details=cls._factor_mapping_details)
df, meta = pyreadstat.read_sav(str(next((ROOT / 'data').glob('*Compute C.sav'))))
qc = pd.read_excel(next((ROOT / 'data').glob('*BS Output.xlsx')), sheet_name='QC Excluded')
excluded = set(zip(qc.SBJNUM, qc.Index1))
clean = df.loc[[key not in excluded for key in zip(df.SBJNUM, df.Index1)]]
report = {'input_rows':len(df),'after_qc_rows':len(clean),'groups':{}}
for label, group in [('users81', clean[(clean.Index1==81)&(clean.S13_AUTO_1==1)]),
                     ('all81',clean[clean.Index1==81]),
                     ('competitor81',clean[(clean.Index1==81)&(clean.S13_AUTO_1==2)])]:
    with contextlib.redirect_stdout(io.StringIO()):
        scores, loadings, mapping, md = cls.perform_factor_analysis(proxy,group)
        _, betas, _, diag = cls.perform_regression_analysis(proxy,group.join(scores),mapping)
    cols=['N_S','N_P','N_C','N_E','ZA']
    basis=betas.Beta.abs()
    direct=sm.OLS(group.ZA,sm.add_constant(group[cols[:4]])).fit()
    report['groups'][label]={'n':len(group),'statistics':group[cols].describe().to_dict(),
        'value_counts':{c:{str(k):int(v) for k,v in group[c].value_counts().items()} for c in cols},
        'correlations':group[cols].corr().to_dict(),'loadings':loadings.to_dict(),
        'mapping':mapping,'mapping_diagnostics':md,'beta':betas.to_dict(),
        'B_ratios':(basis/basis.sum()*100).to_dict(),'diagnostics':diag,
        'direct_standardized_beta':(direct.params.drop('const')*group[cols[:4]].std()/group.ZA.std()).to_dict(),
        'direct_p':direct.pvalues.to_dict(),
        'vif':{c:float(variance_inflation_factor(sm.add_constant(group[cols[:4]]).values,i+1)) for i,c in enumerate(cols[:4])}}
    if label=='users81':
        print('USERS81 FACTOR LOADINGS\n',loadings.to_string())
        print('BETAS\n',betas.to_string())
        print('VALUE COUNTS',report['groups'][label]['value_counts'])
        print('CORRELATIONS\n',group[cols].corr().to_string())
    print(label, 'RATIOS',report['groups'][label]['B_ratios'],'diagnostics',diag,'mapping',mapping)
raw, _ = pyreadstat.read_sav(str(next(p for p in (ROOT/'data').glob('*.sav') if 'Compute' not in p.name)))
settings=pd.read_excel(next((ROOT/'data').glob('8_Setting*')),sheet_name='Settings')
evars=[v for v in settings.E.dropna() if str(v).endswith('$81')]
raw_e=raw.set_index('SBJNUM')[evars].mean(axis=1)
users=clean[(clean.Index1==81)&(clean.S13_AUTO_1==1)]
report['raw_E_check']={'variables':evars,'max_abs_difference':float(np.max(np.abs(users.N_E.to_numpy()-users.SBJNUM.map(raw_e).to_numpy())))}
print('RAW E CHECK',report['raw_E_check'])
ratios=[]
invalid=[]
for idx in users.index:
    sample=users.drop(index=idx)
    constant=[c for c in ['N_S','N_P','N_C','N_E'] if sample[c].nunique()==1]
    if constant:
        invalid.append({'removed_SBJNUM':float(users.loc[idx,'SBJNUM']),'constant_variables':constant})
        continue
    with contextlib.redirect_stdout(io.StringIO()):
        scores, _, mapping, _ = cls.perform_factor_analysis(proxy,sample)
        _, b, _, d = cls.perform_regression_analysis(proxy,sample.join(scores),mapping)
    basis=b.Beta.abs()
    ratios.append({'removed_SBJNUM':float(users.loc[idx,'SBJNUM']),'B_E':float(basis.N_E/basis.sum()*100),'r_squared':d['r_squared']})
report['leave_one_out']=ratios
report['leave_one_out_invalid']=invalid
print('LEAVE ONE OUT B.E MIN MAX',min(r['B_E'] for r in ratios),max(r['B_E'] for r in ratios))
print('LEAVE ONE OUT INVALID',invalid)
proxy._qc_use_model_safeguards=True
with contextlib.redirect_stdout(io.StringIO()):
    scores, _, mapping, md = cls.perform_factor_analysis(proxy,users)
    _, b, _, _ = cls.perform_regression_analysis(proxy,users.join(scores),mapping)
report['safeguard_check']={'mapping_diagnostics':md,'B_E':float(b.Beta.abs().N_E/b.Beta.abs().sum()*100)}
print('SAFEGUARD',report['safeguard_check'])
summary=pd.read_excel(next((ROOT/'data').glob('*BS Output.xlsx')),sheet_name='Summary')
errors=[]
for _,row in summary.iterrows():
    sample=clean
    if row['Code Index1']:
        sample=sample[sample.Index1==row['Code Index1']]
    if 'S13_AUTO_1=' in row.Filter:
        code=2 if 'Competitor' in row.Filter else 1
        sample=sample[sample.S13_AUTO_1==code]
    with contextlib.redirect_stdout(io.StringIO()):
        scores, _, mapping, _ = cls.perform_factor_analysis(proxy,sample)
        _, b, _, _ = cls.perform_regression_analysis(proxy,sample.join(scores),mapping)
    ratios_full=b.Beta.abs()/b.Beta.abs().sum()*100
    errors.extend(abs(float(ratios_full[c])-row[out]) for c,out in [('N_S','B.S'),('N_P','B.P'),('N_C','B.C'),('N_E','B.E')])
report['summary_reconciliation']={'groups':len(summary),'max_abs_difference':max(errors)}
print('ALL SUMMARY CHECK',report['summary_reconciliation'])
output=ROOT/'outputs/brandsense_set8_diagnosis.json'
output.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print('REPORT',output)
