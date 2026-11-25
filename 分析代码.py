#!/usr/bin/env python3
"""Reproduce the exploratory review analysis from the supplied Excel workbook.

Usage: python 分析代码.py /path/to/卖家评论.xlsx --out ./数据结果
Only reads the input workbook; requires pandas, numpy and openpyxl.
All tests are exploratory, with independent-review assumptions.
"""
import argparse
import hashlib
import html
import json
import math
import platform
import re
from pathlib import Path
from statistics import NormalDist
from urllib.parse import urlparse, parse_qs
import numpy as np
import pandas as pd

Z = NormalDist().inv_cdf(.975)

def wilson(k, n):
    if not n:
        return [None, None]
    p = k/n
    center = (p + Z*Z/(2*n))/(1 + Z*Z/n)
    half = Z*math.sqrt(p*(1-p)/n + Z*Z/(4*n*n))/(1 + Z*Z/n)
    return [max(0, center-half), min(1, center+half)]

def fisher_two_sided(a,b,c,d):
    """Sum fixed-margin hypergeometric probabilities <= observed probability."""
    n=a+b+c+d; r=a+b; k=a+c
    def lc(n,k):
        return math.lgamma(n+1)-math.lgamma(k+1)-math.lgamma(n-k+1)
    def logp(x):
        return lc(k,x)+lc(n-k,r-x)-lc(n,r)
    observed=logp(a)
    vals=[math.exp(logp(x)) for x in range(max(0,r-(n-k)),min(r,k)+1)
          if logp(x)<=observed+1e-10]
    return min(1.0, math.fsum(vals))

def comparison(df, cutoff=2):
    x=df['image_marked'].astype(bool); y=df['rating'].le(cutoff)
    a=int((x&y).sum()); b=int((x&~y).sum())
    c=int((~x&y).sum()); d=int((~x&~y).sum())
    n1=a+b; n0=c+d; n=n1+n0
    if not n1 or not n0:
        return {'n':n,'table':[[a,b],[c,d]],'estimable':False}
    p1=a/n1; p0=c/n0; diff=p1-p0
    l1,u1=wilson(a,n1); l0,u0=wilson(c,n0)
    # Newcombe hybrid score interval without continuity correction.
    diff_ci=[diff-math.sqrt((p1-l1)**2+(u0-p0)**2),
             diff+math.sqrt((u1-p1)**2+(p0-l0)**2)]
    expected=np.outer([n1,n0],[a+c,b+d])/n
    observed=np.array([[a,b],[c,d]],dtype=float)
    chi=float(((observed-expected)**2/expected).sum()) if (expected>0).all() else None
    result={'n':n,'table':[[a,b],[c,d]],'estimable':True,
      'image_rate':p1,'unmarked_rate':p0,'image_wilson95':[l1,u1],
      'unmarked_wilson95':[l0,u0],'difference':diff,'difference_newcombe95':diff_ci,
      'fisher_p':fisher_two_sided(a,b,c,d),'pearson_chi2':chi,
      'pearson_df':1,'pearson_p':math.erfc(math.sqrt(chi/2)) if chi is not None else None,
      'minimum_expected':float(expected.min())}
    if min(a,b,c,d)>0:
        rr=p1/p0; rr_se=math.sqrt(1/a-1/n1+1/c-1/n0)
        odds=a*d/(b*c); or_se=math.sqrt(1/a+1/b+1/c+1/d)
        result.update(ratio=rr,ratio_log95=[math.exp(math.log(rr)-Z*rr_se),math.exp(math.log(rr)+Z*rr_se)],
                      odds_ratio=odds,odds_ratio_log95=[math.exp(math.log(odds)-Z*or_se),math.exp(math.log(odds)+Z*or_se)])
    return result

def cmh(tables):
    total_num=total_den=delta=variance=0.
    valid=0
    for t in tables:
        a,b=t[0];c,d=t[1];n=a+b+c+d
        if n<=1 or not (a+b)*(c+d)*(a+c)*(b+d):continue
        valid+=1
        total_num+=a*d/n;total_den+=b*c/n
        delta+=a-(a+b)*(a+c)/n
        variance+=(a+b)*(c+d)*(a+c)*(b+d)/(n*n*(n-1))
    chi=delta*delta/variance if variance else None
    return {'informative_strata':valid,'common_odds_ratio':total_num/total_den if total_den else None,
            'cmh_chi2_no_correction':chi,'cmh_p':math.erfc(math.sqrt(chi/2)) if chi is not None else None}

def plain(v):
    if pd.isna(v):return ''
    v=re.sub(r'<br\s*/?>','. ',str(v),flags=re.I)
    v=re.sub(r'<[^>]+>',' ',v)
    return re.sub(r'\s+',' ',html.unescape(v)).strip().strip('.').strip()

def parse_source(path):
    xls=pd.ExcelFile(path); raw=pd.read_excel(xls,sheet_name=0)
    required=['型号','标题','标题 (翻译)','内容','内容(翻译)','星级','图片数量','评论时间','链接']
    if missing:=set(required)-set(raw.columns):raise ValueError(f'Missing columns: {missing}')
    df=raw.copy();df['source_row']=range(2,len(df)+2)
    df['review_id']=df['链接'].str.extract(r'/customer-reviews/([^/?]+)',expand=False)
    if df.review_id.isna().any():raise ValueError('Missing review identifier')
    conflicting=df.groupby('review_id')[required].nunique(dropna=False).gt(1).any(axis=1)
    if conflicting.any():raise ValueError('Conflicting duplicated review IDs')
    df=df.drop_duplicates('review_id').copy()
    df['rating']=pd.to_numeric(df['星级'],errors='coerce')
    df['date']=pd.to_datetime(df['评论时间'],errors='coerce')
    if (~df.rating.isin([1,2,3,4,5])).any() or df.date.isna().any():
        raise ValueError('Invalid rating/date; specify exclusions explicitly')
    pic=pd.to_numeric(df['图片数量'],errors='coerce')
    if pic.lt(0).any():raise ValueError('Negative image count')
    df['image_marked']=pic.gt(0)
    df['image_count_missing']=pic.isna()
    df['low_rating']=df.rating.le(2)
    df['month']=df.date.dt.strftime('%Y-%m')
    for col,target in [('标题','title'),('标题 (翻译)','title_translation'),('内容','body'),('内容(翻译)','translation')]:
        df[target]=df[col].map(plain)
    df['body_normalized']=df.body.str.casefold().str.replace(r'\s+',' ',regex=True).str.strip()
    df['model']=df['型号'].map(plain)
    df['size']=df.model.str.extract(r'Size:\s*(.*)',expand=False).fillna('未识别').str.casefold().str.replace(r'^for\s+','',regex=True).str.strip()
    df['asin_parameter']=df['链接'].map(lambda s:parse_qs(urlparse(str(s)).query).get('ASIN',[''])[0])
    quality={'source_filename':path.name,'source_sha256':hashlib.sha256(path.read_bytes()).hexdigest(),
       'sheet':xls.sheet_names[0],'raw_n':len(raw),'analysis_n':len(df),'duplicate_ids_removed':len(raw)-len(df),
       'start':str(df.date.min().date()),'end':str(df.date.max().date()),'original_missing':raw.isna().sum().astype(int).to_dict(),
       'raw_model_count':int(df.model.nunique()),'size_count':int(df['size'].nunique()),
       'normalized_duplicate_body_rows':int(df.body_normalized.duplicated().sum()),
       'asin_parameters':df.asin_parameter.value_counts().to_dict()}
    return df,quality

def themes(df,config,original_only=False):
    summary=[];evidence=[]
    for group in config['groups']:
        eligible=df[(df.body.ne('')|df.translation.ne(''))].copy()
        if group['population']=='low':eligible=eligible[eligible.rating.le(2)]
        elif group['population']=='high':eligible=eligible[eligible.rating.ge(4)]
        fields=['title','body'] if original_only else ['title','body','title_translation','translation']
        for rule in group['rules']:
            allp=[re.compile(p,re.I) for p in rule['all_of']]
            excl=[re.compile(p,re.I) for p in rule.get('exclude',[])]
            count=0
            for row in eligible.to_dict('records'):
                text='. '.join(row[k] for k in fields if row[k]).replace('’',"'").replace('‘',"'")
                sentences=[s.strip() for s in re.split(r'[.!?。！？;；\n]+',text) if s.strip()]
                hit=next((s for s in sentences if all(p.search(s) for p in allp) and not any(p.search(s) for p in excl)),None)
                if hit is not None:
                    count+=1
                    evidence.append({'group':group['id'],'rule_id':rule['id'],'topic':rule['label'],
                      'review_id':row['review_id'],'source_row':row['source_row'],'rating':row['rating'],'matched_sentence':hit})
            summary.append({'group':group['id'],'rule_id':rule['id'],'topic':rule['label'],
                            'count':count,'denominator':len(eligible),'share':count/len(eligible) if len(eligible) else None})
    return pd.DataFrame(summary),pd.DataFrame(evidence)

def run(path,out):
    out.mkdir(parents=True,exist_ok=True)
    df,quality=parse_source(path)
    config_path=Path(__file__).with_name('主题规则.json')
    config=json.loads(config_path.read_text(encoding='utf-8'))
    main=comparison(df)
    monthly={m:comparison(g) for m,g in df.groupby('month',sort=True)}
    model_tables=[comparison(g)['table'] for _,g in df.groupby('size')]
    dedup=df.drop_duplicates('body_normalized',keep='first')
    sensitivities={
        '正文规范化去重':comparison(dedup),
        '低评分改为仅1星':comparison(df,1),
        '排除3星后比较':comparison(df[df.rating.ne(3)]),
        '仅分析7月与8月记录':comparison(df[df.month.isin(['2026-07','2026-08'])]),
        '按月分层':cmh([v['table'] for v in monthly.values()]),
        '按手机型号分层':cmh(model_tables)}
    stars=[{'star':s,'count':int(df.rating.eq(s).sum()),'share':float(df.rating.eq(s).mean())} for s in range(1,6)]
    theme,evidence=themes(df,config)
    original,_=themes(df,config,True)
    negative=theme[theme.group.eq('negative')].sort_values('count',ascending=False)
    matched_negative=evidence[evidence.group.eq('negative')].review_id.nunique()
    result={'quality':quality,'overall':{'low_n':int(df.low_rating.sum()),'n':len(df),'rate':float(df.low_rating.mean()),
          'wilson95':wilson(int(df.low_rating.sum()),len(df))},'stars':stars,'main':main,'monthly':monthly,
          'sensitivity':sensitivities,'negative_themes':negative.to_dict('records'),
          'negative_covered_reviews':int(matched_negative),'negative_uncovered_reviews':int(df.low_rating.sum()-matched_negative),
          'theme_evidence_rows':len(evidence),'rules_sha256':hashlib.sha256(config_path.read_bytes()).hexdigest(),
          'environment':{'python':platform.python_version(),'pandas':pd.__version__,'numpy':np.__version__},
          'notes':['Observational convenience export, not random sampling.','Image missing means unmarked, not confirmed no-image.',
                   'Exploratory hypotheses chosen after inspecting the dataset.','No human gold-standard topic labels available.']}
    (out/'统计结果.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,allow_nan=False),encoding='utf-8')
    pd.DataFrame(stars).to_csv(out/'星级分布.csv',index=False,encoding='utf-8-sig')
    df[['review_id','source_row','rating','date','month','model','size','image_marked','image_count_missing','low_rating']].to_csv(out/'评论分析明细.csv',index=False,encoding='utf-8-sig')
    theme.to_csv(out/'主题统计.csv',index=False,encoding='utf-8-sig')
    evidence.to_csv(out/'主题命中证据.csv',index=False,encoding='utf-8-sig')
    original.to_csv(out/'仅原文主题敏感性.csv',index=False,encoding='utf-8-sig')
    pd.DataFrame([{'month':m,**{k:v[k] for k in ['n','image_rate','unmarked_rate','difference','fisher_p']}} for m,v in monthly.items()]).to_csv(out/'按月比较.csv',index=False,encoding='utf-8-sig')
    # Uniform random low-rating sample for future human validation. Labels remain blank.
    sample=df[df.low_rating].sample(n=min(120,int(df.low_rating.sum())),random_state=6165).sort_values('source_row')
    sample=sample[['review_id','source_row','title','title_translation','body','translation']].copy()
    for name in ['做工耐用性','磁性部件','充电','颜色','打滑','是否涉及其他商品']:
        sample[name+'_标注员A']='';sample[name+'_标注员B']='';sample[name+'_协商结果']=''
    sample['备注']=''
    sample.to_csv(out/'人工复核样本_尚未标注.csv',index=False,encoding='utf-8-sig')
    print(json.dumps({'overall':result['overall'],'main':main,'sensitivity':sensitivities,
                      'negative_coverage':[matched_negative,int(df.low_rating.sum())]},ensure_ascii=False,indent=2))

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('input',type=Path);parser.add_argument('--out',type=Path,default=Path('数据结果'))
    args=parser.parse_args();run(args.input,args.out)
