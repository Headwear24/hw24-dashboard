import pandas as pd, os, json
 
# Global state — set by initialise()
df = afs27 = afs26 = None
UPLIFT=1.19; WD_PER=9/20; prov_col='Province Orginal'; sp_col='Sales Person Use'
 
def initialise(filepath, wd_per=9/20, w2_fraction=4/5):
    """Load data file and set global config. Call before calc()."""
    global df, afs27, afs26, UPLIFT, WD_PER, afs26_ytd, afs26_jun
    global sp_lu, branch_dict
    WD_PER = wd_per
    df = pd.read_excel(filepath, sheet_name='AFS 26 Data', header=0)
    afs27 = df[df['AFS']=='AFS27'].copy()
    afs26 = df[df['AFS']=='AFS26'].copy()
    for d_ in [afs27, afs26]:
        d_['Region'] = d_[prov_col].map(REGION_MAP).fillna('Other')
    afs26_ytd = afs26[afs26['Period1'].isin(['Mar','Apr','May'])]
    afs26_jun = afs26[afs26['Period1']=='Jun']
    sp_lu = df[['ALYSHA','DURBAN']].dropna().drop_duplicates()
    sp_lu.columns = ['sp','branch']
    sp_lu['branch'] = sp_lu['branch'].str.strip()
    branch_dict.update(dict(zip(sp_lu['sp'], sp_lu['branch'])))
REGION_MAP={'GP':'Gauteng','LP':'Gauteng','NW':'Gauteng','MP':'Gauteng','FS':'Gauteng',
            'KZN':'KZN','ZN':'KZN','WC':'Western Cape','NC':'Western Cape',
            'EC':'Eastern Cape','ZZZ':'International'}
branch_dict={}
branch_reg={'DURBAN':'KZN','CAPE TOWN':'Western Cape','JOHANNESBURG':'Gauteng','PORT ELIZABETH':'Eastern Cape'}
GCOLORS={'LPCTO':'#2E75B6','STPRO':'#ED7D31','LPMTO':'#375623','LPIMP':'#843C0C',
         'UFLEX':'#595959','RAWMT':'#7B6E58','REPLEN':'#888780','OPP':'#C00000'}
SUB_PROVS={'Gauteng':[('GP','Gauteng core'),('FS','Free State'),('MP','Mpumalanga'),('NW','North West'),('LP','Limpopo')],
           'KZN':[('KZN','KZN core')],'Western Cape':[('WC','WC core'),('NC','Northern Cape')],
           'Eastern Cape':[('EC','EC core')],'International':[]}
COUNTRY_MAP={'SZ':'Swaziland','BW':'Botswana','MZ':'Mozambique','ZW':'Zimbabwe','ZM':'Zambia',
             'NA':'Namibia','LS':'Lesotho','MU':'Mauritius','MW':'Malawi','GB':'Great Britain'}
 
def pct_str(a,t): p=(a-t)/t*100 if t else 0; return f"({abs(p):.1f}%)" if p<0 else f"+{p:.1f}%"
def delta_str(a,t): d=a-t; return f"(R{abs(d)/1e6:.2f}M)" if d<0 else f"+R{d/1e6:.2f}M"
def col(a,t): return '#C00000' if a<t else '#375623'
def fmtR(v): return f"R{v/1e6:.2f}M" if v>=1e6 else f"R{v/1000:.1f}K"
def fmtN(v): return f"{v:,.0f}"
 
def calc(region):
    r27=afs27[afs27['Region']==region]; r26y=afs26_ytd[afs26_ytd['Region']==region]
    r26j=afs26_jun[afs26_jun['Region']==region]
    r27y=r27[r27['Period1'].isin(['Mar','Apr','May'])]; r27j=r27[r27['Period1']=='Jun']
    total_rev=r27['Line Revenue'].sum(); total_units=r27['Line Inv Qty'].sum()
    avg_price=total_rev/total_units if total_units else 0
    clients=r27['Cust Name'].nunique(); invoices=r27['Inv no'].nunique()
    py_ytd=r26y['Line Revenue'].sum(); py_jun=r26j['Line Revenue'].sum()
    tgt_ytd=round(py_ytd*UPLIFT); tgt_jun_full=round(py_jun*UPLIFT)
    tgt_jun_pro=round(py_jun*UPLIFT*WD_PER); weekly_tgt=round(tgt_jun_full/4)
    w1_tgt=weekly_tgt; w2_tgt=round(weekly_tgt*4/5)
    ytd_act=r27y['Line Revenue'].sum(); jun_act=r27j['Line Revenue'].sum()
    w1_act=int(r27[(r27['Period1']=='Jun')&(r27['Week']=='Week 1')]['Line Revenue'].sum())
    w2_act=int(r27[(r27['Period1']=='Jun')&(r27['Week']=='Week 2')]['Line Revenue'].sum())
    chart_max=int(max(w1_act, round(py_jun*UPLIFT/4), w2_act, round(py_jun*UPLIFT/4*2/5)) * 1.2) if py_jun>0 else max(w1_act,w2_act,1000)*2
    monthly={}
    for p1 in ['Mar','Apr','May','Jun']:
        md=r27[r27['Period1']==p1]; pd26=afs26[afs26['Region']==region][afs26['Period1']==p1]['Line Revenue'].sum()
        act=md['Line Revenue'].sum(); units=md['Line Inv Qty'].sum()
        avg=act/units if units else 0
        tgt=round(pd26*UPLIFT*(WD_PER if p1=='Jun' else 1))
        monthly[p1]={'act':act,'tgt':tgt,'avg':avg,'units':units}
    classes=['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN ','OPP']
    prod={}
    for cls in classes:
        cd=r27[r27['CLASS']==cls]; rev=cd['Line Revenue'].sum(); qty=cd['Line Inv Qty'].sum()
        py_cd=r26y[r26y['CLASS']==cls]['Line Revenue'].sum()+r26j[r26j['CLASS']==cls]['Line Revenue'].sum()*WD_PER
        prod[cls.strip()]={'rev':float(rev),'avg':float(rev/qty if qty else 0),'tgt':int(round(py_cd*UPLIFT))}
    ti=r27[r27['CLASS']!='LPMTO'].groupby(['Item No (Stock)','CLASS']).agg(rev=('Line Revenue','sum'),qty=('Line Inv Qty','sum')).reset_index()
    ti['avg']=ti['rev']/ti['qty']; ti['pct']=ti['rev']/total_rev*100 if total_rev else 0
    top8=ti.nlargest(min(8,len(ti)),'rev').to_dict('records')
    bot8=ti[(ti['rev']>=2000)&(ti['qty']>=5)].nsmallest(min(8,len(ti)),'rev').to_dict('records')
    top10=r27.groupby('Cust Name')['Line Revenue'].sum().nlargest(10)
    cp26=r26y.groupby('Cust Name')['Line Revenue'].sum()
    rev27_all=r27.groupby('Cust Name').agg(rev=('Line Revenue','sum'),inv=('Inv no','nunique')).reset_index()
    mg=rev27_all.merge(cp26.rename('py_rev'),on='Cust Name',how='left'); mg['py_rev']=mg['py_rev'].fillna(0)
    new_cl=mg[mg['py_rev']==0].nlargest(min(8,len(mg[mg['py_rev']==0])),'rev')
    both_cl=mg[mg['py_rev']>0].copy(); both_cl['drop']=both_cl['rev']-both_cl['py_rev']
    worst_cl=both_cl.nsmallest(min(8,len(both_cl)),'drop')
    excl=['WEBSITE','COMPANY','DIRECT','']
    bnames=[b for b,rg in branch_reg.items() if rg==region]
    sp_branch=afs27[afs27[sp_col].map(lambda x: branch_dict.get(x,'') in bnames)&~afs27[sp_col].isin(excl)].groupby(sp_col)['Line Revenue'].sum().nlargest(8)
    sp26_b=afs26_ytd[afs26_ytd[sp_col].map(lambda x: branch_dict.get(x,'') in bnames)&~afs26_ytd[sp_col].isin(excl)].groupby(sp_col)['Line Revenue'].sum()
    sp_reg=r27[~r27[sp_col].isin(excl)].groupby(sp_col)['Line Revenue'].sum().nlargest(8)
    sp26_r=r26y[~r26y[sp_col].isin(excl)].groupby(sp_col)['Line Revenue'].sum()
    sub_data=[]
    for pc,pn in SUB_PROVS.get(region,[]):
        cv=r27[r27[prov_col]==pc]['Line Revenue'].sum()
        pyv=r26y[r26y[prov_col]==pc]['Line Revenue'].sum()+r26j[r26j[prov_col]==pc]['Line Revenue'].sum()*WD_PER
        sub_data.append({'code':pc,'name':pn,'v':cv,'t':round(pyv*UPLIFT)})
    if region=='International':
        for ctry,rev in r27.groupby('Country')['Line Revenue'].sum().nlargest(5).items():
            sub_data.append({'code':ctry,'name':COUNTRY_MAP.get(ctry,ctry),'v':rev,'t':0})
    py_s=r26y[r26y['Line Revenue']>0].groupby('Cust Name').agg(py_rev=('Line Revenue','sum'),py_qty=('Line Inv Qty','sum')).reset_index()
    pt=r26y.groupby('Cust Name')['Line Revenue'].sum().reset_index(); pt.columns=['Cust Name','py_total']
    py_s=py_s.merge(pt,on='Cust Name'); py_s=py_s[py_s['py_total']>=20000].copy(); py_s['py_avg']=py_s['py_rev']/py_s['py_qty']
    cy_s=r27[r27['Line Revenue']>0].groupby('Cust Name').agg(cy_rev=('Line Revenue','sum'),cy_qty=('Line Inv Qty','sum')).reset_index()
    cy_s['cy_avg']=cy_s['cy_rev']/cy_s['cy_qty']; avg_m=py_s.merge(cy_s,on='Cust Name',how='inner')
    valid_avg=avg_m[avg_m['cy_avg']>=5.0]
    top_avg=valid_avg.nlargest(min(8,len(valid_avg)),'cy_avg'); bot_avg=valid_avg.nsmallest(min(8,len(valid_avg)),'cy_avg')
    cum_act=[0,0,0,0]; cum_tgt=[0,0,0,0]; cum_py=[0,0,0,0]; ra=rt=rp=0
    for i,p1 in enumerate(['Mar','Apr','May','Jun']):
        ra+=monthly[p1]['act']; rt+=monthly[p1]['tgt']
        pm=afs26[afs26['Region']==region][afs26['Period1']==p1]['Line Revenue'].sum()
        if p1=='Jun': pm*=WD_PER
        rp+=pm; cum_act[i]=int(round(ra)); cum_tgt[i]=int(round(rt)); cum_py[i]=int(round(rp))
    # Avg unit prices for period table
    r27_ytd2 = r27[r27['Period1'].isin(['Mar','Apr','May'])]
    r27_jun2 = r27[r27['Period1']=='Jun']
    ytd_qty2 = float(r27_ytd2['Line Inv Qty'].sum()); jun_qty2 = float(r27_jun2['Line Inv Qty'].sum())
    ytd_avg_price = float(r27_ytd2['Line Revenue'].sum()/ytd_qty2 if ytd_qty2 else 0)
    jun_avg_price = float(r27_jun2['Line Revenue'].sum()/jun_qty2 if jun_qty2 else 0)
    tot_avg_price = float(float(r27['Line Revenue'].sum())/float(r27['Line Inv Qty'].sum()) if r27['Line Inv Qty'].sum() else 0)
 
    # Units targets for period table
    ytd_units_act = int(r27_ytd2['Line Inv Qty'].sum())
    jun_units_act = int(r27_jun2['Line Inv Qty'].sum())
    tot_units_act = int(r27['Line Inv Qty'].sum())
    ytd_units_py  = int(r26y['Line Inv Qty'].sum())
    jun_units_py  = int(r26j['Line Inv Qty'].sum())
    ytd_units_tgt = int(round(ytd_units_py * UPLIFT))
    jun_units_tgt = int(round(jun_units_py * UPLIFT * WD_PER))
    tot_units_tgt = ytd_units_tgt + jun_units_tgt
 
    # Monthly avg prices for chart secondary axis
    monthly_avg_prices = []
    for p1 in ['Mar','Apr','May','Jun']:
        md2 = r27[r27['Period1']==p1]
        q2 = md2['Line Inv Qty'].sum()
        monthly_avg_prices.append(float(round(md2['Line Revenue'].sum()/q2, 2)) if q2>0 else 0.0)
 
    return dict(region=region,total_rev=total_rev,total_units=total_units,avg_price=avg_price,
                ytd_avg_price=ytd_avg_price,jun_avg_price=jun_avg_price,tot_avg_price=tot_avg_price,
                ytd_units_act=ytd_units_act,ytd_units_tgt=ytd_units_tgt,
                jun_units_act=jun_units_act,jun_units_tgt=jun_units_tgt,chart_max=chart_max,
                tot_units_act=tot_units_act,tot_units_tgt=tot_units_tgt,
                monthly_avg_prices=monthly_avg_prices,
                clients=clients,invoices=invoices,ytd_act=ytd_act,tgt_ytd=tgt_ytd,
                ytd_delta=ytd_act-tgt_ytd,tgt_jun_pro=tgt_jun_pro,tgt_jun_full=tgt_jun_full,
                weekly_tgt=weekly_tgt,w1_act=w1_act,w1_tgt=w1_tgt,w2_act=w2_act,w2_tgt=w2_tgt,
                jun_act=jun_act,total_tgt_mtd=tgt_ytd+tgt_jun_pro,monthly=monthly,prod=prod,
                top8=top8,bot8=bot8,top10=top10,cp26=cp26,new_cl=new_cl,worst_cl=worst_cl,
                sp_branch=sp_branch,sp26_b=sp26_b,sp_reg=sp_reg,sp26_r=sp26_r,
                sub_data=sub_data,top_avg=top_avg,bot_avg=bot_avg,
                cum_act=cum_act,cum_tgt=cum_tgt,cum_py=cum_py)
 
REG_COLS={'Gauteng':'#185FA5','KZN':'#ED7D31','Western Cape':'#375623','Eastern Cape':'#843C0C','International':'#444444'}
MONTH_LABELS={'Mar':'March','Apr':'April','May':'May','Jun':'June'}
 
def build_html(d, report_date=None):
    import datetime
    if report_date is None:
        report_date = datetime.date.today().strftime('%-d %b %Y')
    r=d['region']; c=REG_COLS[r]; m=d['monthly']
    ytd_d=d['ytd_delta']; ytd_p=ytd_d/d['tgt_ytd']*100 if d['tgt_ytd'] else 0
    tgt_mtd=d['w1_tgt']+d['w2_tgt']; mtd_d=d['jun_act']-tgt_mtd
    total_d=d['total_rev']-d['tgt_ytd']-tgt_mtd
 
    # Avg prices + units for period table
    ytd_avg_price = d['ytd_avg_price']; jun_avg_price = d['jun_avg_price']; tot_avg_price = d['tot_avg_price']
    ytd_units_act = d['ytd_units_act']; ytd_units_tgt = d['ytd_units_tgt']
    jun_units_act = d['jun_units_act']; jun_units_tgt = d['jun_units_tgt']
    tot_units_act = d['tot_units_act']; tot_units_tgt = d['tot_units_tgt']
    ytd_units_d = ytd_units_act - ytd_units_tgt; jun_units_d = jun_units_act - jun_units_tgt
    tot_units_d = tot_units_act - tot_units_tgt
    chart_max_mtd = d['chart_max']
 
    # Monthly avg prices for charts
    avg_prices = d['monthly_avg_prices']
    # Determine sensible y2 axis range
    valid_avgs = [v for v in avg_prices if v > 0]
    y2_min = max(0, round(min(valid_avgs)/10)*10 - 10) if valid_avgs else 0
    y2_max = round(max(valid_avgs)/10)*10 + 20 if valid_avgs else 100
 
    # Monthly values for f-string injection
    mar_act=int(m['Mar']['act']); mar_tgt=int(m['Mar']['tgt'])
    apr_act=int(m['Apr']['act']); apr_tgt=int(m['Apr']['tgt'])
    may_act=int(m['May']['act']); may_tgt=int(m['May']['tgt'])
    jun_act_v=int(d['jun_act']);  jun_full_tgt=int(d['tgt_jun_full'])
    ytd_act=int(d['ytd_act']);   tgt_ytd=int(d['tgt_ytd']); ytd_delta=int(d['ytd_delta'])
    w1_act=int(d['w1_act']);     w1_tgt=int(d['w1_tgt'])
    w2_act=int(d['w2_act']);     w2_tgt=int(d['w2_tgt'])
    total_rev=int(d['total_rev']); jun_mtd_tgt=int(d['w1_tgt']+d['w2_tgt'])
 
    # Remaining month targets from pre-computed dict (fast)
    def _rm(p1): return round(afs26[afs26['Region']==r][afs26['Period1']==p1]['Line Revenue'].sum()*1.19)
    jul_tgt=_rm('Jul'); aug_tgt=_rm('Aug'); sep_tgt=_rm('Sept')
    oct_tgt=_rm('Oct'); nov_tgt=_rm('Nov'); dec_tgt=_rm('Dec')
    jan_tgt=_rm('Jan'); feb_tgt=_rm('Feb')
 
    def row_color(a,t): return '#C00000' if a<t else '#375623'
 
    # Product class variables for matrix
    def pv(cls,k): return d['prod'].get(cls,{}).get(k,0)
    lpcto_rev=pv('LPCTO','rev'); lpcto_avg=pv('LPCTO','avg')
    stpro_rev=pv('STPRO','rev'); stpro_avg=pv('STPRO','avg')
    lpmto_rev=pv('LPMTO','rev'); lpmto_avg=pv('LPMTO','avg')
    lpimp_rev=pv('LPIMP','rev'); lpimp_avg=pv('LPIMP','avg')
    uflex_rev=pv('UFLEX','rev'); uflex_avg=pv('UFLEX','avg')
    rawmt_rev=pv('RAWMT','rev'); rawmt_avg=pv('RAWMT','avg')
    replen_rev=pv('REPLEN','rev'); replen_avg=pv('REPLEN','avg')
    opp_rev=pv('OPP','rev');   opp_avg=pv('OPP','avg')
    def neg(v): return f"({fmtR(abs(v))})" if v<0 else f"+{fmtR(v)}"
    def negpct(v,t): p=v/t*100 if t else 0; return f"({abs(p):.1f}%)" if p<0 else f"+{p:.1f}%"
    def tbl_row(label,act,tgt,avg=None,indent=False,bold=False,total=False):
        d_=act-tgt; p_=d_/tgt*100 if tgt else 0
        dc=row_color(act,tgt); style='class="total-row"' if total else ('class="bold-row"' if bold else '')
        lbl=f'<span class="italic-lbl">{label}</span>' if indent else f'<span>{label}</span>'
        avg_cell=f'<span class="avg-col" style="text-align:right">R{avg:.2f}</span>' if avg is not None else '<span style="text-align:right;color:#888">—</span>'
        return f'<div {style} style="display:grid;grid-template-columns:1.9fr 1.1fr 1.1fr 1fr .9fr 1fr;font-size:11px;padding:6px 12px;gap:8px;border-bottom:1px solid #f0f0f0;background:{"#DCE6F1" if total else ("" if not bold else "")}">'+\
               f'{lbl}<span style="text-align:right">{fmtN(act)}</span><span style="text-align:right">{fmtN(tgt)}</span>'+\
               f'<span style="text-align:right;color:{dc}">{"("+fmtN(abs(d_))+")" if d_<0 else "+"+fmtN(d_)}</span>'+\
               f'<span style="text-align:right;color:{dc}">{"("+str(abs(round(p_,1)))+"%" if p_<0 else "+"+str(round(p_,1))+"%"}</span>'+\
               f'{avg_cell}</div>'
 
    def sp_rows(sp_ser, sp26_ser, max_=8):
        mx=max(list(sp_ser.values())+[v for v in sp26_ser.values()]) if len(sp_ser) else 1
        rows=[]
        for i,(name,rev) in enumerate(list(sp_ser.items())[:max_]):
            py=sp26_ser.get(name,0); yoy=(rev-py)/py*100 if py else None
            ys='<span style="background:#375623;color:#fff;font-size:8px;padding:1px 4px;border-radius:3px">NEW</span>' if not py \
               else (f'<span style="color:#375623;font-weight:700">+{yoy:.1f}%</span>' if yoy>=0 \
               else f'<span style="color:#C00000;font-weight:700">({abs(yoy):.1f}%)</span>')
            branch=branch_dict.get(name,'—')
            cy_w=round(rev/mx*100); py_w=round(py/mx*100) if py else 0
            bg='#fff' if i%2==0 else '#f9f9f9'
            rows.append(f'''<div style="display:grid;grid-template-columns:24px 1fr 80px 70px 70px 60px;gap:6px;align-items:center;padding:7px 10px;border-bottom:1px solid #f0f0f0;font-size:11px;background:{bg}">
  <span style="color:#888;font-size:10px">{i+1}</span>
  <div><div style="font-weight:700">{name}</div>
  <div style="display:flex;flex-direction:column;gap:2px;margin-top:3px">
    <div style="display:flex;align-items:center;gap:3px"><span style="font-size:8px;color:#888;width:16px">CY</span><div style="flex:1;height:4px;background:#e0e0e0;border-radius:2px;overflow:hidden"><div style="width:{cy_w}%;height:4px;background:{c};border-radius:2px"></div></div><span style="font-size:8px;color:#666;min-width:28px;text-align:right">{fmtR(rev)}</span></div>
    {"" if not py else f'<div style="display:flex;align-items:center;gap:3px"><span style="font-size:8px;color:#888;width:16px">PY</span><div style="flex:1;height:4px;background:#e0e0e0;border-radius:2px;overflow:hidden"><div style="width:{py_w}%;height:4px;background:#A9C4E4;border-radius:2px"></div></div><span style="font-size:8px;color:#666;min-width:28px;text-align:right">{fmtR(py)}</span></div>'}
  </div></div>
  <span style="font-size:10px;color:#444">{branch}</span>
  <span style="text-align:right;font-weight:700">{fmtR(rev)}</span>
  <span style="text-align:right;color:#666">{fmtR(py) if py else "—"}</span>
  <span style="text-align:right">{ys}</span>
</div>''')
        return ''.join(rows)
 
    def client_rows(top10, cp26, max_=10):
        maxR_=max(top10.values) if len(top10) else 1; rows=[]
        for i,(name,rev) in enumerate(list(top10.items())[:max_]):
            py=cp26.get(name,0); yoy=(rev-py)/py*100 if py else None
            ys='<span style="background:#375623;color:#fff;font-size:8px;padding:1px 4px;border-radius:3px">NEW</span>' if not py \
               else (f'<span style="color:#375623;font-weight:700">+{yoy:.1f}%</span>' if yoy>=0 \
               else f'<span style="color:#C00000;font-weight:700">({abs(yoy):.1f}%)</span>')
            bg='#fff' if i%2==0 else '#f9f9f9'; w=round(rev/maxR_*100)
            rows.append(f'''<div style="display:grid;grid-template-columns:22px 1fr 70px 70px 60px;gap:8px;align-items:center;padding:6px 12px;border-bottom:1px solid #f0f0f0;font-size:11px;background:{bg}">
  <span style="color:#888;font-size:10px">{i+1}</span>
  <div><div style="font-weight:700">{name}</div><div style="width:{w}%;height:3px;background:{c};border-radius:2px;margin-top:3px"></div></div>
  <span style="text-align:right;font-weight:700">{fmtR(rev)}</span>
  <span style="text-align:right;color:#666">{"—" if not py else fmtR(py)}</span>
  <span style="text-align:right">{ys}</span>
</div>''')
        return ''.join(rows)
 
    def avg_rows(df_a, is_top, max_=8):
        if len(df_a)==0: return '<div style="padding:10px;color:#888;font-size:10px">Insufficient data</div>'
        mx=max(max(df_a['cy_avg']),max(df_a['py_avg'])); rows=[]
        for i,row in enumerate(df_a.head(max_).itertuples()):
            chg=row.cy_avg-row.py_avg; cy_col='#375623' if is_top else '#C00000'; chg_col='#375623' if chg>=0 else '#C00000'
            bg='#fff' if i%2==0 else '#f9f9f9'
            py_w=round(row.py_avg/mx*100); cy_w=round(row.cy_avg/mx*100)
            rows.append(f'''<div style="display:grid;grid-template-columns:20px 1fr 70px 70px 62px 62px 56px;gap:5px;align-items:center;padding:5px 10px;border-bottom:1px solid #f0f0f0;font-size:10px;background:{bg}">
  <span style="color:#888">{i+1}</span>
  <div><div style="font-weight:700;font-size:10px">{row._1}</div>
    <div style="display:flex;flex-direction:column;gap:2px;margin-top:2px">
      <div style="display:flex;align-items:center;gap:3px"><span style="font-size:7px;color:#888;width:16px">PY</span><div style="flex:1;height:3px;background:#e0e0e0;border-radius:2px;overflow:hidden"><div style="width:{py_w}%;height:3px;background:#A9C4E4;border-radius:2px"></div></div></div>
      <div style="display:flex;align-items:center;gap:3px"><span style="font-size:7px;color:#888;width:16px">CY</span><div style="flex:1;height:3px;background:#e0e0e0;border-radius:2px;overflow:hidden"><div style="width:{cy_w}%;height:3px;background:{cy_col};border-radius:2px"></div></div></div>
    </div>
  </div>
  <span style="text-align:right;color:#666">R{row.py_r:.0f}</span>
  <span style="text-align:right">R{row.cy_rev:.0f}</span>
  <span style="text-align:right;color:#888">R{row.py_avg:.2f}</span>
  <span style="text-align:right;font-weight:700;color:{cy_col}">R{row.cy_avg:.2f}</span>
  <span style="text-align:right;font-weight:700;color:{chg_col}">{("+") if chg>=0 else ""}{chg:.2f}</span>
</div>''')
        return ''.join(rows)
 
    # Sub-province / country section
    sub_html=''
    for s in d['sub_data']:
        pct_=(s['v']/s['t']*100) if s['t']>0 else 100; ahead=s['v']>=s['t'] if s['t']>0 else True
        prog_col='#375623' if ahead else ('#C00000' if pct_<70 else '#ED7D31')
        gap_str=f"Gap: ({fmtR(abs(s['v']-s['t']))}) — {abs(100-pct_):.1f}% behind" if not ahead and s['t']>0 else \
                (f"Ahead: +{fmtR(s['v']-s['t'])}" if s['t']>0 else "Top export market")
        gap_col='#C00000' if not ahead and s['t']>0 else '#375623'
        sub_html+=f'''<div style="border:1px solid #e0e0e0;border-radius:6px;padding:10px 12px;background:#fff">
  <div style="font-size:12px;font-weight:700;margin-bottom:3px">{s["name"]}</div>
  <div style="font-size:16px;font-weight:700;color:{gap_col};margin-bottom:2px">{fmtR(s["v"])}</div>
  <div style="font-size:10px;color:{gap_col};margin-bottom:5px">{gap_str}</div>
  <div style="height:4px;background:#ddd;border-radius:2px;margin-bottom:3px"><div style="width:{min(100,round(pct_))}%;height:4px;background:{prog_col};border-radius:2px"></div></div>
</div>'''
 
    # Products horizontal bar JS data
    cls_list=['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN','OPP']
    prod_act=[d['prod'].get(k,{}).get('rev',0) for k in cls_list]
    prod_tgt=[d['prod'].get(k,{}).get('tgt',0) for k in cls_list]
    prod_max=max(max(prod_act),max(prod_tgt))*1.1
 
    items_rows=''.join([f'''<div style="display:grid;grid-template-columns:3fr 36px 62px 44px 34px;gap:3px;padding:4px 7px;border-bottom:1px solid #f0f0f0;font-size:8.5px;background:{"#fff" if i%2==0 else "#f9f9f9"}">
  <span style="font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{it["Item No (Stock)"]}</span>
  <span style="text-align:center"><span style="font-size:7.5px;padding:1px 2px;border-radius:8px;background:{GCOLORS.get(it["CLASS"].strip(),"#888")}22;color:{GCOLORS.get(it["CLASS"].strip(),"#555")};font-weight:700">{it["CLASS"].strip()}</span></span>
  <span style="text-align:right">{it["rev"]:,.0f}</span>
  <span style="text-align:right;color:#666">{it["avg"]:.2f}</span>
  <span style="text-align:right;color:#666">{it["pct"]:.2f}%</span>
</div>''' for i,it in enumerate(d['top8'])])
 
    bot_rows=''.join([f'''<div style="display:grid;grid-template-columns:3fr 36px 62px 44px 34px;gap:3px;padding:4px 7px;border-bottom:1px solid #f0f0f0;font-size:8.5px;background:{"#fff" if i%2==0 else "#f9f9f9"}">
  <span style="font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{it["Item No (Stock)"]}</span>
  <span style="text-align:center"><span style="font-size:7.5px;padding:1px 2px;border-radius:8px;background:{GCOLORS.get(it["CLASS"].strip(),"#888")}22;color:{GCOLORS.get(it["CLASS"].strip(),"#555")};font-weight:700">{it["CLASS"].strip()}</span></span>
  <span style="text-align:right">{it["rev"]:,.0f}</span>
  <span style="text-align:right;color:#666">{it["avg"]:.2f}</span>
  <span style="text-align:right;color:#666">{it["pct"]:.2f}%</span>
</div>''' for i,it in enumerate(d['bot8'])])
 
    top10_rows = client_rows(d['top10'], d['cp26'])
    new_cl_rows = ''.join([f'''<div style="display:grid;grid-template-columns:22px 1fr 70px 58px;gap:8px;align-items:center;padding:6px 12px;border-bottom:1px solid #f0f0f0;font-size:11px;background:{"#fff" if i%2==0 else "#f9f9f9"}">
  <span style="color:#888;font-size:10px">{i+1}</span>
  <div><div style="font-weight:700">{row.Cust_Name}</div><div style="width:{min(100,round(row.rev/max(d["new_cl"]["rev"].max(),1)*100))}%;height:3px;background:{c};border-radius:2px;margin-top:3px"></div></div>
  <span style="text-align:right;font-weight:700">{fmtR(row.rev)}</span>
  <span style="text-align:right;color:#666">{row.inv}</span>
</div>''' for i,row in enumerate(d['new_cl'].rename(columns={'Cust Name':'Cust_Name'}).itertuples())])
 
    worst_rows = ''.join([f'''<div style="display:grid;grid-template-columns:22px 1fr 70px 70px 70px 62px;gap:8px;align-items:center;padding:6px 12px;border-bottom:1px solid #f0f0f0;font-size:11px;background:{"#fff" if i%2==0 else "#f9f9f9"}">
  <span style="color:#888;font-size:10px">{i+1}</span>
  <div><div style="font-weight:700">{row.Cust_Name}</div>{"" if row.rev>0 else "<span style=\"font-size:9px;background:#C0000022;color:#C00000;padding:1px 5px;border-radius:3px;display:inline-block;margin-top:2px\">Lost</span>"}</div>
  <span style="text-align:right;color:#666">{fmtR(row.py_rev)}</span>
  <span style="text-align:right">{"—" if row.rev==0 else fmtR(row.rev)}</span>
  <span style="text-align:right;font-weight:700;color:#C00000">({fmtR(abs(row.drop))})</span>
  <span style="text-align:right;font-weight:700;color:#C00000">({abs(round(row.drop/row.py_rev*100,1))}%)</span>
</div>''' for i,row in enumerate(d['worst_cl'].rename(columns={'Cust Name':'Cust_Name'}).itertuples())])
 
    sp_branch_rows = sp_rows(d['sp_branch'].to_dict(), d['sp26_b'].to_dict())
    sp_region_rows = sp_rows(d['sp_reg'].to_dict(), d['sp26_r'].to_dict())
    top_avg_rows   = avg_rows(d['top_avg'].rename(columns={'Cust Name': '_1','py_rev':'py_r'}), True)
    bot_avg_rows   = avg_rows(d['bot_avg'].rename(columns={'Cust Name': '_1','py_rev':'py_r'}), False)
 
    mar=m['Mar']; apr=m['Apr']; may_=m['May']; jun=m['Jun']
    ytd_act_d=d['ytd_act']-d['tgt_ytd']; ytd_pct_v=ytd_act_d/d['tgt_ytd']*100 if d['tgt_ytd'] else 0
    w1d=d['w1_act']-d['w1_tgt']; w2d=d['w2_act']-d['w2_tgt']
    mtd_act=d['jun_act']; mtd_tgt=d['w1_tgt']+d['w2_tgt']; mtd_d=mtd_act-mtd_tgt
    total_tgt=d['tgt_ytd']+mtd_tgt; total_d_=d['total_rev']-total_tgt
 
    def dc(v): return '#C00000' if v<0 else '#375623'
    def dfmt(v): return f"({fmtN(abs(v))})" if v<0 else f"+{fmtN(v)}"
    def dpct(v,t): p=v/t*100 if t else 0; return f"({abs(p):.2f}%)" if p<0 else f"+{p:.2f}%"
 
    sub_cols = f"grid-template-columns:repeat({min(len(d['sub_data']),5)},1fr)"
 
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>HW24 {r} Sales Dashboard — {report_date}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',Calibri,Arial,sans-serif;background:#f4f5f7;color:#222;padding:16px}}
.db{{max-width:1100px;margin:0 auto;background:#fff;border-radius:8px;box-shadow:0 2px 12px rgba(0,0,0,.10);overflow:hidden;padding-bottom:24px}}
.title-bar{{background:{c};color:#fff;padding:16px 24px;text-align:center;font-size:20px;font-weight:700;letter-spacing:.5px}}
.sub-bar{{display:flex;justify-content:space-between;padding:6px 20px;font-size:12px;color:#666;border-bottom:1px solid #e0e0e0;background:#f9f9f9}}
.kpi-row{{display:grid;grid-template-columns:repeat(5,1fr);border:1px solid #e0e0e0;margin:14px 14px 0}}
.kpi{{padding:10px 8px;text-align:center;border-right:1px solid #e0e0e0}}.kpi:last-child{{border-right:none}}
.kpi-label{{font-size:10px;font-weight:700;color:#fff;padding:5px 4px;margin:-10px -8px 8px;text-align:center;line-height:1.3}}
.kpi-val{{font-size:20px;font-weight:700;line-height:1}}
.section-hdr{{background:#1F3864;color:#fff;padding:8px 14px;font-size:12px;font-weight:700;margin:14px 14px 0;border-radius:4px}}
.tbl-hdr{{display:grid;background:#2E75B6;color:#fff;font-size:11px;font-weight:700;padding:7px 12px;gap:8px}}
.bold-row{{font-weight:700}}.italic-lbl{{font-style:italic;padding-left:16px;color:#666}}
.total-row{{background:#DCE6F1!important;font-weight:700}}
.avg-col{{color:#185FA5;font-weight:700}}
.chart-pair{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:14px 14px 0}}
.chart-box{{border:1px solid #e0e0e0;border-radius:6px;padding:12px}}
.chart-title{{font-size:11px;font-weight:700;text-align:center;color:#1F3864;margin-bottom:6px}}
.items-hdr{{display:grid;grid-template-columns:3fr 36px 62px 44px 34px;gap:3px;padding:5px 7px;background:#2E75B6;color:#fff;font-size:8.5px;font-weight:700;border-radius:4px 4px 0 0}}
.footer{{text-align:center;color:#aaa;font-size:10px;margin:20px 14px 0;padding-top:12px;border-top:1px solid #eee}}
.mtd-split{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:8px 14px 0;align-items:stretch}}
.loc-tbl-wrap{{display:flex;flex-direction:column}}
.loc-tbl-hdr{{display:grid;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;padding:7px 8px;gap:4px;border-radius:4px 4px 0 0}}
.loc-tbl-row{{display:grid;font-size:10px;padding:7px 8px;gap:4px;border-bottom:1px solid #f0f0f0}}
.loc-tbl-row:nth-child(even){{background:#f9f9f9}}
.loc-total{{background:#DCE6F1!important;font-weight:700}}
.mtd-chart-wrap{{display:flex;flex-direction:column;padding:10px;border:1px solid #e0e0e0;border-radius:6px;background:#fff}}
.mtd-chart-title{{font-size:12px;font-weight:700;text-align:center;color:#1F3864;margin-bottom:4px}}
.mtd-leg{{display:flex;gap:14px;justify-content:center;margin-bottom:6px;font-size:11px;color:#666}}
</style>
</head>
<body>
<div class="db">
<div class="title-bar">HW24 {r.upper()} SALES — DASHBOARD</div>
<div class="sub-bar"><span>{report_date}</span><span>March 2026 – June 2026 &nbsp;|&nbsp; MTD June W1+W2 &nbsp;|&nbsp; Targets = AFS26 ×1.19</span></div>
 
<div class="kpi-row">
  <div class="kpi"><div class="kpi-label" style="background:#2E75B6">Total Turnover</div><div class="kpi-val" style="color:#1F3864">R{d['total_rev']:,.0f}</div></div>
  <div class="kpi"><div class="kpi-label" style="background:#ED7D31">Units Sold</div><div class="kpi-val" style="color:#ED7D31">{d['total_units']:,.0f}</div></div>
  <div class="kpi"><div class="kpi-label" style="background:#375623">Avg Price / Unit</div><div class="kpi-val" style="color:#375623">R{d['avg_price']:.2f}</div></div>
  <div class="kpi"><div class="kpi-label" style="background:#843C0C">Active Clients</div><div class="kpi-val" style="color:#843C0C">{d['clients']:,}</div></div>
  <div class="kpi"><div class="kpi-label" style="background:#595959">Invoices</div><div class="kpi-val" style="color:#595959">{d['invoices']:,}</div></div>
</div>
 
<!-- Product class matrix -->
<div style="margin:10px 14px 0">
  <div style="display:grid;grid-template-columns:repeat(5,1fr);border:1px solid #e0e0e0">
    <div style="border:1px solid #e0e0e0"><div style="background:#2E75B6;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">LPCTO</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#2E75B6">{lpcto_rev:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{lpcto_avg:.2f}</div></div>
    <div style="border:1px solid #e0e0e0"><div style="background:#ED7D31;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">STPRO</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#ED7D31">{stpro_rev:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{stpro_avg:.2f}</div></div>
    <div style="border:1px solid #e0e0e0"><div style="background:#375623;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">LPMTO</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#375623">{lpmto_rev:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{lpmto_avg:.2f}</div></div>
    <div style="border:1px solid #e0e0e0"><div style="background:#843C0C;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">LPIMP</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#843C0C">{lpimp_rev:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{lpimp_avg:.2f}</div></div>
    <div style="border:1px solid #e0e0e0"><div style="background:#595959;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">UFLEX</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#595959">{uflex_rev:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{uflex_avg:.2f}</div></div>
  </div>
  <div style="display:flex;justify-content:center;border:1px solid #e0e0e0;border-top:none">
    <div style="flex:0 0 20%;border:1px solid #e0e0e0"><div style="background:#7B6E58;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">RAWMT</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#7B6E58">{rawmt_rev:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{rawmt_avg:.2f}</div></div>
    <div style="flex:0 0 20%;border:1px solid #e0e0e0"><div style="background:#888780;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">REPLEN</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#888780">{replen_rev:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{replen_avg:.2f}</div></div>
    <div style="flex:0 0 20%;border:1px solid #e0e0e0"><div style="background:#C00000;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">OPP</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#C00000">{opp_rev:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#C00000">Avg R{opp_avg:.2f}</div></div>
  </div>
</div>
 
 
<!-- Monthly Mar-Feb table -->
<div style="margin:10px 14px 0;overflow-x:auto">
  <table id="monthly-table" style="width:100%;border-collapse:collapse;font-size:10px">
    <thead>
      <tr style="background:#2E75B6;color:#fff">
        <td style="padding:6px 10px;font-weight:700;min-width:70px">Metric</td>
        <td style="padding:6px 8px;text-align:right;font-weight:700;white-space:nowrap">Mar</td>
        <td style="padding:6px 8px;text-align:right;font-weight:700;white-space:nowrap">Apr</td>
        <td style="padding:6px 8px;text-align:right;font-weight:700;white-space:nowrap">May</td>
        <td style="padding:6px 8px;text-align:right;font-weight:700;white-space:nowrap;background:{c}">Jun ▶</td>
        <td style="padding:6px 8px;text-align:right;font-weight:700;white-space:nowrap;opacity:.7">Jul</td>
        <td style="padding:6px 8px;text-align:right;font-weight:700;white-space:nowrap;opacity:.7">Aug</td>
        <td style="padding:6px 8px;text-align:right;font-weight:700;white-space:nowrap;opacity:.7">Sep</td>
        <td style="padding:6px 8px;text-align:right;font-weight:700;white-space:nowrap;opacity:.7">Oct</td>
        <td style="padding:6px 8px;text-align:right;font-weight:700;white-space:nowrap;opacity:.7">Nov</td>
        <td style="padding:6px 8px;text-align:right;font-weight:700;white-space:nowrap;opacity:.7">Dec</td>
        <td style="padding:6px 8px;text-align:right;font-weight:700;white-space:nowrap;opacity:.7">Jan</td>
        <td style="padding:6px 8px;text-align:right;font-weight:700;white-space:nowrap;opacity:.7">Feb</td>
        <td style="padding:6px 8px;text-align:right;font-weight:700;white-space:nowrap;background:#1F3864">TOTAL</td>
      </tr>
    </thead>
    <tbody id="monthly-table-body"></tbody>
  </table>
</div>
 
<div class="section-hdr" style="margin-top:14px">Period Performance</div>
<div style="margin:0 14px">
<div class="tbl-hdr" style="grid-template-columns:1.9fr 1.1fr 1.1fr 1fr .9fr 1fr"><span>Period</span><span style="text-align:right">Actual</span><span style="text-align:right">Target</span><span style="text-align:right">Delta</span><span style="text-align:right">% Delta</span><span style="text-align:right">Avg Unit Price</span></div>
{tbl_row("YTD Turnover (Mar–May)", d['ytd_act'], d['tgt_ytd'], ytd_avg_price, bold=True)}
{tbl_row("March",   mar['act'], mar['tgt'], mar['avg'],  indent=True)}
{tbl_row("April",   apr['act'], apr['tgt'], apr['avg'],  indent=True)}
{tbl_row("May",     may_['act'],may_['tgt'],may_['avg'], indent=True)}
{tbl_row("MTD June (W1+W2)", d['jun_act'], mtd_tgt, jun_avg_price, bold=True)}
{tbl_row("Week 1",  d['w1_act'], d['w1_tgt'], indent=True)}
{tbl_row("Week 2",  d['w2_act'], d['w2_tgt'], indent=True)}
{tbl_row("Total incl. June MTD", d['total_rev'], total_tgt, tot_avg_price, total=True)}
<div class="bold-row" style="display:grid;grid-template-columns:1.9fr 1.1fr 1.1fr 1fr .9fr 1fr;font-size:11px;padding:6px 12px;gap:8px;border-bottom:1px solid #f0f0f0;height:8px"></div>
{tbl_row("YTD Units (Mar–May)", ytd_units_act, ytd_units_tgt, bold=True)}
{tbl_row("MTD Units (June W1+W2)", jun_units_act, jun_units_tgt, indent=True)}
{tbl_row("Total Units incl. June", tot_units_act, tot_units_tgt, total=True)}
</div>
 
<div style="margin:6px 14px 0;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden" id="ins1-block"></div>
 
<div class="chart-pair">
  <div class="chart-box">
    <div class="chart-title">CUMULATIVE ACTUAL vs PY (MAR–JUN)</div>
    <div class="leg" style="display:flex;gap:10px;justify-content:center;margin-bottom:6px;font-size:9px;color:#666;flex-wrap:wrap">
      <span style="display:flex;align-items:center;gap:3px"><span style="display:inline-block;width:18px;height:2px;background:{c}"></span>Actual</span>
      <span style="display:flex;align-items:center;gap:3px"><span style="display:inline-block;width:18px;height:0;border-top:2px dashed #888780"></span>Prior Year</span>
      <span style="display:flex;align-items:center;gap:3px"><span style="display:inline-block;width:18px;height:0;border-top:1.5px dotted #888780"></span><span style="color:#888">Avg Price</span></span>
    </div>
    <div style="position:relative;height:190px"><canvas id="cPY"></canvas></div>
  </div>
  <div class="chart-box">
    <div class="chart-title">CUMULATIVE ACTUAL vs TARGET (MAR–JUN)</div>
    <div class="leg" style="display:flex;gap:10px;justify-content:center;margin-bottom:6px;font-size:9px;color:#666;flex-wrap:wrap">
      <span style="display:flex;align-items:center;gap:3px"><span style="display:inline-block;width:18px;height:0;border-top:2px dashed #C00000"></span>Target</span>
      <span style="display:flex;align-items:center;gap:3px"><span style="display:inline-block;width:18px;height:2px;background:{c}"></span>Actual</span>
      <span style="display:flex;align-items:center;gap:3px"><span style="display:inline-block;width:18px;height:0;border-top:1.5px dotted #888780"></span><span style="color:#888">Avg Price</span></span>
    </div>
    <div style="position:relative;height:190px"><canvas id="cYTD"></canvas></div>
  </div>
</div>
 
<div class="section-hdr" style="margin-top:14px">Sub-Region Breakdown — YTD (Mar–Jun)</div>
<div style="display:grid;{sub_cols};gap:10px;margin:8px 14px 0">{sub_html}</div>
 
<div style="margin:6px 14px 0;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden" id="ins2-block"></div>
 
<div class="section-hdr" style="margin-top:14px">National Performance by Region — MTD (June 2026 W1+W2) &nbsp;<span style="font-size:10px;font-weight:400;opacity:.8">W1 = 5 days (100%) = R7,370,437 &nbsp;|&nbsp; W2 = 2 days (40%) = R2,948,175 &nbsp;|&nbsp; Daily = R1,403,893</span></div>
<div class="mtd-split">
  <div class="loc-tbl-wrap" id="mtd-tbl-wrap">
    <div class="loc-tbl-hdr" style="grid-template-columns:1.6fr .95fr .95fr .8fr .7fr">
      <span>Province</span><span style="text-align:right">MTD Actual</span><span style="text-align:right">Jun Target</span><span style="text-align:right">Delta</span><span style="text-align:right">% Delta</span>
    </div>
    <div id="mtd-loc-tbl"></div>
  </div>
  <div class="mtd-chart-wrap" id="mtd-chart-wrap">
    <div class="mtd-chart-title">MTD ACTUAL VS TARGET REVENUE — JUNE 2026</div>
    <div class="mtd-leg">
      <span><span style="display:inline-block;width:12px;height:12px;background:#22A548;border-radius:2px;vertical-align:middle;margin-right:4px"></span>Actual R</span>
      <span><span style="display:inline-block;width:12px;height:12px;background:#8B0000;border-radius:2px;vertical-align:middle;margin-right:4px"></span>Weekly Target</span>
    </div>
    <div id="mtd-canvas-wrap" style="position:relative;flex:1;min-height:160px"><canvas id="cMTD"></canvas></div>
  </div>
</div>
 
<div class="section-hdr" style="margin-top:14px">Top 10 Clients — {r}</div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:22px 1fr 70px 70px 60px;gap:8px;align-items:center;padding:6px 12px;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0">
    <span>#</span><span>Client</span><span style="text-align:right">CY Sales</span><span style="text-align:right">PY YTD</span><span style="text-align:right">YoY %</span>
  </div>
  {top10_rows}
</div>
 
<div class="section-hdr" style="margin-top:14px">New Clients — {r} (no prior year)</div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:22px 1fr 70px 58px;gap:8px;align-items:center;padding:6px 12px;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0">
    <span>#</span><span>Client</span><span style="text-align:right">Revenue</span><span style="text-align:right">Invoices</span>
  </div>
  {new_cl_rows if new_cl_rows else '<div style="padding:10px;color:#888;font-size:10px">No new clients</div>'}
</div>
 
<div class="section-hdr" style="margin-top:14px">Worst Performing Clients — {r}</div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:22px 1fr 70px 70px 70px 62px;gap:8px;align-items:center;padding:6px 12px;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0">
    <span>#</span><span>Client</span><span style="text-align:right">AFS26</span><span style="text-align:right">AFS27</span><span style="text-align:right">Drop (R)</span><span style="text-align:right">%</span>
  </div>
  {worst_rows}
</div>
 
<div style="margin:6px 14px 0;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden" id="ins3-block"></div>
 
<div class="section-hdr" style="margin-top:14px">Top Avg Price — {r} &nbsp;<span style="font-size:10px;font-weight:400;opacity:.8">PY rev &gt; R20K</span></div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:20px 1fr 70px 70px 62px 62px 56px;gap:5px;align-items:center;padding:6px 10px;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0">
    <span>#</span><span>Client</span><span style="text-align:right">PY Rev</span><span style="text-align:right">CY Rev</span><span style="text-align:right">PY Avg</span><span style="text-align:right">CY Avg</span><span style="text-align:right">Change</span>
  </div>
  {top_avg_rows}
</div>
 
<div class="section-hdr" style="margin-top:14px">Bottom Avg Price — {r} &nbsp;<span style="font-size:10px;font-weight:400;opacity:.8">PY rev &gt; R20K</span></div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:20px 1fr 70px 70px 62px 62px 56px;gap:5px;align-items:center;padding:6px 10px;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0">
    <span>#</span><span>Client</span><span style="text-align:right">PY Rev</span><span style="text-align:right">CY Rev</span><span style="text-align:right">PY Avg</span><span style="text-align:right">CY Avg</span><span style="text-align:right">Change</span>
  </div>
  {bot_avg_rows}
</div>
 
<div class="section-hdr" style="margin-top:14px">YTD Turnover by Product Group &amp; Top/Bottom 8 Items (excl. LPMTO)</div>
<div style="display:grid;grid-template-columns:1fr 1.3fr 1.3fr;gap:8px;margin:8px 14px 0;align-items:stretch">
  <div style="border:1px solid #e0e0e0;border-radius:6px;padding:10px;display:flex;flex-direction:column">
    <div class="chart-title">YTD TURNOVER BY PRODUCT GROUP</div>
    <div id="hbar-wrap" style="position:relative;flex:1;min-height:200px"><canvas id="cHBar"></canvas></div>
  </div>
  <div style="display:flex;flex-direction:column;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden">
    <div style="font-size:10px;font-weight:700;color:#1F3864;padding:6px 10px 4px;background:#f9f9f9;border-bottom:1px solid #e0e0e0">Top 8 Items (excl. LPMTO)</div>
    <div class="items-hdr" style="border-radius:0"><span>Product</span><span style="text-align:center">Grp</span><span style="text-align:right">Turnover</span><span style="text-align:right">Avg</span><span style="text-align:right">%</span></div>
    <div style="flex:1">{items_rows}</div>
  </div>
  <div style="display:flex;flex-direction:column;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden">
    <div style="font-size:10px;font-weight:700;color:#1F3864;padding:6px 10px 4px;background:#f9f9f9;border-bottom:1px solid #e0e0e0">Bottom 8 Items (excl. LPMTO)</div>
    <div class="items-hdr" style="border-radius:0"><span>Product</span><span style="text-align:center">Grp</span><span style="text-align:right">Turnover</span><span style="text-align:right">Avg</span><span style="text-align:right">%</span></div>
    <div style="flex:1">{bot_rows}</div>
  </div>
</div>
 
<div class="section-hdr" style="margin-top:14px">Salespersons — Branch-Based ({r})</div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:24px 1fr 80px 70px 70px 60px;gap:6px;align-items:center;padding:6px 10px;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0">
    <span>#</span><span>Salesperson</span><span>Branch</span><span style="text-align:right">CY Sales</span><span style="text-align:right">PY YTD</span><span style="text-align:right">YoY %</span>
  </div>
  {sp_branch_rows if sp_branch_rows else '<div style="padding:10px;color:#888;font-size:10px">No branch-based salespersons for this region</div>'}
</div>
 
<div class="section-hdr" style="margin-top:14px">Salespersons — Sold to {r} Clients</div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:24px 1fr 80px 70px 70px 60px;gap:6px;align-items:center;padding:6px 10px;background:#843C0C;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0">
    <span>#</span><span>Salesperson</span><span>Branch</span><span style="text-align:right">CY Sales</span><span style="text-align:right">PY YTD</span><span style="text-align:right">YoY %</span>
  </div>
  {sp_region_rows}
</div>
 
<div class="footer">HW24 {r} Sales Dashboard &bull; Confidential &bull; Generated {report_date} &bull; Data: March–June 2026 (June W1+W2) &bull; Targets = AFS26 ×1.19</div>
</div>
 
<script>
const gridC='rgba(128,128,128,0.1)', REG_COL='{c}', GREY='#888780';
 
// ── Charts ────────────────────────────────────────────────────────────────────
const avgPrices = {avg_prices};
const y2Axis = {{
  position:'right', min:{y2_min}, max:{y2_max},
  ticks:{{font:{{size:9}},color:GREY,stepSize:5,callback:v=>'R'+v.toFixed(0)}},
  grid:{{display:false}},border:{{color:GREY}},
  title:{{display:true,text:'Avg Price (R)',font:{{size:9}},color:GREY}}
}};
const avgDs = {{
  label:'Avg Price (R)',data:avgPrices,type:'line',
  borderColor:GREY,borderWidth:1.5,borderDash:[2,3],
  pointRadius:3,pointStyle:'circle',pointBackgroundColor:GREY,
  tension:.3,fill:false,yAxisID:'y2'
}};
 
new Chart(document.getElementById('cPY'),{{type:'line',
  data:{{labels:['March','April','May','June (W1+W2)'],datasets:[
    {{label:'Actual',data:{d['cum_act']},borderColor:REG_COL,borderWidth:2,borderDash:[],pointRadius:4,pointStyle:'circle',pointBackgroundColor:REG_COL,tension:.3,fill:false,yAxisID:'y'}},
    {{label:'Prior Year',data:{d['cum_py']},borderColor:GREY,borderWidth:2,borderDash:[5,4],pointRadius:4,pointStyle:'triangle',pointBackgroundColor:GREY,tension:.3,fill:false,yAxisID:'y'}},
    {{...avgDs}}
  ]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},
    scales:{{
      x:{{ticks:{{font:{{size:9}}}},grid:{{display:false}}}},
      y:{{position:'left',ticks:{{font:{{size:9}},callback:v=>'R'+(v/1e6).toFixed(1)+'M'}},grid:{{color:gridC}}}},
      y2:{{...y2Axis}}
    }}}}
}});
 
new Chart(document.getElementById('cYTD'),{{type:'line',
  data:{{labels:['March','April','May','June (W1+W2)'],datasets:[
    {{label:'Target',data:{d['cum_tgt']},borderColor:'#C00000',borderWidth:2,borderDash:[5,4],pointRadius:4,pointStyle:'rectRot',pointBackgroundColor:'#C00000',tension:.3,fill:false,yAxisID:'y'}},
    {{label:'Actual',data:{d['cum_act']},borderColor:REG_COL,borderWidth:2,borderDash:[],pointRadius:4,pointStyle:'circle',pointBackgroundColor:REG_COL,tension:.3,fill:false,yAxisID:'y'}},
    {{...avgDs}}
  ]}},
  options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},
    scales:{{
      x:{{ticks:{{font:{{size:9}}}},grid:{{display:false}}}},
      y:{{position:'left',ticks:{{font:{{size:9}},callback:v=>'R'+(v/1e6).toFixed(1)+'M'}},grid:{{color:gridC}}}},
      y2:{{...y2Axis}}
    }}}}
}});
 
// ── Product group horizontal bar ──────────────────────────────────────────────
const vp={{id:'hv',afterDatasetsDraw(c){{const ctx=c.ctx;c.data.datasets.forEach((ds,di)=>{{c.getDatasetMeta(di).data.forEach((b,i)=>{{
  if(di===0){{const v=ds.data[i];ctx.save();ctx.font='700 9px Segoe UI,Arial';ctx.fillStyle='#333';
  ctx.textAlign='left';ctx.textBaseline='middle';ctx.fillText('R'+(v/1e6).toFixed(1)+'M',b.x+4,b.y);ctx.restore();}}}});}});}}}};
window.addEventListener('load',function(){{
  const iT=document.getElementById('items-tbl'),iH=document.querySelector('.items-hdr');
  document.getElementById('hbar-wrap').style.height=Math.max(200,300)+'px';
  new Chart(document.getElementById('cHBar'),{{type:'bar',plugins:[vp],
    data:{{labels:{json.dumps([k for k in ['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN','OPP']])},
      datasets:[
        {{label:'Actual',data:{json.dumps([round(d['prod'].get(k,{}).get('rev',0)) for k in ['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN','OPP']])},backgroundColor:'{c}',barPercentage:.45,categoryPercentage:.85}},
        {{label:'Target (PY+19%)',data:{json.dumps([d['prod'].get(k,{}).get('tgt',0) for k in ['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN','OPP']])},backgroundColor:'rgba(200,200,200,0.5)',borderColor:'#999',borderWidth:1,barPercentage:.45,categoryPercentage:.85}}
      ]}},
    options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:true,position:'bottom',labels:{{font:{{size:9}},boxWidth:10}}}},
               tooltip:{{callbacks:{{label:ctx=>'R'+Math.round(ctx.raw).toLocaleString()}}}}}},
      scales:{{x:{{min:0,ticks:{{font:{{size:9}},callback:v=>'R'+(v/1e6).toFixed(0)+'M'}},grid:{{color:'rgba(180,180,180,0.2)'}},border:{{display:false}}}},
               y:{{ticks:{{font:{{size:10,weight:'700'}},color:'#333'}},grid:{{display:false}},border:{{display:false}}}}}},
      layout:{{padding:{{right:42,top:2,bottom:2,left:2}}}}}}
  }});
}});
 
// ── Monthly table ─────────────────────────────────────────────────────────────
const monthActs  = {{Mar:{mar_act}, Apr:{apr_act}, May:{may_act}, Jun:{jun_act_v}, Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0,Jan:0,Feb:0}};
const monthTgts  = {{Mar:{mar_tgt}, Apr:{apr_tgt}, May:{may_tgt}, Jun:{jun_full_tgt}, Jul:{jul_tgt},Aug:{aug_tgt},Sep:{sep_tgt},Oct:{oct_tgt},Nov:{nov_tgt},Dec:{dec_tgt},Jan:{jan_tgt},Feb:{feb_tgt}}};
const mths_order = ['Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb'];
const totalAct   = mths_order.reduce((s,m)=>s+monthActs[m],0);
// YTD target: only sum months where we have actual data (excludes future months)
const totalTgt   = mths_order.reduce((s,m)=>s+(monthActs[m]>0?monthTgts[m]:0),0);
const totalDelta = totalAct - totalTgt;
 
const mbody = document.getElementById('monthly-table-body');
[['Actual (R)','act'],['Target (R)','tgt'],['Delta (R)','delta']].forEach(([lbl,key])=>{{
  const tr=document.createElement('tr'); tr.style.borderBottom='1px solid #f0f0f0';
  const td0=document.createElement('td');
  td0.style.cssText=`padding:5px 10px;font-weight:${{key==='delta'?'700':'400'}};white-space:nowrap;background:#f9f9f9`;
  td0.textContent=lbl; tr.appendChild(td0);
  mths_order.forEach(m=>{{
    const act=monthActs[m],tgt=monthTgts[m],delta=act-tgt;
    const isJun=m==='Jun',isFuture=act===0&&m!=='Mar'&&m!=='Apr'&&m!=='May'&&m!=='Jun';
    let v,col;
    if(key==='act')  {{v=act;  col='#1F3864';}}
    if(key==='tgt')  {{v=tgt;  col='#595959';}}
    if(key==='delta'){{v=isFuture?null:delta; col=delta<0?'#C00000':'#375623';}}
    const td=document.createElement('td');
    td.style.cssText=`padding:5px 8px;text-align:right;white-space:nowrap;font-weight:${{key==='delta'?'700':'400'}};${{isJun?'background:#f0f8f0':''}}`;
    if(v===null||(key==='act'&&isFuture)){{td.innerHTML='<span style="color:#ccc">—</span>';}}
    else{{const fmt='R'+(Math.abs(v)/1e6).toFixed(1)+'M';
      td.innerHTML=v<0?`<span style="color:#C00000">(${{fmt}})</span>`:`<span style="color:${{col}}">${{fmt}}</span>`;}}
    tr.appendChild(td);
  }});
  const tdTot=document.createElement('td');
  tdTot.style.cssText='padding:5px 8px;text-align:right;white-space:nowrap;font-weight:700;background:#DCE6F1';
  if(key==='act') tdTot.innerHTML=`<span style="color:#1F3864">R${{(totalAct/1e6).toFixed(1)}}M</span>`;
  else if(key==='tgt') tdTot.innerHTML=`<span style="color:#595959">R${{(totalTgt/1e6).toFixed(1)}}M</span>`;
  else{{const col=totalDelta<0?'#C00000':'#375623',fmt='R'+(Math.abs(totalDelta)/1e6).toFixed(1)+'M';
    tdTot.innerHTML=`<span style="color:${{col}}">${{totalDelta<0?'('+fmt+')':fmt}}</span>`;}}
  tr.appendChild(tdTot); mbody.appendChild(tr);
}});
 
// ── Executive summary ─────────────────────────────────────────────────────────
const ytdAct={ytd_act}, ytdTgt={tgt_ytd}, ytdDelta={ytd_delta};
const ytdPct=Math.abs(ytdDelta/ytdTgt*100).toFixed(1);
const junAct={jun_act_v}, junTgt={jun_mtd_tgt};
const junDelta=junAct-junTgt, junPct=Math.abs(junDelta/junTgt*100).toFixed(1);
const totalRev={total_rev};
const negCol='#C00000', posCol='#375623';
const ytdOk = ytdAct >= ytdTgt;
const junOk = junAct >= junTgt;
 
const insights = [
  {{
    num:1, col: ytdOk?posCol:negCol,
    title: ytdOk
      ? `YTD Revenue Ahead of Target — +R${{Math.abs(ytdDelta/1e6).toFixed(2)}}M`
      : `YTD Revenue Shortfall — R${{Math.abs(ytdDelta/1e6).toFixed(2)}}M Behind Target`,
    body: `Region tracking <strong>R${{(ytdAct/1e6).toFixed(2)}}M YTD (Mar–May)</strong> vs target <strong>R${{(ytdTgt/1e6).toFixed(2)}}M</strong> — ` +
      (ytdOk ? `<strong style="color:${{posCol}}">+${{ytdPct}}% ahead</strong>. Strong regional performance with consistent monthly delivery.`
             : `<strong style="color:${{negCol}}">(${{ytdPct}}%) behind</strong>. Requires focused recovery to close the gap before year-end.`),
    stats: [
      {{label:`YTD Actual R${{(ytdAct/1e6).toFixed(2)}}M`, neg:!ytdOk}},
      {{label:`Target R${{(ytdTgt/1e6).toFixed(2)}}M`, neg:false}},
      {{label:`Gap ${{ytdOk?'+':''}}R${{Math.abs(ytdDelta/1e6).toFixed(2)}}M (${{ytdOk?'+':'-'}}${{ytdPct}}%)`, neg:!ytdOk}}
    ]
  }},
  {{
    num:2, col: junOk?posCol:negCol,
    title: junOk
      ? `June MTD Tracking Ahead — W1+W2 above pro-rated target`
      : `June MTD Behind Pro-Rated Target — Action Required`,
    body: `June MTD (W1+W2) <strong>R${{(junAct/1e6).toFixed(2)}}M</strong> vs pro-rated target <strong>R${{(junTgt/1e6).toFixed(2)}}M</strong> (7/20 working days). ` +
      (junOk ? `Region is tracking well in June — maintain current sales momentum.`
             : `Currently <strong style="color:${{negCol}}">(${{junPct}}%) below</strong> the day-weighted June target. Weekly target R${{(monthTgts['Jun']/4/1e6).toFixed(2)}}M requires focus in remaining weeks.`),
    stats: [
      {{label:`W1 R${{({w1_act}/1e6).toFixed(2)}}M vs R${{({w1_tgt}/1e6).toFixed(2)}}M`,neg:{str(w1_act<w1_tgt).lower()}}},
      {{label:`W2 R${{({w2_act}/1e6).toFixed(2)}}M vs R${{({w2_tgt}/1e6).toFixed(2)}}M`,neg:{str(w2_act<w2_tgt).lower()}}}
    ]
  }},
  {{
    num:3, col:REG_COL,
    title:`Full-Year Target: R${{(totalTgt/1e6).toFixed(1)}}M — Implied Run Rate R${{((totalTgt-totalAct)/8/1e6).toFixed(2)}}M/month`,
    body:`To achieve the annual target of <strong>R${{(totalTgt/1e6).toFixed(1)}}M</strong>, the region needs to average <strong>R${{((totalTgt-totalAct)/8/1e6).toFixed(2)}}M per month</strong> across the remaining 8 months (Jul–Feb). Current run rate (Mar–Jun) is R${{(totalAct/4/1e6).toFixed(2)}}M/month. ` +
      (totalAct/4 >= (totalTgt-totalAct)/8 ? `Current pace is sufficient — maintain consistency.` : `<strong style="color:${{negCol}}">A step-up is required</strong> to meet annual target.`),
    stats:[{{label:`Annual Target R${{(totalTgt/1e6).toFixed(1)}}M`,neg:false}},{{label:`YTD+MTD R${{(totalRev/1e6).toFixed(1)}}M`,neg:false}},{{label:`Remaining R${{((totalTgt-totalRev)/1e6).toFixed(1)}}M`,neg:totalRev<totalTgt}}]
  }}
];
 
function renderInsight(elId, ins) {{
  const el = document.getElementById(elId);
  if (!el) return;
  const statsHtml = ins.stats.map(s=>`<span style="padding:3px 8px;border-radius:3px;font-size:9px;font-weight:700;background:${{s.neg?'#C0000015':'#2E75B615'}};color:${{s.neg?'#C00000':'#2E75B6'}}">${{s.label}}</span>`).join('');
  el.innerHTML = `<div style="display:grid;grid-template-columns:42px 1fr">
    <div style="background:${{ins.col}};display:flex;align-items:flex-start;justify-content:center;padding-top:14px">
      <span style="width:24px;height:24px;border-radius:50%;background:rgba(255,255,255,0.25);display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;color:#fff">${{ins.num}}</span>
    </div>
    <div style="padding:12px 14px;background:#fff">
      <div style="font-size:11px;font-weight:700;color:${{ins.col}};margin-bottom:4px">${{ins.title}}</div>
      <div style="font-size:10px;line-height:1.65;color:#444;margin-bottom:6px">${{ins.body}}</div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">${{statsHtml}}</div>
    </div>
  </div>`;
}}
renderInsight('ins1-block', insights[0]);
renderInsight('ins2-block', insights[1]);
renderInsight('ins3-block', insights[2]);
 
// ── National MTD by Region table + chart ─────────────────────────────────────
const THIS_REGION = '{r}';
const mtdRegions=[
  {{name:'Gauteng',      note:'incl. LP, NW, MP & FS', act:1971168, tgt:3674928}},
  {{name:'KZN',          note:'KZN only',               act:1510211, tgt:2731443}},
  {{name:'Western Cape', note:'incl. Northern Cape',    act:1568645, tgt:3344950}},
  {{name:'Eastern Cape', note:'EC only',                act:235863,  tgt:395707}},
  {{name:'International',note:'Export clients',         act:174263,  tgt:171584}},
];
const tA=mtdRegions.reduce((s,r)=>s+r.act,0);
const tT=mtdRegions.filter(r=>r.tgt>0).reduce((s,r)=>s+r.tgt,0);
const tD=tA-tT, tPct=tT?tD/tT:null;
 
document.getElementById('mtd-loc-tbl').innerHTML=mtdRegions.map((r,i)=>{{
  const isThisRegion = r.name === THIS_REGION;
  const d_=r.tgt?r.act-r.tgt:null, p_=r.tgt?d_/r.tgt:null;
  const isNeg=d_!==null&&d_<0, c_=isNeg?'#C00000':'#375623', nt=r.tgt===0;
  // Highlight this dashboard's region with coloured left border + light background
  const highlight = isThisRegion
    ? `border-left:4px solid {c};background:#f0f4fa!important;`
    : (i%2===0?'':'background:#f9f9f9');
  const nameBold = isThisRegion
    ? `<div style="display:flex;align-items:center;gap:5px"><span style="font-weight:700;color:{c}">${{r.name}}</span><span style="font-size:8px;background:{c};color:#fff;padding:1px 5px;border-radius:8px">THIS REGION</span></div>`
    : `<div style="font-weight:700">${{r.name}}</div>`;
  return`<div class="loc-tbl-row" style="grid-template-columns:1.6fr .95fr .95fr .8fr .7fr;${{highlight}}">
    <div>${{nameBold}}<div style="font-size:8px;color:#aaa">${{r.note}}</div></div>
    <span style="text-align:right;font-weight:${{isThisRegion?'700':'400'}}">R${{(r.act/1e6).toFixed(2)}}M</span>
    <span style="text-align:right;color:#888">${{nt?'—':'R'+(r.tgt/1e6).toFixed(2)+'M'}}</span>
    <span style="text-align:right;color:${{nt?'#888':c_}};font-weight:${{isThisRegion?'700':'400'}}">${{nt?'—':(isNeg?'(R'+(Math.abs(d_)/1e6).toFixed(2)+'M)':'+'+(d_/1e6).toFixed(2)+'M')}}</span>
    <span style="text-align:right;color:${{nt?'#888':c_}};font-weight:${{isThisRegion?'700':'400'}}">${{nt?'N/A':(isNeg?'('+Math.abs(Math.round(p_*100))+'%)':'+'+Math.round(p_*100)+'%')}}</span>
  </div>`;
}}).join('')+`<div class="loc-tbl-row loc-total" style="grid-template-columns:1.6fr .95fr .95fr .8fr .7fr">
  <span>TOTAL</span>
  <span style="text-align:right">R${{(tA/1e6).toFixed(2)}}M</span>
  <span style="text-align:right">R${{(tT/1e6).toFixed(2)}}M</span>
  <span style="text-align:right;color:${{tD<0?'#C00000':'#375623'}}">${{tD<0?'(R'+(Math.abs(tD)/1e6).toFixed(2)+'M)':'+'+(tD/1e6).toFixed(2)+'M'}}</span>
  <span style="text-align:right;color:${{tD<0?'#C00000':'#375623'}}">${{tPct!==null?(tD<0?'('+Math.abs(Math.round(tPct*100))+'%)':'+'+Math.round(tPct*100)+'%'):'—'}}</span>
</div>`;
 
window.addEventListener('load',function(){{
  const tw=document.getElementById('mtd-tbl-wrap'),cw=document.getElementById('mtd-chart-wrap');
  const cvw=document.getElementById('mtd-canvas-wrap');
  const tH=tw.offsetHeight; cw.style.minHeight=tH+'px';
  new Chart(document.getElementById('cMTD'),{{type:'bar',
    data:{{
      labels:['Week 1','Week 2'],
      datasets:[
        {{label:'Actual R',  data:[{w1_act},{w2_act}], backgroundColor:'#22A548',borderRadius:0,barPercentage:1.0,categoryPercentage:0.85}},
        {{label:'Weekly Target',data:[{w1_tgt},{w2_tgt}],backgroundColor:'#8B0000',borderRadius:0,barPercentage:1.0,categoryPercentage:0.85}}
      ]
    }},
    options:{{
      responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>'R'+Math.round(ctx.raw).toLocaleString()}}}}}},
      scales:{{
        x:{{ticks:{{font:{{size:11}},autoSkip:false}},grid:{{display:false}},border:{{display:false}}}},
        y:{{min:0,max:{chart_max_mtd},ticks:{{font:{{size:9}},callback:v=>'R'+(v/1e6).toFixed(2)+'M'}},grid:{{color:'rgba(180,180,180,0.25)'}},border:{{display:false}}}}
      }},
      layout:{{padding:{{top:2}}}}
    }}
  }});
}});
</script>
</body>
</html>"""
    return html
 
regions_cfg = [
    ('Gauteng','#185FA5'),('KZN','#ED7D31'),('Western Cape','#375623'),
    ('Eastern Cape','#843C0C'),('International','#444444'),
]
 
if __name__ == '__main__':
    output_dir = '/mnt/user-data/outputs'
    import os; os.makedirs(output_dir, exist_ok=True)
    files_created = []
    for region, col in regions_cfg:
        d = calc(region)
        d['col'] = col
        html = build_html(d)
        safe_name = region.replace(' ','_')
        path = f"{output_dir}/HW24_{safe_name}_Dashboard.html"
        with open(path,'w',encoding='utf-8') as f:
            f.write(html)
        files_created.append(path)
        print(f"✓ {region}: {len(html):,} bytes → {path}")
    print(f"\nAll {len(files_created)} regional dashboards created.")
 
 
# ══════════════════════════════════════════════════════════════════════════════
# NATIONAL DASHBOARD GENERATOR
# ══════════════════════════════════════════════════════════════════════════════
 
def generate_national(report_date="12 Jun 2026", week_label="Week 2", days_elapsed=4):
    """Generate the complete national dashboard HTML. Call initialise() first."""
    import pandas as pd
 
    # ── Data references from global state ─────────────────────────────────────
    r27 = afs27.copy(); r26 = afs26.copy()
    r27_ytd = r27[r27['Period1'].isin(['Mar','Apr','May'])]
    r27_jun = r27[r27['Period1']=='Jun']
    r26_ytd = r26[r26['Period1'].isin(['Mar','Apr','May'])]
    r26_jun = r26[r26['Period1']=='Jun']
 
    w2_frac = days_elapsed / 5
    wd_el   = (int(week_label.split()[1]) - 1) * 5 + days_elapsed
    wd_tot  = 20
    pro_rat = wd_el / wd_tot
 
    # ── National KPIs ──────────────────────────────────────────────────────────
    total_rev   = float(r27['Line Revenue'].sum())
    total_units = float(r27['Line Inv Qty'].sum())
    avg_price   = total_rev / total_units if total_units else 0
    clients     = int(r27['Cust Name'].nunique())
    invoices    = int(r27['Inv no'].nunique())
 
    # ── Monthly ────────────────────────────────────────────────────────────────
    def mrev(df, p): return float(df[df['Period1']==p]['Line Revenue'].sum())
    def munit(df,p): return float(df[df['Period1']==p]['Line Inv Qty'].sum())
    def mavg(df, p):
        r=mrev(df,p); u=munit(df,p); return r/u if u else 0
 
    mar_act=mrev(r27,'Mar'); apr_act=mrev(r27,'Apr'); may_act=mrev(r27,'May'); jun_act=mrev(r27,'Jun')
    mar_py =mrev(r26,'Mar'); apr_py =mrev(r26,'Apr'); may_py =mrev(r26,'May'); jun_py =mrev(r26,'Jun')
    mar_avg=mavg(r27,'Mar'); apr_avg=mavg(r27,'Apr'); may_avg=mavg(r27,'May'); jun_avg=mavg(r27,'Jun')
 
    ytd_act = mar_act+apr_act+may_act
    ytd_py  = mar_py+apr_py+may_py
 
    # ── Targets ────────────────────────────────────────────────────────────────
    mar_tgt=round(mar_py*UPLIFT); apr_tgt=round(apr_py*UPLIFT); may_tgt=round(may_py*UPLIFT)
    jun_tgt_full=round(jun_py*UPLIFT)
    weekly_tgt=round(jun_tgt_full/4)
    w1_tgt=weekly_tgt; w2_tgt=round(weekly_tgt*w2_frac)
    jun_tgt_pro=w1_tgt+w2_tgt
    ytd_tgt=mar_tgt+apr_tgt+may_tgt
 
    w1_act=float(r27_jun[r27_jun['Week']=='Week 1']['Line Revenue'].sum())
    w2_act=float(r27_jun[r27_jun['Week']=='Week 2']['Line Revenue'].sum())
 
    # Remaining monthly targets
    def rm(p): return round(mrev(r26,p)*UPLIFT)
    jul_tgt=rm('Jul'); aug_tgt=rm('Aug'); sep_tgt=rm('Sept'); oct_tgt=rm('Oct')
    nov_tgt=rm('Nov'); dec_tgt=rm('Dec'); jan_tgt=rm('Jan'); feb_tgt=rm('Feb')
 
    # ── Units ──────────────────────────────────────────────────────────────────
    ytd_u_act=int(munit(r27_ytd.assign(**{}),'Mar')+munit(r27_ytd,'Apr')+munit(r27_ytd,'May'))
    ytd_u_act=int(r27_ytd['Line Inv Qty'].sum())
    ytd_u_py =int(r26_ytd['Line Inv Qty'].sum())
    ytd_u_tgt=round(ytd_u_py*UPLIFT)
    jun_u_act=int(r27_jun['Line Inv Qty'].sum())
    jun_u_py =int(r26_jun['Line Inv Qty'].sum())
    jun_u_tgt=round(jun_u_py*UPLIFT*pro_rat)
    tot_u_tgt=ytd_u_tgt+jun_u_tgt
 
    # ── Product class ──────────────────────────────────────────────────────────
    def pdata(cls):
        cd=r27[r27['CLASS']==cls]; rev=float(cd['Line Revenue'].sum()); qty=float(cd['Line Inv Qty'].sum())
        py_cd=float(r26_ytd[r26_ytd['CLASS']==cls]['Line Revenue'].sum())+float(r26_jun[r26_jun['CLASS']==cls]['Line Revenue'].sum())*pro_rat
        return {'rev':rev,'avg':rev/qty if qty else 0,'tgt':round(py_cd*UPLIFT)}
 
    classes=['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN ','OPP']
    prod={c.strip(): pdata(c) for c in classes}
 
    # ── Top 8 items ────────────────────────────────────────────────────────────
    ti=r27[r27['CLASS']!='LPMTO'].groupby(['Item No (Stock)','CLASS']).agg(rev=('Line Revenue','sum'),qty=('Line Inv Qty','sum')).reset_index()
    ti['avg']=ti['rev']/ti['qty']; ti['pct']=ti['rev']/total_rev*100
    top8=ti.nlargest(8,'rev').to_dict('records')
    bot8=ti[(ti['rev']>=10000)&(ti['qty']>=50)].nsmallest(8,'rev').to_dict('records')
 
    # ── Clients ────────────────────────────────────────────────────────────────
    cp26=r26_ytd.groupby('Cust Name')['Line Revenue'].sum()
    pmap=r27.groupby('Cust Name')[prov_col].agg(lambda x: x.mode()[0])
    top10=r27.groupby('Cust Name')['Line Revenue'].sum().nlargest(10)
 
    rev27all=r27.groupby('Cust Name').agg(rev=('Line Revenue','sum'),inv=('Inv no','nunique'),prov=(prov_col,lambda x:x.mode()[0])).reset_index()
    mg=rev27all.merge(cp26.rename('py_rev'),on='Cust Name',how='left'); mg['py_rev']=mg['py_rev'].fillna(0)
    new_cl=mg[mg['py_rev']==0].nlargest(10,'rev')
    both=mg[mg['py_rev']>0].copy(); both['drop']=both['rev']-both['py_rev']
    worst=both.nsmallest(10,'drop')
 
    # ── Avg price ─────────────────────────────────────────────────────────────
    py_s=r26_ytd[r26_ytd['Line Revenue']>0].groupby('Cust Name').agg(py_rev=('Line Revenue','sum'),py_qty=('Line Inv Qty','sum'),prov=(prov_col,lambda x:x.mode()[0])).reset_index()
    pt=r26_ytd.groupby('Cust Name')['Line Revenue'].sum().reset_index(); pt.columns=['Cust Name','py_total']
    py_s=py_s.merge(pt,on='Cust Name'); py_s=py_s[py_s['py_total']>=50000].copy(); py_s['py_avg']=py_s['py_rev']/py_s['py_qty']
    cy_s=r27[r27['Line Revenue']>0].groupby('Cust Name').agg(cy_rev=('Line Revenue','sum'),cy_qty=('Line Inv Qty','sum')).reset_index()
    cy_s['cy_avg']=cy_s['cy_rev']/cy_s['cy_qty']
    avg_m=py_s.merge(cy_s,on='Cust Name',how='inner'); valid=avg_m[avg_m['cy_avg']>=5.0]
    top_avg=valid.nlargest(10,'cy_avg'); bot_avg=valid.nsmallest(10,'cy_avg')
 
    # ── Region ─────────────────────────────────────────────────────────────────
    r27_all=pd.concat([r27_ytd, r27_jun]); r26_all=pd.concat([r26_ytd, r26_jun])
    PROV_FULL={'GP':'Gauteng','KZN':'KZN','WC':'Western Cape','EC':'Eastern Cape','FS':'Free State',
               'MP':'Mpumalanga','LP':'Limpopo','NW':'North West','NC':'Northern Cape','ZZZ':'International'}
    REG_COLS={'Gauteng':'#185FA5','KZN':'#ED7D31','Western Cape':'#375623','Eastern Cape':'#843C0C','International':'#444444'}
 
    reg_data={}
    for reg in ['Gauteng','KZN','Western Cape','Eastern Cape','International']:
        cy=float(r27_all[r27_all['Region']==reg]['Line Revenue'].sum())
        py_y=float(r26_ytd[r26_ytd['Region']==reg]['Line Revenue'].sum())
        py_j=float(r26_jun[r26_jun['Region']==reg]['Line Revenue'].sum())
        tgt=round((py_y+py_j*pro_rat)*UPLIFT)
        mtd_act=float(r27_jun[r27_jun['Region']==reg]['Line Revenue'].sum())
        mtd_tgt=round(py_j*UPLIFT*pro_rat)
        reg_data[reg]={'cy':cy,'tgt':tgt,'mtd_act':mtd_act,'mtd_tgt':mtd_tgt}
 
    SUB_MAP={'GP':('Gauteng core','Gauteng'),'LP':('Limpopo','Gauteng'),'NW':('North West','Gauteng'),
             'MP':('Mpumalanga','Gauteng'),'FS':('Free State','Gauteng'),
             'KZN':('KZN core','KZN'),'WC':('WC core','Western Cape'),
             'NC':('Northern Cape','Western Cape'),'EC':('EC core','Eastern Cape')}
    sub_data={}
    for p,(name,reg) in SUB_MAP.items():
        cy=float(r27_all[r27_all[prov_col]==p]['Line Revenue'].sum())
        py=float(r26_ytd[r26_ytd[prov_col]==p]['Line Revenue'].sum())+float(r26_jun[r26_jun[prov_col]==p]['Line Revenue'].sum())*pro_rat
        sub_data[p]={'name':name,'v':cy,'t':round(py*UPLIFT)}
 
    # Top 3 intl countries
    CMAP={'SZ':'Swaziland','BW':'Botswana','MZ':'Mozambique','ZW':'Zimbabwe','ZM':'Zambia','NA':'Namibia'}
    top3_intl=r27_all[r27_all[prov_col]=='ZZZ'].groupby('Country')['Line Revenue'].sum().nlargest(3)
 
    # ── Salesperson ────────────────────────────────────────────────────────────
    excl=['WEBSITE','COMPANY','DIRECT','']
    sp27=r27[~r27[sp_col].isin(excl)].groupby(sp_col)['Line Revenue'].sum().nlargest(10)
    sp26y=r26_ytd.groupby(sp_col)['Line Revenue'].sum()
    top_sp=[(n,float(v),float(sp26y.get(n,0)),branch_dict.get(n,'—')) for n,v in sp27.items()]
 
    sp27a=r27[~r27[sp_col].isin(excl)].groupby(sp_col)['Line Revenue'].sum().reset_index()
    sp27a['py']=sp27a[sp_col].map(sp26y).fillna(0)
    bot_sp=[(r[sp_col],float(r['Line Revenue']),float(r['py']),branch_dict.get(r[sp_col],'—'))
            for _,r in sp27a[sp27a['py']>=100000].nsmallest(5,'Line Revenue').iterrows()]
 
    # ── Cumulative chart data ──────────────────────────────────────────────────
    cum_act=[int(mar_act), int(mar_act+apr_act), int(mar_act+apr_act+may_act), int(total_rev)]
    cum_tgt=[mar_tgt, mar_tgt+apr_tgt, mar_tgt+apr_tgt+may_tgt, mar_tgt+apr_tgt+may_tgt+round(jun_py*UPLIFT*pro_rat)]
    cum_py =[int(mar_py), int(mar_py+apr_py), int(mar_py+apr_py+may_py), int(mar_py+apr_py+may_py+jun_py*pro_rat)]
 
    # ── Helper formatters ──────────────────────────────────────────────────────
    def fmtN(v): return f"{int(v):,}"
    def fmtM(v): return f"R{v/1e6:.2f}M"
    def fmtK(v): return f"R{v/1e6:.2f}M" if v>=1e6 else f"R{v/1000:.1f}K"
    def dc(a,t): return '#C00000' if a<t else '#375623'
    def dv(a,t): d=a-t; return f"({fmtN(abs(d))})" if d<0 else f"+{fmtN(d)}"
    def dp(a,t): p=(a-t)/t*100 if t else 0; return f"({abs(p):.1f}%)" if p<0 else f"+{p:.1f}%"
 
    GCOLORS={'LPCTO':'#2E75B6','STPRO':'#ED7D31','LPMTO':'#375623','LPIMP':'#843C0C',
             'UFLEX':'#595959','RAWMT':'#7B6E58','REPLEN':'#888780','OPP':'#C00000'}
 
    def tbl_row(label, act, tgt, avg=None, indent=False, bold=False, total=False):
        d_=act-tgt; p_=d_/tgt*100 if tgt else 0; clr=dc(act,tgt)
        bg='background:#DCE6F1;' if total else ''
        fw='font-weight:700;' if (bold or total) else ''
        lbl=f'<span style="font-style:italic;padding-left:16px;color:#666">{label}</span>' if indent else f'<span>{label}</span>'
        avgc=f'<span style="text-align:right;color:#185FA5;font-weight:700">R{avg:.2f}</span>' if avg else '<span style="text-align:right;color:#888">—</span>'
        return f'''<div style="display:grid;grid-template-columns:1.9fr 1.1fr 1.1fr 1fr .9fr 1fr;{fw}{bg}font-size:11px;padding:6px 12px;gap:8px;border-bottom:1px solid #f0f0f0">
      {lbl}
      <span style="text-align:right">{fmtN(act)}</span>
      <span style="text-align:right">{fmtN(tgt)}</span>
      <span style="text-align:right;color:{clr}">{dv(act,tgt)}</span>
      <span style="text-align:right;color:{clr}">{dp(act,tgt)}</span>
      {avgc}
    </div>'''
 
    def sp_rows(sp_list, max_=10):
        if not sp_list: return '<div style="padding:10px;color:#888;font-size:10px">No data</div>'
        mx=max(max(r[1] for r in sp_list), max(r[2] for r in sp_list if r[2]>0), 1)
        rows=[]
        for i,(name,rev,py,branch) in enumerate(sp_list[:max_]):
            yoy=(rev-py)/py*100 if py else None
            ys='<span style="background:#375623;color:#fff;font-size:8px;padding:1px 4px;border-radius:3px">NEW</span>' if not py else (
               f'<span style="color:#375623;font-weight:700">+{yoy:.1f}%</span>' if yoy>=0 else
               f'<span style="color:#C00000;font-weight:700">({abs(yoy):.1f}%)</span>')
            bg='#fff' if i%2==0 else '#f9f9f9'
            cy_w=round(rev/mx*100); py_w=round(py/mx*100) if py else 0
            branch_full={'DURBAN':'KZN','CAPE TOWN':'WC','JOHANNESBURG':'GP','PORT ELIZABETH':'EC'}.get(branch,'')
            rows.append(f'''<div style="display:grid;grid-template-columns:24px 1fr 130px 72px 72px 62px;gap:6px;align-items:center;padding:7px 10px;border-bottom:1px solid #f0f0f0;font-size:11px;background:{bg}">
      <span style="color:#888;font-size:10px">{i+1}</span>
      <div><div style="font-weight:700">{name}</div>
        <div style="display:flex;flex-direction:column;gap:2px;margin-top:3px">
          <div style="display:flex;align-items:center;gap:3px"><span style="font-size:8px;color:#888;width:16px">CY</span><div style="flex:1;height:4px;background:#e0e0e0;border-radius:2px;overflow:hidden"><div style="width:{cy_w}%;height:4px;background:#2E75B6"></div></div><span style="font-size:8px;color:#666;min-width:36px;text-align:right">{fmtK(rev)}</span></div>
          {f'<div style="display:flex;align-items:center;gap:3px"><span style="font-size:8px;color:#888;width:16px">PY</span><div style="flex:1;height:4px;background:#e0e0e0;border-radius:2px;overflow:hidden"><div style="width:{py_w}%;height:4px;background:#A9C4E4"></div></div><span style="font-size:8px;color:#666;min-width:36px;text-align:right">{fmtK(py)}</span></div>' if py else ''}
        </div>
      </div>
      <div><div style="font-size:10px;font-weight:700;color:#444">{branch}</div><div style="font-size:9px;color:#888">{branch_full}</div></div>
      <span style="text-align:right;font-weight:700">{fmtK(rev)}</span>
      <span style="text-align:right;color:#666">{fmtK(py) if py else "—"}</span>
      <span style="text-align:right">{ys}</span>
    </div>''')
        return ''.join(rows)
 
    def client_rows(clients_series, cp26_series, max_=10):
        mx=max(clients_series.values) if len(clients_series) else 1
        rows=[]
        for i,(name,rev) in enumerate(list(clients_series.items())[:max_]):
            py=float(cp26_series.get(name,0)); prov=pmap.get(name,'?')
            yoy=(rev-py)/py*100 if py else None
            ys='<span style="background:#375623;color:#fff;font-size:8px;padding:1px 4px;border-radius:3px">NEW</span>' if not py else (
               f'<span style="color:#375623;font-weight:700">+{yoy:.1f}%</span>' if yoy>=0 else
               f'<span style="color:#C00000;font-weight:700">({abs(yoy):.1f}%)</span>')
            w=round(rev/mx*100); bg='#fff' if i%2==0 else '#f9f9f9'
            rows.append(f'''<div style="display:grid;grid-template-columns:22px 1fr 110px 72px 72px 68px;gap:8px;align-items:center;padding:6px 12px;border-bottom:1px solid #f0f0f0;font-size:11px;background:{bg}">
      <span style="color:#888;font-size:10px">{i+1}</span>
      <div><div style="font-weight:700">{name}</div><div style="width:{w}%;height:3px;background:#2E75B6;border-radius:2px;margin-top:3px"></div></div>
      <span style="font-size:10px;color:#444">{PROV_FULL.get(prov,prov)}</span>
      <span style="text-align:right;font-weight:700">{fmtK(rev)}</span>
      <span style="text-align:right;color:#666">{'—' if not py else fmtK(py)}</span>
      <span style="text-align:right">{ys}</span>
    </div>''')
        return ''.join(rows)
 
    def items_rows(items, bg_even='#fff'):
        rows=[]
        for i,it in enumerate(items):
            cls=it["CLASS"].strip(); bg=bg_even if i%2==0 else '#f9f9f9'
            rows.append(f'''<div style="display:grid;grid-template-columns:2.5fr 36px 62px 44px 34px;gap:3px;padding:4px 7px;border-bottom:1px solid #f0f0f0;font-size:8.5px;background:{bg}">
      <span style="font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{it["Item No (Stock)"]}</span>
      <span style="text-align:center"><span style="font-size:7.5px;padding:1px 2px;border-radius:8px;background:{GCOLORS.get(cls,"#888")}22;color:{GCOLORS.get(cls,"#555")};font-weight:700">{cls}</span></span>
      <span style="text-align:right">{int(it["rev"]):,}</span>
      <span style="text-align:right;color:#666">{it["avg"]:.2f}</span>
      <span style="text-align:right;color:#666">{it["pct"]:.2f}%</span>
    </div>''')
        return ''.join(rows)
 
    def avg_rows(df_a, is_top, max_=10):
        if len(df_a)==0: return '<div style="padding:10px;color:#888">Insufficient data</div>'
        mx=max(df_a["cy_avg"].max(), df_a["py_avg"].max())
        rows=[]
        for i,row in enumerate(df_a.head(max_).itertuples()):
            chg=row.cy_avg-row.py_avg; cy_col='#375623' if is_top else '#C00000'; chg_col='#375623' if chg>=0 else '#C00000'
            bg='#fff' if i%2==0 else '#f9f9f9'
            py_w=round(row.py_avg/mx*100); cy_w=round(row.cy_avg/mx*100)
            cn = getattr(row, '_1', getattr(row, 'Cust Name', '?'))
            rows.append(f'''<div style="display:grid;grid-template-columns:20px 1fr 110px 78px 78px 62px 62px 62px;gap:6px;align-items:center;padding:6px 10px;border-bottom:1px solid #f0f0f0;font-size:10px;background:{bg}">
      <span style="color:#888">{i+1}</span>
      <div><div style="font-weight:700;font-size:10px">{cn}</div>
        <div style="display:flex;flex-direction:column;gap:2px;margin-top:2px">
          <div style="display:flex;align-items:center;gap:3px"><span style="font-size:7px;color:#888;width:16px">PY</span><div style="flex:1;height:3px;background:#e0e0e0;border-radius:2px;overflow:hidden"><div style="width:{py_w}%;height:3px;background:#A9C4E4"></div></div></div>
          <div style="display:flex;align-items:center;gap:3px"><span style="font-size:7px;color:#888;width:16px">CY</span><div style="flex:1;height:3px;background:#e0e0e0;border-radius:2px;overflow:hidden"><div style="width:{cy_w}%;height:3px;background:{cy_col}"></div></div></div>
        </div>
      </div>
      <span style="font-size:10px;color:#444">{PROV_FULL.get(row.prov,row.prov)}</span>
      <span style="text-align:right;color:#666">R{row.py_rev:.0f}</span>
      <span style="text-align:right">R{row.cy_rev:.0f}</span>
      <span style="text-align:right;color:#888">R{row.py_avg:.2f}</span>
      <span style="text-align:right;font-weight:700;color:{cy_col}">R{row.cy_avg:.2f}</span>
      <span style="text-align:right;font-weight:700;color:{chg_col}">{"+" if chg>=0 else ""}{chg:.2f}</span>
    </div>''')
        return ''.join(rows)
 
    # ── Region YTD cards ───────────────────────────────────────────────────────
    def region_cards():
        cards=[]
        sub_groups={'Gauteng':['GP','FS','MP','NW','LP'],'KZN':['KZN'],
                    'Western Cape':['WC','NC'],'Eastern Cape':['EC'],'International':[]}
        intl_countries=[(CMAP.get(c,c),float(v)) for c,v in top3_intl.items()]
        for reg in ['Gauteng','KZN','Western Cape','Eastern Cape','International']:
            d=reg_data[reg]; gap=d['cy']-d['tgt']; pct_=(d['cy']/d['tgt']*100) if d['tgt'] else 100
            isNeg=gap<0; col_='#C00000' if isNeg else '#375623'
            prog=min(100,round(d['cy']/d['tgt']*100)) if d['tgt'] else 100
            pc='#C00000' if isNeg and pct_<70 else ('#ED7D31' if isNeg else '#375623')
            gstr=f"Gap: ({fmtM(abs(gap))}) — {abs(100-pct_):.1f}% behind" if isNeg and d['tgt'] else (f"Ahead: +{fmtM(gap)}" if d['tgt'] else "No target — bonus revenue")
            subs=sub_groups.get(reg,[])
            sub_html=''
            if reg=='International':
                mx=max([v for _,v in intl_countries],default=1)
                for cname,cv in intl_countries:
                    w=round(cv/mx*100)
                    sub_html+=f'''<div style="display:flex;align-items:center;gap:5px;margin-top:3px">
                      <span style="font-size:8px;color:#666;min-width:90px">{cname}</span>
                      <div style="flex:1;height:4px;background:#e8e8e8;border-radius:2px;overflow:hidden"><div style="width:{w}%;height:4px;background:#2E75B6"></div></div>
                      <span style="font-size:8px;color:#444;min-width:46px;text-align:right">{fmtK(cv)}</span></div>'''
            else:
                sub_vals=[sub_data.get(p,{'v':0,'t':1}) for p in subs]
                mx=max([max(s['v'],s['t']) for s in sub_vals],default=1)
                for p in subs:
                    sd=sub_data.get(p); 
                    if not sd or sd['v']==0: continue
                    pct2=min(100,round(sd['v']/sd['t']*100)) if sd['t']>0 else 100
                    ahead2=sd['v']>=sd['t'] if sd['t']>0 else True
                    bc='#375623' if ahead2 else '#2E75B6'
                    sub_html+=f'''<div style="display:flex;align-items:center;gap:5px;margin-top:3px">
                      <span style="font-size:8px;color:#666;min-width:90px">{sd["name"]}</span>
                      <div style="flex:1;height:4px;background:#e8e8e8;border-radius:2px;overflow:hidden"><div style="width:{pct2}%;height:4px;background:{bc}"></div></div>
                      <span style="font-size:8px;color:#444;min-width:46px;text-align:right">{fmtK(sd["v"])}</span></div>'''
            cards.append(f'''<div style="border:1px solid #e0e0e0;border-radius:6px;padding:10px 12px;background:#fff">
      <div style="font-size:12px;font-weight:700;margin-bottom:2px">{reg}</div>
      <div style="font-size:16px;font-weight:700;color:{col_};margin-bottom:2px">{fmtM(d["cy"])}</div>
      <div style="font-size:10px;color:{col_};margin-bottom:4px">{gstr}</div>
      <div style="height:4px;background:#ddd;border-radius:2px;margin-bottom:4px"><div style="width:{prog}%;height:4px;background:{pc};border-radius:2px"></div></div>
      <div style="font-size:9px;color:#888;margin-bottom:5px">{"R"+(str(round(d["cy"]/1e6,1)))+"M of R"+(str(round(d["tgt"]/1e6,1)))+"M" if d["tgt"] else "Top 3 export markets"}</div>
      {sub_html}
    </div>''')
        return ''.join(cards)
 
    # ── MTD region table rows ──────────────────────────────────────────────────
    def mtd_rows():
        regions=[
            {'name':'Gauteng','note':'incl. LP, NW, MP & FS'},
            {'name':'KZN','note':'KZN only'},
            {'name':'Western Cape','note':'incl. Northern Cape'},
            {'name':'Eastern Cape','note':'EC only'},
            {'name':'International','note':'Export clients'},
        ]
        tA=sum(reg_data[r['name']]['mtd_act'] for r in regions)
        tT=sum(reg_data[r['name']]['mtd_tgt'] for r in regions)
        tD=tA-tT; tP=tD/tT*100 if tT else 0
        html=''
        for i,r in enumerate(regions):
            d=reg_data[r['name']]; act=d['mtd_act']; tgt=d['mtd_tgt']
            dd=act-tgt; pp=dd/tgt*100 if tgt else None; isNeg=dd<0 if tgt else False
            c='#C00000' if isNeg else '#375623'; nt=tgt==0
            bg='' if i%2==0 else 'background:#f9f9f9'
            html+=f'''<div style="display:grid;grid-template-columns:1.6fr .95fr .95fr .8fr .7fr;font-size:10px;padding:7px 8px;gap:4px;border-bottom:1px solid #f0f0f0;{bg}">
        <div><span style="font-weight:700">{r["name"]}</span><div style="font-size:8px;color:#aaa">{r["note"]}</div></div>
        <span style="text-align:right">{fmtK(act)}</span>
        <span style="text-align:right;color:#888">{"—" if nt else fmtK(tgt)}</span>
        <span style="text-align:right;color:{"#888" if nt else c}">{"—" if nt else ("("+fmtK(abs(dd))+")" if isNeg else "+"+fmtK(dd))}</span>
        <span style="text-align:right;color:{"#888" if nt else c}">{"N/A" if nt else (f"({abs(pp):.0f}%)" if isNeg else f"+{pp:.0f}%")}</span>
      </div>'''
        html+=f'''<div style="display:grid;grid-template-columns:1.6fr .95fr .95fr .8fr .7fr;font-size:10px;padding:7px 8px;gap:4px;background:#DCE6F1;font-weight:700">
        <span>TOTAL</span>
        <span style="text-align:right">{fmtK(tA)}</span>
        <span style="text-align:right">{fmtK(tT)}</span>
        <span style="text-align:right;color:{"#C00000" if tD<0 else "#375623"}">{"("+fmtK(abs(tD))+")" if tD<0 else "+"+fmtK(tD)}</span>
        <span style="text-align:right;color:{"#C00000" if tD<0 else "#375623"}">{"("+f"{abs(tP):.0f}%)" if tD<0 else f"+{tP:.0f}%"}</span>
      </div>'''
        return html
 
    # ── Monthly table JS data ──────────────────────────────────────────────────
    mth_acts=f"Mar:{int(mar_act)},Apr:{int(apr_act)},May:{int(may_act)},Jun:{int(jun_act)},Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0,Jan:0,Feb:0"
    mth_tgts=f"Mar:{mar_tgt},Apr:{apr_tgt},May:{may_tgt},Jun:{jun_tgt_full},Jul:{jul_tgt},Aug:{aug_tgt},Sep:{sep_tgt},Oct:{oct_tgt},Nov:{nov_tgt},Dec:{dec_tgt},Jan:{jan_tgt},Feb:{feb_tgt}"
 
    GREY='#888780'; AC='#1D9E75'
    prod_act_js=[prod.get(k,{'rev':0})['rev'] for k in ['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN','OPP']]
    prod_tgt_js=[prod.get(k,{'tgt':0})['tgt'] for k in ['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN','OPP']]
    prod_delta_js=[(prod.get(k,{'rev':0,'tgt':1})) for k in ['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN','OPP']]
 
    delta_rows=''.join([
        f'''<div style="display:grid;grid-template-columns:44px 1fr 40px;gap:0;padding:4px 6px;border-bottom:1px solid #f0f0f0;background:{"#fff" if i%2==0 else "#f9f9f9"};align-items:center">
      <span style="font-weight:700;color:#333">{k.strip()}</span>
      <span style="text-align:right;font-weight:700;color:#C00000">({fmtK(abs(pd["rev"]-pd["tgt"]))})</span>
      <span style="text-align:right;font-weight:700;color:#C00000">({abs(round((pd["rev"]-pd["tgt"])/pd["tgt"]*100 if pd["tgt"] else 0))}%)</span>
    </div>'''
        for i,(k,pd) in enumerate(zip(['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN','OPP'],[prod.get(k,{'rev':0,'tgt':1}) for k in ['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN','OPP']]))
    ])
 
    # ── New + worst client rows ────────────────────────────────────────────────
    max_new=new_cl['rev'].max() if len(new_cl) else 1
    new_rows=''.join([f'''<div style="display:grid;grid-template-columns:22px 1fr 110px 90px 58px;gap:8px;align-items:center;padding:6px 12px;border-bottom:1px solid #f0f0f0;font-size:11px;background:{"#fff" if i%2==0 else "#f9f9f9"}">
      <span style="color:#888;font-size:10px">{i+1}</span>
      <div><div style="font-weight:700">{r["Cust Name"]}</div><div style="width:{round(r["rev"]/max_new*100)}%;height:3px;background:#2E75B6;border-radius:2px;margin-top:3px"></div></div>
      <span style="font-size:10px;color:#444">{PROV_FULL.get(r["prov"],r["prov"])}</span>
      <span style="text-align:right;font-weight:700">{fmtK(r["rev"])}</span>
      <span style="text-align:right;color:#666">{r["inv"]}</span>
    </div>''' for i,(_,r) in enumerate(new_cl.iterrows())])
 
    worst_rows=''.join([f'''<div style="display:grid;grid-template-columns:22px 1fr 110px 84px 84px 84px 68px;gap:8px;align-items:center;padding:6px 12px;border-bottom:1px solid #f0f0f0;font-size:11px;background:{"#fff" if i%2==0 else "#f9f9f9"}">
      <span style="color:#888;font-size:10px">{i+1}</span>
      <div><div style="font-weight:700">{r["Cust Name"]}</div>{"<span style=\'font-size:9px;background:#C0000022;color:#C00000;padding:1px 5px;border-radius:3px'>Lost</span>" if r["rev"]<=0 else ""}</div>
      <span style="font-size:10px;color:#444">{PROV_FULL.get(r["prov"],r["prov"])}</span>
      <span style="text-align:right;color:#666">{fmtK(r["py_rev"])}</span>
      <span style="text-align:right">{"—" if r["rev"]<=0 else fmtK(r["rev"])}</span>
      <span style="text-align:right;font-weight:700;color:#C00000">({fmtK(abs(r["drop"]))})</span>
      <span style="text-align:right;font-weight:700;color:#C00000">({abs(round(r["drop"]/r["py_rev"]*100,1))}%)</span>
    </div>''' for i,(_,r) in enumerate(worst.iterrows())])
 
    # ── Pre-compute renamed dataframes (can't use dict literals in f-string) ────
    top_avg_r  = top_avg.rename(columns={"Cust Name":"_1"})
    bot_avg_r  = bot_avg.rename(columns={"Cust Name":"_1"})
    if 'prov' not in top_avg_r.columns and 'prov' in top_avg.columns:
        top_avg_r['prov'] = top_avg['prov'].values
    if 'prov' not in bot_avg_r.columns and 'prov' in bot_avg.columns:
        bot_avg_r['prov'] = bot_avg['prov'].values
 
    # ── Pre-compute all JS data strings (avoids f-string brace confusion) ────
    _avgs    = [round(mar_avg,2), round(apr_avg,2), round(may_avg,2), round(jun_avg,2)]
    _valid_a = [v for v in _avgs if v>0]
    _y2_min  = max(0, round(min(_valid_a)/10)*10 - 10) if _valid_a else 0
    _y2_max  = round(max(_valid_a)/10)*10 + 20         if _valid_a else 100
 
    _cum_act = str(cum_act)
    _cum_tgt = str(cum_tgt)
    _cum_py  = str(cum_py)
    _avgs_js = str(_avgs)
 
    _prod_act_js = str([round(float(prod.get(k,{}).get('rev',0))) for k in ['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN','OPP']])
    _prod_tgt_js = str([round(float(prod.get(k,{}).get('tgt',0))) for k in ['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN','OPP']])
 
    _w1_act = int(w1_act); _w2_act = int(w2_act)
    _w1_tgt = int(w1_tgt); _w2_tgt = int(w2_tgt)
    _mtd_max = max(_w1_act, _w1_tgt, _w2_act, _w2_tgt, 1)
 
    # JS object literals for monthly table — pre-built as plain strings
    _mth_acts = "{" + f"Mar:{int(mar_act)},Apr:{int(apr_act)},May:{int(may_act)},Jun:{int(jun_act)},Jul:0,Aug:0,Sep:0,Oct:0,Nov:0,Dec:0,Jan:0,Feb:0" + "}"
    _mth_tgts = "{" + f"Mar:{mar_tgt},Apr:{apr_tgt},May:{may_tgt},Jun:{jun_tgt_full},Jul:{jul_tgt},Aug:{aug_tgt},Sep:{sep_tgt},Oct:{oct_tgt},Nov:{nov_tgt},Dec:{dec_tgt},Jan:{jan_tgt},Feb:{feb_tgt}" + "}"
 
    _daily = round(jun_tgt_full/20)
 
    # ── Assemble HTML ──────────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>HW24 Group Sales Dashboard — {report_date}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',Calibri,Arial,sans-serif;background:#f4f5f7;color:#222;padding:16px}}
.db{{max-width:1100px;margin:0 auto;background:#fff;border-radius:8px;box-shadow:0 2px 12px rgba(0,0,0,.10);overflow:hidden;padding-bottom:24px}}
.title-bar{{background:#1F3864;color:#fff;padding:16px 24px;text-align:center;font-size:20px;font-weight:700}}
.sub-bar{{display:flex;justify-content:space-between;padding:6px 20px;font-size:12px;color:#666;border-bottom:1px solid #e0e0e0;background:#f9f9f9}}
.kpi-row{{display:grid;grid-template-columns:repeat(5,1fr);border:1px solid #e0e0e0;margin:14px 14px 0}}
.kpi{{padding:10px 8px;text-align:center;border-right:1px solid #e0e0e0}}.kpi:last-child{{border-right:none}}
.kpi-label{{font-size:10px;font-weight:700;color:#fff;padding:5px 4px;margin:-10px -8px 8px;text-align:center}}
.kpi-val{{font-size:20px;font-weight:700;line-height:1}}
.section-hdr{{background:#1F3864;color:#fff;padding:8px 14px;font-size:12px;font-weight:700;margin:14px 14px 0;border-radius:4px}}
.tbl-hdr{{display:grid;background:#2E75B6;color:#fff;font-size:11px;font-weight:700;padding:7px 12px;gap:8px}}
.bold-row{{font-weight:700}}
.total-row{{background:#DCE6F1!important;font-weight:700}}
.chart-pair{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:14px 14px 0}}
.chart-box{{border:1px solid #e0e0e0;border-radius:6px;padding:12px}}
.chart-title{{font-size:11px;font-weight:700;text-align:center;color:#1F3864;margin-bottom:6px}}
.items-hdr{{display:grid;grid-template-columns:2.5fr 36px 62px 44px 34px;gap:3px;padding:5px 7px;background:#2E75B6;color:#fff;font-size:8.5px;font-weight:700;border-radius:4px 4px 0 0}}
.footer{{text-align:center;color:#aaa;font-size:10px;margin:20px 14px 0;padding-top:12px;border-top:1px solid #eee}}
</style>
</head>
<body><div class="db">
<div class="title-bar">HW24 GROUP SALES — NATIONAL DASHBOARD</div>
<div class="sub-bar"><span>{report_date}</span><span>March 2026 – June 2026 &nbsp;|&nbsp; MTD: June 2026 — {week_label} ({days_elapsed} days) &nbsp;|&nbsp; All targets = AFS26 ×1.19</span></div>
 
<div class="kpi-row">
  <div class="kpi"><div class="kpi-label" style="background:#2E75B6">Total Turnover</div><div class="kpi-val" style="color:#1F3864">R{total_rev:,.0f}</div></div>
  <div class="kpi"><div class="kpi-label" style="background:#ED7D31">Units Sold</div><div class="kpi-val" style="color:#ED7D31">{total_units:,.0f}</div></div>
  <div class="kpi"><div class="kpi-label" style="background:#375623">Avg Price / Unit</div><div class="kpi-val" style="color:#375623">R{avg_price:.2f}</div></div>
  <div class="kpi"><div class="kpi-label" style="background:#843C0C">Active Clients</div><div class="kpi-val" style="color:#843C0C">{clients:,}</div></div>
  <div class="kpi"><div class="kpi-label" style="background:#595959">Invoices</div><div class="kpi-val" style="color:#595959">{invoices:,}</div></div>
</div>
 
<!-- Product matrix -->
<div style="margin:10px 14px 0">
  <div style="display:grid;grid-template-columns:repeat(5,1fr);border:1px solid #e0e0e0">
    <div style="border:1px solid #e0e0e0"><div style="background:#2E75B6;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">LPCTO</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#2E75B6">{prod["LPCTO"]["rev"]:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{prod["LPCTO"]["avg"]:.2f}</div></div>
    <div style="border:1px solid #e0e0e0"><div style="background:#ED7D31;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">STPRO</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#ED7D31">{prod["STPRO"]["rev"]:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{prod["STPRO"]["avg"]:.2f}</div></div>
    <div style="border:1px solid #e0e0e0"><div style="background:#375623;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">LPMTO</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#375623">{prod["LPMTO"]["rev"]:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{prod["LPMTO"]["avg"]:.2f}</div></div>
    <div style="border:1px solid #e0e0e0"><div style="background:#843C0C;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">LPIMP</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#843C0C">{prod["LPIMP"]["rev"]:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{prod["LPIMP"]["avg"]:.2f}</div></div>
    <div style="border:1px solid #e0e0e0"><div style="background:#595959;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">UFLEX</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#595959">{prod["UFLEX"]["rev"]:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{prod["UFLEX"]["avg"]:.2f}</div></div>
  </div>
  <div style="display:flex;justify-content:center;border:1px solid #e0e0e0;border-top:none">
    <div style="flex:0 0 20%;border:1px solid #e0e0e0"><div style="background:#7B6E58;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">RAWMT</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#7B6E58">{prod["RAWMT"]["rev"]:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{prod["RAWMT"]["avg"]:.2f}</div></div>
    <div style="flex:0 0 20%;border:1px solid #e0e0e0"><div style="background:#888780;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">REPLEN</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#888780">{prod["REPLEN"]["rev"]:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#888">Avg R{prod["REPLEN"]["avg"]:.2f}</div></div>
    <div style="flex:0 0 20%;border:1px solid #e0e0e0"><div style="background:#C00000;color:#fff;font-size:11px;font-weight:700;padding:6px 8px;text-align:center">OPP</div><div style="font-size:16px;font-weight:700;text-align:center;padding:8px 4px 2px;color:#C00000">{prod["OPP"]["rev"]:,.0f}</div><div style="font-size:11px;text-align:center;padding:0 4px 7px;color:#C00000">Avg R{prod["OPP"]["avg"]:.2f}</div></div>
  </div>
</div>
 
<!-- Monthly table -->
<div style="margin:10px 14px 0;overflow-x:auto">
  <table id="monthly-table" style="width:100%;border-collapse:collapse;font-size:10px">
    <thead><tr style="background:#2E75B6;color:#fff">
      <td style="padding:6px 10px;font-weight:700;min-width:70px">Metric</td>
      <td style="padding:6px 8px;text-align:right;font-weight:700">Mar</td><td style="padding:6px 8px;text-align:right;font-weight:700">Apr</td><td style="padding:6px 8px;text-align:right;font-weight:700">May</td>
      <td style="padding:6px 8px;text-align:right;font-weight:700;background:#375623">Jun ▶</td>
      <td style="padding:6px 8px;text-align:right;font-weight:700;opacity:.7">Jul</td><td style="padding:6px 8px;text-align:right;font-weight:700;opacity:.7">Aug</td><td style="padding:6px 8px;text-align:right;font-weight:700;opacity:.7">Sep</td><td style="padding:6px 8px;text-align:right;font-weight:700;opacity:.7">Oct</td><td style="padding:6px 8px;text-align:right;font-weight:700;opacity:.7">Nov</td><td style="padding:6px 8px;text-align:right;font-weight:700;opacity:.7">Dec</td><td style="padding:6px 8px;text-align:right;font-weight:700;opacity:.7">Jan</td><td style="padding:6px 8px;text-align:right;font-weight:700;opacity:.7">Feb</td>
      <td style="padding:6px 8px;text-align:right;font-weight:700;background:#1F3864">TOTAL</td>
    </tr></thead>
    <tbody id="monthly-body"></tbody>
  </table>
</div>
 
<!-- Period table -->
<div class="section-hdr">Period Performance</div>
<div style="margin:0 14px">
<div class="tbl-hdr" style="grid-template-columns:1.9fr 1.1fr 1.1fr 1fr .9fr 1fr"><span>Period</span><span style="text-align:right">Actual</span><span style="text-align:right">Target</span><span style="text-align:right">Delta</span><span style="text-align:right">% Delta</span><span style="text-align:right">Avg Unit Price</span></div>
{tbl_row("YTD Turnover (Mar–May)", int(ytd_act), ytd_tgt, mar_avg, bold=True)}
{tbl_row("March",   int(mar_act), mar_tgt, mar_avg, indent=True)}
{tbl_row("April",   int(apr_act), apr_tgt, apr_avg, indent=True)}
{tbl_row("May",     int(may_act), may_tgt, may_avg, indent=True)}
{tbl_row("MTD June (W1+W2)", int(jun_act), jun_tgt_pro, jun_avg, bold=True)}
{tbl_row("Week 1",  int(w1_act), w1_tgt, indent=True)}
{tbl_row("Week 2",  int(w2_act), w2_tgt, indent=True)}
{tbl_row("Total incl. June MTD", int(total_rev), ytd_tgt+jun_tgt_pro, avg_price, total=True)}
{tbl_row("YTD Units (Mar–May)", ytd_u_act, ytd_u_tgt, bold=True)}
{tbl_row("MTD Units (June)", jun_u_act, jun_u_tgt, indent=True)}
{tbl_row("Total Units", int(total_units), tot_u_tgt, total=True)}
</div>
 
<!-- Charts -->
<div class="chart-pair">
  <div class="chart-box">
    <div class="chart-title">YTD CUMULATIVE ACTUAL vs PY (MAR–JUN, JUN PRO-RATED)</div>
    <div style="display:flex;gap:10px;justify-content:center;margin-bottom:6px;font-size:9px;color:#666">
      <span>── Actual &nbsp;</span><span style="border-top:2px dashed #888780;display:inline-block;width:18px;height:0;vertical-align:middle"></span><span>&nbsp;Prior Year &nbsp;</span><span style="border-top:1.5px dotted #888780;display:inline-block;width:18px;height:0;vertical-align:middle"></span><span style="color:#888780">&nbsp;Avg Price</span>
    </div>
    <div style="position:relative;height:190px"><canvas id="cPY"></canvas></div>
  </div>
  <div class="chart-box">
    <div class="chart-title">YTD CUMULATIVE ACTUAL vs TARGET (MAR–JUN, JUN PRO-RATED)</div>
    <div style="display:flex;gap:10px;justify-content:center;margin-bottom:6px;font-size:9px;color:#666">
      <span style="border-top:2px dashed #C00000;display:inline-block;width:18px;height:0;vertical-align:middle"></span><span>&nbsp;Target &nbsp;</span>── Actual
    </div>
    <div style="position:relative;height:190px"><canvas id="cYTD"></canvas></div>
  </div>
</div>
 
<!-- National Performance YTD -->
<div class="section-hdr">National Performance by Region — YTD (Mar–Jun, Jun pro-rated) &nbsp;<span style="font-size:10px;font-weight:400;opacity:.8">Targets = AFS26 ×1.19</span></div>
<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin:8px 14px 0">{region_cards()}</div>
 
<!-- MTD -->
<div class="section-hdr" style="margin-top:14px">National Performance by Region — MTD (June 2026 {week_label}) &nbsp;<span style="font-size:10px;font-weight:400;opacity:.8">W1={fmtK(w1_tgt)} · W2={fmtK(w2_tgt)} ({days_elapsed}/5 days) · Daily={fmtK(round(jun_tgt_full/20))}</span></div>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:8px 14px 0;align-items:stretch">
  <div>
    <div style="display:grid;grid-template-columns:1.6fr .95fr .95fr .8fr .7fr;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;padding:7px 8px;gap:4px;border-radius:4px 4px 0 0">
      <span>Province</span><span style="text-align:right">MTD Actual</span><span style="text-align:right">Jun Target</span><span style="text-align:right">Delta</span><span style="text-align:right">% Delta</span>
    </div>
    {mtd_rows()}
  </div>
  <div style="display:flex;flex-direction:column;padding:10px;border:1px solid #e0e0e0;border-radius:6px;background:#fff">
    <div style="font-size:12px;font-weight:700;text-align:center;color:#1F3864;margin-bottom:8px">MTD ACTUAL VS TARGET — JUNE 2026</div>
    <div style="flex:1;position:relative;min-height:160px"><canvas id="cMTD"></canvas></div>
  </div>
</div>
 
<!-- Top 10 clients -->
<div class="section-hdr" style="margin-top:14px">Top 10 Clients Nationally — YTD</div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:22px 1fr 120px 72px 72px 68px;gap:8px;align-items:center;padding:6px 12px;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0"><span>#</span><span>Client</span><span>Province</span><span style="text-align:right">YTD Sales</span><span style="text-align:right">PY YTD</span><span style="text-align:right">YoY %</span></div>
  {client_rows(top10, cp26)}
</div>
 
<!-- New clients -->
<div class="section-hdr" style="margin-top:14px">Top 10 New Clients — AFS27 (no prior year history)</div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:22px 1fr 120px 90px 58px;gap:8px;align-items:center;padding:6px 12px;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0"><span>#</span><span>Client</span><span>Province</span><span style="text-align:right">YTD Revenue</span><span style="text-align:right">Invoices</span></div>
  {new_rows}
</div>
 
<!-- Worst clients -->
<div class="section-hdr" style="margin-top:14px">Worst Performing Clients — AFS26 vs AFS27</div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:22px 1fr 110px 84px 84px 84px 68px;gap:8px;align-items:center;padding:6px 12px;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0"><span>#</span><span>Client</span><span>Province</span><span style="text-align:right">AFS26</span><span style="text-align:right">AFS27</span><span style="text-align:right">Value Drop</span><span style="text-align:right">% Drop</span></div>
  {worst_rows}
</div>
 
<!-- Avg price -->
<div class="section-hdr" style="margin-top:14px">Top 10 — Highest Avg Price/Unit &nbsp;<span style="font-size:10px;font-weight:400;opacity:.8">PY rev &gt; R50K</span></div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:20px 1fr 110px 78px 78px 62px 62px 62px;gap:6px;align-items:center;padding:6px 10px;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0"><span>#</span><span>Client</span><span>Province</span><span style="text-align:right">PY Revenue</span><span style="text-align:right">CY Revenue</span><span style="text-align:right">PY Avg</span><span style="text-align:right">CY Avg</span><span style="text-align:right">Change</span></div>
  {avg_rows(top_avg_r, True)}
</div>
<div class="section-hdr" style="margin-top:14px">Bottom 10 — Lowest Avg Price/Unit &nbsp;<span style="font-size:10px;font-weight:400;opacity:.8">PY rev &gt; R50K</span></div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:20px 1fr 110px 78px 78px 62px 62px 62px;gap:6px;align-items:center;padding:6px 10px;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0"><span>#</span><span>Client</span><span>Province</span><span style="text-align:right">PY Revenue</span><span style="text-align:right">CY Revenue</span><span style="text-align:right">PY Avg</span><span style="text-align:right">CY Avg</span><span style="text-align:right">Change</span></div>
  {avg_rows(bot_avg_r, False)}
</div>
 
<!-- Product group -->
<div class="section-hdr" style="margin-top:14px">YTD Turnover by Product Group &amp; Top/Bottom 8 Items (excl. LPMTO)</div>
<div style="display:grid;grid-template-columns:1fr 1.3fr 1.3fr;gap:8px;margin:8px 14px 0;align-items:stretch">
  <div style="border:1px solid #e0e0e0;border-radius:6px;padding:10px;display:flex;flex-direction:column">
    <div class="chart-title">YTD TURNOVER BY PRODUCT GROUP — ACTUAL vs TARGET (PY+19%)</div>
    <div id="hbar-delta" style="display:flex;flex-direction:column;justify-content:center;min-width:130px;padding:2px 0;font-size:9px;margin-bottom:6px">
      <div style="display:grid;grid-template-columns:44px 1fr 40px;background:#2E75B6;color:#fff;font-weight:700;padding:4px 6px;border-radius:4px 4px 0 0;font-size:9px"><span>Group</span><span style="text-align:right">Delta (R)</span><span style="text-align:right">%</span></div>
      {delta_rows}
    </div>
    <div id="hbar-wrap" style="position:relative;flex:1;min-height:160px"><canvas id="cHBar"></canvas></div>
  </div>
  <div style="display:flex;flex-direction:column;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden">
    <div style="font-size:10px;font-weight:700;color:#1F3864;padding:6px 10px 4px;background:#f9f9f9;border-bottom:1px solid #e0e0e0">Top 8 Items (excl. LPMTO)</div>
    <div class="items-hdr" style="border-radius:0"><span>Product</span><span style="text-align:center">Grp</span><span style="text-align:right">Turnover</span><span style="text-align:right">Avg</span><span style="text-align:right">%</span></div>
    <div>{items_rows(top8)}</div>
  </div>
  <div style="display:flex;flex-direction:column;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden">
    <div style="font-size:10px;font-weight:700;color:#1F3864;padding:6px 10px 4px;background:#f9f9f9;border-bottom:1px solid #e0e0e0">Bottom 8 Items (excl. LPMTO)</div>
    <div class="items-hdr" style="border-radius:0"><span>Product</span><span style="text-align:center">Grp</span><span style="text-align:right">Turnover</span><span style="text-align:right">Avg</span><span style="text-align:right">%</span></div>
    <div>{items_rows(bot8)}</div>
  </div>
</div>
 
<!-- Salesperson -->
<div class="section-hdr" style="margin-top:14px">YTD Top 10 Salesperson by Turnover (Branch from AS/AT lookup)</div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:24px 1fr 130px 72px 72px 62px;gap:6px;align-items:center;padding:6px 10px;background:#2E75B6;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0"><span>#</span><span>Salesperson</span><span>Branch / Province</span><span style="text-align:right">YTD Sales</span><span style="text-align:right">PY YTD</span><span style="text-align:right">YoY %</span></div>
  {sp_rows(top_sp)}
</div>
 
<div class="section-hdr" style="margin-top:14px">YTD Bottom 5 Salesperson &nbsp;<span style="font-size:10px;font-weight:400;opacity:.8">Active salespeople only (PY &gt; R100K)</span></div>
<div style="margin:0 14px">
  <div style="display:grid;grid-template-columns:24px 1fr 130px 72px 72px 62px;gap:6px;align-items:center;padding:6px 10px;background:#843C0C;color:#fff;font-size:10px;font-weight:700;border-radius:4px 4px 0 0"><span>#</span><span>Salesperson</span><span>Branch / Province</span><span style="text-align:right">YTD Sales</span><span style="text-align:right">PY YTD</span><span style="text-align:right">YoY %</span></div>
  {sp_rows(bot_sp, max_=5)}
</div>
 
<div class="footer">HW24 Group Sales National Dashboard &bull; Confidential &bull; Generated {report_date} &bull; Targets = AFS26 ×1.19 &bull; {week_label} ({days_elapsed} days elapsed, {wd_el}/{wd_tot} MTD working days)</div>
</div>
 
<script>
const gridC='rgba(128,128,128,0.1)', GREY='#888780', AC='#1D9E75';
const avgs={_avgs_js};
const avgDs={{label:'Avg Price',data:avgs,type:'line',borderColor:GREY,borderWidth:1.5,borderDash:[2,3],pointRadius:3,fill:false,yAxisID:'y2'}};
const y2={{position:'right',min:{_y2_min},max:{_y2_max},ticks:{{font:{{size:9}},color:GREY,callback:v=>'R'+v.toFixed(0)}},grid:{{display:false}},title:{{display:true,text:'Avg Price (R)',font:{{size:9}},color:GREY}}}};
 
new Chart(document.getElementById('cPY'),{{type:'line',data:{{labels:['March','April','May','June (MTD)'],datasets:[
  {{label:'Actual',data:{_cum_act},borderColor:AC,borderWidth:2,pointRadius:4,pointStyle:'circle',pointBackgroundColor:AC,tension:.3,fill:false,yAxisID:'y'}},
  {{label:'PY',data:{_cum_py},borderColor:GREY,borderWidth:2,borderDash:[5,4],pointRadius:4,pointStyle:'triangle',pointBackgroundColor:GREY,tension:.3,fill:false,yAxisID:'y'}},
  {{...avgDs}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{x:{{ticks:{{font:{{size:9}}}},grid:{{display:false}}}},y:{{ticks:{{font:{{size:9}},callback:v=>'R'+(v/1e6).toFixed(1)+'M'}},grid:{{color:gridC}}}},y2}}}}}}}});
 
new Chart(document.getElementById('cYTD'),{{type:'line',data:{{labels:['March','April','May','June (MTD)'],datasets:[
  {{label:'Target',data:{_cum_tgt},borderColor:'#C00000',borderWidth:2,borderDash:[5,4],pointRadius:4,pointStyle:'rectRot',pointBackgroundColor:'#C00000',tension:.3,fill:false,yAxisID:'y'}},
  {{label:'Actual',data:{_cum_act},borderColor:AC,borderWidth:2,pointRadius:4,pointStyle:'circle',pointBackgroundColor:AC,tension:.3,fill:false,yAxisID:'y'}},
  {{...avgDs}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{x:{{ticks:{{font:{{size:9}}}},grid:{{display:false}}}},y:{{ticks:{{font:{{size:9}},callback:v=>'R'+(v/1e6).toFixed(1)+'M'}},grid:{{color:gridC}}}},y2}}}}}}}});
 
new Chart(document.getElementById('cMTD'),{{type:'bar',data:{{labels:['Week 1','Week 2'],datasets:[
  {{label:'Actual',data:[{_w1_act},{_w2_act}],backgroundColor:'#22A548',barPercentage:1.0,categoryPercentage:0.85}},
  {{label:'Target',data:[{_w1_tgt},{_w2_tgt}],backgroundColor:'#8B0000',barPercentage:1.0,categoryPercentage:0.85}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>'R'+Math.round(ctx.raw).toLocaleString()}}}}}},scales:{{x:{{ticks:{{font:{{size:11}}}},grid:{{display:false}}}},y:{{min:0,ticks:{{font:{{size:9}},callback:v=>'R'+(v/1e6).toFixed(2)+'M'}},grid:{{color:'rgba(180,180,180,0.25)'}}}}}}}}}}}});
 
const vp={{id:'hv',afterDatasetsDraw(c){{const ctx=c.ctx;c.getDatasetMeta(0).data.forEach((b,i)=>{{const v=c.data.datasets[0].data[i];ctx.save();ctx.font='700 8px Segoe UI';ctx.fillStyle='#333';ctx.textAlign='left';ctx.textBaseline='middle';ctx.fillText('R'+(v/1e6).toFixed(1)+'M',b.x+4,b.y);ctx.restore();}});}}}};
new Chart(document.getElementById('cHBar'),{{type:'bar',plugins:[vp],data:{{
  labels:['LPCTO','STPRO','LPMTO','UFLEX','LPIMP','RAWMT','REPLEN','OPP'],
  datasets:[
    {{label:'Actual',data:{_prod_act_js},backgroundColor:'#2E75B6',barPercentage:.45,categoryPercentage:.85}},
    {{label:'Target (PY+19%)',data:{_prod_tgt_js},backgroundColor:'rgba(200,200,200,0.55)',borderColor:'#999',borderWidth:1,barPercentage:.45,categoryPercentage:.85}}
  ]
}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,
  plugins:{{legend:{{display:true,position:'bottom',labels:{{font:{{size:9}},boxWidth:10}}}}}},
  scales:{{x:{{min:0,ticks:{{font:{{size:9}},callback:v=>'R'+(v/1e6).toFixed(0)+'M'}},grid:{{color:'rgba(180,180,180,0.2)'}}}},y:{{ticks:{{font:{{size:10,weight:'700'}},color:'#333'}},grid:{{display:false}}}}}},layout:{{padding:{{right:42}}}}}}}}}});
 
const mActs={_mth_acts};
const mTgts={_mth_tgts};
const mths=['Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb'];
const tAct=mths.reduce((s,m)=>s+mActs[m],0);
const tTgt=mths.reduce((s,m)=>s+(mActs[m]>0?mTgts[m]:0),0);
const body=document.getElementById('monthly-body');
[['Actual (R)','act'],['Target (R)','tgt'],['Delta (R)','delta']].forEach(([lbl,key])=>{{
  const tr=document.createElement('tr'); tr.style.borderBottom='1px solid #f0f0f0';
  const td0=document.createElement('td'); td0.style.cssText='padding:5px 10px;font-weight:'+(key==='delta'?'700':'400')+';white-space:nowrap;background:#f9f9f9'; td0.textContent=lbl; tr.appendChild(td0);
  mths.forEach(m=>{{
    const act=mActs[m],tgt=mTgts[m],delta=act-tgt,isFut=act===0&&!['Mar','Apr','May','Jun'].includes(m);
    let v,col;
    if(key==='act'){{v=act;col='#1F3864';}}
    if(key==='tgt'){{v=tgt;col='#595959';}}
    if(key==='delta'){{v=isFut?null:delta;col=delta<0?'#C00000':'#375623';}}
    const td=document.createElement('td'); td.style.cssText='padding:5px 8px;text-align:right;white-space:nowrap;font-weight:'+(key==='delta'?'700':'400')+';'+(['Jun'].includes(m)?'background:#f0f8f0':'')+';';
    if(v===null||(key==='act'&&isFut)){{td.innerHTML='<span style="color:#ccc">—</span>';}}
    else{{const fmt='R'+(Math.abs(v)/1e6).toFixed(1)+'M'; td.innerHTML=v<0?'<span style="color:#C00000">('+fmt+')</span>':\`<span style="color:\${{col}}">\${{fmt}}</span>\`;}}
    tr.appendChild(td);
  }});
  const tdT=document.createElement('td'); tdT.style.cssText='padding:5px 8px;text-align:right;white-space:nowrap;font-weight:700;background:#DCE6F1';
  const tD2=tAct-tTgt;
  if(key==='act')tdT.innerHTML='<span style="color:#1F3864">R'+(tAct/1e6).toFixed(1)+'M</span>';
  else if(key==='tgt')tdT.innerHTML='<span style="color:#595959">R'+(tTgt/1e6).toFixed(1)+'M</span>';
  else{{const c2=tD2<0?'#C00000':'#375623',f2='R'+(Math.abs(tD2)/1e6).toFixed(1)+'M'; tdT.innerHTML='<span style="color:'+c2+'">'+( tD2<0?'('+f2+')':f2)+'</span>';}}
  tr.appendChild(tdT); body.appendChild(tr);
}});
</script></body></html>"""
    return html
