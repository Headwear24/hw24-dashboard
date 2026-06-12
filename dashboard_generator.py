import pandas as pd, os, json

df = pd.read_excel('/mnt/user-data/uploads/Raw_data_uplaod_file_upto_10th_June26.xlsx', sheet_name='AFS 26 Data', header=0)
afs27 = df[df['AFS']=='AFS27'].copy()
afs26 = df[df['AFS']=='AFS26'].copy()
UPLIFT=1.19; WD_PER=9/20; prov_col='Province Orginal'; sp_col='Sales Person Use'
REGION_MAP={'GP':'Gauteng','LP':'Gauteng','NW':'Gauteng','MP':'Gauteng','FS':'Gauteng',
            'KZN':'KZN','ZN':'KZN','WC':'Western Cape','NC':'Western Cape',
            'EC':'Eastern Cape','ZZZ':'International'}
for d_ in [afs27,afs26]: d_['Region']=d_[prov_col].map(REGION_MAP).fillna('Other')
afs26_ytd=afs26[afs26['Period1'].isin(['Mar','Apr','May'])]
afs26_jun=afs26[afs26['Period1']=='Jun']
sp_lu=df[['ALYSHA','DURBAN']].dropna().drop_duplicates(); sp_lu.columns=['sp','branch']
sp_lu['branch']=sp_lu['branch'].str.strip()
branch_dict=dict(zip(sp_lu['sp'],sp_lu['branch']))
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

def build_html(d):
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
<title>HW24 {r} Sales Dashboard — 9 June 2026</title>
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
<div class="sub-bar"><span>9-Jun-26</span><span>March 2026 – June 2026 &nbsp;|&nbsp; MTD June W1+W2 &nbsp;|&nbsp; Targets = AFS26 ×1.19</span></div>

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

<div class="footer">HW24 {r} Sales Dashboard &bull; Confidential &bull; Generated 9 June 2026 &bull; Data: March–June 2026 (June W1+W2) &bull; Targets = AFS26 ×1.19</div>
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

output_dir = '/mnt/user-data/outputs'
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
