import openpyxl,glob,json,datetime,sys,re
D=sys.argv[1]; out=sys.argv[2]
def clean(v):
    if isinstance(v,datetime.datetime): return v.strftime('%Y-%m-%d')
    if isinstance(v,str):
        v=re.sub(r'\s+',' ',v).strip(); return v or None
    if isinstance(v,float) and v.is_integer() and abs(v)>1: return int(v)
    return v
def num(v): return v if isinstance(v,(int,float)) else 0
def normst(s):
    if not s: return 'غير محددة'
    s=s.replace('اجراء','إجراء')
    return s
P=[];META={'main':{},'x':{}}
f=glob.glob(D+'/60f*')[0]
ws=openpyxl.load_workbook(f,data_only=True).active
hdr=[clean(c.value) for c in ws[1]]
META['main']={'hdr':hdr[:15],'sheet':ws.title}
for r in ws.iter_rows(min_row=2,values_only=True):
    if not r[3] or not isinstance(r[0],(int,float)): continue
    raw=[[hdr[i],clean(v)] for i,v in enumerate(r) if hdr[i] and clean(v) is not None]
    ca=num(r[6]) or num(r[5])
    P.append(dict(id='M%d'%len(P),src='الموقف المالي العام',st=normst(clean(r[1])),pl=clean(r[2]) or 'غير محددة',nm=clean(r[3]),co=clean(r[4]) or '—',
      c0=num(r[5]),c=ca,ap=num(r[7]),sp=num(r[9]),adv=num(r[11]),bank=num(r[12]),fin=num(r[8]),phy=r[14] if isinstance(r[14],(int,float)) else None,raw=raw))
f=glob.glob(D+'/7ac*')[0]
wb=openpyxl.load_workbook(f,data_only=True)
for ws in wb.worksheets[:2]:
    hdr=[clean(c.value) for c in ws[2]][:43]
    title=next((clean(c.value) for c in ws[3] if clean(c.value)),ws.title)
    META['x'][ws.title]={'hdr':hdr,'title':title}
    yrs=[i for i,h in enumerate(hdr) if h and h.startswith('المصروف') and ('عام' in h or 'السنة' in h)]
    for r in ws.iter_rows(min_row=4,values_only=True):
        if not r[7] or not isinstance(r[0],(int,float)): continue
        raw=[[hdr[i],clean(v)] for i,v in enumerate(r) if i<len(hdr) and hdr[i] and clean(v) is not None]
        c=num(r[14]) or num(r[13])
        P.append(dict(id='X%d'%len(P),src='ملف '+ws.title,st=normst(clean(r[5])),pl=clean(r[6]) or ws.title,nm=clean(r[7]),co=clean(r[10]) or '—',
          c0=num(r[13]),c=c,ap=0,sp=num(r[16]),adv=0,bank=0,fin=num(r[15]),phy=None,cno=clean(r[1]),cdt=clean(r[2]),dur=num(r[11]),ext=num(r[12]),
          yrs=[[hdr[i].replace('المصروف الكلي ','').replace('المصروف ',''),num(r[i])] for i in yrs],raw=raw))
json.dump({'P':P,'meta':META},open(out,'w'),ensure_ascii=False,separators=(',',':'))
print(len(P),sum(1 for p in P if p['id'][0]=='X'))
