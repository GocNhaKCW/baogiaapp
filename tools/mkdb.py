# Tạo db.json từ file Excel Sapo: python3 -I tools/mkdb.py <file.xlsx> <dd/mm/yyyy>
import openpyxl,json,time,warnings,sys
warnings.filterwarnings('ignore')
rows=list(openpyxl.load_workbook(sys.argv[1],read_only=True).active.iter_rows(values_only=True))
h=rows[0]; assert h[1].startswith('Mã SKU') and 'Giá bán lẻ' in h[6], h
last=''; items={}
for r in rows[1:]:
    sku=str(r[1] or '').strip().upper()
    if not sku: continue
    t=str(r[0] or '').strip() or last; last=t
    g=int(r[6] or 0); ton=int((r[3] or 0)+(r[4] or 0))
    items[sku]={'t':t,'g':g,'i':(r[2] or '').strip(),'s':ton}
d=sys.argv[2]
json.dump({'date':d,'ts':int(time.time()*1000),'source':'Sapo export '+d.replace('/','.'),'items':items},open('db.json','w',encoding='utf-8'),ensure_ascii=False,separators=(',',':'))
print(len(items), sum(1 for v in items.values() if v['s']>0), 'in stock')
