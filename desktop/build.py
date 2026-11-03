# يبني صفحة النسخة المكتبية من anbar-finance.html
# الاستخدام: python3 build.py ../anbar-finance.html xlsx.full.min.js
import sys,os,shutil
page=open(sys.argv[1],encoding='utf-8').read()
page=page.replace('sync.textContent=ok?`متصل ·','sync.textContent=ok?`${window.ANBAR_DESKTOP?"ملف محلي":"متصل"} ·')
os.makedirs('app',exist_ok=True)
shutil.copy(sys.argv[2],'app/xlsx.full.min.js')
head='<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>:root{color-scheme:light}body{margin:0;font-family:system-ui,sans-serif;font-size:14px}img{max-width:100%}[hidden]{display:none!important}</style><script src="xlsx.full.min.js"></script></head><body>'
open('app/index.html','w',encoding='utf-8').write(head+page+'</body></html>')
# ثم: npm install && npx electron-builder --win portable --x64
