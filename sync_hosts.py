"""Compatibility builder: the database is authoritative; never import the old Excel."""
from pathlib import Path
import shutil

source = Path('index.html')
if 'data-qihai-shared="supabase-v1"' not in source.read_text(encoding='utf-8'):
    raise SystemExit('請先上傳共用資料庫版 index.html，停止發布以免退回舊版。')
Path('_site').mkdir(exist_ok=True)
shutil.copyfile(source, '_site/index.html')
print('已準備共用資料庫查詢頁；不從舊 Excel 覆寫線上資料。')
