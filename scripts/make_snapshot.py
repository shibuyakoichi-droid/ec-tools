# -*- coding: utf-8 -*-
"""在処管理の「時点」ファイルを作る。

  python scripts/make_snapshot.py 2026-08-31 "2026年8月末 棚卸し"

いまSupabaseに入っている在庫（生地・商品・坂本繊維）をそのまま
snapshots/<日付>.json に書き出し、snapshots/index.json を更新する。
コミットしてプッシュすると、process_mgmt.html の「時点」セレクタに出る。

棚卸しxlsxを取り込む手順：
  1. process_fabrics / yarn_stock を新しいxlsxの内容に更新する
  2. このスクリプトを棚卸し基準日で実行する
  3. git add snapshots && git commit && git push
"""
import json, os, sys, io, urllib.request, urllib.parse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SB = 'https://yukcdqnnevomhdcpsvms.supabase.co/rest/v1/'
KEY = ('eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inl1a2NkcW5uZXZvbWhkY3Bzdm1z'
       'Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY0Mjc2MTUsImV4cCI6MjA5MjAwMzYxNX0.XBSc2kQNwR_bleSwrfYZD8l0fbKTbL1Z3S7ewEtSkDc')
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SNAPDIR = os.path.join(ROOT, 'snapshots')


def get(path):
    req = urllib.request.Request(SB + path, headers={'apikey': KEY, 'Authorization': 'Bearer ' + KEY})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode('utf-8'))


def build(date, label, fabrics, products, sakamoto, facs):
    return {
        'date': date, 'label': label,
        'fabrics': [{'fabric_no': f.get('fabric_no'), 'color': f.get('color') or '',
                     'qty': f.get('qty') or 0, 'location_id': f.get('location_id'),
                     'location_name': facs.get(f.get('location_id'), ''),
                     'memo': f.get('memo') or ''} for f in fabrics],
        'products': [{'code': p.get('code'), 'name': p.get('name'),
                      'qty': p.get('qty') or 0, 'location_id': p.get('location_id'),
                      'location_name': facs.get(p.get('location_id'), '')} for p in products],
        'sakamoto': [{'name': y.get('color_name'), 'qty': y.get('stock_kg') or 0,
                      'unit': y.get('unit') or '反', 'memo': y.get('notes') or ''} for y in sakamoto],
    }


def write(snap):
    os.makedirs(SNAPDIR, exist_ok=True)
    path = os.path.join(SNAPDIR, snap['date'] + '.json')
    with open(path, 'w', encoding='utf-8', newline='\n') as f:
        json.dump(snap, f, ensure_ascii=False, indent=1)

    idx_path = os.path.join(SNAPDIR, 'index.json')
    idx = []
    if os.path.exists(idx_path):
        idx = json.load(open(idx_path, encoding='utf-8'))
    idx = [e for e in idx if e['date'] != snap['date']]
    idx.append({'date': snap['date'], 'label': snap['label'], 'file': snap['date'] + '.json'})
    idx.sort(key=lambda e: e['date'], reverse=True)
    with open(idx_path, 'w', encoding='utf-8', newline='\n') as f:
        json.dump(idx, f, ensure_ascii=False, indent=1)

    tan = sum(r['qty'] for r in snap['fabrics'])
    mai = sum(r['qty'] for r in snap['products'])
    sak = sum(r['qty'] for r in snap['sakamoto'])
    print(f"{snap['date']} {snap['label']}: 生地{len(snap['fabrics'])}件/{tan}反・"
          f"商品{len(snap['products'])}件/{mai}枚・坂本{len(snap['sakamoto'])}品番/{sak}反 -> {path}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    date = sys.argv[1]
    label = sys.argv[2] if len(sys.argv) > 2 else ''
    facs = {f['id']: f['name'] for f in get('factories?select=id,name')}
    write(build(date, label,
                get('process_fabrics?select=*&order=fabric_no'),
                get('process_products?select=*&order=code'),
                get('yarn_stock?select=*&supplier=eq.' + urllib.parse.quote('坂本繊維') + '&order=color_name'),
                facs))
