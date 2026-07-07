#!/usr/bin/env python3
"""
Bチーム売上報告_LINEインポート.xlsx → スプレッドシート 一括インポートスクリプト
使い方: python3 import_line_data.py
"""

import json
import os
import urllib.request
import urllib.error
import sys

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl', '-q'])
    import openpyxl

# =====================================================
# 設定
# =====================================================
# 【修正F】このGAS_URLはツール本体(sales-manager.html)側のGAS_URLとデプロイIDが異なります。
#   ツール本体（最新）: AKfycbx4rKSQgHJ_w_CxaOGepDE-V1iLUr_Wixnrkxl7H6fJ81lxHKWoz58CeUipJRn9fm9_
#   このスクリプト     : AKfycbxEQIzgiwmRjUn2HubpVLwTaNKrsNSot5HTrsnbLoyBzJGA1h41Fgo9rZr5Ppk8yeP-
#   古い/別のデプロイを指している可能性があるため、実際に一括インポートを実行する前に
#   「このURLが現行の正しいデプロイと一致しているか」をユーザー自身で必ず確認してください。
#   （このスクリプトが勝手にURLを書き換えることはしません）
GAS_URL = 'https://script.google.com/macros/s/AKfycbxEQIzgiwmRjUn2HubpVLwTaNKrsNSot5HTrsnbLoyBzJGA1h41Fgo9rZr5Ppk8yeP-/exec'
EXCEL_PATH = '/Users/kyo/Downloads/Bチーム売上報告_LINEインポート.xlsx'

# 【修正F】bulkImportエンドポイントは専用トークン(IMPORT_TOKEN)で保護されるようになったため、
# 環境変数 SALES_IMPORT_TOKEN からトークンを読み込む（ソースコードに直書きしない）。
# 実行前に `export SALES_IMPORT_TOKEN='（GAS側に設定したものと同じ値）'` を行ってください。
IMPORT_TOKEN = os.environ.get('SALES_IMPORT_TOKEN')
if not IMPORT_TOKEN:
    print("エラー: 環境変数 SALES_IMPORT_TOKEN を設定してください。")
    print("例: export SALES_IMPORT_TOKEN='（GASのスクリプトプロパティ IMPORT_TOKEN と同じ値）'")
    sys.exit(1)

# =====================================================
# Excelを読み込む
# =====================================================
print(f"Excelを読み込んでいます: {EXCEL_PATH}")
wb = openpyxl.load_workbook(EXCEL_PATH)
ws = wb.active

rows = []
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
    date_str, member, route, line_name, result, product, amount, payment, note, recording = row

    if not date_str:
        continue

    date_s = str(date_str)[:10]  # "2026-05-01" 形式

    # 同日内の行順を秒単位で記録（ソート時の順序保持のため）
    h = i % 86400
    hour = h // 3600
    minute = (h % 3600) // 60
    second = h % 60
    timestamp = f"{date_s}T{hour:02d}:{minute:02d}:{second:02d}"

    rows.append({
        'timestamp': timestamp,
        'date': date_s,
        'memberName': str(member) if member else '',
        'category':   str(route)  if route  else '',
        'lineName':   str(line_name) if line_name else '',
        'result':     str(result) if result else '',
        'product':    str(product) if product else '',
        'amount':     int(amount)  if amount  else 0,
        'payment':    int(payment) if payment else 0,
        'note':       str(note)    if note    else '',
        'recording':  str(recording) if recording else '',
        'id':         f'import-{i+1:04d}',
    })

print(f"読み込み完了: {len(rows)}件")

# =====================================================
# GASにPOSTで送信（バッチ単位で分割）
# =====================================================
BATCH_SIZE = 50  # 1回あたり50件ずつ送信

def post_to_gas(batch, batch_num, total_batches):
    # 【修正F】専用トークンを同送する（GAS側でIMPORT_TOKENと一致するかチェックされる）
    payload_dict = {'type': 'bulkImport', 'rows': batch, 'importToken': IMPORT_TOKEN}
    payload_bytes = json.dumps(payload_dict, ensure_ascii=False).encode('utf-8')

    req = urllib.request.Request(
        GAS_URL,
        data=payload_bytes,
        headers={'Content-Type': 'application/json; charset=utf-8'},
        method='POST'
    )

    with urllib.request.urlopen(req, timeout=120) as resp:
        body = resp.read().decode('utf-8')
        result = json.loads(body)
        return result

print(f"\nスプレッドシートに送信中... (バッチサイズ: {BATCH_SIZE}件)")
total_batches = (len(rows) + BATCH_SIZE - 1) // BATCH_SIZE
total_imported = 0

for i in range(0, len(rows), BATCH_SIZE):
    batch = rows[i:i + BATCH_SIZE]
    batch_num = i // BATCH_SIZE + 1
    print(f"  バッチ {batch_num}/{total_batches} 送信中 ({len(batch)}件)...", end=' ')

    try:
        result = post_to_gas(batch, batch_num, total_batches)
        if result.get('ok'):
            total_imported += result.get('imported', len(batch))
            print(f"✅ OK")
        else:
            print(f"❌ エラー: {result.get('error', '不明')}")
            print("処理を中断します。")
            sys.exit(1)
    except urllib.error.HTTPError as e:
        print(f"❌ HTTPエラー: {e.code} {e.reason}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ 例外: {e}")
        sys.exit(1)

print(f"\n✅ インポート完了！ 合計 {total_imported}件 をスプレッドシートに追加しました。")
print("ツールをリロードすると反映されます。")
