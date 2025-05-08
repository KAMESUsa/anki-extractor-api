MAX_USES = 50            # １キーあたりの最大実行回数（24科目＋リトライ分をカバー）
USES_COL_NAME = 'Uses'   # スプレッドシート上の使用回数を格納する列名

import os
import tempfile
import zipfile
from flask import Flask, request, send_file, abort
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from anki_extractor import run_extraction

# --- Google Sheets 認証 & 初期化 ---
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]
# credentials.json はサービスアカウントのキー
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
gc    = gspread.authorize(creds)

SHEET_URL    = "https://docs.google.com/spreadsheets/d/＜シートID＞/edit"
sheet        = gc.open_by_url(SHEET_URL).sheet1
header       = sheet.row_values(1)
uses_col_idx = header.index(USES_COL_NAME) + 1  # 1-based 列番号
# ----------------------------------------

app = Flask(__name__)

@app.route("/extract", methods=["POST"])
def extract():
    try:
        # 0) メール＋キーを受け取って照合
        email = request.form.get('email')
        key   = request.form.get('key')
        if not email or not key:
            abort(400, "email と key が必要です")

        # シートから対応行を検索
        records = sheet.get_all_records()
        for idx, rec in enumerate(records, start=2):  # データは2行目から
            if rec.get('Email','').strip() == email.strip() and rec.get('Key','').strip() == key.strip():
                uses   = int(rec.get(USES_COL_NAME, 0))
                row_num = idx
                break
        else:
            abort(401, "無効な email または key です")

        # 利用回数上限チェック
        if uses >= MAX_USES:
            abort(403, f"このライセンスキーは既に {MAX_USES} 回利用されています。")

        # 1) アップロードされたPDFファイルを受け取る
        if 'front' not in request.files or 'back' not in request.files:
            abort(400, "front と back の両方のファイルが必要です")
        front = request.files['front']
        back  = request.files['back']

        # 2) 一時フォルダを作成してPDFを保存
        tmp = tempfile.mkdtemp()
        f1  = os.path.join(tmp, "front.pdf")
        f2  = os.path.join(tmp, "back.pdf")
        front.save(f1)
        back.save(f2)

        # 3) 出力先フォルダとCSVファイル名を設定
        outdir   = os.path.join(tmp, "output")
        os.makedirs(outdir, exist_ok=True)
        csv_name = "result.csv"

        # 4) 変換処理を呼び出す
        run_extraction(f1, f2, outdir, csv_name)

        # 5) 成功したら Uses をインクリメントしてシートに書き戻し
        sheet.update_cell(row_num, uses_col_idx, uses + 1)

        # 6) 出力結果をZIPにまとめて返却
        zip_path = os.path.join(tmp, "result.zip")
        with zipfile.ZipFile(zip_path, "w") as zf:
            for fn in os.listdir(outdir):
                zf.write(os.path.join(outdir, fn), arcname=fn)

        return send_file(
            zip_path,
            mimetype="application/zip",
            as_attachment=True,
            download_name="result.zip"
        )

    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        return f"<pre>{tb}</pre>", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
