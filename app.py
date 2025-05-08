import os
import tempfile
import zipfile
from flask import Flask, request, send_file, abort
from anki_extractor import run_extraction

app = Flask(__name__)

@app.route("/extract", methods=["POST"])
def extract():
    try:
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

        # 5) 出力結果をZIPにまとめて返却
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
        # エラーとトレースバックを HTML として返す
        return f"<pre>{tb}</pre>", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

