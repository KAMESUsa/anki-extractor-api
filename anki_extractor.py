# anki_extractor.py

import re
import os
import csv
import shutil
from pathlib import Path
import fitz  # PyMuPDF

def run_extraction(pdf1, pdf2, output_base, csv_filename):
    """
    pdf1: Front用PDFのパス
    pdf2: Back用PDFのパス
    output_base: 画像＆CSV出力先フォルダ
    csv_filename: 出力するCSVの名前
    """
    # 定数
    LINE_HEIGHT      = 20
    TOP_MARGIN_LINES = 1
    BOT_MARGIN_LINES = 1
    H_MARGIN_LINES   = 1
    ZOOM_FACTOR      = 3.5
    FILL_MIN         = 0.85
    FILL_MAX         = 0.99

    # 出力フォルダを初期化
    out_base = Path(output_base)
    if out_base.exists():
        shutil.rmtree(out_base)
    out_base.mkdir(parents=True)
    print(f"出力先: {out_base}")

    # 画像抽出用の内部関数
    def extract_images(pdf_path, suffix):
        doc = fitz.open(pdf_path)
        toc = doc.get_toc()
        # 目次からページ番号→タイトルマッピング
        title_map = {e[2]-1: e[1] for e in toc if re.match(r'^\d+\.\d+', e[1])}

        sub_re = re.compile(r'^([B-FＢ-Ｆ])[:：]')
        fw2hw = {'Ｂ':'B','Ｃ':'C','Ｄ':'D','Ｅ':'E','Ｆ':'F'}

        for pidx in range(doc.page_count):
            if pidx not in title_map:
                continue
            title = title_map[pidx]
            safe  = re.sub(r'[\\/*?:"<>|]', '_', title)
            page  = doc.load_page(pidx)

            # グレー領域（問題部分）の検出
            frames = [
                fitz.Rect(d['rect'])
                for d in page.get_drawings()
                if d.get('type')=='f' and d.get('fill')
                   and FILL_MIN <= sum(d['fill'][:3])/3 < FILL_MAX
            ]
            if not frames:
                continue
            frame  = sorted(frames, key=lambda r: r.y0)[0]
            matrix = fitz.Matrix(ZOOM_FACTOR, ZOOM_FACTOR)
            mh     = LINE_HEIGHT * H_MARGIN_LINES
            left   = max(0, frame.x0 - mh)
            right  = min(page.rect.x1, frame.x1 + mh)

            # タイトル行の座標も取得して上下マージンを決定
            ytops = []
            try:
                tr = page.search_for(title)[0]
                ytops.append(tr.y0 - LINE_HEIGHT*TOP_MARGIN_LINES)
            except:
                pass
            ytops.append(frame.y0 - LINE_HEIGHT*TOP_MARGIN_LINES)

            top0    = max(0, min(ytops))
            bottom0 = min(page.rect.y1, frame.y1 + LINE_HEIGHT*BOT_MARGIN_LINES)
            crop    = fitz.Rect(left, top0, right, bottom0)

            # タグ（B〜F など）を検出して切り出し位置を決定
            td    = page.get_text("dict", clip=crop)
            items = []
            for blk in td["blocks"]:
                if blk["type"]!=0: continue
                for ln in blk["lines"]:
                    txt = ''.join(sp['text'] for sp in ln['spans']).strip()
                    m   = sub_re.match(txt)
                    if m:
                        raw = m.group(1)
                        tag = fw2hw.get(raw, raw)
                        y   = ln['bbox'][1]
                        items.append((y, tag))
            items.sort(key=lambda x: x[0])

            boundaries = [crop.y0] + [y - 0.5*LINE_HEIGHT for y,_ in items] + [crop.y1]
            tags       = [None] + [tag for _,tag in items]

            # 画像を切り出して保存
            for i in range(len(boundaries)-1):
                clip_rect = fitz.Rect(left,
                                      boundaries[i],
                                      right,
                                      boundaries[i+1])
                pix = page.get_pixmap(matrix=matrix,
                                      clip=clip_rect,
                                      alpha=False)
                if i == 0:
                    fname = f"{safe}{suffix}.jpg"
                else:
                    fname = f"{safe}_{tags[i]}{suffix}.jpg"
                out_path = out_base / fname
                pix.save(out_path)
                print(f"[OK] ページ{pidx+1}: {fname}")

        doc.close()

    # ① Front/Back をそれぞれ切り出し
    extract_images(pdf1, suffix="")
    extract_images(pdf2, suffix="_解答")

    # ② 切り出した JPEG で CSV を組み立て
    files = list(out_base.glob("*.jpg"))
    unique = {}
    for f in files:
        base = f.stem.replace('_解答','')
        unique.setdefault(base, {
            'front': '',
            'back': '',
            'subject': Path(pdf1).stem
        })
    for f in files:
        stem = f.stem
        base = stem.replace('_解答','')
        if stem.endswith('_解答'):
            unique[base]['back'] = f.name
        else:
            unique[base]['front'] = f.name

    # ソート用のキー関数
    def make_sort_key(s):
        m = re.match(r"^(\d+)\.(\d+)", s)
        if m:
            return (int(m.group(1)), int(m.group(2)))
        m2 = re.match(r"^(\d+)", s)
        if m2:
            return (int(m2.group(1)), 0)
        return (9999, 0)

    # ③ CSVを書き出し
    with open(out_base / csv_filename, 'w', newline='', encoding='utf-8') as wf:
        writer = csv.writer(wf)
        writer.writerow(['front','back','subject'])
        for base in sorted(unique.keys(), key=make_sort_key):
            row = unique[base]
            writer.writerow([row['front'], row['back'], row['subject']])

    print(f"CSV完成: {out_base / csv_filename}")
