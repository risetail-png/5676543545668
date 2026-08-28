#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bridge_pdf_link_app_sono3.py
=============================
橋梁定期点検PDF ナビゲーションボタン追加ツール（その３-１／その３-２ 版）

「データ記録様式（その３-１）損傷図」と「データ記録様式（その３-２）損傷写真」
の間に、径間番号が一致するページ同士を相互にジャンプできるボタンを自動追加します。

v8/v9（その９↔その１０版）をベースに、以下の点を今回のPDF様式に合わせて変更しています。
  - キーワードを「データ記録様式（その３-１）」「データ記録様式（その３-２）」に変更
    （全角括弧表記）
  - ページが /Rotate 90（用紙は縦A4だが表示は横向き）で作成されているPDFに対応。
    ボタン画像・リンク注釈の座標を、ページの回転角（0/90/180/270）に応じて
    正しい向き・位置になるよう変換してから描画するようにした。

起動方法:
    python bridge_pdf_link_app_sono3.py

必要ライブラリ:
    pip install pikepdf pypdf Pillow
    ※ tkinterdnd2 はドラッグ＆ドロップ用（任意）
"""

import io
import os
import queue
import re
import sys
import threading
import tkinter as tk
from collections import defaultdict
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

# ── オプション依存 ────────────────────────────────────────────────────────────
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

# ── 必須ライブラリチェック ────────────────────────────────────────────────────
MISSING = []
try:
    import pikepdf
    from pikepdf import Array, Dictionary, Name, Stream
except ImportError:
    MISSING.append("pikepdf")

try:
    import pypdf
except ImportError:
    MISSING.append("pypdf")

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    MISSING.append("pillow")


# ═══════════════════════════════════════════════════════════════════════════════
#  定数
# ═══════════════════════════════════════════════════════════════════════════════

BTN_Y1, BTN_Y2 = 8.0, 34.0
BTN_H   = BTN_Y2 - BTN_Y1
BTN_GAP = 5.0
IMG_SCALE = 3

COLOR_FORWARD         = (46,  97, 184)   # 損傷図 → 損傷写真（青）
COLOR_OUTLINE_FORWARD = (20,  55, 130)
COLOR_BACK            = (34, 139,  69)   # 損傷写真 → 損傷図（緑）
COLOR_OUTLINE_BACK    = (20,  90,  45)

# 今回リンク対象の様式キーワード（全角括弧表記に注意）
KEYWORD_DIAGRAM = "データ記録様式（その３-１）"   # 損傷図
KEYWORD_PHOTO   = "データ記録様式（その３-２）"   # 損傷写真

RE_PHOTO_PAGE_NUM = re.compile(r'写真番号[\s　]*(\d+)((?:\s+\d+)*)')

BTN_ROW_MAX_BOTTOM = 8   # 青・緑ボタンの1行最大数
ROW_GAP = 2.0            # 段間の隙間(pt)


# ═══════════════════════════════════════════════════════════════════════════════
#  ユーティリティ
# ═══════════════════════════════════════════════════════════════════════════════

def _normalize_text(text):
    text = text.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    return text.replace('－', '-').replace('―', '-')


def _parse_photo_page_nums(text):
    """ページ内に含まれる「写真番号」の実数値（前回写真の参照値は除く）を抽出する。"""
    text = _normalize_text(text)
    work = text
    work = re.sub(r'\d{4}[./]\d{2}[./]\d{2}', '', work)
    work = re.sub(r'\d+\.\d+', '', work)
    work = re.sub(r'写真番号\s*\d+\s*[-－]\s*\d+\s*の\S+', '', work)
    work = re.sub(r'前回\s*[-－]?\s*\d*', '', work)
    work = re.sub(r'[-－]\s*\d+', '', work)
    nums = []
    for m in RE_PHOTO_PAGE_NUM.finditer(work):
        nums.append(int(m.group(1)))
        for extra in re.findall(r'\d+', m.group(2)):
            nums.append(int(extra))
    if not nums:
        return []
    base  = min(nums)
    upper = base + 15
    for m in re.finditer(r'\b(\d{1,2})\b', work):
        n = int(m.group(1))
        if base <= n <= upper:
            nums.append(n)
    return sorted(set(nums))


def find_japanese_font():
    candidates = [
        r"C:\Windows\Fonts\msgothic.ttc",
        r"C:\Windows\Fonts\meiryo.ttc",
        r"C:\Windows\Fonts\YuGothM.ttc",
        r"C:\Windows\Fonts\yugothm.ttc",
        "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
        "/System/Library/Fonts/Hiragino Sans GB.ttc",
        "/Library/Fonts/Osaka.ttf",
        "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
        "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


# ═══════════════════════════════════════════════════════════════════════════════
#  ページ分類・径間番号取得
# ═══════════════════════════════════════════════════════════════════════════════

def classify_pages(pdf_path):
    """
    各ページを分類して返す。
    戻り値:
      diag_pages   : 損傷図ページ index リスト（データ記録様式 その３-１）
      photo_pages  : 損傷写真ページ index リスト（データ記録様式 その３-２）
    """
    reader = pypdf.PdfReader(pdf_path)
    diag, photo = [], []
    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        if KEYWORD_DIAGRAM in text:
            diag.append(i)
        if KEYWORD_PHOTO in text:
            photo.append(i)
    return diag, photo


def get_span_number_fallback(text):
    """テキストから径間番号を文字列で返す（例: "1", "1-1"）。
    このPDF様式では「起点側　終点側　<径間番号>　緯度…」の順で並ぶ。"""
    text = _normalize_text(text)
    m = re.search(r'起点側\s*終点側\s*(\d+[-−]\d+|\d+)', text)
    if m:
        return m.group(1)
    # 予備パターン（径間番号が先に来る様式向け）
    m = re.search(r'(\d+[-−]\d+|\d+)\s*起点側\s*終点側', text)
    return m.group(1) if m else None


# ═══════════════════════════════════════════════════════════════════════════════
#  回転（/Rotate）を考慮した座標変換
# ═══════════════════════════════════════════════════════════════════════════════
#
# このPDFはページ実体（MediaBox）がA4縦のまま /Rotate 90 が指定されており、
# ビューア上では横向き（landscape）で表示される。
# ボタン画像やリンク注釈は「ページ本来の（回転前の）座標系」で配置する必要が
# あるため、以下の関数で「表示上の座標（見た目の横向きページでの座標）」から
# 「ページ本来の座標」への変換を行う。

def get_rotation(pdf, page_idx):
    try:
        r = int(pdf.pages[page_idx].get('/Rotate', 0))
    except Exception:
        r = 0
    return r % 360


def get_native_size(pdf, page_idx):
    mb = pdf.pages[page_idx]['/MediaBox']
    return float(mb[2]) - float(mb[0]), float(mb[3]) - float(mb[1])


def get_display_size(pdf, page_idx):
    """ビューア表示上の幅・高さ（/Rotate適用後）を返す。"""
    w, h = get_native_size(pdf, page_idx)
    r = get_rotation(pdf, page_idx)
    if r in (90, 270):
        return h, w
    return w, h


def visual_rect_to_native(r, w, h, x1, y1, x2, y2):
    """表示上の矩形（x1<x2, y1<y2）をページ本来の座標系の矩形に変換する。"""
    if r == 90:
        return w - y2, x1, w - y1, x2
    if r == 180:
        return w - x2, h - y2, w - x1, h - y1
    if r == 270:
        return y1, h - x2, y2, h - x1
    return x1, y1, x2, y2


def image_cm_matrix(r, w, h, x0, y0, wb, hb):
    """単位正方形の画像を wb×hb に拡大して表示上の位置(x0, y0)（左下基準）へ
    配置するための 'cm' オペランド (a, b, c, d, e, f) を、回転角に応じて返す。
    こうして描いた画像は、ビューアが /Rotate を適用したあとに
    正しい向き・正しい位置で表示される。"""
    if r == 90:
        return 0, wb, -hb, 0, w - y0, x0
    if r == 180:
        return -wb, 0, 0, -hb, w - x0, h - y0
    if r == 270:
        return 0, -wb, hb, 0, y0, h - x0
    return wb, 0, 0, hb, x0, y0


# ═══════════════════════════════════════════════════════════════════════════════
#  ボタン描画・追加
# ═══════════════════════════════════════════════════════════════════════════════

def render_button_jpeg(btn_list, total_w_pt, btn_h_pt,
                       fill_color, outline_color, font_path):
    img_w = int(total_w_pt * IMG_SCALE)
    img_h = int(btn_h_pt  * IMG_SCALE)
    img   = Image.new('RGB', (img_w, img_h), (255, 255, 255))
    draw  = ImageDraw.Draw(img)
    n        = len(btn_list)
    gap_px   = int(BTN_GAP * IMG_SCALE)
    btn_w_px = (img_w - gap_px * (n + 1)) // n
    by_margin = int(2 * IMG_SCALE)
    bh        = img_h - int(4 * IMG_SCALE)
    padding_v = int(3 * IMG_SCALE)
    fsize = bh - padding_v * 2
    if font_path:
        for fs in range(fsize, 4, -1):
            try:
                fnt = ImageFont.truetype(font_path, fs)
            except Exception:
                fnt = ImageFont.load_default()
                break
            ok = True
            for label, _ in btn_list:
                bb = fnt.getbbox(label)
                tw, th = bb[2] - bb[0], bb[3] - bb[1]
                if th > bh - padding_v * 2 or tw > btn_w_px - int(4 * IMG_SCALE):
                    ok = False
                    break
            if ok:
                break
    else:
        fnt = ImageFont.load_default()
    for i, (label, _) in enumerate(btn_list):
        bx = gap_px + i * (btn_w_px + gap_px)
        draw.rounded_rectangle([bx, by_margin, bx + btn_w_px, by_margin + bh],
                               radius=int(4 * IMG_SCALE),
                               fill=fill_color, outline=outline_color, width=2)
        bb = fnt.getbbox(label)
        tw, th = bb[2] - bb[0], bb[3] - bb[1]
        draw.text((bx + (btn_w_px - tw) // 2, by_margin + (bh - th) // 2),
                  label, fill=(255, 255, 255), font=fnt)
    buf = io.BytesIO()
    img.save(buf, format='JPEG', quality=92)
    return buf.getvalue(), img_w, img_h


def _place_button_row(pdf, page_idx, row_btn_list, disp_w, y1, y2,
                      fill_color, outline_color, font_path, xobj_name):
    """1行分のボタン画像をページに描画しアノテーションを追加する（内部共通処理）。
    disp_w: ビューア表示上のページ幅。座標はすべて「表示上の座標系」で受け取り、
    内部でページの /Rotate に応じたページ本来の座標系に変換して書き込む。"""
    page      = pdf.pages[page_idx]
    r         = get_rotation(pdf, page_idx)
    native_w, native_h = get_native_size(pdf, page_idx)

    margin_l  = 64.0
    margin_r  = disp_w - 48.0
    btn_total = margin_r - margin_l
    btn_h     = y2 - y1

    jpeg_bytes, img_w, img_h = render_button_jpeg(
        row_btn_list, btn_total, btn_h, fill_color, outline_color, font_path)

    xobj = Stream(pdf, jpeg_bytes)
    xobj['/Type']             = Name('/XObject')
    xobj['/Subtype']          = Name('/Image')
    xobj['/Width']            = img_w
    xobj['/Height']           = img_h
    xobj['/ColorSpace']       = Name('/DeviceRGB')
    xobj['/BitsPerComponent'] = 8
    xobj['/Filter']           = Name('/DCTDecode')
    xobj_ref = pdf.make_indirect(xobj)

    if '/XObject' not in page['/Resources']:
        page['/Resources']['/XObject'] = pikepdf.Dictionary()
    page['/Resources']['/XObject'][xobj_name] = xobj_ref

    a, b, c, d, e, f = image_cm_matrix(r, native_w, native_h,
                                       margin_l, y1, btn_total, btn_h)
    stream_content = (f"q\n{a:.4f} {b:.4f} {c:.4f} {d:.4f} "
                      f"{e:.4f} {f:.4f} cm\n{xobj_name} Do\nQ\n").encode('latin-1')
    cstream = Stream(pdf, stream_content)
    existing = page['/Contents']
    page['/Contents'] = pikepdf.Array(
        (list(existing) if isinstance(existing, pikepdf.Array) else [existing])
        + [pdf.make_indirect(cstream)]
    )

    n        = len(row_btn_list)
    btn_w_pt = (btn_total - BTN_GAP * (n + 1)) / n
    annots   = list(page.get('/Annots', pikepdf.Array()))
    for i, (_, target_idx) in enumerate(row_btn_list):
        bx1 = margin_l + BTN_GAP + i * (btn_w_pt + BTN_GAP)
        bx2 = bx1 + btn_w_pt
        nx1, ny1, nx2, ny2 = visual_rect_to_native(r, native_w, native_h,
                                                   bx1, y1, bx2, y2)
        dest = pikepdf.Array([pdf.pages[target_idx].obj, Name('/Fit')])
        annots.append(pdf.make_indirect(Dictionary(
            Type=Name('/Annot'), Subtype=Name('/Link'),
            Rect=Array([pikepdf.Real(nx1), pikepdf.Real(ny1),
                        pikepdf.Real(nx2), pikepdf.Real(ny2)]),
            Border=Array([pikepdf.Real(0)] * 3),
            Dest=dest, H=Name('/I'),
        )))
    page['/Annots'] = pikepdf.Array(annots)


def add_buttons_bottom(pdf, page_idx, btn_list, disp_w, disp_h,
                       fill_color, outline_color, font_path, xobj_prefix):
    """ページ下端に配置。BTN_ROW_MAX_BOTTOM個超で下方向に折り返す（2段目が1段目の下）。"""
    rows = [btn_list[i:i+BTN_ROW_MAX_BOTTOM]
            for i in range(0, len(btn_list), BTN_ROW_MAX_BOTTOM)]
    num_rows = len(rows)
    for r_i, row in enumerate(rows):
        row_from_bottom = num_rows - 1 - r_i
        y1 = BTN_Y1 + row_from_bottom * (BTN_H + ROW_GAP)
        y2 = y1 + BTN_H
        _place_button_row(pdf, page_idx, row, disp_w, y1, y2,
                          fill_color, outline_color, font_path,
                          f'/{xobj_prefix}{page_idx}r{r_i}')


# ═══════════════════════════════════════════════════════════════════════════════
#  1ファイル処理
# ═══════════════════════════════════════════════════════════════════════════════

def process_one(input_path, output_path, font_path, log_cb):
    """1つのPDFを処理してボタンを追加する。失敗時は例外を送出。"""
    diag_pages, photo_pages = classify_pages(input_path)

    if not diag_pages:
        raise RuntimeError(f"損傷図ページ（{KEYWORD_DIAGRAM}）が見つかりません。")
    if not photo_pages:
        raise RuntimeError(f"損傷写真ページ（{KEYWORD_PHOTO}）が見つかりません。")

    log_cb(f"  損傷図ページ（その３-１）  : {[p+1 for p in diag_pages]}")
    log_cb(f"  損傷写真ページ（その３-２）: {[p+1 for p in photo_pages]}")

    reader = pypdf.PdfReader(input_path)
    diag_span, photo_span = {}, {}

    for pidx in diag_pages:
        diag_span[pidx] = get_span_number_fallback(reader.pages[pidx].extract_text() or "")
    for pidx in photo_pages:
        photo_span[pidx] = get_span_number_fallback(reader.pages[pidx].extract_text() or "")

    span_to_diag  = defaultdict(list)
    span_to_photo = defaultdict(list)
    for pidx, span in diag_span.items():
        if span is not None:
            span_to_diag[span].append(pidx)
    for pidx, span in photo_span.items():
        if span is not None:
            span_to_photo[span].append(pidx)

    def _span_sort_key(s):
        parts = re.split(r'[-−]', str(s))
        try:
            return tuple(int(p) for p in parts)
        except ValueError:
            return (0, 0)

    all_spans     = sorted(set(span_to_diag.keys()) | set(span_to_photo.keys()),
                           key=_span_sort_key)
    is_multi_span = len(all_spans) > 1
    log_cb(f"  径間: {all_spans}")

    # 損傷写真の写真番号取得（ラベル用）
    photo_page_nums = {}
    for pidx in photo_pages:
        text = reader.pages[pidx].extract_text() or ""
        photo_page_nums[pidx] = _parse_photo_page_nums(text)

    pdf = pikepdf.open(input_path, allow_overwriting_input=True)

    # ── 損傷図 → 損傷写真（青ボタン）────────────────────────────────────────
    for didx in diag_pages:
        span = diag_span.get(didx)
        if span is None:
            continue
        target_photo_pages = span_to_photo.get(span, [])
        if not target_photo_pages:
            continue
        dw, dh = get_display_size(pdf, didx)
        btn_list = []
        for pp in sorted(target_photo_pages):
            pp_span   = photo_span.get(pp)
            page_nums = photo_page_nums.get(pp, [])
            if page_nums:
                nums_str = (f"{min(page_nums)}〜{max(page_nums)}"
                            if len(page_nums) > 1 else f"{page_nums[0]}")
                if is_multi_span and pp_span:
                    if re.search(r'\d+-\d+', str(pp_span)):
                        label = f"{pp_span}\u3000{nums_str}"
                    else:
                        label = (f"{pp_span}-{nums_str}" if len(page_nums) > 1
                                 else f"{pp_span}-{page_nums[0]}")
                else:
                    label = nums_str
            else:
                label = (f"{pp_span}径間・p.{pp+1}"
                         if (is_multi_span and pp_span) else f"p.{pp+1}")
            btn_list.append((label, pp))
        add_buttons_bottom(pdf, didx, btn_list, dw, dh,
                           COLOR_FORWARD, COLOR_OUTLINE_FORWARD,
                           font_path, 'FwdBtn')

    # ── 損傷写真 → 損傷図（緑ボタン）────────────────────────────────────────
    for pp in photo_pages:
        span = photo_span.get(pp)
        if span is None:
            continue
        target_diag_pages = span_to_diag.get(span, [])
        if not target_diag_pages:
            continue
        dw, dh = get_display_size(pdf, pp)
        btn_list = []
        for didx in sorted(target_diag_pages):
            pp_span = photo_span.get(pp)
            label = (f"{pp_span}径間・p.{didx+1}"
                     if (is_multi_span and pp_span) else f"p.{didx+1}")
            btn_list.append((label, didx))
        add_buttons_bottom(pdf, pp, btn_list, dw, dh,
                           COLOR_BACK, COLOR_OUTLINE_BACK,
                           font_path, 'BackBtn')

    pdf.save(output_path)


# ═══════════════════════════════════════════════════════════════════════════════
#  複数ファイル一括処理（バックグラウンドスレッド）
# ═══════════════════════════════════════════════════════════════════════════════

def run_batch(file_list, log_cb, progress_cb, done_cb):
    """file_list: [input_path, ...] → 各ファイルと同フォルダに _linked.pdf を出力"""
    try:
        font_path = find_japanese_font()
        if not font_path:
            raise RuntimeError(
                "日本語フォントが見つかりません。\n"
                "MS ゴシック / ヒラギノ / IPAフォント等をインストールしてください。")
        log_cb(f"フォント: {Path(font_path).name}")

        total   = len(file_list)
        ok_list = []
        ng_list = []

        for i, inp in enumerate(file_list, 1):
            inp = str(inp)
            stem = Path(inp).stem
            out  = str(Path(inp).parent / f"{stem}_linked.pdf")
            log_cb("=" * 48)
            log_cb(f"[{i}/{total}] {Path(inp).name}")
            progress_cb(i, total)
            try:
                process_one(inp, out, font_path, log_cb)
                sz = os.path.getsize(out) / 1024 / 1024
                log_cb(f"  → 保存完了: {Path(out).name}  ({sz:.1f} MB)")
                ok_list.append(Path(inp).name)
            except Exception as e:
                import traceback
                log_cb(f"  エラー: {e}")
                log_cb(traceback.format_exc())
                ng_list.append((Path(inp).name, str(e)))

        log_cb("=" * 48)
        log_cb(f"完了: 成功 {len(ok_list)} / 失敗 {len(ng_list)} / 合計 {total}")
        done_cb(ok_list, ng_list)

    except Exception as e:
        import traceback
        log_cb(f"致命的エラー: {e}")
        log_cb(traceback.format_exc())
        done_cb([], [(str(e), str(e))])


# ═══════════════════════════════════════════════════════════════════════════════
#  GUI
# ═══════════════════════════════════════════════════════════════════════════════

class App(tk.Tk if not HAS_DND else TkinterDnD.Tk):

    BG      = "#1a1f2e"
    PANEL   = "#242938"
    BORDER  = "#2e3548"
    ACCENT  = "#4a7fe8"
    TEXT    = "#e8ecf4"
    SUBTEXT = "#8892aa"
    SUCCESS = "#22a06b"
    ERROR   = "#e8516a"
    WARNING = "#f0a040"
    BTN_HOV = "#5a8ff8"

    def __init__(self):
        super().__init__()
        self.title("橋梁点検PDF リンク追加ツール（その３-１/その３-２版）")
        self.geometry("780x620")
        self.minsize(680, 520)
        self.configure(bg=self.BG)
        self.resizable(True, True)

        self._file_list   = []
        self._status      = tk.StringVar(value="PDFファイルを追加してください")
        self._log_queue   = queue.Queue()
        self._processing  = False

        self._build_ui()
        self._poll_log()

        if MISSING:
            self._show_missing()

    # ── UI構築 ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        hdr = tk.Frame(self, bg=self.BG)
        hdr.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 0))
        tk.Label(hdr, text="橋梁点検PDF", font=("Yu Gothic UI", 10),
                 fg=self.SUBTEXT, bg=self.BG).pack(anchor="w")
        tk.Label(hdr, text="リンク追加ツール（その３-１/その３-２）",
                 font=("Yu Gothic UI Bold", 16, "bold"),
                 fg=self.TEXT, bg=self.BG).pack(anchor="w")
        tk.Label(hdr,
                 text="径間番号をもとに、損傷図（その３-１）↔ 損傷写真（その３-２）間にナビゲーションボタンを自動追加します",
                 font=("Yu Gothic UI", 9), fg=self.SUBTEXT, bg=self.BG
                 ).pack(anchor="w", pady=(2, 0))
        legend = tk.Frame(hdr, bg=self.BG)
        legend.pack(anchor="w", pady=(4, 0))
        tk.Label(legend, text="■", fg="#4a7fe8", bg=self.BG,
                 font=("Yu Gothic UI", 9)).pack(side="left")
        tk.Label(legend, text="損傷図→損傷写真  ",
                 fg=self.SUBTEXT, bg=self.BG, font=("Yu Gothic UI", 9)).pack(side="left")
        tk.Label(legend, text="■", fg="#22a06b", bg=self.BG,
                 font=("Yu Gothic UI", 9)).pack(side="left")
        tk.Label(legend, text="損傷写真→損傷図  ／  出力は各PDFと同じフォルダに _linked.pdf で保存",
                 fg=self.SUBTEXT, bg=self.BG, font=("Yu Gothic UI", 9)).pack(side="left")

        tk.Frame(self, bg=self.BORDER, height=1).grid(
            row=0, column=0, sticky="ew", padx=24, pady=(88, 0))

        main = tk.Frame(self, bg=self.BG)
        main.grid(row=1, column=0, sticky="nsew", padx=24, pady=16)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        file_frame = tk.Frame(main, bg=self.PANEL,
                              highlightbackground=self.BORDER,
                              highlightthickness=1)
        file_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        file_frame.columnconfigure(0, weight=1)

        self._drop_zone = tk.Label(
            file_frame,
            text="📂  ここにPDFをドラッグ＆ドロップ（複数可）\nまたはクリックして選択",
            font=("Yu Gothic UI", 10), fg=self.SUBTEXT, bg=self.PANEL,
            cursor="hand2", pady=16
        )
        self._drop_zone.grid(row=0, column=0, columnspan=2,
                             sticky="ew", padx=16, pady=(12, 4))
        self._drop_zone.bind("<Button-1>", lambda e: self._browse_files())
        self._drop_zone.bind("<Enter>",
            lambda e: self._drop_zone.configure(fg=self.ACCENT))
        self._drop_zone.bind("<Leave>",
            lambda e: self._drop_zone.configure(fg=self.SUBTEXT))

        if HAS_DND:
            self._drop_zone.drop_target_register(DND_FILES)
            self._drop_zone.dnd_bind('<<Drop>>', self._on_drop)

        list_wrap = tk.Frame(file_frame, bg=self.PANEL)
        list_wrap.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 4))
        list_wrap.columnconfigure(0, weight=1)

        self._listbox = tk.Listbox(
            list_wrap, height=5,
            bg="#131720", fg=self.TEXT,
            font=("Yu Gothic UI", 9),
            selectbackground=self.ACCENT,
            relief="flat", bd=0,
            activestyle="none",
        )
        self._listbox.grid(row=0, column=0, sticky="ew")
        sb = ttk.Scrollbar(list_wrap, command=self._listbox.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self._listbox['yscrollcommand'] = sb.set

        btn_row = tk.Frame(file_frame, bg=self.PANEL)
        btn_row.grid(row=2, column=0, sticky="e", padx=16, pady=(0, 10))
        self._mk_small_btn(btn_row, "追加…",     self._browse_files).pack(side="left", padx=(0, 6))
        self._mk_small_btn(btn_row, "選択削除",  self._remove_selected).pack(side="left", padx=(0, 6))
        self._mk_small_btn(btn_row, "全クリア",  self._clear_files).pack(side="left")

        log_frame = tk.Frame(main, bg=self.PANEL,
                             highlightbackground=self.BORDER,
                             highlightthickness=1)
        log_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)

        tk.Label(log_frame, text="処理ログ",
                 font=("Yu Gothic UI", 9), fg=self.SUBTEXT, bg=self.PANEL
                 ).grid(row=0, column=0, sticky="w", padx=12, pady=(8, 2))

        self._log = tk.Text(
            log_frame, bg="#131720", fg=self.SUBTEXT,
            font=("Consolas", 9), relief="flat", bd=0,
            state="disabled", wrap="word",
            insertbackground=self.TEXT,
            selectbackground=self.ACCENT,
        )
        self._log.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        log_sb = ttk.Scrollbar(log_frame, command=self._log.yview)
        log_sb.grid(row=1, column=1, sticky="ns", pady=(0, 8), padx=(0, 4))
        self._log['yscrollcommand'] = log_sb.set

        self._log.tag_configure("info",    foreground=self.SUBTEXT)
        self._log.tag_configure("success", foreground=self.SUCCESS)
        self._log.tag_configure("error",   foreground=self.ERROR)
        self._log.tag_configure("warn",    foreground=self.WARNING)
        self._log.tag_configure("accent",  foreground=self.ACCENT)

        footer = tk.Frame(self, bg=self.BG)
        footer.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 16))
        footer.columnconfigure(0, weight=1)

        tk.Label(footer, textvariable=self._status,
                 font=("Yu Gothic UI", 9), fg=self.SUBTEXT, bg=self.BG,
                 anchor="w").grid(row=0, column=0, sticky="w")

        self._progress = ttk.Progressbar(footer, mode='determinate', length=200)
        self._progress.grid(row=0, column=1, padx=(12, 12))

        self._run_btn = tk.Button(
            footer, text="▶  処理開始",
            font=("Yu Gothic UI Bold", 10, "bold"),
            fg="white", bg=self.ACCENT,
            activeforeground="white", activebackground=self.BTN_HOV,
            relief="flat", bd=0, padx=20, pady=8,
            cursor="hand2", command=self._start
        )
        self._run_btn.grid(row=0, column=2)
        self._run_btn.bind("<Enter>",
            lambda e: self._run_btn.configure(bg=self.BTN_HOV))
        self._run_btn.bind("<Leave>",
            lambda e: self._run_btn.configure(bg=self.ACCENT))

    def _mk_small_btn(self, parent, text, cmd):
        return tk.Button(parent, text=text,
                         font=("Yu Gothic UI", 9),
                         fg=self.TEXT, bg=self.BORDER,
                         activeforeground=self.TEXT, activebackground=self.ACCENT,
                         relief="flat", bd=0, padx=10, pady=3,
                         cursor="hand2", command=cmd)

    # ── ファイル操作 ──────────────────────────────────────────────────────────
    def _browse_files(self):
        paths = filedialog.askopenfilenames(
            title="入力PDFを選択（複数可）",
            filetypes=[("PDFファイル", "*.pdf"), ("すべてのファイル", "*.*")]
        )
        if paths:
            self._add_files(list(paths))

    def _on_drop(self, event):
        raw = event.data.strip()
        paths = re.findall(r'\{([^}]+)\}', raw)
        remaining = re.sub(r'\{[^}]+\}', '', raw).split()
        paths += remaining
        pdf_paths = [p for p in paths if p.lower().endswith('.pdf')]
        if pdf_paths:
            self._add_files(pdf_paths)
        else:
            self._log_msg("PDFファイルをドロップしてください", "warn")

    def _add_files(self, paths):
        added = 0
        for p in paths:
            if p not in self._file_list:
                self._file_list.append(p)
                self._listbox.insert(tk.END, Path(p).name)
                added += 1
        if added:
            n = len(self._file_list)
            self._drop_zone.configure(
                text=f"📄  {n} 件のPDFが追加されています（クリックで追加選択）",
                fg=self.ACCENT)
            self._status.set(f"{n} 件のファイルが登録されています")

    def _remove_selected(self):
        sel = list(self._listbox.curselection())
        for i in reversed(sel):
            self._listbox.delete(i)
            self._file_list.pop(i)
        n = len(self._file_list)
        if n == 0:
            self._reset_drop_zone()
        else:
            self._status.set(f"{n} 件のファイルが登録されています")

    def _clear_files(self):
        self._file_list.clear()
        self._listbox.delete(0, tk.END)
        self._reset_drop_zone()

    def _reset_drop_zone(self):
        self._drop_zone.configure(
            text="📂  ここにPDFをドラッグ＆ドロップ（複数可）\nまたはクリックして選択",
            fg=self.SUBTEXT)
        self._status.set("PDFファイルを追加してください")

    # ── 処理実行 ──────────────────────────────────────────────────────────────
    def _start(self):
        if MISSING:
            self._show_missing()
            return
        if self._processing:
            return
        if not self._file_list:
            messagebox.showwarning("ファイル未登録", "PDFファイルを追加してください。")
            return

        self._processing = True
        self._run_btn.configure(state="disabled", text="処理中…", bg="#333d55")
        self._progress['value'] = 0
        self._progress['maximum'] = len(self._file_list)
        self._status.set(f"処理中… 0 / {len(self._file_list)}")
        self._clear_log()
        self._log_msg(f"処理開始: {len(self._file_list)} 件", "accent")

        thread = threading.Thread(
            target=run_batch,
            args=(
                list(self._file_list),
                lambda msg: self._log_queue.put(("info", msg)),
                lambda cur, tot: self._log_queue.put(("progress", (cur, tot))),
                lambda ok, ng: self._log_queue.put(("done", (ok, ng))),
            ),
            daemon=True
        )
        thread.start()

    # ── ログポーリング ────────────────────────────────────────────────────────
    def _poll_log(self):
        while not self._log_queue.empty():
            kind, payload = self._log_queue.get_nowait()
            if kind == "info":
                tag = ("success" if "完了" in payload or "保存" in payload
                       else "error"   if "エラー" in payload
                       else "warn"    if "警告" in payload or "スキップ" in payload
                       else "accent"  if payload.startswith("[")
                       else "info")
                self._log_msg(payload, tag)
            elif kind == "progress":
                cur, tot = payload
                self._progress['value'] = cur
                self._status.set(f"処理中… {cur} / {tot}")
            elif kind == "done":
                ok_list, ng_list = payload
                self._processing = False
                self._run_btn.configure(state="normal",
                                        text="▶  処理開始", bg=self.ACCENT)
                total = len(ok_list) + len(ng_list)
                if not ng_list:
                    self._status.set(f"✓  完了！  {len(ok_list)} 件すべて成功")
                    self._log_msg("✓  すべて正常に完了しました", "success")
                    messagebox.showinfo("完了",
                        f"{len(ok_list)} 件の処理が完了しました。\n"
                        "各PDFと同じフォルダに _linked.pdf で保存されました。")
                else:
                    self._status.set(
                        f"完了  成功 {len(ok_list)} / 失敗 {len(ng_list)} / 計 {total}")
                    self._log_msg(
                        f"✗  {len(ng_list)} 件でエラーが発生しました", "error")
                    err_detail = "\n".join(f"・{name}" for name, _ in ng_list)
                    messagebox.showwarning("一部エラー",
                        f"成功: {len(ok_list)} 件 / 失敗: {len(ng_list)} 件\n\n"
                        f"失敗したファイル:\n{err_detail}\n\n"
                        "詳細はログを確認してください。")
        self.after(100, self._poll_log)

    # ── ログ操作 ──────────────────────────────────────────────────────────────
    def _log_msg(self, msg, tag="info"):
        self._log.configure(state="normal")
        self._log.insert("end", msg + "\n", tag)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _clear_log(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _show_missing(self):
        libs = "\n".join(f"  pip install {m}" for m in MISSING)
        messagebox.showerror(
            "ライブラリ不足",
            f"以下のライブラリをインストールしてください:\n\n{libs}\n\n"
            "インストール後、アプリを再起動してください。"
        )


# ═══════════════════════════════════════════════════════════════════════════════
#  エントリポイント
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = App()
    app.mainloop()
