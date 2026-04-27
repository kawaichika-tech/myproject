import streamlit as st
import pdfplumber
import io
import re

# ============================================================
# PDF テキスト抽出
# ============================================================
def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=3, y_tolerance=3)
            if text:
                parts.append(text)
            for table in page.extract_tables():
                for row in table:
                    if row and any(c for c in row if c):
                        parts.append(" | ".join([str(c or "").strip() for c in row]))
    return "\n".join(parts)


# ============================================================
# 汎用ヘルパー
# ============================================================
_LABEL_SKIP = {"商品名", "品番", "カラー", "色", "メーカー", "種類", "型番",
               "材料", "張り方向", "備考", "決定", "選択", "メイン", "貼り分け",
               "ー", "-", ""}


def find_value_after(text: str, label: str, stop_chars: int = 200) -> str:
    """ラベルの直後の値を返す。テーブルの '|' 区切りや空セルを考慮して、
    ラベルではない最初の有効値を返す。"""
    idx = text.find(label)
    if idx == -1:
        return ""
    after = text[idx + len(label):idx + len(label) + stop_chars]
    line = after.split("\n")[0]
    parts = [p.strip() for p in line.split("|")]
    for p in parts:
        if p and p not in _LABEL_SKIP and "【" not in p:
            return p
    return ""


def find_value_in_table(text: str, label: str) -> str:
    """ラベルを含むテーブル行を探して値を返す。
    1) 同じ行に '... | label | value | ...' → label の直後の有効セル
    2) 横並びヘッダー '... | label1 | label2 | ...' の場合 → 次の行の同じ列
    3) ラベルの後に同じ行で別ラベルが続く場合 → ラベル位置から数えて同じインデックスの値セル
    """
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if label not in line:
            continue
        if "|" not in line:
            after = line.split(label, 1)[1].lstrip(":：\t 　").strip()
            if after and after not in _LABEL_SKIP and "【" not in after:
                first = re.split(r'[\s　]+', after)[0]
                if first and first not in _LABEL_SKIP:
                    return first
            continue

        cells = [c.strip() for c in line.split("|")]
        label_col = -1
        for j, c in enumerate(cells):
            if c == label:
                label_col = j
                break
        if label_col == -1:
            continue

        # 形式1: ラベル直後のセルが値
        if label_col + 1 < len(cells):
            v = cells[label_col + 1]
            if v and v not in _LABEL_SKIP and "【" not in v:
                return v

        # 形式2: 次の行の同じ列
        if i + 1 < len(lines) and "|" in lines[i + 1]:
            next_cells = [c.strip() for c in lines[i + 1].split("|")]
            if label_col < len(next_cells):
                v = next_cells[label_col]
                if v and v not in _LABEL_SKIP and "【" not in v:
                    return v

        # 形式3: 横並びの複数ラベル（label1 | label2 | label3 | val1 | val2 | val3）
        # ラベルが連続するセル数を数えて、同じ相対位置の値を取る
        label_seq_start = label_col
        while label_seq_start > 0 and cells[label_seq_start - 1] in _LABEL_SKIP - {"", "ー", "-"}:
            label_seq_start -= 1
        label_seq_end = label_col
        while label_seq_end + 1 < len(cells) and cells[label_seq_end + 1] in _LABEL_SKIP - {"", "ー", "-"}:
            label_seq_end += 1
        label_count = label_seq_end - label_seq_start + 1
        offset = label_col - label_seq_start
        # 値は label_seq_end の後ろから始まると仮定し、同じoffsetを取る
        val_idx = label_seq_end + 1 + offset
        if val_idx < len(cells):
            v = cells[val_idx]
            if v and v not in _LABEL_SKIP and "【" not in v:
                return v
    return ""

def search_pattern(text: str, pattern: str) -> str:
    """正規表現で最初のマッチを返す"""
    m = re.search(pattern, text)
    return m.group(1).strip() if m else ""

def find_decided_color(block: str, color_keywords: list) -> str:
    """ブロック内で【決定】に対応するカラーを特定する"""
    lines = block.split("\n")
    # パターン1: 【決定】が含まれる行に直接カラー名がある
    for line in lines:
        if "【決定】" in line:
            for kw in color_keywords:
                if kw in line:
                    return kw
    # パターン2: 【決定】の行の直前の行にカラー名がある
    for i, line in enumerate(lines):
        if "【決定】" in line:
            search_block = "\n".join(lines[max(0, i-5):i+1])
            for kw in color_keywords:
                if kw in search_block:
                    return kw
    return ""

def get_block(text: str, start_kw: str, end_kw: str = None, lines: int = 25) -> str:
    """start_kwからend_kwまで（またはlines行）のブロックを返す"""
    idx = text.find(start_kw)
    if idx == -1:
        return ""
    end_idx = len(text)
    if end_kw:
        e = text.find(end_kw, idx + 1)
        if e != -1:
            end_idx = e
    # 行数制限も適用
    text_lines = text[idx:end_idx].split("\n")
    return "\n".join(text_lines[:lines])


# ============================================================
# チェック本体
# ============================================================
def check_specification(text: str) -> tuple[list, list, dict]:
    errors = []
    passes = []
    meta = {}

    # ========== 基本情報 ==========
    m = re.search(r'【.+?】.+?(?:新築工事|工事)', text)
    meta["物件名"] = m.group(0).strip() if m else "（読み取り不可）"

    m = re.search(r'指定区域\s*(\S+)', text)
    meta["指定区域"] = m.group(1) if m else "（読み取り不可）"
    is_boka = "防火地域" in text
    is_junboka = "準防火地域" in text

    m = re.search(r'風致・景観地区\s*(.+?)(?:\n|$)', text)
    keikan_raw = m.group(1).strip() if m else ""
    meta["景観地区"] = keikan_raw if keikan_raw else "（読み取り不可）"
    is_nishikita = "西北部住宅地" in keikan_raw
    is_fuchi = any(f in text for f in ["第一種風致地区", "第二種風致地区", "第三種風致地区", "第四種風致地区", "第五種風致地区"])

    # ========== 玄関ドア ==========
    door_name = ""
    # 「商品名・種類」ラベルの後から開き戸/引戸を探す
    m = re.search(r'商品名・種類\s*\|?\s*(開き戸\s*\S+|引戸\s*\S+|片開き\s*\S+)', text)
    if m:
        door_name = m.group(1).strip()
    if not door_name:
        # 開き戸＋ドア名パターン
        m = re.search(r'(開き戸\s+(?:ジエスタ\S*|エルムーブ\S*)|引戸\s+\S+)', text)
        if m:
            door_name = m.group(1).strip()
    if not door_name:
        # ジエスタ/エルムーブ単独
        m = re.search(r'(ジエスタ2|ジエスタ|エルムーブ2|エルムーブ)', text)
        if m:
            door_name = m.group(1)

    if not door_name:
        errors.append({"項目": "玄関ドア（商品名・種類）",
                       "ルール": "設計士から引き継いだサイズを玄関ドアの種類部分に記載する",
                       "現状": "記載なし", "理由": "商品名・種類が未記載です"})
    else:
        if is_boka or is_junboka:
            if "防火戸" not in door_name:
                errors.append({"項目": "玄関ドア（防火仕様）",
                               "ルール": "防火・準防火地域の場合「ジエスタ2　防火戸」または「エルムーブ2　防火戸」と記載",
                               "現状": door_name, "理由": "防火仕様の商品名になっていません"})
            else:
                passes.append(f"玄関ドア（商品名・種類）: {door_name}")
        else:
            passes.append(f"玄関ドア（商品名・種類）: {door_name}")

    # ========== 外壁（サイディング）==========
    # メーカー：外壁セクション内の「メーカー」ラベルの後を取得
    wall_maker = ""
    m = re.search(r'(?:外壁|サイディング).{0,200}?メーカー\s*([^\|\n]{5,50})', text, re.DOTALL)
    if m:
        wall_maker = m.group(1).strip().split("|")[0].strip()
        # 長すぎる場合は短く
        if len(wall_maker) > 50:
            wall_maker = wall_maker[:50]
    if not wall_maker:
        for kw in ["KMEW", "ケイミュー", "ニチハ", "旭トステム"]:
            if kw in text:
                idx = text.find(kw)
                wall_maker = text[idx:idx+40].split("\n")[0].split("|")[0].strip()
                break

    if not wall_maker:
        errors.append({"項目": "外壁（メーカー）", "ルール": "メーカー名を記載すること",
                       "現状": "記載なし", "理由": "メーカー名が未記載です"})
    else:
        passes.append(f"外壁（メーカー）: {wall_maker}")

    # 品番："メイン"行 または EW\d+ パターン
    wall_hinban = search_pattern(text, r'(EW\d+[A-Z]*)')
    if not wall_hinban:
        m = re.search(r'メイン\s+\S+\s+(EW\S+|\S{5,})', text)
        wall_hinban = m.group(1) if m else ""

    # 採用カラー："メイン"行から取得（これが【決定】色）
    wall_color = ""
    m = re.search(r'メイン\s+(QW\S+|[^\s\|]{3,30})\s', text)
    if m:
        wall_color = m.group(1).strip()
    if not wall_color:
        # テーブルから探す: メイン | | QWセレノグレージュ EW7532H
        m = re.search(r'メイン\s*\|[^\|]*\|\s*(QW\S+[^\|]{0,30})', text)
        if m:
            wall_color = m.group(1).strip().split("|")[0].strip()

    # 採用色のマンセル値
    wall_mv = ""
    if wall_color:
        mv_m = re.search(rf'{re.escape(wall_color[:6])}.{{0,30}}?(\d\.\d[A-Z]+\s+\d\.?\d*/\d\.?\d*)', text)
        if mv_m:
            wall_mv = mv_m.group(1)

    if not wall_hinban:
        errors.append({"項目": "外壁（品番）", "ルール": "品番を記載すること",
                       "現状": "記載なし", "理由": "品番が未記載です"})
    else:
        passes.append(f"外壁（品番）: {wall_hinban}")

    if not wall_color:
        errors.append({"項目": "外壁（カラー）", "ルール": "カラーを選択・記載すること",
                       "現状": "記載なし", "理由": "採用カラーが特定できません"})
    else:
        display = wall_color
        if wall_mv:
            display += f"（マンセル値: {wall_mv}）"
        passes.append(f"外壁（カラー 決定選択）: {display}")

    # ========== 貼り分け・塗り分けチェック ==========
    # ウィルウォール/WRCの場合は外観指示図面で対応 → 貼り分け図面作成不要
    has_willwall = any(kw in text for kw in ["ウィルウォール", "WRC"])
    if has_willwall:
        passes.append("板張り（ウィルウォール/WRC）: 外観指示図面に指示あり → 貼り分け図面作成不要")

    # 貼り分け指示欄に「塗り分け」が記載されている場合はエラー（同一面の塗り分けは不可）
    m_nuri = re.search(r'(?:貼り分け|指示欄).{0,40}塗り分け', text, re.DOTALL)
    if m_nuri:
        errors.append({"項目": "塗り分け（同一面禁止）",
                       "ルール": "同じ面上での「塗り分け」は不可",
                       "現状": "貼り分け指示欄に塗り分けの記載あり",
                       "理由": "同じ面での塗り分けは不可です。貼り分け方法を再検討してください"})

    # 貼り分け種類を確認
    haritawake_type = ""
    m_type = re.search(r'貼り分け指示欄?\s*[：:\|]?\s*(板張り|貼り分け|塗り分け|ウィルウォール|WRC)', text)
    if m_type:
        haritawake_type = m_type.group(1).strip()

    # 貼り分けの記載チェック
    m = re.search(r'貼り分け\s*([^\|\n]{0,30})', text)
    if m:
        val = m.group(1).strip()
        if "対象外" in val or val == "" or "ー" in val:
            passes.append("外壁貼り分け: 対象外")
        else:
            type_display = f"（種類: {haritawake_type}）" if haritawake_type else ""
            passes.append(f"外壁貼り分け: 記載あり {type_display}")

            # 基本は入隅。出隅で貼り分ける場合はW出隅品番・カラーを貼り分け図面に記載
            if "出隅" in text:
                passes.append("外壁貼り分け（出隅）: W出隅の品番・カラーは貼り分け図面に記載 → 貼り分け図面チェックで確認")
            else:
                passes.append("外壁貼り分け（位置）: 入隅（基本通り）")

            # ウィルウォール/WRCでない場合は貼り分け図面（平面図・立面図）の作成が必要
            if not has_willwall:
                passes.append("外壁貼り分け図面: 平面図・立面図の作成と色付けが必要（目視確認）")

    # 外壁マンセル値チェック（景観地区）
    if (is_nishikita or is_fuchi) and wall_mv:
        mm = re.match(r'([0-9.]+)([A-Z]+)\s+([0-9.]+)/([0-9.]+)', wall_mv)
        if mm:
            value = float(mm.group(3))
            chroma = float(mm.group(4))
            hue = mm.group(2)
            ok, reason = True, ""
            if hue == "Y":
                if value > 8.0:
                    ok, reason = False, f"明度{value}が8.0超のためNG"
                elif value > 5.0 and chroma > 3.0:
                    ok, reason = False, f"明度{value}（5.0超）のとき彩度{chroma}が3.0超のためNG"
                elif value <= 5.0 and chroma > 6.0:
                    ok, reason = False, f"明度{value}（5.0以下）のとき彩度{chroma}が6.0超のためNG"
            if ok:
                passes.append(f"外壁マンセル値 景観基準適合: {wall_mv}")
            else:
                errors.append({"項目": "外壁（マンセル値 景観基準）",
                               "ルール": "外壁の色彩は景観計画色彩基準（基準1-④）に適合すること",
                               "現状": wall_mv, "理由": reason})

    # ========== 屋根・破風 ==========
    # 材料："本屋根"の後を探す
    roof_material = ""
    for kw in ["ガルバリウム鋼板", "ガルバリウム", "コロニアル", "スレート"]:
        if kw in text:
            roof_material = kw
            break
    if not roof_material:
        m = re.search(r'本屋根\s*\|?\s*([^\|\n]{3,20})', text)
        if m:
            roof_material = m.group(1).strip()

    if not roof_material:
        errors.append({"項目": "屋根（メーカー／材料）", "ルール": "メーカーまたは材料名を記載すること",
                       "現状": "記載なし", "理由": "屋根材が未記載です"})
    else:
        passes.append(f"屋根（メーカー／材料）: {roof_material}")

    # 貼り方向
    if "タテヒラ" in text:
        passes.append("屋根（貼り方向）: タテヒラ")
    else:
        errors.append({"項目": "屋根（貼り方向）", "ルール": "貼り方向はタテヒラのみ可",
                       "現状": "タテヒラの記載なし", "理由": "貼り方向の記載が必要です"})

    # 品番と採用カラー：【決定】の前後でK0xxとカラー名を取得
    roof_hinban = ""
    roof_color = ""
    roof_mv = ""
    roof_block = get_block(text, "屋根・破風", "雨樋", 40)
    if not roof_block:
        roof_block = get_block(text, "屋根", "雨樋", 40)

    lines_r = roof_block.split("\n")
    for i, line in enumerate(lines_r):
        if "【決定】" in line:
            search = "\n".join(lines_r[max(0, i-6):i+2])
            hm = re.search(r'(K\d{3})', search)
            if hm:
                roof_hinban = hm.group(1)
            cm = re.search(r'(セピア\w*|ジェットブラック\w*|ブラウン\w*|グレー\w*|ブラック\w*)', search)
            if cm:
                roof_color = cm.group(1)
            mvm = re.search(r'([0-9.]+[A-Z]+\s+[0-9.]+/[0-9.]+)', search)
            if mvm:
                roof_mv = mvm.group(1)
            break

    if not roof_hinban:
        errors.append({"項目": "屋根（品番）", "ルール": "品番を記載すること",
                       "現状": "記載なし", "理由": "品番が未記載です"})
    else:
        passes.append(f"屋根（品番）: {roof_hinban}")

    # 破風カラー
    fufu_note = ""
    if "▼屋根・破風色は共通" in text or "破風色は共通" in text:
        fufu_note = "屋根と共通"

    if roof_color or fufu_note:
        display = roof_color if roof_color else ""
        if roof_mv:
            display += f"（マンセル値: {roof_mv}）"
        if fufu_note:
            display += f" ／ 破風: {fufu_note}"
        passes.append(f"屋根・破風（カラー）: {display}")
    else:
        errors.append({"項目": "屋根・破風（カラー）", "ルール": "カラーの記載が必要",
                       "現状": "記載なし", "理由": "カラーが未記載です"})

    # 屋根マンセル値チェック（景観地区）
    if is_nishikita or is_fuchi:
        if not roof_mv:
            errors.append({"項目": "屋根（マンセル値）",
                           "ルール": "景観・風致地区対象のため屋根のマンセル値の記載と基準適合確認が必要",
                           "現状": "マンセル値の記載なし", "理由": "マンセル値の記載が必要です"})
        else:
            mm = re.match(r'([0-9.]+)([A-Z]+)\s+([0-9.]+)/([0-9.]+)', roof_mv)
            if mm:
                value = float(mm.group(3))
                if value > 4.0:
                    errors.append({"項目": "屋根（マンセル値 景観基準）",
                                   "ルール": "屋根の明度は4.0以下であること（基準1-④）",
                                   "現状": roof_mv, "理由": f"明度{value}が4.0を超えています"})
                else:
                    passes.append(f"屋根マンセル値 景観基準適合: {roof_mv}")

    # ========== 雨樋 ==========
    tooi_block = get_block(text, "雨樋", "サッシ", 30)
    if not tooi_block:
        tooi_block = get_block(text, "軒樋", None, 20)

    # 軒樋品番（KAKU と RK85 が別行になる場合も対応）
    noki_hinban = ""
    # 同一行パターン
    m = re.search(r'KAKU\s*(RK\d+)', text)
    if m:
        noki_hinban = f"KAKU {m.group(1)}"
    elif "KAKU" in text and re.search(r'RK\d+', text):
        rk = re.search(r'(RK\d+)', text)
        noki_hinban = f"KAKU {rk.group(1)}" if rk else "KAKU"

    # 竪樋品番
    tate_hinban = search_pattern(tooi_block, r'(瞬水S\d+)')

    # 軒樋カラー（【決定】の前のカラー）
    noki_color = ""
    tate_color = ""
    lines_t = tooi_block.split("\n")
    decided_count = 0
    for i, line in enumerate(lines_t):
        if "【決定】" in line:
            decided_count += 1
            search = "\n".join(lines_t[max(0, i-4):i+1])
            cm = re.search(r'(しんちゃ|ブラック|ホワイト|ミルク\w*|シルバー|ブラウン)', search)
            if cm:
                if decided_count == 1:
                    noki_color = cm.group(1)
                else:
                    tate_color = cm.group(1)

    hinban_display = " ／ ".join(filter(None, [noki_hinban, tate_hinban]))
    color_display = ""
    if noki_color and tate_color:
        color_display = f"軒樋: {noki_color} ／ 竪樋・呼樋: {tate_color}"
    elif noki_color:
        color_display = noki_color
    elif tate_color:
        color_display = tate_color

    if not hinban_display:
        errors.append({"項目": "雨樋（品番）", "ルール": "軒樋・竪樋・呼樋の品番を記入すること",
                       "現状": "記載なし", "理由": "品番が未記載です"})
    else:
        passes.append(f"雨樋（品番）: {hinban_display}")

    if not color_display:
        errors.append({"項目": "雨樋（カラー）", "ルール": "カラーを記入すること",
                       "現状": "記載なし", "理由": "カラーが未記載です"})
    else:
        passes.append(f"雨樋（カラー 決定選択）: {color_display}")

    # ========== サッシ ==========
    sash_block = get_block(text, "サッシ", None, 15)
    sash_color = ""
    # 「ー 【決定】」パターン：【決定】が2列目 → 2番目のカラーが採用
    # 「【決定】 ー」パターン：【決定】が1列目 → 1番目のカラーが採用
    lines_s = sash_block.split("\n")
    color_candidates = []
    for line in lines_s:
        cm_all = re.findall(r'(ホワイト|ブラック|シルバー|ブロンズ|ゴールド)', line)
        color_candidates.extend(cm_all)
    for i, line in enumerate(lines_s):
        if "【決定】" in line:
            # 【決定】の前にあるカラーを取得（同じ行）
            before_decided = line[:line.find("【決定】")]
            cm = re.search(r'(ホワイト|ブラック|シルバー|ブロンズ|ゴールド)', before_decided)
            if cm:
                sash_color = cm.group(1)
            else:
                # 前の行から探す
                search = "\n".join(lines_s[max(0, i-4):i+1])
                # 「ー | 【決定】」の場合は2番目のカラー
                if line.strip().startswith("ー") or "ー" in line[:line.find("【決定】")]:
                    if len(color_candidates) >= 2:
                        sash_color = color_candidates[1]
                    elif len(color_candidates) >= 1:
                        sash_color = color_candidates[0]
                else:
                    cm = re.search(r'(ホワイト|ブラック|シルバー)', search)
                    if cm:
                        sash_color = cm.group(1)

    if not sash_color:
        errors.append({"項目": "サッシ（カラー）", "ルール": "カラーを記入すること",
                       "現状": "記載なし", "理由": "カラーが未記載です"})
    else:
        passes.append(f"サッシ（カラー 決定選択）: {sash_color}")

    # ========== 内部土間タイル ==========
    # テキスト全体から "内部土間タイル" セクション以降を探す
    tile_block = get_block(text, "内部土間タイル", "巾木", 20)
    if not tile_block:
        tile_block = get_block(text, "内部土間", "巾木", 20)

    # メーカー：既知メーカー名を全文から直接検索（タイル系メーカー）
    tile_maker = ""
    tile_maker_keywords = ["ニッタイ", "名古屋モザイク", "サンワカンパニー", "リビエラ", "ADVAN"]
    for kw in tile_maker_keywords:
        if kw in text:
            tile_maker = kw
            break
    if not tile_maker:
        # 「内部土間タイル」ブロック内の「メーカー」ラベルを探す
        m = re.search(r'内部土間.{0,5}?(?:共通仕様|メーカー).{0,5}?([^\|\n]{2,20})', text)
        if m:
            tile_maker = m.group(1).strip().split("|")[0].strip()
    if not tile_maker:
        # テーブル行でニッタイ等を探す
        for line in text.split("\n"):
            if "メーカー" in line and ("ニッタイ" in line or "モザイク" in line):
                m = re.search(r'メーカー\s*\|?\s*([^\|\n]{2,20})', line)
                if m:
                    tile_maker = m.group(1).strip()
                    break

    # 商品名（テーブル形式優先）
    tile_name = find_value_in_table(tile_block, "商品名") or find_value_after(tile_block, "商品名")
    if not tile_name:
        m = re.search(r'内部土間タイル.{0,300}?商品名[\s\|：:]+([^\|\n]{2,30})', text, re.DOTALL)
        if m:
            cand = m.group(1).strip()
            if cand not in _LABEL_SKIP:
                tile_name = cand

    # 品番（テーブル形式優先）
    tile_hinban = find_value_in_table(tile_block, "品番") or find_value_after(tile_block, "品番")
    if not tile_hinban:
        m = re.search(r'([A-Z]{2,3}-\d{2,3}[-\d]*)', tile_block)
        tile_hinban = m.group(1) if m else ""

    if not tile_maker:
        errors.append({"項目": "内部土間タイル（メーカー名）", "ルール": "①メーカー名を記載すること",
                       "現状": "記載なし", "理由": "メーカー名が未記載です"})
    else:
        passes.append(f"内部土間タイル（メーカー名）: {tile_maker}")

    if not tile_name:
        errors.append({"項目": "内部土間タイル（商品名）", "ルール": "②商品名を記載すること",
                       "現状": "記載なし", "理由": "商品名が未記載です"})
    else:
        passes.append(f"内部土間タイル（商品名）: {tile_name}")

    if not tile_hinban:
        errors.append({"項目": "内部土間タイル（品番）", "ルール": "③品番を記載すること",
                       "現状": "記載なし", "理由": "品番が未記載です"})
    else:
        passes.append(f"内部土間タイル（品番）: {tile_hinban}")

    # ========== 巾木 ==========
    habaki_block = get_block(text, "巾木", None, 12)

    habaki_maker_keywords = [
        "ウッドワン", "WOODONE", "パナソニック", "Panasonic", "大建", "ダイケン",
        "DAIKEN", "永大", "EIDAI", "リクシル", "LIXIL", "ノダ", "NODA",
        "東洋テックス", "朝日ウッドテック", "シンコール", "サンゲツ", "TOTO",
    ]
    habaki_maker = ""
    for kw in habaki_maker_keywords:
        if kw in habaki_block:
            habaki_maker = kw
            break
    # ブロック内に無ければ全文から（巾木がメインの一般的なメーカー優先）
    if not habaki_maker:
        for kw in habaki_maker_keywords:
            if kw in text:
                habaki_maker = kw
                break

    # 巾木行を全部つなげて取得（複数行にまたがるケースに対応）
    habaki_full_lines = [line for line in habaki_block.split("\n") if "巾木" in line and len(line) > 4]
    habaki_full = " ".join(habaki_full_lines) if habaki_full_lines else habaki_block

    # 商品名: テーブルラベル → 既知パターンの順で探す
    habaki_name = find_value_in_table(habaki_block, "商品名")
    if not habaki_name:
        nm = re.search(
            r'(ドレスタ\S*シリーズ\S*|ドレスタ\S+|'
            r'ソフトアートⅡ?巾木\S*|ソフトアート\S*|'
            r'ピノアース\S*巾木\S*|ピノアース\S*|'
            r'コンビット\S*|ジョイハードフロアー\S*|'
            r'スタンダード\S*巾木\S*|MDF巾木\S*|無垢巾木\S*|'
            r'Nカラー巾木\S*|新永大巾木\S*|永大巾木\S*|'
            r'巾木[ⅠⅡⅢⅣⅤ\d]+\S*|'
            r'\S+巾木[A-Z\d]*)',
            habaki_full or habaki_block,
        )
        habaki_name = nm.group(1) if nm else ""

    # カラー: テーブルラベル → 既知パターンの順で探す
    habaki_color = find_value_in_table(habaki_block, "カラー") or find_value_in_table(habaki_block, "色")
    if not habaki_color:
        cm = re.search(
            r'(パールホワイト|ピュアホワイト|ホワイトオーク|ホワイト|'
            r'ブラック|ナチュラル|ブラウン|ダークブラウン|ライトブラウン|'
            r'メープル|チェリー|ウォールナット|オーク|モカ\w*|'
            r'チェスナット|オリーブ\w*|マロン|ベージュ|アイボリー|'
            r'グレー|ライトグレー|ダークグレー|シルバー)\S*色?',
            habaki_full or habaki_block,
        )
        habaki_color = cm.group(0) if cm else ""

    if not habaki_maker:
        errors.append({"項目": "巾木（メーカー）", "ルール": "メーカーを記載すること",
                       "現状": "記載なし", "理由": "メーカーが未記載です"})
    else:
        passes.append(f"巾木（メーカー）: {habaki_maker}")

    if not habaki_name:
        errors.append({"項目": "巾木（商品名）", "ルール": "商品名を記載すること",
                       "現状": "記載なし", "理由": "商品名が未記載です"})
    else:
        passes.append(f"巾木（商品名）: {habaki_name}")

    if not habaki_color:
        errors.append({"項目": "巾木（カラー）", "ルール": "カラーを記載すること",
                       "現状": "記載なし", "理由": "カラーが未記載です"})
    else:
        passes.append(f"巾木（カラー）: {habaki_color}")

    # ========== 内部塗装色 ==========
    # 「【決定】」の直前にある塗装色名を取得
    naito_decided = ""
    m = re.search(r'(クリア塗装|バトンオーク塗装|ウォールナット\S*塗装|\S+塗装)\s*【決定】', text)
    if m:
        naito_decided = m.group(1)
    if not naito_decided:
        # 内部塗装色ブロック内で【決定】の前の塗装名を探す
        naito_block = get_block(text, "内部塗装色", None, 6)
        m = re.search(r'(クリア塗装|バトンオーク塗装|\S+塗装)', naito_block)
        if m and "【決定】" in naito_block:
            naito_decided = m.group(1)

    if naito_decided:
        passes.append(f"内部塗装色（決定）: {naito_decided}")
    else:
        passes.append("内部塗装色: 目視確認が必要")

    # ========== 軒天 ==========
    noki_block = get_block(text, "軒天", None, 8)
    noki_material, noki_color_val = "", ""
    nm = re.search(r'(ウエスタンレッドシダー|レッドシダー|ケイカル板|アルミ|\S+シダー)', noki_block)
    noki_material = nm.group(1) if nm else ""
    cm = re.search(r'(クリア|ホワイト|ナチュラル|ブラウン|グレー)', noki_block)
    noki_color_val = cm.group(1) if cm else ""

    if noki_material or noki_color_val:
        display = " ／ ".join(filter(None, [noki_material, noki_color_val]))
        passes.append(f"軒天: {display}")
    else:
        passes.append("軒天: 目視確認が必要")

    # ========== 化粧柱・木格子 ==========
    # 「化粧柱」の直後のカラーを取得（「ナチュラルホワイト」などの扉色と混同しないよう化粧柱以降を検索）
    kesho_color = ""
    m = re.search(r'化粧柱・木格子\s*[\|]?\s*(オリーブ|ブラック|ブラウン|ホワイト|グレー|ダーク\w*)', text)
    if m:
        kesho_color = m.group(1)
    if not kesho_color:
        # 「化粧柱」の後に「決定」がある行でカラーを探す
        for line in text.split("\n"):
            if "化粧柱" in line and ("決定" in line or "オリーブ" in line):
                cm = re.search(r'化粧柱.{0,30}?(オリーブ|ブラック|ブラウン|グレー)', line)
                if cm:
                    kesho_color = cm.group(1)
                    break

    if kesho_color:
        passes.append(f"化粧柱・木格子（カラー）: {kesho_color}")
    else:
        passes.append("化粧柱・木格子: 目視確認が必要（または対象外）")

    return errors, passes, meta


# ============================================================
# レポート整形
# ============================================================
# ============================================================
# 貼り分け図面チェック
# ============================================================
def check_haritawake(text: str) -> tuple[list, list]:
    errors = []
    passes = []

    # この図面は左右2列レイアウト（外壁メイン｜貼り分け）
    # 商品名・品番カラー・張り方向は1行に両方が並ぶ形式
    # → 各行で「最初の値＝メイン」「2番目の値＝貼り分け」として取得

    # メーカー（全文から）
    makers = re.findall(r'メーカー\s*\|?\s*(KMEW|ケイミュー|ニチハ|旭トステム|エスケー化研)', text)
    main_maker = makers[0] if len(makers) >= 1 else ""
    hare_maker = makers[1] if len(makers) >= 2 else makers[0] if len(makers) == 1 else ""

    # 商品名：全行から「商品名」ラベルの次の値をすべて収集
    main_name, hare_name = "", ""
    name_vals = []
    for line in text.split("\n"):
        if "商品名" not in line:
            continue
        # "|"区切り（テーブル形式）
        if "|" in line:
            parts_line = [p.strip() for p in line.split("|")]
            for i, p in enumerate(parts_line):
                if p == "商品名" and i + 1 < len(parts_line):
                    val = parts_line[i + 1].strip()
                    if val and "商品名" not in val and val not in ["", "ー"]:
                        name_vals.append(val)
        else:
            # スペース区切り（"商品名 xxx 商品名 yyy"）
            found = re.findall(r'商品名\s+([^商\n]+?)(?=\s*商品名|\n|$)', line)
            for v in found:
                v = v.strip()
                if v and len(v) > 2:
                    name_vals.append(v)
    # 重複排除
    seen = []
    for v in name_vals:
        if v not in seen:
            seen.append(v)
    name_vals = seen
    main_name = name_vals[0] if len(name_vals) >= 1 else ""
    hare_name = name_vals[1] if len(name_vals) >= 2 else ""

    # 品番・カラー（QWxxxxパターンを全取得）
    qw_all = re.findall(r'(QW\S+\s+EW\S+|QW\S+)', text)
    main_hinban_color = qw_all[0] if len(qw_all) >= 1 else ""
    hare_hinban_color = qw_all[1] if len(qw_all) >= 2 else ""

    # 張り方向（ヨコ/タテ）
    directions = re.findall(r'張り方向\s*\|?\s*(ヨコ|タテ)', text)
    main_direction = directions[0] if len(directions) >= 1 else ""
    hare_direction = directions[1] if len(directions) >= 2 else directions[0] if len(directions) == 1 else ""

    # --- 外壁メインの判定 ---
    if not main_maker:
        errors.append({"項目": "外壁メイン（メーカー）", "ルール": "メーカー名を記載すること",
                       "現状": "記載なし", "理由": "メーカー名が未記載です"})
    else:
        passes.append(f"外壁メイン（メーカー）: {main_maker}")

    if not main_name:
        errors.append({"項目": "外壁メイン（商品名）", "ルール": "商品名を記載すること",
                       "現状": "記載なし", "理由": "商品名が未記載です"})
    else:
        passes.append(f"外壁メイン（商品名）: {main_name}")

    if not main_hinban_color:
        errors.append({"項目": "外壁メイン（品番・カラー）", "ルール": "品番・カラーを記載すること",
                       "現状": "記載なし", "理由": "品番・カラーが未記載です"})
    else:
        passes.append(f"外壁メイン（品番・カラー）: {main_hinban_color}")

    if not main_direction:
        errors.append({"項目": "外壁メイン（張り方向）", "ルール": "張り方向を記載すること",
                       "現状": "記載なし", "理由": "張り方向が未記載です"})
    else:
        passes.append(f"外壁メイン（張り方向）: {main_direction}")

    # --- 貼り分けの判定 ---
    if not hare_maker:
        errors.append({"項目": "貼り分け（メーカー）", "ルール": "メーカー名を記載すること",
                       "現状": "記載なし", "理由": "メーカー名が未記載です"})
    else:
        passes.append(f"貼り分け（メーカー）: {hare_maker}")

    if not hare_name:
        errors.append({"項目": "貼り分け（商品名）", "ルール": "商品名を記載すること",
                       "現状": "記載なし", "理由": "商品名が未記載です"})
    else:
        passes.append(f"貼り分け（商品名）: {hare_name}")

    if not hare_hinban_color:
        errors.append({"項目": "貼り分け（品番・カラー）", "ルール": "品番・カラーを記載すること",
                       "現状": "記載なし", "理由": "品番・カラーが未記載です"})
    else:
        passes.append(f"貼り分け（品番・カラー）: {hare_hinban_color}")

    if not hare_direction:
        errors.append({"項目": "貼り分け（張り方向）", "ルール": "張り方向を記載すること",
                       "現状": "記載なし", "理由": "張り方向が未記載です"})
    else:
        passes.append(f"貼り分け（張り方向）: {hare_direction}")

    # ========== W出隅 ==========
    w_match = re.search(r'W出隅\s*([^\|\n]{3,40})', text)
    if w_match:
        w_info = w_match.group(1).strip().split("|")[0].strip()
        passes.append(f"W出隅（品番・カラー）: {w_info}")
    elif "出隅" in text:
        # 出隅の品番・カラーを別途探す
        m = re.search(r'(?:出隅|W出隅).{0,5}?([A-Z]\d+\S+).{0,20}?(ブラック|ホワイト|グレー|ブラウン|\S+色)', text)
        if m:
            passes.append(f"W出隅（品番・カラー）: {m.group(1)} {m.group(2)}")
        else:
            errors.append({"項目": "W出隅（品番・カラー）",
                           "ルール": "出隅で張り分けする場合は「W出隅　品番　カラー」の指定をすること",
                           "現状": "品番・カラーの記載なし",
                           "理由": "W出隅の品番・カラーが特定できません"})

    # ========== 貼り分け基本ルール ==========
    # 基本は入隅での貼り分け
    if "入隅" in text and "出隅" not in text:
        passes.append("貼り分け位置: 入隅（基本通り）")
    elif "出隅" in text:
        if "W出隅" in text:
            passes.append("貼り分け位置: 出隅 → W出隅指示あり（上記で確認）")
        else:
            errors.append({"項目": "貼り分け位置（出隅指定）",
                           "ルール": "基本は入隅での貼り分け。出隅で張り分けする場合はどちらに合わせるかを指示すること",
                           "現状": "出隅の記載あり・W出隅指示なし",
                           "理由": "出隅での貼り分けはW出隅の品番・カラー指定が必要です"})

    # 同じ面での塗り分け不可チェック
    if re.search(r'塗り分け', text) and not re.search(r'塗り分け.*?不可|不可.*?塗り分け', text):
        errors.append({"項目": "塗り分け（同一面禁止）",
                       "ルール": "同じ面上での「塗り分け」は不可",
                       "現状": "塗り分けの記載あり",
                       "理由": "同じ面での塗り分けは不可です。貼り分け方法を再検討してください"})

    # 同じ面で貼り分けをする場合の見切り確認（注意事項）
    passes.append("見切り検討: 同じ面で貼り分けをする場合は外壁材の厚みに揃えた見切りを検討すること（段差がある場合はコーキング対応をカネマルさんへ確認）")

    # ========== 平面図・立面図の色付け ==========
    has_plan = "平面図" in text or "パース" in text
    has_elevation = "立面図" in text
    if has_plan and has_elevation:
        passes.append("平面図・立面図: 作成あり（貼り分け部分に色付けがされているか目視確認が必要）")
    else:
        errors.append({"項目": "貼り分け図面（平面図・立面図）",
                       "ルール": "外壁貼り分け図面（平面図・立面図）の作成が必要。貼り分け部分がわかるよう立面図・平面図の該当箇所に色を付けること",
                       "現状": f"平面図: {'あり' if has_plan else 'なし'} / 立面図: {'あり' if has_elevation else 'なし'}",
                       "理由": "図面が不足しています"})

    return errors, passes


def format_haritawake_report(errors: list, passes: list) -> str:
    lines = []
    lines.append("### 📋 外壁貼り分け図面 チェックレポート")
    lines.append("")
    lines.append("#### 🔴 エラー・要確認項目")
    if not errors:
        lines.append("すべてルール通りです")
    else:
        for e in errors:
            lines.append(f"* **{e['項目']}**")
            lines.append(f"  * ルール: {e['ルール']}")
            lines.append(f"  * 現状の記載: {e['現状']}")
            lines.append(f"  * 理由: {e['理由']}")
    lines.append("")
    lines.append("#### 🟢 合格・確認済み項目")
    for p in passes:
        lines.append(f"* {p}")
    lines.append("")
    lines.append("---")
    lines.append("修正が必要な箇所は以上です。修正されたPDFがアップロードされ次第、再度チェックを行います。")
    return "\n".join(lines)


def format_report(errors: list, passes: list, meta: dict) -> str:
    lines = []
    lines.append("### 📋 内外装仕様書 チェックレポート")
    lines.append(f"**対象物件:** {meta.get('物件名', '（読み取り不可）')}")
    lines.append(f"**指定区域:** {meta.get('指定区域', '（読み取り不可）')}")
    lines.append(f"**景観・風致地区:** {meta.get('景観地区', '（読み取り不可）')}")
    lines.append("")
    lines.append("#### 🔴 エラー・要確認項目")
    if not errors:
        lines.append("すべてルール通りです")
    else:
        for e in errors:
            lines.append(f"* **{e['項目']}**")
            lines.append(f"  * ルール: {e['ルール']}")
            lines.append(f"  * 現状の記載: {e['現状']}")
            lines.append(f"  * 理由: {e['理由']}")
    lines.append("")
    lines.append("#### 🟢 合格・確認済み項目")
    for p in passes:
        lines.append(f"* {p}")
    lines.append("")
    lines.append("---")
    lines.append("修正が必要な箇所は以上です。修正されたPDFがアップロードされ次第、再度チェックを行います。")
    return "\n".join(lines)


# ============================================================
# 注文住宅 テーブル解析
# ============================================================
def build_chubun_lookup(pdf_bytes: bytes) -> dict:
    """注文仕様書PDFのテーブル構造から ラベル→値 の辞書を構築する。
    列構成（0始まり）:
      外部側: col0=セクション名, col2=ラベル, col3=サブラベル, col4=値1, col5=値2
      内部側: col8=セクション名, col9=ラベル, col11=値1, col12=値2, col13=値3
    """
    lookup = {}

    # CubePDF等で作成されたPDFに対応するため複数の設定を試みる
    table_strategies = [
        {"vertical_strategy": "lines", "horizontal_strategy": "lines",
         "snap_tolerance": 5, "join_tolerance": 3, "intersection_tolerance": 5},
        {"vertical_strategy": "lines", "horizontal_strategy": "lines",
         "snap_tolerance": 10, "join_tolerance": 5, "intersection_tolerance": 10},
        {"vertical_strategy": "text", "horizontal_strategy": "lines",
         "snap_tolerance": 5, "join_tolerance": 3},
        {"vertical_strategy": "lines", "horizontal_strategy": "text",
         "snap_tolerance": 5, "join_tolerance": 3},
    ]

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            # 物件名は extract_text から取得（新築工事を含む行を優先）
            raw = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
            m = re.search(r'(\S+\s+新築工事)', raw)
            if m and "_物件名" not in lookup:
                lookup["_物件名"] = m.group(0).strip()

            # テーブルを抽出（複数戦略を試みて最も行数が多いものを採用）
            best_tables = []
            for strategy in table_strategies:
                try:
                    tables = page.extract_tables(table_settings=strategy)
                    total_rows = sum(len(t) for t in tables)
                    best_total = sum(len(t) for t in best_tables)
                    if total_rows > best_total:
                        best_tables = tables
                except Exception:
                    continue

            for table in best_tables:
                current_ext_sec = ""
                current_int_sec = ""

                for row in table:
                    if not row or not any(c for c in row if c):
                        continue

                    cells = [str(c or "").strip() for c in row]
                    while len(cells) < 16:
                        cells.append("")

                    # 外部セクション名を更新
                    if cells[0] and cells[0] not in ["外部仕様", "内部仕様", "外部仕様 内部仕様"]:
                        current_ext_sec = cells[0]

                    # 外部項目: col2+col3 → ラベル, col4+col5 → 値
                    ext_l2  = cells[2]
                    ext_l3  = cells[3]
                    ext_val = " ".join(filter(None, [cells[4], cells[5]])).strip()

                    if ext_val:
                        if ext_l2 and ext_l3:
                            lookup[f"{current_ext_sec} {ext_l2} {ext_l3}"] = ext_val
                            if ext_l3 not in lookup:
                                lookup[ext_l3] = ext_val
                        elif ext_l2:
                            lookup[f"{current_ext_sec} {ext_l2}"] = ext_val
                        elif ext_l3:
                            lookup[f"{current_ext_sec} {ext_l3}"] = ext_val
                            if ext_l3 not in lookup:
                                lookup[ext_l3] = ext_val

                    # 内部項目: col8 → セクション名, col9 → ラベル, col10 → サブラベル, col11+col12+col13 → 値
                    int_sec_raw = cells[8]
                    int_label   = cells[9]
                    int_sublbl  = cells[10]
                    int_val     = " ".join(filter(None, [cells[11], cells[12], cells[13]])).strip()

                    # セクション名はマージセルで空になる場合があるので持続させる
                    if int_sec_raw:
                        current_int_sec = int_sec_raw

                    # col9が空でcol10に値がある場合（ラベルがサブラベル列にある構造）に対応
                    eff_label  = int_label or int_sublbl
                    eff_sublbl = int_sublbl if int_label else ""

                    if eff_label and int_val:
                        # セクション+ラベルで保存
                        if current_int_sec:
                            lookup[f"{current_int_sec} {eff_label}"] = int_val
                        # ラベル+サブラベルで保存
                        if eff_sublbl:
                            key_sub = f"{eff_label} {eff_sublbl}"
                            if key_sub not in lookup:
                                lookup[key_sub] = int_val
                        # ラベル単体でも保存（先着優先）
                        if eff_label not in lookup:
                            lookup[eff_label] = int_val

    return lookup


# ============================================================
# 注文住宅 チェック本体
# ============================================================
def check_specification_chubun(lookup: dict, config: dict) -> tuple[list, list, dict]:
    errors = []
    passes = []
    meta = {}

    meta["物件名"] = lookup.get("_物件名", "（読み取り不可）")

    is_boka         = config.get("is_boka_area", False)
    wall_type       = config.get("wall_type", "サイディング")
    has_haritawake  = config.get("has_haritawake", False)
    has_parapet     = config.get("has_parapet", False)
    has_keshow_hari = config.get("has_keshow_hari", False)
    is_nonstandard  = config.get("is_nonstandard_siding", False)
    roof_type       = config.get("roof_type", "ガルバ")
    has_shutter     = config.get("has_garage_shutter", False)

    def g(key):
        return lookup.get(key, "").strip()

    def g_partial(section_kw: str, label: str) -> str:
        for k, v in lookup.items():
            if section_kw in k and label in k:
                return v.strip()
        return ""

    # ========== 玄関ドア ==========
    door_name    = g("玄関ドア 商品名・種類")
    door_model   = g("玄関ドア 型番")
    door_color   = g("玄関ドア カラー")
    outer_handle = g("玄関ドア 外：把手デザイン/カラー")
    inner_handle = g("玄関ドア 内：把手デザイン/カラー")
    shikai       = g("玄関ドア 施解錠")

    if not door_name:
        errors.append({"項目": "玄関ドア（商品名・種類）", "ルール": "商品名・種類を記載すること",
                       "現状": "記載なし", "理由": "商品名・種類が未記載です"})
    elif is_boka and "防火戸" not in door_name:
        errors.append({"項目": "玄関ドア（防火仕様）",
                       "ルール": "防火・準防火地域の場合「防火戸」仕様にすること",
                       "現状": door_name, "理由": "防火仕様の商品名になっていません"})
    else:
        passes.append(f"玄関ドア（商品名・種類）: {door_name}")

    for label, val in [("型番", door_model), ("カラー", door_color),
                       ("外把手", outer_handle), ("内把手", inner_handle), ("施解錠", shikai)]:
        if not val:
            errors.append({"項目": f"玄関ドア（{label}）",
                           "ルール": f"{label}を記載すること",
                           "現状": "記載なし", "理由": f"{label}が未記載です"})
        else:
            passes.append(f"玄関ドア（{label}）: {val}")

    # ========== 外壁（サイディング）==========
    if wall_type in ["サイディング", "両方"]:
        sid_maker  = g("外壁(サイディング) メーカー")
        sid_name   = g("外壁(サイディング) 商品名")
        sid_hinban = g("外壁(サイディング) カラー・品番")
        sid_dir    = g("外壁(サイディング) 張り方向")

        for label, val in [("メーカー", sid_maker), ("商品名", sid_name),
                           ("カラー・品番", sid_hinban), ("張り方向", sid_dir)]:
            if not val:
                errors.append({"項目": f"外壁サイディング（{label}）",
                               "ルール": f"{label}を記載すること",
                               "現状": "記載なし", "理由": f"{label}が未記載です"})
            else:
                passes.append(f"外壁サイディング（{label}）: {val}")

        if is_nonstandard:
            passes.append("外壁サイディング: 標準外サイディング使用（別途承認確認が必要）")

    # ========== 外壁（塗り壁）==========
    if wall_type in ["塗り壁", "両方"]:
        paint_name    = g("外壁(塗り壁） 商品名")
        paint_pattern = g("外壁(塗り壁） パターン")
        paint_color   = g("外壁(塗り壁） カラー")

        for label, val in [("商品名", paint_name), ("パターン", paint_pattern), ("カラー", paint_color)]:
            if not val:
                errors.append({"項目": f"外壁塗り壁（{label}）",
                               "ルール": f"{label}を記載すること",
                               "現状": "記載なし", "理由": f"{label}が未記載です"})
            else:
                passes.append(f"外壁塗り壁（{label}）: {val}")

    # ========== 塗り分け／貼り分け ==========
    if has_haritawake:
        hari_umu      = g("塗り分け／貼り分け 有無")
        hari_material = g("板貼り 材種")

        if hari_umu:
            passes.append(f"塗り分け/貼り分け（有無）: {hari_umu}")
        else:
            errors.append({"項目": "塗り分け/貼り分け（有無）", "ルール": "有無を記載すること",
                           "現状": "記載なし", "理由": "貼り分けの有無が未記載です"})
        if hari_material:
            passes.append(f"板貼り（材種）: {hari_material}")
        else:
            errors.append({"項目": "板貼り（材種）", "ルール": "貼り分けがある場合は材種を記載すること",
                           "現状": "記載なし", "理由": "材種が未記載です"})
    else:
        passes.append("塗り分け/貼り分け: 対象外")

    # ========== 屋根・破風 ==========
    roof_material = g("本屋根")
    roof_dir      = g("屋根・破風 張り方向") or g("張り方向")
    roof_color    = g("屋根・破風 屋根 色・品番") or g("屋根 色・品番")
    fufu_color    = g("屋根・破風 破風 色・品番") or g("破風 色・品番")

    if not roof_material:
        errors.append({"項目": "屋根（材料）", "ルール": "本屋根の材料を記載すること",
                       "現状": "記載なし", "理由": "屋根材が未記載です"})
    else:
        passes.append(f"屋根（材料）: {roof_material}")

    if roof_type == "ガルバ":
        if not roof_dir:
            errors.append({"項目": "屋根（張り方向）",
                           "ルール": "ガルバの場合は張り方向（タテヒラ）を記載すること",
                           "現状": "記載なし", "理由": "張り方向が未記載です"})
        elif roof_dir != "タテヒラ":
            errors.append({"項目": "屋根（張り方向）", "ルール": "貼り方向はタテヒラのみ可",
                           "現状": roof_dir, "理由": "タテヒラ以外の方向が記載されています"})
        else:
            passes.append(f"屋根（張り方向）: {roof_dir}")

    if not roof_color:
        errors.append({"項目": "屋根（色・品番）", "ルール": "屋根の色・品番を記載すること",
                       "現状": "記載なし", "理由": "色・品番が未記載です"})
    else:
        passes.append(f"屋根（色・品番）: {roof_color}")

    if not fufu_color:
        errors.append({"項目": "破風（色・品番）", "ルール": "破風の色・品番を記載すること",
                       "現状": "記載なし", "理由": "破風の色・品番が未記載です"})
    else:
        passes.append(f"破風（色・品番）: {fufu_color}")

    if roof_type == "瓦":
        for label, key in [("破風色記載", "屋根・破風 瓦屋根の場合(破風色記載)"),
                           ("雨押さえ",   "屋根・破風 瓦屋根の場合(雨押さえ)")]:
            val = g(key)
            if val:
                passes.append(f"瓦屋根（{label}）: {val}")
            else:
                errors.append({"項目": f"瓦屋根（{label}）",
                               "ルール": f"瓦屋根の場合は{label}を記載すること",
                               "現状": "記載なし", "理由": f"{label}が未記載です"})

    # ========== 雨樋 ==========
    noki_info = g("雨樋 軒樋")
    tate_info = g("雨樋 竪樋・呼樋")

    if not noki_info:
        errors.append({"項目": "雨樋（軒樋）", "ルール": "軒樋の品番・カラーを記載すること",
                       "現状": "記載なし", "理由": "軒樋が未記載です"})
    else:
        passes.append(f"雨樋（軒樋）: {noki_info}")

    if not tate_info:
        errors.append({"項目": "雨樋（竪樋・呼樋）", "ルール": "竪樋・呼樋の品番・カラーを記載すること",
                       "現状": "記載なし", "理由": "竪樋・呼樋が未記載です"})
    else:
        passes.append(f"雨樋（竪樋・呼樋）: {tate_info}")

    # ========== 外部サッシ色 ==========
    sash_color = (g("外部サッシ色・勝手口 色") or g("外部サッシ色 色")
                  or g_partial("外部サッシ色", "色"))
    if not sash_color:
        errors.append({"項目": "外部サッシ色", "ルール": "外部サッシ色を記載すること",
                       "現状": "記載なし", "理由": "外部サッシ色が未記載です"})
    else:
        passes.append(f"外部サッシ色: {sash_color}")

    water_color = g("外部水切り 色")
    passes.append(f"外部水切り: {water_color}" if water_color else "外部水切り: 目視確認が必要")

    # ========== パラペット笠木 ==========
    if has_parapet:
        parapet_color = g("パラペット笠木 色")
        if not parapet_color:
            errors.append({"項目": "パラペット笠木（色）", "ルール": "パラペット笠木の色を記載すること",
                           "現状": "記載なし", "理由": "パラペット笠木の色が未記載です"})
        else:
            passes.append(f"パラペット笠木（色）: {parapet_color}")
    else:
        passes.append("パラペット笠木: 対象外")

    # ========== 外部化粧梁 ==========
    if has_keshow_hari:
        if any("化粧梁" in k for k in lookup):
            passes.append("外部化粧梁: 記載あり（目視確認が必要）")
        else:
            errors.append({"項目": "外部化粧梁", "ルール": "外部化粧梁の仕様を記載すること",
                           "現状": "記載なし", "理由": "化粧梁の記載が見当たりません"})
    else:
        passes.append("外部化粧梁: 対象外")

    # ========== ガレージシャッター ==========
    if has_shutter:
        if any("シャッター" in k or "ガレージ" in k for k in lookup):
            passes.append("ガレージシャッター: 記載あり（目視確認が必要）")
        else:
            errors.append({"項目": "ガレージシャッター", "ルール": "ガレージシャッターの仕様を記載すること",
                           "現状": "記載なし", "理由": "ガレージシャッターの記載が見当たりません"})
    else:
        passes.append("ガレージシャッター: 対象外")

    # ========== 内部土間仕上げ ==========
    tile_maker  = g_partial("内部土間仕上げ", "メーカー") or g("メーカー")
    tile_name   = g_partial("内部土間仕上げ", "商品名") or g("商品名")
    tile_hinban = g_partial("内部土間仕上げ", "品番") or g("品番")

    for label, val in [("メーカー", tile_maker), ("商品名", tile_name), ("品番", tile_hinban)]:
        if not val:
            errors.append({"項目": f"内部土間仕上げ（{label}）",
                           "ルール": f"{label}を記載すること",
                           "現状": "記載なし", "理由": f"{label}が未記載です"})
        else:
            passes.append(f"内部土間仕上げ（{label}）: {val}")

    # ========== 内部サッシ色 ==========
    naibusasshi_color = (g("内部サッシ色 色") or g("内部サッシ色")
                         or g_partial("内部サッシ色", "色"))
    if not naibusasshi_color:
        errors.append({"項目": "内部サッシ色", "ルール": "内部サッシ色を記載すること",
                       "現状": "記載なし", "理由": "内部サッシ色が未記載です"})
    else:
        passes.append(f"内部サッシ色: {naibusasshi_color}")

    # ========== 巾木 ==========
    habaki_hinban = g("巾木 品番") or g_partial("巾木", "品番")
    habaki_color  = g("巾木 色") or g_partial("巾木", "色")
    if not habaki_hinban:
        errors.append({"項目": "巾木（品番）", "ルール": "巾木の品番を記載すること",
                       "現状": "記載なし", "理由": "巾木の品番が未記載です"})
    else:
        passes.append(f"巾木（品番）: {habaki_hinban}")
    if not habaki_color:
        errors.append({"項目": "巾木（色）", "ルール": "巾木の色を記載すること",
                       "現状": "記載なし", "理由": "巾木の色が未記載です"})
    else:
        passes.append(f"巾木（色）: {habaki_color}")

    # ========== 内部塗装色 ==========
    for label in ["化粧柱", "棚・カウンター・笠木", "造作洗面化粧台・手洗い", "階段(上がり框）", "框 ロイヤル仕上げ"]:
        val = g(label)
        disp_label = label.replace("框 ロイヤル仕上げ", "框ロイヤル仕上げ")
        passes.append(f"内部塗装色（{disp_label}）: {val}" if val
                      else f"内部塗装色（{disp_label}）: 目視確認が必要")

    # ========== 外部塗装色 ==========
    noki_wrc = g("軒裏仕上げ(WRC)")
    passes.append(f"外部塗装色（軒裏仕上げWRC）: {noki_wrc}" if noki_wrc
                  else "外部塗装色（軒裏仕上げWRC）: 目視確認が必要")

    tanchiku = g("単独柱・木格子")
    passes.append(f"外部塗装色（単独柱・木格子）: {tanchiku}" if tanchiku
                  else "外部塗装色（単独柱・木格子）: 目視確認が必要")

    # ========== 枠仕様 ==========
    for room in ["基本", "和室", "脱衣室", "お風呂扉", "玄関ドア"]:
        val = g(room)
        passes.append(f"枠仕様（{room}）: {val}" if val
                      else f"枠仕様（{room}）: 目視確認が必要")

    return errors, passes, meta


def format_report_chubun(errors: list, passes: list, meta: dict) -> str:
    lines = []
    lines.append("### 📋 内外装仕様書 チェックレポート（注文住宅）")
    lines.append(f"**対象物件:** {meta.get('物件名', '（読み取り不可）')}")
    lines.append("")
    lines.append("#### 🔴 エラー・要確認項目")
    if not errors:
        lines.append("すべてルール通りです")
    else:
        for e in errors:
            lines.append(f"* **{e['項目']}**")
            lines.append(f"  * ルール: {e['ルール']}")
            lines.append(f"  * 現状の記載: {e['現状']}")
            lines.append(f"  * 理由: {e['理由']}")
    lines.append("")
    lines.append("#### 🟢 合格・確認済み項目")
    for p in passes:
        lines.append(f"* {p}")
    lines.append("")
    lines.append("---")
    lines.append("修正が必要な箇所は以上です。修正されたPDFがアップロードされ次第、再度チェックを行います。")
    return "\n".join(lines)


# ============================================================
# Streamlit UI
# ============================================================
st.set_page_config(page_title="内外装図面チェッカー", page_icon="🏠", layout="wide")
st.title("🏠 内外装図面チェッカー")
st.caption("アイニコグループ株式会社 ｜ 外部・内部仕様一覧表 自動チェックシステム")
st.divider()

with st.sidebar:
    st.markdown("""
    **【分譲】チェック対象項目（仕様書）**
    - ✅ 玄関ドア（商品名・種類）
    - ✅ 外壁（メーカー・品番・カラー）
    - ✅ 外壁貼り分け条件
    - ✅ 屋根・破風（材料・方向・品番・カラー）
    - ✅ 雨樋（品番・カラー）
    - ✅ サッシ（カラー）
    - ✅ 内部土間タイル（メーカー・商品名・品番）
    - ✅ 巾木（メーカー・商品名・カラー）
    - ✅ 内部塗装色（決定色）
    - ✅ 軒天（材料・カラー）
    - ✅ 化粧柱・木格子（カラー）
    - ✅ 景観・風致地区 マンセル値適合確認

    **【分譲】チェック対象項目（貼り分け図面）**
    - ✅ 外壁メイン（メーカー・商品名・品番カラー・張り方向）
    - ✅ 貼り分け（メーカー・商品名・品番カラー・張り方向）
    - ✅ W出隅（品番・カラー）
    - ✅ 平面図・立面図の作成確認

    **【注文】チェック対象項目**
    - ✅ 玄関ドア（商品名・型番・カラー・把手・施解錠）
    - ✅ 外壁サイディング / 塗り壁（条件選択）
    - ✅ 塗り分け/貼り分け（条件選択）
    - ✅ 屋根・破風（ガルバ/瓦 条件選択）
    - ✅ 雨樋（品番・カラー）
    - ✅ 外部サッシ色・水切り
    - ✅ パラペット笠木（条件選択）
    - ✅ 外部化粧梁（条件選択）
    - ✅ ガレージシャッター（条件選択）
    - ✅ 内部サッシ色
    - ✅ 内部土間仕上げ（メーカー・商品名・品番）
    - ✅ 巾木（品番・色）・内部塗装色・枠仕様
    """)

tab_bunjou, tab_chubun = st.tabs(["📋 分譲", "🏠 注文"])

# ============================================================
# 分譲タブ
# ============================================================
with tab_bunjou:
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("📂 ① 外部・内部仕様一覧表")
        uploaded_file = st.file_uploader("仕様書PDFをドラッグ＆ドロップ", type=["pdf"], key="spec")

        if uploaded_file:
            st.success(f"✅ {uploaded_file.name}")

        st.subheader("📂 ② 外壁貼り分け図面（貼り分けがある場合のみ）")
        uploaded_hare = st.file_uploader("貼り分け図面PDFをドラッグ＆ドロップ", type=["pdf"], key="hare")

        if uploaded_hare:
            st.success(f"✅ {uploaded_hare.name}")

        st.divider()
        debug_mode_bunjou = st.checkbox(
            "🔧 デバッグモード（抽出テキストを表示）", key="debug_mode_bunjou"
        )

        if uploaded_file:
            if st.button("🔍 チェック開始", type="primary", use_container_width=True, key="btn_bunjou"):
                pdf_bytes = uploaded_file.read()
                with st.spinner("📄 仕様書を解析中..."):
                    text = extract_text_from_pdf(pdf_bytes)

                if st.session_state.get("debug_mode_bunjou"):
                    with st.expander("🔧 デバッグ: 抽出された全テキスト", expanded=True):
                        st.text(f"総文字数: {len(text)}")
                        st.code(text, language="text")
                        # 巾木周辺のみをピンポイント表示
                        for kw in ["内部土間タイル", "内部土間", "巾木"]:
                            idx = text.find(kw)
                            if idx != -1:
                                start = max(0, idx - 50)
                                end = min(len(text), idx + 600)
                                st.markdown(f"**「{kw}」周辺（位置 {idx}）:**")
                                st.code(text[start:end], language="text")

                if not text.strip():
                    st.error("仕様書PDFからテキストを抽出できませんでした。")
                else:
                    errors, passes, meta = check_specification(text)
                    result = format_report(errors, passes, meta)

                    hare_result = ""
                    if uploaded_hare:
                        hare_bytes = uploaded_hare.read()
                        with st.spinner("📄 貼り分け図面を解析中..."):
                            hare_text = extract_text_from_pdf(hare_bytes)
                        if hare_text.strip():
                            h_errors, h_passes = check_haritawake(hare_text)
                            hare_result = "\n\n" + format_haritawake_report(h_errors, h_passes)
                            if h_errors:
                                st.warning(f"⚠️ 貼り分け図面に{len(h_errors)}件のエラーがあります")

                    full_result = result + hare_result
                    st.session_state["result"] = full_result
                    st.session_state["filename"] = uploaded_file.name

                    if errors:
                        st.warning(f"⚠️ 仕様書に{len(errors)}件のエラーがあります")
                    else:
                        st.success("✅ 仕様書はすべてルール通りです")
        else:
            st.info("① の仕様書PDFをアップロードしてください")

    with col2:
        st.subheader("📋 チェックレポート")
        if "result" in st.session_state:
            st.markdown(st.session_state["result"])
            st.divider()
            filename_base = st.session_state["filename"].replace(".pdf", "")
            st.download_button(label="📥 レポートをダウンロード", data=st.session_state["result"],
                               file_name=f"チェックレポート_{filename_base}.txt",
                               mime="text/plain", use_container_width=True)
        else:
            st.info("PDFをアップロードして「チェック開始」を押してください")

# ============================================================
# 注文タブ
# ============================================================
with tab_chubun:
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("⚙️ 物件条件の設定")

        wall_type_sel = st.radio(
            "外壁タイプ",
            ["サイディング", "塗り壁", "両方（サイディング＋塗り壁）"],
            horizontal=True,
        )
        wall_type_val = "両方" if "両方" in wall_type_sel else wall_type_sel

        col_a, col_b = st.columns(2)
        with col_a:
            has_haritawake  = st.checkbox("外壁の貼り分けあり")
            has_parapet     = st.checkbox("パラペットあり")
            has_keshow_hari = st.checkbox("外部化粧梁あり")
            is_nonstandard  = st.checkbox("標準外サイディング使用")
        with col_b:
            roof_type_sel   = st.radio("屋根タイプ", ["ガルバリウム鋼板", "瓦"], horizontal=True)
            roof_type_val   = "瓦" if roof_type_sel == "瓦" else "ガルバ"
            has_shutter     = st.checkbox("ガレージシャッターあり")
            is_boka_area    = st.checkbox("防火・準防火地域")

        st.divider()
        debug_mode = st.checkbox("🔧 デバッグモード（テーブル構造を表示）", key="debug_mode")
        st.subheader("📂 外部・内部仕様一覧表（注文）")
        uploaded_chubun = st.file_uploader(
            "仕様書PDFをドラッグ＆ドロップ", type=["pdf"], key="chubun"
        )

        if uploaded_chubun:
            st.success(f"✅ {uploaded_chubun.name}")

        st.divider()

        if uploaded_chubun:
            if st.button("🔍 チェック開始", type="primary", use_container_width=True, key="btn_chubun"):
                config = {
                    "wall_type":             wall_type_val,
                    "has_haritawake":        has_haritawake,
                    "has_parapet":           has_parapet,
                    "has_keshow_hari":       has_keshow_hari,
                    "is_nonstandard_siding": is_nonstandard,
                    "roof_type":             roof_type_val,
                    "has_garage_shutter":    has_shutter,
                    "is_boka_area":          is_boka_area,
                }

                pdf_bytes = uploaded_chubun.read()
                st.info(f"PDFサイズ: {len(pdf_bytes):,} バイト")
                with st.spinner("📄 仕様書を解析中..."):
                    try:
                        lookup = build_chubun_lookup(pdf_bytes)
                    except Exception as _build_err:
                        st.error(f"解析例外: {_build_err}")
                        lookup = {}

                if st.session_state.get("debug_mode"):
                    st.subheader("🔧 デバッグ: lookup辞書の内容")
                    st.text(f"lookup件数: {len(lookup)}")
                    for k, v in sorted(lookup.items()):
                        st.text(f"{k!r}: {v!r}")
                    st.divider()
                    st.subheader("🔧 デバッグ: 生テーブル行（先頭3ページ）")
                    import io as _io
                    import pdfplumber as _pp
                    with _pp.open(_io.BytesIO(pdf_bytes)) as _pdf:
                        for _pi, _page in enumerate(_pdf.pages[:3]):
                            page_text = _page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                            st.text(f"--- ページ{_pi+1} テキスト先頭200文字 ---")
                            st.text(page_text[:200])
                            tables = _page.extract_tables()
                            st.text(f"テーブル数: {len(tables)}")
                            for _ti, _table in enumerate(tables):
                                st.text(f"  テーブル{_ti+1}: {len(_table)}行")
                                for _row in _table[:20]:
                                    if _row and any(c for c in _row if c):
                                        st.text(str([str(c or "").strip() for c in _row]))

                if not lookup:
                    st.error("【v3】PDFからデータを取得できませんでした。")
                    try:
                        with pdfplumber.open(io.BytesIO(pdf_bytes)) as _pdf2:
                            for _pi2, _page2 in enumerate(_pdf2.pages[:2]):
                                raw2 = _page2.extract_text(x_tolerance=3, y_tolerance=3) or ""
                                tbls2 = _page2.extract_tables() or []
                                st.info(f"ページ{_pi2+1}: テキスト{len(raw2)}文字 / テーブル{len(tbls2)}個")
                                if raw2:
                                    st.text(f"テキスト先頭: {raw2[:200]}")
                                for _ti2, _t2 in enumerate(tbls2[:2]):
                                    st.text(f"テーブル{_ti2+1}: {len(_t2)}行")
                                    for _r2 in (_t2 or [])[:5]:
                                        if _r2 and any(c for c in _r2 if c):
                                            st.text(str([str(c or "").strip() for c in _r2]))
                    except Exception as _e2:
                        st.error(f"診断エラー: {_e2}")
                else:

                    errors, passes, meta = check_specification_chubun(lookup, config)
                    result = format_report_chubun(errors, passes, meta)

                    st.session_state["chubun_result"]   = result
                    st.session_state["chubun_filename"] = uploaded_chubun.name

                    if errors:
                        st.warning(f"⚠️ {len(errors)}件のエラーがあります")
                    else:
                        st.success("✅ すべてルール通りです")
        else:
            st.info("仕様書PDFをアップロードしてください")

    with col2:
        st.subheader("📋 チェックレポート")
        if "chubun_result" in st.session_state:
            st.markdown(st.session_state["chubun_result"])
            st.divider()
            filename_base = st.session_state["chubun_filename"].replace(".pdf", "")
            st.download_button(
                label="📥 レポートをダウンロード",
                data=st.session_state["chubun_result"],
                file_name=f"チェックレポート_注文_{filename_base}.txt",
                mime="text/plain",
                use_container_width=True,
            )
        else:
            st.info("PDFをアップロードして「チェック開始」を押してください")
