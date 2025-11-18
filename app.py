import streamlit as st
import pandas as pd
from datetime import datetime, date
import io
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import re
import google.generativeai as genai
import os
from textwrap import dedent
from typing import Union

# =========================
# 設定（APIキーは任意）
# =========================
API_KEY = st.secrets["GEMINI_API_KEY"]  # 本来ならば環境変数を使った方が良いがプロトタイプアプリのため直書きをしている
if API_KEY:
    genai.configure(api_key=API_KEY)

# =========================
# ユーティリティ
# =========================
WORK_PROCESS_MAP = {
    "1": "調査分析、要件定義", "2": "基本（外部）設計", "3": "詳細（内部）設計",
    "4": "コーディング・単体テスト", "5": "IT・ST", "6": "システム運用・保守",
    "7": "サーバー構築・運用管理", "8": "DB構築・運用管理", "9": "ネットワーク運用保守",
    "10": "ヘルプ・サポート", "11": "その他"
}

def safe_str(v) -> str:
    """NaN/NaT/Noneを空に、文字列はtrimし 'nan'/'NaT' も空にする"""
    if v is None or (isinstance(v, float) and pd.isna(v)) or pd.isna(v):
        return ""
    s = str(v).strip()
    if s.lower() in ("nan", "nat"):
        return ""
    return s

def to_str_df(df: pd.DataFrame) -> pd.DataFrame:
    return df.fillna("").astype(str)

def find_first(df_str: pd.DataFrame, keyword: str):
    for r in range(df_str.shape[0]):
        row = df_str.iloc[r]
        for c in range(df_str.shape[1]):
            if keyword in row.iloc[c]:
                return r, c
    return None

def next_right_nonempty(df: pd.DataFrame, r: int, c: int, max_look: int = 20):
    for dc in range(1, max_look + 1):
        cc = c + dc
        if cc >= df.shape[1]:
            break
        v = df.iloc[r, cc]
        s = safe_str(v)
        if s:
            return s
    return ""

def parse_date_like(v) -> Union[date, None]:
    if isinstance(v, (pd.Timestamp, datetime)):
        try:
            return v.date()
        except Exception:
            return None

    if isinstance(v, (int, float)):
        try:                
            # Excelのシリアル値（1900/1/1ベース）として変換を試みる
            # '1899-12-30' はExcelの1900年閏年バグを考慮した起点
            temp_date = pd.to_datetime(v, unit='D', origin='1899-12-30')
            
            return temp_date.date()
        except Exception:
            pass # シリアル値でなかった場合は、下の文字列処理へ
    
    s = safe_str(v)
    if not s:
        return None
    # yyyy/mm/dd, yyyy-mm, yyyy/mm, yyyy.mm を緩く拾う
    # 日が無い場合は1日扱い
    m = re.search(r"(\d{4})[./-](\d{1,2})(?:[./-](\d{1,2}))?", s)
    if not m:
        return None
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3) or 1)
    try:
        return date(y, mo, d)
    except Exception:
        return None

def looks_like_proc_codes(s: str) -> bool:
    #st.write("中身:", s)
    #st.warning(bool(re.fullmatch(r"[0-9０-９.．,､、~〜～ 　]+", s.strip())))
    return bool(re.fullmatch(r"[0-9０-９.．,､、~〜～ 　]+", s.strip()))

def pick_first_nonempty(values):
    for v in values:
        s = safe_str(v)
        if s:
            return s
    return ""

def uniq_join(lines):
    out = []
    seen = set()
    for x in lines:
        s = safe_str(x)
        if not s:
            continue
        s = s.strip()
        # 箇条書き記号の除去
        if s.startswith(("・","-","—","―","–")):
            s = s.lstrip("・-—―– ").strip()
        if s not in seen:
            out.append(s)
            seen.add(s)
    return " / ".join(out)

# =========================
# 解析: シート選択＆読み取り
# =========================
LABELS_LEFT = ["フリガナ", "氏名", "現住所", "最寄駅", "最終学歴"]
LABELS_RIGHT = ["生年月日", "性別", "稼働可能日"]

def choose_best_sheet(xl: pd.ExcelFile) -> pd.DataFrame:
    best_df, best_score = None, -1
    for sh in xl.sheet_names:
        df = xl.parse(sh, header=None, dtype=object)
        df_str = to_str_df(df)
        score = 0
        for k in LABELS_LEFT + LABELS_RIGHT + ["情報処理資格", "項番", "作業期間", "案件名", "案件名称", "作業内容"]:
            if find_first(df_str, k):
                score += 1
        if score > best_score:
            best_df, best_score = df, score
    return best_df

def _collect_rightward_values(df: pd.DataFrame, r: int, c: int, max_cols: int = 12) -> list[str]:
    """行rの列cの右側に連続する非空セルを収集。空セルはスキップ可だが、値が一度も出ない場合は [] を返す。"""
    vals = []
    empties_seen = 0
    for dc in range(1, max_cols + 1):
        cc = c + dc
        if cc >= df.shape[1]:
            break
        s = safe_str(df.iloc[r, cc])
        if s:
            vals.append(s)
            empties_seen = 0
        else:
            empties_seen += 1
            # 空セルが2～3個連続したら打ち切り（適度に早期終了）
            if empties_seen >= 3 and vals:
                break
    return vals

def read_personal(df: pd.DataFrame):
    df_str = to_str_df(df)
    # 左側
    result = {
        "furigana": "", "name": "", "address": "",
        "station": "", "education": "",
        "birth": date(2000,1,1), "gender": "未選択",
        "available": datetime.now().date(),
        "qualification": ""
    }
    for k in LABELS_LEFT:
        pos = find_first(df_str, k)
        if pos:
            r, c = pos
            result_map = {
                "フリガナ": "furigana",
                "氏名": "name",
                "現住所": "address",
                "最寄駅": "station",
                "最終学歴": "education",
            }
            result[result_map[k]] = safe_str(next_right_nonempty(df, r, c, 3))
    # 右側
    # 生年月日
    pos = find_first(df_str, "生年月日")
    if pos:
        r, c = pos
        b = parse_date_like(next_right_nonempty(df, r, c, 20))
        result["birth"] = b or date(2000,1,1)
    # 性別
    pos = find_first(df_str, "性別")
    if pos:
        r, c = pos
        g = safe_str(next_right_nonempty(df, r, c, 20))
        if g in ["男", "男性"]:
            result["gender"] = "男性"
        elif g in ["女", "女性"]:
            result["gender"] = "女性"
        elif g == "その他":
            result["gender"] = "その他"
        else:
            result["gender"] = "未選択"
    # 稼働可能日
    pos = find_first(df_str, "稼働可能日")
    if pos:
        r, c = pos
        s = safe_str(next_right_nonempty(df, r, c, 20))
        if ("即日" in s) or (s in ["-", "--", ""]):
            result["available"] = datetime.now().date()
        else:
            d = parse_date_like(s)
            result["available"] = d or datetime.now().date()
    # 資格（行内の右方向をすべて収集。なければ数行下もスキャン）
    pos = find_first(df_str, "情報処理資格")
    if pos:
        r, c = pos
        vals = _collect_rightward_values(df, r, c, max_cols=12)
        if not vals:
            # 行内に見つからない場合は、下方向（次の5行）で右側の値を探索
            for rr in range(r+1, min(r+6, df.shape[0])):
                vals.extend(_collect_rightward_values(df, rr, c, max_cols=12))
        result["qualification"] = "\n".join([safe_str(v) for v in vals if safe_str(v)])
    return result

def find_header_row(df_str: pd.DataFrame) -> Union[int, None]:
    for r in range(df_str.shape[0]):
        row_vals = [df_str.iloc[r, c] for c in range(df_str.shape[1])]
        cond1 = any("項" in v for v in row_vals) and any("作業期間" in v for v in row_vals)
        cond2 = any(("案件名" in v) or ("案件名称" in v) for v in row_vals)
        cond3 = any("作業内容" in v for v in row_vals)
        if cond1 and cond2 and cond3:
            return r
    return None

def col_at(df_str: pd.DataFrame, r: int, keywords: list[str]) -> Union[int, None]:
    for c in range(df_str.shape[1]):
        cell = df_str.iloc[r, c]
        if any(k in cell for k in keywords):
            return c
    return None

def col_at_multiheader(df_str: pd.DataFrame, rows: list[int], keywords: list[str]) -> Union[int, None]:
    """ヘッダ行とサブヘッダ行の両方を見て最初に一致した列を返す"""
    for r in rows:
        c = col_at(df_str, r, keywords)
        if c is not None:
            return c
    return None

def parse_projects(df: pd.DataFrame) -> list:
    df_str = to_str_df(df)
    header_r = find_header_row(df_str)
    if header_r is None:
        return []

    subheader_r = header_r + 1 if header_r + 1 < df.shape[0] else header_r

    # 列位置の推定（ヘッダ行＋サブヘッダ行を考慮）
    C_ID = col_at_multiheader(df_str, [header_r, subheader_r], ["項番", "項"])
    C_PERIOD = col_at_multiheader(df_str, [header_r, subheader_r], ["作業期間", "期間"])
    C_NAME = col_at_multiheader(df_str, [header_r, subheader_r], ["案件名", "案件名称", "案件"])
    C_CONTENT = col_at_multiheader(df_str, [header_r, subheader_r], ["作業内容", "内容"])
    C_OS = col_at_multiheader(df_str, [header_r, subheader_r], ["OS"])
    C_LANG = col_at_multiheader(df_str, [header_r, subheader_r], ["言語", "ツール"])
    C_DB = col_at_multiheader(df_str, [header_r, subheader_r], ["DB", "DB/DC", "DC"])
    C_PROC = col_at_multiheader(df_str, [header_r, subheader_r], ["作業工程", "工程"])
    C_ROLE = col_at_multiheader(df_str, [header_r, subheader_r], ["役割"])
    C_POS = col_at_multiheader(df_str, [header_r, subheader_r], ["ポジション", "役職"])
    C_SCALE = col_at_multiheader(df_str, [header_r, subheader_r], ["規模", "人数"])

    # 必須
    if any(x is None for x in [C_PERIOD, C_NAME, C_CONTENT]):
        return []

    projects = []
    cur = None

    def cell(r, cidx):
        return safe_str(df.iloc[r, cidx]) if (cidx is not None and cidx < df.shape[1]) else ""

    def flush_cur():
        nonlocal cur
        if not cur:
            return
        # 期間→開始/終了推定
        dates = []
        for s in cur["periods"]:
            d = parse_date_like(s)
            if d:
                dates.append(d)
        start_date = min(dates) if dates else None
        end_date = max(dates) if dates else None
        # 「現」「現在」対策
        txt_all = " ".join(cur["periods"])
        if re.search(r"(現|現在)", txt_all):
            end_date = datetime.now().date()

        # 作業工程（番号→ラベル）
        proc_labels = []
        for s in cur["procs"]:
            s_raw = s.strip()
            
            if looks_like_proc_codes(s_raw):
                s_normalized = s_raw.translate(str.maketrans({
                    # 全角数字 -> 半角数字
                    '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
                    '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
                    # 全角記号 -> 半角または統一記号
                    '．': '.',  # 全角ドット -> 半角ドット
                    '､': ',',  # 全角カンマ -> 半角カンマ
                    '、': ',',  # 読点 -> 半角カンマ
                    '，': ',',   # 全角カンマ(FF0C) -> 半角カンマ
                    '～': '〜',  # 全角チルダ -> 波ダッシュ(範囲記号として統一)
                    '~': '〜',  # 半角チルダ -> 波ダッシュ
                }))

                final_codes = [] 
                
                parts = re.split(r"[.,]+", s_normalized)
                for part in parts:
                    part = part.strip()
                    if not part:
                        continue
                    
                    range_match = re.search(r"^(\d+)\s*〜\s*(\d+)$", part) 
                    
                    if range_match:
                        try:
                            start = int(range_match.group(1)) 
                            end = int(range_match.group(2))   
                            for i in range(start, end + 1):
                                final_codes.append(str(i))
                        except ValueError:
                            pass 
                    else:
                        # 範囲でない場合（単なる数字）
                        if re.fullmatch(r"\d+", part):
                            final_codes.append(part)
                
                for k in [x for x in final_codes if x]:
                    if k in WORK_PROCESS_MAP and WORK_PROCESS_MAP[k] not in proc_labels:
                        proc_labels.append(WORK_PROCESS_MAP[k])
                             
                #for k in [x for x in re.split(r"[.,]+", s2) if x]:
                #    if k in WORK_PROCESS_MAP and WORK_PROCESS_MAP[k] not in proc_labels:
                #        proc_labels.append(WORK_PROCESS_MAP[k])

        projects.append({
            "start_date": start_date or date(2000,1,1),
            "end_date": end_date or datetime.now().date(),
            "project_name": cur["project_name"] or "",
            "industry": cur["industry"] or "",
            "work_content": "\n".join([s for s in cur["contents"] if s]),
            "os": uniq_join(cur["oss"]),
            "db_dc": uniq_join(cur["dbs"]),
            "lang_tool": uniq_join(cur["langs"]),
            "work_process_list": proc_labels,
            "work_process_str": ", ".join(proc_labels),
            "role": cur["role"] or "",
            "position": cur["position"] or "",
            "scale": cur["scale"] or "",
        })

    for r in range(subheader_r + 1, df.shape[0]):
        idv = cell(r, C_ID)
        is_new = bool(re.search(r"\d", idv))  # 数字が入っていれば新案件

        if is_new:
            if cur:
                flush_cur()
            cur = {
                "project_name": None,
                "industry": None,
                "periods": [],
                "contents": [],
                "oss": [],
                "dbs": [],
                "langs": [],
                "procs": [],
                "role": "",
                "position": "",
                "scale": "",
            }

        if not cur:
            continue  # まだヘッダ直下の空行など

        # 基本セル
        period_val = cell(r, C_PERIOD)
        cur["periods"].append(period_val)
        is_firstline = bool(parse_date_like(period_val))  # 案件1行目かどうか

        name_val = cell(r, C_NAME)
        if name_val:
            if is_firstline and cur["project_name"] is None:
                cur["project_name"] = name_val
            elif (not is_firstline) and cur["industry"] is None:
                # 「案件名称の真下」を業種として拾う
                cur["industry"] = name_val
            elif cur["project_name"] is None:
                cur["project_name"] = name_val  # フォールバック

        content_val = cell(r, C_CONTENT)
        if content_val:
            cur["contents"].append(content_val)

        os_val = cell(r, C_OS)
        if os_val:
            cur["oss"].append(os_val)

        lang_val = cell(r, C_LANG)
        if lang_val:
            for t in re.split(r"[、,/\n]+", lang_val):
                t = t.strip().lstrip("-・").strip()
                if t:
                    cur["langs"].append(t)

        db_val = cell(r, C_DB)
        if db_val:
            for t in re.split(r"[、,/\n]+", db_val):
                t = t.strip().lstrip("-・").strip()
                if t:
                    cur["dbs"].append(t)

        proc_val = cell(r, C_PROC)
        if proc_val and is_firstline:
            cur["procs"].append(proc_val)

        role_val = cell(r, C_ROLE)
        if role_val and (not is_firstline) and not cur["role"]:
            cur["role"] = role_val

        pos_val = cell(r, C_POS)
        if pos_val and (not is_firstline) and not cur["position"]:
            cur["position"] = pos_val

        scale_val = cell(r, C_SCALE)
        if scale_val and is_firstline and not cur["scale"]:
            cur["scale"] = scale_val

    if cur:
        flush_cur()

    # 空案件を除去（名称も内容も空）
    projects = [p for p in projects if (p["project_name"] or p["work_content"])]
    return projects

# =========================
# Session 初期化
# =========================
def initialize_session_state():
    ss = st.session_state
    ss.setdefault("pi_name", "")
    ss.setdefault("pi_furigana", "")
    ss.setdefault("pi_birth_date", date(2000,1,1))
    ss.setdefault("pi_gender", "未選択")
    ss.setdefault("pi_address", "")
    ss.setdefault("pi_nearest_station", "")
    ss.setdefault("pi_education", "")
    ss.setdefault("pi_available_date", datetime.now().date())
    ss.setdefault("pi_qualifications_input", "")
    ss.setdefault("pi_summary", "")
    ss.setdefault("projects", [])
    ss.setdefault("generated_overview", "")

initialize_session_state()

# =========================
# コールバック
# =========================
def load_from_excel_callback():
    uploaded_file = st.session_state.excel_uploader
    if uploaded_file is None:
        return
    try:
        xl = pd.ExcelFile(uploaded_file)
        df = choose_best_sheet(xl)
        if df is None:
            st.error("有効なシートが見つかりませんでした。")
            return

        # --- 個人情報＆資格 ---
        pi = read_personal(df)
        st.session_state.pi_furigana = pi["furigana"]
        st.session_state.pi_name = pi["name"]
        st.session_state.pi_address = pi["address"]
        st.session_state.pi_nearest_station = pi["station"]
        st.session_state.pi_education = pi["education"]
        st.session_state.pi_birth_date = pi["birth"]
        st.session_state.pi_gender = pi["gender"]
        st.session_state.pi_available_date = pi["available"]
        st.session_state.pi_qualifications_input = pi["qualification"]

        # --- 業務経歴 ---
        st.session_state.projects = parse_projects(df)

        st.success("Excelの内容を入力欄へ反映しました。")

    except Exception as e:
        st.error(f"読み込み中にエラー: {e}")

def enhance_with_ai_callback():
    if not API_KEY:
        st.warning("Gemini APIキーが未設定のためスキップしました。")
        return
    try:
        model = genai.GenerativeModel("gemini-2.5-flash")
        # サマリ
        prompt1 = dedent("""
            あなたは経験豊富なキャリアアドバイザーです。以下の「開発経験サマリ」を、
            簡潔で専門的な表現に整えてください。出力は修正後の本文のみ。
        """) + "\n" + st.session_state.pi_summary
        st.session_state.pi_summary = model.generate_content(prompt1).text

        # 各案件
        for i, p in enumerate(st.session_state.projects):
            if p.get("work_content"):
                prompt2 = dedent("""
                    あなたは経験豊富なキャリアアドバイザーです。以下の「作業内容」を、
                    実績が伝わる箇条書きに整えてください。出力は本文のみ。
                """) + "\n" + p["work_content"]
                st.session_state.projects[i]["work_content"] = model.generate_content(prompt2).text
        st.success("AIで文章を整形しました。")
    except Exception as e:
        st.error(f"AI処理でエラー: {e}")

def generate_overview_callback():
    try:
        skills = set()
        for p in st.session_state.projects:
            if p.get("os"): skills.update([s.strip() for s in str(p["os"]).split("/") if s.strip()])
            if p.get("db_dc"): skills.update([s.strip() for s in str(p["db_dc"]).split("/") if s.strip()])
            if p.get("lang_tool"): skills.update([s.strip() for s in str(p["lang_tool"]).split("/") if s.strip()])
            if p.get("work_process_list"): skills.update(p["work_process_list"])
        all_work = "\n".join([str(p.get("work_content","")) for p in st.session_state.projects])
        remarks = ""
        if API_KEY:
            model = genai.GenerativeModel("gemini-2.5-flash")
            prompt = dedent("""
                以下の作業内容を要約し、備考として1～2文で出力してください。出力は本文のみ。
            """) + "\n" + all_work
            remarks = model.generate_content(prompt).text
        age = (datetime.now().date() - st.session_state.pi_birth_date).days // 365
        gender_str = {"未選択":"", "男性":"男", "女性":"女", "その他":""}.get(st.session_state.pi_gender, "")

        lines = [
            f"氏名\t:{st.session_state.pi_name}　{age}歳　{gender_str}",
            f"最寄\t:{st.session_state.pi_nearest_station}",
            "開始\t:即日可～",
            "単価\t:",
            f"スキル\t:{', '.join(sorted(list(skills)))}",
            f"資格\t:{st.session_state.pi_qualifications_input.replace(chr(10), ', ')}",
            f"備考\t:{remarks}"
        ]
        overview_text = "\n".join(lines)

        st.session_state.generated_overview = overview_text.strip()
        st.success("概要を作成しました。")
    except Exception as e:
        st.error(f"概要作成エラー: {e}")

# =========================
# UI
# =========================
st.set_page_config(page_title="スキルシート自動入力＆Gemini要約アプリ", layout="centered")
st.title("スキルシート自動入力＆Gemini要約アプリ")
st.caption("経歴書Excelファイルをアップロードしてください")

uploaded_file = st.file_uploader(
    "Excelファイル（.xlsx推奨）",
    type=["xlsx", "csv"],
    key="excel_uploader",
    on_change=load_from_excel_callback
)

st.header("個人情報")
cols = st.columns(2)
with cols[0]:
    st.session_state.pi_furigana = st.text_input("フリガナ", st.session_state.pi_furigana)
    st.session_state.pi_name = st.text_input("氏名", st.session_state.pi_name)
    st.session_state.pi_address = st.text_input("現住所", st.session_state.pi_address)
    st.session_state.pi_nearest_station = st.text_input("最寄駅", st.session_state.pi_nearest_station)
with cols[1]:
    st.session_state.pi_birth_date = st.date_input("生年月日", st.session_state.pi_birth_date)
    st.session_state.pi_gender = st.selectbox("性別", ["未選択","男性","女性","その他"], index=["未選択","男性","女性","その他"].index(st.session_state.pi_gender))
    st.session_state.pi_available_date = st.date_input("稼働可能日", st.session_state.pi_available_date)
    st.session_state.pi_education = st.text_input("最終学歴", st.session_state.pi_education)

st.subheader("情報処理資格")
st.session_state.pi_qualifications_input = st.text_area("（1行1資格）", value=st.session_state.pi_qualifications_input, height=100)

st.subheader("開発経験サマリ")
st.session_state.pi_summary = st.text_area("自由記述", value=st.session_state.pi_summary, height=120)

st.header("業務経歴")
if st.button("新しい案件を追加"):
    st.session_state.projects.append({})
for i, p in enumerate(st.session_state.projects):
    st.subheader(f"案件 {i+1}")
    cols = st.columns(2)
    with cols[0]:
        p["start_date"] = st.date_input(f"開始日 (案件 {i+1})", p.get("start_date", date(2022,4,1)))
        p["end_date"] = st.date_input(f"終了日 (案件 {i+1})", p.get("end_date", datetime.now().date()))
        p["project_name"] = st.text_input(f"案件名称 (案件 {i+1})", p.get("project_name",""))
        p["industry"] = st.text_input(f"業種 (案件 {i+1})", p.get("industry",""))
    with cols[1]:
        p["os"] = st.text_input(f"OS (案件 {i+1})", p.get("os",""))
        p["db_dc"] = st.text_input(f"DB/DC (案件 {i+1})", p.get("db_dc",""))
        p["lang_tool"] = st.text_input(f"言語/ツール (案件 {i+1})", p.get("lang_tool",""))
        p["role"] = st.text_input(f"役割 (案件 {i+1})", p.get("role",""))
        p["position"] = st.text_input(f"ポジション (案件 {i+1})", p.get("position",""))
        p["scale"] = st.text_input(f"規模 (案件 {i+1})", p.get("scale",""))
    p["work_content"] = st.text_area(f"作業内容 (案件 {i+1})", p.get("work_content",""))
    selected = st.multiselect(
        f"作業工程 (案件 {i+1})",
        options=list(WORK_PROCESS_MAP.keys()),
        format_func=lambda k: WORK_PROCESS_MAP[k],
        default=[k for k, v in WORK_PROCESS_MAP.items() if v in p.get("work_process_list", [])]
    )
    p["work_process_list"] = [WORK_PROCESS_MAP[k] for k in selected]
    p["work_process_str"] = ", ".join(p["work_process_list"])
    if st.button(f"この案件を削除 (案件 {i+1})"):
        st.session_state.projects.pop(i)
        st.rerun()
    st.markdown("---")

st.header("生成AIによるスキルシート改善")
st.button("生成AIに改善を依頼", on_click=enhance_with_ai_callback)

st.header("スキルシート概要の抽出")
st.button("概要を抽出", on_click=generate_overview_callback)
if st.session_state.generated_overview:
    st.text_area("抽出された概要", value=st.session_state.generated_overview, height=240)

# ---- Excel出力（添付ファイル形式にレイアウト変更） ----
if st.button("スキルシートを生成 (Excel形式)"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wb = writer.book
        if "Sheet" in wb.sheetnames:
             wb.remove(wb["Sheet"])
        ws = wb.create_sheet("スキルシート")
        wb.active = ws

        # --- スタイル定義 ---
        # (添付ファイル形式に合わせて、背景色などを調整)
        title_font = Font(size=18, bold=True, color="000080")
        section_title_font = Font(bold=True, size=12) # 背景色なし
        work_history_font = Font(bold=False, size=9) # 背景色なし
        numbering_font = Font(bold=False, size=8)
        bold_font = Font(bold=True, size=10)
        data_font = Font(bold=False, size=10)
        # 業務経歴テーブルヘッダ用の背景色
        project_title_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        # 罫線
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        dashdot_border = Border(left=Side(style='dashDot'), right=Side(style='dashDot'), top=None, bottom=None)
        # 折り返し + 上寄せ
        wrap_text_alignment = Alignment(wrapText=True, vertical='top')

        # テーブルの列数（K列まで）
        TABLE_COLS = 11
        # 作業内容を書き込む列 (C列)
        COL_PROJECT_NAME = 5

        cur = 1 # 現在の行番号

        # --- 1行目: 空白 ---
        # ユーザーの指示「A行は何も入れず開けておいてください」を「1行目」と解釈
        cur += 1 # 2行目からスタート

        # --- ヘルパー関数 ---
        def style(cell, font=None, fill=None, border=None, align=None):
            if font: cell.font = font
            if fill: cell.fill = fill
            if border: cell.border = border
            if align: cell.alignment = align

        # --- 2行目: タイトル ---
        cell = ws.cell(row=cur, column=2, value="業務経歴書")
        style(cell, font=title_font, border=thin_border)
        ws.merge_cells('B2:K3')
        cur += 2 # 3行目は空白、4行目から

        # --- 4行目: 1. 個人情報 ---
        rows = [
            ("フリガナ", st.session_state.pi_furigana, "生年月日", st.session_state.pi_birth_date.strftime("%Y/%m/%d")),
            ("氏名", st.session_state.pi_name, "性別", st.session_state.pi_gender),
            ("現住所", st.session_state.pi_address, "稼働可能日", st.session_state.pi_available_date.strftime("%Y/%m/%d")),
            ("最寄駅", st.session_state.pi_nearest_station),
            ("最終学歴", st.session_state.pi_education, ),
        ]
        count = 0
        for row in rows:
            style(ws.cell(row=cur, column=2, value=row[0]), font=bold_font, border=thin_border)
            style(ws.cell(row=cur, column=4, value=row[1]), border=thin_border)

            if len(row) == 4 :
                style(ws.cell(row=cur, column=7, value=row[2]), font=bold_font, border=thin_border)
                
                if count == 0:
                    style(ws.cell(row=cur, column=9, value=row[3]), border=thin_border)
                    count = 1
                else:
                    style(ws.cell(row=cur, column=10, value=row[3]), border=thin_border)
            cur += 1

        # フリガナ
        ws.merge_cells(start_row=4, start_column=2, end_row=4, end_column=3)
        ws.merge_cells(start_row=4, start_column=4, end_row=4, end_column=6)

        # 生年月日
        ws.merge_cells(start_row=4, start_column=7, end_row=4, end_column=8)
        ws.merge_cells(start_row=4, start_column=9, end_row=4, end_column=10)
        
        # 氏名
        ws.merge_cells(start_row=5, start_column=2, end_row=5, end_column=3)
        ws.merge_cells(start_row=5, start_column=4, end_row=5, end_column=6)

        # 性別
        ws.merge_cells(start_row=5, start_column=7, end_row=5, end_column=9)
        ws.merge_cells(start_row=5, start_column=10, end_row=5, end_column=TABLE_COLS)
        
        # 現住所
        ws.merge_cells(start_row=6, start_column=2, end_row=6, end_column=3)
        ws.merge_cells(start_row=6, start_column=4, end_row=6, end_column=6)

        # 稼働可能日
        ws.merge_cells(start_row=6, start_column=7, end_row=6, end_column=9)
        ws.merge_cells(start_row=6, start_column=10, end_row=6, end_column=TABLE_COLS)
        
        # 最寄駅
        ws.merge_cells(start_row=7, start_column=2, end_row=7, end_column=3)
        ws.merge_cells(start_row=7, start_column=4, end_row=7, end_column=6)
        ws.merge_cells(start_row=7, start_column=7, end_row=7, end_column=TABLE_COLS)

        # 最終学歴
        ws.merge_cells(start_row=8, start_column=2, end_row=8, end_column=3)
        ws.merge_cells(start_row=8, start_column=4, end_row=8, end_column=TABLE_COLS)
        
        
        # --- 9行目: 2. 資格 ---        
        qlist = [q.strip() for q in st.session_state.pi_qualifications_input.split("\n") if q.strip()]
        if not qlist: qlist = [""]
        
        for q in qlist:
            style(ws.cell(row=cur, column=2, value="情報処理資格"), font=bold_font, border=thin_border)
            cell = ws.cell(row=cur, column=4, value=f"{q}")
            # 資格欄はテーブル幅(K列)まで結合
            ws.merge_cells(start_row=9, start_column=2, end_row=9, end_column=3)
            ws.merge_cells(start_row=cur, start_column=4, end_row=cur, end_column=TABLE_COLS)
            cur += 1
        cur += 1 # 空白行

        # --- 11行目: 開発経験サマリ ---
        cell = ws.cell(row=cur, column=2, value="開発経験サマリ")
        style(cell, font=section_title_font, border=thin_border)
        cur += 1
        
        # サマリ本文もテーブル幅(K列)まで結合
        ws.merge_cells(start_row=cur, start_column=2, end_row=cur, end_column=TABLE_COLS)
        style(ws.cell(row=cur, column=2, value=st.session_state.pi_summary), border=thin_border, align=wrap_text_alignment)
        cur += 2 # 空白行を1つ挟む

        # --- 13行目: 作業工程・役割 ---
        style(ws.cell(row=cur, column=2, value="業務内容：１．調査分析、要件定義  ２：基本（外部）設計　３．詳細（内部）設計　４．ｺｰﾃﾞｨﾝｸﾞ・単体ﾃｽﾄ　5．ＩＴ・ＳＴ"), font=numbering_font)
        cur += 1

        style(ws.cell(row=cur, column=2, value="6.システム運用・保守　7．サーバー構築・運用管理　8．DB構築・運用管理　9．ネットワーク運用保守　10．ヘルプ・サポート 11．その他"), font=numbering_font)
        cur += 1

        style(ws.cell(row=cur, column=2, value="参加形態：（PM)プロジェクトマネージャー　（PL)プロジェクトリーダー　（SPL）サブリーダー　（SE）システムエンジニア （PG）プログラマー"), font=numbering_font)
        cur += 2 # 空白行を1つ挟む
        
        # --- 17行目: 4. 業務経歴 ---
        cell = ws.cell(row=cur, column=2, value="業務経歴")
        style(cell, font=section_title_font, border=thin_border)
        cur += 1

        # --- 18行目: 業務経歴テーブルヘッダ ---
        headers = [
            "項番", "作業期間", "案件名", "作業内容", "機種", "言語/ツール", "作業工程", "規模",
            "業種", "OS", "DB/DC", "役割", "ポジション"
        ] # B列からK列

        targets = [
            (0,0), (0,1), (0,2), (0,3), (0,5), (0,6), (0,8), (0,9),
            (1,2), (1,5), (1,6), (1,8), (1,9)
        ]

        for i, (row, col) in enumerate(targets):
            if i < len(headers):
                cell = ws.cell(row=cur + row, column=col + 2, value=headers[i])
                style(cell, font=bold_font, fill=project_title_fill, border=thin_border, align=wrap_text_alignment)

        # フリガナ
        ws.merge_cells(start_row=cur, start_column=2, end_row=cur + 1, end_column=2)
        ws.merge_cells(start_row=cur, start_column=3, end_row=cur + 1, end_column=3)
        ws.merge_cells(start_row=cur, start_column=5, end_row=cur + 1, end_column=6)
        ws.merge_cells(start_row=cur, start_column=8, end_row=cur, end_column=9)
        ws.merge_cells(start_row=cur + 1, start_column=8, end_row=cur + 1, end_column=9)
        
        cur += 2

        # --- 21行目以降: 案件ループ (テーブル形式) ---
        for i, p in enumerate(st.session_state.projects):
            start_row = cur # この案件の開始行を記憶

            # 1列目書き込み
            cell = ws.cell(row=start_row, column=2, value=i + 1)
                
            # 1列目は全列に罫線と折り返し、上寄せ
            style(cell, font=work_history_font, border=thin_border, align=wrap_text_alignment)

            # --- 2行目 (作業期間) ---
            start_date_str = p.get("start_date").strftime("%Y/%m/%d") if p.get("start_date") else ""
            end_date_str = p.get("end_date").strftime("%Y/%m/%d") if p.get("end_date") else ""
            delta_txt = ""
            if p.get("start_date") and p.get("end_date"):
                days = (p["end_date"] - p["start_date"]).days
                delta_txt = f"(約{round(days/30.4375,1)}ヶ月)" if days >= 0 else "（0ヶ月）"
            
            style(ws.cell(row=start_row, column=3, value=start_date_str),font=data_font, border=thin_border)
            style(ws.cell(row=start_row + 1, column=3, value="～"),font=data_font, border=thin_border)
            style(ws.cell(row=start_row + 2, column=3, value=end_date_str),font=data_font, border=thin_border)
            style(ws.cell(row=start_row + 3, column=3, value=delta_txt),font=data_font, border=thin_border)

            # --- 3行目 (案件名・業種) ---
            style(ws.cell(row=start_row, column=4, value=p.get("project_name","")), font=work_history_font)
            style(ws.cell(row=start_row + 1, column=4, value=p.get("industry","")), font=work_history_font)
            
            # --- 4行目 (作業内容) ---
            content_lines = [line.strip() for line in str(p.get("work_content", "")).split("\n") if line.strip()]
            if not content_lines:
                content_lines = [""]

             # 空でも4行は確保
            if len(content_lines) < 4:
                padding_needed = 4 - len(content_lines)
                content_lines.extend([""] * padding_needed)
                
            content_count = 0
            
            for line in content_lines:
                # C列 (案件名の真下) に作業内容を書き込む
                cell = ws.cell(row=cur, column=COL_PROJECT_NAME, value=line)
                style(cell, border=dashdot_border, align=wrap_text_alignment)
                
                # 作業内容セルを横に結合 (C列からK列まで)
                ws.merge_cells(start_row=cur, start_column=COL_PROJECT_NAME, end_row=cur, end_column=COL_PROJECT_NAME + 1)
                ws.merge_cells(start_row=cur, start_column=COL_PROJECT_NAME + 3, end_row=cur, end_column=COL_PROJECT_NAME + 4)
                
                cur += 1 # 次の行へ
                content_count += 1

            # --- 7行目 (機種・OS) ---
            os = [s.strip() for s in p.get("os", "").split("/") if s.strip()]
            
            for model in range(len(os)):
                style(ws.cell(row=start_row + model, column=7, value=os[model]), font=work_history_font)
            
            # --- 8行目 (言語/ツール・DB/DC) ---
            lang_tool = [s.strip() for s in p.get("lang_tool", "").split("/") if s.strip()]
            db_dc = [s.strip() for s in p.get("db_dc", "").split("/") if s.strip()]
            
            lang_count = 0
            db_count = 0
            
            for lang in range(len(lang_tool)):
                style(ws.cell(row=start_row + lang, column=8, value=lang_tool[lang]), font=work_history_font)
                lang_count += 1

            if lang_tool != db_dc:
                for db in range(len(db_dc)):
                    style(ws.cell(row=start_row + db + (lang_count + 1), column=8, value=db_dc[db]), font=work_history_font)
                    db_count += 1

            #st.write("変更前:", cur, lang_count, db_count, content_count, lang_count + db_count - content_count, cur + lang_count + db_count - content_count)

            # 空でも4行は確保
            for j in range(8): 
                if (lang_count + db_count - content_count) < 4:
                    if (lang_count + db_count - content_count) < 0:
                        lang_count += (lang_count + db_count - content_count) * -1
                    else:
                        lang_count -= lang_count + db_count - content_count
                else:
                    break
                    
            #st.write("変更後:", cur, lang_count, db_count, content_count, lang_count + db_count - content_count, cur + lang_count + db_count - content_count)
            
            cur += lang_count + db_count - content_count
            
            # --- 10行目 (作業工程・役割) ---
            REVERSE_WORK_PROCESS_MAP = {v: k for k, v in WORK_PROCESS_MAP.items()}
            number_list = []
            
            for label in p.get("work_process_list", []):
                # 逆引きマップに存在するか確認
                if label in REVERSE_WORK_PROCESS_MAP:
                    # 存在すれば番号 (e.g. "1") を追加
                    number_list.append(REVERSE_WORK_PROCESS_MAP[label])

            try:
                number_list.sort(key=int)
            except ValueError:
                number_list.sort() # 万が一数値以外が混じっていたら通常ソート

            style(ws.cell(row=start_row, column=10, value=",".join(number_list)), font=work_history_font)
            style(ws.cell(row=start_row+ 1, column=10, value=p.get("role","")), font=work_history_font)
            
            # --- 11行目 (規模・ポジション) ---
            style(ws.cell(row=start_row, column=TABLE_COLS, value=p.get("scale","")), font=work_history_font)
            style(ws.cell(row=start_row + 1, column=TABLE_COLS, value=p.get("position","")), font=work_history_font)                        
            
            # --- この案件の縦セル結合 ---
            end_row = cur - 1 # この案件の最終行
            if end_row > start_row: # 作業内容などで2行以上になった場合
                # C列 (案件名/作業内容) 以外を縦に結合
                for c_idx in [c for c in range(1, TABLE_COLS + 1) if c != COL_PROJECT_NAME]:
                    ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
                    # 結合したセルのスタイルを再適用 (上寄せ)
                    cell = ws.cell(row=start_row, column=c_idx)
                    style(cell, align=wrap_text_alignment)

            for j in range((end_row + 1) - start_row):
                # 他の列 (A, B, D-K) にも罫線を引く (結合される親セル以外)
                for c_idx in [c for c in range(2, TABLE_COLS) if c != COL_PROJECT_NAME]:
                    style(ws.cell(row=start_row + j, column=c_idx + 1),font=dashdot_border, border=dashdot_border)

        # --- 幅調整 (サンプル形式) ---
        ws.column_dimensions["A"].width = 1.3  # 項番
        ws.column_dimensions["B"].width = 3 # 期間
        ws.column_dimensions["C"].width = 11.5 # 案件名/作業内容
        ws.column_dimensions["D"].width = 15 # 業種
        ws.column_dimensions["E"].width = 11.5 # OS
        ws.column_dimensions["F"].width = 20 # 言語
        ws.column_dimensions["G"].width = 11.5 # DB
        ws.column_dimensions["H"].width = 4.25 # 工程
        ws.column_dimensions["I"].width = 10.25 # 役割
        ws.column_dimensions["J"].width = 8.8 # ポジション
        ws.column_dimensions["K"].width = 11 # 規模

        ws.row_dimensions[1].height = 43
        ws.row_dimensions[2].height = 30
        ws.row_dimensions[3].height = 30
    
    st.download_button(
        label="スキルシートをダウンロード",
        data=output.getvalue(),
        file_name=f"{st.session_state.pi_name or 'スキルシート'}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.success("Excelを生成しました。")
