# -*- coding: utf-8 -*-
# スライド（PDF）→ 日本語原稿 → DOCX を作るStreamlitアプリ
# GitHub + Streamlit Community Cloud想定版
# ・毎回、サイドバーで「共通パスワード」と「OpenAI APIキー」を入力して利用
# ・APIキーは保存しません（セッション内のみ）。パスワードはハッシュ照合対応

import os
import re
import hashlib
from io import BytesIO
from typing import List

import pdfplumber
import streamlit as st
from docx import Document
from docx.shared import Pt
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from openai import OpenAI, APIError, RateLimitError, APITimeoutError

# ======================================================
# 設定：パスワード照合（ハッシュで管理）
# ------------------------------------------------------
# 1) Streamlit Cloudの「Advanced settings → Environment variables」に
#    APP_PASSWORD_HASH（SHA256）を設定すると、その値を使います。
# 2) 未設定なら、下の DEFAULT_PASSWORD_HASH を使います（空ならパス不要）。
#    → 管理しやすいので、基本は環境変数にハッシュ値を入れる運用推奨。
# ======================================================
DEFAULT_PASSWORD_HASH = ""  # 例: "5e884898da28047151d0e56f8dc629...（'password'のSHA256）"

def _sha256(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()

def _get_password_hash() -> str:
    return os.getenv("APP_PASSWORD_HASH", DEFAULT_PASSWORD_HASH).strip()

# ======================================================
# UI 基本
# ======================================================
st.set_page_config(page_title="TCROSS SlideCreator", layout="wide")
st.title("スライドクリエーター（PDF → 日本語原稿 → DOCX）")

# ======================================================
# 認証（毎回、サイドバーで入力）
# ======================================================
with st.sidebar:
    st.header("🔐 アクセス & API")
    input_pw = st.text_input("共通パスワード", type="password", help="管理者から共有されたパスワードを入力")
    input_api_key = st.text_input("OpenAI APIキー", type="password", help="sk- から始まるキーを入力（保存しません）")
    login = st.button("ログイン / 更新")

# パスワードチェック
required_hash = _get_password_hash()
if required_hash:
    if login:
        st.session_state["auth_ok"] = (_sha256(input_pw) == required_hash)
        st.session_state["api_key"] = input_api_key.strip()
    auth_ok = st.session_state.get("auth_ok", False)
else:
    # パスワード未設定運用（誰でも入れる）。APIキーだけ確認
    if login:
        st.session_state["auth_ok"] = True
        st.session_state["api_key"] = input_api_key.strip()
    auth_ok = st.session_state.get("auth_ok", False)

# 未ログイン/未入力時のガード
if not auth_ok:
    st.info("サイドバーで **共通パスワード** と **OpenAI APIキー** を入力して「ログイン / 更新」を押してください。")
    st.stop()

api_key = st.session_state.get("api_key", "").strip()
if not api_key:
    st.error("OpenAI APIキーが未入力です。サイドバーから入力してください。")
    st.stop()

# OpenAIクライアント
client = OpenAI(api_key=api_key)

# =========================
# ルール（最終版）
# =========================
RULES = r"""
原稿クリエーター ルール（最終版）

1. 全体ルール
• スライド（PDF）を日本語原稿に変換する。
• 出力は日本語のみ。
• 論文調・医療学会要旨調で整える。
• 数値・HR・95%CI・p値などは絶対に改変しない。
• p値は必ず「p=」のように小文字の「p」で表記する（大文字Pは禁止）。
• 「か月」は必ず「ヶ月」と表記する。
• 固有名詞・試験名は保持する。
• スライド番号や英語原文は不要。
• スライドに記載されている「LIMITATION（試験の限界）」「CLINICAL IMPLICATIONS（臨床的示唆）」は原稿に含めない。
• 「スライドに表示されている」「スライドに提示された」といった表現は禁止。必ず文章として表現する。
• タイトルやスライドに記載されていない内容は一切記載しない。
• 「スライドに提示された」「報告された」など、スライド依存の表現は禁止。必ず文章として表現する。
• 数字は千の位と百の位の間にカンマを挿入する（例4000は、4,000とする）。

2. タイトル
• 演題タイトルは最初のスライドを参照。タイトルがない場合も勝手につくらない。
• 演題タイトルを日本語訳で記載。
• 演題タイトルの後に「：」をつけて、試験名を書く（例　冠動脈バイパス術後の二剤抗血小板療法のde-escalation戦略：TOP-CABG試験）
• 最後は「減少」「抑制」で止める（「減少させる」「抑制する」は不可）。
• 「試験」の後に「。」は付けない。
• タイトルに「」などの括弧や引用符は付けない。
• 書式例：
アテローム動脈硬化性心血管疾患/CKDを有する2型糖尿病患者における経口セマグルチドの心血管イベント抑制効果: SOUL試験

3. 冒頭文（Conclusionから作成）
• 必ず「○○試験より、○○が○○と比べて、…ことが、アメリカ、○○○○の○○氏により、○○学会○○セッションで発表された。」の形で記載する。
• 無作為化比較試験では必ず「○○が○○と比べて」を入れる（例：クロピドグレルがアスピリンと比べて、TriClipが薬物治療単独と比べて、など）。
• 演者名はファーストネーム、ラストネームを必ず入れる。
• 所属はすべてローマ字又は英語名を使う。
• 国名は「米国」ではなく「アメリカ」。
• 国名は「欧州」ではなく「ヨーロッパ」。
• 国名はタイトルの下に含まれていればその国名を採用。
• 国名は「Korea」は「韓国」、「China」は「中国」。
• 国名と所属の間は「・」ではなく「、」。
• 国名と所属がスライドに掲載されていない場合は、「○○、○○の後に演者名○○氏らにより、」を記載し、その後は発表された学会とセッションを記載。
• 発表場所がスライドに記載されていない場合は「…○○で発表された」で終了する。
• 《》の括弧は使わない。

4. 試験デザイン
• 次の行は必ず「○○試験では、」から始める（「本試験は」不可）。
• 無作為化比較試験（RCT）の場合は、必ず以下の固定書式：
○○試験では、○○年○月から○○年○月までに、○○ヶ国の○○施設より、○○○○を有する○○○人の患者を
無作為に○○群（○○人）と○○群（○○人）に割り付けた。
  - 不明な要素（登録年月、国数、施設数、対象疾患名、全登録患者数、群名・人数）は「○○」のまま空欄。
  - 括弧内は「n=」を用いず、必ず「（○○人）」とする。
  - 割付後に投与量・漸増方法・背景療法・追跡期間などがあれば続けて記載。
• 非無作為化（単群試験・観察研究）の場合は以下の書式：
○○試験では、○○年○月から○○年○月までに、○○ヶ国の○○施設より、○○○○を有する○○○人の患者を登録した。
• 追跡結果のみの報告（例：2年間追跡結果）の場合は、上記いずれかの直後に「本報告は○年間の追跡結果である。」と追記可。
• 主要評価項目、副次評価項目の設定内容はここでは述べない。
• 評価項目の内容（例：MACCE＝死亡、心筋梗塞、心不全入院）は、それぞれの「主要評価項目」「副次評価項目」のセクションで必ず記載する。

5. 患者背景
• 患者背景スライドは必ず文章化する。省略や完全スキップは禁止。
• 両群に差がない場合は主要な項目のみ記載する。
• 年齢、女性割合、BMI、HbA1c、罹病期間、主要合併症（ASCVD、CKD、心房細動など）、薬剤使用、KCCQスコア、6分間歩行距離などを要約する。
• 基礎疾患（糖尿病、高血圧、心房細動、心不全）や病歴（PCI歴、CABG歴など）も含め、両群に差がない場合は数値の羅列は避け、「○○、○○、○○についても両群で差はなかった」とまとめる。
• 文章の終わりは単語で終わらず、「〜であった。」「〜差はなかった。」など、必ず文章として完結させる。
• 各群の割合を記載する場合は必ず「○○群○○%、コントロール群○○%」と表現する。
  - 「・」や「vs」などで区切るのは禁止。
  - 「対照群」は使わず、必ず「コントロール群」とする。
• 両群に差がない場合の表現は以下のいずれかとする：
  - 「両群の主要項目に差はなかった」
  - 「両群でバランスが取れていた」
  - ※「大きな偏りは示されていない」という表現は禁止。
• 有意差がある場合は、その項目の割合を両群ごとに記載し、p値を添える（例：「デバイス群○○%、コントロール群○○%、p=○○」）。
• ベースラインから追跡（例：1年・2年など）で示されている指標は、種類を問わず（例：TR重症度、NYHA分類、LVEF、KCCQスコア、6分間歩行距離など）、ベースライン→1年→2年の順に文章化し、改善や維持の流れを強調する。
• クロスオーバー症例がある場合は、その人数と割合を明記し、結果への影響を説明する。
• 患者背景・病変背景・手技の特徴などは必ず文章化する。
• 有意差がある場合は、両群の数値とp値を必ず記載する。
• 有意差が全くない場合は、代表的な項目（例：年齢、性別、基礎疾患、既往歴）に数値を含めて記載し、最後に「両群に差はなかった」または「両群でバランスが取れていた」とまとめる。
• 「スライドに提示された」「提示されている」といった表現は禁止。読者はスライドを見ないため、常に文章として記述する。

6. 主要評価項目
• グラフや表で提示されていても必ず文章化。
• 書式：
主要評価項目とした○○は、○群で○%、コントロール群で○%であり、ハザード比は○○であった（HR ○○［95%CI ○○―○○］p=○○）。
• 95%CIの後に必ず p値を置く。
• 95%CIやp値は文章に含めることは禁止
• 絶対禁止例①：　両群で有意差なく95%CI○○―○○であり、ｐ＜○○であった。　模範例→両群で有意差はなかった（［95%CI○○―○○］p＜○○）。
• 絶対禁止例②：術後3ヶ月未満はハザード比0.76［95%CI 0.55―1.06］ p=0.10で有意差はなく、3ヶ月以降は0.45［95%CI 0.30―0.69］ p<0.001で有意な減少が示された。　
　模範例→術後3ヶ月未満は有意差なく（HR 0.76［95%CI 0.55―1.06］ p=0.10）、3ヶ月以降は有意な減少が示された（HR 0.45［95%CI 0.30―0.69］ p<0.001）。
• 比較した文章の後に入れる。「○○群と○○群ではそれぞれ○○%と○○%であった（HR ○○［95%CI ○○―○○］ p＝○○）。
• NNTなど補助指標があれば記載。
• Freedom from は「自由度」ではなく「回避」または「回避率」と訳す。
• メインスライドだけでなく、その構成要素（コンポーネント）や分解された結果スライドも必ず文章化する。
• コンポーネントが有意差なしであっても「差はなかった」と明記する。
• Kaplan-Meierや内訳スライド、詳細スライドも必ず文章化する。
• SUMMARYスライドに同じ数値があっても、省略せず全てのスライドを文章化する。

7. 副次評価項目
• 表形式でも必ず文章化。
• 腎イベント、心血管死、下肢イベント、MI、脳卒中、出血などを順次記載。
• 有意差がある場合は割合とp値を記載。
• 書式は主要項目と同じ（HR→95%CI→p値）。
• 差がない場合は「有意差はなかった」と明記。
• 結果の箇条書きはやめ、全ては文章で表現する。
• サブスライドも含めて必ず文章化する（例：安全性の詳細スライド、出血の種類、合併症別解析など）。
• サブグループ解析や安全性解析が副次評価に含まれる場合は、別途「その他」に回さず、副次項目として記載してよい。
• Kaplan-Meierや内訳スライド、詳細スライドも必ず文章化する。
• SUMMARYスライドに同じ数値があっても、省略せず全てのスライドを文章化する。

8. その他
• 主要評価項目・副次評価項目以外に提示されている解析（サブグループ解析、代謝指標、炎症指標、QOLなど）をまとめる。
• 存在する場合のみ記載し、なければ省略。
• サブグループ解析：効果が一貫していれば「一貫していた」、交互作用があれば「交互作用が示された（p interaction=○○）」と記載。
• 代謝・炎症・QOL指標：ETD/ETRや平均変化量を正確に保持して記載。

9. 結語
• 必ず演者名で始める。
• 書式：
○○氏は、「…」と、まとめた。（必ず、「と、まとめた。」のように「と」の後に「、」を入れる。）

10. 論文掲載
• 記載がある場合のみ：
尚、本報告は《ジャーナル名》誌に掲載された。
"""

# ======================================================
# ユーティリティ
# ======================================================
def extract_slides_text(uploaded_pdf) -> List[str]:
    """PDF(UploadedFile) → 各ページのテキスト配列"""
    slides = []
    with pdfplumber.open(uploaded_pdf) as pdf:
        for p in pdf.pages:
            t = p.extract_text() or ""
            slides.append(t)
    return slides

def hash_bytes(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()

def postprocess(text: str) -> str:
    """表記ゆれ・禁止表現の最終統一"""
    if not text:
        return text

    # p値を小文字 p= に統一
    text = re.sub(r"[PｐＰ]\s*=\s*", "p=", text)

    # 「月」→「ヶ月」統一（数値＋月 → ヶ月）
    text = re.sub(r"(?<=\d)月", "ヶ月", text)

    # タイトル行の引用符（「」 “” '' ""）の除去（先頭の非空行のみ）
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if line.strip():
            t = line.strip()
            t = re.sub(r'^[「『“"＂]+', "", t)
            t = re.sub(r'[」『”"＂]+$', "", t)
            lines[i] = t
            break
    text = "\n".join(lines)

    # 語句統一等
    text = text.replace("対照群", "コントロール群")
    text = text.replace("Freedom from", "回避")
    text = re.sub(r"スライドに[^\n。]*?(提示|表示)[^\n。]*?。", "", text)

    return text

def _reset_output():
    for k in ("gen_text", "docx_bytes", "last_key"):
        st.session_state.pop(k, None)

@retry(
    retry=retry_if_exception_type((APIError, RateLimitError, APITimeoutError)),
    wait=wait_exponential(multiplier=1, min=2, max=20),
    stop=stop_after_attempt(4),
    reraise=True,
)
def call_llm(slides_text: str) -> str:
    """OpenAI Responses API で原稿化"""
    prompt = f"""
あなたは医学系学会記事の原稿作成アシスタントです。
以下のルールに厳密に従い、PDFスライドの内容のみから日本語原稿を作成してください。
スライドに無い内容は一切書かないでください。

# ルール
{RULES}

# 入力（各ページの抽出テキスト）
{slides_text}

# 必須要件
- タイトルに括弧・引用符（「」 “” ''）は付けない。タイトルが無ければ作らない。
- 主要/副次の“詳細スライド”（内訳・KM・安全性の内訳等）も必ず文章化。SUMMARYだけで省略しない。
- 患者背景は必ず文章で締める（〜であった／〜差はなかった）。
- 95%CI の後は必ず p 値（小文字 p=）。
- 「スライドに提示された／表示された」等の言い回しは禁止。常に本文として記述。
"""
    resp = client.responses.create(
        model="gpt-4.1",
        input=[
            {"role": "system", "content": "You are a meticulous Japanese medical manuscript assistant."},
            {"role": "user", "content": prompt},
        ],
    )
    return resp.output_text

def to_docx(text: str) -> bytes:
    """テキストをWord（.docx）に"""
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "MS Mincho"
    style.font.size = Pt(11)

    for para in text.split("\n"):
        doc.add_paragraph(para)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ======================================================
# UI：ファイルアップロード
# ======================================================
uploaded = st.file_uploader(
    "学会スライド PDF をアップロード",
    type=["pdf"],
    key="uploader",
    on_change=_reset_output,
)

col1, col2 = st.columns([1, 1])
with col1:
    run_btn = st.button("原稿を生成する", type="primary", use_container_width=True)
with col2:
    st.caption("※ 処理後に Word（.docx）をダウンロードできます。")

# ======================================================
# 実行
# ======================================================
if uploaded and run_btn:
    pdf_bytes = uploaded.read()
    slides = extract_slides_text(BytesIO(pdf_bytes))
    slides_joined = "\n\n--- page break ---\n\n".join(slides)

    cache_key = hash_bytes(uploaded.name.encode("utf-8") + pdf_bytes + RULES.encode("utf-8"))

    with st.spinner("原稿を生成中です…"):
        try:
            raw_text = call_llm(slides_joined)
            final_text = postprocess(raw_text)
            docx_bytes = to_docx(final_text)

            st.session_state["gen_text"] = final_text
            st.session_state["docx_bytes"] = docx_bytes
            st.session_state["last_key"] = cache_key

        except Exception as e:
            st.error(f"生成時にエラーが発生しました: {e}")
            st.stop()

# ======================================================
# 表示
# ======================================================
if "gen_text" in st.session_state and "docx_bytes" in st.session_state:
    st.success("原稿を生成しました。")
    st.download_button(
        "Word（.docx）をダウンロード",
        data=st.session_state["docx_bytes"],
        file_name="原稿.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
        key=f"dl-{st.session_state.get('last_key','')}",
    )
    with st.expander("生成テキスト（確認用）", expanded=False):
        st.write(st.session_state["gen_text"])
