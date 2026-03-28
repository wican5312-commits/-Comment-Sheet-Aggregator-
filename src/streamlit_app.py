import sys
import os

_SRC_DIR = os.path.dirname(os.path.abspath(__file__))
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

import streamlit as st
import tempfile
from datetime import datetime

try:
    import aggregator
except ImportError as e:
    try:
        sys.path.insert(0, os.path.join(_SRC_DIR, '..', 'src'))
        import aggregator
    except ImportError as e2:
        st.error(f"aggregator モジュールが見つかりません。\n{e}\n{e2}")
        st.stop()

# ---------------------------------------------------------------------------
# Translations
# ---------------------------------------------------------------------------
T = {
    "JP": {
        "lang_btn":        "English",
        "app_title":       "コメントシート集計ツール",
        "app_sub":         "コメントシートを集計して出席簿と合わせた Excel を生成します",

        "s1_title":        "コメントシート",
        "s1_required":     "必須",
        "s1_hint":         "集計したい Excel ファイルを選択してください（複数可）",
        "s1_types":        ".xlsx / .xls",
        "s1_n_files":      "個のファイルを選択中",

        "s2_title":        "出席簿",
        "s2_optional":     "任意",
        "s2_hint":         "出席簿（KogibetuSeiseki_から始まるファイル）を追加すると学籍番号順に整列し、書式も引き継ぎます",
        "s2_types":        ".xlsx / .xls",
        "s2_selected":     "選択中:",

        "s3_title":        "設定",
        "year_label":      "対象年度",
        "year_ph":         "例: 2025　（空白 = 全データ）",

        "year_warn_title": "対象年度が空欄です",
        "year_warn_body":  "年度が入力されていません。ヘッダー行や無関係な科目を含む **すべての行** が処理されます。このまま続けますか？",
        "year_warn_yes":   "続ける",
        "year_warn_no":    "戻る",

        "btn_run":         "集計開始",
        "no_files_hint":   "コメントシートを選択してから実行してください。",

        "log_loading":     "ファイルを読み込んでいます...",
        "log_att":         "出席簿を処理しています...",
        "log_running":     "集計処理を実行中...",
        "log_done":        "完了しました。",
        "log_failed":      "失敗: ",
        "log_error":       "エラー: ",

        "result_title":    "集計完了！",
        "result_sub":      "ファイルの準備ができました。",
        "btn_dl":          "結果をダウンロード",

        "footer":          "制作：2025年度院生（有志）",
    },
    "EN": {
        "lang_btn":        "日本語",
        "app_title":       "Comment Sheet Aggregator",
        "app_sub":         "Consolidate comment sheets and merge with attendance data into one Excel file.",

        "s1_title":        "Comment Sheets",
        "s1_required":     "Required",
        "s1_hint":         "Upload one or more Excel comment sheet files",
        "s1_types":        ".xlsx / .xls",
        "s1_n_files":      "file(s) selected",

        "s2_title":        "Attendance Sheet",
        "s2_optional":     "Optional",
        "s2_hint":         "Sorts results by student ID and preserves original formatting (Look for file starts with KogibetuSeiseki_)",
        "s2_types":        ".xlsx / .xls",
        "s2_selected":     "Selected:",

        "s3_title":        "Settings",
        "year_label":      "Target Year",
        "year_ph":         "e.g. 2025  (blank = all data)",

        "year_warn_title": "Target year is empty",
        "year_warn_body":  "No year was entered. **All rows** will be processed, including potential header rows and unrelated courses. Continue?",
        "year_warn_yes":   "Continue",
        "year_warn_no":    "Go back",

        "btn_run":         "Run Aggregation",
        "no_files_hint":   "Please select at least one comment sheet before running.",

        "log_loading":     "Loading files...",
        "log_att":         "Processing attendance sheet...",
        "log_running":     "Running aggregation...",
        "log_done":        "Done.",
        "log_failed":      "Failed: ",
        "log_error":       "Error: ",

        "result_title":    "Done!",
        "result_sub":      "Your file is ready to download.",
        "btn_dl":          "Download Result",

        "footer":          "Developed by 2025 Graduate Students",
    },
}

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
_CSS = """
<style>
/* ── Force light mode ─────────────────────────────────────────── */
html, body,
.stApp, [data-testid="stApp"],
[data-testid="stAppViewContainer"],
[data-testid="stMain"], .main,
[data-testid="block-container"], .block-container {
    background-color: #f2f4f9 !important;
    color: #1c1e26 !important;
}
[data-testid="stSidebar"] { display: none !important; }
#MainMenu, footer, header { visibility: hidden !important; }

.block-container {
    padding-top: 0 !important;
    padding-bottom: 3rem !important;
    max-width: 760px !important;
}

/* ── Typography ───────────────────────────────────────────────── */
.stMarkdown p, .stMarkdown li, label { color: #374151 !important; }

/* ── App header ───────────────────────────────────────────────── */
.app-title {
    font-size: 1.5rem;
    font-weight: 800;
    color: #111827 !important;
    margin: 0 0 0.2rem;
    letter-spacing: -0.02em;
    line-height: 1.25;
}
.app-sub {
    font-size: 0.87rem;
    color: #6b7280 !important;
    margin: 0;
}

/* ── Cards ────────────────────────────────────────────────────── */
.card {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    padding: 1.2rem 1.5rem 1rem;
    margin-bottom: 0.8rem;
    box-shadow: 0 1px 5px rgba(0,0,0,0.05);
}
.card-header {
    display: flex;
    align-items: center;
    gap: 0.6rem;
    margin-bottom: 0.45rem;
}
.card-num {
    width: 24px; height: 24px;
    border-radius: 50%;
    background: #3b48d8;
    color: white !important;
    font-size: 0.75rem;
    font-weight: 700;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}
.card-num-gray {
    width: 24px; height: 24px;
    border-radius: 50%;
    background: #9ca3af;
    color: white !important;
    font-size: 0.75rem;
    font-weight: 700;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}
.card-title {
    font-size: 0.95rem;
    font-weight: 700;
    color: #111827 !important;
    flex: 1;
}
.tag {
    font-size: 0.67rem;
    font-weight: 700;
    padding: 0.15rem 0.55rem;
    border-radius: 20px;
    letter-spacing: 0.03em;
}
.tag-req { background: #fee2e2; color: #b91c1c !important; }
.tag-opt { background: #f1f5f9; color: #475569 !important; }
.card-hint {
    font-size: 0.79rem;
    color: #9ca3af !important;
    margin-bottom: 0.6rem;
}

/* File status pills */
.pill {
    display: inline-flex;
    align-items: center;
    gap: 0.3rem;
    font-size: 0.79rem;
    font-weight: 600;
    padding: 0.25rem 0.75rem;
    border-radius: 20px;
    margin-top: 0.4rem;
}
.pill-green { background: #dcfce7; color: #166534 !important; }
.pill-blue  { background: #dbeafe; color: #1e40af !important; }

/* ── Form widgets ─────────────────────────────────────────────── */
[data-testid="stTextInput"] input {
    background-color: #ffffff !important;
    color: #1c1e26 !important;
    border: 1.5px solid #d1d5db !important;
    border-radius: 8px !important;
}
[data-testid="stTextInput"] label { color: #374151 !important; font-weight: 600; }

[data-testid="stFileUploader"],
[data-testid="stFileUploaderDropzone"] {
    background-color: #f8faff !important;
    border: 2px dashed #a5b4fc !important;
    border-radius: 10px !important;
}
/* すべての子要素のテキストを強制的に暗色に */
[data-testid="stFileUploader"] *,
[data-testid="stFileUploaderDropzone"] * {
    color: #374151 !important;
}
/* ファイル名表示行 */
[data-testid="stFileUploaderFileName"],
[data-testid="stFileUploaderFile"],
[data-testid="stFileUploaderFile"] *,
[data-testid="stFileUploaderFileData"],
[data-testid="stFileUploaderFileData"] * {
    color: #1c1e26 !important;
    background-color: #ffffff !important;
}
/* アップロード済みファイルのコンテナ */
.uploadedFile, .uploadedFile * {
    color: #1c1e26 !important;
    background-color: #f0f4ff !important;
}
/* Browse files / × ボタン */
[data-testid="stFileUploaderDropzone"] button,
[data-testid="stFileUploaderDeleteBtn"] button,
[data-testid="stFileUploaderDeleteBtn"] {
    background-color: #eef2ff !important;
    color: #3730a3 !important;
    border: 1px solid #c7d2fe !important;
    border-radius: 6px !important;
}
/* small / caption テキスト（"Limit 200MB…" など） */
[data-testid="stFileUploader"] small,
[data-testid="stFileUploader"] span {
    color: #6b7280 !important;
}

/* ── Buttons ──────────────────────────────────────────────────── */
[data-testid="stButton"] button {
    border-radius: 9px;
    font-weight: 600;
    transition: all 0.18s ease;
    background-color: #ffffff !important;
    color: #374151 !important;
    border: 1.5px solid #d1d5db !important;
    height: 2.6em;
}
[data-testid="stButton"] button:hover {
    border-color: #818cf8 !important;
    color: #3730a3 !important;
    background-color: #f5f3ff !important;
}
[data-testid="stButton"] button[kind="primary"] {
    background: #3b48d8 !important;
    border-color: #3b48d8 !important;
    color: #ffffff !important;
    font-size: 1.05rem;
    height: 3.5em;
    box-shadow: 0 3px 12px rgba(59,72,216,0.25);
}
[data-testid="stButton"] button[kind="primary"]:hover {
    background: #2f3bbd !important;
    border-color: #2f3bbd !important;
    color: #ffffff !important;
    box-shadow: 0 6px 20px rgba(59,72,216,0.35);
    transform: translateY(-1px);
}
[data-testid="stDownloadButton"] button {
    background: #1a7f5a !important;
    border-color: #1a7f5a !important;
    color: #ffffff !important;
    border-radius: 9px;
    font-size: 1.05rem;
    height: 3.5em;
    font-weight: 600;
    box-shadow: 0 3px 12px rgba(26,127,90,0.22);
    transition: all 0.18s ease;
    width: 100%;
}
[data-testid="stDownloadButton"] button:hover {
    background: #15694a !important;
    box-shadow: 0 6px 20px rgba(26,127,90,0.32);
    transform: translateY(-1px);
}

/* ── Progress ─────────────────────────────────────────────────── */
[data-testid="stProgressBar"] > div {
    background-color: #e0e7ff !important;
    border-radius: 99px; height: 5px !important;
}
[data-testid="stProgressBar"] > div > div {
    background: #3b48d8 !important;
    border-radius: 99px;
}

/* ── Log box ──────────────────────────────────────────────────── */
.log-box {
    background: #1e2030;
    border-radius: 10px;
    padding: 0.85rem 1rem;
    font-family: 'Menlo', 'Consolas', monospace;
    font-size: 0.77rem;
    line-height: 1.75;
    max-height: 160px;
    overflow-y: auto;
    margin-top: 0.7rem;
}
.log-line      { color: #c9d1d9 !important; }
.log-line.ok   { color: #3fb950 !important; }
.log-line.err  { color: #f85149 !important; }
.log-line.info { color: #79c0ff !important; }

/* ── Result box ───────────────────────────────────────────────── */
.result-box {
    background: #f0fdf4;
    border: 1.5px solid #86efac;
    border-radius: 12px;
    padding: 1.1rem 1.5rem;
    text-align: center;
    margin: 0.5rem 0 0.7rem;
}
.result-icon  { font-size: 1.9rem; }
.result-title { font-size: 1.05rem; font-weight: 700; color: #166534 !important; margin: 0.25rem 0 0.1rem; }
.result-sub   { font-size: 0.83rem; color: #15803d !important; }

/* ── Confirm banner ───────────────────────────────────────────── */
.confirm-box {
    background: #fffbeb;
    border: 1.5px solid #fcd34d;
    border-radius: 12px;
    padding: 0.95rem 1.2rem;
    margin-bottom: 0.7rem;
}
.confirm-title { font-size: 0.93rem; font-weight: 700; color: #92400e !important; margin-bottom: 0.35rem; }
.confirm-body  { font-size: 0.83rem; color: #78350f !important; line-height: 1.6; }

/* ── Misc ─────────────────────────────────────────────────────── */
hr { border-color: #e5e7eb !important; margin: 0.9rem 0 !important; }
.app-footer {
    text-align: center;
    color: #9ca3af !important;
    font-size: 0.75rem;
    padding: 1.5rem 0 0.5rem;
}

/* ── Safety net: catch any remaining dark-mode text ──────────── */
/* Streamlit occasionally injects inline color:white via JS;
   these selectors cover the most common offenders. */
[data-testid="stMarkdownContainer"] *,
[data-testid="stText"] *,
[data-testid="stCaptionContainer"] *,
[data-testid="stAlert"] *,
[data-testid="stNotification"] * {
    color: #374151 !important;
}
/* Number/size labels inside file uploader */
[data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] p {
    color: #6b7280 !important;
}

/* ── Mobile ───────────────────────────────────────────────────── */
@media (max-width: 600px) {
    .app-title { font-size: 1.2rem; }
    .card { padding: 1rem 1.1rem 0.85rem; }
    [data-testid="stButton"] button[kind="primary"] { height: 3.2em; }
}
</style>
"""

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------

def _init():
    if st.session_state.get("_init"):
        return
    st.session_state._init = True
    st.session_state.setdefault("lang",          "JP")
    st.session_state.setdefault("logs",          [])
    st.session_state.setdefault("result_bytes",  None)
    st.session_state.setdefault("result_name",   None)
    st.session_state.setdefault("confirm_year",  False)
    st.session_state.setdefault("pending_run",   False)


def _add_log(msg, kind=""):
    st.session_state.logs.append((msg, kind))


def _render_log():
    if not st.session_state.logs:
        return
    lines = "".join(
        f'<div class="log-line {k}">{m}</div>'
        for m, k in st.session_state.logs
    )
    st.markdown(f'<div class="log-box">{lines}</div>', unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Run logic
# ---------------------------------------------------------------------------

def _do_run(uploaded_files, attendance_file, target_year, t):
    st.session_state.result_bytes = None
    st.session_state.result_name  = None
    st.session_state.logs         = []

    prog   = st.progress(0)
    status = st.empty()

    with tempfile.TemporaryDirectory() as tmp:
        in_dir = os.path.join(tmp, "in")
        os.makedirs(in_dir, exist_ok=True)

        _add_log(t["log_loading"], "info")
        status.markdown(f"⏳ {t['log_loading']}")
        paths = []
        for i, uf in enumerate(uploaded_files):
            fp = os.path.join(in_dir, uf.name)
            with open(fp, "wb") as f:
                f.write(uf.getbuffer())
            paths.append(fp)
            prog.progress((i + 1) / len(uploaded_files) * 0.3)

        att_path = None
        if attendance_file:
            att_path = os.path.join(tmp, attendance_file.name)
            with open(att_path, "wb") as f:
                f.write(attendance_file.getbuffer())
            _add_log(t["log_att"], "info")
            status.markdown(f"⏳ {t['log_att']}")

        out_name = f"summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        out_path = os.path.join(tmp, out_name)

        _add_log(t["log_running"], "info")
        status.markdown(f"⏳ {t['log_running']}")
        prog.progress(0.6)

        try:
            success, msg = aggregator.process_files(
                paths, out_path, target_year, att_path,
            )
            prog.progress(1.0)
            status.empty()
            if success and os.path.exists(out_path):
                with open(out_path, "rb") as f:
                    st.session_state.result_bytes = f.read()
                st.session_state.result_name = out_name
                _add_log(t["log_done"], "ok")
            else:
                _add_log(t["log_failed"] + msg, "err")
        except Exception as e:
            status.empty()
            _add_log(t["log_error"] + str(e), "err")

    prog.empty()

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    st.set_page_config(
        page_title="コメントシート集計ツール",
        page_icon="📑",
        layout="centered",
        initial_sidebar_state="collapsed",
    )

    _init()
    st.markdown(_CSS, unsafe_allow_html=True)

    t = T[st.session_state.lang]

    # ── Header ───────────────────────────────────────────────────────────
    col_title, col_lang = st.columns([5, 1])
    with col_title:
        st.markdown(
            f'<div style="padding: 1.4rem 0 0.5rem;">'
            f'  <div class="app-title">{t["app_title"]}</div>'
            f'  <div class="app-sub">{t["app_sub"]}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
    with col_lang:
        st.write("")
        st.write("")
        if st.button(t["lang_btn"], key="lang_btn"):
            st.session_state.lang = "EN" if st.session_state.lang == "JP" else "JP"
            st.rerun()

    # ── Step 1 : Comment sheets ───────────────────────────────────────────
    st.markdown(f"""
    <div class="card">
      <div class="card-header">
        <div class="card-num">1</div>
        <div class="card-title">{t['s1_title']}</div>
        <span class="tag tag-req">{t['s1_required']}</span>
      </div>
      <div class="card-hint">{t['s1_hint']} &nbsp;·&nbsp; {t['s1_types']}</div>
    </div>
    """, unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        t["s1_title"], type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="cs_files", label_visibility="collapsed",
    )
    if uploaded_files:
        st.markdown(
            f'<span class="pill pill-green">✓ &nbsp;{len(uploaded_files)} {t["s1_n_files"]}</span>',
            unsafe_allow_html=True,
        )

    st.write("")

    # ── Step 2 : Attendance sheet ─────────────────────────────────────────
    st.markdown(f"""
    <div class="card">
      <div class="card-header">
        <div class="card-num-gray">2</div>
        <div class="card-title">{t['s2_title']}</div>
        <span class="tag tag-opt">{t['s2_optional']}</span>
      </div>
      <div class="card-hint">{t['s2_hint']} &nbsp;·&nbsp; {t['s2_types']}</div>
    </div>
    """, unsafe_allow_html=True)

    attendance_file = st.file_uploader(
        t["s2_title"], type=["xlsx", "xls"],
        key="att_file", label_visibility="collapsed",
    )
    if attendance_file:
        st.markdown(
            f'<span class="pill pill-blue">✓ &nbsp;{t["s2_selected"]} {attendance_file.name}</span>',
            unsafe_allow_html=True,
        )

    st.write("")

    # ── Step 3 : Year filter ──────────────────────────────────────────────
    st.markdown(f"""
    <div class="card">
      <div class="card-header">
        <div class="card-num-gray">3</div>
        <div class="card-title">{t['s3_title']}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    target_year = st.text_input(
        t["year_label"],
        placeholder=t["year_ph"],
        key="year_input",
    )

    st.divider()

    # ── Year-empty confirm ────────────────────────────────────────────────
    if st.session_state.confirm_year:
        st.markdown(f"""
        <div class="confirm-box">
          <div class="confirm-title">⚠️ &nbsp;{t['year_warn_title']}</div>
          <div class="confirm-body">{t['year_warn_body']}</div>
        </div>
        """, unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button(t["year_warn_no"], use_container_width=True, key="warn_no"):
                st.session_state.confirm_year = False
                st.session_state.pending_run  = False
                st.rerun()
        with c2:
            if st.button(t["year_warn_yes"], type="primary",
                         use_container_width=True, key="warn_yes"):
                st.session_state.confirm_year = False
                st.session_state.pending_run  = True
                st.rerun()

    # ── Run button ────────────────────────────────────────────────────────
    if st.button(
        f"🚀  {t['btn_run']}",
        type="primary",
        use_container_width=True,
        disabled=not bool(uploaded_files),
        key="run_btn",
    ):
        if not target_year.strip():
            st.session_state.confirm_year = True
            st.session_state.pending_run  = False
            st.rerun()
        else:
            st.session_state.pending_run = True

    if not uploaded_files and not st.session_state.confirm_year:
        st.caption(f"⬆️  {t['no_files_hint']}")

    # Execute
    if st.session_state.pending_run and uploaded_files:
        st.session_state.pending_run = False
        _do_run(uploaded_files, attendance_file, target_year.strip(), t)
        st.rerun()

    # ── Log ───────────────────────────────────────────────────────────────
    _render_log()

    # ── Result ────────────────────────────────────────────────────────────
    if st.session_state.result_bytes:
        st.markdown(f"""
        <div class="result-box">
          <div class="result-icon">✅</div>
          <div class="result-title">{t['result_title']}</div>
          <div class="result-sub">{t['result_sub']}</div>
        </div>
        """, unsafe_allow_html=True)
        st.balloons()
        st.download_button(
            label=f"📥  {t['btn_dl']}",
            data=st.session_state.result_bytes,
            file_name=st.session_state.result_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_btn",
        )

    # ── Footer ────────────────────────────────────────────────────────────
    st.markdown(
        f'<div class="app-footer">{t["footer"]}</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
