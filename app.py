"""
Word to LaTeX 変換Webアプリ
Pandoc を使用して .docx ファイルを .tex に変換し、
画像を含む成果物を Zip でダウンロードできる Streamlit アプリケーション。
"""

import os
import subprocess
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path

import streamlit as st

# ──────────────────────────────────────────────
# ページ設定
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="Word → LaTeX Converter",
    page_icon="📄",
    layout="centered",
)

# ──────────────────────────────────────────────
# カスタム CSS
# ──────────────────────────────────────────────
st.markdown(
    """
    <style>
    /* ─── 全体 ─── */
    .stApp {
        background: linear-gradient(135deg, #f0f4ff 0%, #ffffff 50%, #e8efff 100%);
    }

    /* ─── ヘッダー ─── */
    .main-header {
        text-align: center;
        padding: 2rem 1rem 1rem;
    }
    .main-header h1 {
        background: linear-gradient(135deg, #1e3a8a, #3b82f6);
        -webkit-background-clip: text;
        background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.4rem;
        font-weight: 800;
        margin-bottom: 0.25rem;
    }
    .main-header p {
        color: #64748b;
        font-size: 1.05rem;
    }

    /* ─── ステップカード ─── */
    .step-container {
        display: flex;
        justify-content: center;
        gap: 1.2rem;
        margin: 1.5rem auto 2rem;
        flex-wrap: wrap;
    }
    .step-card {
        background: #ffffff;
        border: 1px solid #dbeafe;
        border-radius: 14px;
        padding: 1.2rem 1.6rem;
        width: 200px;
        text-align: center;
        box-shadow: 0 2px 12px rgba(59,130,246,0.08);
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .step-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 6px 20px rgba(59,130,246,0.15);
    }
    .step-num {
        display: inline-block;
        background: linear-gradient(135deg, #3b82f6, #2563eb);
        color: #fff;
        width: 36px; height: 36px;
        line-height: 36px;
        border-radius: 50%;
        font-weight: 700;
        font-size: 1rem;
        margin-bottom: 0.5rem;
    }
    .step-card h3 {
        margin: 0.3rem 0 0.15rem;
        font-size: 1rem;
        color: #1e3a8a;
    }
    .step-card p {
        margin: 0;
        font-size: 0.82rem;
        color: #94a3b8;
    }

    /* ─── 変換結果セクション ─── */
    .result-header {
        color: #1e3a8a;
        font-weight: 700;
        font-size: 1.15rem;
        margin-top: 1.5rem;
    }

    /* ─── フッター ─── */
    .footer {
        text-align: center;
        color: #94a3b8;
        font-size: 0.8rem;
        margin-top: 3rem;
        padding-bottom: 1rem;
    }

    /* ─── ボタン ─── */
    .stButton > button {
        background: linear-gradient(135deg, #3b82f6, #2563eb);
        color: #fff;
        border: none;
        border-radius: 10px;
        padding: 0.6rem 2rem;
        font-weight: 600;
        font-size: 1rem;
        transition: transform 0.15s, box-shadow 0.15s;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 14px rgba(37,99,235,0.35);
        color: #fff;
    }

    .stDownloadButton > button {
        background: linear-gradient(135deg, #10b981, #059669);
        color: #fff;
        border: none;
        border-radius: 10px;
        padding: 0.6rem 2rem;
        font-weight: 600;
        font-size: 1rem;
        transition: transform 0.15s, box-shadow 0.15s;
    }
    .stDownloadButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 14px rgba(5,150,105,0.35);
        color: #fff;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ──────────────────────────────────────────────
# ヘッダー
# ──────────────────────────────────────────────
st.markdown(
    """
    <div class="main-header">
        <h1>📄 Word → LaTeX Converter</h1>
        <p>Pandoc を使用して Word 文書を LaTeX 形式に変換します</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ──────────────────────────────────────────────
# 3ステップガイド
# ──────────────────────────────────────────────
st.markdown(
    """
    <div class="step-container">
        <div class="step-card">
            <span class="step-num">1</span>
            <h3>ファイル選択</h3>
            <p>.docx ファイルを<br>アップロード</p>
        </div>
        <div class="step-card">
            <span class="step-num">2</span>
            <h3>変換実行</h3>
            <p>ボタンを押して<br>LaTeX に変換</p>
        </div>
        <div class="step-card">
            <span class="step-num">3</span>
            <h3>ダウンロード</h3>
            <p>.tex と画像を<br>Zip で取得</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ──────────────────────────────────────────────
# Pandoc 存在チェック
# ──────────────────────────────────────────────

def _check_pandoc() -> bool:
    """pandoc コマンドが利用可能か確認する。"""
    try:
        subprocess.run(
            ["pandoc", "--version"],
            capture_output=True,
            check=True,
        )
        return True
    except (FileNotFoundError, subprocess.CalledProcessError):
        return False


if not _check_pandoc():
    st.error("⚠️ Pandoc が見つかりません。Pandoc をインストールしてください。")
    st.stop()

# ──────────────────────────────────────────────
# サイドバー: 言語設定
# ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ 設定")
    st.markdown("---")
    lang = st.radio(
        "🌐 言語設定（Language Setting）",
        options=["日本語", "English"],
        index=0,
        help="日本語: ltjsarticle クラスで出力\nEnglish: article クラス（標準）で出力",
    )
    if lang == "日本語":
        st.info("📝 `ltjsarticle` + `unicode` オプションで出力します")
    else:
        st.info("📝 標準の `article` クラスで出力します")

# ──────────────────────────────────────────────
# ファイルアップロード
# ──────────────────────────────────────────────
st.markdown("---")
uploaded_file = st.file_uploader(
    "📎 Word ファイルをアップロード",
    type=["docx"],
    help="変換したい .docx ファイルを選択してください。",
)

# ──────────────────────────────────────────────
# 変換処理
# ──────────────────────────────────────────────

def convert_docx_to_latex(
    docx_bytes: bytes,
    filename: str,
    lang: str = "日本語",
) -> tuple[str, bytes]:
    """
    .docx → .tex 変換を行い、tex テキストと zip バイト列を返す。

    Parameters
    ----------
    docx_bytes : bytes
        アップロードされた docx のバイナリ。
    filename : str
        元のファイル名（拡張子付き）。
    lang : str
        言語設定。"日本語" または "English"。

    Returns
    -------
    tuple[str, bytes]
        (LaTeX ソース文字列, Zip バイト列)
    """
    stem = Path(filename).stem

    with tempfile.TemporaryDirectory() as tmpdir:
        # ── docx を一時ファイルに保存 ──
        docx_path = os.path.join(tmpdir, filename)
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        # ── 出力パス ──
        tex_path = os.path.join(tmpdir, f"{stem}.tex")
        media_dir = os.path.join(tmpdir, "media")

        # ── Pandoc 実行 ──
        cmd = [
            "pandoc",
            docx_path,
            "-o", tex_path,
            "--standalone",
            f"--extract-media={tmpdir}",
            "--mathml",
        ]

        # ── 言語別オプション ──
        if lang == "日本語":
            cmd.extend([
                "-V", "documentclass=ltjsarticle",
                "-V", "classoption=unicode",
            ])

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
        )

        if result.returncode != 0:
            raise RuntimeError(
                f"Pandoc 変換に失敗しました:\n{result.stderr}"
            )

        # ── tex ファイル読み込み ──
        with open(tex_path, "r", encoding="utf-8") as f:
            tex_content = f.read()

        # ── Zip パッケージ作成 ──
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            # tex ファイル
            zf.write(tex_path, f"{stem}.tex")

            # 抽出された画像
            if os.path.isdir(media_dir):
                for root, _dirs, files in os.walk(media_dir):
                    for file in files:
                        abs_path = os.path.join(root, file)
                        arc_name = os.path.join(
                            "media",
                            os.path.relpath(abs_path, media_dir),
                        )
                        zf.write(abs_path, arc_name)

        zip_buffer.seek(0)
        return tex_content, zip_buffer.getvalue()


# ──────────────────────────────────────────────
# メインロジック
# ──────────────────────────────────────────────
if uploaded_file is not None:
    st.success(f"✅ **{uploaded_file.name}** を読み込みました")

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        convert_clicked = st.button(
            "🔄 変換を実行",
            use_container_width=True,
        )

    if convert_clicked:
        with st.spinner("Pandoc で変換中..."):
            try:
                tex_content, zip_bytes = convert_docx_to_latex(
                    uploaded_file.getvalue(),
                    uploaded_file.name,
                    lang=lang,
                )
            except RuntimeError as e:
                st.error(str(e))
                st.stop()

        # ── 結果をセッションに保存 ──
        st.session_state["tex_content"] = tex_content
        st.session_state["zip_bytes"] = zip_bytes
        st.session_state["stem"] = Path(uploaded_file.name).stem

    # ── 結果表示（セッション保持） ──
    if "tex_content" in st.session_state:
        tex_content = st.session_state["tex_content"]
        zip_bytes = st.session_state["zip_bytes"]
        stem = st.session_state["stem"]

        st.markdown('<p class="result-header">📝 LaTeX プレビュー（冒頭 30 行）</p>', unsafe_allow_html=True)
        preview_lines = tex_content.splitlines()[:30]
        st.code("\n".join(preview_lines), language="latex")

        if len(tex_content.splitlines()) > 30:
            st.caption(f"…全 {len(tex_content.splitlines())} 行中、冒頭 30 行を表示しています。")

        st.markdown("---")
        dl_col1, dl_col2, dl_col3 = st.columns([1, 2, 1])
        with dl_col2:
            st.download_button(
                label="📦 Zip をダウンロード",
                data=zip_bytes,
                file_name=f"{stem}_latex.zip",
                mime="application/zip",
                use_container_width=True,
            )

# ──────────────────────────────────────────────
# フッター
# ──────────────────────────────────────────────
st.markdown(
    '<div class="footer">Powered by Pandoc &amp; Streamlit</div>',
    unsafe_allow_html=True,
)
