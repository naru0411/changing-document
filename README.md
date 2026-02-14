# Word → LaTeX Converter

Pandoc を使用して Word 文書（.docx）を LaTeX（.tex）に変換する Streamlit Web アプリです。

## 前提条件

- **Python 3.10** がインストール済み（`py -3.10 --version` で確認）
- **Pandoc** がインストール済み（`pandoc --version` で確認）

## セットアップ手順

### 1. 仮想環境の作成

```bash
py -3.10 -m venv .venv
```

### 2. 仮想環境の有効化

```bash
.venv\Scripts\activate
```

### 3. 依存ライブラリのインストール

```bash
pip install streamlit
```

### 4. アプリの起動

```bash
streamlit run app.py
```

ブラウザが自動的に開き、`http://localhost:8501` でアプリにアクセスできます。

## 使い方

1. **ファイル選択** — `.docx` ファイルをアップロード
2. **変換実行** — 「🔄 変換を実行」ボタンをクリック
3. **ダウンロード** — 変換された `.tex` と画像を含む Zip ファイルをダウンロード

## 技術スタック

| 項目 | 技術 |
|------|------|
| UI | Streamlit |
| 変換エンジン | Pandoc |
| 言語 | Python 3.10 |
