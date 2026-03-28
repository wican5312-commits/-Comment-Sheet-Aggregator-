================================================================================
  Comment Sheet Aggregator / コメントシート集計ツール
================================================================================

[日本語]

--------------------------------------------------------------------------------
■ 1. このツールでできること
--------------------------------------------------------------------------------

教務システムからダウンロードした「学生コメントシート (Excel)」を
複数まとめて読み込み、学生ごとに1行に整理した Excel ファイルを出力します。

出席簿 (.xls / .xlsx) を一緒に指定すると：
  - 出席簿の書式（色・罫線・列幅）をそのまま引き継ぎます
  - 学籍番号順に並び替えます
  - 出席簿に載っていない学生も末尾に追加されます


--------------------------------------------------------------------------------
■ 2. 使い方
--------------------------------------------------------------------------------

  ┌─────────────────────────────────────────────────────────┐
  │  【A. Web 版】ブラウザで使う（推奨）                        │
  │  URL: https://shukei.streamlit.app/               │
  │  インストール不要。スマホ・Mac でも使えます。                │
  └─────────────────────────────────────────────────────────┘

    手順:
    1. 「コメントシート」欄にファイルをドラッグ＆ドロップ（複数可）
    2. 「出席簿」欄に出席簿ファイルをドラッグ＆ドロップ（任意）
    3. 「対象年度」に年度を入力（例: 2025）
    4. 「集計開始」ボタンを押す
    5. 「結果をダウンロード」ボタンからファイルを保存

  ┌─────────────────────────────────────────────────────────┐
  │  【B. Windows アプリ版】                                   │
  │  「コメントシート集計ツール.exe」をダブルクリックして起動     │
  └─────────────────────────────────────────────────────────┘

    手順:
    1. 「コメントシートを選択...」ボタンでファイルを選ぶ（複数可）
    2. 「出席表を選択...」ボタンで出席簿を選ぶ（任意）
    3. 「対象年度」に年度を入力（例: 2025）
    4. 「集計開始」ボタンを押し、保存先を指定する


--------------------------------------------------------------------------------
■ 3. アカウント情報（GitHub / Streamlit）
--------------------------------------------------------------------------------

  GitHub および Streamlit のアカウントは、いずれも
  WiCAN の Google アカウントで開設されています。

  ログイン方法：
    各サービスのログイン画面で「Googleでログイン（Sign in with Google）」
    を選択し、WiCAN の Google アカウントでログインしてください。

  Streamlit 管理画面（デプロイ・再起動・設定変更など）へのアクセス：
    方法 1: Google で「streamlit community cloud」と検索する
    方法 2: ブラウザに以下の URL を入力する
            https://share.streamlit.io/

    ※ Streamlit 管理画面ではアプリの再起動・設定変更・ログ確認ができます。


--------------------------------------------------------------------------------
■ 4. 大学のシステムが変わったときの対応方法
--------------------------------------------------------------------------------

Excel の列構成が変わった場合は、設定ファイルを書き換えます。

  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  【Windows アプリ】は画面上で変更できます
  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  アプリを起動して、右上の「⚙ 詳細設定」ボタンをクリックすると
  列や行の設定画面が開きます。変更して「保存」を押せば完了です。

  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  【Web 版】はソースコードを書き換えてから GitHub に上げます
  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  ▼ STEP 1 : 設定の場所を確認する

    「src/aggregator.py」をテキストエディタ（メモ帳でも可）で開き、
    ファイルの上部にある CONFIG という箇所を探します。

    -------------------------------------------------------
    CONFIG = {
        # コメントシートの列
        "COL_SUB_ID":  0,   # ← 提出IDが入っている列
        "COL_COURSE":  2,   # ← 科目名が入っている列
        "COL_NAME":    4,   # ← 学生氏名が入っている列
        "COL_ID":      5,   # ← 学籍番号が入っている列
        "COL_COMMENT": 6,   # ← コメント本文が入っている列

        # 出席簿の設定
        "ATT_SKIP_ROWS": 6, # ← 上から何行読み飛ばすか（ヘッダー行数）
        "ATT_COL_ID":    1, # ← 出席簿で学籍番号が入っている列
        "ATT_COL_NAME":  2, # ← 出席簿で学生氏名が入っている列
    }
    -------------------------------------------------------

  ▼ STEP 2 : 列番号のルールを理解する

    Excel の列と数字の対応は以下の通りです：

      A列 = 0,  B列 = 1,  C列 = 2,  D列 = 3,  E列 = 4,
      F列 = 5,  G列 = 6,  H列 = 7,  I列 = 8,  J列 = 9 ...

    例: 「学生氏名」が今まで E列 だったのが D列 に変わった場合
      変更前: "COL_NAME": 4,
      変更後: "COL_NAME": 3,

    例: 「データが7行目から始まる」出席簿の場合
      "ATT_SKIP_ROWS": 6,  （7-1 = 6 と計算して入力）

  ▼ STEP 3 : GitHub でファイルを更新する（ブラウザだけで完結）

    1. GitHub のリポジトリページを開く
    2. 「src」フォルダをクリック
    3. 「aggregator.py」をクリック
    4. 右上の鉛筆アイコン（Edit this file）をクリック
    5. CONFIG の数字を書き換える
    6. 画面下の「Commit changes」ボタンを押す
    7. 数分後、Web 版が自動的に更新される

    ※ Git コマンドが使える場合は、ローカルで編集後に
      「git add / git commit / git push」でも構いません。


--------------------------------------------------------------------------------
■ 4. Windows アプリを作り直す（EXE の再ビルド）
--------------------------------------------------------------------------------

ソースコードを変更した後、Windows アプリにも反映させたい場合は
EXE を作り直す必要があります。

  手順:
  1. 「build_exe.bat」をダブルクリックする
  2. 黒い画面が開いて自動でビルドが始まる
  3. 「Build complete!」と表示されたらキーを押して閉じる
  4. フォルダに新しい「コメントシート集計ツール.exe」が生成される

  ※ 初回は PyInstaller のインストールが自動で行われます（数分かかります）
  ※ 中間ファイル（build/ dist/ *.spec）は自動で削除されます


--------------------------------------------------------------------------------
■ 5. フォルダ構成（参考）
--------------------------------------------------------------------------------

  comment_sheet_aggregator/
  |
  |-- コメントシート集計ツール.exe  ... Windows アプリ本体
  |-- build_exe.bat               ... EXE を作り直すスクリプト
  |-- run_web_app.bat             ... 自分の PC でWeb版を動かすスクリプト
  |-- requirements.txt            ... 必要なライブラリの一覧
  |-- README.txt                  ... このファイル
  |
  |-- src/
  |   |-- aggregator.py           ... ★ 集計の中身・設定はここを編集
  |   |-- gui_app.py              ... Windows アプリの画面
  |   `-- streamlit_app.py        ... Web 版の画面
  |
  `-- xlrd_legacy/                ... .xls 読み取り用ライブラリ（触らない）


--------------------------------------------------------------------------------
■ 6. 削除しても良いファイル
--------------------------------------------------------------------------------

ビルド後に残っていても問題ありませんが、邪魔であれば削除して構いません：

  * build/          フォルダ（ビルド中間ファイル）
  * dist/           フォルダ（ビルド途中の EXE 置き場）
  * *.spec          ファイル（ビルド設定）
  * __pycache__/    フォルダ（Python の一時キャッシュ）
  * user_config.json ファイル（Windows アプリの設定、削除するとリセット）


================================================================================

[English]

--------------------------------------------------------------------------------
■ 1. What This Tool Does
--------------------------------------------------------------------------------

Reads multiple "Student Comment Sheet" Excel files downloaded from the
university system and generates a consolidated Excel file, one row per student.

When an attendance sheet (.xls / .xlsx) is also provided:
  - The output preserves the attendance sheet's formatting (colors, borders, etc.)
  - Results are sorted by student ID
  - Students not in the attendance sheet are appended at the bottom


--------------------------------------------------------------------------------
■ 2. How to Use
--------------------------------------------------------------------------------

  [A. Web Version — Recommended]
  URL: https://shukei.streamlit.app/
  No installation needed. Works on mobile and Mac.

    Steps:
    1. Drop comment sheet file(s) into the Step 1 upload area
    2. Drop the attendance sheet into Step 2 (optional)
    3. Enter the target year (e.g. 2025) in Step 3
    4. Click "Run Aggregation"
    5. Click "Download Result" to save the file

  [B. Windows Application]
  Double-click "コメントシート集計ツール.exe"

    Steps:
    1. Click "Select Comment Sheets..." and pick your Excel files
    2. Click "Select Attendance Sheet..." (optional)
    3. Enter the target year
    4. Click "Run Aggregation" and choose where to save


--------------------------------------------------------------------------------
■ 3. Account Information (GitHub / Streamlit)
--------------------------------------------------------------------------------

  Both the GitHub and Streamlit accounts were created using the
  WiCAN Google account.

  How to log in:
    On the login page of each service, select "Sign in with Google"
    and use the WiCAN Google account.

  Accessing the Streamlit dashboard (deploy, restart, settings, logs):
    Option 1: Search "streamlit community cloud" on Google
    Option 2: Enter the following URL in your browser
              https://share.streamlit.io/

    * The dashboard lets you restart the app, change settings, and view logs.


--------------------------------------------------------------------------------
■ 4. Updating Settings When the University's Excel Format Changes
--------------------------------------------------------------------------------

If the column positions in the Excel files change, you need to update the
configuration in the source code.

  ====================================================================
  Windows App: use the built-in settings screen
  ====================================================================

  Click the "⚙ Advanced Settings" button in the top-right corner of
  the app window. Change the values and click Save.

  ====================================================================
  Web Version: edit the source code and push to GitHub
  ====================================================================

  STEP 1 — Find the settings

    Open "src/aggregator.py" in any text editor (Notepad is fine).
    Near the top of the file, find the CONFIG section:

    -------------------------------------------------------
    CONFIG = {
        # Comment sheet columns
        "COL_SUB_ID":  0,   # column containing Submission ID
        "COL_COURSE":  2,   # column containing Course name
        "COL_NAME":    4,   # column containing Student name
        "COL_ID":      5,   # column containing Student ID
        "COL_COMMENT": 6,   # column containing Comment text

        # Attendance sheet
        "ATT_SKIP_ROWS": 6, # number of header rows to skip
        "ATT_COL_ID":    1, # column with Student ID in attendance sheet
        "ATT_COL_NAME":  2, # column with Student name in attendance sheet
    }
    -------------------------------------------------------

  STEP 2 — Understand column numbers

    Columns are numbered starting from 0:

      A=0,  B=1,  C=2,  D=3,  E=4,
      F=5,  G=6,  H=7,  I=8,  J=9 ...

    Example: "Student name" moved from column E to column D
      Before: "COL_NAME": 4,
      After:  "COL_NAME": 3,

    Example: Attendance sheet data starts at row 7
      "ATT_SKIP_ROWS": 6,   (= 7 minus 1)

  STEP 3 — Update the file on GitHub (browser only, no Git required)

    1. Open your GitHub repository in a browser
    2. Click on the "src" folder
    3. Click on "aggregator.py"
    4. Click the pencil icon (Edit this file) in the top right
    5. Change the numbers in the CONFIG block
    6. Click "Commit changes" at the bottom
    7. The web app will automatically redeploy within a few minutes


--------------------------------------------------------------------------------
■ 4. Rebuilding the Windows EXE
--------------------------------------------------------------------------------

After changing the source code, rebuild the EXE to apply changes to the
Windows application.

  Steps:
  1. Double-click "build_exe.bat"
  2. A command window opens and the build starts automatically
  3. When "Build complete!" appears, press any key to close
  4. A new "コメントシート集計ツール.exe" is created in the root folder

  Notes:
  - PyInstaller is installed automatically on first run (takes a few minutes)
  - Temporary files (build/, dist/, *.spec) are cleaned up automatically


--------------------------------------------------------------------------------
■ 5. Folder Structure (reference)
--------------------------------------------------------------------------------

  comment_sheet_aggregator/
  |
  |-- コメントシート集計ツール.exe  ... Windows application
  |-- build_exe.bat               ... Script to rebuild the EXE
  |-- run_web_app.bat             ... Script to run the web app locally
  |-- requirements.txt            ... Python dependency list
  |-- README.txt                  ... This file
  |
  |-- src/
  |   |-- aggregator.py           ... ★ Core logic — EDIT THIS for config changes
  |   |-- gui_app.py              ... Windows app UI
  |   `-- streamlit_app.py        ... Web app UI
  |
  `-- xlrd_legacy/                ... Library for reading .xls files (do not edit)


--------------------------------------------------------------------------------
■ 6. Files That Can Be Safely Deleted
--------------------------------------------------------------------------------

These files have no effect on the application and can be deleted if desired:

  * build/           folder (PyInstaller intermediate files)
  * dist/            folder (staging area during EXE build)
  * *.spec           files  (build configuration)
  * __pycache__/     folders (Python cache)
  * user_config.json file   (Windows app settings — deleting resets to defaults)

================================================================================
