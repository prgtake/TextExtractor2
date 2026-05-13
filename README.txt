=====================================================
 TextExtractor2 v1.0.0 - README
=====================================================

人間が整理した「フォルダ分け」をそのままAIの解析単位にする。
バラバラの資料を一つの文脈として読み解き、最強のデータベースを構築します。

-----------------------------------------------------
 1. 主な機能
-----------------------------------------------------

・サブフォルダ単位解析: フォルダ内の複数ファイルを統合して1レコードに集約。
・AI指示文相談チャット: 対話形式で最適な抽出・生成ルールを自動作成。
・SQLite自動保存: 抽出結果をプロ仕様のDBに蓄積し、重複処理は自動スキップ。
・SQLiteツール連携: GeminiがDBを参照・更新しながら高度な推論を実行。
・多彩なファイル対応: PDF, 画像, Word, Excel, PowerPoint, 一太郎, メール等。
・堅牢なエラー処理: API制限時も指数バックオフで粘り強く再試行。

-----------------------------------------------------
 2. 活用シーン
-----------------------------------------------------

・教育/スクール: 生徒ごとのフォルダから指導カルテを自動生成。
・営業/CS: 取引先ごとの議事録や見積書を統合して商談DBを作成。
・医療/介護: 患者ごとの資料を統合し、症状や投薬の推移をサマリー化。
・物件/店舗管理: 拠点ごとの調査資料から、共通基準の比較表を作成。

-----------------------------------------------------
 3. 高精度な解析のためのコツ (重要)
-----------------------------------------------------

AI（Gemini）は「頭の中」だけで複雑な計算や推論を行うと、ミスをしやすくなる性質があります。

・計算の根拠を先に出力させる:
  「合計点」だけを求めるのではなく、必ず「各科目の点数」も出力項目に含め、AIに先に書き出させるようにしてください。
  AIは自分が直前に書き出した正確な数値を参照することで、計算精度が劇的に向上します。
  例：「国語点数, 数学点数, 5科目合計点」の順で項目を作る。

・【抽出】と【生成】の明示:
  指示文内で「【抽出】: 資料からそのまま抜き出す」「【生成】: 資料を元にAIが考える」と役割を分けることで、推測によるハルシネーション（嘘）を防ぐことができます。

-----------------------------------------------------
 4. セットアップ
-----------------------------------------------------

・Python 3.9 以上が必要です。
・Google Gemini API キーが必要です。

ライブラリのインストール：
pip install google-genai pandas openpyxl python-docx python-pptx extract-msg olefile striprtf odfpy

-----------------------------------------------------
 5. 使い方
-----------------------------------------------------

1. 処理対象の親フォルダを選択します。
2. 「AIと指示文相談」で、抜き出したい項目についてAIと決めます。
3. 作成した指示ファイル(.txt)を読み込みます。
4. 「指示文を実行」を押し、完了を待ちます。
5. 「DBをCSV出力」で結果を書き出します。

-----------------------------------------------------
 6. ライセンス
-----------------------------------------------------
Copyright (c) 2026 Datan (データン).
Licensed under the MIT License.

=====================================================
 TextExtractor2 v1.0.0 - README (English)
=====================================================

Treat "Folder Structures" organized by humans as the unit of AI analysis.
Read scattered documents as a single context and build the ultimate database.

-----------------------------------------------------
 1. Main Features
-----------------------------------------------------

- Subfolder-Unit Analysis: Integrate multiple files in a folder into a single record.
- AI Prompt Consultation Chat: Automatically create optimal extraction/generation rules through dialogue.
- Automatic SQLite Saving: Accumulate results in a professional-grade database and skip processed folders.
- SQLite Tool Integration: Gemini performs advanced reasoning while referencing/updating the DB.
- Diverse File Support: PDF, Images, Word, Excel, PowerPoint, Ichitaro, Emails, etc.
- Robust Error Handling: Persistent retry with exponential backoff during API limits.

-----------------------------------------------------
 2. Tips for High Accuracy (Important)
-----------------------------------------------------

AI (Gemini) tends to make mistakes when performing complex calculations only "in its head."

- Output Calculation Sources First:
  Instead of asking only for a "Total," include the "source values" in the output items and have the AI write them out first. AI dramatically improves its calculation accuracy by referring to the precise numbers it has just written.
- Explicit use of 【Extraction】 and 【Generation】:
  Distinguishing between extracting facts and generating insights helps prevent hallucinations.

-----------------------------------------------------
 3. Setup
-----------------------------------------------------

- Python 3.9 or higher is required.
- Google Gemini API Key is required.

Install libraries:
pip install google-genai pandas openpyxl python-docx python-pptx extract-msg olefile striprtf odfpy

-----------------------------------------------------
 4. How to Use
-----------------------------------------------------

1. Select the parent folder for processing.
2. Use "AI Prompt Consultation" to decide output items with AI.
3. Load the created prompt file (.txt).
4. Click "Run Prompt" and wait for completion.
5. Use "Export DB to CSV" to write out the results.

-----------------------------------------------------
 5. License
-----------------------------------------------------
Copyright (c) 2026 Datan.
Licensed under the MIT License.
