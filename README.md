# TextExtractor2 (サブフォルダ単位抽出)

Google Gemini API (2.5 Flash Lite / 2.0 Flash) を活用した、サブフォルダ単位での高度な統合解析・構造化ツールです。
人間が整理した「フォルダ構造」をそのまま AI の解析単位として扱い、バラバラの資料を一瞬で「意味のあるデータベース」に変換します。

## 主要な特徴 (TextExtractor2)

- **サブフォルダ単位の統合解析:** 指定したフォルダ内の各サブフォルダを「一つの人格」や「一つの案件」として認識。中にある複数の PDF、画像、Office 文書、メール等を AI がまとめて読み解き、1 レコードとして抽出します。
- **AI 指示文相談チャット:** 「どう指示を書けばいいかわからない」を AI が解決。チャット形式でやりたいことを伝えるだけで、コピペ可能な最適な指示文（プロンプト）を即座に提案します。
- **SQLite データベース自動生成:** 抽出結果は自動的に SQLite (`.db`) へ保存。数千のフォルダを処理してもデータが散らばらず、いつでも CSV で出力可能です。
- **SQLite ツール連携:** AI が既存のデータベースを直接参照・更新可能。マスタ照合や重複チェックを含めた高度なデータ生成が行えます。
- **堅牢な抽出能力:** 標準的な PDF/画像に加え、一太郎 (.jtd)、古い Office 形式 (.doc, .ppt)、メール (.eml, .msg)、RTF、ODT まで幅広く対応。
- **プロフェッショナルな設計:** API 制限に対する指数バックオフ・リトライ、処理済みスキップ機能、コンパクトで洗練された UI を備えています。

---

## 高精度な解析のためのコツ (重要)

AI（Gemini）は「頭の中」だけで複雑な計算や推論を行うと、ミスをしやすくなる性質があります。

- **計算の根拠を先に出力させる:**
  「合計点」だけを求めるのではなく、必ず「各科目の点数」も出力項目に含め、AIに先に書き出させるようにしてください。
  **AIは自分が直前に書き出した正確な数値を参照することで、計算精度が劇的に向上します。**
  例：「国語点数, 数学点数, 5科目合計点」の順で項目を作る。

- **【抽出】と【生成】の明示:**
  指示文内で「【抽出】: 資料からそのまま抜き出す」「【生成】: 資料を元にAIが考える」と役割を分けることで、推測によるハルシネーション（嘘）を防ぐことができます。

---

## 活用のアイデア

1.  **生徒・顧客カルテの自動作成:** 氏名ごとのフォルダ内のテスト結果や面談記録を統合し、一人ひとりの傾向を DB 化。
2.  **複雑な裁判・法律案件の整理:** 事件ごとのフォルダにある証拠書類や判決文を通読し、争点や時系列を整理。
3.  **年度別の経営・決算分析:** 年度ごとのフォルダに格納された諸表をまとめ、主要指標の推移を一覧化。
4.  **物件・支店管理:** 拠点ごとの写真や報告書から、現場に行かずとも全拠点の状態を同じ基準でスコアリング。

---

## セットアップ

### 必要条件
- Python 3.9 以上 (Python 3.13 対応済み)
- Google Gemini API キー

### インストール
```bash
pip install google-genai pandas openpyxl python-docx python-pptx extract-msg olefile striprtf odfpy pyinstaller
```

## 使い方
1. プログラムを起動します。
2. **「処理対象フォルダを指定」**: サブフォルダ群が含まれる親フォルダを選択します。
3. **「AIと指示文相談」**: AI と相談して「出力項目」や「指示内容」を決め、指示ファイル (.txt) を作成します。
4. **「指示ファイル(txt)を指定」**: 作成したファイルを読み込みます。
5. **「指示文を実行」**: 解析が開始され、結果が SQLite に蓄積されます。
6. **「DBをCSV出力」**: 必要に応じて Excel 等で開ける CSV を書き出します。

## ライセンス
Copyright (c) 2026 Datan (データン).
Licensed under the MIT License.

---

# TextExtractor2 (Subfolder-Unit Extraction)

A highly advanced integrated analysis and structuring tool utilizing Google Gemini API (2.5 Flash Lite / 2.0 Flash). 
It treats the "folder structure" organized by humans as the unit of AI analysis, transforming scattered documents into a "meaningful database" in an instant.

## Key Features (TextExtractor2)

- **Subfolder-Unit Integrated Analysis:** Recognizes each subfolder within a target folder as a "single persona" or "single case." AI reads multiple PDFs, images, Office documents, emails, etc., inside and extracts them as a single record.
- **AI Prompt Consultation Chat:** Solves the problem of "not knowing how to write instructions." Just tell the AI what you want to do in a chat format, and it will immediately propose the best copiable prompt.
- **Automatic SQLite Database Generation:** Extraction results are automatically saved to SQLite (.db). Data stays organized even after processing thousands of folders and can be exported to CSV at any time.
- **SQLite Tool Integration:** AI can directly reference and update existing databases. Allows for advanced data generation including master matching and duplication checks.
- **Robust Extraction Capabilities:** Supports a wide range of formats from standard PDF/Images to Ichitaro (.jtd), old Office formats (.doc, .ppt), Emails (.eml, .msg), RTF, and ODT.
- **Professional Design:** Features exponential backoff and retry for API limits, processed skip functionality, and a compact, polished UI.

---

## Tips for High Accuracy (Important)

AI (Gemini) tends to make mistakes when performing complex calculations or reasoning only "in its head."

- **Output Calculation Sources First:**
  Instead of asking only for the "Total Score," always include the "Scores for each subject" in the output items and have the AI write them out first.
  **AI dramatically improves its calculation accuracy by referring to the precise numbers it has just written.**
  Example: Create items in the order of "Math Score, English Score, Total Score."

- **Explicit use of 【Extraction】 and 【Generation】:**
  By distinguishing roles in the instructions—"【Extraction】: Extract exactly from materials" and "【Generation】: AI thinks based on materials"—you can prevent hallucinations (false information) caused by guesswork.

---

## Usage Ideas

1.  **Automatic Student/Customer Records:** Integrate test results and interview records in subfolders for each individual to create a database of trends.
2.  **Organization of Complex Legal Cases:** Read through evidence documents and judgments in case folders to organize issues and timelines.
3.  **Annual Business/Financial Analysis:** Summarize various statements stored in annual folders to list the transition of key indicators.
4.  **Property/Branch Management:** Score all locations based on the same criteria from survey materials in branch folders without visiting the sites.

---

## Setup

### Requirements
- Python 3.9 or higher (Python 3.13 supported)
- Google Gemini API Key

### Installation
```bash
pip install google-genai pandas openpyxl python-docx python-pptx extract-msg olefile striprtf odfpy pyinstaller
```

## How to Use
1. Start the program.
2. **"Select Target Folder"**: Choose the parent folder containing subfolders.
3. **"AI Prompt Consultation"**: Decide on "Output Items" and "Instructions" with AI and create a prompt file (.txt).
4. **"Select Prompt File (txt)"**: Load the created file.
5. **"Run Prompt"**: Analysis starts, and results are accumulated in SQLite.
6. **"Export DB to CSV"**: Export the results to a CSV file that can be opened in Excel, etc.

## License
Copyright (c) 2026 Datan.
Licensed under the MIT License.
