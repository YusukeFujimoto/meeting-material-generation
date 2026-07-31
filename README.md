# tkc-kpi-slide

TKC形式の科目別月次推移表（科目残高データ）から、経営会議用のKPIサマリーをPowerPoint（16:9ワイド、ネイティブグラフ、コメント欄付き）にまとめる[Claude Code](https://docs.claude.com/claude-code) スキルです。

売上高・限界利益率・労働分配率をTKC科目残高データから自動計算し、受注実績データがあれば受注高も併せて取り込みます。得意先別の受注TOP10のような会社特有の指標にも対応しています。同じTKC会計事務所を使うグループ会社で使い回すことを想定し、会社ごとの科目コードの違いや部門構成の違いを吸収できるように作られています。

## できること

- TKC形式（科目コード／固変区分／勘定科目名＋月ごとの借方・貸方・残高の3列組）のExcel（`.xlsx` / 旧`.xls`）から、売上高・限界利益率・労働分配率・（受注実績データがあれば）受注高を月次で算出
- 1つのファイルに複数事業部・複数部門が入っている場合の分割集計
- 直近月データが未確定の概算値になっていないかの信頼性チェック
- 得意先別TOP10など、会社特有の指標を表・箇条書きで追加できる汎用枠（`extraSection`）
- PowerPointネイティブグラフ（画像ではなく編集可能なグラフ）での出力
- Windows環境でのQA（PowerPoint COM経由のスライド画像書き出し）

## セットアップ

Claude Codeのスキルディレクトリにこのリポジトリをクローン（またはコピー）してください。

```bash
git clone https://github.com/YusukeFujimoto/meeting-material-generation.git ~/.claude/skills/tkc-kpi-slide
```

### 依存パッケージ

初回利用時にClaude Codeが自動でインストールしますが、手動で入れる場合は以下です。

```bash
# Node.js（PowerPoint生成用）
cd scripts
npm install pptxgenjs

# Python（データ読み取り・QA用）
pip install openpyxl xlrd lxml defusedxml Pillow "markitdown[pptx]" python-pptx pywin32
```

- `xlrd` は `.xls` 旧バイナリ形式のTKCファイルを読むのに使用
- `pywin32` はWindows環境でのPowerPoint COM経由のQA画像書き出しに使用（`scripts/export_slides_win32.py`）

## 使い方

Claude Codeに、TKC形式の科目残高データ（Excel）を添付して「経営会議資料を作って」のように依頼するだけで、このスキルが自動的に呼び出されます。

手動でスライド生成だけ試したい場合：

```bash
node scripts/build_deck.js scripts/sample.data.json out.pptx
```

`scripts/sample.data.json` は架空のサンプル会社データです。実データを渡す場合のフィールド仕様は [`SKILL.md`](./SKILL.md) と [`references/design-system.md`](./references/design-system.md) を参照してください。

## ファイル構成

```
SKILL.md                       # スキル本体の指示書（全体の流れ・判断基準）
references/
  calc-rules.md                # KPI計算式・データの読み方・厳守事項
  design-system.md             # スライドの見た目・レイアウト仕様
scripts/
  extract_tkc_kpis.py          # TKCデータの月次集計ヘルパー
  build_deck.js                # PowerPoint生成本体（pptxgenjs）
  export_slides_win32.py       # PowerPoint COM経由のQA用PNG書き出し
  sample.data.json             # 架空データのサンプル（build_deck.jsへの入力例）
evals/
  evals.json                   # スキルの動作評価用シナリオ
```

## 注意

- `scripts/sample.data.json` は動作例を示すための架空データです。実際の取引先名・金額は含まれていません。
- 実データ（顧客の科目残高データ等）はこのリポジトリには含めないでください。
