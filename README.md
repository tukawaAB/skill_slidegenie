# skill_slidegenie

[Claude Code](https://docs.claude.com/en/docs/claude-code) / [Cowork](https://www.anthropic.com) のスキルとして動作する、ローカル完結型AIスライド生成ツール **SlideGenie** のリポジトリです。

各自のGoogle API Keyを通じて **Google Gemini API と1対1で通信** し、コンサルティングファーム品質のPowerPointスライドをローカルに生成します。

---

## 📦 リポジトリ構成

このリポジトリのファイルは、そのまま **Claudeスキルのディレクトリ構造** になっています。
clone / ZIPダウンロード後、フォルダ全体を `.claude/skills/slidegenie/` に配置すればそのまま動作します。

```
skill_slidegenie/                        ← このリポジトリ
├── README.md                            ← このファイル
├── SKILL.md                             ← Claude向けスキル定義
├── .env.example                         ← APIキー設定テンプレート
├── .gitignore                           ← .env 等を除外
├── pyproject.toml                       ← Python依存定義
├── uv.lock                              ← 依存ロックファイル
├── discussion_slides.json               ← バッチ生成サンプル
└── src/slidegenie/                      ← ソースコード一式
    ├── auth.py, cli.py, pipeline.py, gemini_client.py, __main__.py
    ├── image_gen/         (graphic / flow / matrix generators)
    ├── json_gen/          (OCR → JSON 変換、後処理)
    ├── slide_gen/         (JSON → PPTX 組立)
    ├── shapes/            (シェイプ描画・テキストフィッティング)
    ├── utils/             (common / constants / logger)
    ├── prompts/addin/     (18 Jinja2 テンプレート)
    ├── prompts/tone-manner/ (ビジュアルスタイル定義)
    └── templates/template.pptx
```

---

## 🚀 導入手順（3ステップ）

### 1. リポジトリを `.claude/skills/slidegenie/` に配置

**オプション A: git clone（推奨）**

```bash
# Windows (Git Bash) / macOS / Linux
git clone <このリポジトリのURL> ~/.claude/skills/slidegenie

# Windows PowerShell
git clone <このリポジトリのURL> $HOME\.claude\skills\slidegenie
```

**オプション B: ZIPダウンロード**

1. GitHubの「Code」→「Download ZIP」でダウンロード
2. 展開したフォルダ名を `slidegenie` に変更
3. 以下に配置:
   - **Windows**: `C:\Users\<ユーザー名>\.claude\skills\slidegenie\`
   - **macOS / Linux**: `~/.claude/skills/slidegenie/`

配置後、`.claude/skills/slidegenie/SKILL.md` が存在する状態になっていればOKです。

### 2. Python環境のセットアップ

Python 3.12以上が必要です。配置したフォルダ内で実行:

```bash
cd ~/.claude/skills/slidegenie      # Windows: cd $HOME\.claude\skills\slidegenie

# uv を使う場合（推奨・ロックファイル効く）
uv sync

# pip を使う場合
python -m venv .venv
source .venv/bin/activate            # Windows: .\.venv\Scripts\activate
pip install -e .
```

[uv のインストール方法](https://docs.astral.sh/uv/getting-started/installation/):
- Windows (PowerShell): `powershell -c "irm https://astral.sh/uv/install.ps1 | iex"`
- macOS / Linux: `curl -LsSf https://astral.sh/uv/install.sh | sh`

### 3. Google API Key の設定

1. [Google AI Studio](https://aistudio.google.com/app/apikey) でAPIキーを取得
2. `.env.example` を `.env` にコピー
3. `.env` 内の `GOOGLE_API_KEY` に自分のキーを記入

```bash
cp .env.example .env     # Windows PowerShell: Copy-Item .env.example .env
```

```ini
# .env
GOOGLE_API_KEY=AIzaSy...(your key)...
```

> `.env` は `.gitignore` で除外済みです。誤ってコミットされる心配はありません。

### 動作確認

```bash
python -m slidegenie generate -p "DX推進の3つのステップ" -o test.pptx
```

`test.pptx` が生成されれば完了です。Claude Code / Cowork を再起動すればスキルとして自動認識されます。

---

## 🤖 Claudeスキルとしての使い方

`.claude/skills/slidegenie/` に配置済みの状態で、Claude Code や Cowork との会話に以下のように依頼するだけで、Claudeが自動でCLIを実行してスライドを生成します:

- 「slidegenieでDX推進のスライドを作って」
- 「業務フローのpptxを生成して」
- 「競合比較のマトリクススライドを作って」

スキル定義の詳細は `SKILL.md` を参照してください。

---

## 🔒 プライバシー・セキュリティ設計

SlideGenie は **ユーザー × Google Gemini の1対1通信** のみで動作します。

### ✅ データフロー

```
┌────────────────────┐                        ┌──────────────────────┐
│   個別ユーザーのPC  │                       │   Google Gemini API  │
│                     │                       │                      │
│  slidegenie CLI ───┼──── GOOGLE_API_KEY ───→│  gemini-3-pro-image  │
│   ├─ 画像生成要求   │                       │  gemini-3-flash      │
│   ├─ OCR要求        │ ←────  画像/JSON ─────│                      │
│   └─ pptx組立       │                       │                      │
│        ↓            │                       │                      │
│   output.pptx       │                       │                      │
└────────────────────┘                        └──────────────────────┘
```

- プロンプト・生成画像・OCR結果は **ユーザーのマシンとGoogleのAPIエンドポイントの間でしかやり取りされません**
- 当リポジトリが管理する中継サーバー・プロキシは存在しません
- 生成されたPPTXは **ローカルファイルシステム** にのみ保存されます

### ✅ 認証情報の扱い

- 認証は **ユーザー自身のGoogle API Key** または Google ADC（Application Default Credentials）のみ
- APIキーは `.env` にローカル保存され、`.gitignore` 済みで公開リポジトリには入りません

### ⚠️ 注意

- Google の利用規約・プライバシーポリシーに基づき、送信したプロンプト・画像は Google の処理対象になります。機密情報の入力は各自の判断で行ってください
- Gemini API の利用料金は **ユーザー自身のGoogle Cloudアカウント** に課金されます

---

## 🎯 何ができるか

| スライド種別 | 用途 | 例 |
|------------|------|-----|
| `graphic` | 概念図・組織図・関係図 | "組織体制図", "DX推進の全体像" |
| `flow` | プロセス・ワークフロー・タイムライン | "業務改善のステップ", "導入プロセス" |
| `matrix` | 比較表・評価マトリクス | "競合比較", "SWOT分析" |

### デザイン仕様
- サイズ: 13.3333 × 7.5 inch（16:9 ワイドスクリーン）
- フォント: Yu Gothic Light
- カラー: トープ（#6D6358）/ ライトベージュ（#D9D0C1）/ 白背景 / ダークグレー（#333333）テキスト
- スタイル: コンサル系ミニマル（グラデーション・影・3D効果なし）

---

## 💻 CLI 使用例

```bash
# 単一スライド生成（タイプ自動選択）
python -m slidegenie generate -p "DX推進の3つのステップ" -o output.pptx

# タイプ指定
python -m slidegenie generate -p "業務フロー改善" -o flow.pptx -t flow

# 画像のみモード（OCRせず画像を貼り付けただけのpptx）
python -m slidegenie generate -p "組織体制図" -o org.pptx -m image

# 既存画像からPPTXへ変換
python -m slidegenie convert -i slide.png -o converted.pptx

# バッチ生成（JSONコンフィグから複数スライド）
python -m slidegenie batch -c specs.json -o deck.pptx -w 4

# 複数PPTXのマージ
python -m slidegenie merge -i a.pptx -i b.pptx -o merged.pptx
```

バッチコンフィグ（`specs.json`）例:
```json
[
  {"prompt": "DX推進の概要",  "make_type": "graphic"},
  {"prompt": "導入プロセス",  "make_type": "flow"},
  {"prompt": "競合比較",       "make_type": "matrix"}
]
```

---

## ⚠️ 既知の制限

- **scheduleスライドタイプは未対応です。** プロンプトテンプレート（`make_scheduleimage_ja/en.j2`）は同梱されていますが、パイプラインからの正式な呼び出しは未サポートのため、`-t schedule` の実行は避けてください。
- Gemini APIのレート制限に達した場合、自動で最大10回のリトライ（指数バックオフ）が行われます。
- 日本語スライドでは「Yu Gothic Light」フォントがインストールされている必要があります（PowerPointで表示する場合）。

---

## 🏗 アーキテクチャ概要

4ステップ パイプライン:

```
User Prompt
    │
    ├─ ①  Language Detection  (is_english → en / ja)
    ├─ ②  Type Selection      (Gemini LLM → graphic / flow / matrix)
    ├─ ③  Image Generation    (gemini-3-pro-image-preview → PIL Image)
    └─ ④  OCR → PPTX
         ├─ a. OCR            (gemini-3-flash-preview → JSON with texts/icons/shapes)
         ├─ b. Post-process   (font normalization, text alignment)
         └─ c. PPTX Build     (python-pptx + template → .pptx)
```

### 依存パッケージ

```
google-genai   >= 1.55.0   # Gemini API クライアント
google-auth    >= 2.43.0   # Google 認証
python-pptx    >= 0.6.23   # PPTX 生成（ローカル）
pillow         >= 11.0.0   # 画像処理（ローカル）
jinja2         >= 3.1.6    # プロンプトテンプレート
numpy          >= 2.2.0    # 数値処理
click          >= 8.1.8    # CLI フレームワーク
```

### 主要ファイル

| ファイル | 役割 |
|---------|------|
| `src/slidegenie/pipeline.py` | メインオーケストレータ。`generate_slide()`, `batch_generate()`, `merge_pptx()` |
| `src/slidegenie/cli.py` | Click CLI コマンド: `generate` / `convert` / `batch` / `merge` |
| `src/slidegenie/auth.py` | Gemini 認証（API Key / ADC） |
| `src/slidegenie/gemini_client.py` | `GEMINIClient` - 画像生成・OCR・チャット・リトライ処理 |
| `src/slidegenie/image_gen/` | 種別ごとの画像プロンプト生成器 |
| `src/slidegenie/json_gen/builder.py` | OCR結果を構造化JSONへ変換 |
| `src/slidegenie/slide_gen/builder.py` | JSON → PPTX 組立 |
| `src/slidegenie/shapes/object_function.py` | シェイプ描画・テキストフィッティング・バリデーション |

---

## 🔧 カスタマイズ

### 新しいスライドタイプの追加

1. `src/slidegenie/image_gen/newtype.py` を作成（`graphic.py` をテンプレに）
2. `src/slidegenie/prompts/addin/make_newtypeimage_{ja,en}.j2` を作成
3. `src/slidegenie/image_gen/builder.py` の `prompt_to_image()` に分岐追加
4. `src/slidegenie/prompts/addin/select_maketype.j2` に新タイプの説明を追加
5. `src/slidegenie/slide_gen/builder.py` の `json_to_pptx()` の有効タイプチェックに追加

### ビジュアルスタイルの変更

`src/slidegenie/prompts/tone-manner/_tone_and_manner_common_ja.j2`（または `_en.j2`）を編集:
- `color_palette`（カラー）
- `typography`（フォント/サイズ）
- `layout_geometry`（ヘッダー/リード/マージン比率）
- `graphics`（アイコン・チャートスタイル）

### テキストフィッティングの調整

`src/slidegenie/utils/constants.py` の `PowerPointConfig`:
- `MIN_FONT_SIZE` / `MAX_FONT_SIZE`
- `TEXT_PADDING_INCHES`
- `SHAPE_TEXT_AREA_RATIOS`

---

## 🐛 トラブルシューティング

| 症状 | 対処 |
|------|------|
| `GOOGLE_API_KEY not found` | `.env` の配置・値と仮想環境アクティブ状態を確認 |
| `ModuleNotFoundError: slidegenie` | `uv sync` または `pip install -e .` を実行 |
| 画像生成が失敗 | `gemini_client.py` のログ確認、Geminiレート制限に注意 |
| OCR結果が空 | OCRプロンプト・画像サイズを確認。`perform_ai_ocr()` のログ参照 |
| シェイプが欠ける | `shapes/object_function.py` の `validate_shape_data()` ログを確認 |
| テキストが切れる | `utils/constants.py` のフォントサイズ設定を調整 |
| スキルが自動認識されない | Claudeを再起動、`SKILL.md` が `.claude/skills/slidegenie/` 直下にあるか確認 |
| 日本語フォントがずれる | PowerPointに「Yu Gothic Light」がインストールされているか確認 |
| `-t schedule` でエラー | スケジュールタイプは未対応なので使用しないでください |

デバッグログ有効化:
```python
from slidegenie.utils.logger import get_logger
logger = get_logger("pipeline")
logger.setLevel("DEBUG")
```

---

## 📄 ライセンス・利用について

社内・個人利用を想定した実装です。利用前に所属組織のポリシーを確認してください。
Geminiを使う際のデータの扱いは、Googleの[利用規約](https://ai.google.dev/terms)・[プライバシーポリシー](https://policies.google.com/privacy)に従います。

## 🙋 コントリビューション

バグ報告・機能追加は Issue または Pull Request でお知らせください。
