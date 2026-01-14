# Slide Studio

PowerPointプレゼンテーションをXML直接編集で作成・修正するスキル。

## セットアップ

```bash
cd slide-studio
npm install
```

## ディレクトリ構造

```
slide-studio/
├── SKILL.md              # スキル定義（Claude用）
├── README.md             # 開発ドキュメント（このファイル）
├── package.json
├── scripts/
│   ├── unpack.js         # PPTX展開
│   ├── pack.js           # PPTX作成
│   ├── apply-potx.js     # テンプレート適用
│   ├── add-slide.js      # スライド追加
│   ├── edit-text.js      # テキスト編集
│   ├── add-image.js      # 画像追加
│   ├── add-shape.js      # シェイプ自由配置
│   ├── clone-slide.js    # スライド複製
│   ├── delete-slide.js   # スライド削除
│   ├── list-layouts.js   # レイアウト一覧
│   ├── list-slides.js    # スライド一覧
│   ├── visualize-slide.js # ASCII構造表示
│   ├── create-base.js    # base.pptx生成
│   └── lib/
│       ├── constants.js      # 定数定義
│       ├── xml-utils.js      # XML操作
│       ├── relationships.js  # .rels管理
│       └── content-types.js  # Content Types管理
├── references/
│   ├── ooxml-reference.md    # OOXML仕様
│   └── slide-patterns.md     # XMLパターン
└── assets/
    └── base.pptx             # ベースPPTX（11レイアウト）
```

## OOXML構造

PPTXはZIPアーカイブで、以下の構造を持つ:

```
[Content_Types].xml          # コンポーネント登録
_rels/.rels                  # ルート関係
ppt/
├── presentation.xml         # プレゼンメタデータ
├── _rels/presentation.xml.rels
├── slides/
│   ├── slide1.xml
│   └── _rels/slide1.xml.rels
├── slideLayouts/            # 11種類のレイアウト
├── slideMasters/            # マスタースライド
├── theme/                   # テーマ（色、フォント）
└── media/                   # 画像等
```

## テスト方法

```bash
# 1. 展開
node scripts/unpack.js assets/base.pptx test-work

# 2. テンプレート適用（オプション）
node scripts/apply-potx.js test-work /path/to/template.potx

# 3. スライド追加
node scripts/add-slide.js test-work --layout 2

# 4. パック
node scripts/pack.js test-work test-output.pptx

# 5. 目視確認
# PowerPointまたはGoogle Slidesで開いて確認
```

## 依存関係

- **Node.js** >= 14
- **jszip** - ZIP操作

## パッケージング

```bash
cd /path/to/pptx-plus
zip -r slide-studio.skill slide-studio/ -x "*.DS_Store" -x "*node_modules/*"
```
