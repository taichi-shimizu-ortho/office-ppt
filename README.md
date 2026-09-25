# office-ppt

既存の pptx を JSON テンプレートにして、素材画像を差し込んだ pptx（ポスター・スライド）を作るツール。

## セットアップ

```sh
uv sync
```

## 使い方

### 1. もとの pptx からテンプレートを作る

```sh
uv run extract.py ~/…/20261003産業医大学会/poster.pptx
```

同じフォルダに `poster.json` ができる。もとの pptx にある画像はすべて差し替える前提なので、
画像は書き出さず、枠（位置とサイズ）だけが `"path": null` として入る。

### 2. JSON を編集する

```jsonc
{
  "base": "poster.pptx",          // もとの pptx（JSON からの相対パス）
  "output": "poster_out.pptx",    // 出力先。base と同じにはできない
  "image_root": "imput",          // 画像フォルダ（JSON からの相対パス、または絶対パス）
  "unit": "cm",
  "slide_size": [84.1, 118.9],    // 参考情報（変更しても反映されない）
  "slides": [
    {
      "index": 0,
      "elements": [
        {"id": 2, "name": "TextBox 1", "type": "text", "box": [3.0, 3.0, 78.0, 8.0],
         "text": ["タイトル", "著者名"]},
        {"id": 3, "name": "Picture 2", "type": "image", "box": [3.0, 15.0, 38.0, 30.0],
         "path": "fig1.png", "fit": "contain"}
      ]
    }
  ]
}
```

- `id` で pptx 上の図形と対応づける。`name` は目印（PowerPoint の「選択ウィンドウ」に出る名前）。
- `box`: `[left, top, width, height]`（cm）。変更すると移動・リサイズする。
  グループ内の図形（`"group"` あり）はグループ内の座標。
- 画像 `fit`
  - `contain`: 枠内に収める（縦横比維持、既定）
  - `cover`: 枠を埋め、はみ出しをトリミング
  - `stretch`: 枠に合わせて伸縮
- **画像がない場合**（`path` が `null`、またはファイルが存在しない）は点線の枠を残す。
- テキストは段落ごとの配列。変更した段落だけ書き換え、その段落の先頭の文字書式が全体に使われる
  （上付き文字など、段落内で書式が混在する部分は PowerPoint 上で直す）。段落内改行は `\n`。
- JSON から消した要素は、もとの pptx のまま残る（もとの画像を残したいときは、その要素を消す）。
- **要素の追加**: `id` を書かない要素は新しく追加される（`box` 必須、スライド最前面・グループの外に置かれる）。
  `"copy_from": <id>` を付けるとその図形を複製して書式（フォント・塗り・枠線など）を引き継ぐ。
  付けない場合、画像はそのまま貼り付け（`path` がなければ点線の枠）、テキストは標準書式のテキストボックス。

  ```jsonc
  {"name": "Fig2 ラベル MT", "type": "text", "copy_from": 7, "box": [72.56, 22.9, 15.75, 1.966], "text": ["MT"]},
  {"name": "Fig2 MT 弱拡大", "type": "image", "box": [72.56, 24.9, 7.775, 13.82], "path": "capMT 3-1.jpg", "fit": "cover"}
  ```

### 3. 生成する

```sh
uv run build.py ~/…/20261003産業医大学会/poster.json
```

もとの pptx は上書きせず、`output` に保存する。
