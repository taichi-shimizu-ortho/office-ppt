# office-ppt

既存の pptx を JSON テンプレートにして、素材画像を差し込んだ pptx（ポスター・スライド）を作るツール。

## セットアップ

```sh
uv sync   # Python 3.13
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
  `"type": "shape"`（`copy_from` 必須）は線などをそのまま複製し、`box` の位置・サイズだけ変える（スケールバーの線など）。

  ```jsonc
  {"name": "Fig2 ラベル MT", "type": "text", "copy_from": 7, "box": [72.56, 22.9, 15.75, 1.966], "text": ["MT"]},
  {"name": "Fig2 MT 弱拡大", "type": "image", "box": [72.56, 24.9, 7.775, 13.82], "path": "capMT 3-1.jpg", "fit": "cover"}
  ```

### 3. 生成する

```sh
uv run build.py ~/…/20261003産業医大学会/poster.json
```

もとの pptx は上書きせず、`output` に保存する。

## 顕微鏡画像（CZI）を PNG にする

明視野（H-E・MT など）の CZI を、ポスターに貼れる PNG に書き出す。
CZI の読み込みには [025_image_format](https://github.com/taichi-shimizu-ortho/025_image_format) と同じ `pylibCZIrw` を使う
（Python 3.13 までしか配布されていないので、このリポジトリも 3.13 にしている）。

```sh
uv run czi_to_png.py a.czi b.czi -o ~/…/20261003産業医大学会/imput/fig2
```

1. 縁の黒い行・列を切り落とす（撮影時に上や左が 1〜2 px 黒く残ることがある）
2. タイル境界に出る 1 px の黒い線を、両隣の平均で埋める（`--no-seam-fix` で無効）
3. 明るさを調整する
   - 白バランス: 各色の背景（99.5 パーセンタイル）を `--white-level`（既定 245）に合わせる
   - 黒点: 暗い側を `--black`（既定 0.4、0 で無効）の強さで詰める
   - ガンマ: 中間調を `--gamma`（既定 1.3、1 で無効）で締める
4. `--rotate 90|180|270` で時計回りに回転（既定は回転しない）

見た目を整えるための調整なので、画像どうしの定量比較には使わない。
使った設定と µm/pixel は PNG のテキスト情報（`czi_to_png`）に残る。
