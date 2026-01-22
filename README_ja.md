# vicalc

Vimライクなターミナル表計算エディタ

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Latest Release](https://img.shields.io/github/v/release/fukuyori/vicalc)](https://github.com/fukuyori/vicalc/releases/latest)

[English](README.md)

## 概要

vicalcは、VisiCalcとVimにインスパイアされた構造指向の表計算エディタです。マウス操作やダイアログではなく、キーボードコマンドで行・列・セルを明示的に操作します。

**設計思想：**
- キーボードファースト
- 行/列モードを第一級の概念として扱う
- Vimライクなモーダル編集
- ターミナルネイティブ（GUIは不要）

## 特徴

- **Vim風ナビゲーション** - hjkl, gg, G, Ctrl+f/b/d/u
- **行/列モード** - /r と /c で編集方向を切り替え
- **数式エンジン** - 35以上の関数（SUM, VLOOKUP, IF など）
- **絶対/相対参照** - $A$1, $A1, A$1, A1
- **数式の自動補正** - 行・列の挿入・削除時に参照を自動調整
- **コピー＆ペースト** - 内部クリップボード（y/p）とシステムクリップボード（"*y/"*p）
- **ビジュアル選択** - v と V で範囲選択
- **Undo/Redo** - u で無制限のアンドゥ
- **ファイル形式** - JSON（ネイティブ）、CSV/TSVインポート・エクスポート
- **Unicode対応** - 日本語などの全角文字を正しく表示

## インストール

### バイナリから

[GitHub Releases](https://github.com/fukuyori/vicalc/releases/latest)から最新版をダウンロードしてください。

### ソースから

```bash
git clone https://github.com/fukuyori/vicalc.git
cd vicalc
cargo build --release
```

バイナリは `target/release/vicalc` に生成されます。

## クイックスタート

```bash
# 空のシートで起動
vicalc

# 既存ファイルを開く
vicalc data.json

# CSVファイルを開く
vicalc data.csv
```

## キーバインド

### モード切り替え

| キー | 動作 |
|------|------|
| `/r` | 行モードに切り替え |
| `/c` | 列モードに切り替え |
| `v` | ビジュアル選択モード |
| `V` | ビジュアル行/列モード |
| `:` | コマンドモード |
| `Esc` | ノーマルモードに戻る |

### ナビゲーション

| キー | 動作 |
|------|------|
| `h` `j` `k` `l` | 左/下/上/右に移動 |
| `gg` | 左上（A1）に移動 |
| `G` | データのある最後のセルに移動 |
| `0` | 最初の列に移動 |
| `$` | データのある最後の列に移動 |
| `Ctrl+f` | 1ページ下 |
| `Ctrl+b` | 1ページ上 |
| `Ctrl+d` | 半ページ下 |
| `Ctrl+u` | 半ページ上 |

### 編集

| キー | 動作 |
|------|------|
| `r` | セル編集（単発） |
| `R` | セル編集（連続） |
| `F2` | セル編集（内容を保持） |
| `=` | 数式入力 |
| `x` | セルをクリア |
| `dd` | 行/列を削除（モードに依存） |
| `o` | 行/列を下/右に挿入 |
| `O` | 行/列を上/左に挿入 |

### コピー＆ペースト

| キー | 動作 |
|------|------|
| `y` | 内部クリップボードにコピー |
| `p` | 内部クリップボードから貼り付け |
| `"*y` | システムクリップボードにコピー（TSV形式） |
| `"*p` | システムクリップボードから貼り付け |
| `3p` | 3回貼り付け（方向はモードに依存） |

### 列幅

| キー | 動作 |
|------|------|
| `<` | 列幅を縮小 |
| `>` | 列幅を拡大 |
| `:autowidth` | 列幅を内容に合わせて自動調整 |

### 検索

| キー | 動作 |
|------|------|
| `:/pattern` | 前方検索 |
| `:?pattern` | 後方検索 |
| `n` | 次の一致 |
| `N` | 前の一致 |

### コマンド

| コマンド | 動作 |
|----------|------|
| `:w [file]` | 保存 |
| `:e file` | ファイルを開く |
| `:q` | 終了 |
| `:wq` | 保存して終了 |
| `:export file.csv` | CSVでエクスポート |
| `:import file.csv` | CSVをインポート |
| `:goto A1` | セルに移動 |
| `:autowidth` | 全列の幅を自動調整 |
| `:autowidth A:C` | A〜C列の幅を自動調整 |
| `:insrow` | 行を挿入 |
| `:inscol` | 列を挿入 |
| `:delrow` | 行を削除 |
| `:delcol` | 列を削除 |

## サポートされている関数

### 数学・統計
`SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `ABS`, `ROUND`, `INT`, `MOD`, `POWER`, `SQRT`

### 条件付き
`IF`, `SUMIF`, `COUNTIF`, `AVERAGEIF`, `IFERROR`

### 検索・参照
`VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`

### 文字列
`LEFT`, `RIGHT`, `MID`, `LEN`, `TRIM`, `UPPER`, `LOWER`, `CONCATENATE`

### 論理
`AND`, `OR`, `NOT`

### 情報
`ISBLANK`, `ISNUMBER`, `ISTEXT`

## ファイル形式

### ネイティブ形式（JSON）

vicalcはJSONをネイティブ形式として使用し、以下を保存します：
- セルの値と数式
- 列幅
- シート名

```json
{
  "version": "1.0",
  "name": "Sheet1",
  "cells": {
    "A1": "こんにちは",
    "B1": "=SUM(A2:A10)"
  },
  "col_widths": {
    "A": 15
  }
}
```

### CSV/TSV

- インポート: `:import file.csv`
- エクスポート: `:export file.csv`
- システムクリップボードはTSV形式を使用

## 行/列モード

vicalcには「編集軸」という独自の概念があります：

- **行モード** (`/r`): 操作は横方向に展開
  - 連続ペーストは右に展開
  - `dd` は行を削除
  - `o` は下に行を挿入

- **列モード** (`/c`): 操作は縦方向に展開
  - 連続ペーストは下に展開
  - `dd` は列を削除
  - `o` は右に列を挿入

現在のモードはステータスバーに表示されます。

## ライセンス

MITライセンス。詳細は[LICENSE](LICENSE)を参照してください。

## 作者

[@fukuyori](https://github.com/fukuyori)
