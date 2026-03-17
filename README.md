# PhotoOrganizer

Windows向けの写真/動画整理アプリです。  
SDカードを検知してメディアをスキャンし、指定保存先へ以下の形式で自動整理します。

```text
[保存先]\[YYYY]\[YYYY-MM-DD]_[イベント名]\[RAW|JPG|MP4]\
```

## 仕様概要

- OS: Windows（.NET 8 / WinForms）
- 常駐: タスクトレイ常駐
- SD検知: WMI + ポーリングのフォールバック
- 分類: RAW / JPG / MP4
- 日付決定: 対象ファイル（ファイル名順の先頭）からEXIF優先で日付取得
- 保存: 同名ファイルは上書き可能、重複は日時+サイズ一致で自動スキップ

## 主な機能

1. **SDカード自動検知**
   - `Win32_VolumeChangeEvent` で挿入を監視
   - 取りこぼし対策として3秒間隔ポーリング
   - 挿入直後の未準備ドライブに対して最大5回リトライ
2. **メディア自動スキャン**
   - サブフォルダを再帰走査
   - ドットファイル除外
3. **自動振り分け保存**
   - `RAW`, `JPG`, `MP4` フォルダへ分類コピー
4. **進捗表示**
   - 画面上部に `処理中... N/総数`
   - ログは下部テキスト領域に表示（ウィンドウサイズに追従）
5. **整合性チェック**
   - コピー後に「サイズ + 更新日時(UTC)」を検証
6. **失敗ファイル自動再試行**
   - 失敗分のみ1回再実行

## 対応形式

- RAW（設定ファイルで拡張可能）
- JPG: `.jpg`, `.jpeg`
- 動画: `.mp4`, `.mov`

### デフォルトRAW拡張子

`.arw`, `.cr2`, `.cr3`, `.nef`, `.dng`, `.raf`, `.rw2`, `.orf`, `.pef`

## 対応メーカー（デフォルト拡張子基準）

- Sony（ARW）
- Canon（CR2/CR3）
- Nikon（NEF）
- FUJIFILM（RAF）
- Panasonic/LUMIX（RW2）
- OM SYSTEM / Olympus（ORF）
- PENTAX（PEF / DNG）

## 重複判定仕様

保存先に同名ファイルがある場合:

- **サイズ + 更新日時(UTC)が一致** → 重複としてスキップ
- 一致しない → 上書きコピー実行後に整合性チェック

## 設定ファイル（直接編集）

以下の `config.json` をリポジトリに含めて管理します:

```text
config.json
```

例:

```json
{
  "RawExtensions": [".arw", ".cr2", ".cr3", ".nef", ".dng", ".raf", ".rw2", ".orf", ".pef"]
}
```

- 先頭 `.` あり/なし両対応（内部で正規化）
- 変更後はアプリ再起動で反映
- Build/Publish時に `config.json` は出力先へ同梱される
- 実行時は EXE と同じフォルダの `config.json` を参照

## 前回値保存

以下の状態を保存します:

- 保存先パス
- 選択中のSDパス

保存先:

```text
%LocalAppData%\PhotoOrganizer\state.json
```

※イベント名は毎回初期化（保存しない）

## 使い方

1. アプリ起動（トレイ常駐）
2. SDカードを接続（自動検知）  
   または「SDカード選択」で手動選択
3. イベント名を入力
4. 「処理開始」
5. 完了ログを確認

## ビルド/公開

### ビルド

```powershell
dotnet build PhotoOrganizer.csproj -c Release
```

### 単一EXE発行（self-contained）

```powershell
dotnet publish PhotoOrganizer.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o publish\win-x64
```

出力例:

```text
publish\win-x64\PhotoOrganizer.exe
```
