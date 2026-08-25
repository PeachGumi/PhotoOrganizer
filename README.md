# Photo Organizer

> **SDカードの写真・動画を Windows / macOS へ安全に取り込み、日付とイベント単位で整理するデスクトップアプリ。**  
> 単に「コピーが終わった」だけでは SD カードを再利用可能と判定せず、保存先の実データ・永続化・ストレージ識別まで確認してから明示的に安全状態へ移行します。

[![CI](https://github.com/PeachGumi/PhotoOrganizer/actions/workflows/ci.yml/badge.svg)](https://github.com/PeachGumi/PhotoOrganizer/actions/workflows/ci.yml)
[![CodeQL](https://github.com/PeachGumi/PhotoOrganizer/actions/workflows/codeql.yml/badge.svg)](https://github.com/PeachGumi/PhotoOrganizer/actions/workflows/codeql.yml)
[![I/O Benchmark](https://github.com/PeachGumi/PhotoOrganizer/actions/workflows/io-benchmark.yml/badge.svg)](https://github.com/PeachGumi/PhotoOrganizer/actions/workflows/io-benchmark.yml)
[![.NET 10](https://img.shields.io/badge/.NET-10.0-512BD4.svg)](https://dotnet.microsoft.com/)
[![Windows / macOS](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-lightgrey.svg)](#対応環境)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Photo Organizer は、カメラの SD カードにある JPG / RAW / MOV / MP4 を指定した保存先へ取り込み、次のような構造へ整理する Windows / macOS 共通アプリです。

```text
[保存先]/[YYYY]/[YYYY-MM-DD]_[イベント名]/[RAW|JPG|MP4]/
```

設計の中心にあるのは、次の3つです。

- **SDカード上の元データを変更しない**
- **保存先にある既存データを上書きしない**
- **「コピー完了」と「SDカード再利用可能」を別の状態として扱う**

Windows 版と macOS 版を別々に実装するのではなく、取り込み・重複判定・コピー整合性・SD 再利用判定を共有 Core へ集約し、OS 固有処理だけを薄い platform adapter に分離しています。

[まず知りたいこと](#まず知りたいこと) ・ [クイックスタート](#クイックスタート) ・ [対応環境](#対応環境) ・ [対応形式](#対応形式) ・ [安全設計](#データ安全性) ・ [インストール](#インストールと配布状況) ・ [開発](#ソースから実行する) ・ [品質保証](#テストとci) ・ [FAQ](#よくある質問)

> **現在の配布状態**
>
> まだ一般利用者向けの stable release は公開していません。署名済み Prerelease → 実機受け入れ試験 → 同一成果物の stable promotion という手順を完了したものだけを正式版として公開する方針です。

---

## まず知りたいこと

### できること

| 機能 | 内容 |
|---|---|
| SDカード検知 | `DCIM` / `PRIVATE` を持つカメラメディアを検知 |
| カード全体の安全スキャン | `DCIM/100NIKON` などの下位フォルダを選んでも、安全なカード root まで広げて確認 |
| メディア分類 | JPG / JPEG、設定済み RAW、MOV / MP4 を自動分類 |
| 日付・イベント整理 | 撮影日とイベント名から保存フォルダを作成 |
| 内容ベース重複判定 | ファイル名ではなく **サイズ + SHA-256** で既存データと照合 |
| 上書き防止 | 同名・異内容は `_2`, `_3` ... として別ファイルで保存 |
| 検証付きコピー | `.partial-*` 一時ファイル、再ハッシュ、no-replace final move を使用 |
| 保存先永続化確認 | Windows / macOS の durability primitive で保存先への永続化を要求 |
| SD再利用判定 | 取り込み後に SD カードを再スキャンし、全対応メディアをもう一度検証 |
| ストレージ差し替え検知 | volume / physical device / mount session identity を追跡 |
| 常駐動作 | Windows はトレイ、macOS はメニューバーに常駐可能 |
| 複数カード | 処理中のカードを固定し、2枚目は待機キューへ保持 |
| 自動起動 | ログイン時起動とバックグラウンド起動に対応 |

### しないこと / できないこと

- **SDカード上の元ファイルを削除・移動・rename・書き換えしません**
- SDカードをアプリからフォーマット・消去する機能はありません
- XMP / XML / TXT / カメラDBなど、未対応 sidecar・補助ファイルは取り込みません
- 緑の再利用可能表示は、**指定保存先1か所への検証済みコピー**を意味します。3-2-1 バックアップや二重バックアップ完了を意味しません
- symlink / junction / reparse point 経由の保存先を独立したバックアップ先として扱いません
- 同じ物理ディスク・SDカード上の別 partition を「別の保存先」とは認めません
- OS からストレージの物理 identity を安全に取得できない構成では fail-closed で拒否する場合があります
- 故障したストレージや、flush 完了を不正に報告する hardware / firmware まで完全に保証するものではありません

### こんな使い方を想定しています

- 撮影のたびに SD カードを Mac / PC へ挿して写真を取り込みたい
- 取り込み後は次回撮影前にカメラ側で SD カードをフォーマットする
- RAW / JPG / 動画をイベントごとに自動で整理したい
- 前回分が SD カードに残っていても、新しいデータだけ安全に追加したい
- カメラが `DSC_0001.JPG` のようなファイル名を再利用しても事故らないようにしたい
- Finder / Explorer の手動コピーより厳格に「消してよいタイミング」を判断したい

---

## クイックスタート

一般利用者向けの正式配布物が公開された後は、GitHub Releases から自分の OS / CPU に合うパッケージを取得します。

基本操作は次の流れです。

1. Photo Organizer を起動する
2. 写真ライブラリの保存先を選ぶ
3. カメラの SD カードを接続する
4. 自動検知を待つ。必要なら SD カードまたはカード内フォルダを手動選択する
5. RAW / JPG / MP4 の件数と完全スキャン成功を確認する
6. イベント名を入力する
7. 保存先プレビューを確認する
8. **取り込み開始**を押す
9. コピー表示が100%になっても SD カードを抜かず、最終検証を待つ
10. 緑の **「保存先コピー検証済み — SDカード再利用可能」** を確認する
11. 必要なら保存先側の写真を別アプリでも確認する
12. SD カードをカメラへ戻し、次回撮影前にカメラ側で再利用 / フォーマットする

```text
撮影
  ↓
SDカードをPC / Macへ接続
  ↓
カード全体を完全スキャン
  ↓
保存先が別volume・別physical deviceか確認
  ↓
既存データとsize + SHA-256で重複照合
  ↓
新規メディアを検証付きコピー
  ↓
SDカード全体をもう一度完全スキャン
  ↓
保存先の実bytes + durable commitを検証
  ↓
source / destination identityを最終確認
  ↓
「SDカード再利用可能」
```

**どこか1つでも確認できなければ green にはなりません。**

---

## Photo Organizer が解決したいこと

写真を取り込むだけなら、Finder や Explorer でコピーするだけでもできます。

難しいのは、その後です。

- コピーが本当に最後まで成功したのか
- 同名ファイルが既にあった場合に上書きされていないか
- カメラがファイル番号を再利用したとき、別写真を「重複」と誤判定しないか
- SD カードに前回分が残っていても、新しい写真だけを安全に取り込めるか
- コピー後に保存先が取り外された状態で、誤って「SDカードを消してよい」と表示しないか
- OS や storage device の write cache にしか存在しないデータを「保存済み」と誤認しないか
- 同じ mount path に別の SD カードが差し替わっても、以前の承認を使い回してしまわないか

Photo Organizer は、**「いつ SD カードを再利用してよいか」までを1つの取り込みトランザクションとして扱う**ことを目的にしています。

---

## 対応環境

### 利用環境

| OS | Architecture | 想定配布形式 | 状態 |
|---|---|---|---|
| Windows | x64 | self-contained ZIP | 対応 |
| Windows | ARM64 | self-contained ZIP | 対応 |
| macOS | Apple Silicon / arm64 | signed + notarized DMG | 対応 |
| macOS | Intel / x64 | signed + notarized DMG | 対応 |
| Linux | - | - | 非対応 |

正式配布時は、Windows 成果物に Authenticode、macOS 成果物に Developer ID Application 署名・Hardened Runtime・Apple Notarization・stapling を要求します。

### 開発環境

- .NET 10 SDK
- Windows または macOS
- Avalonia 12

Linux 上で Core を扱える部分はありますが、製品としての desktop platform adapter と release acceptance の対象は Windows / macOS です。

---

## 対応形式

### RAW

初期設定では次の拡張子を RAW として扱います。

```text
.arw  .cr2  .cr3  .nef  .dng  .raf  .rw2  .orf  .pef
```

RAW 拡張子はアプリ設定から変更できます。

### JPEG

```text
.jpg  .jpeg
```

### 動画

```text
.mov  .mp4
```

`.jpg` / `.jpeg` / `.mov` / `.mp4` は標準分類として予約されており、RAW 設定へ追加しても RAW として再分類されません。

### 対象外ファイル

たとえば次のようなファイルは、取り込みと SD 再利用判定の対象外です。

```text
.xmp  .xml  .txt  カメラ固有DB  sidecar  その他の未対応形式
```

対象外ファイルがカードに存在すること自体は、対応メディアの再利用判定を妨げません。

一方、**対応形式の 0-byte ファイル**は正常な撮影データとして確認できないため、スキャンを fail-closed で停止します。

---

## 保存先の構造

保存先には、年・日付・イベント名・メディア種別の順で整理されます。

```text
Pictures/
└─ 2026/
   └─ 2026-08-26_東京スナップ/
      ├─ RAW/
      │  ├─ DSC_0001.NEF
      │  └─ DSC_0002.NEF
      ├─ JPG/
      │  ├─ DSC_0001.JPG
      │  └─ DSC_0002.JPG
      └─ MP4/
         └─ DSC_0003.MP4
```

### イベント日付の決め方

今回新しく取り込むメディアがある場合は、そのメディアを基準にイベント日付を決定します。

日付取得の優先順位は次のとおりです。

1. EXIF `DateTimeOriginal`
2. EXIF `DateTimeDigitized`
3. ファイルの最終更新日時

複数の撮影日が含まれる場合は最も古い日付をイベント日として採用し、警告を残します。

イベント名に OS 上ファイル名として使用できない文字が含まれる場合は、安全な文字へ置換します。

### 前回の写真が SD カードに残っている場合

問題ありません。

保存先全体から同じサイズ・同じ SHA-256 の実ファイルが見つかった写真は、既に保存済みとして再コピーを省略します。新しい内容だけが今回のイベントへ取り込まれます。

### カメラが同じファイル名を再利用した場合

ファイル名だけでは重複判定しません。

以前の `DSC_0001.JPG` と今回の `DSC_0001.JPG` が同名・同サイズでも、SHA-256 が異なれば別の写真として扱います。

保存先に同名の別データが存在する場合は、既存ファイルを残したまま次のように保存します。

```text
DSC_0001.JPG
DSC_0001_2.JPG
DSC_0001_3.JPG
```

---

## データ安全性

Photo Organizer では、データ安全性に関するルールを通常の実装詳細ではなく**契約**として扱っています。

厳密な仕様は [`docs/DATA_SAFETY.md`](docs/DATA_SAFETY.md) が正です。ここでは利用者が知るべき要点をまとめます。

### 安全判定の要約

| 条件 | green に必要か |
|---|---:|
| SDカード全体を完全スキャンできた | 必須 |
| 保存先が source と別 volume / 別 physical device | 必須 |
| 全対応メディアに size + SHA-256 一致の保存先コピーがある | 必須 |
| 保存先コピーを durable synchronization できた | 必須 |
| durable sync 後の fresh SHA-256 が source と一致 | 必須 |
| source / destination の mount identity が途中で変わっていない | 必須 |
| 2か所以上にバックアップ済み | 対象外 |

### 1. コピー元は immutable

SD カード上の対応メディアに対して、アプリは次の操作を行いません。

- delete
- move
- rename
- truncate
- replace
- overwrite
- metadata 書き換え

アプリが自動削除する可能性があるのは、アプリ自身が作成し、まだ最終ファイルになっていない `.partial-*` 一時ファイルだけです。

### 2. カード全体を完全に読めなければ進めない

SD カードの列挙中に I/O エラー、権限エラー、metadata 取得エラーなどが発生した場合、「見えたファイルだけ」を取り込んで安全とは判定しません。

対応メディアをカード全体から確認できたことが安全判定の前提です。

hidden directory 内の対応メディアも対象です。一方、symlink / reparse point や別 mounted volume はカード外へスキャンが逃げないよう追跡しません。

### 3. 保存先は SD カードと物理的に別でなければならない

path 文字列が違うだけでは不十分です。

Photo Organizer は source と destination について次を確認します。

- mounted volume identity
- physical storage-device identity
- process-local mount-session identity

同じ SD カードの別 partition や、同じ物理ディスク上の別 volume は独立したバックアップ先として認めません。

また、destination path に symlink / junction / reparse point などの alias が含まれる場合も fail-closed で拒否します。

### 4. 重複は timestamp やファイル名ではなく実 bytes で判定

重複判定は原則として、

```text
file size + SHA-256
```

で行います。

同名でも内容が違えば別ファイルです。別名でも内容が完全一致すれば既に保存済みのデータとして扱えます。

### 5. 新規コピーは一時ファイル経由で確定

新しいファイルは概ね次の順序で処理します。

1. source のサイズと SHA-256 を取得
2. 保存先と同じ directory に `.partial-*` を `CreateNew` で作成
3. source から一時ファイルへコピー
4. 一時ファイルを flush
5. 一時ファイルのサイズと SHA-256 を source と比較
6. source を再度読み、コピー中に内容が変化していないことを確認
7. 既存ファイルを上書きしない final move を実行
8. final metadata を設定
9. OS ごとの durable commit を実行
10. final path をもう一度サイズ・SHA-256 検証

途中で失敗しても、既存の library file を削除して帳尻を合わせることはしません。

### 6. 「SHAが一致した」だけでは green にしない

OS や storage device の cache から読み出せたデータが、必ずしも電源断後も永続媒体に残っているとは限りません。

そのため green の SD 再利用判定に **durable destination commit** を要求します。

#### Windows

- final move は no-replace
- `MOVEFILE_WRITE_THROUGH` を使用
- final metadata 設定後、final file handle を disk へ flush

#### macOS

- no-clobber final move
- parent directory entry を同期
- final file へ `F_FULLFSYNC`

durability を OS へ要求できなかった場合は、現在 SHA-256 が一致していても再利用可能とは判定しません。

### 7. durable sync 後にもう一度 SHA-256 を読む

destination を hash した後に durable sync して終わりではありません。

外部プロセスがその隙間で保存先を変更する可能性まで考慮し、**durable synchronization 後に fresh handle から destination SHA-256 を再計算**します。

これにより、古い hash と新しい bytes の組み合わせで green になる TOCTOU を防ぎます。

### 8. コピー後に SD カードをもう一度完全スキャン

コピー開始時に見えていたファイルだけを確認して終わりではありません。

取り込み後にカード全体を再スキャンし、次を確認します。

- 最初に見えていた対応メディアが消えていない
- 新たに対応メディアが増えていれば、それも最終検証対象に含める
- 現在カードに見えている対応メディア全件について独立した destination copy が存在する
- destination copy が size + SHA-256 一致する
- destination copy が durably synchronized されている
- source / destination storage identity が途中で変化していない

すべて成立した場合だけ green になります。

---

## 緑の「SDカード再利用可能」が意味するもの

緑色の状態は、Photo Organizer が対応対象としている現在の JPG / JPEG・設定済み RAW・MOV / MP4 について、**指定保存先1か所への独立コピーを実 bytes で検証し、永続媒体への同期まで要求できた**ことを意味します。

### green が保証する範囲

- カード上で現在見えている対応メディアを完全スキャン済み
- 対応メディアごとに独立した destination file を確認済み
- destination size / SHA-256 を確認済み
- durable synchronization を OS へ要求済み
- durable sync 後に fresh SHA-256 を再確認済み
- storage identity の途中差し替えを検出していない

### green が意味しないこと

- 2か所以上にバックアップ済み
- クラウドへ同期済み
- 保存先ストレージ自体が壊れない
- RAID / 3-2-1 バックアップが完成した
- 将来の bit rot まで保証される

重要な撮影データでは、green 確認後も別媒体・NAS・クラウドなどへの追加バックアップを推奨します。

---

## ストレージ差し替えへの対策

`D:\` や `/Volumes/CAMERA` といった path は、同じ文字列でも途中で別 device へ差し替わることがあります。

Photo Organizer では path だけを identity として使いません。

### Windows

volume identity に加えて、logical volume → partition → physical disk の対応を追跡します。

### macOS

`diskutil info -plist` から persistent volume identity と whole-disk identity を取得します。

`diskutil` の plist 解釈は platform adapter 内の純粋 parser として分離し、欠損・型違い・fallback 順序をテストしています。外部 command 自体も bounded execution とし、stdout / stderr を同時に drain しながら timeout します。

取得に失敗した場合は「推測で続行」せず identity unavailable として fail-closed にします。

さらに、OS の persistent identifier とは別に process-local な mount session ID を持ち、取り外し・再接続・physical mapping 変更が発生した場合は以前の scan approval を無効化します。

詳しくは [`docs/STORAGE_IDENTITY.md`](docs/STORAGE_IDENTITY.md) を参照してください。

---

## 常駐動作と複数SDカード

Photo Organizer は Windows の system tray / macOS の menu bar に常駐できます。

- idle 中にメイン window を閉じると、監視を続けたまま window を隠します
- tray / menu bar から window を再表示できます
- 明示的な Quit で終了できます
- import 中の通常終了は拒否されます
- `--background` 起動に対応します
- login auto-start と background start を個別に設定できます
- background 中でも有効な camera card を検知すると workflow window を表示できます

処理中に2枚目の camera card が接続されても、active card を勝手に変更しません。

2枚目は queue へ入り、現在の処理が終了・reset された後に扱われます。

mount event や実際の login startup は OS 境界そのものなので、unit test だけに依存せず release acceptance でも実機確認します。

---

## 設定

初期保存先は OS の **Pictures / ピクチャ** directory です。

主な設定:

- 保存先
- RAW 拡張子
- ログイン時自動起動
- バックグラウンド起動

通常設定は OS の Local Application Data 配下にある `PhotoOrganizer/settings.json`、background 設定は `PhotoOrganizer/background-settings.json` に保存されます。

安全判定そのものは設定ファイルへ永続化しません。

**アプリを再起動した時点で、以前の green 状態は引き継がれません。**

新しい process では再度 scan / import / verification が必要です。

---

## インストールと配布状況

正式配布物は GitHub の [Releases](https://github.com/PeachGumi/PhotoOrganizer/releases) から提供します。

### 正式版が公開されている場合

自分の環境に合う成果物を選びます。

| Platform | Architecture | 配布形式 |
|---|---|---|
| Windows | x64 | self-contained ZIP |
| Windows | ARM64 | self-contained ZIP |
| macOS | Apple Silicon / arm64 | signed + notarized DMG |
| macOS | Intel / x64 | signed + notarized DMG |

self-contained build のため、正式配布 ZIP / DMG の実行に .NET SDK を別途インストールすることは想定していません。

### stable release がまだない場合

この repository は、署名済み build が成功しただけでは stable release としません。

```text
main の exact commit
      ↓
署名済み Prerelease candidate
      ↓
clean machine / real SD card acceptance
      ↓
acceptance evidence 記録
      ↓
同じ artifact を stable / Latest へ promotion
```

GitHub Releases に stable release が存在しない場合、一般利用者向け production build はまだ公開前です。

詳しくは [`docs/RELEASE.md`](docs/RELEASE.md) と [`docs/RELEASE_ACCEPTANCE.md`](docs/RELEASE_ACCEPTANCE.md) を参照してください。

---

## ソースから実行する

開発には **.NET 10 SDK** が必要です。

```bash
git clone https://github.com/PeachGumi/PhotoOrganizer.git
cd PhotoOrganizer

dotnet restore PhotoOrganizer.slnx
dotnet run --project src/PhotoOrganizer.App/PhotoOrganizer.App.csproj -c Release
```

### build

```bash
dotnet build src/PhotoOrganizer.App/PhotoOrganizer.App.csproj -c Release
```

### Core tests

```bash
dotnet test tests/PhotoOrganizer.Core.Tests/PhotoOrganizer.Core.Tests.csproj -c Release
```

### platform adapter tests

```bash
dotnet test tests/PhotoOrganizer.App.Tests/PhotoOrganizer.App.Tests.csproj -c Release
```

### I/O benchmark

```bash
dotnet run --project tools/PhotoOrganizer.IoBenchmark/PhotoOrganizer.IoBenchmark.csproj -c Release
```

ローカル build は production 署名・notarization 済み配布物とは別物です。第三者へ正式配布する場合は release workflow を使用してください。

---

## アーキテクチャ

Photo Organizer は、OS によって安全ロジックが分岐しないように設計しています。

```text
┌─────────────────────────────────────────────┐
│ PhotoOrganizer.App                          │
│ Avalonia UI / tray / menu bar / settings    │
│ startup / storage platform adapters          │
└──────────────────────┬──────────────────────┘
                       │
                       v
┌─────────────────────────────────────────────┐
│ PhotoOrganizer.Core                         │
│                                             │
│ - media classification                      │
│ - complete scan                             │
│ - duplicate detection                       │
│ - safe copy transaction                     │
│ - durable destination verification          │
│ - SD reuse safety state machine             │
└──────────────────────┬──────────────────────┘
                       │
                       v
┌─────────────────────────────────────────────┐
│ OS / filesystem                             │
│                                             │
│ Windows: volume / physical disk / registry  │
│ macOS: diskutil / whole disk / LaunchAgent  │
└─────────────────────────────────────────────┘
```

Core は WinForms、AppKit、SwiftUI、Windows Management API、Avalonia UI type へ依存しません。

UI / platform layer から Core へ依存し、Core から OS 固有 layer へ逆依存しない方向を維持します。

詳細は [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) を参照してください。

---

## Repository構成

```text
PhotoOrganizer/
├─ src/
│  ├─ PhotoOrganizer.Core/        # 共有の安全・取り込みロジック
│  └─ PhotoOrganizer.App/         # Avalonia desktop app + platform adapters
├─ tests/
│  ├─ PhotoOrganizer.Core.Tests/  # cross-platform safety regression
│  └─ PhotoOrganizer.App.Tests/   # OS adapter / diskutil / startup / identity tests
├─ tools/
│  └─ PhotoOrganizer.IoBenchmark/ # SHA-256 / destination lookup I/O benchmark
├─ docs/
│  ├─ ARCHITECTURE.md
│  ├─ DATA_SAFETY.md
│  ├─ MIGRATION_PLAN.md
│  ├─ RELEASE.md
│  ├─ RELEASE_ACCEPTANCE.md
│  └─ STORAGE_IDENTITY.md
├─ Scripts/                       # repository / release configuration helpers
└─ .github/workflows/             # CI / CodeQL / benchmark / release
```

---

## テストとCI

Photo Organizer は、**Core の安全契約・platform adapter・実OS・package・release contract・実機 acceptance** を別々の層で検証します。

### テストの役割分担

| 層 | 主な対象 |
|---|---|
| `PhotoOrganizer.Core.Tests` | complete scan、source immutability、重複判定、copy transaction、format safety、mount session、race / failure handling |
| `PhotoOrganizer.App.Tests` | Windows / macOS storage identity、`diskutil` plist parser、bounded process execution、startup registration format、実OS volume resolution |
| Package smoke | Windows x64/ARM64 ZIP、macOS arm64/x64 DMG の生成・展開 / mount |
| CodeQL | C# static analysis |
| I/O Benchmark | SHA-256 / destination lookup の読み取り量回帰 |
| Release acceptance | 実SDカード、実ストレージ、unplug、logout/login、Gatekeeper / Authenticode、durability interruption |

platform adapter tests は、OS API をすべて mock 化する方針ではありません。

たとえば macOS storage identity は、

```text
diskutil subprocess
        ↓
BoundedProcessRunner
        ↓
raw plist
        ↓
MacDiskutilInfoParser
        ↓
volume / physical identity
```

のように、**OS コマンド実行部分と純粋な解釈ロジックを分離**しています。

これにより、実 runner 上で `diskutil` が動くことを integration test しつつ、plist の欠損・型違い・fallback 順序・malformed XML は高速な unit test で網羅できます。

同様に startup registration も、実際の Registry / LaunchAgent mutation と、書き込む command / plist の生成を分けてテストします。

### 通常CIで確認するもの

- shared Core safety tests
- source immutability regression
- same-name / different-bytes collision handling
- SHA-256 duplicate detection
- symlink / junction alias rejection
- nested mounted-volume traversal rejection
- volume / physical-device / mount-session replacement detection
- durability failure時の fail-closed
- durable sync 中に destination が変化した race の再hash検証
- Windows / macOS platform storage identity tests
- `diskutil` plist parser の fallback / malformed input
- subprocess timeout / stdout・stderr pipe handling
- startup registration command / LaunchAgent plist generation
- unified desktop app build
- Windows x64 / ARM64 package smoke
- macOS arm64 / x64 DMG smoke
- release workflow contract validation
- CodeQL static analysis
- SHA-256 I/O benchmark

workflows:

```text
.github/workflows/ci.yml
.github/workflows/codeql.yml
.github/workflows/io-benchmark.yml
.github/workflows/release.yml
.github/workflows/promote-release.yml
```

安全性に関する bug は、可能な限り再現 test を追加してから修正します。

---

## Release設計

Windows と macOS は **同じ version・同じ source commit** から release します。

production release は fail-closed です。

- signing credential が不足 → publish しない
- Windows だけ成功 → publish しない
- macOS だけ成功 → publish しない
- checksum 不一致 → publish しない
- signing / notarization 失敗 → publish しない
- real-device acceptance 未完了 → stable へ promotion しない

release authority も分離されています。

- `production-signing` — signing / notarization credential を保持
- `production-release` — acceptance 済み candidate を stable へ promotion する権限

署名済み candidate 作成と stable promotion を別 workflow へ分けることで、CI 成功だけで一般配布が進まない構成です。

---

## セキュリティと脆弱性報告

次の問題は特に重大な release blocker として扱います。

- source media の削除・変更
- destination 既存ファイルの上書き
- false duplicate detection
- false SD-reuse approval
- selected storage 外への path traversal
- signing bypass
- arbitrary code execution

データ消失の再現詳細、秘密鍵、certificate、password、token、実際のユーザー写真などを public Issue へ投稿しないでください。

報告方法は [`SECURITY.md`](SECURITY.md) を参照してください。

---

## 制限事項と注意点

### 1つの保存先はバックアップ戦略そのものではない

Photo Organizer は選択した保存先コピーの整合性を厳格に確認しますが、保存先 device 自体が壊れればデータを失う可能性があります。

重要データでは別媒体・NAS・cloud 等を組み合わせた追加バックアップを用意してください。

### ネットワーク / 仮想 / 特殊filesystem

このアプリは source と destination の物理的な独立性を安全判定へ利用します。

そのため、physical-device identity を OS から確立できない network share、virtual filesystem、特殊 mount などでは fail-closed となる場合があります。

「書き込めるから安全な保存先として使える」とは限りません。

### storage firmware の保証まではできない

Photo Organizer は Windows / macOS が提供する durability primitive を使用しますが、device firmware が flush 要求に対して不正な成功を返すような hardware failure まで検出できるわけではありません。

production release では、使い捨て destination を用いた abrupt disconnect → remount → independent rehash の real-device acceptance も要求しています。

---

## よくある質問

### Q. 緑になったら SD カードをフォーマットしていい？

Photo Organizer の対応対象ファイルについては、選択保存先への実 bytes 一致と durable synchronization まで確認済みという判定です。

ただし green は二重バックアップの意味ではありません。重要撮影では保存先側の追加バックアップも推奨します。

### Q. コピーが100%になったら SD カードを抜いていい？

いいえ。**green の「SDカード再利用可能」が表示されるまで待ってください。**

コピー後にも complete rescan、fresh SHA-256、durable synchronization、storage identity の最終確認があります。

### Q. Photo Organizer が SD カードを自動で消すことはある？

ありません。アプリに format / erase 機能はありません。

### Q. XMP sidecar も保存される？

現在は保存されません。XMP / XML / TXT などは取り込み・SD 再利用判定の対象外です。

### Q. SD カードに前回の写真を残したまま再度取り込める？

できます。保存先全体から size + SHA-256 一致を確認できた既存データは skip し、新しい bytes だけを取り込みます。

### Q. 同名写真があると上書きされる？

されません。同じ bytes なら duplicate、違う bytes なら `_2`, `_3` ... の別名で保持します。

### Q. 同じ内容の写真が別名で存在したら？

size + SHA-256 が一致する完全同一 bytes であれば、既に保存済みの内容として扱います。ファイル名は重複判定の主キーではありません。

### Q. 保存先を同じ SD カードの別 partition にしてもいい？

できません。同じ physical storage device は独立した backup location とみなさず、取り込み前に拒否します。

### Q. NAS を保存先にできる？

physical-device identity を安全に確立できない network share は fail-closed になる可能性があります。現在の安全契約は、source と destination の物理的独立性を OS から確認できる保存先を前提にしています。

### Q. アプリを再起動したら前回の green は残る？

残りません。安全状態は永続化せず、新しい process では再度検証が必要です。

### Q. 2枚の SD カードを同時に挿したら？

処理中の active card は切り替えません。2枚目は待機 queue へ入り、現在の処理が終わった後に扱います。

### Q. Windows版とmacOS版で安全判定は違う？

基本的な安全契約は共通 Core にあります。OS ごとに異なるのは、volume / physical-device identity、durability primitive、startup / resident integration などの platform adapter 部分です。

---

## 移行について

この repository は、旧 Windows 版・macOS 版を統合した canonical implementation です。

- [`PeachGumi/PhotoOrganizer-win`](https://github.com/PeachGumi/PhotoOrganizer-win) — original Windows implementation
- [`PeachGumi/PhotoOrganizer-mac`](https://github.com/PeachGumi/PhotoOrganizer-mac) — hardened macOS safety reference

旧実装で得た platform 固有の知見を引き継ぎつつ、現在は shared Core を安全仕様の source of truth としています。

旧 repository は、統合版の署名済み candidate が real-device acceptance を完了し stable へ promotion されるまで reference として保持します。

意図的な behavior 差分と retirement 条件は [`docs/MIGRATION_PLAN.md`](docs/MIGRATION_PLAN.md) を参照してください。

---

## 関連ドキュメント

| Document | 内容 |
|---|---|
| [`docs/DATA_SAFETY.md`](docs/DATA_SAFETY.md) | データ安全性の normative contract |
| [`docs/STORAGE_IDENTITY.md`](docs/STORAGE_IDENTITY.md) | volume / physical device / mount session identity |
| [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) | shared Core と platform layer の設計 |
| [`docs/MIGRATION_PLAN.md`](docs/MIGRATION_PLAN.md) | legacy 版から統合版への移行方針 |
| [`docs/RELEASE.md`](docs/RELEASE.md) | signing / notarization / prerelease / stable promotion |
| [`docs/RELEASE_ACCEPTANCE.md`](docs/RELEASE_ACCEPTANCE.md) | clean-machine / real-SD / unplug / login / signing acceptance |
| [`SECURITY.md`](SECURITY.md) | vulnerability reporting / security policy |

---

## License

[MIT License](LICENSE)
