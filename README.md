# Photo Organizer

> **撮影データを、SDカードからWindows / macOSへ安全に取り込み、日付とイベント単位で整理するデスクトップアプリ。**  
> コピーが終わっただけでは「SDカードを再利用してよい」と判定せず、保存先の実データ検証と永続化まで確認してから明示的に再利用可能状態へ移行します。

[![CI](https://github.com/PeachGumi/PhotoOrganizer/actions/workflows/ci.yml/badge.svg)](https://github.com/PeachGumi/PhotoOrganizer/actions/workflows/ci.yml)
[![CodeQL](https://github.com/PeachGumi/PhotoOrganizer/actions/workflows/codeql.yml/badge.svg)](https://github.com/PeachGumi/PhotoOrganizer/actions/workflows/codeql.yml)
[![I/O Benchmark](https://github.com/PeachGumi/PhotoOrganizer/actions/workflows/io-benchmark.yml/badge.svg)](https://github.com/PeachGumi/PhotoOrganizer/actions/workflows/io-benchmark.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Photo Organizerは、カメラのSDカードにある写真・動画を指定した保存先へ取り込み、次の形式で整理するWindows / macOS共通アプリです。

```text
[保存先]/[YYYY]/[YYYY-MM-DD]_[イベント名]/[RAW|JPG|MP4]/
```

単なるファイルコピーではなく、**「元データを変更しない」「既存データを上書きしない」「コピー済みとSDカード再利用可能を分ける」**ことを設計の中心に置いています。

Windows版とmacOS版を別々に保守するのではなく、取り込み・重複判定・コピー整合性・SD再利用判定を共有Coreへ集約し、OS固有処理だけを薄いplatform adapterへ分離しています。

---

## まず知りたいこと

### できること

- **SDカードを自動検知**し、`DCIM` / `PRIVATE` を持つカメラメディアを認識
- 手動で `DCIM/100NIKON` などの下位フォルダを選んでも、**カード全体の安全なrootまで広げてスキャン**
- JPG / JPEG、設定済みRAW、MOV / MP4を自動分類
- 撮影日とイベント名から保存フォルダを自動作成
- 保存先全体を対象に、**ファイルサイズ + SHA-256**で内容ベースの重複判定
- 同名・異内容ファイルを `_2`, `_3` ... で保持し、既存ファイルを上書きしない
- `.partial-*` 一時ファイルを使った検証付きコピー
- コピー後にSDカードを再スキャンし、保存先の実ファイルを再検証
- Windows / macOSそれぞれで保存先への**durable commit（永続媒体への同期）**を要求
- コピー元・保存先のvolume / physical device / mount session identityを追跡
- SDカード差し替え、保存先取り外し、コピー失敗、検証失敗などで**fail-closed**
- 2枚目のSDカードが挿入されても、処理中のカードを勝手に切り替えず待機キューへ保持
- Windowsではトレイ、macOSではメニューバーに常駐
- ログイン時自動起動とバックグラウンド起動

### しないこと / できないこと

- **SDカード上の元ファイルを削除・移動・rename・書き換えしません**
- SDカードをアプリからフォーマット・消去する機能はありません
- XMP / XML / TXT / カメラDBなど、対応範囲外のsidecar・補助ファイルは取り込みません
- 緑の再利用可能表示は、**指定保存先1か所への検証済みコピー**を意味します。3-2-1バックアップや二重バックアップ完了を意味しません
- symlink / junction / reparse point経由の保存先を独立したバックアップ先として扱いません
- 同じ物理ディスク・SDカード上の別partitionを「別の保存先」とは認めません
- OSから物理ストレージidentityを安全に取得できない保存先では、取り込み・再利用判定を拒否する場合があります
- 故障したストレージや、flush完了を正しく報告しないハードウェア/firmwareまで完全に保証するものではありません

---

## Photo Organizerが解決したいこと

写真を取り込むだけなら、FinderやExplorerでコピーするだけでもできます。

問題になるのは、その後です。

- コピーが本当に最後まで成功したのか
- 同名ファイルが既にあった場合に上書きされていないか
- カメラがファイル番号を再利用したとき、別写真を「重複」と誤判定しないか
- SDカードに前回分が残っていても、新しい写真だけを安全に取り込めるか
- コピー後に保存先が破損・取り外しされた状態で、誤って「SDカードを消してよい」と表示しないか
- OSやストレージのwrite cacheにしか存在しないデータを「保存済み」と誤認しないか

Photo Organizerは、この最後の **「いつSDカードを再利用してよいか」** までをひとつの取り込みトランザクションとして扱います。

```text
カメラで撮影
    ↓
SDカードをPC / Macへ接続
    ↓
カード全体を完全スキャン
    ↓
保存先が別volume・別physical deviceか確認
    ↓
既存の実データとSHA-256で重複照合
    ↓
新規メディアを検証付きでコピー
    ↓
SDカード全体をもう一度スキャン
    ↓
保存先のsize + SHA-256 + durable commitを確認
    ↓
source / destination identityを最終確認
    ↓
「保存先コピー検証済み — SDカード再利用可能」
```

どこか1つでも証明できなければ、再利用可能状態にはなりません。

---

## 対応形式

### RAW

初期設定では次の拡張子をRAWとして扱います。

```text
.arw  .cr2  .cr3  .nef  .dng  .raf  .rw2  .orf  .pef
```

RAW拡張子はアプリ設定から変更できます。

### JPEG

```text
.jpg  .jpeg
```

### 動画

```text
.mov  .mp4
```

`.jpg` / `.jpeg` / `.mov` / `.mp4` は標準分類として予約されており、RAW設定へ追加してもRAWとして再分類されません。

### 対象外ファイル

たとえば次のようなファイルは、取り込みとSD再利用判定の対象外です。

```text
.xmp  .xml  .txt  カメラ固有DB  sidecar  その他の未対応形式
```

対象外ファイルがカードに存在すること自体は、対応メディアの再利用判定を妨げません。

一方、**対応形式の0-byteファイル**は正常な撮影データとして確認できないため、スキャンをfail-closedで停止します。

---

## 保存先の構造

保存先には年・日付・イベント名・メディア種別の順で整理されます。

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

イベント日付は、今回新しく取り込むメディアがある場合はその対象を基準に決定します。

日付取得は次の優先順位です。

1. EXIF `DateTimeOriginal`
2. EXIF `DateTimeDigitized`
3. ファイルの最終更新日時

複数の撮影日が含まれる場合は最も古い日付をイベント日として採用し、警告を残します。

イベント名にOS上ファイル名として使用できない文字が含まれる場合は、安全な文字へ置換します。

---

## 使い方

基本操作は次の流れです。

1. Photo Organizerを起動
2. 保存先を選択
3. カメラのSDカードを接続
4. 自動検知を待つか、必要なら手動でSDカード / カード内フォルダを選択
5. RAW / JPG / MP4の件数と、完全スキャン成功を確認
6. イベント名を入力
7. 保存先previewを確認
8. **取り込み開始**を押す
9. コピーが終わってもSDカードを抜かず、再スキャンと最終検証を待つ
10. 緑色の **「保存先コピー検証済み — SDカード再利用可能」** を確認
11. 必要に応じて保存先の写真を画像ビューアやRAW現像ソフトで確認
12. SDカードをカメラへ戻し、次回撮影時にカメラ側で再利用 / フォーマット

### 前回の写真がSDカードに残っている場合

問題ありません。

保存先全体から同じサイズ・同じSHA-256の実ファイルが見つかった写真は、既に保存済みとして再コピーを省略します。

新しい内容だけが今回のイベントへ取り込まれます。

### カメラが同じファイル名を再利用した場合

ファイル名だけでは重複判定しません。

たとえば、以前の `DSC_0001.JPG` と今回の `DSC_0001.JPG` が同名・同サイズでも、SHA-256が異なれば別の写真として扱います。

保存先に同名の別データが存在する場合は、既存ファイルを残したまま次のように保存します。

```text
DSC_0001.JPG
DSC_0001_2.JPG
DSC_0001_3.JPG
```

---

## データ安全性

Photo Organizerでは、データ安全性に関するルールを通常の実装詳細ではなく**契約**として扱っています。

厳密な仕様は [`docs/DATA_SAFETY.md`](docs/DATA_SAFETY.md) が正です。READMEでは重要な考え方を要約します。

### 1. コピー元はimmutable

SDカード上の対応メディアに対して、アプリは次の操作を行いません。

- delete
- move
- rename
- truncate
- replace
- overwrite
- metadata書き換え

アプリが自動削除する可能性があるのは、アプリ自身が作成し、まだ最終ファイルになっていない `.partial-*` 一時ファイルだけです。

### 2. カード全体を完全に読めなければ進めない

SDカードの列挙中にI/Oエラー、権限エラー、metadata取得エラーなどが発生した場合、「見えたファイルだけ」を取り込んで安全とは判定しません。

対応メディアをカード全体から確認できたことが安全判定の前提です。

hidden directory内の対応メディアも対象です。一方、symlink / reparse pointや別mounted volumeはカード外へスキャンが逃げないよう追跡しません。

### 3. 保存先はSDカードと物理的に別でなければならない

path文字列が違うだけでは不十分です。

Photo Organizerは、sourceとdestinationについて次を確認します。

- mounted volume identity
- physical storage-device identity
- process-local mount-session identity

同じSDカードの別partitionや、同じ物理ディスク上の別volumeは独立したバックアップ先として認めません。

また、destination pathにsymlink / junction / reparse pointなどのaliasが含まれる場合もfail-closedで拒否します。

### 4. 重複はtimestampやファイル名ではなく実bytesで判定

既存保存データとの重複判定は、基本的に次の組み合わせです。

```text
file size + SHA-256
```

カメラ側のファイル番号再利用やtimestampの一致だけで、別写真を捨てないためです。

### 5. 新規コピーは一時ファイル経由で確定

新しいファイルは、概ね次の順序で処理します。

1. sourceのサイズとSHA-256を取得
2. 保存先と同じdirectoryに `.partial-*` を `CreateNew` で作成
3. sourceから一時ファイルへコピー
4. 一時ファイルをflush
5. 一時ファイルのサイズとSHA-256をsourceと比較
6. sourceを再度読み、コピー中に内容が変化していないことを確認
7. 既存ファイルを上書きしないfinal moveを実行
8. final metadataを設定
9. OSごとのdurable commitを実行
10. final pathをもう一度サイズ・SHA-256検証

途中で失敗しても、既存のlibrary fileを削除して帳尻を合わせることはしません。

### 6. 「SHAが一致した」だけではgreenにしない

OSやストレージdeviceのcacheから読み出せたデータが、必ずしも電源断後も永続媒体に残っているとは限りません。

そのためPhoto Organizerは、greenのSD再利用判定に**durable destination commit**を要求します。

#### Windows

- final moveはno-replaceで実行
- `MOVEFILE_WRITE_THROUGH` を使用
- final metadata設定後、final file handleをdiskへflush

#### macOS

- no-clobber final move
- parent directory entryを同期
- final fileへ `F_FULLFSYNC`

durabilityをOSへ要求できなかった場合は、現在SHA-256が一致していても再利用可能とは判定しません。

### 7. durable sync後にもう一度SHA-256を読む

最終再利用判定では、destinationをhashした後にdurable syncして終わりではありません。

外部プロセスがその隙間で保存先を変更する可能性まで考慮し、**durable synchronization後にfresh handleからdestination SHA-256を再計算**します。

古いhashと新しいbytesの組み合わせでgreenになることを防ぎます。

### 8. コピー後にSDカードをもう一度完全スキャン

コピー開始時に見えていたファイルだけを確認して終わりではありません。

取り込み後にカード全体を再スキャンし、次を確認します。

- 最初に見えていた対応メディアが消えていない
- 新たに対応メディアが増えていれば、それも最終検証対象に含める
- 現在カードに見えている対応メディア全件について、独立したdestination copyが存在する
- destination copyがsize + SHA-256一致する
- destination copyがdurably synchronizedされている
- source / destination storage identityが途中で変化していない

すべて成立した場合だけgreenになります。

---

## 緑の「SDカード再利用可能」が意味するもの

緑色の状態は、Photo Organizerが対応対象としている現在のJPG / JPEG・設定済みRAW・MOV / MP4について、**指定保存先1か所への独立コピーを実bytesで検証し、永続媒体への同期まで要求できた**ことを意味します。

意味しないもの:

- 2か所以上にバックアップ済み
- クラウドへ同期済み
- ストレージ自体の故障耐性がある
- RAID / 3-2-1バックアップが完成した
- 将来のbit rotまで保証される

重要な撮影データでは、green確認後も別媒体やクラウドなどへの追加バックアップを推奨します。

---

## ストレージ差し替えへの対策

`D:\` や `/Volumes/CAMERA` といったpathは、同じ文字列でも途中で別deviceへ差し替わることがあります。

Photo Organizerではpathだけをidentityとして使いません。

### Windows

volume identityに加えて、logical volume → partition → physical diskの対応を追跡します。

### macOS

`diskutil info -plist` からvolume / partition identityとwhole-disk identityを取得します。

外部commandはbounded executionとし、取得に失敗した場合は「推測で続行」せずidentity unavailableとしてfail-closedにします。

さらに、OSのpersistent identifierとは別にprocess-localなmount session IDを持ち、取り外し・再接続・physical mapping変更が発生した場合は以前のscan approvalを無効化します。

詳しくは [`docs/STORAGE_IDENTITY.md`](docs/STORAGE_IDENTITY.md) を参照してください。

---

## 常駐動作と複数SDカード

Photo OrganizerはWindowsのsystem tray / macOSのmenu barに常駐できます。

- idle中にメインwindowを閉じると、通常は監視を続けたままwindowを隠します
- tray / menu barからwindowを再表示できます
- 明示的なQuitで終了できます
- import中の通常終了は拒否されます
- `--background` 起動に対応します
- login auto-startとbackground startを個別に設定できます
- background中でも有効なcamera cardを検知するとworkflow windowを表示できます

処理中に2枚目のcamera cardが接続されても、active cardを勝手に変更しません。

2枚目はqueueへ入り、現在の処理が終了・resetされた後に扱われます。

---

## 設定

初期保存先はOSの **Pictures / ピクチャ** directoryです。

主な設定:

- 保存先
- RAW拡張子
- ログイン時自動起動
- バックグラウンド起動

通常設定はOSのLocal Application Data配下にある `PhotoOrganizer/settings.json`、background設定は `PhotoOrganizer/background-settings.json` に保存されます。

安全判定そのものは設定ファイルへ永続化しません。

**アプリを再起動した時点で、以前のgreen状態は引き継がれません。**

新しいprocessでは再度scan / import / verificationが必要です。

---

## インストール / 配布状況

正式配布物はGitHubの [Releases](https://github.com/PeachGumi/PhotoOrganizer/releases) から提供する設計です。

production releaseでは同一commitから次の4artifactを生成します。

| Platform | Architecture | 配布形式 |
|---|---|---|
| Windows | x64 | self-contained ZIP |
| Windows | ARM64 | self-contained ZIP |
| macOS | Apple Silicon / arm64 | signed + notarized DMG |
| macOS | Intel / x64 | signed + notarized DMG |

Windows artifactはAuthenticode署名、macOS artifactはDeveloper ID Application署名・Hardened Runtime・Apple Notarization・stapling・Gatekeeper検証をrelease条件とします。

### 安定版がまだ公開されていない場合

このrepositoryは、署名済みbuildが成功しただけではstable releaseとしません。

1. signed Prerelease candidateを生成
2. clean machine / real SD cardで [`docs/RELEASE_ACCEPTANCE.md`](docs/RELEASE_ACCEPTANCE.md) を実施
3. exact candidateのevidenceを記録
4. 別workflowで同じartifactをstable / Latestへpromotion

という2段階releaseを採用しています。

GitHub Releasesにstable releaseが存在しない場合、一般利用者向けproduction buildはまだ公開前です。

詳しくは [`docs/RELEASE.md`](docs/RELEASE.md) を参照してください。

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

### Core test

```bash
dotnet test tests/PhotoOrganizer.Core.Tests/PhotoOrganizer.Core.Tests.csproj -c Release
```

### platform adapter test

```bash
dotnet test tests/PhotoOrganizer.App.Tests/PhotoOrganizer.App.Tests.csproj -c Release
```

ローカルbuildはproduction署名・notarization済み配布物とは別物です。第三者へ正式配布する場合はrelease workflowを使用してください。

---

## アーキテクチャ

Photo Organizerは、OSによって安全ロジックが分岐しないように設計しています。

```text
┌─────────────────────────────────────────────┐
│ PhotoOrganizer.App                          │
│ Avalonia UI / tray / menu bar / settings    │
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
│ - SD reuse safety state machine              │
└──────────────────────┬──────────────────────┘
                       │
                       v
┌─────────────────────────────────────────────┐
│ Platform adapters                           │
│                                             │
│ Windows: volume / physical disk / startup   │
│ macOS: diskutil / whole disk / login item   │
│ packaging / signing / lifecycle             │
└─────────────────────────────────────────────┘
```

CoreはWinForms、AppKit、SwiftUI、Windows Management API、Avalonia UI typeへ依存しません。

UI / platform layerからCoreへ依存し、CoreからOS固有layerへ逆依存しない方向を維持します。

詳細は [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) を参照してください。

---

## Repository構成

```text
PhotoOrganizer/
├─ src/
│  ├─ PhotoOrganizer.Core/       # 共有の安全・取り込みロジック
│  └─ PhotoOrganizer.App/        # Avalonia desktop application
├─ tests/
│  ├─ PhotoOrganizer.Core.Tests/ # cross-platform safety regression
│  └─ PhotoOrganizer.App.Tests/  # platform adapter / storage identity tests
├─ tools/
│  └─ PhotoOrganizer.IoBenchmark/ # SHA-256 / destination lookup I/O benchmark
├─ docs/
│  ├─ ARCHITECTURE.md
│  ├─ DATA_SAFETY.md
│  ├─ MIGRATION_PLAN.md
│  ├─ RELEASE.md
│  ├─ RELEASE_ACCEPTANCE.md
│  └─ STORAGE_IDENTITY.md
├─ Scripts/                      # repository / release configuration helpers
└─ .github/workflows/            # CI / CodeQL / benchmark / release
```

---

## テストとCI

通常CIでは、WindowsとmacOSの両方で安全ロジックとdesktop applicationを検証します。

主な検証内容:

- shared Core safety tests
- source immutability regression
- same-name / different-bytes collision handling
- SHA-256 duplicate detection
- symlink / junction alias rejection
- nested mounted-volume traversal rejection
- volume / physical-device / mount-session replacement detection
- durability failure時のfail-closed
- durable sync中にdestinationが変化したraceの再hash検証
- Windows / macOS platform storage identity tests
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

安全性に関するbugは、可能な限り再現testを追加してから修正します。

---

## Release設計

WindowsとmacOSは**同じversion・同じsource commit**からreleaseします。

production releaseはfail-closedです。

- signing credentialが不足 → publishしない
- Windowsだけ成功 → publishしない
- macOSだけ成功 → publishしない
- checksum不一致 → publishしない
- signing / notarization失敗 → publishしない
- real-device acceptance未完了 → stableへpromotionしない

release authorityも分離されています。

- `production-signing` — signing / notarization credentialを保持
- `production-release` — acceptance済みcandidateをstableへpromotionする権限

署名済みcandidate作成とstable promotionを別workflowへ分けることで、CI成功だけで一般配布が進まない構成です。

---

## セキュリティと脆弱性報告

次の問題は特に重大なrelease blockerとして扱います。

- source mediaの削除・変更
- destination既存ファイルの上書き
- false duplicate detection
- false SD-reuse approval
- selected storage外へのpath traversal
- signing bypass
- arbitrary code execution

データ消失の再現詳細、秘密鍵、certificate、password、token、実際のユーザー写真などをpublic Issueへ投稿しないでください。

報告方法は [`SECURITY.md`](SECURITY.md) を参照してください。

---

## 制限事項と注意点

### 1つの保存先はバックアップ戦略そのものではない

Photo Organizerは、選択した保存先コピーの整合性を厳格に確認しますが、保存先device自体が壊れればデータを失う可能性があります。

重要データでは別媒体・NAS・cloud等を組み合わせた追加バックアップを用意してください。

### ネットワーク / 仮想 / 特殊filesystem

このアプリはsourceとdestinationの物理的な独立性を安全判定へ利用します。

そのため、physical-device identityをOSから確立できないnetwork share、virtual filesystem、特殊mountなどではfail-closedとなる場合があります。

「書き込めるから安全な保存先として使える」とは限りません。

### storage firmwareの保証まではできない

Photo OrganizerはWindows / macOSが提供するdurability primitiveを使用しますが、device firmwareがflush要求に対して不正な成功を返すようなhardware failureまで検出できるわけではありません。

production releaseでは、使い捨てdestinationを用いたabrupt disconnect → remount → independent rehashのreal-device acceptanceも要求しています。

---

## よくある質問

### Q. 緑になったらSDカードをフォーマットしていい？

Photo Organizerの対応対象ファイルについては、選択保存先への実bytes一致とdurable synchronizationまで確認済みという判定です。

ただし、greenは二重バックアップの意味ではありません。重要撮影では保存先側の追加バックアップも推奨します。

### Q. Photo OrganizerがSDカードを自動で消すことはある？

ありません。アプリにformat / erase機能はありません。

### Q. XMP sidecarも保存される？

現在は保存されません。XMP / XML / TXTなどは取り込み・SD再利用判定の対象外です。

### Q. SDカードに前回の写真を残したまま再度取り込める？

できます。保存先全体からsize + SHA-256一致を確認できた既存データはskipし、新しいbytesだけを取り込みます。

### Q. 同名写真があると上書きされる？

されません。同じbytesならduplicate、違うbytesなら `_2`, `_3` ... の別名で保持します。

### Q. 保存先を同じSDカードの別partitionにしてもいい？

できません。同じphysical storage deviceは独立したbackup locationとみなさず、取り込み前に拒否します。

### Q. コピーが100%になったらSDカードを抜いていい？

いいえ。**greenの「SDカード再利用可能」が表示されるまで待ってください。**

コピー後にもcomplete rescan、fresh SHA-256、durable synchronization、storage identityの最終確認があります。

---

## 移行について

このrepositoryは、旧Windows版・macOS版を統合した新しいcanonical implementationです。

- [`PeachGumi/PhotoOrganizer-win`](https://github.com/PeachGumi/PhotoOrganizer-win) — original Windows implementation
- [`PeachGumi/PhotoOrganizer-mac`](https://github.com/PeachGumi/PhotoOrganizer-mac) — hardened macOS safety reference

旧実装で得たplatform固有の知見を引き継ぎつつ、現在はshared Coreを安全仕様のsource of truthとしています。

旧repositoryは、統合版の署名済みcandidateがreal-device acceptanceを完了しstableへpromotionされるまでreferenceとして保持します。

意図的なbehavior差分とretirement条件は [`docs/MIGRATION_PLAN.md`](docs/MIGRATION_PLAN.md) を参照してください。

---

## 関連ドキュメント

| Document | 内容 |
|---|---|
| [`docs/DATA_SAFETY.md`](docs/DATA_SAFETY.md) | データ安全性のnormative contract |
| [`docs/STORAGE_IDENTITY.md`](docs/STORAGE_IDENTITY.md) | volume / physical device / mount session identity |
| [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) | shared Coreとplatform layerの設計 |
| [`docs/MIGRATION_PLAN.md`](docs/MIGRATION_PLAN.md) | legacy版から統合版への移行方針 |
| [`docs/RELEASE.md`](docs/RELEASE.md) | signing / notarization / prerelease / stable promotion |
| [`docs/RELEASE_ACCEPTANCE.md`](docs/RELEASE_ACCEPTANCE.md) | clean-machine / real-SD release acceptance |
| [`SECURITY.md`](SECURITY.md) | vulnerability reporting / security policy |

---

## License

[MIT License](LICENSE)
