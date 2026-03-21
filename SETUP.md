# ふくむすめキャリア LP 公開前セットアップガイド

## ファイル構成

```
fukumusume-career/
├── index.html        … メインLP
├── privacy.html      … プライバシーポリシー
├── tokushoho.html    … 特定商取引法に基づく表記
└── SETUP.md          … このファイル
```

## 公開前に必要な作業（チェックリスト）

### 1. Googleカレンダー予約スケジュールを作成する

1. PC で [Googleカレンダー](https://calendar.google.com) を開く
2. 左上の「＋ 作成」→「予約スケジュール」をクリック
3. 以下を設定：
   - タイトル：「ふくむすめキャリア 無料相談」
   - 所要時間：30分
   - 対応可能時間：土曜10:00〜18:00、平日夜20:00〜21:30 など
   - 予約の間隔：30分バッファ
   - 予約受付期間：1〜30日先
   - 確認メール：ON
4. 保存して **予約ページのURL** を取得

### 2. LINE公式アカウントを作成する

1. [LINE Official Account Manager](https://manager.line.biz/) からアカウント作成
2. アカウント名：「ふくむすめキャリア」
3. プロフィール画像を設定
4. 友だち追加URLを取得（`https://lin.ee/xxxxx` 形式）

### 3. LP内のプレースホルダーを差し替える

以下の文字列を検索して、実際のURLに差し替えてください。

| 検索文字列 | 差し替え先 | 該当ファイル |
|---|---|---|
| `BOOKING_URL` | Googleカレンダー予約URL | index.html（3か所） |
| `LINE_URL` | LINE公式 友だち追加URL | index.html（1か所） |
| `EMAIL_ADDRESS` | 連絡先メールアドレス | tokushoho.html（1か所） |

### 4. 特商法ページの確認

`tokushoho.html` 内のコメント `★ 公開前に記入 ★` の箇所を確認してください。
個人事業主の場合、住所・電話番号は「請求時開示」で法的にOKです（現状そう記載済み）。

### 5. OG画像の準備（推奨）

XやLINEで共有した際のプレビュー画像を作成し、以下のメタタグに設定：
```html
<meta property="og:image" content="画像のURL">
```
推奨サイズ：1200x630px

### 6. 公開方法

**GitHub Pagesの場合：**
1. 新しいリポジトリ `fukumusume-career` を作成
2. このフォルダの中身をpush
3. Settings → Pages → Deploy from branch → main
4. 公開URL：`https://maisan-447.github.io/fukumusume-career/`

**独自ドメインの場合：**
1. ドメインを取得（例：fukumusume-career.com）
2. GitHub Pages のカスタムドメイン設定 or Netlify等にデプロイ

## 運用メモ

- 料金やプラン内容を変更したら `index.html` と `tokushoho.html` の両方を更新する
- 初月返金保証の条件を変更する場合も `tokushoho.html` を更新する
