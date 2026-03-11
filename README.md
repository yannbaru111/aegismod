# Aegis Mod

Twitch チャンネル向け自動モデレーションツール（Windows デスクトップアプリ）

## 機能

- 完全一致・類似コメント連投検知
- 連投速度検知
- NGワードリスト
- 段階的ペナルティ（警告・発言禁止）
- AI モデレーション β（sentence-transformers）
- リスナーコマンド（`!ranking` / `!score` / `!pena`）
- 誤処置取り消し（発言禁止解除・BAN解除・違反カウントリセット）

## ビルド方法

1. Python をインストール（https://www.python.org/ ）
   - インストール時に「Add Python to PATH」にチェック
2. ZIPを解凍し `build-run.bat` をダブルクリック
3. `dist\AegisMod.exe` が生成されます

## 初回設定

1. Twitch でBotアカウントを作成
2. https://twitchtokengenerator.com でACCESS TOKENを取得
3. チャットで `/mod Botアカウント名` を実行
4. AegisMod.exe を起動し「設定」からChannel Name・Bot Username・OAuth Tokenを入力

## データ保存先

```
~\.aegismod\
    config.json       # アプリ設定
    ai_data.json      # AIフィードバックデータ
    training.csv      # 学習データエクスポート（手動出力）
```

## ライセンス

Private
