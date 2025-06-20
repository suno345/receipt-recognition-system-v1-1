# 証憑画像認識システム v1.1統合版

## 📋 概要
証憑画像（レシート・領収書）を自動認識してExcel形式の仕入台帳を生成するシステムです。
複数商品対応、月別シート自動分割機能を搭載しています。

## 🚀 簡単な使用方法

### 1. システム起動
- `🚀 証憑画像認識システム_複数商品対応版.command` をダブルクリック
- 初回実行時にOpenAI APIキーの入力が求められます

### 2. 画像の準備
- `📁 画像フォルダ` に証憑画像を配置してください
- 対応形式: JPG, PNG, HEIC, HEIF

### 3. 結果の確認
- `📊 出力ファイル` に Excel ファイルが生成されます
- 月別にシートが自動分割されます

## 📁 フォルダ構成

```
証憑画像認識システム_v1.1統合版/
├── 🚀 証憑画像認識システム_複数商品対応版.command  # 起動ファイル
├── 📁 画像フォルダ/                               # 証憑画像を入れる
├── 📊 出力ファイル/                               # 結果Excel保存先
├── ⚙️ 設定ファイル/                               # API設定等
├── システムファイル/                              # プログラム本体
│   ├── main.py                                   # メインプログラム
│   ├── requirements.txt                          # 必要ライブラリ
│   └── README.txt                               # 技術情報
├── 使用方法.txt                                  # 詳細手順
└── README.md                                    # このファイル
```

## ⚙️ 必要な環境
- macOS（推奨）/ Windows / Linux
- Python 3.7以上
- インターネット接続（OpenAI API使用）

## 🔑 初回設定
1. OpenAI APIキーを取得: https://platform.openai.com/api-keys
2. システム起動時に表示される入力欄にAPIキーを入力
3. 設定は自動保存されます

## 📊 出力形式
- **ファイル名**: 仕入台帳YYYY.xlsx
- **シート構成**: 月別シート + 年間合計シート
- **項目**: 購入日、商品名、商品価格、送料、合計金額、店舗名、等

## 🛠️ 技術仕様
- **AI技術**: OpenAI GPT-4 Vision
- **処理能力**: 複数商品同時認識
- **精度**: 高精度文字認識
- **言語**: 日本語完全対応

## 📞 サポート
システムに関するお問い合わせは開発元までご連絡ください。

---
**証憑画像認識システム v1.1統合版**  
