# NASDA バージョン管理表

Nelson Project のバージョン管理表をHTMLに変換したものです。

## URL

https://shenhaowei1992-sudo.github.io/nelson-version/

## 更新方法

xlsxを更新したら `node gen_html.js` を実行してHTMLを再生成してください。

または GitHub Actions が毎日自動でSVNからダウンロード＆生成します。

## 使い方

1. `バージョン管理表.xlsx` をこのディレクトリに配置
2. `node gen_html.js バージョン管理表.xlsx version_management.html` を実行
3. `version_management.html` が生成される