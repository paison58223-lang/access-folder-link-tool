<h1 align = "center">
  <img width="40" vertical-align="middle" alt="image" src="https://github.com/user-attachments/assets/8f0bea7c-d3a4-4791-a3e0-37796f1c0016" />
  Access-Explorer 連携自動化ツール
  <img width="40" vertical-align="middle" src="https://github.com/user-attachments/assets/7d77e84f-7afe-4eff-b040-1b4f6e3da729" />
</h1>
～自動でフォルダを開き、ヒューマンエラーを <font color="#e63946">**ゼロ**</font> に。～

## 📖 想定される業務
業務システム上で、表示中のレコードに関連する「特定のWindowsフォルダ」を即座に開くシーンを想定しています。
下記のスクリーンショットでは各顧客IDに紐づいた顧客名、メールアドレスのAccessDBがあり、
顧客へ送るメールアドレスが特定のフォルダに顧客IDの名前で保存されています。

<img width="1920" height="1032" alt="image" src="https://github.com/user-attachments/assets/a9b401ea-f59d-4718-9693-b597597558c8" />


## 🚀 解決する課題
手動でフォルダを探し、案件IDなどをコピー＆ペーストする作業には以下のリスクが伴います。
* **ヒューマンエラー:** 似た名前のフォルダを誤って開いてしまう。
* **データ不整合:** 案件IDのコピー漏れや貼り付けミスにより、別案件のファイルとして処理してしまう。

本ツールは、これらを「操作の自動化」によって構造的に排除します。

## ✨ 実現する体験
* **ボタン1つでジャンプ:** 画面上のボタンをクリックするだけで、対応するフォルダを直接エクスプローラーで表示します。
* **クリップボード自動連携:** ジャンプと同時に「案件番号」や「顧客名」をメモリに保持。ファイル名のリネームや検索に即座に利用可能です。

## 🛠 使用技術
* Access VBA
* Windows API (ShellExecute)
