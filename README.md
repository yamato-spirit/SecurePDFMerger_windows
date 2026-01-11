# Secure PDF Merger & Converter (Python GUI Application)

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)

## 📖 概要 (Overview)
**Secure PDF Merger** は、PDFファイルの結合・ページ編集・フォーマット変換を直感的なGUIで行えるデスクトップアプリケーションです。
「機密情報を扱うため、オンラインの変換サイトを使いたくない」というニーズに応えるため、すべての処理をローカル環境で完結させ、メタデータ除去などのセキュリティ機能を実装しています。

**主なターゲット:** 事務処理の効率化、機密文書を取り扱う業務

## ✨ 主な機能 (Features)
* **直感的な操作:** ドラッグ＆ドロップによるページ並べ替え、ファイルの追加。
* **多機能プレビュー:** 結合前のPDFをページ単位で確認、回転、削除が可能。
* **マルチフォーマット対応:** PDFだけでなく、以下のファイルを自動でPDFに変換して取り込みます。※うまく読み込めない場合は、PDFファイルのみを読み込んでください
    * 画像 (.jpg, .png, .bmp, etc.)
    * Office文書 (.docx, .pptx, .xlsx) ※Windows環境推奨
    * Web/テキスト (.html, .txt, .csv)
* **編集履歴管理:** Undo/Redo（元に戻す/やり直し）機能をスタック管理で実装。
* **セキュアな出力:** 結合時に既存のメタデータを削除し、クリーンな状態で保存。

## 🛠 使用技術 (Tech Stack)
本アプリケーションは、PythonのモダンなGUIライブラリと強力なPDF操作ライブラリを組み合わせて構築されています。

### Languages & Frameworks
* **Python 3.10+**
* **CustomTkinter:** モダンでレスポンシブなUIデザインの実装。

### Key Libraries
* **PDF Manipulation:**
    * `pypdf`: PDFの結合、回転、メタデータ操作。
    * `PyMuPDF (fitz)`: 高速なレンダリングによるプレビュー画像生成。
* **File Conversion:**
    * `reportlab`, `xhtml2pdf`: テキスト、CSV、HTMLからのPDF生成。
    * `docx2pdf`, `comtypes`: Microsoft OfficeファイルのPDF変換 (Windows API活用)。
* **Data Handling:** `pandas` (CSV/Excel処理), `Pillow` (画像処理).

## 🏗 設計・こだわりポイント (Architecture & Highlights)
採用担当者様向けの技術的なポイントです。

1.  **クロスプラットフォーム対応とOS固有機能の分離**
    * Windows特有のCOM操作（Office変換）と、汎用的な処理を `try-import` ブロックやOS判定で分離し、Mac環境でもコア機能が動作するように設計しています。
2.  **ステート管理とUX**
    * 編集操作（削除・回転・並べ替え）ごとに状態をディープコピーして履歴スタック(`self.history`)に保存することで、**Undo/Redo**機能を実装し、ユーザーの誤操作を防止しています。
3.  **パフォーマンス最適化**
    * 大量のページを扱う際、プレビュー画像の生成をキャッシュ化(`self.thumbnail_cache`)し、スクロール時の再描画負荷を軽減しています。
    * ドラッグ中の再描画処理を軽量化し、できるだけカクつきのないスムーズな操作感を実現しました。

---

## 🚀 インストールと実行方法 (Installation & Usage)

### 前提条件 (Prerequisites)
* Python 3.10 以上がインストールされていること
* (Windowsのみ) Officeファイルの変換機能を使用する場合は Microsoft Word/PowerPoint がインストールされていること

### 1. リポジトリのクローン
```bash
git clone [https://github.com/your-username/secure-pdf-merger.git](https://github.com/your-username/secure-pdf-merger.git)
cd secure-pdf-merger