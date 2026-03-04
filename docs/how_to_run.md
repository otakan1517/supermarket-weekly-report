# 実行手順（非エンジニア向け）｜週次レポート自動生成（Python × Excel）

この手順書は、週次レポート（Excel）を **ワンコマンドで作る方法**をまとめたものです。  
基本は「①仮想環境を有効化 → ②コマンド実行 → ③90_quality確認」です。

---

## 1. 事前準備（初回のみ）

### 1-1. 作業フォルダへ移動
ターミナルを開いて、リポジトリ直下へ移動します。

```bash
cd supermarket-weekly-report
1-2. 仮想環境を有効化（Mac/Linux）
source .venv/bin/activate

✅ 有効化できると、ターミナルの先頭に (.venv) が表示されます。

1-3. （初回のみ）依存ライブラリのインストール

※すでに終わっていれば不要です。

python -m pip install -U pip
python -m pip install pandas numpy openpyxl
2. レポート生成（最新週自動）
python scripts/make_weekly_report.py --config configs/config.yml

生成物は outputs/ 配下に出力されます。

3. 週を指定して作る（W01/W02など）

configs/config.yml の target_week_label を変更して実行します。

W01を作る → target_week_label: "2026-W01"

W02を作る → target_week_label: "2026-W02"

最新週自動に戻す → target_week_label: ""

変更後に実行：

python scripts/make_weekly_report.py --config configs/config.yml
4. 配布前チェック（必ずやる）
4-1. まず 90_quality を確認

出力Excelを開いて、90_quality シートを確認します。

INFO ok → 基本チェックは問題なし（配布へ進む）

INFO ok 以外 → 欠損/不整合などの可能性があるため、原因確認してから配布

4-2. 次に 01_summary を確認

01_summary で、売上・粗利・粗利率の全体値が想定と大きくズレていないか確認します。

5. よくあるエラーと対処
5-1. python: command not found

原因：仮想環境が有効になっていない
対処：

source .venv/bin/activate
5-2. No such file or directory

原因：作業ディレクトリが違う（相対パスがズレる）
対処：リポジトリ直下へ戻って実行

cd supermarket-weekly-report
python scripts/make_weekly_report.py --config configs/config.yml
6. 運用メモ（おすすめ）

週指定で作った後は、target_week_label: "" に戻す（最新週自動に戻す）

生成後は必ず 90_quality → 01_summary の順に確認してから配布
