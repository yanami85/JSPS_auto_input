# 使い方

1. このレポジトリをクローン
1. `cp item_list_sample.xlsx item_list.xlsx`
1. コピーした`item_list.xlsx`に自動入力したい品目を記載
1. [ChromeDriver](https://chromedriver.chromium.org/)をダウンロードして解凍
2. 解凍した実行ファイルをこのレポジトリ直下に配置
3. `cp settings_sample.py settings.py`
4. コピーした setting.py にJシステムのログインIDとパスワードを記入（Mac/Linuxユーザーの方は`DRIVER_PATH`から`.exe`を削除してください）
5. 必要なライブラリをインストール（pipやcondaなど）
5. `python auto_input.py`
