# 使い方

## python環境を持っている人向け (CUIメインで利用)
1. このレポジトリをクローン
1. `cp item_list_sample.xlsx item_list.xlsx`
1. コピーした`item_list.xlsx`に自動入力したい品目を記載
1. [ChromeDriver](https://chromedriver.chromium.org/)をダウンロードして解凍
2. 解凍した実行ファイルをこのレポジトリ直下に配置
3. `cp settings_sample.py settings.py`
4. コピーした `setting.py` にJシステムのログインIDとパスワードを記入（必要なら年度を変更）
4. Mac/Linuxの場合は`DRIVER_PATH`から`.exe`を削除
5. 必要なライブラリをインストール（pipやcondaなど）
5. `python auto_input.py`

## python環境を持っていない人向け (GUIメインで利用, windowsユーザーのみ)
1. このレポジトリをクローン
2. [ChromeDriver](https://chromedriver.chromium.org/)をダウンロードして解凍
2. 解凍した実行ファイルをこのレポジトリ直下に配置
3. `item_list_sample.xlsx`に物品名等を入力
4. `auto_input_GUI.exe`を実行
5. 以下のwindowが出るので各項目入力し「次へ」をクリック
 ![image](https://user-images.githubusercontent.com/41857834/161011843-304f2651-a7bd-4fe5-87c1-0b69bb7667f1.png)
