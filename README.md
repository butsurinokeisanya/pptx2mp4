# Python による PowerPoint 動画作成 pptx2mp4

パワーポイントを動画に変換する自作ソフトを作成しました. Pythonを用いてパワーポイントファイルから動画を作成するwebアプリケーションです. Windowsにのみ対応します.
pdf2png: pdfスライド1枚目をyoutube用サムネイルにする.
txt2wav (pyttsx3): テキストファイルを音声ファイルにする.
txt2wav (voicepeak): テキストファイルを音声ファイルにする (voicepeak使用).
pptx2mp4 (pyttsx3): パワーポイントファイルから動画を作成する.
pptx2mp4 (voicepeak): パワーポイントファイルから動画を作成する (voicepeak使用).
marge_mp4
![トップ画面](top.png)

## 実行環境
```
Windows11
Python 3.8.10
```

## requirements
```
Flask==3.0.3
flask_sqlalchemy==3.1.1
moviepy==1.0.3
opencv_python==4.10.0.84
pdf2image==1.17.0
Pillow==9.5.0
Pillow==10.4.0
pydub==0.25.1
pyttsx3==2.90
Werkzeug==3.0.3
```
## Git Clone

```
$ git clone https://github.com/butsurinokeisanya/pptx2mp4.git

```

## 移動
```
cd cd pptx2mp4
```
## パッケージインストール

```
$ pip install -r requirements.txt
```


## アプリケーション起動

```
$ py app.py
```

## ブラウザで下記リンクにアクセス
トップページが表示されれば成功である.
```
http://127.0.0.1:5000/
```

## 次回以降はクリックで起動できる.
boot.batはアプリケーション起動とブラウザでローカルホストhttp://127.0.0.1:5000/にアクセスするバッチファイルである. 次回からはboot.batをクリックで実行するとブラウザでwebアプリケーションが立ち上がる.