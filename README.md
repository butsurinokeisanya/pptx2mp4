# Python による PowerPoint 動画作成 pptx2mp4

パワーポイントを動画に変換する自作ソフトを作成しました.
Pythonを用いてパワーポイントファイルから動画を作成するwebアプリケーションです.
pdf2png: pdfスライド1枚目をyoutube用サムネイルにする.
txt2wav (pyttsx3): テキストファイルを音声ファイルにする.
txt2wav (voicepeak): テキストファイルを音声ファイルにする (voicepeak使用).
pptx2mp4 (pyttsx3): パワーポイントファイルから動画を作成する.
pptx2mp4 (voicepeak): パワーポイントファイルから動画を作成する (voicepeak使用).
marge_mp4
![トップ画面](top.png)
## Git Clone

```
$ git clone https://github.com/butsurinokeisanya/pptx2mp4.git

```

## パッケージインストール

```
(venv) $ pip install -r requirements.txt
```


## アプリケーション起動

```
$ py app.py
```

## ブラウザで下記リンクにアクセス
```
http://127.0.0.1:5000/
```