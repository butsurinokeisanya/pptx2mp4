from flask import Flask
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
app.config.from_object('pptx2mp4.config')

from pptx2mp4.views import pptx2mp4, views, index, pdf2png, txt2wav, txt2wav_voicepeak, pptx2mp4_voicepeak, marge_mp4, tex2mp4, tex2mp4_voicepeak
