from flask import request, redirect, url_for, render_template, flash
from pptx2mp4 import app
from werkzeug.utils import secure_filename
import shutil

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/pdf2png')
def pdf2png():
    return render_template('core/pdf2png.html')

@app.route('/txt2wav')
def txt2wav():
    return render_template('core/txt2wav.html')

@app.route('/txt2wav_voicepeak')
def txt2wav_voicepeak():
    return render_template('core/txt2wav_voicepeak.html')

@app.route('/pptx2mp4')
def pptx2mp4():
    return render_template('core/pptx2mp4.html')

@app.route('/pptx2mp4_voicepeak')
def pptx2mp4_voicepeak():
    return render_template('core/pptx2mp4_voicepeak.html')

@app.route('/marge_mp4')
def marge_mp4():
    return render_template('core/marge_mp4.html')

@app.route('/tex2mp4')
def tex2mp4():
    return render_template('core/tex2mp4.html')

@app.route('/tex2mp4_voicepeak')
def tex2mp4_voicepeak():
    return render_template('core/tex2mp4_voicepeak.html')