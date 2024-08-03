from flask import request, redirect, url_for, render_template, flash, send_file, session
from pptx2mp4 import app
from werkzeug.utils import secure_filename
import os
import cv2
import random
import string
import shutil
from pdf2image import convert_from_path

import os
import pyttsx3

UPLOAD_FOLDER = r'.\pptx2mp4\static\txt'
DOWNLOAD_FOLDER=r'.\pptx2mp4\static\wav'
ALLOWED_EXTENSIONS = {'txt'}


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def randomname(n):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=n))

def convert_text_to_speech(text,output_file):
    engine=pyttsx3.init()
    engine.save_to_file(text,output_file)
    engine.runAndWait()
def process_txt(input_txt, output_folder,image_name):
    # 出力フォルダを作成
    os.makedirs(output_folder, exist_ok=True)

    with open(input_txt, 'r', encoding='utf-8') as file:
        notes_text = file.read()

    if notes_text:
        wav_filename = os.path.join(output_folder, image_name+".wav")
        convert_text_to_speech(notes_text, wav_filename)
        print(f"Created {wav_filename}")


def delete_all_files_in_folder(folder_path):
    try:
        # フォルダの中のすべてのファイルとサブフォルダを削除
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)  # ファイルまたはリンクを削除
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)  # ディレクトリを削除
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")
        print("All files and folders have been deleted successfully.")
    except Exception as e:
        print(f"Failed to access folder {folder_path}. Reason: {e}")

@app.route('/txt2wav_add', methods=['POST'])
def txt2wav_add():
    delete_all_files_in_folder(UPLOAD_FOLDER)
    delete_all_files_in_folder(DOWNLOAD_FOLDER)
    image = request.files['upload_files']
    if image:
        if allowed_file(image.filename):
            image_name =secure_filename(image.filename)
            print("-----------------------------------------------------------------------------")
            print(image_name)
            image.save(os.path.join(UPLOAD_FOLDER, image_name))
        else:
            image_name = None
            flash('txtファイルではなかったので失敗しました.',
                  'alert alert-warning')
    else:
        image_name = None
    
    # 「CoInitialize は呼び出されていません。」で使えない.
    txt_path=os.path.join(UPLOAD_FOLDER, image_name)
    output_folder=DOWNLOAD_FOLDER
    process_txt(txt_path, output_folder,image_name.split(".")[0])

    flash('変換されました', 'alert alert-info')
    session['image_name'] = image_name  # Save image_name in the sess
    return render_template('core/txt2wav.html',image_name=image_name)

@app.route('/txt2wav_download', methods=['GET'])
def txt2wav_download():
    image_name = session.get('image_name')
    print(f"image_name retrieved from session: {image_name}")  # Debugging statement
    if image_name:
        return send_file(f'static/wav/{image_name.split(".")[0]}.wav', as_attachment=True)
    else:
        flash('ファイルが見つかりませんでした。', 'alert alert-warning')
        return redirect(url_for('txt2wav_add'))
