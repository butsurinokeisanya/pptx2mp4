
from flask import request, redirect, url_for, render_template, flash, send_file
from pptx2mp4 import app
from werkzeug.utils import secure_filename
import os
import random
import string
import shutil

from pdf2image import convert_from_path
import glob
from moviepy.editor import *
import pptx
from werkzeug.utils import secure_filename

UPLOAD_FOLDER= r'.\pptx2mp4\static\mp4'
DOWNLOAD_FOLDER=r'.\pptx2mp4\static\mp4'
ALLOWED_EXTENSIONS = {'mp4'}

def allowed_file(filename, ALLOWED_EXTENSIONS):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def randomname(n):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=n))

def margemp4(input_folder, output_path):
    file_list = sorted([f for f in os.listdir(input_folder) if f.endswith('mp4')])
    # 動画ファイルのフルパスリストを作成
    file_paths = [os.path.join(input_folder, f) for f in file_list]

    # VideoFileClipオブジェクト作成
    clips = []
    for fp in file_paths:
        try:
            clip = VideoFileClip(fp)
            clips.append(clip)
        except Exception as e:
            print(f"Error loading video {fp}: {e}")

    # 結合
    if clips:
        final_clip = concatenate_videoclips(clips, method="compose")

        # 結果を保存
        final_clip.write_videofile(output_path, codec="libx264", audio_codec="aac")
    else:
        print("No clips to concatenate.")

def delete_all_files_in_folder(folder_path):
    try:
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")
        print("All files and folders have been deleted successfully.")
    except Exception as e:
        print(f"Failed to access folder {folder_path}. Reason: {e}")

@app.route('/marge_mp4_add', methods=['POST'])
def marge_mp4_add():
    delete_all_files_in_folder(UPLOAD_FOLDER)
    delete_all_files_in_folder(DOWNLOAD_FOLDER)
    if request.files.getlist('upload_files')[0].filename:
        upload_files=request.files.getlist('upload_files')
        for upload_file in upload_files:
            upload_file.save(os.path.join(UPLOAD_FOLDER,secure_filename(upload_file.filename)))

    margemp4(UPLOAD_FOLDER, os.path.join(DOWNLOAD_FOLDER, "marge.mp4"))

    flash('変換されました', 'alert alert-info')
    return render_template('core/marge_mp4.html',image_name="marge.mp4")

@app.route('/marge_mp4_download',methods=['GET'])
def marge_mp4_download():
    return send_file(f'static/mp4/marge.mp4',as_attachment=True)

