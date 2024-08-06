from flask import request, redirect, url_for, render_template, flash, send_file, session
from pptx2mp4 import app
from werkzeug.utils import secure_filename
import os
import cv2
import random
import string
import shutil

from pdf2image import convert_from_path
from PIL import Image
UPLOAD_FOLDER = r'.\pptx2mp4\static\pdf'
DOWNLOAD_FOLDER=r'.\pptx2mp4\static\png'
ALLOWED_EXTENSIONS = {'pdf'}


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def randomname(n):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=n))

def pdf_to_images(pdf_path, output_folder, image_name):
    # Convert PDF to images
    images = convert_from_path(pdf_path)
    # Create output folder if not exists
    os.makedirs(output_folder, exist_ok=True)
    # Save images temporarily
    temp_image_path = os.path.join(output_folder, "temp_" + image_name + ".png")
    images[0].save(temp_image_path, "PNG")

    # Compress the image if it exceeds 2MB
    compress_image(temp_image_path, output_folder, image_name)

def compress_image(input_image_path, output_folder, image_name):
    img = Image.open(input_image_path)
    img.save(os.path.join(output_folder, image_name + ".png"), "PNG", optimize=True, quality=85)

    while os.path.getsize(os.path.join(output_folder, image_name + ".png")) > 2 * 1024 * 1024:
        img = Image.open(os.path.join(output_folder, image_name + ".png"))
        img.save(os.path.join(output_folder, image_name + ".png"), "PNG", optimize=True, quality=75)

    os.remove(input_image_path)  # Remove the temporary image

def make_folder(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

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


@app.route('/pdf2png_add', methods=['POST'])
def pdf2png_add():
    make_folder(UPLOAD_FOLDER)
    make_folder(DOWNLOAD_FOLDER)
    delete_all_files_in_folder(UPLOAD_FOLDER)
    delete_all_files_in_folder(DOWNLOAD_FOLDER)
    image = request.files['upload_files']
    delete_all_files_in_folder(UPLOAD_FOLDER)
    image_name=""
    if image:
        if allowed_file(image.filename):
            image_name = secure_filename(image.filename)
            image.save(os.path.join(UPLOAD_FOLDER, image_name))
        else:
            image_name = None
            flash('pdfファイルではなかったので失敗しました.',
                  'alert alert-warning')
    else:
        image_name = None
    
    pdf_path=os.path.join(UPLOAD_FOLDER, image_name)
    output_folder=DOWNLOAD_FOLDER
    pdf_to_images(pdf_path, output_folder, image_name.split(".")[0])

    flash('変換されました', 'alert alert-info')
    session['image_name'] = image_name  # Save image_name in the session
    return render_template('core/pdf2png.html',image_name=image_name)

@app.route('/pdf2png_download', methods=['GET'])
def pdf2png_download():
    image_name = session.get('image_name')
    print(f"image_name retrieved from session: {image_name}")  # Debugging statement
    if image_name:
        return send_file(f'static/png/{image_name.split(".")[0]}.png', as_attachment=True)
    else:
        flash('ファイルが見つかりませんでした。', 'alert alert-warning')
        return redirect(url_for('pdf2png_add'))

