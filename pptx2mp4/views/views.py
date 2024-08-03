from flask import request, redirect, url_for, render_template, flash, session
from pptx2mp4 import app
from functools import wraps


def login_required(view):
    @wraps(view)
    def wrapped_view(**kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return view(**kwargs)
    return wrapped_view


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        if request.form['username'] != app.config['USERNAME'] \
           or request.form['password'] != app.config['PASSWORD']:
            flash('入力されたユーザー名やパスワードが正しくありません', 'alert alert-danger')
        else:
            session['logged_in'] = True
            flash('ログインしました', 'alert alert-info')
            return redirect(url_for('index'))
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    flash('ログアウトしました', 'alert alert-info')
    return redirect(url_for('index'))
