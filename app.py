from flask import Flask, request, render_template, redirect, url_for
import pandas as pd
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        # if user does not select file, browser also
        # submit an empty part without filename
        if file.filename == '':
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            script_choice = request.form['script']
            
            if script_choice == 'delivery':
                process_delivery(file_path)
            elif script_choice == 'defense':
                process_defense(file_path)
            
            return redirect(url_for('index', processed=True))
    return render_template('index.html')

def process_delivery(file_path):
    from scripts.delivery import process_file
    process_file(file_path)

def process_defense(file_path):
    from scripts.defense import process_file
    process_file(file_path)

if __name__ == '__main__':
    app.run(debug=True)
