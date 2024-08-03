from flask import Flask, request, render_template, redirect, url_for, send_file
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
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            script_choice = request.form['script']
            
            if script_choice == 'delivery':
                processed_file_path = process_delivery(file_path)
            elif script_choice == 'defense':
                processed_file_path = process_defense(file_path)
            
            return send_file(processed_file_path, as_attachment=True)
    return render_template('index.html')

def process_delivery(file_path):
    from scripts.delivery import process_file
    output_path = file_path.replace('.xlsx', '_processed_delivery.xlsx')
    process_file(file_path)
    return output_path

def process_defense(file_path):
    from scripts.defense import process_file
    output_path = file_path.replace('.xlsx', '_processed_defense.xlsx')
    process_file(file_path)
    return output_path

if __name__ == '__main__':
    app.run(debug=True)
