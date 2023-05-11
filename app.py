import os
import openai
from flask import Flask, request, jsonify, render_template, redirect, url_for, flash, session
from werkzeug.utils import secure_filename
import pandas as pd
from io import BytesIO
import openpyxl
from langchain.chat_models import ChatOpenAI
from langchain.agents import create_pandas_dataframe_agent

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 16 MB
agent = None

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    global agent
    api_key = request.form['api_key']
    file = request.files['excel_file']
    if file.filename != '':
        _, ext = os.path.splitext(file.filename)
        if ext.lower() not in ['.xls', '.xlsx', '.csv']:
            flash('Please upload a valid Excel or CSV file', 'error')
            return redirect(url_for('index'))
        df = pd.read_excel(file) if ext.lower() in ['.xls', '.xlsx'] else pd.read_csv(file)
        session['api_key'] = api_key
        session['dataframe'] = df.to_json()
        return redirect(url_for('chat'))
    return redirect(url_for('index'))

@app.route('/preview_file', methods=['POST'])
def preview_file():
    file = request.files['file']
    if file:
        _, ext = os.path.splitext(file.filename)
        try:
            df = pd.read_excel(file) if ext.lower() in ['.xls', '.xlsx'] else pd.read_csv(file)
            table_html = df.head(5).to_html(index=False, classes='preview-table', border=0)
            return table_html
        except Exception as e:
            print(e)
            return "Error", 500
    return "Error", 400

@app.route('/chat', methods=['GET'])
def chat():
    if 'api_key' not in session or 'dataframe' not in session:
        return redirect(url_for('index'))

    openai_api_key = session['api_key']
    os.environ["OPENAI_API_KEY"] = openai_api_key
    df = pd.read_json(session['dataframe'])
    chat = ChatOpenAI(temperature=0.0)
    global agent
    agent = create_pandas_dataframe_agent(chat, df, verbose=True)
    return render_template('chat.html')

@app.route('/ask', methods=['POST'])
def ask_question():
    global agent
    question = request.form.get("message")
    if not question:
        return jsonify({"error": "No question provided"}), 400

    try:
        answer = agent.run(question)
        return jsonify({"answer": answer}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host="127.0.0.1",port=8080,debug=True)
