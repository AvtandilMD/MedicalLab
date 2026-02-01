from flask import Flask, render_template
import os
import sys

def get_base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

base_path = get_base_path()
template_folder = os.path.join(base_path, 'templates')
app = Flask(__name__, template_folder=template_folder)

@app.route('/')
def index():
    return render_template('index.html')

if __name__ == '__main__':
    print("=" * 50)
    print("🏥 Premium Medi - მთავარი გვერდი")
    print("=" * 50)
    print("🌐 გახსენით: http://127.0.0.1:8080")
    print("=" * 50)
    app.run(debug=False, host='127.0.0.1', port=8080)