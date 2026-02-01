from flask import Flask

app = Flask(__name__)

@app.route('/')
def index():
    return '''
    <html>
    <head><meta charset="UTF-8"></head>
    <body style="font-family: Arial; text-align: center; padding: 50px;">
        <h1>✅ სერვერი მუშაობს!</h1>
        <p>Flask წარმატებით გაეშვა</p>
    </body>
    </html>
    '''

if __name__ == '__main__':
    print("სერვერი გაშვებულია: http://127.0.0.1:5000")
    app.run(debug=True, host='127.0.0.1', port=5000)