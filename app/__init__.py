from flask import Flask
app = Flask(__name__)

@app.route('/')
def welcomePage():
    return 'Data Analytics Interface'

@app.route('/SecurtiesService')
def listFuncs():
    return 'Securities Service Mainpage'

if __name__ == '__main__':
    app.run(debug=True,host="0.0.0.0")

