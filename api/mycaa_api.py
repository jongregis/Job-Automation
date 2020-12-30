from flask import Flask
from JobAutomation.mycaa_main import runProgram

app = Flask(__name__)


def test():
    print("this is running a function")
    return 'this is a test'


@app.route('/run-mycaa')
def run_mycaa():
    return runProgram('09')


if __name__ == '__main__':
    app.run(debug=True)
