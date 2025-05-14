from flask import Flask
import subprocess

app = Flask(__name__)

@app.route("/executar", methods=["GET"])
def executar_script():
    resultado = subprocess.run(["python3", "relatorio.py"], capture_output=True, text=True)
    return f"<pre>{resultado.stdout or resultado.stderr}</pre>"

if __name__ == "__main__":
    app.run(debug=True)