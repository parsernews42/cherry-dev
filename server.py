import os
import subprocess
from datetime import datetime
from flask import Flask, request, Response, send_from_directory

app = Flask(__name__)

PYTHON_EXE = "python3"
SCRIPTS_DIR = os.path.join(os.path.dirname(__file__), "parsers")

SCRIPTS = {
    "INTERFAX_Business_news": "INTERFAX_Business_news.py",
    "MASH_First_100_news": "MASH_First_100_news.py",
    "RIA_Ekonomika_news": "RIA_Ekonomika_news.py",
    "TASS_news": "TASS_news.py",
    "RGru_news": "RGru_news.py",
    "PRIME_news": "PRIME_news.py",
}

@app.route("/")
def index():
    return send_from_directory('.', 'index.html')

def stream_process_output(process, logfile_path):
    """Вывод stdout скрипта в SSE и лог-файл"""
    with open(logfile_path, "w", encoding="utf-8") as logfile:
        for line in iter(process.stdout.readline, ''):
            clean_line = line.rstrip()
            logfile.write(clean_line + "\n")
            logfile.flush()
            yield f"data: {clean_line}\n\n"
        yield "data: [Завершено]\n\n"

@app.route("/run-script-stream")
def run_script_stream():
    script_name = request.args.get('name')
    if not script_name or script_name not in SCRIPTS:
        return "Неверное имя скрипта", 400

    script_path = os.path.join(SCRIPTS_DIR, SCRIPTS[script_name])
    if not os.path.isfile(script_path):
        return f"Файл скрипта не найден: {script_path}", 404

    # Создаём папку Logs, если нет
    logs_dir = os.path.join(os.path.dirname(__file__), "Logs")
    os.makedirs(logs_dir, exist_ok=True)

    now = datetime.now().strftime("%d.%m.%Y %H.%M.%S")
    logfile_name = f"log {now}.txt"
    logfile_path = os.path.join(logs_dir, logfile_name)

    # Запуск скрипта
    process = subprocess.Popen(
        [PYTHON_EXE, script_path],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        universal_newlines=True,
    )

    return Response(stream_process_output(process, logfile_path), mimetype='text/event-stream')

if __name__ == "__main__":
    app.run(debug=True, threaded=True, port=5000)
