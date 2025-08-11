# run_streamlit.py
import os, sys, time, webbrowser, threading
from streamlit.web import bootstrap

def resource_path(name: str) -> str:
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, name)
    return os.path.join(os.path.dirname(__file__), name)

script = resource_path("app.py")
port = "8501"

os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
os.environ["STREAMLIT_SERVER_PORT"] = port
os.environ["STREAMLIT_BROWSER_GATHERUSAGESTATS"] = "false"

def open_browser():
    time.sleep(1.5)
    webbrowser.open(f"http://localhost:{port}")

threading.Thread(target=open_browser, daemon=True).start()
bootstrap.run(script, "", [], flag_options=set())