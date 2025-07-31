import webview
import os
from api import *

if __name__ == '__main__':
    html = os.path.abspath("templates/index.html")
    api = Api()
    webview.create_window("Consumo Interno", f"file://{html}", js_api=api, maximized=True)
    webview.start()