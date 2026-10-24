"""Data Trust Dashboard: a small Flask app around the Purview Trust Score engine.

    paths    where files live (works from source and from the .exe)
    state    the one shared "what is loaded" record
    scoring  turns rows into the numbers the page shows
    loaders  the two sources: live Purview and a raw JSON file
    reports  the Excel report
    routes   the HTTP API
    web/     the page (index.html + static/ css and ES modules)
"""
from flask import Flask

from .paths import WEB_DIR
from .routes import api


def create_app():
    app = Flask(__name__, static_folder=str(WEB_DIR / "static"), static_url_path="/static")
    app.register_blueprint(api)
    return app
