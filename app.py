# -*- coding: utf-8 -*-
"""Local runner entry point"""
from api.index import app

if __name__ == "__main__":
    app.run(debug=False, host="127.0.0.1", port=5000)
