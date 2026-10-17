#!/bin/sh
cd /opt/random-question-bank
exec env HOST=0.0.0.0 PORT=8044 /opt/random-question-bank/.venv/bin/waitress-serve --listen=0.0.0.0:8044 app:app
