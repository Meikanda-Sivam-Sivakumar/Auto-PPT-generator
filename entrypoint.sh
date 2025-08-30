#!/bin/sh
exec gunicorn app:app --host 0.0.0.0 --port ${PORT:-8000}