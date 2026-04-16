#!/bin/bash
# Ensure jobs.db is a file, not a directory (Docker volume quirk)
touch /app/jobs.db
exec uvicorn main:app --host 0.0.0.0 --port 8000
