FROM python:3.11-slim

WORKDIR /app

RUN apt-get update && apt-get install -y \
    libreoffice \
    curl \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code only (test files excluded via .dockerignore)
COPY citations.py layout_formatter.py pipeline.py style_profiles.py main.py ./

# Copy CSL files if present, else download them
RUN mkdir -p csl_styles && \
    curl -L -o csl_styles/ieee.csl    https://www.zotero.org/styles/ieee && \
    curl -L -o csl_styles/apa-7.csl   https://www.zotero.org/styles/apa-7th-edition && \
    curl -L -o csl_styles/harvard.csl https://www.zotero.org/styles/harvard-university-of-warwick

COPY entrypoint.sh .
RUN chmod +x entrypoint.sh

RUN mkdir -p uploads outputs

EXPOSE 8000
ENTRYPOINT ["./entrypoint.sh"]
