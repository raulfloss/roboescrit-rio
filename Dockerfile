FROM mcr.microsoft.com/playwright/python:v1.60.0-jammy

WORKDIR /app

# Instala Google Chrome real (fingerprint autentico, bypassa PerimeterX)
RUN apt-get update && apt-get install -y wget gnupg xvfb \
    && wget -q -O - https://dl-ssl.google.com/linux/linux_signing_key.pub | apt-key add - \
    && echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" \
       > /etc/apt/sources.list.d/google-chrome.list \
    && apt-get update && apt-get install -y google-chrome-stable \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .

ENV PYTHONUNBUFFERED=1

CMD python server.py
