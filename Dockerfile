FROM mcr.microsoft.com/playwright/python:v1.60.0-jammy

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
RUN apt-get update && apt-get install -y xvfb && rm -rf /var/lib/apt/lists/*
COPY . .

ENV PYTHONUNBUFFERED=1

CMD python server.py
