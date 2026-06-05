FROM mcr.microsoft.com/playwright/python:v1.44.0-jammy

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .

ENV PYTHONUNBUFFERED=1

CMD python server.py
