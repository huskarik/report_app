FROM python:3.12-slim

ENV TZ Europe/Moscow

RUN apt update

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV FLASK_APP=app.app:app

EXPOSE 5000

CMD ["flask", "run", "--host=0.0.0.0", "--port=5000"]