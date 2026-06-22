FROM python:3.10-slim

WORKDIR /app

COPY requirements.txt .

# Install build dependencies for prophet/statsmodels
RUN apt-get update && apt-get install -y \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir -r requirements.txt

CMD ["tail", "-f", "/dev/null"]
