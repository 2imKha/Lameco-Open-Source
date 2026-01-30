FROM python:3.10-slim

# System dependencies for lxml/pandas
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential libxml2-dev libxslt1-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install requirements first (better build cache)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy code and model weights
COPY . .

EXPOSE 5000

# Adjust FLASK_APP to match the team entrypoint
ENV FLASK_APP=prepare_and_train.py
CMD ["flask", "run", "--host=0.0.0.0", "--port=5000"]
