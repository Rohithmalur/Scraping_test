FROM python:3.10-slim

# Install system dependencies and the official Google Chrome binary
RUN apt-get update && apt-get install -y \
    wget curl unzip gnupg libglib2.0-0 libnss3 libgconf-2-4 libfontconfig1 \
    && wget -q -O - https://google.com | apt-key add - \
    && sh -c 'echo "deb [arch=amd64] http://google.com stable main" >> /etc/apt/sources.list.d/google-chrome.list' \
    && apt-get update && apt-get install -y google-chrome-stable

# Set up the working directory inside the container
WORKDIR /app
COPY . /app

# Install the Python libraries
RUN pip install --no-cache-dir -r requirements.txt

# Command to execute your script
CMD ["python", "scraper.py"]
