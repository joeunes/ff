# Use official lightweight Python image
FROM python:3.12-slim

# Set working directory inside container
WORKDIR /app

# Prevent Python from writing pyc files to disc
ENV PYTHONDONTWRITEBYTECODE 1
# Prevent Python from buffering stdout and stderr
ENV PYTHONUNBUFFERED 1

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements file first to leverage cache
COPY requirements.txt /app/

# Install python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy all application files
COPY . /app/

# Create uploads and outputs directories
RUN mkdir -p uploads outputs

# Expose port
EXPOSE 5000

# Run with Gunicorn WSGI server
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "web_app:app", "--timeout", "120"]
