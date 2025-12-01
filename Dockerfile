# Use official Python
FROM python:3.10-slim

# Prevent Python from making .pyc files
ENV PYTHONDONTWRITEBYTECODE=1

# Don’t buffer stdout/stderr
ENV PYTHONUNBUFFERED=1

# Set working directory inside container
WORKDIR /app

# Install required system dependencies
RUN apt-get update && apt-get install -y \
    libglib2.0-0 \
    libsm6 \
    libxrender1 \
    libxext6 \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements
COPY requirements.txt .

# Install Python libs
RUN pip install --no-cache-dir -r requirements.txt

# Copy whole project (including your PNG)
COPY . .

# Expose Streamlit port
EXPOSE 8501

# Run Streamlit
ENTRYPOINT ["streamlit", "run", "main.py", "--server.port=8501", "--server.address=0.0.0.0"]
