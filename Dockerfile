# Use slim Python image
FROM python:3.10-slim

# Disable .pyc files and enable unbuffered logging
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Set working directory
WORKDIR /app

# Install reportlab dependencies
RUN apt-get update && apt-get install -y \
    libglib2.0-0 \
    libsm6 \
    libxrender1 \
    libxext6 \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first (for caching)
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy all project files
COPY . .

# Expose streamlit port
EXPOSE 8501

# Run Streamlit
ENTRYPOINT ["streamlit", "run", "streamlit.py", "--server.port=8501", "--server.address=0.0.0.0"]
