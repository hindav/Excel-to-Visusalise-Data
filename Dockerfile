# Dockerfile
FROM python:3.8

# Install system dependencies
COPY setup_chrome.sh /setup_chrome.sh
RUN chmod +x /setup_chrome.sh && /setup_chrome.sh

# Set working directory
WORKDIR /app

# Copy and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the app
COPY . .

# Expose Streamlit port
EXPOSE 8501

# Run Streamlit
ENTRYPOINT ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
