# Use a lightweight, stable Python runtime image
FROM python:3.11-slim

# Set the working directory inside the container
WORKDIR /app

# Prevent Python from writing .pyc files and buffer outputs for real-time logs
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Copy the requirements file first to take advantage of Docker layer caching
COPY requirements.txt /app/requirements.txt

# Install dependencies cleanly without storing a local cache
RUN pip install --no-cache-dir -r requirements.txt

# Copy your main application file
COPY main.py /app/main.py

# Expose port 8080 to match your target environment
EXPOSE 8080

# Run using the official fastapi production command
# This safely binds to 0.0.0.0 and overrides your script's local 127.0.0.1 settings
CMD ["fastapi", "run", "main.py", "--port", "8080", "--host", "0.0.0.0"]
