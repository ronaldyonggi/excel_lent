FROM python:3.9-slim

# Set the working directory
WORKDIR /app

# Copy the current directory contents into the container at /app
COPY . /app

# Install required packages specified in requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Install LibreOffice
RUN apt-get update && \
    apt-get install -y libreoffice && \
    apt-get clean

# Make port 80 available 
EXPOSE 80

# Run main.py when the container launches
CMD ["python", "main.py"]