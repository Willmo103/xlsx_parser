# Flask Web Interface for Excel File Processing

This application provides a web interface to upload an Excel file, process the file using a predefined script, and download the processed file.

## Building and Running with Docker

1. Build the Docker image:

   ```bash
   docker build -t flask-container.
   ```

2. Run the Docker container:

   ```bash
   docker run -p 9876:9876 flask-container
   ```

3. [Access the web interface: `http://localhost:9876/`](http://localhost:9876/)
