# Excel-Lent: Distributed Data Processing with LibreOffice Calc

## Overview

**Excel-Lent** is a Python-based application designed to process Pandas DataFrames, create Excel files using LibreOffice Calc in headless mode, and compute the sum of all numeric columns. The application leverages distributed computing to efficiently handle large datasets by processing each column in parallel using multiple headless Calc instances.

## Key Features

- **Core IP Implementation:** Processes DataFrames and generates Excel files.
- **Distributed Computing:** Utilizes `ProcessPoolExecutor` and `subprocess` to distribute column processing across multiple Calc instances.
- **Resource Monitoring:** Uses `psutil` to monitor process resource utilization.
- **Performance Metrics:** Captures and logs process-level performance metrics.
- **Containerization:** Docker Compose setup for easy deployment.

## Prerequisites

- **Python 3.9+**
- **LibreOffice Installed Locally** (for running the application outside of Docker)
- **Docker and Docker Compose** (for containerized deployment)

## Installation

### Local Installation

1. **Clone the Repository:**

   ```sh
   git clone https://github.com/yourusername/excel-lent.git
   cd excel-lent
   ```

2. **Set up a virtual Environment:**

   ```sh
   python -m venv .venv
   source .venv/bin/activate  # On Windows use `.venv\Scripts\activate`
   ```

3. **Install dependencies:**

   ```sh
   pip install -r requirements.txt
   ```

4. **Configure LibreOffice Path:**

   Ensure that the `SOFFICE_PATH` in `excel_lent/config.py` points to the correct path of the LibreOffice binary on your system. For example:

   ```python
   SOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
   ```

### Docker Installation

1. **Clone the Repository:**

   ```sh
   git clone https://github.com/yourusername/excel-lent.git
   cd excel-lent
   ```

2. **Build the Docker Image:**

   ```sh
   docker-compose build
   ```

3. **Run the Docker Container:**

   ```sh
   docker-compose up
   ```

## Usage

### Local Execution

1. **Activate the Virtual Environment:**

   ```sh
   source .venv/bin/activate  # On Windows use `.venv\Scripts\activate`
   ```

2. **Run the Application:**

   ```sh
   python main.py
   ```

   This will process a sample DataFrame, generate an Excel file with the results, and print the column sums.

### Docker Execution

1. **Build the Docker Image:**

   ```sh
   docker-compose build
   ```

2. **Run the Docker Container:**

   ```sh
   docker-compose up
   ```

   This will start the container, process a sample DataFrame, generate an Excel file with the results, and print the column sums.

## Architecture Overview

### Core Components

- **ExcelLent Class (`core.py`):** Manages the processing of DataFrames, creates Excel files, and computes column sums.
- **Worker Function (`worker.py`):** Processes individual columns using headless LibreOffice Calc.
- **Utility Functions (`utils.py`):** Provides helper functions, such as generating a sample DataFrame.
- **Configuration (`config.py`):** Stores configuration settings, such as the path to the LibreOffice binary.

### Distributed Computing

- **ProcessPoolExecutor:** Distributes column processing across multiple worker processes.
- **Subprocess:** Manages headless LibreOffice Calc instances.
- **Logging:** Captures and logs resource usage and performance metrics.

## Performance Metrics

- **Resource Monitoring:** Uses `psutil` to monitor memory and CPU usage.
- **Performance Logging:** Logs the time taken to process each column and overall resource usage.

## Deployment

### Docker Compose

1. **Build the Docker Image:**

   ```sh
   docker-compose build
   ```

2. **Run the Docker Container:**

   ```sh
   docker-compose up
   ```

3. **Stop the Docker Container:**

   ```sh
   docker-compose down
   ```

4. **View Logs:**

   ```sh
   docker-compose logs
   ```

## Acknowledgments

- [Pandas](https://pandas.pydata.org/)
- [LibreOffice](https://www.libreoffice.org/)
- [Docker](https://www.docker.com/)
- [Docker Compose](https://docs.docker.com/compose/)

---
