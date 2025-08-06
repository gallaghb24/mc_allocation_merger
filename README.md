# Media Centre Allocation Merger
=======

A simple Streamlit application that consolidates multiple Media Centre allocation
exports into a single Excel workbook. Upload one or more allocation files and the
app returns a consolidated allocation sheet ready for download.

## Project Structure

```
.
├── streamlit_app.py   # Streamlit web application
├── requirements.txt   # Python dependencies
└── README.md          # Project documentation
```

## Installation

Use Python 3.11+.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Running the App

```bash
streamlit run streamlit_app.py
```

## Tests

The project uses `pytest`.

```bash
pytest
```

## Docker

A minimal container image can be built with a `Dockerfile` similar to:

```Dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 8501
CMD ["streamlit", "run", "streamlit_app.py"]
```

Build and run:

```bash
docker build -t mc_allocation_merger .
docker run -p 8501:8501 mc_allocation_merger
```

## Configuration

Streamlit can be configured via `.streamlit/config.toml`. Example:

```toml
[server]
port = 8501
```

When running in Docker, mount the configuration directory:

```bash
docker run -p 8501:8501 -v $(pwd)/.streamlit:/app/.streamlit mc_allocation_merger
```

## Running locally

Install dependencies:

```bash
pip install -r requirements.txt
```

Start the Streamlit app:

```bash
cd backend
./start.sh
```

Configuration values can be adjusted in `backend/config.yaml`.
