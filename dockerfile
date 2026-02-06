FROM mcr.microsoft.com/devcontainers/python:3.11

# Install system dependencies required by netCDF4 and lxml
RUN apt-get update && apt-get install -y --no-install-recommends \
    libhdf5-dev \
    libnetcdf-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy and install Python dependencies
COPY requirements.txt /tmp/
RUN pip install --upgrade pip && pip install --requirement /tmp/requirements.txt
