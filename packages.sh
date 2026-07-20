#!/bin/bash

# Update and install system dependencies
apt-get update
apt-get install -y \
    gcc \
    g++ \
    make \
    python3-dev \
    libxml2-dev \
    libxslt1-dev
