#!/bin/bash
# Script to run the Financial Automation Web Application

echo "Starting Financial Automation Web Application..."
echo "The app will open in your default web browser."
echo "Press Ctrl+C to stop the server."
echo ""

# Install dependencies if needed
if ! python3 -c "import streamlit" 2>/dev/null; then
    echo "Installing dependencies..."
    pip install -r requirements.txt
fi

# Run the Streamlit app
streamlit run app.py






