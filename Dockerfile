# Use a lightweight Python image
FROM python:3.9

# Set working directory
WORKDIR /app

# Copy files into container
COPY . .

# Install dependencies
RUN pip install flask python-docx openpyxl PyMuPDF

# Expose Flask port
EXPOSE 5000

# Run the app
CMD ["python", "app.py"]
