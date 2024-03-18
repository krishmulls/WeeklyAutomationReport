FROM python:3.12

# Set the working directory in the container
WORKDIR /app

# Copy the requirements.txt file to the container
COPY requirements.txt /app/

# Install dependencies using pip
RUN pip install -r requirements.txt

# Copy the Streamlit app files to the container
COPY . /app

# Expose the port where the Streamlit app will run
EXPOSE 8501

# Set environment variables
ENV LC_ALL=C.UTF-8
ENV LANG=C.UTF-8

# Start the Streamlit app
CMD ["streamlit", "run", "Weekly_Meeting.py"]