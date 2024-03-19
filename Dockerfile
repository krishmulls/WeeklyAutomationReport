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
ENV APP_MODULE="app.py"

EXPOSE 8501

HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health

CMD ["streamlit", "run", "app.py", "--server.port=8501"]