FROM python:3.12

# Set the working directory in the container
WORKDIR /app

# Copy the poetry.lock and pyproject.toml files to the container
COPY poetry.lock pyproject.toml /app/

# Install dependencies using Poetry
RUN pip install poetry && poetry config virtualenvs.create false && poetry install

# Copy the Streamlit app files to the container
COPY . /app

# Expose the port where the Streamlit app will run
EXPOSE 8501

# Set environment variables
ENV LC_ALL=C.UTF-8
ENV LANG=C.UTF-8

# Start the Streamlit app
CMD ["streamlit", "run", "Weekly_Meeting.py"]