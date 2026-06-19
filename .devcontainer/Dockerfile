# Use the official Python 3.10 image
FROM python:3.10-buster

RUN git config --global user.email "alain.seys@outlook.com"
RUN git config --global user.name "Alain Seys"

# Install PIP requirements


# Optional: If you need to install Python packages here, uncomment the following lines
 COPY requirements.txt .
 RUN pip install -r requirements.txt

 RUN git config --global http.sslverify false
