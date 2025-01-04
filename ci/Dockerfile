FROM python:3.10.2-bullseye
LABEL authors="a-arseniev"

RUN apt-get update && apt-get upgrade -y && \
  apt-get -y install \
    pip \
    bash \
    expect \
    curl \
    git \
    htop \
    vim \
  && apt-get clean

RUN mkdir /app

RUN pip install --no-cache poetry==1.4.0

COPY pyproject.toml poetry.lock /app/
COPY templates/ /app/templates/
COPY main.py/ /app/

ENV PYTHONPATH "${PYTHONPATH}:/app/"
ENV PYTHONUNBUFFERED=TRUE

RUN cd /app && poetry config virtualenvs.create false && poetry install

WORKDIR /app/