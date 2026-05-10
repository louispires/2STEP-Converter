FROM continuumio/miniconda3:latest

RUN conda install -y -c conda-forge pythonocc-core && \
    conda clean -afy

WORKDIR /app
COPY 2STEP-Converter.py .
COPY entrypoint.sh .
RUN chmod +x entrypoint.sh

RUN mkdir -p /app/models

VOLUME ["/app/models"]

ENTRYPOINT ["/app/entrypoint.sh"]
