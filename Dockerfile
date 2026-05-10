FROM continuumio/miniconda3:latest

RUN conda install -y -c conda-forge pythonocc-core && \
    pip install --no-cache-dir fastapi "uvicorn[standard]" python-multipart jinja2 && \
    conda clean -afy

WORKDIR /app

COPY 2STEP-Converter.py .
COPY app.py .
COPY templates/ templates/
COPY entrypoint.sh .
RUN chmod +x entrypoint.sh

RUN mkdir -p /app/uploads /app/output

EXPOSE 8000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
