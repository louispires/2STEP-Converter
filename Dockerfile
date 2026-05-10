# Stage 1: compile Tailwind CSS
FROM node:20-alpine AS css-builder

WORKDIR /build
COPY tailwind.config.js .
COPY templates/ templates/
COPY static/input.css static/input.css

RUN npm init -y && \
    npm install --save-dev tailwindcss@3 && \
    npx tailwindcss -i static/input.css -o static/styles.css --minify

# Stage 2: Python app
FROM continuumio/miniconda3:latest

RUN conda install -y -c conda-forge pythonocc-core && \
    pip install --no-cache-dir fastapi "uvicorn[standard]" python-multipart jinja2 aiofiles && \
    conda clean -afy

WORKDIR /app

COPY 2STEP-Converter.py app.py entrypoint.sh ./
COPY templates/ templates/
COPY static/ static/
COPY --from=css-builder /build/static/styles.css static/styles.css

RUN chmod +x entrypoint.sh && mkdir -p /app/uploads /app/output

EXPOSE 8000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
