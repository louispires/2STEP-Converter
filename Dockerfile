# Stage 1: compile Tailwind CSS
FROM node:22-alpine AS css-builder

WORKDIR /build
COPY tailwind.config.js .
COPY templates/ templates/
COPY static/input.css static/input.css

RUN npm init -y && \
    npm install --save-dev tailwindcss@3 && \
    npx tailwindcss -i static/input.css -o static/styles.css --minify

# Stage 2: Python app
FROM mambaorg/micromamba:latest

ARG MAMBA_DOCKERFILE_ACTIVATE=1

USER root
RUN apt-get update && \
    apt-get upgrade -y --no-install-recommends && \
    rm -rf /var/lib/apt/lists/*

USER $MAMBA_USER

RUN micromamba install -y -n base -c conda-forge python=3.12 pythonocc-core pip && \
    micromamba clean -afy && \
    pip install --no-cache-dir fastapi "uvicorn[standard]" python-multipart jinja2 aiofiles

USER root

WORKDIR /app

COPY converter.py app.py entrypoint.sh ./
COPY templates/ templates/
COPY static/ static/
COPY --from=css-builder /build/static/styles.css static/styles.css

RUN chmod +x entrypoint.sh && mkdir -p /app/uploads /app/output /app/data

EXPOSE 8000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
