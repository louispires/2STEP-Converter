# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

Fork of [yaneony/2STEP-Converter](https://github.com/yaneony/2STEP-Converter) — a CLI tool that converts mesh files (STL, 3MF, OBJ, AMF, IGES) to clean STEP solids using OpenCASCADE. This fork adds a Docker-hosted web UI on top of the upstream converter.

## Architecture

Two largely independent layers:

**Upstream CLI tool** (`converter.py`)
- Standalone Python script; reads config from `data/config.json` at import time (module-level globals)
- CLI args: `--tolerance`, `--output`, `--format`, `--reduce`, `--no-preview`, etc.
- No `--angular-tolerance` CLI arg — angular tolerance is config-file-only
- Spawns itself in subprocesses for sewing/refining to isolate OCC crashes
- All `input()` calls throw `EOFError` when stdin is closed — use output-file existence as the success signal

**Web UI layer** (this fork's additions)
- `app.py` — FastAPI server; runs `converter.py` as a subprocess per job via `ThreadPoolExecutor(max_workers=2)`
- `templates/index.html` — single-page Alpine.js UI; no build step
- `static/input.css` → compiled to `static/styles.css` by Tailwind at Docker build time

Key design decision: `angular_tolerance` has no CLI arg, so `app.py` writes it to `data/config.json` behind `_config_lock` before each subprocess call. This lock serialises conversions (one at a time despite max_workers=2) to prevent tolerance stomping between concurrent jobs.

Job state is persisted to `/app/output/jobs.json` so the queue survives container restarts.

## Build & run

**Local Docker (primary workflow):**
```bash
docker compose up --build
```
UI at http://localhost:8000. Uploads and outputs persist via bind mounts `./uploads` and `./output`.

**Rebuild CSS only** (after editing `templates/index.html` or `static/input.css`):
The Tailwind compile happens inside the Docker build (`node:22-alpine` stage). There is no local npm workflow — CSS is only compiled at `docker build` time.

**Push to GHCR:** CI (`docker-publish.yml`) builds and pushes on every push to `main` or on version tags (`v*`). Tagged releases get semver image tags; `main` pushes get `latest`.

## Upstreaming

`upstream` remote tracks `yaneony/2STEP-Converter`. Merge with:
```bash
git fetch upstream && git merge upstream/main
```
`converter.py` is upstream-owned — avoid modifying it. All fork-specific logic lives in `app.py`, `templates/`, `static/`, `Dockerfile`, and `docker-compose.yml`.

## Dockerfile structure

Two-stage build:
1. `node:22-alpine` — compiles Tailwind CSS, artifact is `static/styles.css`
2. `mambaorg/micromamba:latest` (Debian bookworm) — installs `pythonocc-core` via conda-forge as `$MAMBA_USER`, then switches back to `root` for runtime so Docker bind mounts are writable

The `apt-get upgrade` in stage 2 patches system CVEs beyond what the base image ships with.

## Key constraints

- `converter.py` reads `data/config.json` at **import time** — config must be written before the subprocess is launched, not during
- The `/upload` endpoint accepts `tolerance`, `angular_tolerance`, `format` as query params; these are snapshotted into the job dict so in-flight jobs are unaffected by slider changes
- Tailwind classes used in `index.html` must exist at build time — the config scans `templates/**/*.html` only; no JIT from JavaScript strings
