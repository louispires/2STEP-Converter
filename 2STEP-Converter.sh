#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OS="$(uname -s)"
if [ "$OS" = "Darwin" ]; then
    LOCAL_ROOT="$HOME/Library/Application Support/STLtoSTP"
else
    LOCAL_ROOT="${XDG_DATA_HOME:-$HOME/.local/share}/STLtoSTP"
fi

if [ -d "$SCRIPT_DIR/lib" ]; then
    MM_ROOT="$SCRIPT_DIR/lib"
elif [ -d "$LOCAL_ROOT" ]; then
    MM_ROOT="$LOCAL_ROOT"
else
    echo "No existing environment found. Where should the environment be installed?"
    echo ""
    echo "  [1] Next to this script  (portable)"
    echo "  [2] $LOCAL_ROOT"
    echo ""
    read -rp "Your choice (1/2): " _choice
    if [ "$_choice" = "2" ]; then
        MM_ROOT="$LOCAL_ROOT"
    else
        MM_ROOT="$SCRIPT_DIR/lib"
    fi
    echo ""
fi

MM="$MM_ROOT/micromamba"
ENV="$MM_ROOT/env"
PY="$ENV/bin/python"
export MAMBA_ROOT_PREFIX="$MM_ROOT"
export CONDA_PKGS_DIRS="$MM_ROOT"
export PYTHONNOUSERSITE=1

export PATH="$ENV/bin:$PATH"

if [ ! -f "$MM" ]; then
    ARCH="$(uname -m)"
    if [ "$OS" = "Darwin" ]; then
        if [ "$ARCH" = "arm64" ]; then
            MM_URL="https://github.com/mamba-org/micromamba-releases/releases/download/2.6.0-0/micromamba-osx-arm64"
        else
            MM_URL="https://github.com/mamba-org/micromamba-releases/releases/download/2.6.0-0/micromamba-osx-64"
        fi
    else
        if [ "$ARCH" = "aarch64" ] || [ "$ARCH" = "arm64" ]; then
            MM_URL="https://github.com/mamba-org/micromamba-releases/releases/download/2.6.0-0/micromamba-linux-aarch64"
        else
            MM_URL="https://github.com/mamba-org/micromamba-releases/releases/download/2.6.0-0/micromamba-linux-64"
        fi
    fi

    mkdir -p "$MM_ROOT"
    echo "Downloading portable Python manager (one-time, ~10 MB) ..."
    curl -L --progress-bar -o "$MM" "$MM_URL"
    if [ $? -ne 0 ]; then
        echo "[ERROR] Download failed. Check your internet connection."
        exit 1
    fi
    MM_SIZE=$(stat -f%z "$MM" 2>/dev/null || stat -c%s "$MM")
    if [ "$MM_SIZE" -lt 5000000 ]; then
        echo "[ERROR] Download corrupt ($MM_SIZE bytes). Delete micromamba and retry."
        rm -f "$MM"
        exit 1
    fi
    chmod +x "$MM"
fi

if [ ! -f "$PY" ]; then
    echo "Setting up Python environment (one-time download, ~500 MB) ..."
    "$MM" create --prefix "$ENV" -c conda-forge python=3.12 pythonocc-core trimesh fast-simplification --yes
    if [ $? -ne 0 ]; then
        echo "[ERROR] Failed to create Python environment."
        exit 1
    fi
else
    if ! "$PY" -c "from OCC.Core.StlAPI import StlAPI_Reader" >/dev/null 2>&1; then
        echo "OpenCASCADE not found or broken -- reinstalling ..."
        "$MM" install --prefix "$ENV" -c conda-forge pythonocc-core --yes
        if [ $? -ne 0 ]; then
            echo "[ERROR] Failed to install pythonocc-core."
            exit 1
        fi
    fi
fi

if ! "$PY" -c "import trimesh; import fast_simplification" >/dev/null 2>&1; then
    echo "Installing trimesh ..."
    "$MM" install --prefix "$ENV" -c conda-forge trimesh fast-simplification --yes || "$PY" -m pip install trimesh fast-simplification
fi

"$PY" "$SCRIPT_DIR/converter.py" "$@"
