#!/bin/bash
set -e

# Redirect stdin from /dev/null to bypass any "Press Enter to exit" prompts
exec python /app/2STEP-Converter.py "$@" </dev/null
