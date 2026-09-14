# Dockerfile
# Python dependencies are pinned in requirements.txt. See DEPLOYMENT.md for
# production image-digest pinning and the distinction from apt package versions.
FROM python:3.12-slim-bookworm

LABEL org.opencontainers.image.title="Naqaa | نقاء" \
    org.opencontainers.image.description="Arabic lead cleaning and reporting platform"

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    HOME=/home/app

WORKDIR /app

# Keep Debian dependencies identical to Streamlit Community Cloud.
COPY packages.txt /app/packages.txt
RUN apt-get update \
    && xargs -r -a /app/packages.txt apt-get install -y --no-install-recommends \
    && rm -rf /var/lib/apt/lists/* \
    && groupadd --gid 1000 app \
    && useradd --uid 10001 --gid 1000 --create-home --shell /usr/sbin/nologin app

COPY requirements.txt /app/requirements.txt
RUN python -m pip install --no-cache-dir -r /app/requirements.txt \
    && python -m pip check

# Explicit copies: never COPY uploads, exports, .env, or secrets into the image.
COPY --chown=10001:1000 app.py /app/app.py
COPY --chown=10001:1000 .streamlit/config.toml /app/.streamlit/config.toml
COPY --chown=10001:1000 assets/Cairo.ttf.b64 assets/OFL-Cairo.txt /app/assets/

# Only the bundled public font belongs in the static directory. No lead data.
RUN mkdir -p /app/static \
    && python -c "from pathlib import Path; from app import _export_font_bytes; Path('/app/static/Cairo.ttf').write_bytes(_export_font_bytes())" \
    && chown -R 10001:1000 /app /home/app

# Group 1000 can read Render runtime secret files; the process is never root.
USER 10001:1000

# A hint for local tooling, not a fixed Render port. Render injects PORT.
EXPOSE 8501

HEALTHCHECK --interval=30s --timeout=5s --start-period=60s --retries=3 \
    CMD python -c "import os, urllib.request; p = os.environ.get('PORT') or '8501'; r = urllib.request.urlopen('http://127.0.0.1:' + p + '/_stcore/health', timeout=4); assert r.status == 200"

CMD ["/bin/sh", "-c", "PORT=\"${PORT:-8501}\"; export PORT; exec streamlit run app.py --server.address=0.0.0.0 --server.port=\"$PORT\""]
