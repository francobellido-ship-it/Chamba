# syntax=docker/dockerfile:1
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 CERTIFICATE_STORAGE_DIR=/data/certificates
WORKDIR /app
RUN --mount=type=secret,id=build_ca \
    sed -i 's|http://deb.debian.org|https://deb.debian.org|g' /etc/apt/sources.list.d/debian.sources \
    && if [ -n "$HTTPS_PROXY" ]; then export https_proxy="$HTTPS_PROXY" http_proxy="$HTTP_PROXY"; fi \
    && if [ -f /run/secrets/build_ca ]; then cp /run/secrets/build_ca /tmp/qcp-build-ca.crt; chmod 644 /tmp/qcp-build-ca.crt; export APT_CA="-o Acquire::https::CaInfo=/tmp/qcp-build-ca.crt"; fi \
    && apt-get ${APT_CA:-} update \
    && apt-get ${APT_CA:-} install -y --no-install-recommends libreoffice-calc fonts-liberation fonts-dejavu-core fonts-crosextra-carlito \
    && rm -rf /var/lib/apt/lists/* /tmp/qcp-build-ca.crt
COPY requirements.txt ./
RUN --mount=type=secret,id=build_ca \
    if [ -f /run/secrets/build_ca ]; then export PIP_CERT=/run/secrets/build_ca; fi; \
    pip install --no-cache-dir --require-hashes -r requirements.txt \
    && useradd --uid 10001 --create-home portal \
    && mkdir -p /data/certificates \
    && chown -R portal:portal /data
COPY --chown=portal:portal app ./app
COPY --chown=portal:portal scripts/smoke.py ./scripts/smoke.py
USER portal
EXPOSE 8000
CMD ["sh", "-c", "exec python -m uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-8000} --workers 1 --no-proxy-headers"]
