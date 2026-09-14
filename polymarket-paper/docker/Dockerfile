# Containerised from day one so moving to a VPS is a deploy, not a rewrite.
FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    TZ=UTC

WORKDIR /app

COPY pyproject.toml README.md ./
COPY src ./src
COPY scripts ./scripts

RUN pip install --no-cache-dir -e .

# State lives on disk and must survive a restart, so it is a volume, never
# a layer. Phase 1 puts the SQLite database here.
VOLUME ["/app/data"]

# Non-root: this process never needs to write outside its own data dir.
RUN useradd --create-home --uid 10001 polypaper && chown -R polypaper /app
USER polypaper

ENTRYPOINT ["python", "-m", "polypaper.scripts_phase0"]
CMD ["--discover"]
