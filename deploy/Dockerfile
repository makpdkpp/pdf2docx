# CPU-only image for the pdf2docx web UI (no GPU, no vLLM).
# Build context: the project root (see docker-compose.yml).
# The model weights are NOT in the image: mount the model folder at /app/model.
FROM python:3.12-slim

ARG PANDOC_VERSION=3.6.4
RUN apt-get update \
 && apt-get install -y --no-install-recommends curl ca-certificates fonts-thai-tlwg \
 && curl -fsSL --retry 5 --retry-delay 5 --retry-all-errors -o /tmp/pandoc.deb \
      "https://github.com/jgm/pandoc/releases/download/${PANDOC_VERSION}/pandoc-${PANDOC_VERSION}-1-amd64.deb" \
 && dpkg -i /tmp/pandoc.deb && rm /tmp/pandoc.deb \
 && rm -rf /var/lib/apt/lists/*

# CPU builds of PyTorch (the default PyPI wheels bundle several GB of CUDA libraries)
RUN pip install --no-cache-dir --index-url https://download.pytorch.org/whl/cpu \
      "torch>=2.11.0,<2.14" "torchvision>=0.26.0,<0.29"

# The nemotron_parse package that ships inside the model repository
COPY model/pyproject.toml model/LICENSE model/THIRD_PARTY_NOTICES.md model/TOKENIZER_ATTRIBUTION.md \
     model/postprocessing.py model/latex2html.py /tmp/nemotron-package/
COPY model/nemotron_parse /tmp/nemotron-package/nemotron_parse
COPY model/vllm_tied_patch /tmp/nemotron-package/vllm_tied_patch
RUN pip install --no-cache-dir "/tmp/nemotron-package[client]" \
      pymupdf python-docx markdown fastapi uvicorn python-multipart \
 && pip check && rm -rf /tmp/nemotron-package

WORKDIR /app
COPY pdf2docx.py webapp.py /app/
COPY web /app/web

# Runs as whatever user docker-compose sets; keep caches somewhere always writable.
ENV PYTHONUNBUFFERED=1 PYTHONUTF8=1 HOME=/tmp HF_HOME=/tmp/hf HF_HUB_OFFLINE=1 \
    JOBS_DIR=/jobs MODEL_DIR=/app/model PORT=8080
EXPOSE 8080
HEALTHCHECK --interval=30s --timeout=5s --start-period=120s --retries=3 \
  CMD curl -fsS http://127.0.0.1:8080/healthz || exit 1
CMD ["python3", "/app/webapp.py"]
