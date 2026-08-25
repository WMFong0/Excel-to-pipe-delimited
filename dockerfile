FROM python:3.14-slim@sha256:ce40764625a4ff50df3548277632e7f96c4e77fe75fa848aae9885476e7df5a4

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV OPENBLAS_NUM_THREADS=1
ENV OMP_NUM_THREADS=1
ENV MKL_NUM_THREADS=1
ENV NUMEXPR_NUM_THREADS=1

COPY requirements.txt requirements.lock /app/

RUN pip install --no-cache-dir --require-hashes -r requirements.lock \
    && useradd --no-create-home --uid 10001 --shell /usr/sbin/nologin appuser

COPY main.py /app/main.py

USER 10001:10001

EXPOSE 8080

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8080/health', timeout=3).read(1)"

CMD ["uvicorn", "main:app", "--port", "8080", "--host", "0.0.0.0", "--limit-concurrency", "8", "--limit-max-requests", "1000", "--timeout-graceful-shutdown", "30", "--no-server-header"]
