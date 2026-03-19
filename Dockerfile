FROM python:3.12-slim

RUN pip install --no-cache-dir pyyaml

COPY convert.py /app/convert.py

WORKDIR /data

ENTRYPOINT ["python", "/app/convert.py"]
CMD ["--help"]
