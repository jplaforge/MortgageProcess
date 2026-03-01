FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY src/ src/

ENV PORT=8000
ENV PYTHONPATH=/app/src

EXPOSE ${PORT}

CMD ["python", "-m", "mortgage_mcp.server"]
