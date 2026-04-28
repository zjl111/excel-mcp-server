ARG BASE_IMAGE=docker.m.daocloud.io/library/python:3.12-slim
FROM ${BASE_IMAGE}

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_INDEX_URL=https://pypi.tuna.tsinghua.edu.cn/simple \
    FASTMCP_HOST=0.0.0.0 \
    FASTMCP_PORT=8017 \
    EXCEL_FILES_PATH=/app/excel_files

WORKDIR /app

COPY pyproject.toml README.md ./
COPY src ./src

RUN pip install --upgrade pip \
    && pip install .

RUN mkdir -p /app/excel_files

EXPOSE 8017

CMD ["excel-mcp-server", "streamable-http"]
