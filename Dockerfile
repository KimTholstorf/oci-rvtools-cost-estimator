FROM python:3.12-slim

LABEL org.opencontainers.image.source="https://github.com/KimTholstorf/oci-rvtools-cost-estimator"
LABEL org.opencontainers.image.description="Convert RVTools Excel exports into an Oracle Cloud monthly cost estimate workbook"
LABEL org.opencontainers.image.licenses="MIT"

RUN pip install --no-cache-dir oci-rvtools

COPY docs/ /app/docs/

COPY Docker/entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

WORKDIR /data

ENTRYPOINT ["/entrypoint.sh"]
