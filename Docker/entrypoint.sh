#!/bin/sh
# oci-rvtools-cost-estimator — dual-mode entrypoint
# No arguments: serve the web app on port 8080
# Any arguments: run the oci-rvtools CLI
if [ $# -eq 0 ]; then
    echo "Starting web app on http://localhost:8080"
    exec python3 -m http.server 8080 --directory /app/docs
else
    exec oci-rvtools "$@"
fi
