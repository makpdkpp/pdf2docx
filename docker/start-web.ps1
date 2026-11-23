# Start the web UI, then open http://localhost:8080
#   .\docker\start-web.ps1        OCR on the CPU (no vLLM server needed; 2-3 minutes per page)
#   .\docker\start-web.ps1 -Gpu   OCR on the vLLM server from start-server.ps1 (seconds per page)
# Uploaded PDFs and results are kept in the webdata\ folder of the project.
# Stop with: docker stop pdf2docx-web
param([int]$Port = 8080, [switch]$Gpu)

$root = Split-Path $PSScriptRoot -Parent
$data = Join-Path $root "webdata"
New-Item -ItemType Directory -Force $data | Out-Null

$ocrUrl = ""
if ($Gpu) {
    try { Invoke-WebRequest -UseBasicParsing -TimeoutSec 5 http://127.0.0.1:8000/v1/models | Out-Null }
    catch { Write-Host "The vLLM server is not answering. Start it first: .\docker\start-server.ps1"; exit 1 }
    $ocrUrl = "http://host.docker.internal:8000/v1"
}

docker rm -f pdf2docx-web 2>$null | Out-Null
docker run -d --name pdf2docx-web `
  -p "127.0.0.1:${Port}:8080" `
  -v "${root}\pdf2docx.py:/app/pdf2docx.py:ro" -v "${root}\webapp.py:/app/webapp.py:ro" `
  -v "${root}\web:/app/web:ro" -v "${root}\model:/app/model:ro" -v "${data}:/jobs" `
  -e JOBS_DIR=/jobs -e PYTHONUNBUFFERED=1 -e "OCR_BASE_URL=$ocrUrl" `
  --entrypoint python3 pdf2docx-client /app/webapp.py | Out-Null

Write-Host "Waiting for the web server..."
for ($i = 0; $i -lt 60; $i++) {
    if ((docker inspect -f '{{.State.Status}}' pdf2docx-web) -ne "running") {
        Write-Host "Web server stopped. Last log lines:"; docker logs --tail 30 pdf2docx-web; exit 1
    }
    try { Invoke-WebRequest -UseBasicParsing -TimeoutSec 3 "http://127.0.0.1:${Port}/api/jobs" | Out-Null
          Write-Host "Ready: http://localhost:${Port}   (stop with: docker stop pdf2docx-web)"; exit 0 }
    catch { Start-Sleep -Seconds 2 }
}
Write-Host "The web server did not answer in time. Check: docker logs pdf2docx-web"; exit 1
