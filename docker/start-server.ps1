# Start the vLLM OCR server on the local GPU (Docker Desktop must be running).
# Settings are sized for an RTX 3060 12GB:
#   - FlashAttention 2 (version 3 needs a newer GPU generation)
#   - 60% of VRAM (~7 GB), up to 8 concurrent pages
# Usage:  .\docker\start-server.ps1                  (model folder: model)
#         .\docker\start-server.ps1 -Model model-v20  (another model folder in the project)
param([string]$Model = "model")
$model = Join-Path (Split-Path $PSScriptRoot -Parent) $Model
if (-not (Test-Path (Join-Path $model "model.safetensors"))) { Write-Host "No model.safetensors in $model"; exit 1 }

docker rm -f nemotron-parse 2>$null | Out-Null
docker run -d --name nemotron-parse --gpus all --shm-size 8g `
  -p 127.0.0.1:8000:8000 -v "${model}:/workspace:ro" `
  nemotron-parse:3.5.0.post8 `
  --model /workspace --served-model-name nemotron-parse --trust-remote-code `
  --dtype bfloat16 --distributed-executor-backend uni --tensor-parallel-size 1 `
  --max-model-len 16384 --max-num-seqs 8 --max-num-batched-tokens 16384 `
  --gpu-memory-utilization 0.6 --no-enable-prefix-caching `
  --attention-config '{\"flash_attn_version\":2}' `
  --limit-mm-per-prompt '{\"image\":1}' `
  --chat-template /workspace/chat_template.jinja --host 0.0.0.0 --port 8000 | Out-Null

Write-Host "Waiting for the server (about 2 minutes)..."
while ($true) {
    if ((docker inspect -f '{{.State.Status}}' nemotron-parse) -ne "running") {
        Write-Host "Server stopped. Last log lines:"; docker logs --tail 30 nemotron-parse; exit 1
    }
    try { Invoke-WebRequest -UseBasicParsing -TimeoutSec 3 http://127.0.0.1:8000/v1/models | Out-Null; break }
    catch { Start-Sleep -Seconds 5 }
}
Write-Host "Ready: http://127.0.0.1:8000/v1   (stop with: docker stop nemotron-parse)"
