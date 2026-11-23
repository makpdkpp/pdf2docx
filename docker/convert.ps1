# Convert a PDF to Word using the client container.
# Usage:  .\docker\convert.ps1 C:\path\file.pdf [extra pdf2docx.py options, e.g. --pages 1-10 --workers 6]
# Add --local to run on the CPU without the vLLM server (about 2-3 minutes per page).
# The .docx and the .pages cache folder are written next to the PDF.
param([Parameter(Mandatory = $true)][string]$Pdf,
      [Parameter(ValueFromRemainingArguments = $true)][string[]]$Rest)

$pdfPath = (Resolve-Path -LiteralPath $Pdf).Path
$dir     = Split-Path $pdfPath -Parent
$name    = Split-Path $pdfPath -Leaf
$root    = Split-Path $PSScriptRoot -Parent
$script  = Join-Path $root "pdf2docx.py"
$model   = Join-Path $root "model"          # only read when --local is used

docker run --rm -v "${dir}:/data" -v "${script}:/app/pdf2docx.py:ro" -v "${model}:/app/model:ro" `
  --entrypoint python3 pdf2docx-client /app/pdf2docx.py "/data/$name" `
  --base-url http://host.docker.internal:8000/v1 @Rest
