# Developer helper: export a .docx to PDF with Microsoft Word, to check how a
# converted file really looks (page count, overflow). Needs Word installed.
# Usage:  .\docker\word-preview.ps1 C:\path\file.docx C:\path\out.pdf
param([Parameter(Mandatory = $true)][string]$Docx, [Parameter(Mandatory = $true)][string]$Pdf)

$docxPath = (Resolve-Path -LiteralPath $Docx).Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($docxPath, $false, $true)      # read-only
    $doc.ExportAsFixedFormat($Pdf, 17)                         # 17 = PDF
    Write-Output "word pages: $($doc.ComputeStatistics(2))"
    # Each source page is one section; list the ones that spill onto more than one Word page.
    $prevEnd = 0; $spill = @()
    for ($i = 1; $i -le $doc.Sections.Count; $i++) {
        $end = $doc.Sections.Item($i).Range.Information(3)      # 3 = page number at the end
        if ($end - $prevEnd -gt 1) { $spill += "$i(+$($end - $prevEnd - 1))" }
        $prevEnd = $end
    }
    Write-Output "source pages: $($doc.Sections.Count); spilling onto extra pages: $(if ($spill) { $spill -join ' ' } else { 'none' })"
    $doc.Close($false)
} finally {
    $word.Quit()
}
