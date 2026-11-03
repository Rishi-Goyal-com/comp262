param([string]$Pptx, [string]$Pdf)
# Export a deck to PDF with PowerPoint so it can be checked visually
$ppt = New-Object -ComObject PowerPoint.Application
try {
    $pres = $ppt.Presentations.Open($Pptx, $true, $false, $false)
    $pres.SaveAs($Pdf, 32)
    $pres.Close()
} finally {
    $ppt.Quit()
}
