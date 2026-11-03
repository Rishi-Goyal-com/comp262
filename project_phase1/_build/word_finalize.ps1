param([string]$Docx, [string]$Pdf)
# Open the report in Word, fill in the table of contents, save, and export a PDF copy for checking
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($Docx)
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
    $doc.Save()
    $doc.SaveAs([ref]$Pdf, [ref]17)
    $doc.Close()
} finally {
    $word.Quit()
}
