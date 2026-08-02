function Build-EntraWordReport {
    <#
        STEP 2 (Word) — Placeholder. Will turn the exported JSON into a Word
        documentation file (tenant overview, tables per section). Not implemented yet.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ExportsRoot,
        [string]$SourceFolder
    )

    Write-Host "Word report is not implemented yet." -ForegroundColor Yellow
    Write-Host "The exported JSON under $ExportsRoot will be the input source." -ForegroundColor DarkGray
}
