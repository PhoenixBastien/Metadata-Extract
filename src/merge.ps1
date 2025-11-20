function Merge-CsvFiles {
    param (
        [string]$MergeDir
    )

    $timestamp = Get-Date -f yyyyMMddHHmmss
    if (!(Test-Path $mergeDir)) { return }
    $outDir = ".\out"
    if (!(Test-Path $outDir)) { New-Item $outDir -ItemType Directory }
    $outPath = "$outDir\merged - $timestamp.csv"

    Get-ChildItem $MergeDir -Filter *.csv | Select-Object -ExpandProperty FullName |
    Import-Csv | Export-Csv $outPath -Encoding ansi -NoTypeInformation -Append
}

Merge-CsvFiles ".\out\merge"