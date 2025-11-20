function Merge-CsvFiles {
    param (
        [string]$MergeDir
    )

    if (!(Test-Path $mergeDir)) { return }

    $timestamp = Get-Date -f yyyyMMddHHmmss
    $outDir = ".\out"
    if (!(Test-Path $outDir)) { New-Item $outDir -ItemType Directory }
    $outPath = "$outDir\merged - $timestamp.csv"

    Get-ChildItem $MergeDir -Filter *.csv | Select-Object -ExpandProperty FullName |
    Import-Csv | Export-Csv $outPath -Encoding ansi -NoTypeInformation -Append
}

Merge-CsvFiles ".\out\merge"