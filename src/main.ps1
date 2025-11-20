function Get-Metadata {
    param (
        [string]$RootPath
    )

    $timestamp = Get-Date -f yyyyMMddHHmmss
    $folderName = Split-Path $RootPath -Leaf
    $outDir = .\out
    if (!(Test-Path $outDir)) { New-Item $outDir -ItemType Directory }
    $outPath = "$outDir\$folderName - $timestamp.csv"
    $folders = Get-ChildItem -LiteralPath $RootPath -Recurse -Directory

    # Define file types that usually have extended properties
    $supportedExtensions = @(".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx")

    $results = $folders | ForEach-Object -Parallel {
        # Get files in the folder
        $files = Get-ChildItem -LiteralPath $_ -File

        # Check if any file has a supported extension
        $hasSupportedFiles = $files | Where-Object {
            $using:supportedExtensions -contains $_.Extension.ToLower()
        } | Select-Object -First 1

        # Only create COM object if needed
        if ($hasSupportedFiles) {
            $shell = New-Object -ComObject Shell.Application
            $folderObj = $shell.NameSpace($_.FullName)
        }

        foreach ($file in $files) {
            $ext = $file.Extension.ToLower()

            # Default values
            $author = $lastAuthor = $revisionNumber = $dateCreated = $dateSaved = $null

            # Only call ExtendedProperty if supported
            if ($folderObj -and ($using:supportedExtensions -contains $ext)) {
                $item = $folderObj.ParseName($file.Name)
                if ($item) {
                    # https://learn.microsoft.com/en-us/windows/win32/properties/props
                    $fmtid = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}"
                    $author = [string]$item.ExtendedProperty("$fmtid 4")
                    $lastAuthor = $item.ExtendedProperty("$fmtid 8")
                    $revisionNumber = $item.ExtendedProperty("$fmtid 9")
                    $dateCreated = try { $item.ExtendedProperty("$fmtid 12").ToLocalTime() } catch {}
                    $dateSaved = try { $item.ExtendedProperty("$fmtid 13").ToLocalTime() } catch {}
                }
            }

            [PSCustomObject]@{
                "File Name"       = $file.Name
                "Location"        = $file.DirectoryName
                "Size (B)"        = $file.Length
                "Created"         = $file.CreationTime.ToLocalTime()
                "Modified"        = $file.LastWriteTime.ToLocalTime()
                "Accessed"        = $file.LastAccessTime.ToLocalTime()
                "Author"          = $author
                "Last Saved By"   = $lastAuthor
                "Revision Number" = $revisionNumber
                "Content Created" = $dateCreated
                "Date Last Saved" = $dateSaved
            }
        }
    } -ThrottleLimit 8

    $results | Export-Csv -LiteralPath $outPath -Encoding ansi -NoTypeInformation
}

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

Measure-Command { Get-Metadata "\\$Env:USERDNSDOMAIN\dfs\Groups\RECORDS" }
# Measure-Command { Get-Metadata "$Env:USERPROFILE\Documents" }

# Merge-CsvFiles ".\out\merge"