function Move-Files {
    param (
        [string]$CsvPath,
        [string]$RootPath
    )

    # import each line of csv into pipeline
    Import-Csv $CsvPath | Foreach-Object {
        # get filename and location (replace non-breaking space with space)
        $FileName = $_."File Name" -replace [char]0x00A0, " "
        $Location = $_."Location" -replace [char]0x00A0, " "

        # skip to next iteration if location doesn't exist
        if (!(Test-Path -LiteralPath $Location)) { return }

        # get file by finding matching location and filename
        $File = Get-ChildItem -LiteralPath $Location | Where-Object {
            # match escaped filename with optional leading whitespaces (since csv import trims strings)
            $_.Name -match "^\s*$([regex]::escape($FileName))$"
        }

        # skip to next iteration if no matching file found in location
        if ($null -eq $File) { return }
        
        # define delete dir's path in root, create it if it doesn't exist and define file's new path in delete dir
        $DeletePath = Join-Path $RootPath delete
        $DeleteDir = New-Item $DeletePath -ItemType Directory -Force
        $NewPath = Join-Path $DeleteDir.FullName $File.Name
        $i = 2

        # check if file exists in delete dir
        while (Test-Path -LiteralPath $NewPath) {
            # add number to file's base name and post-increment it
            $NewPath = Join-Path $DeleteDir.FullName ($File.BaseName + " " + $i++ + $File.Extension)
        }

        # move file to location
        Move-Item -LiteralPath $File.FullName -Destination $NewPath
    }
}

# https://stackoverflow.com/questions/28631419/how-to-recursively-remove-all-empty-folders-in-powershell
function Remove-EmptyFolders {
    param (
        [string]$Path
    )

    # Recursively search folder structure excluding "delete" folder
    $childDirs = Get-ChildItem -LiteralPath $Path -Directory -Exclude "delete"
    foreach ($childDir in $childDirs) {
        Remove-EmptyFolders $childDir.FullName
    }

    # Delete folder structure if no files in any level below it
    $isEmpty = !(Get-ChildItem -LiteralPath $Path)
    if ($isEmpty) {
        Write-Verbose "Removing empty folder: $Path" -Verbose
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
}

$CsvPath = ".\other\Delete.csv"
$RootPath = "\\$Env:USERDNSDOMAIN\DFS\Groups\PR\PARLIAMENTARY AFFAIRS\Parliamentary Affairs - Shared Drive Clean-up"
$RootPath = "$Env:USERPROFILE\Documents"

Move-Files $CsvPath $RootPath
Remove-EmptyFolders $RootPath

$Locations = $(Import-Csv $Path).Location | Sort-Object | Get-Unique
foreach ($Location in $Locations) {
    if (Test-Path $Location) {
        Write-Host "$Location stills exists"
    }
}