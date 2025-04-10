#Requires -Version 7

<#
    .SYNOPSIS
        Copy or move files from one folder to another folder.

    .DESCRIPTION
        This script selects all files in the source folder that match the
        'MatchFileNameRegex' and filters these based on the creation time
        defined in 'ProcessFilesCreatedInTheLastNumberOfDays'.

        The selected files are copied or moved, depending on the 'Action' value
        from the source folder to the destination folder.

        This script is intended to be triggered by a scheduled task that has
        permissions in the source and destination folder.

        The script will only save errors in the log folder.

    .PARAMETER ImportFile
        A .JSON file that contains all the parameters used by the script.

    .PARAMETER Tasks
        One copy or move job for each task.

    .PARAMETER Task.Action
        - 'copy' : Copy files from the source folder to the destination folder.
        - 'move' : Move files from the source folder to the destination folder.

        Action value is not case sensitive.

    .PARAMETER Task.Source.Folder
        The source folder.

    .PARAMETER Task.Source.Recurse
        - TRUE  : search root folder and child folders for files.
        - FALSE : search only in root folder for files.

    .PARAMETER Task.Source.MatchFileNameRegex
        Only files that match the regex will be copied.

        Example:
        - '.*'        : process all files.
        - '.*\.xlsx$' : process only Excel files.

    .PARAMETER Task.Destination.Folder
        The destination folder.

    .PARAMETER Task.Destination.OverWriteFile
        - TRUE  : overwrite duplicate files in the destination folder.
        - FALSE : do not overwrite duplicate files in the destination folder
                  and log an error.

    .PARAMETER Task.ProcessFilesCreatedInTheLastNumberOfDays
        Process files that are created in the last x days.

        Example:
        - 0 : Process all files in the source folder, no filter
        - 1 : Process files created today
        - 2 : Process files created since yesterday morning
        - 5 : Process files created in the last 4 days

    .PARAMETER LogFolder
        The folder where the error log files will be saved.
#>

[CmdLetBinding()]
param (
    [Parameter(Mandatory)]
    [string]$ImportFile,
    [string]$ScriptName = 'Copy or move file',
    [string]$LogFolder = "$PSScriptRoot\..\Errors"
)

begin {
    $ErrorActionPreference = 'stop'

    try {
        $scriptStartTime = Get-Date

        #region Create log folder
        try {
            $logFolderItem = New-Item -Path $LogFolder -ItemType 'Directory' -Force -EA Stop

            $baseLogName = Join-Path -Path $logFolderItem.FullName -ChildPath (
                '{0} - {1}' -f $scriptStartTime.ToString('yyyy_MM_dd_HHmmss_dddd'), $ScriptName
            )

            $logFile = '{0} - Error.txt' -f $baseLogName
        }
        catch {
            Write-Warning "Failed creating the log folder '$LogFolder': $_"

            try {
                "Failed creating the log folder '$LogFolder': $_" |
                Out-File -FilePath "$PSScriptRoot\..\Error.txt"
            }
            catch {
                Write-Warning "Failed creating fallback error file: $_"
            }

            exit 1
        }
        #endregion

        try {
            #region Import .json file
            Write-Verbose "Import .json file '$ImportFile'"

            $jsonFileContent = Get-Content $ImportFile -Raw -Encoding UTF8 |
            ConvertFrom-Json
            #endregion

            $Tasks = $jsonFileContent.Tasks

            if (-not $Tasks) {
                throw "Property 'Tasks' cannot be empty"
            }

            foreach ($task in $Tasks) {
                #region Test .json file properties
                @(
                    'Folder', 'MatchFileNameRegex'
                ).where(
                    { -not $task.Source.$_ }
                ).foreach(
                    { throw "Property 'Source.$_' not found" }
                )

                @(
                    'Folder'
                ).where(
                    { -not $task.Destination.$_ }
                ).foreach(
                    { throw "Property 'Destination.$_' not found" }
                )

                #region Test Action value
                if ($task.Action -notmatch '^copy$|^move$') {
                    throw "Action value '$($task.Action)' is not supported. Supported Action values are: 'copy' or 'move'."
                }
                #endregion

                #region Test integer value
                try {
                    if ([string]::IsNullOrEmpty($task.ProcessFilesCreatedInTheLastNumberOfDays)) {
                        throw 'a blank string or null is not supported'
                    }

                    [int]$ProcessFilesCreatedInTheLastNumberOfDays = $task.ProcessFilesCreatedInTheLastNumberOfDays

                    if ($task.ProcessFilesCreatedInTheLastNumberOfDays -lt 0) {
                        throw 'a negative number is not supported'
                    }
                }
                catch {
                    throw "Property 'ProcessFilesCreatedInTheLastNumberOfDays' must be 0 or a positive number. Number 0 processes all files in the source folder. The value '$($task.ProcessFilesCreatedInTheLastNumberOfDays)' is not supported."
                }
                #endregion

                #region Test boolean values
                foreach (
                    $boolean in
                    @(
                        'Recurse'
                    )
                ) {
                    try {
                        $null = [Boolean]::Parse($task.Source.$boolean)
                    }
                    catch {
                        throw "Property 'Source.$boolean' is not a boolean value"
                    }
                }

                foreach (
                    $boolean in
                    @(
                        'OverWriteFile'
                    )
                ) {
                    try {
                        $null = [Boolean]::Parse($task.Destination.$boolean)
                    }
                    catch {
                        throw "Property 'Destination.$boolean' is not a boolean value"
                    }
                }
                #endregion
                #endregion

                #region Test folders exist
                @{
                    'Source.Folder'      = $task.Source.Folder
                    'Destination.Folder' = $task.Destination.Folder
                }.GetEnumerator().ForEach(
                    {
                        $key = $_.Key
                        $value = $_.Value

                        if (!(Test-Path -LiteralPath $value -PathType Container)) {
                            throw "$key '$value' not found"
                        }
                    }
                )
                #endregion
            }
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
    }
    catch {
        Write-Warning $_
        "Failure:`r`n`r`n- $_" | Out-File -FilePath $logFile -Append
        exit 1
    }
}

process {
    foreach ($task in $Tasks) {
        try {
            $Action = $task.Action
            $SourceFolder = $task.Source.Folder
            $MatchFileNameRegex = $task.Source.MatchFileNameRegex
            $Recurse = $task.Source.Recurse
            $DestinationFolder = $task.Destination.Folder
            $OverWriteFile = $task.Destination.OverWriteFile

            #region Get files from source folder
            Write-Verbose "Get all files in source folder '$SourceFolder'"

            $params = @{
                LiteralPath = $SourceFolder
                Recurse     = $Recurse
                File        = $true
            }
            $allSourceFiles = @(Get-ChildItem @params | Where-Object {
                    $_.Name -match $MatchFileNameRegex
                }
            )

            if (!$allSourceFiles) {
                Write-Verbose 'No files found in source folder'
                continue
            }
            #endregion

            #region Select files to process
            if ($ProcessFilesCreatedInTheLastNumberOfDays -eq 0) {
                Write-Verbose 'Process all files in source folder'
                $filesToProcess = $allSourceFiles
            }
            else {
                $compareDate = (Get-Date).AddDays(
                    - ($ProcessFilesCreatedInTheLastNumberOfDays - 1)
                ).Date

                $filesToProcess = $allSourceFiles.Where(
                    { $_.CreationTime.Date -ge $compareDate }
                )
            }

            Write-Verbose "Found $($filesToProcess.Count) file(s) to process"

            if (!$filesToProcess) {
                Write-Verbose "Found $($allSourceFiles.Count) files in source folder, but no file has a creation date in the last  $ProcessFilesCreatedInTheLastNumberOfDays days"
                continue
            }
            #endregion

            foreach ($file in $filesToProcess) {
                try {
                    #region Copy file to destination folder
                    try {
                        $params = @{
                            LiteralPath = $file.FullName
                            Destination = "$($DestinationFolder)\$($file.Name)"
                            Force       = $OverWriteFile
                        }

                        Write-Verbose "$Action file '$($params.LiteralPath)' to '$($params.Destination)'"

                        if ($Action -eq 'copy') {
                            Copy-Item @params

                        }
                        else {
                            Move-Item @params
                        }
                    }
                    catch {
                        throw "Failed to $Action file '$($params.LiteralPath)' to '$($params.Destination)': $_"
                    }
                    #endregion
                }
                catch {
                    Write-Warning $_
                    "Failure:`r`n`r`n- Action: $Action`r`n- Source folder: $SourceFolder`r`n- Destination folder: $DestinationFolder`r`n-File $($file.FullName)`r`n`r`nError: $_" | Out-File -FilePath $logFile -Append
                }
            }
        }
        catch {
            Write-Warning $_
            "Failure:`r`n`r`n- Action: $Action`r`n- Source folder: $SourceFolder`r`n- Destination folder: $DestinationFolder `r`n`r`nError: $_" | Out-File -FilePath $logFile -Append
        }
    }
}
