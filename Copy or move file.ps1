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

    .PARAMETER Action
        - 'copy' : Copy files from the source folder to the destination folder.
        - 'move' : Move files from the source folder to the destination folder.

        Action value is not case sensitive.

    .PARAMETER Source.Folder
        The source folder.

    .PARAMETER Source.Recurse
        - TRUE  : search root folder and child folders for files.
        - FALSE : search only in root folder for files.

    .PARAMETER Source.MatchFileNameRegex
        Only files that match the regex will be copied.

        Example:
        - '*.*'    : process all files.
        - '*.xlsx' : process only Excel files.

    .PARAMETER Destination.Folder
        The destination folder.

    .PARAMETER Destination.OverWriteFile
        - TRUE  : overwrite duplicate files in the destination folder.
        - FALSE : do not overwrite duplicate files in the destination folder
                  and log an error.

    .PARAMETER ProcessFilesCreatedInTheLastNumberOfDays
        Process files that are created in the last x days.

        Example:
        - 0 : Process all files in the source folder, no filter
        - 1 : Process files created since yesterday morning
        - 5 : Process files created in the last 5 days

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
    function New-LogFileNameHC {
        <#
    .SYNOPSIS
        Generates strings that can be used as a file name.

    .DESCRIPTION
        Converts strings or paths to usable formats for file names and adds the
        date if required. It filters out all the unaccepted characters by
        Windows to use a UNC-path or local-path as a file name. It's also
        useful for adding the date to a string. In case a path is provided, the
        first letter will be in upper case and the rest will be in lower case.
        It will also check if the log file already exists, and if so, create a
        new one with an increased number [0], [1], ..

    .PARAMETER LogFolder
        Folder path where the log files are located.

    .PARAMETER Name
        Can be a path name or just a string.

    .PARAMETER Date
        Adds the date to the name. When using one of the 'Script-options', make
        sure to use 'Get-ScriptRuntime (Start/Stop)' in your script.

        Valid options:
        'ScriptStartTime' : Start time of the script
        'ScriptEndTime'   : End time of the script
        'CurrentTime'     : Time when the command ran

    .PARAMETER Location
        Places the selected date in front or at the end of the name.

        Valid options:
        - Begin : 2014-09-25 - Name (default)
        - End   : Name - 2014-09-25

    .PARAMETER Format
        Format used for the selected date.

        Valid options:
        - yyyy-MM-dd HHmm (DayOfWeek)   : 2014-09-25 1431 (Thursday) (default)
        - yyyy-MM-dd HHmmss (DayOfWeek) : 2014-09-25 143121 (Thursday)
        - yyyyMMdd HHmm (DayOfWeek)     : 20140925 1431 (Thursday)
        - yyyy-MM-dd HHmm               : 2014-09-25 1431
        - yyyyMMdd HHmm                 : 20140925 1431
        - yyyy-MM-dd                    : 2014-09-25
        - yyyyMMdd                      : 20140925

    .PARAMETER NoFormatting
        Doesn't change the string to phrase format with a capital in the
        beginning. However, it still removes/replaces all characters that are
        not allowed in a Windows file name.

    .PARAMETER Unique
        When this switch is set, we will first check if a file exists with the
        same name. If it does, we add a number to the file, every time it runs
        the counter will go up.

    .EXAMPLE
        $params = @{
            LogFolder = 'T:\Log folder'
            Name      = 'Drivers.log'
            Date      = 'CurrentTime'
            Position  = 'End'
        }
        New-LogFileNameHC @params

        Create the string 'T:\Log folder\Drivers - 2015-01-26 1028 (Monday).log'

    .EXAMPLE
        $params = @{
            LogFolder = 'T:\Log folder'
            Format    = 'yyyyMMdd'
            Name      = 'Drivers.log'
            Date      = 'CurrentTime'
            Position  = 'Begin'
        }
        New-LogFileNameHC @params

        Create the string 'T:\Log folder\20220621 - Drivers.log'

    .EXAMPLE
        $params = @{
            LogFolder = 'T:\Log folder'
            Name      = 'Drivers.log'
        }
        New-LogFileNameHC @params

        Create the string 'T:\Log folder\Drivers.log'
    #>

        [CmdletBinding()]
        param (
            [parameter(Mandatory = $true, Position = 0, ParameterSetName = 'Set1')]
            [parameter(Mandatory = $true, Position = 0, ParameterSetName = 'Set2')]
            [ValidateScript({ Test-Path $_ -PathType Container })]
            [String]$LogFolder,
            [parameter(Mandatory = $true, Position = 1, ParameterSetName = 'Set1', ValueFromPipeline = $true)]
            [parameter(Mandatory = $true, Position = 1, ParameterSetName = 'Set2', ValueFromPipeline = $true)]
            [ValidateNotNullOrEmpty()]
            [alias('Path')]
            [String[]]$Name,
            [parameter(Mandatory = $true, Position = 2, ParameterSetName = 'Set2')]
            [ValidateSet('ScriptStartTime', 'ScriptEndTime', 'CurrentTime')]
            [String]$Date,
            [parameter(Mandatory = $false, Position = 3, ParameterSetName = 'Set2')]
            [ValidateSet('Begin', 'End')]
            [alias('Location')]
            [String]$Position = 'Begin',
            [parameter(Mandatory = $false, Position = 4, ParameterSetName = 'Set2')]
            [ValidateSet('yyyy-MM-dd HHmm (DayOfWeek)', 'yyyy-MM-dd HHmmss (DayOfWeek)',
                'yyyyMMdd HHmm (DayOfWeek)', 'yyyy-MM-dd HHmm', 'yyyyMMdd HHmm',
                'yyyy-MM-dd', 'yyyyMMdd')]
            [String]$Format = 'yyyy-MM-dd HHmm (DayOfWeek)',
            [Switch]$NoFormatting,
            [Switch]$Unique
        )

        begin {
            if ($Date) {
                switch ($Date) {
                    'ScriptStartTime' { $d = $ScriptStartTime; break }
                    'ScriptEndTime' { $d = $ScriptEndTime; break }
                    'CurrentTime' { $d = Get-Date; break }
                }

                switch ($Format) {
                    'yyyy-MM-dd HHmm (DayOfWeek)' {
                        $DateFormat = "{0:00}-{1:00}-{2:00} {3:00}{4:00} ({5})" `
                            -f $d.Year, $d.Month, $d.Day, $d.Hour, $d.Minute, $d.DayOfWeek
                        break
                    }
                    'yyyy-MM-dd HHmmss (DayOfWeek)' {
                        $DateFormat = "{0:00}-{1:00}-{2:00} {3:00}{4:00}{5:00} ({6})" `
                            -f $d.Year, $d.Month, $d.Day, $d.Hour, $d.Minute, $d.Second, $d.DayOfWeek
                        break
                    }
                    'yyyyMMdd HHmm (DayOfWeek)' {
                        $DateFormat = "{0:00}{1:00}{2:00} {3:00}{4:00} ({5})" `
                            -f $d.Year, $d.Month, $d.Day, $d.Hour, $d.Minute, $d.DayOfWeek
                        break
                    }
                    'yyyy-MM-dd HHmm' {
                        $DateFormat = ($d).ToString("yyyy-MM-dd HHmm")
                        break
                    }
                    'yyyyMMdd HHmm' {
                        $DateFormat = ($d).ToString("yyyyMMdd HHmm")
                        break
                    }
                    'yyyy-MM-dd' {
                        $DateFormat = ($d).ToString("yyyy-MM-dd")
                        break
                    }
                    'yyyyMMdd' {
                        $DateFormat = ($d).ToString("yyyyMMdd")
                        break
                    }
                }

                switch ($Position) {
                    'Begin' { $Prefix = "$DateFormat - "; break }
                    'End' { $Postfix = " - $DateFormat"; break }
                }
            }
        }

        process {
            foreach ($N in $Name) {
                if ($N -match '[.]...$|[.]....$') {
                    $Extension = ".$($N.Split('.')[-1])"
                    $N = $N.Replace("$Extension", '')
                }

                if ($N -match '[\\]') {
                    $Path = $N -replace '\\', '_'
                    $Path = $Path -replace ':', ''
                    $Path = $Path -replace ' ', ''
                    $Path = $Path.TrimStart("__")

                    if ($NoFormatting) {
                        $N = $Path
                    }
                    else {
                        if ($Path -match '^[a-z]_') {
                            $N = $Path.Substring(0, 1).ToUpper() + $Path.Substring(1, 1) +
                            $Path.Substring(2, 1).ToUpper() + $Path.Substring(3).ToLower() # Local path
                        }
                        else {
                            $N = $Path.Substring(0, 1).ToUpper() + $Path.Substring(1).ToLower() # UNC-path
                        }
                    }
                }
                else {
                    if ($NoFormatting) {
                        $N = $N
                    }
                    else {
                        $N = $N.Substring(0, 1).ToUpper() + $N.Substring(1).ToLower()
                    }
                }

                if ($Unique) {
                    $templateFileName = "$LogFolder\$Prefix$N$Postfix - {0}$Extension"

                    $number = 0

                    $FileName = $templateFileName -f $number

                    while (Test-Path -LiteralPath $FileName) {
                        $number++
                        $FileName = $templateFileName -f $number
                    }
                }
                else {
                    $FileName = "$LogFolder\$Prefix$N$Postfix$Extension"
                }

                Write-Output $FileName
            }
        }
    }

    $ErrorActionPreference = 'stop'

    try {
        $ScriptStartTime = Get-Date

        #region Create log folder
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -EA Stop
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = '{0} - Error.txt' -f (New-LogFileNameHC @LogParams)
        }
        catch {
            Write-Warning $_

            $params = @{
                FilePath = "$PSScriptRoot\..\Error.txt"
            }
            "Failure:`r`n`r`n- Failed creating the log folder '$LogFolder': $_" | Out-File @params

            exit
        }
        #endregion

        try {
            #region Import .json file
            Write-Verbose "Import .json file '$ImportFile'"

            $jsonFileContent = Get-Content $ImportFile -Raw -Encoding UTF8 |
            ConvertFrom-Json
            #endregion

            $SourceFolder = $jsonFileContent.Source.Folder
            $MatchFileNameRegex = $jsonFileContent.Source.MatchFileNameRegex
            $DestinationFolder = $jsonFileContent.Destination.Folder
            $Recurse = $jsonFileContent.Source.Recurse
            $Action = $jsonFileContent.Action
            $OverWriteFile = $jsonFileContent.Destination.OverWriteFile

            #region Test .json file properties
            @(
                'Folder', 'MatchFileNameRegex'
            ).where(
                { -not $jsonFileContent.Source.$_ }
            ).foreach(
                { throw "Property 'Source.$_' not found" }
            )

            @(
                'Folder'
            ).where(
                { -not $jsonFileContent.Destination.$_ }
            ).foreach(
                { throw "Property 'Destination.$_' not found" }
            )

            #region Test Action value
            if ($Action -notmatch '^copy$|^move$') {
                throw "Action value '$Action' is not supported. Supported Action values are: 'copy' or 'move'."
            }
            #endregion

            #region Test integer value
            try {
                if ($jsonFileContent.ProcessFilesCreatedInTheLastNumberOfDays -eq '') {
                    throw 'a blank string is not supported'
                }

                [int]$ProcessFilesCreatedInTheLastNumberOfDays = $jsonFileContent.ProcessFilesCreatedInTheLastNumberOfDays

                if ($jsonFileContent.ProcessFilesCreatedInTheLastNumberOfDays -lt 0) {
                    throw 'a negative number is not supported'
                }
            }
            catch {
                throw "Property 'ProcessFilesCreatedInTheLastNumberOfDays' must be 0 or a positive number. Number 0 processes all files in the source folder. The value '$($jsonFileContent.ProcessFilesCreatedInTheLastNumberOfDays)' is not supported."
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
                    $null = [Boolean]::Parse($jsonFileContent.Source.$boolean)
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
                    $null = [Boolean]::Parse($jsonFileContent.Destination.$boolean)
                }
                catch {
                    throw "Property 'Destination.$boolean' is not a boolean value"
                }
            }
            #endregion
            #endregion

            #region Test folders exist
            @{
                'Source.Folder'      = $SourceFolder
                'Destination.Folder' = $DestinationFolder
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
        catch {
            throw "Input file '$ImportFile': $_"
        }
    }
    catch {
        Write-Warning $_
        "Failure:`r`n`r`n- $_" | Out-File -FilePath $logFile -Append
        exit
    }
}

process {
    try {
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
            Write-Verbose 'No files found, exit script'
            exit
        }
        #endregion

        #region Select files to process
        if ($ProcessFilesCreatedInTheLastNumberOfDays -eq 0) {
            $filesToProcess = $allSourceFiles
        }
        else {
            $compareDate = (Get-Date).AddDays(
                - $ProcessFilesCreatedInTheLastNumberOfDays
            ).Date

            $filesToProcess = $allSourceFiles.Where(
                { $_.CreationTime.Date -ge $compareDate }
            )
        }

        Write-Verbose "Found $($filesToProcess.Count) file(s) to process"

        if (!$filesToProcess) {
            Write-Verbose 'No files found, exit script'
            exit
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
                "Failure for source file '$($file.FullName)':`r`n`r`n- $_" | Out-File -FilePath $logFile -Append
            }
        }
    }
    catch {
        Write-Warning $_
        "Failure:`r`n`r`n- $_" | Out-File -FilePath $logFile -Append
    }
}
