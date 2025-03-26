#Requires -Version 7

<#
    .SYNOPSIS
        Copy, rename and move files between the process computer and the local computer.

    .DESCRIPTION
        This script is triggered by a scheduled task that is executed by a user account with permissions on the SMB file share of the process computer.

        The script copies files based on their creation date from the source
        folder to a subfolder in the source folder 'yyyy\mm', on the process
        computer.

        After the copy, the files are moved from the source folder on the process computer to the destination folder on the local computer and
        renamed 'FileName_ddmmyyyy.pdf'.

        The script will only save errors in the log folder

    .PARAMETER ImportFile
        A .JSON file that contains all the parameters used by the script.

    .PARAMETER SourceFolder
        The source folder on the process computer.

    .PARAMETER ArchiveFolder
        The folder on the process computer where a copy each file will be saved
        for later troubleshooting.

    .PARAMETER DestinationFolder
        The folder where the files in the source folder will be moved too.

    .PARAMETER FileExtensions
        Only the files with a matching file extension will be processed. If
        blank, all files will be processed.

    .PARAMETER LogFolder
        The folder where the log files will be saved.

#>

[CmdLetBinding()]
param (
    [Parameter(Mandatory)]
    [string]$ImportFile,
    [string]$ScriptName = 'Process computer actions',
    [string]$LogFolder = "$PSScriptRoot\Log"
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
            $logFile = New-LogFileNameHC @LogParams
        }
        catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import .json file
        Write-Verbose "Import .json file '$ImportFile'"

        $jsonFileContent = Get-Content $ImportFile -Raw -Encoding UTF8 |
        ConvertFrom-Json
        #endregion

        $SourceFolder = $jsonFileContent.SourceFolder
        $ArchiveFolder = $jsonFileContent.ArchiveFolder
        $DestinationFolder = $jsonFileContent.DestinationFolder
        $FileExtensions = $jsonFileContent.FileExtensions

        #region Test .json file properties
        if (!$SourceFolder) {
            throw "The property 'SourceFolder' is missing in the .json file"
        }
        if (!$ArchiveFolder) {
            throw "The property 'ArchiveFolder' is missing in the .json file"
        }
        if (!$DestinationFolder) {
            throw "The property 'DestinationFolder' is missing in the .json file"
        }
        #endregion

        #region Test folders exist
        if (!(Test-Path -LiteralPath $SourceFolder -PathType Container)) {
            throw "The source folder '$SourceFolder' does not exist"
        }
        if (!(Test-Path -LiteralPath $ArchiveFolder -PathType Container)) {
            throw "The archive folder '$ArchiveFolder' does not exist"
        }
        if (!(Test-Path -LiteralPath $DestinationFolder -PathType Container)) {
            throw "The destination folder '$DestinationFolder' does not exist"
        }
        #endregion
    }
    catch {
        Write-Warning $_

        $params = @{
            FilePath = if ($logFile) { $logFile }
            else { "$PSScriptRoot\Error.log" }
        }
        "Failure:`r`n`r`n- $_" | Out-File @params
        exit
    }
}

process {
    try {
        #region Get files from source folder
        Write-Verbose "Get all files in source folder '$SourceFolder'"

        $allSourceFiles = Get-ChildItem -LiteralPath $SourceFolder -File

        if (!$allSourceFiles) {
            Write-Verbose 'No files found, exit script'

            exit
        }
        #endregion

        #region Select files to process
        $filesToProcess = $allSourceFiles

        if ($FileExtensions) {
            Write-Verbose "Select files with extension '$FileExtensions'"

            $filesToProcess = $allSourceFiles.where(
                { $FileExtensions -contains $_.Extension }
            )
        }
        else {
            Write-Verbose 'Select all files'
        }

        Write-Verbose "Found $($filesToProcess.Count) file(s) to process"

        if (!$filesToProcess) {
            Write-Verbose 'No files found, exit script'

            exit
        }
        #endregion

        #region Copy files from source folder to archive folder


        #endregion

        #region Move files from Source folder to Destination folder and rename


        #endregion
    }
    catch {
        Write-Warning $_
        "Failure:`r`n`r`n- $_" | Out-File -FilePath $logFile
    }
}
