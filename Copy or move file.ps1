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
    [string]$ConfigurationJsonFile
)

begin {
    $ErrorActionPreference = 'stop'

    $systemErrors = @()
    $logFileData = [System.Collections.Generic.List[PSObject]]::new()
    $eventLogData = [System.Collections.Generic.List[PSObject]]::new()
    $scriptStartTime = Get-Date

    try {
        $eventLogData.Add(
            [PSCustomObject]@{
                DateTime  = $scriptStartTime
                Message   = 'Script started'
                EntryType = 'Information'
                EventID   = '100'
            }
        )

        #region Import .json file
        Write-Verbose "Import .json file '$ConfigurationJsonFile'"

        $jsonFileItem = Get-Item -LiteralPath $ConfigurationJsonFile -ErrorAction Stop

        $jsonFileContent = Get-Content $jsonFileItem -Raw -Encoding UTF8 |
        ConvertFrom-Json
        #endregion

        $Tasks = $jsonFileContent.Tasks

        #region Test .json file properties
        if (-not $Tasks) {
            throw "Property 'Tasks' cannot be empty"
        }

        foreach ($task in $Tasks) {
            @(
                'Action', 'Source', 'Destination'
            ).where(
                { -not $task.$_ }
            ).foreach(
                { throw "Property 'Tasks.$_' not found" }
            )

            @(
                'Folder', 'MatchFileNameRegex'
            ).where(
                { -not $task.Source.$_ }
            ).foreach(
                { throw "Property 'Tasks.Source.$_' not found" }
            )

            @(
                'Folder'
            ).where(
                { -not $task.Destination.$_ }
            ).foreach(
                { throw "Property 'Tasks.Destination.$_' not found" }
            )

            #region Test Action value
            if ($task.Action -notmatch '^copy$|^move$') {
                throw "'Tasks.Action' value '$($task.Action)' is not supported. Supported Action values are: 'copy' or 'move'."
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
                throw "Property 'Tasks.ProcessFilesCreatedInTheLastNumberOfDays' must be 0 or a positive number. Number 0 processes all files in the source folder. The value '$($task.ProcessFilesCreatedInTheLastNumberOfDays)' is not supported."
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
                    throw "Property 'Tasks.Source.$boolean' is not a boolean value"
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
                    throw "Property 'Tasks.Destination.$boolean' is not a boolean value"
                }
            }
            #endregion
        }
        #endregion
    }
    catch {
        $systemErrors += [PSCustomObject]@{
            DateTime = Get-Date
            Message  = "Input file '$ConfigurationJsonFile': $_"
        }

        Write-Warning $systemErrors[-1].Message

        return
    }
}

process {
    if ($systemErrors) { return }

    foreach ($task in $Tasks) {
        try {
            $Action = $task.Action
            $SourceFolder = $task.Source.Folder
            $MatchFileNameRegex = $task.Source.MatchFileNameRegex
            $Recurse = $task.Source.Recurse
            $DestinationFolder = $task.Destination.Folder
            $OverWriteFile = $task.Destination.OverWriteFile

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

                Write-Verbose "Process files created since '$compareDate'"

                $filesToProcess = $allSourceFiles.Where(
                    { $_.CreationTime.Date -ge $compareDate }
                )
            }

            $eventLogData.Add(
                [PSCustomObject]@{
                    DateTime  = $scriptStartTime
                    Message   = ("Found {0} file{1} to process" -f $filesToProcess.Count,
                        $(if ($filesToProcess.Count -ne 1) { 's' }))
                    EntryType = 'Information'
                    EventID   = '4'
                }
            )

            Write-Verbose $eventLogData[-1].Message

            if (!$filesToProcess) {
                Write-Verbose "Found $($allSourceFiles.Count) files in source folder, but no file has a creation date in the last  $ProcessFilesCreatedInTheLastNumberOfDays days"
                continue
            }
            #endregion

            #region Copy or move file to destination folder
            foreach ($file in $filesToProcess) {
                try {
                    Write-Verbose "Processing file '$($file.FullName)'"

                    $result = [PSCustomObject]@{
                        DateTime            = Get-Date
                        Action              = $Action
                        SourceFolder        = $SourceFolder
                        DestinationFolder   = $DestinationFolder
                        MatchFileNameRegex  = $MatchFileNameRegex
                        Recurse             = $Recurse
                        OverWriteFile       = $OverWriteFile
                        SourceFilePath      = $file.FullName
                        DestinationFilePath = $null
                        Success             = $false
                        Error               = $null
                    }

                    $params = @{
                        LiteralPath = $file.FullName
                        Destination = "$($DestinationFolder)\$($file.Name)"
                        Force       = $OverWriteFile
                    }

                    $result.DestinationFilePath = $params.Destination

                    Write-Verbose "$Action file '$($params.LiteralPath)' to '$($params.Destination)'"

                    if ($Action -eq 'copy') {
                        Copy-Item @params
                    }
                    else {
                        Move-Item @params
                    }

                    $result.Success = $true
                }
                catch {
                    Write-Warning $_
                    $result.Error = $_
                }
                finally {
                    $logFileData.Add($result)
                }
            }
            #endregion
        }
        catch {
            $systemErrors += [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Action '$Action' SourceFolder '$SourceFolder' DestinationFolder '$DestinationFolder': $_"
            }

            Write-Warning $systemErrors[-1].Message
        }
    }
}

end {
    function Out-LogFileHC {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [PSCustomObject[]]$DataToExport,
            [Parameter(Mandatory)]
            [String]$PartialPath,
            [Parameter(Mandatory)]
            [String[]]$FileExtensions
        )

        $allLogFilePaths = @()

        foreach (
            $fileExtension in
            $FileExtensions | Sort-Object -Unique
        ) {
            $logFilePath = "$PartialPath{0}" -f $fileExtension

            $M = "Export {0} object{1} to '$logFilePath'" -f
            $DataToExport.Count,
            $(if ($DataToExport.Count -ne 1) { 's' })
            Write-Verbose $M

            switch ($fileExtension) {
                '.csv' {
                    $params = @{
                        LiteralPath       = $logFilePath
                        Delimiter         = ';'
                        NoTypeInformation = $true
                    }
                    $DataToExport | Export-Csv @params

                    $allLogFilePaths += $logFilePath
                    break
                }
                '.json' {
                    #region Convert error object to error message string
                    $convertedDataToExport = foreach (
                        $exportObject in
                        $DataToExport
                    ) {
                        foreach ($property in $exportObject.PSObject.Properties) {
                            $name = $property.Name
                            $value = $property.Value
                            if (
                                $value -is [System.Management.Automation.ErrorRecord]
                            ) {
                                if (
                                    $value.Exception -and $value.Exception.Message
                                ) {
                                    $exportObject.$name = $value.Exception.Message
                                }
                                else {
                                    $exportObject.$name = $value.ToString()
                                }
                            }
                        }
                        $exportObject
                    }
                    #endregion

                    $convertedDataToExport |
                    ConvertTo-Json -Depth 7 |
                    Out-File -LiteralPath $logFilePath

                    $allLogFilePaths += $logFilePath
                    break
                }
                '.txt' {
                    $DataToExport |
                    Format-List -Property * -Force |
                    Out-File -LiteralPath $logFilePath

                    $allLogFilePaths += $logFilePath
                    break
                }
                '.xlsx' {
                    $excelParams = @{
                        Path          = $logFilePath
                        AutoNameRange = $true
                        AutoSize      = $true
                        FreezeTopRow  = $true
                        WorksheetName = 'Overview'
                        TableName     = 'Overview'
                        Verbose       = $false
                    }
                    $DataToExport | Export-Excel @excelParams

                    $allLogFilePaths += $logFilePath
                    break
                }
                default {
                    throw "Log file extension '$_' not supported. Supported values are '.xlsx', '.txt' or '.csv'."
                }
            }
        }

        $allLogFilePaths
    }

    function Get-LogFolderHC {
        <#
        .SYNOPSIS
            Ensures that a specified path exists, creating it if it doesn't.
            Supports absolute paths and paths relative to $PSScriptRoot. Returns
            the full path of the folder.

        .DESCRIPTION
            This function takes a path as input and checks if it exists. If
            the path does not exist, it attempts to create the folder. It
            handles both absolute paths and paths relative to the location of
            the currently running script ($PSScriptRoot).

        .PARAMETER Path
            The path to ensure exists. This can be an absolute path (ex.
            C:\MyFolder\SubFolder) or a path relative to the script's
            directory (ex. Data\Logs).

        .EXAMPLE
            Get-LogFolderHC -Path 'C:\MyData\Output'
            # Ensures the directory 'C:\MyData\Output' exists.

        .EXAMPLE
            Get-LogFolderHC -Path 'Logs\Archive'
            # If the script is in 'C:\Scripts', this ensures 'C:\Scripts\Logs\Archive' exists.
        #>

        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        if ($Path -match '^[a-zA-Z]:\\' -or $Path -match '^\\') {
            $fullPath = $Path
        }
        else {
            $fullPath = Join-Path -Path $PSScriptRoot -ChildPath $Path
        }

        if (-not (Test-Path -Path $fullPath -PathType Container)) {
            try {
                Write-Verbose "Create log folder '$fullPath'"
                $null = New-Item -Path $fullPath -ItemType Directory -Force
            }
            catch {
                throw "Failed creating log folder '$fullPath': $_"
            }
        }

        $fullPath
    }

    function Send-MailKitMessageHC {
        <#
            .SYNOPSIS
                Send an email using MailKit and MimeKit assemblies.

            .DESCRIPTION
                This function sends an email using the MailKit and MimeKit
                assemblies. It requires the assemblies to be installed before
                calling the function:

                $params = @{
                    Source           = 'https://www.nuget.org/api/v2'
                    SkipDependencies = $true
                    Scope            = 'AllUsers'
                }
                Install-Package @params -Name 'MailKit'
                Install-Package @params -Name 'MimeKit'

            .PARAMETER MailKitAssemblyPath
                The path to the MailKit assembly.

            .PARAMETER MimeKitAssemblyPath
                The path to the MimeKit assembly.

            .PARAMETER SmtpServerName
                The name of the SMTP server.

            .PARAMETER SmtpPort
                The port of the SMTP server.

            .PARAMETER SmtpConnectionType
                The connection type for the SMTP server.

                Valid values are:
                - 'None'
                - 'Auto'
                - 'SslOnConnect'
                - 'StartTlsWhenAvailable'
                - 'StartTls'

            .PARAMETER Credential
                The credential object containing the username and password.

            .PARAMETER From
                The sender's email address.

            .PARAMETER FromDisplayName
                The display name to show for the sender.

                Email clients may display this differently. It is most likely
                to be shown if the sender's email address is not recognized
                (e.g., not in the address book).

            .PARAMETER To
                The recipient's email address.

            .PARAMETER Body
                The body of the email, HTML is supported.

            .PARAMETER Subject
                The subject of the email.

            .PARAMETER Attachments
                An array of file paths to attach to the email.

            .PARAMETER Priority
                The email priority.

                Valid values are:
                - 'Low'
                - 'Normal'
                - 'High'

            .EXAMPLE
                # Send an email with StartTls and credential

                $SmtpUserName = 'smtpUser'
                $SmtpPassword = 'smtpPassword'

                $securePassword = ConvertTo-SecureString -String $SmtpPassword -AsPlainText -Force
                $credential = New-Object System.Management.Automation.PSCredential($SmtpUserName, $securePassword)

                $params = @{
                    SmtpServerName      = 'SMT_SERVER@example.com'
                    SmtpPort            = 587
                    SmtpConnectionType  = 'StartTls'
                    Credential          = $credential
                    From                = 'm@example.com'
                    To                  = '007@example.com'
                    Body                = '<p>Mission details in attachment</p>'
                    Subject             = 'For your eyes only'
                    Priority            = 'High'
                    Attachments         = @('c:\Mission.ppt', 'c:\ID.pdf')
                    MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                    MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
                }

                Send-MailKitMessageHC @params

            .EXAMPLE
                # Send an email without authentication

                $params = @{
                    SmtpServerName      = 'SMT_SERVER@example.com'
                    SmtpPort            = 25
                    From                = 'hacker@example.com'
                    FromDisplayName     = 'White hat hacker'
                    Bcc                 = @('james@example.com', 'mike@example.com')
                    Body                = '<h1>You have been hacked</h1>'
                    Subject             = 'Oops'
                    MailKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                    MimeKitAssemblyPath = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
                }

                Send-MailKitMessageHC @params
        #>

        [CmdletBinding()]
        param (
            [parameter(Mandatory)]
            [string]$MailKitAssemblyPath,
            [parameter(Mandatory)]
            [string]$MimeKitAssemblyPath,
            [parameter(Mandatory)]
            [string]$SmtpServerName,
            [parameter(Mandatory)]
            [ValidateSet(25, 465, 587, 2525)]
            [int]$SmtpPort,
            [parameter(Mandatory)]
            [string]$Body,
            [parameter(Mandatory)]
            [string]$Subject,
            [parameter(Mandatory)]
            [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
            [string]$From,
            [string]$FromDisplayName,
            [string[]]$To,
            [string[]]$Bcc,
            [int]$MaxAttachmentSize = 20MB,
            [ValidateSet(
                'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable'
            )]
            [string]$SmtpConnectionType = 'None',
            [ValidateSet('Normal', 'Low', 'High')]
            [string]$Priority = 'Normal',
            [string[]]$Attachments,
            [PSCredential]$Credential
        )

        begin {
            function Test-IsAssemblyLoaded {
                param (
                    [String]$Name
                )
                foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
                    if ($assembly.FullName -like "$Name, Version=*") {
                        return $true
                    }
                }
                return $false
            }

            function Add-Attachments {
                param (
                    [string[]]$Attachments,
                    [MimeKit.Multipart]$BodyMultiPart
                )

                $attachmentList = New-Object System.Collections.ArrayList($null)

                $tempFolder = "$env:TEMP\Send-MailKitMessageHC {0}" -f (Get-Random)
                $totalSizeAttachments = 0

                foreach (
                    $attachmentPath in
                    $Attachments | Sort-Object -Unique
                ) {
                    try {
                        #region Test if file exists
                        try {
                            $attachmentItem = Get-Item -LiteralPath $attachmentPath -ErrorAction Stop

                            if ($attachmentItem.PSIsContainer) {
                                Write-Warning "Attachment '$attachmentPath' is a folder, not a file"
                                continue
                            }
                        }
                        catch {
                            Write-Warning "Attachment '$attachmentPath' not found"
                            continue
                        }
                        #endregion

                        $totalSizeAttachments += $attachmentItem.Length

                        if ($attachmentItem.Extension -eq '.xlsx') {
                            #region Copy Excel file, open file cannot be sent
                            if (-not(Test-Path $tempFolder)) {
                                $null = New-Item $tempFolder -ItemType 'Directory'
                            }

                            $params = @{
                                LiteralPath = $attachmentItem.FullName
                                Destination = $tempFolder
                                PassThru    = $true
                                ErrorAction = 'Stop'
                            }

                            $copiedItem = Copy-Item @params

                            $null = $attachmentList.Add($copiedItem)
                            #endregion
                        }
                        else {
                            $null = $attachmentList.Add($attachmentItem)
                        }

                        #region Check size of attachments
                        if ($totalSizeAttachments -ge $MaxAttachmentSize) {
                            $M = "The maximum allowed attachment size of {0} MB has been exceeded ({1} MB). No attachments were added to the email. Check the log folder for details." -f
                            ([math]::Round(($MaxAttachmentSize / 1MB))),
                            ([math]::Round(($totalSizeAttachments / 1MB), 2))

                            Write-Warning $M

                            return [PSCustomObject]@{
                                AttachmentLimitExceededMessage = $M
                            }
                        }
                    }
                    catch {
                        Write-Warning "Failed to add attachment '$attachmentPath': $_"
                    }
                }
                #endregion

                foreach (
                    $attachmentItem in
                    $attachmentList
                ) {
                    try {
                        Write-Verbose "Add mail attachment '$($attachmentItem.Name)'"

                        $attachment = New-Object MimeKit.MimePart

                        $attachment.Content = New-Object MimeKit.MimeContent(
                            [System.IO.File]::OpenRead($attachmentItem.FullName)
                        )

                        $attachment.ContentDisposition = New-Object MimeKit.ContentDisposition

                        $attachment.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64

                        $attachment.FileName = $attachmentItem.Name

                        $bodyMultiPart.Add($attachment)
                    }
                    catch {
                        Write-Warning "Failed to add attachment '$attachmentItem': $_"
                    }
                }
            }

            try {
                #region Test To or Bcc required
                if (-not ($To -or $Bcc)) {
                    throw "Either 'To' to 'Bcc' is required for sending emails"
                }
                #endregion

                #region Test To
                foreach ($email in $To) {
                    if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                        throw "To email address '$email' not valid."
                    }
                }
                #endregion

                #region Test Bcc
                foreach ($email in $Bcc) {
                    if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                        throw "Bcc email address '$email' not valid."
                    }
                }
                #endregion

                #region Load MimeKit assembly
                if (-not(Test-IsAssemblyLoaded -Name 'MimeKit')) {
                    try {
                        Write-Verbose "Load MimeKit assembly '$MimeKitAssemblyPath'"
                        Add-Type -Path $MimeKitAssemblyPath
                    }
                    catch {
                        throw "Failed to load MimeKit assembly '$MimeKitAssemblyPath': $_"
                    }
                }
                #endregion

                #region Load MailKit assembly
                if (-not(Test-IsAssemblyLoaded -Name 'MailKit')) {
                    try {
                        Write-Verbose "Load MailKit assembly '$MailKitAssemblyPath'"
                        Add-Type -Path $MailKitAssemblyPath
                    }
                    catch {
                        throw "Failed to load MailKit assembly '$MailKitAssemblyPath': $_"
                    }
                }
                #endregion
            }
            catch {
                throw "Failed to send email to '$To': $_"
            }
        }

        process {
            try {
                $message = New-Object -TypeName 'MimeKit.MimeMessage'

                #region Create body with attachments
                $bodyPart = New-Object MimeKit.TextPart('html')
                $bodyPart.Text = $Body

                $bodyMultiPart = New-Object MimeKit.Multipart('mixed')
                $bodyMultiPart.Add($bodyPart)

                if ($Attachments) {
                    $params = @{
                        Attachments   = $Attachments
                        BodyMultiPart = $bodyMultiPart
                    }
                    $addAttachments = Add-Attachments @params

                    if ($addAttachments.AttachmentLimitExceededMessage) {
                        $bodyPart.Text += '<p><i>{0}</i></p>' -f
                        $addAttachments.AttachmentLimitExceededMessage
                    }
                }

                $message.Body = $bodyMultiPart
                #endregion

                $fromAddress = New-Object MimeKit.MailboxAddress(
                    $FromDisplayName, $From
                )
                $message.From.Add($fromAddress)

                foreach ($email in $To) {
                    $message.To.Add($email)
                }

                foreach ($email in $Bcc) {
                    $message.Bcc.Add($email)
                }

                $message.Subject = $Subject

                #region Set priority
                switch ($Priority) {
                    'Low' {
                        $message.Headers.Add('X-Priority', '5 (Lowest)')
                        break
                    }
                    'Normal' {
                        $message.Headers.Add('X-Priority', '3 (Normal)')
                        break
                    }
                    'High' {
                        $message.Headers.Add('X-Priority', '1 (Highest)')
                        break
                    }
                    default {
                        throw "Priority type '$_' not supported"
                    }
                }
                #endregion

                $smtp = New-Object -TypeName 'MailKit.Net.Smtp.SmtpClient'

                try {
                    $smtp.Connect(
                        $SmtpServerName, $SmtpPort,
                        [MailKit.Security.SecureSocketOptions]::$SmtpConnectionType
                    )
                }
                catch {
                    throw "Failed to connect to SMTP server '$SmtpServerName' on port '$SmtpPort' with connection type '$SmtpConnectionType': $_"
                }

                if ($Credential) {
                    try {
                        $smtp.Authenticate(
                            $Credential.UserName,
                            $Credential.GetNetworkCredential().Password
                        )
                    }
                    catch {
                        throw "Failed to authenticate with user name '$($Credential.UserName)' to SMTP server '$SmtpServerName': $_"
                    }
                }

                Write-Verbose "Send mail to '$To' with subject '$Subject'"

                $null = $smtp.Send($message)
                $smtp.Disconnect($true)
                $smtp.Dispose()
            }
            catch {
                throw "Failed to send email to '$To': $_"
            }
        }
    }

    function Write-EventsToEventLogHC {
        <#
        .SYNOPSIS
            Write events to the event log.

        .DESCRIPTION
            The use of this function will allow standardization in the Windows
            Event Log by using the same EventID's and other properties across
            different scripts.

            Custom Windows EventID's based on the PowerShell standard streams:

            PowerShell Stream     EventIcon    EventID   EventDescription
            -----------------     ---------    -------   ----------------
            [i] Info              [i] Info     100       Script started
            [4] Verbose           [i] Info     4         Verbose message
            [1] Output/Success    [i] Info     1         Output on success
            [3] Warning           [w] Warning  3         Warning message
            [2] Error             [e] Error    2         Fatal error message
            [i] Info              [i] Info     199       Script ended successfully

        .PARAMETER Source
            Specifies the script name under which the events will be logged.

        .PARAMETER LogName
            Specifies the name of the event log to which the events will be
            written. If the log does not exist, it will be created.

        .PARAMETER Events
            Specifies the events to be written to the event log. This should be
            an array of PSCustomObject with properties: DateTime, Message,
            EntryType, and EventID.
        #>

        [CmdLetBinding()]
        param (
            [Parameter(Mandatory)]
            [String]$Source,
            [Parameter(Mandatory)]
            [String]$LogName,
            [PSCustomObject[]]$Events
        )

        try {
            if (
                -not(
                    ([System.Diagnostics.EventLog]::Exists($LogName)) -and
                    [System.Diagnostics.EventLog]::SourceExists($Source)
                )
            ) {
                Write-Verbose "Create event log '$LogName' and source '$Source'"
                New-EventLog -LogName $LogName -Source $Source -ErrorAction Stop
            }

            foreach ($eventItem in $Events) {
                $params = @{
                    LogName     = $LogName
                    Source      = $Source
                    EntryType   = $eventItem.EntryType
                    EventID     = $eventItem.EventID
                    Message     = '{0}: {1}' -f $eventItem.DateTime, $eventItem.Message
                    ErrorAction = 'Stop'
                }

                Write-Verbose "Write event to log '$LogName' source '$Source' message '$($params.Message)'"

                Write-EventLog @params
            }
        }
        catch {
            throw "Failed to write to event log '$LogName' source '$Source': $_"
        }
    }

    function Get-StringValueHC {
        <#
        .SYNOPSIS
            Retrieve a string from the environment variables or a regular string.

        .DESCRIPTION
            This function checks the 'Name' property. If the value starts with
            'ENV:', it attempts to retrieve the string value from the specified
            environment variable. Otherwise, it returns the value directly.

        .PARAMETER Name
            Either a string starting with 'ENV:'; a plain text string or NULL.

        .EXAMPLE
            Get-StringValueHC -Name 'ENV:passwordVariable'

            # Output: the environment variable value of $ENV:passwordVariable
            # or an error when the variable does not exist

        .EXAMPLE
            Get-StringValueHC -Name 'mySecretPassword'

            # Output: mySecretPassword

        .EXAMPLE
            Get-StringValueHC -Name ''

            # Output: NULL
        #>
        param (
            [String]$Name
        )

        if (-not $Name) {
            return $null
        }
        elseif (
            $Name.StartsWith('ENV:', [System.StringComparison]::OrdinalIgnoreCase)
        ) {
            $envVariableName = $Name.Substring(4).Trim()
            $envStringValue = Get-Item -Path "Env:\$envVariableName" -EA Ignore
            if ($envStringValue) {
                return $envStringValue.Value
            }
            else {
                throw "Environment variable '$envVariableName' not found."
            }
        }
        else {
            return $Name
        }
    }

    try {
        $settings = $jsonFileContent.Settings

        $scriptName = $settings.ScriptName
        $saveInEventLog = $settings.SaveInEventLog
        $sendMail = $settings.SendMail
        $saveLogFiles = $settings.SaveLogFiles

        $allLogFilePaths = @()
        $logFileDataErrors = $logFileData.Where({ $_.Error })
        $baseLogName = $null
        $logFolderPath = $null

        #region Get script name
        if (-not $scriptName) {
            Write-Warning "No 'Settings.ScriptName' found in import file."
            $scriptName = 'Default script name'
        }
        #endregion

        #region Create log files
        try {
            $logFolder = Get-StringValueHC $saveLogFiles.Where.Folder

            $logFileExtensions = $saveLogFiles.Where.FileExtensions
            $isLog = @{
                systemErrors     = $saveLogFiles.What.SystemErrors
                allActions       = $saveLogFiles.What.AllActions
                onlyActionErrors = $saveLogFiles.What.OnlyActionErrors
            }

            if ($logFolder -and $logFileExtensions) {
                #region Get log folder
                try {
                    $logFolderPath = Get-LogFolderHC -Path $logFolder

                    Write-Verbose "Log folder '$logFolderPath'"

                    $baseLogName = Join-Path -Path $logFolderPath -ChildPath (
                        '{0} - {1} ({2})' -f
                        $scriptStartTime.ToString('yyyy_MM_dd_HHmmss_dddd'),
                        $ScriptName,
                        $jsonFileItem.BaseName
                    )
                }
                catch {
                    throw "Failed creating log folder '$LogFolder': $_"
                }
                #endregion

                #region Create log file
                if ($logFileData) {
                    if ($isLog.allActions) {
                        $params = @{
                            DataToExport   = $logFileData
                            PartialPath    = if ($logFileDataErrors) {
                                "$baseLogName - Actions with errors"
                            }
                            else {
                                "$baseLogName - Actions"
                            }
                            FileExtensions = $logFileExtensions
                        }
                        $allLogFilePaths += Out-LogFileHC @params
                    }
                    elseif ($isLog.onlyActionErrors) {
                        if ($logFileDataErrors) {
                            $params = @{
                                DataToExport   = $logFileDataErrors
                                PartialPath    = "$baseLogName - Action errors"
                                FileExtensions = $logFileExtensions
                            }
                            $allLogFilePaths += Out-LogFileHC @params
                        }
                    }
                }

                if ($systemErrors -and $isLog.SystemErrors) {
                    $params = @{
                        DataToExport   = $systemErrors
                        PartialPath    = "$baseLogName - Errors"
                        FileExtensions = $logFileExtensions
                    }
                    $allLogFilePaths += Out-LogFileHC @params
                }
                #endregion
            }
        }
        catch {
            $systemErrors += [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Failed creating log file in folder '$($saveLogFiles.Where.Folder)': $_"
            }

            Write-Warning $systemErrors[-1].Message
        }
        #endregion

        #region Remove old log files
        if ($saveLogFiles.DeleteLogsAfterDays -gt 0 -and $logFolderPath) {
            $cutoffDate = (Get-Date).AddDays(-$saveLogFiles.DeleteLogsAfterDays)

            Write-Verbose "Removing log files older than $cutoffDate from '$logFolderPath'"

            Get-ChildItem -Path $logFolderPath -File |
            Where-Object { $_.LastWriteTime -lt $cutoffDate } |
            ForEach-Object {
                try {
                    $fileToRemove = $_
                    Write-Verbose "Deleting old log file '$_''"
                    Remove-Item -Path $_.FullName -Force
                }
                catch {
                    $systemErrors += [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = "Failed to remove file '$fileToRemove': $_"
                    }

                    Write-Warning $systemErrors[-1].Message

                    if ($baseLogName -and $isLog.systemErrors) {
                        $params = @{
                            DataToExport   = $systemErrors[-1]
                            PartialPath    = "$baseLogName - Errors"
                            FileExtensions = $logFileExtensions
                        }
                        $allLogFilePaths += Out-LogFileHC @params -EA Ignore
                    }
                }
            }
        }
        #endregion

        #region Write events to event log
        try {
            $saveInEventLog.LogName = Get-StringValueHC $saveInEventLog.LogName

            if ($saveInEventLog.Save -and $saveInEventLog.LogName) {
                $systemErrors | ForEach-Object {
                    $eventLogData.Add(
                        [PSCustomObject]@{
                            DateTime  = $_.DateTime
                            Message   = $_.Message
                            EntryType = 'Error'
                            EventID   = '2'
                        }
                    )
                }

                $eventLogData.Add(
                    [PSCustomObject]@{
                        DateTime  = Get-Date
                        Message   = 'Script ended'
                        EntryType = 'Information'
                        EventID   = '199'
                    }
                )

                $params = @{
                    Source  = $scriptName
                    LogName = $saveInEventLog.LogName
                    Events  = $eventLogData
                }
                Write-EventsToEventLogHC @params

            }
            elseif ($saveInEventLog.Save -or $saveInEventLog.LogName) {
                throw "Both 'Settings.SaveInEventLog.Save' and 'Settings.SaveInEventLog.LogName' are required to save events in the event log."
            }
        }
        catch {
            $systemErrors += [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Failed writing events to event log: $_"
            }

            Write-Warning $systemErrors[-1].Message

            if ($baseLogName -and $isLog.systemErrors) {
                $params = @{
                    DataToExport   = $systemErrors[-1]
                    PartialPath    = "$baseLogName - Errors"
                    FileExtensions = $logFileExtensions
                }
                $allLogFilePaths += Out-LogFileHC @params -EA Ignore
            }
        }
        #endregion

        #region Send email
        try {
            $isSendMail = $false

            switch ($sendMail.When) {
                'Never' {
                    break
                }
                'Always' {
                    $isSendMail = $true
                    break
                }
                'OnError' {
                    if ($systemErrors -or $logFileDataErrors) {
                        $isSendMail = $true
                    }
                    break
                }
                'OnErrorOrAction' {
                    if ($systemErrors -or $logFileDataErrors -or $logFileData) {
                        $isSendMail = $true
                    }
                    break
                }
                default {
                    throw "SendMail.When '$($sendMail.When)' not supported. Supported values are 'Never', 'Always', 'OnError' or 'OnErrorOrAction'."
                }
            }

            if ($isSendMail) {
                #region Test mandatory fields
                @{
                    'From'                 = $sendMail.From
                    'Smtp.ServerName'      = $sendMail.Smtp.ServerName
                    'Smtp.Port'            = $sendMail.Smtp.Port
                    'AssemblyPath.MailKit' = $sendMail.AssemblyPath.MailKit
                    'AssemblyPath.MimeKit' = $sendMail.AssemblyPath.MimeKit
                }.GetEnumerator() |
                Where-Object { -not $_.Value } | ForEach-Object {
                    throw "Input file property 'Settings.SendMail.$($_.Key)' cannot be blank"
                }
                #endregion

                $mailParams = @{
                    From                = Get-StringValueHC $sendMail.From
                    Subject             = '{0} action{1}' -f
                    $logFileData.Count,
                    $(if ($logFileData.Count -ne 1) { 's' })
                    SmtpServerName      = Get-StringValueHC $sendMail.Smtp.ServerName
                    SmtpPort            = Get-StringValueHC $sendMail.Smtp.Port
                    MailKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MailKit
                    MimeKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MimeKit
                }

                $mailParams.Body = @"
<!DOCTYPE html>
<html>
    <head>
        <style type="text/css">
            body {
                font-family:verdana;
                font-size:14px;
                background-color:white;
            }
            h1 {
                margin-bottom: 0;
            }
            h2 {
                margin-bottom: 0;
            }
            h3 {
                margin-bottom: 0;
            }
            p.italic {
                font-style: italic;
                font-size: 12px;
            }
            table {
                border-collapse:collapse;
                border:0px none;
                padding:3px;
                text-align:left;
            }
            td, th {
                border-collapse:collapse;
                border:1px none;
                padding:3px;
                text-align:left;
            }
            #aboutTable th {
                color: rgb(143, 140, 140);
                font-weight: normal;
            }
            #aboutTable td {
                color: rgb(143, 140, 140);
                font-weight: normal;
            }
            base {
                target="_blank"
            }
        </style>
    </head>
    <body>
        <table>
            <h1>$scriptName</h1>
            <hr size="2" color="#06cc7a">

            $($sendMail.Body)

            <table>
                <tr>
                    <th>Actions</th>
                    <td>$($logFileData.Count)</td>
                </tr>
                $(
                    if($logFileDataErrors.Count) {
                        '<tr style="background-color: #f78474;">'
                    } else {
                        '<tr>'
                    }
                )
                    <th>Action errors</th>
                    <td>$($logFileDataErrors.Count)</td>
                </tr>
                $(
                    if($systemErrors.Count) {
                        '<tr style="background-color: #f78474;">'
                    } else {
                        '<tr>'
                    }
                )
                    <th>System errors</th>
                    <td>$($systemErrors.Count)</td>
                </tr>
            </table>

            $(
                if ($allLogFilePaths) {
                    '<p><i>* Check the attachment(s) for details</i></p>'
                }
            )

            <hr size="2" color="#06cc7a">
            <table id="aboutTable">
                $(
                    if ($scriptStartTime) {
                        '<tr>
                            <th>Start time</th>
                            <td>{0:00}/{1:00}/{2:00} {3:00}:{4:00} ({5})</td>
                        </tr>' -f
                        $scriptStartTime.Day,
                        $scriptStartTime.Month,
                        $scriptStartTime.Year,
                        $scriptStartTime.Hour,
                        $scriptStartTime.Minute,
                        $scriptStartTime.DayOfWeek
                    }
                )
                $(
                    if ($scriptStartTime) {
                        $runTime = New-TimeSpan -Start $scriptStartTime -End (Get-Date)
                        '<tr>
                            <th>Duration</th>
                            <td>{0:00}:{1:00}:{2:00}</td>
                        </tr>' -f
                        $runTime.Hours, $runTime.Minutes, $runTime.Seconds
                    }
                )
                $(
                    if ($logFolderPath) {
                        '<tr>
                            <th>Log files</th>
                            <td><a href="{0}">Open log folder</a></td>
                        </tr>' -f $logFolderPath
                    }
                )
                <tr>
                    <th>Host</th>
                    <td>$($host.Name)</td>
                </tr>
                <tr>
                    <th>PowerShell</th>
                    <td>$($PSVersionTable.PSVersion.ToString())</td>
                </tr>
                <tr>
                    <th>Computer</th>
                    <td>$env:COMPUTERNAME</td>
                </tr>
                <tr>
                    <th>Account</th>
                    <td>$env:USERDNSDOMAIN\$env:USERNAME</td>
                </tr>
            </table>
        </table>
    </body>
</html>
"@

                if ($sendMail.FromDisplayName) {
                    $mailParams.FromDisplayName = $sendMail.FromDisplayName
                }

                if ($sendMail.Subject) {
                    $mailParams.Subject = '{0}, {1}' -f
                    $mailParams.Subject, $sendMail.Subject
                }

                if ($sendMail.To) {
                    $mailParams.To = $sendMail.To
                }

                if ($sendMail.Bcc) {
                    $mailParams.Bcc = $sendMail.Bcc
                }

                if ($systemErrors -or $logFileDataErrors) {
                    $totalErrorCount = $systemErrors.Count + $logFileDataErrors.Count

                    $mailParams.Priority = 'High'
                    $mailParams.Subject = '{0} error{1}, {2}' -f
                    $totalErrorCount,
                    $(if ($totalErrorCount -ne 1) { 's' }),
                    $mailParams.Subject
                }

                if ($allLogFilePaths) {
                    $mailParams.Attachments = $allLogFilePaths |
                    Sort-Object -Unique
                }

                if ($sendMail.Smtp.ConnectionType) {
                    $mailParams.SmtpConnectionType = Get-StringValueHC $sendMail.Smtp.ConnectionType
                }

                #region Create SMTP credential
                $smtpUserName = Get-StringValueHC $sendMail.Smtp.UserName
                $smtpPassword = Get-StringValueHC $sendMail.Smtp.Password

                if ( $smtpUserName -and $smtpPassword) {
                    try {
                        $securePassword = ConvertTo-SecureString -String $smtpPassword -AsPlainText -Force

                        $credential = New-Object System.Management.Automation.PSCredential($smtpUserName, $securePassword)

                        $mailParams.Credential = $credential
                    }
                    catch {
                        throw "Failed to create credential: $_"
                    }
                }
                elseif ($smtpUserName -or $smtpPassword) {
                    throw "Both 'Settings.SendMail.Smtp.Username' and 'Settings.SendMail.Smtp.Password' are required when authentication is needed."
                }
                #endregion

                Write-Verbose "Send email to '$($mailParams.To)' subject '$($mailParams.Subject)'"

                Send-MailKitMessageHC @mailParams
            }
        }
        catch {
            $systemErrors += [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Failed sending email: $_"
            }

            Write-Warning $systemErrors[-1].Message

            if ($baseLogName -and $isLog.systemErrors) {
                $params = @{
                    DataToExport   = $systemErrors[-1]
                    PartialPath    = "$baseLogName - Errors"
                    FileExtensions = $logFileExtensions
                }
                $null = Out-LogFileHC @params -EA Ignore
            }
        }
        #endregion
    }
    catch {
        $systemErrors += [PSCustomObject]@{
            DateTime = Get-Date
            Message  = $_
        }

        Write-Warning $systemErrors[-1].Message
    }
    finally {
        if ($systemErrors) {
            $M = 'Found {0} system error{1}' -f
            $systemErrors.Count,
            $(if ($systemErrors.Count -ne 1) { 's' })
            Write-Warning $M

            $systemErrors | ForEach-Object {
                Write-Warning $_.Message
            }

            Write-Warning 'Exit script with error code 1'
            exit 1
        }
        else {
            Write-Verbose 'Script finished successfully'
        }
    }
}