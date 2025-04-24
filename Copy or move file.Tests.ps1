#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $realCmdLet = @{
        OutFile = Get-Command Out-File
    }

    $testInputFile = @{
        Tasks    = @(
            @{
                Action                                   = 'copy'
                Source                                   = @{
                    Folder             = (New-Item 'TestDrive:/s' -ItemType Directory).FullName
                    MatchFileNameRegex = 'Analyse_[0-9]{8}.xlsx'
                    Recurse            = $true
                }
                Destination                              = @{
                    Folder        = (New-Item 'TestDrive:/d' -ItemType Directory).FullName
                    OverWriteFile = $false
                }
                ProcessFilesCreatedInTheLastNumberOfDays = 1
            }
        )
        Settings = @{
            ScriptName     = 'Test (Brecht)'
            SendMail       = @{
                When         = 'Always'
                From         = 'm@example.com'
                To           = '007@example.com'
                Subject      = 'Email subject'
                Body         = 'Email body'
                Smtp         = @{
                    ServerName     = 'SMTP_SERVER'
                    Port           = 25
                    ConnectionType = 'StartTls'
                    UserName       = 'bob'
                    Password       = 'pass'
                }
                AssemblyPath = @{
                    MailKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                    MimeKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
                }
            }
            SaveLogFiles   = @{
                What                = @{
                    SystemErrors     = $true
                    AllActions       = $true
                    OnlyActionErrors = $false
                }
                Where               = @{
                    Folder         = (New-Item 'TestDrive:/log' -ItemType Directory).FullName
                    FileExtensions = @('.json', '.csv')
                }
                deleteLogsAfterDays = 1
            }
            SaveInEventLog = @{
                Save    = $true
                LogName = 'Scripts'
            }
        }
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ConfigurationJsonFile = $testOutParams.FilePath
    }

    function Copy-ObjectHC {
        <#
        .SYNOPSIS
            Make a deep copy of an object using JSON serialization.

        .DESCRIPTION
            Uses ConvertTo-Json and ConvertFrom-Json to create an independent
            copy of an object. This method is generally effective for objects
            that can be represented in JSON format.

        .PARAMETER InputObject
            The object to copy.

        .EXAMPLE
            $newArray = Copy-ObjectHC -InputObject $originalArray
        #>
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [Object]$InputObject
        )

        $jsonString = $InputObject | ConvertTo-Json -Depth 100

        $deepCopy = $jsonString | ConvertFrom-Json

        return $deepCopy
    }
    function Send-MailKitMessageHC {
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
            [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
            [string]$From,
            [parameter(Mandatory)]
            [string]$Body,
            [parameter(Mandatory)]
            [string]$Subject,
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
    }

    Mock Send-MailKitMessageHC
    Mock New-EventLog
    Mock Write-EventLog
    Mock Out-File
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ConfigurationJsonFile') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'create an error log file when' {
    It 'the log folder cannot be created' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Settings.SaveLogFiles.Where.Folder = 'x:\notExistingLocation'

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams

        $LASTEXITCODE | Should -Be 1

        Should -Not -Invoke Out-File
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ConfigurationJsonFile = 'nonExisting.json'

            .$testScript @testNewParams

            $LASTEXITCODE | Should -Be 1

            Should -Not -Invoke Out-File
        }
        Context 'property' {
            It 'Tasks.<_> not found' -ForEach @(
                'Action', 'Source', 'Destination'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].$_ = $null

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                $LASTEXITCODE | Should -Be 1

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($LiteralPath -like '* - Errors.json') -and
                    ($InputObject -like "*Property 'Tasks.$_' not found*")
                }
            }
            It 'Tasks.Source.<_> not found' -ForEach @(
                'Folder', 'MatchFileNameRegex'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Source.$_ = $null

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                $LASTEXITCODE | Should -Be 1

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($LiteralPath -like '* - Errors.json') -and
                    ($InputObject -like "*Property 'Tasks.Source.$_' not found*")
                }
            }
            It 'Tasks.Destination.<_> not found' -ForEach @(
                'Folder'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Destination.$_ = $null

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                $LASTEXITCODE | Should -Be 1

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($LiteralPath -like '* - Errors.json') -and
                    ($InputObject -like "*Property 'Tasks.Destination.$_' not found*")
                }
            }
            Context 'Tasks.ProcessFilesCreatedInTheLastNumberOfDays' {
                It 'is not a number' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].ProcessFilesCreatedInTheLastNumberOfDays = 'a'

                    & $realCmdLet.OutFile @testOutParams -InputObject (
                        $testNewInputFile | ConvertTo-Json -Depth 7
                    )

                    .$testScript @testParams

                    Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                        ($LiteralPath -like '* - Errors.json') -and
                        ($InputObject -like "*$($testParams.ConfigurationJsonFile.replace('\','\\'))*Property 'Tasks.ProcessFilesCreatedInTheLastNumberOfDays' must be 0 or a positive number. Number 0 processes all files in the source folder. The value 'a' is not supported*")
                    }
                }
                It 'is a negative number' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].ProcessFilesCreatedInTheLastNumberOfDays = -1

                    & $realCmdLet.OutFile @testOutParams -InputObject (
                        $testNewInputFile | ConvertTo-Json -Depth 7
                    )

                    .$testScript @testParams

                    Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                        ($LiteralPath -like '* - Errors.json') -and
                        ($InputObject -like "*$($testParams.ConfigurationJsonFile.replace('\','\\'))*Property 'Tasks.ProcessFilesCreatedInTheLastNumberOfDays' must be 0 or a positive number. Number 0 processes all files in the source folder. The value '-1' is not supported*")
                    }
                }
                It 'is an empty string' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].ProcessFilesCreatedInTheLastNumberOfDays = ''

                    & $realCmdLet.OutFile @testOutParams -InputObject (
                        $testNewInputFile | ConvertTo-Json -Depth 7
                    )

                    .$testScript @testParams

                    Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                        ($LiteralPath -like '* - Errors.json') -and
                        ($InputObject -like "*$($testParams.ConfigurationJsonFile.replace('\','\\'))*Property 'Tasks.ProcessFilesCreatedInTheLastNumberOfDays' must be 0 or a positive number. Number 0 processes all files in the source folder. The value '' is not supported*")
                    }
                }
                It 'is missing' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].ProcessFilesCreatedInTheLastNumberOfDays = $null

                    & $realCmdLet.OutFile @testOutParams -InputObject (
                        $testNewInputFile | ConvertTo-Json -Depth 7
                    )

                    .$testScript @testParams

                    Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                        ($LiteralPath -like '* - Errors.json') -and
                        ($InputObject -like "*$($testParams.ConfigurationJsonFile.replace('\','\\'))*Property 'Tasks.ProcessFilesCreatedInTheLastNumberOfDays' must be 0 or a positive number. Number 0 processes all files in the source folder. The value '' is not supported*")
                    }
                }
                It '0 is accepted' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].ProcessFilesCreatedInTheLastNumberOfDays = '0'

                    & $realCmdLet.OutFile @testOutParams -InputObject (
                        $testNewInputFile | ConvertTo-Json -Depth 7
                    )

                    .$testScript @testParams

                    Should -Invoke -Not Out-File -ParameterFilter {
                        ($LiteralPath -like '* - Errors.json') -and
                        ($InputObject -like "*ProcessFilesCreatedInTheLastNumberOfDays*")
                    }
                }
            }
            It "Tasks.Action is not value 'copy' or 'move'" {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Action = 'wrong'

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($LiteralPath -like '* - Errors.json') -and
                    ($InputObject -like "*$($testParams.ConfigurationJsonFile.replace('\','\\'))*'Tasks.Action' value 'wrong' is not supported. Supported Action values are: 'copy' or 'move'.*")
                }
            }
            It "Tasks.Source.Recurse is not a boolean" {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Source.Recurse = 'wrong'

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($LiteralPath -like '* - Errors.json') -and
                    ($InputObject -like "*$($testParams.ConfigurationJsonFile.replace('\','\\'))*Property 'Tasks.Source.Recurse' is not a boolean value*")
                }
            }
            It "Tasks.Destination.OverWriteFile is not a boolean" {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Destination.OverWriteFile = 'wrong'

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($LiteralPath -like '* - Errors.json') -and
                    ($InputObject -like "*$($testParams.ConfigurationJsonFile.replace('\','\\'))*Property 'Tasks.Destination.OverWriteFile' is not a boolean value*")
                }
            }
        }
    }
}
Describe 'when the source folder is empty' {
    It 'no error log file is created' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Source.Folder = (New-Item 'TestDrive:/empty' -ItemType Directory).FullName

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams

        Should -Not -Invoke Out-File
    }
}
Describe 'when there is a file in the source folder' {
    Context 'and Action is copy' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile

            $testNewInputFile.Tasks[0].Action = 'copy'

            $testNewInputFile.Tasks[0].Source.Folder = (New-Item 'TestDrive:/source' -ItemType Directory).FullName
            $testNewInputFile.Tasks[0].Destination.Folder = (New-Item 'TestDrive:/destination' -ItemType Directory).FullName

            $testSourceFile = New-Item "$($testNewInputFile.Tasks[0].Source.Folder)\Analyse_26032025.xlsx" -ItemType File

            & $realCmdLet.OutFile @testOutParams -InputObject (
                $testNewInputFile | ConvertTo-Json -Depth 7
            )

            .$testScript @testParams
        }
        It 'the file is copied to the destination folder' {
            "$($testNewInputFile.Tasks[0].Destination.Folder)\Analyse_26032025.xlsx" |
            Should -Exist
        }
        It 'the source file is left untouched' {
            $testSourceFile | Should -Exist
        }
    }
    Context 'and Action is move' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile

            $testNewInputFile.Tasks[0].Action = 'move'

            $testNewInputFile.Tasks[0].Source.Folder = (New-Item 'TestDrive:/source' -ItemType Directory).FullName
            $testNewInputFile.Tasks[0].Destination.Folder = (New-Item 'TestDrive:/destination' -ItemType Directory).FullName

            $testSourceFile = New-Item "$($testNewInputFile.Tasks[0].Source.Folder)\Analyse_26032025.xlsx" -ItemType File

            & $realCmdLet.OutFile @testOutParams -InputObject (
                $testNewInputFile | ConvertTo-Json -Depth 7
            )

            .$testScript @testParams
        }
        It 'the file is present in the destination folder' {
            "$($testNewInputFile.Tasks[0].Destination.Folder)\Analyse_26032025.xlsx" |
            Should -Exist
        }
        It 'the source file is no longer there' {
            $testSourceFile | Should -Not -Exist
        }
    }
}
Describe 'when a file fails to copy' {
    BeforeAll {
        Mock Copy-Item {
            throw 'Oops'
        }

        $testNewInputFile = Copy-ObjectHC $testInputFile

        $testNewInputFile.Tasks[0].Source.Folder = (New-Item 'TestDrive:/source' -ItemType Directory).FullName
        $testNewInputFile.Tasks[0].Destination.Folder = (New-Item 'TestDrive:/destination' -ItemType Directory).FullName

        $testFile = New-Item "$($testNewInputFile.Tasks[0].Source.Folder)\Analyse_26032025.xlsx" -ItemType File

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams
    }
    It 'an error log file is created' {
        Should -Invoke Out-File -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($LiteralPath -like '* - Actions with errors.json') -and
            ($InputObject -like "*$($testFile.FullName.Replace('\','\\'))*Oops*")
        }
    } -Tag test
}