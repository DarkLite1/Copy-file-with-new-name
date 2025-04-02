#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $realCmdLet = @{
        OutFile = Get-Command Out-File
    }

    $testInputFile = @{
        Action                                   = 'copy'
        Source                                   = @{
            Folder             = (New-Item 'TestDrive:/s' -ItemType Directory).FullName
            MatchFileNameRegex = 'Analyse_[0-9]{8}.xlsx'
            Recurse            = $true
        }
        Destination                              = @{
            Folder        = (New-Item 'TestDrive:/d' -ItemType Directory).FullName
            OverWriteFile = $true
        }
        ProcessFilesCreatedInTheLastNumberOfDays = 1
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
    }

    Mock Out-File
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ImportFile') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'create an error log file when' {
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
            ($FilePath -like '*\Error.txt') -and
            ($InputObject -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'

            .$testScript @testNewParams

            Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                ($FilePath -like '* - Error.txt') -and
                ($InputObject -like '*Cannot find path*nonExisting.json*')
            }
        }
        Context 'property' {
            It '<_> not found' -ForEach @(
                'Folder', 'MatchFileNameRegex'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Source.$_ = $null

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($FilePath -like '* - Error.txt') -and
                    ($InputObject -like "*$ImportFile*Property 'Source.$_' not found*")
                }
            }
            It '<_> not found' -ForEach @(
                'Folder'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Destination.$_ = $null

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($FilePath -like '* - Error.txt') -and
                    ($InputObject -like "*$ImportFile*Property 'Destination.$_' not found*")
                }
            }
            It 'Folder <_> not found' -ForEach @(
                'Source', 'Destination'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile[$_]['Folder'] = 'TestDrive:\nonExisting'

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($FilePath -like '* - Error.txt') -and
                    ($InputObject -like "*$ImportFile*$_.Folder 'TestDrive:\nonExisting' not found*")
                }
            }
            Context 'ProcessFilesCreatedInTheLastNumberOfDays' {
                It 'is not a number' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.ProcessFilesCreatedInTheLastNumberOfDays = 'a'

                    & $realCmdLet.OutFile @testOutParams -InputObject (
                        $testNewInputFile | ConvertTo-Json -Depth 7
                    )

                    .$testScript @testParams

                    Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                        ($FilePath -like '* - Error.txt') -and
                        ($InputObject -like "*$ImportFile*Property 'ProcessFilesCreatedInTheLastNumberOfDays' must be 0 or a positive number. Number 0 processes all files in the source folder. The value 'a' is not supported*")
                    }
                }
                It 'is a negative number' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.ProcessFilesCreatedInTheLastNumberOfDays = -1

                    & $realCmdLet.OutFile @testOutParams -InputObject (
                        $testNewInputFile | ConvertTo-Json -Depth 7
                    )

                    .$testScript @testParams

                    Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                        ($FilePath -like '* - Error.txt') -and
                        ($InputObject -like "*$ImportFile*Property 'ProcessFilesCreatedInTheLastNumberOfDays' must be 0 or a positive number. Number 0 processes all files in the source folder. The value '-1' is not supported*")
                    }
                }
                It 'is an empty string' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.ProcessFilesCreatedInTheLastNumberOfDays = ''

                    & $realCmdLet.OutFile @testOutParams -InputObject (
                        $testNewInputFile | ConvertTo-Json -Depth 7
                    )

                    .$testScript @testParams

                    Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                        ($FilePath -like '* - Error.txt') -and
                        ($InputObject -like "*$ImportFile*Property 'ProcessFilesCreatedInTheLastNumberOfDays' must be 0 or a positive number. Number 0 processes all files in the source folder. The value '' is not supported*")
                    }
                }
                It 'is missing' {
                    $testNewInputFile = @{
                        Action      = 'copy'
                        Source      = $testInputFile.Source
                        Destination = $testInputFile.Destination
                    }

                    & $realCmdLet.OutFile @testOutParams -InputObject (
                        $testNewInputFile | ConvertTo-Json -Depth 7
                    )

                    .$testScript @testParams

                    Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                        ($FilePath -like '* - Error.txt') -and
                        ($InputObject -like "*$ImportFile*Property 'ProcessFilesCreatedInTheLastNumberOfDays' must be 0 or a positive number. Number 0 processes all files in the source folder. The value '' is not supported*")
                    }
                }
                It '0 is accepted' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.ProcessFilesCreatedInTheLastNumberOfDays = '0'

                    & $realCmdLet.OutFile @testOutParams -InputObject (
                        $testNewInputFile | ConvertTo-Json -Depth 7
                    )

                    .$testScript @testParams

                    Should -Invoke -Not Out-File -ParameterFilter {
                        ($FilePath -like '* - Error.txt') -and
                        ($InputObject -like "*ProcessFilesCreatedInTheLastNumberOfDays*")
                    }
                }
            }
        }
    }
}
Describe 'when the source folder is empty' {
    It 'no error log file is created' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Source.Folder = (New-Item 'TestDrive:/empty' -ItemType Directory).FullName

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams

        Should -Not -Invoke Out-File
    }
}
Describe 'when there is a file in the source folder' {
    It 'the file is copied to the destination folder' {
        $testNewInputFile = Copy-ObjectHC $testInputFile

        $testNewInputFile.Source.Folder = (New-Item 'TestDrive:/source' -ItemType Directory).FullName
        $testNewInputFile.Destination.Folder = (New-Item 'TestDrive:/destination' -ItemType Directory).FullName

        New-Item "$($testNewInputFile.Source.Folder)\Analyse_26032025.xlsx" -ItemType File

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams

        Get-Item "$($testNewInputFile.Destination.Folder)\Analyse_26032025.xlsx" |
        Should -Not -BeNullOrEmpty
    }
}
Describe 'when a file fails to copy' {
    BeforeAll {
        Mock Copy-Item {
            throw 'Oops'
        }

        $testNewInputFile = Copy-ObjectHC $testInputFile

        $testNewInputFile.Source.Folder = (New-Item 'TestDrive:/source' -ItemType Directory).FullName
        $testNewInputFile.Destination.Folder = (New-Item 'TestDrive:/destination' -ItemType Directory).FullName

        $testFile = New-Item "$($testNewInputFile.Source.Folder)\Analyse_26032025.xlsx" -ItemType File

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams
    }
    It 'an error log file is created' {
        Should -Invoke Out-File -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($FilePath -like '* - Error.txt') -and
            ($InputObject -like "*Failure for source file*Failed to copy file '$($testFile.FullName)'*")
        }
    }
}