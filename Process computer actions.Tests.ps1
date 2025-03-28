#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $realCmdLet = @{
        OutFile = Get-Command Out-File
    }

    $testInputFile = @{
        SourceFolder      = (New-Item 'TestDrive:/s' -ItemType Directory).FullName
        DestinationFolder = (New-Item 'TestDrive:/d' -ItemType Directory).FullName
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
                'SourceFolder', 'DestinationFolder'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($FilePath -like '* - Error.txt') -and
                    ($InputObject -like "*$ImportFile*Property '$_' not found*")
                }
            }
            It 'Folder <_> not found' -ForEach @(
                'SourceFolder', 'DestinationFolder'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = 'TestDrive:\nonExisting'

                & $realCmdLet.OutFile @testOutParams -InputObject (
                    $testNewInputFile | ConvertTo-Json -Depth 7
                )

                .$testScript @testParams

                Should -Invoke Out-File -Times 1 -Exactly -ParameterFilter {
                    ($FilePath -like '* - Error.txt') -and
                    ($InputObject -like "*$ImportFile*$_ 'TestDrive:\nonExisting' not found*")
                }
            }
        }
    }
}
Describe 'when the source folder is empty' {
    It 'no error log file is created' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.SourceFolder = (New-Item 'TestDrive:/empty' -ItemType Directory).FullName

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams

        Should -Not -Invoke Out-File
    }
}
Describe 'when there is a file in the source folder' {
    It 'the file is copied to the destination folder with the correct name' {
        $testNewInputFile = Copy-ObjectHC $testInputFile

        $testNewInputFile.SourceFolder = (New-Item 'TestDrive:/source' -ItemType Directory).FullName
        $testNewInputFile.DestinationFolder = (New-Item 'TestDrive:/destination' -ItemType Directory).FullName

        New-Item "$($testNewInputFile.SourceFolder)\Analyse_26032025.xlsx" -ItemType File

        & $realCmdLet.OutFile @testOutParams -InputObject (
            $testNewInputFile | ConvertTo-Json -Depth 7
        )

        .$testScript @testParams

        Get-Item "$($testNewInputFile.DestinationFolder)\2025\AnalysesJour_20250326.xlsx" |
        Should -Not -BeNullOrEmpty
    }
}
Describe 'when a file fails to copy' {
    BeforeAll {
        Mock Copy-Item {
            throw 'Oops'
        }

        $testNewInputFile = Copy-ObjectHC $testInputFile

        $testNewInputFile.SourceFolder = (New-Item 'TestDrive:/source' -ItemType Directory).FullName
        $testNewInputFile.DestinationFolder = (New-Item 'TestDrive:/destination' -ItemType Directory).FullName

        $testFile = New-Item "$($testNewInputFile.SourceFolder)\Analyse_26032025.xlsx" -ItemType File

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