#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $realCmdLet = @{
        OutFile = Get-Command Out-File
    }

    $testInputFile = @{
        SourceFolder      = (New-Item 'TestDrive:/s' -ItemType Directory).FullName
        ArchiveFolder     = (New-Item 'TestDrive:/a' -ItemType Directory).FullName
        DestinationFolder = (New-Item 'TestDrive:/d' -ItemType Directory).FullName
        FileExtensions    = @()
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
                'SourceFolder', 'ArchiveFolder', 'DestinationFolder'
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
                'SourceFolder', 'ArchiveFolder', 'DestinationFolder'
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