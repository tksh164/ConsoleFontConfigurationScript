#Requires -Version 5.0

function Restart-PowerShellConsoleAsAdmin
{
    [OutputType([void])]
    param ()

    $curretUserId = [System.Security.Principal.WindowsIdentity]::GetCurrent();
    $currentWindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($curretUserId);
    $adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator;

    if (-not $currentWindowsPrincipal.IsInRole($adminRole))
    {
        # Relaunch this script in the administrator's PowerShell console.
        $params = @{
            FilePath     = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
            ArgumentList = '& "{0}"' -f $PSCommandPath
            Verb         = 'runas'
        }
        Start-Process @params

        # Exit the current PowerShell console.
        exit
    }
}

function Get-FontArchiveFile
{
    [OutputType([string])]
    param (
        [string] $WorkFolderPath,
        [string] $Uri
    )

    # Create a directory for work.
    [void](New-Item -Path $WorkFolderPath -ItemType Directory)

    # Download the font archive file from the given URI.
    $downloadFilePath = Join-Path -Path $WorkFolderPath -ChildPath 'fonts.zip'
    Invoke-WebRequest -Method Get -Uri $Uri -OutFile $downloadFilePath

    return $downloadFilePath
}

function Install-DownloadedFontFile
{
    [OutputType([void])]
    param (
        [string] $DownloadFilePath
    )

    # Expand the zip file.
    $expandDestPath = Join-Path -Path $workFolderPath -ChildPath 'fonts'
    Expand-Archive -LiteralPath $DownloadFilePath -DestinationPath $expandDestPath

    # Copy the font files from expanded files.
    Get-ChildItem -LiteralPath $expandDestPath -Filter 'migu-1m-*.ttf' -File -Recurse -Depth 2 |
        ForEach-Object -Process {
            $fontPath = Join-Path -Path 'C:\Windows\Fonts' -ChildPath $_.Name

            # Delete installed font file.
            $params = @{
                FilePath     = 'C:\Windows\System32\cmd.exe'
                ArgumentList = '/c del "{0}"' -f $fontPath
                Wait         = $true
                WindowStyle  = [System.Diagnostics.ProcessWindowStyle]::Hidden
            }
            Start-Process @params

            # Copy new font file.
            Copy-Item -LiteralPath $_.FullName -Destination $fontPath
        }
}

function Set-ConsoleFontRegistry
{
    [OutputType([void])]
    param (
        [string] $WorkFolderPath
    )

    # Content of the reg file.
    $regFileContent = @"
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts]
"Migu 1M Regular (TrueType)"="migu-1m-regular.ttf"
"Migu 1M Bold (TrueType)"="migu-1m-bold.ttf"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Console\TrueTypeFont]
"932"="Migu 1M"
"@

    # Create a reg file.
    $regFilePath = Join-Path -Path $WorkFolderPath -ChildPath 'ttf.reg'
    Set-Content -LiteralPath $regFilePath -Value $regFileContent -Encoding Unicode

    # Import the registry settings from a reg file.
    Start-Process -FilePath 'C:\Windows\System32\reg.exe' -ArgumentList 'import',$regFilePath -Wait
}


# Relaunch this script in the administrator's PowerShell console if needed.
Restart-PowerShellConsoleAsAdmin

# URI of the font archive file.
$uri = 'http://osdn.jp/frs/redir.php?m=iij&f=%2Fmix-mplus-ipa%2F63545%2Fmigu-1m-20150712.zip'

# Download the font archive file from the given URI.
$workFolderPath = Join-Path -Path $env:Temp -ChildPath (New-Guid).ToString()
$downloadFilePath = Get-FontArchiveFile -WorkFolderPath $workFolderPath -Uri $uri

# Install the font files from a downloaded file.
Install-DownloadedFontFile -DownloadFilePath $downloadFilePath

# Setting the console font registry.
Set-ConsoleFontRegistry -WorkFolderPath $workFolderPath

# Delete the working directory.
Remove-Item -Path $workFolderPath -Recurse
