<#
#########################################################
#
# michael.masuch@faps.fau.de
#
# MMa v1 2023-05-31
# MMa v2 2023-10-17: Enhance resilience of script
#
# 
# Description:
#   Silent Installation of 7-Zip
#   and filetype assoziations
#
# Instruction:
#   - Change $FileName to your .exe or .msi of the 7-Zip install file
#   - Execute script in the same folder as the 7-Zip install file
#
# composed with some help of Chat-GPT-4
#
#   
#   Computer\HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\ApplicationAssociationToasts
#   {E88DCCE0-B7B3-11d1-A9F0-00AA0060FA31}
#   
#
#########################################################
#>
Invoke-Command -ScriptBlock {
    # $FileName = '7z2301-x64.msi'
    $FileName = '7z2301-x64.exe'

    # 7zip Installer Path

    if ($PSScriptRoot) {
        $InstallerPath = Join-Path -Path $PSScriptRoot -ChildPath $FileName
    } else {
        $InstallerPath = Join-Path -Path (Get-Location).Path -ChildPath $FileName
    }


    # Install 7zip in silent mode
    # Start-Process -FilePath 'msiexec.exe' -ArgumentList "/i $InstallerPath /qn" -Wait
    Start-Process -FilePath $InstallerPath -ArgumentList '/S' -Wait
    # Invoke-Command -ScriptBlock {
    
    # ------ ▼▼▼ start here when already installed ▼▼▼ 
    # Get 7-Zip install location

    $sevenZipPath = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' | `
        Where-Object { ($_.GetValue('DisplayName') -like '*7-zip*') -and ($_.GetValue('InstallLocation')) } | `
        Get-ItemPropertyValue -Name 'InstallLocation'

    if ($sevenZipPath -is [System.Array]) {
        $sevenZipPath = $sevenZipPath[0]
    }

    if (-not $sevenZipPath) {
        $sevenZipPath = 'C:\Program Files\7-Zip\'
    }

    $sevenZipExePath = Join-Path -Path $sevenZipPath -ChildPath '7zFM.exe'
    $sevenZipDllPath = Join-Path -Path $sevenZipPath -ChildPath '7z.dll'

    # Create PSDrive for HKLM if it doesn't exist
    # Get-PSDrive -PSProvider 'Registry' | Select-Object @('Name', 'Root')

    if (-not (Test-Path -Path 'HKLM:')) {
        New-PSDrive -PSProvider 'Registry' -Root 'HKEY_LOCAL_MACHINE' -Name 'HKLM'
    }
 
    <#
        # Create PSDrive for HKCR if it doesn't exist
        if (-not (Test-Path -Path 'HKCR:')) {
            New-PSDrive -PSProvider 'Registry' -Root 'HKEY_CLASSES_ROOT' -Name 'HKCR'
        }
    #>

    if (Test-Path -Path 'HKLM:\SOFTWARE\Classes\CompressedFolder') {
        Rename-Item -Path 'HKLM:\SOFTWARE\Classes\CompressedFolder' -NewName 'CompressedFolder.BackUp' | Out-Null
    }


    <#
            Rename-Item -Path 'HKCR:\CompressedFolder' -NewName 'CompressedFolder.BackUp'
            Rename-Item -Path 'HKCR:\SystemFileAssociations\.zip' -NewName '.zip.BackUp' -force
            Remove-Item -Path 'HKCR:\SystemFileAssociations\.zip' -force
            Rename-Item -Path 'HKLM:\SOFTWARE\Classes\SystemFileAssociations\.zip' -NewName '.zip.BackUp' -force
            Remove-Item -Path 'HKCR:\SystemFileAssociations\.zip' -force
    #>

    # Define file types and icon indexes

    # icon mapping extracted from the file resource.rc in the archive 7z2201-src.7z path: 7z2201-src.7z\CPP\7zip\Bundles\Format7zF\resource.rc
    $fileTypes = @{
        '7z'       = '0'
        'zip'      = '1'
        'rar'      = '3'
        '001'      = '9'
        #'cab'      = '7'
        #'iso'      = '8'
        'xz'       = '23'
        'txz'      = '23'
        'lzma'     = '16'
        'tar'      = '13'
        'cpio'     = '12'
        'bz2'      = '2'
        'bzip2'    = '2'
        'tbz2'     = '2'
        'tbz'      = '2'
        'gz'       = '14'
        'gzip'     = '14'
        'tgz'      = '14'
        'tpz'      = '14'
        'z'        = '5'
        'taz'      = '5'
        'lzh'      = '6'
        'lha'      = '6'
        'rpm'      = '10'
        'deb'      = '11'
        'arj'      = '4'
        #'vhd'      = '20'
        #'vhdx'     = '20'
        'wim'      = '15'
        'swm'      = '15'
        'esd'      = '15'
        'fat'      = '21'
        'ntfs'     = '22'
        'dmg'      = '17'
        'hfs'      = '18'
        'xar'      = '19'
        'squashfs' = '24'
        'apfs'     = '25'
    }


    # Create registry entries
    foreach ($entry in $fileTypes.GetEnumerator()) {
        $fileType = $entry.Key
        $iconIndex = $entry.Value

        # Paths
        $fileTypePath = "HKLM:\SOFTWARE\Classes\.$fileType"
        $progIdPath = "HKLM:\SOFTWARE\Classes\7-Zip.$fileType"
        $defaultIconPath = "$progIdPath\DefaultIcon"
        $shellPath = "$progIdPath\shell"
        $openPath = "$shellPath\open"
        $commandPath = "$openPath\command"

        # FileType
        if (-not (Test-Path -Path $fileTypePath)) {
            New-Item -Path $fileTypePath -Force | Out-Null
        }
        Set-ItemProperty -Path $fileTypePath -Name '(Default)' -Value "7-Zip.$fileType" -Force | Out-Null

        if ((Test-Path -Path "$fileTypePath\PersistentHandler")) {
            Remove-Item -Path "$fileTypePath\PersistentHandler" -Force | Out-Null
        }

        # ProgId
        if (-not (Test-Path -Path $progIdPath)) {
            New-Item -Path $progIdPath -Force | Out-Null
        }
        Set-ItemProperty -Path $progIdPath -Name '(Default)' -Value "$fileType Archive" -Force | Out-Null

        # DefaultIcon
        if (-not (Test-Path -Path $defaultIconPath)) {
            New-Item -Path $defaultIconPath -Force | Out-Null
        }
        Set-ItemProperty -Path $defaultIconPath -Name '(Default)' -Value "$sevenZipDllPath,$iconIndex" -Force | Out-Null

        # shell
        if (-not (Test-Path -Path $shellPath)) {
            New-Item -Path $shellPath -Force | Out-Null
        }
        Set-ItemProperty -Path $shellPath -Name '(Default)' -Value 'open' -Force | Out-Null

        # open
        if (-not (Test-Path -Path $openPath)) {
            New-Item -Path $openPath -Force | Out-Null
        }

        # command
        if (-not (Test-Path -Path $commandPath)) {
            New-Item -Path $commandPath -Force | Out-Null
        }
        Set-ItemProperty -Path $commandPath -Name '(Default)' -Value "`"$sevenZipExePath`" `"%1`"" -Force | Out-Null
    }
}