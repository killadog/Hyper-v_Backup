<#
.SYNOPSIS 
Hyper-v backup (export) tool

.EXAMPLE
Backup Hyper-v virtual machine 'vm_debian' to 'C:\Backup'
with Log file in folder 'C:\Logs'
and delete backups older than 5 days
and send e-mail with Log to example@example.com

PS> .\hvb.ps1 -VmName vm_debian -To C:\Backup -LogDir C:\Logs -Log -DaysBack 5 -Email example@example.com

or simple

PS> .\hvb.ps1 -VmName vm_debian -To C:\Backup

.NOTES
Author: Rad
Date: July 14, 2022
URL: https://github.com/killadog
#>

param ([parameter(Mandatory = $false)][switch] $Info ## Information about all VMs
    , [parameter(Mandatory = $false)][string[]] $VmName ## One VM name for backup. (Use without -VmFile and -All)
    , [parameter(Mandatory = $false)][string[]] $VmFile ## File with VM names for backup. (Use without -VmName and -All)
    , [parameter(Mandatory = $false)][switch] $All ## Backup all VMs. (Use without -VmName and -VmFile)
    , [parameter(Mandatory = $false)][string[]] $To ## Folder for backup
    , [parameter(Mandatory = $false)][ValidateRange(1, 365)][int] $DaysBack ## Delete backups older than N days [1..365]
    , [parameter(Mandatory = $false)][string[]] $Email ## Send e-mail with Log to recepient(s) (comma separated)
    , [parameter(Mandatory = $false)][string[]] $LogDir = $env:temp ## Log to folder (default is %TEMP%)
    , [parameter(Mandatory = $false)][switch] $Log ## Save Log
    , [parameter(Mandatory = $false)][switch] $Help ## This help screen
)

#(Get-History)[-1]
#exit

function help () {
    Get-Command -Syntax $PSCommandPath
    Get-Help $PSCommandPath -Parameter * | Format-Table -Property @{name = 'Option'; Expression = { $($PSStyle.Foreground.BrightGreen) + "-" + $($_.'name') } },
    @{name = 'Type'; Expression = { $($PSStyle.Foreground.BrightWhite) + $($_.'parameterValue') } },
    @{name = 'Default'; Expression = { if ($($_.'defaultValue' -notlike 'String')) { $($PSStyle.Foreground.BrightWhite) + $($_.'defaultValue') } }; align = 'Center' },
    @{name = 'Explanation'; Expression = { $($PSStyle.Foreground.BrightYellow) + $($_.'description').Text } }
    exit
}

function Format-FileSize() {
    Param ([uint64]$size)
    If ($size -gt 1TB) { [string]::Format("{0:0.00} TB", $size / 1TB) }
    ElseIf ($size -gt 1GB) { [string]::Format("{0:0.00} GB", $size / 1GB) }
    ElseIf ($size -gt 1MB) { [string]::Format("{0:0.00} MB", $size / 1MB) }
    ElseIf ($size -gt 1KB) { [string]::Format("{0:0.00} kB", $size / 1KB) }
    ElseIf ($size -gt 0) { [string]::Format("{0:0.00} B", $size) }
    Else { "" }
}

function Information {
    $VmList = @()
    $Vms | ForEach-Object {
        $VmState = Get-VM -VMName $_ | Select-Object -ExpandProperty State
        $VmUptime = (Get-VM -VMName $_ | Select-Object -ExpandProperty Uptime).ToString("dd\.hh\:mm\:ss")
        $VmFileSize = Format-FileSize(Get-VM -VMName $_ | Select-Object VMid | Get-VHD | Select-Object -ExpandProperty FileSize)
        $VmPath = Get-VM -VMName $_ | Select-Object VMid | Get-VHD | Select-Object -ExpandProperty Path
        $VmVhdType = Get-VHD $VmPath | Select-Object -ExpandProperty VhdType
        $VmNotes = Get-VM -VMName $_ | Select-Object -ExpandProperty Notes
        $VmList += [PSCustomObject] @{
            'Number'   = $i++ + 1
            'VM name'  = $_
            'State'    = $VmState
            'Uptime'   = $VmUptime
            'Vhd type' = $VmVhdType
            'VHD size' = $VmFileSize
            'VHD path' = $VmPath
            'Notes'    = $VmNotes
        }
    }

    $VmList | Format-Table -Property @{n = "Number"; e = { $_.Number }; a = "center" },
    @{n = "VM name"; e = { $($PSStyle.Foreground.BrightYellow) + $_.'VM name' }; a = "center" },
    @{n = "State"; e = { if ($_.State -eq 'Running') { $($PSStyle.Foreground.BrightGreen) + $_.State } else { $($PSStyle.Foreground.BrightRed) + $_.State } }; a = "center" },
    @{n = "Uptime"; e = { $_.Uptime }; a = "center" },
    @{n = "VHD type"; e = { $_.'Vhd type' }; a = "center" },
    @{n = "VHD size"; e = { $_.'VHD size' }; a = "center" },
    @{n = "VHD path"; e = { $_.'VHD path' }; a = "left" },
    @{n = "Notes"; e = { $_.Notes }; a = "left" }
}

if ([System.Version]"$($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)" -lt [System.Version]"7.2") {
    $PSStyle.OutputRendering = 'PlainText'
}
else {
    $PSStyle.OutputRendering = 'Ansi'
    $PSStyle.Formatting.TableHeader = $PSStyle.Foreground.BrightBlack + $PSStyle.Italic
}

if ($Info) {
    $Vms = get-vm | Select-Object -ExpandProperty Name
    Information
    Exit
}

#if ($Help -or (!$args.Count)) {
if ($Help) {
    help
}

if ((($VmName -and $VmFile) -or ($VmName -and $All)) -or (($VmFile -and $All)) -or (!$VmName -and !$VmFile -and !$All)) {
    Write-Error "Choose only ONE from [-VmName | -VmFile | -All]"
    help
}

if (!$To) {
    Write-Error "Set destination folder to backup [-To]"
    help  
}

$To = $To.trimend('\')
$LogDir = $LogDir.trimend('\')

If (Test-Path -Path $LogDir) {
    Write-Host "$LogDir already exists"
}
else {
    Write-Host "Create $LogDir"
    New-Item -ItemType Directory $LogDir -Force | Out-Null
}

$DateFormat = Get-Date -uformat "%Y%m%d_%H%M%S"
$LogPrefix = 'hvb'
$LogFile = "$LogDir\${LogPrefix}_log_$DateFormat.txt"
$EmailFrom = 'example@example.com'
$EmailPassword = 'SECRETPASSWORD!'
$EmailSmtpServer = 'smtp.example.com'
$EmailSmtpPort = 587
$EmailSubject = "Hyper-V backup. $($env:COMPUTERNAME)"

$str = "$($PSStyle.Foreground.BrightYellow)-" * 70

Start-Transcript -Append $LogFile -UseMinimalHeader 

if ($VmFile) {
    $VMs = @(Get-Content $VmFile | Where-Object { ($_.trim() -ne "") -and (!$_.StartsWith("#")) })
}
if ($VmName) {
    $Vms = @($VmName)
}
if ($All) {
    $Vms = get-vm | Select-Object -ExpandProperty Name
}

Write-Host $str
Write-Host "VM(s) to backup:" -NoNewline

Information

if ($Info) {
    Exit
}

$TotalTime = Measure-Command {
    $VmBackup = @()
    $VMs | ForEach-Object {
        $Counter++
        Write-Host $str
        Write-Host "VM '$_' [$Counter/$($VMs.Length)]"

        try {
            $Exists = get-vm -name $_ -ErrorAction Stop
        }
        catch {
            #$Exists = False
        }
        
        If ($Exists) {
            $VmBackupDir = "$To\$_\${_}_$DateFormat"
            $BackupTime = Measure-Command {
                Write-Host "`nCreate folder '$VmBackupDir'"
                New-Item -ItemType Directory $VmBackupDir -Force
                        
                try {
                    Write-Host "Start to backup"
                    Export-VM -Name $_ -Path $VmBackupDir -Passthru -ErrorAction Stop
                    $Result = $true
                    Write-Host "Backup successful"            
                }
                catch {
                    #$_.Exception.Message | Write-Host "$_ $_"
                    #Write-Host "Error: $_"
                    $Result = $false
                }
                if ($Result) {
                    Move-Item $To\$_\${_}_$DateFormat\$_\* $To\$_\${_}_$DateFormat\
                    Remove-Item $To\$_\${_}_$DateFormat\$_ -Recurse -Force -Confirm:$false

                    If ($DaysBack) {
                        Write-Host "`nDelete backups older than $DaysBack days:"
                        $DeleteFromDate = (Get-Date).AddDays(-$DaysBack)
                        $FoldersToDelete = Get-ChildItem -Path $To\$_  | Where-Object { $_.CreationTime -le $DeleteFromDate } | Select-Object -ExpandProperty FullName 
                        if ($FoldersToDelete) {
                            Write-Host $FoldersToDelete -Separator "`n"
                            Get-ChildItem -Path $To\$_ | Where-Object { $_.CreationTime -le $DeleteFromDate } | Remove-Item -Recurse -Force
                        }
                        else {
                            Write-Host "Nothing to delete"
                        }
                    }
                    Write-Host "`nCurrent folders with backups in '$To\$_'"
                    Get-ChildItem -Path $To\$_ | Select-Object -ExpandProperty FullName | Out-Host
                    $Result = 'Success'
                }
                else { 
                    Write-Error "Backup failed" 
                }
            }

            $BackupTime = $BackupTime.ToString("hh\:mm\:ss")
            Write-Host "`nVM '$_' backup time: $BackupTime`n"
            $VmNotes = Get-VM -VMName $_ | Select-Object -ExpandProperty Notes
        }
        Else {
            Write-Host "Error - no VM $_"
            $EmailSubject += " Alarm!"
            $Result = 'Failed'
            $VmNotes = ''
        }
        $VmBackup += [PSCustomObject] @{
            'Number'      = $Counter
            'VM name'     = $_
            'Result'      = $Result
            'Backup time' = $BackupTime
            'Notes'       = $VmNotes
        }
    }
}

Write-Host $str
Write-Host 'Results:'
$VmBackup | Format-Table -Property @{n = "Number"; e = { $_.Number }; a = "center" },
@{n = "VM name"; e = { $_.'VM name' }; a = "center" },
@{n = "Result"; e = { if ($_.Result -eq 'Success') { $($PSStyle.Foreground.BrightGreen) + $_.Result } else { $($PSStyle.Foreground.BrightRed) + $_.Result } }; a = "center" } ,
@{n = "Backup time"; e = { $_.'Backup time' }; a = "center" },
@{n = "Notes"; e = { $_.Notes }; a = "left" }

$TotalTime = $TotalTime.ToString("hh\:mm\:ss")
Write-Host "Total backup time: $TotalTime"

Stop-Transcript

if ($Email) {
    $body = Get-Content -Path $LogFile | Out-String 
    $secpasswd = ConvertTo-SecureString $EmailPassword -AsPlainText -Force
    $mycreds = New-Object System.Management.Automation.PSCredential ($EmailFrom, $secpasswd)
    $encoding = [System.Text.Encoding]::UTF8
    [System.Net.ServicePointManager]::SecurityProtocol = "Tls, TLS11, TLS12" # Uncomment it to use TLS not SSL
    Send-MailMessage -To $Email -Subject $EmailSubject -Body $body -SmtpServer $EmailSmtpServer -Credential $mycreds -Port $EmailSmtpPort -UseSsl -from $EmailFrom -Encoding $encoding -WarningAction:SilentlyContinue
    Write-Host "`nSend e-mail to $Email`n"
}

if (!$Log) {
    Write-Host "Delete '$LogFile'"
    Remove-Item $LogFile -Force
}