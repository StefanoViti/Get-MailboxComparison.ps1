<#
.SYNOPSIS
Get-MailboxComparison.ps1 - In a Microsoft 365 T2T migration, this script performs a mailbox migration assessment and monitoring.

.DESCRIPTION 
In a Microsoft 365 T2T migration project, this script makes the comparison between a mailbox on the source tenant (e.g., john.smith@contoso.com)
and its equivalent on the target tenant (e.g., john.smith@fabrikam.com), to verify if everything (in terms of items and data size) has been
correctly transferred.

.INPUTS
You have to provide a csv containing the list of the source mailboxes named .\SourceMailboxes.csv, in the same folder where this cript is saved
with "PrimarySMTPAddres" as header.
You must have user credentials for connecting to Exchange Online on both tenants: the script will propmpt you for credentials.

.OUTPUTS
A final csv comparison report will be provided at ".\Delta Migration\DeltaMigration_$Date.csv".

.REQUIREMENTS
You must have reading permission on mailboxes on both tenants. The script will propmpt you for credentials.

.PARAMETER SourceDomain
Insert the Source Domain of the mailboxes (e.g., @contoso.com).

.PARAMETER TargetDomain
Insert the Target Domain of the mailboxes (e.g., @fabrikam.com).

.EXAMPLE
.\Get-MailboxComparison.ps1 -SourceDomain "@contoso.com" -TargetDomain "@fabrikam.com"

.NOTES
Written by: Stefano Viti - stefano.viti1995@gmail.com
Follow me at https://www.linkedin.com/in/stefano-viti/

#>

param(
    [Parameter(Mandatory=$True)]
    [string]$SourceDomain,

	[Parameter(Mandatory=$True)]
    [switch]$TargetDomain
)

clear

Write-Host "Insert Source-Tenant Credential" -ForegroundColor White
Connect-ExchangeOnline -Prefix "ST"
Write-Host "Insert Target-Tenant Credential" -ForegroundColor White
Connect-ExchangeOnline -Prefix "TT"

clear

$Date = Get-Date -Format "yyyy_MM_dd"
$Mailboxes = Import-csv -path ".\SourceMailboxes.csv" -Encoding UTF8

$Results = @()
$ResultFilePath = ".\Delta Migration\DeltaMigration_" + $Date + ".csv"

$i = 0
$tot = $Mailboxes.count

Foreach ($Mailbox in $Mailboxes){
    $i ++
    Write-Host "Analyzing the data of the mailbox $($Mailbox.PrimarySMTPAddress). - [$($i)/$($tot)]" -ForegroundColor Yellow

    $SourcePrimarySMTP = $Mailbox.PrimarySMTPAddress
    $TargetPrimarySMTP = $Mailbox.PrimarySMTPAddress.replace($SourceDomain,$TargetDomain)

    $checkTT = $null
    $ErrorActionPreference = "SilentlyContinue"
    $checkTT = Get-TTMailbox $TargetPrimarySMTP
    $ErrorActionPreference = "Continue" 

    $checkST = $null
    $ErrorActionPreference = "SilentlyContinue"
    $checkST = Get-STMailbox $SourcePrimarySMTP
    $ErrorActionPreference = "Continue"

    $checkSTArchive = $null
    $ErrorActionPreference = "SilentlyContinue"
    $checkSTArchive = Get-STMailbox -Archive $SourcePrimarySMTP
    $ErrorActionPreference = "Continue"

    if ($checkTT -ne $null -and $checkST -ne $null -and $checkSTArchive -ne $null){

        #The mailbox has been already migrated and it still exists on the source tenant with the archive

        $SourceData = Get-STMailboxStatistics $SourcePrimarySMTP
        $TargetData = Get-TTMailboxStatistics $TargetPrimarySMTP
        $SourceDataArchive = Get-STMailboxStatistics $SourcePrimarySMTP -Archive
        $TargetDataArchive = Get-TTMailboxStatistics $TargetPrimarySMTP -Archive

        $hash = [ordered]@{
            PrimarySMTPAddress = $TargetPrimarySMTP

            SourcePrimaryMBXItem = $SourceData.ItemCount
            TargetPrimaryMBXItem = $TargetData.ItemCount
            DeltaItem = ($SourceData.ItemCount - $TargetData.ItemCount)
            'DeltaItemPercentage [%]' = (($SourceData.ItemCount - $TargetData.ItemCount)/$SourceData.ItemCount)*100

            'SourcePrimaryMBXSize [GB]' = [math]::Round(($SourceData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            'TargetPrimaryMBXSize [GB]' = [math]::Round(($TargetData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            'DeltaPrimaryMBXSize [GB]' = [math]::Round(($SourceData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2) - [math]::Round(($TargetData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            'DeltaPrimaryMBXPercentage [%]' = (([math]::Round(($SourceData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2) - [math]::Round(($TargetData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2))/[math]::Round(($SourceData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2))*100

            SourceArchiveMBXItem = $SourceDataArchive.ItemCount
            TargetArchiveMBXItem = $TargetDataArchive.ItemCount
            DeltaArchiveItem = ($SourceDataArchive.ItemCount - $TargetDataArchive.ItemCount)
            'DeltaArchiveItemPercentage [%]' = (($SourceDataArchive.ItemCount - $TargetDataArchive.ItemCount)/$SourceDataArchive.ItemCount)*100

            'SourceArchiveMBXSize [GB]' = [math]::Round(($SourceDataArchive.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            'TargetArchiveMBXSize [GB]' = [math]::Round(($TargetDataArchive.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            'DeltaArchiveMBXSize [GB]' = [math]::Round(($SourceDataArchive.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2) - [math]::Round(($TargetDataArchive.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            'DeltaArchiveMBXPercentage [%]' = (([math]::Round(($SourceDataArchive.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2) - [math]::Round(($TargetDataArchive.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2))/[math]::Round(($SourceDataArchive.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2))*100
            }

        $Item = New-Object PSObject -Property $hash
        $Results = $Results + $Item
    }
    if ($checkTT -ne $null -and $checkST -ne $null -and $checkSTArchive -eq $null){

        #The mailbox has been already migrated and it still exists on the source tenant without the archive

        $SourceData = Get-STMailboxStatistics $SourcePrimarySMTP
        $TargetData = Get-TTMailboxStatistics $TargetPrimarySMTP

        $hash = [ordered]@{
            PrimarySMTPAddress = $TargetPrimarySMTP

            SourcePrimaryMBXItem = $SourceData.ItemCount
            TargetPrimaryMBXItem = $TargetData.ItemCount
            DeltaItem = ($SourceData.ItemCount - $TargetData.ItemCount)
            'DeltaItemPercentage [%]' = (($SourceData.ItemCount - $TargetData.ItemCount)/$SourceData.ItemCount)*100

            'SourcePrimaryMBXSize [GB]' = [math]::Round(($SourceData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            'TargetPrimaryMBXSize [GB]' = [math]::Round(($TargetData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            'DeltaPrimaryMBXSize [GB]' = [math]::Round(($SourceData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2) - [math]::Round(($TargetData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2)
            'DeltaPrimaryMBXPercentage [%]' = (([math]::Round(($SourceData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2) - [math]::Round(($TargetData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2))/[math]::Round(($SourceData.TotalItemSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1GB),2))*100

            SourceArchiveMBXItem = "Archivio non attivo sul source"
            TargetArchiveMBXItem = "Archivio non attivo sul source"
            DeltaArchiveItem = "Archivio non attivo sul source"
            'DeltaArchiveItemPercentage [%]' = "Archivio non attivo sul source"

            'SourceArchiveMBXSize [GB]' = "Archivio non attivo sul source"
            'TargetArchiveMBXSize [GB]' = "Archivio non attivo sul source"
            'DeltaArchiveMBXSize [GB]' = "Archivio non attivo sul source"
            'DeltaArchiveMBXPercentage [%]' = "Archivio non attivo sul source"
            }

        $Item = New-Object PSObject -Property $hash
        $Results = $Results + $Item
    }
    if($checkST -eq $null){

        #The mailbox does not exist anymore on the source tenant

        Write-Host "The mailbox $($SourcePrimarySMTP) does not exist, please verify on the source tenant! - [$($i)/$($tot)]" -ForegroundColor Red
    }
}

$Results | Export-csv -Path $ResultFilePath -Encoding UTF8 -NoTypeInformation

$Prompt = New-Object -ComObject wscript.shell
$UserInput = $Prompt.popup("Do you want to open output file?", 0,"Open Output File",4)
If ($UserInput -eq 6){
    Invoke-Item $ResultFilePath
}