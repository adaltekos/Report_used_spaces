#Install-Module -Name SharePointPnPPowerShellOnline -RequiredVersion 3.2.1810.0
#Register-PnPAzureADApp -ApplicationName "PnP App" -Tenant "duondystrybucja.onmicrosoft.com" -Interactive -Store CurrentUser


$filename 		= '' #Complete with filename (ex. Raport_used_spaces.xlsx)
$localPath 		= '' #Complete with local path (ex. C:\Raporty\)
$siteUrl		= '' #Complete with Url site (ex. https://company.sharepoint.com/sites/it-dep)
$onlinePath		= '' #Complete with path where file is on sharepoint (ex. Shared Documents/Global/)
$tenant			= '' #Complete with tenant name (ex. company.onmicrosoft.com)
$appId			= '' #Complete with ClientId (which is ID of application registered in Azure AD)
$thumbprint		= '' #Complete with Thumbprint (which is certificate thumbprint)


#Pobranie pliku z Teams
Connect-PnPOnline -Url $siteUrl -Tenant $tenant -ClientId $appId -Thumbprint $thumbprint

Get-PnPFile -Url ($onlinePath + $filename) -Path $localPath -Filename $filename -AsFile -Force

Start-Sleep -s 3


#OneDrve
Get-PnPTenantSite -IncludeOneDriveSites |
    Where-Object {($_.Template -eq 'SPSPERS#10')} |
    Select @{n=$(Get-Date -Format dd.MM.yyyy).ToString();e={}}, Title, Owner, @{n="StorageUsageCurrent";e={[math]::round(($_.StorageUsageCurrent / 1024),2)}}, @{n="StorageQuota";e={[math]::round(($_.StorageQuota / 1024),2)}} |
    Export-Excel -Path ($localPath + $filename) -WorkSheetname OneDrive -AutoSize

Start-Sleep -s 3


#Sharepoint
Get-PnPTenantSite |
    Where-Object {($_.Template -eq 'GROUP#0') -or ($_.Template -eq 'SITEPAGEPUBLISHING#0')} |
    Select @{n=$(Get-Date -Format dd.MM.yyyy).ToString();e={}}, Title, Url, @{n="StorageUsageCurrent";e={[math]::round(($_.StorageUsageCurrent / 1024),2)}} |
    Export-Excel -Path ($localPath + $filename) -WorkSheetname SharePoint -AutoSize

Start-Sleep -s 3


#Exchange
Connect-ExchangeOnline -AppID $appId -CertificateThumbPrint $thumbprint -Organization $tenant

$mailboxes = Get-Mailbox -ResultSize Unlimited

$exchange = foreach ($mail in $mailboxes) {
    Get-MailboxStatistics -identity $mail.UserPrincipalName | 
    select @{n=$(Get-Date -Format dd.MM.yyyy).ToString();e={}},
	       @{n='UserPrincipalName';e={$mail.UserPrincipalName}}, 
           @{n='RecipientTypeDetails';e={$mail.RecipientTypeDetails}}, 
           @{n='TotalItemSize';e={[math]::round((($_.TotalItemSize.Value.ToString().Split('(')[1].Trim(" bytes)") -replace ',','') / (1024*1024*1024)),2)}}
}

$exchange | Export-Excel -Path ($localPath + $filename) -WorkSheetname Exchange -AutoSize

Start-Sleep -s 3


#Servers
$servers1 = @('server1', 'server2') #Windows 2012
$servers2 = @('server3', 'server4', 'server5', 'server6', 'server7', 'server8', 'server9', 'server10') #Windows 2016-2019 / Windows 10
$servers3 = @('server11', 'server12', 'server13') #Windows 2022 / Windows 11
$servers = $servers1 + $servers2 + $servers3

function WhichServer($server) {
    if ($script:servers1 -contains $server) {
        return 1
    }
    elseif ($script:servers2 -contains $server) {
        return 2
    }
    elseif ($script:servers3 -contains $server) {
        return 3
    }
}

function GetDriveLetters($server) {
    $drives = @{}
    $driveletters = Invoke-Command -ComputerName $server {
        (fsutil fsinfo drives)[1].Split().Trim()
    }
    foreach ($driveletter in $driveletters) {
        if ($driveletter.EndsWith(':\') -and ((Invoke-Command -ComputerName $server {fsutil fsinfo drivetype $Using:driveletter}).EndsWith('Fixed Drive'))) {
            $drive = Invoke-Command -ComputerName $server {
                Write-Output ("Date : ")
                Write-Output ("Partition : " + $Using:driveletter)
                (fsutil volume diskfree $Using:driveletter)
            } | ConvertFrom-String -PropertyNames ('Name', 'Size') -Delimiter " : "
			$drive[0].Size = (Get-Date -Format dd.MM.yyyy)
            $drives[$drive[1].Size] = $drive
        } 
    }
    return $drives
}

function ConvertToGb($line, $server) {
    
        switch (WhichServer($server)) {
            1 {
                return ([math]::round(($line.Size / (1024*1024*1024)),2))
            }
            2 {
                return ([math]::round(((($line.Size).Trim().Split()[0].Trim() -replace "˙","") / (1024*1024*1024)),2))
            }
            3 {
                return ([math]::round(((($line.Size).Trim().Split()[0].Trim() -replace ",","") / (1024*1024*1024)),2))
            }
        } 
}

function CreateVolume($drive, $server) {
    $volume = New-Object -TypeName psobject
    foreach ($line in $drive) {
        if ($line.Name.Trim() -like 'Date' -or $line.Name.Trim() -like 'Partition') {
            $volume | Add-Member -NotePropertyName $line.Name.Trim() -NotePropertyValue $line.Size
        }
        else {
            $volume | Add-Member -NotePropertyName $line.Name.Trim() -NotePropertyValue (ConvertToGb -line $line -server $server)
        }
    }
    return $volume
}

function ExportToExcel($volume, $server) {
    Clear-Variable -Name data
    try {
        $data = Import-Excel -Path ($script:localPath + $script:filename) -WorksheetName $server
    }
    catch {
        Write-Warning "An error occurred while importing Excel data: $($_.Exception.Message)"
    }
    if ($data -eq $null) {
        Export-Excel -InputObject $volume -Path ($script:localPath + $script:filename) -WorkSheetname $server -AutoSize
    }
    else {
        if (!$data.GetType().IsArray) {
            $data = @($data)
        }
        $data = @($volume) + @($data)
        Export-Excel -InputObject $data -Path ($script:localPath + $script:filename) -WorkSheetname $server -AutoSize
    }
}

foreach ($server in $servers) {
    $drives = GetDriveLetters($server)
    foreach ($drive in $drives.Keys) {
        $volume = (CreateVolume -drive ($drives[$drive]) -server $server)
        ExportToExcel -volume $volume -server $server
    }	
}


#Synology


Start-Sleep -s 3

#Add file to Sharepoint
Add-PnPFile -Folder $onlinePath -Path ($localPath + $filename)