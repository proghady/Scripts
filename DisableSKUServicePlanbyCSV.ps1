#Install-Module -name Microsoft.Graph -Scope CurrentUser -Force
#Import-Module Microsoft.Graph

Connect-Graph -Scopes  "User.Read.All","Group.ReadWrite.All","Directory.Read.All","User.ReadWrite.All"
$CSVRecords = Import-CSV "C:\Temp\users.csv"
$i = 0;
$TotalRecords = $CSVRecords.Count
$UpdateResult = @()

#https://learn.microsoft.com/en-us/microsoftteams/sku-reference-edu
#Get-MgSubscribedSku | Select SkuPartNumber, SkuId

#MICROSOFTBOOKINGS

$A5Sku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq 'M365EDU_A5_FACULTY'
$disabledPlans = $A5Sku.ServicePlans | `
    Where ServicePlanName -in ("YAMMER_EDU") | `
    Select -ExpandProperty ServicePlanId

$addLicenses = @(
    @{
        SkuId = $A5Sku.SkuId
        DisabledPlans = $disabledPlans
    }
)

Foreach($CSVRecord in $CSVRecords){

$UserPrincipalName = $CSVRecord.userPrincipalName
 
$i++;
 
Write-Progress -activity "Processing $UserPrincipalName " -status "$i out of $TotalRecords users completed"

    try
    {
           
        
        Set-MgUserLicense -UserId $UserPrincipalName -AddLicenses $addLicenses -RemoveLicenses @()
        $ResetStatus = "Service Disabled"

    }#try
    catch
    {
        $ResetStatus = "Failed: $_"
    }#catch
    $UpdateResult += New-Object PSObject -property $([ordered]@{
            UserPrincipalName=$UserPrincipalName
            Status=$ResetStatus
        })#updateresult

}

#Display result
 
$UpdateResult | Select UserPrincipalName,Status| FT
#Export status report to a CSV file
$UpdateResult | Export-CSV -Path $env:USERPROFILE\$((Get-Date).ToString("yyyyMMdd_HHmmss"))"_ReportbyCSV.csv" -NoTypeInformation -Encoding UTF8

 
#Disconnect-MgGraph