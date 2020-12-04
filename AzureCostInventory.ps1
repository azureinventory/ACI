##########################################################################################
#                                                                                        #
#                        * Azure Grinder Cost Report Generator *                         #
#                                                                                        #
#       Version: 0.0.3                                                                   #
#       Authors: Claudio Merola <clvieira@microsoft.com>                                 #
#                Renato Gregio <renato.gregio@microsoft.com>                             #
#                                                                                        #
#       Date: 12/04/2020                                                                 #
#                                                                                        #
#           https://github.com/RenatoGregio/AzureCostInventory                           #
#                                                                                        #
#                                                                                        #
#        DISCLAIMER:                                                                     #
#        Please note that while being developed by Microsoft employees,                  #
#        Azure Grinder Inventory is not a Microsoft service or product.                  #
#                                                                                        #         
#        Azure Grinder Inventory is a personal driven project, there are none implicit   # 
#        or explicit obligations related to this project, it is provided 'as is' with    #
#        no warranties and confer no rights.                                             #
#                                                                                        #
##########################################################################################



if ($DebugPreference -eq 'Inquire') {
    $DebugPreference = 'Continue'
}

$ErrorActionPreference = "silentlycontinue"
$DesktopPath = "C:\AzureGrinder"
$CSPath = "$HOME/AzureGrinder"
$Global:Subscriptions = ''

<######################################### Help ################################################>

function usageMode() {
    Write-Output "" 
    Write-Output "" 
    Write-Output "Usage: "
    Write-Output "For CloudShell:"
    Write-Output "./AzureCostInventory.ps1"   
    Write-Output ""
    Write-Output "For PowerShell Desktop:"      
    Write-Output "./AzureCostInventory.ps1 -TenantID <Azure Tenant ID> "
    Write-Output "" 
    Write-Output "" 
}

<###################################################### Environment ######################################################################>

    function checkAzCli() {
        $azcli = az --version
        if ($null -eq $azcli) {
            throw "Azure Cli not found!"
            $host.Exit()
        }
        $azcliExt = az extension list --output json | ConvertFrom-Json
        if ($azcliExt.name -notin 'resource-graph') {
            az extension add --name resource-graph 
        }
        if ($null -eq (Get-InstalledModule -Name ImportExcel | Out-Null)) {
            Write-Debug ('ImportExcel Module is not installed, installing..')
            Install-Module -Name ImportExcel -Force
        }
    }

    function LoginSession() {
        $Global:DefaultPath = "$DesktopPath\"
        if ($TenantID -eq '' -or $null -eq $TenantID) {
            write-host "Tenant ID not specified. Use -TenantID parameter if you want to specify directly. "        
            write-host "Authenticating Azure"
            write-host ""
            az account clear | Out-Null
            az login | Out-Null
            write-host ""
            write-host ""
            $Tenants = az account list --query [].homeTenantId -o tsv --only-show-errors | Get-Unique
                
            if ($Tenants.Count -eq 1) {
                write-host "You have privileges only in One Tenant "
                write-host ""
                $TenantID = $Tenants
            }
            else { 
                write-host "Select the the Azure Tenant ID that you want to connect : "
                write-host ""
                $SequenceID = 1
                foreach ($TenantID in $Tenants) {
                    write-host "$SequenceID)  $TenantID"
                    $SequenceID ++ 
                }
                write-host ""
                [int]$SelectTenant = read-host "Select Tenant ( default 1 )"
                $defaultTenant = --$SelectTenant
                $TenantID = $Tenants[$defaultTenant]
            }
    
            write-host "Extracting from Tenant $TenantID"
            $Global:Subscriptions = az account list --output json --only-show-errors | ConvertFrom-Json
            $Global:Subscriptions = $Subscriptions | Where-Object { $_.tenantID -eq $TenantID }
            if($SubscriptionID)
                {
                    $Global:Subscriptions = $Subscriptions | Where-Object { $_.ID -eq $SubscriptionID}
                }
        }
    
        else {
            az account clear | Out-Null
            az login -t $TenantID | Out-Null
            $Global:Subscriptions = az account list --output json --only-show-errors | ConvertFrom-Json
            $Global:Subscriptions = $Subscriptions | Where-Object { $_.tenantID -eq $TenantID }
            if($SubscriptionID)
            {
                $Global:Subscriptions = $Subscriptions | Where-Object { $_.ID -eq $SubscriptionID}
            }
        }
    }

    function checkPS() {
        if ($PSVersionTable.PSEdition -eq 'Desktop') {
            $Global:PSEnvironment = "Desktop"
            write-host "PowerShell Desktop Identified."
            write-host ""
            LoginSession
        }
        else {
            $Global:PSEnvironment = "CloudShell"
            write-host 'Azure CloudShell Identified.'
            write-host ""
            <#### For Azure CloudShell change your StorageAccount Name, Container and SAS for Grid Extractor transfer. ####>
            $Global:DefaultPath = "$CSPath/" 
            $Global:Subscriptions = az account list --output json --only-show-errors | ConvertFrom-Json
        }
    }

    <###################################################### Checking PowerShell ######################################################################>

    checkPS

    <###################################################### Subscriptions ######################################################################>

    if ((Test-Path -Path $DefaultPath -PathType Container) -eq $false) {
        New-Item -Type Directory -Force -Path $DefaultPath | Out-Null
    }

####### Dates 

<#
$Today = [DateTime]::Today.ToString("yyyy-MM-dd")

$Today = Get-Date

$CurrentYear = $Todays.Year
$CurrentMonth = $Todays.Month -1

$startOfMonth = Get-Date -Year $CurrentYear -Month $CurrentMonth  -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0

$endOfMonth = ($startOfMonth).AddMonths(1).AddTicks(-1)

$Last2Month = [DateTime]::Today.AddDays(-60).ToString("yyyy-MM-dd")

$Last3Month = [DateTime]::Today.AddDays(-90).ToString("yyyy-MM-dd")

$TodayM = (Get-Culture).DateTimeFormat.GetMonthName([DateTime]::Today.ToString("MM"))
$LastMonthM = (Get-Culture).DateTimeFormat.GetMonthName([DateTime]::Today.AddDays(-30).ToString("MM"))
$Last2MonthM = (Get-Culture).DateTimeFormat.GetMonthName([DateTime]::Today.AddDays(-60).ToString("MM"))
$Last3MonthM = (Get-Culture).DateTimeFormat.GetMonthName([DateTime]::Today.AddDays(-90).ToString("MM"))

(Get-Culture).DateTimeFormat.GetMonthName($Today)

#>

Get-Job | Remove-Job

$Today = Get-Date
$Counter = 0
$Months = 3

While ($Counter -le $Months)
{

    if ($Counter -ge 1 -and $Today.Month -eq 1)
        {
            $CurrentYear = $Today.Year -1
        }
    else 
        {
            $CurrentYear = $Today.Year
        }
    $CurrentMonth = $Today.Month -($Counter)

    $StartOfMonth = Get-Date -Year $CurrentYear -Month $CurrentMonth  -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    $EndOfMonth = ($startOfMonth).AddMonths(1).AddTicks(-1)

    $StartOfMonth = $StartOfMonth.ToString("yyyy-MM-dd")
    $EndOfMonth = $EndOfMonth.ToString("yyyy-MM-dd")

    Start-Job -Name ('Cost Inventory'+$CurrentMonth) -ScriptBlock {
    
    $job = @()
 
    Foreach ($Subscription in $($args[0])) {
 
    $Sub = $Subscription.id
 
    New-Variable -Name ('SubRun'+$Sub)
 
    New-Variable -Name ('SubJob'+$Sub)
 
    Set-Variable -Name ('SubRun'+$Sub) -Value ([PowerShell]::Create()).AddScript({param($Sub,$EndOfMonth,$StartOfMonth)az consumption usage list --subscription $Sub --start-date $StartOfMonth --end-date $EndOfMonth --include-meter-details | ConvertFrom-Json}).AddArgument($Sub).AddArgument($($args[1])).AddArgument($($args[2]))
 
    Set-Variable -Name ('SubJob'+$Sub) -Value ((get-variable -name ('SubRun'+$Sub)).Value).BeginInvoke()

    $job += (get-variable -name ('SubJob'+$Sub)).Value
 
    }
 
     while ($Job.Runspace.IsCompleted -contains $false) {}
 
    Foreach ($Subscription in $($args[0])) {     
 
    $Sub = $Subscription.id
 
    New-Variable -Name ('SubValue'+$Sub)
 
    Set-Variable -Name ('SubValue'+$Sub) -Value (((get-variable -name ('SubRun'+$Sub)).Value).EndInvoke((get-variable -name ('SubJob'+$Sub)).Value))
 
    ((get-variable -name ('SubRun'+$Sub)).Value).Dispose()
 
    }
 
    $Result = @()
 
    Foreach ($Subscription in $($args[0])) {     
 
    $Sub = $Subscription.id

    $TempSub = (get-variable -name ('SubValue'+$Sub)).Value
 
    $Rest = @{
 
             'Subscription' = $TempSub
 
    }
 
    $Result += $Rest
 
    }
 
    $Result

        } -ArgumentList $Subscriptions,$EndOfMonth,$StartOfMonth | Out-Null


    $Counter ++

}


get-job | Wait-Job | Out-Null

$CostData = @()

$Counter = 0

While ($Counter -le $Months)
{
    $CurrentMonth = $Today.Month -($Counter)

    $CostDataTemp = Receive-Job -Name ('Cost Inventory'+$CurrentMonth) 
    $CostData += $CostDataTemp

    $Counter ++
}


Start-Job -Name 'Cost Processing' -ScriptBlock {
    
    $job = @()
    $JobCounter = 0
    $ProcCounter = 0
    $ResultCounter = 0

    Foreach ($Subscription in $args) {
        
        New-Variable -Name ('SubRun_'+$JobCounter)

        New-Variable -Name ('SubJob_'+$JobCounter)

        Set-Variable -Name ('SubRun_'+$JobCounter) -Value ([PowerShell]::Create()).AddScript({param($Sub)
            $tmp = @()
            Foreach($2 in $Sub.Values)
            {
                Foreach ($1 in $2)
                    {
                        Foreach ($0 in $1)
                            {
                                $Date = [string]$0.usageStart.split('Z')[0]
                                $Date = ([datetime]$Date).ToString("MM/dd/yyyy")
                                $obj = @{
                                    'Subscription' = $0.subscriptionName;
                                    'Resource Type' = $0.consumedService;
                                    'Resource' = $0.instanceName;
                                    'Location' = $0.instanceLocation;
                                    'Date' = $Date;
                                    'Usage Detail' = $0.meterDetails.meterCategory;
                                    'Usage Operation' = $0.meterDetails.meterName;
                                    'Usage Time (Hours)' = [decimal]$0.usageQuantity;
                                    'Currency' = $0.currency;
                                    'Cost' = [decimal]$0.pretaxCost;  
                                }
                                $tmp += $obj
                            }
                     }
            }
            
            $tmp
        }).AddArgument($Subscription)

    Set-Variable -Name ('SubJob_'+$JobCounter) -Value (((get-variable -name ('SubRun_'+$JobCounter)).Value).BeginInvoke())

    $job += (get-variable -name ('SubJob_'+$JobCounter)).Value
 
    $JobCounter ++

    }
 
    while ($Job.Runspace.IsCompleted -contains $false) {}
 
    Foreach ($Subscription in $args) {     
 
    New-Variable -Name ('SubValue_'+$ProcCounter)
 
    Set-Variable -Name ('SubValue_'+$ProcCounter) -Value (((get-variable -name ('SubRun_'+$ProcCounter)).Value).EndInvoke((get-variable -name ('SubJob_'+$ProcCounter)).Value))

    ((get-variable -name ('SubRun_'+$ProcCounter)).Value).Dispose()
 
    $ProcCounter ++

    }
 
    $Result = @()
 
    Foreach ($Subscription in $args) {     

    $TempSub = (get-variable -name ('SubValue_'+$ResultCounter)).Value
 
    Add-Content -Path 'C:\Temp\File.txt' -Value $ResultCounter

    $Result += $TempSub
 
    $ResultCounter ++

    }
 
    $Result

        } -ArgumentList $CostData | Out-Null



    Get-Job -Name 'Cost Processing' | Wait-Job | Out-Null

    $ResultData = Receive-Job -Name 'Cost Processing'

        $Style0 = New-ExcelStyle -HorizontalAlignment Center -AutoSize 
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 'Currency' -Range J:J
        $Style2 = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 'Text' -Range H:H
        $Style3 = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 'Date-Time' -Range E:E

        $ResultData | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object 'Subscription',
            'Resource Type',
            'Resource',
            'Location',
            'Date',
            'Usage Detail',
            'Usage Operation',
            'Usage Time (Hours)',
            'Currency',
            'Cost'| 
            Export-Excel -Path 'C:\AzureGrinder\Costs.xlsx' -WorksheetName 'Costs' -AutoSize -TableName 'Costs' -TableStyle "Light20" -Style $Style0,$Style,$Style2,$Style3



            <############################################################   CHARTS  ###############################################################################>



