##########################################################################################
#                                                                                        #
#                      * Azure Cost Inventory Report Generator *                         #
#                                                                                        #
#       Version: 0.0.63                                                                  #
#       Authors: Claudio Merola <clvieira@microsoft.com>                                 #
#                Renato Gregio <renato.gregio@microsoft.com>                             #
#                                                                                        #
#       Date: 12/11/2020                                                                 #
#                                                                                        #
#           https://github.com/RenatoGregio/AzureCostInventory                           #
#                                                                                        #
#                                                                                        #
#        DISCLAIMER:                                                                     #
#        Please note that while being developed by Microsoft employees,                  #
#        Azure Cost Inventory is not a Microsoft service or product.                     #
#                                                                                        #         
#        Azure Cost Inventory is a personal driven project, there are none implicit      # 
#        or explicit obligations related to this project, it is provided 'as is' with    #
#        no warranties and confer no rights.                                             #
#                                                                                        #
##########################################################################################


param ($TenantID, $PassedMonhs = 2,$SubscriptionID) 

if ($DebugPreference -eq 'Inquire') {
    $DebugPreference = 'Continue'
}

$ErrorActionPreference = "silentlycontinue"
$DesktopPath = "C:\AzureInventory"
$CSPath = "$HOME/AzureInventory"
$Global:tableStyle = "Light20"
$Global:Subscriptions = ''
$Global:Today = Get-Date
$Global:Months = $PassedMonhs

$Runtime = Measure-Command {

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

function Extractor 
    {
        function checkAzCli() 
            {
                $azcli = az --version
                if ($null -eq $azcli) {
                    throw "Azure Cli not found!"
                    $host.Exit()
                }
                $azcliExt = az extension list --output json | ConvertFrom-Json
                if ($azcliExt.name -notin 'costmanagement') {
                    az extension add --name costmanagement
                }
                $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
                if ($null -eq (Get-InstalledModule -Name ImportExcel | Out-Null)) {
                    if($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
                    {
                        Install-Module -Name ImportExcel -Force
                    }
                    else 
                    {
                        Write-Host 'Impossible to install ImportExcel Module if not running as Admin'
                        Write-Host ''
                        Write-Host 'Exiting now.'
                        $host.Exit()
                    }
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
                    az login -t $TenantID | Out-Null
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
                    az login -t $TenantID | Out-Null
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

        <###################################################### Path ######################################################################>

        if ((Test-Path -Path $DefaultPath -PathType Container) -eq $false) {
            New-Item -Type Directory -Force -Path $DefaultPath | Out-Null
        }
    }

    function Inventory {

        $Global:File = ($DefaultPath + "AzureCostInventory_Report_" + (get-date -Format "yyyy-MM-dd_HH_mm") + ".xlsx")

        <###################################################### JOBs ######################################################################>

        Get-Job | Remove-Job

        Write-host ('Starting First Jobs')

        Start-Job -Name 'Resource Group Inventory' -ScriptBlock {
            
            Foreach ($Subscription in $args)
                {   
                    $Sub = $Subscription.id

                    New-Variable -Name ('SubRun'+$Sub)
                    New-Variable -Name ('SubJob'+$Sub)

                    Set-Variable -Name ('SubRun'+$Sub) -Value ([PowerShell]::Create()).AddScript({param($Sub)
                        az group list --subscription $Sub | ConvertFrom-Json
                    }).AddArgument($Sub)
                    
                    Set-Variable -Name ('SubJob'+$Sub) -Value ((get-variable -name ('SubRun'+$Sub)).Value).BeginInvoke()

                    $job += (get-variable -name ('SubJob'+$Sub)).Value
                }

            while ($Job.Runspace.IsCompleted -contains $false) {}

            Foreach ($Subscription in $args) 
                {     
                $Sub = $Subscription.id
             
                New-Variable -Name ('SubValue'+$Sub)
                Set-Variable -Name ('SubValue'+$Sub) -Value (((get-variable -name ('SubRun'+$Sub)).Value).EndInvoke((get-variable -name ('SubJob'+$Sub)).Value))

                }

            $Result = @()

            Foreach ($Subscription in $args) 
                {     
                    $Sub = $Subscription.id
                    $Result += (get-variable -name ('SubValue'+$Sub)).Value
                }
                
            $Result
            } -ArgumentList $Subscriptions

            Get-Job | Wait-Job | Out-Null

            $ResourceGroups = Receive-Job -Name 'Resource Group Inventory'

        $EndDate = Get-Date -Year $Today.Year -Month $Today.Month -Day ($Today.Day -1) -Hour 23 -Minute 59 -Second 59 -Millisecond 0
        $EndDateMode = Get-Date -Year $Today.Year -Month $Today.Month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
        $StartDate = ($EndDateMode).AddMonths(-$Months)

        $EndDate = ($EndDate.ToString("yyyy-MM-dd")+'T23:59:59').ToString()
        $StartDate = ($StartDate.ToString("yyyy-MM-dd")+'T00:00:00').ToString()
       
        Foreach ($Subscription in $Subscriptions)
            { 

                Start-Job -Name ('Usage Inventory'+$Subscription.id) -ScriptBlock {
            
                $Dateset = @'
"{\"totalCost\":{\"name\":\"PreTaxCost\",\"function\":\"Sum\"}}"
'@
                $Sub = $($args[0]).id
                $RGS = @()
                Foreach($ResourceG in $($args[3]))
                    {
                        $RGS +=  $ResourceG | Where-Object {$_.id.split('/')[2] -eq $Sub}
                    }

                        foreach ($rg in $RGS)
                            {
                                        $Scope = $rg.ID

                                        New-Variable -Name ('SubRun'+$rg.Name)
                                        New-Variable -Name ('SubJob'+$rg.Name)

                                        Set-Variable -Name ('SubRun'+$rg.Name) -Value ([PowerShell]::Create()).AddScript({param($Scope,$StartDate,$EndDate,$Dateset)
                                            az costmanagement query --type "Usage" --dataset-aggregation $Dateset --dataset-grouping name="ResourceGroup" type="Dimension" --timeframe "Custom" --time-period from=$StartDate to=$EndDate --scope $Scope | ConvertFrom-Json
                                        }).AddArgument($Scope).AddArgument($($args[1])).AddArgument($($args[2])).AddArgument($Dateset)
                                        
                                        Set-Variable -Name ('SubJob'+$rg.Name) -Value ((get-variable -name ('SubRun'+$rg.Name)).Value).BeginInvoke()

                                        $job += (get-variable -name ('SubJob'+$rg.Name)).Value
                            }

                while ($Job.Runspace.IsCompleted -contains $false) {}

                        foreach ($rg in $RGS)
                            {   
                                        New-Variable -Name ('SubValue'+$rg.Name)
                                        Set-Variable -Name ('SubValue'+$rg.Name) -Value (((get-variable -name ('SubRun'+$rg.Name)).Value).EndInvoke((get-variable -name ('SubJob'+$rg.Name)).Value))
                            }        

                $Result = @()

                        foreach ($rg in $RGS)
                            {
                                        $Results = (get-variable -name ('SubValue'+$rg.Name)).Value
                                        $obj = @{
                                                'ID' = $rg.id;
                                                'Subscription' = $($args[0]).name;
                                                'Resource Group' = $rg.name;
                                                'Location' = $rg.Location;
                                                'Usage' = $Results
                                                }
                                $Result += $obj
                            }
                $Result
                } -ArgumentList $Subscription,$StartDate,$EndDate,$ResourceGroups | Out-Null
            }
            Write-host ('Waiting First Jobs')
            Get-Job | Wait-Job | Out-Null
        }

    function DataProcessor 
    {
        Write-host ('Starting Second Jobs')

        Foreach ($Subscription in $Subscriptions)
            { 
                $InvSub = Receive-Job -Name ('Usage Inventory'+$Subscription.id)

                Start-Job -Name ('Cost Processing'+$Subscription.id) -ScriptBlock {

                        $tmp = @()
                        Foreach ($RG in $args) 
                            {
                                if($RG.Usage.Rows)
                                    {
                                        $SubName = $RG.Subscription
                                        $ResourceGroup = $RG.'Resource Group'
                                        $Location = $RG.Location
                                        $ID = $RG.ID
                                        if($RG.Usage.rows[1].count -eq 4)
                                            {
                                                Foreach ($Row in $RG.Usage.rows)
                                                    {
                                                        $Date0 = [datetime]::ParseExact($Row[1], 'yyyyMMdd', $null)
                                                        $WeekDay = $Date0.DayOfWeek
                                                        $Date = (([datetime]$Date0).ToString("MM/dd/yyyy")).ToString()
                                                        $DateMonth = ((Get-Culture).DateTimeFormat.GetMonthName(([datetime]$Date0).ToString("MM"))).ToString()
                                                        $DateYear = (([datetime]$Date0).ToString("yyyy")).ToString()

                                                        $obj = @{
                                                                'ID' = $ID;
                                                                'Subscription' = $SubName;
                                                                'Resource Group' = $ResourceGroup;
                                                                'Location' = $Location;
                                                                'Date' = $Date;
                                                                'Day of Week' = [string]$WeekDay;
                                                                'Month' = $DateMonth;
                                                                'Year' = $DateYear;
                                                                'Currency' = $Row[3];
                                                                'Cost' = '{0:C}' -f $Row[0]  
                                                            }
                                                        $tmp += $obj
                                                    }
                                            }
                                        else 
                                            {
                                                $Date0 = [datetime]::ParseExact($RG.Usage.Rows[1], 'yyyyMMdd', $null)
                                                $WeekDay = $Date0.DayOfWeek
                                                $Date = (([datetime]$Date0).ToString("MM/dd/yyyy")).ToString()
                                                $DateMonth = ((Get-Culture).DateTimeFormat.GetMonthName(([datetime]$Date0).ToString("MM"))).ToString()
                                                $DateYear = (([datetime]$Date0).ToString("yyyy")).ToString()

                                                $obj = @{
                                                        'ID' = $ID;
                                                        'Subscription' = $SubName;
                                                        'Resource Group' = $ResourceGroup;
                                                        'Location' = $Location;
                                                        'Date' = $Date;
                                                        'Day of Week' = [string]$WeekDay;
                                                        'Month' = $DateMonth;
                                                        'Year' = $DateYear;
                                                        'Currency' = $RG.Usage.Rows[3];
                                                        'Cost' = '{0:C}' -f $RG.Usage.Rows[0]  
                                                    }
                                                $tmp += $obj
                                            }
                                    }   
                            }
                        $tmp
                    } -ArgumentList $InvSub | Out-Null
            }

        Write-host ('Waiting Second Jobs')
        Get-Job | Wait-Job | Out-Null
    }

function DataConsolidation 
    {
        Get-Job | Wait-Job | Out-Null
        Write-host ('Consolidating Data')
        $Global:Data = @()
        Foreach ($Subscription in $Subscriptions)
            { 
                $DataTemp = Receive-Job -Name ('Cost Processing'+$Subscription.id)
                $Global:Data += $DataTemp
            }
    }


function Report 
    {

        Write-host ('Starting Reporting')

        $StyleOver = New-ExcelStyle -Range A1:AF1 -Bold -FontSize 28 -BackgroundColor ([System.Drawing.Color]::YellowGreen) -Merge -HorizontalAlignment Center
        ('Currency: '+$Data.currency[0]) | Export-Excel -Path $File -WorksheetName 'Overview' -Style $StyleOver -MoveToStart -KillExcel

        $Style0 = New-ExcelStyle -HorizontalAlignment Center -AutoSize 
        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 'Currency' -Range I:I
        $Style3 = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 'Date-Time' -Range D:D

        $Data | 
            ForEach-Object { [PSCustomObject]$_ } | 
            Select-Object 'Subscription',
            'Resource Group',
            'Location',
            'Date',
            'Day of Week',
            'Month',
            'Year',
            'Currency',
            'Cost'| 
            Export-Excel -Path $File -WorksheetName 'Usage' -AutoSize -TableName 'Usage' -TableStyle $tableStyle -Style $Style0,$Style,$Style3



            <############################################################   CHARTS  ###############################################################################>

            Write-Debug ('Starting to Generate Charts..')
            $excel = Open-ExcelPackage -Path $File -KillExcel


            $PTParams = @{
                PivotTableName    = "P0"
                Address           = $excel.Overview.cells["A6"] # top-left corner of the table
                SourceWorkSheet   = $excel.Usage
                PivotRows         = @("Month")
                PivotData         = @{"Cost" = "sum" }
                PivotTableStyle   = $tableStyle
                IncludePivotChart = $true
                ChartType         = "ColumnClustered3D"
                ShowCategory      = $false
                ChartRow          = 2 # place the chart below row 22nd
                ChartColumn       = 3
                PivotFilter       = 'Subscription', 'Resource Group'
                ChartTitle        = ('Cost by Month (Currency: '+$Data.currency[0]+')')
                Activate          = $true
                NoLegend          = $true
                ShowPercent       = $true
                ChartHeight       = 500
                ChartWidth        = 500
                PivotNumberFormat = "Currency"
            }
    
            Add-PivotTable @PTParams

            $PTParams = @{
                PivotTableName    = "P1"
                Address           = $excel.Overview.cells["A32"] # top-left corner of the table
                SourceWorkSheet   = $excel.Usage
                PivotRows         = @("Subscription")
                PivotData         = @{"Cost" = "sum" }
                PivotTableStyle   = $tableStyle
                IncludePivotChart = $true
                ChartType         = "Pie3D"
                ChartRow          = 28 # place the chart below row 22nd
                ChartColumn       = 3
                Activate          = $true
                PivotFilter       = 'Month', 'Resource Group'
                ChartTitle        = ('Cost by Subscription (Currency: '+$Data.currency[0]+')')
                NoLegend          = $true
                ShowPercent       = $true
                ShowCategory      = $false
                ChartHeight       = 500
                ChartWidth        = 500
                PivotNumberFormat = "Currency"
            }
    
            Add-PivotTable @PTParams

            $PTParams = @{
                PivotTableName    = "P2"
                Address           = $excel.Overview.cells["M6"] # top-left corner of the table
                SourceWorkSheet   = $excel.Usage
                PivotRows         = @("Location")
                PivotData         = @{"Cost" = "Sum" }
                PivotTableStyle   = $tableStyle
                IncludePivotChart = $true
                ChartType         = "BarStacked3D"
                ChartRow          = 2 # place the chart below row 22nd
                ChartColumn       = 15
                Activate          = $true
                ChartTitle        = ('Cost by Location (Currency: '+$Data.currency[0]+')')
                PivotFilter       = 'Month', 'Subscription'
                ShowPercent       = $true
                ChartHeight       = 500
                ChartWidth        = 500
                PivotNumberFormat = "Currency"
                NoLegend          = $true
            }
    
            Add-PivotTable @PTParams

            $PTParams = @{
                PivotTableName    = "P3"
                Address           = $excel.Overview.cells["M32"] # top-left corner of the table
                SourceWorkSheet   = $excel.Usage
                PivotRows         = @("Day of Week")
                PivotData         = @{"Cost" = "sum" }
                PivotTableStyle   = $tableStyle
                IncludePivotChart = $true
                ChartType         = "ColumnClustered3D"
                ChartRow          = 28 # place the chart below row 22nd
                ChartColumn       = 15
                Activate          = $true
                PivotFilter       = 'Month', 'Subscription'
                ChartTitle        = ('Cost by Day of Week (Currency: '+$Data.currency[0]+')')
                NoLegend          = $true
                ShowPercent       = $true
                ShowCategory      = $false
                ChartHeight       = 500
                ChartWidth        = 500
                PivotNumberFormat = "Currency"
            }
    
            Add-PivotTable @PTParams
            
            Close-ExcelPackage $excel 

    }            


Extractor
Inventory
DataProcessor
DataConsolidation 
Report


}
$Measure = $Runtime.Totalminutes.ToString('#######.##')

Write-Host ('Report Complete. Total Runtime was: ' + $Measure + ' Minutes')

Write-Host ''
Write-Host ('Excel file saved at: ') -NoNewline
write-host $File -ForegroundColor Cyan
Write-Host ''