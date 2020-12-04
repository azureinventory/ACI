##########################################################################################
#                                                                                        #
#                        * Azure Grinder Cost Report Generator *                         #
#                                                                                        #
#       Version: 0.0.1                                                                   #
#       Authors: Claudio Merola <clvieira@microsoft.com>                                 #
#                Renato Gregio <renato.gregio@microsoft.com>                             #
#                                                                                        #
#       Date: 12/03/2020                                                                 #
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


az login

$Subscriptions = az account list --output json --only-show-errors | ConvertFrom-Json

$Today = [DateTime]::Today.ToString("yyyy-MM-dd")

$LastMonth = [DateTime]::Today.AddDays(-30).ToString("yyyy-MM-dd")

Get-Job | Remove-Job

Foreach ($Subscription in $Subscriptions) {

        $Sub = $Subscription.id

       Start-Job -Name ('Costs'+$Sub) -ScriptBlock {

       az consumption usage list --subscription $($args[0]) --start-date $($args[2]) --end-date $($args[1]) --include-meter-details | ConvertFrom-Json

        } -ArgumentList $Sub,$Today,$LastMonth | Out-Null

        }

        Get-Job | Wait-Job  

        $Res = @()

        Foreach ($Subscription in $Subscriptions) {

        $Sub = $Subscription.id

        $TempSub = Receive-Job -Name ('Costs'+$Sub)

        $SubCosts = @{

                    'Subscription' = $TempSub
                    }


        $Res += $SubCosts

        }

        Foreach ($Subscri in $Res)
        {
            Foreach($1 in $Subscri.Values)
                {
                    $tmp = @()
                    Foreach ($0 in $1)
                        {
                                $obj = @{
                                    'Subscription' = $0.subscriptionName;
                                    'Resource Type' = $0.consumedService;
                                    'Resource' = $0.instanceName;
                                    'Location' = $0.instanceLocation;
                                    'Date' = [string]$0.usageStart.split('T')[0];
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


        $Style = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 'Currency' -Range J:J
        $Style2 = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 'Text' -Range H:H
        $Style3 = New-ExcelStyle -HorizontalAlignment Center -AutoSize -NumberFormat 'Date-Time' -Range E:E


        $tmp | 
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
            Export-Excel -Path 'C:\AzureGrinder\Costs.xlsx' -WorksheetName 'Costs' -AutoSize -TableName 'Costs' -TableStyle "Light20" -Style $Style,$Style2,$Style3