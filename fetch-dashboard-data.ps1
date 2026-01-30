# CTI Executive Dashboard - Data Fetcher
# Fetches data from SyteLine GL (Ledger) and generates dashboard JSON
# SOURCE OF TRUTH: SLLedgers table for all GL data
# Runs daily at 6am via Moltbot cron

param(
    [string]$OutputPath = "C:\Users\justi\clawd\dashboard\data"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Web

# Ensure output directory exists
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# SyteLine API Configuration
$SyteLineConfig = @{
    BaseUrl = "https://csi10g.erpsl.inforcloudsuite.com/IDORequestService/ido"
    Tenant = "GVNDYXUFKHB5VMB6_PRD_CTI"
    Username = "gary.phillips@godlan.com"
    Password = 'Crwthtithing2$'
}

# GL Account Mapping (SOURCE OF TRUTH)
$GLAccounts = @{
    "401000" = "Product"
    "402000" = "Service"
    "495000" = "Miscellaneous"
    "495400" = "Freight"
}

# US Federal Holidays 2025-2026
$Holidays = @(
    "2025-01-01", "2025-01-20", "2025-02-17", "2025-05-26", "2025-07-04",
    "2025-09-01", "2025-10-13", "2025-11-11", "2025-11-27", "2025-12-25",
    "2026-01-01", "2026-01-19", "2026-02-16", "2026-05-25", "2026-07-03",
    "2026-09-07", "2026-10-12", "2026-11-11", "2026-11-26", "2026-12-25"
) | ForEach-Object { [DateTime]::Parse($_) }

function Get-SyteLineToken {
    $tokenUrl = "$($SyteLineConfig.BaseUrl)/token/$($SyteLineConfig.Tenant)/$($SyteLineConfig.Username)/$([System.Web.HttpUtility]::UrlEncode($SyteLineConfig.Password))"
    $response = Invoke-RestMethod -Uri $tokenUrl
    if (-not $response.Success) {
        throw "Failed to get SyteLine token: $($response.Message)"
    }
    return $response.Token
}

function Invoke-SyteLineAPI {
    param(
        [string]$Token,
        [string]$IDO,
        [string]$Properties = "",
        [string]$Filter = "",
        [int]$RecordCap = 1000,
        [string]$OrderBy = ""
    )
    
    $uri = "$($SyteLineConfig.BaseUrl)/load/$IDO"
    $params = @("recordcap=$RecordCap")
    if ($Properties) { $params += "properties=$Properties" }
    if ($Filter) { $params += "filter=$([System.Web.HttpUtility]::UrlEncode($Filter))" }
    if ($OrderBy) { $params += "orderby=$([System.Web.HttpUtility]::UrlEncode($OrderBy))" }
    $uri += "?" + ($params -join "&")
    
    $headers = @{ "Authorization" = $Token }
    $response = Invoke-RestMethod -Uri $uri -Headers $headers
    
    if (-not $response.Success) {
        Write-Warning "API call failed for $IDO : $($response.Message)"
        return @()
    }
    return $response.Items
}

function Get-BusinessDays {
    param(
        [DateTime]$StartDate,
        [DateTime]$EndDate
    )
    $days = 0
    $current = $StartDate
    while ($current -le $EndDate) {
        if ($current.DayOfWeek -notin @('Saturday', 'Sunday') -and $current -notin $Holidays) {
            $days++
        }
        $current = $current.AddDays(1)
    }
    return $days
}

function Get-AdjustedSalesDate {
    # Weekend sales go to Monday
    param([DateTime]$Date)
    if ($Date.DayOfWeek -eq 'Saturday') { return $Date.AddDays(2) }
    if ($Date.DayOfWeek -eq 'Sunday') { return $Date.AddDays(1) }
    return $Date
}

# Main execution
Write-Host "$(Get-Date) - Starting dashboard data fetch from GL (SLLedgers)..."

try {
    $token = Get-SyteLineToken
    Write-Host "Got SyteLine token"
    
    $today = (Get-Date).Date
    $yesterday = $today.AddDays(-1)
    $monthStart = Get-Date -Day 1
    $yearStart = Get-Date -Month 1 -Day 1
    
    # Adjust yesterday if it was a weekend
    while ($yesterday.DayOfWeek -eq 'Sunday' -or $yesterday.DayOfWeek -eq 'Saturday') {
        $yesterday = $yesterday.AddDays(-1)
    }
    
    Write-Host "Fetching GL ledger data (revenue accounts)..."
    
    # Get ALL revenue transactions from GL Ledger for the current year
    $currentYear = (Get-Date).Year
    $filter = "(Acct = '401000' OR Acct = '402000' OR Acct = '495000' OR Acct = '495400') AND ControlYear >= $currentYear"
    $ledgerData = Invoke-SyteLineAPI -Token $token -IDO "SLLedgers" `
        -Properties "Acct,DomAmount,TransDate,Ref,ChaDescription" `
        -Filter $filter `
        -RecordCap 10000 `
        -OrderBy "TransDate DESC"
    
    Write-Host "Got $($ledgerData.Count) GL ledger records"
    
    # Process ledger data and aggregate by date and category
    $salesByDateCategory = @{}
    $dailyInvoices = @{}  # Track invoices per day for customer lookup
    
    foreach ($trans in $ledgerData) {
        $acct = $trans.Acct.Trim()
        if ($GLAccounts.ContainsKey($acct)) {
            $category = $GLAccounts[$acct]
            # Revenue is credit (negative in GL), convert to positive
            $amount = [decimal]$trans.DomAmount * -1
            
            # Parse transaction date
            if ($trans.TransDate -and $trans.TransDate.Length -ge 8) {
                try {
                    $transDate = [DateTime]::ParseExact($trans.TransDate.Substring(0,8), "yyyyMMdd", $null)
                    $salesDate = Get-AdjustedSalesDate $transDate
                    $dateKey = $salesDate.ToString("yyyy-MM-dd")
                    
                    # Initialize date entry if needed
                    if (-not $salesByDateCategory.ContainsKey($dateKey)) {
                        $salesByDateCategory[$dateKey] = @{
                            Product = 0
                            Service = 0
                            Freight = 0
                            Miscellaneous = 0
                            Total = 0
                        }
                    }
                    
                    $salesByDateCategory[$dateKey][$category] += $amount
                    $salesByDateCategory[$dateKey].Total += $amount
                    
                    # Track invoice numbers for customer lookup
                    if ($trans.Ref -match "ARI\s+(\d+)") {
                        $invNum = $matches[1].Trim()
                        if (-not $dailyInvoices.ContainsKey($dateKey)) {
                            $dailyInvoices[$dateKey] = @{}
                        }
                        if (-not $dailyInvoices[$dateKey].ContainsKey($invNum)) {
                            $dailyInvoices[$dateKey][$invNum] = $amount
                        } else {
                            $dailyInvoices[$dateKey][$invNum] += $amount
                        }
                    }
                } catch {
                    Write-Warning "Failed to parse date: $($trans.TransDate)"
                }
            }
        }
    }
    
    Write-Host "Processed data for $($salesByDateCategory.Count) unique dates"
    
    # Get invoice details for customer names
    Write-Host "Fetching invoice details for customer names..."
    $arinvs = Invoke-SyteLineAPI -Token $token -IDO "SLArinvs" `
        -Properties "InvNum,CustNum,CadbtName,InvDate,Type" `
        -RecordCap 5000
    
    # Build invoice to customer lookup
    $invoiceCustomers = @{}
    foreach ($inv in $arinvs) {
        if ($inv.Type -eq "I") {
            $invNum = $inv.InvNum.Trim()
            $invoiceCustomers[$invNum] = $inv.CadbtName
        }
    }
    
    # Get AR transactions for customer tracking using correct field names
    Write-Host "Fetching AR transaction history for complete customer data..."
    $arTrans = Invoke-SyteLineAPI -Token $token -IDO "SLArTrans" `
        -Properties "InvNum,CustNum,CadName,InvDate,Amount,Ref" `
        -Filter "InvDate >= '$($yearStart.ToString("yyyyMMdd"))'" `
        -RecordCap 10000
    
    Write-Host "Got $($arTrans.Count) AR transaction records"
    
    # Build customer sales tracking from AR transactions
    $salesByCustomer = @{}
    foreach ($trans in $arTrans) {
        if ($trans.InvDate -and $trans.Ref -match "^ARI") {  # Only invoice transactions
            try {
                $transDate = [DateTime]::ParseExact($trans.InvDate.Substring(0,8), "yyyyMMdd", $null)
                $salesDate = Get-AdjustedSalesDate $transDate
                $custName = $trans.CadName
                $amount = [decimal]$trans.Amount
                
                if ($custName -and $amount -gt 0) {
                    if (-not $salesByCustomer.ContainsKey($custName)) {
                        $salesByCustomer[$custName] = @{ MTD = 0; Yesterday = 0; YTD = 0 }
                    }
                    
                    if ($salesDate -ge $yearStart -and $salesDate -le $today) {
                        $salesByCustomer[$custName].YTD += $amount
                    }
                    if ($salesDate -ge $monthStart -and $salesDate -le $today) {
                        $salesByCustomer[$custName].MTD += $amount
                    }
                    if ($salesDate -eq $yesterday) {
                        $salesByCustomer[$custName].Yesterday += $amount
                    }
                }
            } catch {}
        }
    }
    
    # Get top products from order items
    Write-Host "Fetching order items for product analysis..."
    $coItems = Invoke-SyteLineAPI -Token $token -IDO "SLCoItems" `
        -Properties "DerItem,DerItemDescription,DerExtInvoicedPrice,DerDueDate" `
        -RecordCap 10000
    
    $salesByProduct = @{}
    foreach ($item in $coItems) {
        if ($item.DerExtInvoicedPrice -and [decimal]$item.DerExtInvoicedPrice -gt 0) {
            $prodKey = $item.DerItem.Trim()
            $amount = [decimal]$item.DerExtInvoicedPrice
            
            if (-not $salesByProduct.ContainsKey($prodKey)) {
                $salesByProduct[$prodKey] = @{
                    Description = $item.DerItemDescription
                    MTD = 0
                    YTD = 0
                    Yesterday = 0
                }
            }
            $salesByProduct[$prodKey].MTD += $amount
            $salesByProduct[$prodKey].YTD += $amount
        }
    }
    
    # Get shipment addresses for heat map
    Write-Host "Fetching shipment addresses..."
    $shipments = Invoke-SyteLineAPI -Token $token -IDO "SLShipments" `
        -Properties "ConsigneeCity,ConsigneeState,ConsigneeZip,ConsigneeCountry,ShipDate" `
        -RecordCap 3000 -OrderBy "ShipDate DESC"
    
    $locationCounts = @{}
    foreach ($ship in $shipments) {
        if ($ship.ConsigneeCity -and $ship.ConsigneeState) {
            $locKey = "$($ship.ConsigneeCity), $($ship.ConsigneeState)"
            if (-not $locationCounts.ContainsKey($locKey)) {
                $locationCounts[$locKey] = @{
                    City = $ship.ConsigneeCity
                    State = $ship.ConsigneeState
                    Zip = $ship.ConsigneeZip
                    Count = 0
                }
            }
            $locationCounts[$locKey].Count++
        }
    }
    $shippingLocations = $locationCounts.Values | Sort-Object -Property Count -Descending | Select-Object -First 100
    
    # Calculate summaries from GL data
    $yesterdaySales = if ($salesByDateCategory.ContainsKey($yesterday.ToString("yyyy-MM-dd"))) {
        $salesByDateCategory[$yesterday.ToString("yyyy-MM-dd")]
    } else {
        @{ Product = 0; Service = 0; Freight = 0; Miscellaneous = 0; Total = 0 }
    }
    
    $mtdSales = @{ Product = 0; Service = 0; Freight = 0; Miscellaneous = 0; Total = 0 }
    $ytdSales = @{ Product = 0; Service = 0; Freight = 0; Miscellaneous = 0; Total = 0 }
    
    foreach ($dateKey in $salesByDateCategory.Keys) {
        $date = [DateTime]::Parse($dateKey)
        $data = $salesByDateCategory[$dateKey]
        
        if ($date -ge $monthStart -and $date -le $today) {
            foreach ($cat in @("Product", "Service", "Freight", "Miscellaneous", "Total")) {
                $mtdSales[$cat] += $data[$cat]
            }
        }
        if ($date -ge $yearStart -and $date -le $today) {
            foreach ($cat in @("Product", "Service", "Freight", "Miscellaneous", "Total")) {
                $ytdSales[$cat] += $data[$cat]
            }
        }
    }
    
    # Calculate daily averages (excluding weekends/holidays)
    $mtdBusinessDays = Get-BusinessDays $monthStart $today
    $ytdBusinessDays = Get-BusinessDays $yearStart $today
    
    $mtdDailyAvg = if ($mtdBusinessDays -gt 0) { [math]::Round($mtdSales.Total / $mtdBusinessDays, 2) } else { 0 }
    $ytdDailyAvg = if ($ytdBusinessDays -gt 0) { [math]::Round($ytdSales.Total / $ytdBusinessDays, 2) } else { 0 }
    
    # Top customers and products
    $topCustomersMTD = $salesByCustomer.GetEnumerator() | 
        Where-Object { $_.Value.MTD -gt 0 } |
        Sort-Object { $_.Value.MTD } -Descending | 
        Select-Object -First 10 | 
        ForEach-Object { @{ Name = $_.Key; Amount = [math]::Round($_.Value.MTD, 2) } }
    
    $topCustomersYesterday = $salesByCustomer.GetEnumerator() | 
        Where-Object { $_.Value.Yesterday -gt 0 } |
        Sort-Object { $_.Value.Yesterday } -Descending | 
        Select-Object -First 10 | 
        ForEach-Object { @{ Name = $_.Key; Amount = [math]::Round($_.Value.Yesterday, 2) } }
    
    $topProductsMTD = $salesByProduct.GetEnumerator() | 
        Sort-Object { $_.Value.MTD } -Descending | 
        Select-Object -First 10 | 
        ForEach-Object { @{ Item = $_.Key; Description = $_.Value.Description; Amount = [math]::Round($_.Value.MTD, 2) } }
    
    # Daily trend data (last 30 days)
    $dailyTrend = @()
    for ($i = 30; $i -ge 0; $i--) {
        $date = $today.AddDays(-$i)
        $dateKey = $date.ToString("yyyy-MM-dd")
        $data = if ($salesByDateCategory.ContainsKey($dateKey)) { $salesByDateCategory[$dateKey] } 
                else { @{ Product = 0; Service = 0; Freight = 0; Miscellaneous = 0; Total = 0 } }
        $dailyTrend += @{
            Date = $dateKey
            DayOfWeek = $date.DayOfWeek.ToString()
            Product = [math]::Round($data.Product, 2)
            Service = [math]::Round($data.Service, 2)
            Freight = [math]::Round($data.Freight, 2)
            Miscellaneous = [math]::Round($data.Miscellaneous, 2)
            Total = [math]::Round($data.Total, 2)
        }
    }
    
    # Build final dashboard data
    $dashboardData = @{
        GeneratedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        DataSource = "SyteLine GL Ledger (SLLedgers) - Source of Truth"
        Period = @{
            Yesterday = $yesterday.ToString("yyyy-MM-dd")
            MonthStart = $monthStart.ToString("yyyy-MM-dd")
            YearStart = $yearStart.ToString("yyyy-MM-dd")
            MTDBusinessDays = $mtdBusinessDays
            YTDBusinessDays = $ytdBusinessDays
        }
        Summary = @{
            Yesterday = @{
                Product = [math]::Round($yesterdaySales.Product, 2)
                Service = [math]::Round($yesterdaySales.Service, 2)
                Freight = [math]::Round($yesterdaySales.Freight, 2)
                Miscellaneous = [math]::Round($yesterdaySales.Miscellaneous, 2)
                Total = [math]::Round($yesterdaySales.Total, 2)
            }
            MTD = @{
                Product = [math]::Round($mtdSales.Product, 2)
                Service = [math]::Round($mtdSales.Service, 2)
                Freight = [math]::Round($mtdSales.Freight, 2)
                Miscellaneous = [math]::Round($mtdSales.Miscellaneous, 2)
                Total = [math]::Round($mtdSales.Total, 2)
            }
            YTD = @{
                Product = [math]::Round($ytdSales.Product, 2)
                Service = [math]::Round($ytdSales.Service, 2)
                Freight = [math]::Round($ytdSales.Freight, 2)
                Miscellaneous = [math]::Round($ytdSales.Miscellaneous, 2)
                Total = [math]::Round($ytdSales.Total, 2)
            }
            MTDDailyAverage = $mtdDailyAvg
            YTDDailyAverage = $ytdDailyAvg
        }
        TopCustomers = @{
            Yesterday = $topCustomersYesterday
            MTD = $topCustomersMTD
        }
        TopProducts = @{
            MTD = $topProductsMTD
        }
        DailyTrend = $dailyTrend
        ShippingHeatMap = $shippingLocations
        ServiceTechLocations = @()  # Requires Salesforce
        TopSalesReps = @()  # Requires sales rep data
    }
    
    # Save data
    $outputFile = Join-Path $OutputPath "dashboard-data.json"
    $dashboardData | ConvertTo-Json -Depth 10 | Set-Content $outputFile -Encoding UTF8
    
    Write-Host "`n$(Get-Date) - Dashboard data saved to $outputFile"
    Write-Host "=============================================="
    Write-Host "SUMMARY (From GL - Source of Truth):"
    Write-Host "  Yesterday ($yesterday): `$$([math]::Round($yesterdaySales.Total, 2))"
    Write-Host "  MTD:                    `$$([math]::Round($mtdSales.Total, 2))"
    Write-Host "  YTD:                    `$$([math]::Round($ytdSales.Total, 2))"
    Write-Host "  MTD Daily Avg:          `$$mtdDailyAvg"
    Write-Host "  YTD Daily Avg:          `$$ytdDailyAvg"
    Write-Host "=============================================="
    
} catch {
    Write-Error "Dashboard data fetch failed: $_"
    throw
}
