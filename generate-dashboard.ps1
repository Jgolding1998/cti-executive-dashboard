# CTI Executive Dashboard Generator
# Generates a self-contained HTML dashboard with embedded data
# Uses itmuf_cti_product_category field for Service/Product classification
# SOURCE OF TRUTH: GL Ledger with invoice-level categorization

param(
    [string]$OutputPath = "C:\Users\justi\clawd\dashboard"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Web

# SyteLine API Configuration - Load from config file or environment variables
$configPath = Join-Path $PSScriptRoot "config.json"
if (Test-Path $configPath) {
    $config = Get-Content $configPath | ConvertFrom-Json
    $SyteLineConfig = @{
        BaseUrl = $config.syteline.baseUrl
        Tenant = $config.syteline.tenant
        Username = $config.syteline.username
        Password = $config.syteline.password
    }
} elseif ($env:SYTELINE_USERNAME -and $env:SYTELINE_PASSWORD) {
    # Use environment variables (for GitHub Actions)
    $SyteLineConfig = @{
        BaseUrl = if ($env:SYTELINE_BASEURL) { $env:SYTELINE_BASEURL } else { "https://csi10g.erpsl.inforcloudsuite.com/IDORequestService/ido" }
        Tenant = if ($env:SYTELINE_TENANT) { $env:SYTELINE_TENANT } else { "GVNDYXUFKHB5VMB6_PRD_CTI" }
        Username = $env:SYTELINE_USERNAME
        Password = $env:SYTELINE_PASSWORD
    }
} else {
    throw "No credentials found. Create config.json or set SYTELINE_USERNAME/SYTELINE_PASSWORD environment variables."
}

# Service Product Codes (from Alex's definition - see Teams chat 1/16)
# LFTR = Service - Field Labor
# LGAS = Service - Gas
# LIHL = Service - In-house Labor
$ServiceProductCodes = @("LFTR", "LGAS", "LIHL")

# US Federal Holidays
$Holidays = @(
    "2025-01-01", "2025-01-20", "2025-02-17", "2025-05-26", "2025-07-04",
    "2025-09-01", "2025-10-13", "2025-11-11", "2025-11-27", "2025-12-25",
    "2026-01-01", "2026-01-19", "2026-02-16", "2026-05-25", "2026-07-03",
    "2026-09-07", "2026-10-12", "2026-11-11", "2026-11-26", "2026-12-25"
) | ForEach-Object { [DateTime]::Parse($_) }

function Get-SyteLineToken {
    $tokenUrl = "$($SyteLineConfig.BaseUrl)/token/$($SyteLineConfig.Tenant)/$($SyteLineConfig.Username)/$([System.Web.HttpUtility]::UrlEncode($SyteLineConfig.Password))"
    $response = Invoke-RestMethod -Uri $tokenUrl
    if (-not $response.Success) { throw "Failed to get SyteLine token" }
    return $response.Token
}

function Invoke-SyteLineAPI {
    param([string]$Token, [string]$IDO, [string]$Properties = "", [string]$Filter = "", [int]$RecordCap = 1000, [string]$OrderBy = "")
    $uri = "$($SyteLineConfig.BaseUrl)/load/$IDO"
    $params = @("recordcap=$RecordCap")
    if ($Properties) { $params += "properties=$Properties" }
    if ($Filter) { $params += "filter=$([System.Web.HttpUtility]::UrlEncode($Filter))" }
    if ($OrderBy) { $params += "orderby=$([System.Web.HttpUtility]::UrlEncode($OrderBy))" }
    $uri += "?" + ($params -join "&")
    $headers = @{ "Authorization" = $Token }
    $response = Invoke-RestMethod -Uri $uri -Headers $headers
    if (-not $response.Success) { Write-Warning "API failed for $IDO"; return @() }
    return $response.Items
}

function Get-BusinessDays { param([DateTime]$Start, [DateTime]$End)
    $days = 0; $cur = $Start
    while ($cur -le $End) { if ($cur.DayOfWeek -notin @('Saturday','Sunday') -and $cur -notin $Holidays) { $days++ }; $cur = $cur.AddDays(1) }
    return $days
}

function Get-AdjustedDate { param([DateTime]$Date)
    if ($Date.DayOfWeek -eq 'Saturday') { return $Date.AddDays(2) }
    if ($Date.DayOfWeek -eq 'Sunday') { return $Date.AddDays(1) }
    return $Date
}

function Get-WeekRange { param([DateTime]$Date)
    $monday = $Date.AddDays(-([int]$Date.DayOfWeek - 1))
    if ($Date.DayOfWeek -eq 'Sunday') { $monday = $Date.AddDays(-6) }
    $friday = $monday.AddDays(4)
    return @{ Start = $monday; End = $friday; Label = "Week of $($monday.ToString('MMM d'))" }
}

Write-Host "$(Get-Date) - Generating CTI Executive Dashboard..."

$token = Get-SyteLineToken
Write-Host "Connected to SyteLine"

$today = (Get-Date).Date
$yesterday = $today.AddDays(-1)
while ($yesterday.DayOfWeek -in @('Saturday','Sunday')) { $yesterday = $yesterday.AddDays(-1) }
$monthStart = Get-Date -Day 1
$yearStart = Get-Date -Month 1 -Day 1
$currentYear = (Get-Date).Year

# === STEP 1: Build item ProductCode lookup ===
Write-Host "Fetching item master for service/product categorization..."
$items = Invoke-SyteLineAPI -Token $token -IDO "SLItems" -Properties "Item,ProductCode" -RecordCap 5000
$itemProductCode = @{}
foreach ($item in $items) {
    if ($item.Item) { $itemProductCode[$item.Item.Trim()] = $item.ProductCode }
}
Write-Host "Built ProductCode lookup for $($itemProductCode.Count) items"

# === STEP 2: Get line items and build order breakdown ===
Write-Host "Fetching line items..."
$coItems = Invoke-SyteLineAPI -Token $token -IDO "SLCoItems" -Properties "CoNum,Item,ItDescription,DerExtInvoicedPrice" -RecordCap 15000
Write-Host "Got $($coItems.Count) line items"

$orderBreakdown = @{}
$salesByProduct = @{}

foreach ($line in $coItems) {
    if ($line.CoNum -and $line.Item -and $line.DerExtInvoicedPrice -and [decimal]$line.DerExtInvoicedPrice -gt 0) {
        $orderNum = $line.CoNum.Trim()
        $itemKey = $line.Item.Trim()
        $amount = [decimal]$line.DerExtInvoicedPrice
        $prodCode = $itemProductCode[$itemKey]
        
        # Service = ProductCode in (LFTR, LGAS, LIHL)
        $isService = $prodCode -in $ServiceProductCodes
        
        # Order breakdown
        if (-not $orderBreakdown.ContainsKey($orderNum)) {
            $orderBreakdown[$orderNum] = @{ Service = 0; Product = 0; Total = 0 }
        }
        if ($isService) { $orderBreakdown[$orderNum].Service += $amount }
        else { $orderBreakdown[$orderNum].Product += $amount }
        $orderBreakdown[$orderNum].Total += $amount
        
        # Product sales
        if (-not $salesByProduct.ContainsKey($itemKey)) {
            $salesByProduct[$itemKey] = @{ Description = $line.ItDescription; MTD = 0; Category = if ($isService) { "Service" } else { "Product" } }
        }
        $salesByProduct[$itemKey].MTD += $amount
    }
}
Write-Host "Built breakdown for $($orderBreakdown.Count) orders, $($salesByProduct.Count) products"

# === STEP 3: Get invoice to order mapping and salesperson data ===
Write-Host "Fetching AR transactions..."
$arTrans = Invoke-SyteLineAPI -Token $token -IDO "SLArTrans" -Properties "InvNum,CoNum,CadName,InvDate,DueDate,Amount,Type,ApplyToInvNum" -RecordCap 10000
Write-Host "Got $($arTrans.Count) AR transaction records"

$invToOrder = @{}
$invToCustomer = @{}

# Build invoice balances: Original Amount - Payments - Credits
$invoiceBalances = @{}

# First pass: Add invoices (Type I)
foreach ($ar in $arTrans) {
    if ($ar.Type -eq "I" -and $ar.InvNum) {
        $invNum = $ar.InvNum.Trim()
        if (-not $invoiceBalances.ContainsKey($invNum)) {
            $invoiceBalances[$invNum] = @{ Customer = $ar.CadName; DueDate = $ar.DueDate; CoNum = $ar.CoNum; OriginalAmount = 0; Applied = 0 }
        }
        $invoiceBalances[$invNum].OriginalAmount += [decimal]$ar.Amount
        if ($ar.CoNum) { $invToOrder[$invNum] = $ar.CoNum.Trim() }
        if ($ar.CadName) { $invToCustomer[$invNum] = $ar.CadName }
    }
}

# Second pass: Apply payments (Type P) and credits (Type C)
foreach ($ar in $arTrans) {
    if ($ar.Type -in @("P", "C") -and $ar.Amount) {
        $applyTo = if ($ar.ApplyToInvNum) { $ar.ApplyToInvNum.Trim() } else { if ($ar.InvNum) { $ar.InvNum.Trim() } else { $null } }
        if ($applyTo -and $invoiceBalances.ContainsKey($applyTo)) {
            $invoiceBalances[$applyTo].Applied += [decimal]$ar.Amount
        }
    }
}

# Build open AR list with remaining balances
$openAR = @()
foreach ($inv in $invoiceBalances.GetEnumerator()) {
    $balance = $inv.Value.OriginalAmount - $inv.Value.Applied
    if ($balance -gt 1 -and $inv.Value.DueDate) {  # >$1 threshold for rounding
        try {
            $dueDate = [DateTime]::ParseExact($inv.Value.DueDate.Substring(0,8), "yyyyMMdd", $null)
            $openAR += @{ 
                InvNum = $inv.Key
                Customer = $inv.Value.Customer
                Amount = $balance
                DueDate = $dueDate
                DaysOverdue = ($today - $dueDate).Days
            }
        } catch {}
    }
}
Write-Host "Found $($openAR.Count) open invoices (after payments), $($invToCustomer.Count) invoice-customer mappings"

# === STEP 3.5: Get Order TakenBy (Salesperson) data with Territory Info ===
Write-Host "Fetching order entry data for salesperson and territory tracking..."
$orderFilter = "Invoiced = 1 AND OrderDate >= '$($yearStart.ToString("yyyyMMdd"))'"
$orders = Invoke-SyteLineAPI -Token $token -IDO "SLCoS" -Properties "CoNum,TakenBy,Price,OrderDate,Invoiced,ShipToCity,ShipToState" -Filter $orderFilter -RecordCap 5000
Write-Host "Got $($orders.Count) invoiced orders"

# Build order to salesperson mapping AND territory data
$orderToSalesperson = @{}
$territoryByPerson = @{}
foreach ($order in $orders) {
    if ($order.CoNum -and $order.TakenBy) {
        $takenBy = $order.TakenBy.Trim()
        $orderToSalesperson[$order.CoNum.Trim()] = $takenBy
        
        # Track territory data
        if (-not $territoryByPerson.ContainsKey($takenBy)) {
            $territoryByPerson[$takenBy] = @{ States=@{}; Cities=@{}; Customers=@{}; OrderCount=0 }
        }
        if ($order.ShipToState) {
            $territoryByPerson[$takenBy].States[$order.ShipToState.Trim()] = $true
            if ($order.ShipToCity) {
                $city = "$($order.ShipToCity.Trim()), $($order.ShipToState.Trim())"
                $territoryByPerson[$takenBy].Cities[$city] = $true
            }
        }
        $territoryByPerson[$takenBy].OrderCount++
    }
}
Write-Host "Mapped $($orderToSalesperson.Count) orders to salespeople with territory data"

# === STEP 4: Get GL Ledger and calculate sales by date with proper categorization ===
Write-Host "Fetching GL ledger data..."
$filter = "(Acct = '401000' OR Acct = '402000' OR Acct = '495000' OR Acct = '495400') AND ControlYear >= $currentYear"
$ledgerData = Invoke-SyteLineAPI -Token $token -IDO "SLLedgers" -Properties "Acct,DomAmount,TransDate,Ref" -Filter $filter -RecordCap 10000
Write-Host "Got $($ledgerData.Count) GL ledger records"

$salesByDate = @{}
$salesByCustomer = @{}
$salesBySalesperson = @{}  # NEW: Track sales by salesperson

foreach ($gl in $ledgerData) {
    $acct = $gl.Acct.Trim()
    if ($gl.TransDate.Length -ge 8) {
        try {
            $transDate = [DateTime]::ParseExact($gl.TransDate.Substring(0,8), "yyyyMMdd", $null)
            $salesDate = Get-AdjustedDate $transDate
            $dateKey = $salesDate.ToString("yyyy-MM-dd")
            $amount = [decimal]$gl.DomAmount * -1  # Revenue is credit
            
            if (-not $salesByDate.ContainsKey($dateKey)) {
                $salesByDate[$dateKey] = @{ Product=0; Service=0; Freight=0; Miscellaneous=0; Total=0 }
            }
            
            if ($acct -eq "495400") {
                $salesByDate[$dateKey].Freight += $amount
                $salesByDate[$dateKey].Total += $amount
            } elseif ($acct -eq "495000") {
                $salesByDate[$dateKey].Miscellaneous += $amount
                $salesByDate[$dateKey].Total += $amount
            } elseif ($acct -in @("401000", "402000")) {
                # Handle both invoices (ARI) and credit memos (ARC)
                $invNum = $null
                $orderNum = $null
                $svcRatio = 0
                $prodRatio = 1  # Default to product
                
                if ($gl.Ref -match "ARI\s+(\d+)") {
                    $invNum = $matches[1].Trim()
                    $orderNum = $invToOrder[$invNum]
                } elseif ($gl.Ref -match "ARC") {
                    # AR Credit - reduces revenue (amount is already handled by * -1)
                    # These don't have order linkage, so default to product
                    $invNum = $null
                    $orderNum = $null
                }
                
                if ($orderNum -and $orderBreakdown.ContainsKey($orderNum) -and $orderBreakdown[$orderNum].Total -gt 0) {
                    $breakdown = $orderBreakdown[$orderNum]
                    $svcRatio = $breakdown.Service / $breakdown.Total
                    $prodRatio = $breakdown.Product / $breakdown.Total
                    
                    $salesByDate[$dateKey].Service += $amount * $svcRatio
                    $salesByDate[$dateKey].Product += $amount * $prodRatio
                } else {
                    $salesByDate[$dateKey].Product += $amount  # Default to product (includes ARC credits)
                }
                $salesByDate[$dateKey].Total += $amount
                
                # Customer tracking (using pre-built lookup)
                $custName = $invToCustomer[$invNum]
                if ($custName) {
                    if (-not $salesByCustomer.ContainsKey($custName)) {
                        $salesByCustomer[$custName] = @{ MTD=0; Yesterday=0; YTD=0 }
                    }
                    if ($salesDate -ge $yearStart -and $salesDate -le $today) { $salesByCustomer[$custName].YTD += $amount }
                    if ($salesDate -ge $monthStart -and $salesDate -le $today) { $salesByCustomer[$custName].MTD += $amount }
                    if ($salesDate -eq $yesterday) { $salesByCustomer[$custName].Yesterday += $amount }
                }
                
                # Salesperson tracking with Product/Service breakdown
                $salesperson = if ($orderNum -and $orderToSalesperson.ContainsKey($orderNum)) { $orderToSalesperson[$orderNum] } else { "Unassigned" }
                if (-not $salesBySalesperson.ContainsKey($salesperson)) {
                    $salesBySalesperson[$salesperson] = @{ MTD=0; Yesterday=0; YTD=0; InvoiceCount=0; MTDProduct=0; MTDService=0; YesterdayProduct=0; YesterdayService=0 }
                }
                # Calculate product/service split for this transaction
                $txnSvcAmt = $amount * $svcRatio
                $txnProdAmt = $amount * $prodRatio
                
                if ($salesDate -ge $yearStart -and $salesDate -le $today) { $salesBySalesperson[$salesperson].YTD += $amount }
                if ($salesDate -ge $monthStart -and $salesDate -le $today) { 
                    $salesBySalesperson[$salesperson].MTD += $amount 
                    $salesBySalesperson[$salesperson].MTDProduct += $txnProdAmt
                    $salesBySalesperson[$salesperson].MTDService += $txnSvcAmt
                    $salesBySalesperson[$salesperson].InvoiceCount++
                }
                if ($salesDate -eq $yesterday) { 
                    $salesBySalesperson[$salesperson].Yesterday += $amount
                    $salesBySalesperson[$salesperson].YesterdayProduct += $txnProdAmt
                    $salesBySalesperson[$salesperson].YesterdayService += $txnSvcAmt
                }
            }
        } catch {}
    }
}

# === STEP 5: Calculate summaries ===
$yesterdaySales = if ($salesByDate.ContainsKey($yesterday.ToString("yyyy-MM-dd"))) { $salesByDate[$yesterday.ToString("yyyy-MM-dd")] } else { @{Product=0;Service=0;Freight=0;Miscellaneous=0;Total=0} }
$mtdSales = @{Product=0;Service=0;Freight=0;Miscellaneous=0;Total=0}
$ytdSales = @{Product=0;Service=0;Freight=0;Miscellaneous=0;Total=0}

foreach ($dateKey in $salesByDate.Keys) {
    $date = [DateTime]::Parse($dateKey)
    $data = $salesByDate[$dateKey]
    if ($date -ge $monthStart -and $date -le $today) {
        foreach ($cat in @("Product","Service","Freight","Miscellaneous","Total")) { $mtdSales[$cat] += $data[$cat] }
    }
    if ($date -ge $yearStart -and $date -le $today) {
        foreach ($cat in @("Product","Service","Freight","Miscellaneous","Total")) { $ytdSales[$cat] += $data[$cat] }
    }
}

$mtdDays = Get-BusinessDays $monthStart $today
$ytdDays = Get-BusinessDays $yearStart $today
$mtdAvg = if ($mtdDays -gt 0) { [math]::Round($mtdSales.Total / $mtdDays, 2) } else { 0 }
$ytdAvg = if ($ytdDays -gt 0) { [math]::Round($ytdSales.Total / $ytdDays, 2) } else { 0 }

if ($mtdSales.Total -gt 0) {
    Write-Host "Service/Product split: $([math]::Round($mtdSales.Service / $mtdSales.Total * 100, 1))% Service, $([math]::Round($mtdSales.Product / $mtdSales.Total * 100, 1))% Product"
} else {
    Write-Host "No MTD sales yet (month just started)"
}

# === STEP 5.5: Calculate Day of Week Averages (FIX for empty chart) ===
Write-Host "Calculating day of week averages..."
$dowTotals = @{ Monday=@(); Tuesday=@(); Wednesday=@(); Thursday=@(); Friday=@() }

# Use last 60 days of data for day-of-week patterns
for ($i = 60; $i -ge 0; $i--) {
    $date = $today.AddDays(-$i)
    if ($date.DayOfWeek -in @('Saturday','Sunday')) { continue }
    $dateKey = $date.ToString("yyyy-MM-dd")
    $dow = $date.DayOfWeek.ToString()
    if ($salesByDate.ContainsKey($dateKey) -and $salesByDate[$dateKey].Total -gt 0) {
        $dowTotals[$dow] += $salesByDate[$dateKey].Total
    }
}

$avgByDayOfWeek = @{}
foreach ($dow in @("Monday","Tuesday","Wednesday","Thursday","Friday")) {
    $values = $dowTotals[$dow]
    if ($values.Count -gt 0) {
        $avgByDayOfWeek[$dow] = [math]::Round(($values | Measure-Object -Sum).Sum / $values.Count, 2)
    } else {
        $avgByDayOfWeek[$dow] = $mtdAvg  # Default to MTD average
    }
}
Write-Host "Day of week averages calculated: Mon=$($avgByDayOfWeek['Monday']), Tue=$($avgByDayOfWeek['Tuesday']), Wed=$($avgByDayOfWeek['Wednesday']), Thu=$($avgByDayOfWeek['Thursday']), Fri=$($avgByDayOfWeek['Friday'])"

# === STEP 6: AR Aging (use SyteLine calculated data - Invoice minus Payments) ===
Write-Host "Loading AR aging data..."
$sytelineArPath = Join-Path $OutputPath "data\ar-aging-net.json"
if (Test-Path $sytelineArPath) {
    $sytelineAR = Get-Content $sytelineArPath | ConvertFrom-Json
    $arAging = @{
        Current = [decimal]$sytelineAR.summary.buckets.current
        Days1_30 = [decimal]$sytelineAR.summary.buckets.'1-30'
        Days31_60 = [decimal]$sytelineAR.summary.buckets.'31-60'
        Days61_90 = [decimal]$sytelineAR.summary.buckets.'61-90'
        Days90Plus = [decimal]$sytelineAR.summary.buckets.'91+'
        Total = [decimal]$sytelineAR.summary.totalAR
    }
    Write-Host "Using SyteLine AR aging (Invoice-Payments): Total = `$$([math]::Round($arAging.Total, 2))"
} else {
    # Calculate from transactions if Birst data not available
    $arAging = @{ Current=0; Days1_30=0; Days31_60=0; Days61_90=0; Days90Plus=0; Total=0 }
    foreach ($inv in $openAR) {
        $days = $inv.DaysOverdue
        if ($days -le 0) { $arAging.Current += $inv.Amount }
        elseif ($days -le 30) { $arAging.Days1_30 += $inv.Amount }
        elseif ($days -le 60) { $arAging.Days31_60 += $inv.Amount }
        elseif ($days -le 90) { $arAging.Days61_90 += $inv.Amount }
        else { $arAging.Days90Plus += $inv.Amount }
        $arAging.Total += $inv.Amount
    }
    Write-Host "Calculated AR aging from transactions: Total = `$$([math]::Round($arAging.Total, 2))"
}

# === STEP 7: IMPROVED Cash Flow Prediction ===
Write-Host "Running improved cash flow prediction..."

# Historical average daily collections (based on actual GL data)
# This represents actual revenue posting to GL, which correlates with cash flow
$historicalDailyAvg = $mtdAvg  # Use MTD average as baseline

# AR-based collection expectations by bucket (more conservative rates)
# These represent what % of each bucket we expect to COLLECT in a week
$weeklyCollectionRates = @{
    Current = 0.20     # 20% of current AR collects each week (5-week cycle)
    Days1_30 = 0.15    # 15% of 1-30 day AR collects per week
    Days31_60 = 0.10   # 10% of 31-60 day AR collects per week
    Days61_90 = 0.05   # 5% of 61-90 day AR collects per week
    Days90Plus = 0.02  # 2% of 90+ day AR collects per week (mostly write-offs)
}

# Day of week weight for when collections are received
$dowWeights = @{
    Monday = 0.15
    Tuesday = 0.25
    Wednesday = 0.25
    Thursday = 0.20
    Friday = 0.15
}

# Holidays reduce collections
$holidays2026 = @(
    "2026-01-01", "2026-01-19", "2026-02-16", "2026-05-25", 
    "2026-07-03", "2026-09-07", "2026-11-26", "2026-11-27", "2026-12-25"
) | ForEach-Object { [DateTime]::Parse($_) }

# Build AR due by week
$arByDueWeek = @{}
foreach ($inv in $openAR) {
    $dueDate = $inv.DueDate
    $weekStart = $dueDate.AddDays(-[int]$dueDate.DayOfWeek + 1)
    if ($dueDate.DayOfWeek -eq 'Sunday') { $weekStart = $dueDate.AddDays(-6) }
    $weekKey = $weekStart.ToString("yyyy-MM-dd")
    if (-not $arByDueWeek[$weekKey]) { $arByDueWeek[$weekKey] = 0 }
    $arByDueWeek[$weekKey] += $inv.Amount
}

# Calculate 6-week forecast
$thisWeek = Get-WeekRange $today
$cashFlowPrediction = @()

# Rolling AR balance per bucket
$rollingAR = @{
    Current = $arAging.Current
    Days1_30 = $arAging.Days1_30
    Days31_60 = $arAging.Days31_60
    Days61_90 = $arAging.Days61_90
    Days90Plus = $arAging.Days90Plus
}

for ($w = 0; $w -lt 6; $w++) {
    $weekStart = $thisWeek.Start.AddDays($w * 7)
    $weekEnd = $weekStart.AddDays(4)
    $weekLabel = "Week of $($weekStart.ToString('MMM d'))"
    $weekKey = $weekStart.ToString("yyyy-MM-dd")
    
    $predicted = 0; $actual = 0; $dayBreakdown = @()
    
    # Expected collections from each aging bucket
    $expectedFromCurrent = $rollingAR.Current * $weeklyCollectionRates.Current
    $expectedFrom1_30 = $rollingAR.Days1_30 * $weeklyCollectionRates.Days1_30
    $expectedFrom31_60 = $rollingAR.Days31_60 * $weeklyCollectionRates.Days31_60
    $expectedFrom61_90 = $rollingAR.Days61_90 * $weeklyCollectionRates.Days61_90
    $expectedFrom90Plus = $rollingAR.Days90Plus * $weeklyCollectionRates.Days90Plus
    
    # Total expected from AR
    $arBasedExpectation = $expectedFromCurrent + $expectedFrom1_30 + $expectedFrom31_60 + $expectedFrom61_90 + $expectedFrom90Plus
    
    # Blend AR expectation with historical daily average
    # Weight: 60% historical pattern, 40% AR-based
    $weekExpectedCollections = ($historicalDailyAvg * 5 * 0.6) + ($arBasedExpectation * 0.4)
    
    # New AR coming due this week
    $newARDue = if ($arByDueWeek[$weekKey]) { $arByDueWeek[$weekKey] } else { 0 }
    
    for ($d = 0; $d -lt 5; $d++) {
        $day = $weekStart.AddDays($d)
        $dayKey = $day.ToString("yyyy-MM-dd")
        $dow = $day.DayOfWeek.ToString()
        
        if ($day -lt $today) {
            # Actual data from GL
            $dayAmount = if ($salesByDate.ContainsKey($dayKey)) { $salesByDate[$dayKey].Total } else { 0 }
            $actual += $dayAmount
            $dayBreakdown += @{ Date=$dayKey; DayOfWeek=$dow; Type="Actual"; Amount=[math]::Round($dayAmount,2) }
        } elseif ($day -eq $today) {
            $dayAmount = if ($salesByDate.ContainsKey($dayKey)) { $salesByDate[$dayKey].Total } else { 0 }
            $actual += $dayAmount
            $dayBreakdown += @{ Date=$dayKey; DayOfWeek=$dow; Type="Today"; Amount=[math]::Round($dayAmount,2) }
        } else {
            # Predicted: use day-of-week weighted distribution
            $dowWeight = if ($dowWeights[$dow]) { $dowWeights[$dow] } else { 0.20 }
            
            # Holiday factor
            $isHoliday = $day -in $holidays2026
            $holidayFactor = if ($isHoliday) { 0.2 } else { 1.0 }
            
            # Use actual day-of-week average if available, blended with weekly expectation
            $dowAvgForDay = if ($avgByDayOfWeek[$dow]) { $avgByDayOfWeek[$dow] } else { $historicalDailyAvg }
            
            # Blend: 70% historical DOW average, 30% week-spread
            $dayPrediction = ($dowAvgForDay * 0.7 + ($weekExpectedCollections * $dowWeight) * 0.3) * $holidayFactor
            
            $predicted += $dayPrediction
            $dayBreakdown += @{ Date=$dayKey; DayOfWeek=$dow; Type="Predicted"; Amount=[math]::Round($dayPrediction,2) }
        }
    }
    
    # Age the AR buckets for next week
    $rollingAR.Days90Plus = [math]::Max(0, $rollingAR.Days90Plus - $expectedFrom90Plus + $rollingAR.Days61_90 * 0.25)
    $rollingAR.Days61_90 = [math]::Max(0, $rollingAR.Days61_90 * 0.75 - $expectedFrom61_90 + $rollingAR.Days31_60 * 0.25)
    $rollingAR.Days31_60 = [math]::Max(0, $rollingAR.Days31_60 * 0.75 - $expectedFrom31_60 + $rollingAR.Days1_30 * 0.25)
    $rollingAR.Days1_30 = [math]::Max(0, $rollingAR.Days1_30 * 0.75 - $expectedFrom1_30 + $rollingAR.Current * 0.2)
    $rollingAR.Current = [math]::Max(0, $rollingAR.Current * 0.8 - $expectedFromCurrent + $newARDue)
    
    $weekTotal = $actual + $predicted
    
    # Confidence: decreases further out, adjusted by how much is actual vs predicted
    $actualPct = if ($weekTotal -gt 0) { $actual / $weekTotal } else { 0 }
    $baseConfidence = if ($w -eq 0) { 90 } elseif ($w -eq 1) { 75 } elseif ($w -lt 4) { 60 } else { 45 }
    $confidence = [math]::Round($baseConfidence * (0.5 + $actualPct * 0.5))
    
    $cashFlowPrediction += @{
        Week=$weekLabel; WeekNum=$w+1
        StartDate=$weekStart.ToString("yyyy-MM-dd"); EndDate=$weekEnd.ToString("yyyy-MM-dd")
        Actual=[math]::Round($actual,2); Predicted=[math]::Round($predicted,2); Total=[math]::Round($weekTotal,2)
        ARDue=[math]::Round($newARDue,2)
        ExpectedCollections=[math]::Round($weekExpectedCollections,2)
        Confidence=$confidence
        DayBreakdown=$dayBreakdown
        Sources=@{
            FromCurrent=[math]::Round($expectedFromCurrent,2)
            From1_30=[math]::Round($expectedFrom1_30,2)
            From31_60=[math]::Round($expectedFrom31_60,2)
            From61_90=[math]::Round($expectedFrom61_90,2)
            From90Plus=[math]::Round($expectedFrom90Plus,2)
        }
    }
}

Write-Host "Cash flow prediction complete"
$cf6WeekTotal = 0
$cashFlowPrediction | ForEach-Object { $cf6WeekTotal += $_.Total }
Write-Host "6-week forecast total: `$$([math]::Round($cf6WeekTotal, 2))"

# === STEP 8: Shipments ===
Write-Host "Fetching shipment data..."
$shipments = Invoke-SyteLineAPI -Token $token -IDO "SLShipments" -Properties "ConsigneeCity,ConsigneeState,ConsigneeZip" -RecordCap 10000
$locationCounts = @{}
foreach ($ship in $shipments) {
    if ($ship.ConsigneeCity -and $ship.ConsigneeState) {
        $locKey = "$($ship.ConsigneeCity), $($ship.ConsigneeState)"
        if (-not $locationCounts.ContainsKey($locKey)) { $locationCounts[$locKey] = @{City=$ship.ConsigneeCity; State=$ship.ConsigneeState; Zip=$ship.ConsigneeZip; Count=0} }
        $locationCounts[$locKey].Count++
    }
}
Write-Host "Total shipping locations: $($locationCounts.Count)"

# === STEP 8.5: Service Tech Locations ===
Write-Host "Building service tech location data..."
# Get service items (LFTR, LGAS, LIHL = Field/Gas/In-house Labor)
$serviceItemFilter = "ProductCode IN ('LFTR','LGAS','LIHL')"
$serviceItemsData = Invoke-SyteLineAPI -Token $token -IDO "SLItems" -Properties "Item" -Filter $serviceItemFilter -RecordCap 500
$serviceItemSet = @{}
foreach ($si in $serviceItemsData) { if ($si.Item) { $serviceItemSet[$si.Item.Trim()] = $true } }
Write-Host "Service items: $($serviceItemSet.Count)"

# Find orders with service items using existing $coItems data
$serviceOrderNums = @{}
foreach ($line in $coItems) {
    if ($line.Item -and $serviceItemSet.ContainsKey($line.Item.Trim()) -and $line.CoNum) {
        $serviceOrderNums[$line.CoNum.Trim()] = $true
    }
}
Write-Host "Orders with service items: $($serviceOrderNums.Count)"

# Get service locations from orders
$serviceLocationCounts = @{}
foreach ($order in $orders) {
    if ($order.CoNum -and $serviceOrderNums.ContainsKey($order.CoNum.Trim()) -and $order.ShipToCity -and $order.ShipToState) {
        $locKey = "$($order.ShipToCity.Trim()), $($order.ShipToState.Trim())"
        if (-not $serviceLocationCounts.ContainsKey($locKey)) { 
            $serviceLocationCounts[$locKey] = @{City=$order.ShipToCity.Trim(); State=$order.ShipToState.Trim(); Count=0} 
        }
        $serviceLocationCounts[$locKey].Count++
    }
}
Write-Host "Service tech locations: $($serviceLocationCounts.Count)"

# === STEP 9: Build output ===
$dailyTrend = @()
for ($i = 60; $i -ge 0; $i--) {
    $date = $today.AddDays(-$i)
    # Skip weekends - CTI doesn't operate on weekends
    if ($date.DayOfWeek -eq 'Saturday' -or $date.DayOfWeek -eq 'Sunday') { continue }
    $dateKey = $date.ToString("yyyy-MM-dd")
    $data = if ($salesByDate.ContainsKey($dateKey)) { $salesByDate[$dateKey] } else { @{Product=0;Service=0;Freight=0;Miscellaneous=0;Total=0} }
    $dailyTrend += @{ Date=$dateKey; DayOfWeek=$date.DayOfWeek.ToString(); Product=[math]::Round($data.Product,2); Service=[math]::Round($data.Service,2); Freight=[math]::Round($data.Freight,2); Miscellaneous=[math]::Round($data.Miscellaneous,2); Total=[math]::Round($data.Total,2) }
}

$topCustomersMTD = $salesByCustomer.GetEnumerator() | Where-Object { $_.Value.MTD -gt 0 } | Sort-Object { $_.Value.MTD } -Descending | Select-Object -First 15 | ForEach-Object { @{Name=$_.Key; Amount=[math]::Round($_.Value.MTD,2)} }
$topCustomersYesterday = $salesByCustomer.GetEnumerator() | Where-Object { $_.Value.Yesterday -gt 0 } | Sort-Object { $_.Value.Yesterday } -Descending | Select-Object -First 15 | ForEach-Object { @{Name=$_.Key; Amount=[math]::Round($_.Value.Yesterday,2)} }
$topProductsMTD = $salesByProduct.GetEnumerator() | Sort-Object { $_.Value.MTD } -Descending | Select-Object -First 15 | ForEach-Object { @{Item=$_.Key; Description=$_.Value.Description; Amount=[math]::Round($_.Value.MTD,2)} }
# All shipping locations (no limit)
$shippingLocations = $locationCounts.Values | Sort-Object -Property Count -Descending
# All service tech locations (no limit)
$serviceTechLocations = $serviceLocationCounts.Values | Sort-Object -Property Count -Descending

# Team member roles and labels
$teamRoles = @{
    'Joel' = @{ Role='COO'; Type='Executive' }
    'Allison' = @{ Role='Accounting'; Type='Admin' }
    'lphil' = @{ Role='Sales Rep'; Type='Sales' }
    'cjord' = @{ Role='Sales Rep'; Type='Sales' }
    'mrand' = @{ Role='Sales Rep'; Type='Sales' }
    'nandy' = @{ Role='Sales Rep'; Type='Sales' }
    'marissa' = @{ Role='Sales Rep'; Type='Sales' }
    'seric' = @{ Role='Service Rep'; Type='Service' }
    'mjaso' = @{ Role='Service Rep'; Type='Service' }
    'bcoli' = @{ Role='Sales Rep'; Type='Sales' }
    'tbolt' = @{ Role='Sales Rep'; Type='Sales' }
    'hnich' = @{ Role='Service Rep'; Type='Service' }
    'ljami' = @{ Role='Sales Rep'; Type='Sales' }
    'Greg' = @{ Role='Sales Rep'; Type='Sales' }
    'Anna' = @{ Role='Sales Rep'; Type='Sales' }
    'Rachel L' = @{ Role='Sales Rep'; Type='Sales' }
}

# Build comprehensive team data with Product/Service breakdown AND territory info
$allTeamMTD = $salesBySalesperson.GetEnumerator() | Where-Object { $_.Value.MTD -gt 0 } | Sort-Object { $_.Value.MTD } -Descending | ForEach-Object { 
    $roleInfo = if ($teamRoles.ContainsKey($_.Key)) { $teamRoles[$_.Key] } else { @{ Role='Team Member'; Type='Sales' } }
    $territory = if ($territoryByPerson.ContainsKey($_.Key)) { $territoryByPerson[$_.Key] } else { @{ States=@{}; Cities=@{}; OrderCount=0 } }
    @{
        Name=$_.Key
        Role=$roleInfo.Role
        Type=$roleInfo.Type
        Amount=[math]::Round($_.Value.MTD,2)
        ProductAmount=[math]::Round($_.Value.MTDProduct,2)
        ServiceAmount=[math]::Round($_.Value.MTDService,2)
        InvoiceCount=$_.Value.InvoiceCount
        StateCount=$territory.States.Count
        CityCount=$territory.Cities.Count
        States=($territory.States.Keys | Sort-Object) -join ", "
        TopCities=($territory.Cities.Keys | Select-Object -First 5) -join "; "
    }
}

$allTeamYesterday = $salesBySalesperson.GetEnumerator() | Where-Object { $_.Value.Yesterday -gt 0 } | Sort-Object { $_.Value.Yesterday } -Descending | ForEach-Object { 
    $roleInfo = if ($teamRoles.ContainsKey($_.Key)) { $teamRoles[$_.Key] } else { @{ Role='Team Member'; Type='Sales' } }
    $territory = if ($territoryByPerson.ContainsKey($_.Key)) { $territoryByPerson[$_.Key] } else { @{ States=@{}; Cities=@{}; OrderCount=0 } }
    @{
        Name=$_.Key
        Role=$roleInfo.Role
        Type=$roleInfo.Type
        Amount=[math]::Round($_.Value.Yesterday,2)
        ProductAmount=[math]::Round($_.Value.YesterdayProduct,2)
        ServiceAmount=[math]::Round($_.Value.YesterdayService,2)
        StateCount=$territory.States.Count
        CityCount=$territory.Cities.Count
    }
}

# Filter for sales reps only (excludes Executive, Admin types)
$salespersonMTD = $allTeamMTD | Where-Object { $_.Type -eq 'Sales' }
$salespersonYesterday = $allTeamYesterday | Where-Object { $_.Type -eq 'Sales' }

# Filter for service reps only
$serviceRepMTD = $allTeamMTD | Where-Object { $_.Type -eq 'Service' }
$serviceRepYesterday = $allTeamYesterday | Where-Object { $_.Type -eq 'Service' }

Write-Host "`n=== TEAM SUMMARY ==="
Write-Host "All Team MTD (with roles and territories):"
$allTeamMTD | ForEach-Object { Write-Host "  $($_.Name) ($($_.Role)): `$$($_.Amount) | $($_.StateCount) states, $($_.CityCount) cities" }
Write-Host "`nSales Reps Only:"
$salespersonMTD | ForEach-Object { Write-Host "  $($_.Name): `$$($_.Amount)" }
Write-Host "`nService Reps Only:"
$serviceRepMTD | ForEach-Object { Write-Host "  $($_.Name): `$$($_.Amount)" }

# Load Salesforce territory data if available
$salesforceTerritoryPath = Join-Path $OutputPath "data\salesforce-territory.json"
$salesforceTerritory = $null
if (Test-Path $salesforceTerritoryPath) {
    try {
        $salesforceTerritory = Get-Content $salesforceTerritoryPath | ConvertFrom-Json
        Write-Host "Loaded Salesforce territory data: $($salesforceTerritory.salesTeam.Count) sales reps"
    } catch {
        Write-Host "Warning: Could not load Salesforce territory data"
    }
}

$dashboardData = @{
    GeneratedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    DataSource = "SyteLine GL Ledger with ProductCode (LFTR/LGAS/LIHL=Service)"
    Period = @{ Yesterday=$yesterday.ToString("yyyy-MM-dd"); MonthStart=$monthStart.ToString("yyyy-MM-dd"); YearStart=$yearStart.ToString("yyyy-MM-dd"); MTDBusinessDays=$mtdDays; YTDBusinessDays=$ytdDays }
    Summary = @{
        Yesterday = @{ Product=[math]::Round($yesterdaySales.Product,2); Service=[math]::Round($yesterdaySales.Service,2); Freight=[math]::Round($yesterdaySales.Freight,2); Miscellaneous=[math]::Round($yesterdaySales.Miscellaneous,2); Total=[math]::Round($yesterdaySales.Total,2) }
        MTD = @{ Product=[math]::Round($mtdSales.Product,2); Service=[math]::Round($mtdSales.Service,2); Freight=[math]::Round($mtdSales.Freight,2); Miscellaneous=[math]::Round($mtdSales.Miscellaneous,2); Total=[math]::Round($mtdSales.Total,2) }
        YTD = @{ Product=[math]::Round($ytdSales.Product,2); Service=[math]::Round($ytdSales.Service,2); Freight=[math]::Round($ytdSales.Freight,2); Miscellaneous=[math]::Round($ytdSales.Miscellaneous,2); Total=[math]::Round($ytdSales.Total,2) }
        MTDDailyAverage = $mtdAvg; YTDDailyAverage = $ytdAvg
    }
    ARaging = @{ Current=[math]::Round($arAging.Current,2); Days1_30=[math]::Round($arAging.Days1_30,2); Days31_60=[math]::Round($arAging.Days31_60,2); Days61_90=[math]::Round($arAging.Days61_90,2); Days90Plus=[math]::Round($arAging.Days90Plus,2); Total=[math]::Round($arAging.Total,2) }
    CashFlowPrediction = $cashFlowPrediction
    TopCustomers = @{ Yesterday=$topCustomersYesterday; MTD=$topCustomersMTD }
    TopProducts = @{ MTD=$topProductsMTD }
    Salesperson = @{ Yesterday=$salespersonYesterday; MTD=$salespersonMTD }
    ServiceReps = @{ Yesterday=$serviceRepYesterday; MTD=$serviceRepMTD }
    AllTeam = @{ Yesterday=$allTeamYesterday; MTD=$allTeamMTD }
    DailyTrend = $dailyTrend
    ShippingHeatMap = $shippingLocations
    ServiceTechLocations = $serviceTechLocations
    DayOfWeekAverages = $avgByDayOfWeek
    SalesforceTerritory = $salesforceTerritory
}

# Save JSON
$jsonPath = Join-Path $OutputPath "data\dashboard-data.json"
if (-not (Test-Path (Split-Path $jsonPath))) { New-Item -ItemType Directory -Path (Split-Path $jsonPath) -Force | Out-Null }
$dashboardData | ConvertTo-Json -Depth 10 | Set-Content $jsonPath -Encoding UTF8
Write-Host "Saved JSON: $jsonPath"

Write-Host "`n=== SUMMARY ==="
Write-Host "Yesterday: `$$([math]::Round($yesterdaySales.Total,2)) (Product: `$$([math]::Round($yesterdaySales.Product,2)), Service: `$$([math]::Round($yesterdaySales.Service,2)))"
Write-Host "MTD: `$$([math]::Round($mtdSales.Total,2)) (Product: `$$([math]::Round($mtdSales.Product,2)), Service: `$$([math]::Round($mtdSales.Service,2)))"
Write-Host "AR Total: `$$([math]::Round($arAging.Total,2))"
Write-Host "6-Week Cash Flow Forecast:"
$cashFlowPrediction | ForEach-Object { Write-Host "  $($_.Week): `$$($_.Total) ($($_.Confidence)% confidence)" }

# Create self-contained HTML
Write-Host "`nCreating self-contained dashboard..."
# Use cti-dashboard.html as template (has latest AllTeam, ServiceReps, Salesforce features)
$htmlTemplatePath = Join-Path $OutputPath "cti-dashboard.html"
if (Test-Path $htmlTemplatePath) {
    $htmlTemplate = Get-Content $htmlTemplatePath -Raw
    $jsonData = $dashboardData | ConvertTo-Json -Depth 10
    $embeddedHtml = $htmlTemplate -replace 'let DATA = null;', "let DATA = $jsonData;"
    $liveHtmlPath = Join-Path $OutputPath "cti-dashboard-live.html"
    $embeddedHtml | Set-Content $liveHtmlPath -Encoding UTF8
    Write-Host "Dashboard saved: $liveHtmlPath"
}

Write-Host "`n$(Get-Date) - Dashboard generation complete!"
