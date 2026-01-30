# CTI Executive Dashboard Generator
# Generates a self-contained HTML dashboard with embedded data
# Uses itmuf_cti_product_category field for Service/Product classification
# SOURCE OF TRUTH: GL Ledger with invoice-level categorization

param(
    [string]$OutputPath = "C:\Users\justi\clawd\dashboard"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Web

# SyteLine API Configuration
$SyteLineConfig = @{
    BaseUrl = "https://csi10g.erpsl.inforcloudsuite.com/IDORequestService/ido"
    Tenant = "GVNDYXUFKHB5VMB6_PRD_CTI"
    Username = "gary.phillips@godlan.com"
    Password = 'Crwthtithing2$'
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

# === STEP 3: Get invoice to order mapping ===
Write-Host "Fetching AR transactions..."
# Get AR transactions for aging calculation (Invoice - Payments - Credits = Balance)
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

# === STEP 4: Get GL Ledger and calculate sales by date with proper categorization ===
Write-Host "Fetching GL ledger data..."
$filter = "(Acct = '401000' OR Acct = '402000' OR Acct = '495000' OR Acct = '495400') AND ControlYear >= $currentYear"
$ledgerData = Invoke-SyteLineAPI -Token $token -IDO "SLLedgers" -Properties "Acct,DomAmount,TransDate,Ref" -Filter $filter -RecordCap 10000
Write-Host "Got $($ledgerData.Count) GL ledger records"

$salesByDate = @{}
$salesByCustomer = @{}

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
            } elseif ($acct -in @("401000", "402000") -and $gl.Ref -match "ARI\s+(\d+)") {
                $invNum = $matches[1].Trim()
                $orderNum = $invToOrder[$invNum]
                
                if ($orderNum -and $orderBreakdown.ContainsKey($orderNum) -and $orderBreakdown[$orderNum].Total -gt 0) {
                    $breakdown = $orderBreakdown[$orderNum]
                    $svcRatio = $breakdown.Service / $breakdown.Total
                    $prodRatio = $breakdown.Product / $breakdown.Total
                    
                    $salesByDate[$dateKey].Service += $amount * $svcRatio
                    $salesByDate[$dateKey].Product += $amount * $prodRatio
                } else {
                    $salesByDate[$dateKey].Product += $amount  # Default to product
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

Write-Host "Service/Product split: $([math]::Round($mtdSales.Service / $mtdSales.Total * 100, 1))% Service, $([math]::Round($mtdSales.Product / $mtdSales.Total * 100, 1))% Product"

# === STEP 6: AR Aging (calculated from transactions) ===
Write-Host "Calculating AR aging from open invoices..."

# Include unapplied credits as part of AR (customer credits on account)
$unappliedCredits = $arTrans | Where-Object { $_.Type -eq "C" -and (!$_.ApplyToInvNum -or $_.ApplyToInvNum -eq "0") }
$unappliedCreditTotal = ($unappliedCredits | Measure-Object -Property Amount -Sum).Sum
if (!$unappliedCreditTotal) { $unappliedCreditTotal = 0 }

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

# Add back unapplied credits (they weren't subtracted from any invoice, but they reduce future AR)
# Actually subtract them from current since they're credits on account
$arAging.Current = [math]::Max(0, $arAging.Current - $unappliedCreditTotal)
$arAging.Total = $arAging.Current + $arAging.Days1_30 + $arAging.Days31_60 + $arAging.Days61_90 + $arAging.Days90Plus

Write-Host "AR Total: `$$([math]::Round($arAging.Total, 2)) (unapplied credits: `$$([math]::Round($unappliedCreditTotal, 2)))"

# === STEP 7: Smart Cash Flow Prediction ===
Write-Host "Running smart cash flow prediction..."

# --- COLLECTION PROBABILITY BY AGING BUCKET ---
# Based on typical B2B collection patterns
$collectionRates = @{
    Current = 0.70      # 70% of current AR collects within the week it's due
    Days1_30 = 0.25     # 25% of 1-30 day AR collects per week
    Days31_60 = 0.15    # 15% of 31-60 day AR collects per week
    Days61_90 = 0.08    # 8% of 61-90 day AR collects per week
    Days90Plus = 0.03   # 3% of 90+ day AR collects per week
}

# --- DAY OF WEEK PAYMENT PATTERNS ---
# Payments typically concentrate on certain days (Tuesday/Wednesday heavy)
$dowWeights = @{
    Monday = 0.15       # Start of week - some payments
    Tuesday = 0.28      # Heavy payment day
    Wednesday = 0.25    # Heavy payment day
    Thursday = 0.20     # Moderate
    Friday = 0.12       # Light - people wrap up
}

# --- HOLIDAYS (reduce collections) ---
$holidays2026 = @(
    "2026-01-01", "2026-01-20", "2026-02-17", "2026-05-25", 
    "2026-07-03", "2026-09-07", "2026-11-26", "2026-11-27", "2026-12-25"
) | ForEach-Object { [DateTime]::Parse($_) }

# --- BUILD AR PIPELINE BY DUE DATE ---
$arByDueWeek = @{}
foreach ($inv in $openAR) {
    $dueDate = $inv.DueDate
    $weekStart = $dueDate.AddDays(-[int]$dueDate.DayOfWeek + 1)
    $weekKey = $weekStart.ToString("yyyy-MM-dd")
    if (-not $arByDueWeek[$weekKey]) { $arByDueWeek[$weekKey] = 0 }
    $arByDueWeek[$weekKey] += $inv.Amount
}

# --- CALCULATE WEEKLY EXPECTED COLLECTIONS ---
$thisWeek = Get-WeekRange $today
$cashFlowPrediction = @()

# Track rolling AR for each bucket
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
    
    # Calculate expected collections from each aging bucket
    $expectedFromCurrent = $rollingAR.Current * $collectionRates.Current
    $expectedFrom1_30 = $rollingAR.Days1_30 * $collectionRates.Days1_30
    $expectedFrom31_60 = $rollingAR.Days31_60 * $collectionRates.Days31_60
    $expectedFrom61_90 = $rollingAR.Days61_90 * $collectionRates.Days61_90
    $expectedFrom90Plus = $rollingAR.Days90Plus * $collectionRates.Days90Plus
    
    $weekExpectedCollections = $expectedFromCurrent + $expectedFrom1_30 + $expectedFrom31_60 + $expectedFrom61_90 + $expectedFrom90Plus
    
    # New AR coming due this week (adds to current bucket for next week)
    $newARDue = if ($arByDueWeek[$weekKey]) { $arByDueWeek[$weekKey] } else { 0 }
    
    for ($d = 0; $d -lt 5; $d++) {
        $day = $weekStart.AddDays($d)
        $dayKey = $day.ToString("yyyy-MM-dd")
        $dow = $day.DayOfWeek.ToString()
        
        if ($day -lt $today) {
            # Actual data
            $dayAmount = if ($salesByDate.ContainsKey($dayKey)) { $salesByDate[$dayKey].Total } else { 0 }
            $actual += $dayAmount
            $dayBreakdown += @{ Date=$dayKey; DayOfWeek=$dow; Type="Actual"; Amount=[math]::Round($dayAmount,2) }
        } elseif ($day -eq $today) {
            $dayAmount = if ($salesByDate.ContainsKey($dayKey)) { $salesByDate[$dayKey].Total } else { 0 }
            $actual += $dayAmount
            $dayBreakdown += @{ Date=$dayKey; DayOfWeek=$dow; Type="Today"; Amount=[math]::Round($dayAmount,2) }
        } else {
            # Predicted: distribute weekly collections by day-of-week weight
            $dowWeight = if ($dowWeights[$dow]) { $dowWeights[$dow] } else { 0.20 }
            
            # Check for holiday - reduce by 80%
            $isHoliday = $holidays2026 | Where-Object { $_.Date -eq $day.Date }
            $holidayFactor = if ($isHoliday) { 0.2 } else { 1.0 }
            
            $dayPrediction = $weekExpectedCollections * $dowWeight * $holidayFactor
            $predicted += $dayPrediction
            $dayBreakdown += @{ Date=$dayKey; DayOfWeek=$dow; Type="Predicted"; Amount=[math]::Round($dayPrediction,2) }
        }
    }
    
    # Update rolling AR for next week (age the buckets)
    $collected = $predicted
    $rollingAR.Days90Plus = $rollingAR.Days90Plus - $expectedFrom90Plus + $rollingAR.Days61_90 * 0.3
    $rollingAR.Days61_90 = $rollingAR.Days61_90 * 0.7 - $expectedFrom61_90 + $rollingAR.Days31_60 * 0.3
    $rollingAR.Days31_60 = $rollingAR.Days31_60 * 0.7 - $expectedFrom31_60 + $rollingAR.Days1_30 * 0.3
    $rollingAR.Days1_30 = $rollingAR.Days1_30 * 0.7 - $expectedFrom1_30 + $rollingAR.Current * 0.3
    $rollingAR.Current = [math]::Max(0, $rollingAR.Current * 0.7 - $expectedFromCurrent + $newARDue)
    
    $weekTotal = $actual + $predicted
    
    # Confidence based on how far out + data quality
    $confidence = if ($w -eq 0) { 85 } elseif ($w -eq 1) { 70 } elseif ($w -lt 4) { 55 } else { 40 }
    
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

Write-Host "Cash flow prediction complete (AR-based model)"

# === STEP 8: Shipments ===
Write-Host "Fetching shipment data..."
$shipments = Invoke-SyteLineAPI -Token $token -IDO "SLShipments" -Properties "ConsigneeCity,ConsigneeState,ConsigneeZip" -RecordCap 3000
$locationCounts = @{}
foreach ($ship in $shipments) {
    if ($ship.ConsigneeCity -and $ship.ConsigneeState) {
        $locKey = "$($ship.ConsigneeCity), $($ship.ConsigneeState)"
        if (-not $locationCounts.ContainsKey($locKey)) { $locationCounts[$locKey] = @{City=$ship.ConsigneeCity; State=$ship.ConsigneeState; Zip=$ship.ConsigneeZip; Count=0} }
        $locationCounts[$locKey].Count++
    }
}

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
$shippingLocations = $locationCounts.Values | Sort-Object -Property Count -Descending | Select-Object -First 100

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
    DailyTrend = $dailyTrend
    ShippingHeatMap = $shippingLocations
    DayOfWeekAverages = $avgByDayOfWeek
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
