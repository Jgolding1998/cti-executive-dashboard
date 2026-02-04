# CTI Executive Dashboard Generator v2
# FIXES:
# - Uses itmuf_cti_product_category for Service/Product classification
# - Calculates AR aging from invoice due dates (not Birst)
# - Includes invoice-level detail for drill-down
# - Adds P&L expense data
# - Pushes to GitHub after generation

param(
    [string]$OutputPath = "C:\Users\justi\clawd\dashboard",
    [switch]$PushToGitHub
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Web

# Load config
$configPath = Join-Path $PSScriptRoot "config.json"
if (Test-Path $configPath) {
    $config = Get-Content $configPath | ConvertFrom-Json
    $SyteLineConfig = @{
        BaseUrl = $config.syteline.baseUrl
        Tenant = $config.syteline.tenant
        Username = $config.syteline.username
        Password = $config.syteline.password
    }
} else {
    throw "config.json not found"
}

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
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 120
        if (-not $response.Success) { Write-Warning "API failed for $IDO"; return @() }
        return $response.Items
    } catch {
        Write-Warning "API error for $IDO`: $_"
        return @()
    }
}

function Get-BusinessDays { param([DateTime]$Start, [DateTime]$End)
    $days = 0; $cur = $Start
    while ($cur -le $End) { if ($cur.DayOfWeek -notin @('Saturday','Sunday')) { $days++ }; $cur = $cur.AddDays(1) }
    return $days
}

Write-Host "$(Get-Date) - CTI Dashboard Generator v2 Starting..."

$token = Get-SyteLineToken
Write-Host "Connected to SyteLine"

$today = (Get-Date).Date
$yesterday = $today.AddDays(-1)
while ($yesterday.DayOfWeek -in @('Saturday','Sunday')) { $yesterday = $yesterday.AddDays(-1) }
$monthStart = Get-Date -Day 1
$yearStart = Get-Date -Month 1 -Day 1
$currentYear = (Get-Date).Year

# === STEP 1: Build item category lookup using itmuf_cti_product_category ===
Write-Host "Fetching item master with itmuf_cti_product_category..."
$items = Invoke-SyteLineAPI -Token $token -IDO "SLItems" -Properties "Item,ProductCode,itmuf_cti_product_category" -RecordCap 10000

$itemCategory = @{}
$serviceCount = 0
$productCount = 0

foreach ($item in $items) {
    if ($item.Item) {
        $itemKey = $item.Item.Trim()
        $category = if ($item.itmuf_cti_product_category) { $item.itmuf_cti_product_category.Trim() } else { "" }
        
        # Service = category starts with "Service -"
        $isService = $category -like "Service -*"
        $itemCategory[$itemKey] = @{
            Category = $category
            IsService = $isService
            ProductCode = $item.ProductCode
        }
        if ($isService) { $serviceCount++ } else { $productCount++ }
    }
}
Write-Host "Built category lookup: $serviceCount Service items, $productCount Product items"

# === STEP 2: Calculate AR Aging from SLArTrans (Invoice - Payments) ===
Write-Host "Fetching AR transactions for aging calculation..."
$arTransAll = Invoke-SyteLineAPI -Token $token -IDO "SLArTrans" -Properties "InvNum,CadName,CoNum,InvDate,DueDate,Amount,Type,ApplyToInvNum" -RecordCap 15000
Write-Host "Got $($arTransAll.Count) AR transactions"

# Build invoice balances: Original Amount - Payments - Credits
$invoiceBalances = @{}
$openInvoices = @()
$arByCustomer = @{}
$arAging = @{ Current=0; Days1_30=0; Days31_60=0; Days61_90=0; Days90Plus=0; Total=0 }

# First pass: Add invoices (Type I)
foreach ($ar in ($arTransAll | Where-Object { $_.Type -eq "I" })) {
    if ($ar.InvNum) {
        $invNum = $ar.InvNum.Trim()
        if (-not $invoiceBalances.ContainsKey($invNum)) {
            $invoiceBalances[$invNum] = @{ Customer = $ar.CadName; DueDate = $ar.DueDate; CoNum = $ar.CoNum; Amount = 0; Paid = 0 }
        }
        $invoiceBalances[$invNum].Amount += [decimal]$ar.Amount
    }
}

# Second pass: Apply payments (Type P) and credits (Type C)
foreach ($ar in ($arTransAll | Where-Object { $_.Type -in @("P", "C") })) {
    if ($ar.Amount) {
        $applyTo = if ($ar.ApplyToInvNum) { $ar.ApplyToInvNum.Trim() } else { $null }
        if ($applyTo -and $invoiceBalances.ContainsKey($applyTo)) {
            $invoiceBalances[$applyTo].Paid += [decimal]$ar.Amount
        }
    }
}

# Calculate aging from open balances
foreach ($inv in $invoiceBalances.GetEnumerator()) {
    $balance = $inv.Value.Amount - $inv.Value.Paid
    if ($balance -gt 1) {  # Open invoice with >$1 balance
        $daysOverdue = 0
        $dueDate = $null
        if ($inv.Value.DueDate -and $inv.Value.DueDate.Length -ge 8) {
            try {
                $dueDate = [DateTime]::ParseExact($inv.Value.DueDate.Substring(0,8), "yyyyMMdd", $null)
                $daysOverdue = ($today - $dueDate).Days
            } catch {}
        }
        
        $bucket = if ($daysOverdue -le 0) { "Current" }
                  elseif ($daysOverdue -le 30) { "1-30 Days" }
                  elseif ($daysOverdue -le 60) { "31-60 Days" }
                  elseif ($daysOverdue -le 90) { "61-90 Days" }
                  else { "90+ Days" }
        
        $openInvoices += @{
            InvNum = $inv.Key
            Customer = $inv.Value.Customer
            CoNum = $inv.Value.CoNum
            DueDate = $inv.Value.DueDate
            Balance = $balance
            DaysOverdue = $daysOverdue
            Bucket = $bucket
        }
        
        # Add to bucket totals
        if ($daysOverdue -le 0) { $arAging.Current += $balance }
        elseif ($daysOverdue -le 30) { $arAging.Days1_30 += $balance }
        elseif ($daysOverdue -le 60) { $arAging.Days31_60 += $balance }
        elseif ($daysOverdue -le 90) { $arAging.Days61_90 += $balance }
        else { $arAging.Days90Plus += $balance }
        $arAging.Total += $balance
        
        # By customer
        $custName = $inv.Value.Customer
        if ($custName) {
            if (-not $arByCustomer.ContainsKey($custName)) {
                $arByCustomer[$custName] = @{ Current=0; Days1_30=0; Days31_60=0; Days61_90=0; Days90Plus=0; Total=0 }
            }
            if ($daysOverdue -le 0) { $arByCustomer[$custName].Current += $balance }
            elseif ($daysOverdue -le 30) { $arByCustomer[$custName].Days1_30 += $balance }
            elseif ($daysOverdue -le 60) { $arByCustomer[$custName].Days31_60 += $balance }
            elseif ($daysOverdue -le 90) { $arByCustomer[$custName].Days61_90 += $balance }
            else { $arByCustomer[$custName].Days90Plus += $balance }
            $arByCustomer[$custName].Total += $balance
        }
    }
}
Write-Host "Open invoices: $($openInvoices.Count)"
Write-Host "AR Aging: Current=`$$([math]::Round($arAging.Current,0)), 1-30=`$$([math]::Round($arAging.Days1_30,0)), 31-60=`$$([math]::Round($arAging.Days31_60,0)), 61-90=`$$([math]::Round($arAging.Days61_90,0)), 90+=`$$([math]::Round($arAging.Days90Plus,0)), Total=`$$([math]::Round($arAging.Total,0))"

# === STEP 3: Get invoice line items for drill-down ===
Write-Host "Fetching invoice line items..."
$coItems = Invoke-SyteLineAPI -Token $token -IDO "SLCoItems" -Properties "CoNum,Item,ItDescription,QtyInvoiced,DerExtInvoicedPrice" -RecordCap 20000
Write-Host "Got $($coItems.Count) line items"

# Build order breakdown with item-level detail
$orderBreakdown = @{}
$salesByProduct = @{}

foreach ($line in $coItems) {
    if ($line.CoNum -and $line.Item -and $line.DerExtInvoicedPrice -and [decimal]$line.DerExtInvoicedPrice -gt 0) {
        $orderNum = $line.CoNum.Trim()
        $itemKey = $line.Item.Trim()
        $amount = [decimal]$line.DerExtInvoicedPrice
        
        # Use itmuf_cti_product_category for classification
        $catInfo = $itemCategory[$itemKey]
        $isService = if ($catInfo) { $catInfo.IsService } else { $false }
        $categoryName = if ($catInfo -and $catInfo.Category) { $catInfo.Category } else { "Uncategorized" }
        
        # Order breakdown
        if (-not $orderBreakdown.ContainsKey($orderNum)) {
            $orderBreakdown[$orderNum] = @{ Service = 0; Product = 0; Total = 0; Items = @() }
        }
        if ($isService) { $orderBreakdown[$orderNum].Service += $amount }
        else { $orderBreakdown[$orderNum].Product += $amount }
        $orderBreakdown[$orderNum].Total += $amount
        $orderBreakdown[$orderNum].Items += @{
            Item = $itemKey
            Description = $line.ItDescription
            Qty = $line.QtyInvoiced
            Amount = $amount
            Category = $categoryName
            IsService = $isService
        }
        
        # Product sales tracking
        if (-not $salesByProduct.ContainsKey($itemKey)) {
            $salesByProduct[$itemKey] = @{ Description = $line.ItDescription; MTD = 0; Category = $categoryName; IsService = $isService }
        }
        $salesByProduct[$itemKey].MTD += $amount
    }
}
Write-Host "Built order breakdown for $($orderBreakdown.Count) orders"

# === STEP 4: Get AR transactions to map invoices to orders ===
Write-Host "Fetching AR transactions for invoice mapping..."
$arTrans = Invoke-SyteLineAPI -Token $token -IDO "SLArTrans" -Properties "InvNum,CoNum,CadName,InvDate,DueDate,Amount,Type" -Filter "Type = 'I'" -RecordCap 10000

$invToOrder = @{}
$invToCustomer = @{}
foreach ($ar in $arTrans) {
    if ($ar.InvNum -and $ar.CoNum) {
        $invToOrder[$ar.InvNum.Trim()] = $ar.CoNum.Trim()
    }
    if ($ar.InvNum -and $ar.CadName) {
        $invToCustomer[$ar.InvNum.Trim()] = $ar.CadName
    }
}
Write-Host "Mapped $($invToOrder.Count) invoices to orders"

# === STEP 5: Get order entry data (salesperson, ship-to) ===
Write-Host "Fetching orders..."
$orderFilter = "Invoiced = 1 AND OrderDate >= '$($yearStart.ToString("yyyyMMdd"))'"
$orders = Invoke-SyteLineAPI -Token $token -IDO "SLCoS" -Properties "CoNum,TakenBy,Price,OrderDate,Invoiced,ShipToCity,ShipToState" -Filter $orderFilter -RecordCap 8000
Write-Host "Got $($orders.Count) invoiced orders"

$orderToSalesperson = @{}
$orderToShipTo = @{}
foreach ($order in $orders) {
    if ($order.CoNum) {
        $coNum = $order.CoNum.Trim()
        if ($order.TakenBy) { $orderToSalesperson[$coNum] = $order.TakenBy.Trim() }
        if ($order.ShipToState) { 
            $orderToShipTo[$coNum] = @{ City = $order.ShipToCity; State = $order.ShipToState }
        }
    }
}

# === STEP 6: Get GL Ledger for sales AND expenses ===
Write-Host "Fetching GL ledger (revenue and expenses)..."

# Revenue accounts
$revenueFilter = "(Acct = '401000' OR Acct = '402000' OR Acct = '495000' OR Acct = '495400') AND ControlYear >= $currentYear"
$revenueData = Invoke-SyteLineAPI -Token $token -IDO "SLLedgers" -Properties "Acct,DomAmount,TransDate,Ref,ControlPeriod" -Filter $revenueFilter -RecordCap 15000

# Expense accounts (COGS and Operating Expenses)
$expenseFilter = "(Acct >= '500000' AND Acct < '600000') AND ControlYear >= $currentYear"
$expenseData = Invoke-SyteLineAPI -Token $token -IDO "SLLedgers" -Properties "Acct,DomAmount,TransDate,Ref,ControlPeriod" -Filter $expenseFilter -RecordCap 10000

Write-Host "Got $($revenueData.Count) revenue records, $($expenseData.Count) expense records"

# Process revenue with invoice-level tracking
$salesByDate = @{}
$salesByCustomer = @{}
$salesBySalesperson = @{}
$invoiceDetails = @()  # For drill-down

foreach ($gl in $revenueData) {
    $acct = $gl.Acct.Trim()
    if ($gl.TransDate.Length -ge 8) {
        try {
            $transDate = [DateTime]::ParseExact($gl.TransDate.Substring(0,8), "yyyyMMdd", $null)
            # Adjust weekend dates to Monday
            if ($transDate.DayOfWeek -eq 'Saturday') { $transDate = $transDate.AddDays(2) }
            if ($transDate.DayOfWeek -eq 'Sunday') { $transDate = $transDate.AddDays(1) }
            $dateKey = $transDate.ToString("yyyy-MM-dd")
            $amount = [decimal]$gl.DomAmount * -1  # Revenue is credit (negative)
            
            if ($amount -le 0) { continue }  # Skip credits/adjustments
            
            if (-not $salesByDate.ContainsKey($dateKey)) {
                $salesByDate[$dateKey] = @{ Product=0; Service=0; Freight=0; Miscellaneous=0; Total=0; Invoices=@() }
            }
            
            $category = "Product"
            $invNum = $null
            $orderNum = $null
            $customer = $null
            $salesperson = $null
            $svcRatio = 0
            $prodRatio = 1
            
            if ($acct -eq "495400") {
                $category = "Freight"
                $salesByDate[$dateKey].Freight += $amount
            } elseif ($acct -eq "495000") {
                $category = "Miscellaneous"
                $salesByDate[$dateKey].Miscellaneous += $amount
            } elseif ($acct -in @("401000", "402000") -and $gl.Ref -match "ARI\s+(\d+)") {
                $invNum = $matches[1].Trim()
                $orderNum = $invToOrder[$invNum]
                $customer = $invToCustomer[$invNum]
                $salesperson = if ($orderNum) { $orderToSalesperson[$orderNum] } else { $null }
                
                # Calculate service/product ratio from order line items
                if ($orderNum -and $orderBreakdown.ContainsKey($orderNum) -and $orderBreakdown[$orderNum].Total -gt 0) {
                    $breakdown = $orderBreakdown[$orderNum]
                    $svcRatio = $breakdown.Service / $breakdown.Total
                    $prodRatio = $breakdown.Product / $breakdown.Total
                    
                    $salesByDate[$dateKey].Service += $amount * $svcRatio
                    $salesByDate[$dateKey].Product += $amount * $prodRatio
                } else {
                    $salesByDate[$dateKey].Product += $amount
                }
                
                $category = if ($svcRatio -gt 0.5) { "Service" } else { "Product" }
            }
            
            $salesByDate[$dateKey].Total += $amount
            
            # Track invoice for drill-down
            if ($invNum) {
                $salesByDate[$dateKey].Invoices += @{
                    InvNum = $invNum
                    OrderNum = $orderNum
                    Customer = $customer
                    Salesperson = $salesperson
                    Amount = [math]::Round($amount, 2)
                    ServiceAmount = [math]::Round($amount * $svcRatio, 2)
                    ProductAmount = [math]::Round($amount * $prodRatio, 2)
                    Category = $category
                }
            }
            
            # Customer tracking
            if ($customer) {
                if (-not $salesByCustomer.ContainsKey($customer)) {
                    $salesByCustomer[$customer] = @{ MTD=0; Yesterday=0; YTD=0; Invoices=@() }
                }
                if ($transDate -ge $yearStart -and $transDate -le $today) { $salesByCustomer[$customer].YTD += $amount }
                if ($transDate -ge $monthStart -and $transDate -le $today) { 
                    $salesByCustomer[$customer].MTD += $amount
                    if ($invNum) { $salesByCustomer[$customer].Invoices += $invNum }
                }
                if ($transDate -eq $yesterday) { $salesByCustomer[$customer].Yesterday += $amount }
            }
            
            # Salesperson tracking
            if ($salesperson) {
                if (-not $salesBySalesperson.ContainsKey($salesperson)) {
                    $salesBySalesperson[$salesperson] = @{ MTD=0; Yesterday=0; YTD=0; MTDProduct=0; MTDService=0; InvoiceCount=0 }
                }
                if ($transDate -ge $yearStart) { $salesBySalesperson[$salesperson].YTD += $amount }
                if ($transDate -ge $monthStart) { 
                    $salesBySalesperson[$salesperson].MTD += $amount
                    $salesBySalesperson[$salesperson].MTDProduct += $amount * $prodRatio
                    $salesBySalesperson[$salesperson].MTDService += $amount * $svcRatio
                    $salesBySalesperson[$salesperson].InvoiceCount++
                }
                if ($transDate -eq $yesterday) { $salesBySalesperson[$salesperson].Yesterday += $amount }
            }
        } catch { Write-Warning "Error processing GL record: $_" }
    }
}

# Process expenses for P&L
$expensesByMonth = @{}
foreach ($gl in $expenseData) {
    if ($gl.TransDate.Length -ge 8) {
        try {
            $transDate = [DateTime]::ParseExact($gl.TransDate.Substring(0,8), "yyyyMMdd", $null)
            $monthKey = $transDate.ToString("yyyy-MM")
            $amount = [decimal]$gl.DomAmount  # Expenses are debits (positive)
            
            if (-not $expensesByMonth.ContainsKey($monthKey)) {
                $expensesByMonth[$monthKey] = @{ COGS=0; Operating=0; Total=0 }
            }
            
            $acct = [int]$gl.Acct
            if ($acct -ge 500000 -and $acct -lt 510000) {
                $expensesByMonth[$monthKey].COGS += $amount
            } else {
                $expensesByMonth[$monthKey].Operating += $amount
            }
            $expensesByMonth[$monthKey].Total += $amount
        } catch {}
    }
}

# === STEP 7: Calculate summaries ===
$yesterdaySales = if ($salesByDate.ContainsKey($yesterday.ToString("yyyy-MM-dd"))) { $salesByDate[$yesterday.ToString("yyyy-MM-dd")] } else { @{Product=0;Service=0;Freight=0;Miscellaneous=0;Total=0;Invoices=@()} }
$mtdSales = @{Product=0;Service=0;Freight=0;Miscellaneous=0;Total=0}
$ytdSales = @{Product=0;Service=0;Freight=0;Miscellaneous=0;Total=0}
$mtdInvoices = @()
$ytdInvoices = @()

foreach ($dateKey in $salesByDate.Keys) {
    $date = [DateTime]::Parse($dateKey)
    $data = $salesByDate[$dateKey]
    if ($date -ge $monthStart -and $date -le $today) {
        foreach ($cat in @("Product","Service","Freight","Miscellaneous","Total")) { $mtdSales[$cat] += $data[$cat] }
        $mtdInvoices += $data.Invoices
    }
    if ($date -ge $yearStart -and $date -le $today) {
        foreach ($cat in @("Product","Service","Freight","Miscellaneous","Total")) { $ytdSales[$cat] += $data[$cat] }
    }
}

$mtdDays = Get-BusinessDays $monthStart $today
$ytdDays = Get-BusinessDays $yearStart $today
$mtdAvg = if ($mtdDays -gt 0) { [math]::Round($mtdSales.Total / $mtdDays, 2) } else { 0 }
$ytdAvg = if ($ytdDays -gt 0) { [math]::Round($ytdSales.Total / $ytdDays, 2) } else { 0 }

Write-Host "MTD: `$$([math]::Round($mtdSales.Total,0)) (P:`$$([math]::Round($mtdSales.Product,0)) S:`$$([math]::Round($mtdSales.Service,0)))"
Write-Host "YTD: `$$([math]::Round($ytdSales.Total,0))"

# === STEP 8: Day of Week Averages ===
$dowTotals = @{ Monday=@(); Tuesday=@(); Wednesday=@(); Thursday=@(); Friday=@() }
for ($i = 90; $i -ge 0; $i--) {
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
        $avgByDayOfWeek[$dow] = [math]::Round(($values | Measure-Object -Average).Average, 2)
    } else {
        $avgByDayOfWeek[$dow] = 0
    }
}
Write-Host "DOW Averages: Mon=`$$($avgByDayOfWeek['Monday']) Tue=`$$($avgByDayOfWeek['Tuesday']) Wed=`$$($avgByDayOfWeek['Wednesday']) Thu=`$$($avgByDayOfWeek['Thursday']) Fri=`$$($avgByDayOfWeek['Friday'])"

# === STEP 9: Build output data ===
$dailyTrend = @()
for ($i = 60; $i -ge 0; $i--) {
    $date = $today.AddDays(-$i)
    if ($date.DayOfWeek -in @('Saturday','Sunday')) { continue }
    $dateKey = $date.ToString("yyyy-MM-dd")
    $data = if ($salesByDate.ContainsKey($dateKey)) { $salesByDate[$dateKey] } else { @{Product=0;Service=0;Freight=0;Miscellaneous=0;Total=0;Invoices=@()} }
    $dailyTrend += @{ 
        Date=$dateKey
        DayOfWeek=$date.DayOfWeek.ToString()
        Product=[math]::Round($data.Product,2)
        Service=[math]::Round($data.Service,2)
        Freight=[math]::Round($data.Freight,2)
        Miscellaneous=[math]::Round($data.Miscellaneous,2)
        Total=[math]::Round($data.Total,2)
        InvoiceCount=$data.Invoices.Count
        Invoices=$data.Invoices
    }
}

# Top customers/products
$topCustomersMTD = $salesByCustomer.GetEnumerator() | Where-Object { $_.Value.MTD -gt 0 } | 
    Sort-Object { $_.Value.MTD } -Descending | Select-Object -First 20 | 
    ForEach-Object { @{Name=$_.Key; Amount=[math]::Round($_.Value.MTD,2); InvoiceCount=$_.Value.Invoices.Count} }

$topCustomersYesterday = $salesByCustomer.GetEnumerator() | Where-Object { $_.Value.Yesterday -gt 0 } | 
    Sort-Object { $_.Value.Yesterday } -Descending | Select-Object -First 15 | 
    ForEach-Object { @{Name=$_.Key; Amount=[math]::Round($_.Value.Yesterday,2)} }

$topProductsMTD = $salesByProduct.GetEnumerator() | Sort-Object { $_.Value.MTD } -Descending | Select-Object -First 20 | 
    ForEach-Object { @{Item=$_.Key; Description=$_.Value.Description; Amount=[math]::Round($_.Value.MTD,2); Category=$_.Value.Category; IsService=$_.Value.IsService} }

# AR by customer for table
$arCustomerTable = $arByCustomer.GetEnumerator() | Sort-Object { $_.Value.Total } -Descending | Select-Object -First 50 |
    ForEach-Object { @{
        Customer = $_.Key
        Current = [math]::Round($_.Value.Current, 2)
        Days1_30 = [math]::Round($_.Value.Days1_30, 2)
        Days31_60 = [math]::Round($_.Value.Days31_60, 2)
        Days61_90 = [math]::Round($_.Value.Days61_90, 2)
        Days90Plus = [math]::Round($_.Value.Days90Plus, 2)
        Total = [math]::Round($_.Value.Total, 2)
    }}

# Shipping locations
$shippingLocations = @()
$locationCounts = @{}
foreach ($order in $orders) {
    if ($order.ShipToCity -and $order.ShipToState) {
        $locKey = "$($order.ShipToCity.Trim()), $($order.ShipToState.Trim())"
        if (-not $locationCounts.ContainsKey($locKey)) {
            $locationCounts[$locKey] = @{ City=$order.ShipToCity; State=$order.ShipToState; Count=0 }
        }
        $locationCounts[$locKey].Count++
    }
}
$shippingLocations = $locationCounts.Values | Sort-Object -Property Count -Descending | Select-Object -First 100

# State breakdown for shipping
$stateBreakdown = @{}
foreach ($loc in $shippingLocations) {
    $state = $loc.State.Trim()
    if (-not $stateBreakdown.ContainsKey($state)) { $stateBreakdown[$state] = 0 }
    $stateBreakdown[$state] += $loc.Count
}

# Salesperson data with roles
$teamRoles = @{
    'Joel' = @{ Name='Joel EuDaly'; Role='COO'; Type='Executive' }
    'Allison' = @{ Name='Allison White'; Role='Accounting'; Type='Admin' }
    'lphil' = @{ Name='Phil Libbert'; Role='Inside Sales'; Type='Sales' }
    'cjord' = @{ Name='Jordan Cline'; Role='Client Care'; Type='ClientCare' }
    'nandy' = @{ Name='Andy Neptune'; Role='Outside Sales'; Type='Sales' }
    'marissa' = @{ Name='Marissa King'; Role='Client Care'; Type='ClientCare' }
    'ebrow' = @{ Name='Emily Brown'; Role='Client Care'; Type='ClientCare' }
    'bcoli' = @{ Name='Colin Bowles'; Role='Scheduling Coordinator'; Type='ServiceTeam' }
    'hnich' = @{ Name='Nicholas Hegger'; Role='In-House Service Supervisor'; Type='ServiceTeam' }
    'Greg' = @{ Name='Greg Smith'; Role='Inside Sales'; Type='Sales' }
}

$allTeamMTD = $salesBySalesperson.GetEnumerator() | Where-Object { $_.Value.MTD -gt 0 } | 
    Sort-Object { $_.Value.MTD } -Descending | ForEach-Object {
        $roleInfo = if ($teamRoles.ContainsKey($_.Key)) { $teamRoles[$_.Key] } else { @{ Name=$_.Key; Role='Team Member'; Type='Sales' } }
        @{
            Username = $_.Key
            Name = $roleInfo.Name
            Role = $roleInfo.Role
            Type = $roleInfo.Type
            Amount = [math]::Round($_.Value.MTD, 2)
            ProductAmount = [math]::Round($_.Value.MTDProduct, 2)
            ServiceAmount = [math]::Round($_.Value.MTDService, 2)
            InvoiceCount = $_.Value.InvoiceCount
        }
    }

$allTeamYesterday = $salesBySalesperson.GetEnumerator() | Where-Object { $_.Value.Yesterday -gt 0 } |
    Sort-Object { $_.Value.Yesterday } -Descending | ForEach-Object {
        $roleInfo = if ($teamRoles.ContainsKey($_.Key)) { $teamRoles[$_.Key] } else { @{ Name=$_.Key; Role='Team Member'; Type='Sales' } }
        @{
            Username = $_.Key
            Name = $roleInfo.Name
            Role = $roleInfo.Role
            Type = $roleInfo.Type
            Amount = [math]::Round($_.Value.Yesterday, 2)
        }
    }

# P&L Summary
$currentMonthKey = $today.ToString("yyyy-MM")
$mtdExpenses = if ($expensesByMonth.ContainsKey($currentMonthKey)) { $expensesByMonth[$currentMonthKey] } else { @{COGS=0;Operating=0;Total=0} }
$mtdGrossProfit = $mtdSales.Total - $mtdExpenses.COGS
$mtdNetIncome = $mtdGrossProfit - $mtdExpenses.Operating

# === Build final data object ===
$dashboardData = @{
    GeneratedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    DataSource = "SyteLine GL + AR (itmuf_cti_product_category for Service/Product)"
    Period = @{
        Yesterday = $yesterday.ToString("yyyy-MM-dd")
        MonthStart = $monthStart.ToString("yyyy-MM-dd")
        YearStart = $yearStart.ToString("yyyy-MM-dd")
        MTDBusinessDays = $mtdDays
        YTDBusinessDays = $ytdDays
    }
    Summary = @{
        Yesterday = @{
            Product = [math]::Round($yesterdaySales.Product, 2)
            Service = [math]::Round($yesterdaySales.Service, 2)
            Freight = [math]::Round($yesterdaySales.Freight, 2)
            Miscellaneous = [math]::Round($yesterdaySales.Miscellaneous, 2)
            Total = [math]::Round($yesterdaySales.Total, 2)
            Invoices = $yesterdaySales.Invoices
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
        MTDDailyAverage = $mtdAvg
        YTDDailyAverage = $ytdAvg
    }
    PnL = @{
        MTD = @{
            Revenue = [math]::Round($mtdSales.Total, 2)
            COGS = [math]::Round($mtdExpenses.COGS, 2)
            GrossProfit = [math]::Round($mtdGrossProfit, 2)
            OperatingExpenses = [math]::Round($mtdExpenses.Operating, 2)
            NetIncome = [math]::Round($mtdNetIncome, 2)
        }
    }
    ARaging = @{
        Current = [math]::Round($arAging.Current, 2)
        Days1_30 = [math]::Round($arAging.Days1_30, 2)
        Days31_60 = [math]::Round($arAging.Days31_60, 2)
        Days61_90 = [math]::Round($arAging.Days61_90, 2)
        Days90Plus = [math]::Round($arAging.Days90Plus, 2)
        Total = [math]::Round($arAging.Total, 2)
    }
    ARbyCustomer = $arCustomerTable
    TopCustomers = @{ Yesterday = $topCustomersYesterday; MTD = $topCustomersMTD }
    TopProducts = @{ MTD = $topProductsMTD }
    AllTeam = @{ Yesterday = $allTeamYesterday; MTD = $allTeamMTD }
    Salesperson = @{ 
        Yesterday = ($allTeamYesterday | Where-Object { $_.Type -eq 'Sales' })
        MTD = ($allTeamMTD | Where-Object { $_.Type -eq 'Sales' })
    }
    ServiceTeam = @{
        Yesterday = ($allTeamYesterday | Where-Object { $_.Type -eq 'ServiceTeam' })
        MTD = ($allTeamMTD | Where-Object { $_.Type -eq 'ServiceTeam' })
    }
    ClientCare = @{
        Yesterday = ($allTeamYesterday | Where-Object { $_.Type -eq 'ClientCare' })
        MTD = ($allTeamMTD | Where-Object { $_.Type -eq 'ClientCare' })
    }
    DailyTrend = $dailyTrend
    DayOfWeekAverages = $avgByDayOfWeek
    ShippingHeatMap = $shippingLocations
    ShippingStats = @{
        TotalUSStates = $stateBreakdown.Count
        StateBreakdown = $stateBreakdown
    }
    MTDInvoices = $mtdInvoices | Select-Object -First 500
    FieldTechnicians = @()
    FieldTechSummary = @{ TotalTechs=0; TotalAppointmentsMTD=0; TotalStatesCovered=0; TopPerformer="--" }
    CashFlowPrediction = @()  # TODO: Add proper cash flow
}

# Save JSON
$jsonPath = Join-Path $OutputPath "data\dashboard-data.json"
if (-not (Test-Path (Split-Path $jsonPath))) { New-Item -ItemType Directory -Path (Split-Path $jsonPath) -Force | Out-Null }
$dashboardData | ConvertTo-Json -Depth 15 | Set-Content $jsonPath -Encoding UTF8
Write-Host "Saved JSON: $jsonPath"

# Create embedded HTML
$htmlTemplatePath = Join-Path $OutputPath "cti-dashboard.html"
if (Test-Path $htmlTemplatePath) {
    $htmlTemplate = Get-Content $htmlTemplatePath -Raw
    $jsonData = $dashboardData | ConvertTo-Json -Depth 15
    $embeddedHtml = $htmlTemplate -replace 'let DATA = null;', "let DATA = $jsonData;"
    $liveHtmlPath = Join-Path $OutputPath "cti-dashboard-live.html"
    $embeddedHtml | Set-Content $liveHtmlPath -Encoding UTF8
    Write-Host "Dashboard saved: $liveHtmlPath"
}

# === STEP 10: Push to GitHub if requested ===
if ($PushToGitHub) {
    Write-Host "`nPushing to GitHub..."
    $webRepoPath = "C:\Users\justi\clawd\cti-dashboard-web"
    
    if (Test-Path $webRepoPath) {
        # Copy updated files
        Copy-Item $liveHtmlPath (Join-Path $webRepoPath "index.html") -Force
        Copy-Item $jsonPath (Join-Path $webRepoPath "data\dashboard-data.json") -Force
        
        # Git commit and push
        Push-Location $webRepoPath
        try {
            git add -A
            $commitMsg = "Auto-update: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
            git commit -m $commitMsg 2>&1 | Out-Null
            git push origin main 2>&1
            Write-Host "Pushed to GitHub successfully"
        } catch {
            Write-Warning "Git push failed: $_"
        }
        Pop-Location
    } else {
        Write-Warning "Web repo not found at $webRepoPath"
    }
}

Write-Host "`n$(Get-Date) - Dashboard generation complete!"
Write-Host "=== FINAL SUMMARY ==="
Write-Host "Yesterday: `$$([math]::Round($yesterdaySales.Total,0))"
Write-Host "MTD: `$$([math]::Round($mtdSales.Total,0)) ($mtdDays days)"
Write-Host "YTD: `$$([math]::Round($ytdSales.Total,0)) ($ytdDays days)"
Write-Host "AR Total: `$$([math]::Round($arAging.Total,0))"
