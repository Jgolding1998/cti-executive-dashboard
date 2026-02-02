# Generate Power BI Data with Invoice-Level Detail for Drill-Down
# This creates proper fact/dimension tables matching the HTML dashboard

param(
    [string]$OutputPath = "C:\Users\justi\clawd\dashboard\powerbi-data"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Web

# Load config
$configPath = "C:\Users\justi\clawd\dashboard\config.json"
$config = Get-Content $configPath | ConvertFrom-Json
$SyteLineConfig = @{
    BaseUrl = $config.syteline.baseUrl
    Tenant = $config.syteline.tenant
    Username = $config.syteline.username
    Password = $config.syteline.password
}

function Get-SyteLineToken {
    $tokenUrl = "$($SyteLineConfig.BaseUrl)/token/$($SyteLineConfig.Tenant)/$($SyteLineConfig.Username)/$([System.Web.HttpUtility]::UrlEncode($SyteLineConfig.Password))"
    $response = Invoke-RestMethod -Uri $tokenUrl
    if (-not $response.Success) { throw "Failed to get SyteLine token" }
    return $response.Token
}

function Invoke-SyteLineAPI {
    param([string]$Token, [string]$IDO, [string]$Properties = "", [string]$Filter = "", [int]$RecordCap = 5000, [string]$OrderBy = "")
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

Write-Host "$(Get-Date) - Generating Power BI data with invoice-level detail..."

$token = Get-SyteLineToken
Write-Host "Connected to SyteLine"

$today = (Get-Date).Date
$yearStart = Get-Date -Month 1 -Day 1

# Team roles mapping
$teamRoles = @{
    'Joel' = @{ FullName='Joel EuDaly'; Role='COO'; Type='Executive' }
    'Allison' = @{ FullName='Allison White'; Role='Accounting'; Type='Admin' }
    'lphil' = @{ FullName='Phil Libbert'; Role='Inside Sales'; Type='Sales' }
    'mrand' = @{ FullName='Randy Mobley'; Role='Inside Sales'; Type='Sales' }
    'tbolt' = @{ FullName='Trevor Bolte'; Role='Inside Sales'; Type='Sales' }
    'Greg' = @{ FullName='Greg Smith'; Role='Inside Sales'; Type='Sales' }
    'nandy' = @{ FullName='Andy Neptune'; Role='Outside Sales'; Type='Sales' }
    'cjord' = @{ FullName='Jordan Cline'; Role='Client Care'; Type='ClientCare' }
    'marissa' = @{ FullName='Marissa King'; Role='Client Care'; Type='ClientCare' }
    'ebrow' = @{ FullName='Emily Brown'; Role='Client Care'; Type='ClientCare' }
    'seric' = @{ FullName='Erica Thurmond'; Role='Service Scheduling'; Type='ServiceTeam' }
    'bcoli' = @{ FullName='Colin Bowles'; Role='Scheduling Coordinator'; Type='ServiceTeam' }
    'hnich' = @{ FullName='Nicholas Hegger'; Role='In-House Service Supervisor'; Type='ServiceTeam' }
    'ljami' = @{ FullName='Jami Lorenz'; Role='Service Director'; Type='ServiceTeam' }
    'Rachel L' = @{ FullName='Rachel Lively'; Role='Service Support'; Type='ServiceTeam' }
}

# === 1. Get Item Master for categorization ===
Write-Host "Fetching item master..."
$items = Invoke-SyteLineAPI -Token $token -IDO "SLItems" -Properties "Item,Description,ProductCode" -RecordCap 5000
$itemLookup = @{}
$ServiceProductCodes = @("LFTR", "LGAS", "LIHL")
$MiscProductCodes = @("LMSC", "XNIO")
foreach ($item in $items) {
    if ($item.Item) {
        $cat = "Product"
        if ($item.ProductCode -in $ServiceProductCodes) { $cat = "Service" }
        elseif ($item.ProductCode -in $MiscProductCodes) { $cat = "Miscellaneous" }
        $itemLookup[$item.Item.Trim()] = @{ Description=$item.Description; Category=$cat; ProductCode=$item.ProductCode }
    }
}
Write-Host "Built item lookup: $($itemLookup.Count) items"

# === 2. Get Order Line Items ===
Write-Host "Fetching order line items..."
$coItems = Invoke-SyteLineAPI -Token $token -IDO "SLCoItems" -Properties "CoNum,CoLine,Item,ItDescription,QtyInvoiced,DerExtInvoicedPrice" -RecordCap 15000
$orderLineItems = @{}
foreach ($line in $coItems) {
    if ($line.CoNum -and $line.Item -and $line.DerExtInvoicedPrice) {
        $orderNum = $line.CoNum.Trim()
        $amount = [decimal]$line.DerExtInvoicedPrice
        if ($amount -le 0) { continue }
        
        if (-not $orderLineItems.ContainsKey($orderNum)) {
            $orderLineItems[$orderNum] = @()
        }
        $itemInfo = $itemLookup[$line.Item.Trim()]
        $cat = if ($itemInfo) { $itemInfo.Category } else { "Product" }
        $orderLineItems[$orderNum] += @{
            Item = $line.Item.Trim()
            Description = $line.ItDescription
            Category = $cat
            Amount = $amount
            Qty = [decimal]$line.QtyInvoiced
        }
    }
}
Write-Host "Built order line items: $($orderLineItems.Count) orders"

# === 3. Get Orders with Salesperson ===
Write-Host "Fetching orders..."
$orderFilter = "Invoiced = 1 AND OrderDate >= '20260101'"
$orders = Invoke-SyteLineAPI -Token $token -IDO "SLCoS" -Properties "CoNum,TakenBy,OrderDate,ShipToCity,ShipToState,CustNum" -Filter $orderFilter -RecordCap 5000
$orderLookup = @{}
foreach ($order in $orders) {
    if ($order.CoNum) {
        $orderLookup[$order.CoNum.Trim()] = @{
            TakenBy = $order.TakenBy
            OrderDate = $order.OrderDate
            ShipToCity = $order.ShipToCity
            ShipToState = $order.ShipToState
            CustNum = $order.CustNum
        }
    }
}
Write-Host "Built order lookup: $($orderLookup.Count) orders"

# === 4. Get AR Transactions (Invoices) ===
Write-Host "Fetching AR transactions..."
$arTrans = Invoke-SyteLineAPI -Token $token -IDO "SLArTrans" -Properties "InvNum,CoNum,CadName,InvDate,DueDate,Amount,Type,ApplyToInvNum" -RecordCap 10000

# Build invoice balances
$invoiceBalances = @{}
foreach ($ar in $arTrans) {
    if ($ar.Type -eq "I" -and $ar.InvNum) {
        $invNum = $ar.InvNum.Trim()
        if (-not $invoiceBalances.ContainsKey($invNum)) {
            $invoiceBalances[$invNum] = @{ 
                Customer = $ar.CadName
                CoNum = $ar.CoNum
                InvDate = $ar.InvDate
                DueDate = $ar.DueDate
                OriginalAmount = 0
                Applied = 0 
            }
        }
        $invoiceBalances[$invNum].OriginalAmount += [decimal]$ar.Amount
    }
}
foreach ($ar in $arTrans) {
    if ($ar.Type -in @("P", "C") -and $ar.Amount) {
        $applyTo = if ($ar.ApplyToInvNum) { $ar.ApplyToInvNum.Trim() } else { $null }
        if ($applyTo -and $invoiceBalances.ContainsKey($applyTo)) {
            $invoiceBalances[$applyTo].Applied += [decimal]$ar.Amount
        }
    }
}
Write-Host "Built invoice balances: $($invoiceBalances.Count) invoices"

# === 5. Get GL Ledger for actual sales by date ===
Write-Host "Fetching GL ledger..."
$glFilter = "(Acct = '401000' OR Acct = '402000' OR Acct = '495000' OR Acct = '495400') AND ControlYear >= 2026"
$ledgerData = Invoke-SyteLineAPI -Token $token -IDO "SLLedgers" -Properties "Acct,DomAmount,TransDate,Ref" -Filter $glFilter -RecordCap 10000
Write-Host "Got $($ledgerData.Count) GL records"

# === 6. Build Invoice Detail Table (FACT) ===
Write-Host "Building invoice detail table..."
$invoiceDetails = @()

foreach ($gl in $ledgerData) {
    if ($gl.TransDate.Length -lt 8) { continue }
    
    try {
        $transDate = [DateTime]::ParseExact($gl.TransDate.Substring(0,8), "yyyyMMdd", $null)
    } catch { continue }
    
    $amount = [decimal]$gl.DomAmount * -1  # Revenue is credit
    if ($amount -le 0) { continue }
    
    $acct = $gl.Acct.Trim()
    $invNum = ""
    $customer = ""
    $salesperson = ""
    $salespersonName = ""
    $role = ""
    $type = ""
    $product = 0
    $service = 0
    $freight = 0
    $misc = 0
    $city = ""
    $state = ""
    
    # Parse invoice number from Ref
    if ($gl.Ref -match "ARI\s+(\d+)") {
        $invNum = $matches[1].Trim()
        
        # Get customer from invoice
        if ($invoiceBalances.ContainsKey($invNum)) {
            $customer = $invoiceBalances[$invNum].Customer
            $coNum = $invoiceBalances[$invNum].CoNum
            
            # Get salesperson and location from order
            if ($coNum -and $orderLookup.ContainsKey($coNum.Trim())) {
                $orderInfo = $orderLookup[$coNum.Trim()]
                $salesperson = $orderInfo.TakenBy
                $city = $orderInfo.ShipToCity
                $state = $orderInfo.ShipToState
                
                if ($teamRoles.ContainsKey($salesperson)) {
                    $salespersonName = $teamRoles[$salesperson].FullName
                    $role = $teamRoles[$salesperson].Role
                    $type = $teamRoles[$salesperson].Type
                } else {
                    $salespersonName = $salesperson
                    $role = "Team Member"
                    $type = "Sales"
                }
            }
            
            # Calculate product/service split from line items
            if ($coNum -and $orderLineItems.ContainsKey($coNum.Trim())) {
                $lines = $orderLineItems[$coNum.Trim()]
                $total = 0
                $svcAmt = 0
                $miscAmt = 0
                $prodAmt = 0
                foreach ($ln in $lines) {
                    $total += $ln.Amount
                    if ($ln.Category -eq "Service") { $svcAmt += $ln.Amount }
                    elseif ($ln.Category -eq "Miscellaneous") { $miscAmt += $ln.Amount }
                    else { $prodAmt += $ln.Amount }
                }
                if ($total -gt 0) {
                    $product = $amount * ($prodAmt / $total)
                    $service = $amount * ($svcAmt / $total)
                    $misc = $amount * ($miscAmt / $total)
                }
            }
        }
    }
    
    # Categorize by account
    if ($acct -eq "495400") {
        $freight = $amount
        $product = 0; $service = 0; $misc = 0
    } elseif ($acct -eq "495000") {
        $misc = $amount
        $product = 0; $service = 0; $freight = 0
    } elseif ($product -eq 0 -and $service -eq 0 -and $misc -eq 0) {
        # Default to product if no breakdown
        $product = $amount
    }
    
    $invoiceDetails += [PSCustomObject]@{
        Date = $transDate.ToString("yyyy-MM-dd")
        InvoiceNum = $invNum
        Customer = $customer
        Salesperson = $salespersonName
        SalespersonUsername = $salesperson
        Role = $role
        TeamType = $type
        City = $city
        State = $state
        Product = [math]::Round($product, 2)
        Service = [math]::Round($service, 2)
        Freight = [math]::Round($freight, 2)
        Miscellaneous = [math]::Round($misc, 2)
        Total = [math]::Round($amount, 2)
        Account = $acct
    }
}

Write-Host "Built $($invoiceDetails.Count) invoice detail records"

# === 7. Build AR Aging Detail Table ===
Write-Host "Building AR aging detail..."
$arAgingDetail = @()

foreach ($inv in $invoiceBalances.GetEnumerator()) {
    $balance = $inv.Value.OriginalAmount - $inv.Value.Applied
    if ($balance -le 1) { continue }
    
    try {
        $dueDate = [DateTime]::ParseExact($inv.Value.DueDate.Substring(0,8), "yyyyMMdd", $null)
    } catch { continue }
    
    $daysOverdue = ($today - $dueDate).Days
    $bucket = if ($daysOverdue -le 0) { "Current" }
              elseif ($daysOverdue -le 30) { "1-30 Days" }
              elseif ($daysOverdue -le 60) { "31-60 Days" }
              elseif ($daysOverdue -le 90) { "61-90 Days" }
              else { "90+ Days" }
    
    $arAgingDetail += [PSCustomObject]@{
        InvoiceNum = $inv.Key
        Customer = $inv.Value.Customer
        DueDate = $dueDate.ToString("yyyy-MM-dd")
        DaysOverdue = $daysOverdue
        Bucket = $bucket
        Balance = [math]::Round($balance, 2)
    }
}

Write-Host "Built $($arAgingDetail.Count) AR aging records"

# === 8. Build Summary Tables ===

# Daily Sales
$dailySales = $invoiceDetails | Group-Object Date | ForEach-Object {
    [PSCustomObject]@{
        Date = $_.Name
        Product = [math]::Round(($_.Group | Measure-Object -Property Product -Sum).Sum, 2)
        Service = [math]::Round(($_.Group | Measure-Object -Property Service -Sum).Sum, 2)
        Freight = [math]::Round(($_.Group | Measure-Object -Property Freight -Sum).Sum, 2)
        Miscellaneous = [math]::Round(($_.Group | Measure-Object -Property Miscellaneous -Sum).Sum, 2)
        Total = [math]::Round(($_.Group | Measure-Object -Property Total -Sum).Sum, 2)
    }
} | Sort-Object Date

# Top Customers
$topCustomers = $invoiceDetails | Where-Object { $_.Customer } | Group-Object Customer | ForEach-Object {
    [PSCustomObject]@{
        Customer = $_.Name
        Total = [math]::Round(($_.Group | Measure-Object -Property Total -Sum).Sum, 2)
        InvoiceCount = $_.Count
    }
} | Sort-Object Total -Descending | Select-Object -First 50

# Salesperson Performance
$salespersonPerf = $invoiceDetails | Where-Object { $_.Salesperson } | Group-Object Salesperson | ForEach-Object {
    $group = $_.Group
    [PSCustomObject]@{
        Salesperson = $_.Name
        Role = ($group | Select-Object -First 1).Role
        TeamType = ($group | Select-Object -First 1).TeamType
        Product = [math]::Round(($group | Measure-Object -Property Product -Sum).Sum, 2)
        Service = [math]::Round(($group | Measure-Object -Property Service -Sum).Sum, 2)
        Freight = [math]::Round(($group | Measure-Object -Property Freight -Sum).Sum, 2)
        Miscellaneous = [math]::Round(($group | Measure-Object -Property Miscellaneous -Sum).Sum, 2)
        Total = [math]::Round(($group | Measure-Object -Property Total -Sum).Sum, 2)
        InvoiceCount = $_.Count
    }
} | Sort-Object Total -Descending

# AR Aging Summary
$arAgingSummary = $arAgingDetail | Group-Object Bucket | ForEach-Object {
    [PSCustomObject]@{
        Bucket = $_.Name
        Amount = [math]::Round(($_.Group | Measure-Object -Property Balance -Sum).Sum, 2)
        InvoiceCount = $_.Count
    }
}

# Geography
$geography = $invoiceDetails | Where-Object { $_.State } | Group-Object State | ForEach-Object {
    [PSCustomObject]@{
        State = $_.Name
        Total = [math]::Round(($_.Group | Measure-Object -Property Total -Sum).Sum, 2)
        InvoiceCount = $_.Count
    }
} | Sort-Object Total -Descending

# === 9. Export to CSV ===
Write-Host "Exporting CSVs..."

$invoiceDetails | Export-Csv -Path "$OutputPath\invoice_details.csv" -NoTypeInformation
$arAgingDetail | Export-Csv -Path "$OutputPath\ar_aging_detail.csv" -NoTypeInformation
$dailySales | Export-Csv -Path "$OutputPath\daily_sales.csv" -NoTypeInformation
$topCustomers | Export-Csv -Path "$OutputPath\top_customers.csv" -NoTypeInformation
$salespersonPerf | Export-Csv -Path "$OutputPath\salesperson_performance.csv" -NoTypeInformation
$arAgingSummary | Export-Csv -Path "$OutputPath\ar_aging_summary.csv" -NoTypeInformation
$geography | Export-Csv -Path "$OutputPath\geography.csv" -NoTypeInformation

# Create team dimension table
$teamDim = $teamRoles.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{
        Username = $_.Key
        FullName = $_.Value.FullName
        Role = $_.Value.Role
        TeamType = $_.Value.Type
    }
}
$teamDim | Export-Csv -Path "$OutputPath\team_dimension.csv" -NoTypeInformation

Write-Host "`n=== EXPORT COMPLETE ==="
Write-Host "Invoice Details: $($invoiceDetails.Count) records"
Write-Host "AR Aging Detail: $($arAgingDetail.Count) records"
Write-Host "Daily Sales: $($dailySales.Count) records"
Write-Host "Top Customers: $($topCustomers.Count) records"
Write-Host "Salesperson: $($salespersonPerf.Count) records"
Write-Host "Geography: $($geography.Count) records"

# Summary stats
$totalSales = ($invoiceDetails | Measure-Object -Property Total -Sum).Sum
$totalAR = ($arAgingDetail | Measure-Object -Property Balance -Sum).Sum
Write-Host "`nTotal Sales (2026): `$$([math]::Round($totalSales, 2))"
Write-Host "Total AR Outstanding: `$$([math]::Round($totalAR, 2))"

Write-Host "`nFiles saved to: $OutputPath"
Write-Host "$(Get-Date) - Done!"
