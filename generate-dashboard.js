const axios = require('axios');
const fs = require('fs');
const path = require('path');

// SyteLine API Configuration
const SYTELINE_BASE = 'https://csi10g.erpsl.inforcloudsuite.com/IDORequestService/ido';
const TENANT = 'GVNDYXUFKHB5VMB6_PRD_CTI';
const USERNAME = process.env.SYTELINE_USERNAME || 'gary.phillips@godlan.com';
const PASSWORD = process.env.SYTELINE_PASSWORD || 'Crwthtithing2$';

// Service product codes
const SERVICE_CODES = ['LFTR', 'LGAS', 'LIHL'];

// Collection rates by aging bucket
const COLLECTION_RATES = {
    current: 0.70,
    days1_30: 0.25,
    days31_60: 0.15,
    days61_90: 0.08,
    days90Plus: 0.03
};

// Day of week payment weights
const DOW_WEIGHTS = {
    Monday: 0.15, Tuesday: 0.28, Wednesday: 0.25, Thursday: 0.20, Friday: 0.12
};

// Holidays 2026
const HOLIDAYS = [
    '2026-01-01', '2026-01-20', '2026-02-17', '2026-05-25',
    '2026-07-03', '2026-09-07', '2026-11-26', '2026-11-27', '2026-12-25'
];

async function getToken() {
    const url = `${SYTELINE_BASE}/token/${TENANT}/${encodeURIComponent(USERNAME)}/${encodeURIComponent(PASSWORD)}`;
    const response = await axios.get(url);
    return response.data.Token;
}

async function fetchIDO(token, ido, properties, filter = null, recordCap = 10000) {
    let url = `${SYTELINE_BASE}/load/${ido}?properties=${properties}&recordCap=${recordCap}`;
    if (filter) url += `&filter=${encodeURIComponent(filter)}`;
    const response = await axios.get(url, { headers: { Authorization: token } });
    return response.data.Items || [];
}

function parseDate(dateStr) {
    if (!dateStr || dateStr.length < 8) return null;
    const clean = dateStr.replace(/\s.*/g, '').replace(/-/g, '');
    const y = clean.substring(0, 4);
    const m = clean.substring(4, 6);
    const d = clean.substring(6, 8);
    return new Date(`${y}-${m}-${d}`);
}

function parseDateToKey(dateStr) {
    if (!dateStr || dateStr.length < 8) return null;
    const clean = dateStr.replace(/\s.*/g, '').replace(/-/g, '');
    const y = clean.substring(0, 4);
    const m = clean.substring(4, 6);
    const d = clean.substring(6, 8);
    return `${y}-${m}-${d}`;
}

function formatDate(date) {
    return date.toISOString().split('T')[0];
}

function getWeekStart(date) {
    const d = new Date(date);
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1);
    return new Date(d.setDate(diff));
}

function isWeekend(date) {
    const day = date.getDay();
    return day === 0 || day === 6;
}

function getDayName(date) {
    return ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][date.getDay()];
}

async function generateDashboard() {
    console.log(`${new Date().toISOString()} - Generating CTI Executive Dashboard...`);
    
    const token = await getToken();
    console.log('Connected to SyteLine');
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    let yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);
    while (isWeekend(yesterday)) yesterday.setDate(yesterday.getDate() - 1);
    
    const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
    const yearStart = new Date(today.getFullYear(), 0, 1);
    
    // Fetch item master for product categorization
    console.log('Fetching item master...');
    const items = await fetchIDO(token, 'SLItems', 'Item,ProductCode', null, 15000);
    const itemProductCode = {};
    items.forEach(i => { if (i.Item) itemProductCode[i.Item.trim()] = i.ProductCode || ''; });
    console.log(`Built ProductCode lookup for ${Object.keys(itemProductCode).length} items`);
    
    // Fetch line items for order breakdown
    console.log('Fetching line items...');
    const coItems = await fetchIDO(token, 'SLCoItems', 'CoNum,Item,ItDescription,DerExtInvoicedPrice', null, 15000);
    console.log(`Got ${coItems.length} line items`);
    
    const orderBreakdown = {};
    const salesByProduct = {};
    
    coItems.forEach(line => {
        if (!line.CoNum || !line.DerExtInvoicedPrice) return;
        const orderNum = line.CoNum.trim();
        const amount = Math.abs(parseFloat(line.DerExtInvoicedPrice) || 0);
        if (amount <= 0) return;
        
        const itemKey = (line.Item || '').trim();
        const prodCode = itemProductCode[itemKey] || '';
        const isService = SERVICE_CODES.includes(prodCode);
        
        if (!orderBreakdown[orderNum]) orderBreakdown[orderNum] = { Service: 0, Product: 0, Total: 0 };
        if (isService) orderBreakdown[orderNum].Service += amount;
        else orderBreakdown[orderNum].Product += amount;
        orderBreakdown[orderNum].Total += amount;
        
        if (!salesByProduct[itemKey]) salesByProduct[itemKey] = { Description: line.ItDescription || '', MTD: 0 };
        salesByProduct[itemKey].MTD += amount;
    });
    console.log(`Built breakdown for ${Object.keys(orderBreakdown).length} orders`);
    
    // Fetch AR transactions
    console.log('Fetching AR transactions...');
    const arTrans = await fetchIDO(token, 'SLArTrans', 'InvNum,CoNum,CadName,InvDate,DueDate,Amount,Type,ApplyToInvNum', null, 15000);
    console.log(`Got ${arTrans.length} AR transaction records`);
    
    const invToOrder = {};
    const invToCustomer = {};
    const invoiceBalances = {};
    const invoiceInfo = {}; // Store invoice details for drill-down
    
    // First pass: Add invoices
    arTrans.forEach(ar => {
        if (ar.Type === 'I' && ar.InvNum) {
            const invNum = ar.InvNum.trim();
            if (!invoiceBalances[invNum]) {
                invoiceBalances[invNum] = { Customer: ar.CadName, DueDate: ar.DueDate, InvDate: ar.InvDate, CoNum: ar.CoNum, OriginalAmount: 0, Applied: 0 };
            }
            invoiceBalances[invNum].OriginalAmount += parseFloat(ar.Amount) || 0;
            if (ar.CoNum) invToOrder[invNum] = ar.CoNum.trim();
            if (ar.CadName) invToCustomer[invNum] = ar.CadName;
            
            // Store invoice info for drill-down
            invoiceInfo[invNum] = {
                Customer: ar.CadName || 'Unknown',
                CoNum: ar.CoNum || '',
                InvDate: ar.InvDate,
                Amount: parseFloat(ar.Amount) || 0
            };
        }
    });
    
    // Second pass: Apply payments and credits
    arTrans.forEach(ar => {
        if (['P', 'C'].includes(ar.Type) && ar.Amount) {
            const applyTo = (ar.ApplyToInvNum || ar.InvNum || '').trim();
            if (applyTo && invoiceBalances[applyTo]) {
                invoiceBalances[applyTo].Applied += parseFloat(ar.Amount) || 0;
            }
        }
    });
    
    // Build open AR list
    const openAR = [];
    Object.entries(invoiceBalances).forEach(([invNum, inv]) => {
        const balance = inv.OriginalAmount - inv.Applied;
        if (balance > 1 && inv.DueDate) {
            const dueDate = parseDate(inv.DueDate);
            if (dueDate) {
                const daysOverdue = Math.floor((today - dueDate) / (1000 * 60 * 60 * 24));
                openAR.push({ InvNum: invNum, Customer: inv.Customer, Amount: balance, DueDate: dueDate, DaysOverdue: daysOverdue });
            }
        }
    });
    console.log(`Found ${openAR.length} open invoices`);
    
    // Fetch GL Ledger
    console.log('Fetching GL ledger data...');
    const currentYear = today.getFullYear();
    const filter = `(Acct = '401000' OR Acct = '402000' OR Acct = '495000' OR Acct = '495400') AND ControlYear >= ${currentYear}`;
    const ledgerData = await fetchIDO(token, 'SLLedgers', 'Acct,DomAmount,TransDate,Ref', filter, 15000);
    console.log(`Got ${ledgerData.length} GL ledger records`);
    
    // Process sales by date AND capture invoice details
    const salesByDate = {};
    const invoicesByDate = {}; // NEW: Track invoices per date for drill-down
    const salesByCustomer = {};
    
    ledgerData.forEach(gl => {
        if (!gl.TransDate || !gl.DomAmount) return;
        const dateKey = parseDateToKey(gl.TransDate);
        if (!dateKey) return;
        const amount = Math.abs(parseFloat(gl.DomAmount) || 0);
        const acct = (gl.Acct || '').trim();
        
        if (!salesByDate[dateKey]) salesByDate[dateKey] = { Product: 0, Service: 0, Freight: 0, Miscellaneous: 0, Total: 0 };
        if (!invoicesByDate[dateKey]) invoicesByDate[dateKey] = [];
        
        if (acct === '495400') {
            salesByDate[dateKey].Freight += amount;
            salesByDate[dateKey].Total += amount;
            invoicesByDate[dateKey].push({ InvNum: 'FREIGHT', Customer: 'Freight Revenue', Amount: amount, Category: 'Freight' });
        } else if (acct === '495000') {
            salesByDate[dateKey].Miscellaneous += amount;
            salesByDate[dateKey].Total += amount;
            invoicesByDate[dateKey].push({ InvNum: 'MISC', Customer: 'Miscellaneous', Amount: amount, Category: 'Miscellaneous' });
        } else if (['401000', '402000'].includes(acct) && gl.Ref) {
            const match = gl.Ref.match(/ARI\s+(\d+)/);
            if (match) {
                const invNum = match[1].trim();
                const orderNum = invToOrder[invNum];
                const custName = invToCustomer[invNum] || 'Unknown';
                
                let category = 'Product';
                let svcAmount = 0, prodAmount = 0;
                
                if (orderNum && orderBreakdown[orderNum] && orderBreakdown[orderNum].Total > 0) {
                    const breakdown = orderBreakdown[orderNum];
                    const svcRatio = breakdown.Service / breakdown.Total;
                    const prodRatio = breakdown.Product / breakdown.Total;
                    
                    svcAmount = amount * svcRatio;
                    prodAmount = amount * prodRatio;
                    
                    salesByDate[dateKey].Service += svcAmount;
                    salesByDate[dateKey].Product += prodAmount;
                    salesByDate[dateKey].Total += amount;
                    
                    category = svcRatio > 0.5 ? 'Service' : 'Product';
                    
                    if (custName) {
                        if (!salesByCustomer[custName]) salesByCustomer[custName] = { MTD: 0, Yesterday: 0 };
                        const salesDate = parseDate(gl.TransDate);
                        if (salesDate && salesDate >= monthStart && salesDate <= today) salesByCustomer[custName].MTD += amount;
                        if (salesDate && formatDate(salesDate) === formatDate(yesterday)) salesByCustomer[custName].Yesterday += amount;
                    }
                } else {
                    prodAmount = amount;
                    salesByDate[dateKey].Product += amount;
                    salesByDate[dateKey].Total += amount;
                }
                
                // Add invoice detail for drill-down
                invoicesByDate[dateKey].push({
                    InvNum: invNum,
                    Customer: custName,
                    Amount: Math.round(amount * 100) / 100,
                    Category: category,
                    Product: Math.round(prodAmount * 100) / 100,
                    Service: Math.round(svcAmount * 100) / 100,
                    CoNum: orderNum || ''
                });
            }
        }
    });
    
    // Calculate summaries
    const yesterdayKey = formatDate(yesterday);
    const yesterdaySales = salesByDate[yesterdayKey] || { Product: 0, Service: 0, Freight: 0, Miscellaneous: 0, Total: 0 };
    const yesterdayInvoices = invoicesByDate[yesterdayKey] || [];
    
    const mtdSales = { Product: 0, Service: 0, Freight: 0, Miscellaneous: 0, Total: 0 };
    const ytdSales = { Product: 0, Service: 0, Freight: 0, Miscellaneous: 0, Total: 0 };
    const mtdInvoices = [];
    const ytdInvoices = [];
    
    Object.entries(salesByDate).forEach(([dateKey, data]) => {
        const date = new Date(dateKey);
        if (date >= monthStart && date <= today) {
            ['Product', 'Service', 'Freight', 'Miscellaneous', 'Total'].forEach(cat => mtdSales[cat] += data[cat]);
            if (invoicesByDate[dateKey]) {
                invoicesByDate[dateKey].forEach(inv => mtdInvoices.push({ ...inv, Date: dateKey }));
            }
        }
        if (date >= yearStart && date <= today) {
            ['Product', 'Service', 'Freight', 'Miscellaneous', 'Total'].forEach(cat => ytdSales[cat] += data[cat]);
            if (invoicesByDate[dateKey]) {
                invoicesByDate[dateKey].forEach(inv => ytdInvoices.push({ ...inv, Date: dateKey }));
            }
        }
    });
    
    // Sort invoices by amount descending for drill-down display
    const sortByAmount = (a, b) => b.Amount - a.Amount;
    mtdInvoices.sort(sortByAmount);
    ytdInvoices.sort(sortByAmount);
    yesterdayInvoices.sort(sortByAmount);
    
    // Calculate business days
    let mtdDays = 0, ytdDays = 0;
    for (let d = new Date(monthStart); d <= today; d.setDate(d.getDate() + 1)) {
        if (!isWeekend(d)) mtdDays++;
    }
    for (let d = new Date(yearStart); d <= today; d.setDate(d.getDate() + 1)) {
        if (!isWeekend(d)) ytdDays++;
    }
    
    const mtdAvg = mtdDays > 0 ? mtdSales.Total / mtdDays : 0;
    const ytdAvg = ytdDays > 0 ? ytdSales.Total / ytdDays : 0;
    
    // AR Aging with invoice details
    const arAging = { Current: 0, Days1_30: 0, Days31_60: 0, Days61_90: 0, Days90Plus: 0, Total: 0 };
    const arInvoicesByBucket = { Current: [], Days1_30: [], Days31_60: [], Days61_90: [], Days90Plus: [] };
    
    openAR.forEach(inv => {
        const days = inv.DaysOverdue;
        const invDetail = { InvNum: inv.InvNum, Customer: inv.Customer, Amount: Math.round(inv.Amount * 100) / 100, DaysOverdue: days };
        
        if (days <= 0) {
            arAging.Current += inv.Amount;
            arInvoicesByBucket.Current.push(invDetail);
        } else if (days <= 30) {
            arAging.Days1_30 += inv.Amount;
            arInvoicesByBucket.Days1_30.push(invDetail);
        } else if (days <= 60) {
            arAging.Days31_60 += inv.Amount;
            arInvoicesByBucket.Days31_60.push(invDetail);
        } else if (days <= 90) {
            arAging.Days61_90 += inv.Amount;
            arInvoicesByBucket.Days61_90.push(invDetail);
        } else {
            arAging.Days90Plus += inv.Amount;
            arInvoicesByBucket.Days90Plus.push(invDetail);
        }
        arAging.Total += inv.Amount;
    });
    
    // Sort AR invoices by amount
    Object.values(arInvoicesByBucket).forEach(arr => arr.sort(sortByAmount));
    
    console.log(`AR Total: $${Math.round(arAging.Total)}`);
    
    // Smart Cash Flow Prediction
    console.log('Running smart cash flow prediction...');
    const thisWeekStart = getWeekStart(today);
    const cashFlowPrediction = [];
    
    const rollingAR = { ...arAging };
    
    for (let w = 0; w < 6; w++) {
        const weekStart = new Date(thisWeekStart);
        weekStart.setDate(weekStart.getDate() + w * 7);
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekEnd.getDate() + 4);
        const weekLabel = `Week of ${weekStart.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}`;
        
        const expectedFromCurrent = rollingAR.Current * COLLECTION_RATES.current;
        const expectedFrom1_30 = rollingAR.Days1_30 * COLLECTION_RATES.days1_30;
        const expectedFrom31_60 = rollingAR.Days31_60 * COLLECTION_RATES.days31_60;
        const expectedFrom61_90 = rollingAR.Days61_90 * COLLECTION_RATES.days61_90;
        const expectedFrom90Plus = rollingAR.Days90Plus * COLLECTION_RATES.days90Plus;
        const weekExpected = expectedFromCurrent + expectedFrom1_30 + expectedFrom31_60 + expectedFrom61_90 + expectedFrom90Plus;
        
        let actual = 0, predicted = 0;
        const dayBreakdown = [];
        
        for (let d = 0; d < 5; d++) {
            const day = new Date(weekStart);
            day.setDate(day.getDate() + d);
            const dayKey = formatDate(day);
            const dow = getDayName(day);
            
            if (day < today) {
                const dayAmount = salesByDate[dayKey]?.Total || 0;
                actual += dayAmount;
                dayBreakdown.push({ Date: dayKey, DayOfWeek: dow, Type: 'Actual', Amount: Math.round(dayAmount * 100) / 100 });
            } else if (formatDate(day) === formatDate(today)) {
                const dayAmount = salesByDate[dayKey]?.Total || 0;
                actual += dayAmount;
                dayBreakdown.push({ Date: dayKey, DayOfWeek: dow, Type: 'Today', Amount: Math.round(dayAmount * 100) / 100 });
            } else {
                const dowWeight = DOW_WEIGHTS[dow] || 0.20;
                const isHoliday = HOLIDAYS.includes(dayKey);
                const holidayFactor = isHoliday ? 0.2 : 1.0;
                const dayPrediction = weekExpected * dowWeight * holidayFactor;
                predicted += dayPrediction;
                dayBreakdown.push({ Date: dayKey, DayOfWeek: dow, Type: 'Predicted', Amount: Math.round(dayPrediction * 100) / 100 });
            }
        }
        
        rollingAR.Days90Plus = Math.max(0, rollingAR.Days90Plus - expectedFrom90Plus + rollingAR.Days61_90 * 0.3);
        rollingAR.Days61_90 = Math.max(0, rollingAR.Days61_90 * 0.7 - expectedFrom61_90 + rollingAR.Days31_60 * 0.3);
        rollingAR.Days31_60 = Math.max(0, rollingAR.Days31_60 * 0.7 - expectedFrom31_60 + rollingAR.Days1_30 * 0.3);
        rollingAR.Days1_30 = Math.max(0, rollingAR.Days1_30 * 0.7 - expectedFrom1_30 + rollingAR.Current * 0.3);
        rollingAR.Current = Math.max(0, rollingAR.Current * 0.7 - expectedFromCurrent);
        
        const confidence = w === 0 ? 85 : w === 1 ? 70 : w < 4 ? 55 : 40;
        
        cashFlowPrediction.push({
            Week: weekLabel, WeekNum: w + 1,
            StartDate: formatDate(weekStart), EndDate: formatDate(weekEnd),
            Actual: Math.round(actual * 100) / 100,
            Predicted: Math.round(predicted * 100) / 100,
            Total: Math.round((actual + predicted) * 100) / 100,
            ExpectedCollections: Math.round(weekExpected * 100) / 100,
            Confidence: confidence,
            DayBreakdown: dayBreakdown
        });
    }
    
    // Build daily trend
    const dailyTrend = [];
    for (let i = 60; i >= 0; i--) {
        const date = new Date(today);
        date.setDate(date.getDate() - i);
        if (isWeekend(date)) continue;
        const dateKey = formatDate(date);
        const data = salesByDate[dateKey] || { Product: 0, Service: 0, Freight: 0, Miscellaneous: 0, Total: 0 };
        dailyTrend.push({
            Date: dateKey, DayOfWeek: getDayName(date),
            Product: Math.round(data.Product * 100) / 100,
            Service: Math.round(data.Service * 100) / 100,
            Freight: Math.round(data.Freight * 100) / 100,
            Miscellaneous: Math.round(data.Miscellaneous * 100) / 100,
            Total: Math.round(data.Total * 100) / 100
        });
    }
    
    // Top customers
    const topCustomersMTD = Object.entries(salesByCustomer)
        .filter(([_, v]) => v.MTD > 0)
        .sort((a, b) => b[1].MTD - a[1].MTD)
        .slice(0, 15)
        .map(([name, data]) => ({ Name: name, Amount: Math.round(data.MTD * 100) / 100 }));
    
    const topCustomersYesterday = Object.entries(salesByCustomer)
        .filter(([_, v]) => v.Yesterday > 0)
        .sort((a, b) => b[1].Yesterday - a[1].Yesterday)
        .slice(0, 15)
        .map(([name, data]) => ({ Name: name, Amount: Math.round(data.Yesterday * 100) / 100 }));
    
    // Top products
    const topProductsMTD = Object.entries(salesByProduct)
        .sort((a, b) => b[1].MTD - a[1].MTD)
        .slice(0, 15)
        .map(([item, data]) => ({ Item: item, Description: data.Description, Amount: Math.round(data.MTD * 100) / 100 }));
    
    // Build dashboard data WITH invoice details for drill-down
    const dashboardData = {
        GeneratedAt: new Date().toISOString().replace('T', ' ').substring(0, 19),
        DataSource: 'SyteLine IDO API (Automated)',
        Period: {
            Yesterday: formatDate(yesterday),
            MonthStart: formatDate(monthStart),
            YearStart: formatDate(yearStart),
            MTDBusinessDays: mtdDays,
            YTDBusinessDays: ytdDays
        },
        Summary: {
            Yesterday: { 
                Product: Math.round(yesterdaySales.Product * 100) / 100, 
                Service: Math.round(yesterdaySales.Service * 100) / 100, 
                Freight: Math.round(yesterdaySales.Freight * 100) / 100, 
                Miscellaneous: Math.round(yesterdaySales.Miscellaneous * 100) / 100, 
                Total: Math.round(yesterdaySales.Total * 100) / 100,
                Invoices: yesterdayInvoices.slice(0, 100) // Top 100 for drill-down
            },
            MTD: { 
                Product: Math.round(mtdSales.Product * 100) / 100, 
                Service: Math.round(mtdSales.Service * 100) / 100, 
                Freight: Math.round(mtdSales.Freight * 100) / 100, 
                Miscellaneous: Math.round(mtdSales.Miscellaneous * 100) / 100, 
                Total: Math.round(mtdSales.Total * 100) / 100,
                Invoices: mtdInvoices.slice(0, 200) // Top 200 for drill-down
            },
            YTD: { 
                Product: Math.round(ytdSales.Product * 100) / 100, 
                Service: Math.round(ytdSales.Service * 100) / 100, 
                Freight: Math.round(ytdSales.Freight * 100) / 100, 
                Miscellaneous: Math.round(ytdSales.Miscellaneous * 100) / 100, 
                Total: Math.round(ytdSales.Total * 100) / 100,
                Invoices: ytdInvoices.slice(0, 500) // Top 500 for drill-down
            },
            MTDDailyAverage: Math.round(mtdAvg * 100) / 100,
            YTDDailyAverage: Math.round(ytdAvg * 100) / 100
        },
        ARaging: arAging,
        ARInvoices: arInvoicesByBucket, // NEW: AR invoices by bucket for drill-down
        CashFlowPrediction: cashFlowPrediction,
        TopCustomers: { Yesterday: topCustomersYesterday, MTD: topCustomersMTD },
        TopProducts: { MTD: topProductsMTD },
        DailyTrend: dailyTrend
    };
    
    // Save JSON
    fs.writeFileSync(path.join(__dirname, 'data', 'dashboard-data.json'), JSON.stringify(dashboardData, null, 2));
    
    // Load HTML template and embed data
    const templatePath = path.join(__dirname, 'cti-dashboard.html');
    if (fs.existsSync(templatePath)) {
        let html = fs.readFileSync(templatePath, 'utf8');
        html = html.replace('let DATA = null;', `let DATA = ${JSON.stringify(dashboardData)};`);
        fs.writeFileSync(path.join(__dirname, 'index.html'), html);
        console.log('Dashboard saved: index.html');
    }
    
    console.log('\n=== SUMMARY ===');
    console.log(`Yesterday: $${Math.round(yesterdaySales.Total)} (${yesterdayInvoices.length} invoices)`);
    console.log(`MTD: $${Math.round(mtdSales.Total)} (${mtdInvoices.length} invoices)`);
    console.log(`YTD: $${Math.round(ytdSales.Total)} (${ytdInvoices.length} invoices)`);
    console.log(`AR Total: $${Math.round(arAging.Total)} (${openAR.length} open invoices)`);
    console.log('6-Week Cash Flow Forecast:');
    cashFlowPrediction.forEach(w => console.log(`  ${w.Week}: $${Math.round(w.Total)} (${w.Confidence}% confidence)`));
    
    console.log(`\n${new Date().toISOString()} - Dashboard generation complete!`);
}

if (!fs.existsSync(path.join(__dirname, 'data'))) {
    fs.mkdirSync(path.join(__dirname, 'data'));
}

generateDashboard().catch(err => {
    console.error('Error generating dashboard:', err);
    process.exit(1);
});
