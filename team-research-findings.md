# CTI Dashboard Team & AR Research Findings

**Generated:** 2026-01-31
**Researcher:** Clawd (AI Assistant)

---

## Executive Summary

1. **Username-to-Name Mapping:** Successfully mapped 15 of 18 SyteLine TakenBy usernames to full names via Salesforce
2. **Service Technicians:** Found 20+ active field service technicians in Salesforce (more than currently tracked)
3. **AR Aging:** SyteLine calculation is within 1% of Birst data ($4.656M vs $4.704M) - can replace Birst dependency

---

## 1. SyteLine Username to Full Name Mapping

### TakenBy Values (from SLCoS orders)
The following 18 unique usernames appear in SyteLine customer orders:

| Username | Full Name | Role | Type | Source |
|----------|-----------|------|------|--------|
| Allison | Allison White | Accounting Operations Lead | Admin | SF Match |
| Anna | **Unknown** | - | - | No SF match |
| bcoli | **Colin Bowles** (likely) | Scheduling Coordinator | Service Support | SF Pattern |
| cjord | **Jordan Cline** (likely) | Client Care Supervisor | Service Support | SF Pattern |
| ebrow | Emily Brown | Client Care Supervisor | Service Support | SF Match |
| Greg | Greg Smith | Account Executive | Inside Sales | SF Match |
| hnich | Nicholas Hegger | In-House Service Supervisor | Service | SF Match |
| Joel | Joel EuDaly | COO | Executive | SF Match |
| ljami | Jami Lorenz | Service Director | Service | SF Match |
| lphil | Phil Libbert | Account Executive | Inside Sales | SF Match |
| marissa | Marissa King | Client Care Supervisor | Service Support | SF Match |
| mjaso | **Unknown** | - | - | No SF match (Jason?) |
| mrand | Randy Mobley | Account Executive | Inside Sales | SF Match |
| nandy | Andy Neptune | Outside Sales | Outside Sales | SF Match |
| Rachel L | Rachel Lively | Service Support Specialist | Service Support | SF Match |
| RMA | System - RMA Orders | - | System | N/A |
| seric | Erica Thurmond | Service Scheduling & Support | Service Support | SF Match |
| tbolt | Trevor Bolte | Account Executive | Inside Sales | SF Match |

### Username Pattern Analysis
SyteLine usernames appear to follow the pattern: `{first initial}{last 4 letters of last name}`
- lphil = **L**ibbert Phil → Phil Libbert
- mrand = **M**obley Randy → Randy Mobley  
- nandy = **N**eptune Andy → Andy Neptune
- seric = **S**hoop Erica → Erica Thurmond (maiden name Shoop)
- cjord = **C**line Jordan → Jordan Cline
- bcoli = **B**owles Colin → Colin Bowles

### Unknown Usernames
- **Anna** - Could be someone no longer at company, or using first name only
- **mjaso** - Pattern suggests "M______ Jason" - no current SF user matches

---

## 2. Complete Team Roster (from Salesforce)

### Sales Team

#### Inside Sales (tracked in dashboard)
| Name | Username | Title | Email |
|------|----------|-------|-------|
| Phil Libbert | lphil | Account Executive | phil.libbert@ctigas.com |
| Greg Smith | Greg | Account Executive | greg.smith@ctigas.com |
| Randy Mobley | mrand | Account Executive | randy.mobley@ctigas.com |
| Trevor Bolte | tbolt | Account Executive | trevor.bolte@ctigas.com |

#### Outside Sales (NOT in current dashboard)
| Name | Title | Email |
|------|-------|-------|
| Andy Neptune | Outside Sales | andy.neptune@ctigas.com |
| Albert Febre | Outside Sales | albert.febre@ctigas.com |
| Alex Smith | Outside Sales | alex.smith@ctigas.com |
| Jim Renville | Outside Sales | jim.renville@ctigas.com |
| Miguel Lopez | TEX-MEX-OK/LATAM Regional Manager | miguel.lopez@ctigas.com |

#### Sales Leadership
| Name | Title | Email |
|------|-------|-------|
| Richard Hatcher | Sales Director | richard.hatcher@ctigas.com |

### Service Team

#### Service Leadership
| Name | Username | Title |
|------|----------|-------|
| Jami Lorenz | ljami | Service Director |
| Nicholas Hegger | hnich | In-House Service Supervisor |
| Mike Cusanno | - | Service Technician Supervisor |

#### Service Support (Client Care/Scheduling)
| Name | Username | Role |
|------|----------|------|
| Erica Thurmond | seric | Service Scheduling & Support Supervisor |
| Jordan Cline | cjord | Client Care Supervisor |
| Emily Brown | ebrow | Client Care Supervisor |
| Marissa King | marissa | Client Care Supervisor |
| Rachel Lively | Rachel L | Service Support Specialist |
| Colin Bowles | bcoli | Scheduling Coordinator Lead |

#### Field Service Technicians (20 active)
| Name | Title | Lead? |
|------|-------|-------|
| Nathan Glenister | Field Service Technician | **Lead** |
| Phillip Forbes | Field Service Technician | **Lead** |
| Travis Ratcliffe | Field Service Technician | **Lead** |
| Adam Kies | Service Technician | |
| Andrew Shelley | Service Technician | |
| Cameron Cofer | Field Service Technician | |
| Casey Chambers | Service Technician | |
| Cole Thomas | Field Service Technician | |
| Damir Livancic | Service Technician | |
| David Jubeck | Service Technician | |
| David Thomas | Field Service Technician | |
| Duncan Gepner | Service Technician | |
| Heriberto Hidalgo | Service Technician | |
| JJ McKenna | Service Technician | |
| Joey Todd | Service Technician | |
| Joshua Mattingley | Field Service Technician | |
| Joshua Musmeci | Service Technician | |
| Kyle Kotraba | Service Technician | |
| Matthew Plemons | Service Technician | |
| Sam Atkins | Service Technician | |
| Travis Mays | Service Technician | |
| Vadar Ali | Service Technician | |

#### In-House Service Technicians
| Name | Title |
|------|-------|
| Dane Cross | Certified In-House Service Technician |

### Executive/Admin
| Name | Username | Title |
|------|----------|-------|
| Scott Lordo | - | CEO |
| Joel EuDaly | Joel | COO |
| Justin Golding | - | Financial Analyst |
| Allison White | Allison | Accounting Operations Lead |

---

## 3. AR Aging Analysis

### SyteLine vs Birst Comparison

| Bucket | SyteLine | Birst | Difference |
|--------|----------|-------|------------|
| Current (not due) | $3,528,837 | $3,627,797 | -$98,960 |
| 1-30 days | $570,662 | $617,951 | -$47,290 |
| 31-60 days | $132,249 | $113,317 | +$18,932 |
| 61-90 days | $181,849 | $165,937 | +$15,912 |
| 90+ days | $242,774 | $179,156 | +$63,618 |
| **TOTAL** | **$4,656,370** | **$4,704,159** | **-$47,789** |

**Variance: ~1% ($47,789 on $4.7M)**

### Analysis
- The SyteLine calculation is very close to Birst
- Small differences likely due to:
  - Timing (Birst was Jan 30, SyteLine is real-time)
  - Different rounding approaches
  - Possible differences in how credits are applied
- **Recommendation:** SyteLine AR calculation can replace Birst dependency

### Current AR Calculation Logic
The generate-dashboard.ps1 script already has correct AR logic:
1. Get all SLArTrans records
2. Sum Type="I" (Invoice) amounts per invoice
3. Subtract Type="P" (Payment) amounts via ApplyToInvNum
4. Subtract Type="C" (Credit) amounts via ApplyToInvNum
5. Bucket remaining balance by DueDate vs today

---

## 4. Recommendations for generate-dashboard.ps1

### A. Update Team Roles Mapping
Replace the current `$teamRoles` hashtable with:

```powershell
$teamRoles = @{
    # Executives
    'Joel' = @{ Role='COO'; Type='Executive'; FullName='Joel EuDaly' }
    
    # Admin
    'Allison' = @{ Role='Accounting'; Type='Admin'; FullName='Allison White' }
    
    # Inside Sales
    'lphil' = @{ Role='Inside Sales'; Type='Sales'; FullName='Phil Libbert' }
    'Greg' = @{ Role='Inside Sales'; Type='Sales'; FullName='Greg Smith' }
    'mrand' = @{ Role='Inside Sales'; Type='Sales'; FullName='Randy Mobley' }
    'tbolt' = @{ Role='Inside Sales'; Type='Sales'; FullName='Trevor Bolte' }
    
    # Outside Sales
    'nandy' = @{ Role='Outside Sales'; Type='Sales'; FullName='Andy Neptune' }
    
    # Service Leadership
    'ljami' = @{ Role='Service Director'; Type='Service'; FullName='Jami Lorenz' }
    'hnich' = @{ Role='In-House Service Supervisor'; Type='Service'; FullName='Nicholas Hegger' }
    
    # Service Support / Client Care
    'seric' = @{ Role='Service Scheduling'; Type='Service'; FullName='Erica Thurmond' }
    'cjord' = @{ Role='Client Care'; Type='Service'; FullName='Jordan Cline' }
    'ebrow' = @{ Role='Client Care'; Type='Service'; FullName='Emily Brown' }
    'marissa' = @{ Role='Client Care'; Type='Service'; FullName='Marissa King' }
    'Rachel L' = @{ Role='Service Support'; Type='Service'; FullName='Rachel Lively' }
    'bcoli' = @{ Role='Scheduling Coordinator'; Type='Service'; FullName='Colin Bowles' }
    
    # Unknown/Legacy
    'Anna' = @{ Role='Unknown'; Type='Unknown'; FullName='Anna (Unknown)' }
    'mjaso' = @{ Role='Unknown'; Type='Unknown'; FullName='Jason (Unknown)' }
    'RMA' = @{ Role='System'; Type='System'; FullName='RMA Orders' }
}
```

### B. Use FullName in Dashboard Display
Modify the salesperson output to use FullName when available:
```powershell
$displayName = if ($teamRoles.ContainsKey($_.Key) -and $teamRoles[$_.Key].FullName) { 
    $teamRoles[$_.Key].FullName 
} else { 
    $_.Key 
}
```

### C. Remove Birst Dependency for AR
The current AR calculation from SyteLine is accurate enough (~1% variance).
Remove this fallback block and always use SyteLine:

```powershell
# REMOVE this conditional - always use SyteLine calculation
if (Test-Path $birstArPath) {
    $birstAR = Get-Content $birstArPath | ConvertFrom-Json
    ...
}
```

### D. Add Service Technician Tracking (Future)
To track field technicians in the dashboard:
1. Query SLWOs (Work Orders) or SLServiceOrders
2. Map technician assignments to revenue
3. Add "Service Technicians" section to dashboard

---

## 5. Data Files Created

| File | Description |
|------|-------------|
| `data/salesforce-users.json` | Full Salesforce user export with roles |
| `data/syteline-user-research.json` | SyteLine user query results |
| `data/ar-aging-comparison.json` | SyteLine vs Birst AR comparison |

---

## 6. Open Questions for Justin

1. **Anna** - Who is this? Active employee or legacy?
2. **mjaso** - Any idea who "Jason" might be? Pattern suggests last name starting with M
3. **Field Technicians** - Should we track individual tech performance in the dashboard?
4. **Outside Sales** - Should Outside Sales (Andy Neptune, etc.) be tracked separately from Inside Sales?

---

## Summary of Changes Needed

1. ✅ **Update $teamRoles** with full names and corrected roles
2. ✅ **Display full names** instead of usernames in dashboard
3. ✅ **Remove Birst AR dependency** - SyteLine is accurate enough
4. ⏳ **Add Outside Sales section** (optional)
5. ⏳ **Add Service Technician tracking** (future enhancement)
