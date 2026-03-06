# 📊 Excel Formula Mastery

## Preface

This repository documents a structured, progressive journey through **Microsoft Excel formula proficiency** — from foundational functions to expert-level multi-formula constructions. The content is organized into four tiers: **Basic**, **Intermediate**, **Advanced**, and **Expert**, each building directly on the skills established before it.

The goal of this collection is not merely to list formulas, but to show *how formulas combine* to solve real-world data problems: payroll calculation, lookup across sheets, dynamic filtering, proration logic, and revenue modeling. Each test section captures actual formula solutions applied to realistic datasets, making this a practical reference rather than a theoretical guide.

Whether you are preparing for an Excel assessment, onboarding to a data-heavy role, or building your own formula toolkit — this document maps the full spectrum of what Excel can do when pushed to its limits.

---

## 📁 Structure

```
excel-formula-mastery/
│
├── README.md
├── Basic/
│   ├── excel-test-basic-1/     (7 test sections)
│   └── excel-test-basic-2/     (9 test sections)
├── Intermediate/
│   └── excel-intermediate/     (4 test sections)
├── Advanced/
│   └── excel-advanced/         (4 test sections)
└── Expert/
    └── excel-expert/           (1 comprehensive test)
```

---

## 🟡 Excel Test — Basic 1

### Section 1 — Core Text & Lookup Foundations
| Formula | Purpose |
|---|---|
| `=LEFT(C5,3)` | Extract first 3 characters from a cell |
| `=CONCAT(D5,G5)` | Concatenate two values |
| `=XLOOKUP(C5,'Data Staff'!$C$5:$C$24,'Data Staff'!$D$5:$D$24)` | Cross-sheet employee lookup |
| `=RIGHT(F5,3)` | Extract last 3 characters |
| `=IF(J5>=1250000,(J5*0.3%),0)` | Conditional bonus/tax calculation |

---

### Section 2 — SWITCH & Nested INDEX-MATCH
| Formula | Purpose |
|---|---|
| `=SWITCH(RIGHT(C5,1),"S","Single","M","Married")` | Decode marital status from ID suffix |
| `=IF(E5="Baru",INDEX(...),IF(E5="Junior",...,IF(E5="Senior",...)))` | Three-tier salary lookup by employee level using nested IF + INDEX/MATCH |

---

### Section 3 — Conditional Category Mapping
| Formula | Purpose |
|---|---|
| `=IF(LEFT(A5,1)="a","Beras",IF(LEFT(A5,1)="b","Gula"))` | Map product code prefix to category |
| `=SWITCH(RIGHT(A5,2),"yn","Yuni","rd","Rudi","mn","Muna","yd","Yudi")` | Decode name suffix to full name |

---

### Section 4 — Aggregation, Text Functions & Grade Logic
| Formula | Purpose |
|---|---|
| `=D4*C4` | Unit × price calculation |
| `=SUM(E4:E8)` / `=AVERAGE(E4:E8)` / `=MAX(E4:E8)` / `=MIN(E4:E8)` | Range aggregation |
| `=LEFT(B4,2)` / `=MID(B4,3,6)` / `=RIGHT(B4,1)` | String slicing (prefix, middle, suffix) |
| `=IF(G16>300,"A",IF(G16>=280,"B",IF(G16<280,"C")))` | Grade assignment logic |
| `=VLOOKUP(A3,$A$17:$D$20,2,FALSE)` | Exact vertical lookup |
| `=HLOOKUP(F3,$G$16:$J$19,2,FALSE)` | Horizontal lookup |
| `=COUNTIF(B3:B7,"LULUS")` | Count passing records |
| `=SUMIF($G$3:$G$12,1,$H$3:$H$12)` | Conditional sum by category |
| `UPPER` / `PROPER` / `LOWER` | Text case normalization |

---

### Section 5 — XLOOKUP with LEFT + Conditional SUMIF
| Formula | Purpose |
|---|---|
| `=XLOOKUP(LEFT(C7,3),$B$27:$B$30,$C$27:$C$30)` | Lookup using first 3 characters as key |
| `=SUM(L7:L16)` | Total revenue |
| `=SUMIF(H7:H16,"VIP",L7:L16)` | Revenue from VIP tier only |

---

### Section 6 — Multi-lookup Methods & SWITCH with TRUE
| Formula | Purpose |
|---|---|
| `=XLOOKUP(C5,$B$19:$B$21,$C$19:$C$21)` | Standard XLOOKUP |
| `=VLOOKUP(C5,$B$19:$D$21,3,FALSE)` | Column-indexed VLOOKUP |
| `=HLOOKUP(C5,$G$18:$I$19,2)*G5` | HLOOKUP multiplied by quantity |
| `=IF(B5<DATE(2023,7,10),XLOOKUP(D5,$C$25:$E$25,$C$26:$E$26),"CD Blank")` | Date-conditional lookup |
| `=SWITCH(TRUE,I5>1500000,4,I5>1000000,3,I5>500000,2,I5<=500000,1)` | Tiered salary bracket with SWITCH(TRUE) pattern |

---

### Section 7 — SWITCH + XLOOKUP + HLOOKUP Combination
| Formula | Purpose |
|---|---|
| `=SWITCH(TRUE,E5<=3,(XLOOKUP(...))*E5,E5>3,XLOOKUP(...)*4)` | Quantity-capped pricing: multiply rate × qty, cap at 4 |
| `=IF(D5="NIKAH",XLOOKUP(C5,$B$19:$B$22,$E$19:$E$22),0)` | Marriage allowance lookup |
| `=HLOOKUP(F5,$I$18:$L$19,2)` | Horizontal benefit table lookup |

---

## 🟡 Excel Test — Basic 2

### Section 1 — Marital & Employment Status with SWITCH(TRUE)
| Formula | Purpose |
|---|---|
| `=SWITCH(TRUE,F3="K",D3*1.25,F3="T",D3)` | Married staff gets 25% salary boost |
| `=SWITCH(TRUE,AND(F3="T",E3>10),H3*10%,,,0)` | Overtime bonus for tenured non-married staff |

---

### Section 2 — Publisher Discount Tiers
| Formula | Purpose |
|---|---|
| `=XLOOKUP(LEFT(B5,2),$B$23:$B$32,$C$23:$C$32)` | Map 2-char book code to category |
| `=SWITCH(TRUE,C5="Airlangga",F5*10%,C5="Gunung Agung",F5*15%,C5="Balai Pustaka",F5*20%,,,0)` | Publisher-specific discount rates |

---

### Section 3 — Three Equivalent Lookup Approaches
| Formula | Purpose |
|---|---|
| `=IF(D6="D",XLOOKUP(...),IF(D6="S",...,IF(D6="F",...)))` | Nested IF + XLOOKUP across columns |
| `=SWITCH(D20,"D",XLOOKUP(...),"F",XLOOKUP(...),"S",XLOOKUP(...))` | Cleaner SWITCH + XLOOKUP equivalent |
| `=INDEX($E$30:$G$32,MATCH(E20,...),MATCH(D20,...))` | Two-dimensional INDEX/MATCH lookup |

---

### Section 4 — ID Parsing & Conditional Discount
| Formula | Purpose |
|---|---|
| `LEFT` | Extract department code from employee ID |
| `=SWITCH(TRUE,(MID(B5,4,1)="2"),"Menikah",(MID(B5,4,1)="1"),"Lajang")` | Parse marital status from ID character |
| `=INDEX($G$18:$J$18,MATCH(RIGHT(B5,1),$G$17:$J$17))` | Match ID suffix to benefit table header |
| `=SWITCH(TRUE,OR(F5="Pelajar",D5="SP"),H5*15%,,,0)` | Student/special category discount |

---

### Section 5 — Multi-Sheet Payroll (INDEX/MATCH + DAYS + COUNTIF + SUM)
| Formula | Purpose |
|---|---|
| `=INDEX(Jabatan!$D$5:$D$9,MATCH(...))` | Pull job title from separate sheet |
| `=DAYS($F$4,$D$4)` | Calculate number of work days in period |
| `=COUNTIF('Absen Tidak Masuk Kerja'!$C$4:$C$13,C8)` | Count absences from attendance sheet |
| `=XLOOKUP(E8,Jabatan!$C$5:$C$9,Jabatan!$F$5:$F$9)` | Lookup allowance by position |
| `=SUM(J8,L8,M8)` | Total take-home pay |

---

### Section 6 — Score-Based Hiring Classification
| Formula | Purpose |
|---|---|
| `=SWITCH(TRUE,C4>=95,"Seleksi",C4>=90,"P.Tetap",C4>=85,"P.Kontrak",C4>=80,"P.Lepas",,,"")` | Map assessment score to employment type |

---

### Section 7 — Grade Code to Salary & Benefit
| Formula | Purpose |
|---|---|
| `=SWITCH(TRUE,C7="1A",400000,...,C7="2C",650000,0)` | Six-tier grade-to-salary mapping |
| `=SWITCH(D7,"Nikah","TV","Belum","Radio")` | Marriage status to benefit item mapping |

---

### Section 8 — Multi-Criteria Lookup (Two Methods) + Data Cleaning
| Formula | Purpose |
|---|---|
| `=INDEX($G$13:$G$27,MATCH(1,($C$13:$C$27=$B$8)*($D$13:$D$27=C7),0))` | Array-based two-criteria lookup |
| `=XLOOKUP(B9&C7,$C$13:$C$27&$D$13:$D$27,$G$13:$G$27)` | Concatenated-key XLOOKUP (modern approach) |
| **Data Cleaning Pipeline** | `TRIM` → Text to Columns (space delimiter) → `PROPER` → `UPPER` → Paste as Values |

---

### Section 9 — SUMIFS, Dynamic Lookup & Nested Logic
| Formula | Purpose |
|---|---|
| `=INDEX($N$4:$P$4,MATCH(C3,$N$3:$P$3))` | Header-row lookup to retrieve column label |
| `=XLOOKUP(E3,$M$8:$M$12,$N$8:$N$12)` | Standard rate table lookup |
| `=SWITCH(TRUE,G3>10,I1*2%,G3>20,I1*5%,0)` | Volume-based commission rate |
| `=SUMIFS($K$3:$K$17,$E$3:$E$17,B23)` | Filtered sum by agent/category |

---

## 🟠 Excel Intermediate

### Section 1 — Array Formulas, Date Logic & SUMIFS
| Formula | Purpose |
|---|---|
| `=SUMIF(...,"<200",...)` | Sum orders below threshold |
| `=SUM((F3:F26-G3:G26)*H3:H26)` | Array: revenue = (sell − cost) × qty |
| `=SUM(IF(F3:F26=G3:G26,1,0))` | Count rows where sell price equals cost price |
| `=SUM(C6/(COUNTA(...)-C7))` | Average excluding a specific count |
| `=SUM(IF((H3:H26>=101)*(H3:H26<=250),1,0))` | Array count within a numeric range |
| `="Q"&ROUNDUP(MONTH(A2)/3,0)&" "&YEAR(A2)` | Derive fiscal quarter label from date |
| `=SUMIFS(Table3[Gross Order],Table3[Vendor],"Vendor I",Table3[Item],"Item A",Table3[Date],">=1/1/2020",Table3[Date],"<=1/31/2020")` | Multi-criteria date-range SUMIFS |
| `=(D7-C7)/C7` | Week-on-week growth rate |

---

### Section 2 — Data Cleaning, IFS & Text Formulas
| Formula | Purpose |
|---|---|
| `=SUBSTITUTE([@[OLD IMAGE_URL]],"thumb","detail")` | URL string replacement |
| `=IFS([@[Sale Price]]>49,0,AND(...),[value],TRUE,10.59)` | Three-tier shipping cost logic |
| `=IFS(D11=0,"Qualified",B11<15,"Qualified",C11<3,"Qualified",TRUE,"Not Qualified")` | Multi-condition qualification check |
| `=PROPER(CONCAT(A7," ",B7," ",LEFT(C7,1),"."))` | Full name formatting with initial |
| `="Q"&ROUNDUP(MONTH(R2)/3,0)&" "&YEAR(R2)` | Quarter label from date |
| `=IFS([@TL]="TL 1","OM 1",[@TL]="TL 2","OM 1",...,TRUE,"invalid")` | Team lead to operations manager mapping |

---

### Section 3 — Dynamic Arrays: UNIQUE, FILTER, IFERROR + SMALL
| Formula | Purpose |
|---|---|
| `=IFERROR(INDEX(...,SMALL(IF(Data!B$2:B$82=$B$9,ROW(...)-ROW(...)+1),ROW(1:1))),"")` | Legacy dynamic filter (pre-365): extract matching rows one-by-one |
| `=UNIQUE(Data!B2:B82)` | Distinct list of values (Excel 365) |
| `=FILTER(Data!D2:D82,Data!B2:B82=$B$9)` | Spill filtered results matching a criterion |

---

### Section 4 — Table-Aware Formulas & MoM Analysis
| Formula | Purpose |
|---|---|
| `=XLOOKUP([@MOVIE],Data!$A$2:$A$17,Data!$B$2:$B$17)` | Movie attribute lookup with structured reference |
| `=AVERAGE(Table2[@[Jul-21]:[Jan-22]])` | Row average across 7-month span |
| `=MIN/MAX(Table2[@[Jul-21]:[Jan-22]])` | Row extremes |
| `=[@[Jan-22]]/[@[Dec-21]]-1` | Month-on-month growth rate |
| `=IF([@Average]>AVERAGE([Average]),"Above Average","Below Average")` | Benchmark comparison within column |

---

## 🔴 Excel Advanced

### Section 1 — Conditional Formatting, Goal Seek & Multi-Criteria Lookup
| Formula | Purpose |
|---|---|
| Conditional Formatting | Highlight bottom 10 values in red |
| `=IF(OR($H$6>=500000,$H$7>=25000),"Accept","Reject")` | Accept if either condition is met |
| `=IF(AND($H$6>=500000,$H$7>=25000),"Accept","Reject")` | Accept only if both conditions are met |
| Break-even formulas | `Margin = Price − Cost`, `Units = (Fixed Cost + Target Profit) / Margin` |
| `=INDEX($F$10:$F$21,MATCH(1,($B$10:$B$21=$B$6)*($C$10:$C$21=C5),0))` | Array two-criteria INDEX/MATCH |
| `=XLOOKUP(1,($B$10:$B$21=$B$7)*($C$10:$C$21=C5),$F$10:$F$21,"Not Found",0)` | Two-criteria XLOOKUP with boolean array |

---

### Section 2 — Date Intelligence & Tiered Airline Commission
| Formula | Purpose |
|---|---|
| `=TEXT(B7,"ddd")` | Extract weekday abbreviation |
| `=OR(J6="Sat",J6="Sun")` | Weekend flag |
| `=OR(D6="CGK_SUB",D6="CGK_PKU",...)` | Route membership check |
| `=AND(B6>=DATE(2015,10,1),B6<=DATE(2015,10,15))` | Date range membership |
| `=IFS(E6="Airlines 1",IF(K6,F6*0.04,0),...,E6="Airlines 5",IF(F6<50000000,F6*0.01,...))` | Per-airline multi-tier commission calculation |
| `=SUMIFS($F$6:$F$647,$E$6:$E$647,"Airlines 2",$B$6:$B$647,B6)` | Total airline 2 revenue per date |
| `=SWITCH(I6,"Method 1",0,"Method 2",H6*0.03,"Method 3",4000,"Method 4",3000,0)` | Payment method fee schedule |

---

### Section 3 — Financial Modeling, LARGE/SMALL & Proration
| Formula | Purpose |
|---|---|
| `=IF(E14<=$C$11,D15*(1+$C$6),"")` | Compound growth projection up to year limit |
| `=IF(D$14<=$C$11,D$15*$C$7,"")` | Tax calculation within projection range |
| `=LARGE($F$6:$F$17,$H7)` / `=SMALL($F$6:$F$17,H15)` | Rank-based value extraction |
| `=INDEX($B$6:$B$17,MATCH(I7,$F$6:$F$17,0))` | Reverse lookup: value → name |
| `=XLOOKUP(I15,$F$6:$F$17,$B$6:$B$17,,0)` | XLOOKUP equivalent for reverse lookup |
| **Proration — Full Month** | `=IF($D6<F$5,0,IF($D6>EOMONTH(F$5,0),$C6,$C6*(DAY($D6)/DAY(EOMONTH(F$5,0)))))` |
| **Proration — Partial Month** | `=IF($D6<F$5,0,IF($D6<=EOMONTH(F$5,0),$C6*(DAY($D6)-DAY(F$5)+1)/DAY(EOMONTH(F$5,0)),$C6))` |

---

### Section 4 — Large Dataset Lookup, Tax Classification & SUMIFS
| Formula | Purpose |
|---|---|
| `=XLOOKUP(Test!B2,DTBS[CUSTOMER CODE],DTBS[CUSTOMER NAME],,,-1)` | Last-match XLOOKUP (handles duplicates) |
| `=LOOKUP(2,1/(DTBS[CUSTOMER CODE]=Test!B2),Database!$C$2:$C$54448)` | Classic last-match LOOKUP trick |
| `=TEXTJOIN(", ",TRUE,UNIQUE(FILTER(DTBS[Tax],DTBS[CUSTOMER CODE]=B2)))` | All unique tax types for a customer |
| `=IF(AND(COUNTIFS(...,"PKP")>0,COUNTIFS(...,"PTKP")>0),"Both",TEXTJOIN(...))` | Detect mixed PKP/PTKP tax status |
| `=IF(D2="PKP",(G2+H2)*0.1,IF(D2="PTKP",0,IF(D2="Both",(G2+H2)*(COUNTIFS(...,"PKP")/COUNTIF(...))*0.1,0)))` | Proportional VAT for mixed-status customers |
| `=SUMIFS(DTBS[BOX QTY],DTBS[CUSTOMER NAME],Test!C2)` | Total box quantity by customer name |

---

## 🔵 Excel Expert

### Section 1 — HR Recruitment Revenue Modeling
| Formula | Purpose |
|---|---|
| `="+"&XLOOKUP($O4,$E$2:$E$3002,$H$2:$H$3002,"Not Found")` | Format phone number with + prefix |
| `=LOWER(LEFT(XLOOKUP(...),4))&"***@"&IF(RAND()>0.5,"jossmail.com","yuhuumail.com")` | Generate anonymized email from name |
| `=DATEDIF(XLOOKUP($O4,$E$2:$E$3002,$F$2:$F$3002),$R4,"Y")` | Age calculation via DATEDIF + XLOOKUP |
| `=IF(DATEDIF($R4,$S4,"Y")>=3,0.025,IF(...)>=1,0.015,0))+IF($V4>=50,0.04,IF(...))` | Composite bonus rate: tenure + performance |
| `=COUNTIF($R:$R,"<"&DATE(2021,3,15))` | Count candidates hired before a date |
| `=Z11*VLOOKUP("To be a Sales & Marketing Associate",$J$6:$L$9,3,FALSE)` | Revenue per placement × closing rate |
| `=SUMPRODUCT(($P$3:$P$1500="To be a Sales & Marketing Associate")*($W$3:$W$1500)*VLOOKUP(...))` | Total expected revenue for a role |
| `=SUMPRODUCT(($P="Sales Assoc")*((channel="LinkedIn")*$L$3+(channel="Jobseeker")*$L$4)*VLOOKUP(...)*(1-$W))` | Channel-weighted net revenue after churn |
| `=SUMPRODUCT(($P="Sales Assoc")*IFERROR(INDEX($L$11:$L$16,MATCH(LEFT($T,5)&"*",$J$11:$J$16,0)),0))` | Wildcard MATCH inside SUMPRODUCT for tiered SMS charges |
| `=COUNTIFS($Q,"LinkedIn",$P,"Sales Assoc",$S,">=5/1/2021",$S,"<6/1/2021")` | Monthly source-filtered applicant count |
| `=SUMPRODUCT(($Q="LinkedIn")*($P="Sales Assoc")*($S>=DATE(2021,5,1))*($S<DATE(2021,6,1))*VLOOKUP(...)*(1-$W))` | Monthly channel revenue with date filter |
| `=UNIQUE(SORT(FILTER($V$4:$V$1500,$V$4:$V$1500<>"")))` | Distinct sorted performance score list |
| `=SUMPRODUCT(($V>=50)*(XLOOKUP($O,$E$3:$E$3002,$G$3:$G$3002,"")="Male"))` | Count high-scoring male candidates |

---

## 💡 Key Concepts Covered

| Concept | Functions Used |
|---|---|
| **Text Manipulation** | `LEFT`, `RIGHT`, `MID`, `CONCAT`, `SUBSTITUTE`, `TRIM`, `UPPER`, `PROPER`, `LOWER`, `TEXTJOIN` |
| **Lookup & Reference** | `XLOOKUP`, `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `LOOKUP` |
| **Conditional Logic** | `IF`, `IFS`, `SWITCH`, `AND`, `OR` |
| **Aggregation** | `SUM`, `SUMIF`, `SUMIFS`, `SUMPRODUCT`, `AVERAGE`, `COUNT`, `COUNTIF`, `COUNTIFS` |
| **Date & Time** | `DATE`, `DATEDIF`, `DAYS`, `EOMONTH`, `MONTH`, `YEAR`, `ROUNDUP`, `TEXT` |
| **Dynamic Arrays** | `FILTER`, `UNIQUE`, `SORT`, `SMALL`, `LARGE` |
| **Error Handling** | `IFERROR` |
| **Data Cleaning** | `TRIM`, Text to Columns, `PROPER`, `UPPER`, Paste as Values |
| **Financial Modeling** | Proration, Break-even, MoM Growth, Bonus Tiers, VAT Calculation |

---

> 💬 *Formulas are drawn from real assessment test scenarios. Cell references (e.g. `$C$5`, `Table3[Column]`) reflect actual workbook structure and may need adjustment for your own datasets.*
