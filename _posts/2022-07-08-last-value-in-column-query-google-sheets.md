---
title: "Google Sheets QUERY Return Values in Last Non Empty Row"
excerpt: "Get Last Value in Column via Query with Google Sheets"
categories:
  - Code Snippets
  - Google Sheets
tags:
  - QUERY
  - INDIRECT
  - GOOGLEFINANCE
date: 2022-07-08 09:01:23
---

##  **Google Sheets | `QUERY` Return Values in Last Non Empty Row**

- Get last value in a columns row by using `COUNT` (or `COUNTA`) to get row number, then plugging that into `INDIRECT`.  
- Do calculations on several columns and get just the results row via `INDEX`.

```excel
=QUERY(INDIRECT("Prices!C"&COUNT(Prices!C2:C)+1&":J"),"select D-C,(D/C)-1,H,I,J",0)
=QUERY(INDIRECT("Prices!A"&COUNT(Prices!C2:C)+1&":J"),"select A,B,D-C,(D/C)-1,G,H,I,J label A 'Date',B 'Close',D-C '$Change',(D/C)-1 '%Change',G 'SMA20',H 'SMA50',I 'SMA120',J 'SMA200'",0)
```

### **Example**  
`Prices` Sheet, Cell A1:
 
 ```excel
={ArrayFormula(text({"Date";int(query(query(googlefinance(Variables!$A$2  , "all" , Variables!$C$2 ,IF(Variables!$D$2<>"",Variables!$D$2,Today()), Variables!$B$2  ) ,"Select Col1",1),"offset 1",0))},"MM/DD/YYYY")),query(googlefinance( Variables!$A$2 , "all" , Variables!$C$2 , IF(Variables!$D$2<>"",Variables!$D$2,Today()), Variables!$B$2 ) ,"Select Col4,Col2,Col5,Col3,Col6",1)}
   ```

### **Data**

| Date       | Low   | Open | Close | High  | Volume   | SMA20  | SMA50    | SMA120  | SMA200    |
| ---------- | ----- | ---- | ----- | ----- | -------- | ----- | ------- | ------ | -------- |
| 07/01/2022 | 51.37 | 53   | 57.46 | 57.74 | 13950384 | 64.68 | 63.6046 | 46.133 | 34.76575 |

### **Output**

| Date       | Close | $Change | %Change      | SMA20 | SMA50   | SMA120 | SMA200   |
| ---------- | ----- | ------- | ------------ | ----- | ------- | ------ | -------- |
| 07/01/2022 | 51.37 | 4.46    | 0.0841509434 | 64.68 | 63.6046 | 46.133 | 34.76575 |