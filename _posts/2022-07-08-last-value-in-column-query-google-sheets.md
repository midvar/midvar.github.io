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
date: 2022-07-09 09:01:23-5:00
---

- Get last value in a columns row by using `COUNT` (or `COUNTA`) to get row number, then plugging that into `INDIRECT`.  
- Do calculations on several columns and get just the results row via `INDEX`.

---

### **Example**

**Input**

> Sheetname `Prices`

|     |A|B|C|D|E|F|G|H|I|J|
| --- | ---------- | ----- | ---- | ----- | ----- | -------- | ----- | ------- | ------ | -------- |
|  1  | Date       | Low   | Open | Close | High  | Volume   | SMA20 | SMA50   | SMA120 | SMA200   |
|  2  | 07/01/2022 | 51.37 | 53   | 57.46 | 57.74 | 13950384 | 64.68 | 63.6046 | 46.133 | 34.76575 |

**Formula**

~~~ excel
=QUERY(INDIRECT("Prices!A"&COUNT(Prices!C2:C)+1&":J"),"select A,B,D-C,(D/C)-1,G,H,I,J label A 'Date',B 'Close',D-C '$Change',(D/C)-1 '%Change',G 'SMA20',H 'SMA50',I 'SMA120',J 'SMA200'",0)
~~~

**Output**

> Sheetname `Data`

|     |A|B|C|D|E|F|G|H|
| --- | ---------- | ----- | ------- | ------------ | ----- | ------- | ------ | -------- |
|  1  | Date       | Close | $Change | %Change      | SMA20 | SMA50   | SMA120 | SMA200   |
|  2  | 07/01/2022 | 51.37 | 4.46    | 0.0841509434 | 64.68 | 63.6046 | 46.133 | 34.76575 |

---

### ***Bonus***

{% capture notice-text %}
[Stock_Price_History Sheet with QUERY](https://docs.google.com/spreadsheets/d/1VjLE5rUbDxR17kJVcskBnCeOeFmra4xk_s3TErnBuf4/edit?usp=sharing)

[Try my guide to building a UI for a Stock Price History with Appsheets!](#url-here){: .btn .btn--large .btn--info}
{% endcapture %}

<div class="notice--info text-center">
  <h4 class="no_toc">Sample Worksheet. `QUERY` formula is on Data sheet in cell `A1`</h4>

  {{ notice-text | markdownify }}

</div>