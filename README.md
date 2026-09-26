# EX8: Merging and Retrieving Data Using Lookup Functions

## Aim
To merge data from multiple tables and retrieve information efficiently using Excel lookup functions including **VLOOKUP, HLOOKUP, XLOOKUP, INDEX, and MATCH**.

---

## Objective
- Understand the concept of lookup functions in Excel.
- Retrieve related data from different worksheets or tables.
- Merge datasets using common fields such as Employee ID or Product ID.
- Compare traditional and modern lookup methods.
- Build dynamic and accurate data retrieval formulas.

---

## Software Used
- **Microsoft Excel 2021 / Microsoft 365**

---

# Introduction

Lookup functions are used to search for a value in one table and return the corresponding value from another table. They are essential for combining datasets, creating reports, and avoiding manual data entry.

In this experiment, different lookup techniques were used to merge employee and product data based on unique IDs.

---

# Functions Used

## 1. VLOOKUP

**Purpose:** Searches vertically in the first column of a table and returns a value from another column.

**Syntax**
```excel
=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)
```

**Example**
```excel
=VLOOKUP(A2,$H$2:$K$10,2,FALSE)
```

Returns the employee name corresponding to the Employee ID.

---

## 2. HLOOKUP

**Purpose:** Searches horizontally across the first row of a table.

**Syntax**
```excel
=HLOOKUP(lookup_value, table_array, row_index_num, FALSE)
```

**Example**
```excel
=HLOOKUP(B1,$A$1:$F$5,3,FALSE)
```

Retrieves data stored in a horizontal table.

---

## 3. XLOOKUP

**Purpose:** Searches any direction and returns the matching value.

**Syntax**
```excel
=XLOOKUP(lookup_value, lookup_array, return_array)
```

**Example**
```excel
=XLOOKUP(A2,H2:H10,I2:I10)
```

Returns the employee name without requiring a column index.

---

## 4. INDEX + MATCH

**Purpose:** Performs flexible and dynamic lookups.

**Syntax**
```excel
=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))
```

**Example**
```excel
=INDEX(I2:I10,MATCH(A2,H2:H10,0))
```

Returns the matching value even if columns are rearranged.

---

# Procedure

1. Open the Excel workbook containing two or more related tables.
2. Identify the common key (Employee ID, Product ID, etc.).
3. Apply **VLOOKUP** to retrieve related information.
4. Use **HLOOKUP** for horizontally arranged datasets.
5. Replace VLOOKUP with **XLOOKUP** for improved flexibility.
6. Use **INDEX + MATCH** to perform dynamic lookups.
7. Verify that all retrieved values match the source table.
8. Save the workbook after completing the exercise.

---

# Sample Formulas

| Function | Formula |
|----------|---------|
| VLOOKUP | `=VLOOKUP(A2,$H$2:$K$10,2,FALSE)` |
| HLOOKUP | `=HLOOKUP(B1,$A$1:$F$5,3,FALSE)` |
| XLOOKUP | `=XLOOKUP(A2,H2:H10,I2:I10)` |
| INDEX + MATCH | `=INDEX(I2:I10,MATCH(A2,H2:H10,0))` |

---

# Applications

- Employee information retrieval
- Student record management
- Product inventory lookup
- Sales and customer database merging
- Payroll and HR reporting

---

# Result

Successfully merged data from multiple tables and retrieved accurate records using **VLOOKUP, HLOOKUP, XLOOKUP, and INDEX + MATCH**. The experiment demonstrated efficient methods for data integration and dynamic lookup operations in Microsoft Excel.

---

## Author

**Sakshi Shinde**  
*BE Computer Engineering*
