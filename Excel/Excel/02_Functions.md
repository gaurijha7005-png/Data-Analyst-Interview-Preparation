# 🧮 Excel Functions for Data Analysts

Excel functions help automate calculations, analyze data efficiently, and reduce manual work. Mastering these functions is essential for every Data Analyst.

---

# What is an Excel Function?

An Excel function is a predefined formula that performs a specific calculation using values called **arguments**.

### Example

```excel
=SUM(A1:A10)
```

This function adds all the numbers from cell **A1** to **A10**.

---

# Why are Functions Important?

Functions help you:

- Perform calculations quickly
- Analyze large datasets
- Reduce manual errors
- Automate repetitive tasks
- Build dynamic reports and dashboards

---

# Categories of Excel Functions

## 1. Mathematical Functions

Used for calculations.

| Function | Purpose | Example |
|----------|---------|---------|
| SUM | Adds numbers | =SUM(A1:A10) |
| AVERAGE | Finds average | =AVERAGE(A1:A10) |
| MIN | Smallest value | =MIN(A1:A10) |
| MAX | Largest value | =MAX(A1:A10) |
| ROUND | Rounds a number | =ROUND(A1,2) |

---

## 2. Logical Functions

Used for decision-making.

| Function | Purpose |
|----------|---------|
| IF | Checks a condition |
| IFS | Multiple conditions |
| AND | Returns TRUE if all conditions are TRUE |
| OR | Returns TRUE if any condition is TRUE |
| NOT | Reverses TRUE/FALSE |

### Example

```excel
=IF(B2>=50,"Pass","Fail")
```

---

## 3. Text Functions

Used to manipulate text.

| Function | Purpose |
|----------|---------|
| LEFT | Extracts characters from the left |
| RIGHT | Extracts characters from the right |
| MID | Extracts characters from the middle |
| LEN | Counts characters |
| TRIM | Removes extra spaces |
| UPPER | Converts to uppercase |
| LOWER | Converts to lowercase |
| PROPER | Capitalizes each word |
| CONCAT | Combines text |

Example:

```excel
=CONCAT(A2," ",B2)
```

---

## 4. Lookup Functions

Used to search for values in a dataset.

- XLOOKUP
- VLOOKUP
- HLOOKUP
- INDEX
- MATCH

These functions are commonly used to retrieve information from large tables.

---

## 5. Date & Time Functions

Useful for analyzing dates.

| Function | Purpose |
|----------|---------|
| TODAY | Current date |
| NOW | Current date & time |
| YEAR | Extracts year |
| MONTH | Extracts month |
| DAY | Extracts day |

Example:

```excel
=YEAR(A2)
```

---

# Most Important Functions for Data Analysts

- SUM
- AVERAGE
- IF
- SUMIF
- SUMIFS
- COUNTIF
- COUNTIFS
- XLOOKUP
- INDEX
- MATCH
- IFERROR
- CONCAT
- TEXT
- ROUND
- TODAY

---

# Best Practices

- Use absolute references (`$A$1`) when required.
- Keep formulas simple and readable.
- Use IFERROR to avoid displaying errors.
- Name important ranges for better readability.
- Test formulas on sample data before applying them to the entire dataset.

---

# Common Interview Questions
