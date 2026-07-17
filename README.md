Employee Productivity Score Calculator

Automates employee productivity analysis in Microsoft Excel using Office Scripts.

This project processes employee work metrics from multiple worksheets, performs score calculations, validates data, and generates a consolidated report automatically.

---

## Features

- Calculates employee productivity scores
- Reads data from multiple worksheets
- Handles missing and invalid values
- Replaces negative scores with a configurable default value
- Automatically generates an output report
- Built using Microsoft Office Scripts

---

## Technologies

- Microsoft Excel
- Office Scripts
- TypeScript

---

## Folder Structure

```
src/
sample-data/
docs/
assets/
```

---

## Input

The script expects an Excel workbook containing employee productivity data.

Example columns include:

- Employee Name
- Employee ID
- Days Worked
- Time on Projects
- Private Time
- Daily Estimated Time

---

## Output

The script generates a processed worksheet containing calculated productivity scores and cleaned data.

---

## How to Run

1. Open Excel for the Web.
2. Navigate to the **Automate** tab.
3. Create a new Office Script.
4. Copy the contents of `src/EmployeeProductivityCalculator.ts`.
5. Run the script on a workbook containing employee productivity data.

---

## Example Workflow

```
Input Workbook
        │
        ▼
Employee Data Processing
        │
        ▼
Score Calculation
        │
        ▼
Data Validation
        │
        ▼
Generated Productivity Report
```

## Author

**Arsheya Prasad and Simonne Kulkarni**

Computer Science @ Virginia Tech

