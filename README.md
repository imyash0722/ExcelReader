<div align="center">

# ExcelReader

**Python tool for reading and writing Excel files using xlwings.**

<p>
  <a href="https://github.com/imyash0722/ExcelReader/releases/latest">
    <img alt="Latest Release" src="https://img.shields.io/github/v/release/imyash0722/ExcelReader?style=for-the-badge&logo=github&logoColor=white&label=Latest&color=2196f3">
  </a>
  &nbsp;
  <img alt="Python" src="https://img.shields.io/badge/Python-3.x-333?style=for-the-badge&logo=python&logoColor=white">
  &nbsp;
  <img alt="Library" src="https://img.shields.io/badge/xlwings-library-333?style=for-the-badge&logoColor=white">
  &nbsp;
  <img alt="Platform" src="https://img.shields.io/badge/Windows%20%2F%20macOS-333?style=for-the-badge&logo=windows&logoColor=white">
</p>

</div>

> [!NOTE]
> xlwings requires Microsoft Excel to be installed on your machine. It interfaces directly with a live Excel instance via COM automation.

## Features

|                               |                                                               |
| ----------------------------- | ------------------------------------------------------------- |
| 📖 **Read Excel Data**        | Read individual cells, ranges, or entire tables               |
| ✏️ **Write Excel Data**       | Write single values, lists, or 2D arrays to cells             |
| 📊 **Range Operations**       | Full support for cell ranges like `A1:C4`                     |
| 🔗 **Live Excel Integration** | Operates on a live Excel workbook via COM automation          |

## Quick Start

### Requirements

```bash
pip install xlwings
```

> Microsoft Excel must be installed on your system.

### Run

```python
import xlwings as xw

ws = xw.Book("path/to/your/file.xlsx")
print(ws.range("A1").expand().value)
```

## Usage

```python
import xlwings as xw

ws = xw.Book("Timetable.xlsx")

# Read entire table
print(ws.range("A1").expand().value)

# Write a single value
ws.range("A1").value = "Hello"

# Write a list across a row
ws.range("B1").value = ["a", "b", "c"]

# Write a 2D array
ws.range("A2").value = [[1, 2, 3], ["x", "y", "z"]]
```

## Project Structure

```
ExcelReader/
├── main.py        # Excel read/write examples using xlwings
└── README.md
```

## Tech Stack

- **Python 3.x** — core runtime
- **xlwings** — Excel automation via COM (read + write)
- **Microsoft Excel** — required for xlwings COM interface

## License

[MIT](LICENSE)

<div align="center">
  <sub>ExcelReader v2.0.0</sub>
</div>
