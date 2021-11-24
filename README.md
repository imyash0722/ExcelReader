<div align="center">

# ExcelReader

**Python tool for reading Excel files and extracting data as key-value pairs.**

<p>
  <a href="https://github.com/imyash0722/ExcelReader/releases/latest">
    <img alt="Latest Release" src="https://img.shields.io/github/v/release/imyash0722/ExcelReader?style=for-the-badge&logo=github&logoColor=white&label=Latest&color=4caf50">
  </a>
  &nbsp;
  <img alt="Python" src="https://img.shields.io/badge/Python-3.x-333?style=for-the-badge&logo=python&logoColor=white">
  &nbsp;
  <img alt="Library" src="https://img.shields.io/badge/pyexcel-library-333?style=for-the-badge&logoColor=white">
</p>

</div>

## Features

|                              |                                                             |
| ---------------------------- | ----------------------------------------------------------- |
| 📄 **Multi-format Support**  | Reads both `.xlsx` and `.xls` files                        |
| 🔄 **Auto Conversion**       | Automatically converts `.xls` to `.xlsx` before processing |
| 🗂️ **Sheet Parsing**         | Iterates all sheets and outputs data as key-value pairs     |
| ⚡ **Simple API**            | One function call to read any supported Excel file          |

## Quick Start

### Requirements

```bash
pip install pyexcel pyexcel-xlsx pyexcel-xls
```

### Run

```python
from excel_reader import main_runner

main_runner("path/to/your/file.xlsx")
```

## Usage

```python
# For .xlsx files — reads directly
main_runner("data.xlsx")

# For .xls files — auto-converts to .xlsx, then reads
main_runner("data.xls")
```

Output is printed as `SheetName : value1,value2,...` for each sheet row.

## Project Structure

```
ExcelReader/
├── excel_reader.py        # Core reader logic (main_runner + excel_reader_main class)
├── excel_reader.pyproj    # Visual Studio Python project file
├── excel_reader.sln       # Visual Studio solution file
└── README.md
```

## Tech Stack

- **Python 3.x** — core runtime
- **pyexcel** — unified Excel read interface
- **pyexcel-xlsx / pyexcel-xls** — format backend plugins

## License

[MIT](LICENSE)

<div align="center">
  <sub>ExcelReader v1.0.0</sub>
</div>
