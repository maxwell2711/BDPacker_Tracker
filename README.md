# BDPacker Tracker

This repository contains a Python utility that reads an Excel file containing order/packing data, calculates summary statistics, and writes the results to a new Excel workbook. A minimal Tkinter GUI lets you choose the source file and start the processing.

## Features

- **Packing Time Summary** – sums `packing time` for each packer on each day.
- **Profit Rate Summary** – uses `est net profit` and `packing time` to compute average profit per hour and per minute for each packer/day and also overall totals.
- **Color-coded rows** – each packer is assigned a unique background color for easy visual identification. Colours are derived from a pastel palette that varies by date, and packers on the same day get slight hue/saturation offsets so they stay in the same family while remaining distinguishable.
- **Date separators** – thicker border lines separate data from different dates.
- **Smart filename** – output files are named based on the date range: `summary_output_2-24-26_3-1-26.xlsx` (multiple dates) or `summary_output_2-24-26.xlsx` (single date).
- GUI built with Tkinter for file selection and execution.

## Requirements

- Python 3.8+
- `pandas`
- `openpyxl` (for Excel read/write and formatting)

Install dependencies with pip:

```sh
pip install pandas openpyxl
```

> Tkinter is included with most standard Python installations on Windows and macOS. If you encounter import errors, you may need to install/enable it separately.

## Usage

1. Run the script:
   ```sh
   python bdpacker_tracker.py
   ```
2. Click **Browse...** to select the input Excel file (it should contain the columns described in your example).
3. Click **Process**, choose a location/name for the output workbook (default filename is auto-generated), and the two summary sheets will be generated.

### Columns expected in the input file
The program expects at least the following headers (case-insensitive, leading/trailing spaces ignored):

```
first name,last name,email,date packed,packed by,packing time,est net profit
```

Additional columns are read but ignored.

## Output

The generated workbook contains two worksheets:

1. **Packing Time Summary** 
   - Columns: `date packed`, `packed by`, `total_hours`, `total_minutes`
   - Each packer has a consistent color offset that persists across all dates
   - Different dates use different base colors from the palette, with packer offsets applied to each (the code normalises Excel datetimes back to dates when looking up colours to avoid mismatches).

2. **Profit Rate Summary** 
   - Columns: `date packed`, `packed by`, `total_minutes`, `total_profit`, `profit_per_hour`, `profit_per_minute`
   - Same formatting and coloring as the packing time summary
   - Includes overall daily totals (marked with `<all>`, displayed in white)

## Formatting Details

- **Color scheme**: Base colors rotate through a 7-color pastel palette for each date (pink, blue, green, yellow, lavender, coral, cyan). Each packer is assigned a unique RGB offset that's applied consistently across all dates.
- **Date separators**: Thick single-line borders appear below the last row of each date group for clear visual separation between days.
- **Column widths**: Automatically adjusted to fit content (max 50 chars), so all data is readable without manual adjustment.
- **Minimal borders**: Thin borders on all cells for readability; no column dividing lines between individual columns.
- **Overall totals**: Included in the Profit Rate Summary with `<all>` as the packer name.

Feel free to adjust the script if your column names differ or you need extra calculations.
