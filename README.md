# Online Meeting (ZOOM) Attendance Report Maker

A Python-based utility for automating the processing of online meeting attendance exports into structured Excel reports. The application is designed to simplify attendance analysis by generating attendance, late, absent, and filtered participant reports from Zoom CSV exports.

> **Version:** v0.1.0

## Features

- Attendance report generation
- Late participant detection
- Absent participant identification
- Automatic participant duplicate handling
- Participant name normalization
- Blacklisted participant filtering (CF sheet)
- Color-coded Excel formatting
- Simple drag-and-drop interface
- Batch presets for different meeting schedules

## Output

The application generates an Excel workbook containing:

- **RAW** – Raw attendance data
- **ATT** – Attendance summary
- **ABSENT** – Participants below the attendance threshold
- **LATE** – Participants joining after the configured cutoff time
- **CF** – Filtered participants

## Requirements

- Python 3.10+
- Windows

### Python Packages

- customtkinter
- tkinterdnd2
- pandas
- openpyxl

## Usage

1. Launch the application.
2. Drag and drop an Zoom participants report CSV file or browse manually.
3. Select the desired report options.
4. Configure attendance and late thresholds if required.
5. Choose an output location.
6. Generate the report.

## Project Structure

```
.
├── tool.py        # App run point
├── ui.py          # UI
├── core.py        # Attendance processing unit
├── logo.png
└── README.md
```

## Current Status

This is the initial release (**v0.1.0**) and is under active development. Additional improvements and features will be added in future versions.

## Disclaimer

This project was developed as an internal productivity tool for processing online meeting attendance reports. It may require customization to work with attendance exports from different meeting platforms.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
