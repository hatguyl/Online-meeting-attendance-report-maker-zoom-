# Online Meeting (ZOOM) Attendance Report Maker

Desktop application for automating the processing of Zoom meeting attendance exports into structured Excel reports. Designed as an internal productivity tool, it streamlines attendance analysis by generating organized reports with intelligent participant matching, duplicate resolution, and customizable attendance rules.

> **Version:** **v1.0.0**

---

## ✨ Features

### Attendance Processing
- Generate Attendance, Late, Absent, and Filtered participant reports
- Automatic participant duplicate detection and merging
- Intelligent participant grouping using roll numbers and aliases
- Automatic participant name normalization
- Advanced roll number extraction supporting multiple formats
- Attendance and absent threshold customization
- Configurable late cutoff time
- Batch presets for different meeting schedules

### Report Generation
- Professionally formatted Excel reports
- Automatic worksheet formatting
- Color-coded attendance status
- Alphabetically sorted reports
- RAW report sorted by participant name and join time
- Automatic column sizing
- Excel filters enabled on all worksheets

---

## 📊 Generated Reports

The application creates a single Excel workbook containing:

| Sheet | Description |
|-------|-------------|
| **RAW** | Original participant data sorted by name and join time |
| **ATT** | Attendance summary with merged duplicate participants |
| **ABSENT** | Participants below the configured attendance threshold |
| **LATE** | Participants joining after the configured cutoff time |
| **CF** | Filtered participants matching blacklist rules |

---

## 🚀 Improvements

- Completely redesigned participant sorting
- More accurate duplicate detection
- Improved roll number pattern recognition
- Better participant grouping logic
- Improved report sorting and formatting
- Enhanced Excel highlighting
- Improved drag-and-drop support
- Improved error handling
- Faster, reliable report generation

---

## 🖥 Requirements

- Windows
- Python **3.10+**

### Python Packages

- customtkinter
- tkinterdnd2
- pandas
- openpyxl

## 📦 Usage

1. Launch the application.
2. Drag and drop a Zoom participant report CSV file or browse manually.
3. Configure attendance, absent, and late settings if required.
4. Select the desired report options.
5. Choose an output location.
6. Click **Generate Report**.

---

## 📁 Project Structure

```
.
├── tool.py        # App entry point
├── ui.py          # UI
├── core.py        # Attendance processor
├── logo.png
├── LICENSE
└── README.md
```

---

## 📝 Changelog

### v1.0.0
- Major rewrite of the attendance processing engine
- Improved duplicate participant detection
- Better sorting
- Improved UI 
- Enhanced validation and error handling
- Performance and reliability improvements

---

## ⚠ Disclaimer

This application was developed as an internal productivity tool for processing Zoom meeting attendance reports. While it works with standard Zoom participant exports, additional customization may be required for other meeting platforms or customized report formats.

---

## 📄 License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.
