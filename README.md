# JWFCSS Report Card Generator

Automatically generate personalized report cards for all your students using Excel gradesheets and Word templates.

## 🚀 Quick Start

### Option 1: Streamlit (Web-Based - Recommended for Teams)

```bash
pip install -r requirements_streamlit.txt
streamlit run app_streamlit.py
```

**Benefits:** Modern UI, file drag-and-drop, batch downloads, team sharing

→ See [STREAMLIT_GUIDE.md](STREAMLIT_GUIDE.md) for detailed instructions

### Option 2: Desktop App (Tkinter)

```bash
python3 report_card_gui.py
```

Or for multi-page version with navigation:

```bash
python3 report_card_gui_multipage.py
```

### Option 3: Windows Executable

Download the latest release or build using Docker:

```bash
docker build -f Dockerfile.windows -t report-card-builder .
docker run --rm -v $(pwd)/dist:/app/dist report-card-builder
```

→ See [WINDOWS_BUILD_INSTRUCTIONS.md](WINDOWS_BUILD_INSTRUCTIONS.md) and [DOCKER_BUILD_INSTRUCTIONS.md](DOCKER_BUILD_INSTRUCTIONS.md)

## 📋 Features

✅ Batch generate report cards for entire classes  
✅ Automatic grade-based comments  
✅ Support for multiple subjects and behavioral assessments  
✅ Customizable Word templates  
✅ Process multiple classes in one workbook  
✅ Fast and efficient document creation  
✅ Available as desktop app or web app

## 📁 Project Structure

```
report_card_generator/
├── app_streamlit.py              # Streamlit web version (NEW)
├── report_card_gui.py            # Tkinter desktop version
├── report_card_gui_multipage.py  # Tkinter desktop with navigation
├── requirements.txt              # Python dependencies
├── requirements_streamlit.txt    # Streamlit dependencies
├── USER_MANUAL.md                # User guide
├── STREAMLIT_GUIDE.md            # How to run Streamlit version
├── WINDOWS_BUILD_INSTRUCTIONS.md # Build Windows executable
├── DOCKER_BUILD_INSTRUCTIONS.md  # Build with Docker
└── build_windows_exe.spec        # PyInstaller spec file
```

## 🎯 Choose Your Version

| Version                  | Best For                           | How to Run                             |
| ------------------------ | ---------------------------------- | -------------------------------------- |
| **Streamlit**            | Teams, cloud deployment, modern UI | `streamlit run app_streamlit.py`       |
| **Tkinter (Single)**     | Simple desktop app                 | `python3 report_card_gui.py`           |
| **Tkinter (Multi-page)** | Desktop app with navigation        | `python3 report_card_gui_multipage.py` |
| **Windows Executable**   | Windows users, no Python needed    | `Report_Card_Generator.exe`            |

## 📖 Documentation

- [USER_MANUAL.md](USER_MANUAL.md) - Complete user guide with examples
- [STREAMLIT_GUIDE.md](STREAMLIT_GUIDE.md) - Streamlit version setup and deployment
- [WINDOWS_BUILD_INSTRUCTIONS.md](WINDOWS_BUILD_INSTRUCTIONS.md) - Build Windows executable
- [DOCKER_BUILD_INSTRUCTIONS.md](DOCKER_BUILD_INSTRUCTIONS.md) - Build with Docker

## 🔧 Installation

### Requirements

- Python 3.7 or higher
- pandas
- openpyxl
- python-docx

### Install Dependencies

For desktop version:

```bash
pip install -r requirements.txt
```

For Streamlit version:

```bash
pip install -r requirements_streamlit.txt
```

## 📝 File Format Requirements

### Excel Gradesheet (.xlsx)

- Student ID in first column (numeric values)
- Student data with headers in specific columns
- All grades as numeric values

### Word Template (.docx)

- Include placeholders: `{{NAME}}`, `{{MATH}}`, `{{SAVG}}`, etc.
- Template is filled in for each student
- See [USER_MANUAL.md](USER_MANUAL.md) for complete placeholder list

## 🚀 Deployment

### Local Desktop

1. Install Python and dependencies
2. Run `streamlit run app_streamlit.py` or `python3 report_card_gui.py`

### Team Sharing (Streamlit Cloud)

1. Push code to GitHub
2. Go to https://streamlit.io/cloud
3. Deploy in one click
4. Share URL with team

### Self-Hosted Server

1. Set up on your server
2. Run: `streamlit run app_streamlit.py --server.port 80`
3. Access from `http://your-server-ip`

## 🆘 Troubleshooting

### Streamlit version won't start

```bash
pip install --upgrade streamlit
streamlit run app_streamlit.py
```

### "Module not found" errors

```bash
pip install -r requirements_streamlit.txt  # For Streamlit
# or
pip install -r requirements.txt            # For desktop
```

### File upload issues

- Ensure files are under 200MB
- Check file format (.xlsx and .docx)
- Try refreshing the browser

See [USER_MANUAL.md](USER_MANUAL.md) for more troubleshooting.

## 📊 Example Data

See the documentation for example Excel and Word template formats.

## 🤝 Contributing

Feel free to submit issues and enhancement requests!

## 📄 License

This project is provided as-is for educational and school use.

## 📞 Support

For help:

1. Check [USER_MANUAL.md](USER_MANUAL.md)
2. Check [STREAMLIT_GUIDE.md](STREAMLIT_GUIDE.md) if using web version
3. Review error messages carefully
4. Check the GitHub issues page

---

**Prefer web app?** Use `streamlit run app_streamlit.py`  
**Prefer desktop?** Use `python3 report_card_gui.py` or `python3 report_card_gui_multipage.py`  
**Prefer Windows exe?** Download or build with Docker
