# Running the Streamlit Version

## What is Streamlit?

Streamlit is a modern Python framework that makes it easy to create beautiful web applications with pure Python. No HTML, CSS, or JavaScript knowledge required!

## Benefits of the Streamlit Version

✅ **Modern UI** - Beautiful, professional-looking interface  
✅ **File Upload** - Easy drag-and-drop file uploads  
✅ **Batch Downloads** - Download all reports as a ZIP file  
✅ **Web-Based** - Access from any device with a browser  
✅ **No Installation** - Users don't need to install Python or the app  
✅ **Responsive** - Works on desktop, tablet, and mobile  
✅ **Real-Time Feedback** - See progress and status updates  

## Installation

### Step 1: Install Dependencies

```bash
pip install -r requirements_streamlit.txt
```

Or manually:
```bash
pip install streamlit pandas openpyxl python-docx
```

### Step 2: Run the App

```bash
streamlit run app_streamlit.py
```

The app will open in your browser at `http://localhost:8501`

## Usage

1. **Home Page**
   - Learn about the app
   - See supported features
   - Understand file format requirements

2. **Generate Reports Page**
   - Upload Excel gradesheet
   - Select sheet name from dropdown
   - Upload Word template
   - Click "Generate Report Cards"
   - Download all reports as ZIP or individually

## Comparison: Tkinter vs Streamlit

| Feature | Tkinter | Streamlit |
|---------|---------|-----------|
| File Upload | File browser dialog | Drag-and-drop |
| Download | Manual folder selection | Direct download buttons |
| UI Quality | Basic | Modern & Professional |
| Deployment | Desktop app | Web app / Cloud |
| User Experience | Traditional desktop | Modern web app |
| Mobile Access | No | Yes |
| Installation | Executable needed | Just run script |
| Customization | Moderate | Easy with markdown |

## Deployment Options

### Option 1: Local Network (Easiest)
```bash
streamlit run app_streamlit.py
# Share the URL with your team on the same network
```

### Option 2: Free Cloud Deployment (Streamlit Cloud)

1. Push your code to GitHub
2. Go to https://streamlit.io/cloud
3. Click "New app"
4. Connect your GitHub repository
5. Select `app_streamlit.py`
6. Deploy in one click!

**Benefits:**
- Free hosting
- Your team can use it from anywhere
- No local setup needed
- Automatic updates from GitHub

### Option 3: Self-Hosted (More Control)

Deploy on your own server:
- AWS EC2
- Google Cloud
- Azure
- DigitalOcean
- Your school's server

```bash
# On your server:
streamlit run app_streamlit.py --server.port 80 --server.address 0.0.0.0
```

## File Format Requirements

### Excel File (.xlsx)
- Student ID in first column (numeric)
- Student data starting from row with first student number
- All subject grades as numbers
- Headers in the row immediately before data starts

### Word Template (.docx)
- Include placeholders like `{{NAME}}`, `{{MATH}}`, etc.
- See [Available Placeholders](#available-placeholders) for complete list
- Template will be filled in for each student

## Available Placeholders

See the app for the complete list of placeholders, or check the embedded help in the "Generate Reports" page.

## Troubleshooting

### "Streamlit is not installed"
```bash
pip install streamlit
```

### Port 8501 already in use
```bash
streamlit run app_streamlit.py --server.port 8502
```

### File upload fails
- Check file size (should be under 200MB)
- Ensure file format is correct (.xlsx and .docx)
- Try refreshing the page

### Reports not generating
- Verify Excel file has student data starting with a number
- Check that Word template has valid placeholders
- Look for error messages in the console

## Keyboard Shortcuts

- `R` - Rerun the app
- `C` - Clear cache
- `Q` - Quit (if running locally)

## Tips

1. **Speed Up**: Run locally first to test, then deploy to cloud
2. **Updates**: Just edit the code and save - Streamlit will auto-reload
3. **Sharing**: Deploy to Streamlit Cloud for easy sharing
4. **Mobile**: The app works great on mobile devices too!

## Next Steps

1. Test the app locally with sample data
2. Deploy to Streamlit Cloud if sharing with team
3. Customize styling by editing CSS in the code
4. Add more features as needed

---

**Still prefer Tkinter?** Use `python3 report_card_gui.py` or `python3 report_card_gui_multipage.py` instead.
