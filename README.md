# 🏥 Schedule Management System

A Streamlit-based web application for managing and converting Radiology and Cardiology on-call schedules.

**Project by JA RAD**

## 📋 Overview

This application provides a simple, user-friendly interface for scheduling staff to:
- Generate work and on-call schedules
- Convert existing schedules to import-ready CSV format
- Process multiple schedule files without requiring Python installation

## 🚀 Features

### Radiology Section
1. **Create Rad Work Schedule** *(In Development)*
   - Upload blank or partially filled Work Schedule template
   - Generate completed Work Schedule

2. **Create Rad On-Call Schedule** *(In Development)*
   - Upload blank or partially filled On-Call Schedule template
   - Generate completed On-Call Schedule

3. **Convert Rad Schedules for Import** ✅
   - Upload Work Schedule and On-Call Schedule Excel files
   - Generate `Import_OnCall_Radiology.csv` for system import
   - Processes 5 teams: Gen_CT, IRA, MRI, US, Fluoro

### Cardiology Section
1. **Convert Cardiology Schedules for Import** ✅
   - Upload Cardiovascular and Interventional Cardiologist schedules
   - Generate `Import_OnCall_Cardiology.csv` for system import
   - Processes 2 teams: Cardiovascular (Team 8) and Interventional Cardiologist (Team 94)
   - Expandable for future additional teams

## 🛠️ Installation

### Prerequisites
- Python 3.8 or higher
- Git

### Local Setup

1. Clone the repository:
```bash
git clone https://github.com/sharadgeorge/schedules.git
cd schedules
```

2. Create a virtual environment (optional but recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the application:
```bash
streamlit run app.py
```

The app will open in your default browser at `http://localhost:8501`

## ☁️ Deployment to Streamlit Cloud

1. Push your code to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Sign in with GitHub
4. Click "New app"
5. Select your repository: `sharadgeorge/schedules`
6. Set Main file path: `app.py`
7. **Set App URL**: Enter `schedules` to get: `https://schedules.streamlit.app`
8. Click "Deploy"

Your app is now live at: `https://schedules.streamlit.app`

## 📁 Project Structure

```
schedule/
├── app.py                                      # Main app (Radiology landing page)
├── pages/
│   └── cardiology.py                          # Cardiology page
├── oncall_converter_Radiology_demo_v2.py     # Radiology converter script
├── oncall_converter_Cardiology_demo_v3.py    # Cardiology converter script
├── create_Rad_Work_Schedule.py               # Work schedule generator (placeholder)
├── Create_Rad_OnCall_Schedule.py             # On-call schedule generator (placeholder)
├── requirements.txt                           # Python dependencies
├── .gitignore                                # Git ignore rules
└── README.md                                 # This file
```

## 💡 Usage

### For Radiology Schedule Conversion

1. Navigate to the Radiology page (landing page)
2. Scroll to "Convert Rad Schedules for Import"
3. Upload:
   - Work Schedule Excel file (first upload)
   - On-Call Schedule Excel file (second upload)
4. Click "Convert to Import Format"
5. Download the generated `Import_OnCall_Radiology.csv`

### For Cardiology Schedule Conversion

1. Navigate to the Cardiology page (via sidebar)
2. Upload:
   - Cardiovascular Schedule Excel file
   - Interventional Cardiologist Schedule Excel file
3. Click "Convert to Import Format"
4. Download the generated `Import_OnCall_Cardiology.csv`

## 📊 Output Format

Both converters generate CSV files with caret (^) delimiter containing:
- EMPLOYEE: Username/employee ID
- TEAM: Team ID number
- STARTDATE: Schedule start date (M/D/YYYY)
- STARTTIME: Start time (24-hour format, e.g., 700, 1530)
- ENDDATE: Schedule end date (M/D/YYYY)
- ENDTIME: End time (24-hour format)
- ROLE: Role ID
- NOTES: Optional notes
- ORDER: Optional order
- TEAMCOMMENT: Optional team comments

## 🔧 Configuration

### Radiology Teams
- **Gen_CT** (Team 114): General CT - 3 blocks on weekdays
- **IRA** (Team 115): Interventional Radiology
- **MRI** (Team 116): MRI
- **US** (Team 126): Ultrasound
- **Fluoro** (Team 127): Fluoroscopy

### Cardiology Teams
- **Cardiovascular** (Team 8): Echo Tech (Adult & Pediatric)
- **Interventional Cardiologist** (Team 94): Interventional procedures

### Weekend/Weekday Definition
- **Weekdays**: Sunday - Thursday
- **Weekends**: Friday - Saturday

## 🚧 Development Status

- ✅ Radiology Schedule Converter - **Functional**
- ✅ Cardiology Schedule Converter - **Functional**
- ⏳ Radiology Work Schedule Generator - **In Development**
- ⏳ Radiology On-Call Schedule Generator - **In Development**

## 🤝 Contributing

This is an internal project for JA RAD. For questions or support, please contact the development team.

## 📝 License

Internal use only - JA RAD

## 🐛 Troubleshooting

### Common Issues

1. **File Upload Errors**
   - Ensure Excel files are in `.xlsx` format
   - Check that files contain the expected sheet names
   - Verify file structure matches the template

2. **Conversion Errors**
   - Verify employee names/initials match the configured mappings
   - Check that date formats in Excel are correct
   - Ensure required columns/rows exist in the templates

3. **Month Detection Issues**
   - Include month name or abbreviation in filename
   - Or ensure cell B4 contains month information

### Getting Help

If you encounter issues:
1. Check the error message details in the expandable "Error Details" section
2. Verify your input files match the expected format
3. Contact the development team with:
   - Screenshot of the error
   - Sample input files (with sensitive data removed)
   - Description of what you were trying to do

## 🔄 Future Enhancements

- [ ] Implement Work Schedule generator
- [ ] Implement On-Call Schedule generator
- [ ] Add support for additional Cardiology teams
- [ ] Add schedule validation before conversion
- [ ] Add preview of generated schedules
- [ ] Add export to multiple formats
- [ ] Add schedule history and versioning
- [ ] Add user authentication and access control

## 📞 Support

For technical support or feature requests, please contact the JA RAD development team.

---

**Last Updated**: October 2025  
**Version**: 1.0.0
