


  India Fund SIF NAV Automation
This project automates the process of downloading   NAV (Net Asset Value) data   for India SIF funds from the [AMFI India website](https://www.amfiindia.com/sif/latest-nav/nav-history), processes the downloaded files, and sends them via email using Outlook.


   🚀 Features
- Automatically navigates to AMFI India’s NAV history page.
- Iterates through all available   Funds   and   Strategies  .
- Downloads NAV data in Excel format for a given date range.
- Cleans and converts raw Excel files into a structured format.
- Sends processed NAV data files via Outlook email to a specified recipient.
- Handles missing data gracefully (skips if no NAV data is available).


   ⚙️ Requirements
-   Python 3.8+  
- Google Chrome browser
- ChromeDriver (compatible with your Chrome version)
- Microsoft Outlook (installed and configured)


    Python Libraries
Install required dependencies:
```bash
pip install pandas selenium pywin32
```

---

   📂 Project Structure
```
India-SIF-Automation/
│
├── automation_script.py     Main automation script
├── README.md                Documentation
└── requirements.txt         Python dependencies
```

---

   🔧 Configuration
Update the following   user settings   in the script before running:

```python
FROM_DATE = "03/11/2026"     Start date (MM/DD/YYYY)
TO_DATE = "03/12/2026"       End date (MM/DD/YYYY)

DOWNLOAD_FOLDER = r"C:\Users\asalvi\OneDrive - MORNINGSTAR INC\Documents\India Fund SIF Automation"

EMAIL_RECEIVER = "abhishek.salvi@morningstar.com"

SUBJECT_PREFIX = "India SIF NAV Data"
```

---

   ▶️ Usage
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/India-SIF-Automation.git
   cd India-SIF-Automation
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the script:
   ```bash
   python automation_script.py
   ```

4. The script will:
   - Download NAV files
   - Convert them into structured Excel files
   - Send them via Outlook email

---

   📧 Email Output
Each email sent will include:
- Subject: `India SIF NAV Data - <Investment Name>`
- Body:
  ```
  Hello,

  Please find attached NAV data file for:

  <Investment Name>

  Regards
  Morningstar India Pvt. Ltd.
  ```

- Attachment: Processed NAV Excel file

---

   ⚠️ Notes
- Ensure Outlook is open and configured with the correct account.
- ChromeDriver must match your installed Chrome version.
- The script deletes old raw NAV files before downloading new ones.
- Add delays (`time.sleep`) if Outlook throttles email sending.

---

   🛠️ Future Enhancements
- Add logging for better monitoring.
- Support multiple recipients.
- Error handling improvements.
- Dockerize for deployment.



