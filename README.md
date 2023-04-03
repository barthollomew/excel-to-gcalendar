# Excel to Google Calendar 🗓️

Provides a way for employees to easily add their rosters from Excel spreadsheets to their Google Calendar. The script reads data from the Excel file, extracts event details using regular expressions, and creates a new event in the user's calendar using the Google Calendar API. The script also handles errors and prints informative messages to the user.

## Getting Started 🚀

### Prerequisites 🧰

To use this project, you will need the following:

- Python 3.6 or later 🐍
- An Excel spreadsheet containing the roster data 📊
- A Google Cloud project with the Google Calendar API enabled ☁️
- A service account with a JSON key file that has access to the Google Calendar API 🔑

### Installing 💻

1. Clone the project repository:

   `https://github.com/barthollomew/excel-to-gcalendar`

2. Install the required Python packages:

   `pip install pytz openpyxl google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client`

3. Set the `TIMEZONE`, `CREDENTIALS_FILE`, and `SCOPES` constants in the script to match your timezone, credentials file path, and desired Google Calendar API scopes.

### Usage 🚀

1. Navigate to the project directory.

2. Run the script:

   `python add_roster_to_calendar.py`

3. Follow the prompts to enter your name and select an event from your roster.

4. The script will then add the selected event to your Google Calendar.

### License 📜

This project is licensed under the MIT License. See the `LICENSE` file for details.

### Acknowledgments 🙏

- The Google Calendar API documentation was used as a reference for this project.
- The `pytz`, `openpyxl`, and `google-auth` packages were used in this project. 
