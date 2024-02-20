# SERP Analysis Automation with Google Apps Script

This project automates the process of analyzing search engine results pages (SERP) using the Serpstat API, integrated within Google Sheets via Google Apps Script. It enables users to track task progress, refresh task statuses, retrieve and display keyword analysis data, and manage task creation for SERP analysis in a streamlined and efficient manner.

## Features

- **Task Progress Tracking:** Monitors the progress of SERP analysis tasks by task ID.
- **Automatic Refresh:** Updates the status of tasks within a Google Sheet, indicating completion or ongoing progress.
- **Data Retrieval and Display:** Fetches and displays SERP analysis data for specified keywords, including top results, ad placements, and more.
- **Task Management:** Facilitates the creation of new SERP analysis tasks with customizable parameters.

## Setup

### 1. Google Sheets Preparation
   - Create a new Google Sheet.
   - Name the necessary sheets: `Tasks`, `Run`, `Raw Records` for task management, parameter input, and data display.

### 2. Google Apps Script
   - Open the Google Apps Script editor from your Google Sheet (`Extensions > Apps Script`).
   - Copy and paste the provided code into the script editor.
   - Save and name your project.

### 3. Serpstat API Key
   - Obtain an API key from [Serpstat](https://serpstat.com/).
   - Input the API key in the designated cell (`D5`) on the `Run` sheet.

### 4. Configuration
   - Configure the `Run` sheet with your desired analysis parameters (country, language, region, device).

## Usage

- **Creating a Task:** Use the `createTask()` function to initiate a new SERP analysis task based on the parameters specified in the `Run` sheet.
- **Refreshing Task Status:** Execute the `refresh()` function to update the progress and status of tasks listed in the `Tasks` sheet.
- **Viewing Results:** Access the `Raw Records` sheet to see the detailed results of completed tasks.

## Contributing

Contributions are welcome! Feel free to fork the project and submit pull requests with any enhancements, bug fixes, or improvements.

## License

This project is released under the MIT License. See the `LICENSE` file for more details.

## Disclaimer

This project is not affiliated with Serpstat. It uses the Serpstat API under their terms and conditions. Please ensure you comply with Serpstat's API usage policies.
