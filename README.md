# Google Sheets to Google Calendar Automation

This GitHub repository hosts the code for automating the process of transferring data from a Google Sheet into Google Calendar. This project was created to streamline and automate the process of managing calendar events, significantly reducing the time spent manually entering information.

## Project Description

The goal of this project is to automate the conversion of events stored in a Google Sheets document into Google Calendar events. The automation script identifies cells with a person's name, determines the last time of the day for them, and formats this data into a Google Calendar-readable format.

## Features

- **Automates Event Conversion**: Transfers events from Google Sheets to Google Calendar.
- **Cell Identification**: Identifies cells containing names and their corresponding events and times.
- **Date Handling**: Each sheet in the spreadsheet represents a specific date in the format `month/year`.
- **Time Efficiency**: Saves time by eliminating the need to manually enter event information into Google Calendar.

## Why I Built This

I built this project because I love automation and wanted to reduce the time I spend manually inputting every piece of information into my calendar. This tool helps me manage my schedule more efficiently, allowing me to focus on more important tasks.

## Google Sheets Format

For this script to work correctly, your Google Sheets document must follow a specific format:

- Each sheet in the spreadsheet should represent the date of the event, formatted as `month/year`.
- The sheet should contain columns for names, events, and times.

## Getting Started

### Prerequisites

- A Google Sheets document formatted as described above.
- Access to the Google Calendar API.

### Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/ebod12/auto-sheet-to-calendar.git
    ```
2. Install the required dependencies:
    ```sh
    pip install -r requirements.txt
    ```
3. Set up Google API credentials and save the `credentials.json` file in the project directory.

### Usage

1. Update the script with the specific details of your Google Sheets document.
2. Run the script:
    ```sh
    python sheet_to_calendar.py
    ```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

## Contact

If you have any questions or suggestions, please open an issue or contact me directly at olekabrida@gmail.com.

---

This README provides a clear and detailed description of your project, its purpose, and how to use it, making it easy for others to understand and contribute.
