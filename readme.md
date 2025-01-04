# Email-to-ICS

This Google Apps Script (GAS) project automates the creation of calendar invites (`.ics` files) based on email content. When a new email arrives at a designated inbox, the script processes the email to extract event details using the OpenAI API, generates an `.ics` file, and sends a reply with the event invite attached.

## Features

- **Automated Email Processing**: Automatically scans for unread emails in the designated inbox.
- **AI-Powered Event Extraction**: Uses OpenAI API to extract event details (title, start time, end time, location) from email content.
- **Calendar Invite Generation**: Creates `.ics` files compatible with most calendar applications.
- **Automated Reply**: Sends a reply email with the generated `.ics` file attached.

## How It Works

1. Monitors a specified email address for unread messages.
2. Extracts event details from the most recent email in each thread using the OpenAI API.
3. Generates an `.ics` file with the extracted event details.
4. Replies to the original email with the `.ics` file attached.

## Prerequisites

- A Google Workspace account.
- Access to the Gmail and Google Apps Script APIs.
- An OpenAI API key for accessing the GPT model.

## Setup Instructions

1. **Clone the Repository**: 
   Clone this repository or copy the script file into the Google Apps Script editor.

2. **Configure the Script**:
   - Replace the `TARGET_EMAIL` constant with the email address you want to monitor.
   - Set up your OpenAI API key in the script's properties:
     - Navigate to `File > Project Properties > Script Properties` in the Apps Script editor.
     - Add a property with the name `OPENAI_API_KEY` and your OpenAI API key as the value.

3. **Authorize Permissions**:
   - Run the `createTrigger` function to set up the time-driven trigger (runs every minute).
   - Grant the necessary permissions when prompted.

4. **Deploy the Script**:
   - Deploy the script as a standalone Apps Script project.

## Usage

1. Send an email to the configured `TARGET_EMAIL` address.
2. The script will process the email, extract event details, generate an `.ics` file, and send a reply.
3. Open the `.ics` file to add the event to your calendar.

## Contributing

Contributions, suggestions, and bug reports are welcome! Please open an issue or submit a pull request to improve this project.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

---

**Developer Contact**: For questions or feature suggestions, reach out to [edfagedeveloper@gmail.com](mailto:edfagedeveloper@gmail.com).
