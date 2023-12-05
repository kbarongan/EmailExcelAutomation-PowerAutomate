# Power Automate Desktop - Email Automation 📧🤖

## Overview
This repository contains a Power Automate Desktop script for automating email processing with Microsoft Outlook and Excel. The script retrieves unread emails with attachments from a specified Outlook folder, processes CSV files from these attachments, updates an Excel file, and sends the updated file back via email.

## Features 🌟
- **Email Retrieval** 📥: Automatically fetches unread emails from a specified Outlook folder.
- **Attachment Processing** 📂: Extracts and processes CSV file attachments.
- **Excel Integration** 📊: Updates a master Excel file with data from the CSV files.
- **Email Notification** 📤: Sends an updated Excel file via Outlook email.

## Prerequisites ✅
- Microsoft Power Automate Desktop installed.
- Microsoft Outlook and Excel installed and configured on your machine.
- Access to an Outlook email account.

## Setup Instructions 🛠️
1. Instructions to use in Power Automate Desktop
2. Open Power Automate Desktop and start a new flow.
3. Copy and paste the script text directly into the new flow.
4. **Configure Script Variables** 🔧:
   - `MasterFile`: Path to the Excel file to be updated.
   - `PendingFolder`: Path to the folder where pending CSV files are stored.
   - `CompletedFolder`: Path to the folder for processed files (not used in the current script, but can be integrated for future enhancements).
   - `MyEmail`: Your Outlook email address.
   - `InboxFolderName`: The name of the Outlook folder to retrieve emails from.
5. **Run the Script** ▶️: Open the script with Power Automate Desktop and run it.

## Script Usage 📘
1. **Email Fetching** 📩: The script starts by launching Outlook and fetching unread emails from the specified folder.
2. **CSV Processing** 🔄: It then processes each CSV file attached to the emails.
3. **Excel Updating** 📝: The data from the CSV files are written into the specified Excel file.
4. **Email Sending** 🚀: After processing, the script sends an email with the updated Excel file as an attachment.

## Contributions 🤝
Feel free to fork this repository and contribute to enhancing its functionality. Possible enhancements include error handling, logging, and processing of the `CompletedFolder`.

## License 📜
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Disclaimer ⚠️
This script is provided as-is, and users should modify it to suit their specific needs. Always test the script in a non-production environment first.
