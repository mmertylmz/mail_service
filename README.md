
# Email Attachment Fetcher Service

This Python project implements a Windows service that automatically fetches email attachments from a Gmail account every minute and saves them to a local directory.

## Features

- **Runs as a Windows Service**: The service fetches email attachments in the background.
- **Logs**: Keeps logs of service start, stop, and any errors encountered.
- **Error Handling**: Captures any errors during the email fetch process and logs them.
- **Attachment Saving**: Saves email attachments in a specified folder.

## Prerequisites

- **Python**: Make sure you have Python installed.
- **Libraries**: The following Python libraries are required:
  - `pywin32`: For creating and managing Windows services.
  - `imaplib`: For handling IMAP email protocols.
  - `email`: To process and decode email messages.
  - `configparser`: For reading configurations from a `config.ini` file.

  Install the required libraries using:

  ```bash
  pip install pywin32
  ```

## Setup

1. Clone the repository to your local machine:

   ```bash
   git clone https://github.com/yourusername/yourrepository.git
   cd yourrepository
   ```

2. Create a `config.ini` file in the root directory of your project with the following structure:

   ```ini
   [GMAIL]
   username = your-email@gmail.com
   password = your-email-password
   imap_server = imap.gmail.com
   ```

3. Make sure you have enabled "Less secure app access" in your Gmail account settings or set up an app password for additional security.

4. Create a directory named `attachments` in the project folder where the service will save the email attachments.

## Running the Service

1. Install the service:

   ```bash
   python main.py install
   ```

2. Start the service:

   ```bash
   python main.py start
   ```

3. Stop the service:

   ```bash
   python main.py stop
   ```

4. Uninstall the service:

   ```bash
   python main.py remove
   ```

## Logs

The service logs are saved in the `log` directory under `service.log`. This file contains information about the service status and any issues that occur during its execution.

## How It Works

- The service connects to a Gmail IMAP server using credentials from the `config.ini` file.
- It checks the inbox for unseen emails every minute.
- If there are any unseen emails with attachments, the service downloads the attachments and saves them in the `attachments` folder.
- The service keeps running in the background until manually stopped.
