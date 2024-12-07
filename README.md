📬 Correspondence Registration System (CRS)

The Correspondence Registration System (CRS) automates tracking of incoming and outgoing emails for authorized team members. This system is built using Google Cloud Platform (GCP) and operates through automated process flows (bots) integrated with Google Sheets.
📊 System Overview

The CRS uses two core Google Sheets files:

    HOJA INICIO:
        Initial configuration and authorization setup.
        Links to the PLANILLA sheet for correspondence tracking.

    PLANILLA:
    Contains three tabs:
        Entradas: Records incoming emails.
        Salidas: Records outgoing emails.
        Autorizaciones: Tracks team member authorizations.

⚙️ How It Works

    Authorization:
    Team members authorize email tracking from the HOJA INICIO sheet.

    Daily Automation:
    The system checks email inboxes and outboxes at specified intervals (set in hours).

    Deadline Tracking:
        Updates the response deadline countdown in PLANILLA daily.
        Sends email alerts to project leaders/coordinators when the response deadline is 1 day or less.

    State Updates:
        Automatically marks emails as responded when replies are detected.
        Stops the countdown timer upon response.

    Informative Emails:
        Leaders/coordinators can manually mark emails as informative (no response needed).

🚨 Key Features

    Automated Tracking: No manual logging needed for authorized team members.
    Alerts: Ensures timely follow-ups with email notifications.
    Data Security: The system does not analyze email content; it only tracks correspondence status.

📍 Requirements

    Google Sheets files (HOJA INICIO and PLANILLA) must be located in a specified Google Drive folder for each project.
    GCP integration for automated workflows.

📄 Future Details

For detailed process flows and configuration guidelines, refer to the documentation in subsequent chapters.
