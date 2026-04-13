# Outlook_Automation
This repository contains a dynamic and fully automated solution built in Microsoft Excel using VBA for handling email operations via Outlook.

⚠️ Important: This solution is designed to work with Outlook Classic (Desktop Version). It does not support the New Outlook app, as macros and VBA integration are not available there.

The workbook is structured into two main functional modules:

**Mail Automation (Mailing)**

Enables automated email sending with customizable fields such as recipients, subject, body, and attachments. The solution is designed to be flexible and scalable, supporting dynamic data inputs directly from Excel.

Dynamic To, CC, BCC handling
Custom subject and email body
Attachment support
Loop-based bulk email sending
Fully customizable via Excel inputs


**Attachment Download System (Attachment Sheet)**

Implements a loop-based VBA mechanism to search, extract, and download email attachments from Outlook into a specified local directory. The process is fully dynamic, allowing bulk handling of emails based on defined criteria.

Search emails based on filters (subject, sender, date range)
Loop through emails and extract attachments
Save files automatically to a defined folder
Designed for bulk processing

**🧑‍💻 Requirements**
Microsoft Excel (.xlsm with macros enabled)
Microsoft Outlook Classic (Desktop App)
Macros enabled in Excel
Basic VBA knowledge (optional for customization)

**▶️ How to Use**
Open the Excel file
Enable macros
Fill in required inputs (email details or filters)
Run the macro from the respective sheet
Outlook Classic will be triggered automatically

