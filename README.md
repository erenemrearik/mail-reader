# Mail Reader

![.NET](https://img.shields.io/badge/.NET-512BD4?style=for-the-badge&logo=dotnet&logoColor=white)
![C#](https://img.shields.io/badge/C%23-239120?style=for-the-badge&logo=c-sharp&logoColor=white)
![MailKit](https://img.shields.io/badge/MailKit-FF6F61?style=for-the-badge)
![ClosedXML](https://img.shields.io/badge/ClosedXML-008000?style=for-the-badge)
![JSON](https://img.shields.io/badge/JSON-000000?style=for-the-badge&logo=json&logoColor=white)
![Multi-language](https://img.shields.io/badge/Language-TR%2FENG-yellowgreen?style=for-the-badge)


IMAP Mail Reader is a modular C# application that connects to an IMAP server to retrieve and process emails. The application features spam filtering, attachment handling (with special image processing), and daily Excel report generation with a modern, stylized layout. The project supports bilingual output (Turkish/English) for both console messages and Excel report headers.

## Features

- **IMAP Email Retrieval:**  
  Connects to an IMAP server, authenticates, and retrieves new emails.

- **Spam Filtering:**  
  Filters emails based on configurable spam keywords, blacklisted email addresses, and whitelisted domains defined in a JSON file.

- **Attachment Processing:**  
  Processes email attachments. Image attachments are renamed with a GUID to ensure uniqueness and prevent naming conflicts.

- **Excel Reporting:**  
  Generates daily Excel reports using ClosedXML with a modern, stylized design. Reports include email subject, sender, date, body, and attachment (image) details.

- **Bilingual Support:**  
  Supports both Turkish and English for console outputs and Excel headers. The language is selected at runtime.

## Prerequisites
.NET SDK (Version 5.0 or later)
- ``MailKit``

- ``Newtonsoft.Json``

- ``ClosedXML``

- ``Microsoft.Extensions.Configuration``

- ``Microsoft.Extensions.Configuration.Json``

## Configuration
Create an <strong>appsettings.json</strong> file in the root of your project with the following structure. Update the values according to your environment:
```json
{
  "EmailSettings": {
    "UserMail": "your-email@example.com",
    "AppPassword": "your-app-password",
    "ImapServer": "imap.example.com",
    "ImapPort": 993
  },
  "Paths": {
    "AttachmentPath": "C:\\Path\\To\\Attachments",
    "FiltersFilePath": "C:\\Path\\To\\filters.json",
    "ExcelOutputPath": "C:\\Path\\To\\ExcelReports"
  },
  "TimeSettings": {
    "WaitTime": 300000
  }
}
```

Create a <strong>filters.json</strong> file (the location should match the path specified in your appsettings.json) with your spam filtering configuration:
```json
{
  "spam_keywords": [
    "bedava",
    "kazan",
    "ödül",
    "promosyon",
    "ücretsiz"
  ],
  "blacklist_emails": [
    "spam@example.com",
    "badguy@malicious.com"
  ],
  "whitelist_domains": [
    "gmail.com",
    "company.com",
    "example.org"
  ]
}
```
### Acknowledgments
- [MailKit](https://github.com/jstedfast/MailKit) for robust email handling.
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) for easy Excel report generation.
- [Newtonsoft.Json](https://www.newtonsoft.com/json) for efficient JSON handling.
