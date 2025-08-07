# Outlook Email Routing via VBA and Excel Rules

This project provides a Microsoft Outlook VBA macro that automatically routes incoming emails to designated folders based on predefined keyword rules stored in an Excel file.

## 📦 Project Overview

When a new email arrives in Outlook, this macro:

- Reads subject keywords from the email
- Loads keyword rules from an Excel file (`Email_Routing_Rules.xlsx`)
- Matches keywords in the subject with rules (using multiple keywords)
- Moves the email to a specified Outlook folder dynamically

## 🗂️ Excel Rules Format

The macro reads from the following columns in the Excel file:

| Column A     | Column B     | Column C     | Column D       |
|--------------|--------------|--------------|----------------|
| Keyword1     | Keyword2     | Keyword3     | Target Folder  |

All keywords must appear in the subject for a match. If matched, the mail is moved to the folder specified in Column D.

## 📁 Folder Matching

Folders are searched dynamically and recursively in all Outlook accounts. Partial folder names like `GT13E2` and `GT13E2MU` are treated distinctly based on the full match.

## ✅ Requirements

- Microsoft Outlook (with VBA enabled)
- Microsoft Excel
- Windows OS

## 🛠️ Setup Instructions

1. Open Outlook > Press `ALT + F11` to open the VBA editor
2. Insert a new module and paste the code from `EmailRouting.bas`
3. Adjust the path of the Excel rules file:
   ```vba
   path = "C:\Users\YourUsername\Desktop\Email_Routing_Rules.xlsx"
   ```
4. Save and close the editor
5. Ensure the Excel file is **closed** when a new mail arrives (it will be opened in read-only mode)

## 🚀 Features

- Keyword-based email routing
- Multi-keyword support
- Dynamic folder searching
- Excel-based rule management
- Minimal user interaction required

## 📄 License

This project is open-source and licensed under the [MIT License](LICENSE).

## 🤝 Contributions

Feel free to fork the repository and submit pull requests with improvements or bug fixes.

## 👤 Author

Developed by [Hamed Mahmoudinia]

