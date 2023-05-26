# Email Sending Script
> This script sends an email with an attachment to a list of recipients. The list of recipients is read from an Excel file. The script uses Gmail's SMTP server to send the emails.

## Dependencies
 * Python 3.7 or higher
 * pandas
 * openpyxl
You can install these with pip:

```bash
python3 -m pip install pandas openpyxl
```

## Setup
1. Create an Excel file with one column, containing the email addresses of your recipients. Name this column 'Emails'. Save this file as 'emails.xlsx' in the same directory as the script.

1. Create a text file with the body of your email. Save this file as 'body.txt' in the same directory as the script.

1. Have the file you want to attach to the email in the same directory as the script. This file should be a PDF and named 'Resume.pdf'.

1. Replace the placeholders in the script with your actual Gmail address and application password.

## Usage
You can run the script from the command line with:

```bash
python3 email_script.py
```
