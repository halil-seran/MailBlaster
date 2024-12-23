MailBlaster -> FOR DEVELOPERS

FASTEST MAIL SENDER 

MailBlaster is a simple and efficient software that allows developers to send bulk emails securely via SMTP. It helps automate the process of sending emails to a large group, useful for marketing campaigns, newsletters, or any large-scale communication needs.

Features
	•	Send bulk emails to a specified range of email addresses from an Excel sheet.
	•	Customize the subject and body of the email.
	•	Secure SMTP email sending via Yandex (or your SMTP provider of choice).

Install required dependencies:
pip install pandas


How to Use MailBlaster
	1.	Prepare the Excel file:
	•	Create or download an Excel file (e.g., mailler.xlsx) containing the email addresses in the 8th column (column “I”).
	•	Ensure that your Excel file is properly formatted with valid email addresses.
	2.	Set up your SMTP credentials:
	•	Open the main.py file and enter your SMTP server details and credentials (e.g., for Yandex: SMTP server is smtp.yandex.com, and port is 587).
	•	Set your sender email and password in the script, or allow the script to accept them via GUI.
	3.	Running the Script:
	•	Execute the script using the Python interpreter:
 
 python main.py

	Only complete a few parameters, 

 Important Notes
	•	SMTP Limits: Make sure you are aware of any sending limits imposed by your email provider (e.g., Yandex). For Yandex, the daily limit is typically around 500 emails.
	•	File Structure: Make sure your email addresses are in the correct format in the Excel file, with no empty rows in between.
	•	Security: Ensure that you use app-specific passwords for SMTP connections, especially if you’re using your personal email address, to keep your main password secure.


 
