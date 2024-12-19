MailBlaster -> FOR DEVELOPERS

MailBlaster is a simple and efficient tool that allows developers to send bulk emails securely via SMTP. It helps automate the process of sending emails to a large group, useful for marketing campaigns, newsletters, or any large-scale communication needs.

Features
	•	Send bulk emails to a specified range of email addresses from an Excel sheet.
	•	Customize the subject and body of the email.
	•	Secure SMTP email sending via Yandex (or your SMTP provider of choice).
	•	Simple GUI to easily input ranges and email content.

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

 Using the GUI to Send Emails:
	•	The GUI will prompt you to input:
	•	Start Index: The starting email address number from your list (e.g., 1 for the first email).
	•	End Index: The last email address number in your list (e.g., 100 if you want to send to the first 100 addresses).
	•	Subject: The subject of the email.
	•	Body: The body content of the email.
	•	After entering these details, click “Send”, and the program will send the email range to each specified address.
	5.	Check the Results:
	•	After sending emails, the application will notify you when the emails are successfully sent, or if there were any errors.


 Important Notes
	•	SMTP Limits: Make sure you are aware of any sending limits imposed by your email provider (e.g., Yandex). For Yandex, the daily limit is typically around 500 emails.
	•	File Structure: Make sure your email addresses are in the correct format in the Excel file, with no empty rows in between.
	•	Security: Ensure that you use app-specific passwords for SMTP connections, especially if you’re using your personal email address, to keep your main password secure.


 
