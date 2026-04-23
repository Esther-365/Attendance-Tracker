# Attendance-Tracker
This project is an idea from [the GeeksforGeeks Python projects](https://www.geeksforgeeks.org/python/simple-attendance-tracker-using-python/). It is designed to alert the student and the professor about the student's attendance when nearing or exceeding the minimum leaves allowed. It uses the openpyxl library as it works with Excel, and the smtplib module to send emails.

How it works:
  It uses openpyxl to read the number of leaves the student has acquired for each course. It sends either a warning email or   an email to inform both the student and the professor that the student won’t be permitted to write the final exam given minimum allowed leaves are exceeded. I used [Mailtrap](https://mailtrap.io) to simulate the sending of emails.

```
The .env variables
HOST=sandbox.smtp.mailtrap.io 
PORT=587
MAIL_USERNAME=your_username
MAIL_PASSWORD=your_password
ATTENDANCE_PATH=path/to/your/file.xlsx
```
```
Libraries to install:
Openpyxl
Python-dotenv
pip install openpyxl python-dotenv
```
