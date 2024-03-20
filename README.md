# Python-Email-Automation

The idea was to create a daily tasks spreadsheets, where the user checks a checkbox when the task is completed. At the end of the week, if all tasks were completed, the user receives an email with a special random gift. 

## Process

After creating the desired spread sheet using google sheets, the API was enabled using Google Cloud. After enabling the API, it was possible to download a jason file of the google sheets file. This file was then used in python. 

## Python 

Google Sheets authorization by using a pygsheets library and accessing the spreadsheet and worksheet:
```
import pygsheets
import smtplib
import random
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage

gc = pygsheets.authorize(service_file='/task_reminder_sheets.json')
print("Authorization successful!")
spreadsheet = gc.open('Task Rewards')
worksheet = spreadsheet.sheet1
```

Defined a list of row numbers that correspond to specific checkboxes in the Google Sheet and fetched cell values. These values represent whether the corresponding task is completed.

```
target_rows = [3, 4, 5, 6, 7, 8, 12, 13, 14, 15, 16, 17, 21, 22, 23, 24, 25, 26, 30, 31, 32, 33, 34, 35, 39, 40, 41, 42, 43, 44]

selected_values = []
for row_number in target_rows:
    cell_value = worksheet.cell(f'B{row_number}').value

selected_values.append(cell_value)
print(f"Values from specified rows in column B: {selected_values}")
```

Checked task completion: It is set to True if all the retrieved values are equal to 'True'.
```
all_tasks_completed = all(value.lower() == 'true' for value in selected_values)
```

Defined a list of reward messages that will be randomly selected. The element of surprise increases the will of completing the tasks.
```
reward_messages = [
    "Congratulations! You've earned a matcha coffee at Starbucks!",
    "Great job! Enjoy a relaxing evening with a movie ticket!",
    "Well done! Treat yourself to a nice dinner!",
    "Fantastic work! Go for a date night!",
    "Impressive! Take a break and indulge in a spa day!"]

selected_reward = random.choice(reward_messages)
```

Email configuration and message:
```
if all_tasks_completed:
    sender_email = 'email'
    sender_password = 'password'
    recipient_email = 'email'
    subject = 'Open to see your reward'
    body = f'{selected_reward}\n\nYou have completed all your tasks for this week.'

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    msg.attach(img)

    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
```

The code includes a try block to handle exeptions during the email sendinf process. If successful, it prints a success message; otherwise, any errors will be caught.
```
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
            print('Email sent successfully!')
    except Exception as e:
        print(f'Error sending email: {e}')
else:
    print('Not all tasks are completed. No email sent.')
```
