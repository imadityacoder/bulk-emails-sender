import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
import pandas as pd

my_email = "admoha.hub@gmail.com"
my_password = "aiug poje bfbt mzou"
my_name = "Admoha.com"

# Read the Excel file and extract email addresses
excel_file = 'bulk email 1.xlsx'  # Replace with your Excel file path
print("starting...")
df = pd.read_excel(excel_file)

# Assuming the email addresses are in a column named 'Email'
recipients = df['Email'].tolist()

# Setting up the SMTP server
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(my_email, my_password)
subject = "Explore the world of Information - Admoha.com!"
body = """
<!DOCTYPE html>
<html>
<head>
    <title>Discover the World of Information at Admoha.com!</title>
    <style>
        .email-container {
            border: 1px solid #ddd;
            border-radius: 10px;
            color:black;
            padding: 20px;
            font-family: Arial, sans-serif;
            max-width: 600px;
            margin: 0 auto;
            background-color: #f9f9f9;
        }
        .email-container a {
            color: #1a73e8;
            text-decoration: none;
        }
        .email-container ul {
            padding-left: 20px;
        }
        .email-container li {
            margin-bottom: 10px.
        }
        .email-footer {
            margin-top: 20px;
        }
        .logo {
            display: block;
            max-width: 100px;
            margin: 5px;
        }
    </style>
</head>
<body>
    <div class="email-container">
        <p>Dear Reader,</p>

        <p>I hope this email finds you well.</p>

        <p>My name is Vishal Kumar, and I am thrilled to introduce you to <a href="https://www.admoha.com">Admoha.com</a>, an exciting new blogging platform designed to bring you a wealth of information on a variety of topics.</p>

        <p>At Admoha.com, we are passionate about providing our readers with engaging, well-researched, and insightful content. Whether you are interested in technology, science, how to, programming, social science, or any other subject, our blog has something for everyone.</p>

        <p>Here’s why you should explore Admoha.com:</p>
        <ul>
            <li><strong>Wide Range of Topics</strong>: From the latest tech trends to social information, we cover a diverse array of subjects.</li>
            <li><strong>Expert Opinions</strong>: Our contributors are experts in their fields, offering unique insights and perspectives.</li>
            <li><strong>Engaging Content</strong>: We strive to present information in a way that is both informative and enjoyable to read.</li>
        </ul>

        <p>We would love for you to visit <a href="https://www.admoha.com">Admoha.com</a> and see what we have to offer. Your feedback and engagement are incredibly valuable to us as we continue to grow and improve our platform.</p>

        <p>Thank you for your time, and I look forward to welcoming you to our community of readers.</p>

        <p class="email-footer">
            Vishal Kumar<br>
            Founder, <a href="https://www.admoha.com">Admoha.com<br>
            <img src="https://www.admoha.com/media/icon/favicon.ico" alt="Admoha Logo" class="logo"></a><br>
            <a href="mailto:admoha.hub@gmail.com">admoha.hub@gmail.com</a><br>
            7033702514
        </p>
    </div>
</body>
</html>

"""

for recipient in recipients[:300]:
    msg = MIMEMultipart()
    msg['From'] = formataddr((my_name, my_email))
    msg['To'] = recipient
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'html'))

    # Send the email
    try:
        server.sendmail(msg['From'], msg['To'], msg.as_string())
        print(f"Email sent to {recipient} successfully!")
    except Exception as e:
        print(f"Failed to send email to {recipient}")

# Terminate the SMTP session
server.quit()
print('All emails are sent successfully!')
