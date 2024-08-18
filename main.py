import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Load the Excel file
df = pd.read_excel('companies.xlsx')

# Email credentials (replace with your actual credentials or use environment variables)
your_email = "your_email@example.com"
your_password = "your_password"

# Set up the SMTP server (example for Gmail)
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(your_email, your_password)

for index, row in df.iterrows():
    company_name = row['Company Name']
    hr_email = row['HR Email']

    # Create the email content
    msg = MIMEMultipart()
    msg['From'] = your_email
    msg['To'] = hr_email
    msg['Subject'] = "Exploring Job Opportunities for Data Analyst/Machine Learning Enthusiast"

    body = f"""
Dear HR Team,

I hope this message finds you well.

My name is [Your Name], and I am interested in exploring job opportunities at {company_name}. I am enthusiastic about the possibility of contributing to your team and would love to discuss how my skills and experience align with your needs.

Thank you for considering my application. I look forward to the opportunity to discuss how I can contribute to the success of {company_name}.

Best regards,

[Your Name]  
[Your Email]  
LinkedIn: [Your LinkedIn Profile]  
[Your Phone Number]  
[Your Location]

    """

    msg.attach(MIMEText(body, 'plain'))

    # Send the email
    server.send_message(msg)
    print(f"Email sent to {company_name} at {hr_email}")

    # Clear the message after sending
    msg.clear()

# Disconnect from the server
server.quit()
