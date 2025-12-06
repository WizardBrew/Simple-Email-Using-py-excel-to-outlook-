import pandas as pd
import win32com.client as win32

# Load Excel file
df = pd.read_excel("emails.xlsx")  # Columns: Email, Subject, Body

# Connect to Outlook
outlook = win32.Dispatch("Outlook.Application")

for _, row in df.iterrows():
    recipient = str(row["Email"]).strip()
    subject = str(row["Subject"]).strip()
    body = str(row["Body"]).strip()

    if not recipient:
        print("⚠️ Skipping empty recipient row.")
        continue

    # Create email
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.Body = body  # Use .HTMLBody for rich formatting
    mail.Send()       # Use .Display() to preview before sending

    print(f"✅ Sent to {recipient}")
