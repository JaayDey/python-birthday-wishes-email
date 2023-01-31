import pandas as pd
import datetime
import smtplib
from email.message import EmailMessage
import os


def sendEmail(to, sub, msg):
    print(f"email to {to} \nsend with subject: {sub}\n message: {msg}")
    email = EmailMessage()
    email["from"] = "Jeevan Devkota"
    email["to"] = f"{to}"
    email["subject"] = f"{sub}"
    email.set_content(f"{msg}")

    with smtplib.SMTP(host="smtp.gmail.com", port=587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login("jeevandevkota2020@gmail.com", "xlhmybutfecuodwj")
        smtp.send_message(email)
        print("Email Sent")
    pass


if __name__ == "__mail__":
    df = pd.read_excel("data.xlsx")

    today = datetime.datetime.now().strftime("%d-%m")

    update = []

    yearNow = datetime.datetime.now().strftime("%Y")

    for index, item in df.iterrows():
        print(index, item["Birthday"])
        bday = item["Birthday"].strftime("%d-%m")
        if (bday == today) and yearNow not in str(item["Year"]):
            sendEmail(item["Email"], "Happy Birthday " + item["Name"], item["Message"])
            update.append(index)
            print(update.append(index))

    for i in update:
        yr = df.loc[i, "Year"]

        df.loc[i, "Year"] = f"{yr}, {yearNow}"
    df.to_excel("data.xlsx", index=False)
