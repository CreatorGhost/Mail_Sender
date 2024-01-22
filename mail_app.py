from smtplib import SMTP_SSL
from email.message import EmailMessage
import os
import json
from mimetypes import guess_type
import pandas as pd
from pretty_html_table import build_table

###################### IMoprtant Note   ######################
#     Make sure to turn on 2-factor authentication and get app passsword from Google its required in order to make this code work}


# put your email id and app password in the credentials.json file in current directory

file = open("credentials.json")

data = json.load(file)

SENDER_EMAIL = data["email"]
MAIL_PASSWORD = data["password"]

file_names = ["Reddit.jpg", "Archive.zip", "sample_data.xlsx"]

df = pd.read_excel("sample_data.xlsx")

name = "Aditya"
body = """
<html>
<head>
</head>

<body>
Hi {1},

<br><br>

Please find below the Report 
        {0}
</body>

</html>
""".format(
    build_table(
        df,
        "blue_light",
        width="auto",
        font_family="Open Sans",
        font_size="13px",
        text_align="justify",
    ),
    name,
)


def send_mail(
    SENDER_EMAIL, RECEIVER_EMAIL, MAIL_PASSWORD, subject, file_names, body
):
    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECEIVER_EMAIL
    # msg["Cc"] = CC_EMAIL
    # msg["Bcc"] = BCC_Email
    msg["Subject"] = subject
    msg.add_alternative(body, subtype="html")
    # for file_name in file_names:
    #     mime_type, encoding = guess_type(file_name)
    #     app_type, sub_type = mime_type.split("/")[0], mime_type.split("/")[1]

    #     with open(file_name, "rb") as file:
    #         file_data = file.read()
    #         msg.add_attachment(
    #             file_data, maintype=app_type, subtype=sub_type, filename=file_name
    #         )
    #         file.close()
    # Sending mail via SMTP server
    with SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(SENDER_EMAIL, MAIL_PASSWORD)
        smtp.send_message(msg)
        smtp.close()
    print("Mail Sent Sucessfully")


RECEIVER_EMAIL = "reacevermail@gmail.com"
subject = "Test Multiple  Attachment subject"
#file_names = list of file you want to send as attachement | remove it if you dont want to have any attachments
send_mail(
    SENDER_EMAIL, RECEIVER_EMAIL, MAIL_PASSWORD, subject, file_names, body
)
