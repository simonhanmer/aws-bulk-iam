#! /usr/bin/env python3
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd
import secrets
import os
import pyminizip
import boto3
from botocore.exceptions import ClientError
import settings


excelFile = "./userlist.xlsx"

mailserver = smtplib.SMTP('smtp.office365.com', 587)
mailserver.ehlo()
mailserver.starttls()
mailserver.login(settings.SMTP_USER, settings.SMTP_PASS)


def processExcelFile():
    try:
        excel = pd.read_excel(excelFile)
    except BaseException as e:
        print(f"{format(e)}")

    for idx, column in excel.loc[:, ['name', 'email']].iterrows():
        processUser(column)


def processUser(data):
    name = data['name'].strip()
    email = data['email']
    print(f"Processing {name:>20}: ", end='')
    info = createIAM(name)
    print(f" IAM: {info[0]}")
    if info[0] == "OK":
        emailInfo(name, email, info)


def createIAM(name):
    hyphenName = f"{settings.IAM_GROUP}-{name.replace(' ', '-')}".lower()
    firstName = name.split(' ')[0]

    iam = boto3.client(
            'iam',
            aws_access_key_id=settings.AWS_ACCESS_ID,
            aws_secret_access_key=settings.AWS_ACCESS_SECRET,
            region_name='eu-west-2')

    try:
        iam.create_user(UserName=hyphenName)
    except ClientError as e:
        return ['FAIL', e]

    try:
        iam.add_user_to_group(UserName=hyphenName, GroupName=settings.IAM_GROUP)
    except ClientError as e:
        return ['FAIL', e]

    try:
        response = iam.create_access_key(UserName=hyphenName)
    except ClientError as e:
        return ['FAIL', e]

    return['OK', response['AccessKey']['AccessKeyId'], response['AccessKey']['SecretAccessKey']]


def emailInfo(name, email, info):
    # Generate zipfile
    zipPass = secrets.token_hex(6)

    txtFile = "secret.txt"
    zipfile = "info.zip"
    with open(txtFile, "w") as s:
        s.write(f"Secret Access Key: {info[2]}\n")
        s.close()
        pyminizip.compress(txtFile, None, zipfile, zipPass, 0)
        os.remove(txtFile)

    from_email = "simon.hanmer@ecs.co.uk"
    to_email = email
    msg = MIMEMultipart()
    msg['Subject'] = 'Test Email 1'
    msg['From'] = from_email
    msg['To'] = to_email
    msgText = MIMEText(
        f"Hi {name},<br />\n"
        f"we've created an AWS IAM user to use with the Academy.<br /><br />\n"
        f"The Access Key ID for the user is {info[1]}.<br /><br />\n"
        f"We'll send the secret in a zip file, the password will be {zipPass}<br />\n",
        'html')
    msg.attach(msgText)
    mailserver.sendmail(from_email, to_email, msg.as_string())

    msg = MIMEMultipart()
    msg['Subject'] = 'Test Email 2'
    msg['From'] = from_email
    msg['To'] = to_email
    msgText = MIMEText(
        f"Hi {name},<br />\n"
        f"we've created an AWS IAM user to use with the Academy.<br /><br />\n"
        f"Here's the zip file<br />"
        f"We sent the access ID and password for the zip file in a separate email<br />\n",
        'html')
    msg.attach(msgText)
    zip = MIMEApplication(open(zipfile, 'rb').read())
    zip.add_header('Content-Disposition', 'attachment', filename=zipfile)
    msg.attach(zip)

    mailserver.sendmail(from_email, to_email, msg.as_string())
    os.remove(zipfile)


if __name__ == "__main__":
    processExcelFile()
