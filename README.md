# aws-bulk-iam

This script will read through an excel spreadsheet called `userlist.xlsx` containing two columns, name and email.

For each row in the spreadsheet, it will:
* create an iam user
* assign the user to a specified group
* create access credentials
* email the access key id to the associated email along with a random password
* add the secret to a text file, zip it up with the random random password