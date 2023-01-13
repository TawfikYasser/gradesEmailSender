# Not Production Code.
# Contact for more info.

# We need to use the redmail package
# !sudo pip install redmail
# !sudo pip install openpyxl

# Some imports
import pandas as pd
from redmail import outlook

# outlook username and password of the sender
outlook.username = ""
outlook.password = ""


# Read the sheet of grades using pandas
df = pd.read_csv('')

# Iterate over the list of students to get each student's mail, 
# then extract the student's grades and create a csv file from the grades and send them finally.

for row in df.values:
    # A student record
    studentData = pd.DataFrame([row], columns=list(df.columns))    

    # Email Content Message
    emailContent =""

    # Sending the email 
    outlook.send(receivers=[str(row[3])], 
    subject="SUBJECT OF THE EMAIL", text=emailContent,
    attachments={fn: studentData})

# This script was created by [Tawfik Yasser]: tyasser@nu.edu.eg
# Takes about 1 min
