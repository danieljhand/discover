from __future__ import print_function
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from datetime import datetime, timedelta, date
from jinja2 import Environment, FileSystemLoader, select_autoescape
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from apiclient import errors
import base64
import csv
import mimetypes
import os

# Update as appropriate.
RECIPIENT_EMAIL = "recipient@domain.com"

# Update to the user authenticating against the gmail api
SENDER_EMAIL = "sender@domain.com"


def CreateDraft(service, user_id, message_body):
  """Create and insert a draft email. Print the returned draft's message and id.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message_body: The body of the email message, including headers.

  Returns:
    Draft object, including draft id and message meta data.
  """
  try:
    message = {'message': message_body}
    draft = service.users().drafts().create(userId=user_id, body=message).execute()

    print('Draft id: %s\nDraft message: %s' % (draft['id'], draft['message']))

    return draft
  except ValueError:
    print('An error occurred: %s' % ValueError)
    return None


def CreateMessage(sender, to, subject, message_text):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEText(message_text, 'html')
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject

  b64_bytes = base64.urlsafe_b64encode(message.as_bytes())
  b64_string = b64_bytes.decode()

  print(message)
  return {'raw': b64_string}


# define environment variables used by jinja
env = Environment(
    loader=FileSystemLoader('./templates/'),
    autoescape=select_autoescape(['html', 'xml']),
)

# load Jinja template
template = env.get_template('emailTemplate.html')

# define csv file to store historical data used to train the ML model
data_csv_file = 'data.csv'

# Google Drive API access
DRIVE_SCOPES = 'https://www.googleapis.com/auth/drive.metadata.readonly'
driveStore = file.Storage('drivecredentials.json')
driveCreds = driveStore.get()

if not driveCreds or driveCreds.invalid:
    driveFlow = client.flow_from_clientsecrets('client_secret.json', DRIVE_SCOPES)
    driveCreds = tools.run_flow(driveFlow, driveStore)

driveService = build('drive', 'v3', http=driveCreds.authorize(Http()))

# Get today's date and time
now = datetime.today() - timedelta(days=0)
nowStr = now.replace(microsecond=0).isoformat('T')

# Get date and time from the past (n days ago)
then = now - timedelta(days=1)
thenStr = then.replace(microsecond=0).isoformat('T')


# Construct a query string to be used as a search filter
queryString = "name contains 'Discovery' and modifiedTime > \'%s\' and modifiedTime <  \'%s\'"%(thenStr,nowStr) + " or ""name contains 'Design' and modifiedTime > \'%s\' and modifiedTime <  \'%s\'"%(thenStr,nowStr)

# Call the Drive v3 API
results = driveService.files().list(
    pageSize=25, fields="nextPageToken, files(id, name, webViewLink, owners, createdTime, modifiedTime)", q=queryString, orderBy='folder,modifiedTime').execute()
items = results.get('files', [])

emailMessageBody = ''
if not items:
    print('No files found.')
else:
    # Open csv file to append data_csv_file
    with open(data_csv_file, 'a') as csvfile:
        dataWriter = csv.writer(csvfile, delimiter=' ', quotechar='|', quoting=csv.QUOTE_MINIMAL)
        searchResults=[]
        for item in items:
            # TODO - Determine stage e.g. Discovery, Design etc using expert system or ML function
            deliveryStage = 'Design'

            # TODO - Determine consulting solution e.g. IMS, Labs, PaaS using expert system of ML function.
            consultingSolution = 'Infrastructure Migration Solution (IMS)'

            searchResults.append({'href':item['webViewLink'], 'name': item['name'], 'displayName': item['owners'][0]['displayName'],'emailAddress': item['owners'][0]['emailAddress'], 'solution':consultingSolution, 'stage':deliveryStage })

            # Append data to a file for the purpose of possible future ML training.
            # TODO Supervised learning labels are curretly missing.
            dataWriter.writerow([item['createdTime'], item['modifiedTime'], item['name'], item['owners'][0]['displayName'],item['owners'][0]['emailAddress'], item['webViewLink']])
        emailMessageBody = template.render(searchResults=searchResults)
    csvfile.close()

# Call the Gmail v3 API
MAIL_SCOPES = 'https://www.googleapis.com/auth/gmail.compose'
mailStore = file.Storage('mailcredentials.json')
mailCreds = mailStore.get()
if not mailCreds or mailCreds.invalid:
    mailFlow = client.flow_from_clientsecrets('client_secret.json', MAIL_SCOPES)
    mailCreds = tools.run_flow(mailFlow, mailStore)

# Create e-mail message body
mailService = build('gmail', 'v1', http=mailCreds.authorize(Http()))
mailMessage = CreateMessage(SENDER_EMAIL, RECIPIENT_EMAIL, 'Weekly Design and Discovery Digest - ' + now.strftime('%A %d %B %Y'), emailMessageBody)

# Create draft e-mail ready for review
CreateDraft(mailService, SENDER_EMAIL, mailMessage)
