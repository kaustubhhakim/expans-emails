import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
from jinja2 import Environment, BaseLoader
import html2text
import datetime
from getpass import getpass
from habanero import cn

if __name__ == '__main__':

  # Google sheet details
  SHEET_ID = '1j-0B_rhVoIYdrKuq0abf4UWWFYSfonRSK-jGX2wF_TQ' 
  SHEET_NAME = 'Expanse Papers (Responses)'
  #url = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}'
  url = 'https://docs.google.com/spreadsheets/d/1j-0B_rhVoIYdrKuq0abf4UWWFYSfonRSK-jGX2wF_TQ/export?format=csv'
  df = pd.read_csv(url).dropna()
  df['Timestamp'] = pd.to_datetime(df['Timestamp'])

  # filter dataframe with for last n days
  df = df[df['Timestamp'] > datetime.datetime.now() - pd.to_timedelta("30day")]
  dois = df.values[:,1]

  # Retrieve bibliography using habanero via crossref api
  bibs = []
  for doi in dois:
    #bib = cn.content_negotiation(ids=doi, format='text', style='elsevier-harvard')
    bib = cn.content_negotiation(ids=doi, format='text', style='apa')
    bibs.append(bib)

  # Sender and receiver details
  sender_email = "kaustubh.hakim@kuleuven.be"
  receiver_email = "kaustubh.hakim@kuleuven.be"
  sender_change = input("Current sender: kaustubh.hakim@kuleuven.be. Do you want to change? (if yes press 'y' and enter, or else enter): ")
  if sender_change == 'y':
    sender_email = input("Type new sender email: ")
  receiver_change = input("Current receiver: kaustubh.hakim@kuleuven.be. Do you want to change? (if yes press 'y' and enter, or else enter): ") 
  if receiver_change == 'y':
    receiver_email = input("Type new receiver email: ")
     
  username = input("Type your u-number (e.g., u123456) and press enter: ")
  password = getpass(prompt="Type your password and press enter: ")

  # Create message
  message = MIMEMultipart("alternative")
  message["Subject"] = "Expanse Research Newsletter (test)"
  message["From"] = sender_email
  message["To"] = receiver_email

  # Create the HTML and plain-text version of the message
  html_in = """\
  <html>
    <head>
      <meta http-equiv="content-type" content="text/html; charset=UTF-8">
    </head>
    <body>
      <p>Dear Reader,</p>
      <p>The following exoplanet and planetary science papers were read/recently published at KU Leuven (Institute of Astronomy) and Royal Observatory of Belgium.</p>
      {% for bib in bibs %}
          <li>{{ bib }}</li>
      {% endfor %}
      <p></br></p>
      <p>Kind regards,</br>
      <a href="mailto:expanse-admin@kuleuven.be">expanse-admin@kuleuven.be</a><br>
      Expanse Admin</p>
    </body>
  </html>
  """

  # jinja2 rendering
  template = Environment(loader=BaseLoader).from_string(html_in)
  template_vars = {"bibs": bibs ,}
  html = template.render(template_vars)
  # Convert html to text
  text = html2text.html2text(html)
  # Turn these into plain/html MIMEText objects
  part1 = MIMEText(text, "plain")
  part2 = MIMEText(html, "html")

  # Add HTML/plain-text parts to MIMEMultipart message
  # The email client will try to render the last part first
  message.attach(part1)
  message.attach(part2)

  # Create secure connection with server and send email
  context = ssl.create_default_context()
  with smtplib.SMTP_SSL("smtps.kuleuven.be", 443, context=context) as server:
      server.login(username, password)
      server.sendmail(
          sender_email, receiver_email, message.as_string()
      )