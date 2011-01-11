#!/usr/bin/python

import smtplib
import sys
import os
import logging.handlers
import time

from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE, formatdate
from email import Encoders


now = time.strftime("%Y-%m-%dT%H:%M:%S")

CURRENT_DIR = os.path.dirname(sys.argv[0])
LOG_DIR = CURRENT_DIR + 'logs/'

if os.path.exists(LOG_DIR) :
    pass
else :
    os.mkdir(LOG_DIR)
    
LOG_FILE = os.path.join(LOG_DIR, os.path.splitext(os.path.basename(sys.argv[0]))[0] + '.log')

my_logger = logging.getLogger('MyLogger')
my_logger.setLevel(logging.DEBUG)
handler = logging.handlers.RotatingFileHandler(LOG_FILE, maxBytes=5000, backupCount=5)
my_logger.addHandler(handler)



def send_mail(send_from, send_to, subject, text, files=[], server="smtp.google.com"):
    
  my_logger.info(now + ' [INFO] : Building new mail ')
    
  assert type(send_to) == list
  assert type(files) == list

  msg = MIMEMultipart()
  msg['From'] = send_from
  my_logger.info(now + ' [INFO] : From added ' + msg['From'])
  msg['To'] = COMMASPACE.join(send_to)
  my_logger.info(now + ' [INFO] : To added ' + msg['To'])
  msg['Date'] = formatdate(localtime=True)
  my_logger.info(now + ' [INFO] : date added ' + msg['Date']) 
  msg['Subject'] = subject
  my_logger.info(now + ' [INFO] : Subject added ' + msg['Subject'])  

  msg.attach(MIMEText(text))

  for f in files:
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(f, "rb").read())
    Encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
    msg.attach(part)
    my_logger.info(now + ' [INFO] : Attach ' + f + ' added')

  smtp = smtplib.SMTP(server)
  my_logger.info(now + ' [INFO] : connection opened')
  smtp.sendmail(send_from, send_to, msg.as_string())
  my_logger.info(now + ' [INFO] : Mail sent ')
  smtp.close()
  my_logger.info(now + ' [INFO] : connection closed')


