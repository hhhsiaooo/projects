#!/usr/bin/env python
# coding: utf-8

# In[9]:


import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import ssl


# In[2]:


#
gmail_user = 'dwarfxiao@gmail.com'
pwd = 'l048o737v592e'
from_address = gmail_user
to_address = gmail_user
subject = 'COVID19 info from CSSE'
contents = 'FYI'
attachment = 'path\\covid19.xlsx'


# In[10]:


mail = MIMEMultipart()
mail['From'] = from_address
mail['To'] = to_address
mail['subject'] = subject

mail.attach(MIMEText(contents))

part = MIMEApplication(open('covid19.xlsx', 'rb').read())
part.add_header('Content-Disposition', 'attachment', filename='covid19.xlsx')
mail.attach(part)

s = smtplib.SMTP_SSL('smtp.gmail.com', 465)
s.ehlo()
s.login(gmail_user, pwd)
s.sendmail(from_address, to_address, mail.as_string())
s.quit()


# In[ ]:




