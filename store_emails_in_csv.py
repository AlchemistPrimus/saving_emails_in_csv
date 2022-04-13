"""This project contains a script to extract email from IMAP server

messages are written in a four column csv file.
The columns contains Date, from(sender), subject, message"""
import csv
import email
from email import policy
import imaplib
import logging
import os
import ssl
from bs4 import BeautifulSoup
from dotenv import load_dotenv

#loading environment variables
load_dotenv()

csv_path="mails.csv"

logger=logging.getLogger('imap_poller')
host=os.environ.get('HOST')
port=int(os.environ.get('PORT'))
ssl_context=ssl.create_default_context()

def connect_to_mailbox():
    #mail connection
    mail=imaplib.IMAP4_SSL(host,port,ssl_context=ssl_context)
    user=os.environ.get("USERNAME")
    pwd=os.environ.get("PASSWORD")
    mail.login(user,pwd)
    
    #get mailbox response and select a mail box
    status,messages=mail.select('INBOX')
    return mail,messages

#obtain plain text out of html mails
def get_email_text(email_body):
    soup=BeautifulSoup(email_body,'lxml')
    return soup.get_text(separator='\n',strip=True)

def write_to_csv(mail,writer,N,total_no_of_mails):
    for i in range(total_no_of_mails,total_no_of_mails-N,-1):
        res,data=mail.fetch(str(i),'RFC822')
        response=data[0]
        if isinstance(response,tuple):
            msg=email.message_from_bytes(response[1],policy=policy.default)
            
            #get header
            email_subject=msg['subject']
            email_from=msg['from']
            email_date=msg['date']
            email_text=""
            
            #if email is multipart
            
            if msg.is_multipart():
                #iterating over email parts
                for part in msg.walk():
                    #extract content type of email
                    content_type=part.get_content_type()
                    content_disposition=str(part.get('Content-Disposition'))
                    try:
                        #get email body
                        email_body=part.get_payload(decode=True)
                        if email_body:
                            email_text=get_email_text(email_body.decode('utf-8'))
                    except Exception as exc:
                        logger.warning("Caught exception %r",exc)
                        
                    if (content_type=='text/plain' and 'attachment' not in content_disposition):
                        #print plain text and skip attachments
                        #print(email_text)
                        pass
                    elif 'attachment' in content_disposition:
                        pass
            else:
                #extract content type of email
                content_type=msg.get_content_type()
                #get email email_body
                email_body=msg.get_payload(decode=True)
                if email_body:
                    email_text=get_email_text(email_body.decode('utf-8'))
                    
            if email_text is not None:
                #write data in csv
                row=[email_date,email_from,email_subject,email_text]
                writer.writerow(row)
            else:
                logger.warning('%s:%i: No message extracted',"INBOX",i)
                
def main():
    mail,messages=connect_to_mailbox()
    logging.basicConfig(level=logging.WARNING)
    total_no_of_mails=int(messages[0])
    #number of latest mail fetch
    #set it equal to total_number_of_emails to fetch all mail in the box
    N=2
    with open('csv_path','wt',encoding='utf-8',newline="") as fw:
        writer=csv.writer(fw)
        writer.writerow(["Date","From","Subject","Text Mail"])
        try:
            write_to_csv(mail,writer,N,total_no_of_mails)
        except Exception as exc:
            logger.warning("Caught exception: %r",exc)
            
if __name__=="__main__":
    main()