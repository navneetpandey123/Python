?import imaplib
import email

# Connect to an IMAP server
def connect(server, user, password):
    m = imaplib.IMAP4_SSL(server,993)
    print(m)
    print(user)
    print(password)
    m.login(user, password)
    m.select()
    return m

# Download all attachment files for a given email replace all special char with blank
def downloaAttachmentsInEmail(m, emailid, outputdir):
    resp, data = m.fetch(emailid, "(BODY.PEEK[])")
    email_body = data[0][1]
    #mail = email.message_from_string(email_body)
    mail = email.message_from_bytes(email_body)
    #print(mail)
    if mail.get_content_maintype() != 'multipart':
        return
    for part in mail.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None and part.get_filename() is not None:
            print('To:'+mail['To']+' Subject:'+mail['subject'].replace('#','').replace(':',''))
            print('From:'+mail['from']+' Att:'+part.get_filename()+' Date:'+mail['date'])
            open(outputdir + '/' +mail['From'].replace('>','').replace('<','').replace('.','').replace('"','')+'_'+part.get_filename().replace('\r','').replace('\n',''), 'wb').write(part.get_payload(decode=True))
            

# Download all the attachment files for all emails in the inbox.
def downloadAllAttachmentsInInbox(server, user, password, outputdir):
    m = connect(server, user, password)
    #resp, items = m.search(None, "(ALL)")
    resp, items = m.search(None, '(To "waterfallsupport@ivp.in" SINCE "01-Jan-2018")')
    #resp, items = m.search(None, '(Subject "Re: Sec Master Dashboard major issues")')
    items = items[0].split()
    for emailid in items:
        downloaAttachmentsInEmail(m, emailid, outputdir)

#MAIN
downloadAllAttachmentsInInbox("imap.gmail.com","Email@gmail.com","Password","D:\Files")        
