import imaplib, email
from email.header import decode_header
import os,sys
import setup as s



# TODO: Download the emails
con = imaplib.IMAP4_SSL("imap.gmail.com")
con.login(s.credential['email'],s.credential['password'])


status, data = con.select(mailbox='INBOX')  
# data refers to the number of emails contained in the mailbox
print(status, data)
print(f'There are {int(data[0])} messages in your INBOX')
N = 3

messages = int(data[0])

for i in range(messages, messages-N, -1):
    # fetching the email message by ID
    status, msg = con.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # parsing a bytes email into a message object
            msg = email.message_from_bytes(response[1])

            # decode the email subject
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                # if its a bytes, decode to str
                subject = subject.decode()
            # email sender
            from_ = msg.get("From")
            print("Subject:", subject)
            print("From:", from_)

            # if the email message is multipart
            if msg.is_multipart():
                # iterate over email parts
                for part in msg.walk():
                    # extract content type of email
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        # get the email body
                        body = part.get_payload(decode=True).decode()
                    except:
                        pass
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        # print text/plain emails and skip attachments
                        print(body)
                    elif "attachment" in content_disposition:
                        # download attachment:
                        filename = part.get_filename()
                        if filename:
                            if not os.path.isdir(subject):
                                # make a folder for this email (named after the subject)
                                os.mkdir(subject)
                            filepath = os.path.join(subject, filename)
                            # download attachment and save it
                            open(filepath, 'wb').write(part.get_payload(decode=True))
            else:
                # extract content type of email
                content_type = msg.get_content_type()
                # get the email body
                body = msg.get_payload(decode=True).decode()
                if content_type == "text/plain":
                    # print only text email parts
                    print(body)
                if content_type == "text/html":
                    # if it's HTMLL,create a new HTML and open it in browser
                    if not os.path.isdir(subject):
                        # make a folder for this email (named after the subject)
                        os.mkdir(subject)
                    filename = f"{subject[:50]}.html"
                    filepath = os.path.join(subject, filename)

                    # write the file
                    open(filepath, "w").write(body)
                    # open in the default browser
                    webbrowser.open(filepath)
                print("="*100)

con.close()
con.logout()


        



# also we can search by a specific criteria

# we can use fetch to retrieve the actual messages in the mailbox


# TODO: Extract the info from the email body
# TODO: Store info into a specific format :csv,excel?




