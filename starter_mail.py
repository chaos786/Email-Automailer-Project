import smtplib, sys     # For delivering mails
import pandas as pd     # To read the excel file and for doing manipulations to dataframe
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def automater(email, password, current_month, date, subj, excel_file_location):
    """ Email automater script to send emails to in bulk.
        Parameters:  
            email -> <your email-id>
            password -> <password>
            current_month -> <template specific>
            date -> <template specific>
            subj -> <template specific>
            excel_file_location -> Source of the data
        Returns: None <sends mails to the respective tos and ccs>"""
    
    #me is the variable for the sender's email address
    me = email
    
    #now we import the html files that make up the contents of the mail
    with open('html_template.html') as f: header = f.read()
    with open('footer_template.html') as f: footer = f.read()

    #now we import the data from the excel file that gets selected for the email IDs and other data
    df_mailto = pd.read_excel(excel_file_location, sheet_name=0)
    #now we replace NaN values with actual blanks
    df_mailto[df_mailto.isnull()]=''
    #now we get different variables from the excel file into the program which will later be used for making the emails seem more personal
    names = df_mailto['Full Name'].tolist()
    emails = df_mailto['Email'].tolist()
    cities = df_mailto['City'].tolist()
    ccs = [[i] for i in df_mailto['Alternate Email ID'].tolist()]
    #we print some of the data we extracted from the excel file
    print(emails)
    print(ccs)
    
    msgno = 0
    #now we iterate through the data for every receiver
    for name, emailid, city in zip(names, emails,cities):
        #es has the email IDs of the receiver in a list format
        es =[emailid]+ccs[msgno]
        #page gets customized for the specific receiver
        page = header.format(name, current_month, city, date)
        msg = MIMEMultipart('alternative')
        
        msg['Subject'] = subj
        msg['From'] = me
        msg['To'] = emailid
        msg['Cc'] = ", ".join(ccs[msgno])
        
        page = page + footer
        msg.attach(MIMEText(page, 'html'))
        
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        
        s.login(email, password)
        s.sendmail(me, es, msg.as_string())
        s.quit()
        print('Message {} sent to {}.'.format(msgno, emailid))
 
        msgno += 1