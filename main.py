from flask import Flask, render_template, request,redirect, url_for
from werkzeug.utils import secure_filename

import os, smtplib, ssl
import pandas as pd
from email.message import EmailMessage
from email.utils import make_msgid

#get env variables from local env file
from dotenv import load_dotenv
load_dotenv()


app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv("PASSWORD")

@app.route('/')

@app.route('/home', methods =["POST", "GET"])
def home():
    if request.method == 'POST':
        #Data from the form is inputtted into the variables
        subject = request.form.get('subject')
        name = request.form.get('sendername')
        sheetname =  request.form.get('sheetname')
        mailcolumn =  request.form.get('mailcolumn')
        comapnycolumn = request.form.get('companyname')

        #save excel files to the mailData folder
        excel =  request.files['xlfile']
        excelpath = os.path.join(os.path.abspath(os.path.dirname(__file__)), secure_filename(excel.filename))
        excel.save(excelpath)

        #save pdf files to the mailData folder
        pdf =  request.files['pdffile']
        pdfpath = os.path.join(os.path.abspath(os.path.dirname(__file__)), secure_filename(pdf.filename))
        pdf.save(pdfpath)

        return credentials(subject,name,sheetname,mailcolumn,comapnycolumn,pdfpath,excelpath,pdf.filename) #redirect(url_for('credentials',subject=subject,name=name,sheetname=sheetname,mailcolumn=mailcolumn,comapnycolumn=companycolumn,pdfpath=pdfpath,excelpath=excelpath, filename= pdf.filename))
    else:
        return render_template('index.html')

@app.route('/send') #<subject>,<name>,<sheetname>,<mailcolumn>,<comapnycolumn>,<pdfpath>,<excelpath>,<filename>
def credentials(subject,name,sheetname,mailcolumn,comapnycolumn,pdfpath,excelpath,filename):
    #assign env variables to python variables
    mail = os.getenv("EMAIL")
    password = os.getenv("PASSWORD")

    #email credential details from frontend
    security = ssl.create_default_context()
    Subject = subject
    senderName = name
    sheetName = sheetname
    recipientmailcolumn = mailcolumn
    clientNamecolumn = comapnycolumn

    #extract values from sheet
    emails = pd.read_excel(excelpath,sheet_name=sheetName)
    recipient = emails[recipientmailcolumn].values
    clientname = emails[clientNamecolumn].values
    

    def addattachment():
        with open (pdfpath, 'rb') as f:
            file_data = f.read()
            file_name = filename
            f.close()
        return file_data,file_name

    def emailcredentials(x,y):
        asparagus_cid = make_msgid()
        msg = EmailMessage() 
        msg['Subject'] = Subject
        msg['From'] = 'Iselen Triumph'
        msg['To'] = x
        #g['Bcc'] = 'customercare@rapidvigilsecurity.com'
        #msg.set_content(content.format(name=senderName, company=y))
        msg.add_alternative("""\
            <html lang="en">
            <head>
            </head>
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Poppins&family=Roboto&display=swap');
            </style>
            <body>
                <h3 style="font-family: 'Roboto', sans-serif; color: rgb(59,71,157); text-transform: capitalize;">Goodday, </h3>
                <p style="font-family: 'Poppins', sans-serif;">My name is <b style="font-family: 'Roboto', sans-serif; color: rgb(59,71,157); text-transform: capitalize;">""" +str(senderName)+ """</b>, I am contacting you on behalf of Rapid Vigil Security Company Limited. <br><br>
                    Rapid Vigil is a fast growing security company, which provides safety and security across physical and digital environments. Our mission is to secure the environment (your environment) for people to create wealth without fear or loss. To achieve this we provide Guard services, Health, Safety & Environment (HSE) services, CCTV and Access Control Systems installation and maintenance, as well as so many other security measures in line with the best requirement of our clients.  <br><br>
                    Rapid Vigil Security also offers CCTV Remote Monitoring services, which comes on a lease basis. Our services cut across employee management and time attendance technology. <br><br>
                    We would like to propose our guard services to <b style="font-family: 'Roboto', sans-serif; color: rgb(59,71,157); text-transform: capitalize;">""" +str(y)+ """</b>. <br><br>
                    We would appreciate it if you could grant us a meeting at a mutually convenient time to discuss how we can work together to protect your assets and that of your clients. <br> 	<br>
                    A feedback on the day and time would be highly appreciated. <br><br>
                    Attached to this email is a proposal document, as well as the various security services we provide. <br><br> 
                    Thank you for your time and consideration. We look forward to meeting and working with you soon. <br><br>
                </p>
                <b>
                <p>Regards,</p>
                <p>Triumph Iselen,</p>
                <p>B2C Sales Assistant</p><br>
                <p>Rapid Vigil Security Co. Ltd.<br>
                10 Billy Odumala Street<br>
                Julie Estate, Oregun<br>
                Ikeja<br>
                Lagos</p>



                <p>Tel: +234 807 050 0012<br>
                Tel: +234 807 050 0285<br>
                Tel: +234 807 050 0096<br>
                E-mail: triumph.iselen@rapidvigilsecurity.com<br>
                Website: www.rapidvigilsecurity.com<br></p>


                <p>Stay in contact with us ...<br>
                <a href="https://www.facebook.com/RAPIDVIGIL">Facebook</a> | <a href="https://www.instagram.com/rapidvigil/">Instagram</a> | <a href="https://www.pinterest.com/rapidvigil/rapid-vigil-security/">Pinterest</a></p></b>
            </body>
            </html>
        """.format(asparagus_cid=asparagus_cid[1:-1]), subtype='html')
        msg.add_attachment(addattachment()[0], maintype='application', subtype='octet-stream', filename=addattachment()[1])
        return msg

    #mailing logic
    def sendmail():
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=security) as smtp:

            smtp.login(mail, password) #logs in with mail and password
            
            #inputs the cell value into the respective variables
            x = 0
            for emails in recipient:
                if (not recipient[x]) or ((clientname[x] == '') or (clientname[x] == 0)):
                    render_template('errors.html', content='There was an error')
                else:
                    name = str.capitalize(clientname[x])
                    smtp.send_message(emailcredentials(emails, name))
                    x += 1
            smtp.quit()

        os.remove(pdfpath)
        os.remove(excelpath)

    sendmail()

    return render_template('sent.html')
    

if __name__ == '__main__':
    app.run(debug=os.getenv("DEBUG"))
