from PIL import Image, ImageDraw, ImageFont
from xlrd import open_workbook
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
fromaddr = "sender_address" #Enter Sender Address here.
df = pd.read_excel('mydata.xlsx') #Please create an Excel file as shown in the excel file in the repository. Ensure that the same headers are used.
status=[]
Name = df['Name'].tolist()
College = df['College'].tolist()
mailId=df['mailID'].tolist()
for i in range(0,4):
    image = Image.open('template.jpg') #Template of the certifiacte which you are going to use.
    draw = ImageDraw.Draw(image)
    #Specifying the different font styles along with Size.
    newfont = ImageFont.truetype('Roboto-Italic.ttf', size=120)
    newFont = ImageFont.truetype('Roboto-Italic.ttf', size=90)
    #Before this you might want to change the coordinates where the data has to be entered. These coordinates work for my template.
    draw.text((1400,1300),Name[i], (0,0,+0), font=newfont) 
    draw.text((600,1475),College[i], (0,0,0), font=newFont)
    image.save(Name[i]+".jpg", resolution=100.0) #Here, Name[i].jpg will be the output image. Which we will be attaching withthe email ID.
    fromaddr = "sender_address" #Enter Sender Address here.
    toaddr = mailId[i]

    # instance of MIMEMultipart 
    msg = MIMEMultipart() 
    
    # storing the senders email address 
    msg['From'] = fromaddr 

    # storing the receivers email address 
    msg['To'] = toaddr 

    # storing the subject 
    msg['Subject'] = "Participation Certificate" #You can edit the subject here.

    # string to store the body of the mail 
    body = "Hello, Thank you for your active participation. Please find the attached Participation Certificate." #You can edit the body of the mail here.

    # attach the body with the msg instance 
    msg.attach(MIMEText(body, 'plain')) 

    # open the file to be sent 
    filename = Name[i]+".jpg"
    attachment = open(Name[i]+".jpg", "rb") 

    # instance of MIMEBase and named as p 
    p = MIMEBase('application', 'octet-stream') 

    # To change the payload into encoded form 
    p.set_payload((attachment).read()) 

    # encode into base64 
    encoders.encode_base64(p) 

    p.add_header('Content-Disposition', "attachment; filename= %s" % filename) #Adding the header here.

    # attach the instance 'p' to instance 'msg' 
    msg.attach(p) #Attaching the image to msg.

    # creates SMTP session 
    s = smtplib.SMTP('smtp.gmail.com', 587) #Depending on the mail server, you will have to change this. 587 is the port number for smtp.gmail.com

    # start TLS for security 
    s.starttls() 

    # Authentication 
    s.login(fromaddr, "Application_Password") # Generate an app password you have generated from Google. If you don't have one, visit 'https://bit.ly/googleapppasswords' to create one.
    
    # Converts the Multipart msg into a string 
    text = msg.as_string() 

    # sending the mail 
    s.sendmail(fromaddr, toaddr, text) 
    # terminating the session 
    s.quit() 
