#For email handling
import smtplib

#For getting the password from enviournment
import os

#For Not Displaying the text while taking input
import getpass

#For Efficient Message Passage
from email.message import EmailMessage

#For Sending Images
import imghdr

#For Csv Files
import csv

#Getting Email Id from User and Pass from Env Variable
print("\nWelcome to Automated Email System\n")

#Create a Server of User Choice
while True:

    print("List of Servers To Send Mail From :-\n")
    print("\n\t1. Gmail\n")
    print("\t2. Yahoo\n")
    print("\t3. Outlook\n")
    choice = str(input("\nEnter Your Choice (1/2/3) : "))

    if choice == "1":
        server = smtplib.SMTP("smtp.gmail.com",587)       
        print("\n")
        break
    elif choice == "2":
        server = smtplib.SMTP("smtp.mail.yahoo.com",587)
        print("\n")
        break
    elif choice == "3":
        server = smtplib.SMTP("smtp.office365.com",587)
        print("\n")
        break
    else:
        print("\nWrong Choice Choose Again\n")
        

server.starttls()

sender_mail = str(input("Enter Your Email Account : "))
print("\n")

while True:
    
    pass_choice = str(input("Does Your Email Has 2-Factor-Authentication Enabled (Y/N) : "))

    if pass_choice == "Y" or pass_choice == "y":
        print("\n\tTurn it Off and Turn On Less Secure App Access in your Email Provider Settings if it allows it \n\tOR\n\tAdd an App Password in your Email Provider Settings\n")
        break
    elif pass_choice == "N" or pass_choice == "n":
        print("\n\tTurn on Less Secure App Access in your Email Provider Settings if it Uses it.\n")
        break
    else:
        print("\nWrong Choice Try Again\n")

print("If you have Proceeded with the Above Instructions Correctly, You have to Write Email Password in the Later Case or App Password if you have Setup It\n")

while True:

    pass_choice = str(input("Do You Want To Type the Password (1) or Do You want to Take it From an Enviournment Varialbe (0) : "))

    if pass_choice == "1":
        sender_pass = str(getpass.getpass("\n\tEnter Your Password : "))
        print("\n")
        break
    elif pass_choice == "0":
        env_var = str(getpass.getpass("\n\tEnter Your Enviourment Variable Name : "))
        sender_pass = os.environ.get(env_var)

        while sender_pass == None:
            print("\nIncorrect Enviournment Variable ! Try Again\n")
            env_var = str(getpass.getpass("\n\tEnter Your Enviourment Variable Name : "))
            sender_pass = os.environ.get(env_var)

        print("\n")
        break
    else:
        print("\nWrong Choice, Try Again\n")

server.login(sender_mail,sender_pass)
print("Succesfully Logged In.\n")

#Asks for One or Multiple Recipients Choice
recipent_mail = []
while True:

    multiple = str(input("Do You want to Send Mail to One Person (0) or Multiple (1) : "))

    if multiple == "0":
        recipent_mail.append(str(input("\n\tEnter Recipient Email ID : ")))
        print("\n")
        break
    elif multiple == "1":
        print("\n\tNote That Format of CSV Files Should Be One Email on First Column of Every Row")
        file_name = str(input("\tEnter File Name to Read Multiple Email ID's From : "))
        print("\n")
        
        #FILE READING 
        with open(file_name,'r') as csv_file:
            csv_content = csv.reader(csv_file)
            for a_csv_content in csv_content:
                recipent_mail.append(a_csv_content[0])

        break
    else:
        print("\nWrong Choice Try Again\n")

#Inputs for the subject and body
msg = EmailMessage()
subject = str(input("Enter Subject of Email : "))
content = str(input("Enter Body of Email : "))
print("\n")

#IMAGE ATTACHMENT HERE
while True:
    
    choice = input("Do You Want to Attach Images (Y/N) : ")

    #taking input of images names
    if choice == "y" or choice == "Y":
        
        print("\n\tEnter Images Name One By One with its Format and Type STOP to Stop Attaching\n\tNote That Files Should be in Same Folder as this Script\n")

        capcha = "ASD"
        images = []

        while capcha != "STOP" and capcha != "stop":

            capcha = str(input("\t\t: "))

            if capcha == "STOP" or capcha == "stop":
                break

            images.append(capcha)

        break

    elif choice == "n" or choice == "N":
        break

    else:
        print("\nWrong Input Try Again\n")

print("\n")

#PDF ATTACHMENT HERE
while True:
    
    choice = input("Do You Want to Attach PDF (Y/N) : ")

    #taking input of images names
    if choice == "y" or choice == "Y":
        
        print("\n\tEnter PDF Names One By One with its Format and Type STOP to Stop Attaching\n\tNote That Files Should be in Same Folder as this Script\n")

        capcha = "ASD"
        pdf = []

        while capcha != "STOP" and capcha != "stop":

            capcha = str(input("\t\t: "))

            if capcha == "STOP" or capcha == "stop":
                break

            pdf.append(capcha)

        break

    elif choice == "n" or choice == "N":
        break

    else:
        print("\nWrong Input Try Again\n")

print("\n")

#Send the Mail/s
for one_rec in recipent_mail:
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender_mail
    msg['To'] = one_rec
    msg.set_content = content

    #attaching images one by one
    for one_image in images:
        with open(one_image,'rb') as img_file:
            file_data = img_file.read()
            file_type = imghdr.what(img_file.name)
            file_name = img_file.name
        msg.add_attachment(file_data, maintype='image', subtype=file_type, filename=file_name)

    #attaching pdf one by one
    for one_pdf in pdf:
        with open(one_pdf,'rb') as pdf_file:
            file_data = pdf_file.read()
            file_name = pdf_file.name
        msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

    server.send_message(msg)
    print("The Email has been Succesfully sent to " + one_rec + "\n")
    print("\n")

print("All Jobs are Done !!! Quitting\n")

#Quits the Server
server.quit()
