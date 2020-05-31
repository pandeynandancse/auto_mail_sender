from __future__ import division, print_function
# coding=utf-8

import numpy as np
# Flask utils
from flask import Flask, redirect, url_for, request, render_template

import pandas as pd 
import smtplib
import datetime

from openpyxl import load_workbook
# Define a flask app
app = Flask(__name__)

@app.route('/', methods=['GET'])
def index():
    # Main page
    return render_template('login.html',msg = "")


@app.route('/insert_dob', methods=['POST'])
def insert_dob():
    email = request.form["email"]
    dob = request.form["dob"]

    df = pd.read_excel("mails.xlsx")
    end = df.shape


    df.loc[len(df)] = [email,dob]    # append row in dataframe .important concept
    df.to_excel("mails.xlsx",header = True,index=False)
    return render_template('login.html',msg = "Successfully added")


GMAIL_ID = 'mail@gmail.com'
GMAIL_PSWD = 'mailpassword'





@app.route('/sender', methods=['GET','POST'])
def sender():
    msg=""
    return render_template('sender_registration.html',msg = msg)







@app.route('/sender_registration', methods=['POST'])
def sender_registration():
    email = request.form["email"]
    password = request.form["password"]

    df = pd.read_excel("users.xlsx")
    if "Unnamed: 0" in df.columns:
        df = df.drop('Unnamed: 0',axis=1)
    end = df.shape

    if(len(df)==1):
        msg = "One user is already there Do you want to remove that user?"
    else:
        df.loc[len(df)] = [email,password]    # append row in dataframe .important concept
        df.to_excel("users.xlsx",header = True,index=False)
        msg = "Registered succesfully, allow less secure apps also"
    return render_template('login.html',msg = msg)





@app.route('/delete_users', methods=['GET','POST'])
def delete_user_excel():
    

    df = pd.read_excel("users.xlsx")
    df = df[0:0]
    df.to_excel("users.xlsx",header = True,index=False)
    msg = "Successfully Deleted Existing User"
    return render_template('sender_registration.html',msg = msg)





@app.route('/delete_contacts', methods=['GET','POST'])
def delete_contacts_excel():
    

    df = pd.read_excel("mails.xlsx")
    df = df[0:0]
    df.to_excel("mails.xlsx",header = True,index=False)
    msg = "Successfully Deleted Contacts"
    return render_template('login.html',msg = msg)




def sendEmail(to,sub,msg):

    s = smtplib.SMTP('smtp.gmail.com',587) #port 587
    s.starttls()     #start tls service
    try:
        df = pd.read_excel("users.xlsx")
        global GMAIL_ID
        global GMAIL_PSWD

        GMAIL_ID =  list(df["Email"])[0]
        GMAIL_PSWD = list(df["Password"])[0]
        print(GMAIL_ID,GMAIL_PSWD,to)
        s.login(GMAIL_ID,GMAIL_PSWD)   #allow less secure apps for login
    except:
    	print("error")
    
    if len(to)>0:
        s.sendmail(GMAIL_ID,to,f'Subject: {sub}\n\n{msg}')
    else:
        msg = "No recipients exists"
        return render_template('login.html',msg=msg)
    s.quit()
    



@app.route('/send_mails', methods=['GET','POST'])
def send_mails():   
    df = pd.read_excel("mails.xlsx")
    email_names = list(df["Email"])
    sendEmail(email_names,"Regarding new info", 'Welcome Mr.')
    msg = "Message Sent Successfully"
    return render_template('login.html',msg = msg)

app.run(debug=True)
