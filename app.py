from flask import Flask, render_template, request,redirect,url_for,flash,jsonify,send_file,send_from_directory
import subprocess

from fpdf import FPDF
from os import remove,getcwd

import win32com.client
import pythoncom
from dateutil.parser import *

import pandas as pd
from datetime import datetime




PATH_FILE = getcwd()+"/files/"

app = Flask(__name__)
listClients=[]
datalist = []
informacion =   {
        'numContrato' :  "",
        'strNombre' :  "",
        'strStatus' :  "",
        'strTipoPlan' :  "",
        'strCedula' :  ""
        }
datalist=[informacion]
subprocess.check_call(['pip', 'install', '--upgrade', 'pip'])



@app.route('/')
def index():
    return render_template('index.html', dicData=listClients,contrato=datalist)

@app.route('/calender', methods=['GET'])
def get_calender():
    from datetime import date
    outlook = win32com.client.Dispatch("Outlook.Application",pythoncom.CoInitialize()).GetNamespace("MAPI")
    calendar = outlook.GetDefaultFolder(9)
    appointments = calendar.Items
    appointments.Sort("[Start]")
    appointments.IncludeRecurrences = "True"


    dateFechaActual = date.today()
    intDiaSemana = dateFechaActual.weekday()
    counter = 0
    counter2 = 0
    while (intDiaSemana !=0 ):
        intDiaSemana = intDiaSemana-1
        counter=counter+1

    dateFechaActual = date.today()
    intDiaSemana = dateFechaActual.weekday()

    while (intDiaSemana !=4):
        intDiaSemana = intDiaSemana + 1
        counter2 = counter2 + 1



    begin = date.today()- pd.Timedelta(days=counter)

    end = date.today()+  pd.Timedelta(days=counter2)
    print(f"Activities from: {begin}, to: {end}")
    restriction = "[Start] >= '" + begin.strftime("%d/%m/%Y") + " 01:00 AM' AND [End] <= '" +end.strftime("%d/%m/%Y") + " 23:59 PM'"
    print("restriction:", restriction)
    restrictedItems = appointments.Restrict(restriction)


    apptDict = {}
    item = 0
    for indx, a in enumerate(restrictedItems):
            organizer = str(a.Organizer)
            meetingDate = str(a.Start)
            date = parse(meetingDate).date()
            subject = str(a.Subject)
            duration = str(a.duration)
            apptDict[item] = {"Duration": duration, "Organizer": organizer, "Subject": subject, "Date": date.strftime("%m/%d/%Y")}
            item = item + 1
            apt_df = pd.DataFrame.from_dict(apptDict, orient='index', columns=['Duration', 'Organizer', 'Subject', 'Date'])
            apt_df = apt_df.set_index('Date')
            apt_df['Meetings'] = apt_df[['Duration', 'Organizer', 'Subject']].agg(' | '.join, axis=1)
            grouped_apt_df = apt_df.groupby('Date').agg({'Meetings': ', '.join})
            grouped_apt_df.index = pd.to_datetime(grouped_apt_df.index)
            grouped_apt_df.sort_index()
            filename = 'calendario.csv'
            grouped_apt_df.to_csv(PATH_FILE+filename, index=True, header=True)
    return send_from_directory(PATH_FILE,path=filename,as_attachment=True)


if __name__== '__main__':
    app.run(debug=True)
