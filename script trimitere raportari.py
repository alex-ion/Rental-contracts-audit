import csv, xlsxwriter, os, re, signature, smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders

luna=raw_input("Introdu luna pentru care se efectueaza auditul:\t\t")
zi_limita=raw_input("Introdu data pana la care sa primesti raspuns:\t\t")
optiune=raw_input("Introdu 1 pentru test sau 2 pentru trimiterea mailurilor pe puncte:\t\t")

adrese_email=[
{"Statie":"OTP","Email":""},
{"Statie":"HI2","Email":""},
{"Statie":"B2R","Email":""},
{"Statie":"CND","Email":""},
{"Statie":"CN0","Email":""},
{"Statie":"R2T","Email":""},
{"Statie":"R2Z","Email":""},
{"Statie":"IAS","Email":""},
{"Statie":"BCM","Email":""},
{"Statie":"R72","Email":""},
{"Statie":"C1J","Email":""},
{"Statie":"CLJ","Email":""},
{"Statie":"TGM","Email":""},
{"Statie":"TG2","Email":""},
{"Statie":"SBZ","Email":""},
{"Statie":"OMR","Email":""},
{"Statie":"TSR","Email":""},
{"Statie":"CRA","Email":""},
{"Statie":"KZ7","Email":""},
{"Statie":"BR8","Email":""},
{"Statie":"P36","Email":""},
{"Statie":"Kilometrii depasiti","Email":""}]

raportari_de_trimis = open('de trimis.csv', "rb")
reader1 = csv.reader(raportari_de_trimis, delimiter=",", quotechar='"', quoting=True)
lista_de_trimis=list(reader1)
lista_de_trimis.remove(lista_de_trimis[0])
raportari_de_trimis.close()
raportari=[]

class Raportari():
    def __init__(self, Contract, Remarks, Agent_out, Agent_in, Out_location, In_location, Raportare):
        global raportari
        self.Contract=Contract
        self.Remarks=Remarks
        self.Agent_out=Agent_out
        self.Agent_in=Agent_in
        self.Out_location=Out_location
        self.In_location=In_location
        self.Raportare=Raportare
        raportari.append(self)   

for linie in lista_de_trimis:
    obiect = len(globals())
    globals()[obiect] = Raportari(linie[0], linie[2], linie[3], linie[4], linie[5], linie[6], linie[7])

def creaza_csv(statie,optiune):
    global raportari
    fisiere_de_atasat=[]
    Row=0
    Column=0
    fisier_excel = statie+".xlsx"
    workbook=xlsxwriter.Workbook(fisier_excel)
    worksheet=workbook.add_worksheet()
    worksheet.write(Row,Column,"Contract")
    worksheet.write(Row,Column+1,"Remarks")
    worksheet.write(Row,Column+2,"Agent out")
    worksheet.write(Row,Column+3,"Agent in")   
    worksheet.write(Row,Column+4,"Out location")
    worksheet.write(Row,Column+5,"In location")
    worksheet.write(Row,Column+6,"Explicatii")
    Row+=1
    for element in raportari:
        if element.Raportare==statie:
            worksheet.write(Row,Column,element.Contract)
            worksheet.set_column('A:A', 10)
            if os.path.isfile((element.Contract+".jpg").lower()):
                fisiere_de_atasat.append(element.Contract+".jpg")
                worksheet.write(Row,Column+1,element.Remarks+" (fisier "+element.Contract+".jpg"+" atasat)")
                worksheet.set_column('B:B', 150)
            elif os.path.isfile((element.Contract+".pdf").lower()):
                fisiere_de_atasat.append(element.Contract+".pdf")
                worksheet.write(Row,Column+1,element.Remarks+" (fisier "+element.Contract+".pdf"+" atasat)")
                worksheet.set_column('B:B', 150)
            else:
                worksheet.write(Row,Column+1,element.Remarks)
                worksheet.set_column('B:B', 150)
            worksheet.write(Row,Column+2,element.Agent_out)
            worksheet.set_column('C:C', 30)
            worksheet.write(Row,Column+3,element.Agent_in)
            worksheet.set_column('D:D', 30)
            worksheet.write(Row,Column+4,element.Out_location)
            worksheet.set_column('E:E', 11)
            worksheet.write(Row,Column+5,element.In_location)
            worksheet.set_column('F:G', 11)
            Row+=1
    if Row>1:
        workbook.close()
        trimitere_email(statie,fisier_excel,fisiere_de_atasat,optiune)
        

def trimitere_email(statie,fisier_excel,fisiere_de_atasat,optiune):
    global luna, zi_limita
    fromaddr = "Alexandru Ion <alexandru.ion@avisbudget.ro>"
    for element in adrese_email:
        if element.get("Statie")==statie:
            if optiune=="test":
                toaddr= "28"+element.get("Email")
                toccaddr = ""
            elif optiune=="production":
                toaddr= element.get("Email")
                toccaddr = ""
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['CC'] = toccaddr
    msg['Subject'] = "Raport audit "+luna+" 2020 "+statie
    if statie=="Kilometrii depasiti":
        body="""Salut,<p>
            <p>
            Regasesti in atasament un tabel cu contractele pe care s-au facut mai mult de 500 km pe zi, pe luna din subiect.<p>
            <p>
            Cu stima,
            """
        body = body+signature.semnatura()
    else:
        body="""
            """
        body = body+signature.semnatura()
    msg.attach(MIMEText(body, 'html'))

    #aici adaugam fisierele cu dovezi (print-screenuri, contracte, scanuri)
    if fisiere_de_atasat:
        for fisier in fisiere_de_atasat:
            attachment = open(fisier, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload((attachment).read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', "attachment; filename= %s" % fisier)
            msg.attach(part)

    #aici adaugam fisierul excel cu raportarile
    attachment = open(fisier_excel, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % fisier_excel)
    msg.attach(part)            
    server = smtplib.SMTP_SSL('', 465)
    server.login("", "")
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr.split(",") + toccaddr.split(","), text)
    server.quit()
    print "S-a trimis mailul la "+statie


for statie in adrese_email:
    if optiune=="1":
        creaza_csv(statie.get("Statie"),"test")
    elif optiune=="2":
        creaza_csv(statie.get("Statie"),"production")
raw_input("Apasa <<enter>> pentru a iesi.")
