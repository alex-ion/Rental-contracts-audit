import csv, time, xlsxwriter, codecs
from datetime import datetime
from dateutil import relativedelta

import MySQLdb

db = MySQLdb.connect(host="127.0.0.1",
                     user="root",
                     passwd="",
                     db="")

primul_timp=time.time()
contracte_importate=[]
contracte_scrive=[]
tari_importate=[]
coduri_agenti=[]
masini_importate=[]
ajustari=[]
audit=[]
rezervari_importate=[]
raportari=[]
lista_coduri_tari=[]
lista_coduri_tari_uzuale=["AU","AT","BE","BG","BR","CA","CN","EE","SK","CS","HR","DK","FI","FR","DE","GR","HU","IE","IL","IT","JP","LU","LT","MT","MX","MD","NL","NZ","GB","NO","PL","PT","RO","RU","RS","QV","QS","QC","ES","SE","CH","TR","UA","GB","UE","US"]
lista_coduri_IDP=["AL","AM","AT","AZ","BS","BH","BA","BE","QB","BR","BG","CF","CG","HR","CU","CY","CS","DK","EE","FI","FR","GE","IE","DE","GR","GY","HU","IS","IR","IL","IT","CI","KZ","KE","KW","KZ","LV","LR","LI","LT","LU","QM","MT","MD","MC","MO","ME","MA","MK","NL","NR","NO","PK","PE","PH","PL","PT","QA","RO","RU","QC","SM","SN","RS","SC","QV","QS","ZA","ES","SE","CH","*","TN","TR","TM","UA","GB","UE","UY","UZ","VN","KF"]
grupe_sub_21_ani=["A","B"]
grupe_sub_25_ani=["A","B","C","J"]
grupe_non_TEF=["G","H","O","P","N","L"]
statii_preferred=["OTP","HI2","CLJ","SBZ","TSR"]

fisier_contracte = open('contracte din platforma.csv', "rb")
reader2 = csv.reader(fisier_contracte, delimiter=",", quotechar='"', quoting=True)
lista_contracte=list(reader2)
lista_contracte.remove(lista_contracte[0])
fisier_contracte.close()

fisier_scrive=open('contracte din Scrive.csv',"rb")
reader1 = csv.reader(fisier_scrive, delimiter=",", quotechar='"', quoting=True)
lista_scrive = list(reader1)
lista_scrive.remove(lista_scrive[0])
fisier_scrive.close()

fisier_tari = open('lista tari.csv', "rb")
reader3 = csv.reader(fisier_tari, delimiter=",", quotechar='"', quoting=True)
lista_tari=list(reader3)
fisier_tari.close()

fisier_agenti = open('coduri agenti.csv', "rb")
reader4 = csv.reader(fisier_agenti, delimiter=",", quotechar='"', quoting=True)
lista_agenti=list(reader4)
lista_agenti.remove(lista_agenti[0])
fisier_agenti.close()

fisier_masini = open('masini.csv', "rb")
reader5 = csv.reader(fisier_masini, delimiter=",", quotechar='"', quoting=True)
lista_masini=list(reader5)
lista_masini.remove(lista_masini[0])
fisier_masini.close()

fisier_ajustari = open('ajustari.csv', "rb")
reader6 = csv.reader(fisier_ajustari, delimiter=",", quotechar='"', quoting=True)
lista_ajustari=list(reader6)
lista_ajustari.remove(lista_ajustari[0])
fisier_ajustari.close()

fisier_audit = open('audit.csv', "rb")
reader7 = csv.reader(fisier_audit, delimiter=",", quotechar='"', quoting=True)
lista_audit=list(reader7)
fisier_audit.close()

fisier_rezervari = open('rezervari.csv', "rb")
reader8 = csv.reader(fisier_rezervari, delimiter=",", quotechar='"', quoting=True)
lista_rezervari=list(reader8)
lista_rezervari.remove(lista_rezervari[0])
fisier_rezervari.close()
    

class Contracte():
    def __init__(self,Contract,Brand,Status,LOR,Rezervare,mva,license_No,Customer_Name,Date_of_birth,out_Date,entry_out_date, in_Date, entry_in_date, out_LocT,in_LocT,out_Km,in_Km,Km_drive,credit_card,awd1,mva_repl,grp_Ctr,grp_Crg,sell_status,cnt_prod_01,cnt_prod_02,cnt_prod_03,cnt_prod_04,cnt_prod_05,cnt_prod_06,cnt_prod_07,cnt_prod_08,cnt_prod_09,cnt_prod_10,tarif_1_out, tarif_1_in, agent_out,agent_in,Renter_ADDR1,Renter_ADDR2,Renter_ADDR3, total_charge, net_charges, cdw_Amt, pai_Amt, li_Amt, Remarks, Company_Name, grp_Res):
        global contracte_importate
        self.Contract=Contract
        if Brand=="A":
            self.Brand="Avis"
        elif Brand=="B":
            self.Brand="Budget"
        self.Status=Status
        self.LOR=LOR
        if Rezervare=="00000000  0":
            self.Rezervare=""
        else:
            self.Rezervare=Rezervare
        self.mva=mva
        self.license_No=license_No
        self.Customer_Name=Customer_Name
        self.Date_of_birth=Date_of_birth[0:10].replace("-","/")
        self.out_Date=out_Date.replace("-","/")
        self.entry_out_date=entry_out_date.replace("-","/")
        self.in_Date=in_Date.replace("-","/")
        self.entry_in_date=entry_in_date.replace("-","/")
        self.out_LocT=out_LocT
        self.in_LocT=in_LocT
        self.out_Km=out_Km
        self.in_Km=in_Km
        self.Km_drive=Km_drive
        self.credit_card=credit_card
        self.awd1=awd1
        self.mva_repl=mva_repl
        self.grp_Ctr=grp_Ctr
        self.grp_Crg=grp_Crg
        self.sell_status=sell_status
        self.cnt_prod_01=cnt_prod_01
        self.cnt_prod_02=cnt_prod_02
        self.cnt_prod_03=cnt_prod_03
        self.cnt_prod_04=cnt_prod_04
        self.cnt_prod_05=cnt_prod_05
        self.cnt_prod_06=cnt_prod_06
        self.cnt_prod_07=cnt_prod_07
        self.cnt_prod_08=cnt_prod_08
        self.cnt_prod_09=cnt_prod_09
        self.cnt_prod_10=cnt_prod_10
        self.produse=""
        for element in [cnt_prod_01, cnt_prod_02, cnt_prod_03, cnt_prod_04, cnt_prod_05, cnt_prod_06, cnt_prod_07, cnt_prod_08, cnt_prod_09, cnt_prod_10]:
            if element !="":
                self.produse+=element
        self.tarif_1_out=tarif_1_out
        self.tarif_1_in=tarif_1_in
        for i in coduri_agenti:
            if agent_out==i.cod_agent:
                agent_out=i.nume
            if agent_in==i.cod_agent:
                agent_in=i.nume
        self.agent_out=agent_out
        self.agent_in=agent_in
        self.Renter_ADDR1=Renter_ADDR1
        self.Renter_ADDR2=Renter_ADDR2
        self.Renter_ADDR3=Renter_ADDR3
        self.total_charge=total_charge
        self.net_charges=net_charges
        self.cdw_Amt=cdw_Amt
        self.pai_Amt=pai_Amt
        self.li_Amt=li_Amt
        self.Remarks=Remarks.upper()
        self.cancelled=False
        self.Company_Name=Company_Name
        self.grp_Res=grp_Res
        if net_charges=="0" and Status=="closed":
            self.cancelled=True
        contracte_importate.append(self)

class Scrive():
    def __init__(self,Status,Author,Party_role,Party_name,RA,driverslicense,drivinglicense,registration, party_signed_document, Brand):
        global contracte_scrive
        self.Status=Status
        self.Author=Author
        self.Party_role=Party_role
        self.Party_name=Party_name.upper()
        self.RA=RA
        self.drivinglicense=drivinglicense.upper()
        if drivinglicense=="":
            self.driverslicense=driverslicense.upper() #folosit atunci cand drivinglicense este gol
        self.registration=registration[3:len(registration)]
        temp=Party_name.split(",")
        self.nume=temp[0].upper()
        self.prenume=temp[1].strip(" ").upper()
        self.party_signed_document=party_signed_document
        self.Brand=Brand
        contracte_scrive.append(self)

class Tari():
    def __init__(self,cod_tara,nume_tara):
        global tari_importate,lista_coduri_tari
        self.cod_tara=cod_tara
        self.nume_tara=nume_tara
        tari_importate.append(self)
        lista_coduri_tari.append(self.cod_tara)

class Agenti():
    def __init__(self, prenume, nume, cod_agent):
        global coduri_agenti
        self.nume=str(prenume)+" "+str(nume)
        self.cod_agent=cod_agent
        coduri_agenti.append(self)

class Masini():
    def __init__(self, mva, carburant):
        global masini_importate
        self.mva=mva.rjust(8,"0")
        self.carburant=carburant
        masini_importate.append(self)

class Ajustari():
    def __init__(self, Contract, cod_ajustare, valoare, Road_Exp):
        global ajustari
        self.Contract=Contract
        self.cod_ajustare=cod_ajustare
        self.Road_Exp=Road_Exp
        if float(valoare)>=1:
            self.valoare=valoare
            ajustari.append(self)

class Audit():
    def __init__(self, Contract, remarks):
        global audit
        self.Contract=Contract
        self.remarks=remarks
        audit.append(self)

class Rezervari():
    def __init__(self,resNo,clName,etaDate,rata,CDW,PAI,LI,TP,Remarks):
        self.resNo=resNo
        self.clName=clName
        self.etaDate=etaDate.replace("-","/")
        self.rata=rata
        self.CDW=CDW
        self.PAI=PAI
        self.LI=LI
        self.TP=TP
        self.Remarks=Remarks
        rezervari_importate.append(self)

for linie in lista_agenti:
    obiect = len(globals())
    globals()[obiect] = Agenti(linie[0], linie[1], linie[2])
    
for linie in lista_scrive:
    if linie[10]=="Signing party" and linie[2]!="cancelled":# and "Budget" not in linie[3]:
        obiect = len(globals())
        globals()[obiect] = Scrive(linie[2], linie[3], linie[10], linie[11], linie[17], linie[21], linie[22], linie[25], linie[9],linie[19])
    
for linie in lista_contracte:
    obiect = len(globals())
    globals()[obiect] = Contracte(linie[0], linie[1], linie[2],  linie[6], linie[7], linie[10], linie[11], linie[12], linie[13], linie[27], linie[28], linie[30], linie[31], linie[33], linie[34], linie[35], linie[36], linie[37], linie[40], linie[56],linie[58], linie[68], linie[69],linie[70], linie[72], linie[73], linie[74], linie[75], linie[76], linie[77], linie[78], linie[79], linie[80], linie[81], linie[92], linie[94], linie[97], linie[98], linie[102], linie[103], linie[104], linie[20], linie[23], linie[49], linie[50], linie[51], linie[101], linie[14], linie[67])

for linie in lista_tari:
    obiect = len(globals())
    globals()[obiect] = Tari(linie[0], linie[1])

for linie in lista_masini:
    obiect = len(globals())
    globals()[obiect] = Masini(linie[8], linie[13])

for linie in lista_ajustari:
    obiect = len(globals())
    globals()[obiect] = Ajustari(linie[0], linie[6], linie[7], linie[10])

for linie in lista_audit:
    obiect = len(globals())
    globals()[obiect] = Audit(linie[0], linie[1])

for linie in lista_rezervari:
    obiect = len(globals())
    globals()[obiect] = Rezervari(linie[1], linie[4], linie[7],linie[12],linie[59],linie[60],linie[61],linie[62],linie[15])

def verificare_existenta_contract_in_contracte_importate(contract_cautat):
    brand=""
    for i in contracte_importate:
        if i.Contract==contract_cautat:
            brand=i.Brand
    if brand=="":
        return None
    else:
        return brand
    
def incarcare_raport(contract,brand,remarks):
    global raportari
    dictionar={"Contract":contract,"Brand":brand,"Remarks":remarks}
    raportari.append(dictionar)

def raport1(): #identificare contracte nesemnate in Scrive
    for i in contracte_scrive:
        temp=0
        if i.Status=="timeouted":
            for j in contracte_scrive:
                if i.RA==j.RA:
                    if j.Status=="signed":
                        temp+=1
                else:
                    pass
            if temp==0:
                remarks= "Contractul "+i.RA+" nu este semnat."
                #print remarks
                incarcare_raport(i.RA,i.Brand,remarks) 

def raport2(): #identificare contracte cu permise posibil gresite
    lista_permise_gresite=[]
    lista_coduri_IDP_plus_US_CA=lista_coduri_IDP
    #lista_coduri_IDP_plus_US_CA.append("US") #nu stiu de ce am pus US la un moment dat in lista asta
    #lista_coduri_IDP_plus_US_CA.append("CA") #nu stiu de ce am pus CA la un moment dat in lista asta
    for i in contracte_scrive:
        if i.drivinglicense[0:6]=="ILXXID":
            remarks = "Pe contractul "+i.RA+" s-a introdus la permis numarul de ID al clientului"
            #print remarks
            incarcare_raport(i.RA,i.Brand,remarks)
            lista_permise_gresite.append(i.RA)
        elif i.drivinglicense[0:4]=="USXX":
            remarks = "Pe contractul "+i.RA+" s-a introdus gresit numarul de permis de SUA, lipseste codul statului"
            #print remarks
            incarcare_raport(i.RA,i.Brand,remarks)            
            lista_permise_gresite.append(i.RA)            
        elif i.drivinglicense[0:4]=="CAXX":
            remarks = "Pe contractul "+i.RA+" s-a introdus gresit numarul de permis de CANADA, lipseste codul provinciei"
            #print remarks
            incarcare_raport(i.RA,i.Brand,remarks)          
            lista_permise_gresite.append(i.RA)
        elif i.drivinglicense[0:4]=="GBXX":
            temp="GBXX"+i.nume.replace(" ", "")
            temp=temp[0:9]
            if temp not in i.drivinglicense:
                remarks = "Pe contractul "+i.RA+" este posibil ca numarul de permis de Marea Britanie sa fie gresit"
                #print remarks
                incarcare_raport(i.RA,i.Brand,remarks)                   
                lista_permise_gresite.append(i.RA)
        elif i.drivinglicense[0:4]=="ROXX":
            if i.drivinglicense[4:5].isdigit():
                remarks = "Pe contractul "+i.RA+" este posibil ca numarul de permis de Romania sa fie gresit"
                #print remarks
                incarcare_raport(i.RA,i.Brand,remarks)   
                lista_permise_gresite.append(i.RA)
        elif i.drivinglicense[0:4]=="CNXX":
            remarks = "Pe contractul "+i.RA+" s-a dat masina cu permis de China"
            #print remarks
            incarcare_raport(i.RA,i.Brand,remarks)                  
            lista_permise_gresite.append(i.RA)
        elif i.drivinglicense[0:2] not in lista_coduri_IDP_plus_US_CA:
            remarks = "Pe contractul "+i.RA+" s-a dat masina fara IDP emis de: "+i.drivinglicense[0:2]
            #print remarks
            incarcare_raport(i.RA,i.Brand,remarks)                  
        if (i.drivinglicense[0:2]!="") and (i.drivinglicense[0:2] not in lista_coduri_tari):
            remarks = "Pe contractul "+i.RA+" s-a folosit un cod de tara care nu exista (" +i.drivinglicense[0:2]+")"
            #print remarks
            incarcare_raport(i.RA,i.Brand,remarks)               
        elif (i.drivinglicense[0:2]=="") and (i.driverslicense[0:2] not in lista_coduri_tari):
            remarks = "Pe contractul "+i.RA+" s-a folosit un cod de tara care nu exista (" +i.driverslicense[0:2]+")"
            #print remarks
            incarcare_raport(i.RA,i.Brand,remarks)   
        if len(i.drivinglicense)<8 and (i.RA not in lista_permise_gresite) and (i.drivinglicense!=""):
            remarks = "Pe contractul "+i.RA+" este un permis mult prea scurt"
            #print remarks
            incarcare_raport(i.RA,i.Brand,remarks)   
            lista_permise_gresite.append(i.RA)
    for i in contracte_scrive:
        if "DEL" in i.drivinglicense[5:7]:
            remarks = "Pe contractul "+i.RA+" a ramas cu DEL"
            #print remarks
            incarcare_raport(i.RA,i.Brand,remarks)             
        if "PREF" in i.drivinglicense[5:8]:
            remarks = "Pe contractul "+i.RA+" a ramas cu PREF"
            #print remarks
            incarcare_raport(i.RA,i.Brand,remarks)             
        if i.drivinglicense.isalpha():
            remarks = "Pe contractul "+i.RA+" permisul contine doar litere"
            #print remarks
            incarcare_raport(i.RA,i.Brand,remarks)
            
def raport3(): #identificare soferi seniori si soferi tineri
    for i in contracte_importate:
        if i.out_LocT!="P36":
            if (i.out_Date!="" and i.Date_of_birth !="" and i.license_No != "NOSHOW" and i.cancelled==False):
                first=time.strptime(i.out_Date[0:10],"%d/%m/%Y")
                second=time.strptime(i.Date_of_birth[0:10],"%d/%m/%Y")
                first=datetime(first.tm_year, first.tm_mon, first.tm_mday)
                second=datetime(second.tm_year, second.tm_mon, second.tm_mday)
                difference = relativedelta.relativedelta(first,second)
                if difference.years >= 70 and "SDF" not in i.produse:
                    remarks = "Pe contractul "+i.Contract+" trebuia incasat soferul senior. Soferul are " + str(difference.years) + " ani."
                    #print remarks
                    incarcare_raport(i.Contract,i.Brand,remarks)
                elif difference.years > 21 and difference.years < 25 and i.grp_Ctr not in grupe_sub_25_ani:
                    remarks = "Pe contractul "+i.Contract+" clientul are intre 21 si 25 ani ("+str(difference.years)+"ani) si a inchiriat o grupa pe care nu ar fi trebuit sa o conduca ("+str(i.grp_Ctr)+")"
                    #print remarks
                    incarcare_raport(i.Contract,i.Brand,remarks)   
                elif difference.years < 21:
                    remarks = "Pe contractul "+i.Contract+" clientul are intre 18 si 21 ani si trebuia sa fie taxat pentru sofer sub 21 de ani."
                    #print remarks
                    incarcare_raport(i.Contract,i.Brand,remarks)   
                    if i.grp_Ctr not in grupe_sub_21_ani:
                        remarks = "Pe contractul "+i.Contract+" clientul are intre 18 si 21 ani ("+str(difference.years)+" ani) si a inchiriat o grupa pe care nu ar fi trebuit sa o conduca ("+str(i.grp_Ctr)+")"
                        #print remarks
                        incarcare_raport(i.Contract,i.Brand,remarks)             

def raport4(): #identificare contracte unde conduce altcineva in locul titularului de contract
    for i in contracte_importate:
        temp=0
        if "COND" in i.Remarks:
            temp=1
        elif "DRIVER" in i.Remarks:
            temp=1
        elif "DRV" in i.Remarks:
            temp=1
        elif "SOFER" in i.Remarks:
            temp=1
        if "COND" in i.Renter_ADDR2:
            temp=1
        elif "DRIVER" in i.Renter_ADDR2:
            temp=1
        elif "DRV" in i.Renter_ADDR2:
            temp=1
        elif "SOFER" in i.Renter_ADDR2:
            temp=1
        if i.Company_Name and "COND" in i.Company_Name:
            temp=1
        elif i.Company_Name and "DRIVER" in i.Company_Name:
            temp=1
        elif i.Company_Name and "DRV" in i.Company_Name:
            temp=1
        elif i.Company_Name and "SOFER" in i.Company_Name:
            temp=1              
        if temp==1:
            remarks = "Pe contractul "+i.Contract+" conduce altcineva decat titularul"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)  
            
def raport5(): #identificare contracte unde s-a folosit un permis de conducere dintr-o tara mai putin uzuala
    for i in contracte_scrive:
        gasit=False
        temp=i.drivinglicense[0:2]
        for element in contracte_importate:
            if ((element.Rezervare[8:10]==temp) and (element.Contract==i.RA)):
                gasit=True
            elif temp =="":
                temp2=i.driverslicense.replace("CONDUIRE","")
                temp2=temp2[0:2]
                if ((element.Rezervare[8:10]==temp2) and (element.Contract==i.RA)):
                    gasit=True
        if gasit==False and temp not in lista_coduri_tari_uzuale and i.Status == "signed":
            remarks = str("Pe contractul "+i.RA+" s-a folosit un permis emis de o tara mai putin uzuala: "+temp)
            #print remarks
            incarcare_raport(i.RA,i.Brand,remarks)

def raport6(): #identificare vouchere travel
    for i in contracte_importate:
        if ("T" in i.credit_card and i.license_No != "NOSHOW" and i.cancelled != True):
            remarks = "Pe contractul "+i.Contract+" nu exista card pentru garantare"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)          

def raport7(): #identificare contracte anulate unde nu s-a dat comanda MOD
    for i in contracte_importate:
        if i.cancelled and i.total_charge != "0":
            remarks = "Pe contractul "+i.Contract+" trebuia data comanda MOD dupa anulare"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)  
            
def raport8(): #identificare contracte anulate
    for i in contracte_importate:
        if i.cancelled == True and i.tarif_1_in=="NFI":
            remarks = "Contractul NO SHOW "+i.Contract+" a fost anulat"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)  

def raport9(): #identificare contracte cu AWD de angajat
    for i in contracte_importate:
        if "E1537" in i.awd1  and i.cancelled==False and i.tarif_1_in != "NFI":
            remarks = "Pe contractul "+i.Contract+" s-a folosit AWD de angajat"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)  

def raport10(): #identificare produse vandute pe cont WEB
    contracte_verificate_deja=[]
    for j in rezervari_importate:
        gasit=False
        contract_gasit=0
        for i in contracte_importate:
            if j.resNo==i.Rezervare:
                gasit=True
                contract_gasit=i
                contracte_verificate_deja.append(contract_gasit.Contract)
        if gasit==True:                
            if ("WEB-NO CID" in contract_gasit.Remarks) and contract_gasit.cancelled==False:
                if contract_gasit.li_Amt!="0" and j.LI=="0":
                    remarks = "Pe contractul "+contract_gasit.Contract+" care este pe cont WEB s-a vandut LI"
                    #print remarks
                    incarcare_raport(contract_gasit.Contract,contract_gasit.Brand,remarks) 
                if contract_gasit.pai_Amt!="0"  and j.PAI=="0":
                    remarks =  "Pe contractul "+contract_gasit.Contract+" care este pe cont WEB s-a vandut PAI"
                    #print remarks
                    incarcare_raport(contract_gasit.Contract,contract_gasit.Brand,remarks)
                if contract_gasit.cdw_Amt!="0" and j.CDW!=contract_gasit.cdw_Amt:
                    remarks =  "Pe contractul "+contract_gasit.Contract+" care este pe cont WEB s-a vandut CDW/SCDW"
                    #print remarks
                    incarcare_raport(contract_gasit.Contract,contract_gasit.Brand,remarks)                
                if "UPL" in contract_gasit.produse:
                    remarks =  "Pe contractul "+contract_gasit.Contract+" care este pe cont WEB s-a vandut UPL"
                    #print remarks
                    incarcare_raport(contract_gasit.Contract,contract_gasit.Brand,remarks) 
                if "DIE" in contract_gasit.produse:
                    remarks =  "Pe contractul "+contract_gasit.Contract+" care este pe cont WEB s-a vandut DIE"
                    #print remarks
                    incarcare_raport(contract_gasit.Contract,contract_gasit.Brand,remarks) 
                if "RSN" in contract_gasit.produse:
                    remarks =  "Pe contractul "+contract_gasit.Contract+" care este pe cont WEB s-a vandut RSN"
                    #print remarks
                    incarcare_raport(contract_gasit.Contract,contract_gasit.Brand,remarks)
    for i in contracte_importate:
        if ("WEB-NO CID" in i.Remarks) and i.cancelled==False and i.Contract not in contracte_verificate_deja:
                if i.li_Amt!="0":
                    remarks = "Pe contractul "+i.Contract+" care este pe cont WEB s-a vandut LI"
                    #print remarks
                    incarcare_raport(i.Contract,i.Brand,remarks) 
                if i.pai_Amt!="0":
                    remarks =  "Pe contractul "+i.Contract+" care este pe cont WEB s-a vandut PAI"
                    #print remarks
                    incarcare_raport(i.Contract,i.Brand,remarks)
                if i.cdw_Amt!="0":
                    remarks =  "Pe contractul "+i.Contract+" care este pe cont WEB s-a vandut CDW/SCDW"
                    #print remarks
                    incarcare_raport(i.Contract,i.Brand,remarks)                
                if "UPL" in i.produse:
                    remarks =  "Pe contractul "+i.Contract+" care este pe cont WEB s-a vandut UPL"
                    #print remarks
                    incarcare_raport(i.Contract,i.Brand,remarks) 
                if "DIE" in i.produse:
                    remarks =  "Pe contractul "+i.Contract+" care este pe cont WEB s-a vandut DIE"
                    #print remarks
                    incarcare_raport(i.Contract,i.Brand,remarks) 
                if "RSN" in i.produse:
                    remarks =  "Pe contractul "+i.Contract+" care este pe cont WEB s-a vandut RSN"
                    #print remarks
                    incarcare_raport(i.Contract,i.Brand,remarks)            

def raport11(): #identificare contracte Preferred deschise in timp real
    for i in contracte_importate:
        gasit=False
        rezervarea_gasita=0
        for j in rezervari_importate:
            if i.Rezervare==j.resNo:
                rezervarea_gasita=j
                gasit=True
        if gasit and("PREFERRED" in i.Remarks or "PREFERRED" in rezervarea_gasita.Remarks) and i.cancelled==False and i.out_LocT in statii_preferred and i.agent_out!="Automated No Show Fee" and i.tarif_1_out not in ["5XI","6XI","7XI","8XI","9XI"]:
            first=time.strptime(i.out_Date,"%d/%m/%Y %H:%M")
            second=time.strptime(i.entry_out_date,"%d/%m/%Y %H:%M")
            third=time.strptime(rezervarea_gasita.etaDate,"%Y/%m/%d %H:%M")
            first=datetime(first.tm_year, first.tm_mon, first.tm_mday, first.tm_hour, first.tm_min)
            second=datetime(second.tm_year, second.tm_mon, second.tm_mday, second.tm_hour, second.tm_min)
            third=datetime(third.tm_year, third.tm_mon, third.tm_mday, third.tm_hour, third.tm_min)
            difference = relativedelta.relativedelta(first,second)
            diferenta= difference.hours*60+difference.minutes
            difference2 = relativedelta.relativedelta(first,third)
            diferenta2= difference2.hours*60+difference2.minutes
            if diferenta==0 and diferenta2>=0 and i.Rezervare!="":
                remarks = "Contractul de Preferred "+i.Contract+" a fost deschis in timp real"
                #print remarks
                incarcare_raport(i.Contract,i.Brand,remarks)
        if gasit and("PRESIDENT" in i.Remarks or "PRESIDENT" in rezervarea_gasita.Remarks) and i.cancelled==False and i.out_LocT in statii_preferred and i.agent_out!="Automated No Show Fee" and i.tarif_1_out not in ["5XI","6XI","7XI","8XI","9XI"]:
            first=time.strptime(i.out_Date,"%d/%m/%Y %H:%M")
            second=time.strptime(i.entry_out_date,"%d/%m/%Y %H:%M")
            third=time.strptime(rezervarea_gasita.etaDate,"%Y/%m/%d %H:%M")
            first=datetime(first.tm_year, first.tm_mon, first.tm_mday, first.tm_hour, first.tm_min)
            second=datetime(second.tm_year, second.tm_mon, second.tm_mday, second.tm_hour, second.tm_min)
            third=datetime(third.tm_year, third.tm_mon, third.tm_mday, third.tm_hour, third.tm_min)
            difference = relativedelta.relativedelta(first,second)
            diferenta= difference.hours*60+difference.minutes
            difference2 = relativedelta.relativedelta(first,third)
            diferenta2= difference2.hours*60+difference2.minutes            
            if diferenta==0 and diferenta2>=0 and i.Rezervare!="":           
                remarks = "Contractul de Presidents Club "+i.Contract+" a fost deschis in timp real"
                #print remarks
                incarcare_raport(i.Contract,i.Brand,remarks)                 

def raport12(): #identificare contracte Uncollectable
    for i in contracte_importate:
        if "S/P" in i.credit_card and i.license_No != "NOSHOW" and i.cancelled==False:
            remarks = "Contractul "+i.Contract+" este inchis pe S/P"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)
        if "DI" in i.credit_card and i.license_No != "NOSHOW" and i.cancelled==False:
            remarks =  "Contractul "+i.Contract+" este inchis pe Direct Invoice"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)            
        if "UC" in i.credit_card and i.license_No != "NOSHOW" and i.cancelled==False:
            remarks = "Contractul "+i.Contract+" este inchis pe Uncollectable"
            #print remarks
            incarcare_raport(i.Contracti.Brand,remarks)            

def raport13(): #identificare contracte pe care s-a inchiriat o masina care nu avea voie sa paraseasca tara
    for i in contracte_importate:
        if i.grp_Ctr in grupe_non_TEF and "TEF" in i.produse and "A5555" not in i.awd1:
            remarks = "Pe contractul "+i.Contract+" s-a inchiriat o masina care nu avea voie sa paraseasca tara"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks) 

def raport14(): #identificare contracte cu adrese prea scurte
    for i in contracte_importate:
        if len(i.Renter_ADDR1)<8 and i.license_No != "NOSHOW" and i.cancelled==False:
            remarks = "Pe contractul "+i.Contract+" s-a trecut o adresa AD1 prea scurta"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)              
        if len(i.Renter_ADDR3.replace(" ",""))<6 and i.license_No != "NOSHOW" and i.cancelled==False:
            remarks = "Pe contractul "+i.Contract+" s-a trecut o adresa AD3 prea scurta"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)              

def raport15(): #identificare contracte cu multi km parcursi
    for i in contracte_importate:
        temp=int(i.LOR)*500
        if "-" not in i.Km_drive and temp<int(i.Km_drive):
            remarks = str("Pe contractul "+i.Contract+" s-au facut "+i.Km_drive+" km fata de "+str(temp)+" km cati avea dreptul.")
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)               
    for i in contracte_importate:
        temp2=int(i.LOR)*10   
        if i.Status=="closed" and "-" not in i.Km_drive and temp2>int(i.Km_drive) and i.cancelled==False:
            remarks = str("Pe contractul "+i.Contract+" s-au facut prea putini km ("+i.Km_drive+" km).")
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)              

def raport16(): #identificare contracte deschise/inchise delayed dupa mai multe zile
    for i in contracte_importate:
        first=time.strptime(i.out_Date,"%d/%m/%Y %H:%M")
        second=time.strptime(i.entry_out_date,"%d/%m/%Y %H:%M")
        first=datetime(first.tm_year, first.tm_mon, first.tm_mday, first.tm_hour, first.tm_min)
        second=datetime(second.tm_year, second.tm_mon, second.tm_mday, second.tm_hour, second.tm_min)
        difference = relativedelta.relativedelta(first,second)
        if i.Status!="cancelled" and difference.days>3:
            remarks = str("Contractul "+i.Contract+" s-a deschis in sistem dupa "+str(difference.days)+" zile")
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)              
            
        if i.Status=="closed":
            first=time.strptime(i.in_Date,"%d/%m/%Y %H:%M")
            second=time.strptime(i.entry_in_date,"%d/%m/%Y %H:%M")
            first=datetime(first.tm_year, first.tm_mon, first.tm_mday, first.tm_hour, first.tm_min)
            second=datetime(second.tm_year, second.tm_mon, second.tm_mday, second.tm_hour, second.tm_min)
            difference = relativedelta.relativedelta(second,first)
            if difference.days>3:
                remarks = str("Contractul "+i.Contract+" s-a inchis in sistem dupa "+str(difference.days)+" zile")
                #print remarks
                incarcare_raport(i.Contract,i.Brand,remarks)              


def raport17(): #identificare contracte nesemnate in Scrive
    contracte_fara_masini=[]
    contracte_cu_alte_masini=[]
    contracte_nesemnate=[]
    for i in contracte_importate:
        temp=0
        for j in contracte_scrive:
            if i.Contract==j.RA and i.license_No!=j.registration and i.mva_repl=="0" and j.Status=="signed":
                if j.registration=="":
                    contracte_fara_masini.append(i)
                else:
                    contracte_cu_alte_masini.append((i,j.registration))
            if i.Contract==j.RA and j.Status=="signed":
                temp+=1
        if temp==0 and i.tarif_1_in != "NFI":
            if i.cancelled == False and i.out_LocT != "C1J" and i.out_LocT != "R2Z" and i.out_LocT != "B2I" and i.awd1!="N2812000" and i.awd1!="K6421000":
                contracte_nesemnate.append(i)
                
    for element in contracte_fara_masini:
        remarks = "Contractul "+element.Contract+" este semnat in Scrive fara sa aiba masina "+element.license_No+" trecuta pe el."    
        #print remarks
        incarcare_raport(element.Contract,element.Brand,remarks)

    for element in contracte_cu_alte_masini:
        remarks = "Contractul "+element[0].Contract+" este semnat in Scrive pe "+element[1]+", insa s-a inchiriat "+element[0].license_No
        #print remarks
        incarcare_raport(element[0].Contract,element[0].Brand,remarks)
        
    for element in contracte_nesemnate:
        remarks = "Pentru contractul "+element.Contract+" nu a fost gasit contract semnat in Scrive."
        #print remarks
        incarcare_raport(element.Contract,element.Brand,remarks)
        
def raport18(): #identificare incasare eroanata DIE
    for i in contracte_importate:
        for j in masini_importate:
            if i.mva==j.mva and "DIE" in i.produse and j.carburant=="Gasoline":
                remarks = "Pe contractul "+i.Contract+" s-a incasat DIE desi masina este pe benzina"
                print remarks
                incarcare_raport(i.Contract,i.Brand, remarks)

def raport19(): #identificare incasare NSF pe contracte care nu sunt deschise pentru No Show
    for i in contracte_importate:
        if "NSF" in i.produse and "NFI" != i.tarif_1_in:
            remarks = "Pe contractul "+i.Contract+" s-a incasat NSF pe un contract normal"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)

def raport20(): #identificare contracte pe care difera numele intre contractul din Wizard si contractul din Scrive
    for i in contracte_importate:
        temp=0
        for j in contracte_scrive:
            if i.Contract==j.RA and i.Customer_Name.replace(" ","").replace(",","")!=j.Party_name.replace(" ","").replace(",","") and j.Status=="signed" and (j.Party_name.replace(" ","").replace(",","") not in  i.Customer_Name.replace(" ","").replace(",","")):
                temp=1
                nume_temp=j.Party_name
        if temp==1:
            remarks = "Pe contractul "+i.Contract+" sunt diferente de nume. In Wizard este "+i.Customer_Name+" / In Scrive este "+nume_temp
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)

def raport21(): #incarcare ajustari
    for i in ajustari:
        if i.cod_ajustare == "":
            remarks = str("Pe contractul "+i.Contract+" exista o ajustare de "+i.Road_Exp+" EUR prin Road Expenses")
            incarcare_raport(i.Contract,verificare_existenta_contract_in_contracte_importate(i.Contract),remarks)
        else:
            remarks = str("Pe contractul "+i.Contract+" exista o ajustare de "+i.valoare+" EUR")
            incarcare_raport(i.Contract,verificare_existenta_contract_in_contracte_importate(i.Contract),remarks)
            if i.cod_ajustare!="E":
                remarks = str("Pe contractul "+i.Contract+" ajustarea este introdusa cu alta litera decat E")
                incarcare_raport(i.Contract,verificare_existenta_contract_in_contracte_importate(i.Contract),remarks)
                
           
def raport22(): #incarcari raportari din audit
    lista_temp=[]
    for i in audit:
        temp=False
        for j in contracte_importate:
            if i.Contract == j.Contract:
                temp=True                    
        if temp==True:
            if i.remarks=="Checkin delay analysis":
                pass
            elif i.remarks=="Adjusted T&M":
                pass
            elif i.remarks=="No charge rental":
                pass            
            else:
                lista_temp.append(i)
        else:
            lista_temp.append(i)                
                
    for i in lista_temp:
        remarks = "Pe contractul "+i.Contract+" exista alerta: "+i.remarks
        #print remarks
        incarcare_raport(i.Contract,verificare_existenta_contract_in_contracte_importate(i.Contract),remarks)

        
def raport23(): #long duration rentals
    for i in contracte_importate:
        if int(i.LOR) > 30 and i.Status== "closed" and i.out_LocT != "C1J" and i.out_LocT != "R2Z":
            remarks = "Pe contractul "+i.Contract+" sunt mai mult de 30 de zile"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)

def raport24(): #identificare taxare DIE pe grupe L si O
    for i in contracte_importate:
        if "DIE" in i.produse and i.grp_Res==i.grp_Ctr=="L":
            remarks = "Pe contractul "+i.Contract+" s-a taxat DIE pe grupa rezervata L"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)
        if "DIE" in i.produse and i.grp_Res==i.grp_Ctr=="O":
            remarks = "Pe contractul "+i.Contract+" s-a taxat DIE pe grupa rezervata O"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)

def raport25(): #identificare contracte cu mocangeala mica sau downgrade hotie
    for i in contracte_importate:
        if i. sell_status in ["mocangeala mica","downgrade hotie"] and "NSF" not in i.produse:
            remarks = "Pe contractul "+i.Contract+" s-a rezervat grupa "+i.grp_Res+", s-a platit grupa "+i.grp_Crg+" si s-a primit grupa "+i.grp_Ctr
            if "UPL" in i.produse:
                remarks=remarks+", dar a mai fost aplicat si UPL"
            #print remarks
            incarcare_raport(i.Contract,i.Brand,remarks)
        
def generare_raport():
    global raportari
    raportari = sorted(raportari, key=lambda k: k['Contract'])
    Row=0
    Column=0
    fisier_excel = "Audit.xlsx"
    workbook=xlsxwriter.Workbook(fisier_excel)
    worksheet=workbook.add_worksheet()
    worksheet.write(Row,Column,"Contract")
    worksheet.write(Row,Column+1,"Brand")
    worksheet.write(Row,Column+2,"Remarks")
    worksheet.write(Row,Column+3,"Agent out")
    worksheet.write(Row,Column+4,"Agent in")   
    worksheet.write(Row,Column+5,"Out location")
    worksheet.write(Row,Column+6,"In location")
    Row+=1
    cur = db.cursor()
    for element in raportari:
        contract=element.get("Contract")
        worksheet.write(Row,Column,contract)
        worksheet.write(Row,Column+1,element.get("Brand"))
        worksheet.write(Row,Column+2,element.get("Remarks"))
        query="INSERT INTO audit VALUES('"+contract+"','"+element.get("Remarks")+"')"
        print query
        cur.execute(query)
        db.commit()
        for row in cur.fetchall():
            print row[0]
        agent_out=""
        agent_in=""
        out_LocT=""
        in_LocT=""
        for i in contracte_importate:
            if contract==i.Contract:
                agent_out=i.agent_out
                agent_in=i.agent_in
                out_LocT=i.out_LocT
                in_LocT=i.in_LocT
        worksheet.write(Row,Column+3,agent_out)
        worksheet.write(Row,Column+4,agent_in)   
        worksheet.write(Row,Column+5,out_LocT)
        worksheet.write(Row,Column+6,in_LocT)
        Row+=1
    workbook.close()
    db.close()

    
def toaterapoartele():
    raport1() #identificare contracte nesemnate in Scrive (contract incarcat in Scrive, dar nesemnat)
    raport2() #identificare contracte cu permise posibil gresite sau fara IDP
    raport3() #identificare soferi seniori si soferi tineri
    raport4() #identificare contracte unde conduce altcineva in locul titularului de contract
    raport5() #identificare contracte unde s-a folosit un permis de conducere dintr-o tara mai putin uzuala
    raport6() #identificare vouchere travel fara garantie pe card
    raport7() #identificare contracte anulate unde nu s-a dat comanda MOD
    raport8() #identificare contracte anulate
    raport9() #identificare contracte cu AWD de angajat
    raport10() #identificare produse vandute pe cont WEB
    raport11() #identificare contracte Preferred deschise in timp real
    raport12() #identificare contracte Uncollectable
    raport13() #identificare contracte pe care s-a inchiriat o masina care nu avea voie sa paraseasca tara
    raport14() #identificare contracte cu adrese prea scurte
    raport15() #identificare contracte cu multi km parcursi
    raport16() #identificare contracte deschise/inchise delayed dupa mai multe zile
    raport17() #identificare contracte nesemnate in Scrive (contract neincarcat in Scrive)
    raport18() #identificare incasare eroanata DIE
    raport19() #identificare incasare NSF pe contracte care nu sunt deschise pentru No Show
    raport20() #identificare contracte pe care difera numele intre contractul din Wizard si contractul din Scrive
    raport21() #incarcare ajustari
    raport22() #incarcari raportari de la Avis Europe
    raport23() #long duration rentals
    raport24() #identificare taxare DIE pe grupe rezervate L si O
    raport25() #identificare contracte cu mocangeala mica sau downgrade hotie
    generare_raport()
    
print "Program running, please wait..."
toaterapoartele()

timp_doi=time.time()
diferenta_timp=timp_doi-primul_timp
print "It took "+str(round(diferenta_timp,2))+" seconds to complete the checking"
raw_input("Apasa <enter> pentru iesire")
