######################################################################################################################################
######################################################################################################################################

import salesforce_reporting, datetime,time, ConfigParser, requests, pandas,os
from pandas import ExcelWriter

def get_keys():
    config = ConfigParser.ConfigParser()
    config.read(os.path.join(os.path.dirname((os.path.abspath(__file__))),"keysSFDC.txt"))
    usuario= config.get("KEYS","user")
    password= config.get("KEYS", "pass")
    security_token= config.get("KEYS", "security_token")
    reportid= config.get("KEYS","reportid")
    return usuario,password,security_token,reportid

def func_ppal():

    choice=raw_input("Desea cargar sus credenciales desde keysSFDC.ini? (y/n)")
    if choice=='y':
        llaves=get_keys()
        filename=raw_input("Ingrese el nombre de su file: ")
        print(llaves)
        bajadaSFDC(filename,llaves)
    if choice=='n':
        print("Ingrese sus credenciales manualmente:\n")
        llaves=[]
        i=0
        for i in range(3):
            if i==0:
                llaves.append(raw_input("Username: "))
            if i==1:
                llaves.append(raw_input("\nPassword: "))
            if i==2:
                llaves.append(raw_input("\nSecurity_token: "))
            if i==3:
                llaves.append(raw_input("\nReport ID: "))
        filename=raw_input("\nIngrese el nombre de su file: ")
        while a.isalnum()!=True:
            filename=raw_input("\nIngrese el nombre de su file (DEBE SER ALFANUMERICO: ")
        bajadaSFDC(filename,llaves)

def Loop():
        r = raw_input("Queres volver a correr el programa (y/n)?")
        if r == "yes" or r == "y":
            func_ppal()
        if r == "n" or r == "no":
            print("Cerrando Script. Sayonara.")


def bajadaSFDC(a,ll):

    my_sf = salesforce_reporting.Connection(llaves[0],llaves[1],llaves[2])
    print("Conectado con exito a SFDC.")
    print("http://eu5.salesforce.com/" + llaves[3])

    try:
        SFSF=requests.get("https://eu5.salesforce.com/" + str(llaves[3]))

    except Exception, e:
        print e + (":: No funcionó la request.")


    my_sf = salesforce_reporting.Connection(llaves[0],llaves[1],llaves[2])
    report= my_sf.get_report(llaves[3])

    ### ACA VER

    print(report)
    parser = salesforce_reporting.ReportParser(report)
    parsed=parser.records()
    labels=parser._get_field_labels()
    parsedlabels=[str(labels[x]) for x in range(len(labels))]

    print(parsed)
    print('''\n\n\n''')
    print('''\n\n\n''')
    TABLA=pandas.DataFrame(parsed)
    TABLA.columns = (parsedlabels)
    #TABLA= TABLA.drop_duplicates(subset=['Opportunity Name','Golden Opportunity ID','Requested Delivery Date'],keep='first',inplace=False)

    print('''\n\n\nSin Duplis:\n''')

    print('''\n EOF''')
    try:
        writer = ExcelWriter(a + '.xlsx')
        TABLA.to_excel(writer,a+'.xlsx')
        writer.save()
        print("El archivo " + a + ".xlsx se guardo con exito.")
        Loop()
    except IOError:
        print("\nEl archivo se encuentra en ejecucion, cierre el archivo o cambie el nombre del output file.")
        Loop()



    #' + datetime.date.today().strftime("%B%d%Y")


if __name__=='__main__':

    choice=raw_input("Desea cargar sus credenciales desde keysSFDC.ini? (y/n)")
    if choice=='y':
        llaves=get_keys()
        filename=raw_input("Ingrese el nombre de su file: ")
        print llaves
        bajadaSFDC(filename,llaves)
    if choice=='n':
        print("Ingrese sus credenciales manualmente:\n")
        llaves=[]
        i=0
        for i in range(4):
            if i==0:
                llaves.append(raw_input("Username: "))
            if i==1:
                llaves.append(raw_input("\nPassword: "))
            if i==2:
                llaves.append(raw_input("\nSecurity_token: "))
            if i==3:
                llaves.append(raw_input("\nReport ID: "))
        filename=raw_input("\nIngrese el nombre de su file: ")
        while filename.isalnum()!=True:
            filename=raw_input("\nIngrese el nombre de su file (DEBE SER ALFANUMERICO: ")
        bajadaSFDC(filename,llaves)
