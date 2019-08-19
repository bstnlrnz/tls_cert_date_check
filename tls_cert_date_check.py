###############################################################################
# Author: Bastian Lorenz
# easy_install pyopenssl
# easy_install xlrd
# easy_install termcolor
# easy_install xlsxwriter
# usage:
# python tls_cert_date_check.py /path/to/xlsx-File.xlsx <output.xlsx>
# python tls_cert_date_check.py -d domain.tld
# or configure "FILE". Writing to a xlsx-output file is optional.
# While reading from Excel-Files: Please fill in domain names in first column
###############################################################################
import sys
import xlrd
import numpy as np
import OpenSSL
import ssl
from datetime import datetime
from termcolor import colored
import xlsxwriter
######################################
FILE = "path/to/file.xlsx"
######################################
#Fuer die Ausgabe in ein Excel-File
#WR_FILE="/Users/bastian.lorenz/Desktop/output.xlsx"
######################################
if len(sys.argv) >= 2:
    if sys.argv[1] == "-d":
        DATA_FILE=sys.argv[2]
    else:
        DATA_FILE = sys.argv[1]
    if len(sys.argv) == 3 and sys.argv[1] != "-d":
        #Fuer die Ausgabe in ein Excel-File
        WR_FILE = sys.argv[2]
        book = xlsxwriter.Workbook(WR_FILE)
        worksheet = book.add_worksheet()
        row = 0
        col = 1
    if len(sys.argv) > 3:
        print("Ungueltiger Aufruf!")
        exit()
else:
    DATA_FILE=FILE

# Einlesen der Excel-Datei
def read_excel_file(filename):
    book = xlrd.open_workbook(filename, encoding_override = "utf-8")
    # 1. Arbeitsblatt
    sheet = book.sheet_by_index(0)
    # Einlesen der 1.Spalte
    x_data = np.asarray([sheet.cell(i, 0).value for i in range(0, sheet.nrows)])
    #y_data = np.asarray([sheet.cell(i, 1).value for i in range(1, sheet.nrows)])
    return x_data
    #return x_data, y_data

# Pruefen des Zertifikats
def tls_sni_check(domain):
    hostname = domain
    port = 443
    timeout = 3
    try:
        # Verbindungsaufbau
        conn = ssl.create_connection((hostname, port), timeout)
        context = ssl.SSLContext(ssl.PROTOCOL_SSLv23)
        sock = context.wrap_socket(conn, server_hostname=hostname)

        # Abruf Zertifikats
        #cert_bin = sock.getpeercert(True)
        #x509 = OpenSSL.crypto.load_certificate(OpenSSL.crypto.FILETYPE_ASN1, cert_bin)
        #print("CN=" + x509.get_subject().CN)

        # Umwandeln in PEM Format
        certificate = ssl.DER_cert_to_PEM_cert(sock.getpeercert(True))
        cert = OpenSSL.crypto.load_certificate(OpenSSL.crypto.FILETYPE_PEM, certificate)
        # Abfrage des CNs
        commonname = cert.get_subject().CN
        # Abfrage der Seriennummer
        serial = '{0:x}'.format(int(cert.get_serial_number()))
        # TRUE wenn abgelaufen
        expired = datetime.strptime(cert.get_notAfter().decode('utf-8'), "%Y%m%d%H%M%SZ") <= datetime.utcnow()
        # Fuer Berechnung verbleibender Laufzeit notwendig
        remain = datetime.strptime(cert.get_notAfter().decode('utf-8'), "%Y%m%d%H%M%SZ")
        # date beinhaltet Datum als String
        date = datetime.strftime(datetime.strptime(cert.get_notAfter().decode('utf-8'), "%Y%m%d%H%M%SZ"),"%Y-%m-%d %H:%M:%S")
        if expired == True:
            return ("!! UNGUELTIG !!"), date, remain, serial, commonname
        else:
            return ("GUELTIG"), date, remain, serial, commonname
    except:
        err="Timeout beim Verbindungsaufbau"
        return err, None, None, None, None

# Berechnen der verblebenden Laufzeit
def remaining_days(date):
    current = datetime.today() # war vorher utcnow()
    remaining = remain - current
    remaining =  str(remaining)
    return remaining.split('day')[0]

# Start des Programms
if len(sys.argv) >= 2:
    if sys.argv[1] == "-d":
        x_data=sys.argv[2]
        print("Es wird folgende Domain geprueft: "+x_data)
        size=1
    else:
        x_data = read_excel_file(DATA_FILE)
        size = x_data.size #Gibt Integer-Wert zurueck
        print("Es werden ingesamt %s Domains geprueft!" % size)
else:
    x_data = read_excel_file(DATA_FILE)
    size = x_data.size #Gibt Integer-Wert zurueck
    print("Es werden ingesamt %s Domains geprueft!" % size)

for i in range(0, size):
    # Ausgabe der Domain
    if sys.argv[1] != "-d":
        domain=(x_data[i])
    else:
        domain=x_data
    # Ausgabe GUELTIG, UNGUELTIG oder None
    result, date, remain, serial, commonname = tls_sni_check(domain)
    if len(sys.argv) == 3 and sys.argv[1] != "-d":
        worksheet.write(i, 0, domain) #1. Spalte Name der Domain
        worksheet.write(i, 1, date)   #2. Spalte Datum
        worksheet.write(i, 2, result) #3. Spalte: Gueltig, Ungueltig oder Fehlermeldung
        worksheet.write(i, 4, serial) #4. Spalte: SN der Zertifikats
        worksheet.write(i, 5, commonname) #5. Spalte: CN der Zertifikats
        if remain != None:
            worksheet.write(i,3, "Noch "+remaining_days(remain)+"Tage gueltig!") #4. Spalte: Ausgabe verbleibender Tage
            if result == "!! UNGUELTIG !!":
                worksheet.write(i,3,"Abgelaufen seit "+remaining_days(remain).split("-")[1]+"Tagen")
    else:
    # Ausgabe des Datums
        if result == "!! UNGUELTIG !!":
            #print(date),
            print colored (date+" "+domain+" CN:"+commonname+" "+result+" Abgelaufen seit "+remaining_days(remain).split("-")[1]+"Tagen", 'red')
        else:
            if result == None or remain == None:
                print colored ("Timeout beim Verbindungsaufbau zu "+domain, 'red')
            else:
                #print(date),
                print colored (date+" "+domain+" CN:"+commonname+" "+result+" "+"Restliche Tage: "+remaining_days(remain), 'green')
if len(sys.argv) == 3 and sys.argv[1] != "-d":
    book.close()
else:
    exit()
