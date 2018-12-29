# easy_install pyopenssl
# easy_install xlrd
# easy_install termcolor
# use:
# python tls_cert_date_check.py /path/to/xlsx-File.xlsx
# or configure "DATA_FILE"
import sys
import xlrd
import numpy as np
import OpenSSL
import ssl
from datetime import datetime
from termcolor import colored

if len(sys.argv) > 1 :
    DATA_FILE = sys.argv[1]
else:
    DATA_FILE="/path/to/File.xlsx"

#Einlesen der Excel-Datei

def read_excel_file(filename):
    book = xlrd.open_workbook(filename, encoding_override = "utf-8")
    # 1. Arbeitsblatt
    sheet = book.sheet_by_index(0)
    # Einlesen der 1.Spalte
    x_data = np.asarray([sheet.cell(i, 0).value for i in range(0, sheet.nrows)])
    #y_data = np.asarray([sheet.cell(i, 1).value for i in range(1, sheet.nrows)])
    return x_data
    #return x_data, y_data

#def tls_check(domain):
#    cert=ssl.get_server_certificate((domain, 443))
#    x509 = OpenSSL.crypto.load_certificate(OpenSSL.crypto.FILETYPE_PEM, cert)
#    date = x509.get_notAfter()
#    return date

def tls_sni_check(domain):
    hostname = domain
    port = 443
    timeout = 3

    try:
        conn = ssl.create_connection((hostname, port), timeout)
        context = ssl.SSLContext(ssl.PROTOCOL_SSLv23)
        sock = context.wrap_socket(conn, server_hostname=hostname)
        # Umwandeln in PEM Format
        certificate = ssl.DER_cert_to_PEM_cert(sock.getpeercert(True))
        # Abfragen des Ablaufdatums
        cert = OpenSSL.crypto.load_certificate(OpenSSL.crypto.FILETYPE_PEM, certificate)
        # TRUE wenn abgelaufen
        expired = datetime.strptime(cert.get_notAfter().decode('utf-8'), "%Y%m%d%H%M%SZ") <= datetime.utcnow()
        print (datetime.strptime(cert.get_notAfter().decode('utf-8'), "%Y%m%d%H%M%SZ")),
        if expired == True:
            return ("!! UNGUELTIG !!")
        else:
            return ("GUELTIG")
    except:
        return None

x_data = read_excel_file(DATA_FILE)
size = x_data.size
print("Es werden ingesamt %s Domains geprueft!: " % size)

for i in range(0, size):
    #tmp = tls_check(x_data[i])
    # Ausgabe der Domain
    domain=(x_data[i])
    # Ausgabe GUELTIG oder UNGUELTIG
    tmp = tls_sni_check(x_data[i])
    if tmp == "!! UNGUELTIG !!":
        print colored (domain+" "+tmp, 'red')
    else:
        if tmp == None:
            print colored ("Timeout beim Verbindungsaufbau zu "+domain, 'red')
        else:
            print (domain+" "+tmp)
