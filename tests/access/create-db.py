import os, sys
from win32com.client.gencache import EnsureDispatch as Dispatch

import time
import datetime
import pywintypes
import ConfigParser

if os.path.exists("access.ini"):
    config = ConfigParser.ConfigParser()
    config.read(r"access.ini")

    #print config.get('Auth', 'username', 0)
    #print config.get('Auth', 'password', 0)
    #print config.get('Admin', 'adminuser', 0)
    #print config.get('Admin', 'adminpasswd', 0)

    try:
        DB_FILEPATH = config.get('AccessDB', 'database', 0)
    except:
        DB_FILEPATH = r"test-01.mdb"
else:
    DB_FILEPATH = r"test-01.mdb"

#DB_FILEPATH = r"c:\temp\test.mdb"
CONNECTION_STRING = 'Provider=Microsoft.Jet.OLEDB.4.0;' + \
		'Jet OLEDB:Engine Type=5;' + \
                'data Source=%s' % DB_FILEPATH

adox = Dispatch("ADOX.Catalog")
db = Dispatch('ADODB.Connection')

fields = ['FullName',
            'CompanyName', 
            'MailingAddressStreet',
            'MailingAddressCity', 
            'MailingAddressState', 
            'MailingAddressPostalCode',
            'HomeTelephoneNumber', 
            'BusinessTelephoneNumber', 
            'MobileTelephoneNumber',
            'Email1Address',
            'Body'
        ]

print "Starting ..."
if os.path.exists(DB_FILEPATH):
    try:
        db.Open(CONNECTION_STRING)
        print "DB Opening ..."
    except:
        os.remove(DB_FILEPATH)

    db.Close()
else:
    # Create DataBase
    adox.Create(CONNECTION_STRING)
    db.Open(CONNECTION_STRING)
    adox.ActiveConnection = db

    # Create Tables
    db.Execute('CREATE TABLE contacts (AutoNumber AUTOINCREMENT,' + \
     'OUserName TEXT(50), CDate TIMESTAMP, MUserName TEXT(50), MDate TIMESTAMP,' + \
     'FullName TEXT(50), CompanyName TEXT(50), Email1Address TEXT(50), Body memo with comp)')

    db.Close()
    print "Close."


print "Checking DB ..."
db.Open(CONNECTION_STRING)
adox = None
adox = Dispatch(r'ADOX.Catalog')
adox.ActiveConnection = db

# List Tables    
oTab = adox.Tables
for x in oTab:
    if x.Type == 'TABLE':
        print x.Name

db.Close()
print "Close."
