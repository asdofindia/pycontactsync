#
# Create MySQL DB with fields that have creation and update tiem stamps.
#
# http://gusiev.com/2009/04/update-and-create-timestamps-with-mysql/
#

from win32com.client import Dispatch
from adoconstants import *

# Create Connection object and connect to database.
oConn = Dispatch('ADODB.Connection')
oConn.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};" + \
                        "Server=192.168.2.102;Port=3306;" + \
                        "User=root;Password=password;Database=testdb"

oConn.Open()

# Now prepare the SQL statement to do the data insertion.
sql = "CREATE TABLE IF NOT EXISTS `contacts` (" + \
  "`AutoNumber` int(10) unsigned NOT NULL auto_increment primary key," + \
  "CDate timestamp default '0000-00-00 00:00:00'," + \
  "UDate timestamp default now() on update now()," + \
  "`CUserName` varchar(50) NOT NULL," + \
  "`FullName` varchar(50) default NULL," + \
  "`CompanyName` varchar(50) default NULL," + \
  "`HomeTelephoneNumber` varchar(20) default NULL," + \
  "`BusinessTelephoneNumber` varchar(20) default NULL," + \
  "`MobileTelephoneNumber` varchar(20) default NULL," + \
  "`Email1Address` varchar(50) default NULL," + \
  "`Body` blob);"

# Now execute the SQL statement that we prepared above.
oConn.Execute(sql)

# Close and clean up    
oConn.Close()
oConn = None    
