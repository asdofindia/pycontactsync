import csv
import win32com.client

if win32com.client.gencache.is_readonly == True:
    #allow gencache to create the cached wrapper objects
    win32com.client.gencache.is_readonly = False

    # under p2exe the call in gencache to __init__() does not happen
    # so we use Rebuild() to force the creation of the gen_py folder
    win32com.client.gencache.Rebuild()

    # NB You must ensure that the python...\win32com.client.gen_py dir does not exist
    # to allow creation of the cache in %temp%

from win32com.client.gencache import EnsureDispatch
from win32com.client import constants

objSession = EnsureDispatch("WABAccess.Session", bForDemand=0)

#print objSession

#Open the WAB with an identity. If no identity is already open,
#the Identity Manager is launch to prompt the user to choose an indentity.
#objSession.Open(True)
#

# This can be used to open a specific wab file
objSession.Open(False,'MyWabFile.wab')

#
#The WAB is open without an identity.
#objSession.Open(False)

if objSession.Identities.LastIdentity == "{00000000-0000-0000-0000-000000000000}":
    print "No Identity Selected"
else:
    print "The Identity " + objSession.Identities(objSession.Identities.LastIdentity).Name + " is selected"

print "Default ID ", objSession.Identities.DefaultIdentity
print "Number of ID's ", objSession.Identities.Count
print "Container Count ", objSession.Containers.Count
print "Version ", objSession.Version

def PrintAllContacts(WABSession):
    for oContU in WABSession.Containers:
        print "Number of Contacts ", oContU.Elements.Count
        oItemNum = 0
        for oItem in oContU.Elements:
            oItemNum += 1
            print "Contact Name ", oItem.Name, " Number ", oItemNum, " ID ", oItem.Id
            print "Number of Elements", oItem.Properties.Count
            for oContP in oItem.Properties:
                print "Element ID 0x%08x Element Value %s " % (oContP.Id, oContP.Value)


def PrintAllContactsFax(WABSession):
    for oContU in WABSession.Containers:
        print "Number of Contacts ", oContU.Elements.Count
        oItemNum = 0
        for oItem in oContU.Elements:
            oItemNum += 1
            print "Contact Name ", oItem.Name, " Number ", oItemNum, " ID ", oItem.Id
            print "Number of Elements", oItem.Properties.Count
            for oContP in oItem.Properties:
                #print "Element ID 0x%08x Element Value %s " % (oContP.Id, oContP.Value)
                if oContP.Id == constants.wabPR_PRIMARY_FAX_NUMBER:
                    print oItem.Name
                    print "Primary Fax " + oContP.Value
                elif oContP.Id == constants.wabPR_BUSINESS_FAX_NUMBER:
                    print oItem.Name
                    print "Business Fax " + oContP.Value
                elif oContP.Id == constants.wabPR_HOME_FAX_NUMBER:
                    print oItem.Name
                    print "Home Fax " + oContP.Value


def DeleteAllContacts(WABSession):
    for oContU in WABSession.Containers:
        print "Number of Contacts ", oContU.Elements.Count
        oItemNum = 0
        for oItem in oContU.Elements:
            oItemNum += 1
            print "Contact Name ", oItem.Name, " Number ", oItemNum, " ID ", oItem.Id
            oContU.Elements.Remove(oItem.Id)
    print "Numer of Deleted Contacts ", oItemNum        


def DeleteContact(WABSession, ContactNum):
    oContU = WABSession.Containers(1)
    oItem = oContU.Elements(ContactNum)
    print "Contact Name ", oItem.Name, " ID ", oItem.Id
    oContU.Elements.Remove(oItem.Id)
    print "Deleted"


def PrintContact(WABSession, ContactNum):
    oContU = WABSession.Containers(1)
    oItem = oContU.Elements(ContactNum)
    print "Contact Name ", oItem.Name, " ID ", oItem.Id


def CreateContactOld(WABSession, FirstName, Surname):
    oContU = WABSession.Containers(1)
    oContU.Elements.NewContact( False, FirstName + " " + Surname )
    print "Contact Name ", FirstName


def CreateContact(WABSession, FullName):
    oContU = WABSession.Containers(1)
    oContU.Elements.NewContact( False, FullName )
    print "Contact Name ", FullName


def ModifyContactFax(WABSession, ContactNum):
    oContU = WABSession.Containers(1)
    oItem = oContU.Elements(ContactNum)
    print "Contact Name ", oItem.Name, " ID ", oItem.Id
    oItem.Properties.Add(constants.wabPR_PRIMARY_FAX_NUMBER, "+27 011 680 1234")
    #print "Primary Fax Number ", oItem.Properties.Value
    oItem.Properties.Add(constants.wabPR_BUSINESS_FAX_NUMBER, "+27 011 960 2197")
    #print "Business Fax Number ", oItem.Properties.Value
    oItem.Properties.Add(constants.wabPR_HOME_FAX_NUMBER, "+27 011 960 7777")
    #print "Home Fax Number ", oItem.Properties.Value
    print "Contact Modified"

def ModifyContact(WABSession, ContactNum, WabAttrib, AttribValue):
    oContU = WABSession.Containers(1)
    oItem = oContU.Elements(ContactNum)
    print "Contact Name ", oItem.Name, " ID ", oItem.Id
    oItem.Properties.Add(WabAttrib, AttribValue)
    print "Adding ", AttribValue
    print "Contact Modified"


#DeleteAllContacts(objSession)
#PrintAllContacts(objSession)

#PrintContact(objSession, 2)
#DeleteContact(objSession, 2)
#PrintAllContacts(objSession)

#CreateContactOld(objSession, "Test", "Test")
#ModifyContact(objSession, 2)

#CreateContact(objSession, "Test")

#PrintAllContacts(objSession)
#PrintAllContactsFax(objSession)

#PrintAllContacts(objSession)

#CreateContact(objSession, "Test", "Test")
#ModifyContact(objSession, 1, constants.wabPR_BUSINESS_FAX_NUMBER, "+27 011 960 2197")


#"Name"	"Nickname"	"E-mail Address"	"Home Phone"	"Home Fax"	"Mobile Phone"	"Business Phone"	"Business Fax"	"Company"	"Job Title"
#"Name0" "Nickname1" "E-mail Address2" "Home Phone3" "Home Fax4" "Mobile Phone5"
#"Business Phone6" "Business Fax7" "Company8" "Job Title9"

NumberOfContacts = objSession.Containers(1).Elements.Count
print "Number of Contacts ", NumberOfContacts

def CSVImport(Test):
    ifile  = open('wab-dump.csv', "rb")
    reader = csv.reader(ifile, delimiter='\t', quotechar='"', quoting=csv.QUOTE_ALL)

    rownum = 0
    for row in reader:
        # Save header row.
        if rownum == 0:
            header = row
        else:
            colnum = 0

            for col in row:
                #print '%-8s: %s' % (header[colnum], col)

                if colnum == 0:
                    CreateContact(objSession, col)
                else:
                    if not col == '':
                        if colnum == 1:
                            ModifyContact(objSession, rownum + NumberOfContacts, constants.wabPR_NICKNAME, col)
                        elif colnum == 2:
                            # FixME:
                            # E-Mail in WAB is not stright forward.
                            #ModifyContact(objSession, rownum, constants.wabPR_CONTACT_EMAIL_ADDRESSES, col)
                            #
                            print "Would add e-mail, but we don't!"
                        elif colnum == 3:
                            ModifyContact(objSession, rownum + NumberOfContacts, constants.wabPR_HOME_TELEPHONE_NUMBER, col)
                        elif colnum == 4:
                            ModifyContact(objSession, rownum + NumberOfContacts, constants.wabPR_HOME_FAX_NUMBER, col)
                        elif colnum == 5:
                            ModifyContact(objSession, rownum + NumberOfContacts, constants.wabPR_CELLULAR_TELEPHONE_NUMBER, col)
                        elif colnum == 6:
                            ModifyContact(objSession, rownum + NumberOfContacts, constants.wabPR_HOME_FAX_NUMBER, col)
                        elif colnum == 7:
                            ModifyContact(objSession, rownum + NumberOfContacts, constants.wabPR_BUSINESS_TELEPHONE_NUMBER, col)
                        elif colnum == 8:
                            ModifyContact(objSession, rownum + NumberOfContacts, constants.wabPR_COMPANY_NAME, col)
                        elif colnum == 9:
                            ModifyContact(objSession, rownum + NumberOfContacts, constants.wabPR_PROFESSION, col)
                colnum += 1
        print 'Row number %d: ' % rownum
        rownum += 1

    ifile.close()

CSVImport(1)

objSession.Refresh
objSession.Close
print "Done."
