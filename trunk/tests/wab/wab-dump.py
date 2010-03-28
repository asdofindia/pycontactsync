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
#The WAB is open without an identity.
objSession.Open(False)

if objSession.Identities.LastIdentity == "{00000000-0000-0000-0000-000000000000}":
    print "No Identity Selected"
else:
    print "The Identity " + objSession.Identities(objSession.Identities.LastIdentity).Name + " is selected"

print "Default ID ", objSession.Identities.DefaultIdentity
print "Number of ID's ", objSession.Identities.Count
print "Container Count ", objSession.Containers.Count

for oContU in objSession.Containers:
    print "Number of Contacts ", oContU.Elements.Count
    oItemNum = 0
    for oItem in oContU.Elements:
        oItemNum += 1
        print "Contact Name ", oItem.Name, " Number ", oItemNum, " ID ", oItem.Id
        print "Number of Elements", oItem.Properties.Count
        for oContP in oItem.Properties:
            print "Element ID 0x%08x Element Value %s " % (oContP.Id, oContP.Value)
objSession.Refresh
objSession.Close
print "Done."
