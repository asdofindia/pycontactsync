import win32com.client

#DEBUG = 0
DEBUG = 1

class MSOutlook:
    def __init__(self):
        self.outlookFound = 0
        try:
            self.oOutlookApp = \
                win32com.client.gencache.EnsureDispatch("Outlook.Application")
            self.outlookFound = 1
        except:
            print "MSOutlook: unable to load Outlook"
            self.outlookFound = 0
        
        self.records = []


    def loadContacts(self, keys=None):
        if not self.outlookFound:
            return

        print "Loading MAPI ..."
        try:
            onMAPI = self.oOutlookApp.GetNamespace("MAPI")
        except:
            print "Error loading MAPI!"

        print "Looking to load Redemption ..."
        # Set for Redemption DDL Installed
        RedemptionLib = -1
        try:
            Contact = win32com.client.Dispatch('Redemption.SafeContactItem')
            print "Using Redemption."
            RedemptionLib = 0
        except:
            print "No Redemption, loading standard."
            RedemptionLib = 1
            
        print "Getting Default Folder ..."
        try:
            oContacts = \
                onMAPI.GetDefaultFolder(win32com.client.constants.olFolderContacts)
            #oContacts = onMAPI.GetDefaultFolder(10) # 10=outlook contacts folder
        except:
            print "Error loading Folder Contact"

        if DEBUG:
            print "number of contacts:", len(oContacts.Items)

        for oc in range(len(oContacts.Items)):
            if RedemptionLib:
                # Use Outlook directly
                Contact = oContacts.Items.Item(oc+1)
            else:
                # Use the Redemption Libary
                Contact.Item = oContacts.Items.Item(oc+1)

            #print Contact.Subject

            if Contact.Class == win32com.client.constants.olContact:
                if keys is None:
                    # if we were't give a set of keys to use
                    # then build up a list of keys that we will be
                    # able to process
                    # I didn't include fields of type time, though
                    # those could probably be interpreted
                    keys = []
                    
                    if RedemptionLib:
                        # Use Outlook directly
                        for key in Contact._prop_map_get_:
                            if isinstance(getattr(Contact, key), (int, str, unicode)):
                                keys.append(key)
                    else:
                        # Use the Redemption Libary
                        for key in Contact.Item._prop_map_get_:
                            if isinstance(getattr(Contact, key), (int, str, unicode)):
                                keys.append(key)
                    
                    if DEBUG:
                        keys.sort()
                        print "Fields\n======================================"
                        for key in keys:
                            print key
                record = {}
                for key in keys:
                    record[key] = getattr(Contact, key)
                if DEBUG:
                    print oc, record['FullName'], record['Size']

            self.records.append(record)


if __name__ == '__main__':
    if DEBUG:
        print "Attempting to load Outlook .."
        
    oOutlook = MSOutlook()
    # delayed check for Outlook on win32 box
    if not oOutlook.outlookFound:
        print "Outlook not found!"
        sys.exit(1)

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
                'Size',
                'Body'
                ]

    if DEBUG:
        import time
        print "loading records..."
        startTime = time.time()

    # you can either get all of the data fields
    # or just a specific set of fields which is much faster
    #oOutlook.loadContacts()
    oOutlook.loadContacts(fields)

    if DEBUG:
        print "loading took %f seconds" % (time.time() - startTime)

    print "Number of contacts: %d\n" % len(oOutlook.records)

    print "Contact: %s" % oOutlook.records[0]['FullName']
    print "Size:%s" % oOutlook.records[0]['Size']
    print "Body:\n%s" % oOutlook.records[0]['Body']


    print "Contact: %s" % oOutlook.records[1]['FullName']
    print "Size:%s" % oOutlook.records[1]['Size']
    print "Body:\n%s" % oOutlook.records[1]['Body']

