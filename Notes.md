#General Notes

# Details #
Thinking out aloud!

First big project, working in a language that I believe is best suited for the cross-platform ideas.

I wish to sync contacts on Windows, first, then maybe Linux and move onto Mac.  Python should work nicely, breaking the project into little modules and working on them until they work as expect then merge them into the big project.

Getting to understand Google Project Hosting Site is in deal place to test many ideas.

Test Systems:<br>
- Learn Python<br>
- Find examples of code and get to understand these<br>
- Break examples and make them my own<br>
<ul><li>WAB Dump<br>
</li><li>WAB to CSV<br>
</li><li>CSV to WAB<br>
<br>
</li><li>Outlook Dump<br>
</li><li>Outlook to CSV<br>
</li><li>CSV to Outlook<br>
<br>
</li><li>MySQL Dump<br>
</li><li>MySQL to CSV<br>
</li><li>CSV to MySQL<br>
<br><br>
Still need to think about how to do the sync system. Got a few ideas and need to write them down and see how they might work or not work.<br>
<br>
Stage A<br>
Using WAB as a base client, list all contacts and their properties. Accessing the WAB via a COM interface using something called WABAccess ... Using WABAccess at <a href='http://wabaccess.sourceforge.net/'>http://wabaccess.sourceforge.net/</a> We can also access Outlook Express/Windows Address Book for syncing.</li></ul>


System can be run from source, single binary in Windows or packaged in an install system.  Run from source would need Python and it's libraries. If the system is to be run from a single binary, like in Windows, the development system would need to also install py2exe and run the exe creation script.<br>
<br>
Current requirements for this project<br>
<br>
<b>Install Python and lib's on devs system.<br>
<ul><li>Python for Windows ...</b><br>
<ul><li><a href='http://www.python.org/download/windows/'>http://www.python.org/download/windows/</a> <br>
</li><li><a href='http://www.python.org/download/releases/2.6.5'>http://www.python.org/download/releases/2.6.5</a> <br>
<br>
</li></ul></li><li>Python Win32 Extensions<br>
<ul><li><a href='http://sourceforge.net/projects/pywin32/files/'>http://sourceforge.net/projects/pywin32/files/</a> <br>
</li><li><a href='http://sourceforge.net/projects/pywin32/files/pywin32/Build%20214/pywin32-214.win32-py2.6.exe/download'>http://sourceforge.net/projects/pywin32/files/pywin32/Build%20214/pywin32-214.win32-py2.6.exe/download</a> <br></li></ul></li></ul>

<ul><li>Python py2exe<br>
<ul><li><a href='http://sourceforge.net/projects/py2exe/files/'>http://sourceforge.net/projects/py2exe/files/</a> <br>
</li><li><a href='http://sourceforge.net/projects/py2exe/files/py2exe/0.6.9/py2exe-0.6.9.win32-py2.6.exe/download'>http://sourceforge.net/projects/py2exe/files/py2exe/0.6.9/py2exe-0.6.9.win32-py2.6.exe/download</a> <br></li></ul></li></ul>

<ul><li>WABAccess via Com<br>
<ul><li><a href='http://wabaccess.sourceforge.net/'>http://wabaccess.sourceforge.net/</a></li></ul></li></ul>

<ul><li>MySQL Connector/ODBC<br>
<ul><li><a href='http://dev.mysql.com/downloads/connector/odbc/5.1.html'>http://dev.mysql.com/downloads/connector/odbc/5.1.html</a></li></ul></li></ul>


Stage B<br>
Create a DB to copy contacts into<br>
<blockquote>MS Access has some limitations, that will not easily work for syncing, will need to write extra code, that will give us this functionality on other DB's have this short coming.</blockquote>

<blockquote>MySQL seems the best solution for Remote DataBase Access.</blockquote>

Clear WAB and import DB int contacts<br>
<br>
<br>
Stage C<br>
Setup syncing system to sync to and from client to server<br>
<br>
Stage D<br>
Setup Drupal to access and manage contacts<br>
<br>
<br>
The syncing system can be used with WAB/Outlook, with some very basic mods, which I will put in later.  Would like to make the whole project as modular as possible, so that almost any Contacts/Address Book system can be sync into this system and any backend can be used.<br>
<br>
Later, syncing other parts of Outlook might be possible and worth further investigation, like Calendar, ToDo and so on.<br>
<br>
Ref<br>
Handy URLs<br>
<a href='http://www.mayukhbose.com/python/ado/what-is-ado.php'>http://www.mayukhbose.com/python/ado/what-is-ado.php</a>
<a href='http://www.ecp.cc/pyado.html'>http://www.ecp.cc/pyado.html</a>
<a href='http://www.freelance-developer.com/howto_odbcpy'>http://www.freelance-developer.com/howto_odbcpy</a>

MySQL Ref<br>
<a href='http://gusiev.com/2009/04/update-and-create-timestamps-with-mysql/'>http://gusiev.com/2009/04/update-and-create-timestamps-with-mysql/</a>

ToDo<br>
Current system is using to very special details of MySQL<br>
<blockquote>Auto Inc Field<br>
Creation Date Field<br>
Updated Date Field</blockquote>

If we create for other DB's, we will need to write code to do this client side and not in RMDB.<br>
<br>
<br>
List of All Outlook Fields <br>
<ul><li>Account<br>
</li><li>AssistantName<br>
</li><li>AssistantTelephoneNumber<br>
</li><li>AutoResolvedWinner<br>
</li><li>BillingInformation<br>
</li><li>Body<br>
</li><li>Business2TelephoneNumber<br>
</li><li>BusinessAddress<br>
</li><li>BusinessAddressCity<br>
</li><li>BusinessAddressCountry<br>
</li><li>BusinessAddressPostOfficeBox<br>
</li><li>BusinessAddressPostalCode<br>
</li><li>BusinessAddressState<br>
</li><li>BusinessAddressStreet<br>
</li><li>BusinessFaxNumber<br>
</li><li>BusinessHomePage<br>
</li><li>BusinessTelephoneNumber<br>
</li><li>CallbackTelephoneNumber<br>
</li><li>CarTelephoneNumber<br>
</li><li>Categories<br>
</li><li>Children<br>
</li><li>Class<br>
</li><li>Companies<br>
</li><li>CompanyAndFullName<br>
</li><li>CompanyLastFirstNoSpace<br>
</li><li>CompanyLastFirstSpaceOnly<br>
</li><li>CompanyMainTelephoneNumber<br>
</li><li>CompanyName<br>
</li><li>ComputerNetworkName<br>
</li><li>ConversationIndex<br>
</li><li>ConversationTopic<br>
</li><li>CustomerID<br>
</li><li>Department<br>
</li><li>DownloadState<br>
</li><li>Email1Address<br>
</li><li>Email1AddressType<br>
</li><li>Email1DisplayName<br>
</li><li>Email1EntryID<br>
</li><li>Email2Address<br>
</li><li>Email2AddressType<br>
</li><li>Email2DisplayName<br>
</li><li>Email2EntryID<br>
</li><li>Email3Address<br>
</li><li>Email3AddressType<br>
</li><li>Email3DisplayName<br>
</li><li>Email3EntryID<br>
</li><li>EntryID<br>
</li><li>FTPSite<br>
</li><li>FileAs<br>
</li><li>FirstName<br>
</li><li>FullName<br>
</li><li>FullNameAndCompany<br>
</li><li>Gender<br>
</li><li>GovernmentIDNumber<br>
</li><li>HasPicture<br>
</li><li>Hobby<br>
</li><li>Home2TelephoneNumber<br>
</li><li>HomeAddress<br>
</li><li>HomeAddressCity<br>
</li><li>HomeAddressCountry<br>
</li><li>HomeAddressPostOfficeBox<br>
</li><li>HomeAddressPostalCode<br>
</li><li>HomeAddressState<br>
</li><li>HomeAddressStreet<br>
</li><li>HomeFaxNumber<br>
</li><li>HomeTelephoneNumber<br>
</li><li>IMAddress<br>
</li><li>ISDNNumber<br>
</li><li>Importance<br>
</li><li>Initials<br>
</li><li>InternetFreeBusyAddress<br>
</li><li>IsConflict<br>
</li><li>JobTitle<br>
</li><li>Journal<br>
</li><li>Language<br>
</li><li>LastFirstAndSuffix<br>
</li><li>LastFirstNoSpace<br>
</li><li>LastFirstNoSpaceAndSuffix<br>
</li><li>LastFirstNoSpaceCompany<br>
</li><li>LastFirstSpaceOnly<br>
</li><li>LastFirstSpaceOnlyCompany<br>
</li><li>LastName<br>
</li><li>LastNameAndFirstName<br>
</li><li>MailingAddress<br>
</li><li>MailingAddressCity<br>
</li><li>MailingAddressCountry<br>
</li><li>MailingAddressPostOfficeBox<br>
</li><li>MailingAddressPostalCode<br>
</li><li>MailingAddressState<br>
</li><li>MailingAddressStreet<br>
</li><li>ManagerName<br>
</li><li>MarkForDownload<br>
</li><li>MessageClass<br>
</li><li>MiddleName<br>
</li><li>Mileage<br>
</li><li>MobileTelephoneNumber<br>
</li><li>NetMeetingAlias<br>
</li><li>NetMeetingServer<br>
</li><li>NickName<br>
</li><li>NoAging<br>
</li><li>OfficeLocation<br>
</li><li>OrganizationalIDNumber<br>
</li><li>OtherAddress<br>
</li><li>OtherAddressCity<br>
</li><li>OtherAddressCountry<br>
</li><li>OtherAddressPostOfficeBox<br>
</li><li>OtherAddressPostalCode<br>
</li><li>OtherAddressState<br>
</li><li>OtherAddressStreet<br>
</li><li>OtherFaxNumber<br>
</li><li>OtherTelephoneNumber<br>
</li><li>OutlookInternalVersion<br>
</li><li>OutlookVersion<br>
</li><li>PagerNumber<br>
</li><li>PersonalHomePage<br>
</li><li>PrimaryTelephoneNumber<br>
</li><li>Profession<br>
</li><li>RadioTelephoneNumber<br>
</li><li>ReferredBy<br>
</li><li>Saved<br>
</li><li>SelectedMailingAddress<br>
</li><li>Sensitivity<br>
</li><li>Size<br>
</li><li>Spouse<br>
</li><li>Subject<br>
</li><li>Suffix<br>
</li><li>TTYTDDTelephoneNumber<br>
</li><li>TelexNumber<br>
</li><li>Title<br>
</li><li>UnRead<br>
</li><li>User1<br>
</li><li>User2<br>
</li><li>User3<br>
</li><li>User4<br>
</li><li>UserCertificate<br>
</li><li>WebPage<br>
</li><li>YomiCompanyName<br>
</li><li>YomiFirstName<br>
</li><li>YomiLastName