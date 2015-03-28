#WAB Notes - Research

# Introduction #

Windows Address Book - Notes and research details

# Details #

To access WAB from python, we are using WAB Access, which gives us a COM interface into the WAB. http://wabaccess.sourceforge.net/ or http://sourceforge.net/projects/wabaccess/

Seeing that I am new to python and COM interface, I don't wish to say that there is a fault with any system, unless I have chatted with the Author of WAB Access, I will reserve the right to say I silly ...

All reference to WAB, is from what I can see with python through the COM interface exposed with WABAccess, so please forgive my the use of wrong terminology is this reguard.

Researching dumping the WAB into CSV, I wish to understand as much of the WAB structure, so that we can take advantage of as much as possible.

As I see, the WAB is build on containers, first the Identity, shared or default. Then Contacts (Elements of ID).  Then items of Contacts, which include many things like last modified time (Element ID 0x30080040), but WABAccess does not list this as a nice constant, so I had to go looking for these in MSDN, which I was able to find ...<br><a href='http://msdn.microsoft.com/en-us/library/aa454929.aspx'>http://msdn.microsoft.com/en-us/library/aa454929.aspx</a><br><br>
So, I think it's a good idea to list some of the other MSDN URL's I was able to find, so that I don't need to spend too much time in future finding them ...<br>
<br>
MAPI Properties<br>
<a href='http://msdn.microsoft.com/en-us/library/aa454438.aspx'>http://msdn.microsoft.com/en-us/library/aa454438.aspx</a><br>

Common Non-Transmittable Properties<br>
<a href='http://msdn.microsoft.com/en-us/library/ms879578.aspx'>http://msdn.microsoft.com/en-us/library/ms879578.aspx</a><br>

Common Transmittable Properties<br>
<a href='http://msdn.microsoft.com/en-us/library/ms879579.aspx'>http://msdn.microsoft.com/en-us/library/ms879579.aspx</a><br>

Mail User Properties - Most of what we use ...<br>
<a href='http://msdn.microsoft.com/en-us/library/ms879904.aspx'>http://msdn.microsoft.com/en-us/library/ms879904.aspx</a><br>
<a href='http://wabaccess.sourceforge.net/Schema.htm'>http://wabaccess.sourceforge.net/Schema.htm</a><br>

Other good links<br>
<a href='http://msdn.microsoft.com/en-us/library/ms629361%28VS.85%29.aspx'>http://msdn.microsoft.com/en-us/library/ms629361%28VS.85%29.aspx</a><br>
<a href='http://msdn.microsoft.com/en-us/library/aa155719%28office.10%29.aspx'>http://msdn.microsoft.com/en-us/library/aa155719%28office.10%29.aspx</a><br>

2010/04/10<br>
Done a check-in with basic CSV import, but have run into two problems, first, is that there are multiple e-mail address and I remember there is a procedure to import e-mail address, so I need to find and program this. Have added single e-mail address import.  Don't believe that we should need to worry about multiple e-mail address in your usage, at least not at the moment.<br>
<br>
Second problem, is a bigger problem, very much a Python thing, I believe.  Need to pair the WAB attributes with the CSV fields.  Would be nice to have an array, multi-dimensionally or multiple field descriptions ... I think maybe a dictionary, but I'm not sure how to code it up.  Hoping to ask for some help and see what help I can get.