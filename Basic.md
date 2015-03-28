#Basic outline

# Introduction #
Try and split project into smaller systems

# Details #
Client Contacts (Outlook/WAB)<br>
Sync System<br>
Server Contacts (MySQL/PostGreSQL)<br>

Client Contacts system<br>
<blockquote>Read and write into Outlook/WAB</blockquote>

Server Contacts<br>
<blockquote>Read and write into DB</blockquote>

Sync System<br>
<blockquote>What is needed for client <code>&lt;-&gt;</code> server sync</blockquote>

<br>
Believe what might be easiest to work with is to be able to export<br>
Contacts to csv, clean out Outlook/WAB and re-import from CSV<br>
<br>
Work with an easy selected fields, this might leave open the place for<br>
bugs, but testing before hand with a large set and more eyes on the<br>
project, could quick help move through basic features.<br>
<br>
Need to create a de-duplicator. Single file run-time, so that we don't need to install python and extensions to run for end-user.<br>
<br>
Create an install system, to install and un-install cleanly.