# RDB Pronet: Automate Compliance Document Reminder Emails for Complex Compliance Process
# Excel Desktop application used with RDB Pronet (popular recruitment industry software)
# Small/Medium Business Excel Automation Application
# For compliance documents which need updating every year or regularly, for example: Enhanced DBS criminal records check certificate, passport, industry-specific training certificates

Speed up RDB Pronet compliance document workflow - filter expired documents for each candidate, and automate sending Outlook emails with attachments.

RDB Pronet is a popular recruitment industry software application.
In the small company I worked (name not relevant), I made this Excel Desktop application in VBA to compile list of expired or soon expiring compliance documents, obtained from RDB Pronet database, for each candidate, then send out automated Outlook emails, with Word letter attachments.

There is an Excel workbook here with example data, and the VBA code saved in this Workbook, names and emails auto-generated or made up in this example. Full instructions for use are in first sheet of the Workbook.

The VBA code only was under source control.

I.e. this uses the Excel, Outlook and Word object models in VBA. The email body is formatted in HTML, and the Word attachment uses the 
slightly idiosyncratic formatting code dictated by the Word object model, eg formatting can be set at the paragraph level or at individual 
character level, but not at a "word" level. This is second app I made in VBA and has much room for improvement, any constructive 
criticism welcomed :-)

I wrote this originally with some company private info in the comments, then did a fairly tortuous git interactive rebase to remove all 
references to this in the repository before uploading to Github (was on bitbucket private repo before), hence a few commit
messages may be slightly confusing.

25/3/19 - so far, I have cheated and put most of the text hard coded in the VBA module, except for standard intro text read in from Word file, as it's rather complex to format the variable parts
of each mail, where it varies according to which documents are missing for each candidate. Time was limited, it was developed while I was
doing temp admin and persuaded them to let me have a go at software project, and works fine for the environment in which it is used.
Clear separation of logic and data would be the next step; if anyone reads this and knows of a VBA module to read in a Word file with
all the native Word formatting, and convert this straight into HTML, many thanks to let me know or make pull request!

The reason for the Word attachments repeating the email body is that they may look more of a professional letter feel to the receiver, 
and be easier to print out for a normal non-technical user. An email in itself may well not be enough to get the candidate to take action,
but professional appearance may help.

Any recruitment compliance document process can get pretty complicated to administrate, and ours required saving lots of compliance documents 
within the "Stored Document" section of RDB, which frankly looks like it was never regarded a key feature, requiring about 4 separate button
presses to upload one document, and impossible to load more than one document at a time, and no kind of thumbnail view for the Stored Documents for one candidate.
On top of this, the "Attribute" section, which stores and tracks the expiry dates of each document, has no direct programmatic connection to the "Stored Documents" section.
RDB Pronet has no direct API for uploading and downloading data (ie seems impossible to bulk upload data unless you negotiate a separate
fee with RDB to do it themselves), and they turned off a previous "SQL Query Executer" option which might have enabled extracting more data fields
in queries (it might not, but since they confirmed by phone it cannot be enabled by customers now, I don't know). The GUI looks stuck at about 
Windows 98.

Since everyone uses Excel of course, and this is very much a process internal to the one office, it seemed easier to do this in VBA, and Excel has a fine GUI for this kind of use already there.

To extract lists of candidates from RDB Pronet (detailed instructions in Excel example workbook here), you do a search one Attribute at a
time, eg do a search to get a list of all candidates
whose DBS has already expired, or expires within one month from today. You then copy and paste this list into this Excel application.
Yes, with RDB search functionality, you need to get list for one Attribute at a time, and it won't let you get the actual expiry date
out in the search results, and you can't export this list, you have to copy-paste.
RDB itself warns of expired Attributes when you go individually into each record, but doesn't provide a way to consolidate this list (by
this point it's probably obvious RDB's advantages are for the sales and CRM functionality!). Our RDB had a plugin called Workflow 
designer which is supposed to stop candidates being turned to an Active status if certain documents have expired, but when the compliance
requirements got complicated, like eg candidates needing visa and foreign passport together both in date, or just biometric visa in date alone
etc it seemed even with reps coming to the office to help program the Workflow, this kind of combined condition couldn't be dealt with
exactly right, and workarounds needed in the RDB upload process.

Hence the idea to do it in Excel instead...
