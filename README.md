# MacMail-Archiver
AppleScript to automate archiving email from server IMAP folders to local folders while preserving folder structure from server. A solution to not being able to copy nested mailbox folders.

1. Lets you archive emails older than: {"1 month", "2 months", "3 months", "6 months", "9 months", "1 year", "2 years", "3 years"}
2. Notifies you when it starts archiving a new mailbox
3. Notifies you every 100 messages archived
4. Gives you total archived message count

## Running the script
1. Open the script in Script Editor.
2. Set properites for your mail server.
3. Save as Application.
4. Run it and archive your email.

## Properties to set in the script
### Mail server address
property mailServer : "MAIL.YOURSERVER.COM"
### List of mailboxes to ignore on server (not archived)
property mailboxesToIgnore : {"Deleted Messages", "Drafts", "Outbox", "Trash", "Junk", "Apple Mail To Do", "Contacts", "Emailed Contacts", "Chats"}
### List of mailboxes which will be have an extenstion appended to the name when archived 
property mailboxesToRename : {"INBOX", "Sent"}
### Extension to append to renamed mailboxes
property mailboxRenameExtension : "_Archive"
### Number of times it will retry archiving an email in case of a failure
property maxFail : 3
### Log file name
property logFile : "Archive Email.log"
