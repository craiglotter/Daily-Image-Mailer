Daily Image Mailer
==================

Daily Image Mailer is program that sends a random file from a specified folder to specified email addresses using a supplied address of an available SMTP mail server. The file that is selected is sent as an attachment with the email. 

Once a file has been sent, it is removed from the input pool and moved to a sent folder. Once all files in the input folder have been removed, all the items from the sent folder are moved back to the input folder.

There are two scheduling options: The mail is sent out twice a day, at times specified by the user, or A base time is set and an fixed interval in minutes is used to send mails out based on the base time.

Created by Craig Lotter, August 2007

*********************************

Project Details:

Coded in Visual Basic .NET using Visual Studio .NET 2005
Implements concepts such as email and file manipulation.
Level of Complexity: Very Simple

*********************************

Update 20070817.02:

- Fixed control tab indexes.
- Added ability to schedule a fixed time interval for sending messages, e.g. Send a mail out every 10 minutes

*********************************

Update 20070821.03:

- Stopped from sending system file Thumbs.db out as mail attachment
- Now minimize balloon tip only appears the first time that the app is minimized

*********************************

Update 20070831.04:

- Increased fixed time interval mode's max time to 1440 minutes and min time to 1 minute
- Stopped second time setting from being activated when running in fixed time interval mode.

*********************************

Update 20070928.05:

- Changed Loading and Saving of application settings to a plain text file instead of Microsoft's My.Settings feature
- Fixed Save/Load scheduled times bug

*********************************

Update 20080130.06:

- Now allows you to force send the email outside of the scheduler
- Added a small file description to the email body that is sent out
