# put.io-uploader
vbs script that watches a folder for new files and uploads them to put.io using curl.

##Requirements

Curl is required for use. I downloaded curl for windows and placed the curl executable in C:/Windows/System32/.

##Instructions for setup

###OAUTH Key

Get an OAUTH key from put.io and put it in the top of the watch.vbs file where it has the line oauth = "" so that it reads oauth = "XXXXXX" where XXXXXX is your OAUTH key.

###Black Hole Folder

Set the path to your black hole folder by changing the drive letter in strDrive and the rest of the path in strFolder with escaped backslashes.

###Log file

Set the path to your log file with the logfile variable. 

##Run

Double click the watch.vbs file to run. Check the log file to be sure it is working correctly.
