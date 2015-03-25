' VBScript source code

intInterval = "2"
strDrive = "D:" 
strFolder = "\\Torrents\\"
strComputer = "." 
logfile = "D:\Desktop\watch.log"
oauth = ""

' Connect to WMI

Set objWMIService = GetObject( "winmgmts:" &_ 
    "{impersonationLevel=impersonate}!\\" &_ 
    strComputer & "\root\cimv2" )

' The query string

strQuery =  _
    "Select * From __InstanceOperationEvent" _
    & " Within " & intInterval _
    & " Where Targetinstance Isa 'CIM_DataFile'" _
    & " And TargetInstance.Drive='" & strDrive & "'"_
    & " And TargetInstance.Path='" & strFolder & "'"

' Execute the query

Set colEvents = _
    objWMIService. ExecNotificationQuery (strQuery) 

' The loop

Do 
    ' Wait for the next event  
    ' Get SWbemEventSource object
    ' Get SWbemObject for the target instance
    
    Set objEvent = colEvents.NextEvent()
    Set objTargetInst = objEvent.TargetInstance
    
    ' Check the class name for SWbemEventSource
    ' It cane be one of the following:
    ' - __InstanceCreationEvent
    ' - __INstanceDeletionEvent
    ' - __InstanceModificationEvent
    
    Select Case objEvent.Path_.Class 
        
        Case "__InstanceCreationEvent" 
		upload()
        Case "__InstanceDeletionEvent"  
        	logDeleted()
        Case "__InstanceModificationEvent" 
		logModified()
    End Select 

Loop

Sub upload()
	Set WshShell = WScript.CreateObject ("WScript.shell")
	curlcommand = "curl -i -F file=@""" & objTargetInst.Name & """ ""https://upload.put.io/v2/files/upload?oauth_token=" & oauth & """"

	command = "cmd /C " & curlcommand & " >> " & logfile
	echocommand = "cmd /C echo. & echo " & objTargetInst.Name & " uploaded >> " & logfile
	WshShell.run command
	WshShell.run echocommand
End Sub

Sub logDeleted()
	Set WshShell = WScript.CreateObject ("WScript.shell")
	command = "cmd /C echo. & echo " & objTargetInst.Name & " deleted >> " & logfile
	WshShell.run command
End Sub

Sub logModified()
	Set WshShell = WScript.CreateObject ("WScript.shell")
	command = "cmd /C echo. & echo " & objTargetInst.Name & " modified >> " & logfile
	WshShell.run command
End Sub
