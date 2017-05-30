' Created by: Tom Margrave  At Orasi Support
' File created: Wed May 17 2017
' File Name  CleanTD_80.vbs
' Original location https://github.com/TomMargrave/CleanTD_80
'******************************************************************************
'Description:
'This code is only an example and carries no warranty
'This code will do the following:
' 1. Check to see if QTP/UFT is running and wait for it to stop.
' 2. Locate the Environment Variable TEMP and verify that path exist.
' 3. Clear TD_80 folder if exists
' 4. Output results.
'******************************************************************************
'  Disclaimer

'  While this example may meet the needs of your organization, the sole responsibility
'  for modification and maintenance of the logic is yours and NOT that of the Support Organization.
'
'  The decision to use the information contained herein is done at your own risk.
'
'  The support organization is NOT responsible for any issues encountered as a
'  result of implementing all or any part of the information contained or inferred herein.
'
'  The intent of the information provided here is for educational purposes only.
'  As such, the topics in this document are only guidelines NOT a comprehensive
'  solution, as your own environment will be different.
'
'  This example DOES NOT state or in any way imply that the information
'  conveyed herein provides the solution for your environment.
'
'  The appropriate system technical resources for your enterprise should perform
'  all customization activities.
'
'  Best Practice dictates NO direct changes to be made to any production
'  environment.  It is imperative to perform and thoroughly validate ALL
'  modifications in a Test Environment.  Use the results and knowledge
'  garnered from the Test Environment experience to create a customized
'  Production Deployment Plan for your own environment.
'
'  Always ensure you have a current backup before implementing any solution.
'******************************************************************************

Dim qtApp       'As QuickTest.Application ' Declare the Application object variable
Dim WshShell
Dim blnRdy      ' Boolean to allow script to run when QTP/UFT is correct state.
Dim qtStatus    ' QTP/UFT status
Dim strPath     ' holds path to directory that the contents need to be cleaned.
Dim objFSO      ' File  system object to clean IE cache
Dim fso
Dim sTempFldr   ' holds location of the temp folder path'
Dim ofile
Dim osubfldr
Dim DelDays     ' number of days to keep'
Dim blnUFT      'determines if UFT is installed or not.'

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const strLogFile = "LogDelTD80.log"

' Number of days to keep in the TD_80 folder
DelDays = 14
Log(" ******")
Log("TD_80 Cleaner started.")
'Check to see if UFT is running'
CheckUFTRunning()

'Get the Environement Variable Temp value'
Set WshShell = CreateObject("WScript.Shell")
sTempFldr = WshShell.ExpandEnvironmentStrings("%TEMP%")
Log(" Temp Folder = " & sTempFldr)

disableAllowHPRun(true)

'####################
' DELETE Contents of a td_80 Folder and delete subfolders
'####################

If doesFolderExist(sTempFldr) Then
    strPath = sTempFldr & "\td_80"
    If doesFolderExist(strPath) Then
        DeleteFolderDate strPath, DelDays
    Else
        Log("** DOES NOT EXIST TEMP TD_80 Folder = " & strPath)
    End If
Else
    Log("** DOES NOT EXIST TEMP Folder = " & sTempFldr)
End If

disableAllowHPRun(False)

Set WshShell = nothing

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'

'**********************************************************************
' Sub Name: doesFolderExist
' Purpose:  Checks to see if file exists
' Author: Tom Margrave
' Input:
'	mFile   File path to the check
' Return: Boolean if file exist
' Prerequisites:
'**********************************************************************
Function doesFolderExist(mFile)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(mFile) Then
        doesFolderExist = True
    Else
        doesFolderExist = False
    End If
    Set objFSO = nothing
End Function

'**********************************************************************
' Sub Name:CheckUFTRunning
' Purpose: Check to see if UFT/QTP is running and wait for it to close
' Author: Tom Margrave
' Input: None
' Return: None
' Prerequisites: UFT installed on machine
'**********************************************************************
Function CheckUFTRunning()
    on Error Resume Next
    blnUFT = True
    Set qtApp = CreateObject("QuickTest.Application") ' Create the Application object

    If Err Then
        WScript.StdErr.WriteLine "error " & Err.Number
        Log("No UFT install on this machine.")
        blnUFT = False
        on error goto 0
        Exit function
    End If
    on error goto 0

    blnRdy = Not qtApp.Launched

    'Wait for QTP/UFT to be not running
    Do Until blnRdy
    	qtStatus = qtApp.GetStatus
    	Select Case qtStatus
    		Case "Not launched"    '   Not launched--UFT is not started.
    			blnRdy = true
    			'msgbox qtApp.GetStatus
    			Exit Do
    		Case "Ready"      '   Ready--UFT is idle and ready to perform operations.
    		Case "Busy"       '   Busy--UFT is currently performing an operation.
    		Case "Running"    '   Running--UFT is running a test or component.
    		Case "Recording"  '   Recording--UFT is recording.
    		Case "Waiting"    '   Waiting--UFT is waiting for user input.
    		Case "Paused"     '   Paused--The current run session is paused.
    		case Else
    	End Select
    	'Sleep for given time
    	WScript.Sleep 1000  '60000
    Loop

    Set qtApp = nothing
    ' wait for QTP to shut down
    WScript.Sleep 5000
End Function

'**********************************************************************
' Sub Name: DeleteFolderDate
' Purpose: Delete all files older than a given date recursive
' Author: Tom Margrave
' Input:
'   strFolderPath
'   DelDays
' Return: None
' Prerequisites: None
'**********************************************************************
Function DeleteFolderDate (strFolderPath, DelDays)
'http://stackoverflow.com/a/25081632'
 	Dim objFSO, objFolder
 	Set objFSO = CreateObject ("Scripting.FileSystemObject")
 	If objFSO.FolderExists(strFolderPath) Then
        killdate = date() - DelDays

        arFiles = Array()

        SelectFiles strFolderPath, killdate, arFiles, true
        nDeleted = 0

        For n = 0 to ubound(arFiles)
            on error resume next 'in case of 'in use' files...
            arFiles(n).delete true
            If err.number = 0 Then
             nDeleted = nDeleted + 1
            End If
            on error goto 0
        next
        Log("Total files deleted : "  & nDeleted)
    Else
        Log("Unable to find path: " & strFolderPath)
 	End If
 	Set objFSO = Nothing
End Function

'**********************************************************************
' Function Name: SelectFiles
' Purpose: Select files for deletion and put them in Array
' Author: Tom Margrave
' Input:
'  sPath
'  vKillDate
'  arFilesToKill
'  bIncludeSubFolders
' Return: Array of items to delete
' Prerequisites:
'From https://stackoverflow.com/questions/25081252/how-to-delete-all-files-in-a-directory-tree-older-than-10-days-using-a-vbs-scrip'
'**********************************************************************
Sub SelectFiles(sPath,vKillDate,arFilesToKill,bIncludeSubFolders)
    Set objFSO = CreateObject ("Scripting.FileSystemObject")
  Set folder = objFSO.getfolder(sPath)
  Set files = folder.files

  For each file in files
    dtLastAccessed = null

    'using last accessed to keep files that is being used.
    dtLastAccessed = file.DateLastAccessed

    'below is example if last accessed does not work for you.'
    'dtLastAccessed = file.datelastmodified

    If not isnull(dtLastAccessed) Then
      If dtLastAccessed < vKillDate Then
        count = ubound(arFilesToKill) + 1
        redim preserve arFilesToKill(count)
        Set arFilesToKill(count) = file
      End If
    End If
  next

  If bIncludeSubFolders Then
    For each fldr in folder.subfolders
      SelectFiles fldr.path,vKillDate,arFilesToKill,true
    next
  End If
End Sub

'**********************************************************************
'  Function Name: Log
'  Purpose: Creates log file of last run and number of files deleted.
'  Author: Tom Margrave
'  Input:
'  strText message to be written to log file.
'  Return: None
'  Prerequisites:
'   Const strLogFile
'   Const ForAppending
'**********************************************************************
Function Log(strText)
    Set FSO = wscript.CreateObject("Scripting.FileSystemObject")
    Set TextFile = FSO.OpenTextFile(strLogFile,ForAppending,true)
    If instr(strText,"****") Then
        TextFile.WriteLine("  " )
    End If

    TextFile.WriteLine(GetDateTimeStamp & "  " & strText)
    TextFile.Close
    Set TextFile = Nothing
    Set FSO = Nothing
End Function

'**********************************************************************
' Sub Name: GetDateTimeStamp
' Purpose: Return current date to be used in time stamp
' Author: Tom Margrave
' Input: None
' Return: None
' Prerequisites: None
'**********************************************************************
Function GetDateTimeStamp()
  Dim strNow
  strNow = Now()
  GetDateTimeStamp = Year(strNow) & Pad(Month(strNow)) _
        & Pad(Day(StrNow)) & Pad(Hour(strNow)) _
        & Pad(Minute(strNow)) & Pad(Second(strNow))
End Function

'**********************************************************************
' Sub Name: Pad
' Purpose: Pad value to two digits
' Author: Tom Margrave
' Input:
'   strIn'
' Return: Two digit string
' Prerequisites: None
'**********************************************************************
Function Pad(strIn)
  Do While Len(strIn) < 2
    strIn = "0" & strIn
  Loop
  Pad = strIn
End Function

'**********************************************************************
' Function Name: disableAllowHPRun
' Purpose: Set QTP/UFT  to not allow other HP product to run
' Author: Tom Margrave
' Input:
'   blnOpt
' Return: None
' Prerequisites:
'**********************************************************************
Function disableAllowHPRun(blnOpt)
    If (blnUFT) Then
        If (blnOpt) Then
            WshShell.Regwrite "HKEY_CURRENT_USER\Software\Mercury Interactive\QuickTest Professional\MicTest\AllowTDConnect", 00000000, "REG_DWORD"
            Log("Allow other HP product to run DISABLED")
            Else
            WshShell.Regwrite "HKEY_CURRENT_USER\Software\Mercury Interactive\QuickTest Professional\MicTest\AllowTDConnect", 00000001, "REG_DWORD"
            Log("Allow other HP product to run ENABLED")
        End If
    End If
End Function
