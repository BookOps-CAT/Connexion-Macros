'MacroName:SharedSave
'MacroDescription:Saves bibs in the shared drive file switching seamlessly between usual, deafault save file and
'                 shared drive file
'Macro created by: Tomasz Kalata, BookOps
'Version: v1.0
'Latest update: September 187, 2018

Option Explicit

Sub Main
   Dim CS As Object
   Dim nSaveFileNumber As Integer
   Dim nResult As Integer
   Dim nNumErrors as Integer
   Dim nIndex as Integer
   Dim sDefSaveFile As String
   Dim sShareSaveFile as String
   Dim sType As String
   Dim sLevel As String
   Dim sForm As String
   Dim sErrorList as String
   
   Set CS = GetObject(,"Connex.Client")
   
   'make sure record is displayed or highlighed in a list 
   If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 17 Then
   
      'validate record
       CS.GetFixedField "Type", sType
       CS.GetFixedField "BLvl", sLevel
       CS.GetFixedField "Form", sForm
       
       If sType <> "a" Then
         MsgBox "Attempted to save to the shared local file incorrect record type. Exiting..."
         GoTo Done
       End If
       If sLevel <> "m" Then
         MsgBox "Attempted to save to the shared local file a record with incorrect bibliographic level. Exiting..."
         GoTo Done
       End If
       If sForm <> "o" Then
         MsgBox "Attempted to save to the shared local file a record not coded to have online form. Exiting..."
         Goto Done
       End If
       
       'run standard Connexion validation
       nNumErrors = CS.Validate(sErrorList)
       If nNumErrors > 0 Then
         nIndex = Instr(sErrorList, "|")
         While nIndex > 0
            MsgBox "Validation error: " + Left(sErrorList, nIndex - 1)
            sErrorList = Mid(sErrorList, nIndex + 1)
            nIndex = InStr(sErrorList, "|")
         Wend
         MsgBox "Validation error: " + sErrorList
         Goto Done
       End If
   
      'locate path for user's default safe file
      sDefSaveFile = Mid(Environ(2), 9) + "\OCLC\Connex\DB\DefaultBib.bib.db"
      ' path to shared safe file
      sShareSaveFile = "S:\CATAL\Shared Connex Local File\NYPL_eBooks_project.bib.db"
      
      'MsgBox sDefSaveFile + Chr(10) + sShareSaveFile
      
      'change and access shared local file as default
      nResult = CS.CreateLocalFile (sShareSaveFile)
      'MsgBox nResult
      
      'save file accessed successfuly
      If nResult = 0 Or nResult = 1 Then
         'save record to newly set default local file
         nSaveFileNumber = CS.SaveToLocalFile (False, True)
         If nSaveFileNumber = -1 Then
            MsgBox "Record has been already saved."
         Else
            MsgBox "The record has been added to a local file on the shared S drive in number " + str(nSaveFileNumber)
         End If

      'error when accessing save file
      Else
         MsgBox "Unable to access shared local save file on S drive. Please contact your supervisior."
      End if
      
      'switch back to original, default user save file
      nResult = CS.CreateLocalFile (sDefSaveFile)
      If nResult = 0 or nResult = 1 Then
         GoTo Done
      Else
         MsgBox "Unable to switch back to user's original, default local file. To reverse local save file to your usual destination on your machine, please do it manually by going to File>Local File Manager>[select the default local file]>Set as Default"
      End If
      
   Else
      MsgBox "CAT!SharedSave macro: a bibliographic record must be displayed in order to use this macro."
   End If
Done:
End Sub
