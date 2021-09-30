'MacroName:NewUserDual
'MacroDescription:NewUserDual macro creates dual NYPL & BPL profiles and default settings for each.
'                 Use for onboarding new staff.
'Macro create by: Tomasz Kalata, BookOps
'Last updated: September 30, 2021

Declare Sub EraseFolder(sFolder)
Declare Sub MoveFiles(sSrcFolder, sDstFolder)

Sub Main

   Dim CS As Object
   Dim sUserProfileFolder As String
   Dim sSharedDrive As String
   
   Set CS = CreateObject("Connex.Client")
  
   sConnexFolder = Environ("APPDATA") & "\OCLC\Connex\"
   sUserProfileFolder = sConnexFolder & "Profiles\"
   sSrcFolder = "S:\CATAL\Connex\"
   
   Begin Dialog WarnDialog 150, 90, "Set Up New User"
   
      Text 15, 10, 120, 20, "Warning! This operation will delete all current user settings. Proceed?"
   
      OKButton 42, 40, 60, 15
      CancelButton 42, 60, 60, 15
   
   End Dialog
   
   Dim dWarn As WarnDialog
   
   On Error resume Next
   z = Dialog(dWarn)
   If Err = 102 Then Exit Sub
   
   If z = -1 Then
      'update macros
      MkDir(sConnexFolder & "Macros")
      Call EraseFolder(sConnexFolder & "Macros\")
      CS.RunMacro("newMacros.mbk!CATupdater")
   
      ' delete Connects.ini
      Kill(sConnexFolder & "Connects.ini")
      
      'clean up any Profiles/NYPL & Profiles/BPL folders
      Call EraseFolder(sUserProfileFolder)
      Call EraseFolder(sUserProfileFolder & "BPL\")
      Call EraseFolder(sUserProfileFolder & "NYPL\")
      Call EraseFolder(sUserProfileFolder & "MyProfile\")
      'delete default profile
      RmDir(sUserProfileFolder & "MyProfile\")
      
      'force presence of NYPL & BPL folders
      MkDir(sUserProfileFolder & "BPL")
      MkDir(sUserProfileFolder & "NYPL")

      'recreate settings from the shared drive
      FileCopy sSrcFolder & "onboarding\dual\Connects.ini", sConnexFolder & "Connects.ini"
      FileCopy sSrcFolder & "onboarding\dual\Profiles\" & "Profile.xml", sConnexFolder & "Profiles\" & "Profile.xml"
      Call MoveFiles(sSrcFolder & "onboarding\dual\Profiles\BPL\", sConnexFolder & "Profiles\BPL\")
      Call MoveFiles(sSrcFolder & "onboarding\dual\Profiles\NYPL\", sConnexFolder & "Profiles\NYPL\")
      
      MsgBox "Dual profiles created!"
      
   Else
      MsgBox "Nothing to do here :("   
   End If
   
End Sub

'#########################################
Sub EraseFolder(sFolder)
' deletes all files in given directory

   Dim sFile
   
   sFile = Dir(sFolder)
   Do While sFile <> ""
      Kill(sFolder & sFile)
      sFile = Dir(sFolder)
   Loop

End Sub

'#########################################
Sub MoveFiles(sSrcFolder, sDstFolder)
'Copies files from given source directory to given destination directory

   Dim nIndex As Integer
   Dim i As Integer
   Dim DirArr()
   Dim sFile As String
   
   sFile = Dir(sSrcFolder)
   If Len(sFile) = 0 Then
      MsgBox "Can't connect to the shared folder at: " & sSrcFolder & ", or the shared folder has been emptied. Exiting..."
      Goto Done
   End If
   
   ReDim DirArr(0)
   nIndex = 0
   Do While sFile <> ""
      ReDim Preserve DirArr(nIndex)
      DirArr(nIndex) = sFile
      nIndex = nIndex + 1
      sFile = Dir
   Loop

   For i = 0 To nIndex - 1
      sFile = DirArr(i)
      FileCopy sSrcFolder & sFile, sDstFolder & sFile
   Next

Done:
End Sub
