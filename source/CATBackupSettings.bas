'MacroName:BackupSettings
'MacroDescription:Backs up Connexion user settings
'Last update: October 4, 2021

Declare Sub MoveFiles(sSrcFolder, sDstFolder)
Declare Sub EraseFolder(sFolder)

Sub Main

   Dim sConnexSettingsFolder As String
   Dim sBackupLocalFolder As String
   Dim sUserDbFolder As String
   Dim sUserProfileFolder As String


   sConnexSettingsFolder = Environ("APPDATA") & "\OCLC\Connex\"
   sBackupLocalFolder = "C:\Users\" & Environ("USERNAME") & "\Documents\ConnexBackup\"
   sUserDbFolder = sConnexSettingsFolder & "Db\"
   sUserProfileFolder = sConnexSettingsFolder & "Profiles\"
   
   
   On Error resume Next
   MkDir(sBackupLocalFolder)
   MkDir(sBackupLocalFolder & "Db")
   MkDir(sBackupLocalFolder & "Profiles")
   
   ' erase previous backup
   Call EraseFolder(sBackupLocalFolder)
   Call EraseFolder(sBackupLocalFolder & "Db")
   Call EraseFolder(sBackupLocalFolder & "Profiles\")
   Call EraseFolder(sBackupLocalFolder & "Profiles\NYPL\")
   Call EraseFolder(sBackupLocalFolder & "Profiles\BPL\")
   Call EraseFolder(sBackupLocalFolder & "Profiles\MyProfile\")
   
   ' delete empty folders
   RmDir(sBackupLocalFolder & "Profiles\NYPL\")
   RmDir(sBackupLocalFolder & "Profiles\BPL\")
   RmDir(sBackupLocalFolder & "Profiles\MyProfile\")
   
   If Dir(sUserProfileFolder & "NYPL\") <> "" Then
      MkDir(sBackupLocalFolder & "Profiles\NYPL")
   End If
   If Dir(sUserProfileFolder & "BPL\") <> "" Then
      MkDir(sBackupLocalFolder & "Profiles\BPL")
   End If
   If Dir(sUserProfileFolder & "MyProfile\") <> "" Then
      MkDir(sBackupLocalFolder & "Profiles\MyProfile")
   End If

   ' move general files
   Call MoveFiles(sConnexSettingsFolder, sBackupLocalFolder)
   
   ' move databases
   Call MoveFiles(sUserDbFolder, sBackupLocalFolder & "Db\")
   
   ' move default profiles
   Call MoveFiles(sUserProfileFolder, sBackupLocalFolder & "Profiles\")
   If Dir(sUserProfileFolder & "NYPL\") <> "" Then
      Call MoveFiles(sUserProfileFolder & "NYPL\", sBackupLocalFolder & "Profiles\NYPL\")
   End If
   
    If Dir(sUserProfileFolder & "BPL\") <> "" Then
      Call MoveFiles(sUserProfileFolder & "BPL\", sBackupLocalFolder & "Profiles\BPL\")
   End If
   
   If Dir(sUserProfileFolder & "MyProfile\") <> "" Then
      Call MoveFiles(sUserProfileFolder & "MyProfile\", sBackupLocalFolder & "Profiles\MyProfile\")
   End If
   
   MsgBox "All done! User settings have been backed up to: " & sBackupLocalFolder

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
