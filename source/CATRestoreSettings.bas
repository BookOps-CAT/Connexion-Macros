'MacroName:RestoreSettings
'MacroDescription:Restores Connexion user settings if they were previously backed up in Documents folder

Declare Sub EraseFolder(sFolder)
Declare Sub MoveFiles(sSrcFolder, sDstFolder)

Sub Main

   Dim sConnexSettingsFolder As String
   Dim sBackupLocalFodler As String
   
   sConnexSettingsFolder = Environ("APPDATA") & "\OCLC\Connex\"
   sBackupLocalFolder = "C:\Users\" & Environ("USERNAME") & "\Documents\ConnexBackup\"
   
   Begin Dialog WarnDialog 150, 105, "Restore User Settings"
   
      Text 15, 10, 110, 40, "Warning! This operation will delete all current user settings. This operation does not restore Connexion databases or custom macros. Proceed?"
   
      OKButton 42, 60, 60, 15
      CancelButton 42, 80, 60, 15
   
   End Dialog
   
   Dim dWarn As WarnDialog
   
   On Error resume Next
   z = Dialog(dWarn)
   If Err = 102 Then Exit Sub
   
   If z = -1 Then
   
      'validate backup exists
      If Dir(sBackupLocalFolder) = "" Then
         MsgBox "Unable to find backup folder. Exiting..."
         Goto Done
      End If
      
      ' erase current settings
      Call EraseFolder(sConnexSettingsFolder)
      Call EraseFolder(sConnexSettingsFolder & "Profiles\")
      Call EraseFolder(sConnexSettingsFolder & "Profiles\NYPL\")
      Call EraseFolder(sConnexSettingsFolder & "Profiles\BPL\")
      Call EraseFolder(sConnexSettingsFolder & "Profiles\MyProfile\")
      
      'remove empty directories
      RmDir(sConnexSettingsFolder & "Profiles\NYPL\")
      RmDir(sConnexSettingsFolder & "Profiles\BPL\")
      RmDir(sConnexSettingsFolder & "Profiles\MyProfile\")
      
      ' copy files from backup
      Call MoveFiles(sBackupLocalFolder, sConnexSettingsFolder)
      Call MoveFiles(sBackupLocalFolder & "Profiles\", sConnexSettingsFolder & "Profiles\")
      
      If Dir(sBackupLocalFolder & "Profiles\NYPL\") <> "" Then
         MkDir(sConnexSettingsFolder & "Profiles\NYPL\")
         Call MoveFiles(sBackupLocalFolder & "Profiles\NYPL\", sConnexSettingsFolder & "Profiles\NYPL\")
      End If
      
      If Dir(sBackupLocalFolder & "Profiles\BPL\") <> "" Then
         MkDir(sConnexSettingsFolder & "Profiles\BPL\")
         Call MoveFiles(sBackupLocalFolder & "Profiles\BPL\", sConnexSettingsFolder & "Profiles\BPL\")
      End If
      
      If Dir(sBackupLocalFolder & "Profiles\MyProfile\") <> "" Then
         MkDir(sConnexSettingsFolder & "Profiles\MyProfile\")
         Call MoveFiles(sBackupLocalFolder & "Profiles\MyProfile\", sConnexSettingsFolder & "Profiles\MyProfile\")
      End If
      
      MsgBox "Restore completed."
      
   Else
      Goto Done
   End If


Done:
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