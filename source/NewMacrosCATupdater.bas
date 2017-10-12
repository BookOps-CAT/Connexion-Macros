'MacroName:CATupdater
'MacroDescription:Synchronizes local macro books with those on the shared drive
'Macro written by: Joel Hahn, Niles Public Library District
'Macro modified by: Tomasz Kalata, BookOps
'Last modified: February 02, 2016

Sub Main

 Dim LocalMod
   Dim RemoteMod
   Dim RemoteFolder
   Dim LocalFolder
   Dim sFName As String
   Dim nCount As Integer
   Dim DirArr()
   Dim i
   Dim Synced
  
   Synced = FALSE
   RemoteFolder = "S:\CATAL\BookOps-Cataloging\OCLC Connexion Macros\Macro Updater\"
   LocalFolder = "C:\Program Files (x86)\OCLC\Connexion\Program\Macros\"
    
   sFName = Dir(RemoteFolder & "*.mbk")
   If Len(sFName) = 0 Then
      MsgBox "Can't connect to the shared folder at: " & RemoteFolder & ", or the shared folder has been emptied. Exiting..."
      Goto Done
   End If
   ReDim DirArr(0)
   nCount = 0
   Do While sFName <>""
      ReDim Preserve DirArr(nCount)
      DirArr(nCount) = sFName
      nCount = nCount + 1
      sFName = Dir
   Loop 

   For i = 0 to nCount - 1
      sFName = DirArr(i)
      RemoteMod = FileDateTime(RemoteFolder & sFName)
      If Dir(LocalFolder & sFName) <> "" Then
      LocalMod = FileDateTime(LocalFolder & sFName)
   Else
      FileCopy RemoteFolder & sFName, LocalFolder & sFName
      MsgBox "Copying new macrobook:" & sFName
      Synced = TRUE
   End If
   If RemoteMod > LocalMod Then
      'MsgBox sFName & Chr(10) & "Local: |" & LocalMod & "|" & Chr(10) & "Remote: |" & RemoteMod & "|"
      FileCopy RemoteFolder & sFName, LocalFolder & sFName
      'MsgBox "FileCopy " & RemoteFolder & sFName & ", " & LocalFolder & sFName
      MsgBox "Found newer version of: " & sFName & " . Copying..."
      Synced = TRUE
   End If
   Next

   If Synced <> FALSE Then
      MsgBox "Your local macrobooks have been updated."
   Else
      MsgBox "Your local macrobooks are up-to-date."
   End If
Done:
End Sub