'MacroName:CATupdater
'MacroDescription:Synchronizes local macro books with those on the shared drive
'Macro written by: Joel Hahn, Niles Public Library District
'Macro modified by: Tomasz Kalata, BookOps
'Last modified: October 4, 2021

'v2.0 changes (10/4/2021):
'  * remote folder moved to S:\CATAL\Connex\macros\"
'  * support for Connexion client 2.63 & 3.0
'  * Improved updates messages

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
   RemoteFolder = "S:\CATAL\Connex\macros\"
   If Dir("C:\Program Files (x86)\OCLC\Connexion\Program\Program\") <> "" Then
      LocalFolder = "C:\Program Files (x86)\OCLC\Connexion\Program\Macros\"
   Else
      LocalFolder = Environ("APPDATA") & "\OCLC\Connex\Macros\"
   End If
    
   sFName = Dir(RemoteFolder & "*.mbk")
   If Len(sFName) = 0 Then
      MsgBox "Can't connect to the shared folder at: " & RemoteFolder & ", or the shared folder has been emptied. Exiting..."
      Goto Done
   End If
   ReDim DirArr(0)
   nCount = 0
   Do While sFName <> ""
      If InStr("Bookops.mbk", sFName) = 0 And InStr("newMacros.mbk", SFName) = 0 Then
         ReDim Preserve DirArr(nCount)
         DirArr(nCount) = sFName
         nCount = nCount + 1
         sFName = Dir
      End If
   Loop 

   For i = 0 to nCount - 1
      sFName = DirArr(i)
      RemoteMod = FileDateTime(RemoteFolder & sFName)
      If Dir(LocalFolder & sFName) <> "" Then
         LocalMod = FileDateTime(LocalFolder & sFName)
      Else
         FileCopy RemoteFolder & sFName, LocalFolder & sFName
         MsgBox "Copying new macrobook:" & sFName & ". Click OK to continue."
         Synced = TRUE
      End If
      If RemoteMod > LocalMod Then
         'MsgBox sFName & Chr(10) & "Local: |" & LocalMod & "|" & Chr(10) & "Remote: |" & RemoteMod & "|"
         FileCopy RemoteFolder & sFName, LocalFolder & sFName
         'MsgBox "FileCopy " & RemoteFolder & sFName & ", " & LocalFolder & sFName
         MsgBox "Copying newer version of: " & sFName & ". Click OK to continue."
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

