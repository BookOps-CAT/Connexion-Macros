'MacroName:Updater
'MacroDescription:Synchronizes local macro books with those on the shared drive for both CAT and ACQ departments. Works with Connexion client
'                 2.63 and 3.0.
'Macro written by: Joel Hahn, Niles Public Library District
'Macro modified by: Tomasz Kalata, BookOps
'Last modified: October 4, 2021

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
   
   
   ' Determine where install new macros depending on Connexion version
   ' Connexion 3.0 gets preferrence
    If Dir("C:\Program Files\OCLC\Connexion\Program\") <> "" Then

      LocalFolder = Environ("APPDATA") & "\OCLC\Connex\Macros\"
      
      If Dir("S:\CATAL\Connex\macros\") <> "" Then
         RemoteFolder = "S:\CATAL\Connex\macros\"
      ElseIf Dir("S:\ACQUI\Connex3macros\") <> "" Then
         RemoteFolder = "S:\ACQUI\Connex3macros\"
      End If
      
    ElseIf Dir("C:\Program Files (x86)\OCLC\Connexion\Program\") <> "" Then
      
      LocalFolder = "C:\Program Files (x86)\OCLC\Connexion\Program\Macros\"
      
      If Dir("S:\CATAL\BookOps-Cataloging\OCLC Connexion Macros\Macro Updater\") <> "" Then
         RemoteFolder = "S:\CATAL\BookOps-Cataloging\OCLC Connexion Macros\Macro Updater\"
      ElseIf Dir("S:\ACQUI\ACQ OCLC Macro Updater\") <> "" Then
         RemoteFolder = "S:\ACQUI\ACQ OCLC Macro Updater\"
      End If
    
    End If


   sFName = Dir(RemoteFolder & "*.mbk")
   If Len(sFName) = 0 Then
      MsgBox "Can't connect to the shared folder at: " & RemoteFolder & ". Exiting..."
      Goto Done
   End If
   ReDim DirArr(0)
   nCount = 0
   Do While sFName <> ""
      If InStr("Bookops.mbk", sFName) = 0 And InStr("newMacros.mbk", SFName) = 0 Then
         ReDim Preserve DirArr(nCount)
         DirArr(nCount) = sFName
         nCount = nCount + 1
      End If
      sFName = Dir
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

