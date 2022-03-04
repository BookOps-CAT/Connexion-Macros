'MacroName:Updater
'MacroDescription:Synchronizes local macro books with those on the shared drive for both CAT and ACQ departments. Works with Connexion client 3.0 only
'Macro modified by: Tomasz Kalata, BookOps
'Last modified: March 4, 2022
'version 2.1 (2022-03-04): removed support for client 2.6 and bug fixes

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
   Dim ConnexDefaultDir As String
   
   Synced = FALSE
   ConnexDefaultDir = "C:\Program Files\OCLC\Connexion\Program\"
   
   ' Determine if Connexion installed in its default location
    If Dir(ConnexDefaultDir) <> "" Then
 
      'MsgBox "Found Connexion client in the default location."

      LocalFolder = Environ("APPDATA") & "\OCLC\Connex\Macros\"
      
      If Dir("S:\CATAL\Connex\macros\") <> "" Then
         RemoteFolder = "S:\CATAL\Connex\macros\"
         'MsgBox "Located macro backup directory at: " & RemoteFolder
      ElseIf Dir("S:\ACQUI\Connex3macros\") <> "" Then
         RemoteFolder = "S:\ACQUI\Connex3macros\"
         'MsgBox "Located macro backup directory at: " & RemoteFolder
      Else
         MsgBox "Unable to locate macro backup directory on your shared drive."
      End If
    
    Else
      MsgBox "Connexion client not found in its default directory: " & ConnexDefaultDir
    
    End If
    
   sFName = Dir(RemoteFolder & "*.mbk")
   'MsgBox "first: " & sFName
   If Len(sFName) = 0 Then
      MsgBox "Can't connect to the shared folder at: " & RemoteFolder & ". Exiting..."
      Goto Done
   End If
   
   ' create an array of macrobooks present in the backup directory on the shared drive
   ReDim DirArr(0)
   nCount = 0
   Do While sFName <> ""
      If InStr("Bookops.mbk", sFName) = 0 And InStr("newMacros.mbk", SFName) = 0 Then
         ReDim Preserve DirArr(nCount)
         DirArr(nCount) = sFName
         'MsgBox "added to DirArr: " & DirArr(nCount)
         nCount = nCount + 1
      End If
      sFName = Dir
      'MsgBox "next: " & sFName
   Loop
   
   For i = 0 to nCount -1
      sFName = DirArr(i)
      RemoteMod = FileDateTime(RemoteFolder & sFName)
      
      If Dir(LocalFolder & sFName) <> "" Then
         ' existing macrobook
         LocalMod = FileDateTime(LocalFolder & sFName)
      Else
         ' new macrobook
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