'MacroName:CBHLCCN
'MacroDescription: Create a Brooklyn Public Library call number for CBH records
'Created by: Miriam Gloger

'v. 1.0 (2022-06-10)
'    * copy call number from 050 to 099
'    * Adds 949 command line
'    * Adds cataloger's initials into the 947 tag


Option Explicit

Sub Main
   Dim CS as Object

   On Error Resume Next
   Set CS  = GetObject(,"Connex.Client")
   On Error GoTo 0
   
   If CS.ItemType = 0 Or CS.ItemType = 1 or CS.ItemType = 2 or CS.ItemType = 17 or CS.ItemType = 19 or CS.ItemType = 26 or CS.ItemType = 31 or CS.ItemType = 33 or CS.ItemType = 35 Then
   
      Dim sLccall As String
      Dim sCall As String 
      Dim sCutter As String 
      Dim nPlace As Integer
      Dim sData As String
      Dim nBool As Integer
      Dim filenumber As Integer
      Dim sFileName As String
      Dim sInitials As String
  
      'read default data (initials) from text file stored in macro folder
      sFileName = Mid(Environ(2), 9) + "\OCLC\Connex\Profiles\cat_data.txt"

      If Dir$ (sFileName) <> "" Then
         filenumber = FreeFile
         Open sFileName for Input As filenumber
         Line Input #filenumber, sInitials
         Close #filenumber
      Else
         'create a file with blank values if not present in the main folder
         filenumber = FreeFile
         Open sFileName For Output As filenumber
         Print #filenumber, "XXX"
         Close #filenumber
         sInitials = "XXX"
      End If
                 
  
      nBool = CS.GetField("050",1, sLccall)
   
      If nBool = True Then

         sLccall = Mid(sLccall, 6)
         
         If InStr(sLccall, Chr(223) & "b") <> 0 Then
            nPlace = InStr(sLccall, Chr(223) & "b")
            sCall = Rtrim(Left(sLccall, nPlace-1))
            sCutter = Ltrim(Mid(sLccall, nPlace+2))
            CS.SetField 1, "099  " & Chr(223) & "a " & sCall & Chr(223) & "a " & sCutter
         Else
            MsgBox "Please supply valid Library of Congress Call Number" 
            CS.SetField 1, "099  " & sLccall & Chr(223) & "a " & Chr(252)
            
         End If

      Else
         If nBool = False Then    
            CS.SetField 1, "099  " & Chr(252) & "a " & Chr(252) & "a " & Chr(252)
            MsgBox "Please supply valid Library of Congress Call Number"

         End If

      End If

  
      CS.SetField 1, "947  " & sInitials
      CS.SetField 1, "949  *recs=cbh;ov=."
      CS.Reformat
                
   End If
        
End Sub