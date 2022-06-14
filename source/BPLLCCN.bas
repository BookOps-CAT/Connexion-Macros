'MacroName:CBH_LCCN
'MacroDescription: Create a Brooklyn Public Library call number for CBH records
   
   'LC Call number created by Miriam Gloger version 1.0 (2022-06-10)
   'Copy Call number from 050 to 099 Call number for Brooklyn Public Library CBH 
   'Adds cataloger's initials into the 947 tag and 949 command line 

Option Explicit

Sub Main
   Dim CS as Object
   On Error Resume Next
   Set CS  = GetObject(,"Connex.Client")
   On Error GoTo 0
   If CS  Is Nothing Then
      Set CS  = CreateObject("Connex.Client")
   End If
   
   If CS.ItemType = 0 Or CS.ItemType = 1 or CS.ItemType = 2 or CS.ItemType = 17 or CS.ItemType = 19 or CS.ItemType = 26 or CS.ItemType = 31 or CS.ItemType = 33 or CS.ItemType = 35 Then
   
  Dim sLccall as String
  Dim sCall as String 
  Dim sCutter as String 
  Dim nPlace1 as Integer
  Dim nPlace2 as Integer
  Dim sData as String
  Dim nBool as Integer
  Dim filenumber as Integer
  Dim sFileName as String
  Dim sInitials as String
  
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
      If InStr(sLccall, Chr(223) & "b") <> 0 Then
         sLccall = Mid(sLccall, 6)
         nPlace1 = InStr(sLccall, Chr(223))
         sCall = Rtrim(Left(sLccall, nPlace1-1))
         nPlace2=InStr(sLccall, Chr(223) & "b")
         sCutter=Ltrim(Mid(sLccall, nPlace2+2))
      CS.SetField 1, "099  " & Chr(223) & "a " & sCall & Chr(223) & "a " & sCutter
      
      Else 
      If Len(sCutter) = 0 Then
         CS.SetField 2, "099  " & Chr(252)
         MsgBox "Please supply valid Library of Congress Call Number"
      
      End If
      
      CS.Reformat     
      CS.SetField 1, "947  " & sInitials
      CS.SetField 1, "949  *recs=cbh;ov=."
      End If

   End If
   
   Else
      If nBool = False Then    
         CS.SetField 2, "099  " & "sLccall" & Chr(252)
         MsgBox "Please supply valid Library of Congress Call Number"
      
         CS.SetField 1, "947  " & sInitials
         CS.SetField 1, "949  *recs=cbh;ov=." 
      End If
   End If
   'End If
     CS.Reformat
                
        
End Sub