'MacroName:ValEncoding
'MacroDescription:ValEncoding macro analizes the entire diplayed record for possible MARC-8 encoding validation and displayes
'                 a meaningful error message to catalogers to locate the offending character.
'Version: 1.0.0

Option Explicit

Sub Main
   
   Dim CS As Object
   Dim n As Integer
   Dim nBool, valid
   Dim sData As String
   
   
On Error Resume Next
   Set CS = GetObject(,"Connex.Client")
On Error GoTo 0
   If CS Is Nothing Then
      Set CS = CreateObject("Connex.Client")
   End If
   
   n = 1
   nBool = CS.GetField("...", n, sData)
   valid = TRUE
   Do While nBool = TRUE And n < 600
      
      If InStr(sData, "&#") <> 0 Then
         valid = FALSE
         MsgBox "Tag " & Left(sData, 3) & " includes incorrectly coded character. See &# code in the following string: '" & Mid(sData, 6) & "'"
         
      End If      
   
      n = n + 1
      nBool = CS.GetField("...", n, sData)
   Loop
   
   If valid = TRUE Then
      MsgBox "No character encoding issues detected" 
   End If
   
End Sub
