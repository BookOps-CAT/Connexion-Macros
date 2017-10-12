'MacroName:UpdateExport
'MacroDescription:Updates OCLC holdings then exports a bibliographic record.
'Version: 1.1


Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")
   Dim sErrorList As String
   Dim nIndex As Integer

   nNumErrors = CS.Validate(sErrorList)
   If nNumErrors > 0 Then
      nIndex = Instr(sErrorList, "|")
      While nIndex > 0
         MsgBox "Validation error: " + Left(sErrorList, nIndex - 1)
         sErrorList = Mid(sErrorList, nIndex + 1)
         nIndex = InStr(sErrorList, "|")
      Wend
      MsgBox "Validation error: " + sErrorList
   Else
      CS.UpdateHoldings
      CS.Export
   End If

End Sub
