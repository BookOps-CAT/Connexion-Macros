'MacroName:UpdateExport
'MacroDescription:Updates OCLC holdings then exports a bibliographic record.
'Version: 1.4

'v1.4 (2022-05-06)
'  * adds removal of unspported subject vocabularies from 6xx tags; permits only BIDEX, BOOKOPS, FAST, GSAFD, HOMOIT, LCGFT, and LCSH
'v1.3 (2018-07-02)
'  * added general fiction short stories warning messaged
'v1.2 (2018-02-09)
'  * added enforcement of oclcgw load table for NYPL records

Option Explicit

Declare Function PreferedLoadTable(sBLvl)
Declare Sub CleanSubjectTags()

'###############

Function PreferedLoadTable(sBLvl)

   Dim MonoLoadTable, SerialLoadTable As String
   
   MonoLoadTable = "recs=oclcgw;"
   SerialLoadTable = "recs=oclcgws;"
   
   If InStr("bis", sBLvl) <> 0 Then
      PreferedLoadTable = SerialLoadTable
   Else
      PreferedLoadTable = MonoLoadTable
   End If

End Function

'####################

Sub CleanSubjectTags()

   Dim CS As Object
   Set CS  = GetObject(,"Connex.Client")

   Dim sTag$, lt$, rt$
   Dim nBool
   Dim n, place As Integer
   Dim DelArr(6 to 99) As Integer
   
   'strip unwanted MARC tags:
   'remove subject from unsupported thesauri
  
   n = 6
   nBool = CS.GetFieldLine(n,sTag$)
   Do While nBool = TRUE
      If Left(sTag$, 1) = "6" Then
         If InStr("653,654", Mid(sTag$, 1, 3)) <> 0 Then
            DelArr(n) = n
            'MsgBox "DEL 65x: " & sTag$
         ElseIf InStr("69", Mid(sTag$, 1, 2)) <> 0 Then
            DelArr(n) = n
            'MsgBox "DEL 69x: " & sTag$
         ElseIf InStr("600,610,611,630,648,650,651,654,655,656,657", Mid(sTag$, 1, 3)) <> 0 Then
            'LCSH
            If Mid(sTag$,5,1) = "0" Then
               'do nothing
               'MsgBox "Keep LCSH: " & sTag$ 
            ' other dictionaries
            ElseIf Mid(sTag$,5,1) = "7" Then
               If InStr(sTag$, Chr(223) & "2 gsafd") _
                  Or InStr(sTag$, Chr(223) & "2 fast") Or InStr(sTag$, Chr(223) & "2 lcsh") _
                  Or InStr(sTag$, Chr(223) & "2 bidex") Or InStr(sTag$, Chr(223) & "2 lcgft") _
                  Or InStr(sTag$, Chr(223) & "2 homoit") Or InStr(sTag$, Chr(223) & "2 bookops") Then
                     'MsgBox "Keep list: " & sTag$
                     'do nothing, go to the next one
               Else
                  'MsgBox "DEL not on list: " & sTag$
                  DelArr(n) = n
               End If
            Else
               'MsgBox "DEL other case: " & sTag$
               DelArr(n) = n
            End If
         End If
      End If
      n = n + 1
      nBool = CS.GetFieldLine(n,sTag$)
   Loop
   
   For n = 99 to 6 Step -1
      If DelArr(n) <> 0 Then
         CS.DeleteFieldLine n
      End If
   Next

End Sub

'####################

Sub Main

   Dim CS As Object
   Set CS  = GetObject(,"Connex.Client")

   Dim sErrorList, sValue, s949, lt, rt, sLoadCommand, sBLvl, sPreferedLoadTable As String
   Dim nIndex, n, nPos1, nPos2, nNumErrors As Integer
   Dim bool049, bool949, fieldMissing

   
   If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 17 Then
   
      Call CleanSubjectTags()
   
      'check if NYPL record then apply separate procedure
      bool049 = CS.GetField("049", 1, sValue)
      If bool049 = FALSE Then
         MsgBox "Library code in the 049 field is missing. Please fix before exporting."
         GoTo Done
      Else
         If InStr(sValue, "NYPP") <> 0 Then
            CS.Reformat
            CS.GetFixedField "BLvl", sBLvl
            
            'determine correct load table
            Call PreferedLoadTable(sBLvl)
            sPreferedLoadTable = PreferedLoadTable(sBLvl)

            n = 1
            fieldMissing = True
            Do While CS.GetField("949", n, sValue) And fieldMissing
               If Mid(sValue, 5, 1) = " " Then  'check second indicator to determine if correct 949 field
                  fieldMissing = False
                  'make sure the command field starts with "*"
                  If Mid(sValue, 6, 1) <> "*" Then
                     lt = Left(sValue, 5)
                     rt = Mid(sValue, 6)
                     s949 = lt + "*" + rt
                     CS.SetField n, s949
                  End If
                  
                  'verfy and correct load table
                  CS.GetField "949", n, sValue
                  
                  
                  If InStr(sValue, sPreferedLoadTable) = 0 Then
                     'replace existing load table command with prefered one
                     nPos1 = InStr(sValue, "recs")
                     If nPos1 = 0 Then
                        'load table is completely missing, add it to the end of the string
                        rt = Right(sValue, 1)
                        If rt = ";" Then
                           s949 = sValue + sPreferedLoadTable
                        Else
                           s949 = sValue + ";" + sPreferedLoadTable
                        End If 
                        CS.SetField n, s949
                     Else
                        'load table command is incorrect, replace it in the middle of the string
                        lt = Left(sValue, nPos1 - 1)
                        nPos2 = InStr(Mid(sValue, nPos1), ";")
                        If nPos2 = 0 Then
                           sLoadCommand = Mid(sValue, nPos1)
                        Else
                           sLoadCommand = Mid(sValue, nPos1, nPos2)
                        End If
                        rt = Mid(sValue, Len(lt) + Len(sLoadCommand) + 1)
                        s949 = lt + sPreferedLoadTable + rt
                        CS.SetField n, s949
                     End If

                  End If
               End If
               n = n + 1
               Loop
            End If
         
            If fieldMissing Then
               CS.AddField 1, "949  *" + sPreferedLoadTable
            End If
         
            ' temporary patch for general fiction short stories collections
            CS.GetField "948", 1, sValue
            If InStr(sValue, "808.831") <> 0 Then
               Msgbox "Effective July 1 2018, NYPL has ceased to use '808.831' for general collections of short stories. Use the 'FIC' call number instead. Your record has not been exported."
               GoTo Done
            End If
      
      End If
      
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
         'MsgBox "Updating holdings and exporting..."
         CS.UpdateHoldings
         CS.Export
       End If
    Else
      MsgBox "Bibliographic record must be displayed to launch UpdateExport macro"
    End If

Done:

End Sub
