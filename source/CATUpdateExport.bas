'MacroName:UpdateExport
'MacroDescription:Updates OCLC holdings then exports a bibliographic record.
'Version: 1.10

'v1.10 (2024-10-10)
'  * adds export protection for incomplete call numbers that include a fill character (Chr(252))
'  * adds barcode validation for NYPL RL and BL records
'v1.9 (2024-06-20)
' * permits 653 for SCIPIO records (042 scipio)
'v1.8 (2024-02-26)
' * removes short stories warning flag
'v1.7 (2023-06-23)
'  * fixes a bug causing error in records with more than 200 tags - max set to 400 fields now;
'  * additionally removes non-supported 6xx tags correctly for in longer bibs 
'v1.6 (2022-07-19)
'  * fixes a bug that caused error in records with more than 100 tags
'v1.5 (2022-05-13)
'  * permits 69x tags (used at Schomburg, and some special collections)
'v1.4 (2022-05-06)
'  * adds removal of unspported subject vocabularies from 6xx tags; permits only BIDEX, BOOKOPS, FAST, GSAFD, HOMOIT, LCGFT, and LCSH
'v1.3 (2018-07-02)
'  * added general fiction short stories warning messaged
'v1.2 (2018-02-09)
'  * added enforcement of oclcgw load table for NYPL records

Option Explicit

Declare Function PreferedLoadTable(sBLvl)
Declare Function HasCompleteCallNum(sCallNum)
Declare Function IsValidBarcode(sBarcode, sCollection)
Declare Sub CleanSubjectTags()

'################################

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

'###############################

Function HasCompleteCallNum(sCallNum)

   If InStr(sCallNum, Chr(252)) <> 0 Then
      HasCompleteCallNum = FALSE
   Else
      HasCompleteCallNum = TRUE
   End If
   
End Function

'##############################

Function IsValidBarcode(sBarcode, sCollection)

   sBarcode = Trim(sBarcode)
   IsValidBarcode = TRUE

   If Len(sBarcode) <> 14 Then
      IsValidBarcode = FALSE
   End If
   
   If sCollection = "RL" Then
      If Left(sBarcode, 4) <> "3343" Then
         IsValidBarcode = FALSE
      End If
   ElseIf sCollection = "BL" Then
      If Left(sBarcode, 4) <> "3333" Then
         IsValidBarcode = FALSE
      End If
   End If
   
   If IsNumeric(sBarcode) = FALSE Then
      IsValidBarcode = FALSE
   End If

End Function

'##############################

Sub CleanSubjectTags()

   Dim CS As Object
   Set CS  = GetObject(,"Connex.Client")

   Dim sAuthCode$, sTag$, lt$, rt$
   Dim aBool, nBool
   Dim n, place As Integer
   Dim DelArr(6 to 400) As Integer
   
   'strip unwanted MARC tags:
   'remove subject from unsupported thesauri
   
   'check & store authorization code
   aBool = CS.GetField("042", 1, sAuthCode$)
  
   n = 6
   nBool = CS.GetFieldLine(n,sTag$)
   Do While nBool = TRUE
      'MsgBox n & ", " & sTag$
      If Left(sTag$, 1) = "6" Then
         If InStr("653", Mid(sTag$, 1, 3)) <> 0 Then
            If aBool = TRUE And InStr(sAuthCode$, "scipio") <> 0 Then
               'allow SCIPIO 653s
            Else
               DelArr(n) = n
               'MsgBox "DEL 65x: " & sTag$
            End If
         ElseIf InStr("69", Mid(sTag$, 1, 2)) <> 0 Then
            'do nothing, however these tags are coded
            'MsgBox "Keep 69x: " & sTag$
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
   
   For n = 400 to 6 Step -1
      If DelArr(n) <> 0 Then
         CS.DeleteFieldLine n
      End If
   Next

End Sub

'############################

Sub Main

   Dim CS As Object
   Set CS  = GetObject(,"Connex.Client")

   Dim sBarcode, sErrorList, sValue, s949, lt, rt, sLoadCommand, sBLvl, sPreferedLoadTable, sCallNum, sCallNumTag, sCollection As String
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
            
            'check if valid barcode
            n = 1
            Do While CS.GetField("949", n, sValue)
               If Mid(sValue, 5, 1) = "1" Then
            
                  If InStr(sValue, "CATRL") <> 0 Then
                     sCollection = "RL"
                  Else
                     sCollection = "BL"
                  End If
                  
                  lt = Mid(sValue, InStr(sValue, Chr(223) & "i") + 2)
                  If InStr(lt, Chr(223)) <> 0 Then
                     sBarcode = Left(lt, InStr(lt, Chr(223)) - 1)
                  Else
                     sBarcode = lt
                  End If
              
                  If IsValidBarcode(sBarcode, sCollection) = FALSE Then
                     MsgBox "Invalid item barcode in the occurrence #" & n & " of the 949 field. Please correct and export again. Exiting..."
                     GoTo Done
                  End If
               
               End If
               n = n + 1
            Loop
            
            
            sCallNumTag = "948"
            
            'determine correct load table
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

         Else
            sCallNumTag = "099"
         End If
         
         If fieldMissing Then
            CS.AddField 1, "949  *" + sPreferedLoadTable
         End If
      
      End If
      
      'validation
      'Check if call number is correctly constructed
      CS.GetFieldUnicode sCallNumTag, 1, sCallNum
      If HasCompleteCallNum(sCallNum) = FALSE Then
         MsgBox "ERROR: Incomplete call number. Please provide missing elements in the call number (" & sCallNumTag & " field) and export the record again. Exiting..."
         GoTo Done
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
