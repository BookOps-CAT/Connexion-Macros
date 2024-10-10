'MacroName:UpdateExport_Permissive
'MacroDescription: Cleans up 6xxs, updates OCLC holdings, then exports a bibliographic record using correct for monos or serials Sierra load table.
'  This more permissive UpdateExport macro preserves the following 6xx:
'  + AAT
'  + BIDEX
'  + ESTC
'  + FAST
'  + GMGPC
'  + GSAFD
'  + HOMOIT
'  + LCGFT
'  + LCSH
'  + LOBT
'  + MIGFG
'  + MIM
'  + RBBIN
'  + RBGENR
'  + RBMSCV
'  + RBPAP
'  + RBPRI
'  + RBPROV
'  + RBPUB
'  + RBTYP
'  + RDA
'  + RDAFMN
'  + RDAFNM
'  + TEPT

'v1.4 2024-10-10
'  * adds barcode validation for RL records
'v1.3 2024-06-20
'  * allows 653s if 042 authorization code has SCIPIO (art & rare books sales/auction catalogs)
'v1.2 2024-05-01
'  * expands approved thesauri to include tept (Thesaurus of Ephemera Terms)
'v1.1 2024-03-25
'  * expands list of approved thesauri to include other rare book vocabularies: estc, gmgpc, lobt, migfg,
'     mim, rbbin, rbgenr, rbmscv, rbpap, rbpri, rbprov, rbpup, rbtyp, rda, rdafmn, rdafnm
'v1.0 2024-02-26
'  * a modified UpdateExport macro preserving 6xx from lcsh, fast, gsafd, bidex, lcgft, homoit, aat

Option Explicit

Declare Function PreferedLoadTable(sBLvl)
Declare Sub CleanSubjectTags()
Declare Function IsValidBarcode(sBarcode)

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

Function IsValidBarcode(sBarcode)

   sBarcode = Trim(sBarcode)

   If Len(sBarcode) <> 14 Then
      IsValidBarcode = FALSE
   ElseIf Left(sBarcode, 4) <> "3343" Then
      IsValidBarcode = FALSE
   ElseIf IsNumeric(sBarcode) = FALSE Then
      IsValidBarcode = FALSE
   Else
      IsValidBarcode = TRUE
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
            If aBool = TRUE And InStr(sAuthCode$, "scipio") Then
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
            ' other dictionaries to include aat, estc, gmgpc, lobt, migfg, mim, rbbin, rbgenr, rbmscv, rbpap, rbpri, rbprov, rbpub, rbtyp, rda, rdafmn, rdafnm 
            ElseIf Mid(sTag$,5,1) = "7" Then
               If InStr(sTag$, Chr(223) & "2 gsafd") _
                  Or InStr(sTag$, Chr(223) & "2 fast") Or InStr(sTag$, Chr(223) & "2 lcsh") _
                  Or InStr(sTag$, Chr(223) & "2 bidex") Or InStr(sTag$, Chr(223) & "2 lcgft") _
                  Or InStr(sTag$, Chr(223) & "2 homoit") Or InStr(sTag$, Chr(223) & "2 aat") _
                  Or InStr(sTag$, Chr(223) & "2 estc") Or InStr(sTag$, Chr(223) & "2 gmgpc") _
                  Or InStr(sTag$, Chr(223) & "2 lobt") Or InStr(sTag$, Chr(223) & "2 migfg") _ 
                  Or InStr(sTag$, Chr(223) & "2 mim") Or InStr(sTag$, Chr(223) & "2 rbbin") _
                  Or InStr(sTag$, Chr(223) & "2 rbgenr") Or InStr(sTag$, Chr(223) & "2 rbmscv") _
                  Or InStr(sTag$, Chr(223) & "2 rbpap") Or InStr(sTag$, Chr(223) & "2 rbpri") _
                  Or InStr(sTag$, Chr(223) & "2 rbprov") Or InStr(sTag$, Chr(223) & "2 rbpub") _
                  Or InStr(sTag$, Chr(223) & "2 rbtyp") Or InStr(sTag$, Chr(223) & "2 rda") _
                  Or InStr(sTag$, Chr(223) & "2 rdafmn") Or InStr(sTag$, Chr(223) & "2 rdafnm") _
                  Or InStr(sTag$, Chr(223) & "2 tept") Then
                  
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

'####################

Sub Main

   Dim CS As Object
   Set CS  = GetObject(,"Connex.Client")

   Dim sBarcode, sErrorList, sValue, s949, lt, rt, sLoadCommand, sBLvl, sPreferedLoadTable As String
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
            
                  lt = Mid(sValue, InStr(sValue, Chr(223) & "i") + 2)
                  sBarcode = Left(lt, InStr(lt, Chr(223)) - 1)
              
                  If IsValidBarcode(sBarcode) = FALSE Then
                     MsgBox "Invalid item barcode in the occurrence #" & n & " of the 949 field. Please correct and export again. Exiting..."
                     GoTo Done
                  End If
               
               End If
               n = n + 1
            Loop
            
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
