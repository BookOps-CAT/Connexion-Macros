'MacroName:UpdateExport_EAVie
'MacroDescription: TO BE USED ONLY FOR EAVie records!
'     Cleans up 6xxs, updates OCLC holdings, then exports a bibliographic record using correct for monos or serials Sierra load table.
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

'v1.0 2025-05-26
'  * permits records without a barcode (ditital files)
'  * valides digital barcode if present
'  * validates other mandatory EAVie elements in item record (949 _ 1) such as SPEC ID and AEON eligibility


Option Explicit

Declare Sub CleanSubjectTags()
Declare Function IsValidBarcode(sBarcode)
Declare Function HasSpecId(sValue)
Declare Function SubfieldValues(sSubfield, sValue)


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

Function SubfieldValues(sSubfieldCode, s949)

   Dim temp, sValueArr As String
   
   'strip away irrelevant data & record in an array any repeated subfield values
   
   temp = s949
   
   Do While InStr(temp, Chr(223) & sSubfieldCode) <> 0
      
      'strip away first delimiter and subfield code 
      temp = Mid(temp, InStr(temp, Chr(223) & sSubfieldCode) + 2)
      
      'append to array only the value (remove any subsequent subfields) considering it could be the last subfield
      If InStr(temp, Chr(223)) <> 0 Then
         sValueArr = sValueArr & Left(temp, InStr(temp, Chr(223)) - 1) & Chr(9)
      Else
         sValueArr = sValueArr & Trim(temp) & Chr(9)
      End If
      
   Loop
   
   'clean up array / remove last tab chr
   If sValueArr <> "" Then 
      SubfieldValues = Trim(Left(sValueArr, Len(sValueArr) - 1))
   Else
      SubfieldValues = ""
   End If

End Function

'####################

Function IsValidBarcode(sBarcode)

   If Len(sBarcode) = 0 Then
      'EAVie item records do not require a barcode to be present
      IsValidBarcode = TRUE
   ElseIf Left(sBarcode, 4) = "3363" And Len(sBarcode) = 14 And IsNumeric(sBarcode) Then
      IsValidBarcode = TRUE
   Else
      IsValidBarcode = FALSE
   End If

End Function


'##############################

Function HasSpecId(sNote)

   Dim temp As String

   If InStr(sNote, "(SPEC AMI ID)") = 0 Then
      HasSpecId = FALSE
      GoTo Done
   Else
      temp = Trim(Mid(sNote, InStr(sNote, "(SPEC AMI ID)") + 13))
      
      If InStr(temp, Chr(9)) <> 0 Then
         temp = Trim(Left(temp, InStr(temp, Chr(9)) - 1))
      End If
      
      If IsNumeric(temp) Then
         HasSpecId = TRUE
      Else
         HasSpecID = FALSE
      End If

   End If

Done:
End Function

'##############################

Sub Main

   Dim CS As Object
   Set CS  = GetObject(,"Connex.Client")

   Dim s949, sErrorList, sValue As String
   Dim nIndex, n, nNumErrors As Integer

   
   If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 17 Then
   
      Call CleanSubjectTags()
      
      n = 1
      Do While CS.GetField("949", n, s949)
         If Mid(s949, 5, 1) = "1" Then
           
            'validate barcode
            sValue = SubfieldValues("i", s949)
            If IsValidBarcode(sValue) = FALSE Then
               MsgBox "Invalid item barcode in the occurrence #" & n & " of the 949 field. Please correct and export again. Exiting..."
               GoTo Done
            End If
            
            'validate SPEC AMI ID
            sValue = SubfieldValues("n", s949)
            If HasSpecId(sValue) = False Then
               MsgBox "Invalid of missing SPEC AMI ID in the occurence #" & n & " of the 949 field. Please correct and export again. Exiting..."
            GoTo Done
            End If
            
            'validate AEON note present
            sValue = SubfieldValues("n", s949)
            If InStr(sValue, "AEON eligible") = 0 Then
               MsgBox "Missing required AEON eligibility note in the occurence #" & n & "of the 949 field. Please correect and export egain. Exiting..."            
               GoTo Done
            End If
               
         End If
         n = n + 1
      Loop
      
      'display errors or update and export if okay
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
