'MacroName:BIBCOeval v. 1.0
'MacroDescription: validates displayed record based on BIBCO program requirements
'Macro created by: Tomasz Kalata, BookOps
'Latest update: July 14, 2017
'BibcoEval Macro uses code of the Generate043 macro written by Robert Bremer (bremerr@oclc.org), OCLC WorldCat Quality Division (Revised: Mar. 31, 2013)


Option Explicit

Declare Function RdaBibCoding(t040)
Declare Function Aacr2BibCoding(t040)
Declare Function Normalize(sValue)
Declare Sub DefLangCodes
Declare Sub Add043
Declare Sub UseGenerate043

Dim CS As Object
Dim sLangCode() As String
Dim s043$, sHdg$, sNewGac$
Dim sGac() As String, sGeo() As String


Sub Main

   Set CS = CreateObject("Connex.Client")
   

   If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 17 Then
      'record displayed on the screen
         
      'define variables
      Dim sBibType$, sDesc$, sElvl$, sBibcoReqElem$, sLang$, sLitForm, sSrce$, t040$, t041$, t042$
      Dim PccAuth, NatLib, ClassPres, NonLatin, nBool, Found, Required 'boolean values
      Dim sField$, sValue$, lt$
      Dim x as Integer
      
      sBibcoReqElem = "[!BIBCO MANDATORY ELEMENT!]"

         
      'parse commonly accessed fields
      
      'parse record type
      CS.GetFixedField "Type", sBibType
      
      'parse Description type
      CS.GetFixedField "Desc", sDesc
      
      'parse record level
      CS.GetFixedField "ELvl", sElvl
      
      'parse Source from fixed field
      CS.GetFixedField "Srce", sSrce
      
      'parse 040
      CS.GetField "040", 1, t040
         
      'parse 042 and determine if PCC authenticated bib
      If CS.GetField("042", 1, t042) = True Then
         If InStr(t042, "pcc") Then
            PccAuth = True
         Else
            PccAuth = False
         End If
      Else
         'not PCC Authenticated bib
         PccAuth = False
      End If
         
      'determine if national library bib
      sValue = Mid(t040, 6, 3)
      If InStr("DLC,DNLM,NLM,NLC", sValue) <> 0 Or InStr(t042, "lccopycat") <> 0 Or InStr(t042, "lcode") <> 0 Then
         If sSrce = "d" Then
            CS.SetFixedField "Srce", "c"
         End If
         NatLib = True
      Else
         NatLib = False
      End If
      
      'parse literary form if book format
      If sBibType = "a" Then
         CS.GetFixedField "LitF", sLitForm
      Else
         sLitForm = "NA"
      End If
      
      
      'validate and populate fixed fields, 040, & 042 fields according to
      'BIBCO specifications
         
      ' check if new or existing bib
      CS.QueryRecordStatus "NUM", sValue
      If sValue = "-1" Then
         ' new bibliographic record
         CS.SetFixedField "Srce", "c"
         CS.SetFixedField "ELvl", " "
         t040 = RdaBibCoding(t040)
         CS.SetField 1, t040
      Else
         'confirm to proceed if ELvl is "i"
         If LCase(sElvl) = "i" Then
            Begin Dialog Confirm 210, 85, "Confirmation Request"
            Text 6, 4, 150, 12, "If ELvl (encoding level) was originally set to 'i'"
            Text 6, 12, 200, 12, "it is NOT recommended to upgrade record to BIBCO standard,"
            Text 6, 20, 200, 12, "unless your changes are substantional to justify the upgrade."
            Text 6, 36, 200, 12, "Upgrade to BIBCO anyway?"
            OKButton 35, 55, 60, 20
            CancelButton 105, 55, 60, 20
            End Dialog
            Dim dConfirm As Confirm
            If Dialog(dConfirm) <> -1 Then
               Goto Terminate
            End If
         End If
      
        'existing bibliographic record
         If NatLib = True Then
            If LCase(sElvl) <>  " " Then
               CS.SetFixedField "ELvl", " "
            End If 
         Else
            If LCase(sElvl) <> " " Then
               CS.SetFixedField "ELvl", " "
            End If 
            CS.SetFixedField "Srce", "c"
         End If
         
         If sDesc = "a" Then
            t040 = Aacr2BibCoding(t040)
            CS.SetField 1, t040
         Else
            t040 = RdaBibCoding(t040)
            CS.SetField 1, t040
         End If
      End If
         
      'add PCC authentication code if needed
      If PccAuth = False Then
         If CS.GetField("042", 1, t042) = False Then
            t042 = "042  pcc"
         Else:
            t042 = RTrim(t042) & " " & Chr(223) & "a pcc"
         End If
         CS.SetField 1, t042
      End If

      'add 043 if applicable by running OCLC
      If CS.GetField("043", 1, sField) = False Then
         UseGenerate043
      End If
      
      
      'determine classification & add missing fields if required
      If CS.GetField("050", 1, sField) = True Or CS.GetField("082", 1, sField) = True Then
         ClassPres = True
      End If
      If ClassPres = False Then
         If sBibType = "a" Then
            CS.GetFixedField "LitF", sLitForm
            If InStr("1,f,j", sLitForm) <> 0 Then
               CS.AddField 1, "08204[FIC] " & Chr(223) & "2 23"
            End If
         Else
            CS.AddField 1, "082 " & Chr(252) & sBibcoReqElem
         End If
      End If
      
      
      'determine if Non-Latin laguages are present and if linked fields in original script are mandatory
      CS.GetFixedField "Lang", sLang
      If CS.GetField("041", 1, t041) = True Then
         sValue = t041
         sLang = sLang & "," & Mid(sValue, 6, 3)
         Do While InStr(sValue, Chr(223) & "a") <> 0
            sValue = LTrim(Mid(sValue, InStr(sValue, Chr(223) & "a") + 2))
            sLang = sLang & "," & Mid(sValue, 1, 3)
         Loop
      End If
      
      DefLangCodes
      x = 1
      Do
         If InStr(sLang, sLangCode(x)) <> 0 Then
            NonLatin = True
         End If
         x = x + 1
      Loop Until x >= 21 Or NonLatin = True
      
      If NonLatin = True Then
         
         ' verify if at least one of 245s is in Non-Latin script
         x = 1
         Found = False
         Do
            nBool = CS.GetField("245", x, sField)
            If sField = "Data contains non-latin script" Then
               Found = True
            End If
            x = x + 1
         Loop Until nBool = False Or Found = True
         If Found = False Then
            CS.AddField x, "245 " & Chr(252) & sBibcoReqElem & " -- ADD NON-LATIN LINKED FIELD"
         End If
         
         'if 246 present verify if one of them is in Non-Latin script
         x = 1
         Found = False
         Do
            nBool = CS.GetField("246", x, sField)
            If sField = "Data contains non-latin script" Then
               Found = True
            End If
            x = x + 1
         Loop Until nBool = False
         If x > 2 Then
            '246 is present
            If Found = False Then
               CS.AddField x, "246 " & Chr(252) & sBibcoReqElem & " -- ADD NON-LATIN LINKED FIELD"
            End If
         End If
         
         'if 250 present verify if one of them is in Non-Latin script
         x = 1
         Found = False
         Do
            nBool = CS.GetField("250", x, sField)
            If sField = "Data contains non-latin script" Then
               Found = True
            End If
            x = x + 1
         Loop Until nBool = False
         If x > 2 Then
            '250 is present
            If Found = False Then
               CS.AddField x, "250 " & Chr(252) & sBibcoReqElem & " -- ADD NON-LATIN LINKED FIELD"
            End If
         End If
         
         'if 260 present verify if it has linked Non-Latin field
         x = 1
         Found = False
         Required = False
         Do
            nBool = CS.GetField("260", x, sField)
            If nBool = True Then
               Required = True
            End If
            If sField = "Data contains non-latin script" Then
               Found = True
            End If
            x = x + 1
         Loop Until nBool = False
         If Found = False And Required = True Then
            CS.AddField x, "260 " & Chr(252) & sBibcoReqElem & " -- ADD NON-LATIN LINKED FIELD"
         End If
         
         'if 264 present verify if one of them is in Non-Latin script
         x = 1
         Found = False
         Required = False
         Do
            nBool = CS.GetField("264", x, sField)
            If nBool = True Then
               Required = True
            End If
            If sField = "Data contains non-latin script" Then
               Found = True
            End If
            x = x + 1
         Loop Until nBool = False
         If Found = False And Required = True Then
            CS.AddField x, "264 " & Chr(252) & sBibcoReqElem & " -- ADD NON-LATIN LINKED FIELD"
         End If
         
         'if 490 present verify if one of them is in Non-Latin script
         x = 1
         Found = False
         Do
            nBool = CS.GetField("490", x, sField)
            If sField = "Data contains non-latin script" Then
               Found = True
            End If
            x = x + 1
         Loop Until nBool = False
         If x > 2 Then
            '490 is present
            If Found = False Then
               CS.AddField x, "490 " & Chr(252) & sBibcoReqElem & " -- ADD NON-LATIN LINKED FIELD"
            End If
         End If
      
          'if 740 present verify if one of them is in Non-Latin script
         x = 1
         Found = False
         Do
            nBool = CS.GetField("740", x, sField)
            If sField = "Data contains non-latin script" Then
               Found = True
            End If
            x = x + 1
         Loop Until nBool = False
         If x > 2 Then
            '740 is present
            If Found = False Then
               CS.AddField x, "740 " & Chr(252) & sBibcoReqElem & " -- ADD NON-LATIN LINKED FIELD"
            End If
         End If 
      
      End If
      
      
      'check if 300 present
      If CS.GetField("300", 1, sField) = False Then
         CS.AddField 1, "300 " & Chr(252) & sBibcoReqElem
      End If
      
      
      'add RDA's 3xx fields if needed
      
      If sBibType = "m" Then
         If CS.GetField("347", 1, sField) = False Then
            CS.AddField 1, "347 " & Chr(252) & sBibcoReqElem 
         End If
      ElseIf sBibType = "i" Or sBibType = "j" Then
         If CS.GetField("344", 1, sField) = False Then
            CS.AddField 1, "344 " & Chr(252) & sBibcoReqElem
         End If
         If CS.GetField("347", 1, sField) = False Then
            CS.AddField 1, "347 " & Chr(252) & sBibcoReqElem
         End If
      ElseIf sBibType = "g" Then
         If CS.GetField("344", 1, sField) = False Then
            CS.AddField 1, "344 " & Chr(252) & sBibcoReqElem
         End If
         If CS.GetField("346", 1, sField) = False Then
            CS.AddField 1, "346 " & Chr(252) & sBibcoReqElem
         End If
         If CS.GetField("347", 1, sField) = False Then
            CS.AddField 1, "347 " & Chr(252) & sBibcoReqElem
         End If
         
      End If
      
      
      'evaluate if 546 is needed based on content the 041 field
      CS.GetFixedField "Lang", sLang
      sValue = Mid(t041, 6, 3)
      If sValue = sLang Then
         sValue = ""
      End If
      lt$ = Mid(t041, 9)
      Do While InStr(lt$, Chr(223) & "a") <> 0
         lt$ = LTrim(Mid(lt$, InStr(lt$, Chr(223) & "a") + 2))
         sValue = sValue & "," & Mid(lt$, 1, 3) 
      Loop
      lt$ = Mid(t041, 9)
      Do While InStr(lt$, Chr(223) & "b") <> 0
         lt$ = LTrim(Mid(lt$, InStr(lt$, Chr(223) & "b") + 2))
         sValue = sValue & "," & Mid(lt$, 1, 3) 
      Loop
      lt$ = Mid(t041, 9)
      Do While InStr(lt$, Chr(223) & "d") <> 0
         lt$ = LTrim(Mid(lt$, InStr(lt$, Chr(223) & "d") + 2))
         sValue = sValue & "," & Mid(lt$, 1, 3) 
      Loop
      lt$ = Mid(t041, 9)
      Do While InStr(lt$, Chr(223) & "e") <> 0
         lt$ = LTrim(Mid(lt$, InStr(lt$, Chr(223) & "e") + 2))
         sValue = sValue & "," & Mid(lt$, 1, 3) 
      Loop
      lt$ = Mid(t041, 9)
      Do While InStr(lt$, Chr(223) & "f") <> 0
         lt$ = LTrim(Mid(lt$, InStr(lt$, Chr(223) & "f") + 2))
         sValue = sValue & "," & Mid(lt$, 1, 3) 
      Loop
      lt$ = Mid(t041, 9)
      Do While InStr(lt$, Chr(223) & "g") <> 0
         lt$ = LTrim(Mid(lt$, InStr(lt$, Chr(223) & "g") + 2))
         sValue = sValue & "," & Mid(lt$, 1, 3) 
      Loop
      lt$ = Mid(t041, 9)
      Do While InStr(lt$, Chr(223) & "j") <> 0
         lt$ = LTrim(Mid(lt$, InStr(lt$, Chr(223) & "j") + 2))
         sValue = sValue & "," & Mid(lt$, 1, 3) 
      Loop
      
      If Len(sValue) <> 0 Then
         If CS.GetField("546", 1, sField) = False Then
            CS.AddField 1, "546 " & Chr(252) & sBibcoReqElem
         End If
      End If
      
      
      'determine if at least one 6xx is present and add Novels or Short stories genre headings if needed
      x = 1
      Found = False
      Do
         nBool = CS.GetField("600", x, sField)
         If nBool = True Then
            If Mid(sField, 4, 2) = " 0" Then
               Found = True
               GoTo Done
            End If
         End If
         x = x + 1
      Loop Until nBool = False
      x = 1
      Do
         nBool = CS.GetField("610", x, sField)
         If nBool = True Then
            If Mid(sField, 4, 2) = " 0" Then
               Found = True
               GoTo Done
            End If
         End If
         x = x + 1
      Loop Until nBool = False
      x = 1
      Do
         nBool = CS.GetField("611", x, sField)
         If nBool = True Then
            If Mid(sField, 4, 2) = " 0" Then
               Found = True
               GoTo Done
            End If
         End If
         x = x + 1
      Loop Until nBool = False
      x = 1
      Do
         nBool = CS.GetField("650", x, sField)
         If nBool = True Then
            If Mid(sField, 4, 2) = " 0" Then
               Found = True
               GoTo Done
            End If
         End If
         x = x + 1
      Loop Until nBool = False
      x = 1
      Do
         nBool = CS.GetField("651", x, sField)
         If nBool = True Then
            If Mid(sField, 4, 2) = " 0" Then
               Found = True
               GoTo Done
            End If
         End If
         x = x + 1
      Loop Until nBool = False
      x = 1
      Do
         nBool = CS.GetField("655", x, sField)
         If nBool = True Then
            If Mid(sField, 4, 2) = " 0" Then
               Found = True
               GoTo Done
            ElseIf Mid(sField, 4, 2) = " 7"  Then
               If InStr(sField, Chr(223) & "2") <> 0 Then
                  sValue = Trim(Mid(sField, InStr(sField, Chr(223) & "2")))
                  If sValue = "lcgft" Or sValue = "gsafd" Then
                     Found = True
                     GoTo Done
                  End If
               End If
            End If
         End If
         x = x + 1
      Loop Until nBool = False
      'add 655 for novels & short stories
      If Found = False Then
         If sBibType = "a" and InStr("1,f", sLitForm) <> 0 Then
            CS.AddField 1, "655 7Novels. " & Chr(223) & "2 lcgft"
         ElseIf sBibType = "a" and sLitForm = "j" Then
            CS.AddField 1, "655 7Short stories. " & Chr(223) & "2 lcgft"
         Else
            CS.AddField 1, "650 " & Chr(252) & sBibcoReqElem & " -- ADD AT LEAST ONE SUBJECT (6xx) HEADING"
         End If
      End If
 
Done:

      'remove 263
      x = 1
      Do
         nBool = CS.GetField("263", x, sField)
         If nBool = True Then
            MsgBox "trigger"
            CS.DeleteField "263", x
         End If
         x = x + 1
      Loop Until nBool = False
      'controll all headings
      If CS.ControlHeadingsAll() = True Then
         MsgBox "All controllable headings were controlled"
      Else
         MsgBox "Not able to control all headings"
      End If
   Else
      'no record on Connexion screen
      MsgBox "A bibliographic record must be displayed in order to user this macro."
   End If
   
Terminate:

End Sub


Function RdaBibCoding(t040)

   'ensures correct Desc and the 040 field for RDA record
   Dim tNormalized$, lt$, rt$
   Dim pos1, pos2
 
   
   'make sure Desc is coded as "i"
   CS.SetFixedField "Desc", "i"
   
   tNormalized = Normalize(t040)
   'make sure 040 is properly coded for rda record
   
   If InStr(tNormalized, Chr(223) & "beng" & Chr(223) & "erda") <> 0 Then
      GoTo Done
   Else
      'add or replace subfield $b
      pos1 = InStr(t040, Chr(223) & "b")
      If pos1 <> 0 Then
         'subfield $b present so replace
         lt$ = Trim(Left(t040, pos1 - 1))
         rt$ = Mid(t040, pos1 + 1)
         pos2 = InStr(rt$, Chr(223))
         If pos2 = 0 Then
            rt$ = ""
         Else
            rt$ = Trim(Mid(rt$, pos2))
         End If
         t040 = lt$ & " " & rt$
         lt$ = Left(t040, Instr(t040, Chr(223)) - 1)
         rt$ = Mid(t040, InStr(t040, Chr(223)))
         t040 = RTrim(lt$) & " " & Chr(223) & "b eng " & Trim(rt$)
         
      Else
         'no subfield $b so add
         pos1 = InStr(t040, Chr(223))
         lt$ = Trim(Left(t040, pos1 - 1))
         rt$ = Trim(Mid(t040, pos1))
         t040 = lt$ & " " & Chr(223) & "b eng " & rt$
      End If
      
      'add or replace subfield $e
      pos1 = Instr(t040, Chr(223) & "e")
      If pos1 <> 0 Then
         'remove existing $e coding
         lt$ = Trim(Left(t040, pos1 - 1))
         rt$ = Trim(Mid(t040, pos1 + 1))
         pos2 = InStr(rt$, Chr(223))
         If pos2 = 0 Then
            rt$ = ""
         Else
            rt$ = Mid(rt$, pos2)
         End If
         t040 = lt$ & " " & rt$
         
         'replace $e & instert after $b
         pos1 = InStr(t040, Chr(223) & "b")
         pos2 = InStr(Mid(t040, pos1 + 1), Chr(223))
         lt$ = RTrim(Left(t040, pos1 + pos2 - 1))
         rt$ = LTrim(Mid(t040, pos1 + pos2))
         t040 = lt$ & " " & Chr(223) & "e rda " & rt$   
      Else
         pos1 = InStr(t040, Chr(223) & "b")
         pos2 = InStr(Mid(t040, pos1 + 1), Chr(223))
         lt$ = RTrim(Left(t040, pos1 + pos2 - 1))
         rt$ = LTrim(Mid(t040, pos1 + pos2))
         t040 = lt$ & " " & Chr(223) & "e rda " & rt$
      End If

   End If
   
Done:

   RdaBibCoding = t040

End Function


Function Aacr2BibCoding(t040)

   'ensures correct Desc and the 040 field for AACR2 record
   Dim lt$, rt$
   Dim pos1%, pos2%
   
   'add or replace $b
   pos1 = InStr(t040, Chr(223) & "b")
   If pos1 <> 0 Then
      'subfield $b present so replace
      lt$ = Trim(Left(t040, pos1 - 1))
      rt$ = Mid(t040, pos1 + 1)
      pos2 = InStr(rt$, Chr(223))
      If pos2 = 0 Then
         rt$ = ""
      Else
         rt$ = Trim(Mid(rt$, pos2))
      End If
      t040 = lt$ & " " & rt$
      lt$ = Left(t040, Instr(t040, Chr(223)) - 1)
      rt$ = Mid(t040, InStr(t040, Chr(223)))
      t040 = RTrim(lt$) & " " & Chr(223) & "b eng " & Trim(rt$)
   Else
      'no subfield $b so add
      pos1 = InStr(t040, Chr(223))
      lt$ = Trim(Left(t040, pos1 - 1))
      rt$ = Trim(Mid(t040, pos1))
      t040 = lt$ & " " & Chr(223) & "b eng " & rt$
   End If
   
   Aacr2BibCoding = t040

End Function


Sub DefLangCodes
   
   'defines Non-Latin languge codes

   ReDim sLangCode(1 to 20) As String
   sLangCode(1) = "ara"
   sLangCode(2) = "arm"
   sLangCode(3) = "bel"
   sLangCode(4) = "ben"
   sLangCode(5) = "bul"
   sLangCode(6) = "chi"
   sLangCode(7) = "eth"
   sLangCode(8) = "gre"
   sLangCode(9) = "heb"
   sLangCode(10) = "jpn"
   sLangCode(11) = "kaz"
   sLangCode(12) = "kor"
   sLangCode(13) = "mac"
   sLangCode(14) = "rus"
   sLangCode(15) = "scc"
   sLangCode(16) = "syr"
   sLangCode(17) = "tam"
   sLangCode(18) = "tgk"
   sLangCode(19) = "tha"
   sLangCode(20) = "ukr"
   
   
End Sub


Function Normalize(sValue)

   Dim x as Integer
   Dim CheckChar$, tNormalized$

   'normalized passed string
   x = 1
   While x <= len(sValue):
      CheckChar = LCase(Trim(Mid(sValue, x, 1)))
      tNormalized = tNormalized & CheckChar
      x = x + 1
   Wend
   
   Normalize = tNormalized

End Function



' MacroName: Generate043
' MacroDescription: Generates field 043 based on data in 151 & 781 fields in authority records 
' or 6xx fields in bibliographic records

' Written by Robert Bremer, OCLC WorldCat Quality Division
' Email: bremerr@oclc.org
' Revised: Mar. 31, 2013

' To generate field 043, retrieve an authority record or workform (containing field 151 and/or
' 781) or a bibliographic record or workform (containing LCSH, LC Children's, MeSH, or Canadian
' subject headings), and run the macro.

' OCLC shall not be liable for any loss or damage, lost profits, loss of business, loss of or
' damage to data, downtime or unavailability, of or in connection with use of materials. OCLC
' shall have no liability for any claims arising from use of the materials, based on
' infringement of copyright, patent, trade secret or other right, libel, slander or invasion of
' privacy or claims based on errors, inaccuracies or omissions in or loss of the data.

' OCLC makes no express warranties or representations and disclaims all implied warranties with
' respect to materials as to their accuracy, merchantability or fitness for a particular purpose.
' Macros are supplied "as is."




Function GetGac(sNrmGeo$) As String
   Dim lo%, hi%, x%
   lo = 1
   hi = 909
   x = 1
   Do
      x = (lo + hi)/2  
      If sNrmGeo < sGeo(x) Then
         hi = x - 1
      Else
         lo = x + 1
      End If  
   Loop Until lo > hi or sNrmGeo = sGeo(x)
   If sNrmGeo = sGeo(x) Then
      GetGac = sGac(x)
   Else
      GetGac = "X"
   End If
End Function

Sub DefGac
   ReDim sGac(1 to 909) As String
   sGac(1) = "u-at-ac"
   sGac(2) = "f-cm--- | f-nr---"
   sGac(3) = "a-ye---"
   sGac(4) = "mr-----"
   sGac(5) = "mm-----"
   sGac(6) = "e-gr--- | a-tu---"
   sGac(7) = "a-af---"
   sGac(8) = "f------"
   sGac(9) = "fc-----"
   sGac(10) = "fe-----"
   sGac(11) = "fe-----"
   sGac(12) = "f------"
   sGac(13) = "fq-----"
   sGac(14) = "fw-----"
   sGac(15) = "ff-----"
   sGac(16) = "fh-----"
   sGac(17) = "ff----- | fw-----"
   sGac(18) = "fs-----"
   sGac(19) = "fb-----"
   sGac(20) = "fw-----"
   sGac(21) = "n-us-al"
   sGac(22) = "n-us-al"
   sGac(23) = "n-us-ak"
   sGac(24) = "pn-----"
   sGac(25) = "e-aa---"
   sGac(26) = "n-cn-ab"
   sGac(27) = "f-ae---"
   sGac(28) = "ea-----"
   sGac(29) = "n-cn-ab"
   sGac(30) = "sa-----"
   sGac(31) = "n------ | s------"
   sGac(32) = "poas---"
   sGac(33) = "aa-----"
   sGac(34) = "i------"
   sGac(35) = "sn-----"
   sGac(36) = "e-an---"
   sGac(37) = "f-ao---"
   sGac(38) = "nwxa---"
   sGac(39) = "a-cc-an"
   sGac(40) = "t------"
   sGac(41) = "t------"
   sGac(42) = "nwaq---"
   sGac(43) = "nwaq---"
   sGac(44) = "nw-----"
   sGac(45) = "nwla---"
   sGac(46) = "n-usa--"
   sGac(47) = "n-usa--"
   sGac(48) = "mr-----"
   sGac(49) = "ma-----"
   sGac(50) = "a-mk--- | a-ye---"
   sGac(51) = "ar-----"
   sGac(52) = "au-----"
   sGac(53) = "ps-----"
   sGac(54) = "r------"
   sGac(55) = "r------"
   sGac(56) = "s-ag---"
   sGac(57) = "n-us-az"
   sGac(58) = "n-us-az"
   sGac(59) = "n-us-ar"
   sGac(60) = "n-us-ar"
   sGac(61) = "a-ai--- | a-ir--- | a-ut---"
   sGac(62) = "a-ai---"
   sGac(63) = "a-ai---"
   sGac(64) = "nwaw---"
   sGac(65) = "lsai---"
   sGac(66) = "f-gh---"
   sGac(67) = "u-ac---"
   sGac(68) = "a------"
   sGac(69) = "ac-----"
   sGac(70) = "l------"
   sGac(71) = "n-us---"
   sGac(72) = "fa-----"
   sGac(73) = "u------"
   sGac(74) = "u-at---"
   sGac(75) = "u-atn--"
   sGac(76) = "u-at-ac"
   sGac(77) = "e-au---"
   sGac(78) = "e-au--- | e-hu---"
   sGac(79) = "a-aj---"
   sGac(80) = "a-aj---"
   sGac(81) = "lnaz---"
   sGac(82) = "a-iq---"
   sGac(83) = "nwbf---"
   sGac(84) = "a-ba---"
   sGac(85) = "ed-----"
   sGac(86) = "en-----"
   sGac(87) = "ln-----"
   sGac(88) = "eb-----"
   sGac(89) = "ps-----"
   sGac(90) = "a-bg---"
   sGac(91) = "nwbb---"
   sGac(92) = "nwaq---"
   sGac(93) = "r------"
   sGac(94) = "f-lo---"
   sGac(95) = "e-ne---"
   sGac(96) = "n-cn-bc"
   sGac(97) = "r------"
   sGac(98) = "f-bs---"
   sGac(99) = "a-cc-pe"
   sGac(100) = "e-bw---"
   sGac(101) = "f-cg---"
   sGac(102) = "e-be---"
   sGac(103) = "ncbh---"
   sGac(104) = "el-----"
   sGac(105) = "ab-----"
   sGac(106) = "f-dm---"
   sGac(107) = "s-gy---"
   sGac(108) = "pn-----"
   sGac(109) = "pn-----"
   sGac(110) = "lnbm---"
   sGac(111) = "ln-----"
   sGac(112) = "a-bt---"
   sGac(113) = "f-nr---"
   sGac(114) = "mb-----"
   sGac(115) = "n-usa--"
   sGac(116) = "a-ccp--"
   sGac(117) = "s-bo---"
   sGac(118) = "nwbn---"
   sGac(119) = "a-bn---"
   sGac(120) = "e-bn---"
   sGac(121) = "ln-----"
   sGac(122) = "f-bs---"
   sGac(123) = "lsbv---"
   sGac(124) = "s-bl---"
   sGac(125) = "f-mw---"
   sGac(126) = "n-cn-bc"
   sGac(127) = "s-gy---"
   sGac(128) = "ncbh---"
   sGac(129) = "i-bi---"
   sGac(130) = "e-uk--- | e-ie---"
   sGac(131) = "a-pp---"
   sGac(132) = "pobp---"
   sGac(133) = "f-so---"
   sGac(134) = "nwvb---"
   sGac(135) = "a-bx---"
   sGac(136) = "e-bu---"
   sGac(137) = "f-uv---"
   sGac(138) = "a-br---"
   sGac(139) = "f-bd---"
   sGac(140) = "e-bw---"
   sGac(141) = "mm-----"
   sGac(142) = "n-us-ca"
   sGac(143) = "n-us-ca"
   sGac(144) = "a-cb---"
   sGac(145) = "f-cm---"
   sGac(146) = "f-cm---"
   sGac(147) = "n-cn---"
   sGac(148) = "n-cn-bc | n-cnp--"
   sGac(149) = "nccz---"
   sGac(150) = "lnca---"
   sGac(151) = "lncv---"
   sGac(152) = "cc-----"
   sGac(153) = "nwbn--- | nweu--- | nwsd---"
   sGac(154) = "cc-----"
   sGac(155) = "poci---"
   sGac(156) = "ee-----"
   sGac(157) = "ak-----"
   sGac(158) = "e-urk--"
   sGac(159) = "e-urr--"
   sGac(160) = "a-ai--- | a-aj--- | a-gs---"
   sGac(161) = "nwcj---"
   sGac(162) = "a-io--- | a-my--- | a-ph---"
   sGac(163) = "f-cx---"
   sGac(164) = "f-cx---"
   sGac(165) = "nc-----"
   sGac(166) = "u-atc--"
   sGac(167) = "e-urc--"
   sGac(168) = "a-ce---"
   sGac(169) = "f-cd---"
   sGac(170) = "fq-----"
   sGac(171) = "e-uk-ui"
   sGac(172) = "n-usl--"
   sGac(173) = "n-mx--- | n-ust-- | n-usu--"
   sGac(174) = "s-cl---"
   sGac(175) = "a-cc---"
   sGac(176) = "a-cc-cq"
   sGac(177) = "i-xa---"
   sGac(178) = "r------"
   sGac(179) = "i-xb---"
   sGac(180) = "q------"
   sGac(181) = "n-us-co"
   sGac(182) = "s-ck---"
   sGac(183) = "n-us-co"
   sGac(184) = "n-usp-- | n-mx---"
   sGac(185) = "n-cn-bc | n-us-or | n-us-wa"
   sGac(186) = "n-cn-bc | n-usp--"
   sGac(187) = "b------"
   sGac(188) = "h------"
   sGac(189) = "e-ur--- | ee-----"
   sGac(190) = "i-cq---"
   sGac(191) = "i-cq---"
   sGac(192) = "n-usu--"
   sGac(193) = "f-cg---"
   sGac(194) = "fg-----"
   sGac(195) = "fc-----"
   sGac(196) = "f-cf---"
   sGac(197) = "n-us-ct"
   sGac(198) = "n-us-ct"
   sGac(199) = "pocw---"
   sGac(200) = "u------"
   sGac(201) = "u-cs---"
   sGac(202) = "nccr---"
   sGac(203) = "f-iv---"
   sGac(204) = "e-ci---"
   sGac(205) = "nwcu---"
   sGac(206) = "n-usu--"
   sGac(207) = "nwco---"
   sGac(208) = "a-cy---"
   sGac(209) = "e-xr---"
   sGac(210) = "e-cs---"
   sGac(211) = "f-dm---"
   sGac(212) = "n-us-nd | n-us-sd"
   sGac(213) = "eo-----"
   sGac(214) = "e------"
   sGac(215) = "n-us-dc"
   sGac(216) = "zd-----"
   sGac(217) = "n-us-de"
   sGac(218) = "n-us-de"
   sGac(219) = "s-gy---"
   sGac(220) = "e-dk---"
   sGac(221) = "dd-----"
   sGac(222) = "d------"
   sGac(223) = "n-us-dc"
   sGac(224) = "f-ft---"
   sGac(225) = "nwdq---"
   sGac(226) = "nwdr---"
   sGac(227) = "a-io---"
   sGac(228) = "x------"
   sGac(229) = "f-ke---"
   sGac(230) = "ae-----"
   sGac(231) = "an-----"
   sGac(232) = "as----- | az-----"
   sGac(233) = "e-pl--- | e-ru---"
   sGac(234) = "a-em---"
   sGac(235) = "poea---"
   sGac(236) = "u-ate--"
   sGac(237) = "xa-----"
   sGac(238) = "n-usr--"
   sGac(239) = "s-ec---"
   sGac(240) = "f-ua---"
   sGac(241) = "nces---"
   sGac(242) = "e-uk-en"
   sGac(243) = "ln-----"
   sGac(244) = "w------"
   sGac(245) = "f-eg---"
   sGac(246) = "nl----- | n-cn-on | n-us---"
   sGac(247) = "f-ea---"
   sGac(248) = "s-gy---"
   sGac(249) = "e-er---"
   sGac(250) = "f-et---"
   sGac(251) = "me-----"
   sGac(252) = "e------"
   sGac(253) = "ec-----"
   sGac(254) = "ee-----"
   sGac(255) = "en-----"
   sGac(256) = "es-----"
   sGac(257) = "ew-----"
   sGac(258) = "e------"
   sGac(259) = "e------"
   sGac(260) = "lsfk---"
   sGac(261) = "lnfa---"
   sGac(262) = "am-----"
   sGac(263) = "a-ye---"
   sGac(264) = "pofj---"
   sGac(265) = "e-fi---"
   sGac(266) = "ln-----"
   sGac(267) = "n-us-fl"
   sGac(268) = "n-us-fl"
   sGac(269) = "e-ur---"
   sGac(270) = "e-yu---"
   sGac(271) = "e-fr---"
   sGac(272) = "f-cm---"
   sGac(273) = "h------"
   sGac(274) = "fq-----"
   sGac(275) = "s-fg---"
   sGac(276) = "f-gv---"
   sGac(277) = "a-ii---"
   sGac(278) = "ai-----"
   sGac(279) = "pofp---"
   sGac(280) = "f-ft---"
   sGac(281) = "f-ml---"
   sGac(282) = "f-ft---"
   sGac(283) = "a-cc-fu"
   sGac(284) = "ln-----"
   sGac(285) = "f-gv--- | f-lb--- | f-sl---"
   sGac(286) = "n-us-ga"
   sGac(287) = "f-go---"
   sGac(288) = "pogg---"
   sGac(289) = "f-gm---"
   sGac(290) = "a-cc-ka"
   sGac(291) = "awgz---"
   sGac(292) = "ln-----"
   sGac(293) = "n-us-ga"
   sGac(294) = "a-gs---"
   sGac(295) = "a-gs---"
   sGac(296) = "f-tz---"
   sGac(297) = "a-pp---"
   sGac(298) = "f-sx---"
   sGac(299) = "e-gx---"
   sGac(300) = "e-ge---"
   sGac(301) = "e-gw---"
   sGac(302) = "f-gh---"
   sGac(303) = "e-gi---"
   sGac(304) = "e-gi--- | mm-----"
   sGac(305) = "pokb--- | potv---"
   sGac(306) = "pokb---"
   sGac(307) = "f-gh---"
   sGac(308) = "n-usp--"
   sGac(309) = "e-uk---"
   sGac(310) = "nl-----"
   sGac(311) = "np-----"
   sGac(312) = "fr-----"
   sGac(313) = "e-gr---"
   sGac(314) = "n-us-wy | n-us-ut"
   sGac(315) = "n-gl---"
   sGac(316) = "ln-----"
   sGac(317) = "nwgd---"
   sGac(318) = "nwgp---"
   sGac(319) = "pogu---"
   sGac(320) = "a-cc-kn"
   sGac(321) = "a-cc-kc"
   sGac(322) = "ncgt---"
   sGac(323) = "f-gv---"
   sGac(324) = "ls-----"
   sGac(325) = "fw----- | fc----- | ls-----"
   sGac(326) = "f-pg---"
   sGac(327) = "a-cc-kw"
   sGac(328) = "n-usu--"
   sGac(329) = "s-gy---"
   sGac(330) = "a-cc-ha"
   sGac(331) = "nwht---"
   sGac(332) = "n-us-hi"
   sGac(333) = "i-hm---"
   sGac(334) = "a-cc-hp"
   sGac(335) = "a-cc-he"
   sGac(336) = "a-cc-ho"
   sGac(337) = "ah-----"
   sGac(338) = "nwhi---"
   sGac(339) = "e-ne---"
   sGac(340) = "e------"
   sGac(341) = "ncho---"
   sGac(342) = "a-cc-hk"
   sGac(343) = "ap-----"
   sGac(344) = "fh-----"
   sGac(345) = "a-cc-hh"
   sGac(346) = "n-cnh--"
   sGac(347) = "a-cc-hu"
   sGac(348) = "e-hu---"
   sGac(349) = "nl----- | n-cn-on | n-us-mi"
   sGac(350) = "e-po--- | e-sp---"
   sGac(351) = "e-ic---"
   sGac(352) = "n-us-id"
   sGac(353) = "n-us-id | n-us-mt | n-us-wy"
   sGac(354) = "f-mr---"
   sGac(355) = "n-us-il"
   sGac(356) = "n-us-il"
   sGac(357) = "n-us-in"
   sGac(358) = "a-ii---"
   sGac(359) = "i------"
   sGac(360) = "n-us-ok"
   sGac(361) = "n-us-in"
   sGac(362) = "ai-----"
   sGac(363) = "a-io---"
   sGac(364) = "a-cc-im"
   sGac(365) = "n-us-ia"
   sGac(366) = "a-ir---"
   sGac(367) = "a-iq---"
   sGac(368) = "e-ie---"
   sGac(369) = "ln-----"
   sGac(370) = "a------ | f------"
   sGac(371) = "aw----- | ff----- | e-sp---"
   sGac(372) = "mm-----"
   sGac(373) = "r------"
   sGac(374) = "l------"
   sGac(375) = "ln-----"
   sGac(376) = "i------"
   sGac(377) = "mm-----"
   sGac(378) = "po-----"
   sGac(379) = "ao-----"
   sGac(380) = "e-uk-ui"
   sGac(381) = "a-is---"
   sGac(382) = "fh-----"
   sGac(383) = "f-so---"
   sGac(384) = "e-it---"
   sGac(385) = "nwjm---"
   sGac(386) = "lnjn---"
   sGac(387) = "a-ja---"
   sGac(388) = "pn-----"
   sGac(389) = "a-is--- | awba---"
   sGac(390) = "u-at-ac"
   sGac(391) = "a-cc-ku"
   sGac(392) = "a-cc-ki"
   sGac(393) = "a-cc-kr"
   sGac(394) = "poji---"
   sGac(395) = "a-my---"
   sGac(396) = "a-jo---"
   sGac(397) = "zju----"
   sGac(398) = "f-cm---"
   sGac(399) = "n-us-ks"
   sGac(400) = "n-us-ks"
   sGac(401) = "a-kz---"
   sGac(402) = "a-kz---"
   sGac(403) = "a-my---"
   sGac(404) = "a-my---"
   sGac(405) = "n-us-ky"
   sGac(406) = "f-ke---"
   sGac(407) = "i-fs---"
   sGac(408) = "poki---"
   sGac(409) = "e-bw--- | e-ru--- | e-un---"
   sGac(410) = "a-kg---"
   sGac(411) = "pokb---"
   sGac(412) = "f-ao--- | f-cf--- | f-cg---"
   sGac(413) = "a-kr---"
   sGac(414) = "a-kn---"
   sGac(415) = "a-ko---"
   sGac(416) = "e-kv---"
   sGac(417) = "a-cck--"
   sGac(418) = "a-ku---"
   sGac(419) = "n-us-ky"
   sGac(420) = "a-kg---"
   sGac(421) = "n-us-la"
   sGac(422) = "ln-----"
   sGac(423) = "nl-----"
   sGac(424) = "a-ls---"
   sGac(425) = "e-fi--- | e-no--- | e-sw---"
   sGac(426) = "cl-----"
   sGac(427) = "aw-----"
   sGac(428) = "e-lv---"
   sGac(429) = "a-le---"
   sGac(430) = "nwaw--- | nwbn--- | nwco"
   sGac(431) = "nwli---"
   sGac(432) = "f-lo---"
   sGac(433) = "a-em--- | a-io---"
   sGac(434) = "a-cc-lp"
   sGac(435) = "f-lb---"
   sGac(436) = "f-ly---"
   sGac(437) = "e-lh---"
   sGac(438) = "poln---"
   sGac(439) = "e-li---"
   sGac(440) = "n-us-la"
   sGac(441) = "n-cn-qu"
   sGac(442) = "e-lu---"
   sGac(443) = "l------"
   sGac(444) = "a-cc-mh"
   sGac(445) = "e-bu--- | e-gr--- | e-xn---"
   sGac(446) = "e-xn---"
   sGac(447) = "f-mg---"
   sGac(448) = "lnma---"
   sGac(449) = "n-us-me"
   sGac(450) = "a-my---"
   sGac(451) = "i------"
   sGac(452) = "f-mg---"
   sGac(453) = "f-mw---"
   sGac(454) = "as-----"
   sGac(455) = "am----- | a-th---"
   sGac(456) = "am-----"
   sGac(457) = "am-----"
   sGac(458) = "a-my---"
   sGac(459) = "i-xc---"
   sGac(460) = "f-ml---"
   sGac(461) = "f-ml--- | f-sg---"
   sGac(462) = "e-mm---"
   sGac(463) = "n-cn-mb"
   sGac(464) = "n-cn-mb"
   sGac(465) = "poxd---"
   sGac(466) = "n-cnm--"
   sGac(467) = "zma----"
   sGac(468) = "poxe---"
   sGac(469) = "nwmq---"
   sGac(470) = "n-us-md"
   sGac(471) = "n-us-ma"
   sGac(472) = "n-us-ma"
   sGac(473) = "f-mu---"
   sGac(474) = "i-mf---"
   sGac(475) = "i-my---"
   sGac(476) = "n-us-md"
   sGac(477) = "n-us-me"
   sGac(478) = "mm-----"
   sGac(479) = "mm-----"
   sGac(480) = "ag-----"
   sGac(481) = "pome---"
   sGac(482) = "zme----"
   sGac(483) = "n-mx---"
   sGac(484) = "nm-----"
   sGac(485) = "n-us-mi"
   sGac(486) = "n-us-mi"
   sGac(487) = "nl----- | n-us---"
   sGac(488) = "pott---"
   sGac(489) = "pomi---"
   sGac(490) = "n-usl--"
   sGac(491) = "f-cf---"
   sGac(492) = "aw-----"
   sGac(493) = "n-usc--"
   sGac(494) = "poxf---"
   sGac(495) = "n-us-mn"
   sGac(496) = "n-us-mn"
   sGac(497) = "n-us-ms"
   sGac(498) = "n-us-ms"
   sGac(499) = "n-usm--"
   sGac(500) = "n-us-mo"
   sGac(501) = "n-uss--"
   sGac(502) = "n-us-mo"
   sGac(503) = "e-mv--- | e-rm---"
   sGac(504) = "e-mv---"
   sGac(505) = "e-mv---"
   sGac(506) = "e-mc---"
   sGac(507) = "a-mp---"
   sGac(508) = "n-us-mt"
   sGac(509) = "n-us-mt"
   sGac(510) = "e-mo---"
   sGac(511) = "nwmj---"
   sGac(512) = "zmo----"
   sGac(513) = "f-mr---"
   sGac(514) = "f-mz---"
   sGac(515) = "a-mk---"
   sGac(516) = "f-sx---"
   sGac(517) = "ponu---"
   sGac(518) = "e-fr--- | e-sp---"
   sGac(519) = "n-cn-nk"
   sGac(520) = "n-us-nc"
   sGac(521) = "n-us-nd"
   sGac(522) = "n-us-nb"
   sGac(523) = "n-us-nb"
   sGac(524) = "a-my---"
   sGac(525) = "a-np---"
   sGac(526) = "zne----"
   sGac(527) = "e-ne---"
   sGac(528) = "nwna---"
   sGac(529) = "a-io---"
   sGac(530) = "nwna---"
   sGac(531) = "n-us-nv"
   sGac(532) = "n-us-nv"
   sGac(533) = "nwxi---"
   sGac(534) = "n-cn-nk"
   sGac(535) = "ponl---"
   sGac(536) = "n-usn--"
   sGac(537) = "n------"
   sGac(538) = "a-nw---"
   sGac(539) = "n-us-nh"
   sGac(540) = "ponn---"
   sGac(541) = "n-us-nj"
   sGac(542) = "n-us-nm"
   sGac(543) = "n-us-ny | n-us-nj | n-us-ct"
   sGac(544) = "u-at-ne"
   sGac(545) = "n-us-ny"
   sGac(546) = "u-nz---"
   sGac(547) = "n-cn-nf"
   sGac(548) = "n-cn-nf"
   sGac(549) = "n-cn-nf"
   sGac(550) = "n-us-nh"
   sGac(551) = "ncnq---"
   sGac(552) = "f-ng---"
   sGac(553) = "fi-----"
   sGac(554) = "f-nr---"
   sGac(555) = "fl-----"
   sGac(556) = "a-cc-nn"
   sGac(557) = "poxh---"
   sGac(558) = "n-us-nj"
   sGac(559) = "n-cn-nf"
   sGac(560) = "n-us-nm"
   sGac(561) = "u-at---"
   sGac(562) = "n------"
   sGac(563) = "ln-----"
   sGac(564) = "ln-----"
   sGac(565) = "a-my---"
   sGac(566) = "n-us-nc"
   sGac(567) = "n-us-nd"
   sGac(568) = "pn----- | a------ | n------"
   sGac(569) = "pn-----"
   sGac(570) = "r------"
   sGac(571) = "ln-----"
   sGac(572) = "r------"
   sGac(573) = "n-use--"
   sGac(574) = "xb-----"
   sGac(575) = "e-uk-ni"
   sGac(576) = "poxd---"
   sGac(577) = "f-za---"
   sGac(578) = "u-at-no"
   sGac(579) = "r------"
   sGac(580) = "n-cn-nt"
   sGac(581) = "n-cn---"
   sGac(582) = "n-usc--"
   sGac(583) = "n-usp-- | n-cn-bc"
   sGac(584) = "n-usp--"
   sGac(585) = "e-no---"
   sGac(586) = "ln-----"
   sGac(587) = "n-cn-ns"
   sGac(588) = "n-cn-ns"
   sGac(589) = "u-at-ne"
   sGac(590) = "u-at-no"
   sGac(591) = "n-cn-nu"
   sGac(592) = "n-cn-nt"
   sGac(593) = "n-us-ny"
   sGac(594) = "fe-----"
   sGac(595) = "f-mw---"
   sGac(596) = "u-nz---"
   sGac(597) = "po-----"
   sGac(598) = "n-us-oh"
   sGac(599) = "n-uso--"
   sGac(600) = "n-us-ok"
   sGac(601) = "n-us-ok"
   sGac(602) = "a-mk---"
   sGac(603) = "au-----"
   sGac(604) = "n-cn-on"
   sGac(605) = "n-cn-on"
   sGac(606) = "nl----- | n-cn-on | n-us-ny"
   sGac(607) = "n-us-or"
   sGac(608) = "n-us-or"
   sGac(609) = "n-us-or | n-us-id | n-us-mt | n-us-wa | n-us-wy"
   sGac(610) = "a------"
   sGac(611) = "zo-----"
   sGac(612) = "n-us-ar | n-us-mo | n-us-ok"
   sGac(613) = "n-us-pa"
   sGac(614) = "p------"
   sGac(615) = "poup---"
   sGac(616) = "p------"
   sGac(617) = "n-us-ca | n-us-or | n-us-wa"
   sGac(618) = "a-my---"
   sGac(619) = "a-pk---"
   sGac(620) = "popl---"
   sGac(621) = "a-is--- | awba--- | awgz---"
   sGac(622) = "ncpn---"
   sGac(623) = "e-au--- | e-hu--- | e-yu---"
   sGac(624) = "s-bl--- | s-bo--- | s-py---"
   sGac(625) = "a-pp---"
   sGac(626) = "a-pp---"
   sGac(627) = "aopf---"
   sGac(628) = "s-py---"
   sGac(629) = "n-cn-pi"
   sGac(630) = "n-us-pa"
   sGac(631) = "a-my---"
   sGac(632) = "a-my---"
   sGac(633) = "ap-----"
   sGac(634) = "ar-----"
   sGac(635) = "s-pe---"
   sGac(636) = "pn-----"
   sGac(637) = "a-ph---"
   sGac(638) = "a-my---"
   sGac(639) = "popc---"
   sGac(640) = "zpl----"
   sGac(641) = "e-pl---"
   sGac(642) = "r------ | t------"
   sGac(643) = "pops---"
   sGac(644) = "e-po---"
   sGac(645) = "f-mz---"
   sGac(646) = "f-pg---"
   sGac(647) = "a-em---"
   sGac(648) = "nwpr---"
   sGac(649) = "n-cnp--"
   sGac(650) = "a-ch---"
   sGac(651) = "n-cn-pi"
   sGac(652) = "f-sa---"
   sGac(653) = "a-ye---"
   sGac(654) = "nwpr---"
   sGac(655) = "ep-----"
   sGac(656) = "a-qa---"
   sGac(657) = "a-cc-ts"
   sGac(658) = "u-at-qn"
   sGac(659) = "n-cn-qu"
   sGac(660) = "u-at-qn"
   sGac(661) = "mr-----"
   sGac(662) = "i-re---"
   sGac(663) = "er-----"
   sGac(664) = "n-us-ri"
   sGac(665) = "f-rh--- | f-za---"
   sGac(666) = "f-mw--- | f-rh--- | f-za---"
   sGac(667) = "n-us-ri"
   sGac(668) = "sp-----"
   sGac(669) = "nr-----"
   sGac(670) = "e-rm---"
   sGac(671) = "e------ | aw----- | ff-----"
   sGac(672) = "e-ru---"
   sGac(673) = "f-bd--- | f-rw---"
   sGac(674) = "e-ur---"
   sGac(675) = "e-ru---"
   sGac(676) = "e-ru---"
   sGac(677) = "e-urf--"
   sGac(678) = "e-ru---"
   sGac(679) = "f-rw---"
   sGac(680) = "a-ja---"
   sGac(681) = "u-at-sa"
   sGac(682) = "nwsd---"
   sGac(683) = "a-my---"
   sGac(684) = "fd-----"
   sGac(685) = "f------"
   sGac(686) = "lsxj---"
   sGac(687) = "nwxi---"
   sGac(688) = "nwxi---"
   sGac(689) = "nwxa--- | nwxi---"
   sGac(690) = "n-cn--- | n-us---"
   sGac(691) = "nwxk---"
   sGac(692) = "nwsn--- | nwst---"
   sGac(693) = "n-xl---"
   sGac(694) = "nwxm---"
   sGac(695) = "nwxm---"
   sGac(696) = "nwgp---"
   sGac(697) = "nwst---"
   sGac(698) = "pows---"
   sGac(699) = "posh---"
   sGac(700) = "e-sm---"
   sGac(701) = "f-sf---"
   sGac(702) = "a-my---"
   sGac(703) = "ln-----"
   sGac(704) = "n-cn-sn"
   sGac(705) = "n-cn-sn"
   sGac(706) = "zsa----"
   sGac(707) = "a-su---"
   sGac(708) = "n-us-sc"
   sGac(709) = "ev-----"
   sGac(710) = "ev-----"
   sGac(711) = "e-uk-st"
   sGac(712) = "n-us-sd"
   sGac(713) = "pn-----"
   sGac(714) = "a-my---"
   sGac(715) = "f-sg---"
   sGac(716) = "e-rb---"
   sGac(717) = "e-yu---"
   sGac(718) = "i-se---"
   sGac(719) = "a-cc-ss"
   sGac(720) = "a-cc-sp"
   sGac(721) = "a-cc-sm"
   sGac(722) = "a-cc-sh"
   sGac(723) = "a-th---"
   sGac(724) = "e-urs--"
   sGac(725) = "e-ure--"
   sGac(726) = "e-urw--"
   sGac(727) = "a-cc-sz"
   sGac(728) = "f-sl---"
   sGac(729) = "a-si---"
   sGac(730) = "nweu---"
   sGac(731) = "nwsn---"
   sGac(732) = "ln-----"
   sGac(733) = "ee-----"
   sGac(734) = "e-xo---"
   sGac(735) = "e-xv---"
   sGac(736) = "n-us---"
   sGac(737) = "zs-----"
   sGac(738) = "pobp---"
   sGac(739) = "f-so---"
   sGac(740) = "f-ml--- | f-ng--- | f-nr---"
   sGac(741) = "n-mx--- | n-ust--"
   sGac(742) = "f-sa---"
   sGac(743) = "s------"
   sGac(744) = "az-----"
   sGac(745) = "ls-----"
   sGac(746) = "n-usu--"
   sGac(747) = "u-at-sa"
   sGac(748) = "n-us-sc"
   sGac(749) = "ao-----"
   sGac(750) = "n-us-sd"
   sGac(751) = "lsxs---"
   sGac(752) = "lsfk---"
   sGac(753) = "ps-----"
   sGac(754) = "t------"
   sGac(755) = "f-sd---"
   sGac(756) = "as-----"
   sGac(757) = "f-cm---"
   sGac(758) = "s------"
   sGac(759) = "xc-----"
   sGac(760) = "f-rh---"
   sGac(761) = "n-usu--"
   sGac(762) = "a-ye---"
   sGac(763) = "f-sx---"
   sGac(764) = "n-ust--"
   sGac(765) = "n-usu--"
   sGac(766) = "n-ust-- | n-usu--"
   sGac(767) = "e-ur---"
   sGac(768) = "e-urn--"
   sGac(769) = "e-sp---"
   sGac(770) = "f-eg---"
   sGac(771) = "cc-----"
   sGac(772) = "f-ss---"
   sGac(773) = "f-mr--- | f-ss---"
   sGac(774) = "aoxp---"
   sGac(775) = "a-ce---"
   sGac(776) = "a-my--- | a-si---"
   sGac(777) = "f-sj---"
   sGac(778) = "f-ml---"
   sGac(779) = "fn-----"
   sGac(780) = "fu-----"
   sGac(781) = "a-my--- | a-ph---"
   sGac(782) = "zsu----"
   sGac(783) = "n-ust-- | n-usu-- | n-us-hi"
   sGac(784) = "nl----- | n-cn-on | n-us---"
   sGac(785) = "s-sr---"
   sGac(786) = "lnsb---"
   sGac(787) = "nwsv---"
   sGac(788) = "f-sq---"
   sGac(789) = "e-sw---"
   sGac(790) = "e-sz---"
   sGac(791) = "a-sy---"
   sGac(792) = "a-ch---"
   sGac(793) = "a-ta---"
   sGac(794) = "a-ta---"
   sGac(795) = "f-tz---"
   sGac(796) = "fe-----"
   sGac(797) = "f-tz---"
   sGac(798) = "u-at-tm"
   sGac(799) = "ps-----"
   sGac(800) = "u-at-tm"
   sGac(801) = "n-us-tn"
   sGac(802) = "n-us-tn"
   sGac(803) = "a-my---"
   sGac(804) = "i-fs---"
   sGac(805) = "n-us-tx"
   sGac(806) = "n-us-tx"
   sGac(807) = "a-th---"
   sGac(808) = "af-----"
   sGac(809) = "a-cc-tn"
   sGac(810) = "f-cd--- | f-ly--- | f-ng---"
   sGac(811) = "a-cc-ti"
   sGac(812) = "ac-----"
   sGac(813) = "ac-----"
   sGac(814) = "at-----"
   sGac(815) = "a-em--- | a-io---"
   sGac(816) = "i------"
   sGac(817) = "a-em---"
   sGac(818) = "nwtr---"
   sGac(819) = "f-tg---"
   sGac(820) = "f-gh--- | f-tg---"
   sGac(821) = "potl---"
   sGac(822) = "potl---"
   sGac(823) = "poto---"
   sGac(824) = "a-vt---"
   sGac(825) = "a-pp--- | u-at---"
   sGac(826) = "a-jo---"
   sGac(827) = "nwtr---"
   sGac(828) = "nwtr---"
   sGac(829) = "lstd---"
   sGac(830) = "w------"
   sGac(831) = "a-ts---"
   sGac(832) = "f-ti---"
   sGac(833) = "a-tu---"
   sGac(834) = "a-tk---"
   sGac(835) = "a-tk---"
   sGac(836) = "nwtc---"
   sGac(837) = "potv---"
   sGac(838) = "mm-----"
   sGac(839) = "f-cx---"
   sGac(840) = "f-ug---"
   sGac(841) = "e-un---"
   sGac(842) = "a-ts---"
   sGac(843) = "f-ua--- | a-sy---"
   sGac(844) = "e-ne---"
   sGac(845) = "n-us---"
   sGac(846) = "n-cn-on"
   sGac(847) = "f-uv---"
   sGac(848) = "e-uru--"
   sGac(849) = "zur----"
   sGac(850) = "s-uy---"
   sGac(851) = "n-us---"
   sGac(852) = "n-us-ut"
   sGac(853) = "a-uz---"
   sGac(854) = "a-uz---"
   sGac(855) = "n-us-va"
   sGac(856) = "ponn---"
   sGac(857) = "e-vc---"
   sGac(858) = "s-ve---"
   sGac(859) = "zve----"
   sGac(860) = "n-us-vt"
   sGac(861) = "nwvi---"
   sGac(862) = "u-at-vi"
   sGac(863) = "u-at-vi"
   sGac(864) = "a-vt---"
   sGac(865) = "nwvb--- | nwvi---"
   sGac(866) = "nwvi---"
   sGac(867) = "n-us-va"
   sGac(868) = "e-urp--"
   sGac(869) = "fv-----"
   sGac(870) = "n-us-vt"
   sGac(871) = "n-us-wv"
   sGac(872) = "u-at-we"
   sGac(873) = "powk---"
   sGac(874) = "e-uk-wl"
   sGac(875) = "powf---"
   sGac(876) = "n-us-wa"
   sGac(877) = "n-us-wa"
   sGac(878) = "n-us-wa | n-us-id | n-us-mt | n-us-wy"
   sGac(879) = "n-us-dc | n-us-md | n-us-va"
   sGac(880) = "awba---"
   sGac(881) = "nw-----"
   sGac(882) = "nwla---"
   sGac(883) = "n-us-wv"
   sGac(884) = "u-at-we"
   sGac(885) = "xd-----"
   sGac(886) = "f-ss---"
   sGac(887) = "pows---"
   sGac(888) = "n-usp--"
   sGac(889) = "nwwi---"
   sGac(890) = "n-us-wi"
   sGac(891) = "n-us-wi"
   sGac(892) = "n-us-wy"
   sGac(893) = "n-us-wy"
   sGac(894) = "a-ccs--"
   sGac(895) = "a-cc-su"
   sGac(896) = "a-ccg--"
   sGac(897) = "a-ccy--"
   sGac(898) = "ay-----"
   sGac(899) = "a-ye---"
   sGac(900) = "e-yu---"
   sGac(901) = "n-cn-yk"
   sGac(902) = "n-cn-yk"
   sGac(903) = "a-cc-yu"
   sGac(904) = "f-cg---"
   sGac(905) = "fz-----"
   sGac(906) = "f-za---"
   sGac(907) = "f-tz---"
   sGac(908) = "a-cc-ch"
   sGac(909) = "f-rh---"
End Sub

Sub DefGeo
   ReDim sGeo(1 to 909) As String
   sGeo(1) = "ACT"
   sGeo(2) = "ADAMAWA"
   sGeo(3) = "ADEN"
   sGeo(4) = "ADEN, GULF OF"
   sGeo(5) = "ADRIATIC SEA"
   sGeo(6) = "AEGEAN SEA"
   sGeo(7) = "AFGHANISTAN"
   sGeo(8) = "AFRICA"
   sGeo(9) = "AFRICA, CENTRAL"
   sGeo(10) = "AFRICA, EAST"
   sGeo(11) = "AFRICA, EASTERN"
   sGeo(12) = "AFRICA, FRENCHSPEAKING"
   sGeo(13) = "AFRICA, FRENCHSPEAKING EQUATORIAL"
   sGeo(14) = "AFRICA, FRENCHSPEAKING WEST"
   sGeo(15) = "AFRICA, NORTH"
   sGeo(16) = "AFRICA, NORTHEAST"
   sGeo(17) = "AFRICA, NORTHWEST"
   sGeo(18) = "AFRICA, SOUTHERN"
   sGeo(19) = "AFRICA, SUBSAHARAN"
   sGeo(20) = "AFRICA, WEST"
   sGeo(21) = "ALA"
   sGeo(22) = "ALABAMA"
   sGeo(23) = "ALASKA"
   sGeo(24) = "ALASKA, GULF OF"
   sGeo(25) = "ALBANIA"
   sGeo(26) = "ALBERTA"
   sGeo(27) = "ALGERIA"
   sGeo(28) = "ALPS"
   sGeo(29) = "ALTA"
   sGeo(30) = "AMAZON RIVER"
   sGeo(31) = "AMERICA"
   sGeo(32) = "AMERICAN SAMOA"
   sGeo(33) = "AMUR RIVER"
   sGeo(34) = "ANDAMAN SEA"
   sGeo(35) = "ANDES"
   sGeo(36) = "ANDORRA"
   sGeo(37) = "ANGOLA"
   sGeo(38) = "ANGUILLA"
   sGeo(39) = "ANHUI SHENG"
   sGeo(40) = "ANTARCTIC OCEAN"
   sGeo(41) = "ANTARCTICA"
   sGeo(42) = "ANTIGUA"
   sGeo(43) = "ANTIGUA & BARBUDA"
   sGeo(44) = "ANTILLES, GREATER"
   sGeo(45) = "ANTILLES, LESSER"
   sGeo(46) = "APPALACHIAN"
   sGeo(47) = "APPALACHIAN MOUNTAINS"
   sGeo(48) = "AQABA, GULF OF"
   sGeo(49) = "ARAB COUNTRIES"
   sGeo(50) = "ARABIA, SOUTHERN"
   sGeo(51) = "ARABIAN PENINSULA"
   sGeo(52) = "ARABIAN SEA"
   sGeo(53) = "ARAFURA SEA"
   sGeo(54) = "ARCTIC OCEAN"
   sGeo(55) = "ARCTIC REGIONS"
   sGeo(56) = "ARGENTINA"
   sGeo(57) = "ARIZ"
   sGeo(58) = "ARIZONA"
   sGeo(59) = "ARK"
   sGeo(60) = "ARKANSAS"
   sGeo(61) = "ARMENIA"
   sGeo(62) = "ARMENIAN SSR"
   sGeo(63) = "ARMENIAREP"
   sGeo(64) = "ARUBA"
   sGeo(65) = "ASCENSION ISLAND"
   sGeo(66) = "ASHANTI"
   sGeo(67) = "ASHMORE & CARTIER ISLANDS"
   sGeo(68) = "ASIA"
   sGeo(69) = "ASIA, CENTRAL"
   sGeo(70) = "ATLANTIC OCEAN"
   sGeo(71) = "ATLANTIC STATES"
   sGeo(72) = "ATLAS MOUNTAINS"
   sGeo(73) = "AUSTRALASIA"
   sGeo(74) = "AUSTRALIA"
   sGeo(75) = "AUSTRALIA, NORTHERN"
   sGeo(76) = "AUSTRALIAN CAPITAL TERRITORY"
   sGeo(77) = "AUSTRIA"
   sGeo(78) = "AUSTROHUNGARIAN MONARCHY"
   sGeo(79) = "AZERBAIJAN"
   sGeo(80) = "AZERBAIJAN SSR"
   sGeo(81) = "AZORES"
   sGeo(82) = "BABYLONIA"
   sGeo(83) = "BAHAMAS"
   sGeo(84) = "BAHRAIN"
   sGeo(85) = "BALKAN PENINSULA"
   sGeo(86) = "BALTIC"
   sGeo(87) = "BALTIC SEA"
   sGeo(88) = "BALTIC STATES"
   sGeo(89) = "BANDA SEA"
   sGeo(90) = "BANGLADESH"
   sGeo(91) = "BARBADOS"
   sGeo(92) = "BARBUDA"
   sGeo(93) = "BARENTS SEA"
   sGeo(94) = "BASUTOLAND"
   sGeo(95) = "BATAVIAN REPUBLIC"
   sGeo(96) = "BC"
   sGeo(97) = "BEAUFORT SEA"
   sGeo(98) = "BECHUANALAND"
   sGeo(99) = "BEIJING"
   sGeo(100) = "BELARUS"
   sGeo(101) = "BELGIAN CONGO"
   sGeo(102) = "BELGIUM"
   sGeo(103) = "BELIZE"
   sGeo(104) = "BENELUX COUNTRIES"
   sGeo(105) = "BENGAL, BAY OF"
   sGeo(106) = "BENIN"
   sGeo(107) = "BERBICE"
   sGeo(108) = "BERING SEA"
   sGeo(109) = "BERING STRAIT"
   sGeo(110) = "BERMUDA ISLANDS"
   sGeo(111) = "BERMUDA TRIANGLE"
   sGeo(112) = "BHUTAN"
   sGeo(113) = "BIAFRA"
   sGeo(114) = "BLACK SEA"
   sGeo(115) = "BLUE RIDGE MOUNTAINS"
   sGeo(116) = "BO HAI"
   sGeo(117) = "BOLIVIA"
   sGeo(118) = "BONAIRE"
   sGeo(119) = "BORNEO"
   sGeo(120) = "BOSNIA & HERCEGOVINA"
   sGeo(121) = "BOTHNIA, GULF OF"
   sGeo(122) = "BOTSWANA"
   sGeo(123) = "BOUVET ISLAND"
   sGeo(124) = "BRAZIL"
   sGeo(125) = "BRITISH CENTRAL AFRICA PROTECTORATE"
   sGeo(126) = "BRITISH COLUMBIA"
   sGeo(127) = "BRITISH GUIANA"
   sGeo(128) = "BRITISH HONDURAS"
   sGeo(129) = "BRITISH INDIAN OCEAN TERRITORY"
   sGeo(130) = "BRITISH ISLES"
   sGeo(131) = "BRITISH NEW GUINEA"
   sGeo(132) = "BRITISH SOLOMON ISLANDS"
   sGeo(133) = "BRITISH SOMALILAND"
   sGeo(134) = "BRITISH VIRGIN ISLANDS"
   sGeo(135) = "BRUNEI"
   sGeo(136) = "BULGARIA"
   sGeo(137) = "BURKINA FASO"
   sGeo(138) = "BURMA"
   sGeo(139) = "BURUNDI"
   sGeo(140) = "BYELORUSSIAN SSR"
   sGeo(141) = "BYZANTINE EMPIRE"
   sGeo(142) = "CALIF"
   sGeo(143) = "CALIFORNIA"
   sGeo(144) = "CAMBODIA"
   sGeo(145) = "CAMEROON"
   sGeo(146) = "CAMEROUN"
   sGeo(147) = "CANADA"
   sGeo(148) = "CANADA, WESTERN"
   sGeo(149) = "CANAL ZONE"
   sGeo(150) = "CANARY ISLANDS"
   sGeo(151) = "CAPE VERDE"
   sGeo(152) = "CARIBBEAN AREA"
   sGeo(153) = "CARIBBEAN NETHERLANDS"
   sGeo(154) = "CARIBBEAN SEA"
   sGeo(155) = "CAROLINE ISLANDS"
   sGeo(156) = "CARPATHIAN MOUNTAINS"
   sGeo(157) = "CASPIAN SEA"
   sGeo(158) = "CAUCASUS"
   sGeo(159) = "CAUCASUS, NORTHERN"
   sGeo(160) = "CAUCASUS, SOUTH"
   sGeo(161) = "CAYMAN ISLANDS"
   sGeo(162) = "CELEBES SEA"
   sGeo(163) = "CENTRAL AFRICAN EMPIRE"
   sGeo(164) = "CENTRAL AFRICAN REPUBLIC"
   sGeo(165) = "CENTRAL AMERICA"
   sGeo(166) = "CENTRAL AUSTRALIA"
   sGeo(167) = "CENTRAL CHERNOZEM"
   sGeo(168) = "CEYLON"
   sGeo(169) = "CHAD"
   sGeo(170) = "CHAD, LAKE"
   sGeo(171) = "CHANNEL ISLANDS"
   sGeo(172) = "CHESAPEAKE BAY"
   sGeo(173) = "CHIHUANHUAN DESERT"
   sGeo(174) = "CHILE"
   sGeo(175) = "CHINA"
   sGeo(176) = "CHONGQING"
   sGeo(177) = "CHRISTMAS ISLAND"
   sGeo(178) = "CHUKCHI SEA"
   sGeo(179) = "COCOS ISLANDS"
   sGeo(180) = "COLD REGIONS"
   sGeo(181) = "COLO"
   sGeo(182) = "COLOMBIA"
   sGeo(183) = "COLORADO"
   sGeo(184) = "COLORADO RIVER"
   sGeo(185) = "COLUMBIA RIVER"
   sGeo(186) = "COLUMBIA RIVERWAT"
   sGeo(187) = "COMMONWEALTH COUNTRIES"
   sGeo(188) = "COMMUNAUTE"
   sGeo(189) = "COMMUNIST COUNTRIES"
   sGeo(190) = "COMORO ISLANDS"
   sGeo(191) = "COMOROS"
   sGeo(192) = "CONFEDERATE STATES OF AMERICA"
   sGeo(193) = "CONGO"
   sGeo(194) = "CONGO RIVER"
   sGeo(195) = "CONGO RIVERWAT"
   sGeo(196) = "CONGOBRA"
   sGeo(197) = "CONN"
   sGeo(198) = "CONNECTICUT"
   sGeo(199) = "COOK ISLANDS"
   sGeo(200) = "CORAL SEA"
   sGeo(201) = "CORAL SEA ISLANDS"
   sGeo(202) = "COSTA RICA"
   sGeo(203) = "COTE DIVOIRE"
   sGeo(204) = "CROATIA"
   sGeo(205) = "CUBA"
   sGeo(206) = "CUMBERLAND MOUNTAINS"
   sGeo(207) = "CURACAO"
   sGeo(208) = "CYPRUS"
   sGeo(209) = "CZECH REPUBLIC"
   sGeo(210) = "CZECHOSLOVAKIA"
   sGeo(211) = "DAHOMEY"
   sGeo(212) = "DAKOTA TERRITORY"
   sGeo(213) = "DANUBE RIVER"
   sGeo(214) = "DANUBE RIVERWAT"
   sGeo(215) = "DC"
   sGeo(216) = "DEEP SPACE"
   sGeo(217) = "DEL"
   sGeo(218) = "DELAWARE"
   sGeo(219) = "DEMERARA"
   sGeo(220) = "DENMARK"
   sGeo(221) = "DEVELOPED COUNTRIES"
   sGeo(222) = "DEVELOPING COUNTRIES"
   sGeo(223) = "DISTRICT OF COLUMBIA"
   sGeo(224) = "DJIBOUTI"
   sGeo(225) = "DOMINICA"
   sGeo(226) = "DOMINICAN REPUBLIC"
   sGeo(227) = "DUTCH EAST INDIES"
   sGeo(228) = "EARTH"
   sGeo(229) = "EAST AFRICA PROTECTORATE"
   sGeo(230) = "EAST ASIA"
   sGeo(231) = "EAST CHINA SEA"
   sGeo(232) = "EAST INDIES"
   sGeo(233) = "EAST PRUSSIA"
   sGeo(234) = "EAST TIMOR"
   sGeo(235) = "EASTER ISLAND"
   sGeo(236) = "EASTERN AUSTRALIA"
   sGeo(237) = "EASTERN HEMISPHERE"
   sGeo(238) = "EASTUS"
   sGeo(239) = "ECUADOR"
   sGeo(240) = "EGYPT"
   sGeo(241) = "EL SALVADOR"
   sGeo(242) = "ENGLAND"
   sGeo(243) = "ENGLISH CHANNEL"
   sGeo(244) = "EQUATOR"
   sGeo(245) = "EQUATORIAL GUINEA"
   sGeo(246) = "ERIE, LAKE"
   sGeo(247) = "ERITREA"
   sGeo(248) = "ESSEQUIBO"
   sGeo(249) = "ESTONIA"
   sGeo(250) = "ETHIOPIA"
   sGeo(251) = "EURASIA"
   sGeo(252) = "EUROPE"
   sGeo(253) = "EUROPE, CENTRAL"
   sGeo(254) = "EUROPE, EASTERN"
   sGeo(255) = "EUROPE, NORTHERN"
   sGeo(256) = "EUROPE, SOUTHERN"
   sGeo(257) = "EUROPE, WESTERN"
   sGeo(258) = "EUROPEAN ECONOMIC COMMUNITY COUNTRIES"
   sGeo(259) = "EUROPEAN UNION COUNTRIES"
   sGeo(260) = "FALKLAND ISLANDS"
   sGeo(261) = "FAROE ISLANDS"
   sGeo(262) = "FEDERATED MALAY STATES"
   sGeo(263) = "FEDERATION OF SOUTH ARABIA"
   sGeo(264) = "FIJI"
   sGeo(265) = "FINLAND"
   sGeo(266) = "FINLAND, GULF OF"
   sGeo(267) = "FLA"
   sGeo(268) = "FLORIDA"
   sGeo(269) = "FORMER SOVIET REPUBLICS"
   sGeo(270) = "FORMER YUGOSLAV REPUBLICS"
   sGeo(271) = "FRANCE"
   sGeo(272) = "FRENCH CAMEROONS"
   sGeo(273) = "FRENCH COMMUNITY"
   sGeo(274) = "FRENCH EQUATORIAL AFRICA"
   sGeo(275) = "FRENCH GUIANA"
   sGeo(276) = "FRENCH GUINEA"
   sGeo(277) = "FRENCH INDIA"
   sGeo(278) = "FRENCH INDOCHINA"
   sGeo(279) = "FRENCH POLYNESIA"
   sGeo(280) = "FRENCH SOMALILAND"
   sGeo(281) = "FRENCH SUDAN"
   sGeo(282) = "FRENCH TERRITORY OF THE AFARS & ISSAS"
   sGeo(283) = "FUJIAN SHENG"
   sGeo(284) = "FUNDY, BAY OF"
   sGeo(285) = "FUTA JALLON"
   sGeo(286) = "GA"
   sGeo(287) = "GABON"
   sGeo(288) = "GALAPAGOS ISLANDS"
   sGeo(289) = "GAMBIA"
   sGeo(290) = "GANSU SHENG"
   sGeo(291) = "GAZA STRIP"
   sGeo(292) = "GEORGES BANK"
   sGeo(293) = "GEORGIA"
   sGeo(294) = "GEORGIAN SSR"
   sGeo(295) = "GEORGIAREP"
   sGeo(296) = "GERMAN EAST AFRICA"
   sGeo(297) = "GERMAN NEW GUINEA"
   sGeo(298) = "GERMAN SOUTHWEST AFRICA"
   sGeo(299) = "GERMANY"
   sGeo(300) = "GERMANYEAS"
   sGeo(301) = "GERMANYWES"
   sGeo(302) = "GHANA"
   sGeo(303) = "GIBRALTAR"
   sGeo(304) = "GIBRALTAR, STRAIT OF"
   sGeo(305) = "GILBERT & ELLICE ISLANDS COLONY"
   sGeo(306) = "GILBERT ISLANDS"
   sGeo(307) = "GOLD"
   sGeo(308) = "GREAT BASIN"
   sGeo(309) = "GREAT BRITAIN"
   sGeo(310) = "GREAT LAKES"
   sGeo(311) = "GREAT PLAINS"
   sGeo(312) = "GREAT RIFT"
   sGeo(313) = "GREECE"
   sGeo(314) = "GREEN RIVER"
   sGeo(315) = "GREENLAND"
   sGeo(316) = "GREENLAND SEA"
   sGeo(317) = "GRENADA"
   sGeo(318) = "GUADELOUPE"
   sGeo(319) = "GUAM"
   sGeo(320) = "GUANGDONG SHENG"
   sGeo(321) = "GUANGXI ZHUANGZU ZIZHIQU"
   sGeo(322) = "GUATEMALA"
   sGeo(323) = "GUINEA"
   sGeo(324) = "GUINEA, GULF OF"
   sGeo(325) = "GUINEA, GULF OFREG"
   sGeo(326) = "GUINEABISSAU"
   sGeo(327) = "GUIZHOU SHENG"
   sGeo(328) = "GULF STATES"
   sGeo(329) = "GUYANA"
   sGeo(330) = "HAINAN SHENG"
   sGeo(331) = "HAITI"
   sGeo(332) = "HAWAII"
   sGeo(333) = "HEARD & MCDONALD ISLANDS"
   sGeo(334) = "HEBEI SHENG"
   sGeo(335) = "HEILONGJIANG SHENG"
   sGeo(336) = "HENAN SHENG"
   sGeo(337) = "HIMALAYA MOUNTAINS"
   sGeo(338) = "HISPANIOLA"
   sGeo(339) = "HOLLAND"
   sGeo(340) = "HOLY ROMAN EMPIRE"
   sGeo(341) = "HONDURAS"
   sGeo(342) = "HONG KONG"
   sGeo(343) = "HORMUZ, STRAIT OF"
   sGeo(344) = "HORN OF AFRICA"
   sGeo(345) = "HUBEI SHENG"
   sGeo(346) = "HUDSON BAY"
   sGeo(347) = "HUNAN SHENG"
   sGeo(348) = "HUNGARY"
   sGeo(349) = "HURON, LAKE"
   sGeo(350) = "IBERIAN PENINSULA"
   sGeo(351) = "ICELAND"
   sGeo(352) = "IDAHO"
   sGeo(353) = "IDAHO TERRITORY"
   sGeo(354) = "IFNI"
   sGeo(355) = "ILL"
   sGeo(356) = "ILLINOIS"
   sGeo(357) = "IND"
   sGeo(358) = "INDIA"
   sGeo(359) = "INDIAN OCEAN"
   sGeo(360) = "INDIAN TERRITORY"
   sGeo(361) = "INDIANA"
   sGeo(362) = "INDOCHINA"
   sGeo(363) = "INDONESIA"
   sGeo(364) = "INNER MONGOLIA"
   sGeo(365) = "IOWA"
   sGeo(366) = "IRAN"
   sGeo(367) = "IRAQ"
   sGeo(368) = "IRELAND"
   sGeo(369) = "IRISH SEA"
   sGeo(370) = "ISLAMIC COUNTRIES"
   sGeo(371) = "ISLAMIC EMPIRE"
   sGeo(372) = "ISLANDS OF THE ADRIATIC"
   sGeo(373) = "ISLANDS OF THE ARCTIC"
   sGeo(374) = "ISLANDS OF THE ATLANTIC"
   sGeo(375) = "ISLANDS OF THE BALTIC"
   sGeo(376) = "ISLANDS OF THE INDIAN OCEAN"
   sGeo(377) = "ISLANDS OF THE MEDITERRANEAN"
   sGeo(378) = "ISLANDS OF THE PACIFIC"
   sGeo(379) = "ISLANDS OF THE SOUTH CHINA SEA"
   sGeo(380) = "ISLE OF MAN"
   sGeo(381) = "ISRAEL"
   sGeo(382) = "ITALIAN EAST AFRICA"
   sGeo(383) = "ITALIAN SOMALILAND"
   sGeo(384) = "ITALY"
   sGeo(385) = "JAMAICA"
   sGeo(386) = "JAN MAYEN ISLAND"
   sGeo(387) = "JAPAN"
   sGeo(388) = "JAPAN, SEA OF"
   sGeo(389) = "JERUSALEM"
   sGeo(390) = "JERVIS BAY TERRITORY"
   sGeo(391) = "JIANGSU SHENG"
   sGeo(392) = "JIANGXI SHENG"
   sGeo(393) = "JILIN SHENG"
   sGeo(394) = "JOHNSTON ISLAND"
   sGeo(395) = "JOHOR"
   sGeo(396) = "JORDAN"
   sGeo(397) = "JUPITER"
   sGeo(398) = "KAMERUN"
   sGeo(399) = "KAN"
   sGeo(400) = "KANSAS"
   sGeo(401) = "KAZAKH SSR"
   sGeo(402) = "KAZAKHSTAN"
   sGeo(403) = "KEDAH"
   sGeo(404) = "KELANTAN"
   sGeo(405) = "KENTUCKY"
   sGeo(406) = "KENYA"
   sGeo(407) = "KERGUELEN ISLANDS"
   sGeo(408) = "KERMADEC ISLANDS"
   sGeo(409) = "KIEVAN RUS"
   sGeo(410) = "KIRGHIZ SSR"
   sGeo(411) = "KIRIBATI"
   sGeo(412) = "KONGO KINGDOM"
   sGeo(413) = "KOREA"
   sGeo(414) = "KOREANOR"
   sGeo(415) = "KOREASOU"
   sGeo(416) = "KOSOVO"
   sGeo(417) = "KUNLUN MOUNTAINS"
   sGeo(418) = "KUWAIT"
   sGeo(419) = "KY"
   sGeo(420) = "KYRGYZSTAN"
   sGeo(421) = "LA"
   sGeo(422) = "LABRADOR SEA"
   sGeo(423) = "LAKE STATES"
   sGeo(424) = "LAOS"
   sGeo(425) = "LAPLAND"
   sGeo(426) = "LATIN AMERICA"
   sGeo(427) = "LATIN ORIENT"
   sGeo(428) = "LATVIA"
   sGeo(429) = "LEBANON"
   sGeo(430) = "LEEWARD ANTILLES"
   sGeo(431) = "LEEWARD ISLANDS"
   sGeo(432) = "LESOTHO"
   sGeo(433) = "LESSER SUNDA ISLANDS"
   sGeo(434) = "LIAONING SHENG"
   sGeo(435) = "LIBERIA"
   sGeo(436) = "LIBYA"
   sGeo(437) = "LIECHTENSTEIN"
   sGeo(438) = "LINE ISLANDS"
   sGeo(439) = "LITHUANIA"
   sGeo(440) = "LOUISIANA"
   sGeo(441) = "LOWER CANADA"
   sGeo(442) = "LUXEMBOURG"
   sGeo(443) = "MACARONESIA"
   sGeo(444) = "MACAU"
   sGeo(445) = "MACEDONIA"
   sGeo(446) = "MACEDONIAREP"
   sGeo(447) = "MADAGASCAR"
   sGeo(448) = "MADEIRA ISLANDS"
   sGeo(449) = "MAINE"
   sGeo(450) = "MALACCA"
   sGeo(451) = "MALACCA, STRAIT OF"
   sGeo(452) = "MALAGASY REPUBLIC"
   sGeo(453) = "MALAWI"
   sGeo(454) = "MALAY ARCHIPELAGO"
   sGeo(455) = "MALAY PENINSULA"
   sGeo(456) = "MALAYA"
   sGeo(457) = "MALAYAN UNION"
   sGeo(458) = "MALAYSIA"
   sGeo(459) = "MALDIVES"
   sGeo(460) = "MALI"
   sGeo(461) = "MALI FEDERATION"
   sGeo(462) = "MALTA"
   sGeo(463) = "MAN"
   sGeo(464) = "MANITOBA"
   sGeo(465) = "MARIANA ISLANDS"
   sGeo(466) = "MARITIME PROVINCES"
   sGeo(467) = "MARS"
   sGeo(468) = "MARSHALL ISLANDS"
   sGeo(469) = "MARTINIQUE"
   sGeo(470) = "MARYLAND"
   sGeo(471) = "MASS"
   sGeo(472) = "MASSACHUSETTS"
   sGeo(473) = "MAURITANIA"
   sGeo(474) = "MAURITIUS"
   sGeo(475) = "MAYOTTE"
   sGeo(476) = "MD"
   sGeo(477) = "ME"
   sGeo(478) = "MEDITERRANEAN"
   sGeo(479) = "MEDITERRANEAN SEA"
   sGeo(480) = "MEKONG RIVER"
   sGeo(481) = "MELANESIA"
   sGeo(482) = "MERCURY"
   sGeo(483) = "MEXICO"
   sGeo(484) = "MEXICO, GULF OF"
   sGeo(485) = "MICH"
   sGeo(486) = "MICHIGAN"
   sGeo(487) = "MICHIGAN, LAKE"
   sGeo(488) = "MICRONESIA"
   sGeo(489) = "MICRONESIAFED"
   sGeo(490) = "MIDDLE ATLANTIC STATES"
   sGeo(491) = "MIDDLE CONGO"
   sGeo(492) = "MIDDLE EAST"
   sGeo(493) = "MIDDLE WEST"
   sGeo(494) = "MIDWAY ISLANDS"
   sGeo(495) = "MINN"
   sGeo(496) = "MINNESOTA"
   sGeo(497) = "MISS"
   sGeo(498) = "MISSISSIPPI"
   sGeo(499) = "MISSISSIPPI RIVER"
   sGeo(500) = "MISSOURI"
   sGeo(501) = "MISSOURI RIVER"
   sGeo(502) = "MO"
   sGeo(503) = "MOLDAVIA"
   sGeo(504) = "MOLDAVIAN SSR"
   sGeo(505) = "MOLDOVA"
   sGeo(506) = "MONACO"
   sGeo(507) = "MONGOLIA"
   sGeo(508) = "MONT"
   sGeo(509) = "MONTANA"
   sGeo(510) = "MONTENEGRO"
   sGeo(511) = "MONTSERRAT"
   sGeo(512) = "MOON"
   sGeo(513) = "MOROCCO"
   sGeo(514) = "MOZAMBIQUE"
   sGeo(515) = "MUSCAT & OMAN"
   sGeo(516) = "NAMIBIA"
   sGeo(517) = "NAURU"
   sGeo(518) = "NAVARRE"
   sGeo(519) = "NB"
   sGeo(520) = "NC"
   sGeo(521) = "ND"
   sGeo(522) = "NEB"
   sGeo(523) = "NEBRASKA"
   sGeo(524) = "NEGERI SEMBILAN"
   sGeo(525) = "NEPAL"
   sGeo(526) = "NEPTUNE"
   sGeo(527) = "NETHERLANDS"
   sGeo(528) = "NETHERLANDS ANTILLES"
   sGeo(529) = "NETHERLANDS NEW GUINEA"
   sGeo(530) = "NETHERLANDS WEST INDIES"
   sGeo(531) = "NEV"
   sGeo(532) = "NEVADA"
   sGeo(533) = "NEVIS"
   sGeo(534) = "NEW BRUNSWICK"
   sGeo(535) = "NEW CALEDONIA"
   sGeo(536) = "NEW ENGLAND"
   sGeo(537) = "NEW FRANCE"
   sGeo(538) = "NEW GUINEA"
   sGeo(539) = "NEW HAMPSHIRE"
   sGeo(540) = "NEW HEBRIDES"
   sGeo(541) = "NEW JERSEY"
   sGeo(542) = "NEW MEXICO"
   sGeo(543) = "NEW NETHERLAND"
   sGeo(544) = "NEW SOUTH WALES"
   sGeo(545) = "NEW YORK"
   sGeo(546) = "NEW ZEALAND"
   sGeo(547) = "NEWFOUNDLAND"
   sGeo(548) = "NEWFOUNDLAND & LABRADOR"
   sGeo(549) = "NFLD"
   sGeo(550) = "NH"
   sGeo(551) = "NICARAGUA"
   sGeo(552) = "NIGER"
   sGeo(553) = "NIGER RIVER"
   sGeo(554) = "NIGERIA"
   sGeo(555) = "NILE RIVER"
   sGeo(556) = "NINGXIA HUIZU ZIZHIQU"
   sGeo(557) = "NIUE"
   sGeo(558) = "NJ"
   sGeo(559) = "NL"
   sGeo(560) = "NM"
   sGeo(561) = "NORFOLK ISLAND"
   sGeo(562) = "NORTH AMERICA"
   sGeo(563) = "NORTH ATLANTIC"
   sGeo(564) = "NORTH ATLANTIC OCEAN"
   sGeo(565) = "NORTH BORNEO"
   sGeo(566) = "NORTH CAROLINA"
   sGeo(567) = "NORTH DAKOTA"
   sGeo(568) = "NORTH PACIFIC"
   sGeo(569) = "NORTH PACIFIC OCEAN"
   sGeo(570) = "NORTH POLE"
   sGeo(571) = "NORTH SEA"
   sGeo(572) = "NORTHEAST PASSAGE"
   sGeo(573) = "NORTHEASTERN STATES"
   sGeo(574) = "NORTHERN HEMISPHERE"
   sGeo(575) = "NORTHERN IRELAND"
   sGeo(576) = "NORTHERN MARIANA ISLANDS"
   sGeo(577) = "NORTHERN RHODESIA"
   sGeo(578) = "NORTHERN TERRITORY"
   sGeo(579) = "NORTHWEST PASSAGE"
   sGeo(580) = "NORTHWEST TERRITORIES"
   sGeo(581) = "NORTHWEST, CANADIAN"
   sGeo(582) = "NORTHWEST, OLD"
   sGeo(583) = "NORTHWEST, PACIFIC"
   sGeo(584) = "NORTHWESTERN STATES"
   sGeo(585) = "NORWAY"
   sGeo(586) = "NORWEGIAN SEA"
   sGeo(587) = "NOVA SCOTIA"
   sGeo(588) = "NS"
   sGeo(589) = "NSW"
   sGeo(590) = "NT"
   sGeo(591) = "NUNAVUT"
   sGeo(592) = "NWT"
   sGeo(593) = "NY"
   sGeo(594) = "NYASA, LAKE"
   sGeo(595) = "NYASALAND"
   sGeo(596) = "NZ"
   sGeo(597) = "OCEANIA"
   sGeo(598) = "OHIO"
   sGeo(599) = "OHIO RIVER"
   sGeo(600) = "OKLA"
   sGeo(601) = "OKLAHOMA"
   sGeo(602) = "OMAN"
   sGeo(603) = "OMAN, GULF OF"
   sGeo(604) = "ONT"
   sGeo(605) = "ONTARIO"
   sGeo(606) = "ONTARIO, LAKE"
   sGeo(607) = "OR"
   sGeo(608) = "OREGON"
   sGeo(609) = "OREGON TERRITORY"
   sGeo(610) = "ORIENT"
   sGeo(611) = "OUTER SPACE"
   sGeo(612) = "OZARK MOUNTAINS"
   sGeo(613) = "PA"
   sGeo(614) = "PACIFIC AREA"
   sGeo(615) = "PACIFIC ISLANDS"
   sGeo(616) = "PACIFIC OCEAN"
   sGeo(617) = "PACIFIC STATES"
   sGeo(618) = "PAHANG"
   sGeo(619) = "PAKISTAN"
   sGeo(620) = "PALAU"
   sGeo(621) = "PALESTINE"
   sGeo(622) = "PANAMA"
   sGeo(623) = "PANNONIA"
   sGeo(624) = "PANTANAL"
   sGeo(625) = "PAPUA"
   sGeo(626) = "PAPUA NEW GUINEA"
   sGeo(627) = "PARACEL ISLANDS"
   sGeo(628) = "PARAGUAY"
   sGeo(629) = "PEI"
   sGeo(630) = "PENNSYLVANIA"
   sGeo(631) = "PERAK"
   sGeo(632) = "PERLIS"
   sGeo(633) = "PERSIAN GULF"
   sGeo(634) = "PERSIAN GULF STATES"
   sGeo(635) = "PERU"
   sGeo(636) = "PHILIPPINE SEA"
   sGeo(637) = "PHILIPPINES"
   sGeo(638) = "PINANG"
   sGeo(639) = "PITCAIRN ISLAND"
   sGeo(640) = "PLUTO"
   sGeo(641) = "POLAND"
   sGeo(642) = "POLAR REGIONS"
   sGeo(643) = "POLYNESIA"
   sGeo(644) = "PORTUGAL"
   sGeo(645) = "PORTUGUESE EAST AFRICA"
   sGeo(646) = "PORTUGUESE GUINEA"
   sGeo(647) = "PORTUGUESE TIMOR"
   sGeo(648) = "PR"
   sGeo(649) = "PRAIRIE PROVINCES"
   sGeo(650) = "PRATAS ISLANDS"
   sGeo(651) = "PRINCE EDWARD ISLAND"
   sGeo(652) = "PRINCE EDWARD ISLANDS"
   sGeo(653) = "PROTECTORATE OF SOUTH ARABIA"
   sGeo(654) = "PUERTO RICO"
   sGeo(655) = "PYRENEES"
   sGeo(656) = "QATAR"
   sGeo(657) = "QINGHAI SHENG"
   sGeo(658) = "QLD"
   sGeo(659) = "QUEBEC"
   sGeo(660) = "QUEENSLAND"
   sGeo(661) = "RED SEA"
   sGeo(662) = "REUNION"
   sGeo(663) = "RHINE RIVER"
   sGeo(664) = "RHODE ISLAND"
   sGeo(665) = "RHODESIA"
   sGeo(666) = "RHODESIA & NYASALAND"
   sGeo(667) = "RI"
   sGeo(668) = "RIO DE LA PLATA"
   sGeo(669) = "ROCKY MOUNTAINS"
   sGeo(670) = "ROMANIA"
   sGeo(671) = "ROME"
   sGeo(672) = "RSFSR"
   sGeo(673) = "RUANAURUNDI"
   sGeo(674) = "RUSSIA"
   sGeo(675) = "RUSSIA, NORTHERN"
   sGeo(676) = "RUSSIAFED"
   sGeo(677) = "RUSSIAN FAR EAST"
   sGeo(678) = "RUSSIAN SFSR"
   sGeo(679) = "RWANDA"
   sGeo(680) = "RYUKYU ISLANDS"
   sGeo(681) = "S AUST"
   sGeo(682) = "SABA"
   sGeo(683) = "SABAH"
   sGeo(684) = "SAHARA"
   sGeo(685) = "SAHEL"
   sGeo(686) = "SAINT HELENA"
   sGeo(687) = "SAINT KITTS"
   sGeo(688) = "SAINT KITTS & NEVIS"
   sGeo(689) = "SAINT KITTSNEVISANGUILLA"
   sGeo(690) = "SAINT LAWRENCE RIVER"
   sGeo(691) = "SAINT LUCIA"
   sGeo(692) = "SAINT MARTIN"
   sGeo(693) = "SAINT PIERRE & MIQUELON"
   sGeo(694) = "SAINT VINCENT"
   sGeo(695) = "SAINT VINCENT & THE GRENADINES"
   sGeo(696) = "SAINTBARTHELEMY"
   sGeo(697) = "SAINTMARTIN"
   sGeo(698) = "SAMOA"
   sGeo(699) = "SAMOAN ISLANDS"
   sGeo(700) = "SAN MARINO"
   sGeo(701) = "SAO TOME & PRINCIPE"
   sGeo(702) = "SARAWAK"
   sGeo(703) = "SARGASSO SEA"
   sGeo(704) = "SASK"
   sGeo(705) = "SASKATCHEWAN"
   sGeo(706) = "SATURN"
   sGeo(707) = "SAUDI ARABIA"
   sGeo(708) = "SC"
   sGeo(709) = "SCANDINAVIA"
   sGeo(710) = "SCANDINAVIAN PENINSULA"
   sGeo(711) = "SCOTLAND"
   sGeo(712) = "SD"
   sGeo(713) = "SEA OF OKHOTSK"
   sGeo(714) = "SELANGOR"
   sGeo(715) = "SENEGAL"
   sGeo(716) = "SERBIA"
   sGeo(717) = "SERBIA & MONTENEGRO"
   sGeo(718) = "SEYCHELLES"
   sGeo(719) = "SHAANXI SHENG"
   sGeo(720) = "SHANDONG SHENG"
   sGeo(721) = "SHANGHAI"
   sGeo(722) = "SHANXI SHENG"
   sGeo(723) = "SIAM"
   sGeo(724) = "SIBERIA"
   sGeo(725) = "SIBERIA, EASTERN"
   sGeo(726) = "SIBERIA, WESTERN"
   sGeo(727) = "SICHUAN SHENG"
   sGeo(728) = "SIERRA LEONE"
   sGeo(729) = "SINGAPORE"
   sGeo(730) = "SINT EUSTATIUS"
   sGeo(731) = "SINT MAARTEN"
   sGeo(732) = "SKAGERRAK"
   sGeo(733) = "SLAVIC COUNTRIES"
   sGeo(734) = "SLOVAKIA"
   sGeo(735) = "SLOVENIA"
   sGeo(736) = "SNOWBELT STATES"
   sGeo(737) = "SOLAR SYSTEM"
   sGeo(738) = "SOLOMON ISLANDS"
   sGeo(739) = "SOMALIA"
   sGeo(740) = "SONGHAI EMPIRE"
   sGeo(741) = "SONORAN DESERT"
   sGeo(742) = "SOUTH AFRICA"
   sGeo(743) = "SOUTH AMERICA"
   sGeo(744) = "SOUTH ASIA"
   sGeo(745) = "SOUTH ATLANTIC OCEAN"
   sGeo(746) = "SOUTH ATLANTIC STATES"
   sGeo(747) = "SOUTH AUSTRALIA"
   sGeo(748) = "SOUTH CAROLINA"
   sGeo(749) = "SOUTH CHINA SEA"
   sGeo(750) = "SOUTH DAKOTA"
   sGeo(751) = "SOUTH GEORGIA & SOUTH SANDWICH ISLANDS"
   sGeo(752) = "SOUTH ORKNEY ISLANDS"
   sGeo(753) = "SOUTH PACIFIC OCEAN"
   sGeo(754) = "SOUTH POLE"
   sGeo(755) = "SOUTH SUDAN"
   sGeo(756) = "SOUTHEAST ASIA"
   sGeo(757) = "SOUTHERN CAMEROONS"
   sGeo(758) = "SOUTHERN CONE OF SOUTH AMERICA"
   sGeo(759) = "SOUTHERN HEMISPHERE"
   sGeo(760) = "SOUTHERN RHODESIA"
   sGeo(761) = "SOUTHERN STATES"
   sGeo(762) = "SOUTHERN YEMEN"
   sGeo(763) = "SOUTHWEST AFRICA"
   sGeo(764) = "SOUTHWEST, NEW"
   sGeo(765) = "SOUTHWEST, OLD"
   sGeo(766) = "SOUTHWESTERN STATES"
   sGeo(767) = "SOVIET UNION"
   sGeo(768) = "SOVIET UNION, NORTHWESTERN"
   sGeo(769) = "SPAIN"
   sGeo(770) = "SPANISH GUINEA"
   sGeo(771) = "SPANISH MAIN"
   sGeo(772) = "SPANISH SAHARA"
   sGeo(773) = "SPANISH WEST AFRICA"
   sGeo(774) = "SPRATLY ISLANDS"
   sGeo(775) = "SRI LANKA"
   sGeo(776) = "STRAITS SETTLEMENTS"
   sGeo(777) = "SUDAN"
   sGeo(778) = "SUDANESE REPUBLIC"
   sGeo(779) = "SUDANREG"
   sGeo(780) = "SUEZ CANAL"
   sGeo(781) = "SULU SEA"
   sGeo(782) = "SUN"
   sGeo(783) = "SUNBELT STATES"
   sGeo(784) = "SUPERIOR, LAKE"
   sGeo(785) = "SURINAME"
   sGeo(786) = "SVALBARD"
   sGeo(787) = "SWAN ISLANDS"
   sGeo(788) = "SWAZILAND"
   sGeo(789) = "SWEDEN"
   sGeo(790) = "SWITZERLAND"
   sGeo(791) = "SYRIA"
   sGeo(792) = "TAIWAN"
   sGeo(793) = "TAJIK SSR"
   sGeo(794) = "TAJIKISTAN"
   sGeo(795) = "TANGANYIKA"
   sGeo(796) = "TANGANYIKA, LAKE"
   sGeo(797) = "TANZANIA"
   sGeo(798) = "TAS"
   sGeo(799) = "TASMAN SEA"
   sGeo(800) = "TASMANIA"
   sGeo(801) = "TENN"
   sGeo(802) = "TENNESSEE"
   sGeo(803) = "TERENGGANU"
   sGeo(804) = "TERRES AUSTRALES ET ANTARCTIQUES FRANCAISES"
   sGeo(805) = "TEX"
   sGeo(806) = "TEXAS"
   sGeo(807) = "THAILAND"
   sGeo(808) = "THAILAND, GULF OF"
   sGeo(809) = "TIANJIN"
   sGeo(810) = "TIBESTI MOUNTAINS"
   sGeo(811) = "TIBET"
   sGeo(812) = "TIBET, PLATEAU OF"
   sGeo(813) = "TIBETREG"
   sGeo(814) = "TIEN SHAN"
   sGeo(815) = "TIMOR ISLAND"
   sGeo(816) = "TIMOR SEA"
   sGeo(817) = "TIMORLESTE"
   sGeo(818) = "TOBAGO"
   sGeo(819) = "TOGO"
   sGeo(820) = "TOGOLAND"
   sGeo(821) = "TOKELAU"
   sGeo(822) = "TOKELAU ISLANDS"
   sGeo(823) = "TONGA"
   sGeo(824) = "TONKIN"
   sGeo(825) = "TORRES STRAIT"
   sGeo(826) = "TRANSJORDAN"
   sGeo(827) = "TRINIDAD"
   sGeo(828) = "TRINIDAD & TOBAGO"
   sGeo(829) = "TRISTAN DA CUNHA"
   sGeo(830) = "TROPICS"
   sGeo(831) = "TRUCIAL STATES"
   sGeo(832) = "TUNISIA"
   sGeo(833) = "TURKEY"
   sGeo(834) = "TURKMEN SSR"
   sGeo(835) = "TURKMENISTAN"
   sGeo(836) = "TURKS & CAICOS ISLANDS"
   sGeo(837) = "TUVALU"
   sGeo(838) = "TYRRHENIAN SEA"
   sGeo(839) = "UBANGISHARI"
   sGeo(840) = "UGANDA"
   sGeo(841) = "UKRAINE"
   sGeo(842) = "UNITED ARAB EMIRATES"
   sGeo(843) = "UNITED ARAB REPUBLIC"
   sGeo(844) = "UNITED PROVINCES OF THE NETHERLANDS"
   sGeo(845) = "UNITED STATES"
   sGeo(846) = "UPPER CANADA"
   sGeo(847) = "UPPER VOLTA"
   sGeo(848) = "URAL MOUNTAINS"
   sGeo(849) = "URANUS"
   sGeo(850) = "URUGUAY"
   sGeo(851) = "US"
   sGeo(852) = "UTAH"
   sGeo(853) = "UZBEK SSR"
   sGeo(854) = "UZBEKISTAN"
   sGeo(855) = "VA"
   sGeo(856) = "VANUATU"
   sGeo(857) = "VATICAN CITY"
   sGeo(858) = "VENEZUELA"
   sGeo(859) = "VENUS"
   sGeo(860) = "VERMONT"
   sGeo(861) = "VI"
   sGeo(862) = "VIC"
   sGeo(863) = "VICTORIA"
   sGeo(864) = "VIETNAM"
   sGeo(865) = "VIRGIN ISLANDS"
   sGeo(866) = "VIRGIN ISLANDS OF THE UNITED STATES"
   sGeo(867) = "VIRGINIA"
   sGeo(868) = "VOLGA RIVER"
   sGeo(869) = "VOLTA RIVER"
   sGeo(870) = "VT"
   sGeo(871) = "W VA"
   sGeo(872) = "WA"
   sGeo(873) = "WAKE ISLAND"
   sGeo(874) = "WALES"
   sGeo(875) = "WALLIS & FUTUNA ISLANDS"
   sGeo(876) = "WASH"
   sGeo(877) = "WASHINGTON"
   sGeo(878) = "WASHINGTON TERRITORY"
   sGeo(879) = "WASHINGTONREG"
   sGeo(880) = "WEST BANK"
   sGeo(881) = "WEST INDIES"
   sGeo(882) = "WEST INDIES, FRENCH"
   sGeo(883) = "WEST VIRGINIA"
   sGeo(884) = "WESTERN AUSTRALIA"
   sGeo(885) = "WESTERN HEMISPHERE"
   sGeo(886) = "WESTERN SAHARA"
   sGeo(887) = "WESTERN SAMOA"
   sGeo(888) = "WESTUS"
   sGeo(889) = "WINDWARD ISLANDS"
   sGeo(890) = "WIS"
   sGeo(891) = "WISCONSIN"
   sGeo(892) = "WYO"
   sGeo(893) = "WYOMING"
   sGeo(894) = "XI RIVER"
   sGeo(895) = "XINJIANG UYGUR ZIZHIQU"
   sGeo(896) = "YANGTZE RIVER"
   sGeo(897) = "YELLOW RIVER"
   sGeo(898) = "YELLOW SEA"
   sGeo(899) = "YEMEN"
   sGeo(900) = "YUGOSLAVIA"
   sGeo(901) = "YUKON"
   sGeo(902) = "YUKON TERRITORY"
   sGeo(903) = "YUNNAN SHENG"
   sGeo(904) = "ZAIRE"
   sGeo(905) = "ZAMBEZI RIVER"
   sGeo(906) = "ZAMBIA"
   sGeo(907) = "ZANZIBAR"
   sGeo(908) = "ZHEJIANG SHENG"
   sGeo(909) = "ZIMBABWE"
End Sub

Sub NrmHdgAnd (sText$)
   If InStr(sHdg, sText) <> 0 Then
      sHdg = Left(sHdg, InStr(sHdg, sText) - 1) & Left(sText, InStr(sText, " AND ") - 1) & " & " & Mid(sText, InStr(sText, " AND ") + 5) & Mid(sHdg, InStr(sHdg, sText) + Len(sText))
   End If
End Sub

Sub NrmHdgDel (sText$)
   If InStr(sHdg, sText) <> 0 Then
      sHdg = Left(sHdg, InStr(sHdg, sText) - 1) & " " & Mid(sHdg, InStr(sHdg, sText) + Len(sText))
   End If
End Sub

Sub NrmHdgRep (sText$, sRepl$)
   If InStr(sHdg, sText) <> 0 Then
      sHdg = Left(sHdg, InStr(sHdg, sText) - 1) & " " & sRepl & " " & Mid(sHdg, InStr(sHdg, sText) + Len(sText)) 
   End If
End Sub

Sub NrmHdg
   Dim p%
   p = 1
   Do While InStr(sHdg, "  ") <> 0
      sHdg = Left(sHdg, InStr(sHdg, "  ") - 1) & " " & Mid(sHdg, InStr(sHdg, "  ") + 2)
   Loop
   If InStr(sHdg, " : ") <> 0 and InStr(sHdg, "(") <> 0 and InStr(sHdg, ")") <> 0 Then
      sHdg = Left(sHdg, InStr(sHdg, " : ") - 1) & ")"
   End If
   sHdg = "  " & UCase(sHdg) & "  "
   Do While InStr(sHdg, "( ") <> 0
      sHdg = Left(sHdg, InStr(sHdg, "( ") - 1) & "(" & Mid(sHdg, InStr(sHdg, "( ") + 2)
   Loop
   Do While InStr(sHdg, " )") <> 0
      sHdg = Left(sHdg, InStr(sHdg, " )") - 1) & ")" & Mid(sHdg, InStr(sHdg, " )") + 2)
   Loop
   If InStr(sHdg, "ßZ") <> 0 Then
      sHdg = Left(sHdg, InStr(sHdg, "ßZ") - 1) & "\Z" & Mid(sHdg, InStr(sHdg, "ßZ") + 2, Len(sHdg))
   End If
   If InStr(sHdg, "ß") <> 0 Then
      sHdg = Left(sHdg, InStr(sHdg, "ß") - 2) & "  "
   End If
   Do While p < Len(sHdg)
      If InStr(" ()\,1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZß", Mid(sHdg, p, 1)) = 0 Then
         sHdg = Mid(sHdg, 1, p - 1) & Mid(sHdg, p + 1)
      Else
         p = p + 1
      End If
   Loop
   If InStr(sHdg, " AND ") <> 0 Then
      NrmHdgAnd "ANTIGUA AND BARBUDA"
      NrmHdgAnd "ASHMORE AND CARTIER ISLANDS"
      NrmHdgAnd "BOSNIA AND HERCEGOVINA"
      NrmHdgAnd "FRENCH TERRITORY OF THE AFARS AND ISSAS"
      NrmHdgAnd "GILBERT AND ELLICE ISLANDS COLONY"
      NrmHdgAnd "HEARD AND MCDONALD ISLANDS"
      NrmHdgAnd "MUSCAT AND OMAN"
      NrmHdgAnd "NEWFOUNDLAND AND LABRADOR"
      NrmHdgAnd "RHODESIA AND NYASALAND"
      NrmHdgAnd "SAINT KITTS AND NEVIS"
      NrmHdgAnd "SAINT PIERRE AND MIQUELON"
      NrmHdgAnd "SAINT VINCENT AND THE GRENADINES"
      NrmHdgAnd "SAO TOME AND PRINCIPE"
      NrmHdgAnd "SERBIA AND MONTENEGRO"
      NrmHdgAnd "SOUTH GEORGIA AND SOUTH SANDWICH ISLANDS"
      NrmHdgAnd "TRINIDAD AND TOBAGO"
      NrmHdgAnd "TURKS AND CAICOS ISLANDS"
      NrmHdgAnd "WALLIS AND FUTUNA ISLANDS"
   End If
   If InStr(sHdg, "(") <> 0 and InStr(sHdg, ")") <> 0 Then
      NrmHdgRep "(ARMENIA)", "(ARMENIAREP)"
      NrmHdgRep "(GEORGIA)", "(GEORGIAREP)"
      NrmHdgRep "(MACEDONIA)", "(MACEDONIAREP)"
      NrmHdgRep "(MICRONESIA)", "(MICRONESIAFED)"
      NrmHdgRep "(RUSSIA)", "(RUSSIAFED)"
      If InStr(sHdg, " AND ") <> 0 Then
         NrmHdgRep "(ARMENIA AND ", "(ARMENIAREP AND "
         NrmHdgRep " AND ARMENIA)", " AND ARMENIAREP)"
         NrmHdgRep "(GEORGIA AND ", "(GEORGIAREP AND "
         NrmHdgRep " AND GEORGIA)", " AND GEORGIAREP)"
         NrmHdgRep "(MACEDONIA AND ", "(MACEDONIAREP AND "
         NrmHdgRep " AND MACEDONIA)", " AND MACEDONIAREP)"
         NrmHdgRep "(MICRONESIA AND ", "(MICRONESIAFED AND "
         NrmHdgRep " AND MICRONESIA)", " AND MICRONESIAFED)"
         NrmHdgRep "(RUSSIA AND ", "(RUSSIAFED AND "
         NrmHdgRep " AND RUSSIA)", " AND RUSSIAFED)"
      End If
      NrmHdgRep "ARMENIA (REPUBLIC)", "ARMENIAREP"
      NrmHdgRep "COLORADO RIVER (COLOMEXICO)", "COLORADO RIVER"
      NrmHdgRep "CONGO (BRAZZAVILLE)", "CONGOBRA"
      NrmHdgRep "GONGO (DEMOCRATIC REPUBLIC)", "CONGO"
      NrmHdgRep "EAST (US)", "EASTUS"
      NrmHdgRep "GEORGIA (REPUBLIC)", "GEORGIAREP"
      NrmHdgRep "GERMANY (EAST)", "GERMANYEAS"
      NrmHdgRep "GERMANY (WEST)", "GERMANYWES"
      NrmHdgRep "GREAT LAKES (NORTH AMERICA)", "GREAT LAKES"
      NrmHdgRep "GREEN RIVER (WYOUTAH)", "GREEN RIVER"
      NrmHdgRep "GUINEA (REGION)", "GUINEAREG"
      NrmHdgRep "KOREA (NORTH)", "KOREANOR"
      NrmHdgRep "KOREA (SOUTH)", "KOREASOU"
      NrmHdgRep "MACEDONIA (REPUBLIC)", "MACEDONIAREP"
      NrmHdgRep "MICRONESIA (FEDERATED STATES)", "MICRONESIAFED"
      NrmHdgRep "RUSSIA (FEDERATION)", "RUSSIAFED"
      NrmHdgRep "SUDAN (REGION)", "SUDANREG"
      NrmHdgRep "TOGOLAND (BRITISH)", "GHANA"
      NrmHdgRep "TOGOLAND (FRENCH)", "TOGO"
      NrmHdgRep "WEST (US)", "WESTUS"
      NrmHdgDel " (ATLANTIC OCEAN) "
      NrmHdgDel " (COLLECTIVITY) "
      NrmHdgDel " (COLONY) "
      NrmHdgDel " (DWARF PLANET) "
      NrmHdgDel " (EMIRATE) "
      NrmHdgDel " (FEDERATION) "
      NrmHdgDel " (GRAAFSCHAP) "
      NrmHdgDel " (INDIAN OCEAN) "
      NrmHdgDel " (KEELING) "
      NrmHdgDel " (KINGDOM) "
      NrmHdgDel " (ORGANIZATION) "
      NrmHdgDel " (PEOPLES DEMOCRATIC REPUBLIC) "
      NrmHdgDel " (PLANET) "
      NrmHdgDel " (PROTECTORATE) "
      NrmHdgDel " (PROVINCE) "
      NrmHdgDel " (REGION) "
      NrmHdgDel " (STATE) "
      NrmHdgDel " (TRUST TERRITORY) "
      NrmHdgDel " (WEST INDIES) "
      If InStr(sHdg, " (TERRITORY UNDER") <> 0 Then
         sHdg = Left(sHdg, InStr(sHdg, " (TERRITORY UNDER") - 1) & "  "
      End If
   End If
   If InStr(Trim(sHdg), " ") <> 0 Then
      NrmHdgRep "GUINEA, GULF OF, REGION", "GUINEA, GULF OFREG"
      NrmHdgRep "TIBET REGION", "TIBETREG"
      If InStr(sHdg, " WASHINGTON ") <> 0 Then
         NrmHdgRep "WASHINGTON METROPOLITAN AREA", "WASHINGTONREG"
         NrmHdgRep "WASHINGTON REGION", "WASHINGTONREG"
         NrmHdgRep "WASHINGTON SUBURBAN AREA", "WASHINGTONREG"
      End If
      If InStr(sHdg, " WATERSHED ") <> 0 Then
         NrmHdgRep "COLUMBIA RIVER WATERSHED", "COLUMBIA RIVERWAT"
         NrmHdgRep "CONGO RIVER WATERSHED", "CONGO RIVERWAT"
         NrmHdgRep "DANUBE RIVER WATERSHED", "DANUBE RIVERWAT"
      End If
      NrmHdgDel " AUTONOMOUS "
      NrmHdgDel " COAST "
      NrmHdgDel " DELTA "
      NrmHdgDel " ESTUARY "
      NrmHdgDel " METROPOLITAN AREA "
      NrmHdgDel " REGION "
      NrmHdgDel " SUBURBAN AREA "
      NrmHdgDel " VALLEY "
      NrmHdgDel " WATERSHED "
   End If
   sHdg = Trim(sHdg)
   If Right(sHdg, 1) = "," Then
      sHdg = Left(sHdg, Len(sHdg) - 1)
   End If
End Sub

Sub OddGeo
   If sHdg = "GHANA (EMPIRE)" or sHdg = "GUINEAREG" or sHdg = "MALI (EMPIRE)" Then
      sHdg = "AFRICA, WEST"  
   ElseIf sHdg = "AMUR RIVER (CHINA AND RUSSIAFED)" Then
      sHdg = "AMUR RIVER"
   ElseIf InStr(sHdg, "(AUSTRALIA)") <> 0 or InStr(sHdg, "AUSTRALIA \Z ") <> 0 Then
      If sHdg = "PACIFIC COAST (AUSTRALIA)" or InStr(sHdg, "AUSTRALIA \Z PACIFIC COAST") <> 0 Then
         sHdg = "EASTERN AUSTRALIA"
      ElseIf InStr(sHdg, "AUSTRALIA \Z ") <> 0 Then
         If InStr(sHdg, "(ACT)") <> 0 or InStr(sHdg, "AUSTRALIA \Z AUSTRALIAN CAPITAL TERRITORY") <> 0 Then
            sHdg = "AUSTRALIAN CAPITAL TERRITORY"
         ElseIf InStr(sHdg, "(CENTRAL AUSTRALIA)") <> 0 or InStr(sHdg, "AUSTRALIA \Z CENTRAL AUSTRALIA") <> 0 Then
            sHdg = "CENTRAL AUSTRALIA"
         ElseIf InStr(sHdg, "(JERVIS BAY TERRITORY)") <> 0 or InStr(sHdg, "AUSTRALIA \Z JERVIS BAY TERRITORY") <> 0 Then
            sHdg = "JERVIS BAY TERRITORY"
         ElseIf InStr(sHdg, "(NSW)") <> 0 or InStr(sHdg, "AUSTRALIA \Z NEW SOUTH WALES") <> 0 Then
            sHdg = "NEW SOUTH WALES"
         ElseIf InStr(sHdg, "(NT)") <> 0 or InStr(sHdg, "AUSTRALIA \Z NORTHERN TERRITORY") <> 0 Then
            sHdg = "NORTHERN TERRITORY"
         ElseIf InStr(sHdg, "(QLD)") <> 0 or InStr(sHdg, "AUSTRALIA \Z QUEENSLAND") <> 0 Then
            sHdg = "QUEENSLAND"
         ElseIf InStr(sHdg, "(S AUST)") <> 0 or InStr(sHdg, "AUSTRALIA \Z SOUTH AUSTRALIA") <> 0 Then
            sHdg = "SOUTH AUSTRALIA"
         ElseIf InStr(sHdg, "(TAS)") <> 0 or InStr(sHdg, "AUSTRALIA \Z TASMANIA") <> 0 Then
            sHdg = "TASMANIA"
         ElseIf InStr(sHdg, "(VIC)") <> 0 or InStr(sHdg, "AUSTRALIA \Z VICTORIA") <> 0 Then
            sHdg = "VICTORIA"
         ElseIf InStr(sHdg, "(WA)") <> 0 or InStr(sHdg, "AUSTRALIA \Z WESTERN AUSTRALIA") <> 0 Then
            sHdg = "WESTERN AUSTRALIA"
         ElseIf InStr(sHdg, "\Z") <> 0 and (InStr(sHdg, "(ACT AND ") <> 0 or InStr(sHdg, "(JERVIS BAY TERRITORY AND ") <> 0 or InStr(sHdg, "(NSW AND ") <> 0 or InStr(sHdg, "(NT AND ") <> 0 or InStr(sHdg, "(QLD AND ") <> 0 or InStr(sHdg, "(S AUST AND ") <> 0 or InStr(sHdg, "(TAS AND ") <> 0 or InStr(sHdg, "(VIC AND ") <> 0 or InStr(sHdg, "(WA AND ") <> 0) Then
            sHdg = Mid(sHdg, InStr(sHdg, "\Z") + 3, Len(sHdg))
         End If
      End If
   ElseIf sHdg = "EAST BENGAL (PAKISTAN)" or sHdg = "EAST PAKISTAN (PAKISTAN)" Then
      sHdg = "BANGLADESH"
   ElseIf sHdg = "VIRGIN ISLAND (GREAT BRITAIN)" Then
      sHdg = "BRITISH VIRGIN ISLANDS"
   ElseIf InStr(sHdg, "PANAMA \Z CANAL ZONE") <> 0 Then
      sHdg = "CANAL ZONE"
   ElseIf InStr(sHdg, "(CHINA)") <> 0 or InStr(sHdg, "CHINA \Z ") <> 0 Then
      If sHdg = "ANHUI SHENG (CHINA)" or InStr(sHdg, "CHINA \Z ANHUI SHENG") <> 0 Then
         sHdg = "ANHUI SHENG"
      ElseIf sHdg = "BEIJING (CHINA)" or InStr(sHdg, "CHINA \Z BEIJING") <> 0 Then
         sHdg = "BEIJING"
      ElseIf sHdg = "BO HAI (CHINA)" or InStr(sHdg, "CHINA \Z BO HAI") <> 0 Then
         sHdg = "BO HAI"
      ElseIf sHdg = "CHONGQING (CHINA)" or InStr(sHdg, "CHINA \Z CHONGQING") <> 0 Then
         sHdg = "CHONGQING"
      ElseIf sHdg = "FUJIAN SHENG (CHINA)" or InStr(sHdg, "CHINA \Z FUJIAN SHENG") <> 0 Then
         sHdg = "FUJIAN SHENG"
      ElseIf sHdg = "GANSU SHENG (CHINA)" or InStr(sHdg, "CHINA \Z GANSU SHENG") <> 0 Then
         sHdg = "GANSU SHENG"
      ElseIf sHdg = "GUANGDONG SHENG (CHINA)" or InStr(sHdg, "CHINA \Z GUANGDONG SHENG") <> 0 Then
         sHdg = "GUANGDONG SHENG"
      ElseIf sHdg = "GUANGXI ZHUANGZU ZIZHIQU (CHINA)" or InStr(sHdg, "CHINA \Z GUANGXI ZHUANGZU ZIZHIQU") <> 0 Then
         sHdg = "GUANGXI ZHUANGZU ZIZHIQU"
      ElseIf sHdg = "GUIZHOU SHENG (CHINA)" or InStr(sHdg, "CHINA \Z GUIZHOU SHENG") <> 0 Then
         sHdg = "GUIZHOU SHENG"
      ElseIf sHdg = "HAINAN SHENG (CHINA)" or InStr(sHdg, "CHINA \Z HAINAN SHENG") <> 0 Then
         sHdg = "HAINAN SHENG"
      ElseIf sHdg = "HEBEI SHENG (CHINA)" or InStr(sHdg, "CHINA \Z HEBEI SHENG") <> 0 Then
         sHdg = "HEBEI SHENG"
      ElseIf sHdg = "HEILONGJIANG SHENG (CHINA)" or InStr(sHdg, "CHINA \Z HEILONGJIANG SHENG") <> 0 Then
         sHdg = "HEILONGJIANG SHENG"
      ElseIf sHdg = "HENAN SHENG (CHINA)" or InStr(sHdg, "CHINA \Z HENAN SHENG") <> 0 Then
         sHdg = "HENAN SHENG"
      ElseIf sHdg = "HONG KONG (CHINA)" or InStr(sHdg, "CHINA \Z HONG KONG") <> 0 Then
         sHdg = "HONG KONG"
      ElseIf sHdg = "HUBEI SHENG (CHINA)" or InStr(sHdg, "CHINA \Z HUBEI SHENG") <> 0 Then
         sHdg = "HUBEI SHENG"
      ElseIf sHdg = "HUNAN SHENG (CHINA)" or InStr(sHdg, "CHINA \Z HUNAN SHENG") <> 0 Then
         sHdg = "HUNAN SHENG"
      ElseIf sHdg = "INNER MONGOLIA (CHINA)" or InStr(sHdg, "CHINA \Z INNER MONGOLIA") <> 0 Then
         sHdg = "INNER MONGOLIA"
      ElseIf sHdg = "JIANGSU SHENG (CHINA)" or InStr(sHdg, "CHINA \Z JIANGSU SHENG") <> 0 Then
         sHdg = "JIANGSU SHENG"
      ElseIf sHdg = "JIANGXI SHENG (CHINA)" or InStr(sHdg, "CHINA \Z JIANGXI SHENG") <> 0 Then
         sHdg = "JIANGXI SHENG"
      ElseIf sHdg = "JILIN SHENG (CHINA)" or InStr(sHdg, "CHINA \Z JILIN SHENG") <> 0 Then
         sHdg = "JILIN SHENG"
      ElseIf sHdg = "LIAONING SHENG (CHINA)" or InStr(sHdg, "CHINA \Z LIAONING SHENG") <> 0 Then
         sHdg = "LIAONING SHENG"
      ElseIf sHdg = "MACAU (CHINA)" or InStr(sHdg, "CHINA \Z MACAU") <> 0 Then
         sHdg = "MACAU"
      ElseIf sHdg = "NINGXIA HUIZU ZIZHIQU (CHINA)" or InStr(sHdg, "CHINA \Z NINGXIA HUIZU ZIZHIQU") <> 0 Then
         sHdg = "NINGXIA HUIZU ZIZHIQU"
      ElseIf sHdg = "QINGHAI SHENG (CHINA)" or InStr(sHdg, "CHINA \Z QINGHAI SHENG") <> 0 Then
         sHdg = "QINGHAI SHENG"
      ElseIf sHdg = "SHAANXI SHENG (CHINA)" or InStr(sHdg, "CHINA \Z SHAANXI SHENG") <> 0 Then
         sHdg = "SHAANXI SHENG"
      ElseIf sHdg = "SHANDONG SHENG (CHINA)" or InStr(sHdg, "CHINA \Z SHANDONG SHENG") <> 0 Then
         sHdg = "SHANDONG SHENG"
      ElseIf sHdg = "SHANGHAI (CHINA)" or InStr(sHdg, "CHINA \Z SHANGHAI") <> 0 Then
         sHdg = "SHANGHAI"
      ElseIf sHdg = "SHANXI SHENG (CHINA)" or InStr(sHdg, "CHINA \Z SHANXI SHENG") <> 0 Then
         sHdg = "SHANXI SHENG"
      ElseIf sHdg = "SICHUAN SHENG (CHINA)" or InStr(sHdg, "CHINA \Z SICHUAN SHENG") <> 0 Then
         sHdg = "SICHUAN SHENG"
      ElseIf sHdg = "TIANJIN (CHINA)" or InStr(sHdg, "CHINA \Z TIANJIN") <> 0 Then
         sHdg = "TIANJIN"
      ElseIf sHdg = "TIBET (CHINA)" or InStr(sHdg, "CHINA \Z TIBET") <> 0 Then
         sHdg = "TIBET"
      ElseIf sHdg = "XI RIVER (CHINA)" or InStr(sHdg, "CHINA \Z XI RIVER") <> 0 Then
         sHdg = "XI RIVER"
      ElseIf sHdg = "XINJIANG UYGUR ZIZHIQU (CHINA)" or InStr(sHdg, "CHINA \Z XINJIANG UYGUR ZIZHIQU") <> 0 Then
         sHdg = "XINJIANG UYGUR ZIZHIQU"
      ElseIf sHdg = "YANGTZE RIVER (CHINA)" or sHdg = "YANGTZE RIVER GORGES (CHINA)" or InStr(sHdg, "CHINA \Z YANGTZE RIVER") <> 0 Then
         sHdg = "YANGTZE RIVER"
      ElseIf sHdg = "YELLOW RIVER (CHINA)" or InStr(sHdg, "CHINA \Z YELLOW RIVER") <> 0 Then
         sHdg = "YELLOW RIVER"
      ElseIf sHdg = "YUNNAN SHENG (CHINA)" or InStr(sHdg, "CHINA \Z YUNNAN SHENG") <> 0 Then
         sHdg = "YUNNAN SHENG"
      ElseIf sHdg = "ZHEJIANG SHENG (CHINA)" or InStr(sHdg, "CHINA \Z ZHEJIANG SHENG") <> 0 Then
         sHdg = "ZHEJIANG SHENG"
      End If
   ElseIf sHdg = "BRAZZAVILLE (CONGO)" Then
      sHdg = "CONGOBRA"
   ElseIf sHdg = "ERITREA (ETHIOPIA)" Then
      sHdg = "ERITREA"
   ElseIf InStr(sHdg, "HURON, LAKE (MICH AND ONT)") <> 0 Then
      sHdg = "HURON, LAKE"
   ElseIf InStr(sHdg, "KOSOVO (REPUBLIC)") <> 0 or InStr(sHdg, "KOSOVO (SERBIA)") <> 0 Then
      sHdg = "KOSOVO"
   ElseIf sHdg = "KUNLUN MOUNTAINS (CHINA AND INDIA)" Then
      sHdg = "KUNLUN MOUNTAINS"
   ElseIf InStr(sHdg, "MALAYSIA \Z MALAYA") <> 0 Then
      sHdg = "MALAYA"
   ElseIf sHdg = "KATTEGAT (DENMARK AND SWEDEN)" <> 0 Then
      sHdg = "NORTH ATLANTIC OCEAN"
   ElseIf sHdg = "ALASKA, GULF OF (ALASKA)" or InStr(sHdg, "ALASKA \Z ALASKA, GULF OF") <> 0 Then
      sHdg = "NORTH PACIFIC OCEAN"
   ElseIf InStr(sHdg, "ONTARIO, LAKE (NY AND ONT)") <> 0 Then
      sHdg = "ONTARIO, LAKE"
   ElseIf sHdg = "NEW GUINEA (TERRITORY)" Then
      sHdg = "PAPUA NEW GUINEA"
   ElseIf sHdg = "EAST PRUSSIA (GERMANY)" Then
      sHdg = "PRUSSIA, EAST"
   ElseIf sHdg = "RIO DE LA PLATA (ARGENTINA AND URUGUAY)" Then
      sHdg = "RIO DE LA PLATA"
   ElseIf InStr(sHdg, "(RUSSIA)") <> 0 or InStr(sHdg, "RUSSIAFED") <> 0 Then
      If sHdg = "CAUCASUS, NORTHERN (RUSSIAFED)" or InStr(sHdg, "RUSSIAFED \Z CAUCASUS, NORTHERN") <> 0 Then
         sHdg = "CAUCASUS, NORTHERN"
      ElseIf sHdg = "CENTRAL CHERNOZEM (RUSSIAFED)" or InStr(sHdg, "RUSSIAFED \Z CENTRAL CHERNOZEM") <> 0 Then
        sHdg = "CENTRAL CHERNOZEM"
      ElseIf sHdg = "RUSSIAN FAR EAST (RUSSIAFED)" or InStr(sHdg, "RUSSIAFED \Z RUSSIAN FAR EAST") <> 0 Then
         sHdg = "RUSSIAN FAR EAST"
      ElseIf sHdg = "SIBERIA, EASTERN (RUSSIAFED)" or InStr(sHdg, "RUSSIAFED \Z SIBERIA, EASTERN") <> 0 Then
         sHdg = "SIBERIA, EASTERN"
      ElseIf sHdg = "SIBERIA, WESTERN (RUSSIAFED)" or InStr(sHdg, "RUSSIAFED \Z SIBERIA, WESTERN") <> 0 Then
         sHdg = "SIBERIA, WESTERN"
      ElseIf sHdg = "SIBERIA (RUSSIAFED)" or sHdg = "SIBERIA, NORTHEASTERN (RUSSIAFED)" or sHdg = "SIBERIA, NORTHWESTERN (RUSSIAFED)" or InStr(sHdg, "RUSSIAFED \Z SIBERIA") <> 0 or InStr(sHdg, "RUSSIAFED \Z SIBERIA, NORTHEASTERN") <> 0 or InStr(sHdg, "RUSSIAFED \Z SIBERIA, NORTHWESTERN") <> 0 Then
         sHdg = "SIBERIA"
      ElseIf sHdg = "VOLGA RIVER (RUSSIAFED)" or InStr(sHdg, "RUSSIAFED \Z VOLGA RIVER") <> 0 Then
         sHdg = "VOLGA RIVER"
      ElseIf sHdg = "URAL MOUNTAINS (RUSSIAFED)" or InStr(sHdg, "RUSSIAFED \Z URAL MOUNTAINS") <> 0 Then
         sHdg = "URAL MOUNTAINS"
      End If
   ElseIf sHdg = "CEUTA (SPAIN)" or sHdg = "MELILLA (SPAIN)" or InStr(sHdg, "SPAIN \Z CEUTA") <> 0 or InStr(sHdg, "SPAIN \Z MELILLA") <> 0 Then
      sHdg = "SPANISH NORTH AFRICA"
   ElseIf sHdg = "SUEZ CANAL (EGYPT)" or InStr(sHdg, "EGYPT \Z SUEZ CANAL") <> 0 Then
      sHdg = "SUEZ CANAL"
   ElseIf sHdg = "BEAR ISLAND (NORWAY)" or sHdg = "SPITSBERGEN ISLAND (NORWAY)" or sHdg = "SVALBARD (NORWAY)" or InStr(sHdg, "NORWAY \Z BEAR ISLAND") <> 0 or InStr(sHdg, "NORWAY \Z SPITSBERGEN ISLAND") <> 0 or InStr(sHdg, "NORWAY \Z SVALBARD") <> 0 Then
      sHdg = "SVALBARD"
   ElseIf sHdg = "SWAN ISLANDS (HONDURAS)" or InStr(sHdg, "HONDURAS \Z SWAN ISLANDS") <> 0 Then
      sHdg = "SWAN ISLANDS"
   ElseIf InStr(sHdg, "EAST TIMOR (") <> 0 or sHdg = "TIMOR TIMUR (INDONESIA)" Then
      sHdg = "TIMORLESTE"
   ElseIf sHdg = "ELLICE ISLANDS (GILBERT & ELLICE ISLANDS COLONY)" Then
      sHdg = "TUVALU"
   ElseIf InStr(sHdg, "VIETNAM (") <> 0 Then
      sHdg = "VIETNAM"
   ElseIf sHdg = "VOLTA RIVER (GHANA)" or InStr(sHdg, "GHANA \Z VOLTA RIVER") <> 0 Then
      sHdg = "VOLTA RIVER"
   ElseIf InStr(sHdg, "YEMEN (") <> 0 Then
      sHdg = "YEMEN"
   End If
End Sub

Sub MulGeo
   Dim sHdg1$, sHdg2$
   sHdg1 = Left(sHdg, InStr(sHdg, " AND ") - 1)
   sHdg2 = Mid(sHdg, InStr(sHdg, " AND ") + 5)
   sNewGac = GetGac(sHdg1)
   If InStr(s043, sNewGac) = 0 and sNewGac <> "X" Then
      s043 = s043 & " | " & sNewGac
   End If
   sNewGac = GetGac(sHdg2)
   If InStr(s043, sNewGac) = 0 and sNewGac <> "X" Then
      s043 = s043 & " | " & sNewGac
   End If 
End Sub

Sub Add043
   Dim sText$
   If s043 <> "" Then
      Do While InStr(s043, " | ") <> 0
         sText = Mid(s043, InStr(s043, " | ") + 3, 7)
         s043 = Left(s043, InStr(s043, " | ") - 1) & " ßa " & Mid(s043, InStr(s043, " | ") + 3)
         Do While InStr(s043, " | " & sText) <> 0
            s043 = Left(s043, InStr(s043, " | " & sText) - 1) & Mid(s043, InStr(s043, " | " & sText) + 10)
         Loop
      Loop
      s043 = "043  " & Mid(s043, 5)
      If CS.SetFieldLine(999, s043) = True Then
         CS.Reformat
      End If
   End If
End Sub

Sub UseGenerate043
   Dim n%, nBool%
   Set CS = CreateObject("Connex.Client")
   If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 2 or CS.ItemType = 3 or CS.ItemType = 4 or CS.ItemType = 14 or CS.ItemType = 17 or CS.ItemType = 18 or CS.ItemType = 19 or CS.ItemType = 20 or CS.ItemType = 35 Then
      
      
      s043 = ""
      DefGeo
      DefGac
      If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 2 or CS.ItemType = 17 or CS.ItemType = 19 or CS.ItemType = 35 Then
         n = 6
         nBool = CS.GetFieldLine(n, sHdg)
         Do While nBool = True
            If InStr("600,610,611,630,650,651", Mid(sHdg, 1, 3)) <> 0 and InStr("01235", Mid(sHdg, 5, 1)) <> 0 Then
               If Left(sHdg, 3) = "651" or (Left(sHdg, 3) = "610" and Mid(sHdg, 4, 1) = "1") Then
                  sHdg = Mid(sHdg, 6, Len(sHdg))
                  If InStr(sHdg, "ß") <> 0 Then
                     sHdg = Left(sHdg, InStr(sHdg, "ß") - 1)
                  End If
                  NrmHdg
                  OddGeo
                  If InStr(sHdg, "(") <> 0 and InStr(sHdg, ")") <> 0 Then
                     sHdg = Mid(sHdg, InStr(sHdg, "(") + 1)
                     sHdg = Left(sHdg, Len(sHdg) - 1)
                     Do While InStr(sHdg, ",") <> 0
                        sHdg = Mid(sHdg, InStr(sHdg, ",") + 2)
                     Loop
                  End If
                  If InStr(sHdg, " AND ") = 0 Then
                     sNewGac = GetGac(sHdg)
                     If sHdg <> "X" and sNewGac <> "X" Then
                        s043 = s043 & " | " & sNewGac
                     End If
                  Else
                     MulGeo
                  End If
               End If
               If InStr(sHdg, "ßz") <> 0 Then
                  sHdg = Mid(sHdg, InStr(sHdg, "ßz") + 3)
                  NrmHdg
                  OddGeo
                  If InStr(sHdg, "\Z") <> 0 Then
                     sHdg = Mid(sHdg, 1, InStr(sHdg, "\Z") - 2)
                  End If
                  If InStr(sHdg, "(") <> 0 and InStr(sHdg, ")") <> 0 Then
                     sHdg = Mid(sHdg, InStr(sHdg, "(") + 1)
                     sHdg = Left(sHdg, Len(sHdg) - 1)
                     Do While InStr(sHdg, ",") <> 0
                        sHdg = Mid(sHdg, InStr(sHdg, ",") + 2)
                     Loop
                  End If
                  If InStr(sHdg, " AND ") = 0 Then
                     sNewGac = GetGac(sHdg)
                     If sHdg <> "X" and sNewGac <> "X" Then
                        s043 = s043 & " | " & sNewGac
                     End If
                  Else
                     MulGeo
                  End If
               End If
            End If
            n = n + 1
            nBool = CS.GetFieldLine(n, sHdg)
         Loop
      ElseIf CS.ItemType = 3 or CS.ItemType = 4 or CS.ItemType = 14 or CS.ItemType = 18 or CS.ItemType = 20 Then
         nBool = CS.GetField("781", 1, sHdg)
         If nBool = False Then
            nBool = CS.GetField("151", 1, sHdg)
         End If
         If nBool = True Then
            If Left(sHdg, 3) = "151" or Left(sHdg, 3) = "781" Then
               sHdg = Mid(sHdg, 6, Len(sHdg))
               If InStr(sHdg, "ß") <> 0 Then
                  sHdg = Left(sHdg, InStr(sHdg, "ß") - 1)
               End If
               NrmHdg
               OddGeo
               If InStr(sHdg, "(") <> 0 and InStr(sHdg, ")") <> 0 Then
                  sHdg = Mid(sHdg, InStr(sHdg, "(") + 1)
                  sHdg = Left(sHdg, Len(sHdg) - 1)
                  Do While InStr(sHdg, ",") <> 0
                     sHdg = Mid(sHdg, InStr(sHdg, ",") + 2)
                  Loop
               End If
               If InStr(sHdg, " AND ") = 0 Then
                  sNewGac = GetGac(sHdg)
                  If sHdg <> "X" and sNewGac <> "X" Then
                     s043 = s043 & " | " & sNewGac
                  End If
               Else
                  MulGeo
               End If
               If InStr(sHdg, "ßz") <> 0 Then
                  sHdg = Mid(sHdg, InStr(sHdg, "ßz") + 3)
                  NrmHdg
                  If InStr(sHdg, "ß") <> 0 Then
                     sHdg = Left(sHdg, InStr(sHdg, "ß") - 2)
                  End If
                  OddGeo
                  If InStr(sHdg, "\Z") <> 0 Then
                     sHdg = Left(sHdg, InStr(sHdg, "\Z") - 2)
                  End If
                  If InStr(sHdg, "(") <> 0 and InStr(sHdg, ")") <> 0 Then
                     sHdg = Mid(sHdg, InStr(sHdg, "(") + 1)
                     sHdg = Left(sHdg, Len(sHdg) - 1)
                     Do While InStr(sHdg, ",") <> 0
                        sHdg = Mid(sHdg, InStr(sHdg, ",") + 2)
                     Loop
                  End If
                  If InStr(sHdg, " AND ") = 0 Then
                     sNewGac = GetGac(sHdg)
                     If sHdg <> "X" and sNewGac <> "X" Then
                        s043 = s043 & " | " & sNewGac
                     End If
                  Else
                     MulGeo
                  End If
               End If
            End If
         End If
      End If
      Add043
   Else
      MsgBox "An authority or bibliographic record must be displayed in order to use this macro", 0, "Generate 043 Macro"
   End If
End Sub