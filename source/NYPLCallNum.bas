'MacroName:NYPL CallNum v.2.5.9
'MacroDescription: NYPL macro for creating a complete call number in field 948 based on catalogers selected pattern and information coded in the record
'                  Macro handles call number patterns for English and World Lanugages, fiction, non-fiction, biography and biography with Dewey
'                  incorporates functions of Format macro - populates subfielf $f
'Macro created by: Tomasz Kalata, BookOps
'Latest update: March 07, 2019

'v.2.5.9 update details:
'  * complete elimination of 11 characters rule in cutter for non-latin materials - only last name used or first letter
'    (cutter in subfield $c)

'v.2.5.8 update details:
'  * Bug fix: removes period as the last character in cutter (the bug introduced in 2.5.6)

'v 2.5.7 update details:
'  * Readalong call number prefix added
'  * Validation flats added for Readalong call numbers

'v.2.5.6 update details:
'  * DVD & Bluray call number cutter includes full first word
'  * bug fix: underscore nonspacing character (Chr(246)) handling fixed

'v.2.5.5 update details:
'  * bug fix: proper behaviour if no bibliographic record displayed
'  * added flag for short stories collections

'v.2.5.4 update details:
'  *bug fix: removes flase fiction flag for Graphic Novels
'  *improvement: diactritics function simplified and made more comprehensive
'  *bug fix: corrected broken format error flags
'  *rules change: authors name for literary collection changed from $b to $c
'  *improvement: added validation error flag for Dewey + Name call numbers that are not in 7xx or 8xx range


Option Explicit

Declare Function Dewey(a)
Declare Sub ElemArray(sElemArr, n100, n600)
Declare Sub ElemArrayChoice(sCallType, sElemArr, sNameChoice, n100, n600)
Declare Sub Diacritics(sNameTitle)
Declare Sub Rules(sElemArr, sCallType, sLang, sCutter, sNameChoice)
Declare Sub InsertCallNum(s948, f, sInitials)
Declare Sub Validation(a, f, sAudn, sCallType, sCont, sItemForm, sLang, sRecType, sTmat, sLitF, sBiog, s948)

'temporary variables
Dim place, i as Integer
Dim lt$, rt$, sTemp$

Sub Main
   Dim CS as Object
   Set CS = CreateObject("Connex.Client")
   If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 2 or CS.ItemType = 17 Then
      Dim s300$, s538$, s948$, sAudn$, sAudnLangPrefix$, sBiog$, sCallType$, sCont$, sCutter$, sElemArr$, _
         sFormatPrefix$, sItemForm$, sLang$, sLitF$, sNameChoice$, sRecType$, sTMat$
      Dim bool538
      Dim a, f, n100, n600 As Integer
      Dim sFormat() As String
      Dim sAudience() As String
      Dim sOutput() As String

      Dim filenumber As Integer
      Dim sFileName As String
      Dim sDefaultInitials$
      Dim sInitials$

      CS.GetFixedField "Type", sRecType
      CS.GetFixedField "Audn", sAudn
      CS.GetFixedField "Lang", sLang
      CS.GetFixedField "Form", sItemForm
      CS.GetFixedField "Cont", sCont
      CS.GetFixedField "TMat", sTMat
      If sRecType = "a" Then
         CS.GetFixedField "LitF", sLitF
         CS.GetFixedField "Biog", sBiog
      End If
      CS.GetFieldUnicode "300", 1, s300
      If InStr(s300, Chr(223) & "e") <> 0 Then
         s300 = Mid(s300, InStr(s300, Chr(223) & "e"))
      End If
      s300 = UCase(s300)

      ReDim sFormat(11)
         sFormat(0) = " "
         sFormat(1) = "BLURAY"
         sFormat(2) = "CD"
         sFormat(3) = "CLASSICS"
         sFormat(4) = "DVD"
         sFormat(5) = "FINANCE"
         sFormat(6) = "GRAPHIC"
         sFormat(7) = "HOLIDAY"
         sFormat(8) = "JOBS"
         sFormat(9) = "LG PRINT"
         sFormat(10) = "YR"
         sFormat(11) = "READALONG"
      ReDim sAudience(1)
         sAudience(0) = "JUVENILE"
         sAudience(1) = "ADULT"
      ReDim sOutput(1)
         sOutput(0) = "add to record"
         sOutput(1) = "add to clipboard"

      'read default data (initials) from text file stored in macro folder
      sFileName = "cat_data.txt"
      If Dir$ (sFileName) <> "" Then
         filenumber = FreeFile
         Open sFileName for Input As filenumber
         Line Input #filenumber, sDefaultInitials
         Close #filenumber
      Else
         'create a file with blank values if not present in the main folder
         filenumber = FreeFile
         Open sFileName For Output As filenumber
         Print #filenumber, "XXX"
         Close #filenumber
         sDefaultInitials = "XXX"
      End If


      'Dialog box presenting to a cataloger choices for types of call numbers
      Begin Dialog MainWindow 220, 220, "NYPL Call Number Macro v. 2.5.3"

         'top-left outline
         GroupBox 18, 50, 85, 41, ""
         'bottom-left outline
         GroupBox 18, 90, 85, 100, ""
         'top-right outline
         GroupBox 120, 50, 82, 100, ""
         'bottom-right outline
         GroupBox 120, 149, 82, 41, ""
         OptionGroup .Type
         OptionButton  24,  55, 70, 14, "&EASY BOOKS"
         OptionButton  24,  75, 70, 14, "&PICTURE BOOKS"
         OptionButton  24,  95, 70, 14, "&BIOGRAPHY"
         OptionButton  24,  115, 70, 14, "&DEWEY"
         OptionButton  24,  135, 70, 14, "DEWEY + &NAME"
         OptionButton  24,  155, 70, 14, "&FICTION"
         OptionButton  24,  175, 70, 14, "&GN FICTION"
         OptionButton  125,  55, 70, 14, "&MYSTERY"
         OptionButton  125,  75, 70, 14, "&ROMANCE"
         OptionButton  125,  95, 70, 14, "&SCI FI"
         OptionButton  125,  115, 70, 14, "&URBAN"
         OptionButton  125,  135, 70, 14, "&WESTERN"
         OptionButton  125,  155, 70, 14, "M&OVIE"
         OptionButton  125,  175, 70, 14, "&TV"
         DropListBox  18, 10, 61, 115, sFormat(), .sFormat
         Text 81, 12, 70, 14, "FORMAT"
         DropListBox  18, 30, 61, 40, sAudience(), .sAudience
         Text 81, 32, 78, 14, "AUDIENCE"
         Textbox 130, 12, 20, 15, .sInitials
         Text 160, 14, 30, 30, "INITIALS"
         'Text 150, 12, 70, 14, "Output:"
         DropListBox  130, 30, 68, 40, sOutput(), .sOutput
         OkButton        50, 195,  55, 16
         CancelButton   115, 195,  55, 16
      End Dialog
      Dim dCallNum as MainWindow
      'selects most likely audience, format
      If InStr("abcj", sAudn) <> 0 And sAudn <> "" Then
         dCallNum.sAudience = 0
      Else
         dCallNum.sAudience = 1
      End If
      If sRecType = "a" Then
         If InStr(sCont, "6") <> 0 Then
            dCallNum.sFormat = 6
         ElseIf sItemForm = "d" Then
            dCallNum.sFormat = 9
         End If
      ElseIf sRecType = "i" Then
         If InStr(s300, "AUDIO-ENABLED BOOK") <> 0 Or InStr(s300, "AUDIO ENABLED BOOK") <> 0 Then
            dCallNum.sFormat = 11
         Else
            dCallNum.sFormat = 2
         End If
      ElseIf sRecType = "j" Then
         MsgBox "Please consider using NYPLMusicCD macro instead. The record appeared to be a music CD"
      ElseIf sRecType = "g" And sTMat = "v" Then
         bool538 = CS.GetField("538", 1, s538)
         If bool538 = TRUE Then
            s538 = Mid(LCase(s538), 6)
            If InStr(s538, "blu-ray") <> 0 Or InStr(s538, "bluray") <> 0 Then
               dCallNum.sFormat = 1
            Else
               dCallNum.sFormat = 4
            End If
         Else
            dCallNum.sFormat = 4
         End If
      End If
      'populate INITIALS box with default value
      dCallNum.sInitials = sDefaultInitials

      On Error Resume Next
      Dialog dCallNum
      If Err = 102 Then Exit Sub


'  define subfield $p of 948 field
      a = dCallNum.sAudience
      sLang = UCase(sLang)
      If sLang = "" Or sLang = "UND" Then
         sLang = Chr(252)
         MsgBox "INCOMPLETE: Please correct languge coding in the fixed field and the call number."
      End If
      If a = 0 Then
         If sLang <> "ENG" Then
            sAudnLangPrefix = Chr(223) & "p J " & sLang & " "
         Else
            sAudnLangPrefix = Chr(223) & "p J" & " "
         End If
      Else
         If sLang <> "ENG" Then
            sAudnLangPrefix = Chr(223) & "p " & sLang & " "
         End If
      End If

'  define subfield $f of 948 field
      f = dCallNum.sFormat
      If f <> 0 Then
         sFormatPrefix = Chr(223) & "f " & sFormat(dCallNum.sFormat) & " "
      End If

      Call ElemArray(sElemArr, n100, n600)

'  define the reminder of 948 field
      s948 = "948  " & sAudnLangPrefix & sFormatPrefix
      Select Case dCallNum.Type
         Case 0
            sCallType = "eas"
            dCallNum.sAudience = 0
            s948 = s948 & Chr(223) & "a E"
         Case 1
            sCallType = "pic"
            dCallNum.sAudience = 0
            s948 = s948 & Chr(223) & "a PIC"
         Case 2
            sCallType = "bio"
            'add selection of name if multiple 600 here then pass variable to rules
            Call ElemArrayChoice(sCallType, sElemArr, sNameChoice, n100, n600)
            s948 = s948 & Chr(223) & "a B"
         Case 3
            sCallType = "dew"
            s948 = s948 & Dewey(a)
         Case 4
            sCallType = "den"
            'add selection of name if 100 and 600 here then pass variable to rules
            Call ElemArrayChoice(sCallType, sElemArr, sNameChoice, n100, n600)
            s948 = s948 & Dewey(a)
         Case 5
            sCallType = "fic"
            s948 = s948 & Chr(223) & "a FIC"
         Case 6
            sCallType = "gfi"
            dCallNum.sFormat = 7
            s948 = s948 & Chr(223) & "a GN FIC"
         Case 7
            sCallType = "mys"
            s948 = s948 & Chr(223) & "a MYSTERY"
         Case 8
            sCallType = "rom"
            s948 = s948 & Chr(223) & "a ROMANCE"
         Case 9
            sCallType = "sci"
            s948 = s948 & Chr(223) & "a SCI FI"
         Case 10
            sCallType = "urb"
            s948 = s948 & Chr(223) & "a URBAN"
         Case 11
            sCallType = "wes"
            s948 = s948 & Chr(223) & "a WESTERN"
         Case 12
            sCallType = "mov"
            s948 = s948 & Chr(223) & "a MOVIE"
         Case 13
            sCallType = "tvs"
            s948 = s948 & Chr(223) & "a TV"
      End Select

      Call Rules(sElemArr, sCallType, sLang, sCutter, sNameChoice)
      s948 = s948 & sCutter

      'Output selection
      If dCallNum.sOutput = 1 Then
         s948 = Mid(s948, 6)
         Do While InStr(s948, Chr(223)) <> 0
            place = InStr(s948, Chr(223))
            lt$ = Left(s948, place - 1)
            rt$ = Mid(s948, place + 3)
            s948 = lt$ + rt$
         Loop
         Do While InStr(s948, Chr(252)) <> 0
            place = InStr(s948, Chr(252))
            lt$ = Left(s948, place - 1)
            rt$ = Mid(s948, place + 1)
            s948 = lt$ & "[empty]" & rt$
         Loop
         Clipboard.Clear
         Clipboard.SetText s948
      Else
         'populate default initials
         sInitials = dCallNum.sInitials
         filenumber = FreeFile
         Open sFileName For Output As filenumber
         Print #filenumber, sInitials
         Close #filenumber

         'insert call number & other strings
         Call InsertCallNum(s948, f, sInitials)
      End If

   Else
      MsgBox "INFO: A bibliographic record must be displayed in order to use this macro."
      Goto ReallyDone
   End If

Done:
   Call Validation(a, f, sAudn, sCallType, sCont, sItemForm, sLang, sRecType, sTmat, sLitF, sBiog, s948)
ReallyDone:
End Sub

'########################################################################

Sub Rules(sElemArr, sCallType, sLang, sCutter, sNameChoice)
   'rules of creating reminder of the call number
   Dim sLastChr$, sMainEntry$
   Dim start_point, end_point As Integer


   'find main entry
   If InStr(sElemArr, "100: ") <> 0 Then
      start_point = InStr(sElemArr, "100: ")
      sTemp = Mid(sElemArr, start_point)
      end_point = InStr(sTemp, Chr(9))
      sMainEntry = Left(sTemp, end_point - 1)
   ElseIf InStr(sElemArr, "110: ") <> 0 Then
      start_point = InStr(sElemArr, "110: ")
      sTemp = Mid(sElemArr, start_point)
      end_point = InStr(sTemp, Chr(9))
      sMainEntry = Left(sTemp, end_point - 1)
   Else
      If InStr(sElemArr, "245: ") <> 0 Then
         start_point = InStr(sElemArr, "245: ")
         sTemp = Mid(sElemArr, start_point)
         end_point = InStr(sTemp, Chr(9))
         sMainEntry = Left(sTemp, end_point - 1)
      Else
         MsgBox "MISSING INFO: No valid MARC field for main entry. Please check your record."
         sCutter = Chr(252)
      End If
   End If

   'determine rule to apply
   If InStr("eas,pic,fic,gfi,mys,rom,sci,urb,wes", sCallType) <> 0 Then
      Goto Rule1
   ElseIf sCallType = "dew" Then
      Goto Rule2
   ElseIf sCallType = "bio" Or sCallType = "den" Then
      Goto Rule3
   ElseIf sCallType = "mov" Or sCallType = "tvs" Then
      Goto Rule5
   End If

Rule1:
   'fic, eas, pic, urb, sci, wes, rom, gfi, mys: last name of author of first letter of 110 or 245
   If Left(sMainEntry, 3) = "100" Then
      sMainEntry = Mid(sMainEntry, 6)
      Do While InStr(sMainEntry, ",")
         place = InStr(sMainEntry, ",")
         sMainEntry = RTrim(Left(sMainEntry, place - 1))
      Loop
      sCutter = Left(sMainEntry, 30)
   ElseIf Left(sMainEntry, 3) = "110" Or Left(sMainEntry, 3) = "245" Then
      sCutter = Mid(sMainEntry, 6, 1)
   End If
   sCutter = " " & Chr(223) & "c " & sCutter
   Goto Done

Rule2:
   'dew,den: first letter of main entry
   sCutter = Mid(sMainEntry, 6, 1)
   sCutter = " " & Chr(223) & "c " & sCutter
   Goto Done

Rule3:
   'bio: last name for biographee & first letter of main entry
   'den: last name in subfield $b

   sMainEntry = Mid(sMainEntry, 6, 1)

   If sNameChoice <> Chr(252) Then
      'cut off each section to the right of comma chr until none is left
      Do While InStr(sNameChoice, ",")
         place = InStr(sNameChoice, ",")
         sNameChoice = RTrim(Left(sNameChoice, place - 1))
      Loop
      sNameChoice = Left(sNameChoice, 30)

      'call number type applied based on catalogers selection of element for cuttering 100 vs 600
      If Left(sNameChoice, 3) = "100" Then
         sNameChoice = Mid(sNameChoice, 6)
         sCutter = " " & Chr(223) & "c " & sNameChoice
      ElseIf Left(sNameChoice, 3) = "600" Then
         sNameChoice = Mid(sNameChoice, 6)
         sCutter = " " & Chr(223) & "b " & sNameChoice & " " & Chr(223) & "c " & sMainEntry
      End If
   Else
      'empty element scenarios (have fill character)
      If sCallType = "bio" Then
         sCutter = " " & Chr(223) & "b " & sNameChoice & " " & Chr(223) & "c " & sMainEntry
      Else
         sCutter = " " & Chr(223) & "c " & sNameChoice
      End If
   End If
   Goto Done

Rule4:
   'mov, tvs: Latin - first letter of 245 ; Non-Latin - 11 characters of 245; obsolete efective 08/01/2018 (v.2.5.6)
   If InStr(sElemArr, "245: ") <> 0 Then
      start_point = InStr(sElemArr, "245: ")
      sTemp = Mid(sElemArr, start_point)
      end_point = InStr(sTemp, Chr(9))
      sCutter = Left(sTemp, end_point - 1)
      sCutter = RTrim(" " & Chr(223) & "c " & Mid(sCutter, 6, 1))
   Else
      MsgBox "MISSING INFO: No valid 245 MARC field for a cutter. Please check your record."
      sCutter = " " & Chr(223) & "c " & sCutter
   End If
   Goto Done
Rule5:
   'mov, tvs: first word of 245
   'proposed change to cuttering rules for feature movies and TV shows; effective 08/01/2018 (v.2.5.6)
   If InStr(sElemArr, "245: ") <> 0 Then
      start_point = InStr(sElemArr, "245: ")
      sTemp = Mid(sElemArr, start_point)
      end_point = InStr(sTemp, Chr(9))
      sCutter = Left(sTemp, end_point - 1)
      sCutter = RTrim(Mid(sCutter, 6))
      If InStr(sCutter, " ") <> 0 Then
         end_point = InStr(sCutter, " ")
         sCutter = Left(sCutter, end_point - 1)
      End If
      sCutter = " " & Chr(223) & "c " & sCutter
   Else
      MsgBox "MISSING INFO: No valid 245 MARC field for a cutter. Please check your record."
      sCutter = " " & Chr(223) & "c " & sCutter
   End If
   Goto Done

Done:
End Sub

'########################################################################

Sub ElemArray(sElemArr, n100, n600)
'gather elements of the record that may be used in call number
   Dim CS as Object
   Set CS = CreateObject("Connex.Client")
   Dim sNameTitle$, sIndicator
   Dim sFields() As String
   Dim n, m As Integer

   ReDim sFields(3)
      sFields(0) = "100"
      sFields(1) = "110"
      sFields(2) = "245"
      sFields(3) = "600"

   n100 = 0
   n600 = 0
   For n = 0 To 3
      m = 1
      Do While CS.GetFieldUnicode(sFields(n), m, sNameTitle)
         If InStr(sNameTitle, "&#") = 0 Then
            Call Diacritics(sNameTitle)
            sIndicator = Mid(sNameTitle, 5, 1)
            If n = 0 Then
               n100 = 1
               sNameTitle = Mid(sNameTitle, 6)
            ElseIf n = 1 Then
               'EDIT HERE
               sNameTitle = Mid(sNameTitle, 6)
            ElseIf n = 2 Then
               If sIndicator = "0" Or sIndicator = " " Then
                  sNameTitle = Mid(sNameTitle, 6)
               Else
                  sNameTitle = Mid(sNameTitle, 6 + sIndicator)
               End If
            ElseIf n = 3 Then
               If sIndicator = "0" Then
                  sNameTitle = Mid(sNameTitle, 6)
                  n600 = n600 + 1
               Else
                  Goto NextOccurance
               End If
            End If
            sElemArr = sElemArr & sFields(n) & ": " & sNameTitle & Chr(9)
         End If
NextOccurance:
         m = m + 1
      Loop
   Next n

End Sub

'########################################################################

Sub ElemArrayChoice(sCallType, sElemArr, sNameChoice, n100, n600)

   Dim n, x As Integer
   Dim start_point, end_point As Integer
   Dim sTemp$, sTemp2, sNameArr$
   Dim z As Integer

   If n100 = 0 And n600 = 0 Then
      Goto NoData
   ElseIf n100 = 1 And n600 = 0 Then
      If sCallType = "den" Then
         start_point = InStr(sElemArr, "100: ")
         sTemp = Mid(sElemArr, start_point)
         end_point = InStr(sTemp, Chr(9))
         sNameChoice = Left(sTemp, end_point - 1)
         Goto Done
      ElseIf sCallType = "bio" Then
         Goto NoData
      End If
   ElseIf n100 >= 0 And n600 >= 1 Then
      If sCallType = "bio" Then
         If n600 = 1 Then
            start_point = InStr(sElemArr, "600: ")
            sTemp = Mid(sElemArr, start_point)
            end_point = InStr(sTemp, Chr(9))
            sNameChoice = Left(sTemp, end_point - 1)
            Goto Done
         Else
            sTemp = sElemArr
            Do While InStr(sTemp, "600")
               start_point = InStr(sTemp, "600")
               sTemp = Mid(sTemp, start_point)
               end_point = InStr(sTemp, Chr(9))
               sTemp2 = Left(sTemp, end_point)
               sNameArr = sNameArr & sTemp2
               sTemp = Mid(sTemp, end_point + 1)
            Loop
         End If
      ElseIf sCallType = "den" Then
         If n100 = 0 And n600 = 1 Then
            start_point = InStr(sElemArr, "600: ")
            sTemp = Mid(sElemArr, start_point)
            end_point = InStr(sTemp, Chr(9))
            sNameChoice = Left(sTemp, end_point - 1)
            Goto Done
         Else
            sTemp = sElemArr
            If InStr(sTemp, "100: ") Then
               start_point = InStr(sTemp, "100: ")
               sTemp = Mid(sTemp, start_point)
               end_point = InStr(sTemp, Chr(9))
               sNameArr = Left(sTemp, end_point)
            End If
            sTemp = sElemArr
            Do While InStr(sTemp, "600")
               start_point = InStr(sTemp, "600")
               sTemp = Mid(sTemp, start_point)
               end_point = InStr(sTemp, Chr(9))
               sTemp2 = Left(sTemp, end_point)
               sNameArr = sNameArr & sTemp2
               sTemp = Mid(sTemp, end_point + 1)
            Loop
         End If
      End If
      'dialog box for selection of the name
      Begin Dialog UserDialog 200, 60, "Select Element"
      DropListBox  8, 15, 100, 250, sNameArr, .sNameArr
      OkButton        130, 15,  54, 16
      CancelButton   130, 35,  54, 16
      End Dialog
      Dim dElement as UserDialog
      On Error Resume Next
      'Dialog dElement
      z = Dialog(dElement)
      If z = 0 Then
         'MsgBox "closing"
      End If
      If Err = 102 Then
         sNameArr = ""
         Exit Sub
      End If
      n = dElement.sNameArr + 1
      sTemp = sNameArr
      x = 0
      Do
         place = InStr(sTemp, Chr(9))
         If place <> 0 Then
            lt$ = Left(sTemp, place - 1)
            sTemp = Mid(sTemp, place + 1)
         Else
            lt$ = sTemp
         End If
         x = x + 1
      Loop Until x >= n
      sNameChoice = lt$
   End If
   Goto Done
NoData:
   MsgBox "MISSING INFO: No valid MARC field for this type of call number. Please check your record."
   sNameChoice = Chr(252)
Done:
End Sub

'########################################################################

Function Dewey(a)
'creates string with Dewey number taken from 082 field; 4 digits after period for adult materials, 2 digitst for juvenile; stirps 0s at the end
   Dim CS as Object
   Set CS = CreateObject("Connex.Client")
   Dim s082$, sLastDigit$
   Dim bool082
   Dim x as Integer

   bool082 = CS.GetField("082", 1, s082)
   If bool082 = FALSE Then
      MsgBox "MISSING DATA: No 082 field in the record to create call number."
      Dewey = Chr(223) & "a " & Chr(252)
      Goto Done
   Else
      If a = 0 Then
         Do While InStr(s082, "/")
            place = InStr(s082, "/")
            lt$ = Left(s082, place-1)
            rt$ = Mid(s082, place+1)
            s082 = lt$ + rt$
         Loop
         Do While InStr(s082, "j")
            place = InStr(s082, "j")
            lt$ = Left(s082, place-1)
            rt$ = Mid(s082, place+1)
            s082 = lt$ + rt$
         Loop
         Do While InStr(s082, "C")
            place = InStr(s082, "C")
            lt$ = Left(s082, place-1)
            rt$ = Mid(s082, place+1)
            s082 = lt$ + rt$
         Loop
         place = InStr(s082, Chr(223))
         If place <> 0 Then
            'removes everything on the right of delimiter including space preceding it
            place = InStr(s082, Chr(223))
            s082 = RTrim(Left(s082, place - 1))
         End If
         s082 = RTrim(Mid(s082,6,6))
      Else
         Do While InStr(s082, "/")
            place = InStr(s082, "/")
            lt$ = Left(s082, place-1)
            rt$ = Mid(s082, place+1)
            s082 = lt$ + rt$
         Loop
         Do While InStr(s082, "C")
            place = InStr(s082, "C")
            lt$ = Left(s082, place-1)
            rt$ = Mid(s082, place+1)
            s082 = lt$ + rt$
         Loop
         place = InStr(s082, Chr(223))
         If place <> 0 Then
           'removes everything on the right of delimiter including space preceding it
            place = InStr(s082, Chr(223))
            s082 = RTrim(Left(s082, place - 1))
         End If
         s082 = Mid(s082, 6)
         If Left(s082, 1) <> "8" Then
            s082 = Left(s082, 8)
         End If
   End If
   'removes 0 if it's the last digit and loop
   Do
      x = Len(s082)
      If x <= 3 Then
         Exit Do
      End If
      sLastDigit = Right(s082,1)
      If sLastDigit = "0" Or sLastDigit = "." Then
         s082 = Left(s082, x - 1)
      Else
         Exit Do
      End If
   Loop
   Dewey = Chr(223) & "a " & s082
   End If
Done:
End Function

'########################################################################

Sub Diacritics(sNameTitle)
'removes diacritic marks and other unwanted characters from a string
   Dim CheckChar$, EntryType$

   EntryType = Left(sNameTitle,3)
   If EntryType = "100" Or EntryType = "600" Then
      If InStr(sNameTitle, Chr(223) & "e") <> 0 Then
         place = InStr(sNameTitle, Chr(223) & "e")
         sNameTitle = Left(sNameTitle, place-1)
         sNameTitle = RTrim(sNameTitle)
      End If
      If InStr(sNameTitle, Chr(223) & "d") <> 0 Then
         place = InStr(sNameTitle, Chr(223) & "d")
         sNameTitle = Left(sNameTitle, place-1)
         sNameTitle = RTrim(sNameTitle)
      End If
      If InStr(sNameTitle, Chr(223) & "q") <> 0 Then
         place = InStr(sNameTitle, Chr(223) & "q")
         sNameTitle = Left(sNameTitle, place-1)
         sNameTitle = RTrim(sNameTitle)
      End If
      If InStr(sNameTitle, Chr(223) & "c") <> 0 Then
         place = InStr(sNameTitle, Chr(223) & "c")
         sNameTitle = Left(sNameTitle, place-1)
         sNameTitle = RTrim(sNameTitle)
      End If
      If InStr(sNameTitle, Chr(223) & "x") <> 0 Then
         place = InStr(sNameTitle, Chr(223) & "x")
         sNameTitle = Left(sNameTitle, place-1)
         sNameTitle = RTrim(sNameTitle)
      End If
      Do While InStr(sNameTitle, Chr(223))
         place = InStr(sNameTitle, Chr(223))
         lt$ = Left(sNameTitle, place-2)
         rt$ = Mid(sNameTitle, place+2)
         sNameTitle = lt$ + rt$
      Loop
      sNameTitle = RTrim(Left(sNameTitle, 30))
   Else
   'title/corporate main entry cases
      place = InStr(sNameTitle, Chr(223))
      If place <> 0 Then
         sNameTitle = Left(sNameTitle, place-1)
         sNameTitle = RTrim(sNameTitle)
      End If
      sNameTitle = RTrim(Left(sNameTitle, 30))
   End If

   i = 1
   While i <= Len(sNameTitle)
      CheckChar = Mid(sNameTitle, i, 1)
      Select Case CheckChar
         'characters above letter (example: acute, breve, umlaut, macron, etc.)
         Case Chr(226), Chr(174), Chr(176), Chr(230), Chr(239), Chr(234), Chr(227), Chr(238), Chr(250), Chr(251), Chr(225), Chr(96), Chr(233), Chr(254), Chr(237), Chr(235), Chr(236), Chr(229), Chr(167), Chr(228), Chr(126), Chr(183), Chr(232), Chr(231)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
            i = i - 1
         'characters below letter (example: cedilla, hooks, underscore, etc.)
         Case Chr(240), Chr(248), Chr(247), Chr(241), Chr(244), Chr(242), Chr(243), Chr(245), Chr(246), Chr(249), Chr(246)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
            i = i - 1
         'Ds
         Case Chr(179), Chr(163), Chr(186)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & "d" & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
         'ae lower & uppercase
         Case Chr(181), Chr(165)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & "ae" & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
         'oe lower & uppercase
         Case Chr(182), Chr(166)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & "oe" & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
         'l with slash upper & loawercase
         Case Chr(161), Chr(177)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & "l" & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
         'o with hook or slash
         Case Chr(178), Chr(162), Chr(188), Chr(172)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & "o" & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
         'Turkish i without dot
         Case Chr(183)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & "i" & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
         'u with hook
         Case Chr(189), Chr(173)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & "u" & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
         Case "-"
            sNameTitle = Mid(sNameTitle, 1, i - 1) & " " & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
         'characters to be removed completely
         Case "'", Chr(176), Chr(174), Chr(167)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
            i = i - 1
'        commmented out for update v. 2.5.6. - these characters are allowed in cutters for visual materials effective 08/01/2018
'         Case ".", ":", ";", "/"
'            sNameTitle = Mid(sNameTitle, 1, i - 1) & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
'            i = i - 1

      End Select
      i = i + 1
   Wend
   sNameTitle = UCase(sNameTitle)
'  update v.2.5.8: period should not be allowed as the last element of cutter
'                 - a bug introdued in v. 2.5.6 when period was allowed in DVD call numbers
   If Right(sNameTitle, 1) = "." Then
      sNameTitle = Left(sNameTitle, Len(sNameTitle) - 1)
   End If
End Sub

'########################################################################

Sub Validation(a, f, sAudn, sCallType, sCont, sItemForm, sLang, sRecType, sTmat, sLitF, sBiog, s948)
   Dim place As Integer

   'audience related conflicts
   If a = 0 Then
      If InStr("abcj", sAudn) = 0 Or sAudn = "" Then
         MsgBox "AUDIENCE conflict: Please check record fixed field coding and selected audience."
      ElseIf InStr("mys,sci,rom,urb,wes", sCallType) <> 0 Then
         MsgBox "AUDIENCE conflict: Juvenile material can not have genre call number."
      End If
   ElseIf a = 1 Then
      If InStr("abcj", sAudn) <> 0 And sAudn <> "" Then
         MsgBox "AUDIENCE conflict: Please check record fixed field coding and selected audience."
      ElseIf InStr("eas, pic", sCallType) <> 0 Then
         MsgBox "AUDIENCE conflict: Adult materials can not have Easy or Picture book call number."
      End If
   End If

   'format related conflicts
   If f = 0 Then
      'format empty
      If InStr("gfi,mov,tvs", sCallType) <> 0 Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type."
      ElseIf sItemForm = "d" Then
         MsgBox "FORMAT conflict: Please verify format selection. It appears format should be LG PRINT"
      End If
   ElseIf f = 1 Or f = 4 Then
      'format BlueRay or DVD
      If InStr("eas,pic,fic,gfi,mys,rom,sci,urb,wes", sCallType) <> 0 Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type."
      End If
      If sRecType <> "g" And sTMat <> "v" Then
         MsgBox "FORMAT conflict: Please check record fixed field coding and selected format."
      End If
   ElseIf f = 3 Then
      'format CLASSICS
      If sRecType <> "i" Then
         MsgBox "FORMAT conflict: Please check record fixed field coding and selected format."
      End If
   ElseIf f = 5 Then
      'format 'FINANCE
      If sCallType <> "dew" Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type."
      End If
   ElseIf f = 6 Then
      'GRAPHIC
      If InStr(sCont, "6") = 0 Then
         MsgBox "FORMAT conflict: Please check record fixed field coding and selected format."
      ElseIf InStr("eas,pic,mys,rom,sci,urb,wes,mov,tvs", sCallType) <> 0 Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type."
      End If
   ElseIf f = 7 Then
      'format HOLIDAY
      If InStr("den,bio,gfi,mys,rom,sci,urb,wes,mov,tvs", sCallType) <> 0 Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type."
      ElseIf InStr("abcj", sAudn) = 0 Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type."
      End If
   ElseIf f = 8 Then
      'format JOBS
      If sCallType <> "dew" Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type."
      End If
   ElseIf f = 9 Then
      'format LG PRINT
      If InStr("eas,pic,gfi,mov,tvs", sCallType) <> 0 Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type."
      ElseIf sItemForm <> "d" Then
         MsgBox "FORMAT conflict: Please verify format selection and item form coding."
      End If
   ElseIf f = 10 Then
      'format YR
      If sCallType <> "fic" Or a <> 0 Or InStr("bcj", sAudn) = 0 Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type. YR format is valid only for juvenile fiction call number"
      End If
   ElseIf f = 11 Then
      'format READALONG
      If sCallType <> "fic" And sCallType <> "dew" And sCallType <> "bio" Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type. READALONG format is valid only for juvenile fiction, dewey, and biography call numbers"
      End If
      If a = 1 Then
         MsgBox "AUDIENCE conflict: Please verify audience selection and format. READALONG format is valid only for juvenile materials."
      End If
   End If

   'content validation
   If sRecType = "a" Then
      If InStr("1fj", sLitF) <> 0 Then
         If InStr("eas,pic,fic,gfi,mys,rom,sci,urb,wes", sCallType) = 0 Then
            MsgBox "LITERARY FORM conflict: Fixed field indicates the material is work of fiction. Please verify your selection."
         End If
      Else
         If InSTr("fic,gfi,mys,rom,sci,urb,wes", sCallType) <> 0 Then
            MsgBox "LITERARY FORM conflict: Fixed field indicates the material is non-fiction work. Please verify your selection."
         End If
      End If
      If sBiog = "" Then
         If sCallType = "bio" Then
            MsgBox "WARNING: Fixed field indicates the material doesn't include biographical information. Please verify your call number."
         End If
      End If

      'Dewey + Name call number type
      If sCallType = "den" Then
         place = InStr(s948, Chr(223) & "a")
         If InStr("78", Mid(s948, place + 3, 1)) = 0 Then
            MsgBox "WARNING: Call number Dewey + Name should be used only in 7xx and 8xx Dewey ranges. Please verify your selection."
         End If
      End If

      'short stories collections warning flag
   End If


End Sub

'########################################################################

Sub InsertCallNum(s948, f, sInitials)
   Dim CS as Object
   Set CS = CreateObject("Connex.Client")
   Dim s901$

   CS.SetField 1, s948
   CS.SetField 1, "945  .b" & Chr(252)
   CS.SetField 1, "946  m"

   'sFormat codes in variable f
   If f = 1 Then
      'bluray
      CS.SetField 1, "949  *b2=b;recs=oclcgw;"
   ElseIf f = 2 Then
      'audiobook
      CS.SetField 1, "949  *b2=u;recs=oclcgw;"
   ElseIf f = 4 Then
      'dvd
      CS.SetField 1, "949  *b2=v;recs=oclcgw;"
   ElseIf f = 9 Then
      'large-print
      CS.SetField 1, "949  *b2=l;recs=oclcgw;"
   Else
      'print
      CS.SetField 1, "949  *b2=a;recs=oclcgw;"
   End If

   s901 = "901  " & sInitials & " " & Chr(223) & "b CATBL"

   CS.SetField 1, s901
 CS.EndRecord
End Sub

