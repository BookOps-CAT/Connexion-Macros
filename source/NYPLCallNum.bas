'MacroName:NYPL CallNum
'MacroDescription: NYPL macro for creating a complete call number in field 948 based on catalogers selected pattern and information coded in the record
'                  Macro handles call number patterns for English and World Languages, fiction, non-fiction, biography and biography with Dewey
'                  incorporates functions of Format macro - populates subfield $f 

'v3.3.0 (09-18-2024):
'  * adds fill character (Chr(252)) in lieu of empty or digits in the cutter for non-visual materials (DVDs, BluRay)
'  * fixes dotles i handling by replacing it with uppercase i
'  * fixes READALONG Sierra Material Type code (5)
'  * adds handling of incorrectly coded characters in 100, 110, 245, and 6xx fields
'v3.2.1 (02-06-2023):
'  * fixes invalid cutters consisting of numbers in print materials and visual non-fic
'v3.2.0 (10-25-2022):
'  * adds bib location command in 949 field with a default BL value ("bn=zzzzz;")
'v3.1.1 (09-01-2022):
'  * fixes a bug in cuttering by title when apostrophe is part of skip characters in second indicator of 245 field (example: 245 02 L'Amour)
'v3.1.0 (07-29-2022):
'  * adds reformating routine before any parsing
'  * fixes incorrect flag for Classics
'  * adds a flag if material includes multiple languages, but the call number does not include a laguage prefix
'v3.0.1 (05-13-2022):
'  * fixes bug that includes subfield "v","t", & "z" in call numbers for biography and dewey+subject
'v3.0.0 (05-06-2022):
'  * removes a routine that deletes unsupported subject vocabularies from 6xx fields (moved to CAT!UpdateExport macro)
'v2.9.0 update datails (04-15-2022):
'  * removes aat thesaurus from terms permitted on BL
'v2.8.0 update details (03-07-2022):
'  * permits homosaurus terms (homoit)
'  * changes call number patterns for Dewey/Dewey+Name > Dewey/Dewey+Author/Dewey+Subject
'  * removes time table digits from 8xx for single author works
'  * adds Modern Persian (Farsi) literature time table digits removal
'v2.7.0 update details (09-30-2021):
' * GN FIC call number type eliminated and incorpporated into FIC
' * BRAILLE format added

'v2.6.1 update details (04-05-2021):
'  * removal of catalog headings unapproved for use in NYPL catalog (BISACS, SEARS, etc.)

'v2.6.0 update details (01-29-2021):
'  * first word in cutter (taken from 110) when the main entry is 110 tag in fiction and related call numbers
'v2.5.10 update details (07-20-2020):
'  * catalogers initials saved to OCLC profile folder instead of general Connexion directory (more likely to have write access)

'v2.5.9 update details:
'  * complete elimination of 11 characters rule in cutter for non-Latin materials - only last name used or first letter
'    (cutter in subfield $c)

'v2.5.8 update details:
'  * Bug fix: removes period as the last character in cutter (the bug introduced in 2.5.6)

'v2.5.7 update details:
'  * Readalong call number prefix added
'  * Validation flats added for Readalong call numbers

'v2.5.6 update details:
'  * DVD & Bluray call number cutter includes full first word
'  * bug fix: underscore nonspacing character (Chr(246)) handling fixed

'v2.5.5 update details:
'  * bug fix: proper behavior if no bibliographic record displayed
'  * added flag for short stories collections

'v2.5.4 update details:
'  * bug fix: removes false fiction flag for Graphic Novels
'  * improvement: diacritics function simplified and made more comprehensive
'  * bug fix: corrected broken format error flags
'  * rules change: authors name for literary collection changed from $b to $c
'  * improvement: added validation error flag for Dewey + Name call numbers that are not in 7xx or 8xx range


Option Explicit

Declare Function Dewey(a, sCallType)
Declare Function HasIllegalCutter(sCutter)
Declare Function Normalized(sNameTitle, sTag)
Declare Sub LocalDewey(s082, sCallType)
Declare Sub ElemArray(sElemArr, sCallType)
Declare Sub ElemArrayChoice(sElemArr, sNameChoice, n600)
Declare Sub Diacritics(sNameTitle)
Declare Sub Rules(sElemArr, sCallType, sLang, sCutter, sNameChoice)
Declare Sub InsertCallNum(s948, f, sInitials)
Declare Sub Validation(a, f, sAudn, sCallType, sCont, sCutter, sItemForm, sLang, sRecType, sTmat, sLitF, sBiog, s041, s948, boolLangPrefix)

'temporary variables
Dim place, i as Integer
Dim lt$, rt$, sTemp$

Sub Main

   Dim CS as Object
   On Error Resume Next
   Set CS  = GetObject(,"Connex.Client")
   On Error GoTo 0
 
   If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 2 or CS.ItemType = 17 Then

      ' make sure all fields are reformated and subfields include required spacing
      CS.Reformat
   
      Dim s041$, s300$, s538$, s948$, sAudn$, sAudnLangPrefix$, sBiog$, sCallType$, sCont$, sCutter$, sElemArr$, _
         sFormatPrefix$, sItemForm$, sLang$, sLitF$, sNameChoice$, sRecType$, sTMat$
      Dim bool538
      Dim boolLangPrefix
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
      CS.GetFieldUnicode "041", 1, s041
      If InStr(s300, Chr(223) & "e") <> 0 Then
         s300 = Mid(s300, InStr(s300, Chr(223) & "e"))
      End If
      s300 = UCase(s300)
      
      ReDim sFormat(12)
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
         sFormat(12) = "BRAILLE"
      ReDim sAudience(1)
         sAudience(0) = "JUVENILE"
         sAudience(1) = "ADULT"
      ReDim sOutput(1)
         sOutput(0) = "add to record"
         sOutput(1) = "add to clipboard"
         
      'read default data (initials) from text file stored in macro folder
      sFileName = Mid(Environ(2), 9) + "\OCLC\Connex\Profiles\cat_data.txt"

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
         

      'Dialog box presenting cataloger choices for types of call numbers
      Begin Dialog MainWindow 220, 220, "NYPL Call Number Macro"
         
         'top-left outline
         GroupBox 18, 50, 87, 41, ""
         'bottom-left outline
         GroupBox 18, 90, 87, 100, ""
         'top-right outline
         GroupBox 120, 50, 82, 100, ""
         'bottom-right outline
         GroupBox 120, 149, 82, 41, ""
         OptionGroup .Type
         OptionButton  24,  55, 70, 14, "&EASY BOOKS"
         OptionButton  24,  75, 70, 14, "&PICTURE BOOKS"
         OptionButton  24,  95, 70, 14, "&BIOGRAPHY"
         OptionButton  24,  115, 70, 14, "&DEWEY"
         OptionButton  24,  135, 70, 14, "DEWEY+&NAME"
         OptionButton  24,  155, 75, 14, "DEWE&Y+SUBJECT"
         OptionButton  24,  175, 70, 14, "&FICTION"
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
      
      'selects most likely audience
      If InStr("abcj", sAudn) <> 0 And sAudn <> "" Then
         dCallNum.sAudience = 0
      Else
         dCallNum.sAudience = 1
      End If
      
      'select most likely format
      If sRecType = "a" Then
         If InStr(sCont, "6") <> 0 Then
            dCallNum.sFormat = 6
         ElseIf sItemForm = "d" Then
            dCallNum.sFormat = 9
         ElseIf sItemForm = "f" Then
            dCallNum.sFormat = 12
         End If
      ElseIf sRecType = "i" Then
         If InStr(s300, "AUDIO-ENABLED BOOK") <> 0 Or InStr(s300, "AUDIO ENABLED BOOK") <> 0 Then
            dCallNum.sFormat = 11
         Else
            dCallNum.sFormat = 2
         End If
      ElseIf sRecType = "j" Then
         MsgBox "Please consider using NYPLMusicCD macro instead. The record appears to be a music CD"
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
         MsgBox "INCOMPLETE: Please correct language coding in the fixed field and the call number."
      End If
      If a = 0 Then
         If sLang <> "ENG" Then
            sAudnLangPrefix = Chr(223) & "p J " & sLang & " "
            boolLangPrefix = True
         Else
            sAudnLangPrefix = Chr(223) & "p J" & " "
            boolLangPrefix = False
         End If
      Else
         If sLang <> "ENG" Then
            sAudnLangPrefix = Chr(223) & "p " & sLang & " "
            boolLangPrefix = True
         Else
            boolLangPrefix = False
         End If
      End If
      
'  define subfield $f of 948 field
      f = dCallNum.sFormat
      If f <> 0 Then
         sFormatPrefix = Chr(223) & "f " & sFormat(dCallNum.sFormat) & " "
      End If
      
'  define the remainder of 948 field
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
            Call ElemArrayChoice(sElemArr, sNameChoice, n600)
            s948 = s948 & Chr(223) & "a B"
         Case 3
            sCallType = "dew"
            s948 = s948 & Dewey(a, sCallType)
         Case 4
            sCallType = "den"
            s948 = s948 & Dewey(a, sCallType)
        Case 5
            sCallType = "des"
            'add selection of name from available 600s here then pass variable to rules
            Call ElemArrayChoice(sElemArr, sNameChoice, n600)
            s948 = s948 & Dewey(a, sCallType)
         Case 6
            sCallType = "fic"
            s948 = s948 & Chr(223) & "a FIC"
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
      
      'complete creation of sElemArr
      Call ElemArray(sElemArr, sCallType)
      
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
   Call Validation(a, f, sAudn, sCallType, sCont, sCutter, sItemForm, sLang, sRecType, sTmat, sLitF, sBiog, s041, s948, boolLangPrefix)
ReallyDone:
End Sub

'########################################################################

Sub Rules(sElemArr, sCallType, sLang, sCutter, sNameChoice)
   'rules for creating the remainder of the call number 
   Dim sLastChr$, sMainEntry$
   Dim start_point, end_point As Integer
   
   'find main entry 
   If InStr(sElemArr, "100: " & Chr(252)) = 0 Then
      start_point = InStr(sElemArr, "100: ")
      sTemp = Mid(sElemArr, start_point)
      end_point = InStr(sTemp, Chr(9))
      sMainEntry = Left(sTemp, end_point - 1)
   ElseIf InStr(sElemArr, "110: " & Chr(252)) = 0 Then
      start_point = InStr(sElemArr, "110:")
      sTemp = Mid(sElemArr, start_point)
      end_point = InStr(sTemp, Chr(9))
      sMainEntry = Left(sTemp, end_point - 1)
   Else
      If InStr(sElemArr, "245: " & Chr(252)) = 0 Then
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
   If InStr("eas,den,pic,fic,mys,rom,sci,urb,wes", sCallType) <> 0 Then
      Goto Rule1
   ElseIf sCallType = "dew" Then
      Goto Rule2
   ElseIf sCallType = "bio" Or sCallType = "des" Then
      Goto Rule3
   ElseIf sCallType = "mov" Or sCallType = "tvs" Then
      Goto Rule5
   End If
  
Rule1:
   'eas, den, fic, mys, pic, rom, urb, sci, wes: last name of author, first word of 110 or first letter of 245
   If Left(sMainEntry, 3) = "100" Then
      sMainEntry = Mid(sMainEntry, 6)
      Do While InStr(sMainEntry, ",")
         place = InStr(sMainEntry, ",")
         sMainEntry = RTrim(Left(sMainEntry, place - 1))
      Loop
      sCutter = Left(sMainEntry, 30)
   ElseIf Left(sMainEntry, 3) <> "100" AND sCallType = "den" Then
      sCutter = Chr(252)
      MsgBox "Invalid call number choice. Dewey+Author must have 100 field."
   ElseIf Left(sMainEntry, 3) = "110" Then
      sMainEntry = Mid(sMainEntry, 6)
      Do While InStr(sMainEntry, " ")
         place = InStr(sMainEntry, " ")
         sMainEntry = RTrim(Left(sMainEntry, place - 1))
      Loop
      sCutter = Left(sMainEntry, 30)
   
   ElseIf Left(sMainEntry, 3) = "245" Then
      sCutter = Mid(sMainEntry, 6, 1)
      If HasIllegalCutter(sCutter) = TRUE Then
         sCutter = Chr(252)
      End If
   End If

   sCutter = " " & Chr(223) & "c " & sCutter
   Goto Done

Rule2:
   'dew,den: first letter of main entry
   If HasIllegalCutter(sMainEntry) = TRUE Then
      sCutter = Chr(252)
   Else 
      sCutter = Mid(sMainEntry, 6, 1)
   End If
   sCutter = " " & Chr(223) & "c " & sCutter
   Goto Done

Rule3:
   'bio: last name for biographee & first letter of main entry
   'des: dewey + last name in subfield b + first letter of main entry
   
   sMainEntry = Mid(sMainEntry, 6, 1)
  
   If sNameChoice <> Chr(252) Then
      'cut off each section to the right of comma chr until none are left
      Do While InStr(sNameChoice, ",")
         place = InStr(sNameChoice, ",")
         sNameChoice = RTrim(Left(sNameChoice, place - 1))
      Loop
      sNameChoice = Left(sNameChoice, 30)
      sNameChoice = Mid(sNameChoice, 6)
   End If
  
   sCutter = " " & Chr(223) & "b " & sNameChoice & " " & Chr(223) & "c " & sMainEntry
   Goto Done

Rule4:
   'mov, tvs: Latin - first letter of 245 ; Non-Latin - 11 characters of 245; obsolete effective 08/01/2018 (v.2.5.6)
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
      MsgBox "MISSING INFO: No valid 245 MARC field for cutter. Please check your record."
      sCutter = " " & Chr(223) & "c " & sCutter
   End If
   Goto Done

Done:
End Sub

'########################################################

Function Normalized(sNameTitle, sTag)
   'Checks if Unicode is present and shows warning or adds fill character for empty strings.
   'Activates Diacritics routine.
   
   Dim Indicator

   If InStr(sNameTitle, "&#") <> 0 Then
      MsgBox "Field " & Left(sNameTitle, 3) & " includes a possibly incorrectly coded character (see &#x code). Please replace with ALA diacritics: '" & Mid(sNameTitle, 6) &  "'" 
      Normalized = sTag & ": " & Chr(252)
   ElseIf sNameTitle = "" Then
      Normalized = sTag & ": " & Chr(252)
   Else
      Indicator = Mid(sNameTitle,5,1)
      If Indicator = "0" Or Indicator = " " Then
         lt$ = sTag & ": "
         rt$ = Mid(sNameTitle, 6, 30)
         sNameTitle = lt$ + rt$
      Else
         lt$ = sTag & ": "
         rt$ = Mid(sNameTitle, 6 + Indicator, 10) 'is this 10 a correct limitation? is title cutter being shortened for DVD by this?
         sNameTitle = lt$ + rt$
      End If
      
      If InStr(sNameTitle, Chr(223) & "e") Then
         place = InStr(sNameTitle, Chr(223) & "e")
         sNameTitle = Left(sNameTitle, place-1)
         sNameTitle = RTrim(sNameTitle)
      End If
      If InStr(sNameTitle, Chr(223) & "d") Then
         place = InStr(sNameTitle, Chr(223) & "d")
         sNameTitle = Left(sNameTitle, place-1)
         sNameTitle = RTrim(sNameTitle)
      End If
      If InStr(sNameTitle, Chr(223) & "c") Then
         place = InStr(sNameTitle, Chr(223) & "c")
         sNameTitle = Left(sNameTitle, place-1)
         sNameTitle = RTrim(sNameTitle)
      End If
      If InStr(sNameTitle, Chr(223) & "v") Then
         place = InStr(sNameTitle, Chr(223) & "v")
         sNameTitle = Left(sNameTitle, place-1)
         sNameTitle = RTrim(sNameTitle)
      End If
      Do While InStr(sNameTitle, ",")
         place = InStr(sNameTitle, ",")
         sNameTitle = RTrim(Left(sNameTitle, place - 1))
      Loop
      Do While InStr(sNameTitle, Chr(223))
         place = InStr(sNameTitle, Chr(223))
         lt$ = Left(sNameTitle, place-2)
         rt$ = Mid(sNameTitle, place+2)
         sNameTitle = lt$ + rt$
      Loop
      Do While InStr(sNameTitle, ".")
         place = InStr(sNameTitle, ".")
         lt$ = Left(sNameTitle, place-1)
         rt$ = Mid(sNameTitle, place+1)
         sNameTitle = lt$ + rt$
      Loop
      
      Call Diacritics(sNameTitle)
      sNameTitle = UCase(sNameTitle)
      Normalized = sNameTitle

   End If

End Function

'################################################################

Sub ElemArray(sElemArr, sCallType)
'gather elements of the record that may be used in call number   
   Dim CS as Object
   On Error Resume Next
   Set CS  = GetObject(,"Connex.Client")
   On Error GoTo 0

   Dim sNameTitle$
   Dim i As Integer
   Dim bool

   'linked fields with Non-Latin script are displayed first,
   'they should be ignored for cuttering purposes;
   'occasionally primary Latin script (MARC-8) will have a unicode-encoded
   'character if it is not supported by MARC-8 (for example African inverted e - see #1163881264)
   'and in such situations CutterArray should display a warning but allow to use an element for cuttering
      
   bool = CS.GetFieldUnicode("100", 2, sNameTitle)
      If bool = False Then
      CS.GetFieldUnicode "100", 1, sNameTitle
   End If
   sNameTitle = Normalized(sNameTitle, "100")
   sElemArr = sElemArr & sNameTitle & Chr(9)
   
   bool = CS.GetFieldUnicode("110", 2, sNameTitle)
   If bool = False Then
      CS.GetFieldUnicode "110", 1, sNameTitle
   End If
   sNameTitle = Normalized(sNameTitle, "110")
   sElemArr = sElemArr & sNameTitle & Chr(9)
   
   bool = CS.GetFieldUnicode("245", 2, sNameTitle)
   If bool = False Then
         CS.GetFieldUnicode "245", 1, sNameTitle
   End If
   sNameTitle = Normalized(sNameTitle, "245")
   sElemArr = sElemArr & sNameTitle & Chr(9)

   
   If sCallType = "des" Or sCallType = "bio" Then
      i = 1
      Do While CS.GetFieldUnicode("600", i, sNameTitle)
         If InStr(sNameTitle, "&#") = 0 And Mid(sNameTitle, 5, 1) = "0" Then
            sNameTitle = Normalized(sNameTitle, "600")
            sElemArr = sElemArr & sNameTitle & Chr(9)
         End If
         i = i + 1
      Loop
  End If
   
End Sub  
   


'########################################################################

Sub ElemArrayChoice(sElemArr, sNameChoice, n600)

   Dim n, x As Integer
   Dim start_point, end_point As Integer
   Dim sTemp$, sTemp2, sNameArr$
   Dim z As Integer
   
   If n600 = 0 Then
      Goto NoData
   ElseIf n600 = 1 Then
      'only one 600 field, simply create element
      start_point = InStr(sElemArr, "600: ")
      sTemp = Mid(sElemArr, start_point)
      end_point = InStr(sTemp, Chr(9))
      sNameChoice = Left(sTemp, end_point - 1)
      Goto Done     
   ElseIf n600 >= 1 Then
      'multiple 600 fields, allow cataloger to select
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
   Goto Done
 
NoData:
   MsgBox "MISSING INFO: No valid MARC field for this type of call number. Please check your record."
   sNameChoice = Chr(252)
Done:
End Sub

'########################################################################

Function Dewey(a, sCallType)
'creates string with Dewey number taken from 082 field; 4 digits after period for adult materials, 2 digits for juvenile; strips 0s at the end
   Dim CS as Object
   On Error Resume Next
   Set CS  = GetObject(,"Connex.Client")
   On Error GoTo 0

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
 
         'shorten Dewey to 4 digits after the period if not 8xx
         If Left(s082, 1) <> "8" Then
            s082 = Left(s082, 8)
         End If
   End If
   
   Call LocalDewey(s082, sCallType)
 
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

'########################################################

Function HasIllegalCutter(sCutter)

   If InStr("0123456789", sCutter) <> 0 Then
      MsgBox "INCORRECT call number: a cutter can not consist of a digit. Please use first letter of spelled out number in the language of the cataloged material."
      HasIllegalCutter = TRUE
   Else
      HasIllegalCutter = FALSE
   End If

End Function


'########################################################################

Sub LocalDewey(s082, sCallType)
   Dim s1stDig$, s2ndDig$, s3rdDig$, s5thDig$, s6thDig$, s7thDig$, s8thDig$, s9thDig$
   Dim s1stThreeDig$, s1stFiveDig$, s1stSixDig$, s3rdTo7thDig$
   Dim sLitTimeTableMsg$
   
   If sCallType = "dew" Then
      Goto Done
   End If
   
   s1stDig = Left(s082, 1)
   s2ndDig = Mid(s082, 2, 1)
   s3rdDig = Mid(s082, 3, 1)
   s5thDig = Mid(s082, 5, 1)
   s6thDig = Mid(s082, 6, 1)
   s7thDig = Mid(s082, 7, 1)
   s8thDig = Mid(s082, 8, 1)
   s9thDig = Mid(s082, 9, 1)
   
   s1stThreeDig = Left(s082, 3)
   s1stFiveDig = Left(s082, 5)
   s1stSixDig = Left(s082, 6)
   s3rdTo7thDig = Mid(082, 3, 5)
   
   sLitTimeTableMsg = "INFO: NYPL does not use time tables in literature for works by our about single author. Removing time table digits..."
   

   'removes time periods from American English, English, Spanish, German, French & Italian literature call number
   If s1stDig = "8" And InStr("123456", s2ndDig) <> 0 Then
      If InStr("12345678", s3rdDig) <> 0 And InStr("123456789", s5thDig) <> 0 Then
         If InStr(s082,"08") <> 0 Or InStr(s082, "09") <> 0 Then
            If InStr(s082,"08") <> 0 Then
               place = InStr(s082, "08")
            End If
            If InStr(s082, "09") <> 0 Then
               place = InStr(s082, "09")
            End If
            MsgBox sLitTimeTableMsg
            lt$ = Left(s082, 4)
            rt$ = Mid(s082, place)
            s082 = lt$ + rt$
         Else
            MsgBox sLitTimeTableMsg
            s082 = Left(s082, 3)
         End If
      End If
      If s3rdTo7thDig = "0.800" Or s3rdTo7thDig = "0.900" Then
         MsgBox sLitTimeTableMsg
         s082 = Left(s082,5)
      End If
   End If
   'removes time periods from Slavic literatures call numbers
   If s1stFiveDig = "891.8" And InStr("123456789", s6thDig) <> 0 Then
      If InStr("12345678", s7thDig) <> 0 And InStr("123456789", s8thDig) <> 0 Then
         If InStr(s082,"08") <> 0 Or InStr(s082, "09") <> 0 Then
            If InStr(s082,"08") <> 0 Then
               place = InStr(s082, "08")
            End If
            If InStr(s082, "09") <> 0 Then
               place = InStr(s082, "09")
            End If
            MsgBox sLitTimeTableMsg
            lt$ = Left(s082, 7)
            rt$ = Mid(s082, place)
            s082 = lt$ + rt$
         Else
            MsgBox sLitTimeTableMsg
            s082 = Left(s082, 7)
         End If
      End If
   End If
   'removes time periods from Finnic literatures call numbers
   If s1stSixDig = "894.54" And InStr("15", s7thDig) <> 0 Then
      If InStr("12345678", s8thDig) <> 0 And InStr("1234", s9thDig) <> 0 Then
         If InStr(s082,"08") <> 0 Or InStr(s082, "09") <> 0 Then
            If InStr(s082,"08") <> 0 Then
               place = InStr(s082, "08")
            End If
            If InStr(s082, "09") <> 0 Then
               place = InStr(s082, "09")
            End If
            MsgBox sLitTimeTableMsg
            lt$ = Left(s082, 8)
            rt$ = Mid(s082, place)
            s082 = lt$ + rt$
         Else
            MsgBox sLitTimeTableMsg
            s082 = Left(s082, 8)
         End If
      End If
   End If
   'removes time periods from Russian literature call numbers
   If s1stFiveDig = "891.7" And InStr("12345678", s6thDig) <> 0 Then
      If InStr("12345", s7thDig) <> 0 Then
         If InStr(s082,"08") <> 0 Or InStr(s082, "09") <> 0 Then
            If InStr(s082,"08") <> 0 Then
               place = InStr(s082, "08")
            End If
            If InStr(s082, "09") <> 0 Then
               place = InStr(s082, "09")
            End If
            MsgBox sLitTimeTableMsg
            lt$ = Left(s082, 6)
            rt$ = Mid(s082, place)
            s082 = lt$ + rt$
         Else
            MsgBox sLitTimeTableMsg
            s082 = Left(s082, 6)
         End If
      End If
   End If
  'removes time periods from Ukrainian literatures call numbers
   If s1stSixDig = "891.79" And InStr("12345678", s7thDig) <> 0 Then
      If InStr("12345", s8thDig) <> 0 Then
         If InStr(s082,"08") <> 0 Or InStr(s082, "09") <> 0 Then
            If InStr(s082,"08") <> 0 Then
               place = InStr(s082, "08")
            End If
            If InStr(s082, "09") <> 0 Then
               place = InStr(s082, "09")
            End If
            MsgBox sLitTimeTableMsg
            lt$ = Left(s082, 7)
            rt$ = Mid(s082, place)
            s082 = lt$ + rt$
         Else
            MsgBox sLitTimeTableMsg
            s082 = Left(s082, 7)
         End If
      End If
   End If
   'removes time periods from Japanese literature call numbers
   If s1stFiveDig = "895.6" And InStr("12345678", s6thDig) <> 0 Then
      If InStr("123456", s7thDig) <> 0 Then
         If InStr(s082,"08") <> 0 Or InStr(s082, "09") <> 0 Then
            If InStr(s082,"08") <> 0 Then
               place = InStr(s082, "08")
            End If
            If InStr(s082, "09") <> 0 Then
               place = InStr(s082, "09")
            End If
            MsgBox sLitTimeTableMsg
            lt$ = Left(s082, 6)
            rt$ = Mid(s082, place)
            s082 = lt$ + rt$
         Else
            MsgBox sLitTimeTableMsg
            s082 = Left(s082, 6)
         End If
      End If
   End If
   'removes time periods from Chinese literature call numbers
   If s1stFiveDig = "895.1" And InStr("12345678", s6thDig) <> 0 Then
      If InStr("123456", s7thDig) <> 0 Then
         If InStr(s082,"08") <> 0 Or InStr(s082, "09") <> 0 Then
            If InStr(s082,"08") <> 0 Then
               place = InStr(s082, "08")
            End If
            If InStr(s082, "09") <> 0 Then
               place = InStr(s082, "09")
            End If
            MsgBox sLitTimeTableMsg
            lt$ = Left(s082, 6)
            rt$ = Mid(s082, place)
            s082 = lt$ + "0" + rt$
         Else
            MsgBox sLitTimeTableMsg
            s082 = Left(s082, 6)
         End If
      End If
   End If
   'removes time periods from other Germanic literatures call numbers (includes Yiddish, Swedish, Old Norse, Icelandic)
   If s1stThreeDig = "839" And InStr("124567", s5thDig) <> 0 Then
      If InStr("12345678", s6thDig) <> 0 Then
         If InStr(s082,"08") <> 0 Or InStr(s082, "09") <> 0 Then
            If InStr(s082,"08") <> 0 Then
               place = InStr(s082, "08")
            End If
            If InStr(s082, "09") <> 0 Then
               place = InStr(s082, "09")
            End If
            MsgBox sLitTimeTableMsg
            lt$ = Left(s082, 6)
            rt$ = Mid(s082, place)
            s082 = lt$ + rt$
         Else
            MsgBox sLitTimeTableMsg
            s082 = Left(s082, 6)
         End If
      End If
   End If
   'removes time periods from Portuguese literature
   If s1stThreeDig = "869" And InStr("12345678", s5thDig) <> 0 Then
         If InStr(s082,"08") <> 0 Or InStr(s082, "09") <> 0 Then
            If InStr(s082,"08") <> 0 Then
               place = InStr(s082, "08")
            End If
            If InStr(s082, "09") <> 0 Then
               place = InStr(s082, "09")
            End If
            MsgBox sLitTimeTableMsg
            lt$ = Left(s082, 5)
            rt$ = Mid(s082, place)
            s082 = lt$ + rt$
         Else
            MsgBox sLitTimeTableMsg
            s082 = Left(s082, 5)
         End If
   End If
   'removes time periods from Danish, Norwegian literatures call numbers 
   If s1stFiveDig = "839.8" And InStr("12", s6thDig) <> 0 Then
      If InStr("12345678", s7thDig) <> 0 Then
         If InStr(s082,"08") <> 0 Or InStr(s082, "09") <> 0 Then
            If InStr(s082,"08") <> 0 Then
               place = InStr(s082, "08")
            End If
            If InStr(s082, "09") <> 0 Then
               place = InStr(s082, "09")
            End If
            MsgBox sLitTimeTableMsg
            lt$ = Left(s082, 7)
            rt$ = Mid(s082, place)
            s082 = lt$ + rt$
         Else
            MsgBox sLitTimeTableMsg
            s082 = Left(s082, 7)
         End If
      End If
   End If
   'removes time periods from Arabic literatures call numbers 
   If s1stFiveDig = "892.7" And InStr("12345678", s6thDig) <> 0 And InStr("1234567", s7thDig) <> 0 Then
         If InStr(s082,"08") <> 0 Or InStr(s082, "09") <> 0 Then
            If InStr(s082,"08") <> 0 Then
               place = InStr(s082, "08")
            End If
            If InStr(s082, "09") <> 0 Then
               place = InStr(s082, "09")
            End If
            MsgBox sLitTimeTableMsg
            lt$ = Left(s082, 6)
            rt$ = Mid(s082, place)
            s082 = lt$ + rt$
         Else
            MsgBox sLitTimeTableMsg
            s082 = Left(s082, 6)
         End If
   End If
   'removes time periods from Classic Greek and Latin
   If s1stDig = "8" And InStr("78", s2ndDig) <> 0 Then
      If InStr("12345678", s3rdDig) <> 0 And s5thDig = "0" And InStr("1234", s6thDig) <> 0  Then
         If InStr(s082,"08") <> 0 Or InStr(s082, "09") <> 0 Then
            If InStr(s082,"08") <> 0 Then
               place = InStr(s082, "08")
            End If
            If InStr(s082, "09") <> 0 Then
               place = InStr(s082, "09")
            End If
            MsgBox sLitTimeTableMsg
            lt$ = Left(s082, 4)
            rt$ = Mid(s082, place)
            s082 = lt$ + rt$
         Else
            MsgBox sLitTimeTableMsg
            s082 = Left(s082, 3)
         End If
      End If
   End If
   ' removes time period digits from Modern Persian (Farsi) literature
   If s1stSixDig = "891.55" And InStr("12345678", s7thDig) <> 0 Then
      If InStr(s082,"08") <> 0 Or InStr(s082, "09") <> 0 Then
         If InStr(s082,"08") <> 0 Then
            place = InStr(s082, "08")
         End If
         If InStr(s082, "09") <> 0 Then
            place = InStr(s082, "09")
         End If
         MsgBox sLitTimeTableMsg
         lt$ = Left(s082, 7)
         rt$ = Mid(s082, place)
         s082 = lt$ + rt$
      Else
         MsgBox sLitTimeTableMsg
         s082 = Left(s082, 7)
      End If
   End If

Done:
End Sub

'########################################################################

Sub Diacritics(sNameTitle)
'removes diacritic marks and other unwanted characters from a string
   Dim CheckChar$
   
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
         'l with slash upper & lowercase
         Case Chr(161), Chr(177)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & "l" & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
         'o with hook or slash
         Case Chr(178), Chr(162), Chr(188), Chr(172)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & "o" & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
         'Turkish i without dot
         Case Chr(184)
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
'        commented out for update v. 2.5.6. - these characters are allowed in cutters for visual materials effective 08/01/2018
'        Case ".", ":", ";", "/"
'           sNameTitle = Mid(sNameTitle, 1, i - 1) & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
'           i = i - 1
                 
      End Select
      i = i + 1   
   Wend

End Sub

'########################################################################

Sub Validation(a, f, sAudn, sCallType, sCont, sCutter, sItemForm, sLang, sRecType, sTmat, sLitF, sBiog, s041, s948, boolLangPrefix)
   Dim place As Integer
   Dim sTemp$
   Dim sLangs$
   Dim sValue$
   
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
      If InStr("mov,tvs", sCallType) <> 0 Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type."
      ElseIf sItemForm = "d" Then
         MsgBox "FORMAT conflict: Please verify format selection. It appears that the format should be LG PRINT"
      End If
   ElseIf f = 1 Or f = 4 Then
      'format BlueRay or DVD
      If InStr("eas,pic,fic,mys,rom,sci,urb,wes", sCallType) <> 0 Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type."
      End If
      If sRecType <> "g" And sTMat <> "v" Then
         MsgBox "FORMAT conflict: Please check record fixed field coding and selected format."
      End If
   ElseIf f = 3 Then
      'format CLASSICS
      If InStr("gjopr", sRecType) <> 0 Then
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
      If InStr("den,bio,mys,rom,sci,urb,wes,mov,tvs", sCallType) <> 0 Then
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
      If InStr("eas,pic,mov,tvs", sCallType) <> 0 Then
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
      If sCallType <> "fic" Or sCallType <> "dew" Or sCallType <> "bio" Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type. READALONG format is valid only for juvenile fiction, dewey, and biography call numbers"
      End If
      If a = 1 Then
         MsgBox "AUDIENCE conflict: Please verify audience selection and format. READALONG format is valid only for juvenile materials."
      End If
   ElseIf f = 12 Then
      'format Braille
      If sCallType <> "fic" And sCallType <> "dew" And sCallType <> "bio" And sCallType <> "pic" And sCallType <> "den" Then
         MsgBox "FORMAT conflict: Please verify format selection and call number type. BRAILLE format is valid only for picture books, fiction, dewey, dewey+subject, and biography call numbers."
      End If
   End If
   
   'content validation
   If sRecType = "a" Then
      If InStr("1fj", sLitF) <> 0 Then
         If InStr("eas,pic,fic,mys,rom,sci,urb,wes", sCallType) = 0 Then
            MsgBox "LITERARY FORM conflict: Fixed field indicates the material is a work of fiction. Please verify your selection."
         End If
      Else
         If InSTr("fic,mys,rom,sci,urb,wes", sCallType) <> 0 Then
            MsgBox "LITERARY FORM conflict: Fixed field indicates the material is a non-fiction work. Please verify your selection."
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
   

   'language prefix validation
   'check if material is coded be in more then one language
   
   If s041 <> "" Then
   
      sLangs = UCase(Mid(s041, 6, 3))
      sTemp = Mid(s041, 6)
      
      Do While InStr(sTemp, Chr(223) & "a") <> 0:
        
         place = InStr(sTemp, Chr(223) & "a")
         
         ' assuming subfields are formated correctly and there is space between subfield and value
         sValue = UCase(Trim(Mid(sTemp, place + 3, 3)))
        
         If Len(sValue) = 3 Then
            sLangs = sLangs + Chr(9) + sValue
         End If
        
         sTemp = Mid(sTemp, place + 5)
         
      Loop
      
      If boolLangPrefix = False Then
 
         If sLangs <> "" And InStr(sLangs, sLang) = 0 Then
         
            MsgBox "WARNING: language prefix may be missing in the call number. The 041 tag indicates multiple languages present."
      
         ElseIf sLangs <> "" And sLang = "ENG" And sLangs <> sLang Then
         
            MsgBox "WARNING: language prefix may be missing in the call number. The 041 tag indicates multiple languages present."
      
         End If
      End If
   End If

End Sub


'########################################################################

Sub InsertCallNum(s948, f, sInitials)
   Dim CS as Object
   On Error Resume Next
   Set CS  = GetObject(,"Connex.Client")
   On Error GoTo 0

   Dim s901$, s949$
   Dim nBool
   
   CS.SetField 1, s948
   CS.SetField 1, "945  .b" & Chr(252)
   CS.SetField 1, "946  m"
   
   'sFormat codes in variable f
   If f = 1 Then
      'bluray
      s949 = "949  *b2=b;"
   ElseIf f = 2 Then
      'audiobook
      s949 = "949  *b2=u;"
   ElseIf f = 4 Then
      'dvd
      s949 = "949  *b2=v;"
   ElseIf f = 9 Then
      'large-print
      s949 = "949  *b2=l;"
   ElseIf f = 11 Then
      'readalong
      s949 = "949  *b2=5;"
   Else
      'print
      s949 = "949  *b2=a;"
   End If
   
   s949 = s949 & "recs=oclcgw;bn=zzzzz;"
   CS.SetField 1, s949

   s901 = "901  " & sInitials & " " & Chr(223) & "b CATBL"
   CS.SetField 1, s901
   
  CS.EndRecord
End Sub
