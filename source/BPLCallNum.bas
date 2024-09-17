'MacroName:BPL CallNum macro
'MacroDescription: Revised version of CallNum macro; 
'                  includes format, audience, call type and output selection;
'                  includes call number detection, local Dewey rules, call number conflict routine;
'                  supports selection of names for the call number based on multiple 6xx fields,
'                  improved behavior for computer science call numbers,
'                  added separation of cataloger's initials and code (pulled from a file instead)
'                  overlay string supplied for World Language materials

'v.3.4.0 (09-17-2024):
'  * removes ISBNs with "Kindle" qualifier
'  * adds fill character (Chr(252)) to missing or invalid call number elements
'  * sets Sierra bib format to 8 for READALONGS
'  * sets 4 digits after period limit except for 8xx
'  * fixes handling of dotless i common in Turkish materials by replacing it with uppercase i
'  * fixes duplication of ilocs in main menu
'v3.3.0 (02-06-2023) changes:
'  * removal of flags for 005.258 (programming for specific operating systems), 005.3582 (apps for specific mobile devices),
'        005.432 (specific operating systems), 005.5 (general purpose programs), 005.7585 & 005.7565 (database management systems),
'        006.74 (markup languages)
'v3.2.0 (12-5-2022) changes:
'  * READALONG format autodetection improvements and removal of flags for READALONG J-E pattern and print bib format
'v3.1.0 (07-29-2022) changes:
'  * adds record reformating routine before any parsing
'  * adds a flag if material includes multiple languages, but the call number does not include a language prefix
'v3.0.1 (07-15-2022) changes:
'  * fixes incorrect recognition of DVDs and labeling them as KIT format
'v3.0.0 (05-06-2022) changes:
'  * removes the routine that deletes unsupported subject vocabularies from 6xx fields (moved to CAT!UpdateExport macro)
'  * ceases removal of "Popular works" subdivision from LCSH
'  * introduces semantic versioning (major, minor, patch)
'v2.8 (04-15-2022) changes:
'  * removes AAT thesaurus as acceptable for bibs produced in BookOps/CAT (OCLC began assigning them automatically and they proliferated in records recently)
'  * adds homosaurus thesaurus (homoit)
'v2.7 (03-09-2022) changes:
'  * introduces Dewey+Name for literary works by a single author in a particular form (811 POE - collection of Edgar Poe poems)
'  * adds Modern Persian literature time table digits removal
'v2.6 changes:
'  * removal of the S52 segment from call numbers for Shakespeare
'v2.5 changes:
'  * refactoring of removal of unwanted subject headings
'v2.4 changes:
'  * title entry J-E have cutter consisting of the first letter of the title, switch from Rule 1 to Rule 2 (same as fiction)
'  * cuttering of titles with corporate main entry (110 tag) changed to full first word
'v2.3 changes:
'  * file with cataloger initials stored now in OCLC's profile directory
'v2.2 changes:
'  * autobiography cutter changes: main entry used for the cutter instead of the title entry
'v2.1 changes:
'  * bug fix: corrected removal of time table numbers from Russian literature
'v2.0 changes:
'  * feature DVD call number includes one of two type prefixes: TV and MOVIE
'v1.9 changes:
'  * Readalong format support
'  * Chinese lit time table handling added
'v1.8 changes:
'  * DVD call number cutter includes full first word of title
'  * bug fix: underscore nonspacing character (Chr(246)) handling fixed
'v1.7 changes:
'  *improved and simplified diactriticts function



Declare Function Dewey(sAudn,sCallType)
Declare Function FileDlgFunction(identifier$, action, suppvalue)
Declare Function CleanCuttersLastDigit(s082)
Declare Function Cutter(sCutterArr,sCallType,sBiog,sLTxt)
Declare Function HasIllegalCutter(sCutter)
Declare Function Normalized(sNameTitle, sTag)
Declare Sub CutterArray(sCutterArr,sCallType,sDewey)
Declare Sub SubjectChoice(sSubjectArr)
Declare Sub Diacritics(sNameTitle)
Declare Sub LocalDewey(s082,sCallType)
Declare Sub InsertCallNum(s099,sRecType,sFormItem,sLang,sAudn,f, sInitials)
Declare Sub Conflicts(booLangPrefix, s041, sAudn, sBiog, sCallType, sCutter, sLang, sRecType, sItemForm, sLitF, sTMat, sLTxt, f, a)



Sub Main

   Dim CS as Object
   Set CS  = GetObject(,"Connex.Client")

   If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 2 or CS.ItemType = 17 Then

      ' make sure all fields are reformated and subfields include required spacing      
      CS.Reformat
      
      Dim s041$, s099$, s100$, s130$, s300$, s600$, s650$, s655$, s730$, sCallType$, sAudn$, sLang$, sRecType$, sItemForm$, sFormatPrefix$, sAudiencePrefix$, _
       sLangPrefix$, sCutterArr$, sDewey$, sLitF$, sBiog$, sTMat$, sLTxt$
      Dim boolLangPrefix
      Dim iAudn As Integer
      Dim sFormat() As String
      Dim sAudience() As String
      Dim sOutput() As String
      
      'variables for storage and retrieval of cataloger's initials
      Dim filenumber As Integer
      Dim sFileName As String
      Dim sDefaultInitials$
      Dim sInitials$
      
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
         Print #filenumber, "xxx"
         Close #filenumber
         sDefaultInitials = "xxx"
      End If
      
      CS.GetFixedField "Type", sRecType
      CS.GetFixedField "Form", sItemForm
      CS.GetFixedField "Audn", sAudn
      CS.GetFixedField "Lang", sLang
      CS.GetFieldUnicode "100", 1, s100
      CS.GetFieldUnicode "300", 1, s300
      CS.GetFixedField "LitF", sLitF
      CS.GetFixedField "Biog", sBiog
      CS.GetFixedField "TMat", sTMat
      CS.GetFieldUnicode "041", 1, s041
      If InStr(s300, Chr(223) & "e") <> 0 Then
         s300 = Mid(s300, InStr(s300, Chr(223) & "e"))
      End If
      s300 = UCase(s300)
      
      ReDim sFormat(12)
         sFormat(0) = " "
         sFormat(1) = "AUDIO"
         sFormat(2) = "BOOK & CD"
         sFormat(3) = "BOOK & DVD"
         sFormat(4) = "CD"
         sFormat(5) = "CD-ROM"
         sFormat(6) = "DVD"
         sFormat(7) = "KIT"
         sFormat(8) = "LIB"
         sFormat(9) = "Mu"
         sFormat(10) = "NM"
         sFormat(11) = "READALONG"
         sFormat(12) = "WEB SITE"
      ReDim sAudience(1)
         sAudience(0) = "JUVENILE"
         sAudience(1) = "ADULT/YOUNG"
      ReDim sOutput(1)
         sOutput(0) = "add to record"
         sOutput(1) = "copy to clipboard"

      'Dialog box presenting to a cataloger choices for types of call numbers
      Begin Dialog UserDialog 220, 245, "BPL Call Number Macro" '.FileDlgFunction
         GroupBox 20, 50, 180, 162, ""
         OptionGroup .Type
         OptionButton  24,  55, 80, 14, "&EASY BOOK"
         OptionButton  24,  75, 80, 14, "&FICTION"
         OptionButton  24,  95, 80, 14, "&DEWEY"
         OptionButton  24,  115, 80, 14, "DEWEY+&NAME"
         OptionButton  24,  135, 80, 14, "DEWEY+&SUBJECT"
         OptionButton  24,  155, 80, 14, "&BIOGRAPHY"
         OptionButton  24,  175, 80, 14, "&MOVIE"
         OptionButton  24,  195, 80, 14, "&TV"
         
         Text          150, 57, 40, 24, "(jbb, jje, jer)"
         Text          130, 77, 64, 20, "( _fc, _my, _sf, _sh)"
         Text          134, 97, 64, 20, "( _nf, _ej)"
         Text          155, 117, 64, 14, "(_nf)"
         Text          155, 137, 30, 14, "(anf)"
         Text          155, 157, 30, 10, "( _bi)"
         Text          155, 177, 30, 10, "( _dv)"
         Text          155, 197, 30, 10, "( _dv)"
         DropListBox   20, 10, 70, 115, sFormat(), .sFormat
         Text          92, 12, 40, 14, "FORMAT"
         Textbox       135, 11, 20, 13, .sInitials
         Text          170, 12, 20, 13, "initials"
         DropListBox   20, 30, 70, 40, sAudience(), .sAudience
         Text          92, 32, 40, 14, "AUDIENCE"
         DropListBox   135, 30, 65, 60, sOutput(), .sOutput
         OkButton      40, 220,  54, 16
         CancelButton  120, 220,  54, 16
      End Dialog

      Dim dCallNum as UserDialog

      'populate INITIALS box with default value
      dCallNum.sInitials = sDefaultInitials

      'selects most likely audience, format, and call number choice
      If InStr("abcj", sAudn) <> 0 And sAudn <> "" Then
         dCallNum.sAudience = 0
      Else
         dCallNum.sAudience = 1
      End If

      ' format & callNum type
      If sRecType = "a" Then
         ' microfilm & microfish
         If sItemForm = "a" Or sItemForm = "b" Then
            dCallNum.sFormat = 10
         ElseIf InStr(s300, "AUDIO-ENABLED BOOK") <> 0 Or InStr(s300, "AUDIO ENABLED BOOK") <> 0 Or InStr(s300, "WONDERBOOK") <> 0 Then
            dCallNum.sFormat = 11
         End If
         
         
         If InStr("1fj", sLitF) <> 0 And sLitF <> "" Then
            If InStr("ab", sAudn) <> 0 and sAudn <> "" Then
               ' picture book
               dCallNum.Type = 0
            Else
               ' regular fiction
               dCallNum.Type = 1
            End If

         ElseIf InStr("dehimps", sLitF) <> 0 And s100 <> "" And sBiog <> "a" And sBiog <> "b" Then
            ' dewey + name
            dCallNum.Type = 3

         ElseIf Instr("0i", sLitF) <> 0 And (sBiog = "a" Or sBiog = "b") Then
            dCallNum.Type = 5
            
         Else
            ' dewey
            dCallNum.Type = 2
            
         End If

      ElseIf sRecType = "c" Or sRecType = "d" Then
         dCallNum.sFormat = 9
         dCallNum.Type = 2
      ElseIf sRecType = "i" Then
         CS.GetFixedField "LTxt", sLTxt
         If sLTxt = "f" Then
            dCallNum.Type = 1
         ElseIf InStr("abmt", sLTxt) <> 0 And sLTxt <> "" Then
            dCallNum.Type = 5
         Else
            dCallNum.Type = 2
         End If
         If InStr(s300, "AUDIO-ENABLED BOOK") <> 0 Or InStr(s300, "AUDIO ENABLED BOOK") <> 0 Then
            dCallNum.sFormat = 11
         ElseIf InStr(s300, "CM") <> 0 Or InStr(s300, "P.") <> 0 Or InStr(s300, "PAGES") <> 0 Or InStr(s300, "BOOK") <> 0 Then
            dCallNum.sFormat = 2
         Else
            dCallNum.sFormat = 1
         End If
      ElseIf sRecType = "j" Then
         dCallNum.sFormat = 5
         MsgBox "Please consider using BPLMusicCD macro instead"
      ElseIf sRecType = "g" And (sTMat = "v" Or sTMat = "m") Then
         dCallNum.sFormat = 6
         ' type 6 (movie) or 6 (tv)
         'sample first 650 if Drama
         CS.GetField "650", 1, s650
         If InStr(s650, "Drama.") <> 0 Or InStr(s650, "Juvenile drama.") <> 0 Then
            'fictional work
            CS.GetField "655", 1, s655
            CS.GetField "130", 1, s130
            CS.GetField "730", 1, s730
            If InStr(s655, "television") <> 0 Or InStr(s130, "(Television") <> 0 Or InStr(s730, "(Television") <> 0 Then
               dCallNum.Type = 7
            Else
               dCallNum.Type = 6
            End If
         Else
            bool = CS.GetField("600", 1, s600)
            If bool = TRUE Then
               'suggest biography call type
               dCallNum.Type = 5
            Else
               dCallNum.Type = 2
            End If
         End If
      ElseIf sRecType = "m" And sItemForm = "o" Then
         dCallNum.sFormat = 12
         dCallNum.Type = 2
      ElseIf sRecType = "m" And sItemForm = "q" Then
         dCallNum.sFormat = 5
         dCallNum.Type = 2
      End If

      On Error Resume Next
      Dialog dCallNum 
      If Err = 102 Then Exit Sub
      
      'construct field 099 with call number
      f = dCallNum.sFormat
      a = dCallNum.sAudience
      If dCallNum.sFormat <> 0 Then
         sFormatPrefix = sFormat(dCallNum.sFormat) & " " & Chr(223) & "a "
      Else
         Goto Lang
      End If
Lang:
      'if English material skip, if WL insert appropriate string with language code
      sLang = UCase(sLang)
      If sLang = "ENG" Then
         boolLangPrefix = False
         Goto AudnType:
      Else
         sLangPrefix = sLang & " " & Chr(223) & "a "
         boolLangPrefix = True
      End If
AudnType:     
      If dCallNum.sAudience = 0 Then
         sAudnPrefix = "J " & Chr(223) & "a "
      Else
         Goto CaseType
      End If   
CaseType:
      Select Case dCallNum.Type
      Case 0
      'creates call number for juvenile easy and easy reader books
         sCallType = "eas"
         s099 = "099  " & sFormatPrefix & sLangPrefix & "J-E " & Chr(223) & "a "
      Case 1
      'creates call number for fiction
         sCallType = "fic"
         s099 = "099  " & sFormatPrefix & sLangPrefix & sAudnPrefix & "FIC " & Chr(223) & "a "
      Case 2
      'creates call number for Dewey classification
         sCallType = "dew"
         sDewey = Dewey(sAudn,sCallType)
         s099 = "099  " & sFormatPrefix & sLangPrefix & sAudnPrefix & sDewey
      Case 3
      'creates call number for literary works by a single author (pattern: Dewey + Name)
         sCallType = "den"
         sDewey = Dewey(sAudn,sCallType)
         s099 = "099  " & sFormatPrefix & sLangPrefix & sAudnPrefix & sDewey
      Case 4
      'creates call number for Dewey with subject
         sCallType = "des"
         sDewey = Dewey(sAudn,sCallType)
      'check if this call number is valiable for juvenile material
         s099 = "099  " & sFormatPrefix & sLangPrefix & sAudnPrefix & sDewey
      Case 5
      'creates call number for biography
         sCallType = "bio"
         s099 = "099  " & sFormatPrefix & sLangPrefix & sAudnPrefix & "B " & Chr(223) & "a "
      Case 6
      'creates call number for DVD feature movies
         sCallType = "mov"
         s099 = "099  " & sFormatPrefix & sLangPrefix & sAudnPrefix & "MOVIE " & Chr(223) & "a "
      Case 7
      'creates call number for DVD tv fictional show
         sCallType = "tvs"
         s099 = "099  " & sFormatPrefix & sLangPrefix & sAudnPrefix & "TV " & Chr(223) & "a "
      End Select
      
      Call CutterArray(sCutterArr,sCallType,sDewey)
      sCutter = Cutter(sCutterArr,sCallType,sBiog,sLTxt)

      'Output selection
      If dCallNum.sOutput = 1 Then
         'for in-memory output remove subfields coding
         s099 = s099 & sCutter
         s099 = Mid(s099, 6)
         Do While InStr(s099, Chr(223)) <> 0
            place = InStr(s099, Chr(223))
            lt$ = Left(s099, place - 1)
            rt$ = Mid(s099, place + 3)
            s099 = lt$ + rt$
         Loop
         Clipboard.Clear
         Clipboard.SetText s099
      Else
         'poplulate default intitials
         sInitials = dCallNum.sInitials
         filenumber = FreeFile
         Open sFileName for Output As filenumber
         Print #filenumber, sInitials
         Close #filenumber
         
         'insert call number
         s099 = s099 & sCutter
         Call InsertCallNum(s099,sRecType,sItemForm,sLang,sAudn,f, sInitials)
      End If
   Else
      MsgBox "INFO: A bibliographic record must be displayed in order to use this macro."
      Goto Done
   End If   
Done:

Call Conflicts(boolLangPrefix, s041, sAudn, sBiog, sCallType, sCutter, sLang, sRecType,sItemForm, sLitF, sTMat, sLTxt, f, a)
End Sub
'##############################################################

Function FileDlgFunction(identifier$, action, suppvalue)
   DlgFocus 0
End Function


'##############################################################

Function Dewey(sAudn,sCallType)
'creates string with Dewey number taken from 082 field
   Dim CS as Object
   Set CS  = GetObject(,"Connex.Client")

   Dim s082$
   Dim place as Integer
   
   bool082 = CS.GetField("082", 1, s082)
   If bool082 = FALSE Then
      MsgBox "MISSING DATA: no 082 field in the record. Please supply Dewey classification number manually."
      Dewey = Chr(252) & " " & Chr(223) & "a "
      Goto Done
   Else
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
      If InStr(s082, "Fic") <> 0 Or InStr(s082, "E") <> 0 Then
         MsgBox "INCOMPLETE: Please correct the call number."
         Dewey = Chr(252) & " " & Chr(223) & "a "
         Goto Done
      End If
   End If

   'removes 0 if it's the last digit and loop
   If sAudn = "j" Or sAudn = "a" Or sAudn = "b" Or sAudn = "c" Then
      s082 = RTrim(Mid(s082,6,8))
      Call LocalDewey(s082,sCallType)
      s082 = CleanCuttersLastDigit(s082)
   Else
      'with exception to 8xx limit number of digits after decimal point to 4 only
      If Mid(s082, 6, 1) = "8" Then
         s082 = Mid(s082, 6, 20)
      Else
         s082 = Mid(s082,6,8)
      End If
      Call LocalDewey(s082,sCallType)
      s082 = CleanCuttersLastDigit(s082)

   End If
   Dewey = s082 & " " & Chr(223) & "a "
Done:
End Function

'########################################################
Function CleanCuttersLastDigit(s082)

   Dim sLastDigit$
   Dim x as Integer
   
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
   
   CleanCuttersLastDigit = s082

End Function

'########################################################

Function Cutter(sCutterArr,sCallType,sBiog,sLTxt)
'Cuttering rules:
   'Rule1: full surname from 100 field only
   'Rule2: full surname from 100 field or 1st letter of 245
   'Rule3: (1st letter of 100 or 110) or 1st letter of 245
   'Rule4: full 1st word from 245
   'Rule5: full surname from 600 and 1st letter of main entry
   'Rule6: full surname from 600 and 1st letter of 245
   'Rule7: no cutter
   
   Dim sSubjectArr$
   
   If sCallType = "eas" Then
      Goto Rule2
   End If
   If sCallType = "fic" Then
      Goto Rule2
   End If
   If sCallType = "dew" Then
      Goto Rule3
   End If
   If sCallType = "den" Then
      Goto Rule1
   End If
   If sCallType = "des" Then
      GoTo Rule5
   End If
   If sCallType = "bio" Then
      GoTo Rule5
   End If
   If sCallType = "mov" Or sCallType = "tvs" Then
      Goto Rule4
   End If
   
Rule1:
'full author last name
   If InStr(sCutterArr, "100_" & Chr(252)) <> 0 Then
      MsgBox "INCOMPLETE: Dewey+Name should be used for literary works by a single author (ex.: 811 POE = collection of E.Poe poems)"
      sCutter = Chr(252)
   Else
      start_point = InStr(sCutterArr, "100_")
      temp$ = Mid(sCutterArr, start_point + 4)
      end_point = InStr(temp$, Chr(9))
      sCutter = Left(temp$, end_point - 1)
   End If
   Goto Done
   
Rule2:
'full author last name, first word of 110, or 1st letter of title
   If InStr(sCutterArr, "100_" & Chr(252)) = 0 Then
      start_point = InStr(sCutterArr, "100_")
      temp$ = Mid(sCutterArr, start_point + 4)
      end_point = InStr(temp$, Chr(9))
      sCutter = Left(temp$, end_point - 1)
   ElseIf InStr(sCutterArr, "110_" & Chr(252)) = 0 Then
      start_point = InStr(sCutterArr, "110_")
      temp$ = Mid(sCutterArr, start_point + 4)
      end_point = InStr(temp$, Chr(9))
      sCutter = Left(temp$, end_point - 1)
      If InStr(sCutter, " ") <> 0 Then
         end_point = InStr(sCutter, " ")
         sCutter = Left(sCutter, end_point - 1)
      End If
   Else
      If InStr(sCutterArr, "245_") <> 0 Then
         start_point = InStr(sCutterArr, "245_")
         sCutter = Mid(sCutterArr, start_point + 4, 1)
         
         If HasIllegalCutter(sCutter) = 1 Then
            sCutter = Chr(252)
         End If

      End If
   End If
   Goto Done
   
Rule3:
'1st letter of main entry
   If InStr(sCutterArr, "100_" & Chr(252)) = 0 Then
      start_point = InStr(sCutterArr, "100_")
      sCutter = Mid(sCutterArr, start_point + 4, 1)
   ElseIf InStr(sCutterArr, "110_" & Chr(252)) = 0 Then
      start_point = InStr(sCutterArr, "110_")
      sCutter = Mid(sCutterArr, start_point + 4, 1)
   ElseIf InStr(sCutterArr, "245_" & Chr(252)) = 0 Then
      start_point = InStr(sCutterArr, "245_")
      sCutter = Mid(sCutterArr, start_point + 4, 1)
      If HasIllegalCutter(sCutter) = 1 Then
         sCutter = Chr(252)
      End If
   Else
      MsgBox "MISSING DATA: Unable to create a cutter."
      sCutter = Chr(252)
   End If
   Goto Done
   
Rule4:
'full 1st word of title
   If InStr(sCutterArr, "245_" & Chr(252)) = 0 Then
      start_point = InStr(sCutterArr, "245_")
      temp$ = Mid(sCutterArr, start_point + 4)
      end_point = InStr(temp$, Chr(9))
      sCutter = Left(temp$, end_point - 1)
      If InStr(sCutter, " ") <> 0 Then
         end_point = InStr(sCutter, " ")
         sCutter = Left(sCutter, end_point - 1)
      End If
   End If
   Goto Done
   
Rule5:
'full last name from 6xx and 1st letter of main entry
   temp$ = sCutterArr
   x = 0
   Do While InStr(temp$, "600_") 
      start_point = InStr(temp$, "600_")
      temp$ = Mid(temp$, start_point + 4)
      end_point = InStr(temp$, Chr(9))
      temp2$ = Left(temp$, end_point)
      sSubjectArr = sSubjectArr & temp2$
      x = x + 1
   Loop
   temp$ = sCutterArr
   Do While InStr(temp$, "610_") 
      start_point = InStr(temp$, "610_")
      temp$ = Mid(temp$, start_point + 4)
      end_point = InStr(temp$, Chr(9))
      temp2$ = Left(temp$, end_point)
      sSubjectArr = sSubjectArr & temp2$
      x = x + 1
   Loop
   temp$ = sCutterArr
   Do While InStr(temp$, "650_") 
      start_point = InStr(temp$, "650_")
      temp$ = Mid(temp$, start_point + 4)
      end_point = InStr(temp$, Chr(9))
      temp2$ = Left(temp$, end_point)
      sSubjectArr = sSubjectArr & temp2$
      x = x + 1
   Loop
   temp$ = sCutterArr
   Do While InStr(temp$, "630_") 
      start_point = InStr(temp$, "630_")
      temp$ = Mid(temp$, start_point + 4)
      end_point = InStr(temp$, Chr(9))
      temp2$ = Left(temp$, end_point)
      sSubjectArr = sSubjectArr & temp2$
      x = x + 1
   Loop
   If x = 0 Then
      MsgBox "MISSING DATA: Record doesn't have valid 6xx field for this class number"
      sCutter = Chr(252) 
   ElseIf x = 1 Then
      sCutter = RTrim(sSubjectArr)
   Else
      Call SubjectChoice(sSubjectArr)
      sCutter = RTrim(sSubjectArr)
      If Len(sCutter) = 0 Then
         sCutter = Chr(252)
      End If
   End If
   
   If sBiog = "b" Or (InStr("bt", sLTxt) <> 0 And sBiog <> "") Or (sBiog = "a" And sLTxt = "") Or (sBiog = "" And InStr("am", sLTxt) <> 0 And sLTxt <> "") Or sCallType = "des" Or sCallType = "bio" Then
      If InStr(sCutterArr, "100_" & Chr(252)) = 0 Then
         start_point = InStr(sCutterArr, "100_")
         temp$ = Mid(sCutterArr, start_point + 4, 1)
         sCutter = sCutter & " " & Chr(223) & "a " & temp$
      ElseIf InStr(sCutterArr, "110_" & Chr(252)) = 0 Then
         start_point = InStr(sCutterArr, "110_")
         temp$ = Mid(sCutterArr, start_point + 4, 1)
         sCutter = sCutter & " " &  Chr(223) & "a " & temp$
      ElseIf InStr(sCutterArr, "245_" & Chr(252)) = 0 Then
         start_point = InStr(sCutterArr, "245_")
         temp$ = Mid(sCutterArr, start_point + 4, 1)
         sCutter = sCutter & " " &  Chr(223) & "a " & temp$
      End If
   Else
      MsgBox "MISSING DATA: Could not determine the cutter. Please assign manually."
      sCutter = sCutter & " " &  Chr(223) & "a " & Chr(252)
   End If
   Goto Done
   
Rule6:
'Rule obsolete (cuttering autobiography has changed to main entry on June 14, 2019)
'full last name from 600 and 1st letter of title if autobiography
   If InStr(sCutterArr, "245_" & Chr(252)) = 0 Then
      start_point = InStr(sCutterArr, "245_")
      temp$ = Mid(sCutterArr, start_point + 4, 1)
      sCutter = sCutter & " " & Chr(223) & "a " & temp$
   End If
   
Rule7:
'no Cutter
   Cutter = ""
   
Done:
   Cutter = sCutter
End Function

'########################################################

Function HasIllegalCutter(sCutter)

   If InStr(0123456789, sCutter) <> 0 Then
      MsgBox "INCORRECT call number: a cutter can not consist of a digit. Please use first letter of spelled out number in the language of the cataloged material."
      HasIllegalCutter = 1
   Else
      HasIllegalCutter = 0
   End If

End Function

'########################################################

Sub SubjectChoice(sSubjectArr)
   
   Begin Dialog UserDialog 200, 60, "Select Subject"
      DropListBox  8, 15, 100, 250, sSubjectArr, .sSubjectArr
      OkButton        130, 15,  54, 16
      CancelButton   130, 35,  54, 16
   End Dialog
   Dim dSubject as UserDialog
   On Error Resume Next
   Dialog dSubject
   If Err = 102 Then
      sSubjectArr = ""
      Exit Sub
   End If
   
   n = dSubject.sSubjectArr + 1
   temp$ = sSubjectArr
   x = 0
   Do
      place = InStr(temp$, Chr(9))
      If place <> 0 Then
         lt$ = Left(temp$, place - 1)
         temp$ = Mid(temp$, place + 1)
      Else
         lt$ = temp$
      End If
      x = x + 1
   Loop Until x >= n
   sSubjectArr = lt$

End Sub

'########################################################

Function Normalized(sNameTitle, sTag)
   'Checks if Unicode is present and shows warning or adds fill character for empty strings.
   'Activates Diacritics routine.

   If InStr(sNameTitle, "&#") <> 0 Then
      MsgBox "Field " & Left(sNameTitle, 3) & " includes a possibly incorrectly coded character (see &#x code). Please replace with ALA diacritics: '" & Mid(sNameTitle, 6) &  "'" 
      Normalized = sTag & "_" & Chr(252)
   ElseIf sNameTitle = "" Then
      Normalized = sTag & "_" & Chr(252)
   Else
      Indicator = Mid(sNameTitle,5,1)
      If Indicator = "0" Or Indicator = " " Then
         lt$ = sTag & "_"
         rt$ = Mid(sNameTitle, 6, 30)
         sNameTitle = lt$ + rt$
      Else
         lt$ = sTag & "_"
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

'########################################################

Sub CutterArray(sCutterArr,sCallType,sDewey)
'Creates a string used in cuttering based on author/title field
   Dim CS as Object
   Set CS  = GetObject(,"Connex.Client")
   Dim sNameTitle$, sTag$
   Dim i as Integer
   
   'linked fields with Non-Latin script are displayed first,
   'they should be ignored for cuttering purposes;
   'occasionally primary Latin script (MARC-8) will have a unicode-encoded
   'character if it is not supported by MARC-8 (for example African inverted e - see #1163881264 for example)
   'and in such situations CutterArray should display a warning but allow to use an element for cuttering
      
   sTag = "100"   
   bool100 = CS.GetFieldUnicode(sTag, 2, sNameTitle)
   
   If bool100 = False Then
      CS.GetFieldUnicode sTag, 1, sNameTitle
   End If
   sNameTitle = Normalized(sNameTitle, sTag)
   sCutterArr = sNameTitle & Chr(9)
   
   sTag = "110"
   bool110 = CS.GetFieldUnicode(sTag, 2, sNameTitle)
   If bool110 = False Then
      CS.GetFieldUnicode sTag, 1, sNameTitle
   End If
   sNameTitle = Normalized(sNameTitle, sTag)
   sCutterArr = sCutterArr & sNameTitle & Chr(9)
   
   sTag = "245"
   bool245 = CS.GetFieldUnicode(sTag, 2, sNameTitle)
   If bool245 = False Then
         CS.GetFieldUnicode sTag, 1, sNameTitle
   End If
   sNameTitle = Normalized(sNameTitle, sTag)
   sCutterArr = sCutterArr & sNameTitle & Chr(9)

   
   If sCallType = "des" Or sCallType = "bio" Then
      i = 1
      sTag = "600"
      Do While CS.GetFieldUnicode("600", i, sNameTitle)
         If InStr(sNameTitle, "&#") = 0 And Mid(sNameTitle, 5, 1) = "0" Then
            sNameTitle = Normalized(sNameTitle, sTag)
            sCutterArr = sCutterArr & sNameTitle & Chr(9)
         End If
         i = i + 1
      Loop
      If sCallType = "des" Then
         i = 1
         sTag = "610"
         Do While CS.GetFieldUnicode("610", i, sNameTitle)
            If InStr(sNameTitle, "&#") = 0 And Mid(sNameTitle, 5, 1) = "0" Then
               sNameTitle = Normalized(sNameTitle, sTag)
               If InStr(sNameTitle, "(") Then
                  place = InStr(sNameTitle, "(")
                  sNameTitle = Left(sNameTitle, place-1)
                  sNameTitle = RTrim(sNameTitle)
               End If
               sCutterArr = sCutterArr & sNameTitle & Chr(9)
            End If
         i = i + 1
         Loop
         
         If Left(sDewey, 3) = "004" Or Left(sDewey, 3) = "005" Or Left(sDewey, 3) = "006" Then
            i = 1
            sTag = "630"
            Do While CS.GetFieldUnicode("630", i, sNameTitle)
               If InStr(sNameTitle, "&#") = 0 And Mid(sNameTitle, 5, 1) = "0" Then
                  sNameTitle = Normalized(sNameTitle, sTag)
                  If InStr(sNameTitle, "(") Then
                     place = InStr(sNameTitle, "(")
                     sNameTitle = Left(sNameTitle, place-1)
                     sNameTitle = RTrim(sNameTitle)
                  End If
                  sCutterArr = sCutterArr & sNameTitle & Chr(9)
               End If
               i = i + 1
            Loop
            i = 1
            sTag = "650"
            Do While CS.GetFieldUnicode("650", i, sNameTitle)
               If InStr(sNameTitle, "&#") = 0 And Mid(sNameTitle, 5, 1) = "0" Then
                  sNameTitle = Normalized(sNameTitle, sTag)
                  If InStr(sNameTitle, "(") Then
                     place = InStr(sNameTitle, "(")
                     sNameTitle = Left(sNameTitle, place-1)
                     sNameTitle = RTrim(sNameTitle)
                  End If
                  sCutterArr = sCutterArr & sNameTitle & Chr(9)
               End If
               i = i + 1
            Loop
         End If
      End If
  End If

   
End Sub

'####################################################################

Sub Diacritics(sNameTitle)
   'removes diacritic marks and other unwanted characters from a string
 
   i = 1
   While i <= Len(sNameTitle)
      CheckChar = Mid(sNameTitle, i, 1)
      
      Select Case CheckChar
         'characters above letter (example: acute, breve, umlaut, macron, etc.)
         Case Chr(226), Chr(174), Chr(176), Chr(230), Chr(239), Chr(234), Chr(227), Chr(238), Chr(250), Chr(251), Chr(225), Chr(96), Chr(233), Chr(254), Chr(237), Chr(235), Chr(236), Chr(229), Chr(167), Chr(228), Chr(126), Chr(183), Chr(232), Chr(231)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
            i = i - 1
         'characters below letter (example: cedilla, hooks, etc.)
         Case Chr(240), Chr(248), Chr(247), Chr(241), Chr(244), Chr(242), Chr(243), Chr(245), Chr(249), Chr(246)
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
         Case Chr(184)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & "i" & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
         'u with hook
         Case Chr(189), Chr(173)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & "u" & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
         'characters to be removed completely
         Case Chr(176), Chr(174), Chr(167)
            sNameTitle = Mid(sNameTitle, 1, i - 1) & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
            i = i - 1
'         Case ":", ";", "/"
'            sNameTitle = Mid(sNameTitle, 1, i - 1) & Mid(sNameTitle, i + 1, Len(sNameTitle) - i)
'            i = i - 1
          
      End Select
      i = i + 1   
   Wend
   
End Sub

'##################################################

Sub LocalDewey(s082,sCallType)

   Dim sClassNum$, s1stDig$, s1stTwoDig$, s1stThreeDig$, s1stFiveDig$, s1stSixDig$, s1stSevenDig$, s3rdTo7thDig$
   Dim s2ndDig$, s3rdDig$, s5thDig$, s6thDig$, s7thDig$, s8thDig$, s9thDig$
   Dim sLitTimeTableMsg$
   
   sClassNum = Left(s082,8)
   
   s1stDig = Left(s082,1)
   s2ndDig = Mid(s082,2,1)
   s3rdDig = Mid(s082,3,1)
   s5thDig = Mid(s082,5,1)
   s6thDig = Mid(s082,6,1)
   s7thDig = Mid(s082,7,1)
   s8thDig = Mid(s082,8,1)
   s9thDig = Mid(field082,9,1)
   
   s1stTwoDig = Left(s082,2)
   s1stThreeDig = Left(s082,3)
   s1stFiveDig = Left(s082,5)
   s1stSixDig = Left(s082,6)
   s1stSevenDig = Left(s082,7)
   s3rdTo7thDig = Mid(s082,3,5)
   
   sLitTimeTableMsg = "INFO: BPL does not use the time table in literature classification. Removing time table digits..."
   
   'Shakespeare call number
   If sClassNum = "822.33" Then
      MsgBox "INCOMPLETE: Please correct the cutter. Call number for Shakespeare includes Dewey number, letter and number from Dewey optional table, and picked by cataloger cutter, example: 822.33 S7 G (Hamlet - text, edited by R.Gill). Please do not use cuttering created by the macro."
      Goto Done
   End If
   
   'specific programming languages
   If sCallType <> "des" Then
      If sClassNum = "005.133" Then
         s082 = s082 & " " & Chr(223) & "a " & Chr(252)
         MsgBox "INCOMPLETE: Use Dewey option to arrange alphabetically by name of programming language, example: 005.133 C++ K. Please insert language name before the cutter or run macro again and select Dewey + Subject option."
      End If
   
      'specific mobile devices
      If sClassNum = "004.1675" Then
         s082 = s082 & " " & Chr(223) & "a " & Chr(252)
         MsgBox "INCOMPLETE: Use Dewey option to arrange alphabetically by name of mobile device, example: 004.1675 IPHONE C. Please insert the name of the device before the cutter or run macro again and select Dewey + Subject option."
      End If
   
      'specific operating systems of personal computers
      If sClassNum = "005.446" Then
         s082 = s082 & " " & Chr(223) & "a " & Chr(252)
         MsgBox "INCOMPLETE: Use Dewey option to arrange alphabetically by name of operating system, example: 005.446 VISTA B. Please insert the name of operating system before the cutter or run macro again and select Dewey + Subject option."
      End If
   End If
   
   'travel guides
   If s1stTwoDig = "91" And s3rdDig <> "0" And s3rdDig <> "1" And s3rdDig <> "2" Then
      MsgBox "INFO: 082 field indicates the item is a guidebook. BPL Dewey number for guidebooks stops before 04 notation."
      If InStr(s082, "04") Then
         place = InStr(s082, "04")
         s082 = Left(s082, place-1)
      End If
   End If
   
   'Bible modern translations
   If s1stFiveDig = "220.5" Then
      MsgBox "INCOMPLETE: Please correct the cutter. BPL call number for Bible includes Sanborn cutter B582 and cataloger's supplied cutter (use editor or publisher) of one or more letters."
      s082 = s082 & " " & Chr(223) & "a " & "B582" 
   End If
   
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

'################################################################################

Sub Conflicts(booLangPrefix, s041, sAudn, sBiog, sCallType, sCutter, sLang, sRecType, sItemForm, sLitF, sTMat, sLTxt, f, a)
   Dim CS as Object
   Set CS  = GetObject(,"Connex.Client")

   Dim s082$, s338$, s347$, sFields$
   Dim sTemp$, sLangs$
   'f - numerical code of format
   'a - numerical code of audience

   bool082 = CS.GetField("082", 1, s082)
AudnCheck:
   If sCallType = "eas" Then
      If InStr("ab", sAudn) <> 0 And sAudn <> "" Then
         If bool082 = TRUE Then
            If InStr(s082, "[FIC]") <> 0 Then
               MsgBox "AUDIENCE conflict: The material is classed as fiction for older children (082 field - [Fic]). Please verify your selection."
            End If
         End If
      ElseIf sAudn = "j" Then
         MsgBox "INFO: Caution advised. Record coded as broad juvenile material (fixed field Audn: j)."
      Else
         MsgBox "AUDIENCE conflict: Record not coded as eas material (fixed field Audn). Please verify your selection."
      End If
   Else
      If InStr("cdefgj", sAudn) <> 0 Or sAund = "" Then
         If InStr("cj", sAudn) <> 0 And sAudn <> "" Then
            If a = 0 Then
               If InStr(s082, "[E]") <> 0 And f <> 11 Then
                  MsgBox "AUDIENCE conflict: The material is classed as eas fiction (082 field - [E]). Please verify your selection."
               Else
                  Goto LitCheck
               End If
            Else
               MsgBox "AUDIENCE conflict: Record coded as juvenile material (fixed field Audn). Please verify your selection."
            End If
         Else
            If a = 1 Then
               Goto LitCheck
            Else
               MsgBox "AUDIENCE conflict: Record not coded as adult/young adult material (fixed field Audn). Please verify your selection."
            End If
         End If
      Else
         MsgBox "AUDIENCE conflict: Record coded as material for young readers, age 0-8 (fixed field Audn-a,b). Please consider using J-E call number."
      End If
      If a = 1 and f = 11 Then
         MsgBox "AUDIENCE conflict: Readalong books are intended for juvenile readers and call number must include J prefix."
      End If
   End If
LitCheck:
   If sRecType = "a" Then
      If sCallType = "fic" Then
         If InStr("1fj", sLitF) = 0 Then
            MsgBox "LITERARY FORM conflict: Fixed field indicates the material is non-fiction. Please verify your selection."
         End If
      ElseIf sCallType = "eas" Then
         Goto BioCheck
      Else
         If InStr("1fj", sLitF) <> 0 Then
            MsgBox "LITERARY FORM conflict: Fixed field indicates the material is fiction. Please verify your selection."
         End If
      End If
   End If
BioCheck:
   If sRecType = "a" Then
      If InStr("bio+des", sCallType) <> 0 Then 
         If InStr("abd", sBiog) = 0 Then
            MsgBox "INFO: Caution advised. Record not coded as a biography/autobiography."
         Else
            Goto FormCheck
         End If
      Else
         If InStr("ab", sBiog) <> 0 And sBiog <> "" Then
            MsgBox "INFO: Caution advised. Record coded as a biography/autobiography."
         End If 
      End If
   ElseIf sRecType = "i" Then
      If InStr("abmt", sLTxt) <> 0 and sLTxt <> "" Then
         If InStr("bio+des", sCallType) = 0 Then
            MsgBox "INFO: Caution advised. Record coded as a biography/autobiography."
         End If 
      End If
   End If
FormCheck:
   If f = 0 Then
      If sRecType <> "a" Then
         MsgBox "FORMAT conflict: bibliographical record type is not for language material. Please verify your selection."
      End If
   ElseIf f = 1 Then
      If sRecType <> "i" Then
         MsgBox "FORMAT conflict: bibliographical record type is not for non-musical sound recording. Please verify your selection."
      End If
   ElseIf f = 2 Or f = 3 Then
      If sRecType <> "a" And sRecType <> "i" Then
         MsgBox "FORMAT conflict: bibliographical record type is not for language material or non-musical sound recordsing. Please verify your selection."
      End If
   ElseIf f = 4 Then
      If sRecType <> "j" Then
         MsgBox "FORMAT conflict: bibliographical record type is not for musical sound recording. Please verify your selection."
      End If
   ElseIf f = 5 Then
      If sRecType <> "m" and sItemForm <> "q" Then
         MsgBox "FORMAT conflict: bibliographical record type is not for computer file. Please verify your selection."
      End If
   ElseIf f = 6 Then
      If sRecType <> "g" Then
         MsgBox "FORMAT conflict: bibliographical record type is not for visual materials. Please verify your selection."
      End If
   ElseIf f = 9 Then
      If sRecType <> "c" Or sRecType <> "d" Then
         MsgBox "FORMAT conflict: bibliographical record type is not for notated music. Please verify your selection."
      End If
   ElseIf f = 10 Then
      If sRecType <> "a" And InStr("ab", sItemForm) <> 0 And sItemForm <> "" Then
         MsgBox "FORMAT conflict: bibliographical record type is not for microfilm or mirofiche. Please verify your selection."
      End If
   ElseIf f = 11 Then
      If InStr("ai", sRecType) = 0  Then
         MsgBox "FORMAT conflict: bibliographical record type is not for audio-enabled books. Please verify your selection."
      End If   
   ElseIf f = 12 Then
      If sRecType <> "m" And sItemForm <> "o" Then
         MsgBox "FORMAT conflict: bibliographical record type is not for electronic remote access resources. Please verify your selection."
      End If   
   End If
   
   'check if material format is correct for DVD
   If sRecType = "g" Then
      n = 1
      sFields = ""
      nBool = CS.GetField("347", n, s347)
      Do While nBool = TRUE
         sFields = sFields + s347
         n = n + 1
         nBool = CS.GetField("347", n, s347)
      Loop
      n = 1
      nBool = CS.GetField("338", n, s338)
      Do While nBool = TRUE
         sFields = sFields + s338
         n = n + 1
         nBool = CS.GetField("338", n, s338)
      Loop
      If InStr(sFields, "DVD") = 0 And InStr(sFields, "videodisc") = 0 Then
         MsgBox "FORMAT conflict: record seems to indicate a different material type than DVD. Please verify your selection."
      End If
   End If
   
LangCheck:

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

Sub InsertCallNum(s099,sRecType,sItemForm,sLang,sAudn,f,sInitials)

   Dim CS as Object
   Set CS  = GetObject(,"Connex.Client")

   Dim s949$, sCD, sDVD, sIndcator$, sOverlay$, sLargePrint$, sReadalong$, SierraCode$, s007$, s006$
   Dim n as Integer

   CS.SetField 1, s099
   
   s949 = "949  *"
   sOverlay = "recs=b;ov=."
   sLargePrint = "b2=l;"
   sDVD = "b2=h;"
   sCD = "b2=i;"
   sReadalong = "b2=8;"

   
   If sRecType = "a" Then
      If sItemForm = "d" Then
         MsgBox "INFO: Large print record. Setting Sierra format code to large print."
         SierraCode = s949 & sLargePrint & sOverlay
      ElseIf f = 11 Then
         SierraCode = s949 & sReadalong & sOverlay
      Else
         SierraCode = s949 & sOverlay
      End If
   ElseIf sRecType = "g" Then
      SierraCode = s949 & sDVD & sOverlay
   ElseIf sRecType = "i" Then
      If f = 11 Then
         SierraCode = s949 & sReadalong & sOverlay
      Else
         SierraCode = s949 & sCD & sOverlay
      End If
   End If
   
   CS.SetField 1, SierraCode
   
   'BOOK & CD call numbers lacking 007 or 006
   If f = 2 Then
      bool007 = CS.GetField("007", 1, s007)
      If bool007 <> TRUE Then
         'creates 007 field with default values for audio CD
         s007 = "007  s " & Chr(223) & "b d " & Chr(223) & "d f " & Chr(223) & "e u " & Chr(223) & "f n " & Chr(223) & "g g " & Chr(223) & "h n " & Chr(223) & "i n " & Chr(223) & "k m " & Chr(223) & "m e " & Chr(223) & "n d"
         bool = CS.SetField(1, s007)
      End If
      bool006 = CS.GetField("006", 1, s006)
      If bool006 <> TRUE Then
         If Len(sAudn) = 0 Then
            sAudn = " "
         End If
         If sRecType = "a" Then
            Begin Dialog TypeChoices 170, 80, "Choose type of sound recording"
            OptionGroup .Choice
            OptionButton  24,  10, 140, 14, "&non-musical sound recording"
            OptionButton  24,  28, 140, 14, "&musical sound recording"
            OkButton        24, 50,  54, 16
            CancelButton   100, 50,  54, 16
            End Dialog
            Dim CaseOptions as TypeChoices
            On Error Resume Next
            Dialog CaseOptions 
            If Err = 102 Then Exit Sub
            Select Case CaseOptions.Choice
               Case 0
                  SoundRecType$ = "i"
               Case 1
                  SoundRecType$ = "j"
            End Select
            If SoundRecType$ = "i" Then
               Begin Dialog LTxtChoices 170, 325, "Choose type of literary text"
               OptionGroup .Choice
               OptionButton  24,  10, 140, 14, "&Autobiography"
               OptionButton  24,  28, 140, 14, "&Biography"
               OptionButton  24,  46, 140, 14, "&Drama"
               OptionButton  24,  64, 140, 14, "&Essays"
               OptionButton  24,  82, 140, 14, "&Fiction"
               OptionButton  24,  100, 140, 14, "&History"
               OptionButton  24,  118, 140, 14, "&Instruction (How to...)"
               OptionButton  24,  136, 140, 14, "&Language instruction"
               OptionButton  24,  154, 140, 14, "&Comedy"
               OptionButton  24,  172, 140, 14, "&Lectures, speaches"
               OptionButton  24,  190, 140, 14, "&Memoirs"
               OptionButton  24,  208, 140, 14, "&Folktales"
               OptionButton  24,  226, 140, 14, "&Poetry"
               OptionButton  24,  244, 140, 14, "&Sounds"
               OptionButton  24,  262, 140, 14, "&Interviews"
               OptionButton  24,  280, 140, 14, "&Other"
               OkButton        24, 300,  54, 16
               CancelButton   100, 300,  54, 16
               End Dialog
               Dim LTxtOptions as LTxtChoices
               On Error Resume Next
               Dialog LTxtOptions 
               If Err = 102 Then Exit Sub
               Select Case LTxtOptions.Choice
                  Case 0
                     LTxt$ = "a "
                  Case 1
                     LTxt$ = "b "
                  Case 2
                     LTxt$ = "dl"
                  Case 3
                     LTxt$ = "ez "
                  Case 4
                     LTxt$ = "f "
                  Case 5
                     LTxt$ = "h "
                  Case 6
                     LTxt$ = "i "
                  Case 7
                     LTxt$ = "j "
                  Case 8
                     LTxt$ = "k "
                  Case 9
                     LTxt$ = "l "
                  Case 10
                     LTxt$ = "m "
                  Case 11
                     LTxt$ = "o "
                  Case 12
                     LTxt$ = "p "
                  Case 13
                     LTxt$ = "s "
                  Case 14
                     LTxt$ = "tm"
                  Case 15
                     LTxt$ = "z "
               End Select
               s006 = "006  " & SoundRecType$ & "nnnn" & sAudn & "       " & LTxt$ &  " n "  
            ElseIf srec_type$ = "j" Then
               s006 = "006  " & SoundRecType$ & Chr(252) & Chr(252) & "nn" & sAudn & "          n "
            End If        
         ElseIf sRecType = "i" or sRecType = "j" Then
            bool = CS.GetFixedField("LTxt", sLTxt$)
            If InStr("a ,b , c ,gt, h ,i , j ,l ,m ,o ,kr,s ,tm,z ", sLTxt$) <> 0 Then
               form$ = "0"
            ElseIf sLTxt$ = "a " Or sLTxt$ = "m " Then
               bio$ = "a"
            ElseIf sLTxt$ = "b " Then
               bio$ = "b"
            ElseIf InStr("d ,ez,f ,p ", sLTxt$) <> 0 Then
               form$ = Left(sLTxt$, 1)
               bio$ = " " 
            Else
               bio$ = " "
               form$ = "0"
            End If
            s006 = "006  a    " & sAudn & "      000 " & form$ & bio$
         Else
            MsgBox "Type of bibliographic record not supported. Report to Tomasz immediately"
         End If
         bool = CS.SetField(1, s006)
      End If
   End If
   
   If sItemForm <> "o" Then
      n = 1
      nBool = CS.GetField("020", n, isbn$)
      Do While nBool = TRUE
         isbn$ = LCase(isbn$)
         If InStr(isbn$, "ebk") <> 0 Or InStr(isbn$, "ebook") <> 0 Or InStr(isbn$, "electronic") <> 0 _ 
          Or InStr(isbn$, "e-book") <> 0 Or InStr(isbn$, "e-isbn") <> 0 Or InStr(isbn$, "e-mhid") <> 0 _ 
          Or InStr(isbn$, "pdf") <> 0 Or InStr(isbn$, "epub") <> 0 Or InStr(isbn$, "e-mobi") <> 0 _
          Or InStr(isbn$, "html") <> 0 Or InStr(isbn$, "mobil") <> 0 Or InStr(isbn$, "el.") <> 0 _
          Or InStr(isbn$, "kindle") <> 0 Then
            'remove apostrophe in the beginning of the line below to display deleted isbns
            'MsgBox isbn$
            CS.DeleteField "020", n
         Else
            n = n + 1 
         End If
         nBool = CS.GetField("020", n, isbn$)
      Loop 
   End If
   
   'add field with initials
   s947 = "947  " & sInitials
   CS.SetField 1, s947
   
   CS.EndRecord

End Sub

