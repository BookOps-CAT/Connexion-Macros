'MacroName:NYLP MusicCD
'MacroDescription:Creates call number for NYPL music CDs ; 
                  'call numbers can be insterted into displayed record or copied into clipboard for pasting into MidWest platform
'Macro created by: Tomasz Kalata, BookOps
'Last updated: February 05, 2015 (v. 1.1)
'Existing issues: lower screeen resolution may cause failure for drop-down menu to appear (most likely genre since it's rahter long
                 ' at this time the menues have hard coded lenght, in case of problems lower DropBoxCombo & ListBoxCombo
                 ' last value to 200 or less

Declare Sub Diacritics(sHeading)
Declare Function CutterArray()
Declare Sub CutterManipulation(sHeading)
Declare Sub InsertCallNumber(sField948, sInitials)

Sub Main

   Dim CS as Object
   Set CS = CreateObject("Connex.Client")
   If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 2 or CS.ItemType = 17 Then
      Dim sAudn$, sRecType$, sCutter$, sField948$, sRecNumber$
      Dim sGenre() As String
      Dim sOutput() As String
      
      Dim filenumber As Integer
      Dim sFileName As String
      Dim sDefaultInitials$
      Dim sInitials$
      
      CS.GetFixedField "Type", sRecType
      CS.GetFixedField "Audn", sAudn
      CS.QueryRecordStatus "NUM", sRecNumber
      sRecNumber = "-" & Right(sRecNumber, 4)
      
      If sRecType <> "j" Then
         MsgBox "INFO: Record NOT coded as musical sound recording. Verify fixed field coding. Exiting..."
         Goto Done
      Else
         sCutter = CutterArray
         ReDim sGenre(28)
            sGenre(0) = "BLUES"
            sGenre(1) = "CHAMBER"
            sGenre(2) = "CHILDREN"
            sGenre(3) = "CONCERTO"
            sGenre(4) = "COUNTRY"
            sGenre(5) = "ELECTRONIC"
            sGenre(6) = "GOSPEL"
            sGenre(7) = "HIPHOP"
            sGenre(8) = "HOLIDAY"
            sGenre(9) = "INSTRUMENTAL"
            sGenre(10) = "JAZZ"
            sGenre(11) = "LATIN-POP"
            sGenre(12) = "METAL"
            sGenre(13) = "MUSICAL"
            sGenre(14) = "OPERA"
            sGenre(15) = "ORATORIO"
            sGenre(16) = "ORCHESTRAL"
            sGenre(17) = "POP"
            sGenre(18) = "R&B/SOUL"
            sGenre(19) = "REGGAE"
            sGenre(20) = "REGGAETON"
            sGenre(21) = "ROCK"
            sGenre(22) = "SALSA"
            sGenre(23) = "SOUNDTRACK"
            sGenre(24) = "SYMPHONY"
            sGenre(25) = "US-FOLK"
            sGenre(26) = "VIDEO-GAME"
            sGenre(27) = "VOCAL"
            sGenre(28) = "WORLD"
         ReDim sOutput(1)
            sOutput(0) = "add to record"
            sOutput(1) = "copy to clipboard"
            
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
            Print #filenumber, "xxx"
            Close #filenumber
            sDefaultInitials = "xxx"
         End If
            
         Begin Dialog UserDialog 170, 110, "Select genre, cutter and output"
            Text  8, 5, 80, 8, "Genre"
            DropComboBox  8, 20, 80, 250, sGenre(), .sGenre
            Text 100, 5, 30, 30, "Initials"
            Textbox 100, 20, 20, 15, .sInitials
            Text  8, 35, 80, 8, "Cutter"
            DropListBox  8, 50, 80, 250, sCutter, .sCutter
            Text  8, 65, 80, 8, "Output"
            ListBox 8, 80, 80, 30, sOutput(), .sOutput
            OKButton  100, 65, 54, 14
            CancelButton  100, 83, 54, 14
         End Dialog
         Dim dMusic as UserDialog
         'populate INITIALS box with default value
         dMusic.sInitials = sDefaultInitials
         On Error Resume Next
         Dialog dMusic
         If Err = 102 Then Exit Sub
      End If
      'relates cataloger's cutter choice (numeral in dMusic.sCutter) with appropriate string portion of sCutter$
      n = dMusic.sCutter + 1
      temp$ = sCutter
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
      sCutter = lt$
      sCutter = Mid(sCutter, 6, 8)
      'assembles and inserts call number
      If dMusic.sGenre = "CHILDREN" Then
         sField948 = "948  " & Chr(223) & "p J " & Chr(223) & "f " & "CD " & Chr(223) & "a " & dMusic.sGenre & " " & Chr(223) & "c " & sCutter & sRecNumber
      Else
         sField948 = "948  " & Chr(223) & "f " & "CD " & Chr(223) & "a " & dMusic.sGenre & " " & Chr(223) & "c " & sCutter & sRecNumber
      End If

      If dMusic.sOutput = 1 Then
         x = Len(sField948)
         sField948 = Left(sField948, x - 5)
         If dMusic.sGenre = "CHILDREN" Then
            sField948 = Mid(sField948, 6)
         Else
            sField948 = Mid(sField948, 12)
         End If
         
         Do While InStr(sField948, Chr(223)) <> 0
            place = InStr(sField948, Chr(223))
            lt$ = Left(sField948, place - 1)
            rt$ = Mid(sField948, place + 3)
            sField948 = lt$ + rt$
         Loop
         Clipboard.Clear
         Clipboard.SetText sField948
      Else
         'populate default initials
         sInitials = dMusic.sInitials
         filenumber = FreeFile
         Open sFileName For Output As filenumber
         Print #filenumber, sInitials
         Close #filenumber
         
         'insert call  number
         Call InsertCallNumber(sField948, sInitials)
      End If
      
      If InStr("abcj", sAudn) <> 0 And sAudn <> "" And dMusic.sGenre <> "CHILDREN" Then
         MsgBox "Genre selection different than CHILDREN while Audience fixed field coded as juvenile. Please verify your selection." 
      End If
      
   Else
      MsgBox "INFO: A bibliographic record must be displayed in order to use this macro."
   End If
   
Done: 
End Sub

Function CutterArray()
   
   Dim CS as Object
   Set CS = CreateObject("Connex.Client")
   Dim sCutterOpt$, sHeading$
   Dim i As Integer
   Dim sNonLatMessage$

   sNonLatMessage = "Data conta"
   'finds cutter options
   bool100 = CS.GetField("100", 1, sHeading)
   If bool100 = TRUE Then
      If Left(sHeading, 10) = sNonLatMessage Then
         bool100 = CS.GetField("100", 2, sHeading)
         If bool100 = TRUE Then
            Call CutterManipulation(sHeading)
            sCutterOpt = sHeading
         End If
      Else
         Call CutterManipulation(sHeading)
         sCutterOpt = sHeading
      End If
   End If 
   bool110 = CS.GetField("110", 1, sHeading)
   If bool110 = TRUE Then
      If Left(sHeading, 10) = sNonLatMessage Then
         bool110 = CS.GetField("110", 2, sHeading)
         If bool110 = TRUE Then
            Call CutterManipulation(sHeading)
            sCutterOpt = sHeading
         End If
      Else
         Call CutterManipulation(sHeading)
         sCutterOpt = sHeading
      End If
   End If 
   bool245 = CS.GetField("245", 1, sHeading)
   If bool245 = TRUE Then
      If Left(sHeading, 10) = sNonLatMessage Then
         bool245 = CS.GetField("245", 2, sHeading)
         If bool245 = TRUE Then
            Call CutterManipulation(sHeading)
            If Len(sCutterOpt) = 0 Then
               sCutterOpt = sHeading
            Else
               sCutterOpt = sCutterOpt & Chr(9) & sHeading
            End If
         End If
      Else
         Call CutterManipulation(sHeading)
         If Len(sCutterOpt) = 0 Then
            sCutterOpt = sHeading
         Else
            sCutterOpt = sCutterOpt & Chr(9) & sHeading
         End If
      End If
   End If 
   i = 1
   Do While CS.GetField("700", i, sHeading)
      If Left(sHeading, 10) <> sNonLatMessage Then
         Call CutterManipulation(sHeading)
         sCutterOpt = sCutterOpt & Chr(9) & sHeading
      End If
      i = i + 1
   Loop
   i = 1
   Do While CS.GetField("710", i, sHeading)
      If Left(sHeading, 10) <> sNonLatMessage Then
         Call CutterManipulation(sHeading)
         sCutterOpt = sCutterOpt & Chr(9) & sHeading
      End If
      i = i + 1
   Loop

CutterArray = sCutterOpt
End Function

Sub Diacritics(sHeading)
'removes diacritic marks and other unwanted characters from a string

   Dim CS as Object
   Set CS = CreateObject("Connex.Client")
   Dim i as Integer

   i = 1
   While i <= Len(sHeading)
      CheckChar = Mid(sHeading, i, 1)
      Select Case CheckChar
         Case "à", "ä", "â", "À", "Á", "á", "Â", "Ä", "Å", "å"
            sHeading = Mid(sHeading, 1, i - 1) & Mid(sHeading, i + 1, Len(sHeading) - i)
         Case "ç", "Ç"
            sHeading = Mid(sHeading, 1, i - 1) & Mid(sHeading, i + 1, Len(sHeading) - i)
         Case "Ð", "ð"
            sHeading = Mid(sHeading, 1, i - 1) & Mid(sHeading, i + 1, Len(sHeading) - i)     
         Case "é", "É", "è", "È", "Ê", "ê", "Ë", "ë"
            sHeading = Mid(sHeading, 1, i - 1) & Mid(sHeading, i + 1, Len(sHeading) - i)
         Case "í", "Ì", "ì", "Í", "Î", "î", "Ï", "ï", Chr(230)
            sHeading = Mid(sHeading, 1, i - 1) & Mid(sHeading, i + 1, Len(sHeading) - i)   
         Case "ñ", "Ñ"
            sHeading = Mid(sHeading, 1, i - 1) & Mid(sHeading, i + 1, Len(sHeading) - i)   
         Case "ó", "ö", "Ö", "Ò", "ò", "Ó", "Ô", "ô", "Ö", Chr(227)
            sHeading = Mid(sHeading, 1, i - 1) & Mid(sHeading, i + 1, Len(sHeading) - i)   
         Case "ü", "ú", "Ü", "Ù", "ù", "Ú", "Û", "û", "Ü"
            sHeading = Mid(sHeading, 1, i - 1) & Mid(sHeading, i + 1, Len(sHeading) - i)
         Case Chr(181), Chr(165)
            sHeading = Mid(sHeading, 1, i - 1) & "ae" & Mid(sHeading, i + 1, Len(sHeading) - i)
         Case Chr(182), Chr(166)
            sHeading = Mid(sHeading, 1, i - 1) & "oe" & Mid(sHeading, i + 1, Len(sHeading) - i)
         Case Chr(161), Chr(177)
            sHeading = Mid(sHeading, 1, i - 1) & "l" & Mid(sHeading, i + 1, Len(sHeading) - i) 
         Case Chr(178), Chr(162)
            sHeading = Mid(sHeading, 1, i - 1) & "o" & Mid(sHeading, i + 1, Len(sHeading) - i)
         Case Chr(186), Chr(163), Chr(179), "Ð"
            sHeading = Mid(sHeading, 1, i - 1) & "d" & Mid(sHeading, i + 1, Len(sHeading) - i)
         Case "-"
            sHeading = Mid(sHeading, 1, i - 1) & " " & Mid(sHeading, i + 1, Len(sHeading) - i)
         'characters to be removed completely
         Case Chr(176), Chr(174), Chr(167)
            sHeading = Mid(sHeading, 1, i - 1) & Mid(sHeading, i + 1, Len(sHeading) - i) 
      End Select
      i = i + 1   
   Wend
End Sub

Sub CutterManipulation(sHeading)

   Dim CS as Object
   Set CS = CreateObject("Connex.Client")
   Dim sEntryType$
   
   Call Diacritics(sHeading)
   If InStr(sHeading, Chr(223) & "i") <> 0 Then
      lt$ = Left(sHeading, 5)
      place = InStr(sHeading, Chr(223) & "a")
      rt$ = LTrim(Mid(sHeading, place + 2))
      sHeading = lt$ + rt$ 
   End If
   If InStr(sHeading, Chr(223)) <> 0 Then
      place = InStr(sHeading, Chr(223))
      sHeading = Left(sHeading, place - 1)
      sHeading = RTrim(sHeading)
   End If
   If InStr(sHeading, "(") <> 0 Then
      place = InStr(sHeading, "(")
      sHeading = Left(sHeading, place - 1)
      sHeading = RTrim(sHeading)
   End If
   sEntryType = Left(sHeading,3)
   If sEntryType = "100" Or sEntryType = "700" Then
      If InStr(sHeading, ",") <> 0 Then
         place = InStr(sHeading, ",")
         lt$ = RTrim(Left(sHeading, place - 1))
         rt$ = LTrim(Mid(sHeading, place + 1))
         sHeading = lt$ & "," & rt$
      End If
   ElseIf sEntryType = "245" Then
      Indicator = Mid(sHeading, 5, 1)
      If Indicator <> 0 Then
         lt$ = Left(sHeading, 5)
         rt$ = Mid(sHeading, 6 + Indicator)
         sHeading = lt$ + rt$
      End If
      If InStr(sHeading, ":") <> 0 Then
         place = InStr(sHeading, ":")
         lt$ = RTrim(Left(sHeading, place - 1))
         rt$ = LTrim(Mid(sHeading, place + 1))
         sHeading = lt$ + rt$
      End If
      If InStr(sHeading, ";") <> 0 Then
         place = InStr(sHeading, ";")
         lt$ = RTrim(Left(sHeading, place - 1))
         rt$ = LTrim(Mid(sHeading, place + 1))
         sHeading = lt$ + rt$
      End If
      If InStr(sHeading, "/") <> 0 Then
         place = InStr(sHeading, "/")
         lt$ = RTrim(Left(sHeading, place - 1))
         rt$ = LTrim(Mid(sHeading, place + 1))
         sHeading = lt$ + rt$
      End If
   End If
   lt$ = Left(sHeading, 3)
   rt$ = RTrim(Mid(sHeading, 6, 8))
   If Right(rt$, 1) = "," Or Right(rt$, 1) = "." Then
      rt$ = RTrim(Left(rt$, 8 - 1))
   End If
   sHeading = lt$ & ": " & rt$
   sHeading = UCase(sHeading)

End Sub

Sub InsertCallNumber(sField948, sInitials)
   Dim CS as Object
   Set CS = CreateObject("Connex.Client")
   Dim s901$
  
   CS.SetField 1, sField948
   CS.SetField 1, "945  .o"
   CS.SetField 1, "946  m"
   CS.SetField 1, "949  *b2=y;recs=oclcgw;"
   s901 = "901  " & sInitials & " " & Chr(223) & "b CATBL"
   CS.SetField 1, s901

End Sub
