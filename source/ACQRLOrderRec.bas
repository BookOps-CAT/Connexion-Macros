'MacroName:RLOrderRec
'MacroDescription: used by NYPL Research to create & exports bibs and order records to Sierra;
'                  macro defines controlled values for particular order record fields, displays them in
'                  a window where they can be chosen, creates approriate strings in fields 949,960, & 961 in
'                  displayed on screen record based on a user selection, and exports such record to Sierra;
'                  fields 949, 960, & 961 specify load table and details of the order record (see III documentation);
'                  macro creates and modifies a text file "acq_data.txt" in the main Connexion folder:
'                  "C:\Program Files (x86)\OCLC\Connexion\Program\"    to store staff initials and 4 other values that
'                  populate default values for the next run of the macro, allowing user to apply the same template to
'                  consecutive order records
'Macro created by: Tomasz Kalata, BookOps
'Latest update: December 1, 2020 ; v. 1.9
'
'v1.9 details (12-01-2020): YBP bug fixed
'v1.8 details (09-31-2020): added new vendor - NESTO
'v1.7 details (07-22-2020): data persisted in Connexion AppData folder specific to each user
'v1.6 details (03-10-2020): WNB vendor code added
'v1.5 details: Ugly Duckling Press vendor code (udp) added
'
'
'To add a location to the list go the section marked XXXXXXXX and
'add a new line after the last line "sLocation([next number])" then
'mofify number in parenthesis in section's first line: from "ReDim sLocation([number])" to "ReDim sLocation([next number])
'
'To add a vendor to the list go to the section marked ZZZZZZZZ and
'add a new line after the last line: "sVendor([next number]) = "[vendor code]" then
'modify number in parenthesis in section's first line: from "ReDim sVendor([number]") to "ReDim sVendor([next number])

Option Explicit

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")
   If CS.ItemType = 0 or CS.ItemType = 1 or CS.ItemType = 2 or CS.ItemType = 17 Then
      Dim s949$, s960$, s961$, sForm$, sPrice$, sTodaysDate$, sRDate$, sVendorArr$, sOrderType, sOrderCode3, _
         lt$, rt$, temp$, sType$
      Dim sFund() As String
      Dim sLocation() As String
      Dim sOCode3() As String
      Dim sOType() As String
      Dim sVendor() As String
      Dim i%, Copies%, place%, filenumber%
      Dim sFileName As String
      'declare variables for default values to populate window options
      Dim sDefaultValues$, sDefInitials$, sDefFund$
      Dim DefLocation%, DefOCode3%, DefOType%, DefVendor%
      Dim checkchar$, pricepattern, fundpattern, copiespattern, retvalue%

      'find current date and format it
      sTodaysDate = Date
      Do While InStr(sTodaysDate, "/")
         place = InStr(sTodaysDate, "/")
         lt$ = Left(sTodaysDate, place - 1)
         rt$ = Mid(sTodaysDate, place + 1)
         sTodaysDate = lt$ & "-" & rt$
      Loop
      'find record type to populate order record's Form field
      CS.GetFixedField "Type", sType

      ReDim sOCode3(1)
         sOCode3(0) = "DOMESTIC"
         sOCode3(1) = "FOREIGN"
      ReDim sOType(1)
         sOType(0) = "APPROVAL"
         sOType(1) = "FIRM ORDER"

'XXXXXXXXXX
      ReDim sLocation(18)
         sLocation(0) = " "
         sLocation(1) = "MAB"
         sLocation(2) = "MAE"
         sLocation(3) = "MAF"
         sLocation(4) = "MAG"
         sLocation(5) = "MAL"
         sLocation(6) = "MAO"
         sLocation(7) = "MAP"
         sLocation(8) = "MAR"
         sLocation(9) = "MAS"
         sLocation(10) = "MAQ"
         sLocation(11) = "MAR"
         sLocation(12) = "MAU"
         sLocation(13) = "MYD"
         sLocation(14) = "MYH"
         sLocation(15) = "MYM"
         sLocation(16) = "MYT"
         sLocation(17) = "SC"
         sLocation(18) = "SLR"
'ZZZZZZZZZZ
      ReDim sVendor(39)
         sVendor(0) = " "
         sVendor(1) = "4398"
         sVendor(2) = "ALIBR"
         sVendor(3) = "AMERI"
         sVendor(4) = "BARL"
         sVendor(5) = "BCH"
         sVendor(6) = "BLAN"
         sVendor(7) = "BRN"
         sVendor(8) = "CAMB"
         sVendor(9) = "CROOT"
         sVendor(10) = "CSL"
         sVendor(11) = "DKA"
         sVendor(12) = "ENGH"
         sVendor(13) = "EUR"
         sVendor(14) = "EVP"
         sVendor(15) = "HERIT"
         sVendor(16) = "HOG"
         sVendor(17) = "HRR"
         sVendor(18) = "INC"
         sVendor(19) = "ITR"
         sVendor(20) = "JERUS"
         sVendor(21) = "KAR"
         sVendor(22) = "KEN"
         sVendor(23) = "LAT"
         sVendor(24) = "LEXI"
         sVendor(25) = "LNR"
         sVendor(26) = "MDJ"
         sVendor(27) = "MEXIC"
         sVendor(28) = "MGENS"
         sVendor(29) = "NESTO"
         sVendor(30) = "NORTE"
         sVendor(31) = "PRAGE"
         sVendor(32) = "PVL"
         sVendor(33) = "RETTA"
         sVendor(34) = "SBD"
         sVendor(35) = "SUR"
         sVendor(36) = "TROP"
         sVendor(37) = "WNB"
         sVendor(38) = "YBP"
         sVendor(39) = "UDP"

      'read default data from text file stored in macro folder
      sFileName = "acq_data.txt"
      sFileName = Mid(Environ(2), 9) + "\OCLC\Connex\Profiles\acq_data.txt"
      If Dir$ (sFileName) <> "" Then
         filenumber = FreeFile
         Open sFileName For Input As filenumber
         Line Input #filenumber, sDefaultValues
         Close #filenumber
         'change retrieved values from stings to integers if applicable and assign them to variables
         DefLocation = Val(GetField(sDefaultValues, 1, ":"))
         DefOCode3 = Val(GetField(sDefaultValues, 2, ":"))
         sDefFund = GetField(sDefaultValues, 3, ":")
         DefOType = Val(GetField(sDefaultValues, 4, ":"))
         DefVendor = Val(GetField(sDefaultValues, 5, ":"))
         sDefInitials = GetField(sDefaultValues, 6, ":")
      Else
         'create a file with blank values if non present in the macro folder
         filenumber = FreeFile
         Open sFileName For Output As filenumber
         Print #filenumber, "0:0: :0:0: "
         Close #filenumber
         DefLocation = 0
         DefOCode3 = 0
         sDefFund = " "
         DefOType = 0
         DefVendor = 0
         sDefInitials = " "
      End If

'Dialog box presenting order record options
MenuWindow:
      Begin Dialog MainWindow 247, 108, "NYPL Research Order Record"
         'left panel
         Text 15, 12, 55, 12, "LOCATION:"
         DropListBox 55, 10, 55, 170, sLocation, .sLocation
         Text 15, 29, 55, 12, "VENDOR:"
         DropListBox 55, 27, 55, 170, sVendor, .sVendor
         Text 15, 46, 55, 12, "FUND:"
         TextBox 55, 44, 55, 12, .sFund
         Text 15, 63, 55, 12, "PRICE:"
         TextBox 55, 61, 55, 12, .sPrice
         'right panel
         Text 120, 12, 55, 12, "ORDER TYPE:"
         DropListBox 175, 10, 55, 40, sOType, .sOType
         Text 120, 29, 55, 50, "O CODE 3:"
         DropListBox 175, 27, 55, 40, sOCode3, .sOCode3
         Text 120, 46, 55, 12, "COPIES:"
         TextBox 175, 44, 55, 12, .sCopies
         Text 120, 61, 55, 12, "INITIALS:"
         TextBox 175, 61, 55, 12, .sInitials
         'bottom
         OkButton        55, 80,  55, 16
         CancelButton   120, 80,  55, 16
      End Dialog
      Dim dOrderRec as MainWindow
      'populate window options with values selected for createded previously order record
      dOrderRec.sLocation = DefLocation
      dOrderRec.sCopies = "1"
      dOrderRec.sOCode3 = DefOCode3
      dOrderRec.sFund = sDefFund
      dOrderRec.sOType = DefOType
      dOrderREc.sVendor = DefVendor
      dOrderRec.sInitials = sDefInitials
      'handle erors and activate order record window
      On Error Resume Next
      Dialog dOrderRec
      If Err = 102 Then Exit Sub

      'match selected option with Sierra's value for Form, Order Type, and Order Code 3

      If sType = "a" Then
         sForm = "b"
      Elseif sType = "c" Then
         sForm = "c"
      Elseif sType = "j" Then
         sForm = "w"
      Elseif sType = "g" Then
         sForm = "v"
      Elseif sType = "e" Then
         sForm = "u"
      Else
         MsgBox "Could not match record data with NYPL order record Form code. Please correct manually."
      End If
      If dOrderRec.sOType = 0 Then
         sOrderType = "a"
         sRDate = sTodaysDate
      ElseIf dOrderRec.sOType = 1 Then
         sOrderType = "f"
      End If
      If dOrderRec.sOCode3 = 0 Then
         sOrderCode3 = "d"
      ElseIf dOrderRec.sOCode3 = 1 Then
         sOrderCode3 = "f"
      End If

      'remove any white spaces from user entered values
      dOrderRec.sFund = Trim(dOrderRec.sFund)
      dOrderRec.sPrice = Trim(dOrderRec.sPrice)
      dOrderRec.sCopies = Trim(dOrderRec.sCopies)
      dOrderRec.sInitials = Trim(dOrderRec.sInitials)

      'user input error handling
      If dOrderRec.sLocation = 0 Then
         MsgBox "Forgot about the location! Let's try again..."
         Goto MenuWindow
      End If
      If dOrderRec.sVendor = 0 Then
         MsgBox "Forgot about the vendor! Let's try again..."
         Goto MenuWindow
      End If
      If dOrderRec.sFund = "" Then
         MsgBox "Forgot to indicate the fund! Let's try again..."
         Goto MenuWindow
      Else
         fundpattern = "#####[a-z]*[a-z]"
         retvalue = dOrderRec.sFund LIKE fundpattern
         If retvalue = -1 Then
            If Len(dOrderRec.sFund) > 15 Then      'can not be longer than 15 characters
               MsgBox "It looks fund code is too long. Let's try again..."
               Goto MenuWindow
            Else
               'define illegal characters that fund code should not include
               i = 1
               While i <= Len(dOrderRec.sFund)
                  checkchar = Mid(dOrderRec.sFund, i, 1)
                  Select Case checkchar
                     Case " ",  ".",  ",",  "/", ";", ":", "[", "]", "\", "-", "=", "`", "_", "+", "*", "'", "!", _
                              "@", "#", "$", "%", "^", "&", "*", "(", ")"
                        MsgBox "Fund code includes illegal character. Before I report it ... Let's try again..."
                        GoTo MenuWindow
                  End Select
                     i = i + 1
               Wend
            End If
         ElseIf retvalue = 0 Then
            MsgBox "It looks fund code is incorrect. Let's try again..."
            Goto MenuWindow
         End If
      End If
      If dOrderRec.sPrice = "" Then
         MsgBox "Forgot about price! Let's try again..."
         Goto MenuWindow
      Else
         pricepattern = "#.##"
         retvalue = dOrderRec.sPrice LIKE pricepattern
         If retvalue = 0 Then
            pricepattern = "##.##"
            retvalue = dOrderRec.sPrice LIKE pricepattern
            If retvalue = 0 Then
               pricepattern = "###.##"
               retvalue = dOrderRec.sPrice LIKE pricepattern
               If retvalue = 0 Then
                  MsgBox "Mistyped price. Let's try again..."
                  Goto MenuWindow
               End If
            End If
         End If
      End If
      If dOrderRec.sCopies = "" Then
         MsgBox "Must have at least 1 copy. Let's try again..."
         Goto MenuWindow
      Else
         copiespattern = "#"
         retvalue = dOrderRec.sCopies LIKE copiespattern
         If retvalue = 0 Then
            copiespattern = "##"
            retvalue = dOrderRec.sCopies LIKE copiespattern
            If retvalue = 0 Then
               MsgBox "Number of copies is incorrect. Let's try again."
               Goto MenuWindow
            End If
         End If
      End If

      'create strings of 949 (load table and authorization for load), 960 (order fixed fields), & 961 (order variable fields)
      s949 = "949  *recs=researchord;ins=ckgriffin;"
      s960 = "960  " & _
            Chr(223) & "a p " & _
            Chr(223) & "t " & LCase(sLocation(dOrderRec.sLocation)) & " " & _
            Chr(223) & "d r " & _
            Chr(223) & "g " & sForm & " " & _
            Chr(223) & "e " & sOrderCode3 & " " & _
            Chr(223) & "o " & dOrderRec.sCopies & " " & _
            Chr(223) & "s " & dOrderRec.sPrice & " " & _
            Chr(223) & "u " & dOrderRec.sFund & " " & _
            Chr(223) & "i " & sOrderType & " " & _
            Chr(223) & "r " & sRDate & " " & _
            Chr(223) & "m 1 " & _
            Chr(223) & "v " & LCase(sVendor(dOrderRec.sVendor))
      s961 = "961  " & Chr(223) & "d " & dOrderRec.sInitials

      'insert 949,960, & 961 strings into the displayed record
      CS.SetField 1, s949
      CS.SetField 1, s960
      CS.SetField 1, s961

      'populate default data in file in Connexion Macro folder; each element separated with colon
      filenumber = FreeFile
      Open sFileName For Output As filenumber
      sDefaultValues = Trim(Str(dOrderRec.sLocation)) & ":" & _
                       Trim(Str(dOrderRec.sOCode3)) & ":" & _
                       Trim(dOrderRec.sFund) & ":" & _
                       Trim(Str(dOrderRec.sOType)) & ":" & _
                       Trim(Str(dOrderRec.sVendor)) & ":" & _
                       Trim(dOrderRec.sInitials)
      Print #filenumber, sDefaultValues
      Close #filenumber

      'export the record
      CS.Export

      'determine if price in dollars or euro
      If InStr("CSL,HRR,BLAN", sVendor(dOrderRec.sVendor)) <> 0 Then
         MsgBox "Please modify price in Sierra. This vendor uses euro."
      End If

   Else
      MsgBox "INFO: A bibliographic record must be displayed in order to use this macro."
      Goto Done
   End If
Done:
End Sub
