'MacroName:ToExcel
'MacroDescription: Writes ISBN, title, author and call number of displayed bib. record to Excel spreadsheet
'Last update: March 15, 2021
'version: v.1.2 : Windows 10 compatibility fix
   
   
Sub Main

  Dim CS As Object
  Set CS = CreateObject("Connex.Client")

  'Set the output Excel file
  sFileName = "C:" & Mid(Environ(11), 10) + "\Desktop\LargePrint.xls"


  If Dir(sFileName) = "" Then
    MsgBox "File not found: " & Chr(10) & Chr(9) & sFileName & Chr(10) & Chr(10) & "Exiting..."
    Exit Sub
  End If

  'Get the author, if there is one
  bool = CS.GetField("100", 1, sAuthor)
  If bool = FALSE Then
    bool = CS.GetField("110", 1, sAuthor)
    If bool = FALSE Then
      bool = CS.GetField("111", 1, sAuthor)
    End If
  End If
  If bool <> FALSE Then
    sAuthor = Mid(sAuthor, 6)
  Else
    sAuthor = ""
  End If
  'Strip out subfield codes
  Do While InStr(sAuthor, Chr(223))
    place = InStr(sAuthor, Chr(223))
    sAuthor = RTrim(Left(sAuthor, place - 1)) & " " & LTrim(Mid(sAuthor, place + 2))
  Loop

  'Get the title
  bool = CS.GetField("245", 1, sTitle)
  sTitle = Mid(sTitle, 6)
  'Remove the statement of responsibility, if there is one
  If InStr(sTitle, Chr(223) & "c") Then
    place = InStr(sTitle, Chr(223) & "c")
    sTitle = RTrim(Left(sTitle, place - 1))
    If Right(sTitle, 1) = "/" Then sTitle = RTrim(Left(sTitle, Len(sTitle) - 1))
  End If
  'Strip out subfield codes
  Do While InStr(sTitle, Chr(223))
    place = InStr(sTitle, Chr(223))
    sTitle = RTrim(Left(sTitle, place - 1)) & " " & LTrim(Mid(sTitle, place + 2))
  Loop

  'Get the call number
  bool = CS.GetField("099", 1, sCallnumber)
  sCallnumber = Mid(sCallnumber, 6)
  'Strip out subfield codes
  Do While InStr(sCallnumber, Chr(223))
    place = InStr(sCallnumber, Chr(223))
    sCallnumber = RTrim(Left(sCallnumber, place - 1)) & " " & LTrim(Mid(sCallnumber, place + 2))
  Loop


  'Get the ISBN
  bool = CS.GetField("020", 1, sISBN)
  sISBN = Mid(sISBN, 6)
  'Remove the price if there is one
  If InStr(sISBN, Chr(223) & "c") Then
    place = InStr(sISBN, Chr(223) & "c")
    sISBN = RTrim(Left(sISBN, place - 1))
    If Right(sISBN, 1) = "$" Then sISBN = RTrim(Left(sISBN, Len(sISBN) - 1))
  End If
  'Strip out subfield codes
  Do While InStr(sISBN, Chr(223))
    place = InStr(sISBN, Chr(223))
    sISBN = RTrim(Left(sISBN, place - 1)) & " " & LTrim(Mid(sISBN, place + 2))
  Loop


  'Open a new Excel session  
  Dim XL As Object
  Set XL = CreateObject("Excel.Application")
  
  'Display the Excel window
  XL.Visible = TRUE

  'Open the Excel file
  XL.Workbooks.Open sFileName  

  'Find the next empty row
  Dim xlRange as Object
  Set xlRange = XL.Sheets(1).UsedRange
  If xlRange.Rows.Count = 1 And Trim(XL.Sheets(1).Range("B1").Value) = "" Then
    'If the file is empty, set up the column headers
    nNextRow = xlRange.Rows.Count
    XL.Sheets(1).Range("A" & CStr(nNextRow)).Value = "ISBN"
    XL.Sheets(1).Range("B" & CStr(nNextRow)).Value = "Title"
    XL.Sheets(1).Range("C" & CStr(nNextRow)).Value = "Author"
    XL.Sheets(1).Range("D" & CStr(nNextRow)).Value = "Call #"

    
  End If
  nNextRow = xlRange.Rows.Count + 1

  'Write the data into the Excel file, columns A-E
  XL.Sheets(1).Range("A" & CStr(nNextRow)).Value = sISBN
  XL.Sheets(1).Range("B" & CStr(nNextRow)).Value = sTitle
  XL.Sheets(1).Range("C" & CStr(nNextRow)).Value = sAuthor
  XL.Sheets(1).Range("D" & CStr(nNextRow)).Value = sCallnumber
  

  'Save and close the file  
  XL.Workbooks(1).Close (TRUE)

  'Close the new Excel session
  XL.Quit
End Sub

