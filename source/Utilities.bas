'MacroName:Utilities v1.0.0
'MacroDescription:Tools for creation of default settings, backups, and updating other macros.
'Macro created by: Tomasz Kalata, BookOps
'Last updated: October 4, 2021

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")
   
   Begin Dialog UtilDialog 180, 165, "Utilities"
   
      OptionGroup .Macro
      OptionButton 35, 10, 140, 20, "Update Macros"
      OptionButton 35, 30, 140, 20, "Backup Your Settings"
      OptionButton 35, 50, 140, 20, "Restore Your Settings"
      OptionButton 35, 70, 140, 20, "Create New Dual (NYP-BPL) User"
      OptionButton 35, 90, 140, 20, "Create New NYP User"
   
      OKButton 50, 120, 80, 15
      CancelButton 50, 140, 80, 15
   
   End Dialog
   
   Dim dUtil As UtilDialog
   
   On Error Resume Next
   Dialog dUtil
   If Err = 102 Then Exit Sub
   
   Select Case dUtil.Macro
      Case 0
         BOOL = CS.RunMacro("newMacros.mbk!Updater")
      Case 1
         BOOL = CS.RunMacro("CAT.mbk!BackupSettings")
      Case 2
         BOOL = CS.RunMacro("CAT.mbk!RestoreSettings")
      Case 3
         BOOL = CS.RunMacro("CAT.mbk!NewUserDual")
      Case 4
         BOOL = CS.RunMacro("CAT.mbk!NewUserNYP")
   End Select

End Sub