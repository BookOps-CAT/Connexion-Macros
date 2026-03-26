'MacroName:Utilities v1.1.0
'MacroDescription:Tools for creation of default settings, backups, and updating other macros.
'Last updated: March 26, 2026

Sub Main

   Dim CS As Object
   Set CS  = GetObject(,"Connex.Client")
   
   Begin Dialog UtilDialog 180, 185, "Utilities"
   
      OptionGroup .Macro
      OptionButton 35, 10, 140, 20, "Update Macros"
      OptionButton 35, 30, 140, 20, "Backup Your Settings"
      OptionButton 35, 50, 140, 20, "Restore Your Settings"
      OptionButton 35, 70, 140, 20, "Create CAT New Dual (NYP-BPL) User"
      OptionButton 35, 90, 140, 20, "Create CAT New NYP User"
      OPtionButton 35, 110, 140, 20, "Create New ACQ NYP User"
   
      OKButton 50, 140, 80, 15
      CancelButton 50, 160, 80, 15
   
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
         BOOL = CS.RunMacro("CAT.mbk!NewUserCATDual")
      Case 4
         BOOL = CS.RunMacro("CAT.mbk!NewUserCATNYP")
      Case 5
         BOOL = CS.RunMacro("CAT.mbk!NewUserACQNYP")
   End Select

End Sub
