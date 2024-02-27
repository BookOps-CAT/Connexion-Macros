# BookOps Cataloging Dept. Connexion Macros
A repository of BookOps Cataloging Department macros for the OCLC Connexion client.

## Installation
### Easy Way
To install, copy the macro book (.mbk file in /macrobooks folder) to the `C:/Users/[your-username]/AppData/Roaming/OCLC/Connex/Macros/` directory.

By default library computers hide the `AppData` directory. If you do not see `AppData` folder in your profile directory, follow the instruction below:
1. Open the Windows Folder Explorer application
2. Navigate to your profile folder: `C:/Users/[your-username]`
3. Go to the `View` tab in the toolbar
4. Select `Hidden items` option
5. Verify `AppData` is visible now

![folder-options](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/main/folder-options.png)

### Hard Way (Copy and Paste from Source)
Alternatively, copy the code for the individual macro directly from the .BAS file (/source directory) using the macro editor in Connexion.

1. Tools > Macros > Manage (CTRL-ALT-SHIFT-G)
2. Select a local macro book
3. Click the New Macro button
4. Enter in a description
5. Click OK
6. Enter in a name for the macro
7. Click the Edit button to open the OCLC Connexion Macro Editor and Debugger.
8. Use the mouse to select all the text
9. Paste the text copied from the source into the Editor
10. Click the Check icon in the toolbar
11. Click the Save icon in the toolbar
12. After the macro is saved then it can be mapped to a user tool, key, or key chord as usual (optional).
Done!

## Macros
1. [Bookops.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/macrobooks/Bookops.mbk)
	* [Utilities.bas v1.0.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/Utilities.bas)
2. [Brooklyn.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/macrobooks/Brooklyn.mbk)
	* [BPLCallNum.bas v3.4.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/BPLCallNum.bas)
	* [ToExcel.bas v1.2](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/BPLToExcel.bas)
3. [NYPL.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/macrobooks/NYPL.mbk)
	* [NYPLCallNum.bas v3.2.1](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/NYPLCallNum.bas)
	* [NYPLMusicCD.bas v1.5](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/NYPLMusicCD.bas)
4. [CAT.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/macrobooks/CAT.mbk)
	* [BackupSettings.bas v1.0.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/CATBackupSettings.bas)
	* [CATBibcoEval.bas v1.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/CATBibcoEval.bas)
	* [UpdateExport.bas v1.8](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/CATUpdateExport.bas)
    * [CATShareSave.bas v1.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/CATShareSave.bas)
    * [CATUserDual.bas v1.0.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/CATNewUserDual.bas)
    * [CATUserNYP.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/CATNewUserNYP.bas)
5. [ACQ.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/macrobooks/ACQ.mbk)
	* [RLOrderRec.bas v3.2](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/ACQRLOrderRec.bas)
6. [NewMacros.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/macrobooks/newMacros.mbk)
	* [updater.bas v2.1](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/newMacrosUpdater.bas)
## License
[MIT](https://opensource.org/licenses/MIT)
