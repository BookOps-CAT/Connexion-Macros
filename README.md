# BookOps Cataloging Dept. Connexion Macros
A repository of the BookOps cataloging department Macros for OCLC Connexion client.

## Installation
### Easy Way
To install, copy the macro book (.mbk file in /macrobooks folder) to the "C:/Users/[your-username]/AppData/Roaming/OCLC/Connex/Macros/" directory.

Simple as that!
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
1. Bookops.mbk macrobook
	* [Utilities.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/Utilities.bas)
2. Brooklyn.mbk macrobook
	* [BPLCallNum.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/BPLCallNum.bas)
	* [ToExcel.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/BPLToExcel.bas)
3. NYPL.mbk macrobook
	* [NYPLCallNum.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/NYPLCallNum.bas)
	* [NYPLMusicCD.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/NYPLMusicCD.bas)
4. CAT.mbk macrobook
	* [BackupSettings.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/CATBackupSettings.bas)
	* [CATBibcoEval.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/CATBibcoEval.bas)
	* [UpdateExport.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/CATUpdateExport.bas)
    * [CATShareSave.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/CATShareSave.bas)
    * [CATUserDual.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/CATNewUserDual.bas)
    * [CATUserNYP.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/CATNewUserNYP.bas)
5. ACQ.mbk macrobook
	* [RLOrderRec.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/ACQRLOrderRec.bas)
6. NewMacros.mbk macrobook
	* [updater.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/master/source/newMacrosUpdater.bas)
## License
[MIT](https://opensource.org/licenses/MIT)
