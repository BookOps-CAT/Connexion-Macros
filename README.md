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

![folder-options](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/media/folder-options.png)

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
1. [Bookops.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/macrobooks/Bookops.mbk)
	* [Utilities.bas v1.0.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/Utilities.bas): includes macro updater, single and dual profile creation, and Connexion settings backup and restore functions
2. [Brooklyn.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/macrobooks/Brooklyn.mbk)
	* [BPLCallNum.bas v3.4.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/BPLCallNum.bas): BPL call number creation macro
	* [ToExcel.bas v1.2](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/BPLToExcel.bas): exports elements of displayed in the Connexion client record to the .xls format
3. [NYPL.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/macrobooks/NYPL.mbk)
	* [NYPLCallNum.bas v3.2.1](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/NYPLCallNum.bas): NYPL call number creation macro
	* [NYPLMusicCD.bas v1.5](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/NYPLMusicCD.bas): NYPL music CDs call number creation macro
4. [CAT.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/macrobooks/CAT.mbk)
	* [BackupSettings.bas v1.0.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATBackupSettings.bas): backups Connexion client settings
	* [BibcoEval.bas v1.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATBibcoEval.bas): evaluates and flags elements missing or incorrect for BIBCO records
	* [UpdateExport.bas v1.8](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATUpdateExport.bas): updates holdings, cleans 6xxs, sets appropriate Sierra load table, and exports records to Sierra; to be used in BookOps CAT
	* [UpdateExport_Permissive.bas v1.2](http://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATUpdateExport_Permissive.bas): more permissive version of `UpdateExport` macro allowing a wider range of vocabularies in 6xx fields
	* [ShareSave.bas v1.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATShareSave.bas)
    * [UserDual.bas v1.0.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATNewUserDual.bas): creates dual, BPL & NYPL, profiles in the Connexion client; requires access to CATAL network drive
    * [UserNYP.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATNewUserNYP.bas): creates single NYPL Connexion client profile
5. [ACQ.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/macrobooks/ACQ.mbk)
	* [RLOrderRec.bas v3.2](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/ACQRLOrderRec.bas): encodes order record information in a special MARC field and exports to Sierra
6. [NewMacros.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/macrobooks/newMacros.mbk)
	* [updater.bas v2.1](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/newMacrosUpdater.bas): pulls and updates local macrobooks using a network drive
## License
[MIT](https://opensource.org/licenses/MIT)
