# BookOps Cataloging Dept. Connexion Macros
A repository of BookOps Cataloging Department macros for the OCLC Connexion client.

## Installation
### Easy Way
If you have access to the CAT department shared drive (S: drive) you can find the macrobook files in S:/CATAL/Connex/macros and copy the files from there. Otherwise to install macros, find the macrobook you wish to install in the [`Connexion-Macros/macrobooks/` folder](https://github.com/BookOps-CAT/Connexion-Macros/tree/main/macrobooks), click on it, and download it to your computer.

Example, downloading the `CAT macrobook` which includes UpdateExport, BackupSettings, and other macros:
![macrobook-download-button](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/media/macrobook-download.png)

Then copy the macro book (.mbk file) to the `C:/Users/[your-name]/AppData/Roaming/OCLC/Connex/Macros/` directory on your computer. If a copy of the macro exists already, overwrite it with the downloaded version. It is recommended to do this while your Connexion client is closed.

If you do not see the `AppData` directory in your Windows profile (it is hidden by default on library computers), reveal it using the following instruction:
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
12. After the macro is saved it can be mapped to a User Tool, Keymap, or Shortcut as usual (optional).
Done!

## Setup
After copying the macrobooks to `C:/Users/[your-name]/AppData/Roaming/OCLC/Connex/Macros/` you can set up your macros in ConnexionClient and map them to user tools.

### Create New User Profile
When setting up the Connexion Client for the first time, users will need to create a new user profile. CAT staff onsite at LSC can create NYPL or dual NYPL/BPL profiles by running the NewUserNYP or NewUserDual macro directly from the Macro window in the client. Users from other departments or CAT staff working offsite will not be able to create a profile using this process. The gateway configuration file is `Connects.ini` and is available on the CATAL shared drive under `CATAL/Connex/onboarding/dual` or `CATAL/Connex/onboarding/nypl-only`.

![user-tools](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/media/manage-macros.png)


1. Open Connexion and navigate to Tools > Macros > Manage
2. Under CAT, you will see NewUserDual and NewUserNYP. Select the appropriate macro based on whether you will need profiles for both NYPL and BPL or just NYPL.
3. After selecting the macro, click run. 

### Mapping Macros to User Tools
After setting up a profile users can map macros to shortcuts and add them to the toolbar.

![user-tools](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/media/user-tools.png)

1. Open Connexion and navigate to Tools > User Tools > Manage
2. Select the Macro button at the top of the pop up window. A list of available macros should appear on the left side of the window.
3. Select the macro you would like to assign from the list and select the user tool to map it to on the right. Click Assign Tool. Within the Cataloging Department macros are assigned to the following user tools:
    - **UserTool 1**: CAT!UpdateExport
    - **UserTool 2**: Brooklyn!CallNum
    - **UserTool 3**: OCLC!Add33x
    - **UserTool 4**: OCLC!Generate043
    - **UserTool 5**: unassigned
    - **UserTool 6**: OCLC!GenerateAuthorityRecord
    - **UserTool 7**: unassigned
    - **UserTool 8**: unassigned
    - **UserTool 9**: NYPL!CallNum
    - **UserTool 10**: Bookops!Utilities
4. When finished assigning macros to user tools, click OK 
5. These shortcuts can be added to the tool bar by navigating to Tools > Toolbar Editor. Scroll through the list until you find ToolsUserTools1. Click on the tool you would like to add and drag it to your tool bar. When finished click close.

## Macros
1. [Bookops.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/macrobooks/Bookops.mbk)
	* [Utilities.bas v1.0.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/Utilities.bas): includes macro updater, single and dual profile creation, and Connexion settings backup and restore functions
2. [Brooklyn.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/macrobooks/Brooklyn.mbk)
	* [BPLCallNum.bas v3.4.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/BPLCallNum.bas): BPL call number creation macro
	* [ToExcel.bas v1.2](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/BPLToExcel.bas): exports elements of displayed in the Connexion client record to the .xls format
3. [NYPL.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/macrobooks/NYPL.mbk)
	* [NYPLCallNum.bas v3.3.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/NYPLCallNum.bas): NYPL call number creation macro
	* [NYPLMusicCD.bas v1.5](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/NYPLMusicCD.bas): NYPL music CDs call number creation macro
4. [CAT.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/macrobooks/CAT.mbk)
	* [BackupSettings.bas v1.0.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATBackupSettings.bas): backups Connexion client settings
	* [BibcoEval.bas v1.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATBibcoEval.bas): evaluates and flags elements missing or incorrect for BIBCO records
	* [UpdateExport.bas v1.10](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATUpdateExport.bas): updates holdings, cleans 6xxs, sets appropriate Sierra load table, and exports records to Sierra; to be used in BookOps CAT
	* [UpdateExport_Permissive.bas v1.4](http://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATUpdateExport_Permissive.bas): more permissive version of `UpdateExport` macro allowing a wider range of vocabularies in 6xx fields
	* [ShareSave.bas v1.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATShareSave.bas)
    * [UserDual.bas v1.0.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATNewUserDual.bas): creates dual, BPL & NYPL, profiles in the Connexion client; requires access to CATAL network drive
    * [UserNYP.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATNewUserNYP.bas): creates single NYPL Connexion client profile
    * [ValEncoding.bas](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/CATValEncoding.bas): Analyzes the entire record for possible non MARC-8 encoded characters
5. [ACQ.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/macrobooks/ACQ.mbk)
	* [RLOrderRec.bas v4.0](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/ACQRLOrderRec.bas): encodes order record information in a special MARC field and exports to Sierra
6. [NewMacros.mbk macrobook](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/macrobooks/newMacros.mbk)
	* [updater.bas v2.1](https://github.com/BookOps-CAT/Connexion-Macros/blob/main/source/newMacrosUpdater.bas): pulls and updates local macrobooks using a network drive
## License
[MIT](https://opensource.org/licenses/MIT)
