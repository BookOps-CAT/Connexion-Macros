# BookOps Cataloging Dept. Connexion Macros
A repository of the BookOps cataloging department Macros for OCLC Connexion client.

## Installation
### Easy Way
To install,copy the macro book (.mbk file in /macrobooks folder) to the "C:\Program Files (x86)\OCLC\Connexion\Program\Macros" directory.

Simple as that!
### Hard Way (Copy and Paste from Source)
Alternatively copy the code for the individual macro directly from the .BAS file (/source directory) using the macro editor in Connexion.

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

## Macros
1. Brooklyn.mbk macrobook
	* CallNum (BPLCallNum.bas)
	* ToExcel (needs updating)
2. NYPL.mbk macrobook
	* CallNum (NYPLCallNum.bas)
	* MusicCD (NYPLMusicCD.bas)
3. CAT.mbk macrobook
	* BibcoEval (CATBibcoEval.bas)
	* UpdateExport (CATUpdateExport.bas)
4. ACQ.mbk macrobook
	* RLOrderRec (ACQRLOrderRec.bas)
5. NewMacros.mbk macrobook
	* ACQupdater (NewMacrosACQupdater.bas)
	* CATupdater (NewMacrosCATupdater.bas)
## License
[MIT](https://opensource.org/licenses/MIT)
