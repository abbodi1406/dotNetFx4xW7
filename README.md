# .NET 4.7.2 Repack Installer for Windows 7 SP1 x86/x64

* Repacked slim installer for .NET Framework 4.7.2 with latest updates integrated and without extra setup payload.

* Support creating slim updated installer for .NET 4.6 and later.

* Note:  
to rebuild/recreate *netfx_Full.mzz* file, you need to download [WiX Toolset binaries](https://github.com/wixtoolset/wix3/releases/download/wix3111rtm/wix311-binaries.zip)  
then, extract the files into **BIN** directory (you can delete the big sdk folder, it's not needed).

## Create slim main package:

* Place NDP*-ENU.exe file here next to the scripts (e.g. NDP472-KB4054541-x86-x64-ENU.exe or NDP472-KB4054530-x86-x64-AllOS-ENU.exe).

* If exist, also place any updates exe files here (make sure to get files for both x86 and x64).

* Optional, edit *dotNetFx.cmd* and change the two options values (1 or 0):  
**BuildMzz**: Rebuild/Repack files inside *netfx_Full.mzz* instead administrative directories.  
**ShowMsp**: Show slipstreamed patches in "Control Panel\Programs and Features\Installed Updates".

* Run *dotNetFx.cmd* as administrator.

* Note: you may change **\BIN\NDP\DisplayIcon.ico** prior execution, if you want another ARP icon.

* Optional, use *7zSFX.cmd* to create 7-zip SFX executable installer afterwards.  
you will need to modify **\BIN\7zSFX\7zSD.sfx** module with resource editor to update and set product name and version.

* Available command line switches:  
```
/y  
Passive mode, shows progress but requires no user interaction.  
/ai  
Quiet mode, no user input required, only extraction dialog is shown.  
/ai /gm2  
Quiet mode, no user input required or output shown.  
/sfxlang:  
Set the program display language, if possible. Example: /sfxlang:1031  
/h | /?  
Display this help.  

Examples:  

Automatically install package and display progress:  
NDP472-Slim-x86-x64-ENU.exe /y  
Silently install package and display no progress:  
NDP472-Slim-x86-x64-ENU.exe /ai /gm2
```

## Create slim languages packages:

* Place NDP*.exe files inside **LP** directory (Recommended to put all lang files, or you may put only specific languages).

* Run *dotNetFxLP.cmd* as administrator.

* Note: you may change **\BIN\NDP\LP\DisplayIcon.ico** prior execution, if you want another ARP icon.

* Optional, use *7zSFX.cmd* to create 7-zip SFX executable installer afterwards.  
you will need to modify **\BIN\7zSFXLP\7zSD.sfx** module with resource editor to update and set product name and version.

* Available command line switches:  
```
/y  
Passive mode, shows progress but requires no user interaction.  
/ai  
Quiet mode, no user input required, only extraction dialog is shown.  
/ail  
LP Passive mode. Default or specified LP is installed.  
/sfxlang:  
Language mode. Change default display language and/or language pack to install  
/gm2  
Optional command line switch to disable extraction dialog  
/h | /?  
Display this help.  

Examples:  

Automatically install default language and display progress:  
NDP472-Slim-x86-x64-INTL.exe /y  
Silently install default language and display no progress:  
NDP472-Slim-x86-x64-INTL.exe /ai  
Silently install specific language (1031=German):  
NDP472-Slim-x86-x64-INTL.exe /sfxlang:1031 /ai  
Install French language and display progress:  
NDP472-Slim-x86-x64-INTL.exe /sfxlang:1036 /ail
```

## Credits:

- [ricktendo64](https://forums.mydigitallife.net/members/28038/) (Original installer and msi vbs slimmers).  
- [7-zip](https://www.7-zip.org/).  
- [WiX Toolset](https://wixtoolset.org).  
- Microsoft Windows SDK for Windows 7 (Windows Installer utility scripts).  
- Microsoft .NET Framework is an intellectual property of (c) Microsoft Corporation. All Rights Reserved.
