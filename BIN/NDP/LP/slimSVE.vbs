' Usage:
'
' - Make sure netfx*.msi and DisplayIcon.ico are in the same directory as this file before running:
'
' cscript <name_of_file>.vbs
'
' - Create administrative install to leave the setup files behind:
'
' msiexec /a <netfx_installer>.msi TARGETDIR=<path_to_new_folder>
'
' script by dumpydooby (modded by ricktendo64)
Option Explicit
Dim ws, installer, fs, db, view, record, x
Set ws = WScript.CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
Set installer = WScript.CreateObject("WindowsInstaller.Installer")
If WScript.Arguments.Count <> 0 Then
	For each x in WScript.Arguments
		ProcessMSI x
	Next
Else
	If fs.FileExists("netfx_Full_x64.msi") Then ProcessMSI "netfx_Full_x64.msi"
	If fs.FileExists("netfx_Full_x86.msi") Then ProcessMSI "netfx_Full_x86.msi"
	If fs.FileExists("netfx_FullLP_x64.msi") Then ProcessMSI "netfx_FullLP_x64.msi"
	If fs.FileExists("netfx_FullLP_x86.msi") Then ProcessMSI "netfx_FullLP_x86.msi"
	If fs.FileExists("netfx_Full_GDR_x64.msi") Then ProcessMSI "netfx_Full_GDR_x64.msi"
	If fs.FileExists("netfx_Full_GDR_x86.msi") Then ProcessMSI "netfx_Full_GDR_x86.msi"
	If fs.FileExists("netfx_Full_LDR_x64.msi") Then ProcessMSI "netfx_Full_LDR_x64.msi"
	If fs.FileExists("netfx_Full_LDR_x86.msi") Then ProcessMSI "netfx_Full_LDR_x86.msi"
	If fs.FileExists("netfx_FullLP_GDR_x64.msi") Then ProcessMSI "netfx_FullLP_GDR_x64.msi"
	If fs.FileExists("netfx_FullLP_GDR_x86.msi") Then ProcessMSI "netfx_FullLP_GDR_x86.msi"
	If fs.FileExists("netfx_FullLP_LDR_x64.msi") Then ProcessMSI "netfx_FullLP_LDR_x64.msi"
	If fs.FileExists("netfx_FullLP_LDR_x86.msi") Then ProcessMSI "netfx_FullLP_LDR_x86.msi"
End If
'**********************************************************************
'** Function; Query MSI database                                     **
'**********************************************************************
Function QueryDatabase(arrOpts)
	On Error Resume Next
	Dim query, file, binary : binary = false
	If LCase(TypeName(arrOpts)) = "string" Then
		query = arrOpts
	Else
		If fs.FileExists(arrOpts(0)) Then
			file = arrOpts(0)
			query = arrOpts(1)
		Else
			query = arrOpts(0)
			file = arrOpts(1)
		End If
		binary = true
	End If
	WScript.Echo query
	If binary Then
		Set record = installer.CreateRecord(1)
		record.SetStream 1, file
	End If
	Set view = db.OpenView (query) : CheckError
	If binary Then
		view.Execute record : CheckError
	Else
		view.Execute : CheckError
	End If
	view.close
	Set view = nothing
	If binary Then Set record = nothing
	binary = false
	db.commit : CheckError
End Function
'**********************************************************************
'** Subroutine; Check errors in most recently executed MSI command   **
'**********************************************************************
Sub CheckError
	Dim message, errRec
	If Err = 0 Then Exit Sub
	message = Err.Source & " " & Hex(Err) & ": " & Err.Description
	If Not installer Is Nothing Then
		Set errRec = installer.LastErrorRecord
		If Not errRec Is Nothing Then message = message & vbNewLine & errRec.FormatText
	End If
	Wscript.Echo "" : Wscript.Echo message : Wscript.Echo ""
	Wscript.Quit 2
End Sub
'**********************************************************************
'** Function; Push changes to MSI                                    **
'**********************************************************************
Function ProcessMSI(file)
	Set db = installer.OpenDatabase(file, 1)
	On Error Resume Next
	QueryDatabase("DELETE FROM `Component` WHERE `Directory_` = 'IdentityCacheDir'") 
	QueryDatabase("DELETE FROM `Component` WHERE `Directory_` = 'NetFx_IdentityCacheDir_Graphics'") 
	QueryDatabase("DELETE FROM `Component` WHERE `Directory_` = 'NetFx_Installer_ResourcesDir_enu'") 
	QueryDatabase("DELETE FROM `Component` WHERE `Directory_` = 'NetFx_Installer_ResourcesDir_sve'") 
	QueryDatabase("DELETE FROM `Component` WHERE `Component` = 'NetFxRepair_Resources_sve'") 
	QueryDatabase("DELETE FROM `Component` WHERE `Component` = 'NetFxRepair_exe'") 
	QueryDatabase("DELETE FROM `CreateFolder` WHERE `Directory_` = 'IdentityCacheDir'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory` = 'IdentityCacheDir'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory` = 'Framework_amd64_IronSetupCache'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory` = 'Framework_x86_IronSetupCache'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory` = 'FrameworkSetupCache_amd64_IronSetupCache'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory` = 'FrameworkSetupCache_x86_IronSetupCache'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory` = 'Microsoft.NET_IronSetupCache_amd64'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory` = 'Microsoft.NET_IronSetupCache_x86'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory` = 'WindowsFolder_IronSetupCache_amd64'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory` = 'WindowsFolder_IronSetupCache_x86'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory_Parent` = 'Framework_amd64_IronSetupCache'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory_Parent` = 'Framework_x86_IronSetupCache'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory_Parent` = 'FrameworkSetupCache_amd64_IronSetupCache'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory_Parent` = 'FrameworkSetupCache_x86_IronSetupCache'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory_Parent` = 'FrameworkVersionFolder_amd64_IronSetupCache'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory_Parent` = 'FrameworkVersionFolder_x86_IronSetupCache'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory_Parent` = 'IdentityCacheDir'") 
	QueryDatabase("DELETE FROM `AdminExecuteSequence` WHERE `Action` = 'CA_WindowsFolder_IronSetupCache_amd64'") 
	QueryDatabase("DELETE FROM `AdminExecuteSequence` WHERE `Action` = 'CA_WindowsFolder_IronSetupCache_x86'") 
	QueryDatabase("DELETE FROM `AdminUISequence` WHERE `Action` = 'CA_WindowsFolder_IronSetupCache_amd64'") 
	QueryDatabase("DELETE FROM `AdminUISequence` WHERE `Action` = 'CA_WindowsFolder_IronSetupCache_x86'") 
	QueryDatabase("DELETE FROM `CustomAction` WHERE `Action` = 'CA_WindowsFolder_IronSetupCache_amd64'") 
	QueryDatabase("DELETE FROM `CustomAction` WHERE `Action` = 'CA_WindowsFolder_IronSetupCache_x86'") 
	QueryDatabase("DELETE FROM `CustomAction` WHERE `Action` = 'CA_CompressSetupCache'") 
	QueryDatabase("DELETE FROM `CustomAction` WHERE `Action` = 'CA_SetCompressSetupCache'") 
	QueryDatabase("DELETE FROM `InstallExecuteSequence` WHERE `Action` = 'CA_WindowsFolder_IronSetupCache_amd64'") 
	QueryDatabase("DELETE FROM `InstallExecuteSequence` WHERE `Action` = 'CA_WindowsFolder_IronSetupCache_x86'") 
	QueryDatabase("DELETE FROM `InstallExecuteSequence` WHERE `Action` = 'CA_CompressSetupCache'") 
	QueryDatabase("DELETE FROM `InstallExecuteSequence` WHERE `Action` = 'CA_SetCompressSetupCache'") 
	QueryDatabase("DELETE FROM `InstallUISequence` WHERE `Action` = 'CA_WindowsFolder_IronSetupCache_amd64'") 
	QueryDatabase("DELETE FROM `InstallUISequence` WHERE `Action` = 'CA_WindowsFolder_IronSetupCache_x86'") 
	QueryDatabase("DELETE FROM `FeatureComponents` WHERE `Feature_` = 'F_Installer_IdentityARP'") 
	QueryDatabase("DELETE FROM `FeatureComponents` WHERE `Feature_` = 'MSICache_FullLP_amd64_sve'") 
	QueryDatabase("DELETE FROM `FeatureComponents` WHERE `Feature_` = 'MSICache_FullLP_x86_sve'") 
	QueryDatabase("DELETE FROM `FeatureComponents` WHERE `Feature_` = 'NetFx_Installer_Setup_ddf'") 
	QueryDatabase("DELETE FROM `FeatureComponents` WHERE `Feature_` = 'Netfx_NetFxRepair'") 
	QueryDatabase("DELETE FROM `FeatureComponents` WHERE `Feature_` = 'Netfx_SetupUtility'") 
	QueryDatabase("DELETE FROM `Feature` WHERE `Feature` = 'F_Installer_IdentityARP'") 
	QueryDatabase("DELETE FROM `Feature` WHERE `Feature` = 'MSICache_FullLP_amd64_sve'") 
	QueryDatabase("DELETE FROM `Feature` WHERE `Feature` = 'MSICache_FullLP_x86_sve'") 
	QueryDatabase("DELETE FROM `Feature` WHERE `Feature` = 'NetFx_Installer_Setup_ddf'") 
	QueryDatabase("DELETE FROM `Feature` WHERE `Feature` = 'Netfx_NetFxRepair'") 
	QueryDatabase("DELETE FROM `Feature` WHERE `Feature` = 'Netfx_SetupUtility'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'Installer_MSICache_FullLP_amd64_sve'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'Installer_MSICache_FullLP_x86_sve'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'NetFx_Installer_Core'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'NetFx_Installer_Customizable'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'NetFx_Installer_EULAs_enu'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'NetFx_Installer_EULAs_sve'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'NetFx_Installer_Graphics'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'NetFx_Installer_LocalizedData_enu'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'NetFx_Installer_LocalizedData_sve'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'NetFx_Installer_ParameterInfo'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'NetFx_Installer_Resources_enu'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'NetFx_Installer_Resources_sve'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'NetFxRepair_exe'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'NetFxRepair_Resources_sve'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SetupUtility_exe'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'eula.1033.rtf'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'eula.1053.rtf'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_DHtmlHeader.html'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_DisplayIcon.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_header.bmp'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_LocalizedData_enu.xml'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_LocalizedData_sve.xml'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Print.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Rotate1.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Rotate2.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Rotate3.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Rotate4.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Rotate5.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Rotate6.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Rotate7.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Rotate8.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Rotate9.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Rotate10.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Save.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Setup.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_SetupUi.xsd'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_SplashScreen.bmp'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_stop.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_Strings.xml'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_SysReqMet.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_SysReqNotMet.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_UiInfo.xml'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_warn.ico'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'NetFx_watermark.bmp'") 
	QueryDatabase("DELETE FROM `Registry` WHERE `Component_` = 'Identity_RegEntries_amd64_sve_FullLP'") 
	QueryDatabase("DELETE FROM `Registry` WHERE `Component_` = 'Identity_RegEntries_x86_sve_FullLP'") 
	QueryDatabase("CREATE TABLE `Icon` (`Name` CHAR(72) NOT NULL, `Data` OBJECT NOT NULL PRIMARY KEY `Name`)") 
	QueryDatabase(Array("DisplayIcon.ico", "INSERT INTO `Icon` (`Name`, `Data`) VALUES ('DisplayIcon', ?)")) 
	QueryDatabase("INSERT INTO `Property` (`Property`,`Value`) VALUES ('ARPPRODUCTICON','DisplayIcon')") 
	QueryDatabase("INSERT INTO `Property` (`Property`,`Value`) VALUES ('EXTUI','1')") 
	QueryDatabase("DELETE FROM `Property` WHERE `Property` = 'ARPSYSTEMCOMPONENT'") 
	Set db = nothing
End Function