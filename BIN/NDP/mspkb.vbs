If Wscript.Arguments.Count = 0 Then Wscript.Quit 1
On Error Resume Next
Dim wi, si
Set wi = Wscript.CreateObject("WindowsInstaller.Installer")
Set si = wi.SummaryInformation(Wscript.Arguments(0), 0)
Wscript.Echo si.Property(2)
Wscript.Quit 0
