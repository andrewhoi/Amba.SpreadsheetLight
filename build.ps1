# Set-ExecutionPolicy remotesigned
gci .\source -Recurse "packages.config" |% {
	"Restoring " + $_.FullName
	.\source\.nuget\nuget.exe i $_.FullName -o .\source\packages
}
Import-Module .\Source\packages\psake.4.4.1\tools\psake.psm1

Invoke-Psake -framework '4.0x64'

Remove-Module psake

$nl = [Environment]::NewLine
Write-Host $nl
Write-Host "Press any key to continue..."

$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")