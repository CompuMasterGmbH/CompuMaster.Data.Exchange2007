::REM -UpdateNuGetExecutable not required since it's updated by VS.NET mechanisms
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& 'CompuMaster.Data.Exchange\_CreateNewNuGetPackage\DoNotModify\New-NuGetPackage.ps1' -ProjectFilePath '.\CompuMaster.Data.Exchange\CM.Data.Exchange2007.vbproj' -verbose -NoPrompt -DoNotUpdateNuSpecFile"
pause