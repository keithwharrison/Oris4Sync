call "%VS110COMNTOOLS%"vsvars32.bat

devenv "CmisSync\Windows\CmisSync.sln" /Build "Release|x86" /Project "Installer"
devenv "CmisSync\Windows\CmisSync.sln" /Build "Release|x64" /Project "Installer"
devenv "CmisSync\Windows\CmisSync.sln" /Build "Release|x86" /Project "InstallerBootstrapper"

