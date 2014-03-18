call "%VS110COMNTOOLS%"vsvars32.bat

devenv "CmisSync\Windows\CmisSync.sln" /Clean "Debug|x86" /Project "Installer"
devenv "CmisSync\Windows\CmisSync.sln" /Clean "Debug|x64" /Project "Installer"
devenv "CmisSync\Windows\CmisSync.sln" /Clean "Debug|x86" /Project "InstallerBootstrapper"

devenv "CmisSync\Windows\CmisSync.sln" /Clean "Release|x86" /Project "Installer"
devenv "CmisSync\Windows\CmisSync.sln" /Clean "Release|x64" /Project "Installer"
devenv "CmisSync\Windows\CmisSync.sln" /Clean "Release|x86" /Project "InstallerBootstrapper"

devenv "CmisSync\Windows\CmisSync.sln" /Build "Debug|x86" /Project "Installer"
devenv "CmisSync\Windows\CmisSync.sln" /Build "Debug|x64" /Project "Installer"
devenv "CmisSync\Windows\CmisSync.sln" /Build "Debug|x86" /Project "InstallerBootstrapper"

devenv "CmisSync\Windows\CmisSync.sln" /Build "Release|x86" /Project "Installer"
devenv "CmisSync\Windows\CmisSync.sln" /Build "Release|x64" /Project "Installer"
devenv "CmisSync\Windows\CmisSync.sln" /Build "Release|x86" /Project "InstallerBootstrapper"

