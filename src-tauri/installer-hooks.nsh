!macro NSIS_HOOK_POSTINSTALL
  DetailPrint "Installing Cloudreve Office add-ins..."
  ExecWait '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "$INSTDIR\office-addin\Install-OfficeAddins.ps1" -Action Install -PackageDirectory "$INSTDIR\office-addin"' $0
  ${If} $0 != 0
    DetailPrint "Cloudreve Office add-ins were not installed. Check that the VSTO Runtime is installed."
    MessageBox MB_ICONEXCLAMATION|MB_OK "Cloudreve was installed, but the Office add-ins could not be installed. Install the Microsoft Visual Studio Tools for Office Runtime, then run the installer again."
  ${EndIf}
!macroend

!macro NSIS_HOOK_PREUNINSTALL
  DetailPrint "Removing Cloudreve Office add-ins..."
  ExecWait '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "$INSTDIR\office-addin\Install-OfficeAddins.ps1" -Action Uninstall -PackageDirectory "$INSTDIR\office-addin"' $0
!macroend
