@SET INSTALLDIR="%LOCALAPPDATA%\Viramate\Installer"
@echo Setting up Viramate Installer...
@mkdir %INSTALLDIR%
@copy * "%INSTALLDIR%"
@cd %INSTALLDIR%
@del sfx-bootstrap.bat
@start Viramate.exe
@echo OK.