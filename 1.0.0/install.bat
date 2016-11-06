:: Criado por Janderson Costa em 06/11/2016.
:: Descri��o:
::  Registra os Add-ins no sistema de registros do Windows.
::  Copia os Add-ins (arquivos) para o computador do usu�rio.
::  Cria a pasta de backup no computador do usu�rio.
:: Notas:
::  V�lido apenas para arquivos do Excel.

@echo off


:: VARI�VEIS GLOBAIS

set officeRegistryPath="HKEY_CURRENT_USER\Software\Microsoft\Office"

:: Excel
set excelAddinFileName=OfficeAutoBackup.xla
set excelAddinPath1=%appdata%\Microsoft\AddIns\%excelAddinFileName%
set excelAddinPath2=%appdata%\Microsoft\Suplementos\%excelAddinFileName%


:: INSTALA��O

@echo Instalando registros e arquivos...

for /l %%i in (100,-1,14) do (
	:: itera de 100 � 14 como sendo o n�mero/nome das chaves pesquisadas
	reg query %officeRegistryPath%\%%i.0 > nul

	if not errorlevel 1 (
		reg add %officeRegistryPath%\%%i.0\Excel\Options /v OPEN /t REG_SZ /d "\"%excelAddinFileName%"\" /f
		reg add %officeRegistryPath%\%%i.0\Excel\AddInLoadTimes /v %excelAddinPath1% /t REG_BINARY /d 0 /f
		reg add %officeRegistryPath%\%%i.0\Excel\AddInLoadTimes /v %excelAddinPath2% /t REG_BINARY /d 0 /f

		copy add-ins\excel\%excelAddinFileName% %appdata%\Microsoft\AddIns /y
		copy add-ins\excel\%excelAddinFileName% %appdata%\Microsoft\Suplementos /y

		mkdir %USERPROFILE%\OfficeAutoBackup

		:: para o loop
		goto :break
	)
)
:break
