:: Criado por Janderson Costa em 06/11/2016.
:: Descrição:
::  Registra os Add-ins no sistema de registros do Windows.
::  Copia os Add-ins (arquivos) para o computador do usuário.
::  Cria a pasta de backup no computador do usuário.
:: Notas:
::  Válido apenas para arquivos do Excel.

@echo off


:: VARIÁVEIS GLOBAIS

set officeRegistryPath="HKEY_CURRENT_USER\Software\Microsoft\Office"

:: Excel
set excelAddinFileName=OfficeAutoBackup.xla
set excelAddinPath1=%appdata%\Microsoft\AddIns\%excelAddinFileName%
set excelAddinPath2=%appdata%\Microsoft\Suplementos\%excelAddinFileName%


:: INSTALAÇÃO

@echo Instalando registros e arquivos...

for /l %%i in (100,-1,14) do (
	:: itera de 100 à 14 como sendo o número/nome das chaves pesquisadas
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
