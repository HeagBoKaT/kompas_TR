@echo off
echo Установка редактора технических требований KOMPAS-3D
echo ===================================================

set INSTALL_DIR="%PROGRAMFILES%\KOMPAS-Tech-Req-Editor"
set DESKTOP_SHORTCUT="%USERPROFILE%\Desktop\KOMPAS-Технические требования.lnk"
set START_MENU_DIR="%APPDATA%\Microsoft\Windows\Start Menu\Programs\KOMPAS-Tech-Req-Editor"

echo Создание директории установки...
if not exist %INSTALL_DIR% mkdir %INSTALL_DIR%

echo Копирование файлов...
xcopy /Y /E /I "dist\*.*" %INSTALL_DIR%

echo Создание ярлыка на рабочем столе...
powershell -Command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%DESKTOP_SHORTCUT%'); $Shortcut.TargetPath = '%INSTALL_DIR%\KOMPAS-Технические требования.exe'; $Shortcut.Save()"

echo Создание ярлыка в меню Пуск...
if not exist %START_MENU_DIR% mkdir %START_MENU_DIR%
powershell -Command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%START_MENU_DIR%\KOMPAS-Технические требования.lnk'); $Shortcut.TargetPath = '%INSTALL_DIR%\KOMPAS-Технические требования.exe'; $Shortcut.Save()"

echo Установка завершена!
echo Приложение установлено в %INSTALL_DIR%
echo Ярлыки созданы на рабочем столе и в меню Пуск
echo.
echo Нажмите любую клавишу для завершения...
pause > nul
