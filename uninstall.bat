@echo off
echo Удаление редактора технических требований KOMPAS-3D
echo ===================================================

set INSTALL_DIR="%PROGRAMFILES%\KOMPAS-Tech-Req-Editor"
set DESKTOP_SHORTCUT="%USERPROFILE%\Desktop\KOMPAS-Технические требования.lnk"
set START_MENU_DIR="%APPDATA%\Microsoft\Windows\Start Menu\Programs\KOMPAS-Tech-Req-Editor"

echo Удаление ярлыка с рабочего стола...
if exist %DESKTOP_SHORTCUT% del /F /Q %DESKTOP_SHORTCUT%

echo Удаление ярлыков из меню Пуск...
if exist %START_MENU_DIR% rmdir /S /Q %START_MENU_DIR%

echo Удаление файлов приложения...
if exist %INSTALL_DIR% rmdir /S /Q %INSTALL_DIR%

echo Удаление завершено!
echo.
echo Нажмите любую клавишу для завершения...
pause > nul
