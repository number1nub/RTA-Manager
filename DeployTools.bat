@Echo off
cls
echo.
echo.
echo Preparing to deploy...

:: Destination folder
SET DESTINATION_PATH=%USERPROFILE%\My Documents\Halliburton RTA Manager\Include

:: File paths to copy
SET RTA_TOOLS=%CD%\Include\RTA Sheet Tools.exe
SET CMD_LINE=%CD%\Include\CMDline_Functions.exe

echo Copying file....
echo.

COPY /Y "%RTA_TOOLS%" "%DESTINATION_PATH%"

COPY /Y "%CMD_LINE%" "%DESTINATION_PATH%"

echo.
echo.
echo        Done!