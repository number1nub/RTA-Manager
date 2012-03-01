@Echo off
cls

:: Destination folder
SET DESTINATION_PATH=%USERPROFILE%\Halliburton RTA Manager\Include

:: File paths to copy
SET RTA_TOOLS=%CD%\..\Include\RTA Sheet Tools.exe
SET CMD_LIN=%CD%\..\Include\CMDline_Functions.exe


COPY /Y "%RTA_TOOLS%" "%DESTINATION_PATH%"
COPY /Y "%CMD_LINE%" "%DESTINATION_PATH%"