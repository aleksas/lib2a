
set lib_file= OpenAL32.lib
set dll_file= OpenAL32.dll
set def_file= OpenAL32.def
set   a_file= libopenal32.a


cls
@echo off

set project_name=LIB to A convert v1.0

color 3f
Title %project_name%

cls

echo.
echo.
echo  [ %project_name% ]
echo  Free Open Source program - GNU GPL Licence.
echo.
echo.
echo  Overview
echo.
echo  This tool automatically converts a .LIB file (MS Visual C linker library) 
echo  into .A file (MinGW linker library).
echo  It is useful to make a C/C++ program a reusable component.
echo.
echo.
echo  Contact
echo.
echo  EMail : CorEngine@gmail.com
echo  MSN   : CorEngine@hotmail.com.ar
echo  Blog  : CorEngine.wordpress.com (Spanish Blog).
echo.
echo.
echo  by FX Programmer - 2010.
echo.
echo.
echo  Press any key to continue ...
pause>nul

cd convert

cls

echo.
echo.
echo  [ Making .DEF file ]
echo.
echo.
echo  Hang out while your .DEF file are being generated...
echo.
echo.

..\bin\reimp.exe -d %lib_file%

echo.
echo  Press any key to continue ...
pause>nul

cls

echo.
echo.
echo  [ Making .A linker library ]
echo.
echo.
echo  Hang out while your .A linker library are being generated...
echo.
echo.

..\bin\dlltool.exe -d %def_file% -D %dll_file% -k -l %a_file%

echo.
echo  Press any key to exit ...
pause>nul

cls
exit
