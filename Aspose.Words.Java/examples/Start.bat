@echo off

rem Find the absolute path of the Examples directory
SET EXAMPLES_DIR=%cd%\Examples

rem Create a variable to store the JDK or JRE Home folder
SET SUN_JRE_1_7_HOME=

:jre
echo Trying to find JRE 1.7 in Program Files....
SET SUN_JRE_1_7_HOME=%ProgramFiles%\Java\jre7
IF EXIST "%SUN_JRE_1_7_HOME%\bin\java.exe" (
	goto :foundjre
)

:jre_x86
echo Trying to find JRE 1.7 in Program Files x86....
SET SUN_JRE_1_7_HOME=%ProgramFiles(x86)%\Java\jre7
IF EXIST "%SUN_JRE_1_7_HOME%\bin\java.exe" (
	goto :foundjre
)

:manual
rem Try default Java
SET SUN_JRE_1_7_HOME=
echo Trying default java....
java  -cp "lib\*;.;Aspose.Examples.Launcher.jar;" aspose.examples.launcher.AsposeExamplesLauncher

rem No Java found in any of the above conditions, Set the path manually
rem Un-comment below 2 lines, if you want to set path manually
rem SUN_JRE_1_7_HOME=X:\Program Files (x86)\Java\jre7
rem goto :foundjre

cd App
Call Startapp.bat
cd ..


echo Press any key to exit....
pause > nul


goto :end

:foundjre
echo Found JRE in %SUN_JRE_1_7_HOME%

rem We searched for jre7, there are no versions, so just execute the command directly
"%SUN_JRE_1_7_HOME%\bin\java" -cp "lib\*;.;Aspose.Examples.Launcher.jar;%SUN_JRE_1_7_HOME%\lib\*;" aspose.examples.launcher.AsposeExamplesLauncher
cd App
Call Startapp.bat
cd ..
goto:end


:end
echo Finished Launcher Batch file execution


rem Reset the variables
SET SUN_JRE_1_7_HOME=
SET EXAMPLES_DIR=