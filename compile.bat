@echo off
SETLOCAL ENABLEEXTENSIONS

IF "%1"=="" (
	ECHO.
	SET /P compile_file=请输入需要打包的 Python 脚本文件路径：
) ELSE (
	SET compile_file=%1
)
IF NOT EXIST "%compile_file%" (
	ECHO.
	ECHO 指定的脚本文件不存在！
	GOTO END
)

FOR /F "delims=" %%A IN ('"echo %compile_file%"') do (
	SET compile_path=%%~dpA
	SET compile_name=%%~nA
)

ECHO.
SET /P app_name=请输入打包后的程序文件名，默认为 %compile_name% ：
IF DEFINED app_name (
	SET compile_name=--name %app_name%
) ELSE (
	SET compile_name=--name %compile_name%
)

IF EXIST icon.ico (
	SET package_icon=--icon "%compile_path%\icon.ico"
) ELSE (
	SET package_icon= 
)

ECHO.
choice /C FD /M 请选择打包模式：打包为单文件[F]，还是打包为单文件夹[D]？
IF ERRORLEVEL 2 SET compile_mode=--onedir
IF ERRORLEVEL 1 SET compile_mode=--onefile

ECHO.
choice /C YN /M 运行时是否显示控制台窗口？
IF ERRORLEVEL 2 SET show_console=--console
IF ERRORLEVEL 1 SET show_console=--noconsole

ECHO.
IF EXIST %compile_path%\.env\Scripts\activate.bat (
	ECHO 正在激活脚本运行的 Python 虚拟环境...
	CALL %compile_path%\.env\Scripts\activate.bat
)
IF EXIST %compile_path%\.venv\Scripts\activate.bat (
	ECHO 正在激活脚本运行的 Python 虚拟环境...
	CALL %compile_path%\.venv\Scripts\activate.bat
)

ECHO.
ECHO PyInstaller %compile_mode% %show_console% %package_icon% %compile_name% %compile_file%
pyinstaller %compile_mode% %show_console% %package_icon% %compile_name% %compile_file%
IF ERRORLEVEL 1 (
	ECHO.
	ECHO 打包 Python 脚本时出现错误! 
) ELSE (
	ECHO.
	ECHO  脚本 %compile_file% 打包完成。
)

IF EXIST %compile_path%\.env\Scripts\deactivate.bat (
	ECHO.
	ECHO 正在退出脚本运行的 Python 虚拟环境...
	CALL %compile_path%\.env\Scripts\deactivate.bat
)
IF EXIST %compile_path%\.venv\Scripts\deactivate.bat (
	ECHO.
	ECHO 正在退出脚本运行的 Python 虚拟环境...
	CALL %compile_path%\.venv\Scripts\deactivate.bat
)ECHO.
PAUSE