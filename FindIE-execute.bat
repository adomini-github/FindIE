@echo off
timeout /t 60 /nobreak
if not exist C:\FindIE (
	mkdir C:\FindIE
)
xcopy \\syanaiseserver.syanaise-soudan.local\share\setup\FindIE\FindIE.ps1 C:\FindIE\ /y
cd /d C:\FindIE\
powershell -NoProfile -ExecutionPolicy Unrestricted .\FindIE.ps1 -hour 24 -path \\syanaiseserver\setup\FindIE\output -nolog 1