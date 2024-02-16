@echo off
set hostname 
hostname
echo %hostname%
pause
exit
systeminfo >C:%UserProfile%\Systeminfo.txt
ipconfig /all >>%UserProfile%\Systeminfo.txt
wmic bios get serialnumber | more >> %UserProfile%\Desktop\Systeminfo.txt
