Bcdedit.exe /set nointegritychecks ON
bcdedit.exe -set loadoptions DDISABLE_INTEGRITY_CHECKS
bcdedit.exe -set TESTSIGNING ON
pause
exit