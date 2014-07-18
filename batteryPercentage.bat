set battery=Battery Percentage(%%)
for /f "skip=1" %%a in ('wmic path Win32_Battery get EstimatedChargeRemaining') do if %%a gtr 0 set BATTERYLEVEL=%%a
@echo Battery percentage=%BATTERYLEVEL%


set ref="http://msdn.microsoft.com/en-us/library/windows/desktop/aa394074(v=vs.85).aspx"