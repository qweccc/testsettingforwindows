strComputers = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%COMPUTERNAME%") 
 strComputers = "." 
 If WScript.Arguments.count = 1 Then 
   strComputers = WScript.Arguments(0) 
 End If 

 arrComputers = Split(strComputers,",") 
    For Each objComputer In arrComputers 
    strComputer = trim(objComputer) 
 On Error Resume Next    
 Set objHW = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate, authenticationLevel=pktPrivacy}!\\" _ 
    & strComputer & "\root\cimv2") 

 ' objService.Security_.AuthenticationLevel = 0 
 ' objService.Security_.ImpersonationLevel = 3 

 Set colCPU = objHW.InstancesOf("Win32_Processor") 
 S = "電腦名稱" & vbTab & strComputer & vbCrLf  & "==============================================================" &vbCrLf 
 Set colCS = objHW.InstancesOf("Win32_computerSystem") 
 S = S & "登入使用者" & vbTab 
 For Each objCS in colCS 
   S = S & objCS.Username 
 Next 
 S = S & vbCrLf & "==============================================================" &vbCrLf 


 S = S & "CPU規格" & vbCrLf & "==============================================================" &vbCrLf 
 For Each objCPU in colCPU 
    S = S & "CPU位址:" & right(objCPU.DeviceID,1) & vbTab & vbTab & "型號:"& trim(objCPU.Name) &vbCrLf &vbCrLf 
 Next 

 mTotalSize=0 
 mBuffer="" 
 Set colMemory = objHW.InstancesOf("Win32_PhysicalMemory") 
 S = S & "記憶體規格" & vbCrLf & "==============================================================" &vbCrLf 
 For Each objMemory in colMemory 
    mBuffer = mBuffer & "====>Bank:" & right(objMemory.Tag,1) & vbTab & vbTab & "容量:" & objMemory.Capacity/1024/1024 & "MB" &vbCrLf 
    mTotalSize = mTotalSize + objMemory.Capacity/1024/1024 
 Next 
 S = S & "總容量:" & mTotalSize & "MB" & vbCrLf 
 S = S & mBuffer & vbCrLf 


 Set colDisk = objHW.InstancesOf("Win32_DiskDrive") 
 Set colPartition = objHW.InstancesOf("Win32_DiskPartition") 
 S = S & "硬碟規格" & vbCrLf & "==============================================================" &vbCrLf 
 For Each objDisk in colDisk 
    S = S & "型號:" & objDisk.Model & vbTab & "容量:" & round(objDisk.Size/1024/1024/1024,2) & "GB" &vbCrLf 
    For Each objPartition in colPartition 
       If objDisk.Index = objPartition.DiskIndex Then 
          S = S & "====>分割區:" & objPartition.Index & vbTab & vbTab & "容量:" & round(objPartition.Size/1024/1024/1024,2) & "GB" &vbCrLf 
       End If 
    Next 
 Next 

 Wscript.Echo S 

 Next