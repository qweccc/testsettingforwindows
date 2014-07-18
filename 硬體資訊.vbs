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
 S = "�q���W��" & vbTab & strComputer & vbCrLf  & "==============================================================" &vbCrLf 
 Set colCS = objHW.InstancesOf("Win32_computerSystem") 
 S = S & "�n�J�ϥΪ�" & vbTab 
 For Each objCS in colCS 
   S = S & objCS.Username 
 Next 
 S = S & vbCrLf & "==============================================================" &vbCrLf 


 S = S & "CPU�W��" & vbCrLf & "==============================================================" &vbCrLf 
 For Each objCPU in colCPU 
    S = S & "CPU��}:" & right(objCPU.DeviceID,1) & vbTab & vbTab & "����:"& trim(objCPU.Name) &vbCrLf &vbCrLf 
 Next 

 mTotalSize=0 
 mBuffer="" 
 Set colMemory = objHW.InstancesOf("Win32_PhysicalMemory") 
 S = S & "�O����W��" & vbCrLf & "==============================================================" &vbCrLf 
 For Each objMemory in colMemory 
    mBuffer = mBuffer & "====>Bank:" & right(objMemory.Tag,1) & vbTab & vbTab & "�e�q:" & objMemory.Capacity/1024/1024 & "MB" &vbCrLf 
    mTotalSize = mTotalSize + objMemory.Capacity/1024/1024 
 Next 
 S = S & "�`�e�q:" & mTotalSize & "MB" & vbCrLf 
 S = S & mBuffer & vbCrLf 


 Set colDisk = objHW.InstancesOf("Win32_DiskDrive") 
 Set colPartition = objHW.InstancesOf("Win32_DiskPartition") 
 S = S & "�w�гW��" & vbCrLf & "==============================================================" &vbCrLf 
 For Each objDisk in colDisk 
    S = S & "����:" & objDisk.Model & vbTab & "�e�q:" & round(objDisk.Size/1024/1024/1024,2) & "GB" &vbCrLf 
    For Each objPartition in colPartition 
       If objDisk.Index = objPartition.DiskIndex Then 
          S = S & "====>���ΰ�:" & objPartition.Index & vbTab & vbTab & "�e�q:" & round(objPartition.Size/1024/1024/1024,2) & "GB" &vbCrLf 
       End If 
    Next 
 Next 

 Wscript.Echo S 

 Next