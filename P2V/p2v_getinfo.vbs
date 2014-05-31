On Error Resume Next	
strComputer = "."

strDateDay = Day(date)
strDateYear = Year(date)
strDateMonth = Month(date)
strDelim = chr(34) & ","
strVersion = "2.0"
strClientName = "YOURNAMEHERE"
strClientDomain = "YOURDOMAINHERE"
strAuthorName = "markchristopherwest@gmail.com"
strFileName = "P2V_" & strComputer & "_" & strDateMonth & strDateDay & strDateYear & "_PREP.CSV"
'Wscript.Echo strFileName
strFileOutput = strFileName
' Constants for FileSystemObject
Const FOR_READING = 1
Const FOR_WRITING = 2
Const FOR_APPENDING = 8
' Create a Script Runtime FileSystemObject.
Set objFSO = CreateObject("Scripting.FileSystemObject")
' Check to see if the output file exists. If so, open it for writing or appending.
' If not, create it and open it for writing.
If objFSO.FileExists(strFileOutput) Then
  Set objOutputFile = objFSO.OpenTextFile (strFileOutput, FOR_WRITING)
Else
  Set objOutputFile = objFSO.CreateTextFile(strFileOutput)
End If
If Err <> 0 Then
  Wscript.Echo "Unable to open " & strFileOutput & " for output."
  WScript.Quit
End If


'Bind to WMI
    Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

'Test for success in binding to WMI.

    If Err = 0 Then
	objOutputFile.Write chr(34) & "P2V Prerequisite Checker v." & strVersion & chr(34)
	objOutputFile.Writeline
	objOutputFile.Write chr(34) & "Custom Script Developed for: " & strClientName & chr(34)
	objOutputFile.Writeline
	objOutputFile.Write chr(34) & "Questions?  Comments?  Contact: " & strAuthorName & chr(34)
	objOutputFile.Writeline
	objOutputFile.Write chr(34) & "This log file was generated: " & Now(date) & chr(34) 
	objOutputFile.Write
      
    Else
      objOutputFile.WriteLine VbCrLf & "Error: Unable to connect to WMI." & VbCrLf & _
       "Error Number: " & Err.Number & VbCrLf & _
       "Error Source: " & Err.Source & VbCrLf & _
       "Error Description: " & Err.Description
      Err.Clear
    End If

' List Groups

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_Group  Where LocalAccount = True")
	objOutputFile.Writeline                                    
	objOutputFile.Write chr(34) & "Windows Local Group Information for:" & strComputer & chr(34) 
	objOutputFile.Writeline
    objOutputFile.Write chr(34) & "Caption:" & chr(34) & ","
    objOutputFile.Write chr(34) & "Description:"& chr(34) & ","
    objOutputFile.Write chr(34) & "Domain:" & chr(34) & ","
    objOutputFile.Write chr(34) & "Local Account:" & chr(34) & ","
    objOutputFile.Write chr(34) & "Name:" & chr(34) & ","
    objOutputFile.Write chr(34) & "SID:" & chr(34) & ","
    objOutputFile.Write chr(34) & "SID Type:" & chr(34) & ","
    objOutputFile.Write chr(34) & "Status:" & chr(34)
    objOutputFile.Writeline

For Each objItem in colItems
    objOutputFile.Write chr(34) & objItem.Caption & chr(34) & ","
    objOutputFile.Write chr(34) & objItem.Description & chr(34) & ","
    objOutputFile.Write chr(34) & objItem.Domain & chr(34) & ","
    objOutputFile.Write chr(34) & objItem.LocalAccount & chr(34) & ","
    objOutputFile.Write chr(34) & objItem.Name & chr(34) & ","
    objOutputFile.Write chr(34) & objItem.SID & chr(34) & ","
    objOutputFile.Write chr(34) & objItem.SIDType & chr(34) & ","
    objOutputFile.Write chr(34) & objItem.Status & chr(34)
    objOutputFile.Writeline
Next

' List Users

	objOutputFile.Writeline                                    
	objOutputFile.Write chr(34) & "Windows Local User Account Information for:" & strComputer & chr(34) 
	objOutputFile.Writeline
	objOutputFile.Write chr(34) & "AccountType: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Caption: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Description: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Disabled: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Domain: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "FullName: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "LocalAccount: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Lockout: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Name: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "PasswordChangeable: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "PasswordExpires: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "PasswordRequired: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "SID: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "SIDType: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Status: "  & chr(34) 
	objOutputFile.Writeline     
		
		Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_UserAccount", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)                                        
   For Each objItem In colItems
      objOutputFile.Write chr(34) & objItem.AccountType  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.Caption  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.Description  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.Disabled  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.Domain  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.FullName  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.LocalAccount  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.Lockout  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.Name  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.PasswordChangeable  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.PasswordExpires  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.PasswordRequired  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.SID  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.SIDType  & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.Status  & chr(34) & "," 
      objOutputFile.Writeline
   Next

'Network Configuration      
       	objOutputFile.Writeline                                    
		objOutputFile.Write chr(34) & "Win32_NetworkAdapterConfiguration CIM for:" & strComputer & chr(34) 
		objOutputFile.Writeline
		objOutputFile.Write chr(34) & "Caption: "  & chr(34) & ","
		objOutputFile.Write chr(34) & "Description: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "DHCPEnabled: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "DHCPLeaseExpires: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "DHCPLeaseObtained: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "DHCPServer: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "DNSDomain: "  & chr(34) & ","
		objOutputFile.Write chr(34) & "DNSEnabledForWINSResolution: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "DNSHostName: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "DNSServerSearchOrder: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "DomainDNSRegistrationEnabled:"  & chr(34) & "," 		
		objOutputFile.Write chr(34) & "FullDNSRegistrationEnabled: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "IPAddress: "  & chr(34) & ","
		objOutputFile.Write chr(34) & "IPConnectionMetric: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "IPEnabled: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "IPSubnet: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "MACAddress: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "WINSPrimaryServer: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "WINSScopeID: "  & chr(34) & "," 
		objOutputFile.Write chr(34) & "WINSSecondaryServer: "  & chr(34)
		objOutputFile.Writeline

	
		Set colNicConfigs = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)  
		For Each objItem In colNicConfigs
			objOutputFile.Write chr(34) & objItem.Caption  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.Description  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.DHCPEnabled  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.DHCPLeaseExpires  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.DHCPLeaseObtained  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.DHCPServer  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.DNSDomain  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.DNSEnabledForWINSResolution  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.DNSHostName  & chr(34) & ","

			objOutputFile.Write chr(34) & strIPAddress  & chr(34) & ","
			For Each strIPAddress In objNicConfig.IPAddress
				objOutputFile.WriteLine "        " & strIPAddress
			Next
			objOutputFile.Write chr(34) & objItem.IPConnectionMetric  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.IPEnabled  & chr(34) & ","
			For Each strIPAddress In objNicConfig.IPAddress
				objOutputFile.WriteLine "        " & strIPAddress
			Next


			objOutputFile.Write chr(34) & strIPSubnet  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.Name  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.MACAddress  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.WINSPrimaryServer  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.WINSScopeID  & chr(34) & ","
			objOutputFile.Write chr(34) & objItem.WINSSecondaryServer  & chr(34)
			objOutputFile.Writeline		 
			
      Next
      
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Service", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)
      objOutputFile.Writeline                                    
      objOutputFile.Write chr(34) & "Win32_Service CIM for:" & strComputer & chr(34) 
      objOutputFile.Writeline
      objOutputFile.Write chr(34) & "Name: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "AcceptPause: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "AcceptStop: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "Caption: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "CheckPoint: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "CreationClassName: "  & chr(34) & "," 
      'objOutputFile.Write chr(34) & "Description: "  & chr(34) & "," 
      'objOutputFile.Write chr(34) & "DesktopInteract: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "DisplayName: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "ErrorControl: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "ExitCode: " &  chr(34) & "," 
      'objOutputFile.Write chr(34) & "InstallDate: "  & chr(34) & "," 
      'objOutputFile.Write chr(34) & "Name: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "PathName: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "ProcessId: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "ServiceSpecificExitCode: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "ServiceType: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "Started: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "StartMode: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "StartName: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "State: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "Status: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "SystemCreationClassName: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "SystemName: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "TagId: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "WaitHint: "  & chr(34)
      objOutputFile.Writeline

   For Each objItem In colItems
      objOutputFile.Write chr(34) & objItem.Name & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.AcceptPause & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.AcceptStop & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.Caption & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.CheckPoint & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.CreationClassName & chr(34) & "," 
      'objOutputFile.Write chr(34) & objItem.Description & chr(34) & "," 
      'objOutputFile.Write chr(34) & objItem.DesktopInteract & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.DisplayName & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.ErrorControl & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.ExitCode & chr(34) & "," 
      'objOutputFile.Write chr(34) & WMIDateStringToDate(objItem.InstallDate) & chr(34) & "," 
      'objOutputFile.Write chr(34) & objItem.Name & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.PathName & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.ProcessId & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.ServiceSpecificExitCode & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.ServiceType & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.Started & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.StartMode & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.StartName & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.State & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.Status & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.SystemCreationClassName & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.SystemName & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.TagId & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.WaitHint & chr(34)
      objOutputFile.Writeline
Next


'Query Win32_LogicalDisk CIM

   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDisk", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)
      objOutputFile.Writeline                                    
      objOutputFile.Write chr(34) & "Win32_LogicalDisk CIM for:" & strComputer & chr(34) 
      objOutputFile.Writeline
      objOutputFile.Write chr(34) & "Caption: " & chr(34) & "," 
      objOutputFile.Write chr(34) & "Compressed: " & chr(34) & ","     
      objOutputFile.Write chr(34) & "Description: " & chr(34) & "," 
      objOutputFile.Write chr(34) & "DeviceID: " & chr(34) & "," 
      objOutputFile.Write chr(34) & "DriveType: " & chr(34) & "," 
      objOutputFile.Write chr(34) & "FileSystem: " & chr(34) & "," 
      objOutputFile.Write chr(34) & "FreeSpace: " & chr(34) & "," 
      objOutputFile.Write chr(34) & "MediaType: " & chr(34) & ","
      objOutputFile.Write chr(34) & "QuotasDisabled:"   & chr(34) & ","
      objOutputFile.Write chr(34) & "QuotasIncomplete:"   & chr(34) & ","
      objOutputFile.Write chr(34) & "QuotasRebuilding:"   & chr(34) & ","
      objOutputFile.Write chr(34) & "Size:"   & chr(34) & ","
      objOutputFile.Write chr(34) & "VolumeDirty: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "VolumeName: "  & chr(34) & "," 
      objOutputFile.Write chr(34) & "VolumeSerialNumber: " & chr(34) & "," 
      objOutputFile.Writeline

   For Each objItem In colItems

      objOutputFile.Write chr(34) & objItem.Caption & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.Compressed & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.Description & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.DeviceID & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.DriveType & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.FileSystem & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.FreeSpace & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.MediaType & chr(34) & ","
      objOutputFile.Write chr(34) & objItem.QuotasDisabled & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.QuotasIncomplete & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.QuotasRebuilding & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.Size & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.VolumeDirty & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.VolumeName & chr(34) & "," 
      objOutputFile.Write chr(34) & objItem.VolumeSerialNumber & chr(34)
      objOutputFile.Writeline
      
   Next

' Get Win32_DiskDrive CIM 
    objOutputFile.Writeline                                    
    objOutputFile.Write chr(34) & "Win32_DiskDrive CIM for:" & strComputer & chr(34) 
    objOutputFile.Writeline
	objOutputFile.Write chr(34) & "Availability: " & chr(34) & "," 
	objOutputFile.Write chr(34) & "BytesPerSector: " & chr(34) & "," 
	objOutputFile.Write chr(34) & "CapabilityDescriptions: " & chr(34) & "," 
	objOutputFile.Write chr(34) & "Caption: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "CompressionMethod: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "ConfigManagerErrorCode: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "ConfigManagerUserConfig: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "CreationClassName: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "DefaultBlockSize: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Description: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "DeviceID: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "ErrorCleared: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "ErrorDescription: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "ErrorMethodology: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Index: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "InstallDate: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "InterfaceType: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "LastErrorCode: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Manufacturer: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "MaxBlockSize: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "MaxMediaSize: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "MediaLoaded: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "MediaType: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "MinBlockSize: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Model: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Name: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "NeedsCleaning: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "NumberOfMediaSupported: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Partitions: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "PNPDeviceID: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "PowerManagementCapabilities: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "PowerManagementSupported: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "SCSIBus: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "SCSILogicalUnit: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "SCSIPort: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "SCSITargetId: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "SectorsPerTrack: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Size: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "Status: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "StatusInfo: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "SystemCreationClassName: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "SystemName: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "TotalCylinders: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "TotalHeads: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "TotalSectors: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "TotalTracks: "  & chr(34) & "," 
	objOutputFile.Write chr(34) & "TracksPerCylinder: "  & chr(34)
	objOutputFile.Writeline

Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskDrive",,48)
	
For Each objItem in colItems
	objOutputFile.Write chr(34) & objItem.Availability & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.BytesPerSector & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.CapabilityDescriptions & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.Caption & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.CompressionMethod & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.ConfigManagerErrorCode & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.ConfigManagerUserConfig & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.CreationClassName & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.DefaultBlockSize & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.Description & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.DeviceID & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.ErrorCleared & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.ErrorDescription & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.ErrorMethodology & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.Index & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.InstallDate & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.InterfaceType & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.LastErrorCode & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.Manufacturer & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.MaxBlockSize & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.MaxMediaSize & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.MediaLoaded & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.MediaType & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.MinBlockSize & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.Model & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.Name & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.NeedsCleaning & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.NumberOfMediaSupported & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.Partitions & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.PNPDeviceID & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.PowerManagementCapabilities & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.PowerManagementSupported & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.SCSIBus & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.SCSILogicalUnit & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.SCSIPort & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.SCSITargetId & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.SectorsPerTrack & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.Size & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.Status & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.StatusInfo & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.SystemCreationClassName & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.SystemName & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.TotalCylinders & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.TotalHeads & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.TotalSectors & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.TotalTracks & chr(34) & "," 
	objOutputFile.Write chr(34) & objItem.TracksPerCylinder & chr(34)
	objOutputFile.Writeline
	
	Next


   
   
   
   
objLogFile.Close