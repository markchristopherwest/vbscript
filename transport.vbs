' Name : Transport Package Between SCCM Sites
' Role : Configuration Management
' Note : You will need to customize this for your environment
' Dude : markchristopherwest@gmail.com

'Keep Going
On Error Resume Next

'Set FSO
Set FSO =CreateObject("scripting.FileSystemObject")

'Event Log Variables
Const EVENT_FAILED = 1
Set objShell = Wscript.CreateObject("Wscript.Shell")
	objShell.LogEvent EVENT_FAILED, _
		"SCCM Import Started." _
		& "Additional information is available from http://www.markchristopherwest.com."
		
'Core Assignment Variables 
Dim strNewPackageName,strNewPackageDescription,intNewPackageSourceFlag,strNewPackageSourcePath,strNewPackageVendor,strNewPackageVersion
Dim strExistingPackageID, strExistingPackageDirectory, strNewProgramName, strNewProgramCommandLine, strNewProgramComment, strNewProgramFlags, intNewMaxRunTime
	
Dim net
Dim localConnection
Dim swbemLocator
Dim swbemServices
Dim providerLoc
Dim location

'Work for this 
Dim strSiteServer, strSiteCode, strStandingPackageFolderID
Dim strDPServerName
Dim strExistingFilePackageID

	'Service Account Credentials
		siteSystem_CBT = "SOURCE"
		siteSystem_CDA = "TARGET"
		siteSystem_CDA_userName = "DOMAIN\USER
		siteSystem_CDA_userPassword = "PASSWORD"
		siteSystem_CBT_userName = ""
		siteSystem_CBT_userPassword = ""
		siteSystem_CBT_siteCode= "CBT"
		siteSystem_CDA_siteCode= "CDA"

	'Set Connections to Each Site
		Set connection_X = Connect(siteSystem_CBT,siteSystem_CBT_userName,siteSystem_CBT_userPassword)
		Set connection_Y = Connect(siteSystem_CDA,siteSystem_CDA_userName,siteSystem_CDA_userPassword)


Dim newPackageID,strConsoleFolder,objPackPath,objPackItem
Dim strPackName,strPackType


'Name of Folder in the UI to Place this In
strConsoleFolder = "FilePackage"

strExistingFilePackageID = Wscript.Arguments(0)

Wscript.Echo "The PackageID to be Transported is: " & strExistingFilePackageID 

		'Get the Info from Source Tier & Populate Variables
		strExistingProgramNameToModify = "FilePackage Installer"
		'Get the Package Info
		Call QueryPackProgInfo(connection_X, strExistingFilePackageID, strExistingProgramNameToModify)
		'Purge any Packages with the Same name to Avoid Duplicates
		Call DeletePackagesPreTransport(connection_Y,strNewPackageName,strNewPackageVersion)
		'Copy the PackageSource From Source to Destination
		'Purge Source  on Remote Computer
		Call DeleteThat(siteSystem_CDA,strExistingPackageDirectory)
		'Copy The Source
		Call CopyThis("\\wtopwd03.cbt.tcgic1.com\FilePackage\" & strNewPackageName & "\fp_" & strNewPackageName & "_" & strNewPackageVersion,"\\cbdpwa99.cb.tcgic.com\FilePackage\" & strNewPackageName & "\fp_"  & strNewPackageName & "_" & strNewPackageVersion)
		'Create the New Package
		Call CreatePackageInFolder(connection_Y, strNewPackageName, strNewPackageVendor, strNewPackageVersion, strNewPackageDescription, intNewPackageSourceFlag, strNewPackageSourcePath)
		'Create the Program
		Call CreateProgram(connection_Y, strExistingPackageID, strNewProgramName, strNewProgramComment, strNewProgramCommandLine, strNewProgramFlags, intNewMaxRunTime)
		'Assign the Program to the Primary DP
		strSitefCode
		Call SWDAssignPackageToDistributionPoint(connection_Y, strExistingPackageID, "PDA", "CBDPWS99")
		'Assign the Program to the Secondary DP
		Call SWDAssignPackageToDistributionPoint(connection_Y, strExistingPackageID, "PWN", "WTOPWD04")
		'Assign the Program to the International Asia DP
		Call SWDAssignPackageToDistributionPoint(connection_Y, strExistingPackageID, "SAZ", "APZPWA10")
		'Assign the Program to the International Canada Secondary DP
		Call SWDAssignPackageToDistributionPoint(connection_Y, strExistingPackageID, "SCZ", "TORPWA01")
		'Assign the Program to the International Europe Secondary DP
		Call SWDAssignPackageToDistributionPoint(connection_Y, strExistingPackageID, "SEZ", "EUZAS095")
		'Assign the Program to the International Latin Secondary DP
		Call SWDAssignPackageToDistributionPoint(connection_Y, strExistingPackageID, "SLZ", "SRLPWA01")
			
		Wscript.Echo strExistingPackageID & " is READY for Deployment/Transport on the P-Side or PDA/PWN"
		

Sub PreparePackageSource(PackageName,PackageType)
	'Create the Parent Folder if it Does Not Exist
	thisFilePackageParent = "D:\Package_Source\FilePackage\" & PackageName & "\"
	If  Not FSO.FolderExists(thisFilePackageParent) Then
	   newfolder = FSO.CreateFolder (thisFilePackageParent)
	   Wscript.Echo "A new folder '" & newfolder & "' has been created - FilePackage required this folder to continue processing..." 
	Else
	   Wscript.Echo "The folder '" & thisFilePackageParent & "' already exists on the target... no need to create a new one!"
	
	End If
	
	
End Sub




Sub CreatePackageInFolder(connection, newPackageName, strNewPackageVendor, strNewPackageVersion, newPackageDescription, newPackageSourceFlag, newPackageSourcePath)
	'Wscript.Echo("Initiating Package Creation!<br/>")    
    ' Create the new package object. 
    Set newPackage = connection.Get("SMS_Package").SpawnInstance_

    ' Populate the new package properties.
    newPackage.Name = newPackageName
	newPackage.Manufacturer = strNewPackageVendor
	newPackage.Version = strNewPackageVersion
	newPackage.Description = newPackageDescription
    newPackage.PkgSourceFlag = newPackageSourceFlag
    newPackage.PkgSourcePath = newPackageSourcePath

    ' Save the package.
	'Wscript.Echo("Creating Package: " & strNewPackageVendor & " " & newPackageName & " " & newPackageVersion & "<br/>")    
    Set newPackagePath = newPackage.Put_
    newPackageID = newPackagePath.Keys("PackageID")
	
	
    ' Output the new package name.
	'Wscript.Echo("Package ID was: " & newPackageID & "<br/>")
	strExistingPackageID = newPackageID
	'get Folder Identifier
	For Each objFolder in connection.InstancesOf("SMS_ObjectContainerNode")
	if objFolder.Name=strConsoleFolder and objFolder.ObjectType=2 Then
	FolderIdentifier=objFolder.ContainerNodeID
	End If
	Next

	'move the package to specific folder
	Set objNewConsoleFolderItem = connection.Get _
	("SMS_ObjectContainerItem").SpawnInstance_()

	objNewConsoleFolderItem.InstanceKey = newPackageID
	objNewConsoleFolderItem.ObjectType=2 ' Package Node.
	objNewConsoleFolderItem.ContainerNodeID = FolderIdentifier 'Container
	objNewConsoleFolderItem.Put_
	
	'Wscript.Echo("Package Moved to Admin GUI Folder: " & FolderIdentifier & " aka " & strConsoleFolder & "<br/>")
	
End Sub

Sub CreateProgram(connection, existingPackageID, newProgramName, newProgramComment, newProgramCommandLine, NewProgramFlags, newMaxRunTime)
    
    ' Create the new program object.
    Set newProgram = connection.Get("SMS_Program").SpawnInstance_
    
    ' Populate the program properties.
    newProgram.PackageID = existingPackageID
    newProgram.ProgramName = newProgramName
    newProgram.Comment = newProgramComment
    newProgram.CommandLine = newProgramCommandLine
    newProgram.Duration = newMaxRunTime
	newProgram.ProgramFlags = newProgramFlags
    ' Save the new program and properties.
    newProgram.Put_
    ' Output new program name.
    'Wscript.Echo "Created program: " & newProgramName & "<br/>"
	
End Sub



Sub QueryPackProgInfo(connection, existingPackageID, existingProgramNameToModify)
	
	'Create the Package Object
	Set objPackage = connection.Get("SMS_Package.PackageID='" & existingPackageID & "'")
	
     ' Load the specific program to change (programname is a key value and must be unique).
    Set objProgram = connection.Get("SMS_Program.PackageID='" & existingPackageID & "'" & ",ProgramName='" & existingProgramNameToModify & "'") 
	
	' Output Package Information
    Wscript.Echo "Existing Package Name		: " & objPackage.Name & "<br/>"
	Wscript.Echo "Existing Package Vendor		: " & objPackage.Manufacturer & "<br/>"
	Wscript.Echo "Existing Package Flag  	 	: " & objPackage.PkgSourceFlag & "<br/>"
	Wscript.Echo "Existing Package Source 	: " & objPackage.PkgSourcePath & "<br/>"
	Wscript.Echo "Existing Package Description: " & objPackage.Description & "<br/>"
		
	' Output Program Information
    'Wscript.Echo "Existing Program Name		: " & objProgram.ProgramName & "<br/>"
	'Wscript.Echo "Existing CMD Line			: " & objProgram.CommandLine & "<br/>"
	'Wscript.Echo "Existing Program Comment	: " & objProgram.Comment & "<br/>"
	'Wscript.Echo "Existing Program Duration	: " & objProgram.Duration & "<br/>"
	
	' Set All PackProg Variables for CreatePackageInFolder & CreateProgram
	
	strNewPackageName = objPackage.Name
	strNewPackageDescription = objPackage.Description
	strNewPackageVendor = objPackage.Manufacturer
	strNewPackageVersion = objPackage.Version
	intNewPackageSourceFlag = objPackage.PkgSourceFlag
	strNewPackageSourcePath = objPackage.PkgSourcepath
	strNewProgramName = objProgram.ProgramName
	strNewProgramCommandLine = objProgram.CommandLine
	strNewProgramComment = objProgram.Comment
	strNewProgramFlags = objProgram.ProgramFlags
	intNewMaxRunTime = objProgram.Duration
	
	' Output Completion
	'Wscript.Echo "Pre-Transport Variables Have Been Determined<br/><br/>"
	

	
End Sub

Sub SWDAssignPackageToDistributionPoint(connection, existingPackageID, siteCode, serverName)

    ' Create distribution point object (this is not an actual distribution point).
    Set distributionPoint = connection.Get("SMS_DistributionPoint").SpawnInstance_
    '	 Associate the existing package with the new distribution point object.
    distributionPoint.PackageID = existingPackageID     
    ' This query selects a single distribution point based on the provided SiteCode and ServerName.
    query = "SELECT * FROM SMS_SystemResourceList WHERE RoleName='SMS Distribution Point' AND SiteCode='" & siteCode & "' AND ServerName='" & serverName & "'"
    Set listOfResources = connection.ExecQuery(query, , wbemFlagForwardOnly Or wbemFlagReturnImmediately)
    ' The query returns a collection that needs to be enumerated (although we should only get one instance back).
    For Each resource In ListOfResources      
        distributionPoint.ServerNALPath = Resource.NALPath
        distributionPoint.SiteCode = Resource.SiteCode        
    Next
	' Save the distribution point instance for the package.
    distributionPoint.Put_ 
    ' Display notification text.
    'Wscript.Echo "Assigned package: " & distributionPoint.PackageID & " to DP: " & serverName & "<br/>"
    

End Sub	

	Sub CopyThis(Source,Destination)
		intMonth = Month(Now())
		intYear = Year(Now())
		intDay = Day(Now())

		If intMonth < 10 Then
			intMonth = "0" & intMonth
		End If
			
		If intDay < 10 Then
			intDay = "0" & intDay
		End If

		Wscript.Echo "SOURCE: " & Source
		Wscript.Echo "DESTINATION: " & Destination
		'thisCMD = "psexec.exe \\" & strComputer & " -u	""" & siteSystem_CDA_userName & """	 -p " & siteSystem_CDA_userPassword & " schtasks.exe /create /z /tn """ & strFilePackageName & """	/tr """ & strSchTasksCmd & """  /sc once /st " & strTaskTime & " /sd " & strTaskDate & " /ru """ & siteSystem_CDA_userName & """	 /rp " & siteSystem_CDA_userPassword & ""	
		
		
		'sCopyCommand = "psexec.exe " & strComputer & " -u	""" & siteSystem_CBT_userName & """	 -p " & siteSystem_CBT_userPassword & " robocopy.exe " & Source & """ """ & Destination & """ /E /S /V /LOG+:d:\Package_Source\FilePackage\FilePackage_Logs\FilePackage_TRANSPORT_" & UCASE(strNewPackageName) & "_" & UCASE(strNewPackageVersion) & ".log"
		sCopyCommand = "robocopy.exe " & Source & " " & Destination & " /E /S /V /LOG+:d:\Package_Source\FilePackage\FilePackage_Logs\Transport\FilePackage_TRANSPORT_" & intYear & "-" & intMonth & "-" & intDay  & "_" & UCASE(strNewPackageName) & "_" & UCASE(strNewPackageVersion) & ".log"
		Wscript.Echo "COPYcmd: " & sCopyCommand
		Set objShell = CreateObject("WScript.Shell")
		Set objWshScriptExec = objShell.Exec(sCopyCommand)
		
		Set objStdOut = objWshScriptExec.StdOut
		
		While Not objStdOut.AtEndOfStream
		strLine = objStdOut.ReadLine 
		Wscript.Echo strLine
		Wend
		Set objShell = Nothing
		
		
		
	End Sub
	
Sub DeletePackagesPreTransport(connection,package_name,version_name)

			' SMS_Package object
			dim oPackage
			dim oPackageSet
			' Package of SMS_Package objects
			dim strQuery
			' String for the Query
			'Check to see if returned No Records
					
			strQuery = "select * from SMS_Package where Name = '" & package_name & "'"	& " and Version = '" & version_name & "'" 
			set colObjPackages = connection.ExecQuery(strQuery)
			if colObjPackages.Count = 0 then
			
				'Wscript.Echo "1 Site Searched for: " & package_name

			else
				' Write each one to an <option> tag for the dropdown list
				For Each objPackage In colObjPackages
				Wscript.Echo "Found PackageID: " & objPackage.PackageID & " = " & objPackage.Manufacturer & " " & objPackage.Name & " " & objPackage.Version
		
			' Get the specified package instance (passed in as existingPackageID).
			Set packageToDelete = connection.Get("SMS_Package.PackageID='" & objPackage.PackageID & "'")
			
			' Delete the package.
			PackageToDelete.Delete_
			
			' Output package ID of deleted package.
			Wscript.Echo "Deleted Package ID: " & objPackage.PackageID
			strExistingPackageDirectory = objPackage.PkgSourcePath
			Wscript.Echo "Cleanup Directory: " & strExistingPackageDirectory
			
			

			Next
				end if


	End Sub

	Sub DeleteThat(targetHost,targetDirectory)

		Wscript.Echo "Target Host: " & targetHost
		Wscript.Echo "Target Path: " & targetDirectory
		thisCMD = "psexec.exe -accepteula \\" & targetHost & " -u	""" & siteSystem_CDA_userName & """	 -p " & siteSystem_CDA_userPassword & " cmd.exe /c ""rd /s /q " & targetDirectory &""""
		delCommand = thisCMD
		Wscript.Echo "DELcmd: " & delCommand
		Set objShell = CreateObject("WScript.Shell")
			Set objWshScriptExec = objShell.Exec(delCommand)
		
		Set objStdOut = objWshScriptExec.StdOut
		
		While Not objStdOut.AtEndOfStream
		strLine = objStdOut.ReadLine 
		Wscript.Echo strLine
		Wend
		Set objShell = Nothing
		
		
		
	End Sub

	
Function Connect(server, userName, userPassword)

    On Error Resume Next

    Dim net
    Dim localConnection
    Dim swbemLocator
    Dim swbemServices
    Dim providerLoc
    Dim location

    Set swbemLocator = CreateObject("WbemScripting.SWbemLocator")

    swbemLocator.Security_.AuthenticationLevel = 6 'Packet Privacy

    ' If  the server is local, don't supply credentials.
    Set net = CreateObject("WScript.NetWork")
    If UCase(net.ComputerName) = UCase(server) Then
        localConnection = true
        userName = ""
        userPassword = ""
        server = "."
    End If

    ' Connect to the server.
    Set swbemServices= swbemLocator.ConnectServer _
            (server, "root\sms",userName,userPassword)
    If Err.Number<>0 Then
        'Wscript.Echo "Couldn't connect: " + Err.Description
		Wscript.Echo("Couldn't connect to SCCM SDK! " & Err.Description & " ")
        Connect = null
		Stop
        Exit Function

		Else
		Wscript.Echo("Script has Connected to SCCM SDK Service on " & server & "")
    End If


    ' Determine where the provider is and connect.
    Set providerLoc = swbemServices.InstancesOf("SMS_ProviderLocation")

        For Each location In providerLoc
            If location.ProviderForLocalSite = True Then
                Set swbemServices = swbemLocator.ConnectServer _
                 (server, "root\sms\site_" + _
                    location.SiteCode,userName,userPassword)
                If Err.Number<>0 Then
                    'Wscript.Echo "Couldn't connect:" + Err.Description
					Wscript.Echo("Connection Error: " & Err.Description & "#" & Err.Number &" ")
                    Connect = Null
                    Exit Function
                End If
                Set Connect = swbemServices
                Exit Function
            End If
        Next
    Set Connect = null ' Failed to connect.
End Function

Function ConvertToWMIDate(strDate)
	'Convert from a standard date time to wmi date
	'4/18/2005 11:30:00 AM = 2005041811300.000000+***
	strYear = year(strDate):strMonth = month(strDate)
	strDay = day(strDate):strHour = hour(strDate)
	strMinute = minute(strDate)
	'Pad single digits with leading zero
	if len(strmonth) = 1 then strMonth = "0" & strMonth
	if len(strDay) = 1 then strDay = "0" & strDay
	if len(strHour) = 1 then strHour = "0" & strHour
	if len(strMinute) = 1 then strMinute = "0" & strMinute
	ConvertToWMIDate = strYear & strMonth & strDay & strHour _
	& strMinute & "00.000000+***"
End Function
	
