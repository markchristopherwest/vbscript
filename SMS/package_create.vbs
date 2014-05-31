Dim connection
Dim computer
Dim userName
Dim userPassword
Dim password 'Password object

On Error Resume Next

Const EVENT_FAILED = 1
Set objShell = Wscript.CreateObject("Wscript.Shell")
objShell.LogEvent EVENT_FAILED, _
    "SCCM Import Started." _
    & "Additional information is available from http://www.fabrikam.biz."

Wscript.StdOut.Write "Computer you want to connect to (Enter . for local): "
computer = WScript.StdIn.ReadLine

If computer = "." Then
    userName = ""
    userPassword = ""
Else
    Wscript.StdOut.Write "Please enter the user name: "
    userName = WScript.StdIn.ReadLine
    
    Set password = CreateObject("ScriptPW.Password") 
    WScript.StdOut.Write "Please enter your password:" 
    userPassword = password.GetPassword() 
End If
      
Set connection = Connect(computer,userName,userPassword)

If Err.Number<>0 Then
    Wscript.Echo "Call to connect failed"
End If

Dim strPackageName,strPackageDescription,intPackageSourceFlag,strPackageSourcePath
Dim strExistingPackageID, strNewProgramName, strNewProgramComment, intNewMaxRunTime
Dim strSiteCode, strDPServerName
Dim strExistingCollectionID, strExistingProgramName, strNewAdvertisementComment, strNewAdvertisementStartOfferDateTime, strNewAdvertisementStartOfferEnabled
Dim strExistingAdvertisementID, strIncludeSubCollection
'Package Creation Variables
strNewPackageName = "Sample Package " & Now
strNewPackageDescription = "This was created using SCCM SDK."
intNewPackageSourceFlag = 2
strNewPackageSourcePath = "D:\Package_Source\File_Packs\cmXXXXX"
'Program Creation Variables
strNewProgramName = "Sample Installer"
strNewProgramComment = "Sample Installer Comment"
strNewProgramCommandLine = "your_exiting_file.bat or FilePack_Installer.exe /destination /type /logs"
intNewMaxRunTime = 120
'Collection Creation Variables
strExistingParentCollectionID = "CBT0004D"
strNewCollectionName = "Sample Collection" & Now
strNewCollectionComment = "Any variables can be left as a comment here."
strOwnedByThisSite = "CBT"
'strResourceClassName = 
'strResourceID = 
'SWD Distribution Point Variables
strSiteCode = "CBT"
strDPServerName = "WTOPWD03"
'SWD Create Advertisement Variables
'strExistingCollectionID = 
'strExistingProgramName = 
'strNewAdvertisementComment = 
'strNewAdvertisementStartOfferDateTime = 
'strNewAdvertisementStartOfferEnd =
'SWD Assign Assignment to Collection Variables
'strExistingAdvertisementID = 
'strIncludeSubCollection = 

'Some other Object
Dim objPackageID,strConsoleFolder

strConsoleFolder = "FILEPACKS"
intObjectType = 2


'Main
Call CreatePackageInFolder(connection, strNewPackageName, strNewPackageDescription, intNewPackageSourceFlag, strNewPackageSourcePath)
'Call CreateConsoleFolderItem(connection, strInstanceID, intObjectType, nodeID)
Call CreateProgram(connection, strExistingPackageID, strNewProgramName, strNewProgramComment, strNewProgramCommandLine, intNewMaxRunTime)
'Call CreateStaticCollection(connection, existingParentCollectionID, newCollectionName, newCollectionComment, ownedByThisSite, resourceClassName, resourceID)
Call SWDAssignPackageToDistributionPoint(connection, strExistingPackageID, strSiteCode, strDPServerName)
'Call SWDCreateAdvertisement(connection, existingCollectionID, existingPackageID, existingProgramName, newAdvertisementName, newAdvertisementComment, newAdvertisementFlags, newAdvertisementStartOfferDateTime, newAdvertisementStartOfferEnabled)
'Call SWDAssignAdvertisementCollection(connection, existingAdvertisementID, existingCollectionID, includeSubCollection)

Sub CreatePackageInFolder(connection, newPackageName, newPackageDescription, newPackageSourceFlag, newPackageSourcePath)

    ' Create the new package object. 
    Set newPackage = connection.Get("SMS_Package").SpawnInstance_

    ' Populate the new package properties.
    newPackage.Name = newPackageName
    newPackage.Description = newPackageDescription
    newPackage.PkgSourceFlag = newPackageSourceFlag
    newPackage.PkgSourcePath = newPackageSourcePath

    ' Save the package.
    objPackPath=newPackage.Put_

    ' Output the new package name.
    'wscript.echo "Created package: "  & newPackageName & newPackageDescription
	Set Package=connection.Get(objPackPath)
	newPackageID= Package.PackageID	
	wscript.echo "Package: "  & newPackageName & " ID: " & newPackageID
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
	
	wscript.echo "Package Moved to Folder:" & FolderIdentifier
	
End Sub





Sub CreateProgram(connection, existingPackageID, newProgramName, newProgramComment, newProgramCommandLine, newMaxRunTime)
    
    ' Create the new program object.
    Set newProgram = connection.Get("SMS_Program").SpawnInstance_
    
    ' Populate the program properties.
    newProgram.PackageID = existingPackageID
    newProgram.ProgramName = newProgramName
    newProgram.Comment = newProgramComment
    newProgram.CommandLine = newProgramCommandLine
    newProgram.Duration = newMaxRunTime
    
    ' Save the new program and properties.
    newProgram.Put_
    
    ' Output new program name.
    wscript.echo "Created program: " & newProgramName

End Sub

Sub SWDAssignPackageToDistributionPoint(connection, existingPackageID, siteCode, serverName)

    ' Create distribution point object (this is not an actual distribution point).
    Set distributionPoint = connection.Get("SMS_DistributionPoint").SpawnInstance_

    ' Associate the existing package with the new distribution point object.
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
    Wscript.Echo "Assigned package: " & distributionPoint.PackageID & " to DP: " & serverName

End Sub

Sub SWDCreateAdvertisement(connection, existingCollectionID, existingPackageID, existingProgramName, newAdvertisementName, newAdvertisementComment, newAdvertisementFlags, newAdvertisementStartOfferDateTime, newAdvertisementStartOfferEnabled)

    ' Create the new advertisement object.
    Set newAdvertisement = connection.Get("SMS_Advertisement").SpawnInstance_
    
    ' Populate the advertisement properties.
    newAdvertisement.CollectionID = existingCollectionID
    newAdvertisement.PackageID = existingPackageID
    newAdvertisement.ProgramName = existingProgramName
    newAdvertisement.AdvertisementName = newAdvertisementName
    newAdvertisement.Comment = newAdvertisementComment
    newAdvertisement.AdvertFlags = newAdvertisementFlags
    newAdvertisement.PresentTime = newAdvertisementStartOfferDateTime
    newAdvertisement.PresentTimeEnabled = newAdvertisementStartOfferEnabled
    
    ' Save the new advertisement and properties.
    newAdvertisement.Put_ 
    
    ' Output new advertisement name.
    Wscript.Echo "Created advertisement: " & newAdvertisement.AdvertisementName

End Sub

Sub SWDAssignAdvertisementCollection(connection, existingAdvertisementID, existingCollectionID, includeSubCollection)

    ' Get the specific advertisement object.
    Set advertisementToAssign = connection.Get("SMS_Advertisement.AdvertisementID='" & existingAdvertisementID & "'")
    
    ' Fill the advertisement properties for collection.
    advertisementToAssign.CollectionID = existingCollectionID
    advertisementToAssign.IncludeSubCollection = IncludeSubCollection
    
    ' Save the advertisement.
    advertisementToAssign.Put_
    
    ' Output advertisement and collection information.
    Wscript.Echo "Assigned advertisement: " & existingAdvertisementID
    Wscript.Echo "                        " & advertisementToAssign.AdvertisementName
    Wscript.Echo "To collection:          " & advertisementToAssign.CollectionID

End Sub

Sub CreateStaticCollection(connection, existingParentCollectionID, newCollectionName, newCollectionComment, ownedByThisSite, resourceClassName, resourceID)

    ' Create the collection.
    Set newCollection = connection.Get("SMS_Collection").SpawnInstance_
    newCollection.Comment = newCollectionComment
    newCollection.Name = newCollectionName
    newCollection.OwnedByThisSite = ownedByThisSite
    
    ' Save the new collection and save the collection path for later.
    Set collectionPath = newCollection.Put_    
    
   ' Define to what collection the new collection is subordinate.
   ' IMPORTANT: If you do not specify the relationship, the new collection will not be visible in the console. 
    Set newSubCollectToSubCollect = connection.Get("SMS_CollectToSubCollect").SpawnInstance_
    newSubCollectToSubCollect.parentCollectionID = existingParentCollectionID
    newSubCollectToSubCollect.subCollectionID = CStr(collectionPath.Keys("CollectionID"))
    
    ' Save the subcollection information.
    newSubCollectToSubCollect.Put_
        
    ' Create the direct rule.
    Set newDirectRule = connection.Get("SMS_CollectionRuleDirect").SpawnInstance_
    newDirectRule.ResourceClassName = resourceClassName
    newDirectRule.ResourceID = resourceID
    
    ' Add the new query rule to a variable.
    Set newCollectionRule = newDirectRule
    
    ' Get the collection.
    Set newCollection = connection.Get(collectionPath.RelPath)
    
    ' Add the rules to the collection.
    newCollection.AddMembershipRule newCollectionRule

    ' Call RequestRefresh to initiate the collection evaluator. 
    newCollection.RequestRefresh False
    
End Sub

Sub createAdvertisement()
'Create advertisements for install set properties and commit changes 
     Set newAdvertisement = objSWbemServices.Get("SMS_Advertisement").SpawnInstance_() 
 newAdvertisement.AdvertisementName = VarSAMSInstName 
 newAdvertisement.comment  = VarNAME 
 newAdvertisement.CollectionID  = VarCollectionID 
 newAdvertisement.PackageID  = VarPackageID 
 newAdvertisement.ProgramName  = VarInstName 
 newAdvertisement.PresentTime  = datetime 
 'newAdvertisement.ExpirationTime = expdatetime 
 'newAdvertisement.AdvertFlags  = VarAdFlags 
 newAdvertisement.Priority  = VarAdPriority 
 newAdvertisement.RemoteClientFlags = 8 
      newAdvertisement.Put_ 
     wscript.echo "Created Advertisement " +VarSAMSInstName



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
        Wscript.Echo "Couldn't connect: " + Err.Description
        Connect = null
        Exit Function
    End If
    

    ' Determine where the provider is and connect.
    Set providerLoc = swbemServices.InstancesOf("SMS_ProviderLocation")

        For Each location In providerLoc
            If location.ProviderForLocalSite = True Then
                Set swbemServices = swbemLocator.ConnectServer _
                 (location.Machine, "root\sms\site_" + _
                    location.SiteCode,userName,userPassword)
                If Err.Number<>0 Then
                    Wscript.Echo "Couldn't connect:" + Err.Description
                    Connect = Null
                    Exit Function
                End If
                Set Connect = swbemServices
				Wscript.Echo "connected to SMS CIM"
                Exit Function
            End If
        Next
    Set Connect = null ' Failed to connect.
End Function