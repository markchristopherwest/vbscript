<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<html>
<body>
 <form method="post" action="filepack_import.asp">
      <table border="0">
        <tr>
          <td width="39"><select name="ChangeID" size="10" style="font-family:Tahoma; font-stretch:expanded; font-weight:300; font-size:14px; color:#FFFFFF; background-color:#000000; border-color:#272726; border-style:solid;">
              <%
    'Declare variables
	
	
Dim connection
Dim computer
Dim userName
Dim userPassword
Dim password 'Password object

Dim strPackageName,strPackageDescription,intPackageSourceFlag,strPackageSourcePath
Dim strExistingPackageID, strNewProgramName, strNewProgramComment, intNewMaxRunTime
Dim strSiteCode, strDPServerName
Dim strExistingCollectionID, strExistingProgramName, strNewAdvertisementComment, strNewAdvertisementStartOfferDateTime, strNewAdvertisementStartOfferEnabled
Dim strExistingAdvertisementID, strIncludeSubCollection

On Error Resume Next

'Wscript.StdOut.Write "Computer you want to connect to (Enter . for local): "
'computer = WScript.StdIn.ReadLine
computer = "."


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
    'Wscript.Echo "Call to connect failed"
	Response.Write("Call to connect failed! <br />")
End If

	
	
    Dim fso, ofile, ldata
    'Open the text data file
    Set fso = CreateObject("Scripting.FileSystemObject") 
    Set ofile = fso.OpenTextFile("d:\SCCMWBEM\change\index_change_inbox.txt", 1)
    'Read the file contents line by line and
    '     add to listbox
    Do While Not ofile.AtEndOfStream 
    ldata = ofile.ReadLine
    Response.Write "<OPTION VALUE=""" & ldata & """>" & ldata
    Loop
    oFile.Close
    Set ofile=nothing
    Set fso=nothing
    %>
          </select></td>
          <td width="326"><p>Select the File Pack to Import from Change Inbox</p>
              <p>&nbsp;</p>
            <p>
                <label>
                <input type="submit" value="Submit" />
                </label>
            </p></td>
        </tr>
      </table>
        </form>
<%

On Error Resume Next

Set FSO = CreateObject("Scripting.FileSystemObject")

Dim FilePackName, FilePackType, FilePackDrive

Sub CopySource(ChangeInboxSource,PackageSource)
	const SUCCESS = 0
	const WARNING = 2
	On Error Resume Next
	Response.Write("Source Folder:" & ChangeInboxSource & " <br />")
	Response.Write("Target Folder:" & PackageSource & " <br />")
	Set FSO =CreateObject("scripting.FileSystemObject")
	FSO.CopyFolder ChangeInboxSource, PackageSource ,True
	If Err.Number=0 Then 
	    objShell.LogEvent SUCCESS, file.path & " copied successuflly."  
	  	Response.Write(file.path & " copied successuflly.<br />")
	Else
	   objShell.LogEvent WARNING, file.path & " failed to copy."
	 
	  	Response.Write(file.path & " failed to copy." & "<br />")
		End IF
		Response.Write("Copy Complete." & "<br />")
End Sub


Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
            Response.Write(Folder & " was found to contain folder " & Subfolder.Name & "<br />")
			cmParent = Folder.Path
			cmChild = Subfolder.Path
			'Response.Write("The cmParent is " & cmParent & "<br />")
			Response.Write("Processing FilePack: " & cmChild & "<br />")
			'ShowSubfolders FSO.GetFolder(Subfolder.Path)
			ShowFilePackType FSO.GetFolder(cmParent),FSO.GetFolder(cmChild)
			
			Next
End Sub



Sub ShowFilePackType(Parent,Child)
		Response.Write("The Parent Folder is " & Parent.Path & "<br />")
		Response.Write("The Child Child is " & Child.Path & "<br />")
			If Child.Name = "app" then
			Response.Write("usebat should be A <br />")
			Response.Write("destination should be D: <br />")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			
			FilePackDrive = "D:\"
						
			End If
			If Child.Name = "winapp" then
			Response.Write("usebat should be A <br />")
			Response.Write("destination should be D: <br />")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			
			FilePackDrive = "D:\"
			
			End If
			If Child.Name = "pline" then
			Response.Write("usebat should be A")
			Response.Write("destination should be D:")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			FilePackDrive = "D:"
			
			End If
			If Child.Name = "dat" then
			Response.Write("usebat should be B <br />")
			Response.Write("destination should be E: <br />")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			FilePackDrive = "E:"
			
			End If
			If Child.Name = "pub" then
			Response.Write("usebat should be C <br />")
			Response.Write("destination should be E: <br />")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			FilePackDrive = "E:"
			
			End If
			If Child.Name = "locapp" then
			Response.Write("usebat should be A <br />")
			Response.Write("destination should be E: <br />")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			FilePackDrive = "E:"
			
			End If
			If Child.Name = "mmap" then
			Response.Write("usebat should be A <br />")
			Response.Write("destination should be F: <br />")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			
			FilePackDrive = "F:\"
			Response.Write("FilePackName:  fp_" & Parent.Name & "_" & FilePackType & "<br />")
			End If
			If Child.Name = "mapinfo" then
			Response.Write("usebat should be A <br />")
			Response.Write("destination should be D <br />")
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			FilePackType = Child.Name
			FilePackDrive = "D:\"
			
			End If
			If Child.Name = "home" then
			Response.Write("usebat should be A <br />")
			Response.Write("destination should be D <br />")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			FilePackDrive = "D:\"
			
			End If
			If Child.Name = "tsdat" then
			Response.Write("usebat should be  <br />A")
			Response.Write("destination should be D <br />")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			FilePackDrive = "D:\"
			
			End If
			If Child.Name = "tsapp" then
			Response.Write("usebat should be A  <br />")
			Response.Write("destination should be D  <br />")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			FilePackDrive = "D:\"
			
			End If
			If Child.Name = "img" then
			Response.Write("usebat should be A <br />")
			Response.Write("destination should be D <br />")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			FilePackDrive = "D:\"
			
			End If
			If Child.Name = "UNIXWebsphereAppServer" then
			Response.Write("usebat DOES NOT APPLY! <br />")
			Response.Write("destination should be /  <br />")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			FilePackDrive = "/"
			
			End If
			If Child.Name = "UNIXAppServer" then
			Response.Write("usebat DOES NOT APPLY! <br />")
			Response.Write("destination should be / <br />")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			FilePackDrive = "/"
			
			End If
			If Child.Name = "NTAppServer" then
			Response.Write("usebat should be A <br />")
			Response.Write("destination should be D <br />")
			FilePackType = Child.Name
			FilePackName = "fp_" & Parent.Name & "_" & FilePackType
			FilePackDrive = "D:\"
			
			End If
				If Child.Name = "inetpub" then
			Response.Write("This is an IIS Website <br />")
			Response.Write("IIS Not Supported at this TIme<br />")
			End If
			Response.Write("Package Name:  fp_" & Parent.Name & "_" & FilePackType & "<br />")
			CopySource "\\CHANGE\INBOX\" & ChangeID & "\" & FilePackType,"D:\Package_Source\File_Packs\" & ChangeID & "\"  & FilePackName
			FilePackLocation = "D:\Package_Source\File_Packs\" & ChangeID & "\"  & FilePackName
			'newPackageName = FilePackName
			'newPackageDescription = FilePackType
			'newPackageSourceFlag = 2
			'newPackageSourcePath = FilePackLocation
				'Let's Create Some Packages, Programs & Advertisements, shall we?
				'Package Creation Variables
				strNewPackageName = FilePackName
				strNewPackageDescription = FilePackName & " was created using SCCM SDK on " & Now
				intNewPackageSourceFlag = 2
				strNewPackageSourcePath = FilePackLocation
				'Program Creation Variables
				strNewProgramName = "FilePack Installer"
				strNewProgramComment = "According to the original Tivoli Script, DESTINATION should be " & FilePackDrive
				strNewProgramCommandLine = "legacy.bat or FilePack_Installer.exe /destination=" & FilePackDrive & " /type=" & FilePackType
				intNewMaxRunTime = 180
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
				strExistingCollectionID = "CBT0005C"
				strExistingCollectionName = "BETA: Test Area"
				strExistingProgramName = strNewProgramName
				strNewAdvertisementName = "BETA: SDK-" & strNewPackageName & " to " & strExistingCollectionName
				strNewAdvertisementComment = "Advertisement of " &strNewPackagename & " to " & strExistingCollectionName & " created by SCCM SDK"
				'strNewAdvertisementFlags = ""
				strNewAvailableAfter = Now()
				strNewAdvertisementStartOfferDateTime = ConvertToWMIDate(strNewAvailableAfter)
				'Wscript.Echo strNewAdvertisementStartOfferDateTime
				Response.Write(strNewAdvertisementStartOfferDateTime & " <br />")
				strNewAdvertisementStartOfferEnabled = TRUE
				'SWD Assign Assignment to Collection Variables
				'strExistingAdvertisementID = 
				'strIncludeSubCollection = 

			CreatePackageInFolder(connection, strNewPackageName, strNewPackageDescription, intNewPackageSourceFlag, strNewPackageSourcePath)
			CreateProgram(connection, strExistingPackageID, strNewProgramName, strNewProgramComment, strNewProgramCommandLine, intNewMaxRunTime)
			'Call CreateStaticCollection(connection, existingParentCollectionID, newCollectionName, newCollectionComment, ownedByThisSite, resourceClassName, resourceID)
			SWDAssignPackageToDistributionPoint(connection, strExistingPackageID, strSiteCode, strDPServerName)
			SWDCreateAdvertisement(connection, strExistingCollectionID, strExistingPackageID, strExistingProgramName, strNewAdvertisementName, strNewAdvertisementComment, strNewAdvertisementFlags, strNewAdvertisementStartOfferDateTime, strNewAdvertisementStartOfferEnabled)
			'Call SWDAssignAdvertisementCollection(connection, existingAdvertisementID, existingCollectionID, includeSubCollection)
			
End Sub

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
	'wscript.echo "Package: "  & newPackageName & " ID: " & newPackageID
	Response.Write("Package: "  & newPackageName & " ID: " & newPackageID & " <br />")
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
	
	'wscript.echo "Package Moved to Folder:" & FolderIdentifier
	Response.Write("Package Moved to Folder:" & FolderIdentifier & "<br />")
	
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
    'wscript.echo "Created program: " & newProgramName
	Response.Write("Created program: " & newProgramName & " <br />")

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
    'Wscript.Echo "Assigned package: " & distributionPoint.PackageID & " to DP: " & serverName
	Response.Write("Assigned package: " & distributionPoint.PackageID & " to DP: " & serverName & " <br />")
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
    'Wscript.Echo "Created advertisement: " & newAdvertisement.AdvertisementName & " with exsiting package " & newAdvertisement.PackageID & " with existing program name  " & newAdvertisement.ProgramName
	Response.Write("Created advertisement: " & newAdvertisement.AdvertisementName & " with exsiting package " & newAdvertisement.PackageID & " with existing program name  " & newAdvertisement.ProgramName & " <br />")
	
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
		Response.Write("Couldn't connect to SCCM SDK! " & Err.Description & " <br />")
        Connect = null
        Exit Function
    End If
    Response.Write("Web Page is Connected to SCCM SDK Service :)<br />")

    ' Determine where the provider is and connect.
    Set providerLoc = swbemServices.InstancesOf("SMS_ProviderLocation")

        For Each location In providerLoc
            If location.ProviderForLocalSite = True Then
                Set swbemServices = swbemLocator.ConnectServer _
                 (location.Machine, "root\sms\site_" + _
                    location.SiteCode,userName,userPassword)
                If Err.Number<>0 Then
                    'Wscript.Echo "Couldn't connect:" + Err.Description
					Response.Write("Connection Error: " & Err.Description & " <br />")
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
end function


dim ChangeID
ChangeID=Request.Form("ChangeID")
If ChangeID<>"" Then
	Response.Write("Determining File Packs to Copy for: " & ChangeID & "<br />")
	PackageSourceFolder= "D:\Package_Source\File_Packs\" & ChangeID
	FSO.CreateFolder(PackageSourceFolder)
			
	ShowSubfolders FSO.GetFolder("\\CHANGE\INBOX\" & ChangeID)
	
	Response.Write("FilePackName equals: " & FilePackName & "<br />")
	Response.Write("FilePackType equals: " & FilePackType & "<br />")
	Response.Write("FilePackDrive equals: " & FilePackDrive & "<br />")
	
	  	Response.Write("No More File Packs Under: " & ChangeID & "<br />")
	  
	 Else
	 Response.Write("No Change ID Has Been Selected!<br />")
	  
	 
End If
%>
</body>
</html>

   
    