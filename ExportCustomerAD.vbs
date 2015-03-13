Dim exportFileName
Dim count
Dim FileObject
Dim LogFile, LogFile2, LogFile2a, LogFile3, LogFile4, LogFile5
Dim Connection
Dim Command
Dim RecordSet

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!---READ ME---!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'You will find commented sections marked at the beginning and end with
' "'+++++++" These sections are places in which you will want to update
' the file in order to match up with the customer you are migrating.
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'These items are to be modified to fit the needs of creating a new 
'user on the hosted domain. These are unchanging values throughout
'the duration of the script
' 
'strCustomerNumber = 6-digit customer number that corresponds to the Customer Number.
'strShortName = Friendly name of Customer.
'strTargetOU = Customer OU to target for generate list of users. This helps to
'  exclude certain parts of Active Directory such as the "Users" container.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
strCustomerNumber		= "XXXXXX"
strShortName			= "ShortName"
strTargetOU				= "OU=XXXXX," 'Do not forget to use a "," ad the end.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
strShortNameCount 		= LEN(strShortName)

Const Writable = 2

exportFileName = InputBox("Enter the beginning name for the export files;" & _
 " .csv will be added to the file name.", "Customer Export Script",_
 "XXXXXX-Export")
 
Set FileObject = CreateObject("Scripting.FileSystemObject")
Set LogFile = FileObject.OpenTextFile(exportFileName & "-USERS.csv", Writable, True)
Set LogFile2 = FileObject.OpenTextFile(exportFileName & "-USERS-SMTP.csv", Writable, True)
Set LogFile2a = FileObject.OpenTextFile(exportFileName & "-USERS-X500.csv", Writable, True)
Set LogFile3 = FileObject.OpenTextFile(exportFileName & "-CONTACTS.csv", Writable, True)
Set LogFile4 = FileObject.OpenTextFile(exportFileName & "-GROUPS.csv", Writable, True)
Set LogFile5 = FileObject.OpenTextFile(exportFileName & "-GROUPS-SMTP.csv", Writable, True)

LogFile.Write "displayName,givenName,sn,facsimileTelephoneNumber,homePhone,mobile,pager,telephoneNumber" & vbCrLf
LogFile2.Write "displayName,proxyAddresses" & vbCrLf
LogFile2a.Write "displayName,X500" & vbCrLf
LogFile3.Write "displayName,mail,givenName,sn,mailNickname,Company" & vbCrLf
LogFile4.Write "displayName" & vbCrLf
LogFile5.Write "displayName,proxyAddresses" & vbCrLf

UsersExport()
ContactsExport()
GroupsExport()

Sub UsersExport()
	Set Connection   = CreateObject("ADODB.Connection")
	Set Command      = CreateObject("ADODB.Command")
	Set RecordSet    = CreateObject("ADODB.RecordSet") 

	With Connection
		.Provider    = "ADsDSOObject"
		.Open "Active Directory Provider"
	End With

	'************************************************************************ 
	'This is the query that looks up specific information from the customers
	'that we want to harvest from their Active Directory.
	'************************************************************************ 

	Set Command.ActiveConnection     = Connection
	Set objRootDSE                   = GetObject("LDAP://rootDSE")
	Command.CommandText              = "<LDAP://" & strTargetOU & _
										objRootDSE.Get("defaultNamingContext") _
										& ">;(&(&(objectCategory=user)" _
										& "(objectClass=user))" _
										& "(proxyAddresses=*))" _
										& ";distinguishedName,cn,givenName,sn," _
										& "facsimileTelephoneNumber,homePhone," _
										& "mobile,pager,telephoneNumber," _
										& "proxyAddresses,sAMAccountName,legacyExchangeDN;subtree" 
	Set RecordSet                    = Command.Execute 
	While Not RecordSet.EOF

	On Error Resume Next

	'************************************************************************ 
	'Utilizes Customers Active Directory Data and determines what to record
	'into the CSV file.
	'************************************************************************ 
	strLegacyExchangeDN = RecordSet.Fields("legacyExchangeDN")
	strFirstName = RecordSet.Fields("givenName")
	strLastName = RecordSet.Fields("sn")
		If IsNull(strLastName) Then
			If LCASE(Left(strFirstName,strShortNameCount)) = LCASE(strShortName) Then
				strUserCN = strFirstName
			Else
				strUserCN = strShortName & " " & strFirstName
				strFirstName = strUserCN
			End If
			strLastName = "blank"
		Else
			strUserCN = strFirstName & " " & strLastName
		End If
	strFaxNumber = RecordSet.Fields("facsimileTelephoneNumber")
		If IsNull(strFaxNumber) = True Then
			strFaxNumber = "blank"
		Else
			strFaxNumber = ValidPhone(strFaxNumber)
		End If
	strHomeNumber = RecordSet.Fields("homePhone")
		If IsNull(strHomeNumber) = True Then
			strHomeNumber = "blank"
		Else
			strHomeNumber = ValidPhone(strHomeNumber)
		End If
	strMobileNumber = RecordSet.Fields("mobile")
		If IsNull(strMobileNumber) = True Then
			strMobileNumber = "blank"
		Else
			strMobileNumber	= ValidPhone(strMobileNumber)
		End If
	strPageNumber = RecordSet.Fields("pager")
		If IsNull(strPageNumber) = True Then
			strPageNumber = "blank"
		Else
			strPageNumber = ValidPhone(strPageNumber)
		End If
	strOfficeNumber	= RecordSet.Fields("telephoneNumber")
		If IsNull(strOfficeNumber) = True Then
			strOfficeNumber = "blank"
		Else
			strOfficeNumber	= ValidPhone(strOfficeNumber)
		End If
	strProxyAddresses = RecordSet.Fields("proxyAddresses")
	strSAMAccountName = RecordSet.Fields("sAMAccountName")
	
	'************************************************************************ 
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'Modify these lines to exclude out all SMTP addresses that will be
	'included as part of the Hosted Side Recipiant policy for this customer.
	'Main SMTP address for Hosted Side will be <firstname>.<lastname>@<domain> so
	'if this was their default email address be sure to exclude it.
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'These exclude entire domains
	strExcludeSMTP1            = LCase("domain.com")
	strExcludeSMTP2            = LCase("domain.local")
	strExcludeSMTP3            = LCase("domain.local")
	'These exclude specific email addresses
	strExcludeFullSMTP1        = LCase(strSAMAccountName & "@domain.com")
	strExcludeFullSMTP2        = LCase(strSAMAccountName & "@domain.com")
	strExcludeFullSMTP3        = LCase(strSAMAccountName & "domain.com")
	strExcludeFullSMTP4        = LCase(Left(strFirstName,1) & strLastName & "@domain.com")
	strExcludeFullSMTP5        = LCase(Left(strFirstName,1) & strLastName & "@domain.com")
	strExcludeFullSMTP6        = LCase(strFirstName & "." & strLastName & "@domain.com")
	strExcludeFullSMTP7        = LCase(strFirstName & "." & strlastName & "@domain.com")
	strExcludeFullSMTP8        = LCase(strFirstName & strLastname & "@domain.com")
	
	'Examples using John Doe as name
	'%1g%s@domain.com = JDoe@domain.com = LCase(Left(strFirstName,1) & strLastName & "@domain.com")
	'%g%1s@domain.com = JohnD@domain.com = LCase(strFirstName & Left(strLastName,1) & "@domain.com")
	'%m@domain.com = strSAMAccountName@domain.com = LCase(strSAMAccountName & "@domain.com")
	'%g.%s@domain.com = John.Doe@domain.com = LCase(strFirstName & "." & strLastName & "@domain.com")
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	
	'************************************************************************ 
	'Determines the length of the values inputted in strExcludeSMTP1,2,3
	'************************************************************************ 
	strSMTP1Count              = Len(strExcludeSMTP1)
	strSMTP2Count              = Len(strExcludeSMTP2)
	strSMTP3Count              = Len(strExcludeSMTP3)
	strSMTP4Count              = Len(strExcludeFullSMTP1)
	strSMTP5Count              = Len(strExcludeFullSMTP2)
	strSMTP6Count              = Len(strExcludeFullSMTP3)
	strSMTP7Count              = Len(strExcludeFullSMTP4)
	strSMTP8Count              = Len(strExcludeFullSMTP5)
	strSMTP9Count              = Len(strExcludeFullSMTP6)
	strSMTP10Count             = Len(strExcludeFullSMTP7)
	strSMTP11Count             = Len(strExcludeFullSMTP8)
	'************************************************************************
	If IsNull(strFirstName) Then
	   count = count
	   RecordSet.MoveNext
	 Else
		LogFile.Write strUserCN & "," & strFirstName & "," & strLastName & "," & strFaxNumber & "," & _
			strHomeNumber & "," & strMobileNumber & "," & strPageNumber & "," & strOfficeNumber & vbCrLf
		For Each Item in strProxyAddresses
		   If ((Left(Item,5) = "SMTP:") OR (Left(Item,5) = "smtp:")) AND _ 
				 ((Right(LCase(Item),strSMTP1Count) = strExcludeSMTP1) OR _ 
				 (Right(LCase(Item),strSMTP2Count) = strExcludeSMTP2) OR _
				 (Right(LCase(Item),strSMTP3Count) = strExcludeSMTP3) OR _
				 (Right(LCase(Item),strSMTP4Count) = strExcludeFullSMTP1) OR _
				 (Right(LCase(Item),strSMTP5Count) = strExcludeFullSMTP2) OR _
				 (Right(LCase(Item),strSMTP6Count) = strExcludeFullSMTP3) OR _
				 (Right(LCase(Item),strSMTP7Count) = strExcludeFullSMTP4) OR _ 
				 (Right(LCase(Item),strSMTP8Count) = strExcludeFullSMTP5) OR _ 
				 (Right(LCase(Item),strSMTP9Count) = strExcludeFullSMTP6) OR _ 
				 (Right(LCase(Item),strSMTP10Count) = strExcludeFullSMTP7) OR _
				 (Right(LCase(Item),strSMTP11Count) = strExcludeFullSMTP8)) Then
'				 'Do Nothing
			Else
				If ((Left(Item,5) = "SMTP:") OR (Left(Item,5) = "smtp:")) Then			 
					Logfile2.Write strUserCN & "," & LCase(Mid(Item,6)) & vbCrLF
				End If
		   End If
		Next
		Logfile2a.Write strUserCN & "," & strLegacyExchangeDN & vbCrLF
		RecordSet.MoveNext
	End if
	Wend
	LogFile.Close
	LogFile2.Close
End Sub

Sub ContactsExport()
	Set Connection   = CreateObject("ADODB.Connection")
	Set Command      = CreateObject("ADODB.Command")
	Set RecordSet    = CreateObject("ADODB.RecordSet") 
	With Connection
		.Provider    = "ADsDSOObject"
		.Open "Active Directory Provider"
	End With
	'*****************************************************************************
	'** This is the query that looks up specific information from the customers **
	'** Active Directory. If the Query is done here then you would define it    **
	'** in the next section.                                                    **
	'*****************************************************************************
	'displayName,mail,givenName,sn,mailNickname,Company
	Set Command.ActiveConnection     = Connection
	Set objRootDSE                   = GetObject("LDAP://rootDSE")
	Command.CommandText              = "<LDAP://" & _
										objRootDSE.Get("defaultNamingContext")_
										& ">;(&(&(objectCategory=Person)(objectClass=Contact))(proxyAddresses=*))" & _
										";givenName,sn,mail,company;subtree" 
	Set RecordSet                    = Command.Execute 
	While Not RecordSet.EOF
	On Error Resume Next
	strFirstName = RecordSet.Fields("givenName")
	 	strFirstName = Replace(strFirstName,",","")
	strLastName = RecordSet.Fields("sn")
		If IsNull(strLastName) = True Then
			strLastName = "blank"
		Else
			strLastName = Replace(strLastName,",","")
		End If
	strMail = RecordSet.Fields("mail")	
	If strLastName = "blank" Then
		strContactNewCN = strFirstName
	Else
		strContactNewCN = strFirstName & " " & strLastName
	End If
	strMailNick	= strFirstName & strLastName & "-" & strCustomerNumber
		strMailNick = Replace(strMailNick," ","")
		strMailNick = Replace(strMailNick,".","")
		strMailNick = Replace(strMailNick,"(","")
		strMailNick = Replace(strMailNick,")","")
	strCompany = RecordSet.Fields("company")
		If IsNull(strCompany) = True Then
			strCompany = "blank"
		Else
			'Do Nothing
		End If
	'*****************************************************************************
	'** Creation of the LDIFDE file begins                                      **
	'***************************************************************************** 
	If IsNull(strFirstName) Then
	   count = count
	   RecordSet.MoveNext
	 Else
	   LogFile3.Write strContactNewCN & "," & LCASE(strMail) & "," & strFirstName & "," & strLastName & "," & LCASE(strMailNick) & vbCrLf
	   RecordSet.MoveNext
	End If
	Wend
	LogFile3.Close
	Set objRootDSE = Nothing
	Set Connection = Nothing
	Set Command = Nothing
	Set RecordSet = Nothing
End Sub

Sub GroupsExport()
	Set Connection = CreateObject("ADODB.Connection")
	Set Command = CreateObject("ADODB.Command")
	Set RecordSet = CreateObject("ADODB.RecordSet") 

	With Connection
		.Provider = "ADsDSOObject"
		.Open "Active Directory Provider"
	End With
	'************************************************************************ 
	'This is the query that looks up specific information from the customers
	'Active Directory. If the Query is done here then you would define it
	'in the next section.
	'************************************************************************
	Set Command.ActiveConnection = Connection
	Set objRootDSE = GetObject("LDAP://rootDSE")
	Command.CommandText = "<LDAP://" & _
		objRootDSE.Get("defaultNamingContext") & _
		">;(&(&(objectCategory=group)(objectClass=group))(proxyAddresses=*))" & _
		";distinguishedName,cn,proxyAddresses;subtree"
	Set RecordSet = Command.Execute
	While Not RecordSet.EOF	
	
	strCN = RecordSet.Fields("cn")
	strProxyAddresses = RecordSet.Fields("proxyAddresses")
	strGroupCN = strShortName & " " & strCN	

	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'Modify these lines to exclude out domain specific SMTP addresses.
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'These exclude entire domains
	strExcludeSMTP1            = LCase("domian.com")
	strExcludeSMTP2            = LCase("domain.com")
	strExcludeSMTP3            = LCase("domain.local")
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	
	'************************************************************************ 
	'Determines the length of the values inputted in strExcludeSMTP1,2,3
	'************************************************************************ 
	strSMTP1Count              = Len(strExcludeSMTP1)
	strSMTP2Count              = Len(strExcludeSMTP2)
	strSMTP3Count              = Len(strExcludeSMTP3)
	'************************************************************************ 
	LogFile4.Write strGroupCN & vbCrLf
	For Each Item in strProxyAddresses
		If ((Left(Item,5) = "SMTP:") OR (Left(Item,5) = "smtp:")) AND _ 
			((Right(LCase(Item),strSMTP1Count) = strExcludeSMTP1) OR _ 
			(Right(LCase(Item),strSMTP2Count) = strExcludeSMTP2) OR _
			(Right(LCase(Item),strSMTP3Count) = strExcludeSMTP3)) Then
			'Do Nothing
		Else
			If ((Left(Item,5) = "SMTP:") OR (Left(Item,5) = "smtp:")) Then
				LogFile5.Write strGroupCN & "," & LCase(Mid(Item,6)) & vbCrLF
			End If
		End If		
	Next
	RecordSet.MoveNext
	Wend
	LogFile4.Close
	LogFile5.Close
End Sub



Function ValidPhone(strNumber)
	Dim tmp
	Dim objRegEx : Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Global = True
	objRegEx.Pattern = "[^0-9]"
	strNumber = objRegEx.Replace(strNumber,"")
	strNumber = CStr(strNumber)
	Select Case Len(strNumber)
		Case 1
			tmp = strNumber
		Case 2
			tmp = strNumber
		Case 3
			tmp = strNumber
		Case 4
			tmp = strNumber		
		Case 10
			tmp = tmp & "(" & Mid( strNumber, 1, 3 ) & ") "
			tmp = tmp & Mid( strNumber, 4, 3 ) & "-"
			tmp = tmp & Mid( strNumber, 7, 4 )
		Case 11
			tmp = tmp & "(" & Mid( strNumber, 2, 3 ) & ") "
			tmp = tmp & Mid( strNumber, 5, 3 ) & "-"
			tmp = tmp & Mid( strNumber, 8, 4 )
		Case 12
			tmp = tmp & "(" & Mid( strNumber, 1, 3 ) & ") "
			tmp = tmp & Mid( strNumber, 4, 3 ) & "-"
			tmp = tmp & Mid( strNumber, 7, 4 ) & " ext. "
			tmp = tmp & Mid( strNumber, 11, 4 )
		Case 13
			tmp = tmp & "(" & Mid( strNumber, 1, 3 ) & ") "
			tmp = tmp & Mid( strNumber, 4, 3 ) & "-"
			tmp = tmp & Mid( strNumber, 7, 4 ) & " ext. "
			tmp = tmp & Mid( strNumber, 11, 4 )
		Case 14
			tmp = tmp & "(" & Mid( strNumber, 1, 3 ) & ") "
			tmp = tmp & Mid( strNumber, 4, 3 ) & "-"
			tmp = tmp & Mid( strNumber, 7, 4 ) & " ext. "
			tmp = tmp & Mid( strNumber, 11, 4 )
		Case Else
			tmp = "blank"
	End Select
	ValidPhone = tmp
End Function

Set objRootDSE = Nothing
Set FileObject = Nothing
Set LogFile = Nothing
Set LogFile2 = Nothing
Set LogFile3 = Nothing
Set LogFile4 = Nothing
Set LogFile5 = Nothing
Set Connection = Nothing
Set Command = Nothing
Set RecordSet = Nothing