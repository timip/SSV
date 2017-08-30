'==========================================================================
'	Name:		SAT
'	Section:	win.user_ldap
'	Author:		Tim IP
'	Build:		20131009A
'==========================================================================

strComputer="."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colComputer = objWMIService.ExecQuery ("Select DomainRole from Win32_ComputerSystem")
For Each oComputer in colComputer
	iDR = oComputer.DomainRole
Next


On Error Resume Next
strCurTime = Now
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")
Set objDomain = GetObject("LDAP://" & strDNSDomain)
Set objRootDSE = Nothing
Set objDomain = Nothing
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"

Set aioCommand = CreateObject("ADSystemInfo")
domainNETBios = aioCommand.DomainShortName

adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection
Err.Clear
On Error Goto 0

If (strDNSDomain <> "") Then
	On Error Resume Next
	strBase = "<LDAP://" & strDNSDomain & ">"
	strFilter = "(&(objectCategory=person)(objectClass=user))"
	strAttributes = "sAMAccountName, name, userAccountControl, whenChanged, whenCreated, accountExpires, description, lastLogonTimestamp, lastLogon, logonCount, pwdLastSet, badPwdCount, badPasswordTime, LockoutTime, lockoutDuration, distinguishedName, objectSid, primaryGroupID"
	strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	adoCommand.CommandText = strQuery
	adoCommand.Properties("Page Size") = 1000
	adoCommand.Properties("Timeout") = 60
	adoCommand.Properties("Cache Results") = False
	Set adoRecordset = adoCommand.Execute
	Err.Clear
	On Error Goto 0
	
	Do Until adoRecordset.EOF	
		On Error Resume Next
		strSAMAccountName = adoRecordset.Fields("sAMAccountName").Value
		strName = adoRecordset.Fields("name").Value
		strUserAccountControl = adoRecordset.Fields("userAccountControl").Value
		strWhenChanged = adoRecordset.Fields("whenChanged").Value
		strWhenCreated = adoRecordset.Fields("whenCreated").Value
		strAccountExpires = LargeIntegerToDate(adoRecordset.Fields("accountExpires").Value)
		arrDescription = adoRecordset.Fields("description").Value
		strDescription = ""
		For each item in arrDescription
			strDescription = strDescription & item
		Next
		strLastLogonTimestamp = LargeIntegerToDate(adoRecordset.Fields("lastLogonTimestamp").Value)
		strLogonCount = adoRecordset.Fields("logonCount").Value
		strPwdLastSet = LargeIntegerToDate(adoRecordset.Fields("pwdLastSet").Value)
		strBadPwdCount = adoRecordset.Fields("badPwdCount").Value
		strBadPasswordTime = LargeIntegerToDate(adoRecordset.Fields("badPasswordTime").Value)
		strLockoutTime = LargeIntegerToDate(adoRecordset.Fields("lockoutTime").Value)
		strDistinguishedName = adoRecordset.Fields("distinguishedName").Value
		strLastlogon = LargeIntegerToDate(adoRecordset.Fields("lastLogon").Value)
		strPrimaryGroupID = adoRecordset.Fields("primaryGroupID").Value
		Err.Clear
		On Error Goto 0	
		strObjectSid = adoRecordset.Fields("objectSid").Value
		
		domainrole = 0
		Set objWMIService = GetObject("winmgmts:\\.\root\cimv2") 
		Set colItems = objWMIService.ExecQuery("Select DomainRole from Win32_ComputerSystem",,48) 
		For Each objItem in colItems 
			domainrole = objItem.DomainRole 
		Next 

		If (domainrole > 3) Then
			strAccountLocked = "F"	
			strAccountCantChgPwd = "F"	
			Set user = GetObject("WinNT://" & CreateObject("ADSystemInfo").DomainShortName & "/" & adoRecordset.Fields("samAccountName"))
			If (user.IsAccountLocked) then
				strAccountLocked = "T"
			End if
			If (user.Get("userFlags") And ADS_UF_PASSWD_CANT_CHANGE) <> 0 Then
				strAccountCantChgPwd = "T"
			End if
		Else
			strAccountLocked = "NA"
			strAccountCantChgPwd = "NA"
		End If
		
		Wscript.Echo "SAMAccountName=""" & strSAMAccountName & _
			""" Desc=""" & strName & _
			""" UserAccountControl=""" & strUserAccountControl & _
			""" WhenChangedUTC=""" & strWhenChanged & _
			""" WhenCreatedUTC=""" & strWhenCreated & _
			""" AccountExpires=""" & strAccountExpires & _
			""" Description=""" & strDescription & _
			""" LastLogonTimestamp=""" & strLastLogonTimestamp & _
			""" LastLogon=""" & strLastlogon & _
			""" LogonCount=""" & strLogonCount & _
			""" PwdLastSet=""" & strPwdLastSet & _
			""" BadPwdCount=""" & strBadPwdCount & _
			""" BadPasswordTime=""" & strBadPasswordTime & _
			""" LockoutTime=""" & strLockoutTime & _
			""" CurrentTime=""" & strCurTime & _
			""" Sid=""" & HexStrToDecStr(OctetToHexStr(strObjectSid)) & _
			""" PrimaryGroupID=""" & strPrimaryGroupID & _
			""" DistinguishedName=""" & strDistinguishedName & _
			""" AccountLocked=""" & strAccountLocked & _
			""" AccountCantChgPwd=""" & strAccountCantChgPwd & _
			""""
		adoRecordset.MoveNext
	Loop
Else
	Wscript.echo "Info=""This machine is not joined to a domain or not able to connect to domain controller."""	
End If

On Error Resume Next
adoRecordset.Close
adoConnection.Close
Err.Clear
On Error Goto 0


Function adoGetAtt(adoRecordset, target)
	On Error Resume Next
	set adoGetAtt = adoRecordset.Fields(target).Value
	If (Err.Number <> 0) Then
		Wscript.echo Err.descrption
	End If
End Function

Function LargeIntegerToDate(value)
	'On Error Resume Next
	'takes Microsoft LargeInteger value (Integer8) and returns according the date and time
	Dim sho, timeShiftValue, timeShift, i8High, i8Low
    'first determine the local time from the timezone bias in the registry
    Set sho = CreateObject("Wscript.Shell")
    timeShiftValue = sho.RegRead("HKLM\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
    If IsArray(timeShiftValue) Then
        timeShift = 0
        For i = 0 To UBound(timeShiftValue)
            timeShift = timeShift + (timeShiftValue(i) * 256^i)
        Next
    Else
        timeShift = timeShiftValue
    End If

    'get the large integer into two long values (high part and low part)
    i8High = value.HighPart
    i8Low = value.LowPart

    If (i8Low < 0) Then
       i8High = i8High + 1 
    End If
	
    'calculate the date and time: 100-nanosecond-steps since 12:00 AM, 1/1/1601
    If ((i8High = 0) And (i8Low = 0) Or ((i8High = 2147483648) And (i8Low = -1))) Then 
        LargeIntegerToDate = "Not Defined"
    Else 
        LargeIntegerToDate = #1/1/1601# + (((i8High * 2^32) + i8Low)/600000000 - timeShift)/1440 
    End If
	
End Function

Function OctetToHexStr(arrbytOctet)
	' Function to convert OctetString (Byte Array) to a hex string.
	Dim k
	OctetToHexStr = ""
	For k = 1 To Lenb(arrbytOctet)
	OctetToHexStr = OctetToHexStr _
	& Right("0" & Hex(Ascb(Midb(arrbytOctet, k, 1))), 2)
	Next
End Function

Function HexStrToDecStr(strSid)
	Dim arrbytSid, lngTemp, j
	ReDim arrbytSid(Len(strSid)/2 - 1)
	For j = 0 To UBound(arrbytSid)
	arrbytSid(j) = CInt("&H" & Mid(strSid, 2*j + 1, 2))
	Next
	HexStrToDecStr = "S-" & arrbytSid(0) & "-" & arrbytSid(1) & "-" & arrbytSid(8)
	lngTemp = arrbytSid(15)
	lngTemp = lngTemp * 256 + arrbytSid(14)
	lngTemp = lngTemp * 256 + arrbytSid(13)
	lngTemp = lngTemp * 256 + arrbytSid(12)
	HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)
	lngTemp = arrbytSid(19)
	lngTemp = lngTemp * 256 + arrbytSid(18)
	lngTemp = lngTemp * 256 + arrbytSid(17)
	lngTemp = lngTemp * 256 + arrbytSid(16)
	HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)
	lngTemp = arrbytSid(23)
	lngTemp = lngTemp * 256 + arrbytSid(22)
	lngTemp = lngTemp * 256 + arrbytSid(21)
	lngTemp = lngTemp * 256 + arrbytSid(20)
	HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)
	lngTemp = arrbytSid(25)
	lngTemp = lngTemp * 256 + arrbytSid(24)
	HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)
End Function
