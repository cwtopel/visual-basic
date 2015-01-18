'CREATE A Reports Sub Folder since the script doesn't do it for you.
On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2
Const ForAppending = 8
Const TriStateFalse = 0

TIME_DATE_FORMAT = "." & Year(Now()) & Month(Now()) & Day(Now()) & "." & Hour(Now()) & Minute(Now()) & Second(Now())
FILE_NAME = ".\Reports\AD_Users" & TIME_DATE_FORMAT & ".csv"

Dim fso, cn, cmd, ou, rs, objStream, OutputBuilder, outFile

Set fso = CreateObject("Scripting.FileSystemObject")

Set cn = CreateObject("ADODB.Connection")
cn.Provider = "ADsDSOObject"
cn.Open "Active Directory Provider"
Set cmd = CreateObject("ADODB.Command")
Set cmd.ActiveConnection = cn
ou = "OU=Computers,DC=DOMAINNAME,DC=com"

cmd.CommandText = "SELECT name FROM 'LDAP://" & ou & "' WHERE objectClass='Computer' ORDER BY name"
cmd.Properties("Page Size") = 1000
cmd.Properties("Searchscope") = ADS_SCOPE_SUBTREE

Set rs = cmd.Execute
rs.MoveFirst
LineCount = rs.RecordCount
rs.MoveFirst

Set objStream = FSO.CreateTextFile(FILE_NAME, True, TristateFalse)

msgbox ("Beginning Domain Scan (" & LineCount & ")")

Do Until rs.EOF
	strComputer = rs(0)
	if strComputer = "" then
		wscript.quit
	else
		OutputBuilder = "" 'RESET OUTPUT BUILDER
		OutputBuilder = strComputer & ","
		Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy}").ExecQuery ("select * from Win32_PingStatus where address = '" & strComputer & "'")
		For Each objStatus in objPing
			If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then
				'REQUEST TIME OUT
				OutputBuilder = OutputBuilder & "DID NOT REPLY" & "," & "NA"
			else
				'PING RESPONSE
				set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
				Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
				For Each objComputer in colSettings
					If strComputer <> objComputer.name Then
						OutputBuilder = OutputBuilder & "NAME MISSMATCH"
					else
						OutputBuilder = OutputBuilder & objComputer.name & "," & objComputer.username

						'Computer Info.
						Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
						Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
						For Each objOS in colOperatingSystems
							dtmBootup = objOS.LastBootUpTime
							dtmLastBootUpTime = WMIDateStringToDate(dtmBootup)
							dtmSystemUptime = DateDiff("h", dtmLastBootUpTime, Now)
							OutputBuilder = OutputBuilder & "," & dtmSystemUptime

							'OutputBuilder = OutputBuilder & "," & objOS.BootDevice
							'OutputBuilder = OutputBuilder & "," & objOS.BuildNumber
							'OutputBuilder = OutputBuilder & "," & objOS.BuildType
							'OutputBuilder = OutputBuilder & "," & objOS.Caption
							'OutputBuilder = OutputBuilder & "," & objOS.CodeSet
							'OutputBuilder = OutputBuilder & "," & objOS.CountryCode
							'OutputBuilder = OutputBuilder & "," & objOS.Debug
							'OutputBuilder = OutputBuilder & "," & objOS.EncryptionLevel
							'dtmConvertedDate.Value = objOS.InstallDate
							'dtmInstallDate = dtmConvertedDate.GetVarDate
							'OutputBuilder = OutputBuilder & "," & & dtmInstallDate
							'OutputBuilder = OutputBuilder & "," & objOS.NumberOfLicensedUsers
							'OutputBuilder = OutputBuilder & "," & objOS.Organization
							'OutputBuilder = OutputBuilder & "," & objOS.OSLanguage
							'OutputBuilder = OutputBuilder & "," & objOS.OSProductSuite
							'OutputBuilder = OutputBuilder & "," & objOS.OSType
							'OutputBuilder = OutputBuilder & "," & objOS.Primary
							'OutputBuilder = OutputBuilder & "," & objOS.RegisteredUser
							'OutputBuilder = OutputBuilder & "," & objOS.SerialNumber
							OutputBuilder = OutputBuilder & "," & objOS.Description 'Computer Description
							'OutputBuilder = OutputBuilder & "," & objOS.Version 'OS VERSION
						Next

						'IPADDRESS
						Set IPConfigSet = objWMIService.ExecQuery ("Select IPAddress from Win32_NetworkAdapterConfiguration ")
						For Each IPConfig in IPConfigSet
							If Not IsNull(IPConfig.IPAddress) Then
								For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
									If Not IPConfig.IPAddress(i) = "" Then
										OutputBuilder = OutputBuilder & "," & IPConfig.IPAddress(i)
									Else
										OutputBuilder = OutputBuilder & ", NA"
									End If
								Next
							End If
						Next
					End If
				Next
			end if
		next
		objStream.WriteLine OutputBuilder 'Builds the Computer Information Output
	end if
	rs.MoveNext
Loop

outFile.Close
objStream.Close
msgbox ("Domain Has Been Processed")

Set outFile = Nothing
Set objStream = Nothing
Set fso = Nothing

Function WMIDateStringToDate(dtmBootup)
	WMIDateStringToDate =  CDate(Mid(dtmBootup, 5, 2) & "/" & _
		Mid(dtmBootup, 7, 2) & "/" & Left(dtmBootup, 4) _
		& " " & Mid (dtmBootup, 9, 2) & ":" & Mid(dtmBootup, 11, 2) & ":" & Mid(dtmBootup, 13, 2))
End Function
