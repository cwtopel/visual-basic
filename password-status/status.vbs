On Error Resume Next

Dim oFSO, oGroup, oUser, oDomain, oMaxPwdAge, oFile
Dim iUserAccountControl, dtmValue, iTimeInterval, dblMaxPwdNano
Set oFSO = CreateObject("scripting.filesystemobject")
Set oGroup = GetObject("LDAP://ou=Departments,dc=domain,dc=somesite,dc=org")
Const ForWriting = 2
Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
Const E_ADS_PROPERTY_NOT_FOUND = &h8000500D
Const ONE_HUNDRED_NANOSECOND = .000000100
Const SECONDS_IN_DAY = 86400

Sub enumMembers(oGroup)
   For Each oUser In oGroup
      If oUser.Class = "user" Then
            iUserAccountControl = oUser.Get("userAccountControl")
            If iUserAccountControl And ADS_UF_DONT_EXPIRE_PASSWD Then
                oFile.WriteLine oUser.department & "/" & oUser.displayName & vbTab & "password does not expire"
            Else
                dtmValue = oUser.PasswordLastChanged
                If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
                    oFile.WriteLine oUser.department & "/" & oUser.displayName & vbTab & "password has never been set"
                Else
                    iTimeInterval = Int(Now - dtmValue)
                    oFile.WriteLine oUser.department & "/" & oUser.displayName & vbTab & iTimeInterval & " days old"
                End If
                Set oDomain = GetObject("LDAP://dc=domain,dc=somesite,dc=org")
                Set oMaxPwdAge = oDomain.Get("maxPwdAge")
                If oMaxPwdAge.LowPart = 0 Then
                    oFile.WriteLine oUser.department & "/" & oUser.displayName & vbTab & "password does not expire"
                Else
                    dblMaxPwdNano = _
                            Abs(oMaxPwdAge.HighPart * 2^32 + oMaxPwdAge.LowPart)
                    dblMaxPwdSecs = dblMaxPwdNano * ONE_HUNDRED_NANOSECOND
                    dblMaxPwdDays = Int(dblMaxPwdSecs / SECONDS_IN_DAY)
                    If iTimeInterval >= dblMaxPwdDays Then
                        oFile.WriteLine oUser.department & "/" & oUser.displayName & vbTab & "password has expired."
                    End If
                End If
            End If
        ElseIf oUser.Class = "organizationalUnit" or oUser.Class = "container" Then
            enumMembers(oUser)
        End If
    Next
End Sub

Set oFile = oFSO.CreateTextFile("PasswordStatus.txt", ForWriting, True)
Call enummembers(ogroup)
