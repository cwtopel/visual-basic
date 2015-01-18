' This VBS will search recursively from the StartKey downwards.
' This will not search upwards in the registry tree
' Use care when searching a large tree as it may consume CPU to do a large search.
' Output will be interactive, use cscript.exe if running locally in test
' Prior to load to iTWatch, change the value of SearchValue and the StartKey to ensure resources are not over utilized.
' Author: kstokes, cwtopel
' Version: 1.18.2015

WScript.Echo ""
WScript.Echo "START"
Const HKCU = &H80000001  
Const HKLM = &H80000002  
Const HKU =  &H80000003

Const StartKey    = ""
Const SearchKey = "System"

Set reg = GetObject("winmgmts://./root/default:StdRegProv")
Set reg1 = GetObject("winmgmts://./root/default:StdRegProv")
' Gets a list of the top level of HKEY_USERs, and then looks in each for the directory
' Stored in array rootSIDs
reg.EnumKey HKU, StartKey, rootSIDs


If Not IsNull(rootSIDs) Then
	For Each SID in rootSIDs
		combined = SID & "\Software\"
		WScript.Echo "Key: " & combined
		'Get all subkeys of this key
		reg1.enumKey HKU, combined, SIDSubKeys
		If Not IsNull (SIDSubKeys) Then
			For Each SubKey in SIDSubKeys
				WScript.Echo "-Subkey is: " & SubKey
					If SubKey = "CryptoLocker" Then
					WScript.Echo "------------------"
					WScript.Echo "ALERT, CRYPTOLOCKER FOUND"
					WScript.Echo "------------------"
					End If
			Next
		End If
		If IsNull (SIDSubKeys) Then
			WScript.Echo "SIDSubKeys is null"
		End If		
	Next
End If

WScript.Echo "END"
