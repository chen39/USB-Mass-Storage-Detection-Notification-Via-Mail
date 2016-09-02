strComputer = "." '(Any computer name or address)
Set wmi = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set wmiEvent = wmi.ExecNotificationQuery("select * from __InstanceOperationEvent within 1 where TargetInstance ISA 'Win32_PnPEntity' and TargetInstance.Description='USB Mass Storage Device'") 
While True
Set usb = wmiEvent.NextEvent()
Select Case usb.Path_.Class
Case "__InstanceCreationEvent" 
On Error Resume Next
strComputer = "."
Dim fso
Dim objEmail
Set objEmail = CreateObject("CDO.Message")
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set f = fso.CreateTextFile("output.txt", True)
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colDevices = objWMIService.ExecQuery ("Select * From Win32_USBControllerDevice")
For Each objDevice in colDevices
 strDeviceName = objDevice.Dependent
 'msgbox strDeviceName
 strQuotes = Chr(34)
 strDeviceName = Replace(strDeviceName, strQuotes, "")
 arrDeviceNames = Split(strDeviceName, "=")
 strDeviceName = arrDeviceNames(1)
 Set colUSBDevices = objWMIService.ExecQuery ("Select * From Win32_PnPEntity Where DeviceID = '" & strDeviceName & "'")
 For Each objUSBDevice in colUSBDevices
 y = objUSBDevice.Caption 
 if instr(1,y,"USB Device") then x = x & "USB Device: " & y & vbcrlf
 Next 
Next
f.WriteLine x

f.WriteLine "---------------"


'Get USB details for DISKS
y=""
for i = 0 to 10
DiskIndex=i
  Set objWMIService = GetObject("winmgmts:\\" & strComputer  & "\root\cimv2")
' WMI Query to the Win32_OperatingSystem
    x = "\\\\.\\PHYSICALDRIVE" & DiskIndex  'for a query we must use \\ for a single \
    x = "Select * from Win32_DiskDrive where InterfaceType = 'USB' AND DeviceID = '" & x & "'"
    Set colItems = objWMIService.ExecQuery(x)
    For Each DD In colItems
        y = y & vbcrlf & "Device " & DiskIndex & ":" & DD.Model
        y = y & " FWARE:" & DD.FirmwareRevision
        y = y & " IFACE_TYPE:" & DD.InterfaceType  'USB
        y = y & " MEDIA_TYPE:" &  DD.MediaType
        if not IsNull(DD.Size) then y = y & " SIZE:" &  DD.size
   Next
Next
f.WriteLine y

f.WriteLine "---------------"


'Find USB drives (or a specific drive model - in this case KINGSTON USB drives)
'If you want all USB drives listed, comment out with a ' the  If line and the End If line
strComputer = "."
TargetPath = ""
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colDiskDrives = objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive WHERE InterfaceType = 'USB'")
For Each objDrive In colDiskDrives
    If Instr(1,ucase(objDrive.Caption), "KINGSTON") > 0 Then
 strDeviceID = Replace(objDrive.DeviceID, "\", "\\")
 Set colPartitions = objWMIService.ExecQuery ("ASSOCIATORS OF {Win32_DiskDrive.DeviceID=""" & strDeviceID & """} WHERE AssocClass = " & "Win32_DiskDriveToDiskPartition")
 For Each objPartition In colPartitions
 Set colLogicalDisks = objWMIService.ExecQuery("ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" & objPartition.DeviceID & """} WHERE AssocClass = " & "Win32_LogicalDiskToPartition")
 For Each objLogicalDisk In colLogicalDisks
 TargetPath = TargetPath & objLogicalDisk.DeviceID & vbtab
 Next
 Next
    End If
Next
f.WriteLine "USB Drive(s) mounted at " & TargetPath

f.WriteLine "---------------"


'Show drive letters associated with each
ComputerName = "."
Set wmiServices  = GetObject ( _
    "winmgmts:{impersonationLevel=Impersonate}!//" _
    & ComputerName)
' Get physical disk drive
Set wmiDiskDrives =  wmiServices.ExecQuery ( "SELECT Caption, DeviceID FROM Win32_DiskDrive WHERE InterfaceType = 'USB'")

For Each wmiDiskDrive In wmiDiskDrives
   ' x = wmiDiskDrive.Caption & Vbtab & " " & wmiDiskDrive.DeviceID 

    'Use the disk drive device id to
    ' find associated partition
    query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" & wmiDiskDrive.DeviceID & "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"    
    Set wmiDiskPartitions = wmiServices.ExecQuery(query)

    For Each wmiDiskPartition In wmiDiskPartitions
        'Use partition device id to find logical disk
        Set wmiLogicalDisks = wmiServices.ExecQuery ("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" _
             & wmiDiskPartition.DeviceID & "'} WHERE AssocClass = Win32_LogicalDiskToPartition") 
 x = ""
        For Each wmiLogicalDisk In wmiLogicalDisks
            x = x & wmiDiskDrive.Caption & " " & wmiDiskPartition.DeviceID & " = " & wmiLogicalDisk.DeviceID
 f.WriteLine x

        Next      
    Next
Next
f.Close

'************************************
'** Seting basic email information **
'************************************
Const EmailFrom = "spiceworks@webline.local"
Const EmailTo = "chetanpr@webline.local"
Const EmailSubject = "USB/Mass Storage Notification" 

'***************************************
'** Setting Mail Server Configuration **
'***************************************
Const MailSendUsing = "2"
Const MailSendServer = "mail.webline.local"
Const MailSendPort = "25"
Const MailSendUsername = "chetanpr@webline.local"
Const MailSendPassword = "abcd@1234"
Const MailSendAuthenticationType = "1"

'**************************************
'** Email Parameters (DO NOT CHANGE) **
'**************************************
objEmail.From = EmailFrom
objEmail.To = EmailTo
objEmail.Subject = EmailSubject
objEmail.Textbody = EmailBody
objEmail.AddAttachment EmailAttachments
objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = MailSendUsing
ObjEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MailSendServer
objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = MailSendPort
objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = MailSendAuthenticationType
objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = MailSendUsername
objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = MailSendPassword

'*******************************************************
'** Setting a text file to be shown in the email Body **
'*******************************************************
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
'** File to be inserted in Body
Const FileToBeUsed = "output.txt" 
Dim fs, fo
Set fs = CreateObject("Scripting.FileSystemObject")
'** Open the file for reading
Set fo = fs.OpenTextFile(FileToBeUsed, ForReading)
'** The ReadAll method reads the entire file into the variable BodyText
objEmail.Textbody = fo.ReadAll
'** Close the file
fo.Close
'** Clear variables
Set fo = Nothing
Set fs = Nothing

'* cdoSendUsingPickup (1)
'* cdoSendUsingPort (2)
'* cdoSendUsingExchange (3)

'********************************
'** Parameters (DO NOT CHANGE) **
'********************************
ObjEmail.Configuration.Fields.Update
objEmail.Send 
End Select
Wend 



