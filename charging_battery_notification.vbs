set oLocator = CreateObject("WbemScripting.SWbemLocator")
set oServices = oLocator.ConnectServer(".","root\wmi")
set oResults = oServices.ExecQuery("select * from batteryfullchargedcapacity")
Set FSO = CreateObject("Scripting.FileSystemObject")
for each oResult in oResults
   iFull = oResult.FullChargedCapacity
next

YamlFilePath = FSO.GetParentFolderName(WScript.ScriptFullName) & "/charging_battery_notification.properties"
Set File = FSO.OpenTextFile(YamlFilePath, 1)
Set Props = CreateObject("Scripting.Dictionary")
Do While Not File.AtEndOfStream
    line = File.ReadLine
    If InStr(line, "=") > 0 Then
        key = Left(line, InStr(line, "=") - 1)
        value = Mid(line, InStr(line, "=") + 1)
        Props.Add key, value
    End If
Loop
File.Close

threshold = Cint(Props.Item("threshold"))
infoStep = Cint(Props.Item("infoStep"))
sleepMillis = CLng(Props.Item("sleepMillis"))
thresholdWarning = CBool(Props.Item("threshold_warning"))
chargingInfo = CBool(Props.Item("charging_info"))

introMessage = "Script charging_battery_notification started." & vbCrLf & "Loaded properties:"
For Each key In Props.Keys
    introMessage = introMessage & vbCrLf & vbTab & key & ": " & CStr(Props.Item(key))
Next

msgbox introMessage

lastPercent = 0
while (1)
  set oResults = oServices.ExecQuery("select * from batterystatus")
  for each oResult in oResults
    iRemaining = oResult.RemainingCapacity
    bCharging = oResult.Charging
  next
  iPercent = ((iRemaining / iFull) * 100) mod 100
  if thresholdWarning and bCharging and (iPercent >= threshold) Then
    msgbox "Battery level is " & iPercent & "% now. Please stop charging for optimal battery life."
  ElseIf chargingInfo and bCharging and ((iPercent-lastPercent) > infoStep) Then
    lastPercent = Int(iPercent/infoStep)*infoStep
    msgbox "Battery level reached " & lastPercent & "%. Keep charging if you want."
  ElseIf not bCharging Then
    lastPercent = 0
  End If
  wscript.sleep sleepMillis
wend