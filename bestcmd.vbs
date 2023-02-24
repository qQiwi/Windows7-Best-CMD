option explicit
Dim vrt10, objWSH, strCommand, objA, strMsg, FSO, TextStream
vrt10 = inputbox("tasklist = cmd"&vbCr&"network = ping"&vbCr&"IP = ipconfig", "cmd")
If vrt10 <> "" Then
If (vrt10="cmd") or (vrt10="/cmd") then
set objWSH = CreateObject("WScript.Shell")
strCommand = "TASKLIST"
set objA = objWSH.Exec(strCommand)
strMsg = objA.StdOut.ReadAll()
WScript.Echo(strMsg)

ElseIf (vrt10="ping") or (vrt10="/ping") then
set objWSH = CreateObject("WScript.Shell")
strCommand = "ping"
set objA = objWSH.Exec(strCommand)
strMsg = objA.StdOut.ReadAll()
WScript.Echo(strMsg)

ElseIf (vrt10="ipconfig") or (vrt10="/ipconfig") then
set objWSH = CreateObject("WScript.Shell")
strCommand = "ipconfig"
set objA = objWSH.Exec(strCommand)
strMsg = objA.StdOut.ReadAll()
WScript.Echo(strMsg)

else
set objWSH = CreateObject("WScript.Shell")
strCommand = vrt10
set objA = objWSH.Exec(strCommand)
strMsg = objA.StdOut.ReadAll()
WScript.Echo(strMsg)
end if
Set FSO = CreateObject("Scripting.FileSystemObject")
Set TextStream = FSO.CreateTextFile(vrt10 + ".txt",True)
TextStream.WriteLine(strMsg)
TextStream.Close
set TextStream = Nothing
set FSO = Nothing
Else
End If
