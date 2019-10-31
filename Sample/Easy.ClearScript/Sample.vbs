Set fso = CreateObject ("Scripting.FileSystemObject")

Set stdout = fso.GetStandardStream (1)
Set stderr = fso.GetStandardStream (2)

stdout.WriteLine "This will go to standard output."
stderr.WriteLine "This will go to error output."

' Open a new debugger instance and attach it to this script (script mode must be selected)
stop

result = MsgBox("Ready?", vbQuestion + vbYesNo, "Let's Go!")
if (result = vbYes) then
    stdout.WriteLine("yes=" & result)
else
    stdout.WriteLine("no=" & result)
end if
