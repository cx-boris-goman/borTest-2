
Option Explicit





'This is separate from clsSecurity so it can be shared with other apps.
Public Function DailyPassword(Optional pdtDate As Date) As String
    Dim hardCodedPassword As String
    Dim hardCodedPasswordCopy As String	
	
	hardCodedPassword = "NotSafe!"
	hardCodedPasswordCopy = hardCodedPassword
     
    'DailyPassword = sRet_param

	Dim Con As ADODB.Connection
	Set Con = New ADODB.Connection
	
	Dim t
	t = InputBox()
	
	Dim Con as ADODB.connection 
    ' Execute
    Cmd.Execute(t)
	'Cmd.Execute(hardCodedPasswordCopy)

	Con.Execute(t)

	'DailyPassword = Con.Open(hardCodedPassword)

	Con.Open(hardCodedPassword)

End Function

