REM  *****  BASIC  *****

Attribute VB_Name = "QRReader"

Sub process1QRCodeInput()
    saveData (getInput())
End Sub

Sub process6QRCodeInput()
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
    saveData (getInput())
End Sub

Public Function getInput()
    getInput = InputBox("Scan QR Code", "Match Scouting Input")
End Function

Sub testSaveData()
    saveData ("s=RV;e=2022flwp;l=qm;m=33;r=r2;t=7521;as=[29];at=N;au=0;al=0;ac=N;tu=0;tl=0;wd=N;wbt=N;cif=x;ss=;c=x;lsr=x;be=N;cn=0;ds=x;dr=x;ba=N;d=N;cf=N;all=N;co=;cnf=a;ts=33")
End Sub

Sub saveData(inp As String)
	
	Dim fields
	Dim par
	Dim key As String
	Dim value As String
	Dim values(30)
	
	Dim Doc As Object
	Set Doc = ThisComponent
	Dim sheet As Object
	Set sheet = Doc.Sheets(0)
	
	If inp = "" Then
		Exit Sub
	End If
	
	fields = Split(inp, ";")
	
	Dim i As Integer
	Dim str as String
	
	Dim headed As Boolean
	headed = (sheet.getCellByPosition(0, 0).String <> "") 'col, row
	
	i = 0
	
	If headed = False Then
		For Each str in fields
		
			par = Split(str, "=")
			key = par(0)
			
			If key = "s" Then
				key = "Scouter"
			ElseIf key = "e" Then
				key = "Event"
			ElseIf key = "l" Then
				key = "Level"
			ElseIf key = "m" Then
				key = "Match"
			ElseIf key = "r" Then
				key = "Robot"
			ElseIf key = "t" Then
				key = "Team"
			ElseIf key = "ts" Then
				key = "Total Score"
			EndIf
			
			sheet.getCellByPosition(i, 0).String = key
			
			values(i) = par(1)
			i = i + 1
			
		Next
		
	Else
		For Each str in fields
		
			par = Split(str, "=")
			
			values(i) = par(1)
			i = i + 1
			
		Next
	EndIf
	
	i = 1
	While sheet.getCellByPosition(0, i).String <> ""
		i = i + 1
	Wend
	
	Dim x As Integer
	x = 0
	
	For Each str In values
		sheet.getCellByPosition(x, i).String = str
		x = x + 1
	Next
		
	
End Sub



