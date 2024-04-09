Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
	On Error Resume Next

	' Sheets
	Dim setSheet As Worksheet
	Set setSheet = ThisWorkbook.Sheets("Set")

	Dim settingsSheet As Worksheet
	Set settingsSheet = ThisWorkbook.Sheets("Settings")

	' Define last row
	Dim lastRow As Long
	lastRow = setSheet.Cells(Rows.Count, "C").End(xlUp).Row

	' Define sort type
	' 0: Sort by All found, then by Amount Left, then by Set Order
	' 1: Sort by Set Order
	' 2: Sort by Part Name
	Dim sortType As Integer
	sortType = settingsSheet.Range("A5").Value

	If (Target.Column <> 2 Or Target.Row = 1) And (Sh.Name <> "Settings" Or Target.Address <> settingsSheet.Range("A5").Address) Then Exit Sub
	If Target.Count > 1 Then Exit Sub

	If sortType = 2 Then
		' Sort by Set Order
		setSheet.Range("A2:I" & lastRow).Sort Key1:=setSheet.Range("H2"), Order1:=xlAscending, _
			Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
	ElseIf sortType = 3 Then
		' Sort by Part Name
		setSheet.Range("A2:I" & lastRow).Sort Key1:=setSheet.Range("G2"), Order1:=xlAscending, _
			Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
	Else
		' Sort by All found, then by Amount Left, then by Set Order
		setSheet.Range("A2:I" & lastRow).Sort Key1:=setSheet.Range("I2"), Order1:=xlAscending, _
			Key2:=setSheet.Range("C2"), Order2:=xlAscending, Key3:=setSheet.Range("H2"), Order3:=xlAscending, _
			Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
	End If
End Sub
