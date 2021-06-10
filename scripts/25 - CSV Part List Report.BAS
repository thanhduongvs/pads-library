Sub Main
	Set wshell = CreateObject("WScript.Shell")
	Debug.Print wshell.CurrentDirectory
	Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim TextFile As Object
 	Dim index As Integer
    Set TextFile = FSO.CreateTextFile(wshell.CurrentDirectory+"\pos.csv", True, True)
    
    TextFile.Write "Item" & ","
    TextFile.Write "Reference" & ","
	TextFile.Write "PartType" & ","
	TextFile.Write "Layer" & ","
	TextFile.Write "Orientation" & ","
	TextFile.Write "X" & ","
	TextFile.Write "Y" & ","
	TextFile.Write "SMD" & ","
	TextFile.Write vbCrLf
	
    index = 1
	'Output table rows
	For Each part In ActiveDocument.Components
		TextFile.Write index & ","
		
		TextFile.Write part.Name & ","
		TextFile.Write part.PartType & ","
		TextFile.Write ActiveDocument.LayerName(part.layer) & ","
		TextFile.Write part.Orientation & ","
		TextFile.Write Format(part.PositionX, "0.000") & ","
		TextFile.Write Format(part.PositionY, "0.000") & ","
		TextFile.Write Format(part.IsSMD, "Yes/No") & ","
		TextFile.Write vbCrLf
		index = index + 1
	Next part

End Sub