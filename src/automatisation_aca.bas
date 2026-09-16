Private Sub UserForm_Initialize()
    ' Remplir le ComboBox avec les codes uniques de la colonne Code de la feuille de données
    Dim ws As Worksheet
    Dim rng As Range
    Dim code As Range
    
    Set ws = ThisWorkbook.Sheets("BASE") ' Remplacez "BASE" par le nom de votre feuille de données
    
    ' Recherche des codes uniques dans la colonne Code
    Set rng = ws.Range("D1:C" & ws.Cells(ws.Rows.Count, "C").End(xlUp).Row)
    
    For Each code In rng
        If code.Value <> "" And Application.WorksheetFunction.CountIf(Me.ComboBoxCode.List, code.Value) = 0 Then
            Me.ComboBoxCode.AddItem code.Value
        End If
    Next code
End Sub
