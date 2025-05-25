Attribute VB_Name = "CleanData"
Sub RemoveBlankRows()
    Dim ws As Worksheet
    Dim rng As Range
    Dim i As Long

    Set ws = ActiveSheet
    Application.ScreenUpdating = False

    For i = ws.UsedRange.Rows.Count To 1 Step -1
        If WorksheetFunction.CountA(ws.Rows(i)) = 0 Then
            ws.Rows(i).Delete
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox "Blank rows removed successfully!", vbInformation
End Sub
