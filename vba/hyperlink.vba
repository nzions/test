Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    ' Makes the follow links work

    If Target.Name = "Go >>>" Then

        Dim myRow As Integer
        myRow = Target.Range.Row

        Cells(4, "H") = Cells(myRow, "A").Value

        ' a dodgey way
        Sheet2.Cells(4, "H").Value = Cells(myRow, "A").Value

        ' a better way
        Sheets("my worksheet").Cells(4, "H").Value = Cells(myRow, "A").Value
       Sheets("my worksheet").Activate
       ActiveSheet.Cells(4, "H").Select
    End If
End Sub