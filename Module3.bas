Attribute VB_Name = "Module3"
Sub 貼り付け()
    Dim 日報貼付列 As Long
    Worksheets("プロジェクト時間記録").Activate
    For i = 2 To 2000
        日報貼付列 = 12
        If Worksheets("プロジェクト時間記録").Cells(i, 日報貼付列).Value <> "" Then
            Worksheets("プロジェクト時間記録").Cells(i, 日報貼付列 + 1).Value = Worksheets("プロジェクト時間記録").Cells(i, 日報貼付列).Value
        End If
    Next i
End Sub
