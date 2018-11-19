VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufm日報生成 
   Caption         =   "日報生成"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5400
   OleObjectBlob   =   "ufm日報生成.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufm日報生成"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    'ウィンドウ位置中央(Excel基準)
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - Me.Height - 4
    Me.Left = Application.Left + (Application.Width / 2) - Me.Width - 4
    If Me.Top < 0 Then
        Me.Top = 0
    End If
    If Me.Left < 0 Then
        Me.Left = 0
    End If
    
    '年月日リスト初期化
    cmb年.Clear
    For i = Year(Date) - 2 To Year(Date) + 1
        cmb年.AddItem i
    Next i
    cmb年.Value = Year(Date)

    cmb月.Clear
    For i = 1 To 12
        cmb月.AddItem i
    Next i
    cmb月.Value = Month(Date)
    
    Dim 当月日数, 曜日 As String
    当月日数 = Day(DateSerial(Year(cmb年.Value), Month(cmb月.Value) + 1, 0))
    cmb日.Clear
    For i = 1 To 当月日数
        曜日 = Format(Weekday(DateSerial(cmb年.Value, cmb月.Value, i)), "aaa")
        cmb日.AddItem i
        cmb日.List(cmb日.ListCount - 1, 0) = i
        cmb日.List(cmb日.ListCount - 1, 1) = 曜日
        cmb日.List(cmb日.ListCount - 1, 2) = i & "(" & 曜日 & ")"
    Next i
    cmb日.Value = Day(DateAdd("d", -1, Date))
    '日曜の場合に金曜日に設定
    If cmb日.List(cmb日.ListIndex, 1) = "日" Then
        cmb日.Value = Day(DateAdd("d", -3, Date))
    End If

End Sub

Private Sub btn生成_Click()

    Dim 記録日付 As String
    If cmb年.Value = "" Or cmb月.Value = "" Or cmb日.Value = "" Then
        MsgBox "記録日付を入力してください。", vbExclamation
        cmb年.SetFocus
        Exit Sub
    End If
    記録日付 = cmb年.Value & "/" & Format(cmb月.Value, "00") & "/" & Format(cmb日.Value, "00")
    
    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, True)
    On Error GoTo ErrRSOpen
    
    sql = "SELECT 日報貼付,時間数" _
        & " FROM 時間管理" _
        & " WHERE 記録日付 =#" & 記録日付 & "#" _
        & " AND 削除フラグ = False" _
        & " ORDER BY [開始時間]"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    txt日報生成.Value = ""
    txt合計時間.Value = "0"
    Do Until oRS.EOF
        If txt日報生成.Value <> "" Then
            txt日報生成.Value = Trim(txt日報生成.Value) & vbCrLf & Trim(oRS.Fields("日報貼付").Value)
        Else
            txt日報生成.Value = "進捗など" & vbCrLf & Trim(oRS.Fields("日報貼付").Value)
        End If
        txt合計時間.Value = val(txt合計時間.Value) + val(oRS.Fields("時間数"))
        oRS.MoveNext
    Loop
    
    sql = "SELECT 今後の作業,プロジェクト名" _
        & " FROM チケット管理" _
        & " LEFT JOIN プロジェクト管理 ON プロジェクト管理.プロジェクト番号 = チケット管理.プロジェクト番号" _
        & " WHERE チケット管理.ステータス <> " & ステータス_終了 _
        & " AND チケット管理.削除フラグ <> True" _
        & " ORDER BY チケット管理.プロジェクト番号,開始"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    Dim プロジェクト名 As String
    txt今後作業生成.Value = ""
    Do Until oRS.EOF
        If Not IsNull(oRS.Fields("今後の作業").Value) Then
            If txt今後作業生成.Value <> "" Then
                If プロジェクト名 <> oRS.Fields("プロジェクト名").Value Then
                    txt今後作業生成.Value = Trim(txt今後作業生成.Value) & vbCrLf & vbCrLf & Trim(oRS.Fields("プロジェクト名")) & vbCrLf & Trim(Null2Blank(oRS.Fields("今後の作業").Value))
                Else
                    txt今後作業生成.Value = Trim(txt今後作業生成.Value) & vbCrLf & Trim(Null2Blank(oRS.Fields("今後の作業").Value))
                End If
            Else
                txt今後作業生成.Value = "今後の作業" & vbCrLf & vbCrLf & Trim(oRS.Fields("プロジェクト名").Value) & vbCrLf & Trim(Null2Blank(oRS.Fields("今後の作業").Value))
            End If
            プロジェクト名 = oRS.Fields("プロジェクト名").Value
        End If
        oRS.MoveNext
    Loop
    
    txt合計時間.BackColor = &HFFFFFF
    If txt合計時間.Value <> 7.75 Then
        txt合計時間.BackColor = &H8080FF
    End If
    
    Dim h As Double
    h = val(txt合計時間.Value) - 7.75
    If h > 0 Then
        txt残業時間.Value = h
    Else
        txt残業時間.Value = "0"
    End If
    
    
    If txt日報生成.Value = "" Then
        MsgBox "この日付の日報がありません。ご確認ください。", vbExclamation
    End If
    
    Exit Sub

ErrDataInvalid:
    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    
    MsgBox "データの読出に失敗しました。再度実行してください。", vbExclamation
    Unload Me
    Exit Sub
    
ErrRSOpen:
    'データベース切断
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    
    MsgBox "データの読出に失敗しました。再度実行してください。(" & Err.Number & ")", vbExclamation
    Unload Me
    Exit Sub
    
ErrDBOpen:
    MsgBox "データベースの接続に失敗しました。(" & Err.Number & ")", vbCritical
    Unload Me
    
End Sub

Private Sub cmb月_Change()

    cmb日.Clear
    cmb日.AddItem
    Dim 当月日数, 曜日 As String

    当月日数 = Day(DateSerial(cmb年.Value, cmb月.Value + 1, 0))
    For i = 1 To 当月日数
        曜日 = Format(Weekday(DateSerial(cmb年.Value, cmb月.Value, i)), "aaa")
        cmb日.AddItem
        cmb日.List(cmb日.ListCount - 1, 0) = i
        cmb日.List(cmb日.ListCount - 1, 1) = 曜日
        cmb日.List(cmb日.ListCount - 1, 2) = i & "(" & 曜日 & ")"
    Next i

End Sub
