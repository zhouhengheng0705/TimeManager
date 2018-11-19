VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufmスケジュール 
   Caption         =   "スケジュール"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5280
   OleObjectBlob   =   "ufmスケジュール.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufmスケジュール"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub UserForm_Initialize()

    'ウィンドウ位置中央(Excel基準)
    If ufmカレンダー.Visible = False Then
        Me.StartUpPosition = 0
        Me.Top = Application.Top + (Application.Height / 2) - Me.Height - 4
        Me.Left = Application.Left + (Application.Width / 2) - Me.Width - 4
        If Me.Top < 0 Then
            Me.Top = 0
        End If
        If Me.Left < 0 Then
            Me.Left = 0
        End If
    Else
        Me.StartUpPosition = 0
        Me.Top = ufmカレンダー.Top
        Me.Left = ufmカレンダー.Left + ufmカレンダー.Width - 10
        Me.Height = 142.5
        Me.Width = 272.5
    End If
    
    'ウィンドウ最小化ボタン有効
    Call FrmDec(Me.Caption, True, True, True)
    
    'リストビュー初期化
    lvスケジュール.ColumnHeaders.Clear
    lvスケジュール.ColumnHeaders.Add , "日付", "日付", 65
    lvスケジュール.ColumnHeaders.Add , "開始時間", "開始時間", 45
    lvスケジュール.ColumnHeaders.Add , "内容", "内容", 180
    lvスケジュール.ColumnHeaders.Add , "スケジュール番号", "スケジュール番号", 0

    '開始終了時間初期化
    cmb開始_時.Clear
    cmb開始_時.AddItem
    For i = 8 To 22
        cmb開始_時.AddItem Format(i, "00")
    Next i

    cmb開始_分.Clear
    cmb開始_分.AddItem
    For i = 0 To 45 Step 15
        cmb開始_分.AddItem Format(i, "00")
    Next i

End Sub

Sub UserForm_Initialize2(日付 As String)

    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, True)
    
    txt日付.Value = 日付
    On Error GoTo ErrRSOpen
    sql = "SELECT *" _
        & " FROM スケジュール" _
        & " WHERE スケジュール.日付 =#" & 日付 & "#"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    lvスケジュール.ListItems.Clear
    Do Until oRS.EOF
        With lvスケジュール.ListItems.Add
            .Text = oRS.Fields("日付").Value
            .SubItems(1) = Format(oRS.Fields("開始時間").Value, "hh:mm")
            .SubItems(2) = oRS.Fields("内容").Value
            .SubItems(3) = oRS.Fields("スケジュール番号").Value
        End With
        oRS.MoveNext
    Loop
    
    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0
    
    Exit Sub
    
ErrRSOpen:
    'データベース切断
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing

    MsgBox "データの読出に失敗しました。再度実行してください。(" & Err.Number & ")", vbExclamation
    Exit Sub

ErrDBOpen:
    MsgBox "データベースの接続に失敗しました。(" & Err.Number & ")", vbCritical

End Sub

Private Sub btn追加_Click()

   '開始時間入力チェック
    Dim 開始時間 As String
    If cmb開始_時.Value = "" Or cmb開始_分.Value = "" Then
        MsgBox "開始時間を入力してください。", vbExclamation
        cmb開始_時.SetFocus
        Exit Sub
    End If
    開始時間 = Format(cmb開始_時.Value, "00") & ":" & Format(cmb開始_分.Value, "00")
    
    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, False)
    On Error GoTo ErrRSOpen
    
    Set oRS = oDB.OpenRecordset("スケジュール", dbOpenTable)
    
    '新規レコード追加
    oRS.AddNew
    txtスケジュール番号.Value = Right(Format(oRS.Fields("ID").Value, "0000"), 4)
    oRS.Fields("スケジュール番号").Value = txtスケジュール番号.Value
    oRS.Fields("日付").Value = Null2Blank(txt日付.Value)
    oRS.Fields("開始時間").Value = Null2Blank(開始時間)
    oRS.Fields("内容").Value = Null2Blank(txt内容.Value)
    oRS.Update
    
    '予定日付リストクリア
    Const 参照シート名 = "予定日付"
    Const 出力開始行 = 2
    Const 出力開始列 = 1
    Dim 参照シート As Worksheet, 最終行 As Long, 最終列 As Long
    Set 参照シート = ThisWorkbook.Worksheets(参照シート名)
    最終行 = 参照シート.UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    最終列 = 参照シート.UsedRange.Columns(参照シート.UsedRange.Columns.Count).Column
    If 最終行 >= 出力開始行 Then
        '再読込が必要場合のみクリア
        参照シート.Rows(出力開始行 & ":" & 最終行).Delete Shift:=xlUp
    End If
    
    Dim sql As String
    '予定日付リスト取得
    sql = "SELECT 日付" _
        & " FROM スケジュール"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    参照シート.Cells(出力開始行, 出力開始列).CopyFromRecordset oRS
    

    Call ufmカレンダー.初期化

    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0
    
    MsgBox "データの登録に成功しました", vbInformation
    
    UserForm_Initialize2 txt日付.Value
    
    Exit Sub
    
ErrRSOpen:
    'データベース切断
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    
    MsgBox "データの登録に失敗しました。再度実行してください。(" & Err.Number & ")", vbExclamation
    Exit Sub
    
ErrDBOpen:
    MsgBox "データベースの接続に失敗しました。(" & Err.Number & ")", vbCritical

End Sub

Private Sub btn削除_Click()

    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, False)
    
    Dim スケジュール番号 As String, n As Long
    n = 0
    If lvスケジュール.ListItems.Count <> 0 Then
        For i = 1 To lvスケジュール.ListItems.Count
            If lvスケジュール.ListItems.Item(i).Checked = True Then
                スケジュール番号 = lvスケジュール.ListItems(i).SubItems(3)
                On Error GoTo ErrRSOpen
                Set oRS = oDB.OpenRecordset("スケジュール", dbOpenTable)
                oRS.Index = "スケジュール番号検索"
                oRS.Seek "=", スケジュール番号
                oRS.Delete
                n = n + 1
            End If
        Next i
    Else
        Exit Sub
    End If
    
    If n = 0 Then
        MsgBox "削除するデータを選択してください。", vbExclamation
        Exit Sub
    End If
    
    '予定日付リストクリア
    Const 参照シート名 = "予定日付"
    Const 出力開始行 = 2
    Const 出力開始列 = 1
    Dim 参照シート As Worksheet, 最終行 As Long, 最終列 As Long
    Set 参照シート = ThisWorkbook.Worksheets(参照シート名)
    最終行 = 参照シート.UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    最終列 = 参照シート.UsedRange.Columns(参照シート.UsedRange.Columns.Count).Column
    If 最終行 >= 出力開始行 Then
        '再読込が必要場合のみクリア
        参照シート.Rows(出力開始行 & ":" & 最終行).Delete Shift:=xlUp
    End If
    
    Dim sql As String
    '予定日付リスト取得
    sql = "SELECT 日付" _
        & " FROM スケジュール"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    参照シート.Cells(出力開始行, 出力開始列).CopyFromRecordset oRS
    
    For i = 1 To 42
        If 予定チェック(ufmカレンダー.Controls("Label" & i).Tag) Then
            ufmカレンダー.Controls("Label" & i).BackColor = &HFF80FF
        End If
    Next i
    
    Call ufmカレンダー.初期化
    
    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0
    
    MsgBox "選択したデータを削除しました。", vbInformation
    
    UserForm_Initialize2 txt日付.Value
    
    Exit Sub
    
ErrRSOpen:
    'データベース切断
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    
    MsgBox "データの登録に失敗しました。再度実行してください。(" & Err.Number & ")", vbExclamation
    Exit Sub
    
ErrDBOpen:
    MsgBox "データベースの接続に失敗しました。(" & Err.Number & ")", vbCritical

End Sub

Private Sub btn最小化_Click()
    
    Me.Height = 142.5
    Me.Width = 272.5
    btn最小化.Visible = False
    btn最大化.Visible = True

End Sub

Private Sub btn最大化_Click()
    
    Me.Height = 227.5
    Me.Width = 272.5
    btn最大化.Visible = False
    btn最小化.Visible = True

End Sub

