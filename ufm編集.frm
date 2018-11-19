VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufm編集 
   Caption         =   "編集"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7635
   OleObjectBlob   =   "ufm編集.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufm編集"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const 勤務設定_本社 = 1
Const 勤務設定_日勤早番 = 2
Const 勤務設定_日勤遅番 = 3
Const 勤務設定_スライド = 4
Const 勤務設定_本社10時 = 5

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
    cmb編集_年.Clear
    For i = Year(Date) - 2 To Year(Date) + 1
        cmb編集_年.AddItem i
    Next i
    cmb編集_年.Value = Year(Date)

    cmb編集_月.Clear
    For i = 1 To 12
        cmb編集_月.AddItem i
    Next i
    cmb編集_月.Value = Month(Date)

    cmb編集_日.Clear
    Dim 当月日数, 曜日 As String
    当月日数 = Day(DateSerial(Year(cmb編集_年.Value), Month(cmb編集_月.Value) + 1, 0))
    For i = 1 To 当月日数
        曜日 = Format(Weekday(DateSerial(cmb編集_年.Value, cmb編集_月.Value, i)), "aaa")
        cmb編集_日.AddItem
        cmb編集_日.List(cmb編集_日.ListCount - 1, 0) = i
        cmb編集_日.List(cmb編集_日.ListCount - 1, 1) = 曜日
        cmb編集_日.List(cmb編集_日.ListCount - 1, 2) = i & "(" & 曜日 & ")"
    Next i
    cmb編集_日.Value = Day(DateAdd("d", -1, Date))
    '日曜の場合に金曜日に設定
    If cmb編集_日.List(cmb編集_日.ListIndex, 1) = "日" Then
        cmb編集_日.Value = Day(DateAdd("d", -3, Date))
    Else
        cmb編集_日.Value = Day(DateAdd("d", -1, Date))
    End If
    
    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, True)
    On Error GoTo ErrRSOpen
    
    '勤務設定初期化
    sql = "SELECT 項目名,値" _
        & " FROM V_勤務設定"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    Do Until oRS.EOF
        cmb勤務設定.AddItem
        cmb勤務設定.List(cmb勤務設定.ListCount - 1, 0) = oRS.Fields("値")
        cmb勤務設定.List(cmb勤務設定.ListCount - 1, 1) = oRS.Fields("項目名")
        oRS.MoveNext
    Loop
    cmb勤務設定.Value = cmb勤務設定.List(0, 0)
    
        
    'プロジェクトリスト初期化
    sql = "SELECT プロジェクト名, プロジェクト番号,項目名" _
        & " FROM プロジェクト管理 LEFT JOIN V_部門コード ON V_部門コード.値 = プロジェクト管理.部門コード"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    cmbプロジェクト.Clear
    cmbプロジェクト.AddItem
    Do Until oRS.EOF
        cmbプロジェクト.AddItem
        cmbプロジェクト.List(cmbプロジェクト.ListCount - 1, 0) = oRS.Fields("プロジェクト番号").Value
        cmbプロジェクト.List(cmbプロジェクト.ListCount - 1, 1) = oRS.Fields("プロジェクト名").Value
        cmbプロジェクト.List(cmbプロジェクト.ListCount - 1, 2) = oRS.Fields("項目名").Value
        oRS.MoveNext
    Loop
    cmbプロジェクト.Value = cmbプロジェクト.List(1, 0)
    
    
    'チケット名リスト取得
    sql = "SELECT チケット番号,チケット名,(チケット番号 & "" "" & チケット名) As 表示名,ステータス,項目名" _
        & " FROM チケット管理" _
        & " LEFT JOIN V_ステータス ON val(V_ステータス.値) = チケット管理.ステータス" _
        & " WHERE チケット管理.チケット番号 <> '" & "" & "'" _
        & " AND チケット管理.プロジェクト番号 ='" & cmbプロジェクト.Value & "'" _
        & " ORDER BY チケット管理.ステータス,チケット番号"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    cmbチケット名.Clear
    Do Until oRS.EOF
        cmbチケット名.AddItem
        cmbチケット名.List(cmbチケット名.ListCount - 1, 0) = oRS.Fields("項目名").Value
        cmbチケット名.List(cmbチケット名.ListCount - 1, 1) = oRS.Fields("チケット番号").Value
        cmbチケット名.List(cmbチケット名.ListCount - 1, 2) = oRS.Fields("チケット名").Value
        cmbチケット名.List(cmbチケット名.ListCount - 1, 3) = oRS.Fields("表示名").Value
        oRS.MoveNext
    Loop

    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0
    
    '開始終了時間初期化
    cmb開始_時.Clear
    cmb終了_時.Clear
    cmb開始_時.AddItem
    cmb終了_時.AddItem
    For i = 8 To 22
        cmb開始_時.AddItem Format(i, "00")
        cmb終了_時.AddItem Format(i, "00")
    Next i

    cmb開始_分.Clear
    cmb終了_分.Clear
    cmb開始_分.AddItem
    cmb終了_分.AddItem
    For i = 0 To 45 Step 15
        cmb開始_分.AddItem Format(i, "00")
        cmb終了_分.AddItem Format(i, "00")
    Next i
    
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

Private Sub cmbプロジェクト_Change()

    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, False)
    On Error GoTo ErrRSOpen


    'チケット名リスト取得
    sql = "SELECT チケット番号,チケット名,(チケット番号 & "" "" & チケット名) As 表示名,ステータス,項目名" _
        & " FROM チケット管理" _
        & " LEFT JOIN V_ステータス ON val(V_ステータス.値) = チケット管理.ステータス" _
        & " WHERE チケット管理.チケット番号 <> '" & "" & "'" _
        & " AND チケット管理.プロジェクト番号 ='" & cmbプロジェクト.Value & "'" _
        & " ORDER BY チケット管理.ステータス"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    cmbチケット名.Clear
    Do Until oRS.EOF
        cmbチケット名.AddItem
        cmbチケット名.List(cmbチケット名.ListCount - 1, 0) = oRS.Fields("項目名").Value
        cmbチケット名.List(cmbチケット名.ListCount - 1, 1) = oRS.Fields("チケット番号").Value
        cmbチケット名.List(cmbチケット名.ListCount - 1, 2) = oRS.Fields("チケット名").Value
        cmbチケット名.List(cmbチケット名.ListCount - 1, 3) = oRS.Fields("表示名").Value
        oRS.MoveNext
    Loop
    
    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0
    
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

Sub UserForm_Initialize2(記録番号 As String)
    
    txt記録番号.Value = 記録番号
    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, True)
    On Error GoTo ErrRSOpen
    
    'レコードセット取得
    Set oRS = oDB.OpenRecordset("時間管理", dbOpenTable)
    oRS.Index = "記録番号"
    oRS.Seek "=", txt記録番号.Value
    If oRS.NoMatch Then
        GoTo ErrDataInvalid
    End If
    
    cmbチケット名.Value = Null
    
    '編集部へ出力
    cmb編集_年.Value = Year(oRS.Fields("記録日付").Value)
    cmb編集_月.Value = Month(oRS.Fields("記録日付").Value)
    cmb編集_日.Value = Day(oRS.Fields("記録日付").Value)
    cmb勤務設定.Value = oRS.Fields("勤務設定").Value
    cmbプロジェクト.Value = oRS.Fields("プロジェクト番号").Value
    cmbチケット名.Value = Null2Blank(oRS.Fields("チケット番号"))
    cmb開始_時.Value = Format(Hour(oRS.Fields("開始時間")), "00")
    cmb開始_分.Value = Format(Minute(oRS.Fields("開始時間")), "00")
    cmb終了_時.Value = Format(Hour(oRS.Fields("終了時間")), "00")
    cmb終了_分.Value = Format(Minute(oRS.Fields("終了時間")), "00")
    cb無効.Value = oRS.Fields("削除フラグ").Value
    txtコメント.Value = oRS.Fields("コメント").Value
    
    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0
    
    '更新ボタン有効化
    btn追加.Visible = False
    btn更新.Visible = True
    
    Exit Sub
    
ErrDataInvalid:
    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    
    MsgBox "データの読出に失敗しました。最新の情報を確認してください。", vbExclamation
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

    '日付入力チェック
    Dim 記録日付 As String
    If cmb編集_年.Value = "" Or cmb編集_月 = "" Or cmb編集_日 = "" Then
        MsgBox "記録日付を入力してください。", vbExclamation
    End If
    記録日付 = cmb編集_年.Value & "/" & Format(cmb編集_月.Value, "00") & "/" & Format(cmb編集_日, "00")
    
    '勤務設定入力チェック
    Dim 勤務設定 As Long
    If cmb勤務設定.Text = "" Then
        MsgBox "勤務設定を入力してください。", vbExclamation
        cmb勤務設定.SetFocus
        Exit Sub
    Else
        勤務設定 = cmb勤務設定.Value
    End If
    
    'プロジェクト入力チェック
    Dim プロジェクト As String
    If cmbプロジェクト.Text = "" Then
        プロジェクト = ""
    ElseIf IsNull(cmbプロジェクト.Value) Then
        If Not プロジェクト書式チェック(cmbプロジェクト.Text) Then
            MsgBox "プロジェクト入力が不正です。", vbExclamation
            cmbプロジェクト.SetFocus
            Exit Sub
        Else
            プロジェクト = cmbプロジェクト.Text
        End If
    Else
        プロジェクト = cmbプロジェクト.Value
    End If
    
    'チケット名入力チェック
    Dim チケット名 As String
    If cmbチケット名.Text = "" Then
        チケット名 = ""
    ElseIf IsNull(cmbチケット名.Value) Then
        If Not チケット名書式チェック(cmbチケット名.Text) Then
            MsgBox "チケット名入力が不正です。", vbExclamation
            cmbチケット名.SetFocus
            Exit Sub
        Else
            チケット名 = cmbチケット名.Text
        End If
    Else
        チケット名 = cmbチケット名.Value
    End If
    
    '開始時間入力チェック
    Dim 開始時間 As String
    If cmb開始_時.Value = "" Or cmb開始_分.Value = "" Then
        MsgBox "開始時間を入力してください。", vbExclamation
        cmb開始_時.SetFocus
        Exit Sub
    End If
    開始時間 = Format(cmb開始_時.Value, "00") & ":" & Format(cmb開始_分.Value, "00")
    
    '終了時間入力チェック
    Dim 終了時間 As String
    If cmb終了_時.Value = "" Or cmb終了_分.Value = "" Then
        MsgBox "終了時間を入力してください。", vbExclamation
        cmb終了_時.SetFocus
        Exit Sub
    End If
    終了時間 = Format(cmb終了_時.Value, "00") & ":" & Format(cmb終了_分.Value, "00")
    
    If 開始時間 > 終了時間 Then
        MsgBox "終了時間を入力してください。", vbExclamation
        cmb終了_時.SetFocus
        Exit Sub
    End If
    
    Dim 時間数 As Double
    If txt時間数.Value <> "" Then
        時間数 = txt時間数.Value
    End If
    
    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, False)
    On Error GoTo ErrRSOpen
    
    Set oRS = oDB.OpenRecordset("時間管理", dbOpenTable)
    
    'トランザクション開始
    oWks.BeginTrans
    
    '新規レコード追加
    oRS.AddNew
    txt記録番号.Value = "K" & Format(Now(), "yyyymmdd") & "-" & Right(Format(oRS.Fields("ID").Value, "0000"), 4)
    oRS.Fields("記録番号").Value = txt記録番号.Value
    oRS.Fields("記録日付").Value = 記録日付
    oRS.Fields("プロジェクト番号").Value = プロジェクト
    oRS.Fields("チケット番号").Value = チケット名
    oRS.Fields("開始時間").Value = 開始時間
    oRS.Fields("終了時間").Value = 終了時間
    oRS.Fields("時間数").Value = Null2Blank(時間数)
    oRS.Fields("勤務設定").Value = 勤務設定
    oRS.Fields("コメント").Value = txtコメント.Value
    oRS.Fields("削除フラグ").Value = cb無効.Value
    If Left(oRS.Fields("チケット番号").Value, 1) = "#" Then
        oRS.Fields("日報貼付").Value = 開始時間 & "～" & 終了時間 & "[" & Format(時間数, "00.00") & "H" & "]" & cmbチケット名.List(cmbチケット名.ListIndex, 3)
    Else
        oRS.Fields("日報貼付").Value = 開始時間 & "～" & 終了時間 & "[" & Format(時間数, "00.00") & "H" & "]" & Trim(txtコメント.Value)
    End If
    oRS.Update
    
    'トランザクション完了
    oWks.CommitTrans
    
    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0
    
    ufm時間管理ツール.btn検索_Click
    MsgBox "データの登録に成功しました。", vbInformation
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

Private Sub btn更新_Click()

    '日付入力チェック
    Dim 記録日付 As String
    If cmb編集_年.Value = "" Or cmb編集_月 = "" Or cmb編集_日 = "" Then
        MsgBox "記録日付を入力してください。", vbExclamation
    End If
    記録日付 = cmb編集_年.Value & "/" & Format(cmb編集_月.Value, "00") & "/" & Format(cmb編集_日, "00")
    
    '勤務設定入力チェック
    Dim 勤務設定 As Long
    If cmb勤務設定.Text = "" Then
        MsgBox "勤務設定を入力してください。", vbExclamation
        cmb勤務設定.SetFocus
        Exit Sub
    Else
        勤務設定 = cmb勤務設定.Value
    End If
    
    
    'プロジェクト入力チェック
    Dim プロジェクト As String
    If cmbプロジェクト.Text = "" Then
        プロジェクト = ""
    ElseIf IsNull(cmbプロジェクト.Value) Then
        If Not プロジェクト書式チェック(cmbプロジェクト.Text) Then
            MsgBox "プロジェクト入力が不正です。", vbExclamation
            cmbプロジェクト.SetFocus
            Exit Sub
        Else
            プロジェクト = cmbプロジェクト.Text
        End If
    Else
        プロジェクト = cmbプロジェクト.Value
    End If
    
    'チケット名入力チェック
    Dim チケット名 As String
    If cmbチケット名.Text = "" Then
        チケット名 = ""
    ElseIf IsNull(cmbチケット名.Value) Then
        If Not チケット名書式チェック(cmbチケット名.Text) Then
            MsgBox "チケット名入力が不正です。", vbExclamation
            cmbチケット名.SetFocus
            Exit Sub
        Else
            チケット名 = cmbチケット名.Text
        End If
    Else
        チケット名 = cmbチケット名.Value
    End If
    
    '開始時間入力チェック
    Dim 開始時間 As String
    If cmb開始_時.Value = "" Or cmb開始_分.Value = "" Then
        MsgBox "開始時間を入力してください。", vbExclamation
        cmb開始_時.SetFocus
        Exit Sub
    End If
    開始時間 = Format(cmb開始_時.Value, "00") & ":" & Format(cmb開始_分.Value, "00")
    
    '終了時間入力チェック
    Dim 終了時間 As String
    If cmb終了_時.Value = "" Or cmb終了_分.Value = "" Then
        MsgBox "終了時間を入力してください。", vbExclamation
        cmb終了_時.SetFocus
        Exit Sub
    End If
    終了時間 = Format(cmb終了_時.Value, "00") & ":" & Format(cmb終了_分.Value, "00")
    
    If 開始時間 > 終了時間 Then
        MsgBox "終了時間を入力してください。", vbExclamation
        cmb終了_時.SetFocus
        Exit Sub
    End If
    
    Dim 時間数 As Double
    If txt時間数.Value <> "" Then
        時間数 = txt時間数.Value
    End If
    
    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, False)
    On Error GoTo ErrRSOpen
    
    Set oRS = oDB.OpenRecordset("時間管理", dbOpenTable)
    oRS.Index = "記録番号"
    oRS.Seek "=", txt記録番号.Value
    If oRS.NoMatch Then
        GoTo ErrDataInvalid
    End If
    
    'トランザクション開始
    oWks.BeginTrans
    
    '新規レコード追加
    oRS.Edit
    oRS.Fields("記録日付").Value = 記録日付
    oRS.Fields("プロジェクト番号").Value = プロジェクト
    oRS.Fields("チケット番号").Value = チケット名
    oRS.Fields("開始時間").Value = 開始時間
    oRS.Fields("終了時間").Value = 終了時間
    oRS.Fields("時間数").Value = Null2Blank(時間数)
    oRS.Fields("勤務設定").Value = 勤務設定
    oRS.Fields("コメント").Value = txtコメント.Value
    oRS.Fields("削除フラグ").Value = cb無効.Value
    If Left(oRS.Fields("チケット番号").Value, 1) <> "" Then
        oRS.Fields("日報貼付").Value = 開始時間 & "～" & 終了時間 & "[" & Format(時間数, "00.00") & "H" & "]" & cmbチケット名.List(cmbチケット名.ListIndex, 3)
    Else
        oRS.Fields("日報貼付").Value = 開始時間 & "～" & 終了時間 & "[" & Format(時間数, "00.00") & "H" & "]" & Trim(txtコメント.Value)
    End If
    oRS.Update
    
    'トランザクション完了
    oWks.CommitTrans
    
    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0
    
    ufm時間管理ツール.btn検索_Click
    MsgBox "データの更新に成功しました。", vbInformation
    Exit Sub
    
ErrDataInvalid:
    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    
    MsgBox "記録案件が未登録です。", vbExclamation
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

Private Sub cmb開始_時_Change()

    Dim 参照シート As Worksheet
    Set 参照シート = ThisWorkbook.Worksheets("出勤時間設定")
    
    If cmb開始_時.Value = "" Or cmb開始_分.Value = "" Or cmb終了_時.Value = "" Or cmb終了_分.Value = "" Then
        Exit Sub
    End If
    
    Dim 開始時刻, 終了時刻 As String
    開始時刻 = Format(cmb開始_時.Value, "00") & ":" & Format(cmb開始_分.Value, "00")
    終了時刻 = Format(cmb終了_時.Value, "00") & ":" & Format(cmb終了_分.Value, "00")
    If 開始時刻 > 終了時刻 Then
        Exit Sub
    End If
    
    Dim 開始時刻数値, 終了時刻数値, 開始時間数, 終了時間数 As Double
    
    開始時刻数値 = val(cmb開始_時.Value) + val(cmb開始_分.Value / 60)
    終了時刻数値 = val(cmb終了_時.Value) + val(cmb終了_分.Value / 60)
    
    
    Select Case cmb勤務設定.Value
        
        '----------------'
        Case 勤務設定_本社
        
        For i = 3 To 61
            If 開始時刻数値 = 参照シート.Cells(i, 2).Value Then
                開始時間数 = 参照シート.Cells(i, 3).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 2).Value Then
                終了時間数 = 参照シート.Cells(i, 3).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
    
        '----------------'
        Case 勤務設定_日勤早番
        
        For i = 3 To 63
            If 開始時刻数値 = 参照シート.Cells(i, 4).Value Then
                開始時間数 = 参照シート.Cells(i, 5).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 4).Value Then
                終了時間数 = 参照シート.Cells(i, 5).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
        '----------------'
        Case 勤務設定_スライド
        
        For i = 3 To 43
            If 開始時刻数値 = 参照シート.Cells(i, 6).Value Then
                開始時間数 = 参照シート.Cells(i, 7).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 6).Value Then
                終了時間数 = 参照シート.Cells(i, 7).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
        '-----------------'
        Case 勤務設定_本社10時
        
        For i = 3 To 49
            If 開始時刻数値 = 参照シート.Cells(i, 8).Value Then
                開始時間数 = 参照シート.Cells(i, 9).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 49
            If 終了時刻数値 = 参照シート.Cells(i, 8).Value Then
                終了時間数 = 参照シート.Cells(i, 9).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
    End Select
        
End Sub

Private Sub cmb開始_分_Change()

    Dim 参照シート As Worksheet
    Set 参照シート = ThisWorkbook.Worksheets("出勤時間設定")
    
    If cmb開始_時.Value = "" Or cmb開始_分.Value = "" Or cmb終了_時.Value = "" Or cmb終了_分.Value = "" Then
        Exit Sub
    End If
    
    Dim 開始時刻, 終了時刻 As String
    開始時刻 = Format(cmb開始_時.Value, "00") & ":" & Format(cmb開始_分.Value, "00")
    終了時刻 = Format(cmb終了_時.Value, "00") & ":" & Format(cmb終了_分.Value, "00")
    If 開始時刻 > 終了時刻 Then
        Exit Sub
    End If
    
    Dim 開始時刻数値, 終了時刻数値, 開始時間数, 終了時間数 As Double
    
    開始時刻数値 = val(cmb開始_時.Value) + val(cmb開始_分.Value / 60)
    終了時刻数値 = val(cmb終了_時.Value) + val(cmb終了_分.Value / 60)
    
    
    Select Case cmb勤務設定.Value
        
        '----------------'
        Case 勤務設定_本社
        
        For i = 3 To 61
            If 開始時刻数値 = 参照シート.Cells(i, 2).Value Then
                開始時間数 = 参照シート.Cells(i, 3).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 2).Value Then
                終了時間数 = 参照シート.Cells(i, 3).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
        '----------------'
        Case 勤務設定_日勤早番
        
        For i = 3 To 63
            If 開始時刻数値 = 参照シート.Cells(i, 4).Value Then
                開始時間数 = 参照シート.Cells(i, 5).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 4).Value Then
                終了時間数 = 参照シート.Cells(i, 5).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
        '----------------'
        Case 勤務設定_スライド
        
        For i = 3 To 43
            If 開始時刻数値 = 参照シート.Cells(i, 6).Value Then
                開始時間数 = 参照シート.Cells(i, 7).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 6).Value Then
                終了時間数 = 参照シート.Cells(i, 7).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
        '-----------------'
        Case 勤務設定_本社10時
        
        For i = 3 To 49
            If 開始時刻数値 = 参照シート.Cells(i, 8).Value Then
                開始時間数 = 参照シート.Cells(i, 9).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 49
            If 終了時刻数値 = 参照シート.Cells(i, 8).Value Then
                終了時間数 = 参照シート.Cells(i, 9).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
    End Select
        
End Sub

Private Sub cmb終了_時_Change()

    Dim 参照シート As Worksheet
    Set 参照シート = ThisWorkbook.Worksheets("出勤時間設定")
    
    If cmb開始_時.Value = "" Or cmb開始_分.Value = "" Or cmb終了_時.Value = "" Or cmb終了_分.Value = "" Then
        Exit Sub
    End If
    
    Dim 開始時刻, 終了時刻 As String
    開始時刻 = Format(cmb開始_時.Value, "00") & ":" & Format(cmb開始_分.Value, "00")
    終了時刻 = Format(cmb終了_時.Value, "00") & ":" & Format(cmb終了_分.Value, "00")
    If 開始時刻 > 終了時刻 Then
        Exit Sub
    End If
    
    Dim 開始時刻数値, 終了時刻数値, 開始時間数, 終了時間数 As Double
    
    開始時刻数値 = val(cmb開始_時.Value) + val(cmb開始_分.Value / 60)
    終了時刻数値 = val(cmb終了_時.Value) + val(cmb終了_分.Value / 60)
    
    
    Select Case cmb勤務設定.Value
        
        '----------------'
        Case 勤務設定_本社
        
        For i = 3 To 61
            If 開始時刻数値 = 参照シート.Cells(i, 2).Value Then
                開始時間数 = 参照シート.Cells(i, 3).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 2).Value Then
                終了時間数 = 参照シート.Cells(i, 3).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
        '----------------'
        Case 勤務設定_日勤早番
        
        For i = 3 To 63
            If 開始時刻数値 = 参照シート.Cells(i, 4).Value Then
                開始時間数 = 参照シート.Cells(i, 5).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 4).Value Then
                終了時間数 = 参照シート.Cells(i, 5).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
        '----------------'
        Case 勤務設定_スライド
        
        For i = 3 To 43
            If 開始時刻数値 = 参照シート.Cells(i, 6).Value Then
                開始時間数 = 参照シート.Cells(i, 7).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 6).Value Then
                終了時間数 = 参照シート.Cells(i, 7).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
        '-----------------'
        Case 勤務設定_本社10時
        
        For i = 3 To 49
            If 開始時刻数値 = 参照シート.Cells(i, 8).Value Then
                開始時間数 = 参照シート.Cells(i, 9).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 49
            If 終了時刻数値 = 参照シート.Cells(i, 8).Value Then
                終了時間数 = 参照シート.Cells(i, 9).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
    End Select
        
End Sub

Private Sub cmb終了_分_Change()

    Dim 参照シート As Worksheet
    Set 参照シート = ThisWorkbook.Worksheets("出勤時間設定")
    
    If cmb開始_時.Value = "" Or cmb開始_分.Value = "" Or cmb終了_時.Value = "" Or cmb終了_分.Value = "" Then
        Exit Sub
    End If
    
    Dim 開始時刻, 終了時刻 As String
    開始時刻 = Format(cmb開始_時.Value, "00") & ":" & Format(cmb開始_分.Value, "00")
    終了時刻 = Format(cmb終了_時.Value, "00") & ":" & Format(cmb終了_分.Value, "00")
    If 開始時刻 > 終了時刻 Then
        Exit Sub
    End If
    
    Dim 開始時刻数値, 終了時刻数値, 開始時間数, 終了時間数 As Double
    
    開始時刻数値 = val(cmb開始_時.Value) + val(cmb開始_分.Value / 60)
    終了時刻数値 = val(cmb終了_時.Value) + val(cmb終了_分.Value / 60)
    
    
    Select Case cmb勤務設定.Value
        
        '----------------'
        Case 勤務設定_本社
        
        For i = 3 To 61
            If 開始時刻数値 = 参照シート.Cells(i, 2).Value Then
                開始時間数 = 参照シート.Cells(i, 3).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 2).Value Then
                終了時間数 = 参照シート.Cells(i, 3).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
        '----------------'
        Case 勤務設定_日勤早番
        
        For i = 3 To 63
            If 開始時刻数値 = 参照シート.Cells(i, 4).Value Then
                開始時間数 = 参照シート.Cells(i, 5).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 4).Value Then
                終了時間数 = 参照シート.Cells(i, 5).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
        '----------------'
        Case 勤務設定_スライド
        
        For i = 3 To 43
            If 開始時刻数値 = 参照シート.Cells(i, 6).Value Then
                開始時間数 = 参照シート.Cells(i, 7).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 6).Value Then
                終了時間数 = 参照シート.Cells(i, 7).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
        '-----------------'
        Case 勤務設定_本社10時
        
        For i = 3 To 49
            If 開始時刻数値 = 参照シート.Cells(i, 8).Value Then
                開始時間数 = 参照シート.Cells(i, 9).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 49
            If 終了時刻数値 = 参照シート.Cells(i, 8).Value Then
                終了時間数 = 参照シート.Cells(i, 9).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
    End Select
        
End Sub

Private Sub cmb勤務設定_Change()
    
    Dim 参照シート As Worksheet
    Set 参照シート = ThisWorkbook.Worksheets("出勤時間設定")
    
    If cmb開始_時.Value = "" Or cmb開始_分.Value = "" Or cmb終了_時.Value = "" Or cmb終了_分.Value = "" Then
        Exit Sub
    End If
    
    Dim 開始時刻, 終了時刻 As String
    開始時刻 = Format(cmb開始_時.Value, "00") & ":" & Format(cmb開始_分.Value, "00")
    終了時刻 = Format(cmb終了_時.Value, "00") & ":" & Format(cmb終了_分.Value, "00")
    If 開始時刻 > 終了時刻 Then
        Exit Sub
    End If
    
    Dim 開始時刻数値, 終了時刻数値, 開始時間数, 終了時間数 As Double
    
    開始時刻数値 = val(cmb開始_時.Value) + val(cmb開始_分.Value / 60)
    終了時刻数値 = val(cmb終了_時.Value) + val(cmb終了_分.Value / 60)
    
    
    Select Case cmb勤務設定.Value
        
        '----------------'
        Case 勤務設定_本社
        
        For i = 3 To 61
            If 開始時刻数値 = 参照シート.Cells(i, 2).Value Then
                開始時間数 = 参照シート.Cells(i, 3).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 2).Value Then
                終了時間数 = 参照シート.Cells(i, 3).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
        '----------------'
        Case 勤務設定_日勤早番
        
        For i = 3 To 63
            If 開始時刻数値 = 参照シート.Cells(i, 4).Value Then
                開始時間数 = 参照シート.Cells(i, 5).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 4).Value Then
                終了時間数 = 参照シート.Cells(i, 5).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
        '----------------'
        Case 勤務設定_スライド
        
        For i = 3 To 43
            If 開始時刻数値 = 参照シート.Cells(i, 6).Value Then
                開始時間数 = 参照シート.Cells(i, 7).Value
                Exit For
            End If
        Next i
        
        For i = 3 To 61
            If 終了時刻数値 = 参照シート.Cells(i, 6).Value Then
                終了時間数 = 参照シート.Cells(i, 7).Value
                Exit For
            End If
        Next i
        
        txt時間数.Value = Format(終了時間数 - 開始時間数, "00.00")
        
    End Select
    
End Sub

Private Sub cmb編集_月_Change()

    cmb編集_日.Clear
    cmb編集_日.AddItem
    Dim 当月日数, 曜日 As String

    当月日数 = Day(DateSerial(cmb編集_年.Value, cmb編集_月.Value + 1, 0))
    For i = 1 To 当月日数
        曜日 = Format(Weekday(DateSerial(cmb編集_年.Value, cmb編集_月.Value, i)), "aaa")
        cmb編集_日.AddItem
        cmb編集_日.List(cmb編集_日.ListCount - 1, 0) = i
        cmb編集_日.List(cmb編集_日.ListCount - 1, 1) = 曜日
        cmb編集_日.List(cmb編集_日.ListCount - 1, 2) = i & "(" & 曜日 & ")"
    Next i

End Sub

Private Sub cmbチケット名_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmbチケット名.Text = "" Then
            Exit Sub
        End If
        If チケット名書式チェック(cmbチケット名.Text) Then
            Exit Sub
        End If
        'キーワード入力でエンターキーが押されたらテキスト検索
        Dim チケット名 As String
        チケット名 = SearchComboboxText(cmbチケット名, 2, cmbチケット名.Text)
        If チケット名 <> "" Then
            cmbチケット名.Value = チケット名
        End If
    End If
End Sub

