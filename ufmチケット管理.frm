VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufmチケット管理 
   Caption         =   "チケット管理"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9795
   OleObjectBlob   =   "ufmチケット管理.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufmチケット管理"
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
    
    'ウィンドウ最小化ボタン有効
    Call FrmDec(Me.Caption, True, True, True)
    
    'リストビュー初期化
    lvチケット.ColumnHeaders.Clear
    lvチケット.ColumnHeaders.Add , "トラッカー", "トラッカー", 50
    lvチケット.ColumnHeaders.Add , "優先度", "優先度", 40
    lvチケット.ColumnHeaders.Add , "プロジェクト名", "プロジェクト名", 85
    lvチケット.ColumnHeaders.Add , "チケット番号", "チケット番号", 55
    lvチケット.ColumnHeaders.Add , "チケット名", "チケット名", 180
    lvチケット.ColumnHeaders.Add , "ステータス", "ステータス", 45
    lvチケット.ColumnHeaders.Add , "進捗率", "進捗率", 40
    lvチケット.ColumnHeaders.Add , "開始", "開始", 80
    lvチケット.ColumnHeaders.Add , "期日", "期日", 80
    lvチケット.ColumnHeaders.Add , "リリース予定日", "リリース予定日", 80
    lvチケット.ColumnHeaders.Add , "予定工数", "予定工数", 65
    lvチケット.ColumnHeaders.Add , "記録工数", "記録工数", 65
    lvチケット.ColumnHeaders.Add , "無効", "無効", 40
    lvチケット.ColumnHeaders.Add , "チケット管理番号", "チケット管理番号", 0
    lvチケット.ColumnHeaders.Add , "今後の作業", "今後の作業", 0
    lvチケット.ColumnHeaders.Add , "備考", "備考"
    
    '検索年月リスト初期化
    cmb検索条件_年.Clear
    cmb検索条件_年.AddItem
    For i = Year(Date) - 2 To Year(Date) + 1
        cmb検索条件_年.AddItem i
    Next i
    
    '検索範囲:当月から3ヶ月前を検索開始月とする
    If Month(Date) < 4 Then
        cmb検索条件_年.Value = Year(Date) - 1
    Else
        cmb検索条件_年.Value = Year(Date)
    End If
    
    cmb検索条件_月.Clear
    cmb検索条件_月.AddItem
    For i = 1 To 12
        cmb検索条件_月.AddItem i
    Next i
    cmb検索条件_月.Value = Month(DateAdd("m", -3, Date))
    
    
    '年リスト初期化
    cmb開始_年.Clear
    cmb期日_年.Clear
    cmbリリース予定日_年.Clear
    cmb開始_年.AddItem
    cmb期日_年.AddItem
    cmbリリース予定日_年.AddItem
    For i = Year(Date) - 2 To Year(Date) + 1
        cmb開始_年.AddItem i
        cmb期日_年.AddItem i
        cmbリリース予定日_年.AddItem i
    Next i
    cmb開始_年.Value = Year(Date)
    cmb期日_年.Value = Year(Date)
    cmbリリース予定日_年.Value = ""

    
    '月リスト初期化
    cmb開始_月.Clear
    cmb期日_月.Clear
    cmbリリース予定日_月.Clear
    cmb開始_月.AddItem
    cmb期日_月.AddItem
    cmbリリース予定日_月.AddItem
    For i = 1 To 12
        cmb開始_月.AddItem i
        cmb期日_月.AddItem i
        cmbリリース予定日_月.AddItem i
    Next i
    cmb開始_月.Value = Month(Date)
    cmb期日_月.Value = Month(Date)
    cmbリリース予定日_月.Value = ""
    
    '日リスト初期化
    cmb開始_日.Clear
    cmb期日_日.Clear
    cmbリリース予定日_日.Clear
    cmb開始_日.AddItem
    cmb期日_日.AddItem
    cmbリリース予定日_日.AddItem
    For i = 1 To 31
        cmb開始_日.AddItem i
        cmb期日_日.AddItem i
        cmbリリース予定日_日.AddItem i
    Next i
    cmb開始_日.Value = Day(Date)
    cmb期日_日.Value = Day(Date)
    cmbリリース予定日_日.Value = ""
    
    '進捗率リスト初期化
    cmb進捗率.Clear
    For i = 0 To 1 Step 0.1
        cmb進捗率.AddItem
        cmb進捗率.List(cmb進捗率.ListCount - 1, 0) = i
        cmb進捗率.List(cmb進捗率.ListCount - 1, 1) = FormatPercent(i, 0, vbTrue, vbFalse, vbFalse)
    Next i
    cmb進捗率.Value = cmb進捗率.List(0, 0)
    
    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, True)
    On Error GoTo ErrRSOpen
    
    'トラッカーリスト初期化
    sql = "SELECT 項目名,値" _
        & " FROM V_トラッカー"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    Do Until oRS.EOF
        cmbトラッカー.AddItem
        cmbトラッカー.List(cmbトラッカー.ListCount - 1, 0) = oRS.Fields("値")
        cmbトラッカー.List(cmbトラッカー.ListCount - 1, 1) = oRS.Fields("項目名")
        oRS.MoveNext
    Loop
    cmbトラッカー.Value = cmbトラッカー.List(0, 0)
        
    '優先度リスト初期化
    sql = "SELECT 項目名,値" _
        & " FROM V_優先度"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    Do Until oRS.EOF
        cmb優先度.AddItem
        cmb優先度.List(cmb優先度.ListCount - 1, 0) = oRS.Fields("値")
        cmb優先度.List(cmb優先度.ListCount - 1, 1) = oRS.Fields("項目名")
        oRS.MoveNext
    Loop
    cmb優先度.Value = cmb優先度.List(1, 0)
    
    'ステータスリスト初期化
    sql = "SELECT 項目名,値" _
        & " FROM V_ステータス"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    Do Until oRS.EOF
        cmbステータス.AddItem
        cmbステータス.List(cmbステータス.ListCount - 1, 0) = oRS.Fields("値")
        cmbステータス.List(cmbステータス.ListCount - 1, 1) = oRS.Fields("項目名")
        oRS.MoveNext
    Loop
    cmbステータス.Value = cmbステータス.List(0, 0)
    
    
    
    'プロジェクトリスト初期化
    sql = "SELECT プロジェクト名, プロジェクト番号,項目名" _
        & " FROM プロジェクト管理 LEFT JOIN V_部門コード ON V_部門コード.値 = プロジェクト管理.部門コード"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    cmbプロジェクト.Clear
    cmb検索条件_プロジェクト.Clear
    cmbプロジェクト.AddItem
    cmb検索条件_プロジェクト.AddItem
    Do Until oRS.EOF
        cmbプロジェクト.AddItem
        cmb検索条件_プロジェクト.AddItem
        cmbプロジェクト.List(cmbプロジェクト.ListCount - 1, 0) = oRS.Fields("プロジェクト番号").Value
        cmbプロジェクト.List(cmbプロジェクト.ListCount - 1, 1) = oRS.Fields("プロジェクト名").Value
        cmbプロジェクト.List(cmbプロジェクト.ListCount - 1, 2) = oRS.Fields("項目名").Value
        cmb検索条件_プロジェクト.List(cmbプロジェクト.ListCount - 1, 0) = oRS.Fields("プロジェクト番号").Value
        cmb検索条件_プロジェクト.List(cmbプロジェクト.ListCount - 1, 1) = oRS.Fields("プロジェクト名").Value
        cmb検索条件_プロジェクト.List(cmbプロジェクト.ListCount - 1, 2) = oRS.Fields("項目名").Value
        oRS.MoveNext
    Loop
    cmbプロジェクト.Value = cmbプロジェクト.List(1, 0)
    
    btn追加.Enabled = True
    btn更新.Enabled = False
    
    btn検索_Click
    
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

Private Sub UserForm_Resize()
    
    Dim 幅 As Integer, 高 As Integer
    幅 = Me.InsideWidth - lvチケット.Left * 2
    高 = Me.InsideHeight - fm検索条件.Height - fm編集.Height - 25
    If 幅 >= fm検索条件.Width + lbl件数.Width Then
        lbl件数.Left = 幅 - lbl件数.Width
        lvチケット.Width = 幅
        For j = 1 To lvチケット.ColumnHeaders.Count - 1
            If 幅 - 5 <= lvチケット.ColumnHeaders.Item(j).Width Then
                Exit For
            End If
            幅 = 幅 - lvチケット.ColumnHeaders.Item(j).Width
        Next
        If lvチケット.ColumnHeaders.Count >= 1 Then
            lvチケット.ColumnHeaders.Item(lvチケット.ColumnHeaders.Count).Width = 幅 - 5
        End If
    End If
    If 高 >= 100 Then
        lvチケット.Height = 高
        fm編集.Top = fm検索条件.Height + lvチケット.Height + 25
    End If
    
End Sub

Private Sub btn全選択_Click()

    Dim i, j As Long
    j = 0
    If lvチケット.ListItems.Count = 0 Then
        MsgBox "データなし、検索してください。", vbInformation
        Exit Sub
    End If
    For i = 1 To lvチケット.ListItems.Count
         If lvチケット.ListItems.Item(i).Checked = False Then
            Exit For
        End If
        j = j + 1
    Next i
    If j = lvチケット.ListItems.Count Then
        For i = 1 To lvチケット.ListItems.Count
            lvチケット.ListItems.Item(i).Checked = False
        Next i
    Else
        For i = 1 To lvチケット.ListItems.Count
             lvチケット.ListItems.Item(i).Checked = True
        Next i
    End If

End Sub

Private Function フィルタ条件() As String

    'フィルタ条件生成
    Dim 開始年月日 As String
    If cmb検索条件_年.Value = "" Or cmb検索条件_月.Value = "" Then
        MsgBox "検索範囲を入力してください。", vbExclamation
    End If
    開始年月日 = cmb検索条件_年.Value & "/" & Format(cmb検索条件_月.Value, "00") & "/" & "01"
    If 開始年月日 <> "" Then
        If フィルタ条件 <> "" Then
            フィルタ条件 = フィルタ条件 & " And "
        End If
        フィルタ条件 = フィルタ条件 & "[チケット管理.開始] >=#" & 開始年月日 & "#"
    End If

    Dim プロジェクト As String
    If cmb検索条件_プロジェクト.Text = "" Then
        プロジェクト = ""
    ElseIf IsNull(cmb検索条件_プロジェクト.Value) Then
        If Not プロジェクト書式チェック(cmb検索条件_プロジェクト.Text) Then
            MsgBox "プロジェクト入力が不正です。", vbExclamation
            cmb検索条件_プロジェクト.SetFocus
            Exit Function
        Else
            プロジェクト = cmb検索条件_プロジェクト.Text
        End If
    Else
        プロジェクト = cmb検索条件_プロジェクト.Value
    End If

    If プロジェクト <> "" Then
        If フィルタ条件 <> "" Then
            フィルタ条件 = フィルタ条件 & " And "
        End If
        フィルタ条件 = フィルタ条件 & "[チケット管理.プロジェクト番号] ='" & プロジェクト & "'"
    End If

    Dim チケット番号 As String
    If txt検索条件_チケット番号.Value <> "" Then
        チケット番号 = StrConv(txt検索条件_チケット番号, vbNarrow)
        If Left(チケット番号, 1) <> "#" Then
            チケット番号 = "#" & チケット番号
        End If
        If チケット名書式チェック(チケット番号) = False Then
            MsgBox "チケット番号が不正です。", vbExclamation
            txtチケット番号.SetFocus
            Exit Function
        End If
    End If
    If チケット番号 <> "" Then
            If フィルタ条件 <> "" Then
            フィルタ条件 = フィルタ条件 & " And "
        End If
        フィルタ条件 = フィルタ条件 & "[チケット管理.チケット番号] ='" & チケット番号 & "'"
    End If
    
    If Not cb検索条件_無効.Value Then
        If フィルタ条件 <> "" Then
            フィルタ条件 = フィルタ条件 & " And "
        End If
        フィルタ条件 = フィルタ条件 & "[チケット管理.削除フラグ] = False"
    End If
        
    Dim チェックフィルタ As String
    If cbトラッカー_1.Value Or _
        cbトラッカー_2.Value Or _
        cbトラッカー_3.Value Or _
        cbトラッカー_4.Value Or _
        cbトラッカー_5.Value Then

        If フィルタ条件 <> "" Then
            フィルタ条件 = フィルタ条件 & " And "
        End If
        チェックフィルタ = ""
        If cbトラッカー_1.Value Then
            チェックフィルタ = チェックフィルタ & "[チケット管理.トラッカー] = 1"
        End If
        If cbトラッカー_2.Value Then
            If チェックフィルタ <> "" Then
                チェックフィルタ = チェックフィルタ & " Or "
            End If
            チェックフィルタ = チェックフィルタ & "[チケット管理.トラッカー] = 2"
        End If
        If cbトラッカー_3.Value Then
            If チェックフィルタ <> "" Then
                チェックフィルタ = チェックフィルタ & " Or "
            End If
            チェックフィルタ = チェックフィルタ & "[チケット管理.トラッカー] = 3"
        End If
        If cbトラッカー_4.Value Then
            If チェックフィルタ <> "" Then
                チェックフィルタ = チェックフィルタ & " Or "
            End If
            チェックフィルタ = チェックフィルタ & "[チケット管理.トラッカー] = 4"
        End If
        If cbトラッカー_5.Value Then
            If チェックフィルタ <> "" Then
                チェックフィルタ = チェックフィルタ & " Or "
            End If
            チェックフィルタ = チェックフィルタ & "[チケット管理.トラッカー] = 5"
        End If
        フィルタ条件 = フィルタ条件 & " ( " & チェックフィルタ & " ) "
    End If
    
    If cbステータス_1.Value Or _
        cbステータス_2.Value Or _
        cbステータス_3.Value Or _
        cbステータス_4.Value Then

        If フィルタ条件 <> "" Then
            フィルタ条件 = フィルタ条件 & " And "
        End If
        チェックフィルタ = ""
        If cbステータス_1.Value Then
            チェックフィルタ = チェックフィルタ & "[チケット管理.ステータス] = 1"
        End If
        If cbステータス_2.Value Then
            If チェックフィルタ <> "" Then
                チェックフィルタ = チェックフィルタ & " Or "
            End If
            チェックフィルタ = チェックフィルタ & "[チケット管理.ステータス] = 2"
        End If
        If cbステータス_3.Value Then
            If チェックフィルタ <> "" Then
                チェックフィルタ = チェックフィルタ & " Or "
            End If
            チェックフィルタ = チェックフィルタ & "[チケット管理.ステータス] = 3"
        End If
        If cbステータス_4.Value Then
            If チェックフィルタ <> "" Then
                チェックフィルタ = チェックフィルタ & " Or "
            End If
            チェックフィルタ = チェックフィルタ & "[チケット管理.ステータス] = 4"
        End If
        フィルタ条件 = フィルタ条件 & " ( " & チェックフィルタ & " ) "
    End If
    
    If フィルタ条件 <> "" Then
        フィルタ条件 = " WHERE " & フィルタ条件
    End If
 
End Function

Private Sub btn検索_Click()

    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, True)

    'レコードセット取得
    On Error GoTo ErrRSOpen
    sql = "SELECT *" _
        & " FROM (((チケット管理" _
        & " LEFT JOIN プロジェクト管理 ON プロジェクト管理.プロジェクト番号 = チケット管理.プロジェクト番号)" _
        & " LEFT JOIN V_ステータス ON V_ステータス.値 = CStr(チケット管理.ステータス))" _
        & " LEFT JOIN V_トラッカー ON V_トラッカー.値 = CStr(チケット管理.トラッカー))" _
        & " LEFT JOIN V_優先度 ON V_優先度.値 = CStr(チケット管理.優先度)"
        
    Dim sort As String
    sort = " ORDER BY チケット管理.ステータス,チケット番号"
    
    If フィルタ条件 <> "" Then
        sql = sql & フィルタ条件
    End If
    
    sql = sql & sort
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)

    lvチケット.ListItems.Clear
    lvチケット.Sorted = False
    Do Until oRS.EOF
        With lvチケット.ListItems.Add
            .Text = oRS.Fields("V_トラッカー.項目名").Value
            .SubItems(1) = oRS.Fields("V_優先度.項目名").Value
            .SubItems(2) = Null2Blank(oRS.Fields("プロジェクト名").Value)
            .SubItems(3) = Null2Blank(oRS.Fields("チケット番号").Value)
            .SubItems(4) = Null2Blank(oRS.Fields("チケット名").Value)
            .SubItems(5) = oRS.Fields("V_ステータス.項目名").Value
            .SubItems(6) = FormatPercent(oRS.Fields("進捗率").Value, 0, vbTrue, vbFalse, vbFalse)
            .SubItems(7) = oRS.Fields("開始").Value & "(" & Format(Weekday(oRS.Fields("開始").Value), "aaa") & ")"
            .SubItems(8) = oRS.Fields("期日").Value & "(" & Format(Weekday(oRS.Fields("期日").Value), "aaa") & ")"
            If IsNull(oRS.Fields("リリース予定日").Value) Then
                .SubItems(9) = ""
            Else
                .SubItems(9) = oRS.Fields("リリース予定日").Value & "(" & Format(Weekday(oRS.Fields("リリース予定日").Value), "aaa") & ")"
            End If
            If IsNull(oRS.Fields("予定工数").Value) Then
                .SubItems(10) = ""
            Else
                .SubItems(10) = Format(oRS.Fields("予定工数").Value, "00.00") & "H"
            End If
            If IsNull(oRS.Fields("記録工数").Value) Then
                .SubItems(11) = ""
            Else
                .SubItems(11) = Format(oRS.Fields("記録工数").Value, "00.00") & "H"
            End If
            If oRS.Fields("チケット管理.削除フラグ").Value = True Then
                .SubItems(12) = "●"
            End If
            .SubItems(13) = Null2Blank(oRS.Fields("チケット管理番号").Value)
            .SubItems(14) = Null2Blank(oRS.Fields("今後の作業").Value)
            .SubItems(15) = Null2Blank(oRS.Fields("備考").Value)
        End With
        oRS.MoveNext
    Loop

    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0

    lbl件数.Caption = lvチケット.ListItems.Count & " 件"
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
    
    'チケット番号入力チェック
    Dim チケット番号 As String
    If txtチケット番号.Value <> "" Then
        チケット番号 = StrConv(txtチケット番号, vbNarrow)
        If Left(チケット番号, 1) <> "#" Then
            チケット番号 = "#" & チケット番号
        End If
        If チケット名書式チェック(チケット番号) = False Then
            MsgBox "チケット番号が不正です。", vbExclamation
            txtチケット番号.SetFocus
            Exit Sub
        End If
    Else
        チケット番号 = ""
    End If
    
    'チケット名入力チェック
    Dim チケット名 As String
    If txtチケット名.Value <> "" Then
        チケット名 = txtチケット名.Value
    Else
        MsgBox "チケット名が未入力。", vbExclamation
    End If
    
    '日付入力チェック
    Dim 開始 As String
    If cmb開始_年.Value = "" Or cmb開始_月 = "" Or cmb開始_日 = "" Then
        MsgBox "開始日付を入力してください。", vbExclamation
    End If
    開始 = cmb開始_年.Value & "/" & Format(cmb開始_月.Value, "00") & "/" & Format(cmb開始_日, "00")
    
    Dim 期日 As String
    If cmb期日_年.Value = "" Or cmb期日_月 = "" Or cmb期日_日 = "" Then
        MsgBox "期日日付を入力してください。", vbExclamation
    End If
    期日 = cmb期日_年.Value & "/" & Format(cmb期日_月.Value, "00") & "/" & Format(cmb期日_日, "00")
    
    Dim リリース予定日 As String
    If cmbリリース予定日_年.Value = "" Or cmbリリース予定日_月 = "" Or cmbリリース予定日_日 = "" Then
        リリース予定日 = ""
    Else
        リリース予定日 = cmbリリース予定日_年.Value & "/" & Format(cmbリリース予定日_月.Value, "00") & "/" & Format(cmbリリース予定日_日, "00")
    End If
    
    Dim 予定工数 As Double
    If txt予定工数.Value <> "" Then
        予定工数 = txt予定工数.Value
    End If
    
    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, False)
    On Error GoTo ErrRSOpen
    
    Set oRS = oDB.OpenRecordset("チケット管理", dbOpenTable)
    'トランザクション開始
    oWks.BeginTrans
    
    '新規レコード追加
    oRS.AddNew
    oRS.Fields("プロジェクト番号").Value = Null2Blank(プロジェクト)
    oRS.Fields("トラッカー").Value = cmbトラッカー.Value
    oRS.Fields("ステータス").Value = cmbステータス.Value
    oRS.Fields("チケット番号").Value = Null2Blank(チケット番号)
    oRS.Fields("チケット名").Value = Null2Blank(txtチケット名.Value)
    oRS.Fields("優先度").Value = cmb優先度.Value
    oRS.Fields("進捗率").Value = cmb進捗率.Value
    oRS.Fields("開始").Value = 開始
    oRS.Fields("期日").Value = 期日
    oRS.Fields("リリース予定日").Value = DBDateInput(リリース予定日)
    oRS.Fields("予定工数").Value = Null2Blank(予定工数)
    txtチケット管理番号.Value = "C" & Format(Now(), "yyyymmdd") & "-" & Right(Format(oRS.Fields("ID").Value, "0000"), 4)
    oRS.Fields("チケット管理番号").Value = txtチケット管理番号.Value
    oRS.Fields("今後の作業").Value = Null2Blank(チケット番号) & "―" & Trim(Null2Blank(txtチケット名.Value)) & "[" & Trim(開始) & "～" & Trim(期日) & "]"
    oRS.Fields("備考").Value = Null2Blank(txt備考.Value)
    oRS.Update
    
    'トランザクション完了
    oWks.CommitTrans
    
    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0
    
    '操作メニュー表示
    MsgBox "データの登録に成功しました。", vbInformation
    btn追加.Enabled = False
    btn更新.Enabled = True
    btn検索_Click
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

Private Sub lvチケット_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvチケット
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvチケット_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim チケット管理番号 As String
    チケット管理番号 = Item.ListSubItems(13).Text
    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, True)
    On Error GoTo ErrRSOpen
    
    'レコードセット取得
    Set oRS = oDB.OpenRecordset("チケット管理", dbOpenTable)
    oRS.Index = "チケット検索"
    oRS.Seek "=", チケット管理番号
    If oRS.NoMatch Then
        GoTo ErrDataInvalid
    End If
    
    '編集部へ出力
    txtチケット管理番号.Value = oRS.Fields("チケット管理番号").Value
    cmbプロジェクト.Value = oRS.Fields("プロジェクト番号").Value
    cmbトラッカー.Value = oRS.Fields("トラッカー").Value
    cmbステータス.Value = oRS.Fields("ステータス").Value
    txtチケット番号.Value = oRS.Fields("チケット番号").Value
    txtチケット名.Value = oRS.Fields("チケット名").Value
    cmb優先度.Value = oRS.Fields("優先度").Value
    cmb進捗率.Value = oRS.Fields("進捗率").Value
    txt予定工数.Value = oRS.Fields("予定工数").Value
    cmb開始_年.Value = Year(oRS.Fields("開始").Value)
    cmb開始_月.Value = Month(oRS.Fields("開始").Value)
    cmb開始_日.Value = Day(oRS.Fields("開始").Value)
    cmb期日_年.Value = Year(oRS.Fields("期日").Value)
    cmb期日_月.Value = Month(oRS.Fields("期日").Value)
    cmb期日_日.Value = Day(oRS.Fields("期日").Value)
    '----------------------------------------------'
    cmb開始_年.Value = Year(oRS.Fields("開始").Value)
    cmb開始_月.Value = Month(oRS.Fields("開始").Value)
    cmb開始_日.Value = Day(oRS.Fields("開始").Value)
    cmb期日_年.Value = Year(oRS.Fields("期日").Value)
    cmb期日_月.Value = Month(oRS.Fields("期日").Value)
    cmb期日_日.Value = Day(oRS.Fields("期日").Value)
    '----------------------------------------------'
    cmbリリース予定日_年.Value = Year(oRS.Fields("リリース予定日").Value)
    cmbリリース予定日_月.Value = Month(oRS.Fields("リリース予定日").Value)
    cmbリリース予定日_日.Value = Day(oRS.Fields("リリース予定日").Value)
    cb無効.Value = oRS.Fields("削除フラグ").Value
    txt備考.Value = oRS.Fields("備考").Value
    
    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0
    
    '更新ボタン有効化
    btn更新.Enabled = True
    
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

Private Sub btnクリア_Click()

    '編集部を初期化
    cmbプロジェクト.Value = cmbプロジェクト.List(1, 0)
    cmbトラッカー.Value = cmbトラッカー.List(0, 0)
    cmbステータス.Value = cmbステータス.List(0, 0)
    cmb進捗率.Value = cmb進捗率.List(0, 0)
    txtチケット番号.Value = ""
    txtチケット名.Value = ""
    txt予定工数.Value = ""
    cmb開始_年.Value = Year(Date)
    cmb開始_月.Value = Month(Date)
    cmb開始_日.Value = Day(Date)
    cmb期日_年.Value = Year(Date)
    cmb期日_月.Value = Month(Date)
    cmb期日_日.Value = Day(Date)
    cmbリリース予定日_年.Value = ""
    cmbリリース予定日_月.Value = ""
    cmbリリース予定日_日.Value = ""
    txt備考.Value = ""
    
    '更新ボタン無効化
    btn更新.Enabled = False
    btn追加.Enabled = True
    
End Sub

Private Sub btn更新_Click()

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
    
    'チケット番号入力チェック
    Dim チケット番号 As String
    If txtチケット番号.Value <> "" Then
        チケット番号 = StrConv(txtチケット番号, vbNarrow)
        If Left(チケット番号, 1) <> "#" Then
            チケット番号 = "#" & チケット番号
        End If
        If チケット名書式チェック(チケット番号) = False Then
            MsgBox "チケット番号が不正です。", vbExclamation
            txtチケット番号.SetFocus
            Exit Sub
        End If
    Else
        チケット番号 = ""
    End If
    
    '日付入力チェック
    Dim 開始 As String
    If cmb開始_年.Value = "" Or cmb開始_月 = "" Or cmb開始_日 = "" Then
        MsgBox "開始日付を入力してください。", vbExclamation
    End If
    開始 = cmb開始_年.Value & "/" & Format(cmb開始_月.Value, "00") & "/" & Format(cmb開始_日, "00")
    
    Dim 期日 As String
    If cmb期日_年.Value = "" Or cmb期日_月 = "" Or cmb期日_日 = "" Then
        MsgBox "期日日付を入力してください。", vbExclamation
    End If
    期日 = cmb期日_年.Value & "/" & Format(cmb期日_月.Value, "00") & "/" & Format(cmb期日_日, "00")
    
    Dim リリース予定日 As String
    If cmbリリース予定日_年.Value = "" Or cmbリリース予定日_月 = "" Or cmbリリース予定日_日 = "" Then
        リリース予定日 = ""
    Else
        リリース予定日 = cmbリリース予定日_年.Value & "/" & Format(cmbリリース予定日_月.Value, "00") & "/" & Format(cmbリリース予定日_日, "00")
    End If
    
    Dim 予定工数 As Double
    If txt予定工数.Value <> "" Then
        予定工数 = txt予定工数.Value
    End If
    
    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, False)
    On Error GoTo ErrRSOpen
    
    'トランザクション開始
    oWks.BeginTrans
    
    'チケット管理番号存在チェック
    Set oRS = oDB.OpenRecordset("チケット管理", dbOpenTable)
    oRS.Index = "チケット検索"
    oRS.Seek "=", txtチケット管理番号.Value
    If oRS.NoMatch Then
        GoTo ErrDataInvalid
    End If
    
    'レコード更新
    oRS.Edit
    oRS.Fields("プロジェクト番号").Value = Null2Blank(プロジェクト)
    oRS.Fields("トラッカー").Value = cmbトラッカー.Value
    oRS.Fields("ステータス").Value = cmbステータス.Value
    oRS.Fields("チケット番号").Value = Null2Blank(チケット番号)
    oRS.Fields("チケット名").Value = Null2Blank(txtチケット名.Value)
    oRS.Fields("優先度").Value = cmb優先度.Value
    oRS.Fields("進捗率").Value = cmb進捗率.Value
    oRS.Fields("開始").Value = 開始
    oRS.Fields("期日").Value = 期日
    oRS.Fields("リリース予定日").Value = DBDateInput(リリース予定日)
    oRS.Fields("予定工数").Value = Null2Blank(予定工数)
    oRS.Fields("削除フラグ").Value = cb無効.Value
    oRS.Fields("今後の作業").Value = Null2Blank(チケット番号) & "―" & Trim(Null2Blank(txtチケット名.Value)) & "[" & Trim(開始) & "～" & Trim(期日) & "]"
    oRS.Fields("備考").Value = Null2Blank(txt備考.Value)
    oRS.Update
    
    'トランザクション完了
    oWks.CommitTrans
    
    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0
    
    '操作メニュー表示
    MsgBox "データの更新に成功しました。", vbInformation
    btn検索_Click
    Exit Sub
     
ErrDataInvalid:
    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    
    MsgBox "チケットが未登録です。", vbExclamation
    Exit Sub
    
ErrRSOpen:
    'データベース切断
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    
    MsgBox "データの更新に失敗しました。再度実行してください。(" & Err.Number & ")", vbExclamation
    Exit Sub
    
ErrDBOpen:
    MsgBox "データベースの接続に失敗しました。(" & Err.Number & ")", vbCritical

End Sub
