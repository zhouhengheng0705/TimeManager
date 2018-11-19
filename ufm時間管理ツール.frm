VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufm時間管理ツール 
   Caption         =   "時間管理ツール"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690
   OleObjectBlob   =   "ufm時間管理ツール.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufm時間管理ツール"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private 時間管理一覧配置横Dic As Object
Const 時間管理出力行 = 4
Const 時間管理出力列 = 2
Const 時間管理ヘッダ行 = 3
Const 時間管理一覧読出行数 = 50

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
    
    'リストビュー初期化
    lv検索一覧.ColumnHeaders.Clear
    lv検索一覧.ColumnHeaders.Add , "記録日付", "記録日付", 78
    lv検索一覧.ColumnHeaders.Add , "開始時間", "開始時間", 48
    lv検索一覧.ColumnHeaders.Add , "終了時間", "終了時間", 48
    lv検索一覧.ColumnHeaders.Add , "時間数", "時間数", 38
    lv検索一覧.ColumnHeaders.Add , "プロジェクト名", "プロジェクト名", 90
    lv検索一覧.ColumnHeaders.Add , "チケット番号", "チケット番号", 55
    lv検索一覧.ColumnHeaders.Add , "チケット名", "チケット名", 240
    lv検索一覧.ColumnHeaders.Add , "コメント", "コメント", 120
    lv検索一覧.ColumnHeaders.Add , "勤務設定", "勤務設定", 70
    lv検索一覧.ColumnHeaders.Add , "日報貼付", "日報貼付", 0
    lv検索一覧.ColumnHeaders.Add , "記録番号", "記録番号", 0
    lv検索一覧.ColumnHeaders.Add , "無効", "無効", 40
    lv検索一覧.ColumnHeaders.Add , "備考", "備考"
    
    
    '検収予定年月日リスト初期化
    cmb検索条件_開始_年.Clear
    cmb検索条件_終了_年.Clear
    cmb検索条件_開始_年.AddItem
    cmb検索条件_終了_年.AddItem
    For i = Year(Date) - 2 To Year(Date) + 1
        cmb検索条件_開始_年.AddItem i
        cmb検索条件_終了_年.AddItem i
    Next i
    cmb検索条件_開始_年.Value = Year(DateAdd("d", -7, Date)) 'Year(DateAdd("m", -1, Date))
    cmb検索条件_終了_年.Value = Year(Date)

    cmb検索条件_開始_月.Clear
    cmb検索条件_終了_月.Clear
    cmb検索条件_開始_月.AddItem
    cmb検索条件_終了_月.AddItem
    For i = 1 To 12
        cmb検索条件_開始_月.AddItem i
        cmb検索条件_終了_月.AddItem i
    Next i
    cmb検索条件_開始_月.Value = Month(DateAdd("d", -7, Date)) 'Month(DateAdd("m", -1, Date))
    cmb検索条件_終了_月.Value = Month(Date)
    
    cmb検索条件_開始_日.Clear
    cmb検索条件_終了_日.Clear
    cmb検索条件_開始_日.AddItem
    cmb検索条件_終了_日.AddItem
    For i = 1 To 31
        cmb検索条件_開始_日.AddItem i
        cmb検索条件_終了_日.AddItem i
    Next i
    cmb検索条件_開始_日.Value = Day(DateAdd("d", -7, Date))
    cmb検索条件_終了_日.Value = Day(Date)
    
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, True)
    On Error GoTo ErrRSOpen
    
    'プロジェクトリスト初期化
    sql = "SELECT プロジェクト名, プロジェクト番号,項目名" _
        & " FROM プロジェクト管理 LEFT JOIN V_部門コード ON V_部門コード.値 = プロジェクト管理.部門コード"
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    cmb検索条件_プロジェクト.Clear
    cmb検索条件_プロジェクト.AddItem
    Do Until oRS.EOF
        cmb検索条件_プロジェクト.AddItem
        cmb検索条件_プロジェクト.List(cmb検索条件_プロジェクト.ListCount - 1, 0) = oRS.Fields("プロジェクト番号").Value
        cmb検索条件_プロジェクト.List(cmb検索条件_プロジェクト.ListCount - 1, 1) = oRS.Fields("プロジェクト名").Value
        cmb検索条件_プロジェクト.List(cmb検索条件_プロジェクト.ListCount - 1, 2) = oRS.Fields("項目名").Value
        oRS.MoveNext
    Loop

    'ウィンドウ最小化ボタン有効
    Call FrmDec(Me.Caption, True, True, True)

    'ウィンドウリサイズ
    UserForm_Resize
    
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
    幅 = Me.InsideWidth - lv検索一覧.Left * 2
    高 = Me.InsideHeight - fm検索条件.Height - 25
    If 幅 >= fm検索条件.Width + lbl件数.Width Then
        lbl件数.Left = 幅 - lbl件数.Width
        lv検索一覧.Width = 幅
        For j = 1 To lv検索一覧.ColumnHeaders.Count - 1
            If 幅 - 5 <= lv検索一覧.ColumnHeaders.Item(j).Width Then
                Exit For
            End If
            幅 = 幅 - lv検索一覧.ColumnHeaders.Item(j).Width
        Next
        If lv検索一覧.ColumnHeaders.Count >= 1 Then
            lv検索一覧.ColumnHeaders.Item(lv検索一覧.ColumnHeaders.Count).Width = 幅 - 5
        End If
    End If
    
    If 高 >= 100 Then
        lv検索一覧.Height = 高
    End If
    
End Sub

Private Function フィルタ条件() As String

    'フィルタ条件生成
    Dim 開始年月日, 終了年月日 As String
    If cmb検索条件_開始_年.Value = "" Or cmb検索条件_開始_月.Value = "" Or cmb検索条件_開始_日.Value = "" Then
        MsgBox "開始年月日を正しく入力してください。", vbExclamation
        Exit Function
    End If
    開始年月日 = cmb検索条件_開始_年.Value & "/" & Format(cmb検索条件_開始_月.Value, "00") & "/" & Format(cmb検索条件_開始_日.Value, "00")
    If 開始年月日 <> "" Then
        If フィルタ条件 <> "" Then
            フィルタ条件 = フィルタ条件 & " And "
        End If
        フィルタ条件 = フィルタ条件 & "[時間管理.記録日付] >=#" & 開始年月日 & "#"
    End If
    
    If cmb検索条件_終了_年.Value = "" Or cmb検索条件_終了_月.Value = "" Or cmb検索条件_終了_日.Value = "" Then
        MsgBox "終了年月日を正しく入力してください。", vbExclamation
        Exit Function
    End If
    終了年月日 = cmb検索条件_終了_年.Value & "/" & Format(cmb検索条件_終了_月.Value, "00") & "/" & Format(cmb検索条件_終了_日.Value, "00")
    If 終了年月日 <> "" Then
        If フィルタ条件 <> "" Then
            フィルタ条件 = フィルタ条件 & " And "
        End If
        フィルタ条件 = フィルタ条件 & "[時間管理.記録日付] <=#" & 終了年月日 & "#"
    End If
    
    If 開始年月日 > 終了年月日 Then
        MsgBox "終了年月日を正しく入力してください。", vbExclamation
        Exit Function
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
        フィルタ条件 = フィルタ条件 & "[時間管理.プロジェクト番号] ='" & プロジェクト & "'"
    End If
    
    Dim チケット番号 As String
    If txt検索条件_チケット番号.Value <> "" Then
        チケット番号 = StrConv(txt検索条件_チケット番号.Value, vbNarrow)
        If Left(チケット番号, 1) <> "#" Then
            チケット番号 = "#" & チケット番号
        End If
        If チケット名書式チェック(チケット番号) = False Then
            MsgBox "チケット番号が不正です。", vbExclamation
            txt検索条件_チケット番号.SetFocus
            Exit Function
        End If
    End If
        
    If チケット番号 <> "" Then
        If フィルタ条件 <> "" Then
            フィルタ条件 = フィルタ条件 & " And "
        End If
        フィルタ条件 = フィルタ条件 & "[時間管理.チケット番号] ='" & チケット番号 & "'"
    End If
    
    If Not cb検索条件_無効.Value Then
        If フィルタ条件 <> "" Then
            フィルタ条件 = フィルタ条件 & " And "
        End If
        フィルタ条件 = フィルタ条件 & "[時間管理.削除フラグ] = False"
    End If

    If フィルタ条件 <> "" Then
        フィルタ条件 = " WHERE " & フィルタ条件
    End If
    
End Function
'
Sub btn検索_Click()

    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, True)

    'レコードセット取得
    On Error GoTo ErrRSOpen
    sql = "SELECT *" _
        & " FROM (((時間管理" _
        & " LEFT JOIN チケット管理 ON チケット管理.チケット番号= 時間管理.チケット番号)" _
        & " LEFT JOIN プロジェクト管理 ON プロジェクト管理.プロジェクト番号 = 時間管理.プロジェクト番号)" _
        & " LEFT JOIN V_勤務設定 ON V_勤務設定.値 = 時間管理.勤務設定)" _
        & " LEFT JOIN V_部門コード ON V_部門コード.値 = プロジェクト管理.部門コード"

    Dim sort As String
    sort = " ORDER BY 時間管理.記録日付 DESC,時間管理.開始時間 "
    sql = sql & フィルタ条件 & sort
    
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    lv検索一覧.ListItems.Clear
    lv検索一覧.Sorted = False
    Do Until oRS.EOF
        With lv検索一覧.ListItems.Add
            .Text = oRS.Fields("記録日付").Value & "(" & Format(Weekday(oRS.Fields("記録日付").Value), "aaa") & ")"
            .SubItems(1) = Format(oRS.Fields("開始時間").Value, "hh:mm:ss")
            .SubItems(2) = Format(oRS.Fields("終了時間").Value, "hh:mm:ss")
            .SubItems(3) = Format(oRS.Fields("時間数").Value, "00.00") & "H"
            If Null2Blank(oRS.Fields("プロジェクト名")) = "" Then
                .SubItems(4) = "---"
            Else
                .SubItems(4) = Null2Blank(oRS.Fields("プロジェクト名"))
            End If
            If Null2Blank(oRS.Fields("チケット管理.チケット番号").Value) = "" Then
                .SubItems(5) = "---"
            Else
                .SubItems(5) = Null2Blank(oRS.Fields("チケット管理.チケット番号").Value)
            End If
            If Null2Blank(oRS.Fields("チケット名").Value) = "" Then
                .SubItems(6) = "---"
            Else
                .SubItems(6) = Null2Blank(oRS.Fields("チケット名").Value)
            End If
            .SubItems(7) = Null2Blank(oRS.Fields("コメント").Value)
            .SubItems(8) = oRS.Fields("V_勤務設定.項目名").Value
            .SubItems(9) = Null2Blank(oRS.Fields("日報貼付").Value)
            .SubItems(10) = oRS.Fields("記録番号").Value
            If oRS.Fields("時間管理.削除フラグ").Value Then
                .SubItems(11) = "●"
            End If
            .SubItems(12) = Null2Blank(oRS.Fields("時間管理.備考").Value)
        End With
        oRS.MoveNext
    Loop

    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0

    lbl件数.Caption = lv検索一覧.ListItems.Count & " 件"
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

Private Sub btn全選択_Click()

    Dim i, j As Long
    j = 0
    If lv検索一覧.ListItems.Count = 0 Then
        MsgBox "データなし、検索してください。", vbInformation
        Exit Sub
    End If
    For i = 1 To lv検索一覧.ListItems.Count
         If lv検索一覧.ListItems.Item(i).Checked = False Then
            Exit For
        End If
        j = j + 1
    Next i
    If j = lv検索一覧.ListItems.Count Then
        For i = 1 To lv検索一覧.ListItems.Count
            lv検索一覧.ListItems.Item(i).Checked = False
        Next i
    Else
        For i = 1 To lv検索一覧.ListItems.Count
             lv検索一覧.ListItems.Item(i).Checked = True
        Next i
    End If

End Sub

Private Sub lv検索一覧_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lv検索一覧
        .Sorted = True
        .SortKey = ColumnHeader.Index - 1
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Function 時間管理一覧配置横Dic生成()

    Dim 最終行 As Long, 最終列 As Long, ヘッダ行Dic As Object, 参照キー As Variant, データ行Dic As Object, 参照シート As Worksheet
    Dim i As Integer, j As Integer
    
    '集計表Dic生成
    Set 時間管理一覧配置横Dic = CreateObject("Scripting.Dictionary")
    Set 参照シート = ThisWorkbook.Worksheets("時間管理一覧配置横")
    最終行 = 参照シート.UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    最終列 = 参照シート.UsedRange.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
    Set ヘッダ行Dic = CreateObject("Scripting.Dictionary")
    For j = 1 To 最終列
        ヘッダ行Dic.Add 参照シート.Cells(1, j).Value, j
    Next j
    For i = 2 To 最終行
        Set データ行Dic = CreateObject("Scripting.Dictionary")
        For Each 参照キー In ヘッダ行Dic
            データ行Dic.Add 参照キー, 参照シート.Cells(i, ヘッダ行Dic(参照キー)).Value
        Next 参照キー
        時間管理一覧配置横Dic.Add 参照シート.Cells(i, 1).Value, データ行Dic
    Next i

End Function

Private Sub btnエクスポート_Click()

'画面描画更新停止(処理高速化)
    Application.ScreenUpdating = False
    
    時間管理一覧配置横Dic生成
    フィルタ条件
    Dim Key As Variant
    Dim sql As String

    sql = "SELECT *," _
        & " IIF(時間管理.削除フラグ = True,'●','') As 無効" _
        & " FROM (((時間管理" _
        & " LEFT JOIN チケット管理 ON チケット管理.チケット番号 = 時間管理.チケット番号)" _
        & " LEFT JOIN プロジェクト管理 ON プロジェクト管理.プロジェクト番号 = 時間管理.プロジェクト番号)" _
        & " LEFT JOIN V_勤務設定 ON V_勤務設定.値 = 時間管理.勤務設定)" _
        & " LEFT JOIN V_部門コード ON V_部門コード.値 = プロジェクト管理.部門コード"
       
    Dim sort As String
    sort = " ORDER BY 時間管理.記録日付 DESC,時間管理.開始時間 "

    If フィルタ条件 <> "" Then
        sql = sql & フィルタ条件 & sort
    Else
        sql = sql & sort
    End If
        
    Dim Sq As String
    Sq = "SELECT "
    For Each Key In 時間管理一覧配置横Dic
        Sq = Sq & 時間管理一覧配置横Dic(Key)("フィールドパス") & ", "
    Next
    
    sql = Left(Sq, Len(Sq) - 2) & " FROM(" & sql & ")"
    
    Dim ヘッダ行, 出力行, i As Long
    出力行 = 時間管理出力行
    ヘッダ行 = 時間管理ヘッダ行
    
    Dim 参照シート, データシート As Worksheet
    Set 参照シート = ThisWorkbook.Worksheets("エクスポート")
    Set データシート = ThisWorkbook.Worksheets("時間管理一覧配置横")
    
    '既存データクリア
    参照シート.Activate
    ActiveWindow.FreezePanes = False
    If 参照シート.AutoFilterMode Then
        参照シート.AutoFilterMode = False
    End If
    ActiveSheet.Cells.Select
    Selection.Clear
    
    '初期書式設定
    With 参照シート
        .Rows(2).RowHeight = 21
        .Range("B2").Value = "時間管理一覧"
        .Range("B2").Font.Size = 14
        .Range("B2").Font.Bold = True
        .Range("B2").HorizontalAlignment = xlLeft
        .Range("B2:C2").Select
        Selection.Merge
        .Range("D2:F2").Select
        Selection.Merge
    End With
    
    'ヘッダ行初期化
    For Each Key In 時間管理一覧配置横Dic
            参照シート.Range(時間管理一覧配置横Dic(Key)("列") & ヘッダ行) = 時間管理一覧配置横Dic(Key)("ヘッダ名")
            参照シート.Range(時間管理一覧配置横Dic(Key)("列") & ヘッダ行).Interior.Color = 時間管理一覧配置横Dic(Key)("ヘッダ色")
            参照シート.Range(時間管理一覧配置横Dic(Key)("列") & ヘッダ行).Borders.LineStyle = xlContinuous
    Next Key

    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, True)
    On Error GoTo ErrRSOpen
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    参照シート.Activate
    Do Until oRS.EOF
        参照シート.Cells(出力行, 時間管理出力列).CopyFromRecordset oRS, 時間管理一覧読出行数
        If oRS.EOF Then
            Exit Do
        End If
        oRS.MoveNext
        出力行 = 出力行 + 時間管理一覧読出行数
    Loop
    参照シート.Range("D2").Value = oRS.RecordCount & "件" & "              " & Now()
    
    'データ行表示形式設定
    最終行 = 参照シート.UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    最終列 = 参照シート.UsedRange.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
    For Each Key In 時間管理一覧配置横Dic
        
        'データ行表示形式設定
        If 時間管理一覧配置横Dic(Key)("表示形式") <> "" Then
            参照シート.Columns(時間管理一覧配置横Dic(Key)("列")).NumberFormatLocal = 時間管理一覧配置横Dic(Key)("表示形式")
        End If
        'データ行配置設定
        If 時間管理一覧配置横Dic(Key)("配置") <> "" Then
            参照シート.Range(Range(時間管理一覧配置横Dic(Key)("列") & 時間管理出力行), Range(時間管理一覧配置横Dic(Key)("列") & 最終行)).HorizontalAlignment = 時間管理一覧配置横Dic(Key)("配置")
        End If
        '列幅自動調整
        参照シート.Columns(時間管理一覧配置横Dic(Key)("列")).AutoFit
        'データ行枠線設定
        参照シート.Range(Range(時間管理一覧配置横Dic(Key)("列") & 時間管理出力行), Range(時間管理一覧配置横Dic(Key)("列") & 最終行)).Borders.LineStyle = xlContinuous
    Next Key
    参照シート.Range("B4").Value = Format(参照シート.Range("B4").Value, "yyyy/mm/dd")
    
    
    With 参照シート
        'フィルター効くにする
        .Range("B3:M" & 最終行).Select
        Selection.AutoFilter
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
        .Rows(時間管理出力行).Select
        ActiveWindow.FreezePanes = True
        .Range("B2").Select
    End With

  'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0
    MsgBox "時間管理のエクスポートが完了しました。", vbInformation
    Exit Sub

ErrRSOpen:
'   データベース切断
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing

    MsgBox "データの読出に失敗しました。再度実行してください。(" & Err.Number & ")", vbExclamation
    Exit Sub

ErrDBOpen:

    MsgBox "データベースの接続に失敗しました。(" & Err.Number & ")", vbCritical

End Sub

Private Sub btn新規_Click()

    ufm編集.Show
    ufm編集.Caption = "新規追加"
 
End Sub

Private Sub btn編集_Click()

    Dim 選択行 As Integer, 記録番号 As String
    For 選択行 = 1 To lv検索一覧.ListItems.Count - 1
        If lv検索一覧.ListItems(選択行).Selected Then
            Exit For
        End If
    Next 選択行
    記録番号 = lv検索一覧.ListItems(選択行).ListSubItems(10).Text

    ufm編集.Show
    ufm編集.Caption = "編集"
    ufm編集.UserForm_Initialize2 記録番号
    
End Sub

Private Sub btn削除_Click()
    
    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, False)
    
    Dim 記録番号 As String
    For i = 1 To lv検索一覧.ListItems.Count
        If lv検索一覧.ListItems.Item(i).Checked = True Then
            記録番号 = lv検索一覧.ListItems(i).ListSubItems(10).Text
            
            'データの完全削除を確認
            Dim 結果 As Boolean
            確認ダイアログ.表示 "このデータをデータベースから削除します。" & vbCrLf & "よろしいですか? ※削除したデータが復元できません。", "データの完全消去に同意"
            結果 = 確認ダイアログ.結果
            If 結果 Then
                Set oRS = oDB.OpenRecordset("時間管理", dbOpenTable)
                oRS.Index = "記録番号"
                oRS.Seek "=", 記録番号
                oRS.Delete
            Else
                Exit Sub
            End If
        End If
    Next i
    
    btn検索_Click
    
    MsgBox "選択したデータを削除しました。", vbInformation
    
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
    Unload Me
    Exit Sub
    
ErrDBOpen:
    MsgBox "データベースの接続に失敗しました。(" & Err.Number & ")", vbCritical
    Unload Me

End Sub

Private Sub btn日報生成_Click()

    ufm日報生成.Show
End Sub

Private Sub btnメール_Click()

    Shell "EXPLORER.EXE https://www.tcomm.jp/webmail/"
    
End Sub
