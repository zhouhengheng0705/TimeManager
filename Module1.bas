Attribute VB_Name = "Module1"
' 消防保守契約管理ツール Macro
' マクロ記録日 : 2017/07/05  ユーザー名 :周恒恒
Public Const システムデータ開始行 = 5
Public Const システムデータ開始列 = 2
Public Const システム開始年 = 2015
Public 設定Dic As Object
Public Const システム種別_バージョン = 1
Public Const システム種別_伝票種別換算補正倍率 = 11
Public Const システム状況_未確定 = 1
Public Const システム状況_新規 = 2
Public Const システム状況_継続 = 3
Public Const システム状況_全更新 = 4
Public Const システム状況_部分更新指令 = 5 '未使用F
Public Const システム状況_部分更新 = 6
Public Const システム状況_解約 = 7

Public Const 勤務設定_本社 = 1
Public Const 勤務設定_日勤早番 = 2
Public Const 勤務設定_日勤遅番 = 3
Public Const 勤務設定_スライド = 4
Public Const 勤務設定_本社10時 = 5

Public Const ステータス_新規 = 1
Public Const ステータス_進行中 = 2
Public Const ステータス_リリース待ち = 3
Public Const ステータス_終了 = 4


Public Const デバッグモード = False
Public Const パスワード = "zhh99072"
Public Const データベース名 = "時間管理.accdb"
Public Const DAOエンジン = "DAO.DBEngine.120"
Public Const DBパスワード = "tcomm"
Public Const 保護パスワード = "tcomm"
'---------------------------------------------------------------------------------------------
'WindowsAPI定義
'ウィンドウ制御用
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, _
        ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Const GWL_STYLE = (-16) 'ウィンドウスタイルを取得
Private Const WS_THICKFRAME = &H40000 'ウィンドウのサイズ変更
Private Const WS_MINIMIZEBOX = &H20000 '最小化ボタン
Private Const WS_MAXIMIZEBOX = &H10000 '最大化ボタン
Private Const LP_CLASSNAME = "ThunderDFrame"

'hWndInsertAfterの設定
Private Const TEMAE_SET = 0 '手前にセット
Private Const USIRO_SET = 1 '後ろにセット
Private Const TUNENI_TEMAE_SET = -1 '常に手前にセット
Private Const KAIJYO = -2 '解除
'wFlagsの設定
Private Const HYOUZI_SURU = &H40 '表示する
Private Const NO_SIZE = &H1 'サイズを設定しない
Private Const NO_MOVE = &H2 '位置を設定しない

'アイコン制御用
Private Declare Function LoadIconBynum Lib "user32" Alias _
        "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" _
        (ByVal hWnd As Long) As Long
Private Declare Function DrawIcon Lib "user32" _
        (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function ReleaseDC Lib "user32" _
        (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" _
        (ByVal hIcon As Long) As Long

Private Const IDI_ASTERISK = 32516&     '1:情報
Private Const IDI_EXCLAMATION = 32515&  '2:注意
Private Const IDI_HAND = 32513&         '3:警告
Private Const IDI_QUESTION = 32514&     '4:問い合わせ
Private Const IDI_APPLICATION = 32512&  '未使用
Private Const IDI_WINLOGO = 32517&      '未使用

'---------------------------------------------------------------------------------------------
'ユーザーフォームウィンドウ制御設定
Sub FrmDec(フォーム名 As String, Optional 最小化 As Boolean = False, Optional 最大化 As Boolean = False, Optional サイズ可変 As Boolean = False)
    Dim fRet As Long
    Dim hWnd As Long
    Dim fStyle As Long

    hWnd = FindWindow(LP_CLASSNAME, フォーム名)
    fStyle = GetWindowLong(hWnd, GWL_STYLE)
    If 最小化 Then
        fStyle = (fStyle Or WS_MINIMIZEBOX)
    End If
    If 最大化 Then
        fStyle = (fStyle Or WS_MAXIMIZEBOX)
    End If
    If サイズ可変 Then
        fStyle = (fStyle Or WS_THICKFRAME)
    End If
    fRet = SetWindowLong(hWnd, GWL_STYLE, fStyle)
    fRet = DrawMenuBar(hWnd)

End Sub

'ユーザーフォームウィンドウ階層設定
Sub FrmPos(フォーム名 As String, Optional 最前面 As Boolean = False)
    Dim hWnd As Long

    hWnd = FindWindow(LP_CLASSNAME, フォーム名)
    If 最前面 Then
        Call SetWindowPos(hWnd, TUNENI_TEMAE_SET, 0, 0, 0, 0, HYOUZI_SURU Or NO_MOVE Or NO_SIZE)
    Else
        Call SetWindowPos(hWnd, KAIJYO, 0, 0, 0, 0, HYOUZI_SURU Or NO_MOVE Or NO_SIZE)
    End If
End Sub

'ユーザーフォームウィンドウ最小化
Sub FrmMin(フォーム名 As String, 最小化 As Boolean)
    Dim hWnd As Long

    hWnd = FindWindow(LP_CLASSNAME, フォーム名)
    If 最小化 Then
        Call CloseWindow(hWnd)
    Else
        Call OpenIcon(hWnd)
    End If
End Sub

'---------------------------------------------------------------------------------------------
Function DB接続(oWks As DAO.Workspace, Optional ReadOnly As Boolean = True) As DAO.Database

    Set DB接続 = oWks.OpenDatabase(ThisWorkbook.Path & "\" & データベース名, False, ReadOnly, ";PWD=" & DBパスワード)

End Function

'---------------------------------------------------------------------------------------------
Sub システム_最新_Click()

    '画面描画更新停止(処理高速化)
    Application.ScreenUpdating = False

    Dim システムシート As Worksheet
    Set システムシート = ThisWorkbook.Worksheets("システム")
    Dim 最終行 As Long
    最終行 = システムシート.UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    If 最終行 >= システムデータ開始行 Then
        システムシート.Rows(システムデータ開始行 & ":" & 最終行).Delete Shift:=xlUp
    End If

    'SQL生成
    Dim sql As String
    sql = "SELECT 項目名, 値" _
        & " FROM M_システム" _
        & " WHERE [種別] = " & システム種別_バージョン

    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, True)

    'レコードセット取得
    On Error GoTo ErrRSOpen

    'システム情報更新
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    システムシート.Cells(システムデータ開始行, システムデータ開始列).CopyFromRecordset oRS

    'バージョンチェック＆アップデート
    Dim 現バージョン As String, 新バージョン As String
    oRS.MoveFirst
    現バージョン = システムシート.Range("C5").Value
    新バージョン = システムシート.Range("C2").Value

    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0

    If 現バージョン <> 新バージョン Then
        If MsgBox("データベースのバージョンが一致しません。アップグレードしますか？" & vbCrLf & "※実行前にデータベースのバックアップをお勧めします。", vbOKCancel + vbInformation) = vbOK Then
            DBバージョン更新 現バージョン
        End If
    End If
    Exit Sub

ErrRSOpen:
    'データベース切断
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing

    MsgBox "データベース情報の読出に失敗しました。再度実行してください。(" & Err.Number & ")", vbExclamation
    Exit Sub

ErrDBOpen:
    MsgBox "データベースの接続に失敗しました。(" & Err.Number & ")", vbCritical

End Sub

Function DBバージョン確認() As Boolean

    Dim システムシート As Worksheet
    Set システムシート = ThisWorkbook.Worksheets("システム")
    Dim 現バージョン As String, 新バージョン As String
    現バージョン = システムシート.Range("C5").Value
    新バージョン = システムシート.Range("C2").Value
    DBバージョン確認 = (現バージョン = 新バージョン)

End Function

Private Sub DBバージョン更新(現バージョン As String)

    Dim システムシート As Worksheet
    Set システムシート = ThisWorkbook.Worksheets("システム")

    'データベース接続
    On Error GoTo ErrDBOpen
    Dim oWks As DAO.Workspace, oDB As DAO.Database, oRS As DAO.Recordset
    Set oWks = CreateObject(DAOエンジン).Workspaces(0)
    Set oDB = DB接続(oWks, False)

    On Error GoTo ErrRSOpen

    'トランザクション開始
    oWks.BeginTrans

    'バージョンに応じて更新処理を記述

    'トランザクション完了
    oWks.CommitTrans

    'システム情報更新
    Dim 最終行 As Long
    最終行 = システムシート.UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    If 最終行 >= システムデータ開始行 Then
        システムシート.Rows(システムデータ開始行 & ":" & 最終行).Delete Shift:=xlUp
    End If
    Dim sql As String
    sql = "SELECT 項目名, 値" _
        & " FROM M_システム" _
        & " WHERE [種別] = " & システム種別_バージョン
    Set oRS = oDB.OpenRecordset(sql, dbOpenDynaset)
    システムシート.Cells(システムデータ開始行, システムデータ開始列).CopyFromRecordset oRS

    'データベース切断
    oRS.Close
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing
    On Error GoTo 0

    MsgBox "データベースのアップグレードが完了しました。", vbInformation
    Exit Sub

ErrRSOpen:
    'データベース切断
    Set oRS = Nothing
    oDB.Close
    Set oDB = Nothing

    MsgBox "データベースのアップグレードに失敗しました。再度実行してください。(" & Err.Number & ")", vbExclamation
    Exit Sub

ErrDBOpen:
    MsgBox "データベースの接続に失敗しました。(" & Err.Number & ")", vbCritical

End Sub

Private Sub DBフィールド追加(oDB As DAO.Database, テーブル名 As String, フィールド名 As String, 型 As Variant, Optional 位置 As Integer = -1)

    Dim tdef As DAO.TableDef
    Dim fld As DAO.Field
    Dim prp As DAO.Property

    'フィールド追加
    Set tdef = oDB.TableDefs(テーブル名)
    tdef.Fields.Append tdef.CreateField(フィールド名, 型)

    '位置変更
    If 位置 = -1 Then
        'デフォルトは末尾へ設定
        位置 = tdef.Fields.Count - 1
    End If
    Set fld = tdef.Fields(フィールド名)
    fld.OrdinalPosition = 位置
    tdef.Fields.Refresh

End Sub

'---------------------------------------------------------------------------------------------
Function ID番号書式チェック(接頭辞 As String, ID番号 As String) As Boolean

    Dim RE, reMatch
    Set RE = CreateObject("VBScript.RegExp")
    With RE
        .Pattern = "^" & 接頭辞 & "\d\d\d\d\d\d\d\d-\d\d\d\d$" '検索パターンを設定
        .IgnoreCase = False '大文字と小文字を区別する
        .Global = True '文字列全体を検索
        Set reMatch = .Execute(ID番号)
        If reMatch.Count = 0 Then
            ID番号書式チェック = False
            Exit Function
        End If
    End With
    Set RE = Nothing
    ID番号書式チェック = True
    
End Function

'---------------------------------------------------------------------------------------------
Function プロジェクト書式チェック(プロジェクト As String) As Boolean

    Dim RE, reMatch
    Set RE = CreateObject("VBScript.RegExp")
    With RE
        .Pattern = "\d\d\d\d$" '検索パターンを設定
        .IgnoreCase = False '大文字と小文字を区別する
        .Global = True '文字列全体を検索
        Set reMatch = .Execute(プロジェクト)
        If reMatch.Count = 0 Then
            プロジェクト書式チェック = False
            Exit Function
        End If
    End With
    Set RE = Nothing
    プロジェクト書式チェック = True
    
End Function

'---------------------------------------------------------------------------------------------
Function チケット名書式チェック(チケット名 As String) As Boolean

    Dim RE, reMatch
    Set RE = CreateObject("VBScript.RegExp")
    With RE
        .Pattern = "^#" & "\d\d\d\d(||\d)$" '検索パターンを設定
        .IgnoreCase = False '大文字と小文字を区別する
        .Global = True '文字列全体を検索
        Set reMatch = .Execute(チケット名)
        If reMatch.Count = 0 Then
            チケット名書式チェック = False
            Exit Function
        End If
    End With
    Set RE = Nothing
    チケット名書式チェック = True
    
End Function

' NULLブランク文字列変換
Function Null2Blank(val As Variant) As Variant

    If IsNull(val) Then
        Null2Blank = ""
    Else
        Null2Blank = val
    End If

End Function

' ブランク文字列NULL変換
Function Blank2Null(val As Variant) As Variant
    
    If val = "" Then
        Blank2Null = Null
    Else
        Blank2Null = val
    End If

End Function

'---------------------------------------------------------------------------------------------
' 列番号アルファベット変換
'
' 引数   lngColNum : 列番号
' 戻り値 列アルファベット文字列
'---------------------------------------------------------------------------------------------
Function ColNum2Txt(lngColNum As Long) As String

  On Error GoTo ErrHandler

  Dim strAddr As String

  strAddr = ThisWorkbook.ActiveSheet.Cells(1, lngColNum).Address(False, False)
  ColNum2Txt = Left(strAddr, Len(strAddr) - 1)

  Exit Function

ErrHandler:

  ColNum2Txt = ""

End Function

' 日付項目DB設定用整形
Function DBDateInput(dateStr As String) As Variant
    If dateStr = "" Or dateStr = "-" Then
        DBDateInput = Null
    Else
        DBDateInput = dateStr
    End If
End Function

'ユーザーフォームアイコン描画
Sub ShowSystemIcon(フォーム名 As String, IDI As Long)
  Dim objHandle As Long
  Dim hWnd As Long
  Dim PictDC As Long

  objHandle& = LoadIconBynum(0, IDI&)
  hWnd = FindWindow(LP_CLASSNAME, フォーム名)
  PictDC = GetWindowDC(hWnd)
  DrawIcon PictDC, 10, 30, objHandle&
  ReleaseDC hWnd, PictDC
  DestroyIcon objHandle&
End Sub

' 通貨項目DB表示用整形
Function SearchComboboxText(Combobox As Variant, SearchIdx As Long, Keyword As String) As String

    SearchComboboxText = ""
    For i = 0 To Combobox.ListCount - 1
        If Combobox.List(i, SearchIdx) Like ("*" & Keyword & "*") Then
            SearchComboboxText = Combobox.List(i)
            Exit Function
        End If
    Next i
    
End Function

Function 予定チェック(日付 As String) As Boolean
    
    Dim 参照シート As Worksheet, 出力行, 出力列 As Long
    Set 参照シート = ThisWorkbook.Worksheets("予定日付")
    予定チェック = False
    出力行 = 2
    出力列 = 1
    'データ解析
    最終行 = 参照シート.UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    最終列 = 参照シート.UsedRange.Columns(参照シート.UsedRange.Columns.Count).Column
    
    For i = 出力行 To 最終行
        If 日付 = 参照シート.Cells(i, 1).Value Then
            予定チェック = True
            Exit For
        End If
    Next i
    
End Function

Sub 記録日付()

    Dim 参照シート As Worksheet, 出力行 As Long
    Set 参照シート = ThisWorkbook.Worksheets("プロジェクト時間記録")
    
    'データ解析
    最終行 = 参照シート.UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    最終列 = 参照シート.UsedRange.Columns(参照シート.UsedRange.Columns.Count).Column
    
    '日付設定
    ThisWorkbook.Worksheets("日付").Activate
    For i = 1 To 8
        Worksheets("日付").Cells(i, 1).Value = DateAdd("d", i - 8, Date)
    Next i

    
    '日付初期化
    出力行 = 2
    参照シート.Activate
    参照シート.Unprotect (パスワード)
    With 参照シート.Range("A2", "A1000").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=日付!$A$1:$A$8"
    End With
    
    With 参照シート.Range("E2", "E1000").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=勤務設定!$A$2:$A$3"
    End With
        
    参照シート.Protect Password:=保護パスワード, UserInterfaceOnly:=True, AllowFiltering:=True
    
End Sub


