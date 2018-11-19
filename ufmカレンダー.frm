VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufmカレンダー 
   Caption         =   "カレンダー"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3540
   OleObjectBlob   =   "ufmカレンダー.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "ufmカレンダー"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const 開始年 = 2012

Private Sub Label1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Sub UserForm_Initialize()

    Dim i As Long
    
    cmb月.Clear
    For i = 1 To 12
        cmb月.AddItem i
    Next i
    cmb月.Value = Month(Date)
    cmb年.Clear
    For i = 開始年 To Year(Date) + 10
        cmb年.AddItem i
    Next i
    cmb年.Value = Year(Date)
'    txt現在時刻.Value = Format(Now(), "hh:mm")
    
    Call 初期化
    
End Sub

'Sub Sample1()
'
'    Application.OnTime (TimeValue(Format(Now(), "hh:mm:ss")) + TimeValue("00:00:02")), "Test"
'
'End Sub
'
'Sub Test()
'
'    txt現在時刻.Value = Format(Now(), "hh:mm")
'
'End Sub

Private Sub SpinButton1_SpinUp()

    If cmb年.Value = "" Or cmb月.Value = "" Then
        MsgBox "年月を正しく入力してください。", vbExclamation
        Exit Sub
    End If
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Dim 現在日付, 前月日付 As String
    現在日付 = DateSerial(cmb年.Value, cmb月.Value, 1)
    前月日付 = DateAdd("m", -1, 現在日付)
    cmb年.Value = Year(前月日付)
    cmb月.Value = Month(前月日付)
    
    Call 初期化

End Sub

Private Sub SpinButton1_SpinDown()

    If cmb年.Value = "" Or cmb月.Value = "" Then
        MsgBox "年月を正しく入力してください。", vbExclamation
        Exit Sub
    End If
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Dim 現在日付, 次月日付 As String
    現在日付 = DateSerial(cmb年.Value, cmb月.Value, 1)
    次月日付 = DateAdd("m", 1, 現在日付)
    cmb年.Value = Year(次月日付)
    cmb月.Value = Month(次月日付)
    
    Call 初期化

End Sub

Private Sub cmb年_Change()
    
    If cmb年.Value <> "" And cmb月.Value <> "" Then
        Call カレンダーフォーム初期化
        Call 初期化
    End If

End Sub

Private Sub cmb月_Change()
    
    If cmb年.Value <> "" And cmb月.Value <> "" Then
        Call カレンダーフォーム初期化
        Call 初期化
    End If

End Sub

Private Sub Label1_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label1.SpecialEffect = fmSpecialEffectBump
    Label1.Font.Bold = True
    
End Sub

Private Sub Label2_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label2.SpecialEffect = fmSpecialEffectBump
    Label2.Font.Bold = True
    
End Sub

Private Sub Label3_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label3.SpecialEffect = fmSpecialEffectBump
    Label3.Font.Bold = True

End Sub

Private Sub Label4_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label4.SpecialEffect = fmSpecialEffectBump
    Label4.Font.Bold = True
    
End Sub

Private Sub Label5_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label5.SpecialEffect = fmSpecialEffectBump
    Label5.Font.Bold = True
    
End Sub

Private Sub Label6_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label6.SpecialEffect = fmSpecialEffectBump
    Label6.Font.Bold = True

End Sub

Private Sub Label7_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label7.SpecialEffect = fmSpecialEffectBump
    Label7.Font.Bold = True
    
End Sub

Private Sub Label8_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label8.SpecialEffect = fmSpecialEffectBump
    Label8.Font.Bold = True
    
End Sub

Private Sub Label9_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label9.SpecialEffect = fmSpecialEffectBump
    Label9.Font.Bold = True

End Sub

Private Sub Label10_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label10.SpecialEffect = fmSpecialEffectBump
    Label10.Font.Bold = True
    
End Sub

Private Sub Label11_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label11.SpecialEffect = fmSpecialEffectBump
    Label11.Font.Bold = True
    
End Sub

Private Sub Label12_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label12.SpecialEffect = fmSpecialEffectBump
    Label12.Font.Bold = True

End Sub

Private Sub Label13_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label13.SpecialEffect = fmSpecialEffectBump
    Label13.Font.Bold = True
    
End Sub

Private Sub Label14_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label14.SpecialEffect = fmSpecialEffectBump
    Label14.Font.Bold = True
    
End Sub

Private Sub Label15_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label15.SpecialEffect = fmSpecialEffectBump
    Label15.Font.Bold = True

End Sub

Private Sub Label16_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label16.SpecialEffect = fmSpecialEffectBump
    Label16.Font.Bold = True
    
End Sub

Private Sub Label17_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label17.SpecialEffect = fmSpecialEffectBump
    Label17.Font.Bold = True
    
End Sub

Private Sub Label18_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label18.SpecialEffect = fmSpecialEffectBump
    Label18.Font.Bold = True

End Sub

Private Sub Label19_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label19.SpecialEffect = fmSpecialEffectBump
    Label19.Font.Bold = True
    
End Sub

Private Sub Label20_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label20.SpecialEffect = fmSpecialEffectBump
    Label20.Font.Bold = True
    
End Sub

Private Sub Label21_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label21.SpecialEffect = fmSpecialEffectBump
    Label21.Font.Bold = True

End Sub

Private Sub Label22_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label22.SpecialEffect = fmSpecialEffectBump
    Label22.Font.Bold = True
    
End Sub

Private Sub Label23_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label23.SpecialEffect = fmSpecialEffectBump
    Label23.Font.Bold = True
    
End Sub

Private Sub Label24_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label24.SpecialEffect = fmSpecialEffectBump
    Label24.Font.Bold = True

End Sub

Private Sub Label25_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label25.SpecialEffect = fmSpecialEffectBump
    Label25.Font.Bold = True
    
End Sub

Private Sub Label26_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label26.SpecialEffect = fmSpecialEffectBump
    Label26.Font.Bold = True
    
End Sub

Private Sub Label27_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label27.SpecialEffect = fmSpecialEffectBump
    Label27.Font.Bold = True

End Sub

Private Sub Label28_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label28.SpecialEffect = fmSpecialEffectBump
    Label28.Font.Bold = True
    
End Sub

Private Sub Label29_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label29.SpecialEffect = fmSpecialEffectBump
    Label29.Font.Bold = True
    
End Sub

Private Sub Label30_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label30.SpecialEffect = fmSpecialEffectBump
    Label30.Font.Bold = True

End Sub

Private Sub Label31_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label31.SpecialEffect = fmSpecialEffectBump
    Label31.Font.Bold = True
    
End Sub

Private Sub Label32_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label32.SpecialEffect = fmSpecialEffectBump
    Label32.Font.Bold = True
    
End Sub

Private Sub Label33_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label33.SpecialEffect = fmSpecialEffectBump
    Label33.Font.Bold = True

End Sub

Private Sub Label34_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label34.SpecialEffect = fmSpecialEffectBump
    Label34.Font.Bold = True
    
End Sub

Private Sub Label35_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label35.SpecialEffect = fmSpecialEffectBump
    Label35.Font.Bold = True
    
End Sub

Private Sub Label36_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label36.SpecialEffect = fmSpecialEffectBump
    Label36.Font.Bold = True

End Sub

Private Sub Label37_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label37.SpecialEffect = fmSpecialEffectBump
    Label37.Font.Bold = True
    
End Sub

Private Sub Label38_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label38.SpecialEffect = fmSpecialEffectBump
    Label38.Font.Bold = True
    
End Sub

Private Sub Label39_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label39.SpecialEffect = fmSpecialEffectBump
    Label39.Font.Bold = True

End Sub

Private Sub Label40_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化

    'クリック時書式変化
    Label40.SpecialEffect = fmSpecialEffectBump
    Label40.Font.Bold = True
    
End Sub

Private Sub Label41_Click()
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label41.SpecialEffect = fmSpecialEffectBump
    Label41.Font.Bold = True
    
End Sub

Private Sub Label42_Click()

    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label42.SpecialEffect = fmSpecialEffectBump
    Label42.Font.Bold = True

End Sub


Private Sub Label1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label1.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label1.Tag
    ufmスケジュール.Caption = Label1.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label1.SpecialEffect = fmSpecialEffectBump
    Label1.Font.Bold = True

End Sub

Private Sub Label2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label2.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label2.Tag
    ufmスケジュール.Caption = Label2.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label2.SpecialEffect = fmSpecialEffectBump
    Label2.Font.Bold = True
    
End Sub


Private Sub Label3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label3.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label3.Tag
    ufmスケジュール.Caption = Label3.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label3.SpecialEffect = fmSpecialEffectBump
    Label3.Font.Bold = True
    
End Sub


Private Sub Label4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label4.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label4.Tag
    ufmスケジュール.Caption = Label4.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label4.SpecialEffect = fmSpecialEffectBump
    Label4.Font.Bold = True
End Sub

Private Sub Label5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label5.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label5.Tag
    ufmスケジュール.Caption = Label5.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    'クリック時書式変化
    Label5.SpecialEffect = fmSpecialEffectBump
    Label5.Font.Bold = True
End Sub

Private Sub Label6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label6.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label6.Tag
    ufmスケジュール.Caption = Label6.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label6.SpecialEffect = fmSpecialEffectBump
    Label6.Font.Bold = True
End Sub

Private Sub Label7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label7.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label7.Tag
    ufmスケジュール.Caption = Label7.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label7.SpecialEffect = fmSpecialEffectBump
    Label7.Font.Bold = True
End Sub

Private Sub Label8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label8.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label8.Tag
    ufmスケジュール.Caption = Label8.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label8.SpecialEffect = fmSpecialEffectBump
    Label8.Font.Bold = True
End Sub

Private Sub Label9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label9.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label9.Tag
    ufmスケジュール.Caption = Label9.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label9.SpecialEffect = fmSpecialEffectBump
    Label9.Font.Bold = True
End Sub

Private Sub Label10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label10.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label10.Tag
    ufmスケジュール.Caption = Label10.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label10.SpecialEffect = fmSpecialEffectBump
    Label10.Font.Bold = True
End Sub

Private Sub Label11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label11.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label11.Tag
    ufmスケジュール.Caption = Label11.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label11.SpecialEffect = fmSpecialEffectBump
    Label11.Font.Bold = True
End Sub

Private Sub Label12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label12.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label12.Tag
    ufmスケジュール.Caption = Label12.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label12.SpecialEffect = fmSpecialEffectBump
    Label12.Font.Bold = True
End Sub

Private Sub Label13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label13.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label13.Tag
    ufmスケジュール.Caption = Label13.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label13.SpecialEffect = fmSpecialEffectBump
    Label13.Font.Bold = True
End Sub

Private Sub Label14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label14.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label14.Tag
    ufmスケジュール.Caption = Label14.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label14.SpecialEffect = fmSpecialEffectBump
    Label14.Font.Bold = True
End Sub

Private Sub Label15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label15.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label15.Tag
    ufmスケジュール.Caption = Label15.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label15.SpecialEffect = fmSpecialEffectBump
    Label15.Font.Bold = True
End Sub

Private Sub Label16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label16.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label16.Tag
    ufmスケジュール.Caption = Label16.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label16.SpecialEffect = fmSpecialEffectBump
    Label16.Font.Bold = True
End Sub

Private Sub Label17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label17.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label17.Tag
    ufmスケジュール.Caption = Label17.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label17.SpecialEffect = fmSpecialEffectBump
    Label17.Font.Bold = True
End Sub

Private Sub Label18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label18.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label18.Tag
    ufmスケジュール.Caption = Label18.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label18.SpecialEffect = fmSpecialEffectBump
    Label18.Font.Bold = True
End Sub

Private Sub Label19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label19.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label19.Tag
    ufmスケジュール.Caption = Label19.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label19.SpecialEffect = fmSpecialEffectBump
    Label19.Font.Bold = True
End Sub

Private Sub Label20_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label20.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label20.Tag
    ufmスケジュール.Caption = Label20.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label20.SpecialEffect = fmSpecialEffectBump
    Label20.Font.Bold = True
End Sub

Private Sub Label21_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label21.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label21.Tag
    ufmスケジュール.Caption = Label21.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label21.SpecialEffect = fmSpecialEffectBump
    Label21.Font.Bold = True
End Sub

Private Sub Label22_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label22.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label22.Tag
    ufmスケジュール.Caption = Label22.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label22.SpecialEffect = fmSpecialEffectBump
    Label22.Font.Bold = True
End Sub

Private Sub Label23_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label23.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label23.Tag
    ufmスケジュール.Caption = Label23.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label23.SpecialEffect = fmSpecialEffectBump
    Label23.Font.Bold = True
End Sub

Private Sub Label24_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label24.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label24.Tag
    ufmスケジュール.Caption = Label24.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label24.SpecialEffect = fmSpecialEffectBump
    Label24.Font.Bold = True
End Sub

Private Sub Label25_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label25.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label25.Tag
    ufmスケジュール.Caption = Label25.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label25.SpecialEffect = fmSpecialEffectBump
    Label25.Font.Bold = True
End Sub

Private Sub Label26_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label26.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label26.Tag
    ufmスケジュール.Caption = Label26.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label26.SpecialEffect = fmSpecialEffectBump
    Label26.Font.Bold = True
End Sub

Private Sub Label27_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label27.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label27.Tag
    ufmスケジュール.Caption = Label27.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label27.SpecialEffect = fmSpecialEffectBump
    Label27.Font.Bold = True
End Sub

Private Sub Label28_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label28.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label28.Tag
    ufmスケジュール.Caption = Label28.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label28.SpecialEffect = fmSpecialEffectBump
    Label28.Font.Bold = True
End Sub

Private Sub Label29_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label29.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label29.Tag
    ufmスケジュール.Caption = Label29.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label29.SpecialEffect = fmSpecialEffectBump
    Label29.Font.Bold = True
End Sub

Private Sub Label30_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label30.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label30.Tag
    ufmスケジュール.Caption = Label30.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label30.SpecialEffect = fmSpecialEffectBump
    Label30.Font.Bold = True
End Sub

Private Sub Label31_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label31.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label31.Tag
    ufmスケジュール.Caption = Label31.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label31.SpecialEffect = fmSpecialEffectBump
    Label31.Font.Bold = True
End Sub

Private Sub Label32_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label32.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label32.Tag
    ufmスケジュール.Caption = Label32.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label32.SpecialEffect = fmSpecialEffectBump
    Label32.Font.Bold = True
End Sub

Private Sub Label33_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label33.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label33.Tag
    ufmスケジュール.Caption = Label33.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label33.SpecialEffect = fmSpecialEffectBump
    Label33.Font.Bold = True
End Sub

Private Sub Label34_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label34.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label34.Tag
    ufmスケジュール.Caption = Label34.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label34.SpecialEffect = fmSpecialEffectBump
    Label34.Font.Bold = True
End Sub

Private Sub Label35_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label35.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label35.Tag
    ufmスケジュール.Caption = Label35.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label35.SpecialEffect = fmSpecialEffectBump
    Label35.Font.Bold = True
End Sub

Private Sub Label36_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label36.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label36.Tag
    ufmスケジュール.Caption = Label36.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label36.SpecialEffect = fmSpecialEffectBump
    Label36.Font.Bold = True
End Sub

Private Sub Label37_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label37.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label37.Tag
    ufmスケジュール.Caption = Label37.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label37.SpecialEffect = fmSpecialEffectBump
    Label37.Font.Bold = True
End Sub

Private Sub Label38_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label38.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label38.Tag
    ufmスケジュール.Caption = Label38.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label38.SpecialEffect = fmSpecialEffectBump
    Label38.Font.Bold = True
End Sub

Private Sub Label39_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label39.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label39.Tag
    ufmスケジュール.Caption = Label39.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label39.SpecialEffect = fmSpecialEffectBump
    Label39.Font.Bold = True
End Sub

Private Sub Label40_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label40.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label40.Tag
    ufmスケジュール.Caption = Label40.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label40.SpecialEffect = fmSpecialEffectBump
    Label40.Font.Bold = True
End Sub

Private Sub Label41_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label41.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label41.Tag
    ufmスケジュール.Caption = Label41.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    'カレンダーフォーム初期化
    Call カレンダーフォーム初期化
    
    Label41.SpecialEffect = fmSpecialEffectBump
    Label41.Font.Bold = True
End Sub

Private Sub Label42_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    '日付の曜日取得
    Dim 曜日 As String
    曜日 = Format(Weekday(Label42.Tag), "aaa")
    
    'スケジュール呼び出し
    ufmスケジュール.Show
    ufmスケジュール.UserForm_Initialize2 Label42.Tag
    ufmスケジュール.Caption = Label42.Tag & "(" & 曜日 & ")" & "  スケジュール"
    
    Call カレンダーフォーム初期化

    Label42.SpecialEffect = fmSpecialEffectBump
    Label42.Font.Bold = True
End Sub

Sub 初期化()

    Dim 月初日 As String, 月初日曜日, n As Long
    月初日 = DateSerial(cmb年.Value, cmb月.Value, 1)
    前月日数 = Day(DateAdd("d", -1, 月初日))
    当月日数 = Day(DateAdd("d", -1, DateAdd("m", 1, 月初日)))
    月初日曜日 = Weekday(月初日, vbSunday)
    '当月初日から末日まで格納
    n = 1
    For i = 月初日曜日 To (月初日曜日 + 当月日数 - 1)
        Me.Controls("Label" & i).Caption = n
        Me.Controls("Label" & i).Tag = DateSerial(Year(月初日), Month(月初日), n)
        Me.Controls("Label" & i).BackColor = &HC0FFC0
        '当日日付取得
        If Me.Controls("Label" & i).Tag = Trim(Date) Then
            Me.Controls("Label" & i).BackColor = vbYellow
        End If
        n = n + 1
    Next i
    '翌月日付格納
    n = 1
    For i = (月初日曜日 + 当月日数) To 42
        Me.Controls("Label" & i).Caption = n
        Me.Controls("Label" & i).Tag = DateSerial(Year(DateAdd("m", 1, 月初日)), Month(DateAdd("m", 1, 月初日)), n)
        Me.Controls("Label" & i).BackColor = &HE0E0E0
        n = n + 1
    Next i
    '前月日付格納
    n = 前月日数
    For i = (月初日曜日 - 1) To 1 Step -1
        Me.Controls("Label" & i).Caption = n
        Me.Controls("Label" & i).Tag = DateSerial(Year(DateAdd("m", -1, 月初日)), Month(DateAdd("m", -1, 月初日)), n)
        Me.Controls("Label" & i).BackColor = &HE0E0E0
        n = n - 1
    Next i
    '土曜日カラーはブルー設定
    For i = 7 To 42 Step 7
        Me.Controls("Label" & i).ForeColor = vbBlue
    Next i
    '日曜日カラーは赤色設定
    For i = 1 To 36 Step 7
        Me.Controls("Label" & i).ForeColor = vbRed
    Next i
    
    For i = 1 To 42
        If 予定チェック(Me.Controls("Label" & i).Tag) Then
            Me.Controls("Label" & i).BackColor = &HFF80FF
        End If
        Me.Controls("Label" & i).SpecialEffect = fmSpecialEffectFlat
    Next i
    
End Sub

Sub カレンダーフォーム初期化()

    'カレンダーフォーム初期化
    For i = 1 To 42
        Me.Controls("Label" & i).SpecialEffect = fmSpecialEffectFlat
        Me.Controls("Label" & i).FontBold = False
    Next i

End Sub
