VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 確認ダイアログ 
   Caption         =   "Microsoft Excel"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4875
   OleObjectBlob   =   "確認ダイアログ.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "確認ダイアログ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public 結果 As Boolean
Dim FormHeight As Long 'フォームの高さ

Private Sub UserForm_Initialize()
    If Me.Tag = "" Then Me.Tag = 1
    FormHeight = Me.Height
    Me.StartUpPosition = 2 '中央表示固定
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnキャンセル_Click
    End If
End Sub

Private Sub UserForm_Activate()
    Me.Height = IIf(Me.Height = FormHeight, FormHeight + 1, FormHeight)
End Sub

Private Sub UserForm_Resize()
    Dim 高 As Integer
    高 = lblメッセージ.Height + 65
    If 高 > Me.Height Then
        Me.Height = 高
        cb最終確認.Height = lblメッセージ.Height + 20
        btnOK.Top = lblメッセージ.Height + cb最終確認.Height + 20
        btnキャンセル.Top = lblメッセージ.Height + cb最終確認.Height + 20
    End If
    Me.Repaint
    ShowSystemIcon Me.Caption, Me.Tag
End Sub

Sub 表示(メッセージ As String, Optional 最終確認 As String = "")
    結果 = False
    lblメッセージ.Caption = メッセージ
    If 最終確認 = "" Then
        cb最終確認.Value = True
        cb最終確認.Visible = False
    Else
        cb最終確認.Caption = 最終確認
        cb最終確認.Value = False
        cb最終確認.Visible = True
    End If
    Me.Show
    ShowSystemIcon Me.Caption, Me.Tag

End Sub

Private Sub cb最終確認_Click()
    btnOK.Enabled = cb最終確認.Value
End Sub

Private Sub btnキャンセル_Click()
    'ここにキャンセル処理を記述
    結果 = False
    Me.Hide
'    Unload Me
End Sub

Private Sub btnOK_Click()
    'ここにOK処理を記述
    結果 = True
    Me.Hide
'    Unload Me
End Sub

