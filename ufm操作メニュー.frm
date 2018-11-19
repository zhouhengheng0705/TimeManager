VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufm操作メニュー 
   Caption         =   "操作メニュー"
   ClientHeight    =   2115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3060
   OleObjectBlob   =   "ufm操作メニュー.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufm操作メニュー"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    
    'ウィンドウ位置右上(Excel基準)
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 5
    Me.Left = Application.Left + Application.Width - Me.Width - 92
    
    'ウィンドウ最小化ボタン有効
    Call FrmDec(Me.Caption, True)
    
'    cb最前面表示_Change
    cbExcel表示.Value = True
'    cbExcel表示_Change
    
End Sub

Private Sub btnパスワード変更_Click()
    ufmパスワード変更.Show
End Sub

Private Sub btn時間管理_Click()
    ufm時間管理ツール.Show
End Sub

Private Sub btnチケット管理_Click()
    ufmチケット管理.Show
End Sub

Private Sub btnカレンダー_Click()
    ufmカレンダー.Show
End Sub
