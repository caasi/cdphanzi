VERSION 5.00
Begin VB.Form frm圖片設定 
   Caption         =   "設定複製圖片"
   ClientHeight    =   4140
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   5460
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox lst自訂 
      Height          =   1848
      ItemData        =   "frm圖片設定.frx":0000
      Left            =   1128
      List            =   "frm圖片設定.frx":0025
      TabIndex        =   10
      Top             =   2028
      Width           =   912
   End
   Begin VB.TextBox txt自訂 
      Height          =   264
      HideSelection   =   0   'False
      Left            =   1128
      TabIndex        =   9
      Top             =   1620
      Width           =   876
   End
   Begin VB.OptionButton opt自訂 
      Caption         =   "自訂"
      Height          =   384
      Left            =   312
      TabIndex        =   8
      Top             =   1560
      Width           =   1296
   End
   Begin VB.OptionButton opt印表機 
      Caption         =   "印表機    600"
      Height          =   384
      Left            =   312
      TabIndex        =   7
      Top             =   1044
      Width           =   1296
   End
   Begin VB.OptionButton opt顯示器 
      Caption         =   "顯示器    120"
      Height          =   384
      Left            =   312
      TabIndex        =   6
      Top             =   528
      Width           =   1296
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   4284
      TabIndex        =   3
      Top             =   996
      Width           =   864
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   375
      Left            =   4284
      TabIndex        =   2
      Top             =   468
      Width           =   864
   End
   Begin VB.ListBox lstFontsize 
      Height          =   3108
      ItemData        =   "frm圖片設定.frx":0063
      Left            =   2904
      List            =   "frm圖片設定.frx":00B1
      TabIndex        =   1
      Top             =   864
      Width           =   936
   End
   Begin VB.TextBox txtFontsize 
      Height          =   264
      HideSelection   =   0   'False
      Left            =   2904
      TabIndex        =   0
      Top             =   468
      Width           =   876
   End
   Begin VB.Label lblResolution 
      Caption         =   "解析度(像素/英吋)"
      Height          =   276
      Left            =   324
      TabIndex        =   5
      Top             =   156
      Width           =   1536
   End
   Begin VB.Label lblfontsize 
      Caption         =   "大小"
      Height          =   252
      Left            =   2904
      TabIndex        =   4
      Top             =   156
      Width           =   480
   End
End
Attribute VB_Name = "frm圖片設定"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CancelButton_Click()

Unload Me

End Sub

Private Sub Form_Load()

Dim i As Integer

Me.Top = (Screen.Height * 0.85) \ 2 - Me.Height \ 2
Me.Left = Screen.Width \ 2 - Me.Width \ 2

txtFontsize = 圖片字型大小
If 圖片解析度 = 120 Then
    opt顯示器.Value = True
    txt自訂 = 圖片解析度
ElseIf 圖片解析度 = 600 Then
    opt印表機.Value = True
    txt自訂 = 圖片解析度
Else
    opt自訂.Value = True
    txt自訂 = 圖片解析度
End If

End Sub


Private Sub lstFontsize_Click()

If lstFontsize.ListIndex > -1 Then
    txtFontsize.Text = lstFontsize.List(lstFontsize.ListIndex)
    txtFontsize.SelStart = 0
    txtFontsize.SelLength = Len(txtFontsize.Text)
End If


End Sub


Private Sub lst自訂_Click()

If lst自訂.ListIndex > -1 Then
    txt自訂 = lst自訂.List(lst自訂.ListIndex)
    txt自訂.SelStart = 0
    txt自訂.SelLength = Len(txt自訂.Text)
    opt自訂.Value = True
End If

End Sub

Private Sub OKButton_Click()

Dim 字體大小 As Integer


字體大小 = Val(txtFontsize)
If 字體大小 > 0 Then 圖片字型大小 = 字體大小
If opt顯示器 Then
    圖片解析度 = 120
ElseIf opt印表機 Then
    圖片解析度 = 600
ElseIf opt自訂 Then
    If Val(txt自訂) > 0 Then 圖片解析度 = Val(txt自訂)
End If

mdi漢字字形.cbo圖片大小 = 圖片字型大小
mdi漢字字形.cbo解析度 = 圖片解析度

Unload Me

End Sub

