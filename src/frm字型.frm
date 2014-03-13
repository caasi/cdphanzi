VERSION 5.00
Begin VB.Form frm字型 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "字型"
   ClientHeight    =   4176
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   5472
   Icon            =   "frm字型.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4176
   ScaleWidth      =   5472
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox txt楚文字 
      Enabled         =   0   'False
      Height          =   264
      Left            =   900
      TabIndex        =   15
      Text            =   "中研院楚系簡帛文字"
      Top             =   3540
      Width           =   1788
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Left            =   900
      TabIndex        =   13
      Text            =   "中研院甲骨文"
      Top             =   3048
      Width           =   1788
   End
   Begin VB.TextBox txtFontsize 
      Height          =   264
      HideSelection   =   0   'False
      Left            =   2964
      TabIndex        =   12
      Top             =   576
      Width           =   876
   End
   Begin VB.ListBox lstFontsize 
      Height          =   2928
      ItemData        =   "frm字型.frx":030A
      Left            =   2964
      List            =   "frm字型.frx":0358
      TabIndex        =   11
      Top             =   972
      Width           =   912
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   375
      Left            =   4344
      TabIndex        =   10
      Top             =   576
      Width           =   864
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   4344
      TabIndex        =   9
      Top             =   1104
      Width           =   864
   End
   Begin VB.TextBox txt金文 
      Enabled         =   0   'False
      Height          =   264
      Left            =   900
      TabIndex        =   8
      Text            =   "中研院金文"
      Top             =   2556
      Width           =   1788
   End
   Begin VB.TextBox txt小篆 
      Enabled         =   0   'False
      Height          =   264
      Left            =   900
      TabIndex        =   7
      Text            =   "北師大說文小篆"
      Top             =   2064
      Width           =   1788
   End
   Begin VB.ListBox lst楷書 
      Height          =   768
      ItemData        =   "frm字型.frx":03B0
      Left            =   900
      List            =   "frm字型.frx":03BA
      TabIndex        =   6
      Top             =   948
      Width           =   1788
   End
   Begin VB.TextBox txt楷書 
      Height          =   264
      HideSelection   =   0   'False
      Left            =   900
      TabIndex        =   5
      Top             =   588
      Width           =   1788
   End
   Begin VB.Label lbl楚文字 
      Caption         =   "楚文字"
      Height          =   252
      Left            =   156
      TabIndex        =   16
      Top             =   3540
      Width           =   588
   End
   Begin VB.Label lbl甲骨文 
      Caption         =   "甲骨文"
      Height          =   252
      Left            =   156
      TabIndex        =   14
      Top             =   3048
      Width           =   588
   End
   Begin VB.Label lblfontsize 
      Caption         =   "大小"
      Height          =   252
      Left            =   2964
      TabIndex        =   4
      Top             =   264
      Width           =   480
   End
   Begin VB.Label lbl金文 
      Caption         =   "金文"
      Height          =   252
      Left            =   156
      TabIndex        =   3
      Top             =   2556
      Width           =   480
   End
   Begin VB.Label lbl小篆 
      Caption         =   "小篆"
      Height          =   252
      Left            =   156
      TabIndex        =   2
      Top             =   2064
      Width           =   480
   End
   Begin VB.Label lbl楷書 
      Caption         =   "楷書"
      Height          =   252
      Left            =   180
      TabIndex        =   1
      Top             =   594
      Width           =   480
   End
   Begin VB.Label lblFontName 
      Caption         =   "字型"
      Height          =   276
      Left            =   900
      TabIndex        =   0
      Top             =   264
      Width           =   588
   End
End
Attribute VB_Name = "frm字型"
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

txt楷書.Text = 顯示字型
txt楷書.SelStart = 0
txt楷書.SelLength = Len(txt楷書.Text)

txtFontsize.Text = 顯示字型大小
txtFontsize.SelStart = 0
txtFontsize.SelLength = Len(txt楷書.Text)

For i = 0 To lst楷書.ListCount - 1
    If lst楷書.List(i) = 顯示字型 Then
        lst楷書.ListIndex = i
        Exit For
    End If
Next i

For i = 0 To lstFontsize.ListCount - 1
    If lstFontsize.List(i) = 顯示字型大小 Then
        lstFontsize.ListIndex = i
        Exit For
    End If
Next i

End Sub

Private Sub lstFontsize_Click()

If lstFontsize.ListIndex > -1 Then
    txtFontsize.Text = lstFontsize.List(lstFontsize.ListIndex)
    txtFontsize.SelStart = 0
    txtFontsize.SelLength = Len(txtFontsize.Text)
End If


End Sub


Private Sub lst楷書_Click()

If lst楷書.ListIndex > -1 Then
    txt楷書.Text = lst楷書.List(lst楷書.ListIndex)
    txt楷書.SelStart = 0
    txt楷書.SelLength = Len(txt楷書.Text)
End If

End Sub

Private Sub OKButton_Click()

Dim 字體大小 As Integer

If txt楷書.Text = "細明體" Or txt楷書.Text = "標楷體" Then
    If txt楷書.Text <> 顯示字型 Then
        mdi漢字字形.cbo字型名稱.Text = txt楷書.Text
    End If
End If

字體大小 = Val(txtFontsize.Text)
If 字體大小 <> 顯示字型大小 Then
    If 字體大小 < 10 Then
        字體大小 = 10
    'ElseIf 字體大小 > 72 Then
    '    字體大小 = 72
    End If
    mdi漢字字形.cbo字體大小.Text = 字體大小
End If

Unload Me

End Sub

