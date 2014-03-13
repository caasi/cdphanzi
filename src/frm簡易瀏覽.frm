VERSION 5.00
Begin VB.Form frm簡易瀏覽 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "預設開啟(基本)"
   ClientHeight    =   4608
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   7200
   Icon            =   "frm簡易瀏覽.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4608
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic預設瀏覽 
      Height          =   1620
      Index           =   1
      Left            =   3168
      ScaleHeight     =   1572
      ScaleWidth      =   1968
      TabIndex        =   10
      Top             =   420
      Width           =   2016
      Begin VB.Label lbl字形結構 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "字形結構"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1032
         Index           =   1
         Left            =   1404
         TabIndex        =   13
         Top             =   336
         Width           =   360
      End
      Begin VB.Label lbl字形孳乳 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "部件檢字"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1008
         Index           =   1
         Left            =   276
         TabIndex        =   11
         Top             =   324
         Width           =   276
      End
      Begin VB.Shape shp字形孳乳 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  '不透明
         Height          =   1296
         Index           =   1
         Left            =   132
         Top             =   156
         Width           =   580
      End
      Begin VB.Label lbl異體字表 
         Appearance      =   0  '平面
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "異體字表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1080
         Index           =   1
         Left            =   852
         TabIndex        =   12
         Top             =   324
         Width           =   336
      End
      Begin VB.Shape shp異體字表 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  '不透明
         Height          =   1296
         Index           =   1
         Left            =   684
         Top             =   156
         Width           =   580
      End
      Begin VB.Shape shp字形結構 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  '不透明
         Height          =   1296
         Index           =   1
         Left            =   1236
         Top             =   156
         Width           =   580
      End
   End
   Begin VB.PictureBox pic預設瀏覽 
      Height          =   1620
      Index           =   3
      Left            =   3168
      ScaleHeight     =   1572
      ScaleWidth      =   1968
      TabIndex        =   8
      Top             =   2484
      Width           =   2016
      Begin VB.Label lbl字形索引 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "字形索引"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1020
         Index           =   3
         Left            =   1400
         TabIndex        =   16
         Top             =   288
         Width           =   348
      End
      Begin VB.Label lbl字形孳乳 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "部件檢字"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1008
         Index           =   3
         Left            =   276
         TabIndex        =   9
         Top             =   288
         Width           =   276
      End
      Begin VB.Shape shp字形孳乳 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  '不透明
         Height          =   1296
         Index           =   3
         Left            =   132
         Top             =   156
         Width           =   576
      End
      Begin VB.Label lbl字形結構 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "字形結構"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1032
         Index           =   3
         Left            =   864
         TabIndex        =   15
         Top             =   288
         Width           =   360
      End
      Begin VB.Shape shp字形結構 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  '不透明
         Height          =   1296
         Index           =   3
         Left            =   696
         Top             =   156
         Width           =   576
      End
      Begin VB.Shape shp字形索引 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   1  '不透明
         Height          =   1296
         Index           =   3
         Left            =   1237
         Top             =   156
         Width           =   576
      End
   End
   Begin VB.PictureBox pic預設瀏覽 
      Height          =   1620
      Index           =   0
      Left            =   600
      ScaleHeight     =   1572
      ScaleWidth      =   1932
      TabIndex        =   2
      Top             =   408
      Width           =   1980
      Begin VB.Label lbl字形索引 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "字形結構"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1020
         Index           =   0
         Left            =   840
         TabIndex        =   17
         Top             =   324
         Width           =   348
      End
      Begin VB.Label lbl構字部件 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "部件符號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1020
         Index           =   0
         Left            =   1400
         TabIndex        =   4
         Top             =   324
         Width           =   348
      End
      Begin VB.Label lbl字形孳乳 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "部件檢字"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1008
         Index           =   0
         Left            =   264
         TabIndex        =   3
         Top             =   324
         Width           =   276
      End
      Begin VB.Shape shp字形孳乳 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  '不透明
         Height          =   1296
         Index           =   0
         Left            =   132
         Top             =   156
         Width           =   576
      End
      Begin VB.Shape shp構字部件 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  '不透明
         Height          =   1296
         Index           =   0
         Left            =   1237
         Top             =   156
         Width           =   576
      End
      Begin VB.Shape shp字形索引 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  '不透明
         Height          =   1296
         Index           =   0
         Left            =   696
         Top             =   156
         Width           =   576
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   5712
      TabIndex        =   1
      Top             =   984
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "確定"
      Height          =   375
      Left            =   5700
      TabIndex        =   0
      Top             =   444
      Width           =   1215
   End
   Begin VB.PictureBox pic預設瀏覽 
      Height          =   1620
      Index           =   2
      Left            =   576
      ScaleHeight     =   1572
      ScaleWidth      =   1968
      TabIndex        =   5
      Top             =   2484
      Width           =   2016
      Begin VB.Label lbl字形索引 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "字形索引"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1020
         Index           =   2
         Left            =   1400
         TabIndex        =   14
         Top             =   324
         Width           =   348
      End
      Begin VB.Label lbl字形孳乳 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "部件檢字"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1008
         Index           =   2
         Left            =   264
         TabIndex        =   7
         Top             =   324
         Width           =   276
      End
      Begin VB.Label lbl異體字表 
         Appearance      =   0  '平面
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '透明
         Caption         =   "異體字表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1080
         Index           =   2
         Left            =   840
         TabIndex        =   6
         Top             =   324
         Width           =   336
      End
      Begin VB.Shape shp字形孳乳 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  '不透明
         Height          =   1296
         Index           =   2
         Left            =   120
         Top             =   156
         Width           =   580
      End
      Begin VB.Shape shp異體字表 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  '不透明
         Height          =   1296
         Index           =   2
         Left            =   679
         Top             =   156
         Width           =   580
      End
      Begin VB.Shape shp字形索引 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   1  '不透明
         Height          =   1296
         Index           =   2
         Left            =   1237
         Top             =   156
         Width           =   580
      End
   End
   Begin VB.Shape shp預設瀏覽 
      BorderWidth     =   2
      Height          =   2076
      Index           =   1
      Left            =   2892
      Top             =   216
      Visible         =   0   'False
      Width           =   2592
   End
   Begin VB.Shape shp預設瀏覽 
      BorderWidth     =   2
      Height          =   2076
      Index           =   3
      Left            =   2904
      Top             =   2280
      Visible         =   0   'False
      Width           =   2592
   End
   Begin VB.Shape shp預設瀏覽 
      BorderWidth     =   2
      Height          =   2076
      Index           =   2
      Left            =   324
      Top             =   2280
      Visible         =   0   'False
      Width           =   2592
   End
   Begin VB.Shape shp預設瀏覽 
      BorderWidth     =   2
      Height          =   2076
      Index           =   0
      Left            =   324
      Top             =   216
      Visible         =   0   'False
      Width           =   2592
   End
End
Attribute VB_Name = "frm簡易瀏覽"
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

For i = 0 To 簡易瀏覽總數 - 1
    If i = 預設瀏覽模式 - 1 And 簡易瀏覽模式 Then
        shp預設瀏覽(i).Visible = True
    Else
        shp預設瀏覽(i).Visible = False
    End If
Next i

改變預設瀏覽 = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

計算現用視窗

End Sub

Private Sub lbl字形孳乳_Click(Index As Integer)

Dim i As Integer

For i = 0 To 簡易瀏覽總數 - 1
    If i = Index Then
        shp預設瀏覽(i).Visible = True
    Else
        shp預設瀏覽(i).Visible = False
    End If
Next i

End Sub

Private Sub lbl字形結構_Click(Index As Integer)

Dim i As Integer

For i = 0 To 簡易瀏覽總數 - 1
    If i = Index Then
        shp預設瀏覽(i).Visible = True
    Else
        shp預設瀏覽(i).Visible = False
    End If
Next i

End Sub

Private Sub lbl異體字表_Click(Index As Integer)

Dim i As Integer

For i = 0 To 簡易瀏覽總數 - 1
    If i = Index Then
        shp預設瀏覽(i).Visible = True
    Else
        shp預設瀏覽(i).Visible = False
    End If
Next i

End Sub

Private Sub lbl字形索引_Click(Index As Integer)

Dim i As Integer

For i = 0 To 簡易瀏覽總數 - 1
    If i = Index Then
        shp預設瀏覽(i).Visible = True
    Else
        shp預設瀏覽(i).Visible = False
    End If
Next i

End Sub

Private Sub lbl構字部件_Click(Index As Integer)

Dim i As Integer

For i = 0 To 簡易瀏覽總數 - 1
    If i = Index Then
        shp預設瀏覽(i).Visible = True
    Else
        shp預設瀏覽(i).Visible = False
    End If
Next i

End Sub

Private Sub OKButton_Click()

Dim i As Integer

For i = 0 To 簡易瀏覽總數 - 1
    If shp預設瀏覽(i).Visible Then
        預設瀏覽模式 = i + 1
    End If
Next i

簡易瀏覽模式 = True
改變預設瀏覽 = True

Unload Me

End Sub

Private Sub pic預設瀏覽_Click(Index As Integer)

Dim i As Integer

For i = 0 To 簡易瀏覽總數 - 1
    If i = Index Then
        shp預設瀏覽(i).Visible = True
    Else
        shp預設瀏覽(i).Visible = False
    End If
Next i

End Sub
