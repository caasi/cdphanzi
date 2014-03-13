VERSION 5.00
Begin VB.Form frm版本 
   BorderStyle     =   1  '單線固定
   Caption         =   "關於漢字構形資料庫"
   ClientHeight    =   4644
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7788
   Icon            =   "frm版本.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   7788
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   2316
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "cdpservice@iis.sinica.edu.tw"
      Top             =   3060
      Width           =   4200
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   2340
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "http://www.sinica.edu.tw/~cdp/cdphanzi/"
      Top             =   2472
      Width           =   4200
   End
   Begin VB.CommandButton cmd確定 
      Caption         =   "確　定"
      Height          =   375
      Left            =   3287
      TabIndex        =   5
      Top             =   3984
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "服務信箱："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1092
      TabIndex        =   6
      Top             =   3120
      Width           =   1212
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   7320
      Y1              =   3744
      Y2              =   3744
   End
   Begin VB.Label lbl版本_5 
      Caption         =   "更新網址："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1092
      TabIndex        =   4
      Top             =   2544
      Width           =   1212
   End
   Begin VB.Label lbl版本_4 
      Caption         =   "漢字庫字型為葉健欣製作"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1092
      TabIndex        =   3
      Top             =   1968
      Width           =   3375
   End
   Begin VB.Label lbl版本_3 
      Caption         =   "小篆字型原為北京師範大學製作"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1092
      TabIndex        =   2
      Top             =   1392
      Width           =   5652
   End
   Begin VB.Label lbl版本_2 
      Caption         =   "研發單位：中央研究院資訊科學研究所文獻處理實驗室"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1092
      TabIndex        =   1
      Top             =   816
      Width           =   5892
   End
   Begin VB.Label lbl版本_1 
      Caption         =   "漢字構形資料庫 2.51(2007/12/24)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1092
      TabIndex        =   0
      Top             =   240
      Width           =   5772
   End
   Begin VB.Image img版本 
      Height          =   384
      Left            =   240
      Picture         =   "frm版本.frx":030A
      Top             =   120
      Width           =   384
   End
End
Attribute VB_Name = "frm版本"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd確定_Click()
Unload Me
End Sub


Private Sub Form_Load()
frm版本.Top = mdi漢字字形.Top + (mdi漢字字形.Height - frm版本.Height) / 2 + mdi漢字字形.pic構字符號.Height + mdi漢字字形.pic字形屬性.Height
frm版本.Left = mdi漢字字形.Left + (mdi漢字字形.Width - frm版本.Width) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
計算現用視窗

End Sub

