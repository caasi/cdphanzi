VERSION 5.00
Begin VB.Form frm���� 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "����~�r�c�θ�Ʈw"
   ClientHeight    =   4644
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7788
   Icon            =   "frm����.frx":0000
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
   Begin VB.CommandButton cmd�T�w 
      Caption         =   "�T�@�w"
      Height          =   375
      Left            =   3287
      TabIndex        =   5
      Top             =   3984
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "�A�ȫH�c�G"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
   Begin VB.Label lbl����_5 
      Caption         =   "��s���}�G"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
   Begin VB.Label lbl����_4 
      Caption         =   "�~�r�w�r���������Y�s�@"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
   Begin VB.Label lbl����_3 
      Caption         =   "�p�f�r���쬰�_�ʮv�d�j�ǻs�@"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
   Begin VB.Label lbl����_2 
      Caption         =   "��o���G������s�|��T��Ǭ�s�Ҥ��m�B�z�����"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
   Begin VB.Label lbl����_1 
      Caption         =   "�~�r�c�θ�Ʈw 2.51(2007/12/24)"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
   Begin VB.Image img���� 
      Height          =   384
      Left            =   240
      Picture         =   "frm����.frx":030A
      Top             =   120
      Width           =   384
   End
End
Attribute VB_Name = "frm����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd�T�w_Click()
Unload Me
End Sub


Private Sub Form_Load()
frm����.Top = mdi�~�r�r��.Top + (mdi�~�r�r��.Height - frm����.Height) / 2 + mdi�~�r�r��.pic�c�r�Ÿ�.Height + mdi�~�r�r��.pic�r���ݩ�.Height
frm����.Left = mdi�~�r�r��.Left + (mdi�~�r�r��.Width - frm����.Width) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
�p��{�ε���

End Sub

