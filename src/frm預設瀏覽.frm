VERSION 5.00
Begin VB.Form frm�w�]�s�� 
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "�w�]�}��(�i��)"
   ClientHeight    =   4596
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   7200
   Icon            =   "frm�w�]�s��.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4596
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic�w�]�s�� 
      Height          =   1620
      Index           =   2
      Left            =   600
      ScaleHeight     =   1572
      ScaleWidth      =   1968
      TabIndex        =   17
      Top             =   2484
      Width           =   2016
      Begin VB.Label lbl�r�κt�� 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�r�κt��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   552
         Index           =   2
         Left            =   1296
         TabIndex        =   22
         Top             =   880
         Width           =   480
      End
      Begin VB.Label lbl�r�ί��� 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�r�ί���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   516
         Index           =   2
         Left            =   1300
         TabIndex        =   21
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl�r�δF�� 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�����˦r"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Left            =   276
         TabIndex        =   20
         Top             =   288
         Width           =   276
      End
      Begin VB.Label lbl����r�� 
         Appearance      =   0  '����
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "����r��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   528
         Index           =   2
         Left            =   744
         TabIndex        =   19
         Top             =   888
         Width           =   540
      End
      Begin VB.Label lbl�r�ε��c 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�r�ε��c"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   516
         Index           =   2
         Left            =   744
         TabIndex        =   18
         Top             =   240
         Width           =   480
      End
      Begin VB.Shape shp�r�δF�� 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  '���z��
         Height          =   1300
         Index           =   2
         Left            =   132
         Top             =   156
         Width           =   580
      End
      Begin VB.Shape shp�r�ε��c 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  '���z��
         Height          =   660
         Index           =   2
         Left            =   698
         Top             =   156
         Width           =   576
      End
      Begin VB.Shape shp����r�� 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  '���z��
         Height          =   650
         Index           =   2
         Left            =   698
         Top             =   800
         Width           =   576
      End
      Begin VB.Shape shp�r�ί��� 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   1  '���z��
         Height          =   660
         Index           =   2
         Left            =   1260
         Top             =   156
         Width           =   576
      End
      Begin VB.Shape shp�r�κt�� 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  '���z��
         Height          =   648
         Index           =   2
         Left            =   1260
         Top             =   800
         Width           =   576
      End
   End
   Begin VB.PictureBox pic�w�]�s�� 
      Height          =   1620
      Index           =   0
      Left            =   600
      ScaleHeight     =   1572
      ScaleWidth      =   1968
      TabIndex        =   11
      Top             =   420
      Width           =   2016
      Begin VB.Label lbl�r�κt�� 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�r�κt��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   552
         Index           =   0
         Left            =   1308
         TabIndex        =   23
         Top             =   876
         Width           =   480
      End
      Begin VB.Label lbl����r�� 
         Appearance      =   0  '����
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "����r��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1080
         Index           =   0
         Left            =   852
         TabIndex        =   14
         Top             =   324
         Width           =   336
      End
      Begin VB.Label lbl�r�δF�� 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�����˦r"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   13
         Top             =   324
         Width           =   276
      End
      Begin VB.Label lbl�r�ε��c 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�r�ε��c"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   516
         Index           =   0
         Left            =   1300
         TabIndex        =   12
         Top             =   240
         Width           =   480
      End
      Begin VB.Shape shp�r�ε��c 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  '���z��
         Height          =   660
         Index           =   0
         Left            =   1260
         Top             =   156
         Width           =   576
      End
      Begin VB.Shape shp�r�κt�� 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  '���z��
         Height          =   648
         Index           =   0
         Left            =   1260
         Top             =   800
         Width           =   576
      End
      Begin VB.Shape shp����r�� 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  '���z��
         Height          =   1296
         Index           =   0
         Left            =   698
         Top             =   156
         Width           =   576
      End
      Begin VB.Shape shp�r�δF�� 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  '���z��
         Height          =   1296
         Index           =   0
         Left            =   132
         Top             =   156
         Width           =   580
      End
   End
   Begin VB.PictureBox pic�w�]�s�� 
      Height          =   1620
      Index           =   1
      Left            =   3168
      ScaleHeight     =   1572
      ScaleWidth      =   1968
      TabIndex        =   8
      Top             =   420
      Width           =   2016
      Begin VB.Label lbl�r�κt�� 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�r�κt��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   552
         Index           =   1
         Left            =   1308
         TabIndex        =   16
         Top             =   876
         Width           =   480
      End
      Begin VB.Label lbl����r�� 
         Appearance      =   0  '����
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "����r��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   528
         Index           =   1
         Left            =   1296
         TabIndex        =   15
         Top             =   252
         Width           =   540
      End
      Begin VB.Label lbl�r�ε��c 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�r�ε��c"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Left            =   852
         TabIndex        =   10
         Top             =   312
         Width           =   360
      End
      Begin VB.Label lbl�r�δF�� 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�����˦r"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   9
         Top             =   288
         Width           =   276
      End
      Begin VB.Shape shp�r�δF�� 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  '���z��
         Height          =   1296
         Index           =   1
         Left            =   132
         Top             =   156
         Width           =   580
      End
      Begin VB.Shape shp�r�ε��c 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  '���z��
         Height          =   1296
         Index           =   1
         Left            =   698
         Top             =   156
         Width           =   580
      End
      Begin VB.Shape shp����r�� 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  '���z��
         Height          =   660
         Index           =   1
         Left            =   1260
         Top             =   156
         Width           =   576
      End
      Begin VB.Shape shp�r�κt�� 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  '���z��
         Height          =   648
         Index           =   1
         Left            =   1260
         Top             =   800
         Width           =   576
      End
   End
   Begin VB.PictureBox pic�w�]�s�� 
      Height          =   1620
      Index           =   3
      Left            =   3168
      ScaleHeight     =   1572
      ScaleWidth      =   1968
      TabIndex        =   2
      Top             =   2484
      Width           =   2016
      Begin VB.Label lbl����r�� 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "����r��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   492
         Index           =   3
         Left            =   1296
         TabIndex        =   7
         Top             =   880
         Width           =   504
      End
      Begin VB.Label lbl�r�ε��c 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�r�ε��c"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   516
         Index           =   3
         Left            =   744
         TabIndex        =   5
         Top             =   240
         Width           =   480
      End
      Begin VB.Shape shp�r�ε��c 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  '���z��
         Height          =   660
         Index           =   3
         Left            =   698
         Top             =   156
         Width           =   576
      End
      Begin VB.Label lbl����r�� 
         Appearance      =   0  '����
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '�z��
         Caption         =   "����r��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   528
         Index           =   3
         Left            =   744
         TabIndex        =   4
         Top             =   888
         Width           =   540
      End
      Begin VB.Label lbl�r�δF�� 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "�����˦r"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         TabIndex        =   3
         Top             =   288
         Width           =   276
      End
      Begin VB.Shape shp�r�δF�� 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  '���z��
         Height          =   1296
         Index           =   3
         Left            =   132
         Top             =   156
         Width           =   580
      End
      Begin VB.Shape shp����r�� 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  '���z��
         Height          =   650
         Index           =   3
         Left            =   698
         Top             =   800
         Width           =   576
      End
      Begin VB.Label lbl�c�r���� 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         BackStyle       =   0  '�z��
         Caption         =   "����Ÿ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   528
         Index           =   3
         Left            =   1300
         TabIndex        =   6
         Top             =   240
         Width           =   480
      End
      Begin VB.Shape shp�c�r���� 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  '���z��
         Height          =   660
         Index           =   3
         Left            =   1260
         Top             =   156
         Width           =   576
      End
      Begin VB.Shape shp����r�� 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  '���z��
         Height          =   650
         Index           =   3
         Left            =   1260
         Top             =   800
         Width           =   576
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "����"
      Height          =   375
      Left            =   5712
      TabIndex        =   1
      Top             =   984
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "�T�w"
      Height          =   375
      Left            =   5700
      TabIndex        =   0
      Top             =   444
      Width           =   1215
   End
   Begin VB.Shape shp�w�]�s�� 
      BorderWidth     =   2
      Height          =   2076
      Index           =   1
      Left            =   2892
      Top             =   216
      Visible         =   0   'False
      Width           =   2592
   End
   Begin VB.Shape shp�r�ε��c1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  '���z��
      Height          =   660
      Index           =   2
      Left            =   4428
      Top             =   576
      Width           =   636
   End
   Begin VB.Shape shp�r�κt��1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  '���z��
      Height          =   648
      Index           =   2
      Left            =   4440
      Top             =   1236
      Width           =   624
   End
   Begin VB.Shape shp�r�κt��1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  '���z��
      Height          =   648
      Index           =   0
      Left            =   4452
      Top             =   3300
      Width           =   624
   End
   Begin VB.Shape shp�r�ε��c1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  '���z��
      Height          =   660
      Index           =   3
      Left            =   4440
      Top             =   2640
      Width           =   636
   End
   Begin VB.Shape shp�w�]�s�� 
      BorderWidth     =   2
      Height          =   2076
      Index           =   3
      Left            =   2904
      Top             =   2280
      Visible         =   0   'False
      Width           =   2592
   End
   Begin VB.Shape shp�w�]�s�� 
      BorderWidth     =   2
      Height          =   2076
      Index           =   2
      Left            =   336
      Top             =   2292
      Visible         =   0   'False
      Width           =   2592
   End
   Begin VB.Shape shp�w�]�s�� 
      BorderWidth     =   2
      Height          =   2076
      Index           =   0
      Left            =   336
      Top             =   240
      Visible         =   0   'False
      Width           =   2592
   End
End
Attribute VB_Name = "frm�w�]�s��"
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

For i = 0 To �w�]�s���`�� - 1
    If i = �w�]�s���Ҧ� - 1 And Not ²���s���Ҧ� Then
        shp�w�]�s��(i).Visible = True
    Else
        shp�w�]�s��(i).Visible = False
    End If
Next i

���ܹw�]�s�� = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

�p��{�ε���

End Sub

Private Sub lbl�r�δF��_Click(Index As Integer)

Dim i As Integer

For i = 0 To �w�]�s���`�� - 1
    If i = Index Then
        shp�w�]�s��(i).Visible = True
    Else
        shp�w�]�s��(i).Visible = False
    End If
Next i

End Sub

Private Sub lbl�r�ε��c_Click(Index As Integer)

Dim i As Integer

For i = 0 To �w�]�s���`�� - 1
    If i = Index Then
        shp�w�]�s��(i).Visible = True
    Else
        shp�w�]�s��(i).Visible = False
    End If
Next i

End Sub

Private Sub lbl�r�κt��_Click(Index As Integer)

Dim i As Integer

For i = 0 To �w�]�s���`�� - 1
    If i = Index Then
        shp�w�]�s��(i).Visible = True
    Else
        shp�w�]�s��(i).Visible = False
    End If
Next i

End Sub

Private Sub lbl����r��_Click(Index As Integer)

Dim i As Integer

For i = 0 To �w�]�s���`�� - 1
    If i = Index Then
        shp�w�]�s��(i).Visible = True
    Else
        shp�w�]�s��(i).Visible = False
    End If
Next i

End Sub

Private Sub lbl����r��_Click(Index As Integer)

Dim i As Integer

For i = 0 To �w�]�s���`�� - 1
    If i = Index Then
        shp�w�]�s��(i).Visible = True
    Else
        shp�w�]�s��(i).Visible = False
    End If
Next i

End Sub

Private Sub lbl�r�ί���_Click(Index As Integer)

Dim i As Integer

For i = 0 To �w�]�s���`�� - 1
    If i = Index Then
        shp�w�]�s��(i).Visible = True
    Else
        shp�w�]�s��(i).Visible = False
    End If
Next i

End Sub

Private Sub lbl�c�r����_Click(Index As Integer)

Dim i As Integer

For i = 0 To �w�]�s���`�� - 1
    If i = Index Then
        shp�w�]�s��(i).Visible = True
    Else
        shp�w�]�s��(i).Visible = False
    End If
Next i

End Sub

Private Sub OKButton_Click()

Dim i As Integer

For i = 0 To �w�]�s���`�� - 1
    If shp�w�]�s��(i).Visible Then
        �w�]�s���Ҧ� = i + 1
    End If
Next i

²���s���Ҧ� = False
���ܹw�]�s�� = True

Unload Me

End Sub

Private Sub pic�w�]�s��_Click(Index As Integer)

Dim i As Integer

For i = 0 To �w�]�s���`�� - 1
    If i = Index Then
        shp�w�]�s��(i).Visible = True
    Else
        shp�w�]�s��(i).Visible = False
    End If
Next i

End Sub
