VERSION 5.00
Begin VB.Form frm�r�� 
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "�r��"
   ClientHeight    =   4176
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   5472
   Icon            =   "frm�r��.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4176
   ScaleWidth      =   5472
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.TextBox txt����r 
      Enabled         =   0   'False
      Height          =   264
      Left            =   900
      TabIndex        =   15
      Text            =   "����|���t²����r"
      Top             =   3540
      Width           =   1788
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Left            =   900
      TabIndex        =   13
      Text            =   "����|�Ұ���"
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
      ItemData        =   "frm�r��.frx":030A
      Left            =   2964
      List            =   "frm�r��.frx":0358
      TabIndex        =   11
      Top             =   972
      Width           =   912
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "�T�w"
      Default         =   -1  'True
      Height          =   375
      Left            =   4344
      TabIndex        =   10
      Top             =   576
      Width           =   864
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "����"
      Height          =   375
      Left            =   4344
      TabIndex        =   9
      Top             =   1104
      Width           =   864
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   264
      Left            =   900
      TabIndex        =   8
      Text            =   "����|����"
      Top             =   2556
      Width           =   1788
   End
   Begin VB.TextBox txt�p�f 
      Enabled         =   0   'False
      Height          =   264
      Left            =   900
      TabIndex        =   7
      Text            =   "�_�v�j����p�f"
      Top             =   2064
      Width           =   1788
   End
   Begin VB.ListBox lst���� 
      Height          =   768
      ItemData        =   "frm�r��.frx":03B0
      Left            =   900
      List            =   "frm�r��.frx":03BA
      TabIndex        =   6
      Top             =   948
      Width           =   1788
   End
   Begin VB.TextBox txt���� 
      Height          =   264
      HideSelection   =   0   'False
      Left            =   900
      TabIndex        =   5
      Top             =   588
      Width           =   1788
   End
   Begin VB.Label lbl����r 
      Caption         =   "����r"
      Height          =   252
      Left            =   156
      TabIndex        =   16
      Top             =   3540
      Width           =   588
   End
   Begin VB.Label lbl�Ұ��� 
      Caption         =   "�Ұ���"
      Height          =   252
      Left            =   156
      TabIndex        =   14
      Top             =   3048
      Width           =   588
   End
   Begin VB.Label lblfontsize 
      Caption         =   "�j�p"
      Height          =   252
      Left            =   2964
      TabIndex        =   4
      Top             =   264
      Width           =   480
   End
   Begin VB.Label lbl���� 
      Caption         =   "����"
      Height          =   252
      Left            =   156
      TabIndex        =   3
      Top             =   2556
      Width           =   480
   End
   Begin VB.Label lbl�p�f 
      Caption         =   "�p�f"
      Height          =   252
      Left            =   156
      TabIndex        =   2
      Top             =   2064
      Width           =   480
   End
   Begin VB.Label lbl���� 
      Caption         =   "����"
      Height          =   252
      Left            =   180
      TabIndex        =   1
      Top             =   594
      Width           =   480
   End
   Begin VB.Label lblFontName 
      Caption         =   "�r��"
      Height          =   276
      Left            =   900
      TabIndex        =   0
      Top             =   264
      Width           =   588
   End
End
Attribute VB_Name = "frm�r��"
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

txt����.Text = ��ܦr��
txt����.SelStart = 0
txt����.SelLength = Len(txt����.Text)

txtFontsize.Text = ��ܦr���j�p
txtFontsize.SelStart = 0
txtFontsize.SelLength = Len(txt����.Text)

For i = 0 To lst����.ListCount - 1
    If lst����.List(i) = ��ܦr�� Then
        lst����.ListIndex = i
        Exit For
    End If
Next i

For i = 0 To lstFontsize.ListCount - 1
    If lstFontsize.List(i) = ��ܦr���j�p Then
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


Private Sub lst����_Click()

If lst����.ListIndex > -1 Then
    txt����.Text = lst����.List(lst����.ListIndex)
    txt����.SelStart = 0
    txt����.SelLength = Len(txt����.Text)
End If

End Sub

Private Sub OKButton_Click()

Dim �r��j�p As Integer

If txt����.Text = "�ө���" Or txt����.Text = "�з���" Then
    If txt����.Text <> ��ܦr�� Then
        mdi�~�r�r��.cbo�r���W��.Text = txt����.Text
    End If
End If

�r��j�p = Val(txtFontsize.Text)
If �r��j�p <> ��ܦr���j�p Then
    If �r��j�p < 10 Then
        �r��j�p = 10
    'ElseIf �r��j�p > 72 Then
    '    �r��j�p = 72
    End If
    mdi�~�r�r��.cbo�r��j�p.Text = �r��j�p
End If

Unload Me

End Sub

