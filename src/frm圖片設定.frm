VERSION 5.00
Begin VB.Form frm�Ϥ��]�w 
   Caption         =   "�]�w�ƻs�Ϥ�"
   ClientHeight    =   4140
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   5460
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.ListBox lst�ۭq 
      Height          =   1848
      ItemData        =   "frm�Ϥ��]�w.frx":0000
      Left            =   1128
      List            =   "frm�Ϥ��]�w.frx":0025
      TabIndex        =   10
      Top             =   2028
      Width           =   912
   End
   Begin VB.TextBox txt�ۭq 
      Height          =   264
      HideSelection   =   0   'False
      Left            =   1128
      TabIndex        =   9
      Top             =   1620
      Width           =   876
   End
   Begin VB.OptionButton opt�ۭq 
      Caption         =   "�ۭq"
      Height          =   384
      Left            =   312
      TabIndex        =   8
      Top             =   1560
      Width           =   1296
   End
   Begin VB.OptionButton opt�L��� 
      Caption         =   "�L���    600"
      Height          =   384
      Left            =   312
      TabIndex        =   7
      Top             =   1044
      Width           =   1296
   End
   Begin VB.OptionButton opt��ܾ� 
      Caption         =   "��ܾ�    120"
      Height          =   384
      Left            =   312
      TabIndex        =   6
      Top             =   528
      Width           =   1296
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "����"
      Height          =   375
      Left            =   4284
      TabIndex        =   3
      Top             =   996
      Width           =   864
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "�T�w"
      Default         =   -1  'True
      Height          =   375
      Left            =   4284
      TabIndex        =   2
      Top             =   468
      Width           =   864
   End
   Begin VB.ListBox lstFontsize 
      Height          =   3108
      ItemData        =   "frm�Ϥ��]�w.frx":0063
      Left            =   2904
      List            =   "frm�Ϥ��]�w.frx":00B1
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
      Caption         =   "�ѪR��(����/�^�T)"
      Height          =   276
      Left            =   324
      TabIndex        =   5
      Top             =   156
      Width           =   1536
   End
   Begin VB.Label lblfontsize 
      Caption         =   "�j�p"
      Height          =   252
      Left            =   2904
      TabIndex        =   4
      Top             =   156
      Width           =   480
   End
End
Attribute VB_Name = "frm�Ϥ��]�w"
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

txtFontsize = �Ϥ��r���j�p
If �Ϥ��ѪR�� = 120 Then
    opt��ܾ�.Value = True
    txt�ۭq = �Ϥ��ѪR��
ElseIf �Ϥ��ѪR�� = 600 Then
    opt�L���.Value = True
    txt�ۭq = �Ϥ��ѪR��
Else
    opt�ۭq.Value = True
    txt�ۭq = �Ϥ��ѪR��
End If

End Sub


Private Sub lstFontsize_Click()

If lstFontsize.ListIndex > -1 Then
    txtFontsize.Text = lstFontsize.List(lstFontsize.ListIndex)
    txtFontsize.SelStart = 0
    txtFontsize.SelLength = Len(txtFontsize.Text)
End If


End Sub


Private Sub lst�ۭq_Click()

If lst�ۭq.ListIndex > -1 Then
    txt�ۭq = lst�ۭq.List(lst�ۭq.ListIndex)
    txt�ۭq.SelStart = 0
    txt�ۭq.SelLength = Len(txt�ۭq.Text)
    opt�ۭq.Value = True
End If

End Sub

Private Sub OKButton_Click()

Dim �r��j�p As Integer


�r��j�p = Val(txtFontsize)
If �r��j�p > 0 Then �Ϥ��r���j�p = �r��j�p
If opt��ܾ� Then
    �Ϥ��ѪR�� = 120
ElseIf opt�L��� Then
    �Ϥ��ѪR�� = 600
ElseIf opt�ۭq Then
    If Val(txt�ۭq) > 0 Then �Ϥ��ѪR�� = Val(txt�ۭq)
End If

mdi�~�r�r��.cbo�Ϥ��j�p = �Ϥ��r���j�p
mdi�~�r�r��.cbo�ѪR�� = �Ϥ��ѪR��

Unload Me

End Sub

