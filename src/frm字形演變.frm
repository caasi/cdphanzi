VERSION 5.00
Object = "{65996203-3B87-11D4-A21F-00E029189826}#6.9#0"; "TLIST6.OCX"
Begin VB.Form frm�r�κt�� 
   Caption         =   "�r�κt��"
   ClientHeight    =   6228
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9060
   Icon            =   "frm�r�κt��.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6228
   ScaleWidth      =   9060
   Begin TListProLibCtl.TList tree�r�ξ𪬵��c 
      DragIcon        =   "frm�r�κt��.frx":030A
      Height          =   2652
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   3852
      _Version        =   393225
      _ExtentX        =   6800
      _ExtentY        =   4683
      _StockProps     =   228
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�з���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      Appearance      =   1
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      SelForeColor    =   -2147483634
      SelBackColor    =   -2147483635
      ShiftStep       =   300
      ItemImageDefHeight=   228
      ItemImageDefWidth=   228
      WidthOfText     =   0
      TabStopDistance =   0
      MarkHeight      =   0
      MarkWidth       =   0
      TitleHeight     =   0
      XOffset         =   0
      TriggerEvents   =   0
      PathSeparator   =   "\"
      Caption         =   ""
      FixedSize       =   0   'False
      ShowChildren    =   0   'False
      ExpandChildren  =   0   'False
      ExpandNewItem   =   0   'False
      Scrollbars      =   3
      PictureOpen     =   "frm�r�κt��.frx":074C
      PictureClosed   =   "frm�r�κt��.frx":085E
      PictureLeaf     =   "frm�r�κt��.frx":0970
      PictureMark     =   "frm�r�κt��.frx":0A82
      ImageStretch    =   0   'False
      NoIntegralHeight=   -1  'True
      DisableNoScroll =   0   'False
      NoPictureRoot   =   0   'False
      MSOutlineAdd    =   0   'False
      BackwardCompatibility=   0   'False
      InvStyle        =   0
      ViewStyle       =   0
      PictureType     =   0
      CurrentIndexMethod=   0
      ViewStyleEx     =   1
      AutoExpand      =   1
      TreeLinesStyle  =   0
      PicInMultiLine  =   0
      ShowCaption     =   0
      ShowTitles      =   0
      AutoScrDuringDragDrop=   0
      DragHighlight   =   0
      MousePointer    =   0
      DefMultiLine    =   0   'False
      SmartDragDrop   =   0   'False
      WidthOfTextMin  =   0
      DrawFocusRect   =   0   'False
      LcPresent       =   -1  'True
      WebTargetFrame  =   ""
      WebURLBase      =   ""
      GradientStyle   =   0
      TransparentBackground=   0   'False
      DefBorderStyle  =   0
      DefPictureAlignment=   5
      DefAlignment    =   0
      DefTextAlignment=   2
      ShowHiddenItems =   0   'False
      DefItemCellBackColor=   583057600
      _InternalVersion=   524290
      ExchangeSerialNumber=   "frm�r�κt��.frx":0B7C
      DragIconStyle   =   0
      ExchangeDefItemCellDef=   "frm�r�κt��.frx":0BC9
      _ChkCounter     =   -1
      TreeLinesHighlightColor=   -2113929196
      TreeLinesShadowColor=   -2113929200
   End
   Begin VB.Image imgPinPush 
      Height          =   264
      Left            =   4440
      Picture         =   "frm�r�κt��.frx":0CD0
      Top             =   240
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imgPin 
      Height          =   264
      Left            =   4440
      Picture         =   "frm�r�κt��.frx":0E5A
      Top             =   720
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imglock 
      Height          =   264
      Left            =   0
      Picture         =   "frm�r�κt��.frx":0FE4
      Tag             =   "0"
      ToolTipText     =   "��w"
      Top             =   240
      Width           =   288
   End
End
Attribute VB_Name = "frm�r�κt��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private �����N�X As Integer, ���� As String
Private �ϰ�r��}�C(0 To �r��Ӽ�) As Variant
Private XCheck As Single, YCheck As Single

Private Sub Form_Activate()

�{�ε��� = ����
'�{�ε����N�X = �����N�X
�{�ε����N�X = �r�κt�ܥN�X
��������r�Τu��C���A �{�ε����N�X
tree�r�ξ𪬵��c_Click
mdi�~�r�r��.txt���A = ���c���A�C

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim �r�ڧ� As String, �s�� As Long

�Ұʦr�κt�� = True
If ��lfirst <> 1 Then
   If �w���J�e�� = 0 Then
      If �t��winstate = 0 Then
         frm�r�κt��.Left = �t��left
         frm�r�κt��.Top = �t��top
         frm�r�κt��.Height = �t��height
         frm�r�κt��.Width = �t��width
      Else
         frm�r�κt��.WindowState = �t��winstate
      End If
   ElseIf �Ұʦr�δF�� Then
         frm�r�κt��.Left = frm�r�δF��.Left + frm�r�δF��.Width
         frm�r�κt��.Top = frm�r�δF��.Top
         frm�r�κt��.Height = frm�r�δF��.Height
         frm�r�κt��.Width = frm�r�δF��.Width
   End If
End If

tree�r�ξ𪬵��c.FontSize = CInt(��ܦr���j�p)
'Me.Tag = �@�ε����N�X
Me.Tag = �r�κt�ܥN�X
�����N�X = �@�ε����N�X
���� = �@�ε���(�@�ε����N�X)
tree�r�ξ𪬵��c.AddItem ""
'tree�r�ξ𪬵��c.ListIndex = 0
tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureLeaf

i = 0
Do While �r��}�C(i) <> ""
   �ϰ�r��}�C(i) = �r��}�C(i)
   i = i + 1
Loop

If �O�_���զr�Ÿ�(mdi�~�r�r��.txt�r��.Text, 1, 14) = 0 Then
   'If Mid(mdi�~�r�r��.txt�r��.FontName, 1, 3) <> "hzk" Then
   '   �˦r��.Index = "�r��"
   '   �˦r��.Seek "=", mdi�~�r�r��.txt�r��.Text
   'Else
      �����˦r��.Index = "�s��"
      �����˦r��.Seek "=", �t�νs��
   'End If
   
   If Not �����˦r��.NoMatch Then
      Do While Not �����˦r��.NoMatch
         �s�� = �����˦r��.Fields("�s��")
         Exit Do
         �����˦r��.MoveNext
       Loop
    End If
    ���J�r�� "�з���", mdi�~�r�r��.txt�r��.Text, �s��
End If

End Sub

Private Sub Form_Resize()

If Me.ScaleHeight - tree�r�ξ𪬵��c.Top * 2 > 0 Then tree�r�ξ𪬵��c.Height = Me.ScaleHeight - tree�r�ξ𪬵��c.Top * 2
If Me.ScaleWidth - tree�r�ξ𪬵��c.Left * 2 > 0 Then tree�r�ξ𪬵��c.Width = Me.ScaleWidth - tree�r�ξ𪬵��c.Left * 2

End Sub


Private Sub Form_Unload(Cancel As Integer)
mdi�~�r�r��.mnu_�r�κt��.Enabled = True
�Ұʦr�κt�� = False
�p��{�ε���

End Sub

Public Sub ���J�r��(�t�Φr�� As String, �r�� As String, �s�� As Long)
Dim ����� As String, �r��s�� As Integer, �r���� As String, ���� As Integer
Dim i As Integer, �r�� As String
Dim ���ѽs�� As Long, �p�f�s�� As Integer, ����s�� As Long, �Ұ���s�� As Integer, ���t��r�s�� As Long
Dim ���Ѧr�� As String, �p�f�r�� As String, ����r�� As String, �Ұ���r�� As String, ���t��r�r�� As String
Dim �p�f�r�� As String, ����r�� As String, �Ұ���r�� As String, ���t��r�r�� As String
Dim �Ұ������ As Long, ������� As Long, ���t��r���� As Long, �p�f���� As Long
Dim �������� As String, ���W�ʦr As Boolean, RTF�ʦr As String
Dim ����r�� As String, �Ұ���r�� As String, ���t��r�r�� As String

If imglock.Tag = 1 Then Exit Sub

If �s�� <= 0 Then Exit Sub

Screen.MousePointer = ccHourglass

If �t�Φr�� = "�_�v�j����p�f" Or �t�Φr�� = "�_�v�j���孫��" Then
    �p�f�s�� = �s��
    �p�f�˦r��.Index = "�s��"
    �p�f�˦r��.Seek "=", �s��
    ���ѽs�� = �p�f�˦r��.Fields("���ѽs��")
ElseIf ����|����(�t�Φr��) Then
    ����s�� = �s��
    If �t�Φr�� <> "����|����" Then
        ���岧�g�r��.Seek "=", �s��
        ���ѽs�� = ���岧�g�r��.Fields("���ѽs��")
    Else
        �����˦r��.Index = "�s��"
        �����˦r��.Seek "=", �s��
        ���ѽs�� = �����˦r��.Fields("���ѽs��")
    End If
ElseIf ����|�Ұ���(�t�Φr��) Then
    �Ұ���s�� = �s��
    If �t�Φr�� <> "����|�Ұ���" Then
        �Ұ��岧�g�r��.Seek "=", �s��
        ���ѽs�� = �Ұ��岧�g�r��.Fields("���ѽs��")
    Else
        �Ұ����˦r��.Index = "�s��"
        �Ұ����˦r��.Seek "=", �s��
        ���ѽs�� = �Ұ����˦r��.Fields("���ѽs��")
    End If
ElseIf ����|���t��r(�t�Φr��) Then
    ���t��r�s�� = �s��
    If �t�Φr�� <> "����|���t²����r" Then
        ���t��r���g�r��.Seek "=", �s��
        ���ѽs�� = ���t��r���g�r��.Fields("���ѽs��")
    Else
        ���t��r�˦r��.Index = "�s��"
        ���t��r�˦r��.Seek "=", �s��
        ���ѽs�� = ���t��r�˦r��.Fields("���ѽs��")
    End If
Else
    ���ѽs�� = �s��
End If

�����˦r��.Index = "�s��"
�����˦r��.Seek "=", ���ѽs��
���Ѧr�� = �����˦r��.Fields("�r�X")
�r��s�� = �����˦r��.Fields("�r��")
�r�� = �ϰ�r��}�C(�r��s��)

If �t�Φr�� <> "�_�v�j����p�f" And �t�Φr�� <> "�_�v�j���孫��" Then
    If Not IsNull(�����˦r��.Fields("�p�f�s��")) Then
        �p�f�s�� = �����˦r��.Fields("�p�f�s��")
    Else
        �p�f�s�� = 0
    End If
End If

If Not ����|����(�t�Φr��) Then
    If Not IsNull(�����˦r��.Fields("����s��")) Then
        ����s�� = �����˦r��.Fields("����s��")
    Else
        ����s�� = 0
    End If
End If

If Not ����|�Ұ���(�t�Φr��) Then
    If Not IsNull(�����˦r��.Fields("�Ұ���s��")) Then
        �Ұ���s�� = �����˦r��.Fields("�Ұ���s��")
    Else
        �Ұ���s�� = 0
    End If
End If

If Not ����|���t��r(�t�Φr��) Then
    If Not IsNull(�����˦r��.Fields("���t��r�s��")) Then
        ���t��r�s�� = �����˦r��.Fields("���t��r�s��")
    Else
        ���t��r�s�� = 0
    End If
End If

�Ұ������ = -1
������� = -1
���t��r���� = -1
�p�f���� = -1

tree�r�ξ𪬵��c.Clear
tree�r�ξ𪬵��c.Redraw = False
tree�r�ξ𪬵��c.AddItem ���Ѧr��
tree�r�ξ𪬵��c.ItemFontName(0) = �ഫ��ܦr��(�r��)
tree�r�ξ𪬵��c.ItemLngValue(0) = ���ѽs��
If Not IsNull(�˦r��.Fields("�r��")) Then
    tree�r�ξ𪬵��c.ItemTag(0) = �r�θ`�I�аO
Else
    tree�r�ξ𪬵��c.ItemTag(0) = �c�r���`�I�аO
End If


If (�p�f�s�� = 0 Or Not mdi�~�r�r��.mnu_�p�f�ﶵ.Checked) And (����s�� = 0 Or Not mdi�~�r�r��.mnu_����ﶵ.Checked) And (�Ұ���s�� = 0 Or Not mdi�~�r�r��.mnu_�Ұ���ﶵ.Checked) And (���t��r�s�� = 0 Or Not mdi�~�r�r��.mnu_���t��r�ﶵ.Checked) Then
    tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureLeaf
    tree�r�ξ𪬵��c.Redraw = True
    Screen.MousePointer = ccDefault
    Exit Sub
End If

tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureOpen

If �Ұ���s�� <> 0 And mdi�~�r�r��.mnu_�Ұ���ﶵ.Checked Then
    �Ұ����˦r��.Index = "�s��"
    �Ұ����˦r��.Seek "=", �Ұ���s��
    
    If Not �Ұ����˦r��.NoMatch Then
        �Ұ���r�� = �Ұ����˦r��.Fields("�r�X")
        �Ұ���r�� = "����|�Ұ���"
        �Ұ��岧�g�r��.Seek "=", �Ұ���s��
    Else
        �Ұ��岧�g�r��.Seek "=", �Ұ���s��
        �Ұ���r�� = �Ұ��岧�g�r��.Fields("�r�X")
        �Ұ���r�� = �t�Φr��
    End If

    If Not IsNull(�Ұ��岧�g�r��.Fields("�X�B")) Then
        �Ұ���r�� = �Ұ��岧�g�r��.Fields("�X�B")
        If IsNumeric(Left(�Ұ���r��, 1)) Then �Ұ���r�� = "�X��" & �Ұ���r��
    Else
        �Ұ���r�� = ""
    End If
    
    �Ұ������ = tree�r�ξ𪬵��c.ListCount
    tree�r�ξ𪬵��c.AddItem �Ұ���r��, 0
    tree�r�ξ𪬵��c.ItemFontName(�Ұ������) = �Ұ���r��
    tree�r�ξ𪬵��c.ItemLngValue(�Ұ������) = �Ұ���s��
    If Not IsNull(�Ұ����˦r��.Fields("�r��")) Then
        tree�r�ξ𪬵��c.ItemTag(�Ұ������) = �r�θ`�I�аO
    Else
        tree�r�ξ𪬵��c.ItemTag(�Ұ������) = �c�r���`�I�аO
    End If

    If Len(�Ұ���r��) = 0 Then
        tree�r�ξ𪬵��c.Image(�Ұ������) = tree�r�ξ𪬵��c.PictureLeaf
    Else
        tree�r�ξ𪬵��c.Image(�Ұ������) = tree�r�ξ𪬵��c.PictureOpen
        tree�r�ξ𪬵��c.AddItem �Ұ���r��, �Ұ������
        tree�r�ξ𪬵��c.ItemFontName(�Ұ������ + 1) = "�з���"
        tree�r�ξ𪬵��c.ItemTag(�Ұ������ + 1) = ��L�`�I�аO
        tree�r�ξ𪬵��c.ItemFontSize(�Ұ������ + 1) = 12
        tree�r�ξ𪬵��c.ItemLngValue(�Ұ������ + 1) = -9999
        tree�r�ξ𪬵��c.Image(�Ұ������ + 1) = tree�r�ξ𪬵��c.PictureLeaf
    End If
End If

If ����s�� <> 0 And mdi�~�r�r��.mnu_����ﶵ.Checked Then
    �����˦r��.Index = "�s��"
    �����˦r��.Seek "=", ����s��
    If Not �����˦r��.NoMatch Then
        ����r�� = �����˦r��.Fields("�r�X")
        ����r�� = "����|����"
        ���岧�g�r��.Seek "=", ����s��
        �������� = ���岧�g�r��.Fields("����")
    Else
        ���岧�g�r��.Seek "=", ����s��
        ����r�� = ���岧�g�r��.Fields("�r�X")
        ����r�� = �t�Φr��
        �������� = ���岧�g�r��.Fields("����")
    End If
    
    ���嶰�����W.Seek "=", ��������
    If �������� <> 15000 Then ����r�� = "����" & ��������
    If Not ���嶰�����W.NoMatch Then
        ����r�� = ����r�� & "(" & ���嶰�����W.Fields("���W") & ")"
        If ���嶰�����W.Fields("�ʦr") = 1 Then
            ���W�ʦr = True
        Else
            ���W�ʦr = False
        End If
    Else
        ����r�� = ""
    End If
    
    ������� = tree�r�ξ𪬵��c.ListCount
    tree�r�ξ𪬵��c.AddItem ����r��, 0
    tree�r�ξ𪬵��c.ItemFontName(�������) = ����r��
    tree�r�ξ𪬵��c.ItemLngValue(�������) = ����s��
    If Not IsNull(�����˦r��.Fields("�r��")) Then
        tree�r�ξ𪬵��c.ItemTag(�������) = �r�θ`�I�аO
    Else
        tree�r�ξ𪬵��c.ItemTag(�������) = �c�r���`�I�аO
    End If

    If Len(����r��) = 0 Then
        tree�r�ξ𪬵��c.Image(�������) = tree�r�ξ𪬵��c.PictureLeaf
    Else
        tree�r�ξ𪬵��c.Image(�������) = tree�r�ξ𪬵��c.PictureOpen
        If Not ���W�ʦr Then
            tree�r�ξ𪬵��c.AddItem ����r��, �������
            tree�r�ξ𪬵��c.ItemFontName(������� + 1) = "�з���"
            tree�r�ξ𪬵��c.ItemTag(������� + 1) = ��L�`�I�аO
        Else
            RTF�ʦr = �ഫRTF�ʦr(����r��, ��ܦr��)
            tree�r�ξ𪬵��c.AddItem RTF�ʦr, �������
            tree�r�ξ𪬵��c.ItemCell(������� + 1).RTFStyle = 1
            tree�r�ξ𪬵��c.ItemTag(������� + 1) = ���W�`�I�аO & ����r��
        End If
        tree�r�ξ𪬵��c.ItemFontSize(������� + 1) = 12
        tree�r�ξ𪬵��c.ItemLngValue(������� + 1) = -9999
        tree�r�ξ𪬵��c.Image(������� + 1) = tree�r�ξ𪬵��c.PictureLeaf
    End If
End If

If ���t��r�s�� <> 0 And mdi�~�r�r��.mnu_���t��r�ﶵ.Checked Then
    ���t��r�˦r��.Index = "�s��"
    ���t��r�˦r��.Seek "=", ���t��r�s��
    
    If Not ���t��r�˦r��.NoMatch Then
        ���t��r�r�� = ���t��r�˦r��.Fields("�r�X")
        ���t��r�r�� = "����|���t²����r"
        ���t��r���g�r��.Seek "=", ���t��r�s��
    Else
        ���t��r���g�r��.Seek "=", ���t��r�s��
        ���t��r�r�� = ���t��r���g�r��.Fields("�r�X")
        ���t��r�r�� = �t�Φr��
    End If
    
    If Not IsNull(���t��r���g�r��.Fields("�X�B")) Then
        ���t��r�r�� = ���t��r���g�r��.Fields("�X�B")
    Else
        ���t��r�r�� = ""
    End If
    
    ���t��r���� = tree�r�ξ𪬵��c.ListCount
    tree�r�ξ𪬵��c.AddItem ���t��r�r��, 0
    tree�r�ξ𪬵��c.ItemFontName(���t��r����) = ���t��r�r��
    tree�r�ξ𪬵��c.ItemLngValue(���t��r����) = ���t��r�s��
    If Not IsNull(���t��r�˦r��.Fields("�r��")) Then
        tree�r�ξ𪬵��c.ItemTag(���t��r����) = �r�θ`�I�аO
    Else
        tree�r�ξ𪬵��c.ItemTag(���t��r����) = �c�r���`�I�аO
    End If

    If Len(���t��r�r��) = 0 Then
        tree�r�ξ𪬵��c.Image(���t��r����) = tree�r�ξ𪬵��c.PictureLeaf
    Else
        tree�r�ξ𪬵��c.Image(���t��r����) = tree�r�ξ𪬵��c.PictureOpen
        tree�r�ξ𪬵��c.AddItem ���t��r�r��, ���t��r����
        tree�r�ξ𪬵��c.ItemFontName(���t��r���� + 1) = "�з���"
        tree�r�ξ𪬵��c.ItemTag(���t��r���� + 1) = ��L�`�I�аO
        tree�r�ξ𪬵��c.ItemFontSize(���t��r���� + 1) = 12
        tree�r�ξ𪬵��c.ItemLngValue(���t��r���� + 1) = -9999
        tree�r�ξ𪬵��c.Image(���t��r���� + 1) = tree�r�ξ𪬵��c.PictureLeaf
    End If
End If

If �p�f�s�� <> 0 And mdi�~�r�r��.mnu_�p�f�ﶵ.Checked Then
    �p�f�˦r��.Index = "�s��"
    �p�f�˦r��.Seek "=", �p�f�s��
    �p�f�r�� = �p�f�˦r��.Fields("�r�X")
    If Not IsNull(�p�f�˦r��.Fields("�r��")) Then
        �p�f�r�� = �p�f�˦r��.Fields("�r��")
    Else
        �p�f�r�� = ""
    End If

    �p�f���� = tree�r�ξ𪬵��c.ListCount
    
    tree�r�ξ𪬵��c.AddItem �p�f�r��, 0
    If �t�Φr�� = "�_�v�j���孫��" Then
        tree�r�ξ𪬵��c.ItemFontName(�p�f����) = "�_�v�j���孫��"
    Else
        tree�r�ξ𪬵��c.ItemFontName(�p�f����) = "�_�v�j����p�f"
    End If
    tree�r�ξ𪬵��c.ItemLngValue(�p�f����) = �p�f�s��

    If Not IsNull(�p�f�˦r��.Fields("�r��")) Then
        tree�r�ξ𪬵��c.ItemTag(�p�f����) = �r�θ`�I�аO
    Else
        tree�r�ξ𪬵��c.ItemTag(�p�f����) = �c�r���`�I�аO
    End If

    If Len(�p�f�r��) = 0 Then
        tree�r�ξ𪬵��c.Image(�p�f����) = tree�r�ξ𪬵��c.PictureLeaf
    Else
        tree�r�ξ𪬵��c.Image(�p�f����) = tree�r�ξ𪬵��c.PictureOpen
        tree�r�ξ𪬵��c.AddItem �p�f�r��, �p�f����
        tree�r�ξ𪬵��c.ItemFontName(�p�f���� + 1) = �ഫ��ܦr��("�з���")
        tree�r�ξ𪬵��c.ItemFontSize(�p�f���� + 1) = 12
        tree�r�ξ𪬵��c.ItemLngValue(�p�f���� + 1) = -9999
        tree�r�ξ𪬵��c.ItemTag(�p�f���� + 1) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(�p�f���� + 1) = tree�r�ξ𪬵��c.PictureLeaf
    End If
End If

tree�r�ξ𪬵��c.Expand(0) = True
If �Ұ������ > -1 Then tree�r�ξ𪬵��c.Expand(�Ұ������) = True
If ������� > -1 Then tree�r�ξ𪬵��c.Expand(�������) = True
If ���t��r���� > -1 Then tree�r�ξ𪬵��c.Expand(���t��r����) = True
If �p�f���� > -1 Then tree�r�ξ𪬵��c.Expand(�p�f����) = True


tree�r�ξ𪬵��c.Redraw = True
���c���A�C = ""
mdi�~�r�r��.txt���A = ���c���A�C

Screen.MousePointer = ccDefault

End Sub

Private Sub imglock_Click()

If imglock.Tag = 0 Then
    imglock.Tag = 1
    imglock.Picture = imgPinPush.Picture
    imglock.ToolTipText = "�Ѱ���w"
    frm�r�κt��.Caption = "�r�κt��(��w)"
Else
    imglock.Tag = 0
    imglock.Picture = imgPin.Picture
    imglock.ToolTipText = "��w"
    frm�r�κt��.Caption = "�r�κt��"
End If

End Sub

Private Sub tree�r�ξ𪬵��c_Click()
Dim �r�� As String
Dim �r�� As String
Dim �s�� As Long

If tree�r�ξ𪬵��c.ListIndex <> -1 Then
   If Len(tree�r�ξ𪬵��c.List(tree�r�ξ𪬵��c.ListIndex)) = 1 Then
      �r�� = tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.ListIndex)
      �r�� = tree�r�ξ𪬵��c.List(tree�r�ξ𪬵��c.ListIndex)
      �s�� = tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.ListIndex)
      �^���ݩ� �r��, �r��, �s��
      �^���c�r�� �r��, �r��, �s��
      If mdi�~�r�r��.txt�r��.font.Name = "�з���" Then �즲�r�� = tree�r�ξ𪬵��c.List(tree�r�ξ𪬵��c.ListIndex)
      If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� �r��, �r��, �s��
      If �Ұʲ���r�� Then frm����r��.���J�r�� �r��, �r��, �s��
      If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� �r��, �r��, �s��
      If �Ұʲ���r�� Then frm����r��.���J�r�� �r��, �r��, �s��
    End If
End If

End Sub

Private Sub tree�r�ξ𪬵��c_GotFocus()

�{�ε����N�X = �r�κt�ܥN�X

End Sub

Private Sub tree�r�ξ𪬵��c_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' Initiates dragging only after moving at least 100 twips with the mouse depressed
If (Button And 1) And (XCheck > 0) And (YCheck > 0) And ((Abs(XCheck - x) > 150) Or (Abs(YCheck - y) > 150)) Then
    XCheck = 0
    YCheck = 0  ' Reset mouse coordinates
    If tree�r�ξ𪬵��c.ListIndex >= 0 Then
        tree�r�ξ𪬵��c.BeforeDrag
        tree�r�ξ𪬵��c.Drag 1         ' Start drag
    End If
End If

'If Button = 1 Then
'    tree�r�ξ𪬵��c.BeforeDrag
'    tree�r�ξ𪬵��c.Drag 1
'End If

End Sub

Private Sub tree�r�ξ𪬵��c_DragOver(Source As Control, x As Single, y As Single, State As Integer)

tree�r�ξ𪬵��c.OnDragOver x, y, State

End Sub


Private Sub tree�r�ξ𪬵��c_LostFocus()

tree�r�ξ𪬵��c.ListIndex = -1

End Sub

Private Sub tree�r�ξ𪬵��c_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

XCheck = x
YCheck = y

End Sub
