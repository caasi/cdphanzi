VERSION 5.00
Object = "{65996203-3B87-11D4-A21F-00E029189826}#6.9#0"; "TLIST6.OCX"
Begin VB.Form frm�r�ί��� 
   Caption         =   "�r�ί���"
   ClientHeight    =   6228
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9060
   Icon            =   "frm�r�ί���.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6228
   ScaleWidth      =   9060
   Begin TListProLibCtl.TList tree�r�ξ𪬵��c 
      DragIcon        =   "frm�r�ί���.frx":030A
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
      PictureOpen     =   "frm�r�ί���.frx":074C
      PictureClosed   =   "frm�r�ί���.frx":085E
      PictureLeaf     =   "frm�r�ί���.frx":0970
      PictureMark     =   "frm�r�ί���.frx":0A82
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
      ExchangeSerialNumber=   "frm�r�ί���.frx":0B7C
      DragIconStyle   =   0
      ExchangeDefItemCellDef=   "frm�r�ί���.frx":0BC9
      _ChkCounter     =   -1
      TreeLinesHighlightColor=   -2113929196
      TreeLinesShadowColor=   -2113929200
   End
   Begin VB.Image imgPinPush 
      Height          =   264
      Left            =   4440
      Picture         =   "frm�r�ί���.frx":0CD0
      Top             =   240
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imgPin 
      Height          =   264
      Left            =   4440
      Picture         =   "frm�r�ί���.frx":0E5A
      Top             =   720
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imglock 
      Height          =   264
      Left            =   0
      Picture         =   "frm�r�ί���.frx":0FE4
      Tag             =   "0"
      ToolTipText     =   "��w"
      Top             =   240
      Width           =   288
   End
End
Attribute VB_Name = "frm�r�ί���"
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
�{�ε����N�X = �r�ί��ޥN�X
��������r�Τu��C���A �{�ε����N�X
tree�r�ξ𪬵��c_Click
mdi�~�r�r��.txt���A = ���c���A�C

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim �r�ڧ� As String, �s�� As Long

�Ұʦr�ί��� = True
If ��lfirst <> 1 Then
   If �w���J�e�� = 0 Then
      If ����winstate = 0 Then
         frm�r�ί���.Left = ����left
         frm�r�ί���.Top = ����top
         frm�r�ί���.Height = ����height
         frm�r�ί���.Width = ����width
      Else
         frm�r�ί���.WindowState = ����winstate
      End If
   ElseIf �Ұʦr�δF�� Then
         frm�r�ί���.Left = frm�r�δF��.Left + frm�r�δF��.Width
         frm�r�ί���.Top = frm�r�δF��.Top
         frm�r�ί���.Height = frm�r�δF��.Height
         frm�r�ί���.Width = frm�r�δF��.Width
   End If
End If

tree�r�ξ𪬵��c.FontSize = CInt(��ܦr���j�p)
'Me.Tag = �@�ε����N�X
Me.Tag = �r�ί��ޥN�X
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
mdi�~�r�r��.mnu_�r�ί���.Enabled = True
�Ұʦr�ί��� = False
�p��{�ε���

End Sub

Public Sub ���J�r��(�t�Φr�� As String, �r�� As String, �s�� As Long)
Dim ����� As String, �r��s�� As Integer, �r���� As String, ���� As Integer
Dim i As Integer, �r�� As String
Dim ���ѽs�� As Long, �p�f�s�� As Integer, ����s�� As Long, �Ұ���s�� As Integer, ���t��r�s�� As Long
Dim ���Ѧr�� As String, �p�f�r�� As String, ����r�� As String, �Ұ���r�� As String, ���t��r�r�� As String
Dim �p�f�r�� As String, ����r�� As String, �Ұ���r�� As String, ���t��r�r�� As String
Dim �����˦r As Boolean, ����ɿ� As Boolean
Dim ���t��r�˦r As Boolean, ���t��r�ɿ� As Boolean
Dim ItemNo As Integer, ItemBig5 As Integer, ItemUnicode As Integer
Dim Item�~�y�j�r�� As Integer, Item���F�~�y�j�r�� As Integer, Item�ا��~�y�j�r�� As Integer
Dim Item����j��� As Integer
Dim Item���� As Integer, Item������L As Integer, Item���夤�� As Integer
Dim Item���� As Integer, Item����s As Integer, Item������L As Integer, Item���徹�� As Integer, Item����ޱo As Integer
Dim Item�Ұ��� As Integer, Item�Ұ�������ġ As Integer, Item�Ұ���r���L As Integer, Item�Ұ���r���� As Integer
Dim Item���t��r, Item���t²����r�s As Integer, Item���t��r�X�B As Integer
Dim ���� As String, ����r�Y As Integer

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
    ���岧�g�r��.Index = "�s��"
    ���岧�g�r��.Seek "=", �s��
    If Not IsNull(���岧�g�r��.Fields("���ѽs��")) Then
        ���ѽs�� = ���岧�g�r��.Fields("���ѽs��")
    Else
        �����˦r��.Index = "�s��"
        �����˦r��.Seek "=", �s��
        ���ѽs�� = �����˦r��.Fields("���ѽs��")
    End If
ElseIf ����|�Ұ���(�t�Φr��) Then
    �Ұ���s�� = �s��
    �Ұ��岧�g�r��.Index = "�s��"
    �Ұ��岧�g�r��.Seek "=", �s��
    If Not IsNull(�Ұ��岧�g�r��.Fields("���ѽs��")) Then
        ���ѽs�� = �Ұ��岧�g�r��.Fields("���ѽs��")
    Else
        �Ұ����˦r��.Index = "�s��"
        �Ұ����˦r��.Seek "=", �s��
        ���ѽs�� = �Ұ����˦r��.Fields("���ѽs��")
    End If
ElseIf ����|���t��r(�t�Φr��) Then
    ���t��r�s�� = �s��
    ���t��r�˦r��.Index = "�s��"
    ���t��r�˦r��.Seek "=", �s��
    If Not IsNull(���t��r���g�r��.Fields("���ѽs��")) Then
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

'If �t�Φr�� <> "�_�v�j����p�f" And �t�Φr�� <> "�_�v�j���孫��" Then
'    If Not IsNull(�����˦r��.Fields("�p�f�s��")) Then
'        �p�f�s�� = �����˦r��.Fields("�p�f�s��")
'    Else
'        �p�f�s�� = 0
'    End If
'End If

'If �t�Φr�� <> "����|����" Then
'    If Not IsNull(�����˦r��.Fields("����s��")) Then
'        ����s�� = �����˦r��.Fields("����s��")
'    Else
'        ����s�� = 0
'    End If
'End If

tree�r�ξ𪬵��c.Clear
tree�r�ξ𪬵��c.Redraw = False
tree�r�ξ𪬵��c.FontName = ��ܦr��

tree�r�ξ𪬵��c.AddItem ���Ѧr��
tree�r�ξ𪬵��c.ItemFontName(0) = �ഫ��ܦr��(�r��)
tree�r�ξ𪬵��c.ItemLngValue(0) = ���ѽs��
If Not IsNull(�����˦r��.Fields("�r��")) Then
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �r�θ`�I�аO
Else
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �c�r���`�I�аO
End If

'If �p�f�s�� = 0 And ����s�� = 0 Then
'    tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureLeaf
'    tree�r�ξ𪬵��c.Redraw = True
'    Screen.MousePointer = ccDefault
'    Exit Sub
'End If
tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureOpen

ItemNo = 1
Item�~�y�j�r�� = -1
Item���F�~�y�j�r�� = -1
Item�ا��~�y�j�r�� = -1
Item����j��� = -1
Item���� = -1
Item������L = -1
Item���夤�� = -1
Item���� = -1
Item����s = -1
Item������L = -1
Item���徹�� = -1
Item����ޱo = -1
Item�Ұ��� = -1
Item�Ұ�������ġ = -1
Item�Ұ���r���L = -1
Item�Ұ���r���� = -1
Item���t��r = -1
Item���t²����r�s = -1
Item���t��r�X�B = -1
ItemBig5 = -1
ItemUnicode = -1

If Not IsNull(�����˦r��.Fields("�~�y�j�r��")) And (mdi�~�r�r��.mnu_���F�~�y�j�r��ﶵ.Checked Or mdi�~�r�r��.mnu_�ا��~�y�j�r��ﶵ.Checked) Then
    Item�~�y�j�r�� = tree�r�ξ𪬵��c.ListCount
    tree�r�ξ𪬵��c.AddItem "�~�y�j�r��", 0
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
    If Not IsNull(�˦r��.Fields("����")) Then ���� = �M�䳡��(�����˦r��.Fields("����"))
    tree�r�ξ𪬵��c.AddItem "����:" & ����, Item�~�y�j�r��
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    
    If Not IsNull(�����˦r��.Fields("���F�~�y�j�r��")) And mdi�~�r�r��.mnu_���F�~�y�j�r��ﶵ.Checked Then
        Item���F�~�y�j�r�� = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "���F�ϮѤ��q", Item�~�y�j�r��
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "�U-��-�r:" & �����˦r��.Fields("���F�~�y�j�r��"), Item���F�~�y�j�r��
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If

    If Not IsNull(�����˦r��.Fields("�ا��~�y�j�r��")) And mdi�~�r�r��.mnu_�ا��~�y�j�r��ﶵ.Checked Then
        Item�ا��~�y�j�r�� = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "�ا��X����", Item�~�y�j�r��
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "��-�r:" & �����˦r��.Fields("�ا��~�y�j�r��"), Item�ا��~�y�j�r��
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If

End If

If Not IsNull(�����˦r��.Fields("����j������")) And mdi�~�r�r��.mnu_����j���ﶵ.Checked Then
    Item����j��� = tree�r�ξ𪬵��c.ListCount
    tree�r�ξ𪬵��c.AddItem "����j���(���ؾǳN�|)", 0
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
    If Not IsNull(�˦r��.Fields("����j��峡��")) Then ���� = �M�䳡��(�����˦r��.Fields("����j��峡��"))
    tree�r�ξ𪬵��c.AddItem "����:" & ����, Item����j���
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    
    tree�r�ξ𪬵��c.AddItem "�U-�s��:" & �����˦r��.Fields("����j������"), Item����j���
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
End If

If Not IsNull(�����˦r��.Fields("�p�f�s��")) And (mdi�~�r�r��.mnu_����Ѧr���L�ﶵ.Checked Or mdi�~�r�r��.mnu_���ػ���Ѧr�ﶵ.Checked) Then
    �p�f�s�� = �����˦r��.Fields("�p�f�s��")
    �p�f�˦r��.Index = "�s��"
    �p�f�˦r��.Seek "=", �p�f�s��

    Item���� = tree�r�ξ𪬵��c.ListCount
    tree�r�ξ𪬵��c.AddItem "����Ѧr", 0
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
    ���� = �M�仡�峡��(�p�f�˦r��.Fields("����"))
    tree�r�ξ𪬵��c.AddItem "����:" & ����, Item����
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    
    tree�r�ξ𪬵��c.AddItem "��:" & �p�f�˦r��.Fields("��"), Item����
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    
    If Not IsNull(�p�f�˦r��.Fields("���L����")) And mdi�~�r�r��.mnu_����Ѧr���L�ﶵ.Checked Then
        Item������L = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "����Ѧr���L(����ѧ�)", Item����
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "�U-��:" & �p�f�˦r��.Fields("���L����"), Item������L
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If
    
    If Not IsNull(�p�f�˦r��.Fields("���د���")) And mdi�~�r�r��.mnu_���ػ���Ѧr�ﶵ.Checked Then
        Item���夤�� = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "����Ѧr(���خѧ�)", Item����
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "��(�W/�U):" & �p�f�˦r��.Fields("���د���"), Item���夤��
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If
End If

�����˦r = False
����ɿ� = False

If Not IsNull(�����˦r��.Fields("����s��")) Then
    �����˦r = True
Else
    ���嶰���ޱo.Seek "=", ���ѽs��
    If Not ���嶰���ޱo.NoMatch Then
        ����ɿ� = True
    Else
        ����ɿ��.Seek "=", ���ѽs��
        If Not ����ɿ��.NoMatch Then ����ɿ� = True
    End If
End If

If �����˦r And (mdi�~�r�r��.mnu_����s�ﶵ.Checked Or mdi�~�r�r��.mnu_������L�ﶵ.Checked Or mdi�~�r�r��.mnu_��P���嶰�������ﶵ.Checked Or mdi�~�r�r��.mnu_��P���嶰���ޱo�ﶵ.Checked) Then
    If Not ����|����(�t�Φr��) Then
        ����s�� = �����˦r��.Fields("����s��")
    End If
    ���岧�g�r��.Index = "�s��"
    ���岧�g�r��.Seek "=", ����s��
    
    ����r�Y = ���岧�g�r��.Fields("�ո�")
    
    Item���� = tree�r�ξ𪬵��c.ListCount
    tree�r�ξ𪬵��c.AddItem "����", 0
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
    If Not ���岧�g�r��.NoMatch And mdi�~�r�r��.mnu_����s�ﶵ.Checked Then
        Item����s = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "����s(���خѧ�)", Item����
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "�s��-��-�r:" & ���岧�g�r��.Fields("����s����"), Item����s
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If
    
    ������L.Seek "=", ����r�Y
    If Not ������L.NoMatch And mdi�~�r�r��.mnu_������L�ﶵ.Checked Then
        Item������L = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "������L(���䤤��j��)", Item����
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "��-��:" & ������L.Fields("����"), Item������L
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If
    
    'If Not IsNull(���岧�g�r��.Fields("����")) And mdi�~�r�r��.mnu_��P���嶰�������ﶵ.Checked Then
        'Item���徹�� = tree�r�ξ𪬵��c.ListCount
        'tree�r�ξ𪬵��c.AddItem "��P���嶰��(�������|�ҥj��)", Item����
        'tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        'tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        'tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        'tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        'tree�r�ξ𪬵��c.AddItem "����:" & ���岧�g�r��.Fields("����"), Item���徹��
        'tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        'tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        'tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        'tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    'End If
    
    ���嶰���ޱo.Seek "=", ���ѽs��
    If Not ���嶰���ޱo.NoMatch And mdi�~�r�r��.mnu_��P���嶰���ޱo�ﶵ.Checked Then
        Item����ޱo = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "��P���嶰���ޱo(���خѧ�)", Item����
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "�s��-���X:" & ���嶰���ޱo.Fields("����"), Item����ޱo
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If
End If

If ����ɿ� And (mdi�~�r�r��.mnu_����s�ﶵ.Checked Or mdi�~�r�r��.mnu_������L�ﶵ.Checked Or mdi�~�r�r��.mnu_��P���嶰�������ﶵ.Checked Or mdi�~�r�r��.mnu_��P���嶰���ޱo�ﶵ.Checked) Then

    
    Item���� = tree�r�ξ𪬵��c.ListCount
    tree�r�ξ𪬵��c.AddItem "����", 0
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
    ����ɿ��.Seek "=", ���ѽs��
    If ����ɿ��.NoMatch Then GoTo ��g�����ޱo����
    
    If Not IsNull(����ɿ��.Fields("����s����")) And mdi�~�r�r��.mnu_����s�ﶵ.Checked Then
        Item����s = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "����s(���خѧ�)", Item����
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "�s��-��-�r:" & ����ɿ��.Fields("����s����"), Item����s
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If
    
    If Not IsNull(����ɿ��.Fields("������L����")) And mdi�~�r�r��.mnu_������L�ﶵ.Checked Then
        Item������L = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "������L(���䤤��j��)", Item����
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "��-��:" & ����ɿ��.Fields("������L����"), Item������L
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If
    
��g�����ޱo����:

    ���嶰���ޱo.Seek "=", ���ѽs��
    If Not ���嶰���ޱo.NoMatch And mdi�~�r�r��.mnu_��P���嶰���ޱo�ﶵ.Checked Then
        Item����ޱo = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "��P���嶰���ޱo(���خѧ�)", Item����
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "�s��-���X:" & ���嶰���ޱo.Fields("����"), Item����ޱo
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If
End If

If Not IsNull(�����˦r��.Fields("�Ұ���s��")) And (mdi�~�r�r��.mnu_�Ұ�������ġ�ﶵ.Checked Or mdi�~�r�r��.mnu_�Ұ���r���L�ﶵ.Checked Or mdi�~�r�r��.mnu_�Ұ���r�����ﶵ.Checked) Then
    If Not ����|�Ұ���(�t�Φr��) Then
        �Ұ���s�� = �����˦r��.Fields("�Ұ���s��")
    End If
    �Ұ��岧�g�r��.Index = "�s��"
    �Ұ��岧�g�r��.Seek "=", �Ұ���s��

    Item�Ұ��� = tree�r�ξ𪬵��c.ListCount
    tree�r�ξ𪬵��c.AddItem "�Ұ���", 0
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
    If Not IsNull(�Ұ��岧�g�r��.Fields("�Ұ�������ġ")) And mdi�~�r�r��.mnu_�Ұ�������ġ�ﶵ.Checked Then
        Item�Ұ�������ġ = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "��V�Ұ�������ġ(�N�L�j��)", Item�Ұ���
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "�U-��:" & �Ұ��岧�g�r��.Fields("�Ұ�������ġ"), Item�Ұ�������ġ
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If
    
    If Not IsNull(�Ұ��岧�g�r��.Fields("�Ұ���r���L")) And mdi�~�r�r��.mnu_�Ұ���r���L�ﶵ.Checked Then
        Item�Ұ���r���L = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "�Ұ���r���L(�N�L�j��)", Item�Ұ���
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "�U-��:" & �Ұ��岧�g�r��.Fields("�Ұ���r���L"), Item�Ұ���r���L
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If
    
    If Not IsNull(�Ұ��岧�g�r��.Fields("�Ұ���r����")) And mdi�~�r�r��.mnu_�Ұ���r�����ﶵ.Checked Then
        Item�Ұ���r���� = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "�Ұ���r����(������s�|)", Item�Ұ���
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "��-��:" & �Ұ��岧�g�r��.Fields("�Ұ���r����"), Item�Ұ���r����
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If
    
End If

���t��r�˦r = False
���t��r�ɿ� = False

If Not IsNull(�����˦r��.Fields("���t��r�s��")) Then
    ���t��r�˦r = True
Else
    ���t��r�ɿ��.Seek "=", ���ѽs��
    If Not ���t��r�ɿ��.NoMatch Then ���t��r�ɿ� = True
End If

If ���t��r�˦r And (mdi�~�r�r��.mnu_���t²����r�s�ﶵ.Checked Or mdi�~�r�r��.mnu_���t��r�X�B�ﶵ.Checked) Then
    If Not ����|���t��r(�t�Φr��) Then
        ���t��r�s�� = �����˦r��.Fields("���t��r�s��")
    End If
    ���t��r���g�r��.Index = "�s��"
    ���t��r���g�r��.Seek "=", ���t��r�s��

    Item���t��r = tree�r�ξ𪬵��c.ListCount
    tree�r�ξ𪬵��c.AddItem "���t��r", 0
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
    If Not IsNull(���t��r���g�r��.Fields("���t²����r�s")) And mdi�~�r�r��.mnu_���t²����r�s�ﶵ.Checked Then
        Item���t²����r�s = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "���t²����r�s(��_�Ш|�X����)", Item���t��r
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "���X-��r:" & ���t��r���g�r��.Fields("���t²����r�s"), Item���t²����r�s
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If
    
    'If Not IsNull(���t��r�˦r��.Fields("�X�B")) And mdi�~�r�r��.mnu_���t��r�X�B�ﶵ.Checked Then
    '    Item���t��r�X�B = tree�r�ξ𪬵��c.ListCount
    '    tree�r�ξ𪬵��c.AddItem "�X�B", Item���t��r
    '    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    '    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    '    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    '    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
    '    tree�r�ξ𪬵��c.AddItem "�Ӹ�-²��:" & ���t��r�˦r��.Fields("�X�B"), Item���t��r�X�B
    '    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    '    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    '    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    '    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    'End If

End If

If ���t��r�ɿ� And (mdi�~�r�r��.mnu_���t²����r�s�ﶵ.Checked Or mdi�~�r�r��.mnu_���t��r�X�B�ﶵ.Checked) Then

    Item���t��r = tree�r�ξ𪬵��c.ListCount
    tree�r�ξ𪬵��c.AddItem "���t��r", 0
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
    If Not IsNull(���t��r�ɿ��.Fields("���t²����r�s")) And mdi�~�r�r��.mnu_���t²����r�s�ﶵ.Checked Then
        Item���t²����r�s = tree�r�ξ𪬵��c.ListCount
        tree�r�ξ𪬵��c.AddItem "���t²����r�s(��_�Ш|�X����)", Item���t��r
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
        tree�r�ξ𪬵��c.AddItem "���X-��r:" & ���t��r�ɿ��.Fields("���t²����r�s"), Item���t²����r�s
        tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
        tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
    End If

End If

If Not IsNull(�����˦r��.Fields("Unicode")) And mdi�~�r�r��.mnu_Unicode�ﶵ.Checked Then
    ItemUnicode = tree�r�ξ𪬵��c.ListCount
    tree�r�ξ𪬵��c.AddItem "Unicode 3.2", 0
    'tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.NewIndex) = �ഫ��ܦr��("�з���")
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
    tree�r�ξ𪬵��c.AddItem �����˦r��.Fields("Unicode"), ItemUnicode
    'tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.NewIndex) = �ഫ��ܦr��("�з���")
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
End If

If ���ѽs�� > 0 And ���ѽs�� <= 13060 And mdi�~�r�r��.mnu_Big5�ﶵ.Checked Then
    ItemBig5 = tree�r�ξ𪬵��c.ListCount
    tree�r�ξ𪬵��c.AddItem "Big5", 0
    'tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.NewIndex) = �ഫ��ܦr��("�з���")
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureOpen
    
    tree�r�ξ𪬵��c.AddItem �����˦r��.Fields("Big5"), ItemBig5
    'tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.NewIndex) = �ഫ��ܦr��("�з���")
    tree�r�ξ𪬵��c.ItemFontSize(tree�r�ξ𪬵��c.NewIndex) = 12
    tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = -9999
    tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = ��L�`�I�аO
    tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
End If

tree�r�ξ𪬵��c.Expand(0) = True
If Item�~�y�j�r�� > -1 Then tree�r�ξ𪬵��c.Expand(Item�~�y�j�r��) = True
If Item���F�~�y�j�r�� > -1 Then tree�r�ξ𪬵��c.Expand(Item���F�~�y�j�r��) = True
If Item�ا��~�y�j�r�� > -1 Then tree�r�ξ𪬵��c.Expand(Item�ا��~�y�j�r��) = True
If Item����j��� > -1 Then tree�r�ξ𪬵��c.Expand(Item����j���) = True
If Item���� > -1 Then tree�r�ξ𪬵��c.Expand(Item����) = True
If Item������L > -1 Then tree�r�ξ𪬵��c.Expand(Item������L) = True
If Item���夤�� > -1 Then tree�r�ξ𪬵��c.Expand(Item���夤��) = True
If Item���� > -1 Then tree�r�ξ𪬵��c.Expand(Item����) = True
If Item����s > -1 Then tree�r�ξ𪬵��c.Expand(Item����s) = True
If Item������L > -1 Then tree�r�ξ𪬵��c.Expand(Item������L) = True
If Item���徹�� > -1 Then tree�r�ξ𪬵��c.Expand(Item���徹��) = True
If Item����ޱo > -1 Then tree�r�ξ𪬵��c.Expand(Item����ޱo) = True
If Item�Ұ��� > -1 Then tree�r�ξ𪬵��c.Expand(Item�Ұ���) = True
If Item�Ұ�������ġ > -1 Then tree�r�ξ𪬵��c.Expand(Item�Ұ�������ġ) = True
If Item�Ұ���r���L > -1 Then tree�r�ξ𪬵��c.Expand(Item�Ұ���r���L) = True
If Item�Ұ���r���� > -1 Then tree�r�ξ𪬵��c.Expand(Item�Ұ���r����) = True
If Item���t��r > -1 Then tree�r�ξ𪬵��c.Expand(Item���t��r) = True
If Item���t²����r�s > -1 Then tree�r�ξ𪬵��c.Expand(Item���t²����r�s) = True
If Item���t��r�X�B > -1 Then tree�r�ξ𪬵��c.Expand(Item���t��r�X�B) = True
If ItemBig5 > -1 Then tree�r�ξ𪬵��c.Expand(ItemBig5) = True
If ItemUnicode > -1 Then tree�r�ξ𪬵��c.Expand(ItemUnicode) = True
If tree�r�ξ𪬵��c.ListCount = 1 Then tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureLeaf


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
    frm�r�ί���.Caption = "�r�ί���(��w)"
Else
    imglock.Tag = 0
    imglock.Picture = imgPin.Picture
    imglock.ToolTipText = "��w"
    frm�r�ί���.Caption = "�r�ί���"
End If

End Sub

Private Sub tree�r�ξ𪬵��c_Click()
Dim �r�� As String
Dim �r�� As String
Dim �s�� As Long

�{�ε����N�X = �r�ί��ޥN�X

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
      If �Ұʲ���r�� Then frm����r��.���J�r�� �r��, �r��, �s��
    End If
End If

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
