VERSION 5.00
Object = "{65996203-3B87-11D4-A21F-00E029189826}#6.9#0"; "TLIST6.OCX"
Begin VB.Form frm�X�B�˦r 
   Caption         =   "�X�B�˦r"
   ClientHeight    =   5844
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   4812
   Icon            =   "frm�X�B�˦r.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5844
   ScaleWidth      =   4812
   Begin VB.ComboBox cbo�X�B 
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4392
   End
   Begin TListProLibCtl.TList tree�r�ξ𪬵��c 
      DragIcon        =   "frm�X�B�˦r.frx":030A
      Height          =   2652
      Left            =   240
      TabIndex        =   0
      Top             =   660
      Width           =   4368
      _Version        =   393225
      _ExtentX        =   7705
      _ExtentY        =   4678
      _StockProps     =   228
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
      PictureOpen     =   "frm�X�B�˦r.frx":074C
      PictureClosed   =   "frm�X�B�˦r.frx":085E
      PictureLeaf     =   "frm�X�B�˦r.frx":0970
      PictureMark     =   "frm�X�B�˦r.frx":0A82
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
      ExchangeSerialNumber=   "frm�X�B�˦r.frx":0B7C
      DragIconStyle   =   0
      ExchangeDefItemCellDef=   "frm�X�B�˦r.frx":0BC9
      _ChkCounter     =   -1
      TreeLinesHighlightColor=   -2113929196
      TreeLinesShadowColor=   -2113929200
   End
End
Attribute VB_Name = "frm�X�B�˦r"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private �����N�X As Integer, ���� As String
Private �ϰ�r��}�C(0 To �r��Ӽ�) As Variant
Private ���_ As Boolean
Private �`��  As Long
Private XCheck As Single, YCheck As Single


Private Sub cbo�X�B_GotFocus()

�{�α���N�X = �X�B�˦r_�˦r���

End Sub

Public Sub cbo�X�B_KeyPress(KeyAscii As Integer)

Dim i As Integer

If KeyAscii = vbKeyReturn Then
    Screen.MousePointer = 11
   
    i = 0
    Do While �r��}�C(i) <> ""
        �ϰ�r��}�C(i) = �r��}�C(i)
        i = i + 1
    Loop
   
    ���_ = False
    'If cbo�X�B.Text <> "����" And cbo�X�B.Text <> "�X��" Then
        ���J�� Trim(cbo�X�B.Text)
    'End If
    Screen.MousePointer = 0
End If

End Sub

Private Sub Form_Activate()

�{�ε��� = ����
�{�ε����N�X = �X�B�˦r�N�X
'�{�ε����N�X = �����N�X
��������r�Τu��C���A �{�ε����N�X
mdi�~�r�r��.txt���A = �F�Ū��A�C

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyEscape Then
   DoEvents
   ���_ = True
End If

End Sub

Private Sub Form_Load()
Dim i As Integer

�ҰʥX�B�˦r = True
If ��lfirst <> 1 Then
   If �w���J�e�� = 0 Then
      If ����winstate = 0 Then
         frm�X�B�˦r.Left = �X�Bleft
         frm�X�B�˦r.Top = �X�Btop
         frm�X�B�˦r.Height = �X�Bheight
         frm�X�B�˦r.Width = �X�Bwidth
      Else
         frm�X�B�˦r.WindowState = �X�Bwinstate
      End If
   ElseIf �Ұʦr�δF�� Then
         frm�X�B�˦r.Left = frm�r�δF��.Left
         frm�X�B�˦r.Top = frm�r�δF��.Top
         frm�X�B�˦r.Height = frm�r�δF��.Height
         frm�X�B�˦r.Width = frm�r�δF��.Width
   End If
End If

tree�r�ξ𪬵��c.FontSize = CInt(��ܦr���j�p)
Me.Tag = �X�B�˦r�N�X
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

cbo�X�B.List(0) = "�X��(�Ұ���X��)"
cbo�X�B.List(1) = "��(�p�٫n�a�Ұ�)"
cbo�X�B.List(2) = "�^(�^����åҰ���)"
cbo�X�B.List(3) = "�h(�h�S�󵥩��åҰ���)"
cbo�X�B.List(4) = "����(��P���嶰��)"
cbo�X�B.List(5) = "����"
cbo�X�B.List(6) = "�������"
cbo�X�B.List(7) = "����j��"
cbo�X�B.List(8) = "����ó��"
cbo�X�B.List(9) = "����f��"
cbo�X�B.List(10) = "����U�r"
cbo�X�B.List(11) = "����_�r"
cbo�X�B.List(12) = "�ѤR(�����ѬP�[1���ӤR�b)"
cbo�X�B.List(13) = "�ѵ�(�����ѬP�[1���ӻ���)"
cbo�X�B.List(14) = "�]2(����]�s2����)"
cbo�X�B.List(15) = "��25(���F���Ѵ�25����)"
cbo�X�B.List(16) = "����(���F�l�u�w�����ѥҽg)"
cbo�X�B.List(17) = "���A(���F�l�u�w�����ѤA�g)"
cbo�X�B.List(18) = "�B21(�����B�O�s21����)"
cbo�X�B.List(19) = "�H1(�H��1���Ӧˮ�)"
cbo�X�B.List(20) = "�H2(�H��1���ӻ���)"
cbo�X�B.List(21) = "�S27(�����S�a�Y27����)"
cbo�X�B.List(22) = "��1(�������a�L1����)"
cbo�X�B.List(23) = "��13(�������a�L13����)"
cbo�X�B.List(24) = "��99(�������a�L99����)"
cbo�X�B.List(25) = "��1(�������s1����)"
cbo�X�B.List(26) = "�`2(�`�w���w�s�i���Y2����)"
cbo�X�B.List(27) = "��1(������s1����)"
cbo�X�B.List(28) = "��2(������s2����)"
cbo�X�B.List(29) = "��(���ԤA��)"
cbo�X�B.List(30) = "�P406(���F�����P406����)"
cbo�X�B.List(31) = "�j370(�����j�˼t370����)"

End Sub

Private Sub Form_Resize()
Dim frm���� As Integer

frm���� = Me.ScaleHeight - cbo�X�B.Height - cbo�X�B.Top * 3

If frm���� > 0 Then
   tree�r�ξ𪬵��c.Height = frm����
End If

If (Me.ScaleWidth - tree�r�ξ𪬵��c.Left * 2) > 0 Then
   tree�r�ξ𪬵��c.Width = Me.ScaleWidth - tree�r�ξ𪬵��c.Left * 2
End If

cbo�X�B.Width = tree�r�ξ𪬵��c.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)

mdi�~�r�r��.mnu_�X�B�˦r.Enabled = True
�p��{�ε���
�ҰʥX�B�˦r = False

End Sub

Private Sub ���J��(�X�B As String)

Dim �r�Ϊ� As Recordset, SQL���z�� As String
Dim i As Integer
Dim ��� As String, �r�� As String
Dim �ʤ��� As Long, �F���`�� As Long
Dim �X���s�� As Long, �������� As Long
Dim leftpos As Integer, rightpos As Integer

mdi�~�r�r��.txt���A = "�r�θ��J��  ......       " & " , �����_�Ы� Esc ��"
Screen.MousePointer = ccHourglass

tree�r�ξ𪬵��c.Clear

leftpos = InStr(1, �X�B, "(")
rightpos = InStr(1, �X�B, ")")

If leftpos > 0 Then
    ��� = Left(�X�B, leftpos - 1)
    If rightpos < Len(�X�B) Then ��� = ��� & Right(�X�B, Len(�X�B) - rightpos)
Else
    ��� = �X�B
End If

DoEvents

��ڧǼ� = 0
�F���`�� = 0
�`�� = 0

�X�B��r�� (���)

If �X�B���Ұ��� Then
    If �X�B���Ұ���X�� Then
        If �X�B�����ǰt Then
            �X���s�� = CStr(CLng(Right(���, Len(���) - 2)))
            SQL���z�� = "SELECT * From ���g�r�� Where (�X�B='" & �X���s�� & "') And (�W�u>0) And (�W�u<6) And (�W�u<>3) ORDER BY �s��"
        Else
            SQL���z�� = "SELECT * From ���g�r�� Where (�X��=1) And (�W�u>0) And (�W�u<6) And (�W�u<>3) ORDER BY �s��"
        End If
    Else
        If �X�B�����ǰt Then
            SQL���z�� = "SELECT * From ���g�r�� Where (�X�B='" & ��� & "') And (�W�u>0) And (�W�u<6) And (�W�u<>3) ORDER BY �s��"
        Else
            SQL���z�� = "SELECT * From ���g�r�� Where (�X�B like '" & ��� & "*') And (�W�u>0) And (�W�u<6) And (�W�u<>3) ORDER BY �s��"
        End If
    End If
    Set �r�Ϊ� = �Ұ����Ʈw.OpenRecordset(SQL���z��)
ElseIf �X�B������ Then
    If �X�B�����ǰt Then
        �������� = CLng(Right(���, Len(���) - 2))
        SQL���z�� = "SELECT * From ���g�r�� Where (����=" & �������� & ") And (�W�u>0) And (�W�u<6) And (�W�u<>3) ORDER BY �s��"
    Else
        SQL���z�� = "SELECT * From ���g�r�� Where (�W�u>0) And (�W�u<6) And (�W�u<>3) ORDER BY �s��"
    End If
    Set �r�Ϊ� = �����Ʈw.OpenRecordset(SQL���z��)
ElseIf �X�B���p�f Then
    SQL���z�� = "SELECT * From �˦r�� Where �r��='" & ��� & "' ORDER BY �s��"
    Set �r�Ϊ� = �p�f��Ʈw.OpenRecordset(SQL���z��)
Else
    If InStr(1, �X�B, ".") > 0 Then
        SQL���z�� = "SELECT * From ���g�r�� Where (�X�B='" & ��� & "') And (�W�u>0) And (�W�u<6) And (�W�u<>3) ORDER BY �s��"
    Else
        SQL���z�� = "SELECT * From ���g�r�� Where (�X�B like'" & ��� & "*') And (�W�u>0) and (�W�u<6) And (�W�u<>3) ORDER BY �s��"
    End If
    Set �r�Ϊ� = ���t��r��Ʈw.OpenRecordset(SQL���z��)
End If
   
If Not �r�Ϊ�.EOF Then
   
    �r�Ϊ�.MoveLast
    �`�� = �r�Ϊ�.RecordCount
            
    tree�r�ξ𪬵��c.AddItem ���
      
    tree�r�ξ𪬵��c.ItemLngValue(0) = -999999
    tree�r�ξ𪬵��c.ItemTag(0) = ��L�`�I�аO
    tree�r�ξ𪬵��c.ItemFontName(0) = ��ܦr��
    tree�r�ξ𪬵��c.Expand(0) = True
    
    �r�Ϊ�.MoveFirst

    Do Until �r�Ϊ�.EOF
        If ���_ = True Then Exit Do
         
        ��ڧǼ� = ��ڧǼ� + 1
        �ʤ��� = ��ڧǼ� / �`�� * 100
        mdi�~�r�r��.txt���A = "�r�θ��J�w���� " & �ʤ��� & " % , �����_�Ы� Esc ��"
         
        �r�� = �r�Ϊ�.Fields("�r�X")
        tree�r�ξ𪬵��c.AddItem �r��, 0
        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
        tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.NewIndex) = �ϰ�r��}�C(�r�Ϊ�.Fields("�r��"))
        tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = �r�Ϊ�.Fields("�s��")
                  'If Not IsNull(�r�Ϊ�.Fields("�r��")) Then
                   '  tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �r�θ`�I�аO
                  'Else
                   '  tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �c�r���`�I�аO
                  'End If
        �r�Ϊ�.MoveNext
        
        If (tree�r�ξ𪬵��c.ListCount + 10) Mod 50 = 0 Then
            tree�r�ξ𪬵��c.Redraw = True
            Screen.MousePointer = ccDefault
            DoEvents
            tree�r�ξ𪬵��c.Redraw = False
            Screen.MousePointer = ccHourglass
        End If
         
      Loop
Else
      
      �`�� = 0

End If


tree�r�ξ𪬵��c.Redraw = True
If ���_ = True Then
    �F�Ū��A�C = "�ϥΪ̤��_�I �w���J " & ��ڧǼ� & " �Ӧr��"
Else
    �F�Ū��A�C = "���J�����I �@ " & �`�� & " �Ӧr��"
End If
mdi�~�r�r��.txt���A = �F�Ū��A�C

Screen.MousePointer = ccDefault

End Sub

Private Sub tree�r�ξ𪬵��c_Click()

Dim �r�� As String
Dim �r�� As String
Dim �s�� As Long

If tree�r�ξ𪬵��c.ListIndex <> -1 Then
   If tree�r�ξ𪬵��c.List(0) <> "" Then
      �r�� = tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.ListIndex)
      �r�� = tree�r�ξ𪬵��c.List(tree�r�ξ𪬵��c.ListIndex)
      If Len(�r��) = 1 Then
      �s�� = tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.ListIndex)
      
      �^���ݩ� �r��, �r��, �s��
      �^���c�r�� �r��, �r��, �s��
      If mdi�~�r�r��.txt�r��.font.Name = "�з���" Then �즲�r�� = tree�r�ξ𪬵��c.List(tree�r�ξ𪬵��c.ListIndex)
      
      If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� �r��, �r��, �s��
      If �Ұʲ���r�� Then frm����r��.���J�r�� �r��, �r��, �s��
      If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� �r��, �r��, �s��
      If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� �r��, �r��, �s��
      If �Ұʲ���r�� Then frm����r��.���J�r�� �r��, �r��, �s��
      mdi�~�r�r��.txt���A = �F�Ū��A�C
      End If
   End If
End If

End Sub

Private Sub tree�r�ξ𪬵��c_GotFocus()

�{�α���N�X = �X�B�˦r_�𪬵��c

End Sub


