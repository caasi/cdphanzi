VERSION 5.00
Object = "{65996203-3B87-11D4-A21F-00E029189826}#6.9#0"; "TLIST6.OCX"
Begin VB.Form frm����d�� 
   Caption         =   "����"
   ClientHeight    =   3132
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "�з���"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm�c�r����.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3132
   ScaleWidth      =   5640
   Begin TListProLibCtl.TList tree�r�ξ𪬵��c 
      DragIcon        =   "frm�c�r����.frx":030A
      Height          =   2652
      Left            =   240
      TabIndex        =   0
      Top             =   252
      Width           =   3852
      _Version        =   393225
      _ExtentX        =   6800
      _ExtentY        =   4683
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
      PictureOpen     =   "frm�c�r����.frx":074C
      PictureClosed   =   "frm�c�r����.frx":085E
      PictureLeaf     =   "frm�c�r����.frx":0970
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
      SmartDragDrop   =   -1  'True
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
      ExchangeSerialNumber=   "frm�c�r����.frx":0A82
      DragIconStyle   =   0
      ExchangeDefItemCellDef=   "frm�c�r����.frx":0ACF
      _ChkCounter     =   -1
      TreeLinesHighlightColor=   -2113929196
      TreeLinesShadowColor=   -2113929200
   End
End
Attribute VB_Name = "frm����d��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ���e As Integer, ���� As Integer
Private �����N�X As Integer, ���� As String, ���A�C As String
Private XCheck As Single, YCheck As Single

Public Sub ����d��(���󵧵e As Integer, ���󭺵� As Integer)
Dim �r�Ϊ� As Recordset, �r�� As String
Dim ���� As String

Screen.MousePointer = ccHourglass

���e = ���󵧵e
���� = ���󭺵�

tree�r�ξ𪬵��c.Clear
tree�r�ξ𪬵��c.Redraw = False

If ���e = 0 Then
    ���� = "1-99�e"
Else
    ���� = ���e & "�e"
End If

If ���� > 0 Then ���� = ���� & "(" & mdi�~�r�r��.cbo����.List(����) & ")"
If ���� = "�c�r�Ÿ�" Then
    tree�r�ξ𪬵��c.AddItem "�c�r�Ÿ�"
ElseIf ���� = "�ϧΤ�r" Then
    tree�r�ξ𪬵��c.AddItem "�ϧΤ�r"
ElseIf ���� = "�K��" Then
    tree�r�ξ𪬵��c.AddItem "�K��"
ElseIf ���� = "²�|" Then
    tree�r�ξ𪬵��c.AddItem "²�|"
Else
    tree�r�ξ𪬵��c.AddItem ����
End If
tree�r�ξ𪬵��c.ItemTag(0) = ��L�`�I�аO

Set �r�Ϊ� = �t�θ�Ʈw.OpenRecordset(����~�rSQL���z��(����, ���e, ����))

tree�r�ξ𪬵��c.Font.Name = ��ܦr��

Set �˦r�� = �����˦r��
�˦r��.Index = "�r��"

If Not �r�Ϊ�.EOF Then
   Do Until �r�Ϊ�.EOF
      �r�� = �r�Ϊ�.Fields("�r��")
      �˦r��.Seek "=", �r��
      tree�r�ξ𪬵��c.AddItem �r��, 0
      tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = �˦r��.Fields("�s��")
      tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �r�θ`�I�аO
      �r�Ϊ�.MoveNext
   Loop
   'If tree�r�ξ𪬵��c.ListCount > 0 Then tree�r�ξ𪬵��c.ListIndex = -1
End If
���󪬺A�C = tree�r�ξ𪬵��c.ListCount & " �ӳ���"
mdi�~�r�r��.txt���A = ���󪬺A�C

�r�Ϊ�.Close

tree�r�ξ𪬵��c.Expand(0) = True
tree�r�ξ𪬵��c.Redraw = True
Screen.MousePointer = ccDefault

End Sub

Public Sub Form_Activate()
�{�ε��� = ����
�{�ε����N�X = �����N�X
'mdi�~�r�r��.mnu_����N�X(�{�ε����N�X).Checked = True
��������r�Τu��C���A �{�ε����N�X, ���e, ����
mdi�~�r�r��.txt���A = ���󪬺A�C

End Sub

Public Sub Form_Load()

If ��lfirst <> 1 Then
   If �w���J�e�� = 0 Then
      If ����winstate = 0 Then
         frm����d��.Left = ����left
         frm����d��.Top = ����top
         frm����d��.Height = ����height
         frm����d��.Width = ����width
      Else
         frm����d��.WindowState = ����winstate
      End If
   ElseIf �Ұʦr�δF�� And Not �Ұʳ���d�� Then
         frm����d��.Left = frm�r�δF��.Left + frm�r�δF��.Width
         frm����d��.Top = frm�r�δF��.Top
         frm����d��.Height = frm�r�δF��.Height
         frm����d��.Width = frm�r�δF��.Width
   End If
End If

tree�r�ξ𪬵��c.Font.Size = ��ܦr���j�p
���e = mdi�~�r�r��.cbo���e.ListIndex
���� = mdi�~�r�r��.cbo����.ListIndex
�����N�X = �@�ε����N�X
���� = �@�ε���(�@�ε����N�X)
Me.Caption = �@�ε���(�@�ε����N�X)
Me.Tag = �@�ε����N�X

�{�ε��� = ����
�{�ε����N�X = �����N�X
'If �{�ε����N�X > 0 And �{�ε����N�X < 11 Then
'   mdi�~�r�r��.mnu_����N�X(�{�ε����N�X).Checked = True
'End If
��������r�Τu��C���A �{�ε����N�X, ���e, ����

����d�� ���e, ����
�Ұʳ���d�� = True

End Sub

Private Sub Form_Resize()

If Me.ScaleHeight - tree�r�ξ𪬵��c.Top * 2 > 0 Then tree�r�ξ𪬵��c.Height = Me.ScaleHeight - tree�r�ξ𪬵��c.Top * 2
If Me.ScaleWidth - tree�r�ξ𪬵��c.Left * 2 > 0 Then tree�r�ξ𪬵��c.Width = Me.ScaleWidth - tree�r�ξ𪬵��c.Left * 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
'mdi�~�r�r��.mnu_����N�X(�����N�X).Checked = False
mdi�~�r�r��.txt���A = ""
�p��{�ε���
�Ұʳ���d�� = False
End Sub


Private Sub tree�r�ξ𪬵��c_Click()
Dim �r�� As String
Dim �r�� As String
Dim �s�� As Long

If tree�r�ξ𪬵��c.ListIndex > 0 Then
   If tree�r�ξ𪬵��c.List(1) <> "" Then
      �r�� = tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.ListIndex)
      �r�� = tree�r�ξ𪬵��c.List(tree�r�ξ𪬵��c.ListIndex)
      �s�� = tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.ListIndex)
      �^���ݩ� �r��, �r��, �s��
      �^���c�r�� �r��, �r��, �s��
      If mdi�~�r�r��.txt�r��.Font.Name = "�з���" Then �즲�r�� = tree�r�ξ𪬵��c.List(tree�r�ξ𪬵��c.ListIndex)
      If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� �r��, �r��, �s��
      If �Ұʲ���r�� Then frm����r��.���J�r�� �r��, �r��, �s��
      If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� �r��, �r��, �s��
      If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� �r��, �r��, �s��
      If �Ұʲ���r�� Then frm����r��.���J�r�� �r��, �r��, �s��
      mdi�~�r�r��.txt���A = ���󪬺A�C
    End If
End If

End Sub

Private Sub tree�r�ξ𪬵��c_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

tree�r�ξ𪬵��c.OnDragOver X, Y, State

End Sub

Private Sub tree�r�ξ𪬵��c_GotFocus()

'tree�r�ξ𪬵��c_Click

End Sub

Private Sub tree�r�ξ𪬵��c_LostFocus()

tree�r�ξ𪬵��c.ListIndex = -1

End Sub

Private Sub tree�r�ξ𪬵��c_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

XCheck = X
YCheck = Y

End Sub

Private Sub tree�r�ξ𪬵��c_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Initiates dragging only after moving at least 100 twips with the mouse depressed
If (Button And 1) And (XCheck > 0) And (YCheck > 0) And ((Abs(XCheck - X) > 150) Or (Abs(YCheck - Y) > 150)) Then
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
