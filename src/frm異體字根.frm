VERSION 5.00
Object = "{65996203-3B87-11D4-A21F-00E029189826}#6.9#0"; "TLIST6.OCX"
Begin VB.Form frm����r�� 
   Caption         =   "����r��"
   ClientHeight    =   3816
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5100
   Icon            =   "frm����r��.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3816
   ScaleWidth      =   5100
   Begin TListProLibCtl.TList tree�r�ξ𪬵��c 
      DragIcon        =   "frm����r��.frx":030A
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
      PictureOpen     =   "frm����r��.frx":074C
      PictureClosed   =   "frm����r��.frx":085E
      PictureLeaf     =   "frm����r��.frx":0970
      PictureMark     =   "frm����r��.frx":0A82
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
      PicturePalette  =   "frm����r��.frx":0B94
      ExchangeSerialNumber=   "frm����r��.frx":0CA6
      DragIconStyle   =   0
      ExchangeDefItemCellDef=   "frm����r��.frx":0CF3
      _ChkCounter     =   -1
      TreeLinesHighlightColor=   -2113929196
      TreeLinesShadowColor=   -2113929200
   End
   Begin VB.Image imgPinPush 
      Height          =   264
      Left            =   4440
      Picture         =   "frm����r��.frx":0DFA
      Top             =   840
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imgPin 
      Height          =   264
      Left            =   4440
      Picture         =   "frm����r��.frx":0F84
      Top             =   1320
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imglock 
      Height          =   264
      Left            =   0
      Picture         =   "frm����r��.frx":110E
      Tag             =   "0"
      ToolTipText     =   "��w"
      Top             =   240
      Width           =   288
   End
   Begin VB.Image Image1 
      Height          =   192
      Left            =   4440
      Picture         =   "frm����r��.frx":1298
      Top             =   480
      Visible         =   0   'False
      Width           =   192
   End
End
Attribute VB_Name = "frm����r��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private �����N�X As Integer, ���� As String, �r�ڪ� As Recordset
Private �ϰ�r��}�C(0 To �r��Ӽ�) As Variant
Private ���_ As Boolean
Private �`��  As Long, ���� As String
Private XCheck As Single, YCheck As Single

Private Sub Form_Activate()
�{�ε��� = ����
'�{�ε����N�X = �����N�X
�{�ε����N�X = ����r�ڥN�X
��������r�Τu��C���A �{�ε����N�X
mdi�~�r�r��.txt���A = ���骬�A�C

End Sub


Private Sub Form_Load()

Dim i As Integer, �s�� As Long

�Ұʲ���r�� = True

If ��lfirst <> 1 Then
   If �w���J�e�� = 0 Then
      If ����winstate = 0 Then
         frm����r��.Left = ����left
         frm����r��.Top = ����top
         frm����r��.Height = ����height
         frm����r��.Width = ����width
      Else
         frm����r��.WindowState = ����winstate
      End If
   ElseIf �Ұʦr�δF�� Then
         frm����r��.Left = frm�r�δF��.Left + frm�r�δF��.Width
         frm����r��.Top = frm�r�δF��.Top
         frm����r��.Height = frm�r�δF��.Height
         frm����r��.Width = frm�r�δF��.Width
   End If
End If

Set �r�ڪ� = �t�θ�Ʈw.OpenRecordset("�r��")
�r�ڪ�.Index = "�r��"

i = 0
Do While �r��}�C(i) <> ""
   �ϰ�r��}�C(i) = �r��}�C(i)
   i = i + 1
Loop

tree�r�ξ𪬵��c.FontSize = CInt(��ܦr���j�p)
�����N�X = �@�ε����N�X
���� = �@�ε���(�@�ε����N�X)
'Me.Tag = �@�ε����N�X
Me.Tag = ����r�ڥN�X
'tree�r�ξ𪬵��c.AddItem ""
'tree�r�ξ𪬵��c.ListIndex = 0
'tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureLeaf

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


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
   DoEvents
   ���_ = True
End If

End Sub

Private Sub Form_Resize()

If Me.ScaleHeight - tree�r�ξ𪬵��c.Top * 2 > 0 Then tree�r�ξ𪬵��c.Height = Me.ScaleHeight - tree�r�ξ𪬵��c.Top * 2
If Me.ScaleWidth - tree�r�ξ𪬵��c.Left * 2 > 0 Then tree�r�ξ𪬵��c.Width = Me.ScaleWidth - tree�r�ξ𪬵��c.Left * 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
�r�ڪ�.Close
�Ұʲ���r�� = False
mdi�~�r�r��.mnu_����r��.Enabled = True
�p��{�ε���
End Sub

Private Sub imglock_Click()

If imglock.Tag = 0 Then
    imglock.Tag = 1
    imglock.Picture = imgPinPush.Picture
    imglock.ToolTipText = "�Ѱ���w"
    frm����r��.Caption = "����r��(��w)"
Else
    imglock.Tag = 0
    imglock.Picture = imgPin.Picture
    imglock.ToolTipText = "��w"
    frm����r��.Caption = "����r��"
End If

End Sub

Private Sub tree�r�ξ𪬵��c_Click()
Dim �r�� As String
Dim �r�� As String
Dim �s�� As Long

If tree�r�ξ𪬵��c.ListIndex <> -1 Then
   If tree�r�ξ𪬵��c.List(0) <> "" Then
      �r�� = tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.ListIndex)
      �r�� = tree�r�ξ𪬵��c.List(tree�r�ξ𪬵��c.ListIndex)
      �s�� = tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.ListIndex)
    
      �^���ݩ� �r��, �r��, �s��
      �^���c�r�� �r��, �r��, �s��
      If mdi�~�r�r��.txt�r��.font.Name = "�з���" Then �즲�r�� = tree�r�ξ𪬵��c.List(tree�r�ξ𪬵��c.ListIndex)
      
      If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� �r��, �r��, �s��
      If �Ұʲ���r�� Then frm����r��.���J�r�� �r��, �r��, �s��
      If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� �r��, �r��, �s��
      If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� �r��, �r��, �s��
      mdi�~�r�r��.txt���A = ���骬�A�C
   End If
End If

End Sub

Private Sub tree�r�ξ𪬵��c_DragOver(Source As Control, x As Single, y As Single, State As Integer)

tree�r�ξ𪬵��c.OnDragOver x, y, State

End Sub

Private Sub tree�r�ξ𪬵��c_GotFocus()

'tree�r�ξ𪬵��c_Click
�{�ε����N�X = ����r�ڥN�X

End Sub

Private Sub tree�r�ξ𪬵��c_LostFocus()

tree�r�ξ𪬵��c.ListIndex = -1

End Sub

Private Sub tree�r�ξ𪬵��c_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

XCheck = x
YCheck = y

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
'    tree�r�ξ𪬵��c.Drag 1
'End If

End Sub

Public Sub ���J�r��(�t�Φr�� As String, val�r�� As String, �s�� As Long)
Dim �r�Ϊ� As Recordset, �ӳ��r�Ϊ� As Recordset
Dim SQL���z�� As String, ��� As String
Dim i As Integer, �`�� As Long, �r�� As String, �νX As Integer
Dim �r�� As String, �r��s�� As Integer
Dim ���ѽs�� As Long, �p�f�s�� As Integer, ����s�� As Long, �Ұ���s�� As Integer, ���t��r�s�� As Long

If imglock.Tag = 1 Then Exit Sub
If �s�� <= 0 Then Exit Sub

mdi�~�r�r��.txt���A = "�r�θ��J��  ......       " & " , �����_�Ы� Esc ��"
Screen.MousePointer = ccHourglass

�r�� = val�r��

tree�r�ξ𪬵��c.Clear

If �s�� <= 0 Then �s�� = �t�νs��
If �t�Φr�� = "�_�v�j����p�f" Or �t�Φr�� = "�_�v�j���孫��" Then
    �p�f�s�� = �s��
    �p�f�˦r��.Index = "�s��"
    �p�f�˦r��.Seek "=", �s��
    ���ѽs�� = �p�f�˦r��.Fields("���ѽs��")
ElseIf �t�Φr�� = "����|����" Then
    ����s�� = �s��
    �����˦r��.Index = "�s��"
    �����˦r��.Seek "=", �s��
    ���ѽs�� = �����˦r��.Fields("���ѽs��")
ElseIf �t�Φr�� = "����|�Ұ���" Then
    �Ұ���s�� = �s��
    �Ұ����˦r��.Index = "�s��"
    �Ұ����˦r��.Seek "=", �s��
    ���ѽs�� = �Ұ����˦r��.Fields("���ѽs��")
ElseIf �t�Φr�� = "����|���t²����r" Then
    ���t��r�s�� = �s��
    ���t��r�˦r��.Index = "�s��"
    ���t��r�˦r��.Seek "=", �s��
    ���ѽs�� = ���t��r�˦r��.Fields("���ѽs��")
Else
    ���ѽs�� = �s��
End If

Set �˦r�� = �����˦r��
�˦r��.Index = "�s��"
�˦r��.Seek "=", ���ѽs��
�r��s�� = �˦r��.Fields("�r��")
�r�� = �r��}�C(�r��s��)
��� = �˦r��.Fields("�r�X")

SQL���z�� = "SELECT �ո�,First(�νX),�s�� From ����r�� GROUP BY �ո�,�s�� HAVING �s��= " & ���ѽs�� & " order by �ո�"

DoEvents

��ڧǼ� = 0
�`�� = 0

tree�r�ξ𪬵��c.Redraw = False

Set �r�Ϊ� = �t�θ�Ʈw.OpenRecordset(SQL���z��)

If tree�r�ξ𪬵��c.ListCount > 0 Then
   tree�r�ξ𪬵��c.RemoveItem (0)
End If

If Not �r�Ϊ�.EOF Then
   
   �r�Ϊ�.MoveLast
   �`�� = �r�Ϊ�.RecordCount
 
   If �`�� > 1 Or (�`�� = 1 And �r�Ϊ�.Fields(1) <> 1) Then
      tree�r�ξ𪬵��c.AddItem ���
      tree�r�ξ𪬵��c.ItemFontName(0) = �ഫ��ܦr��(�r��)
      tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureOpen
      
      tree�r�ξ𪬵��c.ItemLngValue(0) = ���ѽs��
      If Not IsNull(�˦r��.Fields("�r��")) Then
         tree�r�ξ𪬵��c.ItemTag(0) = �r�θ`�I�аO
      Else
         tree�r�ξ𪬵��c.ItemTag(0) = �c�r���`�I�аO
      End If
      tree�r�ξ𪬵��c.Expand(0) = True
   End If
    
   �r�Ϊ�.MoveFirst
    
   Do Until �r�Ϊ�.EOF
      If ���_ = True Then Exit Do
         
      SQL���z�� = "SELECT * From ����r�� Where �ո�= " & �r�Ϊ�.Fields("�ո�") & " ORDER BY �νX,�s��"
      Set �ӳ��r�Ϊ� = �t�θ�Ʈw.OpenRecordset(SQL���z��)
            
      �ӳ��r�Ϊ�.MoveFirst
      �˦r��.Index = "�s��"

      Do Until �ӳ��r�Ϊ�.EOF
         ��ڧǼ� = ��ڧǼ� + 1

         �˦r��.Seek "=", �ӳ��r�Ϊ�.Fields("�s��")
         If Not �˦r��.NoMatch Then
            If Not IsNull(�˦r��.Fields("�r��")) And �˦r��.Fields("�r��") <> "" Then
               �r�� = �˦r��.Fields("�r��")
            Else
               If Not IsNull(�˦r��.Fields("�r�X")) And �˦r��.Fields("�r�X") <> "" Then
                  �r�� = �˦r��.Fields("�r�X")
               Else
                  �r�� = "��"
               End If
            End If
          End If
   
          If �ӳ��r�Ϊ�.Fields("�νX") <> 1 Then
             tree�r�ξ𪬵��c.AddItem �r��, i
             tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.NewIndex) = �ഫ��ܦr��(�ϰ�r��}�C(�˦r��.Fields("�r��")))
             tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = �˦r��.Fields("�s��")
             tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �r�θ`�I�аO
             If �ӳ��r�Ϊ�.Fields("�νX") = 2 Then
                'tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PicturePalette '²��r
                tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureMark '²��r
             'ElseIf �ӳ��r�Ϊ�.Fields("�νX") > 100 Then   '�νX>100
             '   tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = Image1
             ElseIf �˦r��.Fields("�s���Ÿ�") = 0 Then
                tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureMark  '�r��
             Else
                tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf  '����
             End If
          Else
             'If �`�� > 1 Or (�`�� = 1 And �r�Ϊ�.Fields("�νX") <> 1) Then
             If �`�� > 1 Or (�`�� = 1 And �r�Ϊ�.Fields(1) <> 1) Then
                tree�r�ξ𪬵��c.AddItem �r��, 0
                tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.NewIndex) = �ഫ��ܦr��(�ϰ�r��}�C(�˦r��.Fields("�r��")))
                tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = �˦r��.Fields("�s��")
                tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �r�θ`�I�аO
                tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureClosed
                i = tree�r�ξ𪬵��c.NewIndex
             Else
                tree�r�ξ𪬵��c.AddItem �r��
                tree�r�ξ𪬵��c.ItemFontName(0) = �ഫ��ܦr��(�ϰ�r��}�C(�˦r��.Fields("�r��")))
                tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureOpen
                tree�r�ξ𪬵��c.ItemLngValue(0) = �˦r��.Fields("�s��")
                If Not IsNull(�˦r��.Fields("�r��")) Then
                    tree�r�ξ𪬵��c.ItemTag(0) = �r�θ`�I�аO
                Else
                    tree�r�ξ𪬵��c.ItemTag(0) = �c�r���`�I�аO
                End If

                i = 0
             End If
          End If
          �ӳ��r�Ϊ�.MoveNext
       Loop
       tree�r�ξ𪬵��c.Expand(i) = True
        
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
   tree�r�ξ𪬵��c.AddItem ���
   tree�r�ξ𪬵��c.ItemFontName(0) = �ഫ��ܦr��(�r��)
   tree�r�ξ𪬵��c.ItemLngValue(0) = ���ѽs��
   
   �˦r��.Index = "�s��"
   �˦r��.Seek "=", �s��
   If Not �˦r��.NoMatch Then
      If Not IsNull(�˦r��.Fields("�r��")) Then
         tree�r�ξ𪬵��c.ItemTag(0) = �r�θ`�I�аO
      Else
         tree�r�ξ𪬵��c.ItemTag(0) = �c�r���`�I�аO
      End If

      If �˦r��.Fields("�s���Ÿ�") = 0 Then
         tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureMark
      Else
         tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
      End If
   End If

End If

tree�r�ξ𪬵��c.Redraw = True
���骬�A�C = "���J�����I �@ " & ��ڧǼ� & " �Ӧr��"
mdi�~�r�r��.txt���A = ���骬�A�C
'mdi�~�r�r��.sbar���A�C.Visible = True

Screen.MousePointer = ccDefault

End Sub
