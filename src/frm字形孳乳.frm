VERSION 5.00
Object = "{65996203-3B87-11D4-A21F-00E029189826}#6.9#0"; "TLIST6.OCX"
Begin VB.Form frm�r�δF�� 
   Caption         =   "�����˦r"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "�з���"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm�r�δF��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   8280
   Begin VB.TextBox txt�c�r�� 
      Height          =   360
      HideSelection   =   0   'False
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "��J�@��h�ӳ����A�A��Enter�˦r"
      Top             =   120
      Width           =   4395
   End
   Begin TListProLibCtl.TList tree�r�ξ𪬵��c 
      DragIcon        =   "frm�r�δF��.frx":030A
      Height          =   2652
      Left            =   240
      TabIndex        =   1
      Top             =   684
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
      PictureOpen     =   "frm�r�δF��.frx":074C
      PictureClosed   =   "frm�r�δF��.frx":085E
      PictureLeaf     =   "frm�r�δF��.frx":0970
      PictureMark     =   "frm�r�δF��.frx":0A82
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
      ExchangeSerialNumber=   "frm�r�δF��.frx":0B7C
      DragIconStyle   =   0
      ExchangeDefItemCellDef=   "frm�r�δF��.frx":0BC9
      _ChkCounter     =   -1
      TreeLinesHighlightColor=   -2113929196
      TreeLinesShadowColor=   -2113929200
   End
End
Attribute VB_Name = "frm�r�δF��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private �����N�X As Integer, ���� As String, �r�ڪ� As Recordset, �˦r�� As Recordset
Private �ӳ��t�Φr�� As String
Private �ϰ�r��}�C(0 To �r��Ӽ�) As Variant
Private ���_ As Boolean
Private �`��  As Long, ���� As String
Private XCheck As Single, YCheck As Single

Private Sub Form_Activate()

�{�ε��� = ����
�{�ε����N�X = �r�δF�ťN�X
'�{�ε����N�X = �����N�X
��������r�Τu��C���A �{�ε����N�X
mdi�~�r�r��.txt���A = �F�Ū��A�C

End Sub


Private Sub Form_Load()
Dim i As Integer

�Ұʦr�δF�� = True
If ��lfirst <> 1 Then
   If �w���J�e�� = 0 Then
      If �F��winstate = 0 Then
         frm�r�δF��.Left = �F��left
         frm�r�δF��.Top = �F��top
         frm�r�δF��.Height = �F��height
         frm�r�δF��.Width = �F��width
      Else
         frm�r�δF��.WindowState = �F��winstate
      End If
   End If
Else
   txt�c�r��.Text = "��J�@��h�ӳ����A�A��Enter"
End If

txt�c�r��.SelStart = 0
txt�c�r��.SelLength = Len(txt�c�r��)

If �t�Φr�� = "����" Then
    Set �r�ڪ� = ���Ѧr��
ElseIf �t�Φr�� = "�p�f" Then
    Set �r�ڪ� = �p�f�W��r
    txt�c�r��.FontName = "�_�v�j����p�f"
ElseIf �t�Φr�� = "����" Then
    Set �r�ڪ� = ����r��
    txt�c�r��.FontName = "����|����"
ElseIf �t�Φr�� = "�Ұ���" Then
    Set �r�ڪ� = �Ұ���r��
    txt�c�r��.FontName = "����|�Ұ���"
ElseIf �t�Φr�� = "���t��r" Then
    Set �r�ڪ� = ���t��r�r��
    txt�c�r��.FontName = "����|���t²����r"
End If
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
Me.Tag = �r�δF�ťN�X
tree�r�ξ𪬵��c.AddItem ""
'tree�r�ξ𪬵��c.ListIndex = 0
tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureLeaf
�{�α���N�X = �r�δF��_�˦r���

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
   DoEvents
   ���_ = True
End If

End Sub

Private Sub Form_Resize()
Dim frm���� As Integer

frm���� = Me.ScaleHeight - txt�c�r��.Height - txt�c�r��.Top * 3

If frm���� > 0 Then
   tree�r�ξ𪬵��c.Height = frm����
End If

If (Me.ScaleWidth - tree�r�ξ𪬵��c.Left * 2) > 0 Then
   tree�r�ξ𪬵��c.Width = Me.ScaleWidth - tree�r�ξ𪬵��c.Left * 2
End If

txt�c�r��.Width = tree�r�ξ𪬵��c.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)
'�r�ڪ�.Close
mdi�~�r�r��.mnu_�r�δF��.Enabled = True
�p��{�ε���
�Ұʦr�δF�� = False
End Sub

Private Sub tree�r�ξ𪬵��c_GotFocus()

�{�ε����N�X = �r�δF�ťN�X
�{�α���N�X = �r�δF��_�𪬵��c

End Sub

Private Sub txt�c�r��_DragDrop(Source As Control, x As Single, y As Single)
Dim �r�� As String
Dim ���b�� As String
Dim �k�b�� As String


���b�� = Left(txt�c�r��, txt�c�r��.SelStart)
�k�b�� = Right$(txt�c�r��, Len(txt�c�r��) - txt�c�r��.SelLength - txt�c�r��.SelStart)
 
If TypeOf Source Is ListBox Then
   If Source.ListIndex < 0 Then Exit Sub
   Source.Drag 2       ' End Dragging
   txt�c�r�� = ���b�� & Source.List(Source.ListIndex) & �k�b��
End If

If TypeOf Source Is TList Then
   If Source Is Nothing Then Exit Sub
   Source.Drag 2       ' End Dragging
   Screen.MousePointer = 11
   �r�� = Left(Source, 2)
   'txt�c�r�� = ���b�� & mdi�~�r�r��.txt�r��.Text & �k�b��
   txt�c�r�� = ���b�� & �즲�r�� & �k�b��
   Screen.MousePointer = 0
End If
txt�c�r��.SetFocus
txt�c�r��.SelStart = Len(txt�c�r��)
End Sub

Private Sub txt�c�r��_GotFocus()

�{�ε����N�X = �r�δF�ťN�X
�{�α���N�X = �r�δF��_�˦r���

End Sub

Private Sub txt�c�r��_KeyPress(KeyAscii As Integer)
Dim �զr�� As String
Dim i As Integer

If KeyAscii = vbKeyReturn Then
   Screen.MousePointer = 11
   
   i = 0
   Do While �r��}�C(i) <> ""
      �ϰ�r��}�C(i) = �r��}�C(i)
      i = i + 1
   Loop
   ���_ = False
   ���J�� Trim(txt�c�r��.Text)
   Screen.MousePointer = 0
End If
   
End Sub


Private Sub tree�r�ξ𪬵��c_Expand(ByVal i As Long)

'If tree�r�ξ𪬵��c.ListIndex <> -1 And Not (tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureLeaf And tree�r�ξ𪬵��c.List(0) = "") Then
   If tree�r�ξ𪬵��c.ListCountEx(i) = 1 Then
      If tree�r�ξ𪬵��c.List(i + 1) = "" Then
         Screen.MousePointer = 11
         tree�r�ξ𪬵��c.RemoveItem (i + 1)
         tree�r�ξ𪬵��c.Redraw = False
         ���J�ӳ��𪬵��c i
         tree�r�ξ𪬵��c.Redraw = True
         Screen.MousePointer = 0
      End If
   End If
   tree�r�ξ𪬵��c.Expand(i) = True
'End If

End Sub

Private Sub tree�r�ξ𪬵��c_Click()
Dim �r�� As String
Dim �r�� As String
Dim �s�� As Long

If tree�r�ξ𪬵��c.ListIndex <> -1 Then
   If tree�r�ξ𪬵��c.List(0) <> "" Then
      �r�� = �ഫ�r��(tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.ListIndex))
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

Public Function �غc�զr��(�զr�� As String) As String
Dim �r�νs�����|(100) As Integer, �r�ΰ��|(100) As String
Dim maxid As Integer
Dim i As Integer, j As Integer, temp As Integer, temp2 As String
Dim �c�r�� As String, �Ÿ� As String

If mdi�~�r�r��.mnu_�r�δF�űĥ�SQL�y�k.Checked And �զr�� = "*" Then
    �غc�զr�� = "*"
    Exit Function
End If

If �t�Φr�� = "����" Then
    Set �˦r�� = �����˦r��
    Set �r�ڪ� = ���Ѧr��
ElseIf �t�Φr�� = "�p�f" Then
    Set �˦r�� = �p�f�˦r��
    Set �r�ڪ� = �p�f�W��r
ElseIf �t�Φr�� = "����" Then
    Set �˦r�� = �����˦r��
    Set �r�ڪ� = ����r��
ElseIf �t�Φr�� = "�Ұ���" Then
    Set �˦r�� = �Ұ����˦r��
    Set �r�ڪ� = �Ұ���r��
ElseIf �t�Φr�� = "���t��r" Then
    Set �˦r�� = ���t��r�˦r��
    Set �r�ڪ� = ���t��r�r��
End If
�˦r��.Index = "�r��"
�r�ڪ�.Index = "�r��"
maxid = 0
�Ÿ� = "*"
    
If ���O = "�����" Or ���O = "����ǤG" Then
   �c�r�� = �Ÿ� & �զr�� & �Ÿ�
ElseIf ���O = "�r�ڧ�" Or ���O = "�r�ڧǤG" Then
       If mdi�~�r�r��.mnu_�r�δF�űĥ�SQL�y�k.Checked Then
            �c�r�� = ""
       Else
            �c�r�� = �Ÿ�
       End If
       For i = 1 To Len(�զr��)
           �˦r��.Seek "=", Mid(�զr��, i, 1)
           If Not �˦r��.NoMatch Then
              Do While Not �˦r��.EOF And �˦r��.Fields("�r��") = Mid(�զr��, i, 1)     '�з���
                ' If �˦r��.Fields("�r��") = 0 Then
                    �c�r�� = �c�r�� & �˦r��.Fields(���O)
                    Exit Do
                'End If
                 �˦r��.MoveNext
              Loop
           ElseIf mdi�~�r�r��.mnu_�r�δF�űĥ�SQL�y�k.Checked And �O�_���U�Φr��(Mid(�զr��, i, 1)) Then
                �c�r�� = �c�r�� & Mid(�զr��, i, 1)
           End If
           If Not mdi�~�r�r��.mnu_�r�δF�űĥ�SQL�y�k.Checked Then �c�r�� = �c�r�� & �Ÿ�
       Next
Else
   '���O=�r�ڲ� OR �r�ڲդG
    maxid = 0
    For i = 1 To Len(�զr��)
        �˦r��.Seek "=", Mid(�զr��, i, 1)
        If Not �˦r��.NoMatch Then
           Do While Not �˦r��.EOF
              If Not IsNull(�˦r��.Fields("�r��")) Then
'              If �˦r��.Fields("�r��") = 0 Then
                 If Not IsNull(�˦r��.Fields(���O)) Then
                    For j = 1 To Len(�˦r��.Fields(���O))
                        �r�ΰ��|(maxid) = Mid(�˦r��.Fields(���O), j, 1)
                        maxid = maxid + 1
                    Next j
                 End If
                 Exit Do
              End If
              �˦r��.MoveNext
           Loop
        End If
    Next
 
    For i = 0 To maxid - 1
        �r�ڪ�.Seek "=", �r�ΰ��|(i)
        If Not �r�ڪ�.NoMatch Then
           �r�νs�����|(i) = �r�ڪ�.Fields("�s��")
        End If
    Next
  
    For i = 0 To maxid - 2
        For j = i + 1 To maxid - 1
            If �r�νs�����|(i) > �r�νs�����|(j) Then
               temp = �r�νs�����|(i)
               �r�νs�����|(i) = �r�νs�����|(j)
               �r�νs�����|(j) = temp
               temp2 = �r�ΰ��|(i)
               �r�ΰ��|(i) = �r�ΰ��|(j)
               �r�ΰ��|(j) = temp2
            End If
        Next
    Next
       
    �c�r�� = �Ÿ�
       
    For i = 0 To maxid - 1
        �c�r�� = �c�r�� & �r�ΰ��|(i) & �Ÿ�
    Next
End If

�غc�զr�� = �c�r��
End Function


Private Sub ���J��(�զr�� As String)
Dim �r�Ϊ� As Recordset
Dim SQL���z�� As String
Dim i As Integer, �r�� As Integer
Dim ��� As String
Dim �r�� As String, �r�� As String
Dim �r���� As String
Dim �ʤ��� As Long
Dim wd As String, �Ȧs�զr�� As String, ���� As Integer, times As Integer
Dim �F���`�� As Long, �v�ŦC�X���� As String
Dim �r���z�� As String, �r���զr�r�� As String
Dim �D�r As Boolean, ���� As Boolean, �r����� As String

mdi�~�r�r��.txt���A = "�r�θ��J��  ......       " & " , �����_�Ы� Esc ��"
Screen.MousePointer = ccHourglass

tree�r�ξ𪬵��c.Clear
��� = �զr��

If (Len(�զr��) = 2 And Mid(�զr��, 1, 1) >= "��" And Mid(�զr��, 1, 1) <= "��") Then
    wd = Mid(�զr��, 1, 1)
    If wd = "��" Or wd = "��" Then
       ���� = 2
    ElseIf wd = "��" Or wd = "��" Or wd = "��" Then
       ���� = 3
    ElseIf wd = "��" Or wd = "��" Or wd = "��" Then
       ���� = 4
    End If
    �Ȧs�զr�� = Mid(�զr��, 2, 1)
    For times = 1 To ���� - 1
        �Ȧs�զr�� = �Ȧs�զr�� + �Ȧs�զr��
    Next times
    �զr�� = �Ȧs�զr��
End If

If mdi�~�r�r��.mnu_�r�δF�űĥ�SQL�y�k.Checked = True Then
    If mdi�~�r�r��.mnu_�r�δF�ť]�t���g����.Checked = True Then
        ���O = "�r�ڧǤG"
    Else
        ���O = "�r�ڧ�"
    End If
    GoTo �}�l�غc�զr��
End If

If mdi�~�r�r��.mnu_�r�δF�ť]�t���g����.Checked = True And �t�Φr�� = "����" Then    '�]�t���g
   If Len(�զr��) = 1 Then
     '�ˬd�O�_���ۦ��r��
      ���g�r��.Seek "=", �զr��
      If Not ���g�r��.NoMatch Then
         �զr�� = ���g�r��.Fields("���g")
      End If
      If mdi�~�r�r��.mnu_�r�δF�ųv�ŦC�X��@����.Checked = True Then
         ���O = "����ǤG"
      Else
         ���O = "�r�ڧǤG"
      End If
   Else
      If mdi�~�r�r��.mnu_�r�δF�ſ�ӿ�J���󶶧�.Checked = True Then
         ���O = "�r�ڧǤG"
      Else
         ���O = "�r�ڲդG"
      End If
   End If
Else
   '���]�t���g
   If mdi�~�r�r��.mnu_�r�δF�ųv�ŦC�X��@����.Checked = True And Len(�զr��) = 1 Then
      ���O = "�����"
   Else
      If Len(�զr��) = 1 Or mdi�~�r�r��.mnu_�r�δF�ſ�ӿ�J���󶶧�.Checked = True Then
         ���O = "�r�ڧ�"
      Else
         ���O = "�r�ڲ�"
      End If
   End If
End If

�}�l�غc�զr��:

If �t�Φr�� = "�p�f" Then
    If ���O = "����ǤG" Then ���O = "�����"
    If ���O = "�r�ڧǤG" Then ���O = "�r�ڧ�"
    If ���O = "�r�ڲդG" Then ���O = "�r�ڲ�"
    For i = 1 To Len(�զr��)
        �r�� = Mid(�զr��, i, 1)
        �p�f���g�r��.Seek "=", �r��
        If Not �p�f���g�r��.NoMatch Then Mid(�զr��, i, 1) = �p�f���g�r��.Fields("���g")
    Next i
    ��� = �զr��
End If

If �t�Φr�� = "����" Then
    If ���O = "����ǤG" Then ���O = "�����"
    If ���O = "�r�ڧǤG" Then ���O = "�r�ڧ�"
    If ���O = "�r�ڲդG" Then ���O = "�r�ڲ�"
    For i = 1 To Len(�զr��)
        �r�� = Mid(�զr��, i, 1)
        ���岧�g�r��.Seek "=", �r��
        If Not ���岧�g�r��.NoMatch Then Mid(�զr��, i, 1) = ���岧�g�r��.Fields("���g")
    Next i
    ��� = �զr��
End If

If �t�Φr�� = "�Ұ���" Then
    If ���O = "����ǤG" Then ���O = "�����"
    If ���O = "�r�ڧǤG" Then ���O = "�r�ڧ�"
    If ���O = "�r�ڲդG" Then ���O = "�r�ڲ�"
    For i = 1 To Len(�զr��)
        �r�� = Mid(�զr��, i, 1)
        �Ұ��岧�g�r��.Seek "=", �r��
        If Not �Ұ��岧�g�r��.NoMatch Then Mid(�զr��, i, 1) = �Ұ��岧�g�r��.Fields("���g")
    Next i
    ��� = �զr��
End If

If �t�Φr�� = "���t��r" Then
    If ���O = "����ǤG" Then ���O = "�����"
    If ���O = "�r�ڧǤG" Then ���O = "�r�ڧ�"
    If ���O = "�r�ڲդG" Then ���O = "�r�ڲ�"
    For i = 1 To Len(�զr��)
        �r�� = Mid(�զr��, i, 1)
        ���t��r���g�r��.Seek "=", �r��
        If Not ���t��r���g�r��.NoMatch Then Mid(�զr��, i, 1) = ���t��r���g�r��.Fields("���g")
    Next i
    ��� = �զr��
End If

�զr�� = �غc�զr��(�զr��)

If mdi�~�r�r��.mnu_�`�Φr.Checked = True Then
    �r����� = "�`�Φr"
    �r���z�� = "�`�Φr > 0"
    �r���զr�r�� = "[�զr�r��(�`�Φr)]"
ElseIf mdi�~�r�r��.mnu_Big5.Checked = True Then
    �r����� = "Big5"
    �r���z�� = "�s�� > 0 and �s�� <= 13053"
    �r���զr�r�� = "[�զr�r��(Big5)]"
ElseIf mdi�~�r�r��.mnu_²�Ʀr�`��.Checked = True Then
    �r����� = "²�Ʀr"
    �r���z�� = "²�Ʀr > 0"
    �r���զr�r�� = "[�զr�r��(²�Ʀr)]"
ElseIf mdi�~�r�r��.mnu_�~�y�j�r��.Checked = True Then
    �r����� = "�~�y�j�r��"
    �r���z�� = "�~�y�j�r�� > 0"
    �r���զr�r�� = "[�զr�r��(�~�y�j�r��)]"
ElseIf mdi�~�r�r��.mnu_����ϧΤ�r.Checked = True Then
    �r����� = "�s��"
    �r���z�� = "�s��>0 and �ϧΤ�r>0"
    �r���զr�r�� = "[�զr�r��(�ϧΤ�r)]"
Else
    �r����� = "�s��"
    �r���z�� = "�s�� > 0"
    �r���զr�r�� = "�զr�r��"
End If


If ��l�v�ŦC�X = 1 And Mid(���O, 1, 2) = "����" Then
   '�v�ŦC�X���� = " ( " & �r���զr�r�� & " > 1 or ( " & �r���զr�r�� & " = 1 and " & �r���z�� & " ) )"
   �v�ŦC�X���� = " ( " & �r���զr�r�� & " > 0 )"
Else
   �v�ŦC�X���� = " (  " & �r���z�� & " ) "
End If

If Mid(���O, 1, 2) = "�r��" Then
   If �t�Φr�� = "����" Then
        SQL���z�� = "SELECT * From �˦r�� Where �s��>0 and " & �v�ŦC�X���� & " and " & ���O & " Like '" & �զr�� & "'  ORDER BY �������e�Ƨ�,�զr�r�� DESC "
   Else
        SQL���z�� = "SELECT * From �˦r�� Where �s��>0 and " & �v�ŦC�X���� & " and " & ���O & " Like '" & �զr�� & "'  ORDER BY �s��,�զr�r�� DESC "
   End If
Else
   If �t�Φr�� = "����" Then
        SQL���z�� = "SELECT * From �˦r�� Where (�r��='" & ��� & "') or ( " & �v�ŦC�X���� & " And " & ���O & " Like '" & �զr�� & "')  ORDER BY �������e�Ƨ�,�զr�r�� DESC "
   Else
        SQL���z�� = "SELECT * From �˦r�� Where (�r��='" & ��� & "') or ( " & �v�ŦC�X���� & " And " & ���O & " Like '" & �զr�� & "')  ORDER BY �s��,�զr�r�� DESC "
   End If
End If

DoEvents

��ڧǼ� = 0
�F���`�� = 0
�`�� = 0

'If �զr�� <> "" And �զr�� <> "*" And �զr�� <> "**" Then
If �զr�� <> "" And �զr�� <> "**" Then
   tree�r�ξ𪬵��c.Redraw = False
   
    If �t�Φr�� = "����" Then
        Set �r�Ϊ� = �t�θ�Ʈw.OpenRecordset(SQL���z��)
    ElseIf �t�Φr�� = "�p�f" Then
        Set �r�Ϊ� = �p�f��Ʈw.OpenRecordset(SQL���z��)
    ElseIf �t�Φr�� = "����" Then
        Set �r�Ϊ� = �����Ʈw.OpenRecordset(SQL���z��)
    ElseIf �t�Φr�� = "�Ұ���" Then
        Set �r�Ϊ� = �Ұ����Ʈw.OpenRecordset(SQL���z��)
    ElseIf �t�Φr�� = "���t��r" Then
        Set �r�Ϊ� = ���t��r��Ʈw.OpenRecordset(SQL���z��)
    End If
   
   If Not �r�Ϊ�.EOF Then
   
      �r�Ϊ�.MoveLast
      �`�� = �r�Ϊ�.RecordCount
            
      If �զr�� <> "*" Then
        tree�r�ξ𪬵��c.AddItem ���
      Else
        tree�r�ξ𪬵��c.AddItem "��"
      End If
      If �t�Φr�� = "����" Then
        tree�r�ξ𪬵��c.ItemFontName(0) = ��ܦr�� '�{�Φr��
      ElseIf �t�Φr�� = "�p�f" Then
        tree�r�ξ𪬵��c.ItemFontName(0) = "�_�v�j����p�f"
      ElseIf �t�Φr�� = "����" Then
        tree�r�ξ𪬵��c.ItemFontName(0) = "����|����"
      ElseIf �t�Φr�� = "�Ұ���" Then
        tree�r�ξ𪬵��c.ItemFontName(0) = "����|�Ұ���"
      ElseIf �t�Φr�� = "���t��r" Then
        tree�r�ξ𪬵��c.ItemFontName(0) = "����|���t²����r"
      End If
      
      tree�r�ξ𪬵��c.ItemLngValue(0) = -999999
      If Len(���) = 1 Then
         tree�r�ξ𪬵��c.ItemTag(0) = �r�θ`�I�аO
      Else
         tree�r�ξ𪬵��c.ItemTag(0) = ��L�`�I�аO
      End If
        
      tree�r�ξ𪬵��c.Expand(0) = True
    
      �r�Ϊ�.MoveFirst

      Do Until �r�Ϊ�.EOF
         If ���_ = True Then Exit Do
         
         �ʤ��� = ��ڧǼ� / �`�� * 100
         mdi�~�r�r��.txt���A = "�r�θ��J�w���� " & �ʤ��� & " % , �����_�Ы� Esc ��"
         
         If Not IsNull(�r�Ϊ�.Fields("�r��")) And �r�Ϊ�.Fields("�r��") <> "" Then
            �r�� = �r�Ϊ�.Fields("�r��")
         Else
            If Not IsNull(�r�Ϊ�.Fields("�r�X")) And �r�Ϊ�.Fields("�r�X") <> "" Then
               �r�� = �r�Ϊ�.Fields("�r�X")
            Else
               �r�� = "��"
            End If
         End If
           
         �r�� = �r�Ϊ�.Fields(�r���զr�r��)
         If �r����� <> "Big5" Then
            If IsNull(�r�Ϊ�.Fields(�r�����)) Then
                �D�r = True
            ElseIf �r�Ϊ�.Fields(�r�����) = 0 Then
                �D�r = True
            Else
                �D�r = False
            End If
         Else
            If �r�Ϊ�.Fields("�s��") > 0 And �r�Ϊ�.Fields("�s��") <= 13053 Then
                �D�r = False
            ElseIf �r�Ϊ�.Fields("�s��") > 13053 Then
                �D�r = True
            Else
                �D�r = False
            End If
         End If
         
         If �r�� > 1 Then
            ���� = True
         ElseIf �r�� = 1 And �D�r Then
            ���� = True
         Else
            ���� = False
         End If
                    
         If Not (�ϰ�r��}�C(�r�Ϊ�.Fields("�r��")) = �ഫ�r��(tree�r�ξ𪬵��c.ItemFontName(0)) And �r�� = tree�r�ξ𪬵��c.List(0)) Then
            If Not ((���O = "�����" Or ���O = "����ǤG") And (�r�Ϊ�.Fields("�s���Ÿ�") = 0)) Then
               ��ڧǼ� = ��ڧǼ� + 1
               If ���� Then
                  tree�r�ξ𪬵��c.AddItem �r��, 0
                  tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.NewIndex) = �ഫ��ܦr��(�ϰ�r��}�C(�r�Ϊ�.Fields("�r��")))
                  tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = �r�Ϊ�.Fields("�s��")
                  If Not IsNull(�r�Ϊ�.Fields("�r��")) Then
                     tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �r�θ`�I�аO
                  Else
                     tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �c�r���`�I�аO
                  End If

                  If Mid(���O, 1, 2) = "�r��" Then
                     tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
                  Else
                     If �r�Ϊ�.Fields("�s���Ÿ�") = 0 Then
                        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
                     Else
                        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureClosed
                        tree�r�ξ𪬵��c.AddItem "", tree�r�ξ𪬵��c.NewIndex
                     End If
                  End If
               Else
                  tree�r�ξ𪬵��c.AddItem �r��, 0
                  tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
                  tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.NewIndex) = �ഫ��ܦr��(�ϰ�r��}�C(�r�Ϊ�.Fields("�r��")))
                  tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = �r�Ϊ�.Fields("�s��")
                  If Not IsNull(�r�Ϊ�.Fields("�r��")) Then
                     tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �r�θ`�I�аO
                  Else
                     tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �c�r���`�I�аO
                  End If

               End If
             End If
         Else
            If ���O = "�����" Or �t�Φr�� <> "����" Then
                �F���`�� = �r�Ϊ�.Fields("�զr�r��")
            ElseIf ���O = "����ǤG" Then
                �F���`�� = �r�Ϊ�.Fields("�զr�r�ƤG")
            Else
                ��ڧǼ� = ��ڧǼ� + 1
            End If
            
            If ���� Then
               tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureOpen
            Else
               If Len(���) = 1 Then
                  tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureLeaf
               Else
                  tree�r�ξ𪬵��c.Image(0) = tree�r�ξ𪬵��c.PictureOpen
               End If
            End If
            tree�r�ξ𪬵��c.ItemFontName(0) = �ഫ��ܦr��(�ϰ�r��}�C(�r�Ϊ�.Fields("�r��")))
            tree�r�ξ𪬵��c.ItemLngValue(0) = �r�Ϊ�.Fields("�s��")
            If Len(���) = 1 Then
               tree�r�ξ𪬵��c.ItemTag(0) = �r�θ`�I�аO
            Else
               tree�r�ξ𪬵��c.ItemTag(0) = ��L�`�I�аO
            End If
         End If
         
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
End If

If �`�� > 0 Then
    If tree�r�ξ𪬵��c.ItemLngValue(0) = -999999 And Len(tree�r�ξ𪬵��c.List(0)) = 1 Then

        If �t�Φr�� = "�p�f" Then
            Set �˦r�� = �p�f�˦r��
        ElseIf �t�Φr�� = "����|����" Then
            Set �˦r�� = �����˦r��
        ElseIf �t�Φr�� = "����|�Ұ���" Then
            Set �˦r�� = �Ұ����˦r��
        ElseIf �t�Φr�� = "����|���t²����r" Then
            Set �˦r�� = ���t��r�˦r��
        Else
            Set �˦r�� = �����˦r��
        End If
   
        �˦r��.Index = "�r��"
        �˦r��.Seek "=", tree�r�ξ𪬵��c.List(0)
        If Not �˦r��.NoMatch Then
            If �˦r��.Fields("�s��") = 0 And �t�Φr�� = "�p�f" Then
                tree�r�ξ𪬵��c.ItemFontName(0) = "�з���"
                tree�r�ξ𪬵��c.ItemLngValue(0) = �˦r��.Fields("���ѽs��")
            Else
                tree�r�ξ𪬵��c.ItemLngValue(0) = �˦r��.Fields("�s��")
            End If
        End If
    End If

End If

If Mid(���O, 1, 3) <> "�����" Then �F���`�� = ��ڧǼ�

tree�r�ξ𪬵��c.Redraw = True
�F�Ū��A�C = "���J�����I �@ " & �`�� & " �Ӧr��(����)"
mdi�~�r�r��.txt���A = �F�Ū��A�C
'mdi�~�r�r��.sbar���A�C.Visible = True

�ӳ��t�Φr�� = �t�Φr��

Screen.MousePointer = ccDefault

End Sub

Public Sub ���J�ӳ��𪬵��c(i As Long)
Dim �ӳ��r�Ϊ� As Recordset
Dim �զr�� As String, �r�� As String
Dim SQL���z�� As String, �ӳ��`�� As Long
Dim �r���� As String, �v�ŦC�X���� As String
Dim �s�� As Long, �r�� As Integer
Dim �r���z�� As String, �r���զr�r�� As String
Dim �D�r As Boolean, ���� As Boolean, �r����� As String


If mdi�~�r�r��.mnu_�r�δF�ť]�t���g����.Checked = True Then    '�]�t���g
   If Len(tree�r�ξ𪬵��c.List(i)) = 1 Then
      '�ˬd�O�_�����g�r��
      ���g�r��.Seek "=", tree�r�ξ𪬵��c.List(i)
      If Not ���g�r��.NoMatch Then
         �զr�� = ���g�r��.Fields("���g")
      Else
         �զr�� = tree�r�ξ𪬵��c.List(i)
      End If
      If mdi�~�r�r��.mnu_�r�δF�ųv�ŦC�X��@����.Checked = True Then
         ���O = "����ǤG"
      Else
         ���O = "�r�ڧǤG"
      End If
   Else
      ���O = "�r�ڲդG"
   End If
Else
   '���]�t���g
   �զr�� = tree�r�ξ𪬵��c.List(i)
   If mdi�~�r�r��.mnu_�r�δF�ųv�ŦC�X��@����.Checked = True And Len(�զr��) = 1 Then
      ���O = "�����"
   Else
      If Len(�զr��) = 1 Then
         ���O = "�r�ڧ�"
      Else
         ���O = "�r�ڲ�"
      End If
   End If
End If

If �ӳ��t�Φr�� = "�p�f" Or �ӳ��t�Φr�� = "����" Or �ӳ��t�Φr�� = "�Ұ���" Or �ӳ��t�Φr�� = "���t��r" Then
    If ���O = "����ǤG" Then ���O = "�����"
    If ���O = "�r�ڧǤG" Then ���O = "�r�ڧ�"
    If ���O = "�r�ڲդG" Then ���O = "�r�ڲ�"
    'For i = 1 To Len(�զr��)
    '    �r�� = Mid(�զr��, i, 1)
    '    �p�f���g�r��.Seek "=", �r��
    '    If Not �p�f���g�r��.NoMatch Then Mid(�զr��, i, 1) = �p�f���g�r��.Fields("���g")
    'Next i
    '��� = �զr��
End If

�զr�� = �غc�զr��(�զr��)

'If ��l�v�ŦC�X = 1 Then
'   �v�ŦC�X���� = " ( �զr�r�� > 1 or ( �զr�r�� =1 and �r�W >= " & Trim(��l�r�W) & " ) )"
'Else
'   �v�ŦC�X���� = " ( �r�W >= " & Trim(��l�r�W) & " ) "
'End If

If mdi�~�r�r��.mnu_�`�Φr.Checked = True Then
    �r����� = "�`�Φr"
    �r���z�� = "�`�Φr > 0"
    �r���զr�r�� = "[�զr�r��(�`�Φr)]"
ElseIf mdi�~�r�r��.mnu_Big5.Checked = True Then
    �r����� = "Big5"
    �r���z�� = "�s�� > 0 and �s�� <= 13053"
    �r���զr�r�� = "[�զr�r��(Big5)]"
ElseIf mdi�~�r�r��.mnu_²�Ʀr�`��.Checked = True Then
    �r����� = "²�Ʀr"
    �r���z�� = "²�Ʀr > 0"
    �r���զr�r�� = "[�զr�r��(²�Ʀr)]"
ElseIf mdi�~�r�r��.mnu_�~�y�j�r��.Checked = True Then
    �r����� = "�~�y�j�r��"
    �r���z�� = "�~�y�j�r�� > 0"
    �r���զr�r�� = "[�զr�r��(�~�y�j�r��)]"
Else
    �r����� = "�s��"
    �r���z�� = "�s�� > 0"
    �r���զr�r�� = "�զr�r��"
End If


If ��l�v�ŦC�X = 1 And Mid(���O, 1, 2) = "����" Then
   '�v�ŦC�X���� = " ( " & �r���զr�r�� & " > 1 or ( " & �r���զr�r�� & " = 1 and " & �r���z�� & " ) )"
   �v�ŦC�X���� = " ( " & �r���զr�r�� & " > 0 )"
Else
   �v�ŦC�X���� = " (  " & �r���z�� & " ) "
End If

If Mid(���O, 1, 2) = "�r��" Then
    If �ӳ��t�Φr�� = "����" Then
        SQL���z�� = "SELECT * From �˦r�� Where �s��>0 and " & �v�ŦC�X���� & " and " & ���O & " Like '" & �զr�� & "'  ORDER BY �������e�Ƨ�,�զr�r�� DESC "
    Else
        SQL���z�� = "SELECT * From �˦r�� Where �s��>0 and " & �v�ŦC�X���� & " and " & ���O & " Like '" & �զr�� & "'  ORDER BY �s��,�զr�r�� DESC "
    End If
Else
    If �ӳ��t�Φr�� = "����" Then
        SQL���z�� = "SELECT * From �˦r�� Where (�r��='" & tree�r�ξ𪬵��c.List(i) & "') or ( " & �v�ŦC�X���� & " And " & ���O & " Like '" & �զr�� & "')  ORDER BY �������e�Ƨ�,�զr�r�� DESC "
    Else
        SQL���z�� = "SELECT * From �˦r�� Where (�r��='" & tree�r�ξ𪬵��c.List(i) & "') or ( " & �v�ŦC�X���� & " And " & ���O & " Like '" & �զr�� & "')  ORDER BY �s��,�զr�r�� DESC "
    End If
End If

If �ӳ��t�Φr�� = "����" Then
    Set �ӳ��r�Ϊ� = �t�θ�Ʈw.OpenRecordset(SQL���z��)
ElseIf �ӳ��t�Φr�� = "�p�f" Then
    Set �ӳ��r�Ϊ� = �p�f��Ʈw.OpenRecordset(SQL���z��)
ElseIf �ӳ��t�Φr�� = "����" Then
    Set �ӳ��r�Ϊ� = �����Ʈw.OpenRecordset(SQL���z��)
ElseIf �ӳ��t�Φr�� = "�Ұ���" Then
    Set �ӳ��r�Ϊ� = �Ұ����Ʈw.OpenRecordset(SQL���z��)
ElseIf �ӳ��t�Φr�� = "���t��r" Then
    Set �ӳ��r�Ϊ� = ���t��r��Ʈw.OpenRecordset(SQL���z��)
End If
   
�ӳ��r�Ϊ�.MoveLast
�ӳ��`�� = �ӳ��r�Ϊ�.RecordCount

�ӳ��r�Ϊ�.MoveFirst

   If Not �ӳ��r�Ϊ�.EOF Then
      
      Do Until �ӳ��r�Ϊ�.EOF
        
         If Not IsNull(�ӳ��r�Ϊ�.Fields("�r��")) Then
            �r�� = �ӳ��r�Ϊ�.Fields("�r��")
         Else
            If Not IsNull(�ӳ��r�Ϊ�.Fields("�r�X")) Then
              �r�� = �ӳ��r�Ϊ�.Fields("�r�X")
            Else
              �r�� = "��"
            End If
         End If
         
         '�r�� = �ӳ��r�Ϊ�.Fields("�զr�r��")
         �r�� = �ӳ��r�Ϊ�.Fields(�r���զr�r��)
         If �r����� <> "Big5" Then
            If IsNull(�ӳ��r�Ϊ�.Fields(�r�����)) Then
                �D�r = True
            ElseIf �ӳ��r�Ϊ�.Fields(�r�����) = 0 Then
                �D�r = True
            Else
                �D�r = False
            End If
         Else
            If �ӳ��r�Ϊ�.Fields("�s��") > 0 And �ӳ��r�Ϊ�.Fields("�s��") <= 13053 Then
                �D�r = False
            ElseIf �ӳ��r�Ϊ�.Fields("�s��") > 13053 Then
                �D�r = True
            Else
                �D�r = False
            End If
         End If
         
         If �r�� > 1 Then
            ���� = True
         ElseIf �r�� = 1 And �D�r Then
            ���� = True
         Else
            ���� = False
         End If
         
         If Not (�ϰ�r��}�C(�ӳ��r�Ϊ�.Fields("�r��")) = �ഫ�r��(tree�r�ξ𪬵��c.ItemFontName(i)) And �r�� = tree�r�ξ𪬵��c.List(i)) Then
            If Not ((���O = "�����" Or ���O = "����ǤG") And (�ӳ��r�Ϊ�.Fields("�s���Ÿ�") = 0)) Then
               If ���� Then
                  tree�r�ξ𪬵��c.AddItem �r��, i
                  tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.NewIndex) = �ഫ��ܦr��(�ϰ�r��}�C(�ӳ��r�Ϊ�.Fields("�r��")))
                  tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = �ӳ��r�Ϊ�.Fields("�s��")
                  If Not IsNull(�ӳ��r�Ϊ�.Fields("�r��")) Then
                     tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �r�θ`�I�аO
                  Else
                     tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �c�r���`�I�аO
                  End If
                  
                  If Mid(���O, 1, 2) = "�r��" Then
                     tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
                  Else
                     If �ӳ��r�Ϊ�.Fields("�s���Ÿ�") = 0 Then
                        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
                     Else
                        tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureClosed
                        tree�r�ξ𪬵��c.AddItem "", tree�r�ξ𪬵��c.NewIndex
                     End If
                  End If
               Else
                  tree�r�ξ𪬵��c.AddItem �r��, i
                  tree�r�ξ𪬵��c.Image(tree�r�ξ𪬵��c.NewIndex) = tree�r�ξ𪬵��c.PictureLeaf
                  tree�r�ξ𪬵��c.ItemFontName(tree�r�ξ𪬵��c.NewIndex) = �ഫ��ܦr��(�ϰ�r��}�C(�ӳ��r�Ϊ�.Fields("�r��")))
                  tree�r�ξ𪬵��c.ItemLngValue(tree�r�ξ𪬵��c.NewIndex) = �ӳ��r�Ϊ�.Fields("�s��")
                  If Not IsNull(�ӳ��r�Ϊ�.Fields("�r��")) Then
                     tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �r�θ`�I�аO
                  Else
                     tree�r�ξ𪬵��c.ItemTag(tree�r�ξ𪬵��c.NewIndex) = �c�r���`�I�аO
                  End If
               End If
            End If
         End If
         �ӳ��r�Ϊ�.MoveNext
      Loop
   End If
  
   �F�Ū��A�C = "���J�����I �@ " & �ӳ��`�� & " �Ӧr��(����)"
   mdi�~�r�r��.txt���A = �F�Ū��A�C
   'mdi�~�r�r��.sbar���A�C.Visible = True

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

Public Sub �C�X��w�r�������Ҧ��r��()

Dim ���� As String, LikeSQL As Boolean

���� = txt�c�r��
LikeSQL = mdi�~�r�r��.mnu_�r�δF�űĥ�SQL�y�k.Checked

txt�c�r�� = "*"
mdi�~�r�r��.mnu_�r�δF�űĥ�SQL�y�k.Checked = True
txt�c�r��_KeyPress (vbKeyReturn)

txt�c�r�� = ����
mdi�~�r�r��.mnu_�r�δF�űĥ�SQL�y�k.Checked = LikeSQL

End Sub
