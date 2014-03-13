VERSION 5.00
Object = "{65996203-3B87-11D4-A21F-00E029189826}#6.9#0"; "TLIST6.OCX"
Begin VB.Form frm¦r§Î´F¨Å 
   Caption         =   "³¡¥óÀË¦r"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "¼Ğ·¢Åé"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm¦r§Î´F¨Å.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   8280
   Begin VB.TextBox txtºc¦r¦¡ 
      Height          =   360
      HideSelection   =   0   'False
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "¿é¤J¤@¨ì¦h­Ó³¡¥ó«á¡A¦A«öEnterÀË¦r"
      Top             =   120
      Width           =   4395
   End
   Begin TListProLibCtl.TList tree¦r§Î¾ğª¬µ²ºc 
      DragIcon        =   "frm¦r§Î´F¨Å.frx":030A
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
      PictureOpen     =   "frm¦r§Î´F¨Å.frx":074C
      PictureClosed   =   "frm¦r§Î´F¨Å.frx":085E
      PictureLeaf     =   "frm¦r§Î´F¨Å.frx":0970
      PictureMark     =   "frm¦r§Î´F¨Å.frx":0A82
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
      ExchangeSerialNumber=   "frm¦r§Î´F¨Å.frx":0B7C
      DragIconStyle   =   0
      ExchangeDefItemCellDef=   "frm¦r§Î´F¨Å.frx":0BC9
      _ChkCounter     =   -1
      TreeLinesHighlightColor=   -2113929196
      TreeLinesShadowColor=   -2113929200
   End
End
Attribute VB_Name = "frm¦r§Î´F¨Å"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private µøµ¡¥N½X As Integer, µøµ¡ As String, ¦r®Úªí As Recordset, ÀË¦rªí As Recordset
Private ²Ó³¡¨t²Î¦rÅé As String
Private °Ï°ì¦rÅé°}¦C(0 To ¦rÅé­Ó¼Æ) As Variant
Private ¤¤Â_ As Boolean
Private Á`¼Æ  As Long, ³¡¼Æ As String
Private XCheck As Single, YCheck As Single

Private Sub Form_Activate()

²{¥Îµøµ¡ = µøµ¡
²{¥Îµøµ¡¥N½X = ¦r§Î´F¨Å¥N½X
'²{¥Îµøµ¡¥N½X = µøµ¡¥N½X
¤Á´«¿ï¨ú¦r§Î¤u¨ã¦Cª¬ºA ²{¥Îµøµ¡¥N½X
mdiº~¦r¦r§Î.txtª¬ºA = ´F¨Åª¬ºA¦C

End Sub


Private Sub Form_Load()
Dim i As Integer

±Ò°Ê¦r§Î´F¨Å = True
If ªì©lfirst <> 1 Then
   If ¤w¸ü¤Jµe­± = 0 Then
      If ´F¨Åwinstate = 0 Then
         frm¦r§Î´F¨Å.Left = ´F¨Åleft
         frm¦r§Î´F¨Å.Top = ´F¨Åtop
         frm¦r§Î´F¨Å.Height = ´F¨Åheight
         frm¦r§Î´F¨Å.Width = ´F¨Åwidth
      Else
         frm¦r§Î´F¨Å.WindowState = ´F¨Åwinstate
      End If
   End If
Else
   txtºc¦r¦¡.Text = "¿é¤J¤@¨ì¦h­Ó³¡¥ó«á¡A¦A«öEnter"
End If

txtºc¦r¦¡.SelStart = 0
txtºc¦r¦¡.SelLength = Len(txtºc¦r¦¡)

If ¨t²Î¦rÅé = "·¢®Ñ" Then
    Set ¦r®Úªí = ·¢®Ñ¦r®Ú
ElseIf ¨t²Î¦rÅé = "¤p½f" Then
    Set ¦r®Úªí = ¤p½f¿WÅé¦r
    txtºc¦r¦¡.FontName = "¥_®v¤j»¡¤å¤p½f"
ElseIf ¨t²Î¦rÅé = "ª÷¤å" Then
    Set ¦r®Úªí = ª÷¤å¦r®Ú
    txtºc¦r¦¡.FontName = "¤¤¬ã°|ª÷¤å"
ElseIf ¨t²Î¦rÅé = "¥Ò°©¤å" Then
    Set ¦r®Úªí = ¥Ò°©¤å¦r®Ú
    txtºc¦r¦¡.FontName = "¤¤¬ã°|¥Ò°©¤å"
ElseIf ¨t²Î¦rÅé = "·¡¨t¤å¦r" Then
    Set ¦r®Úªí = ·¡¨t¤å¦r¦r®Ú
    txtºc¦r¦¡.FontName = "¤¤¬ã°|·¡¨tÂ²©­¤å¦r"
End If
¦r®Úªí.Index = "¦r§Î"

i = 0
Do While ¦rÅé°}¦C(i) <> ""
   °Ï°ì¦rÅé°}¦C(i) = ¦rÅé°}¦C(i)
   i = i + 1
Loop

tree¦r§Î¾ğª¬µ²ºc.FontSize = CInt(Åã¥Ü¦r«¬¤j¤p)
µøµ¡¥N½X = ¦@¥Îµøµ¡¥N½X
µøµ¡ = ¦@¥Îµøµ¡(¦@¥Îµøµ¡¥N½X)
'Me.Tag = ¦@¥Îµøµ¡¥N½X
Me.Tag = ¦r§Î´F¨Å¥N½X
tree¦r§Î¾ğª¬µ²ºc.AddItem ""
'tree¦r§Î¾ğª¬µ²ºc.ListIndex = 0
tree¦r§Î¾ğª¬µ²ºc.Image(0) = tree¦r§Î¾ğª¬µ²ºc.PictureLeaf
²{¥Î±±¨î¶µ¥N½X = ¦r§Î´F¨Å_ÀË¦r¤è¶ô

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
   DoEvents
   ¤¤Â_ = True
End If

End Sub

Private Sub Form_Resize()
Dim frm°ª«× As Integer

frm°ª«× = Me.ScaleHeight - txtºc¦r¦¡.Height - txtºc¦r¦¡.Top * 3

If frm°ª«× > 0 Then
   tree¦r§Î¾ğª¬µ²ºc.Height = frm°ª«×
End If

If (Me.ScaleWidth - tree¦r§Î¾ğª¬µ²ºc.Left * 2) > 0 Then
   tree¦r§Î¾ğª¬µ²ºc.Width = Me.ScaleWidth - tree¦r§Î¾ğª¬µ²ºc.Left * 2
End If

txtºc¦r¦¡.Width = tree¦r§Î¾ğª¬µ²ºc.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)
'¦r®Úªí.Close
mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å.Enabled = True
­pºâ²{¥Îµøµ¡
±Ò°Ê¦r§Î´F¨Å = False
End Sub

Private Sub tree¦r§Î¾ğª¬µ²ºc_GotFocus()

²{¥Îµøµ¡¥N½X = ¦r§Î´F¨Å¥N½X
²{¥Î±±¨î¶µ¥N½X = ¦r§Î´F¨Å_¾ğª¬µ²ºc

End Sub

Private Sub txtºc¦r¦¡_DragDrop(Source As Control, x As Single, y As Single)
Dim ¦r§Î As String
Dim ¥ª¥b³¡ As String
Dim ¥k¥b³¡ As String


¥ª¥b³¡ = Left(txtºc¦r¦¡, txtºc¦r¦¡.SelStart)
¥k¥b³¡ = Right$(txtºc¦r¦¡, Len(txtºc¦r¦¡) - txtºc¦r¦¡.SelLength - txtºc¦r¦¡.SelStart)
 
If TypeOf Source Is ListBox Then
   If Source.ListIndex < 0 Then Exit Sub
   Source.Drag 2       ' End Dragging
   txtºc¦r¦¡ = ¥ª¥b³¡ & Source.List(Source.ListIndex) & ¥k¥b³¡
End If

If TypeOf Source Is TList Then
   If Source Is Nothing Then Exit Sub
   Source.Drag 2       ' End Dragging
   Screen.MousePointer = 11
   ¦r§Î = Left(Source, 2)
   'txtºc¦r¦¡ = ¥ª¥b³¡ & mdiº~¦r¦r§Î.txt¦r§Î.Text & ¥k¥b³¡
   txtºc¦r¦¡ = ¥ª¥b³¡ & ©ì¦²¦r¦ê & ¥k¥b³¡
   Screen.MousePointer = 0
End If
txtºc¦r¦¡.SetFocus
txtºc¦r¦¡.SelStart = Len(txtºc¦r¦¡)
End Sub

Private Sub txtºc¦r¦¡_GotFocus()

²{¥Îµøµ¡¥N½X = ¦r§Î´F¨Å¥N½X
²{¥Î±±¨î¶µ¥N½X = ¦r§Î´F¨Å_ÀË¦r¤è¶ô

End Sub

Private Sub txtºc¦r¦¡_KeyPress(KeyAscii As Integer)
Dim ²Õ¦r¦¡ As String
Dim i As Integer

If KeyAscii = vbKeyReturn Then
   Screen.MousePointer = 11
   
   i = 0
   Do While ¦rÅé°}¦C(i) <> ""
      °Ï°ì¦rÅé°}¦C(i) = ¦rÅé°}¦C(i)
      i = i + 1
   Loop
   ¤¤Â_ = False
   ¸ü¤J¾ğª¬ Trim(txtºc¦r¦¡.Text)
   Screen.MousePointer = 0
End If
   
End Sub


Private Sub tree¦r§Î¾ğª¬µ²ºc_Expand(ByVal i As Long)

'If tree¦r§Î¾ğª¬µ²ºc.ListIndex <> -1 And Not (tree¦r§Î¾ğª¬µ²ºc.Image(0) = tree¦r§Î¾ğª¬µ²ºc.PictureLeaf And tree¦r§Î¾ğª¬µ²ºc.List(0) = "") Then
   If tree¦r§Î¾ğª¬µ²ºc.ListCountEx(i) = 1 Then
      If tree¦r§Î¾ğª¬µ²ºc.List(i + 1) = "" Then
         Screen.MousePointer = 11
         tree¦r§Î¾ğª¬µ²ºc.RemoveItem (i + 1)
         tree¦r§Î¾ğª¬µ²ºc.Redraw = False
         ¸ü¤J²Ó³¡¾ğª¬µ²ºc i
         tree¦r§Î¾ğª¬µ²ºc.Redraw = True
         Screen.MousePointer = 0
      End If
   End If
   tree¦r§Î¾ğª¬µ²ºc.Expand(i) = True
'End If

End Sub

Private Sub tree¦r§Î¾ğª¬µ²ºc_Click()
Dim ¦rÅé As String
Dim ¦r§Î As String
Dim ½s¸¹ As Long

If tree¦r§Î¾ğª¬µ²ºc.ListIndex <> -1 Then
   If tree¦r§Î¾ğª¬µ²ºc.List(0) <> "" Then
      ¦rÅé = Âà´«¦rÅé(tree¦r§Î¾ğª¬µ²ºc.ItemFontName(tree¦r§Î¾ğª¬µ²ºc.ListIndex))
      ¦r§Î = tree¦r§Î¾ğª¬µ²ºc.List(tree¦r§Î¾ğª¬µ²ºc.ListIndex)
      If Len(¦r§Î) = 1 Then
      ½s¸¹ = tree¦r§Î¾ğª¬µ²ºc.ItemLngValue(tree¦r§Î¾ğª¬µ²ºc.ListIndex)
      
      Â^¨úÄİ©Ê ¦rÅé, ¦r§Î, ½s¸¹
      Â^¨úºc¦r¦¡ ¦rÅé, ¦r§Î, ½s¸¹
      If mdiº~¦r¦r§Î.txt¦r§Î.font.Name = "¼Ğ·¢Åé" Then ©ì¦²¦r¦ê = tree¦r§Î¾ğª¬µ²ºc.List(tree¦r§Î¾ğª¬µ²ºc.ListIndex)
      
      If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î ¦rÅé, ¦r§Î, ½s¸¹
      If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î ¦rÅé, ¦r§Î, ½s¸¹
      If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î ¦rÅé, ¦r§Î, ½s¸¹
      If ±Ò°Ê¦r§Î¯Á¤Ş Then frm¦r§Î¯Á¤Ş.¸ü¤J¦r§Î ¦rÅé, ¦r§Î, ½s¸¹
      If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î ¦rÅé, ¦r§Î, ½s¸¹
      mdiº~¦r¦r§Î.txtª¬ºA = ´F¨Åª¬ºA¦C
      End If
   End If
End If

End Sub

Private Sub tree¦r§Î¾ğª¬µ²ºc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' Initiates dragging only after moving at least 100 twips with the mouse depressed
If (Button And 1) And (XCheck > 0) And (YCheck > 0) And ((Abs(XCheck - x) > 150) Or (Abs(YCheck - y) > 150)) Then
    XCheck = 0
    YCheck = 0  ' Reset mouse coordinates
    If tree¦r§Î¾ğª¬µ²ºc.ListIndex >= 0 Then
        tree¦r§Î¾ğª¬µ²ºc.BeforeDrag
        tree¦r§Î¾ğª¬µ²ºc.Drag 1         ' Start drag
    End If
End If

'If Button = 1 Then
'    tree¦r§Î¾ğª¬µ²ºc.BeforeDrag
'    tree¦r§Î¾ğª¬µ²ºc.Drag 1
'End If

End Sub

Public Function «Øºc²Õ¦r¦¡(²Õ¦r¦¡ As String) As String
Dim ¦r§Î½s¸¹°ïÅ|(100) As Integer, ¦r§Î°ïÅ|(100) As String
Dim maxid As Integer
Dim i As Integer, j As Integer, temp As Integer, temp2 As String
Dim ºc¦r¦¡ As String, ²Å¸¹ As String

If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Checked And ²Õ¦r¦¡ = "*" Then
    «Øºc²Õ¦r¦¡ = "*"
    Exit Function
End If

If ¨t²Î¦rÅé = "·¢®Ñ" Then
    Set ÀË¦rªí = ·¢®ÑÀË¦rªí
    Set ¦r®Úªí = ·¢®Ñ¦r®Ú
ElseIf ¨t²Î¦rÅé = "¤p½f" Then
    Set ÀË¦rªí = ¤p½fÀË¦rªí
    Set ¦r®Úªí = ¤p½f¿WÅé¦r
ElseIf ¨t²Î¦rÅé = "ª÷¤å" Then
    Set ÀË¦rªí = ª÷¤åÀË¦rªí
    Set ¦r®Úªí = ª÷¤å¦r®Ú
ElseIf ¨t²Î¦rÅé = "¥Ò°©¤å" Then
    Set ÀË¦rªí = ¥Ò°©¤åÀË¦rªí
    Set ¦r®Úªí = ¥Ò°©¤å¦r®Ú
ElseIf ¨t²Î¦rÅé = "·¡¨t¤å¦r" Then
    Set ÀË¦rªí = ·¡¨t¤å¦rÀË¦rªí
    Set ¦r®Úªí = ·¡¨t¤å¦r¦r®Ú
End If
ÀË¦rªí.Index = "¦r§Î"
¦r®Úªí.Index = "¦r§Î"
maxid = 0
²Å¸¹ = "*"
    
If Äæ¦ì§O = "³¡¥ó§Ç" Or Äæ¦ì§O = "³¡¥ó§Ç¤G" Then
   ºc¦r¦¡ = ²Å¸¹ & ²Õ¦r¦¡ & ²Å¸¹
ElseIf Äæ¦ì§O = "¦r®Ú§Ç" Or Äæ¦ì§O = "¦r®Ú§Ç¤G" Then
       If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Checked Then
            ºc¦r¦¡ = ""
       Else
            ºc¦r¦¡ = ²Å¸¹
       End If
       For i = 1 To Len(²Õ¦r¦¡)
           ÀË¦rªí.Seek "=", Mid(²Õ¦r¦¡, i, 1)
           If Not ÀË¦rªí.NoMatch Then
              Do While Not ÀË¦rªí.EOF And ÀË¦rªí.Fields("¦r§Î") = Mid(²Õ¦r¦¡, i, 1)     '¼Ğ·¢Åé
                ' If ÀË¦rªí.Fields("¦rÅé") = 0 Then
                    ºc¦r¦¡ = ºc¦r¦¡ & ÀË¦rªí.Fields(Äæ¦ì§O)
                    Exit Do
                'End If
                 ÀË¦rªí.MoveNext
              Loop
           ElseIf mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Checked And ¬O§_¬°¸U¥Î¦r¤¸(Mid(²Õ¦r¦¡, i, 1)) Then
                ºc¦r¦¡ = ºc¦r¦¡ & Mid(²Õ¦r¦¡, i, 1)
           End If
           If Not mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Checked Then ºc¦r¦¡ = ºc¦r¦¡ & ²Å¸¹
       Next
Else
   'Äæ¦ì§O=¦r®Ú²Õ OR ¦r®Ú²Õ¤G
    maxid = 0
    For i = 1 To Len(²Õ¦r¦¡)
        ÀË¦rªí.Seek "=", Mid(²Õ¦r¦¡, i, 1)
        If Not ÀË¦rªí.NoMatch Then
           Do While Not ÀË¦rªí.EOF
              If Not IsNull(ÀË¦rªí.Fields("¦r§Î")) Then
'              If ÀË¦rªí.Fields("¦rÅé") = 0 Then
                 If Not IsNull(ÀË¦rªí.Fields(Äæ¦ì§O)) Then
                    For j = 1 To Len(ÀË¦rªí.Fields(Äæ¦ì§O))
                        ¦r§Î°ïÅ|(maxid) = Mid(ÀË¦rªí.Fields(Äæ¦ì§O), j, 1)
                        maxid = maxid + 1
                    Next j
                 End If
                 Exit Do
              End If
              ÀË¦rªí.MoveNext
           Loop
        End If
    Next
 
    For i = 0 To maxid - 1
        ¦r®Úªí.Seek "=", ¦r§Î°ïÅ|(i)
        If Not ¦r®Úªí.NoMatch Then
           ¦r§Î½s¸¹°ïÅ|(i) = ¦r®Úªí.Fields("½s¸¹")
        End If
    Next
  
    For i = 0 To maxid - 2
        For j = i + 1 To maxid - 1
            If ¦r§Î½s¸¹°ïÅ|(i) > ¦r§Î½s¸¹°ïÅ|(j) Then
               temp = ¦r§Î½s¸¹°ïÅ|(i)
               ¦r§Î½s¸¹°ïÅ|(i) = ¦r§Î½s¸¹°ïÅ|(j)
               ¦r§Î½s¸¹°ïÅ|(j) = temp
               temp2 = ¦r§Î°ïÅ|(i)
               ¦r§Î°ïÅ|(i) = ¦r§Î°ïÅ|(j)
               ¦r§Î°ïÅ|(j) = temp2
            End If
        Next
    Next
       
    ºc¦r¦¡ = ²Å¸¹
       
    For i = 0 To maxid - 1
        ºc¦r¦¡ = ºc¦r¦¡ & ¦r§Î°ïÅ|(i) & ²Å¸¹
    Next
End If

«Øºc²Õ¦r¦¡ = ºc¦r¦¡
End Function


Private Sub ¸ü¤J¾ğª¬(²Õ¦r¦¡ As String)
Dim ¦r§Îªí As Recordset
Dim SQL³¯­z¦¡ As String
Dim i As Integer, ¦r¼Æ As Integer
Dim ¾ğ®Ú As String
Dim ¦r§Î As String, ¦r®Ú As String
Dim ¦r«¬ÀÉ As String
Dim ¦Ê¤À¤ñ As Long
Dim wd As String, ¼È¦s²Õ¦r¦¡ As String, ¦¸¼Æ As Integer, times As Integer
Dim ´F¨ÅÁ`¼Æ As Long, ³v¯Å¦C¥X±ø¥ó As String
Dim ¦r¶°¿z¿ï As String, ¦r¶°²Õ¦r¦r¼Æ As String
Dim «D¦r As Boolean, ³¡¥ó As Boolean, ¦r¶°Äæ¦ì As String

mdiº~¦r¦r§Î.txtª¬ºA = "¦r§Î¸ü¤J¤¤  ......       " & " , ±ı¤¤Â_½Ğ«ö Esc Áä"
Screen.MousePointer = ccHourglass

tree¦r§Î¾ğª¬µ²ºc.Clear
¾ğ®Ú = ²Õ¦r¦¡

If (Len(²Õ¦r¦¡) = 2 And Mid(²Õ¦r¦¡, 1, 1) >= "ô" And Mid(²Õ¦r¦¡, 1, 1) <= "û") Then
    wd = Mid(²Õ¦r¦¡, 1, 1)
    If wd = "õ" Or wd = "ô" Then
       ¦¸¼Æ = 2
    ElseIf wd = "ø" Or wd = "÷" Or wd = "ö" Then
       ¦¸¼Æ = 3
    ElseIf wd = "û" Or wd = "ú" Or wd = "ù" Then
       ¦¸¼Æ = 4
    End If
    ¼È¦s²Õ¦r¦¡ = Mid(²Õ¦r¦¡, 2, 1)
    For times = 1 To ¦¸¼Æ - 1
        ¼È¦s²Õ¦r¦¡ = ¼È¦s²Õ¦r¦¡ + ¼È¦s²Õ¦r¦¡
    Next times
    ²Õ¦r¦¡ = ¼È¦s²Õ¦r¦¡
End If

If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Checked = True Then
    If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó.Checked = True Then
        Äæ¦ì§O = "¦r®Ú§Ç¤G"
    Else
        Äæ¦ì§O = "¦r®Ú§Ç"
    End If
    GoTo ¶}©l«Øºc²Õ¦r¦¡
End If

If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó.Checked = True And ¨t²Î¦rÅé = "·¢®Ñ" Then    '¥]§t²§¼g
   If Len(²Õ¦r¦¡) = 1 Then
     'ÀË¬d¬O§_¬°¬Û¦ü¦r®Ú
      ²§¼g¦r®Ú.Seek "=", ²Õ¦r¦¡
      If Not ²§¼g¦r®Ú.NoMatch Then
         ²Õ¦r¦¡ = ²§¼g¦r®Ú.Fields("²§¼g")
      End If
      If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Checked = True Then
         Äæ¦ì§O = "³¡¥ó§Ç¤G"
      Else
         Äæ¦ì§O = "¦r®Ú§Ç¤G"
      End If
   Else
      If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å¿í·Ó¿é¤J³¡¥ó¶¶§Ç.Checked = True Then
         Äæ¦ì§O = "¦r®Ú§Ç¤G"
      Else
         Äæ¦ì§O = "¦r®Ú²Õ¤G"
      End If
   End If
Else
   '¤£¥]§t²§¼g
   If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Checked = True And Len(²Õ¦r¦¡) = 1 Then
      Äæ¦ì§O = "³¡¥ó§Ç"
   Else
      If Len(²Õ¦r¦¡) = 1 Or mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å¿í·Ó¿é¤J³¡¥ó¶¶§Ç.Checked = True Then
         Äæ¦ì§O = "¦r®Ú§Ç"
      Else
         Äæ¦ì§O = "¦r®Ú²Õ"
      End If
   End If
End If

¶}©l«Øºc²Õ¦r¦¡:

If ¨t²Î¦rÅé = "¤p½f" Then
    If Äæ¦ì§O = "³¡¥ó§Ç¤G" Then Äæ¦ì§O = "³¡¥ó§Ç"
    If Äæ¦ì§O = "¦r®Ú§Ç¤G" Then Äæ¦ì§O = "¦r®Ú§Ç"
    If Äæ¦ì§O = "¦r®Ú²Õ¤G" Then Äæ¦ì§O = "¦r®Ú²Õ"
    For i = 1 To Len(²Õ¦r¦¡)
        ¦r®Ú = Mid(²Õ¦r¦¡, i, 1)
        ¤p½f²§¼g¦r®Ú.Seek "=", ¦r®Ú
        If Not ¤p½f²§¼g¦r®Ú.NoMatch Then Mid(²Õ¦r¦¡, i, 1) = ¤p½f²§¼g¦r®Ú.Fields("²§¼g")
    Next i
    ¾ğ®Ú = ²Õ¦r¦¡
End If

If ¨t²Î¦rÅé = "ª÷¤å" Then
    If Äæ¦ì§O = "³¡¥ó§Ç¤G" Then Äæ¦ì§O = "³¡¥ó§Ç"
    If Äæ¦ì§O = "¦r®Ú§Ç¤G" Then Äæ¦ì§O = "¦r®Ú§Ç"
    If Äæ¦ì§O = "¦r®Ú²Õ¤G" Then Äæ¦ì§O = "¦r®Ú²Õ"
    For i = 1 To Len(²Õ¦r¦¡)
        ¦r®Ú = Mid(²Õ¦r¦¡, i, 1)
        ª÷¤å²§¼g¦r®Ú.Seek "=", ¦r®Ú
        If Not ª÷¤å²§¼g¦r®Ú.NoMatch Then Mid(²Õ¦r¦¡, i, 1) = ª÷¤å²§¼g¦r®Ú.Fields("²§¼g")
    Next i
    ¾ğ®Ú = ²Õ¦r¦¡
End If

If ¨t²Î¦rÅé = "¥Ò°©¤å" Then
    If Äæ¦ì§O = "³¡¥ó§Ç¤G" Then Äæ¦ì§O = "³¡¥ó§Ç"
    If Äæ¦ì§O = "¦r®Ú§Ç¤G" Then Äæ¦ì§O = "¦r®Ú§Ç"
    If Äæ¦ì§O = "¦r®Ú²Õ¤G" Then Äæ¦ì§O = "¦r®Ú²Õ"
    For i = 1 To Len(²Õ¦r¦¡)
        ¦r®Ú = Mid(²Õ¦r¦¡, i, 1)
        ¥Ò°©¤å²§¼g¦r®Ú.Seek "=", ¦r®Ú
        If Not ¥Ò°©¤å²§¼g¦r®Ú.NoMatch Then Mid(²Õ¦r¦¡, i, 1) = ¥Ò°©¤å²§¼g¦r®Ú.Fields("²§¼g")
    Next i
    ¾ğ®Ú = ²Õ¦r¦¡
End If

If ¨t²Î¦rÅé = "·¡¨t¤å¦r" Then
    If Äæ¦ì§O = "³¡¥ó§Ç¤G" Then Äæ¦ì§O = "³¡¥ó§Ç"
    If Äæ¦ì§O = "¦r®Ú§Ç¤G" Then Äæ¦ì§O = "¦r®Ú§Ç"
    If Äæ¦ì§O = "¦r®Ú²Õ¤G" Then Äæ¦ì§O = "¦r®Ú²Õ"
    For i = 1 To Len(²Õ¦r¦¡)
        ¦r®Ú = Mid(²Õ¦r¦¡, i, 1)
        ·¡¨t¤å¦r²§¼g¦r®Ú.Seek "=", ¦r®Ú
        If Not ·¡¨t¤å¦r²§¼g¦r®Ú.NoMatch Then Mid(²Õ¦r¦¡, i, 1) = ·¡¨t¤å¦r²§¼g¦r®Ú.Fields("²§¼g")
    Next i
    ¾ğ®Ú = ²Õ¦r¦¡
End If

²Õ¦r¦¡ = «Øºc²Õ¦r¦¡(²Õ¦r¦¡)

If mdiº~¦r¦r§Î.mnu_±`¥Î¦r.Checked = True Then
    ¦r¶°Äæ¦ì = "±`¥Î¦r"
    ¦r¶°¿z¿ï = "±`¥Î¦r > 0"
    ¦r¶°²Õ¦r¦r¼Æ = "[²Õ¦r¦r¼Æ(±`¥Î¦r)]"
ElseIf mdiº~¦r¦r§Î.mnu_Big5.Checked = True Then
    ¦r¶°Äæ¦ì = "Big5"
    ¦r¶°¿z¿ï = "½s¸¹ > 0 and ½s¸¹ <= 13053"
    ¦r¶°²Õ¦r¦r¼Æ = "[²Õ¦r¦r¼Æ(Big5)]"
ElseIf mdiº~¦r¦r§Î.mnu_Â²¤Æ¦rÁ`ªí.Checked = True Then
    ¦r¶°Äæ¦ì = "Â²¤Æ¦r"
    ¦r¶°¿z¿ï = "Â²¤Æ¦r > 0"
    ¦r¶°²Õ¦r¦r¼Æ = "[²Õ¦r¦r¼Æ(Â²¤Æ¦r)]"
ElseIf mdiº~¦r¦r§Î.mnu_º~»y¤j¦r¨å.Checked = True Then
    ¦r¶°Äæ¦ì = "º~»y¤j¦r¨å"
    ¦r¶°¿z¿ï = "º~»y¤j¦r¨å > 0"
    ¦r¶°²Õ¦r¦r¼Æ = "[²Õ¦r¦r¼Æ(º~»y¤j¦r¨å)]"
ElseIf mdiº~¦r¦r§Î.mnu_ª÷¤å¹Ï§Î¤å¦r.Checked = True Then
    ¦r¶°Äæ¦ì = "½s¸¹"
    ¦r¶°¿z¿ï = "½s¸¹>0 and ¹Ï§Î¤å¦r>0"
    ¦r¶°²Õ¦r¦r¼Æ = "[²Õ¦r¦r¼Æ(¹Ï§Î¤å¦r)]"
Else
    ¦r¶°Äæ¦ì = "½s¸¹"
    ¦r¶°¿z¿ï = "½s¸¹ > 0"
    ¦r¶°²Õ¦r¦r¼Æ = "²Õ¦r¦r¼Æ"
End If


If ªì©l³v¯Å¦C¥X = 1 And Mid(Äæ¦ì§O, 1, 2) = "³¡¥ó" Then
   '³v¯Å¦C¥X±ø¥ó = " ( " & ¦r¶°²Õ¦r¦r¼Æ & " > 1 or ( " & ¦r¶°²Õ¦r¦r¼Æ & " = 1 and " & ¦r¶°¿z¿ï & " ) )"
   ³v¯Å¦C¥X±ø¥ó = " ( " & ¦r¶°²Õ¦r¦r¼Æ & " > 0 )"
Else
   ³v¯Å¦C¥X±ø¥ó = " (  " & ¦r¶°¿z¿ï & " ) "
End If

If Mid(Äæ¦ì§O, 1, 2) = "¦r®Ú" Then
   If ¨t²Î¦rÅé = "·¢®Ñ" Then
        SQL³¯­z¦¡ = "SELECT * From ÀË¦rªí Where ½s¸¹>0 and " & ³v¯Å¦C¥X±ø¥ó & " and " & Äæ¦ì§O & " Like '" & ²Õ¦r¦¡ & "'  ORDER BY ³¡­ºµ§µe±Æ§Ç,²Õ¦r¦r¼Æ DESC "
   Else
        SQL³¯­z¦¡ = "SELECT * From ÀË¦rªí Where ½s¸¹>0 and " & ³v¯Å¦C¥X±ø¥ó & " and " & Äæ¦ì§O & " Like '" & ²Õ¦r¦¡ & "'  ORDER BY ½s¸¹,²Õ¦r¦r¼Æ DESC "
   End If
Else
   If ¨t²Î¦rÅé = "·¢®Ñ" Then
        SQL³¯­z¦¡ = "SELECT * From ÀË¦rªí Where (¦r§Î='" & ¾ğ®Ú & "') or ( " & ³v¯Å¦C¥X±ø¥ó & " And " & Äæ¦ì§O & " Like '" & ²Õ¦r¦¡ & "')  ORDER BY ³¡­ºµ§µe±Æ§Ç,²Õ¦r¦r¼Æ DESC "
   Else
        SQL³¯­z¦¡ = "SELECT * From ÀË¦rªí Where (¦r§Î='" & ¾ğ®Ú & "') or ( " & ³v¯Å¦C¥X±ø¥ó & " And " & Äæ¦ì§O & " Like '" & ²Õ¦r¦¡ & "')  ORDER BY ½s¸¹,²Õ¦r¦r¼Æ DESC "
   End If
End If

DoEvents

¾ğ®Ú§Ç¼Æ = 0
´F¨ÅÁ`¼Æ = 0
Á`¼Æ = 0

'If ²Õ¦r¦¡ <> "" And ²Õ¦r¦¡ <> "*" And ²Õ¦r¦¡ <> "**" Then
If ²Õ¦r¦¡ <> "" And ²Õ¦r¦¡ <> "**" Then
   tree¦r§Î¾ğª¬µ²ºc.Redraw = False
   
    If ¨t²Î¦rÅé = "·¢®Ñ" Then
        Set ¦r§Îªí = ¨t²Î¸ê®Æ®w.OpenRecordset(SQL³¯­z¦¡)
    ElseIf ¨t²Î¦rÅé = "¤p½f" Then
        Set ¦r§Îªí = ¤p½f¸ê®Æ®w.OpenRecordset(SQL³¯­z¦¡)
    ElseIf ¨t²Î¦rÅé = "ª÷¤å" Then
        Set ¦r§Îªí = ª÷¤å¸ê®Æ®w.OpenRecordset(SQL³¯­z¦¡)
    ElseIf ¨t²Î¦rÅé = "¥Ò°©¤å" Then
        Set ¦r§Îªí = ¥Ò°©¤å¸ê®Æ®w.OpenRecordset(SQL³¯­z¦¡)
    ElseIf ¨t²Î¦rÅé = "·¡¨t¤å¦r" Then
        Set ¦r§Îªí = ·¡¨t¤å¦r¸ê®Æ®w.OpenRecordset(SQL³¯­z¦¡)
    End If
   
   If Not ¦r§Îªí.EOF Then
   
      ¦r§Îªí.MoveLast
      Á`¼Æ = ¦r§Îªí.RecordCount
            
      If ²Õ¦r¦¡ <> "*" Then
        tree¦r§Î¾ğª¬µ²ºc.AddItem ¾ğ®Ú
      Else
        tree¦r§Î¾ğª¬µ²ºc.AddItem "¡¯"
      End If
      If ¨t²Î¦rÅé = "·¢®Ñ" Then
        tree¦r§Î¾ğª¬µ²ºc.ItemFontName(0) = Åã¥Ü¦r«¬ '²{¥Î¦rÅé
      ElseIf ¨t²Î¦rÅé = "¤p½f" Then
        tree¦r§Î¾ğª¬µ²ºc.ItemFontName(0) = "¥_®v¤j»¡¤å¤p½f"
      ElseIf ¨t²Î¦rÅé = "ª÷¤å" Then
        tree¦r§Î¾ğª¬µ²ºc.ItemFontName(0) = "¤¤¬ã°|ª÷¤å"
      ElseIf ¨t²Î¦rÅé = "¥Ò°©¤å" Then
        tree¦r§Î¾ğª¬µ²ºc.ItemFontName(0) = "¤¤¬ã°|¥Ò°©¤å"
      ElseIf ¨t²Î¦rÅé = "·¡¨t¤å¦r" Then
        tree¦r§Î¾ğª¬µ²ºc.ItemFontName(0) = "¤¤¬ã°|·¡¨tÂ²©­¤å¦r"
      End If
      
      tree¦r§Î¾ğª¬µ²ºc.ItemLngValue(0) = -999999
      If Len(¾ğ®Ú) = 1 Then
         tree¦r§Î¾ğª¬µ²ºc.ItemTag(0) = ¦r§Î¸`ÂI¼Ğ°O
      Else
         tree¦r§Î¾ğª¬µ²ºc.ItemTag(0) = ¨ä¥L¸`ÂI¼Ğ°O
      End If
        
      tree¦r§Î¾ğª¬µ²ºc.Expand(0) = True
    
      ¦r§Îªí.MoveFirst

      Do Until ¦r§Îªí.EOF
         If ¤¤Â_ = True Then Exit Do
         
         ¦Ê¤À¤ñ = ¾ğ®Ú§Ç¼Æ / Á`¼Æ * 100
         mdiº~¦r¦r§Î.txtª¬ºA = "¦r§Î¸ü¤J¤w§¹¦¨ " & ¦Ê¤À¤ñ & " % , ±ı¤¤Â_½Ğ«ö Esc Áä"
         
         If Not IsNull(¦r§Îªí.Fields("¦r§Î")) And ¦r§Îªí.Fields("¦r§Î") <> "" Then
            ¦r§Î = ¦r§Îªí.Fields("¦r§Î")
         Else
            If Not IsNull(¦r§Îªí.Fields("¦r½X")) And ¦r§Îªí.Fields("¦r½X") <> "" Then
               ¦r§Î = ¦r§Îªí.Fields("¦r½X")
            Else
               ¦r§Î = "¡´"
            End If
         End If
           
         ¦r¼Æ = ¦r§Îªí.Fields(¦r¶°²Õ¦r¦r¼Æ)
         If ¦r¶°Äæ¦ì <> "Big5" Then
            If IsNull(¦r§Îªí.Fields(¦r¶°Äæ¦ì)) Then
                «D¦r = True
            ElseIf ¦r§Îªí.Fields(¦r¶°Äæ¦ì) = 0 Then
                «D¦r = True
            Else
                «D¦r = False
            End If
         Else
            If ¦r§Îªí.Fields("½s¸¹") > 0 And ¦r§Îªí.Fields("½s¸¹") <= 13053 Then
                «D¦r = False
            ElseIf ¦r§Îªí.Fields("½s¸¹") > 13053 Then
                «D¦r = True
            Else
                «D¦r = False
            End If
         End If
         
         If ¦r¼Æ > 1 Then
            ³¡¥ó = True
         ElseIf ¦r¼Æ = 1 And «D¦r Then
            ³¡¥ó = True
         Else
            ³¡¥ó = False
         End If
                    
         If Not (°Ï°ì¦rÅé°}¦C(¦r§Îªí.Fields("¦rÅé")) = Âà´«¦rÅé(tree¦r§Î¾ğª¬µ²ºc.ItemFontName(0)) And ¦r§Î = tree¦r§Î¾ğª¬µ²ºc.List(0)) Then
            If Not ((Äæ¦ì§O = "³¡¥ó§Ç" Or Äæ¦ì§O = "³¡¥ó§Ç¤G") And (¦r§Îªí.Fields("³s±µ²Å¸¹") = 0)) Then
               ¾ğ®Ú§Ç¼Æ = ¾ğ®Ú§Ç¼Æ + 1
               If ³¡¥ó Then
                  tree¦r§Î¾ğª¬µ²ºc.AddItem ¦r§Î, 0
                  tree¦r§Î¾ğª¬µ²ºc.ItemFontName(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = Âà´«Åã¥Ü¦r«¬(°Ï°ì¦rÅé°}¦C(¦r§Îªí.Fields("¦rÅé")))
                  tree¦r§Î¾ğª¬µ²ºc.ItemLngValue(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = ¦r§Îªí.Fields("½s¸¹")
                  If Not IsNull(¦r§Îªí.Fields("¦r§Î")) Then
                     tree¦r§Î¾ğª¬µ²ºc.ItemTag(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = ¦r§Î¸`ÂI¼Ğ°O
                  Else
                     tree¦r§Î¾ğª¬µ²ºc.ItemTag(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = ºc¦r¦¡¸`ÂI¼Ğ°O
                  End If

                  If Mid(Äæ¦ì§O, 1, 2) = "¦r®Ú" Then
                     tree¦r§Î¾ğª¬µ²ºc.Image(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = tree¦r§Î¾ğª¬µ²ºc.PictureLeaf
                  Else
                     If ¦r§Îªí.Fields("³s±µ²Å¸¹") = 0 Then
                        tree¦r§Î¾ğª¬µ²ºc.Image(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = tree¦r§Î¾ğª¬µ²ºc.PictureLeaf
                     Else
                        tree¦r§Î¾ğª¬µ²ºc.Image(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = tree¦r§Î¾ğª¬µ²ºc.PictureClosed
                        tree¦r§Î¾ğª¬µ²ºc.AddItem "", tree¦r§Î¾ğª¬µ²ºc.NewIndex
                     End If
                  End If
               Else
                  tree¦r§Î¾ğª¬µ²ºc.AddItem ¦r§Î, 0
                  tree¦r§Î¾ğª¬µ²ºc.Image(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = tree¦r§Î¾ğª¬µ²ºc.PictureLeaf
                  tree¦r§Î¾ğª¬µ²ºc.ItemFontName(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = Âà´«Åã¥Ü¦r«¬(°Ï°ì¦rÅé°}¦C(¦r§Îªí.Fields("¦rÅé")))
                  tree¦r§Î¾ğª¬µ²ºc.ItemLngValue(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = ¦r§Îªí.Fields("½s¸¹")
                  If Not IsNull(¦r§Îªí.Fields("¦r§Î")) Then
                     tree¦r§Î¾ğª¬µ²ºc.ItemTag(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = ¦r§Î¸`ÂI¼Ğ°O
                  Else
                     tree¦r§Î¾ğª¬µ²ºc.ItemTag(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = ºc¦r¦¡¸`ÂI¼Ğ°O
                  End If

               End If
             End If
         Else
            If Äæ¦ì§O = "³¡¥ó§Ç" Or ¨t²Î¦rÅé <> "·¢®Ñ" Then
                ´F¨ÅÁ`¼Æ = ¦r§Îªí.Fields("²Õ¦r¦r¼Æ")
            ElseIf Äæ¦ì§O = "³¡¥ó§Ç¤G" Then
                ´F¨ÅÁ`¼Æ = ¦r§Îªí.Fields("²Õ¦r¦r¼Æ¤G")
            Else
                ¾ğ®Ú§Ç¼Æ = ¾ğ®Ú§Ç¼Æ + 1
            End If
            
            If ³¡¥ó Then
               tree¦r§Î¾ğª¬µ²ºc.Image(0) = tree¦r§Î¾ğª¬µ²ºc.PictureOpen
            Else
               If Len(¾ğ®Ú) = 1 Then
                  tree¦r§Î¾ğª¬µ²ºc.Image(0) = tree¦r§Î¾ğª¬µ²ºc.PictureLeaf
               Else
                  tree¦r§Î¾ğª¬µ²ºc.Image(0) = tree¦r§Î¾ğª¬µ²ºc.PictureOpen
               End If
            End If
            tree¦r§Î¾ğª¬µ²ºc.ItemFontName(0) = Âà´«Åã¥Ü¦r«¬(°Ï°ì¦rÅé°}¦C(¦r§Îªí.Fields("¦rÅé")))
            tree¦r§Î¾ğª¬µ²ºc.ItemLngValue(0) = ¦r§Îªí.Fields("½s¸¹")
            If Len(¾ğ®Ú) = 1 Then
               tree¦r§Î¾ğª¬µ²ºc.ItemTag(0) = ¦r§Î¸`ÂI¼Ğ°O
            Else
               tree¦r§Î¾ğª¬µ²ºc.ItemTag(0) = ¨ä¥L¸`ÂI¼Ğ°O
            End If
         End If
         
         ¦r§Îªí.MoveNext
         
         If (tree¦r§Î¾ğª¬µ²ºc.ListCount + 10) Mod 50 = 0 Then
            tree¦r§Î¾ğª¬µ²ºc.Redraw = True
            Screen.MousePointer = ccDefault
            DoEvents
            tree¦r§Î¾ğª¬µ²ºc.Redraw = False
            Screen.MousePointer = ccHourglass
         End If
         
      Loop
   Else
      Á`¼Æ = 0
   End If
End If

If Á`¼Æ > 0 Then
    If tree¦r§Î¾ğª¬µ²ºc.ItemLngValue(0) = -999999 And Len(tree¦r§Î¾ğª¬µ²ºc.List(0)) = 1 Then

        If ¨t²Î¦rÅé = "¤p½f" Then
            Set ÀË¦rªí = ¤p½fÀË¦rªí
        ElseIf ¨t²Î¦rÅé = "¤¤¬ã°|ª÷¤å" Then
            Set ÀË¦rªí = ª÷¤åÀË¦rªí
        ElseIf ¨t²Î¦rÅé = "¤¤¬ã°|¥Ò°©¤å" Then
            Set ÀË¦rªí = ¥Ò°©¤åÀË¦rªí
        ElseIf ¨t²Î¦rÅé = "¤¤¬ã°|·¡¨tÂ²©­¤å¦r" Then
            Set ÀË¦rªí = ·¡¨t¤å¦rÀË¦rªí
        Else
            Set ÀË¦rªí = ·¢®ÑÀË¦rªí
        End If
   
        ÀË¦rªí.Index = "¦r§Î"
        ÀË¦rªí.Seek "=", tree¦r§Î¾ğª¬µ²ºc.List(0)
        If Not ÀË¦rªí.NoMatch Then
            If ÀË¦rªí.Fields("½s¸¹") = 0 And ¨t²Î¦rÅé = "¤p½f" Then
                tree¦r§Î¾ğª¬µ²ºc.ItemFontName(0) = "¼Ğ·¢Åé"
                tree¦r§Î¾ğª¬µ²ºc.ItemLngValue(0) = ÀË¦rªí.Fields("·¢®Ñ½s¸¹")
            Else
                tree¦r§Î¾ğª¬µ²ºc.ItemLngValue(0) = ÀË¦rªí.Fields("½s¸¹")
            End If
        End If
    End If

End If

If Mid(Äæ¦ì§O, 1, 3) <> "³¡¥ó§Ç" Then ´F¨ÅÁ`¼Æ = ¾ğ®Ú§Ç¼Æ

tree¦r§Î¾ğª¬µ²ºc.Redraw = True
´F¨Åª¬ºA¦C = "¸ü¤J§¹²¦¡I ¦@ " & Á`¼Æ & " ­Ó¦r§Î(³¡¥ó)"
mdiº~¦r¦r§Î.txtª¬ºA = ´F¨Åª¬ºA¦C
'mdiº~¦r¦r§Î.sbarª¬ºA¦C.Visible = True

²Ó³¡¨t²Î¦rÅé = ¨t²Î¦rÅé

Screen.MousePointer = ccDefault

End Sub

Public Sub ¸ü¤J²Ó³¡¾ğª¬µ²ºc(i As Long)
Dim ²Ó³¡¦r§Îªí As Recordset
Dim ²Õ¦r¦¡ As String, ¦r§Î As String
Dim SQL³¯­z¦¡ As String, ²Ó³¡Á`¼Æ As Long
Dim ¦r«¬ÀÉ As String, ³v¯Å¦C¥X±ø¥ó As String
Dim ½s¸¹ As Long, ¦r¼Æ As Integer
Dim ¦r¶°¿z¿ï As String, ¦r¶°²Õ¦r¦r¼Æ As String
Dim «D¦r As Boolean, ³¡¥ó As Boolean, ¦r¶°Äæ¦ì As String


If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó.Checked = True Then    '¥]§t²§¼g
   If Len(tree¦r§Î¾ğª¬µ²ºc.List(i)) = 1 Then
      'ÀË¬d¬O§_¬°²§¼g¦r®Ú
      ²§¼g¦r®Ú.Seek "=", tree¦r§Î¾ğª¬µ²ºc.List(i)
      If Not ²§¼g¦r®Ú.NoMatch Then
         ²Õ¦r¦¡ = ²§¼g¦r®Ú.Fields("²§¼g")
      Else
         ²Õ¦r¦¡ = tree¦r§Î¾ğª¬µ²ºc.List(i)
      End If
      If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Checked = True Then
         Äæ¦ì§O = "³¡¥ó§Ç¤G"
      Else
         Äæ¦ì§O = "¦r®Ú§Ç¤G"
      End If
   Else
      Äæ¦ì§O = "¦r®Ú²Õ¤G"
   End If
Else
   '¤£¥]§t²§¼g
   ²Õ¦r¦¡ = tree¦r§Î¾ğª¬µ²ºc.List(i)
   If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Checked = True And Len(²Õ¦r¦¡) = 1 Then
      Äæ¦ì§O = "³¡¥ó§Ç"
   Else
      If Len(²Õ¦r¦¡) = 1 Then
         Äæ¦ì§O = "¦r®Ú§Ç"
      Else
         Äæ¦ì§O = "¦r®Ú²Õ"
      End If
   End If
End If

If ²Ó³¡¨t²Î¦rÅé = "¤p½f" Or ²Ó³¡¨t²Î¦rÅé = "ª÷¤å" Or ²Ó³¡¨t²Î¦rÅé = "¥Ò°©¤å" Or ²Ó³¡¨t²Î¦rÅé = "·¡¨t¤å¦r" Then
    If Äæ¦ì§O = "³¡¥ó§Ç¤G" Then Äæ¦ì§O = "³¡¥ó§Ç"
    If Äæ¦ì§O = "¦r®Ú§Ç¤G" Then Äæ¦ì§O = "¦r®Ú§Ç"
    If Äæ¦ì§O = "¦r®Ú²Õ¤G" Then Äæ¦ì§O = "¦r®Ú²Õ"
    'For i = 1 To Len(²Õ¦r¦¡)
    '    ¦r®Ú = Mid(²Õ¦r¦¡, i, 1)
    '    ¤p½f²§¼g¦r®Ú.Seek "=", ¦r®Ú
    '    If Not ¤p½f²§¼g¦r®Ú.NoMatch Then Mid(²Õ¦r¦¡, i, 1) = ¤p½f²§¼g¦r®Ú.Fields("²§¼g")
    'Next i
    '¾ğ®Ú = ²Õ¦r¦¡
End If

²Õ¦r¦¡ = «Øºc²Õ¦r¦¡(²Õ¦r¦¡)

'If ªì©l³v¯Å¦C¥X = 1 Then
'   ³v¯Å¦C¥X±ø¥ó = " ( ²Õ¦r¦r¼Æ > 1 or ( ²Õ¦r¦r¼Æ =1 and ¦rÀW >= " & Trim(ªì©l¦rÀW) & " ) )"
'Else
'   ³v¯Å¦C¥X±ø¥ó = " ( ¦rÀW >= " & Trim(ªì©l¦rÀW) & " ) "
'End If

If mdiº~¦r¦r§Î.mnu_±`¥Î¦r.Checked = True Then
    ¦r¶°Äæ¦ì = "±`¥Î¦r"
    ¦r¶°¿z¿ï = "±`¥Î¦r > 0"
    ¦r¶°²Õ¦r¦r¼Æ = "[²Õ¦r¦r¼Æ(±`¥Î¦r)]"
ElseIf mdiº~¦r¦r§Î.mnu_Big5.Checked = True Then
    ¦r¶°Äæ¦ì = "Big5"
    ¦r¶°¿z¿ï = "½s¸¹ > 0 and ½s¸¹ <= 13053"
    ¦r¶°²Õ¦r¦r¼Æ = "[²Õ¦r¦r¼Æ(Big5)]"
ElseIf mdiº~¦r¦r§Î.mnu_Â²¤Æ¦rÁ`ªí.Checked = True Then
    ¦r¶°Äæ¦ì = "Â²¤Æ¦r"
    ¦r¶°¿z¿ï = "Â²¤Æ¦r > 0"
    ¦r¶°²Õ¦r¦r¼Æ = "[²Õ¦r¦r¼Æ(Â²¤Æ¦r)]"
ElseIf mdiº~¦r¦r§Î.mnu_º~»y¤j¦r¨å.Checked = True Then
    ¦r¶°Äæ¦ì = "º~»y¤j¦r¨å"
    ¦r¶°¿z¿ï = "º~»y¤j¦r¨å > 0"
    ¦r¶°²Õ¦r¦r¼Æ = "[²Õ¦r¦r¼Æ(º~»y¤j¦r¨å)]"
Else
    ¦r¶°Äæ¦ì = "½s¸¹"
    ¦r¶°¿z¿ï = "½s¸¹ > 0"
    ¦r¶°²Õ¦r¦r¼Æ = "²Õ¦r¦r¼Æ"
End If


If ªì©l³v¯Å¦C¥X = 1 And Mid(Äæ¦ì§O, 1, 2) = "³¡¥ó" Then
   '³v¯Å¦C¥X±ø¥ó = " ( " & ¦r¶°²Õ¦r¦r¼Æ & " > 1 or ( " & ¦r¶°²Õ¦r¦r¼Æ & " = 1 and " & ¦r¶°¿z¿ï & " ) )"
   ³v¯Å¦C¥X±ø¥ó = " ( " & ¦r¶°²Õ¦r¦r¼Æ & " > 0 )"
Else
   ³v¯Å¦C¥X±ø¥ó = " (  " & ¦r¶°¿z¿ï & " ) "
End If

If Mid(Äæ¦ì§O, 1, 2) = "¦r®Ú" Then
    If ²Ó³¡¨t²Î¦rÅé = "·¢®Ñ" Then
        SQL³¯­z¦¡ = "SELECT * From ÀË¦rªí Where ½s¸¹>0 and " & ³v¯Å¦C¥X±ø¥ó & " and " & Äæ¦ì§O & " Like '" & ²Õ¦r¦¡ & "'  ORDER BY ³¡­ºµ§µe±Æ§Ç,²Õ¦r¦r¼Æ DESC "
    Else
        SQL³¯­z¦¡ = "SELECT * From ÀË¦rªí Where ½s¸¹>0 and " & ³v¯Å¦C¥X±ø¥ó & " and " & Äæ¦ì§O & " Like '" & ²Õ¦r¦¡ & "'  ORDER BY ½s¸¹,²Õ¦r¦r¼Æ DESC "
    End If
Else
    If ²Ó³¡¨t²Î¦rÅé = "·¢®Ñ" Then
        SQL³¯­z¦¡ = "SELECT * From ÀË¦rªí Where (¦r§Î='" & tree¦r§Î¾ğª¬µ²ºc.List(i) & "') or ( " & ³v¯Å¦C¥X±ø¥ó & " And " & Äæ¦ì§O & " Like '" & ²Õ¦r¦¡ & "')  ORDER BY ³¡­ºµ§µe±Æ§Ç,²Õ¦r¦r¼Æ DESC "
    Else
        SQL³¯­z¦¡ = "SELECT * From ÀË¦rªí Where (¦r§Î='" & tree¦r§Î¾ğª¬µ²ºc.List(i) & "') or ( " & ³v¯Å¦C¥X±ø¥ó & " And " & Äæ¦ì§O & " Like '" & ²Õ¦r¦¡ & "')  ORDER BY ½s¸¹,²Õ¦r¦r¼Æ DESC "
    End If
End If

If ²Ó³¡¨t²Î¦rÅé = "·¢®Ñ" Then
    Set ²Ó³¡¦r§Îªí = ¨t²Î¸ê®Æ®w.OpenRecordset(SQL³¯­z¦¡)
ElseIf ²Ó³¡¨t²Î¦rÅé = "¤p½f" Then
    Set ²Ó³¡¦r§Îªí = ¤p½f¸ê®Æ®w.OpenRecordset(SQL³¯­z¦¡)
ElseIf ²Ó³¡¨t²Î¦rÅé = "ª÷¤å" Then
    Set ²Ó³¡¦r§Îªí = ª÷¤å¸ê®Æ®w.OpenRecordset(SQL³¯­z¦¡)
ElseIf ²Ó³¡¨t²Î¦rÅé = "¥Ò°©¤å" Then
    Set ²Ó³¡¦r§Îªí = ¥Ò°©¤å¸ê®Æ®w.OpenRecordset(SQL³¯­z¦¡)
ElseIf ²Ó³¡¨t²Î¦rÅé = "·¡¨t¤å¦r" Then
    Set ²Ó³¡¦r§Îªí = ·¡¨t¤å¦r¸ê®Æ®w.OpenRecordset(SQL³¯­z¦¡)
End If
   
²Ó³¡¦r§Îªí.MoveLast
²Ó³¡Á`¼Æ = ²Ó³¡¦r§Îªí.RecordCount

²Ó³¡¦r§Îªí.MoveFirst

   If Not ²Ó³¡¦r§Îªí.EOF Then
      
      Do Until ²Ó³¡¦r§Îªí.EOF
        
         If Not IsNull(²Ó³¡¦r§Îªí.Fields("¦r§Î")) Then
            ¦r§Î = ²Ó³¡¦r§Îªí.Fields("¦r§Î")
         Else
            If Not IsNull(²Ó³¡¦r§Îªí.Fields("¦r½X")) Then
              ¦r§Î = ²Ó³¡¦r§Îªí.Fields("¦r½X")
            Else
              ¦r§Î = "¡´"
            End If
         End If
         
         '¦r¼Æ = ²Ó³¡¦r§Îªí.Fields("²Õ¦r¦r¼Æ")
         ¦r¼Æ = ²Ó³¡¦r§Îªí.Fields(¦r¶°²Õ¦r¦r¼Æ)
         If ¦r¶°Äæ¦ì <> "Big5" Then
            If IsNull(²Ó³¡¦r§Îªí.Fields(¦r¶°Äæ¦ì)) Then
                «D¦r = True
            ElseIf ²Ó³¡¦r§Îªí.Fields(¦r¶°Äæ¦ì) = 0 Then
                «D¦r = True
            Else
                «D¦r = False
            End If
         Else
            If ²Ó³¡¦r§Îªí.Fields("½s¸¹") > 0 And ²Ó³¡¦r§Îªí.Fields("½s¸¹") <= 13053 Then
                «D¦r = False
            ElseIf ²Ó³¡¦r§Îªí.Fields("½s¸¹") > 13053 Then
                «D¦r = True
            Else
                «D¦r = False
            End If
         End If
         
         If ¦r¼Æ > 1 Then
            ³¡¥ó = True
         ElseIf ¦r¼Æ = 1 And «D¦r Then
            ³¡¥ó = True
         Else
            ³¡¥ó = False
         End If
         
         If Not (°Ï°ì¦rÅé°}¦C(²Ó³¡¦r§Îªí.Fields("¦rÅé")) = Âà´«¦rÅé(tree¦r§Î¾ğª¬µ²ºc.ItemFontName(i)) And ¦r§Î = tree¦r§Î¾ğª¬µ²ºc.List(i)) Then
            If Not ((Äæ¦ì§O = "³¡¥ó§Ç" Or Äæ¦ì§O = "³¡¥ó§Ç¤G") And (²Ó³¡¦r§Îªí.Fields("³s±µ²Å¸¹") = 0)) Then
               If ³¡¥ó Then
                  tree¦r§Î¾ğª¬µ²ºc.AddItem ¦r§Î, i
                  tree¦r§Î¾ğª¬µ²ºc.ItemFontName(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = Âà´«Åã¥Ü¦r«¬(°Ï°ì¦rÅé°}¦C(²Ó³¡¦r§Îªí.Fields("¦rÅé")))
                  tree¦r§Î¾ğª¬µ²ºc.ItemLngValue(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = ²Ó³¡¦r§Îªí.Fields("½s¸¹")
                  If Not IsNull(²Ó³¡¦r§Îªí.Fields("¦r§Î")) Then
                     tree¦r§Î¾ğª¬µ²ºc.ItemTag(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = ¦r§Î¸`ÂI¼Ğ°O
                  Else
                     tree¦r§Î¾ğª¬µ²ºc.ItemTag(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = ºc¦r¦¡¸`ÂI¼Ğ°O
                  End If
                  
                  If Mid(Äæ¦ì§O, 1, 2) = "¦r®Ú" Then
                     tree¦r§Î¾ğª¬µ²ºc.Image(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = tree¦r§Î¾ğª¬µ²ºc.PictureLeaf
                  Else
                     If ²Ó³¡¦r§Îªí.Fields("³s±µ²Å¸¹") = 0 Then
                        tree¦r§Î¾ğª¬µ²ºc.Image(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = tree¦r§Î¾ğª¬µ²ºc.PictureLeaf
                     Else
                        tree¦r§Î¾ğª¬µ²ºc.Image(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = tree¦r§Î¾ğª¬µ²ºc.PictureClosed
                        tree¦r§Î¾ğª¬µ²ºc.AddItem "", tree¦r§Î¾ğª¬µ²ºc.NewIndex
                     End If
                  End If
               Else
                  tree¦r§Î¾ğª¬µ²ºc.AddItem ¦r§Î, i
                  tree¦r§Î¾ğª¬µ²ºc.Image(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = tree¦r§Î¾ğª¬µ²ºc.PictureLeaf
                  tree¦r§Î¾ğª¬µ²ºc.ItemFontName(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = Âà´«Åã¥Ü¦r«¬(°Ï°ì¦rÅé°}¦C(²Ó³¡¦r§Îªí.Fields("¦rÅé")))
                  tree¦r§Î¾ğª¬µ²ºc.ItemLngValue(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = ²Ó³¡¦r§Îªí.Fields("½s¸¹")
                  If Not IsNull(²Ó³¡¦r§Îªí.Fields("¦r§Î")) Then
                     tree¦r§Î¾ğª¬µ²ºc.ItemTag(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = ¦r§Î¸`ÂI¼Ğ°O
                  Else
                     tree¦r§Î¾ğª¬µ²ºc.ItemTag(tree¦r§Î¾ğª¬µ²ºc.NewIndex) = ºc¦r¦¡¸`ÂI¼Ğ°O
                  End If
               End If
            End If
         End If
         ²Ó³¡¦r§Îªí.MoveNext
      Loop
   End If
  
   ´F¨Åª¬ºA¦C = "¸ü¤J§¹²¦¡I ¦@ " & ²Ó³¡Á`¼Æ & " ­Ó¦r§Î(³¡¥ó)"
   mdiº~¦r¦r§Î.txtª¬ºA = ´F¨Åª¬ºA¦C
   'mdiº~¦r¦r§Î.sbarª¬ºA¦C.Visible = True

End Sub

Private Sub tree¦r§Î¾ğª¬µ²ºc_DragOver(Source As Control, x As Single, y As Single, State As Integer)

tree¦r§Î¾ğª¬µ²ºc.OnDragOver x, y, State

End Sub


Private Sub tree¦r§Î¾ğª¬µ²ºc_LostFocus()

tree¦r§Î¾ğª¬µ²ºc.ListIndex = -1

End Sub

Private Sub tree¦r§Î¾ğª¬µ²ºc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

XCheck = x
YCheck = y

End Sub

Public Sub ¦C¥X¿ï©w¦r¶°¤¤ªº©Ò¦³¦r§Î()

Dim ³¡¥ó As String, LikeSQL As Boolean

³¡¥ó = txtºc¦r¦¡
LikeSQL = mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Checked

txtºc¦r¦¡ = "*"
mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Checked = True
txtºc¦r¦¡_KeyPress (vbKeyReturn)

txtºc¦r¦¡ = ³¡¥ó
mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Checked = LikeSQL

End Sub
