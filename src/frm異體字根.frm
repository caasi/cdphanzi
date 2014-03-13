VERSION 5.00
Object = "{65996203-3B87-11D4-A21F-00E029189826}#6.9#0"; "TLIST6.OCX"
Begin VB.Form frm異體字根 
   Caption         =   "異體字根"
   ClientHeight    =   3816
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5100
   Icon            =   "frm異體字根.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3816
   ScaleWidth      =   5100
   Begin TListProLibCtl.TList tree字形樹狀結構 
      DragIcon        =   "frm異體字根.frx":030A
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
         Name            =   "標楷體"
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
      PictureOpen     =   "frm異體字根.frx":074C
      PictureClosed   =   "frm異體字根.frx":085E
      PictureLeaf     =   "frm異體字根.frx":0970
      PictureMark     =   "frm異體字根.frx":0A82
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
      PicturePalette  =   "frm異體字根.frx":0B94
      ExchangeSerialNumber=   "frm異體字根.frx":0CA6
      DragIconStyle   =   0
      ExchangeDefItemCellDef=   "frm異體字根.frx":0CF3
      _ChkCounter     =   -1
      TreeLinesHighlightColor=   -2113929196
      TreeLinesShadowColor=   -2113929200
   End
   Begin VB.Image imgPinPush 
      Height          =   264
      Left            =   4440
      Picture         =   "frm異體字根.frx":0DFA
      Top             =   840
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imgPin 
      Height          =   264
      Left            =   4440
      Picture         =   "frm異體字根.frx":0F84
      Top             =   1320
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imglock 
      Height          =   264
      Left            =   0
      Picture         =   "frm異體字根.frx":110E
      Tag             =   "0"
      ToolTipText     =   "鎖定"
      Top             =   240
      Width           =   288
   End
   Begin VB.Image Image1 
      Height          =   192
      Left            =   4440
      Picture         =   "frm異體字根.frx":1298
      Top             =   480
      Visible         =   0   'False
      Width           =   192
   End
End
Attribute VB_Name = "frm異體字根"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private 視窗代碼 As Integer, 視窗 As String, 字根表 As Recordset
Private 區域字體陣列(0 To 字體個數) As Variant
Private 中斷 As Boolean
Private 總數  As Long, 部數 As String
Private XCheck As Single, YCheck As Single

Private Sub Form_Activate()
現用視窗 = 視窗
'現用視窗代碼 = 視窗代碼
現用視窗代碼 = 異體字根代碼
切換選取字形工具列狀態 現用視窗代碼
mdi漢字字形.txt狀態 = 異體狀態列

End Sub


Private Sub Form_Load()

Dim i As Integer, 編號 As Long

啟動異體字根 = True

If 初始first <> 1 Then
   If 已載入畫面 = 0 Then
      If 異根winstate = 0 Then
         frm異體字根.Left = 異根left
         frm異體字根.Top = 異根top
         frm異體字根.Height = 異根height
         frm異體字根.Width = 異根width
      Else
         frm異體字根.WindowState = 異根winstate
      End If
   ElseIf 啟動字形孳乳 Then
         frm異體字根.Left = frm字形孳乳.Left + frm字形孳乳.Width
         frm異體字根.Top = frm字形孳乳.Top
         frm異體字根.Height = frm字形孳乳.Height
         frm異體字根.Width = frm字形孳乳.Width
   End If
End If

Set 字根表 = 系統資料庫.OpenRecordset("字根")
字根表.Index = "字形"

i = 0
Do While 字體陣列(i) <> ""
   區域字體陣列(i) = 字體陣列(i)
   i = i + 1
Loop

tree字形樹狀結構.FontSize = CInt(顯示字型大小)
視窗代碼 = 共用視窗代碼
視窗 = 共用視窗(共用視窗代碼)
'Me.Tag = 共用視窗代碼
Me.Tag = 異體字根代碼
'tree字形樹狀結構.AddItem ""
'tree字形樹狀結構.ListIndex = 0
'tree字形樹狀結構.Image(0) = tree字形樹狀結構.PictureLeaf

If 是否為組字符號(mdi漢字字形.txt字形.Text, 1, 14) = 0 Then
   'If Mid(mdi漢字字形.txt字形.FontName, 1, 3) <> "hzk" Then
   '   檢字表.Index = "字形"
   '   檢字表.Seek "=", mdi漢字字形.txt字形.Text
   'Else
        楷書檢字表.Index = "編號"
        楷書檢字表.Seek "=", 系統編號
   'End If
   
    If Not 楷書檢字表.NoMatch Then
        Do While Not 楷書檢字表.NoMatch
            編號 = 楷書檢字表.Fields("編號")
            Exit Do
        楷書檢字表.MoveNext
        Loop
    End If
    載入字形 "標楷體", mdi漢字字形.txt字形.Text, 編號
End If

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
   DoEvents
   中斷 = True
End If

End Sub

Private Sub Form_Resize()

If Me.ScaleHeight - tree字形樹狀結構.Top * 2 > 0 Then tree字形樹狀結構.Height = Me.ScaleHeight - tree字形樹狀結構.Top * 2
If Me.ScaleWidth - tree字形樹狀結構.Left * 2 > 0 Then tree字形樹狀結構.Width = Me.ScaleWidth - tree字形樹狀結構.Left * 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
字根表.Close
啟動異體字根 = False
mdi漢字字形.mnu_異體字根.Enabled = True
計算現用視窗
End Sub

Private Sub imglock_Click()

If imglock.Tag = 0 Then
    imglock.Tag = 1
    imglock.Picture = imgPinPush.Picture
    imglock.ToolTipText = "解除鎖定"
    frm異體字根.Caption = "異體字根(鎖定)"
Else
    imglock.Tag = 0
    imglock.Picture = imgPin.Picture
    imglock.ToolTipText = "鎖定"
    frm異體字根.Caption = "異體字根"
End If

End Sub

Private Sub tree字形樹狀結構_Click()
Dim 字體 As String
Dim 字形 As String
Dim 編號 As Long

If tree字形樹狀結構.ListIndex <> -1 Then
   If tree字形樹狀結構.List(0) <> "" Then
      字體 = tree字形樹狀結構.ItemFontName(tree字形樹狀結構.ListIndex)
      字形 = tree字形樹狀結構.List(tree字形樹狀結構.ListIndex)
      編號 = tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.ListIndex)
    
      擷取屬性 字體, 字形, 編號
      擷取構字式 字體, 字形, 編號
      If mdi漢字字形.txt字形.font.Name = "標楷體" Then 拖曳字串 = tree字形樹狀結構.List(tree字形樹狀結構.ListIndex)
      
      If 啟動字形結構 Then frm字形結構.載入字形 字體, 字形, 編號
      If 啟動異體字表 Then frm異體字表.載入字形 字體, 字形, 編號
      If 啟動字形演變 Then frm字形演變.載入字形 字體, 字形, 編號
      If 啟動字形索引 Then frm字形索引.載入字形 字體, 字形, 編號
      mdi漢字字形.txt狀態 = 異體狀態列
   End If
End If

End Sub

Private Sub tree字形樹狀結構_DragOver(Source As Control, x As Single, y As Single, State As Integer)

tree字形樹狀結構.OnDragOver x, y, State

End Sub

Private Sub tree字形樹狀結構_GotFocus()

'tree字形樹狀結構_Click
現用視窗代碼 = 異體字根代碼

End Sub

Private Sub tree字形樹狀結構_LostFocus()

tree字形樹狀結構.ListIndex = -1

End Sub

Private Sub tree字形樹狀結構_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

XCheck = x
YCheck = y

End Sub

Private Sub tree字形樹狀結構_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

' Initiates dragging only after moving at least 100 twips with the mouse depressed
If (Button And 1) And (XCheck > 0) And (YCheck > 0) And ((Abs(XCheck - x) > 150) Or (Abs(YCheck - y) > 150)) Then
    XCheck = 0
    YCheck = 0  ' Reset mouse coordinates
    If tree字形樹狀結構.ListIndex >= 0 Then
        tree字形樹狀結構.BeforeDrag
        tree字形樹狀結構.Drag 1         ' Start drag
    End If
End If

'If Button = 1 Then
'    tree字形樹狀結構.Drag 1
'End If

End Sub

Public Sub 載入字形(系統字體 As String, val字形 As String, 編號 As Long)
Dim 字形表 As Recordset, 細部字形表 As Recordset
Dim SQL陳述式 As String, 樹根 As String
Dim i As Integer, 總數 As Long, 字形 As String, 形碼 As Integer
Dim 字體 As String, 字體編號 As Integer
Dim 楷書編號 As Long, 小篆編號 As Integer, 金文編號 As Long, 甲骨文編號 As Integer, 楚系文字編號 As Long

If imglock.Tag = 1 Then Exit Sub
If 編號 <= 0 Then Exit Sub

mdi漢字字形.txt狀態 = "字形載入中  ......       " & " , 欲中斷請按 Esc 鍵"
Screen.MousePointer = ccHourglass

字形 = val字形

tree字形樹狀結構.Clear

If 編號 <= 0 Then 編號 = 系統編號
If 系統字體 = "北師大說文小篆" Or 系統字體 = "北師大說文重文" Then
    小篆編號 = 編號
    小篆檢字表.Index = "編號"
    小篆檢字表.Seek "=", 編號
    楷書編號 = 小篆檢字表.Fields("楷書編號")
ElseIf 系統字體 = "中研院金文" Then
    金文編號 = 編號
    金文檢字表.Index = "編號"
    金文檢字表.Seek "=", 編號
    楷書編號 = 金文檢字表.Fields("楷書編號")
ElseIf 系統字體 = "中研院甲骨文" Then
    甲骨文編號 = 編號
    甲骨文檢字表.Index = "編號"
    甲骨文檢字表.Seek "=", 編號
    楷書編號 = 甲骨文檢字表.Fields("楷書編號")
ElseIf 系統字體 = "中研院楚系簡帛文字" Then
    楚系文字編號 = 編號
    楚系文字檢字表.Index = "編號"
    楚系文字檢字表.Seek "=", 編號
    楷書編號 = 楚系文字檢字表.Fields("楷書編號")
Else
    楷書編號 = 編號
End If

Set 檢字表 = 楷書檢字表
檢字表.Index = "編號"
檢字表.Seek "=", 楷書編號
字體編號 = 檢字表.Fields("字體")
字體 = 字體陣列(字體編號)
樹根 = 檢字表.Fields("字碼")

SQL陳述式 = "SELECT 組號,First(形碼),編號 From 異體字根 GROUP BY 組號,編號 HAVING 編號= " & 楷書編號 & " order by 組號"

DoEvents

樹根序數 = 0
總數 = 0

tree字形樹狀結構.Redraw = False

Set 字形表 = 系統資料庫.OpenRecordset(SQL陳述式)

If tree字形樹狀結構.ListCount > 0 Then
   tree字形樹狀結構.RemoveItem (0)
End If

If Not 字形表.EOF Then
   
   字形表.MoveLast
   總數 = 字形表.RecordCount
 
   If 總數 > 1 Or (總數 = 1 And 字形表.Fields(1) <> 1) Then
      tree字形樹狀結構.AddItem 樹根
      tree字形樹狀結構.ItemFontName(0) = 轉換顯示字型(字體)
      tree字形樹狀結構.Image(0) = tree字形樹狀結構.PictureOpen
      
      tree字形樹狀結構.ItemLngValue(0) = 楷書編號
      If Not IsNull(檢字表.Fields("字形")) Then
         tree字形樹狀結構.ItemTag(0) = 字形節點標記
      Else
         tree字形樹狀結構.ItemTag(0) = 構字式節點標記
      End If
      tree字形樹狀結構.Expand(0) = True
   End If
    
   字形表.MoveFirst
    
   Do Until 字形表.EOF
      If 中斷 = True Then Exit Do
         
      SQL陳述式 = "SELECT * From 異體字根 Where 組號= " & 字形表.Fields("組號") & " ORDER BY 形碼,編號"
      Set 細部字形表 = 系統資料庫.OpenRecordset(SQL陳述式)
            
      細部字形表.MoveFirst
      檢字表.Index = "編號"

      Do Until 細部字形表.EOF
         樹根序數 = 樹根序數 + 1

         檢字表.Seek "=", 細部字形表.Fields("編號")
         If Not 檢字表.NoMatch Then
            If Not IsNull(檢字表.Fields("字形")) And 檢字表.Fields("字形") <> "" Then
               字形 = 檢字表.Fields("字形")
            Else
               If Not IsNull(檢字表.Fields("字碼")) And 檢字表.Fields("字碼") <> "" Then
                  字形 = 檢字表.Fields("字碼")
               Else
                  字形 = "●"
               End If
            End If
          End If
   
          If 細部字形表.Fields("形碼") <> 1 Then
             tree字形樹狀結構.AddItem 字形, i
             tree字形樹狀結構.ItemFontName(tree字形樹狀結構.NewIndex) = 轉換顯示字型(區域字體陣列(檢字表.Fields("字體")))
             tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = 檢字表.Fields("編號")
             tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 字形節點標記
             If 細部字形表.Fields("形碼") = 2 Then
                'tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PicturePalette '簡體字
                tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureMark '簡體字
             'ElseIf 細部字形表.Fields("形碼") > 100 Then   '形碼>100
             '   tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = Image1
             ElseIf 檢字表.Fields("連接符號") = 0 Then
                tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureMark  '字根
             Else
                tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf  '部件
             End If
          Else
             'If 總數 > 1 Or (總數 = 1 And 字形表.Fields("形碼") <> 1) Then
             If 總數 > 1 Or (總數 = 1 And 字形表.Fields(1) <> 1) Then
                tree字形樹狀結構.AddItem 字形, 0
                tree字形樹狀結構.ItemFontName(tree字形樹狀結構.NewIndex) = 轉換顯示字型(區域字體陣列(檢字表.Fields("字體")))
                tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = 檢字表.Fields("編號")
                tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 字形節點標記
                tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureClosed
                i = tree字形樹狀結構.NewIndex
             Else
                tree字形樹狀結構.AddItem 字形
                tree字形樹狀結構.ItemFontName(0) = 轉換顯示字型(區域字體陣列(檢字表.Fields("字體")))
                tree字形樹狀結構.Image(0) = tree字形樹狀結構.PictureOpen
                tree字形樹狀結構.ItemLngValue(0) = 檢字表.Fields("編號")
                If Not IsNull(檢字表.Fields("字形")) Then
                    tree字形樹狀結構.ItemTag(0) = 字形節點標記
                Else
                    tree字形樹狀結構.ItemTag(0) = 構字式節點標記
                End If

                i = 0
             End If
          End If
          細部字形表.MoveNext
       Loop
       tree字形樹狀結構.Expand(i) = True
        
       字形表.MoveNext
         
       If (tree字形樹狀結構.ListCount + 10) Mod 50 = 0 Then
          tree字形樹狀結構.Redraw = True
          Screen.MousePointer = ccDefault
          DoEvents
          tree字形樹狀結構.Redraw = False
          Screen.MousePointer = ccHourglass
       End If
   Loop
Else
   tree字形樹狀結構.AddItem 樹根
   tree字形樹狀結構.ItemFontName(0) = 轉換顯示字型(字體)
   tree字形樹狀結構.ItemLngValue(0) = 楷書編號
   
   檢字表.Index = "編號"
   檢字表.Seek "=", 編號
   If Not 檢字表.NoMatch Then
      If Not IsNull(檢字表.Fields("字形")) Then
         tree字形樹狀結構.ItemTag(0) = 字形節點標記
      Else
         tree字形樹狀結構.ItemTag(0) = 構字式節點標記
      End If

      If 檢字表.Fields("連接符號") = 0 Then
         tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureMark
      Else
         tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
      End If
   End If

End If

tree字形樹狀結構.Redraw = True
異體狀態列 = "載入完畢！ 共 " & 樹根序數 & " 個字形"
mdi漢字字形.txt狀態 = 異體狀態列
'mdi漢字字形.sbar狀態列.Visible = True

Screen.MousePointer = ccDefault

End Sub
