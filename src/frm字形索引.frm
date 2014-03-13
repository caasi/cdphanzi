VERSION 5.00
Object = "{65996203-3B87-11D4-A21F-00E029189826}#6.9#0"; "TLIST6.OCX"
Begin VB.Form frm字形索引 
   Caption         =   "字形索引"
   ClientHeight    =   6228
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9060
   Icon            =   "frm字形索引.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6228
   ScaleWidth      =   9060
   Begin TListProLibCtl.TList tree字形樹狀結構 
      DragIcon        =   "frm字形索引.frx":030A
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
      PictureOpen     =   "frm字形索引.frx":074C
      PictureClosed   =   "frm字形索引.frx":085E
      PictureLeaf     =   "frm字形索引.frx":0970
      PictureMark     =   "frm字形索引.frx":0A82
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
      ExchangeSerialNumber=   "frm字形索引.frx":0B7C
      DragIconStyle   =   0
      ExchangeDefItemCellDef=   "frm字形索引.frx":0BC9
      _ChkCounter     =   -1
      TreeLinesHighlightColor=   -2113929196
      TreeLinesShadowColor=   -2113929200
   End
   Begin VB.Image imgPinPush 
      Height          =   264
      Left            =   4440
      Picture         =   "frm字形索引.frx":0CD0
      Top             =   240
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imgPin 
      Height          =   264
      Left            =   4440
      Picture         =   "frm字形索引.frx":0E5A
      Top             =   720
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imglock 
      Height          =   264
      Left            =   0
      Picture         =   "frm字形索引.frx":0FE4
      Tag             =   "0"
      ToolTipText     =   "鎖定"
      Top             =   240
      Width           =   288
   End
End
Attribute VB_Name = "frm字形索引"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private 視窗代碼 As Integer, 視窗 As String
Private 區域字體陣列(0 To 字體個數) As Variant
Private XCheck As Single, YCheck As Single

Private Sub Form_Activate()
現用視窗 = 視窗
'現用視窗代碼 = 視窗代碼
現用視窗代碼 = 字形索引代碼
切換選取字形工具列狀態 現用視窗代碼
tree字形樹狀結構_Click
mdi漢字字形.txt狀態 = 結構狀態列

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim 字根序 As String, 編號 As Long

啟動字形索引 = True
If 初始first <> 1 Then
   If 已載入畫面 = 0 Then
      If 索引winstate = 0 Then
         frm字形索引.Left = 索引left
         frm字形索引.Top = 索引top
         frm字形索引.Height = 索引height
         frm字形索引.Width = 索引width
      Else
         frm字形索引.WindowState = 索引winstate
      End If
   ElseIf 啟動字形孳乳 Then
         frm字形索引.Left = frm字形孳乳.Left + frm字形孳乳.Width
         frm字形索引.Top = frm字形孳乳.Top
         frm字形索引.Height = frm字形孳乳.Height
         frm字形索引.Width = frm字形孳乳.Width
   End If
End If

tree字形樹狀結構.FontSize = CInt(顯示字型大小)
'Me.Tag = 共用視窗代碼
Me.Tag = 字形索引代碼
視窗代碼 = 共用視窗代碼
視窗 = 共用視窗(共用視窗代碼)
tree字形樹狀結構.AddItem ""
'tree字形樹狀結構.ListIndex = 0
tree字形樹狀結構.Image(0) = tree字形樹狀結構.PictureLeaf

i = 0
Do While 字體陣列(i) <> ""
   區域字體陣列(i) = 字體陣列(i)
   i = i + 1
Loop

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

Private Sub Form_Resize()

If Me.ScaleHeight - tree字形樹狀結構.Top * 2 > 0 Then tree字形樹狀結構.Height = Me.ScaleHeight - tree字形樹狀結構.Top * 2
If Me.ScaleWidth - tree字形樹狀結構.Left * 2 > 0 Then tree字形樹狀結構.Width = Me.ScaleWidth - tree字形樹狀結構.Left * 2

End Sub


Private Sub Form_Unload(Cancel As Integer)
mdi漢字字形.mnu_字形索引.Enabled = True
啟動字形索引 = False
計算現用視窗

End Sub

Public Sub 載入字形(系統字體 As String, 字形 As String, 編號 As Long)
Dim 部件序 As String, 字體編號 As Integer, 字型檔 As String, 分解 As Integer
Dim i As Integer, 字體 As String
Dim 楷書編號 As Long, 小篆編號 As Integer, 金文編號 As Long, 甲骨文編號 As Integer, 楚系文字編號 As Long
Dim 楷書字形 As String, 小篆字形 As String, 金文字形 As String, 甲骨文字形 As String, 楚系文字字形 As String
Dim 小篆字源 As String, 金文字源 As String, 甲骨文字源 As String, 楚系文字字源 As String
Dim 金文檢字 As Boolean, 金文補遺 As Boolean
Dim 楚系文字檢字 As Boolean, 楚系文字補遺 As Boolean
Dim ItemNo As Integer, ItemBig5 As Integer, ItemUnicode As Integer
Dim Item漢語大字典 As Integer, Item遠東漢語大字典 As Integer, Item建宏漢語大字典 As Integer
Dim Item中文大辭典 As Integer
Dim Item說文 As Integer, Item說文詁林 As Integer, Item說文中華 As Integer
Dim Item金文 As Integer, Item金文編 As Integer, Item金文詁林 As Integer, Item金文器號 As Integer, Item金文引得 As Integer
Dim Item甲骨文 As Integer, Item甲骨刻辭類纂 As Integer, Item甲骨文字詁林 As Integer, Item甲骨文字集釋 As Integer
Dim Item楚系文字, Item楚系簡帛文字編 As Integer, Item楚系文字出處 As Integer
Dim 部首 As String, 金文字頭 As Integer

If imglock.Tag = 1 Then Exit Sub
If 編號 <= 0 Then Exit Sub

Screen.MousePointer = ccHourglass

If 系統字體 = "北師大說文小篆" Or 系統字體 = "北師大說文重文" Then
    小篆編號 = 編號
    小篆檢字表.Index = "編號"
    小篆檢字表.Seek "=", 編號
    楷書編號 = 小篆檢字表.Fields("楷書編號")
ElseIf 中研院金文(系統字體) Then
    金文編號 = 編號
    金文異寫字表.Index = "編號"
    金文異寫字表.Seek "=", 編號
    If Not IsNull(金文異寫字表.Fields("楷書編號")) Then
        楷書編號 = 金文異寫字表.Fields("楷書編號")
    Else
        金文檢字表.Index = "編號"
        金文檢字表.Seek "=", 編號
        楷書編號 = 金文檢字表.Fields("楷書編號")
    End If
ElseIf 中研院甲骨文(系統字體) Then
    甲骨文編號 = 編號
    甲骨文異寫字表.Index = "編號"
    甲骨文異寫字表.Seek "=", 編號
    If Not IsNull(甲骨文異寫字表.Fields("楷書編號")) Then
        楷書編號 = 甲骨文異寫字表.Fields("楷書編號")
    Else
        甲骨文檢字表.Index = "編號"
        甲骨文檢字表.Seek "=", 編號
        楷書編號 = 甲骨文檢字表.Fields("楷書編號")
    End If
ElseIf 中研院楚系文字(系統字體) Then
    楚系文字編號 = 編號
    楚系文字檢字表.Index = "編號"
    楚系文字檢字表.Seek "=", 編號
    If Not IsNull(楚系文字異寫字表.Fields("楷書編號")) Then
        楷書編號 = 楚系文字異寫字表.Fields("楷書編號")
    Else
        楚系文字檢字表.Index = "編號"
        楚系文字檢字表.Seek "=", 編號
        楷書編號 = 楚系文字檢字表.Fields("楷書編號")
    End If
Else
    楷書編號 = 編號
End If

楷書檢字表.Index = "編號"
楷書檢字表.Seek "=", 楷書編號
楷書字形 = 楷書檢字表.Fields("字碼")
字體編號 = 楷書檢字表.Fields("字體")
字體 = 區域字體陣列(字體編號)

'If 系統字體 <> "北師大說文小篆" And 系統字體 <> "北師大說文重文" Then
'    If Not IsNull(楷書檢字表.Fields("小篆編號")) Then
'        小篆編號 = 楷書檢字表.Fields("小篆編號")
'    Else
'        小篆編號 = 0
'    End If
'End If

'If 系統字體 <> "中研院金文" Then
'    If Not IsNull(楷書檢字表.Fields("金文編號")) Then
'        金文編號 = 楷書檢字表.Fields("金文編號")
'    Else
'        金文編號 = 0
'    End If
'End If

tree字形樹狀結構.Clear
tree字形樹狀結構.Redraw = False
tree字形樹狀結構.FontName = 顯示字型

tree字形樹狀結構.AddItem 楷書字形
tree字形樹狀結構.ItemFontName(0) = 轉換顯示字型(字體)
tree字形樹狀結構.ItemLngValue(0) = 楷書編號
If Not IsNull(楷書檢字表.Fields("字形")) Then
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 字形節點標記
Else
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 構字式節點標記
End If

'If 小篆編號 = 0 And 金文編號 = 0 Then
'    tree字形樹狀結構.Image(0) = tree字形樹狀結構.PictureLeaf
'    tree字形樹狀結構.Redraw = True
'    Screen.MousePointer = ccDefault
'    Exit Sub
'End If
tree字形樹狀結構.Image(0) = tree字形樹狀結構.PictureOpen

ItemNo = 1
Item漢語大字典 = -1
Item遠東漢語大字典 = -1
Item建宏漢語大字典 = -1
Item中文大辭典 = -1
Item說文 = -1
Item說文詁林 = -1
Item說文中華 = -1
Item金文 = -1
Item金文編 = -1
Item金文詁林 = -1
Item金文器號 = -1
Item金文引得 = -1
Item甲骨文 = -1
Item甲骨刻辭類纂 = -1
Item甲骨文字詁林 = -1
Item甲骨文字集釋 = -1
Item楚系文字 = -1
Item楚系簡帛文字編 = -1
Item楚系文字出處 = -1
ItemBig5 = -1
ItemUnicode = -1

If Not IsNull(楷書檢字表.Fields("漢語大字典")) And (mdi漢字字形.mnu_遠東漢語大字典選項.Checked Or mdi漢字字形.mnu_建宏漢語大字典選項.Checked) Then
    Item漢語大字典 = tree字形樹狀結構.ListCount
    tree字形樹狀結構.AddItem "漢語大字典", 0
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
    If Not IsNull(檢字表.Fields("部首")) Then 部首 = 尋找部首(楷書檢字表.Fields("部首"))
    tree字形樹狀結構.AddItem "部首:" & 部首, Item漢語大字典
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    
    If Not IsNull(楷書檢字表.Fields("遠東漢語大字典")) And mdi漢字字形.mnu_遠東漢語大字典選項.Checked Then
        Item遠東漢語大字典 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "遠東圖書公司", Item漢語大字典
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "冊-頁-字:" & 楷書檢字表.Fields("遠東漢語大字典"), Item遠東漢語大字典
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If

    If Not IsNull(楷書檢字表.Fields("建宏漢語大字典")) And mdi漢字字形.mnu_建宏漢語大字典選項.Checked Then
        Item建宏漢語大字典 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "建宏出版社", Item漢語大字典
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "頁-字:" & 楷書檢字表.Fields("建宏漢語大字典"), Item建宏漢語大字典
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If

End If

If Not IsNull(楷書檢字表.Fields("中文大辭典索引")) And mdi漢字字形.mnu_中文大辭典選項.Checked Then
    Item中文大辭典 = tree字形樹狀結構.ListCount
    tree字形樹狀結構.AddItem "中文大辭典(中華學術院)", 0
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
    If Not IsNull(檢字表.Fields("中文大辭典部首")) Then 部首 = 尋找部首(楷書檢字表.Fields("中文大辭典部首"))
    tree字形樹狀結構.AddItem "部首:" & 部首, Item中文大辭典
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    
    tree字形樹狀結構.AddItem "冊-編號:" & 楷書檢字表.Fields("中文大辭典索引"), Item中文大辭典
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
End If

If Not IsNull(楷書檢字表.Fields("小篆編號")) And (mdi漢字字形.mnu_說文解字詁林選項.Checked Or mdi漢字字形.mnu_中華說文解字選項.Checked) Then
    小篆編號 = 楷書檢字表.Fields("小篆編號")
    小篆檢字表.Index = "編號"
    小篆檢字表.Seek "=", 小篆編號

    Item說文 = tree字形樹狀結構.ListCount
    tree字形樹狀結構.AddItem "說文解字", 0
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
    部首 = 尋找說文部首(小篆檢字表.Fields("部首"))
    tree字形樹狀結構.AddItem "部首:" & 部首, Item說文
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    
    tree字形樹狀結構.AddItem "卷:" & 小篆檢字表.Fields("卷"), Item說文
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    
    If Not IsNull(小篆檢字表.Fields("詁林索引")) And mdi漢字字形.mnu_說文解字詁林選項.Checked Then
        Item說文詁林 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "說文解字詁林(鼎文書局)", Item說文
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "冊-頁:" & 小篆檢字表.Fields("詁林索引"), Item說文詁林
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If
    
    If Not IsNull(小篆檢字表.Fields("中華索引")) And mdi漢字字形.mnu_中華說文解字選項.Checked Then
        Item說文中華 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "說文解字(中華書局)", Item說文
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "頁(上/下):" & 小篆檢字表.Fields("中華索引"), Item說文中華
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If
End If

金文檢字 = False
金文補遺 = False

If Not IsNull(楷書檢字表.Fields("金文編號")) Then
    金文檢字 = True
Else
    金文集成引得.Seek "=", 楷書編號
    If Not 金文集成引得.NoMatch Then
        金文補遺 = True
    Else
        金文補遺表.Seek "=", 楷書編號
        If Not 金文補遺表.NoMatch Then 金文補遺 = True
    End If
End If

If 金文檢字 And (mdi漢字字形.mnu_金文編選項.Checked Or mdi漢字字形.mnu_金文詁林選項.Checked Or mdi漢字字形.mnu_殷周金文集成器號選項.Checked Or mdi漢字字形.mnu_殷周金文集成引得選項.Checked) Then
    If Not 中研院金文(系統字體) Then
        金文編號 = 楷書檢字表.Fields("金文編號")
    End If
    金文異寫字表.Index = "編號"
    金文異寫字表.Seek "=", 金文編號
    
    金文字頭 = 金文異寫字表.Fields("組號")
    
    Item金文 = tree字形樹狀結構.ListCount
    tree字形樹狀結構.AddItem "金文", 0
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
    If Not 金文異寫字表.NoMatch And mdi漢字字形.mnu_金文編選項.Checked Then
        Item金文編 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "金文編(中華書局)", Item金文
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "編號-行-字:" & 金文異寫字表.Fields("金文編索引"), Item金文編
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If
    
    金文詁林.Seek "=", 金文字頭
    If Not 金文詁林.NoMatch And mdi漢字字形.mnu_金文詁林選項.Checked Then
        Item金文詁林 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "金文詁林(香港中文大學)", Item金文
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "卷-頁:" & 金文詁林.Fields("索引"), Item金文詁林
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If
    
    'If Not IsNull(金文異寫字表.Fields("器號")) And mdi漢字字形.mnu_殷周金文集成器號選項.Checked Then
        'Item金文器號 = tree字形樹狀結構.ListCount
        'tree字形樹狀結構.AddItem "殷周金文集成(中國社科院考古所)", Item金文
        'tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        'tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        'tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        'tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        'tree字形樹狀結構.AddItem "器號:" & 金文異寫字表.Fields("器號"), Item金文器號
        'tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        'tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        'tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        'tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    'End If
    
    金文集成引得.Seek "=", 楷書編號
    If Not 金文集成引得.NoMatch And mdi漢字字形.mnu_殷周金文集成引得選項.Checked Then
        Item金文引得 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "殷周金文集成引得(中華書局)", Item金文
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "編號-頁碼:" & 金文集成引得.Fields("索引"), Item金文引得
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If
End If

If 金文補遺 And (mdi漢字字形.mnu_金文編選項.Checked Or mdi漢字字形.mnu_金文詁林選項.Checked Or mdi漢字字形.mnu_殷周金文集成器號選項.Checked Or mdi漢字字形.mnu_殷周金文集成引得選項.Checked) Then

    
    Item金文 = tree字形樹狀結構.ListCount
    tree字形樹狀結構.AddItem "金文", 0
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
    金文補遺表.Seek "=", 楷書編號
    If 金文補遺表.NoMatch Then GoTo 填寫集成引得索引
    
    If Not IsNull(金文補遺表.Fields("金文編索引")) And mdi漢字字形.mnu_金文編選項.Checked Then
        Item金文編 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "金文編(中華書局)", Item金文
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "編號-行-字:" & 金文補遺表.Fields("金文編索引"), Item金文編
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If
    
    If Not IsNull(金文補遺表.Fields("金文詁林索引")) And mdi漢字字形.mnu_金文詁林選項.Checked Then
        Item金文詁林 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "金文詁林(香港中文大學)", Item金文
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "卷-頁:" & 金文補遺表.Fields("金文詁林索引"), Item金文詁林
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If
    
填寫集成引得索引:

    金文集成引得.Seek "=", 楷書編號
    If Not 金文集成引得.NoMatch And mdi漢字字形.mnu_殷周金文集成引得選項.Checked Then
        Item金文引得 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "殷周金文集成引得(中華書局)", Item金文
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "編號-頁碼:" & 金文集成引得.Fields("索引"), Item金文引得
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If
End If

If Not IsNull(楷書檢字表.Fields("甲骨文編號")) And (mdi漢字字形.mnu_甲骨刻辭類纂選項.Checked Or mdi漢字字形.mnu_甲骨文字詁林選項.Checked Or mdi漢字字形.mnu_甲骨文字集釋選項.Checked) Then
    If Not 中研院甲骨文(系統字體) Then
        甲骨文編號 = 楷書檢字表.Fields("甲骨文編號")
    End If
    甲骨文異寫字表.Index = "編號"
    甲骨文異寫字表.Seek "=", 甲骨文編號

    Item甲骨文 = tree字形樹狀結構.ListCount
    tree字形樹狀結構.AddItem "甲骨文", 0
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
    If Not IsNull(甲骨文異寫字表.Fields("甲骨刻辭類纂")) And mdi漢字字形.mnu_甲骨刻辭類纂選項.Checked Then
        Item甲骨刻辭類纂 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "殷墟甲骨刻辭類纂(吉林大學)", Item甲骨文
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "冊-頁:" & 甲骨文異寫字表.Fields("甲骨刻辭類纂"), Item甲骨刻辭類纂
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If
    
    If Not IsNull(甲骨文異寫字表.Fields("甲骨文字詁林")) And mdi漢字字形.mnu_甲骨文字詁林選項.Checked Then
        Item甲骨文字詁林 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "甲骨文字詁林(吉林大學)", Item甲骨文
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "冊-頁:" & 甲骨文異寫字表.Fields("甲骨文字詁林"), Item甲骨文字詁林
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If
    
    If Not IsNull(甲骨文異寫字表.Fields("甲骨文字集釋")) And mdi漢字字形.mnu_甲骨文字集釋選項.Checked Then
        Item甲骨文字集釋 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "甲骨文字集釋(中央研究院)", Item甲骨文
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "卷-頁:" & 甲骨文異寫字表.Fields("甲骨文字集釋"), Item甲骨文字集釋
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If
    
End If

楚系文字檢字 = False
楚系文字補遺 = False

If Not IsNull(楷書檢字表.Fields("楚系文字編號")) Then
    楚系文字檢字 = True
Else
    楚系文字補遺表.Seek "=", 楷書編號
    If Not 楚系文字補遺表.NoMatch Then 楚系文字補遺 = True
End If

If 楚系文字檢字 And (mdi漢字字形.mnu_楚系簡帛文字編選項.Checked Or mdi漢字字形.mnu_楚系文字出處選項.Checked) Then
    If Not 中研院楚系文字(系統字體) Then
        楚系文字編號 = 楷書檢字表.Fields("楚系文字編號")
    End If
    楚系文字異寫字表.Index = "編號"
    楚系文字異寫字表.Seek "=", 楚系文字編號

    Item楚系文字 = tree字形樹狀結構.ListCount
    tree字形樹狀結構.AddItem "楚系文字", 0
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
    If Not IsNull(楚系文字異寫字表.Fields("楚系簡帛文字編")) And mdi漢字字形.mnu_楚系簡帛文字編選項.Checked Then
        Item楚系簡帛文字編 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "楚系簡帛文字編(湖北教育出版社)", Item楚系文字
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "頁碼-行字:" & 楚系文字異寫字表.Fields("楚系簡帛文字編"), Item楚系簡帛文字編
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If
    
    'If Not IsNull(楚系文字檢字表.Fields("出處")) And mdi漢字字形.mnu_楚系文字出處選項.Checked Then
    '    Item楚系文字出處 = tree字形樹狀結構.ListCount
    '    tree字形樹狀結構.AddItem "出處", Item楚系文字
    '    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    '    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    '    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    '    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
    '    tree字形樹狀結構.AddItem "墓號-簡號:" & 楚系文字檢字表.Fields("出處"), Item楚系文字出處
    '    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    '    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    '    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    '    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    'End If

End If

If 楚系文字補遺 And (mdi漢字字形.mnu_楚系簡帛文字編選項.Checked Or mdi漢字字形.mnu_楚系文字出處選項.Checked) Then

    Item楚系文字 = tree字形樹狀結構.ListCount
    tree字形樹狀結構.AddItem "楚系文字", 0
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
    If Not IsNull(楚系文字補遺表.Fields("楚系簡帛文字編")) And mdi漢字字形.mnu_楚系簡帛文字編選項.Checked Then
        Item楚系簡帛文字編 = tree字形樹狀結構.ListCount
        tree字形樹狀結構.AddItem "楚系簡帛文字編(湖北教育出版社)", Item楚系文字
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
        tree字形樹狀結構.AddItem "頁碼-行字:" & 楚系文字補遺表.Fields("楚系簡帛文字編"), Item楚系簡帛文字編
        tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    End If

End If

If Not IsNull(楷書檢字表.Fields("Unicode")) And mdi漢字字形.mnu_Unicode選項.Checked Then
    ItemUnicode = tree字形樹狀結構.ListCount
    tree字形樹狀結構.AddItem "Unicode 3.2", 0
    'tree字形樹狀結構.ItemFontName(tree字形樹狀結構.NewIndex) = 轉換顯示字型("標楷體")
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
    tree字形樹狀結構.AddItem 楷書檢字表.Fields("Unicode"), ItemUnicode
    'tree字形樹狀結構.ItemFontName(tree字形樹狀結構.NewIndex) = 轉換顯示字型("標楷體")
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
End If

If 楷書編號 > 0 And 楷書編號 <= 13060 And mdi漢字字形.mnu_Big5選項.Checked Then
    ItemBig5 = tree字形樹狀結構.ListCount
    tree字形樹狀結構.AddItem "Big5", 0
    'tree字形樹狀結構.ItemFontName(tree字形樹狀結構.NewIndex) = 轉換顯示字型("標楷體")
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureOpen
    
    tree字形樹狀結構.AddItem 楷書檢字表.Fields("Big5"), ItemBig5
    'tree字形樹狀結構.ItemFontName(tree字形樹狀結構.NewIndex) = 轉換顯示字型("標楷體")
    tree字形樹狀結構.ItemFontSize(tree字形樹狀結構.NewIndex) = 12
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = -9999
    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 其他節點標記
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
End If

tree字形樹狀結構.Expand(0) = True
If Item漢語大字典 > -1 Then tree字形樹狀結構.Expand(Item漢語大字典) = True
If Item遠東漢語大字典 > -1 Then tree字形樹狀結構.Expand(Item遠東漢語大字典) = True
If Item建宏漢語大字典 > -1 Then tree字形樹狀結構.Expand(Item建宏漢語大字典) = True
If Item中文大辭典 > -1 Then tree字形樹狀結構.Expand(Item中文大辭典) = True
If Item說文 > -1 Then tree字形樹狀結構.Expand(Item說文) = True
If Item說文詁林 > -1 Then tree字形樹狀結構.Expand(Item說文詁林) = True
If Item說文中華 > -1 Then tree字形樹狀結構.Expand(Item說文中華) = True
If Item金文 > -1 Then tree字形樹狀結構.Expand(Item金文) = True
If Item金文編 > -1 Then tree字形樹狀結構.Expand(Item金文編) = True
If Item金文詁林 > -1 Then tree字形樹狀結構.Expand(Item金文詁林) = True
If Item金文器號 > -1 Then tree字形樹狀結構.Expand(Item金文器號) = True
If Item金文引得 > -1 Then tree字形樹狀結構.Expand(Item金文引得) = True
If Item甲骨文 > -1 Then tree字形樹狀結構.Expand(Item甲骨文) = True
If Item甲骨刻辭類纂 > -1 Then tree字形樹狀結構.Expand(Item甲骨刻辭類纂) = True
If Item甲骨文字詁林 > -1 Then tree字形樹狀結構.Expand(Item甲骨文字詁林) = True
If Item甲骨文字集釋 > -1 Then tree字形樹狀結構.Expand(Item甲骨文字集釋) = True
If Item楚系文字 > -1 Then tree字形樹狀結構.Expand(Item楚系文字) = True
If Item楚系簡帛文字編 > -1 Then tree字形樹狀結構.Expand(Item楚系簡帛文字編) = True
If Item楚系文字出處 > -1 Then tree字形樹狀結構.Expand(Item楚系文字出處) = True
If ItemBig5 > -1 Then tree字形樹狀結構.Expand(ItemBig5) = True
If ItemUnicode > -1 Then tree字形樹狀結構.Expand(ItemUnicode) = True
If tree字形樹狀結構.ListCount = 1 Then tree字形樹狀結構.Image(0) = tree字形樹狀結構.PictureLeaf


tree字形樹狀結構.Redraw = True
結構狀態列 = ""
mdi漢字字形.txt狀態 = 結構狀態列

Screen.MousePointer = ccDefault

End Sub

Private Sub imglock_Click()

If imglock.Tag = 0 Then
    imglock.Tag = 1
    imglock.Picture = imgPinPush.Picture
    imglock.ToolTipText = "解除鎖定"
    frm字形索引.Caption = "字形索引(鎖定)"
Else
    imglock.Tag = 0
    imglock.Picture = imgPin.Picture
    imglock.ToolTipText = "鎖定"
    frm字形索引.Caption = "字形索引"
End If

End Sub

Private Sub tree字形樹狀結構_Click()
Dim 字體 As String
Dim 字形 As String
Dim 編號 As Long

現用視窗代碼 = 字形索引代碼

If tree字形樹狀結構.ListIndex <> -1 Then
   If Len(tree字形樹狀結構.List(tree字形樹狀結構.ListIndex)) = 1 Then
      字體 = tree字形樹狀結構.ItemFontName(tree字形樹狀結構.ListIndex)
      字形 = tree字形樹狀結構.List(tree字形樹狀結構.ListIndex)
      編號 = tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.ListIndex)
      擷取屬性 字體, 字形, 編號
      擷取構字式 字體, 字形, 編號
      If mdi漢字字形.txt字形.font.Name = "標楷體" Then 拖曳字串 = tree字形樹狀結構.List(tree字形樹狀結構.ListIndex)
      If 啟動字形結構 Then frm字形結構.載入字形 字體, 字形, 編號
      If 啟動異體字表 Then frm異體字表.載入字形 字體, 字形, 編號
      If 啟動異體字根 Then frm異體字根.載入字形 字體, 字形, 編號
    End If
End If

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
'    tree字形樹狀結構.BeforeDrag
'    tree字形樹狀結構.Drag 1
'End If

End Sub

Private Sub tree字形樹狀結構_DragOver(Source As Control, x As Single, y As Single, State As Integer)

tree字形樹狀結構.OnDragOver x, y, State

End Sub


Private Sub tree字形樹狀結構_LostFocus()

tree字形樹狀結構.ListIndex = -1

End Sub

Private Sub tree字形樹狀結構_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

XCheck = x
YCheck = y

End Sub
