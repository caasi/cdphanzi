VERSION 5.00
Object = "{65996203-3B87-11D4-A21F-00E029189826}#6.9#0"; "TLIST6.OCX"
Begin VB.Form frm部件範例 
   Caption         =   "部件"
   ClientHeight    =   3132
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "標楷體"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm構字部件.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3132
   ScaleWidth      =   5640
   Begin TListProLibCtl.TList tree字形樹狀結構 
      DragIcon        =   "frm構字部件.frx":030A
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
      PictureOpen     =   "frm構字部件.frx":074C
      PictureClosed   =   "frm構字部件.frx":085E
      PictureLeaf     =   "frm構字部件.frx":0970
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
      ExchangeSerialNumber=   "frm構字部件.frx":0A82
      DragIconStyle   =   0
      ExchangeDefItemCellDef=   "frm構字部件.frx":0ACF
      _ChkCounter     =   -1
      TreeLinesHighlightColor=   -2113929196
      TreeLinesShadowColor=   -2113929200
   End
End
Attribute VB_Name = "frm部件範例"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private 筆畫 As Integer, 首筆 As Integer
Private 視窗代碼 As Integer, 視窗 As String, 狀態列 As String
Private XCheck As Single, YCheck As Single

Public Sub 部件查詢(部件筆畫 As Integer, 部件首筆 As Integer)
Dim 字形表 As Recordset, 字形 As String
Dim 條件 As String

Screen.MousePointer = ccHourglass

筆畫 = 部件筆畫
首筆 = 部件首筆

tree字形樹狀結構.Clear
tree字形樹狀結構.Redraw = False

If 筆畫 = 0 Then
    條件 = "1-99畫"
Else
    條件 = 筆畫 & "畫"
End If

If 首筆 > 0 Then 條件 = 條件 & "(" & mdi漢字字形.cbo首筆.List(首筆) & ")"
If 視窗 = "構字符號" Then
    tree字形樹狀結構.AddItem "構字符號"
ElseIf 視窗 = "圖形文字" Then
    tree字形樹狀結構.AddItem "圖形文字"
ElseIf 視窗 = "八卦" Then
    tree字形樹狀結構.AddItem "八卦"
ElseIf 視窗 = "簡牘" Then
    tree字形樹狀結構.AddItem "簡牘"
Else
    tree字形樹狀結構.AddItem 條件
End If
tree字形樹狀結構.ItemTag(0) = 其他節點標記

Set 字形表 = 系統資料庫.OpenRecordset(部件外字SQL陳述式(視窗, 筆畫, 首筆))

tree字形樹狀結構.Font.Name = 顯示字型

Set 檢字表 = 楷書檢字表
檢字表.Index = "字形"

If Not 字形表.EOF Then
   Do Until 字形表.EOF
      字形 = 字形表.Fields("字形")
      檢字表.Seek "=", 字形
      tree字形樹狀結構.AddItem 字形, 0
      tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = 檢字表.Fields("編號")
      tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 字形節點標記
      字形表.MoveNext
   Loop
   'If tree字形樹狀結構.ListCount > 0 Then tree字形樹狀結構.ListIndex = -1
End If
部件狀態列 = tree字形樹狀結構.ListCount & " 個部件"
mdi漢字字形.txt狀態 = 部件狀態列

字形表.Close

tree字形樹狀結構.Expand(0) = True
tree字形樹狀結構.Redraw = True
Screen.MousePointer = ccDefault

End Sub

Public Sub Form_Activate()
現用視窗 = 視窗
現用視窗代碼 = 視窗代碼
'mdi漢字字形.mnu_部件代碼(現用視窗代碼).Checked = True
切換選取字形工具列狀態 現用視窗代碼, 筆畫, 首筆
mdi漢字字形.txt狀態 = 部件狀態列

End Sub

Public Sub Form_Load()

If 初始first <> 1 Then
   If 已載入畫面 = 0 Then
      If 部件winstate = 0 Then
         frm部件範例.Left = 部件left
         frm部件範例.Top = 部件top
         frm部件範例.Height = 部件height
         frm部件範例.Width = 部件width
      Else
         frm部件範例.WindowState = 部件winstate
      End If
   ElseIf 啟動字形孳乳 And Not 啟動部件範例 Then
         frm部件範例.Left = frm字形孳乳.Left + frm字形孳乳.Width
         frm部件範例.Top = frm字形孳乳.Top
         frm部件範例.Height = frm字形孳乳.Height
         frm部件範例.Width = frm字形孳乳.Width
   End If
End If

tree字形樹狀結構.Font.Size = 顯示字型大小
筆畫 = mdi漢字字形.cbo筆畫.ListIndex
首筆 = mdi漢字字形.cbo首筆.ListIndex
視窗代碼 = 共用視窗代碼
視窗 = 共用視窗(共用視窗代碼)
Me.Caption = 共用視窗(共用視窗代碼)
Me.Tag = 共用視窗代碼

現用視窗 = 視窗
現用視窗代碼 = 視窗代碼
'If 現用視窗代碼 > 0 And 現用視窗代碼 < 11 Then
'   mdi漢字字形.mnu_部件代碼(現用視窗代碼).Checked = True
'End If
切換選取字形工具列狀態 現用視窗代碼, 筆畫, 首筆

部件查詢 筆畫, 首筆
啟動部件範例 = True

End Sub

Private Sub Form_Resize()

If Me.ScaleHeight - tree字形樹狀結構.Top * 2 > 0 Then tree字形樹狀結構.Height = Me.ScaleHeight - tree字形樹狀結構.Top * 2
If Me.ScaleWidth - tree字形樹狀結構.Left * 2 > 0 Then tree字形樹狀結構.Width = Me.ScaleWidth - tree字形樹狀結構.Left * 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
'mdi漢字字形.mnu_部件代碼(視窗代碼).Checked = False
mdi漢字字形.txt狀態 = ""
計算現用視窗
啟動部件範例 = False
End Sub


Private Sub tree字形樹狀結構_Click()
Dim 字體 As String
Dim 字形 As String
Dim 編號 As Long

If tree字形樹狀結構.ListIndex > 0 Then
   If tree字形樹狀結構.List(1) <> "" Then
      字體 = tree字形樹狀結構.ItemFontName(tree字形樹狀結構.ListIndex)
      字形 = tree字形樹狀結構.List(tree字形樹狀結構.ListIndex)
      編號 = tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.ListIndex)
      擷取屬性 字體, 字形, 編號
      擷取構字式 字體, 字形, 編號
      If mdi漢字字形.txt字形.Font.Name = "標楷體" Then 拖曳字串 = tree字形樹狀結構.List(tree字形樹狀結構.ListIndex)
      If 啟動字形結構 Then frm字形結構.載入字形 字體, 字形, 編號
      If 啟動異體字表 Then frm異體字表.載入字形 字體, 字形, 編號
      If 啟動字形演變 Then frm字形演變.載入字形 字體, 字形, 編號
      If 啟動字形索引 Then frm字形索引.載入字形 字體, 字形, 編號
      If 啟動異體字根 Then frm異體字根.載入字形 字體, 字形, 編號
      mdi漢字字形.txt狀態 = 部件狀態列
    End If
End If

End Sub

Private Sub tree字形樹狀結構_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

tree字形樹狀結構.OnDragOver X, Y, State

End Sub

Private Sub tree字形樹狀結構_GotFocus()

'tree字形樹狀結構_Click

End Sub

Private Sub tree字形樹狀結構_LostFocus()

tree字形樹狀結構.ListIndex = -1

End Sub

Private Sub tree字形樹狀結構_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

XCheck = X
YCheck = Y

End Sub

Private Sub tree字形樹狀結構_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Initiates dragging only after moving at least 100 twips with the mouse depressed
If (Button And 1) And (XCheck > 0) And (YCheck > 0) And ((Abs(XCheck - X) > 150) Or (Abs(YCheck - Y) > 150)) Then
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
