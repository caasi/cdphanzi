VERSION 5.00
Object = "{65996203-3B87-11D4-A21F-00E029189826}#6.9#0"; "TLIST6.OCX"
Begin VB.Form frm出處檢字 
   Caption         =   "出處檢字"
   ClientHeight    =   5844
   ClientLeft      =   48
   ClientTop       =   324
   ClientWidth     =   4812
   Icon            =   "frm出處檢字.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5844
   ScaleWidth      =   4812
   Begin VB.ComboBox cbo出處 
      BeginProperty Font 
         Name            =   "標楷體"
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
   Begin TListProLibCtl.TList tree字形樹狀結構 
      DragIcon        =   "frm出處檢字.frx":030A
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
      PictureOpen     =   "frm出處檢字.frx":074C
      PictureClosed   =   "frm出處檢字.frx":085E
      PictureLeaf     =   "frm出處檢字.frx":0970
      PictureMark     =   "frm出處檢字.frx":0A82
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
      ExchangeSerialNumber=   "frm出處檢字.frx":0B7C
      DragIconStyle   =   0
      ExchangeDefItemCellDef=   "frm出處檢字.frx":0BC9
      _ChkCounter     =   -1
      TreeLinesHighlightColor=   -2113929196
      TreeLinesShadowColor=   -2113929200
   End
End
Attribute VB_Name = "frm出處檢字"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private 視窗代碼 As Integer, 視窗 As String
Private 區域字體陣列(0 To 字體個數) As Variant
Private 中斷 As Boolean
Private 總數  As Long
Private XCheck As Single, YCheck As Single


Private Sub cbo出處_GotFocus()

現用控制項代碼 = 出處檢字_檢字方塊

End Sub

Public Sub cbo出處_KeyPress(KeyAscii As Integer)

Dim i As Integer

If KeyAscii = vbKeyReturn Then
    Screen.MousePointer = 11
   
    i = 0
    Do While 字體陣列(i) <> ""
        區域字體陣列(i) = 字體陣列(i)
        i = i + 1
    Loop
   
    中斷 = False
    'If cbo出處.Text <> "集成" And cbo出處.Text <> "合集" Then
        載入樹狀 Trim(cbo出處.Text)
    'End If
    Screen.MousePointer = 0
End If

End Sub

Private Sub Form_Activate()

現用視窗 = 視窗
現用視窗代碼 = 出處檢字代碼
'現用視窗代碼 = 視窗代碼
切換選取字形工具列狀態 現用視窗代碼
mdi漢字字形.txt狀態 = 孳乳狀態列

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyEscape Then
   DoEvents
   中斷 = True
End If

End Sub

Private Sub Form_Load()
Dim i As Integer

啟動出處檢字 = True
If 初始first <> 1 Then
   If 已載入畫面 = 0 Then
      If 索引winstate = 0 Then
         frm出處檢字.Left = 出處left
         frm出處檢字.Top = 出處top
         frm出處檢字.Height = 出處height
         frm出處檢字.Width = 出處width
      Else
         frm出處檢字.WindowState = 出處winstate
      End If
   ElseIf 啟動字形孳乳 Then
         frm出處檢字.Left = frm字形孳乳.Left
         frm出處檢字.Top = frm字形孳乳.Top
         frm出處檢字.Height = frm字形孳乳.Height
         frm出處檢字.Width = frm字形孳乳.Width
   End If
End If

tree字形樹狀結構.FontSize = CInt(顯示字型大小)
Me.Tag = 出處檢字代碼
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

cbo出處.List(0) = "合集(甲骨文合集)"
cbo出處.List(1) = "屯(小屯南地甲骨)"
cbo出處.List(2) = "英(英國所藏甲骨集)"
cbo出處.List(3) = "懷(懷特氏等所藏甲骨集)"
cbo出處.List(4) = "集成(殷周金文集成)"
cbo出處.List(5) = "說文"
cbo出處.List(6) = "說文或體"
cbo出處.List(7) = "說文古文"
cbo出處.List(8) = "說文籀文"
cbo出處.List(9) = "說文篆文"
cbo出處.List(10) = "說文俗字"
cbo出處.List(11) = "說文奇字"
cbo出處.List(12) = "天卜(江陵天星觀1號墓卜筮)"
cbo出處.List(13) = "天策(江陵天星觀1號墓遣策)"
cbo出處.List(14) = "包2(荊門包山2號墓)"
cbo出處.List(15) = "仰25(長沙仰天湖25號墓)"
cbo出處.List(16) = "帛甲(長沙子彈庫楚帛書甲篇)"
cbo出處.List(17) = "帛乙(長沙子彈庫楚帛書乙篇)"
cbo出處.List(18) = "雨21(江陵雨臺山21號墓)"
cbo出處.List(19) = "信1(信陽1號墓竹書)"
cbo出處.List(20) = "信2(信陽1號墓遣策)"
cbo出處.List(21) = "范27(江陵范家坡27號墓)"
cbo出處.List(22) = "秦1(江陵秦家嘴1號墓)"
cbo出處.List(23) = "秦13(江陵秦家嘴13號墓)"
cbo出處.List(24) = "秦99(江陵秦家嘴99號墓)"
cbo出處.List(25) = "馬1(江陵馬山1號墓)"
cbo出處.List(26) = "常2(常德市德山夕陽坡2號墓)"
cbo出處.List(27) = "望1(江陵望山1號墓)"
cbo出處.List(28) = "望2(江陵望山2號墓)"
cbo出處.List(29) = "曾(曾候乙墓)"
cbo出處.List(30) = "牌406(長沙五里牌406號墓)"
cbo出處.List(31) = "磚370(江陵磚瓦廠370號墓)"

End Sub

Private Sub Form_Resize()
Dim frm高度 As Integer

frm高度 = Me.ScaleHeight - cbo出處.Height - cbo出處.Top * 3

If frm高度 > 0 Then
   tree字形樹狀結構.Height = frm高度
End If

If (Me.ScaleWidth - tree字形樹狀結構.Left * 2) > 0 Then
   tree字形樹狀結構.Width = Me.ScaleWidth - tree字形樹狀結構.Left * 2
End If

cbo出處.Width = tree字形樹狀結構.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)

mdi漢字字形.mnu_出處檢字.Enabled = True
計算現用視窗
啟動出處檢字 = False

End Sub

Private Sub 載入樹狀(出處 As String)

Dim 字形表 As Recordset, SQL陳述式 As String
Dim i As Integer
Dim 樹根 As String, 字形 As String
Dim 百分比 As Long, 孳乳總數 As Long
Dim 合集編號 As Long, 集成器號 As Long
Dim leftpos As Integer, rightpos As Integer

mdi漢字字形.txt狀態 = "字形載入中  ......       " & " , 欲中斷請按 Esc 鍵"
Screen.MousePointer = ccHourglass

tree字形樹狀結構.Clear

leftpos = InStr(1, 出處, "(")
rightpos = InStr(1, 出處, ")")

If leftpos > 0 Then
    樹根 = Left(出處, leftpos - 1)
    If rightpos < Len(出處) Then 樹根 = 樹根 & Right(出處, Len(出處) - rightpos)
Else
    樹根 = 出處
End If

DoEvents

樹根序數 = 0
孳乳總數 = 0
總數 = 0

出處轉字體 (樹根)

If 出處為甲骨文 Then
    If 出處為甲骨文合集 Then
        If 出處完全匹配 Then
            合集編號 = CStr(CLng(Right(樹根, Len(樹根) - 2)))
            SQL陳述式 = "SELECT * From 異寫字表 Where (出處='" & 合集編號 & "') And (上線>0) And (上線<6) And (上線<>3) ORDER BY 編號"
        Else
            SQL陳述式 = "SELECT * From 異寫字表 Where (合集=1) And (上線>0) And (上線<6) And (上線<>3) ORDER BY 編號"
        End If
    Else
        If 出處完全匹配 Then
            SQL陳述式 = "SELECT * From 異寫字表 Where (出處='" & 樹根 & "') And (上線>0) And (上線<6) And (上線<>3) ORDER BY 編號"
        Else
            SQL陳述式 = "SELECT * From 異寫字表 Where (出處 like '" & 樹根 & "*') And (上線>0) And (上線<6) And (上線<>3) ORDER BY 編號"
        End If
    End If
    Set 字形表 = 甲骨文資料庫.OpenRecordset(SQL陳述式)
ElseIf 出處為金文 Then
    If 出處完全匹配 Then
        集成器號 = CLng(Right(樹根, Len(樹根) - 2))
        SQL陳述式 = "SELECT * From 異寫字表 Where (器號=" & 集成器號 & ") And (上線>0) And (上線<6) And (上線<>3) ORDER BY 編號"
    Else
        SQL陳述式 = "SELECT * From 異寫字表 Where (上線>0) And (上線<6) And (上線<>3) ORDER BY 編號"
    End If
    Set 字形表 = 金文資料庫.OpenRecordset(SQL陳述式)
ElseIf 出處為小篆 Then
    SQL陳述式 = "SELECT * From 檢字表 Where 字源='" & 樹根 & "' ORDER BY 編號"
    Set 字形表 = 小篆資料庫.OpenRecordset(SQL陳述式)
Else
    If InStr(1, 出處, ".") > 0 Then
        SQL陳述式 = "SELECT * From 異寫字表 Where (出處='" & 樹根 & "') And (上線>0) And (上線<6) And (上線<>3) ORDER BY 編號"
    Else
        SQL陳述式 = "SELECT * From 異寫字表 Where (出處 like'" & 樹根 & "*') And (上線>0) and (上線<6) And (上線<>3) ORDER BY 編號"
    End If
    Set 字形表 = 楚系文字資料庫.OpenRecordset(SQL陳述式)
End If
   
If Not 字形表.EOF Then
   
    字形表.MoveLast
    總數 = 字形表.RecordCount
            
    tree字形樹狀結構.AddItem 樹根
      
    tree字形樹狀結構.ItemLngValue(0) = -999999
    tree字形樹狀結構.ItemTag(0) = 其他節點標記
    tree字形樹狀結構.ItemFontName(0) = 顯示字型
    tree字形樹狀結構.Expand(0) = True
    
    字形表.MoveFirst

    Do Until 字形表.EOF
        If 中斷 = True Then Exit Do
         
        樹根序數 = 樹根序數 + 1
        百分比 = 樹根序數 / 總數 * 100
        mdi漢字字形.txt狀態 = "字形載入已完成 " & 百分比 & " % , 欲中斷請按 Esc 鍵"
         
        字形 = 字形表.Fields("字碼")
        tree字形樹狀結構.AddItem 字形, 0
        tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
        tree字形樹狀結構.ItemFontName(tree字形樹狀結構.NewIndex) = 區域字體陣列(字形表.Fields("字體"))
        tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = 字形表.Fields("編號")
                  'If Not IsNull(字形表.Fields("字形")) Then
                   '  tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 字形節點標記
                  'Else
                   '  tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 構字式節點標記
                  'End If
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
      
      總數 = 0

End If


tree字形樹狀結構.Redraw = True
If 中斷 = True Then
    孳乳狀態列 = "使用者中斷！ 已載入 " & 樹根序數 & " 個字形"
Else
    孳乳狀態列 = "載入完畢！ 共 " & 總數 & " 個字形"
End If
mdi漢字字形.txt狀態 = 孳乳狀態列

Screen.MousePointer = ccDefault

End Sub

Private Sub tree字形樹狀結構_Click()

Dim 字體 As String
Dim 字形 As String
Dim 編號 As Long

If tree字形樹狀結構.ListIndex <> -1 Then
   If tree字形樹狀結構.List(0) <> "" Then
      字體 = tree字形樹狀結構.ItemFontName(tree字形樹狀結構.ListIndex)
      字形 = tree字形樹狀結構.List(tree字形樹狀結構.ListIndex)
      If Len(字形) = 1 Then
      編號 = tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.ListIndex)
      
      擷取屬性 字體, 字形, 編號
      擷取構字式 字體, 字形, 編號
      If mdi漢字字形.txt字形.font.Name = "標楷體" Then 拖曳字串 = tree字形樹狀結構.List(tree字形樹狀結構.ListIndex)
      
      If 啟動字形結構 Then frm字形結構.載入字形 字體, 字形, 編號
      If 啟動異體字表 Then frm異體字表.載入字形 字體, 字形, 編號
      If 啟動字形演變 Then frm字形演變.載入字形 字體, 字形, 編號
      If 啟動字形索引 Then frm字形索引.載入字形 字體, 字形, 編號
      If 啟動異體字根 Then frm異體字根.載入字形 字體, 字形, 編號
      mdi漢字字形.txt狀態 = 孳乳狀態列
      End If
   End If
End If

End Sub

Private Sub tree字形樹狀結構_GotFocus()

現用控制項代碼 = 出處檢字_樹狀結構

End Sub


