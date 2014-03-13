VERSION 5.00
Object = "{65996203-3B87-11D4-A21F-00E029189826}#6.9#0"; "TLIST6.OCX"
Begin VB.Form frm字形演變 
   Caption         =   "字形演變"
   ClientHeight    =   6228
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9060
   Icon            =   "frm字形演變.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6228
   ScaleWidth      =   9060
   Begin TListProLibCtl.TList tree字形樹狀結構 
      DragIcon        =   "frm字形演變.frx":030A
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
      PictureOpen     =   "frm字形演變.frx":074C
      PictureClosed   =   "frm字形演變.frx":085E
      PictureLeaf     =   "frm字形演變.frx":0970
      PictureMark     =   "frm字形演變.frx":0A82
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
      ExchangeSerialNumber=   "frm字形演變.frx":0B7C
      DragIconStyle   =   0
      ExchangeDefItemCellDef=   "frm字形演變.frx":0BC9
      _ChkCounter     =   -1
      TreeLinesHighlightColor=   -2113929196
      TreeLinesShadowColor=   -2113929200
   End
   Begin VB.Image imgPinPush 
      Height          =   264
      Left            =   4440
      Picture         =   "frm字形演變.frx":0CD0
      Top             =   240
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imgPin 
      Height          =   264
      Left            =   4440
      Picture         =   "frm字形演變.frx":0E5A
      Top             =   720
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imglock 
      Height          =   264
      Left            =   0
      Picture         =   "frm字形演變.frx":0FE4
      Tag             =   "0"
      ToolTipText     =   "鎖定"
      Top             =   240
      Width           =   288
   End
End
Attribute VB_Name = "frm字形演變"
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
現用視窗代碼 = 字形演變代碼
切換選取字形工具列狀態 現用視窗代碼
tree字形樹狀結構_Click
mdi漢字字形.txt狀態 = 結構狀態列

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim 字根序 As String, 編號 As Long

啟動字形演變 = True
If 初始first <> 1 Then
   If 已載入畫面 = 0 Then
      If 演變winstate = 0 Then
         frm字形演變.Left = 演變left
         frm字形演變.Top = 演變top
         frm字形演變.Height = 演變height
         frm字形演變.Width = 演變width
      Else
         frm字形演變.WindowState = 演變winstate
      End If
   ElseIf 啟動字形孳乳 Then
         frm字形演變.Left = frm字形孳乳.Left + frm字形孳乳.Width
         frm字形演變.Top = frm字形孳乳.Top
         frm字形演變.Height = frm字形孳乳.Height
         frm字形演變.Width = frm字形孳乳.Width
   End If
End If

tree字形樹狀結構.FontSize = CInt(顯示字型大小)
'Me.Tag = 共用視窗代碼
Me.Tag = 字形演變代碼
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
mdi漢字字形.mnu_字形演變.Enabled = True
啟動字形演變 = False
計算現用視窗

End Sub

Public Sub 載入字形(系統字體 As String, 字形 As String, 編號 As Long)
Dim 部件序 As String, 字體編號 As Integer, 字型檔 As String, 分解 As Integer
Dim i As Integer, 字體 As String
Dim 楷書編號 As Long, 小篆編號 As Integer, 金文編號 As Long, 甲骨文編號 As Integer, 楚系文字編號 As Long
Dim 楷書字形 As String, 小篆字形 As String, 金文字形 As String, 甲骨文字形 As String, 楚系文字字形 As String
Dim 小篆字源 As String, 金文字源 As String, 甲骨文字源 As String, 楚系文字字源 As String
Dim 甲骨文索引 As Long, 金文索引 As Long, 楚系文字索引 As Long, 小篆索引 As Long
Dim 集成器號 As String, 器名缺字 As Boolean, RTF缺字 As String
Dim 金文字型 As String, 甲骨文字型 As String, 楚系文字字型 As String

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
    If 系統字體 <> "中研院金文" Then
        金文異寫字表.Seek "=", 編號
        楷書編號 = 金文異寫字表.Fields("楷書編號")
    Else
        金文檢字表.Index = "編號"
        金文檢字表.Seek "=", 編號
        楷書編號 = 金文檢字表.Fields("楷書編號")
    End If
ElseIf 中研院甲骨文(系統字體) Then
    甲骨文編號 = 編號
    If 系統字體 <> "中研院甲骨文" Then
        甲骨文異寫字表.Seek "=", 編號
        楷書編號 = 甲骨文異寫字表.Fields("楷書編號")
    Else
        甲骨文檢字表.Index = "編號"
        甲骨文檢字表.Seek "=", 編號
        楷書編號 = 甲骨文檢字表.Fields("楷書編號")
    End If
ElseIf 中研院楚系文字(系統字體) Then
    楚系文字編號 = 編號
    If 系統字體 <> "中研院楚系簡帛文字" Then
        楚系文字異寫字表.Seek "=", 編號
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

If 系統字體 <> "北師大說文小篆" And 系統字體 <> "北師大說文重文" Then
    If Not IsNull(楷書檢字表.Fields("小篆編號")) Then
        小篆編號 = 楷書檢字表.Fields("小篆編號")
    Else
        小篆編號 = 0
    End If
End If

If Not 中研院金文(系統字體) Then
    If Not IsNull(楷書檢字表.Fields("金文編號")) Then
        金文編號 = 楷書檢字表.Fields("金文編號")
    Else
        金文編號 = 0
    End If
End If

If Not 中研院甲骨文(系統字體) Then
    If Not IsNull(楷書檢字表.Fields("甲骨文編號")) Then
        甲骨文編號 = 楷書檢字表.Fields("甲骨文編號")
    Else
        甲骨文編號 = 0
    End If
End If

If Not 中研院楚系文字(系統字體) Then
    If Not IsNull(楷書檢字表.Fields("楚系文字編號")) Then
        楚系文字編號 = 楷書檢字表.Fields("楚系文字編號")
    Else
        楚系文字編號 = 0
    End If
End If

甲骨文索引 = -1
金文索引 = -1
楚系文字索引 = -1
小篆索引 = -1

tree字形樹狀結構.Clear
tree字形樹狀結構.Redraw = False
tree字形樹狀結構.AddItem 楷書字形
tree字形樹狀結構.ItemFontName(0) = 轉換顯示字型(字體)
tree字形樹狀結構.ItemLngValue(0) = 楷書編號
If Not IsNull(檢字表.Fields("字形")) Then
    tree字形樹狀結構.ItemTag(0) = 字形節點標記
Else
    tree字形樹狀結構.ItemTag(0) = 構字式節點標記
End If


If (小篆編號 = 0 Or Not mdi漢字字形.mnu_小篆選項.Checked) And (金文編號 = 0 Or Not mdi漢字字形.mnu_金文選項.Checked) And (甲骨文編號 = 0 Or Not mdi漢字字形.mnu_甲骨文選項.Checked) And (楚系文字編號 = 0 Or Not mdi漢字字形.mnu_楚系文字選項.Checked) Then
    tree字形樹狀結構.Image(0) = tree字形樹狀結構.PictureLeaf
    tree字形樹狀結構.Redraw = True
    Screen.MousePointer = ccDefault
    Exit Sub
End If

tree字形樹狀結構.Image(0) = tree字形樹狀結構.PictureOpen

If 甲骨文編號 <> 0 And mdi漢字字形.mnu_甲骨文選項.Checked Then
    甲骨文檢字表.Index = "編號"
    甲骨文檢字表.Seek "=", 甲骨文編號
    
    If Not 甲骨文檢字表.NoMatch Then
        甲骨文字形 = 甲骨文檢字表.Fields("字碼")
        甲骨文字型 = "中研院甲骨文"
        甲骨文異寫字表.Seek "=", 甲骨文編號
    Else
        甲骨文異寫字表.Seek "=", 甲骨文編號
        甲骨文字形 = 甲骨文異寫字表.Fields("字碼")
        甲骨文字型 = 系統字體
    End If

    If Not IsNull(甲骨文異寫字表.Fields("出處")) Then
        甲骨文字源 = 甲骨文異寫字表.Fields("出處")
        If IsNumeric(Left(甲骨文字源, 1)) Then 甲骨文字源 = "合集" & 甲骨文字源
    Else
        甲骨文字源 = ""
    End If
    
    甲骨文索引 = tree字形樹狀結構.ListCount
    tree字形樹狀結構.AddItem 甲骨文字形, 0
    tree字形樹狀結構.ItemFontName(甲骨文索引) = 甲骨文字型
    tree字形樹狀結構.ItemLngValue(甲骨文索引) = 甲骨文編號
    If Not IsNull(甲骨文檢字表.Fields("字形")) Then
        tree字形樹狀結構.ItemTag(甲骨文索引) = 字形節點標記
    Else
        tree字形樹狀結構.ItemTag(甲骨文索引) = 構字式節點標記
    End If

    If Len(甲骨文字源) = 0 Then
        tree字形樹狀結構.Image(甲骨文索引) = tree字形樹狀結構.PictureLeaf
    Else
        tree字形樹狀結構.Image(甲骨文索引) = tree字形樹狀結構.PictureOpen
        tree字形樹狀結構.AddItem 甲骨文字源, 甲骨文索引
        tree字形樹狀結構.ItemFontName(甲骨文索引 + 1) = "標楷體"
        tree字形樹狀結構.ItemTag(甲骨文索引 + 1) = 其他節點標記
        tree字形樹狀結構.ItemFontSize(甲骨文索引 + 1) = 12
        tree字形樹狀結構.ItemLngValue(甲骨文索引 + 1) = -9999
        tree字形樹狀結構.Image(甲骨文索引 + 1) = tree字形樹狀結構.PictureLeaf
    End If
End If

If 金文編號 <> 0 And mdi漢字字形.mnu_金文選項.Checked Then
    金文檢字表.Index = "編號"
    金文檢字表.Seek "=", 金文編號
    If Not 金文檢字表.NoMatch Then
        金文字形 = 金文檢字表.Fields("字碼")
        金文字型 = "中研院金文"
        金文異寫字表.Seek "=", 金文編號
        集成器號 = 金文異寫字表.Fields("器號")
    Else
        金文異寫字表.Seek "=", 金文編號
        金文字形 = 金文異寫字表.Fields("字碼")
        金文字型 = 系統字體
        集成器號 = 金文異寫字表.Fields("器號")
    End If
    
    金文集成器名.Seek "=", 集成器號
    If 集成器號 <> 15000 Then 金文字源 = "集成" & 集成器號
    If Not 金文集成器名.NoMatch Then
        金文字源 = 金文字源 & "(" & 金文集成器名.Fields("器名") & ")"
        If 金文集成器名.Fields("缺字") = 1 Then
            器名缺字 = True
        Else
            器名缺字 = False
        End If
    Else
        金文字源 = ""
    End If
    
    金文索引 = tree字形樹狀結構.ListCount
    tree字形樹狀結構.AddItem 金文字形, 0
    tree字形樹狀結構.ItemFontName(金文索引) = 金文字型
    tree字形樹狀結構.ItemLngValue(金文索引) = 金文編號
    If Not IsNull(金文檢字表.Fields("字形")) Then
        tree字形樹狀結構.ItemTag(金文索引) = 字形節點標記
    Else
        tree字形樹狀結構.ItemTag(金文索引) = 構字式節點標記
    End If

    If Len(金文字源) = 0 Then
        tree字形樹狀結構.Image(金文索引) = tree字形樹狀結構.PictureLeaf
    Else
        tree字形樹狀結構.Image(金文索引) = tree字形樹狀結構.PictureOpen
        If Not 器名缺字 Then
            tree字形樹狀結構.AddItem 金文字源, 金文索引
            tree字形樹狀結構.ItemFontName(金文索引 + 1) = "標楷體"
            tree字形樹狀結構.ItemTag(金文索引 + 1) = 其他節點標記
        Else
            RTF缺字 = 轉換RTF缺字(金文字源, 顯示字型)
            tree字形樹狀結構.AddItem RTF缺字, 金文索引
            tree字形樹狀結構.ItemCell(金文索引 + 1).RTFStyle = 1
            tree字形樹狀結構.ItemTag(金文索引 + 1) = 器名節點標記 & 金文字源
        End If
        tree字形樹狀結構.ItemFontSize(金文索引 + 1) = 12
        tree字形樹狀結構.ItemLngValue(金文索引 + 1) = -9999
        tree字形樹狀結構.Image(金文索引 + 1) = tree字形樹狀結構.PictureLeaf
    End If
End If

If 楚系文字編號 <> 0 And mdi漢字字形.mnu_楚系文字選項.Checked Then
    楚系文字檢字表.Index = "編號"
    楚系文字檢字表.Seek "=", 楚系文字編號
    
    If Not 楚系文字檢字表.NoMatch Then
        楚系文字字形 = 楚系文字檢字表.Fields("字碼")
        楚系文字字型 = "中研院楚系簡帛文字"
        楚系文字異寫字表.Seek "=", 楚系文字編號
    Else
        楚系文字異寫字表.Seek "=", 楚系文字編號
        楚系文字字形 = 楚系文字異寫字表.Fields("字碼")
        楚系文字字型 = 系統字體
    End If
    
    If Not IsNull(楚系文字異寫字表.Fields("出處")) Then
        楚系文字字源 = 楚系文字異寫字表.Fields("出處")
    Else
        楚系文字字源 = ""
    End If
    
    楚系文字索引 = tree字形樹狀結構.ListCount
    tree字形樹狀結構.AddItem 楚系文字字形, 0
    tree字形樹狀結構.ItemFontName(楚系文字索引) = 楚系文字字型
    tree字形樹狀結構.ItemLngValue(楚系文字索引) = 楚系文字編號
    If Not IsNull(楚系文字檢字表.Fields("字形")) Then
        tree字形樹狀結構.ItemTag(楚系文字索引) = 字形節點標記
    Else
        tree字形樹狀結構.ItemTag(楚系文字索引) = 構字式節點標記
    End If

    If Len(楚系文字字源) = 0 Then
        tree字形樹狀結構.Image(楚系文字索引) = tree字形樹狀結構.PictureLeaf
    Else
        tree字形樹狀結構.Image(楚系文字索引) = tree字形樹狀結構.PictureOpen
        tree字形樹狀結構.AddItem 楚系文字字源, 楚系文字索引
        tree字形樹狀結構.ItemFontName(楚系文字索引 + 1) = "標楷體"
        tree字形樹狀結構.ItemTag(楚系文字索引 + 1) = 其他節點標記
        tree字形樹狀結構.ItemFontSize(楚系文字索引 + 1) = 12
        tree字形樹狀結構.ItemLngValue(楚系文字索引 + 1) = -9999
        tree字形樹狀結構.Image(楚系文字索引 + 1) = tree字形樹狀結構.PictureLeaf
    End If
End If

If 小篆編號 <> 0 And mdi漢字字形.mnu_小篆選項.Checked Then
    小篆檢字表.Index = "編號"
    小篆檢字表.Seek "=", 小篆編號
    小篆字形 = 小篆檢字表.Fields("字碼")
    If Not IsNull(小篆檢字表.Fields("字源")) Then
        小篆字源 = 小篆檢字表.Fields("字源")
    Else
        小篆字源 = ""
    End If

    小篆索引 = tree字形樹狀結構.ListCount
    
    tree字形樹狀結構.AddItem 小篆字形, 0
    If 系統字體 = "北師大說文重文" Then
        tree字形樹狀結構.ItemFontName(小篆索引) = "北師大說文重文"
    Else
        tree字形樹狀結構.ItemFontName(小篆索引) = "北師大說文小篆"
    End If
    tree字形樹狀結構.ItemLngValue(小篆索引) = 小篆編號

    If Not IsNull(小篆檢字表.Fields("字形")) Then
        tree字形樹狀結構.ItemTag(小篆索引) = 字形節點標記
    Else
        tree字形樹狀結構.ItemTag(小篆索引) = 構字式節點標記
    End If

    If Len(小篆字源) = 0 Then
        tree字形樹狀結構.Image(小篆索引) = tree字形樹狀結構.PictureLeaf
    Else
        tree字形樹狀結構.Image(小篆索引) = tree字形樹狀結構.PictureOpen
        tree字形樹狀結構.AddItem 小篆字源, 小篆索引
        tree字形樹狀結構.ItemFontName(小篆索引 + 1) = 轉換顯示字型("標楷體")
        tree字形樹狀結構.ItemFontSize(小篆索引 + 1) = 12
        tree字形樹狀結構.ItemLngValue(小篆索引 + 1) = -9999
        tree字形樹狀結構.ItemTag(小篆索引 + 1) = 其他節點標記
        tree字形樹狀結構.Image(小篆索引 + 1) = tree字形樹狀結構.PictureLeaf
    End If
End If

tree字形樹狀結構.Expand(0) = True
If 甲骨文索引 > -1 Then tree字形樹狀結構.Expand(甲骨文索引) = True
If 金文索引 > -1 Then tree字形樹狀結構.Expand(金文索引) = True
If 楚系文字索引 > -1 Then tree字形樹狀結構.Expand(楚系文字索引) = True
If 小篆索引 > -1 Then tree字形樹狀結構.Expand(小篆索引) = True


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
    frm字形演變.Caption = "字形演變(鎖定)"
Else
    imglock.Tag = 0
    imglock.Picture = imgPin.Picture
    imglock.ToolTipText = "鎖定"
    frm字形演變.Caption = "字形演變"
End If

End Sub

Private Sub tree字形樹狀結構_Click()
Dim 字體 As String
Dim 字形 As String
Dim 編號 As Long

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
      If 啟動字形索引 Then frm字形索引.載入字形 字體, 字形, 編號
      If 啟動異體字根 Then frm異體字根.載入字形 字體, 字形, 編號
    End If
End If

End Sub

Private Sub tree字形樹狀結構_GotFocus()

現用視窗代碼 = 字形演變代碼

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
