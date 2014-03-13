VERSION 5.00
Object = "{65996203-3B87-11D4-A21F-00E029189826}#6.9#0"; "TLIST6.OCX"
Begin VB.Form frm異體字表 
   Caption         =   "異體字表"
   ClientHeight    =   4956
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6804
   Icon            =   "frm異體字表.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4956
   ScaleWidth      =   6804
   Begin TListProLibCtl.TList tree字形樹狀結構 
      DragIcon        =   "frm異體字表.frx":030A
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
      PictureOpen     =   "frm異體字表.frx":074C
      PictureClosed   =   "frm異體字表.frx":085E
      PictureLeaf     =   "frm異體字表.frx":0970
      PictureMark     =   "frm異體字表.frx":0A82
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
      PicturePalette  =   "frm異體字表.frx":0B94
      ExchangeSerialNumber=   "frm異體字表.frx":0CA6
      DragIconStyle   =   0
      ExchangeDefItemCellDef=   "frm異體字表.frx":0CF3
      _ChkCounter     =   -1
      TreeLinesHighlightColor=   -2113929196
      TreeLinesShadowColor=   -2113929200
   End
   Begin VB.Image img楚系文字 
      Height          =   240
      Left            =   6168
      Picture         =   "frm異體字表.frx":0DFA
      Top             =   1872
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Image img奇字 
      Height          =   240
      Left            =   5052
      Picture         =   "frm異體字表.frx":150C
      Top             =   3528
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Image img俗字 
      Height          =   240
      Left            =   5604
      Picture         =   "frm異體字表.frx":1C1E
      Top             =   3144
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Image img籀文 
      Height          =   240
      Left            =   5052
      Picture         =   "frm異體字表.frx":2330
      Top             =   3144
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Image img或體 
      Height          =   240
      Left            =   5064
      Picture         =   "frm異體字表.frx":2A42
      Top             =   2724
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Image img古文 
      Height          =   240
      Left            =   5604
      Picture         =   "frm異體字表.frx":3154
      Top             =   2736
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Image img小篆 
      Height          =   240
      Left            =   5604
      Picture         =   "frm異體字表.frx":3866
      Top             =   2352
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Image img甲骨文 
      Height          =   240
      Left            =   5604
      Picture         =   "frm異體字表.frx":3F78
      Top             =   1872
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Image img訛字 
      Height          =   240
      Left            =   5064
      Picture         =   "frm異體字表.frx":468A
      Top             =   2340
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Image img金文 
      Height          =   240
      Left            =   5052
      Picture         =   "frm異體字表.frx":4D9C
      Top             =   1860
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Image imgNull 
      Height          =   384
      Left            =   4452
      Picture         =   "frm異體字表.frx":54AE
      Top             =   2316
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image img簡化字 
      Height          =   240
      Left            =   4452
      Picture         =   "frm異體字表.frx":5CF0
      Top             =   1848
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.Image imgPinPush 
      Height          =   264
      Left            =   4440
      Picture         =   "frm異體字表.frx":6402
      Top             =   840
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imgPin 
      Height          =   264
      Left            =   4440
      Picture         =   "frm異體字表.frx":658C
      Top             =   1320
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imglock 
      Height          =   264
      Left            =   0
      Picture         =   "frm異體字表.frx":6716
      Tag             =   "0"
      ToolTipText     =   "鎖定"
      Top             =   240
      Width           =   288
   End
   Begin VB.Image Image1 
      Height          =   192
      Left            =   4440
      Picture         =   "frm異體字表.frx":68A0
      Top             =   480
      Visible         =   0   'False
      Width           =   192
   End
End
Attribute VB_Name = "frm異體字表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private 視窗代碼 As Integer, 視窗 As String, 字根表 As Recordset, 檢字表 As Recordset
Private 區域字體陣列(0 To 字體個數) As Variant
Private 異體字表 As Recordset, 異寫字表 As Recordset
Private 中斷 As Boolean
Private 總數  As Long, 部數 As String
Private XCheck As Single, YCheck As Single
Private 標記個數 As Integer
Private Const 簡化字代碼 = 2
Private Const 小篆代碼 = 3, 或體代碼 = 4, 古文代碼 = 5, 籀文代碼 = 6, 俗字代碼 = 7, 奇字代碼 = 8
Private Const 金文代碼 = 11, 甲骨文代碼 = 12, 楚系文字代碼 = 13
Private Const 訛字代碼 = 99


Private Sub Form_Activate()
現用視窗 = 視窗
'現用視窗代碼 = 視窗代碼
現用視窗代碼 = 異體字表代碼
切換選取字形工具列狀態 現用視窗代碼
mdi漢字字形.txt狀態 = 異體狀態列

End Sub


Private Sub Form_Load()
Dim i As Integer
Dim 編號 As Long


啟動異體字表 = True

If 初始first <> 1 Then
   If 已載入畫面 = 0 Then
      If 異體winstate = 0 Then
         frm異體字表.Left = 異體left
         frm異體字表.Top = 異體top
         frm異體字表.Height = 異體height
         frm異體字表.Width = 異體width
      Else
         frm異體字表.WindowState = 異體winstate
      End If
   ElseIf 啟動字形孳乳 Then
         frm異體字表.Left = frm字形孳乳.Left + frm字形孳乳.Width
         frm異體字表.Top = frm字形孳乳.Top
         frm異體字表.Height = frm字形孳乳.Height
         frm異體字表.Width = frm字形孳乳.Width
   End If
End If

tree字形樹狀結構.MarkPicture(0) = imgNull

tree字形樹狀結構.MarkPicture(簡化字代碼) = img簡化字
tree字形樹狀結構.MarkPicture(金文代碼) = img金文
tree字形樹狀結構.MarkPicture(甲骨文代碼) = img甲骨文
tree字形樹狀結構.MarkPicture(楚系文字代碼) = img楚系文字
tree字形樹狀結構.MarkPicture(小篆代碼) = img小篆
tree字形樹狀結構.MarkPicture(或體代碼) = img或體
tree字形樹狀結構.MarkPicture(古文代碼) = img古文
tree字形樹狀結構.MarkPicture(籀文代碼) = img籀文
tree字形樹狀結構.MarkPicture(俗字代碼) = img俗字
tree字形樹狀結構.MarkPicture(奇字代碼) = img奇字
tree字形樹狀結構.MarkPicture(訛字代碼) = img訛字

'If 系統字體 = "楷書" Then
'    Set 字根表 = 楷書字根
'ElseIf 系統字體 = "小篆" Then
'    Set 字根表 = 小篆獨體字
'End If
'字根表.Index = "字形"

i = 0
Do While 字體陣列(i) <> ""
   區域字體陣列(i) = 字體陣列(i)
   i = i + 1
Loop

tree字形樹狀結構.FontSize = CInt(顯示字型大小)
視窗代碼 = 共用視窗代碼
視窗 = 共用視窗(共用視窗代碼)
'Me.Tag = 共用視窗代碼
Me.Tag = 異體字表代碼
tree字形樹狀結構.AddItem ""
'tree字形樹狀結構.ListIndex = 0
tree字形樹狀結構.Image(0) = tree字形樹狀結構.PictureLeaf

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
'字根表.Close
啟動異體字表 = False
mdi漢字字形.mnu_異體字表.Enabled = True
計算現用視窗
End Sub

Private Sub imglock_Click()

If imglock.Tag = 0 Then
    imglock.Tag = 1
    imglock.Picture = imgPinPush.Picture
    imglock.ToolTipText = "解除鎖定"
    frm異體字表.Caption = "異體字表(鎖定)"
Else
    imglock.Tag = 0
    imglock.Picture = imgPin.Picture
    imglock.ToolTipText = "鎖定"
    frm異體字表.Caption = "異體字表"
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
      If 啟動字形演變 Then frm字形演變.載入字形 字體, 字形, 編號
      If 啟動字形索引 Then frm字形索引.載入字形 字體, 字形, 編號
      If 啟動異體字根 Then frm異體字根.載入字形 字體, 字形, 編號
      mdi漢字字形.txt狀態 = 異體狀態列
   End If
End If

End Sub

Private Sub tree字形樹狀結構_Expand(ByVal i As Long)

If tree字形樹狀結構.ListCountEx(i) = 1 Then
    If tree字形樹狀結構.List(i + 1) = "" Then
        Screen.MousePointer = 11
        tree字形樹狀結構.RemoveItem i + 1
        tree字形樹狀結構.Redraw = False
        載入異寫二 i
        tree字形樹狀結構.Redraw = True
        Screen.MousePointer = 0
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

Public Sub 載入字形(系統字體 As String, val字形 As String, 系統編號 As Long)
Dim 字頭表 As Recordset, 異寫型態 As Integer
Dim SQL陳述式 As String, 樹根 As String
Dim i As Long, 總數 As Long, 字形 As String, 形碼 As Integer
Dim 字體編號 As Integer, 字體 As String, 暫存字體 As String
Dim 楷書編號 As Long, 編號 As Long
Dim 顯示字型 As String

If imglock.Tag = 1 Then Exit Sub
編號 = 系統編號
If 編號 <= 0 Then Exit Sub

mdi漢字字形.txt狀態 = "字形載入中  ......       " & " , 欲中斷請按 Esc 鍵"
Screen.MousePointer = ccHourglass

字形 = val字形
tree字形樹狀結構.Clear


If 系統字體 = "北師大說文小篆" Or 系統字體 = "北師大說文重文" Then
    Set 檢字表 = 小篆檢字表
ElseIf 中研院金文(系統字體) Then
    Set 檢字表 = 金文檢字表
ElseIf 中研院甲骨文(系統字體) Then
    Set 檢字表 = 甲骨文檢字表
ElseIf 中研院楚系文字(系統字體) Then
    Set 檢字表 = 楚系文字檢字表
Else
    Set 檢字表 = 楷書檢字表
End If

檢字表.Index = "編號"
檢字表.Seek "=", 編號
If Not 檢字表.NoMatch Then
    字體編號 = 檢字表.Fields("字體")
    字形 = 檢字表.Fields("字碼")
ElseIf 中研院金文(系統字體) Then
    金文異寫字表.Index = "編號"
    金文異寫字表.Seek "=", 編號
    字體編號 = 金文異寫字表.Fields("字體")
    字形 = 金文異寫字表.Fields("字碼")
    編號 = 金文異寫字表.Fields("異寫編號")
ElseIf 中研院甲骨文(系統字體) Then
    甲骨文異寫字表.Index = "編號"
    甲骨文異寫字表.Seek "=", 編號
    字體編號 = 甲骨文異寫字表.Fields("字體")
    字形 = 甲骨文異寫字表.Fields("字碼")
    編號 = 甲骨文異寫字表.Fields("異寫編號")
ElseIf 中研院楚系文字(系統字體) Then
    楚系文字異寫字表.Index = "編號"
    楚系文字異寫字表.Seek "=", 編號
    字體編號 = 楚系文字異寫字表.Fields("字體")
    字形 = 楚系文字異寫字表.Fields("字碼")
    編號 = 楚系文字異寫字表.Fields("異寫編號")
End If

字體 = 區域字體陣列(字體編號)
'If 字體編號 = 0 Then 字體 = "標楷體"
'If 字體編號 = 13 Then 字體 = "北師大說文小篆"
'If 字體編號 = 14 Then 字體 = "北師大說文重文"
'If 字體編號 = 15 Then 字體 = "中研院金文"
'If 字體編號 = 16 Then 字體 = "中研院甲骨文"
'If 字體編號 = 17 Then 字體 = "中研院楚系簡帛文字"
If 系統字體 = "北師大說文小篆" Or 系統字體 = "北師大說文重文" Or 系統字體 = "中研院金文" Or 系統字體 = "中研院甲骨文" Or 系統字體 = "中研院楚系簡帛文字" Then 字形 = 檢字表.Fields("字碼")
樹根 = 字形

If Not 中研院金文(系統字體) And Not 中研院楚系文字(系統字體) Then
    SQL陳述式 = "SELECT 組號 FROM 異體字表 WHERE (編號 = " & 編號 & ")  GROUP BY 組號 ORDER BY 組號;"
Else
    SQL陳述式 = "SELECT 組號 FROM 異體字表 WHERE (編號 = " & 編號 & ") OR (重見 = " & 編號 & ") GROUP BY 組號 ORDER BY 組號;"
End If

DoEvents

標記個數 = 0
樹根序數 = 0
總數 = 0

tree字形樹狀結構.Redraw = False

If 系統字體 = "北師大說文小篆" Or 系統字體 = "北師大說文重文" Then
    Set 字頭表 = 小篆資料庫.OpenRecordset(SQL陳述式)
ElseIf 中研院金文(系統字體) Then
    Set 字頭表 = 金文資料庫.OpenRecordset(SQL陳述式)
ElseIf 中研院甲骨文(系統字體) Then
    Set 字頭表 = 甲骨文資料庫.OpenRecordset(SQL陳述式)
ElseIf 中研院楚系文字(系統字體) Then
    Set 字頭表 = 楚系文字資料庫.OpenRecordset(SQL陳述式)
Else
    Set 字頭表 = 系統資料庫.OpenRecordset(SQL陳述式)
End If

If tree字形樹狀結構.ListCount > 0 Then
   tree字形樹狀結構.RemoveItem (0)
End If

If Not 字頭表.EOF Then
   
   字頭表.MoveLast
   總數 = 字頭表.RecordCount
 
   'If 總數 > 1 Or (總數 = 1 And 字頭表.Fields(1) <> 1) Or 系統字體 = "中研院金文" Then
   If 總數 > 1 Then
      tree字形樹狀結構.AddItem 樹根
      tree字形樹狀結構.ItemFontName(0) = 轉換顯示字型(字體)
      tree字形樹狀結構.Image(0) = tree字形樹狀結構.PictureOpen
      
      tree字形樹狀結構.ItemLngValue(0) = 編號
      If Not IsNull(檢字表.Fields("字形")) Then
        tree字形樹狀結構.ItemTag(0) = 字形節點標記
      Else
        tree字形樹狀結構.ItemTag(0) = 構字式節點標記
      End If
      tree字形樹狀結構.Expand(0) = True
   End If
    
   字頭表.MoveFirst
    
   Do Until 字頭表.EOF
      If 中斷 = True Then Exit Do
         
      SQL陳述式 = "SELECT * From 異體字表 Where 組號= " & 字頭表.Fields("組號") & " ORDER BY 形碼,編號"
      If 系統字體 = "北師大說文小篆" Or 系統字體 = "北師大說文重文" Then
         Set 異體字表 = 小篆資料庫.OpenRecordset(SQL陳述式)
      ElseIf 中研院金文(系統字體) Then
         Set 異體字表 = 金文資料庫.OpenRecordset(SQL陳述式)
      ElseIf 中研院甲骨文(系統字體) Then
         Set 異體字表 = 甲骨文資料庫.OpenRecordset(SQL陳述式)
      ElseIf 中研院楚系文字(系統字體) Then
         Set 異體字表 = 楚系文字資料庫.OpenRecordset(SQL陳述式)
      Else
         Set 異體字表 = 系統資料庫.OpenRecordset(SQL陳述式)
      End If
      
      異體字表.MoveLast
      異體字表.MoveFirst
      檢字表.Index = "編號"

      Do Until 異體字表.EOF
         樹根序數 = 樹根序數 + 1

         檢字表.Seek "=", 異體字表.Fields("編號")
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
   
          If 異體字表.Fields("形碼") > 1 Then
             If 中研院金文(系統字體) Or 中研院甲骨文(系統字體) Or 中研院楚系文字(系統字體) Then
                異寫型態 = 0
                If Not IsNull(異體字表.Fields("型態")) Then 異寫型態 = 異體字表.Fields("型態")
                If 異寫型態 = 1 Then
                    載入異寫字一 系統字體, 異體字表.Fields("編號"), i
                    tree字形樹狀結構.Expand(i) = True
                    字頭表.MoveNext
                    GoTo 異寫字處理完畢
                End If
             End If
             
             If 中研院金文(系統字體) Or 中研院甲骨文(系統字體) Or 中研院楚系文字(系統字體) Then
                If Not IsNull(異體字表.Fields("字體")) Then
                    暫存字體 = 區域字體陣列(異體字表.Fields("字體"))
                    字形 = 異體字表.Fields("字碼")
                Else
                    暫存字體 = 區域字體陣列(檢字表.Fields("字體"))
                End If
             Else
                暫存字體 = 區域字體陣列(檢字表.Fields("字體"))
             End If
             tree字形樹狀結構.AddItem 字形, i

             'If 系統字體 <> "中研院金文" And 暫存字體 = "中研院金文" Then 暫存字體 = "hzkf"
             tree字形樹狀結構.ItemFontName(tree字形樹狀結構.NewIndex) = 轉換顯示字型(暫存字體)
             tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = 異體字表.Fields("編號")
             If Not IsNull(檢字表.Fields("字形")) Then
                tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 字形節點標記
             Else
                tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 構字式節點標記
             End If
             
             If 異寫型態 = 2 Then
                tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureClosed
                tree字形樹狀結構.AddItem "", tree字形樹狀結構.NewIndex
                tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
             End If
             
             If 系統字體 <> "北師大說文小篆" And 系統字體 <> "北師大說文重文" And Not 中研院金文(系統字體) And Not 中研院甲骨文(系統字體) And Not 中研院楚系文字(系統字體) Then
                'tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PicturePalette '簡體字
                If Not IsNull(異體字表.Fields("類別")) Then
                    If 異體字表.Fields("類別") = 簡化字代碼 Then
                        tree字形樹狀結構.ItemMark(tree字形樹狀結構.NewIndex) = 簡化字代碼
                    ElseIf 異體字表.Fields("類別") = 小篆代碼 Then
                        tree字形樹狀結構.ItemMark(tree字形樹狀結構.NewIndex) = 小篆代碼
                    ElseIf 異體字表.Fields("類別") = 或體代碼 Then
                        tree字形樹狀結構.ItemMark(tree字形樹狀結構.NewIndex) = 或體代碼
                    ElseIf 異體字表.Fields("類別") = 古文代碼 Then
                        tree字形樹狀結構.ItemMark(tree字形樹狀結構.NewIndex) = 古文代碼
                    ElseIf 異體字表.Fields("類別") = 籀文代碼 Then
                        tree字形樹狀結構.ItemMark(tree字形樹狀結構.NewIndex) = 籀文代碼
                    ElseIf 異體字表.Fields("類別") = 俗字代碼 Then
                        tree字形樹狀結構.ItemMark(tree字形樹狀結構.NewIndex) = 俗字代碼
                    ElseIf 異體字表.Fields("類別") = 奇字代碼 Then
                        tree字形樹狀結構.ItemMark(tree字形樹狀結構.NewIndex) = 奇字代碼
                    ElseIf 異體字表.Fields("類別") = 金文代碼 Then
                        tree字形樹狀結構.ItemMark(tree字形樹狀結構.NewIndex) = 金文代碼
                    ElseIf 異體字表.Fields("類別") = 甲骨文代碼 Then
                        tree字形樹狀結構.ItemMark(tree字形樹狀結構.NewIndex) = 甲骨文代碼
                    ElseIf 異體字表.Fields("類別") = 楚系文字代碼 Then
                        tree字形樹狀結構.ItemMark(tree字形樹狀結構.NewIndex) = 楚系文字代碼
                    ElseIf 異體字表.Fields("類別") = 訛字代碼 Then
                        tree字形樹狀結構.ItemMark(tree字形樹狀結構.NewIndex) = 訛字代碼
                    End If
                    標記個數 = 標記個數 + 1
                End If
             'ElseIf 異體字表.Fields("形碼") > 100 Then   '形碼>100
             '   tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = Image1
             'ElseIf 檢字表.Fields("連接符號") = 0 Then
                tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf  '字根
             'Else
             '   tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf  '部件
             End If
          ElseIf 異體字表.Fields("形碼") = 0 Then
          
             楷書編號 = 異體字表.Fields("楷書編號")
             楷書檢字表.Index = "編號"
             楷書檢字表.Seek "=", 楷書編號
             If Not IsNull(楷書檢字表.Fields("小篆編號")) Then
                If 總數 > 1 Then
                    tree字形樹狀結構.AddItem 楷書檢字表.Fields("小篆"), 0
                Else
                    tree字形樹狀結構.AddItem 楷書檢字表.Fields("小篆") ', 0
                End If
                tree字形樹狀結構.ItemFontName(tree字形樹狀結構.NewIndex) = "北師大說文小篆"
                tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = 楷書檢字表.Fields("小篆編號")
                tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureClosed
                i = tree字形樹狀結構.NewIndex
             Else
                If 總數 > 1 Then
                    tree字形樹狀結構.AddItem 楷書檢字表.Fields("字碼"), 0
                Else
                    tree字形樹狀結構.AddItem 楷書檢字表.Fields("字碼") ', 0
                End If
                暫存字體 = 區域字體陣列(楷書檢字表.Fields("字體"))
                tree字形樹狀結構.ItemFontName(tree字形樹狀結構.NewIndex) = 轉換顯示字型(暫存字體)
                tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = 楷書檢字表.Fields("編號")
                tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureClosed
                i = tree字形樹狀結構.NewIndex
             End If
             If Not IsNull(楷書檢字表.Fields("字形")) Then
                tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 字形節點標記
             Else
                tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 構字式節點標記
             End If

          Else
             'If 總數 > 1 Or (總數 = 1 And 字頭表.Fields("形碼") <> 1) Then
             If 中研院金文(系統字體) Or 中研院甲骨文(系統字體) Or 中研院楚系文字(系統字體) Then
                異寫型態 = 0
                If Not IsNull(異體字表.Fields("型態")) Then 異寫型態 = 異體字表.Fields("型態")
                If 異寫型態 = 1 Then
                    載入異寫字一 系統字體, 異體字表.Fields("編號"), i
                    tree字形樹狀結構.Expand(i) = True
                    字頭表.MoveNext
                    GoTo 異寫字處理完畢
                End If

                If 中研院金文(系統字體) Then
                    If Not IsNull(異體字表.Fields("字體")) Then
                        顯示字型 = 區域字體陣列(異體字表.Fields("字體"))
                        字形 = 異體字表.Fields("字碼")
                    Else
                        顯示字型 = "中研院金文"
                    End If
                ElseIf 中研院甲骨文(系統字體) Then
                    If Not IsNull(異體字表.Fields("字體")) Then
                        顯示字型 = 區域字體陣列(異體字表.Fields("字體"))
                        字形 = 異體字表.Fields("字碼")
                    Else
                        顯示字型 = "中研院甲骨文"
                    End If
                 ElseIf 中研院楚系文字(系統字體) Then
                    If Not IsNull(異體字表.Fields("字體")) Then
                        顯示字型 = 區域字體陣列(異體字表.Fields("字體"))
                        字形 = 異體字表.Fields("字碼")
                    Else
                        顯示字型 = "中研院楚系簡帛文字"
                    End If
                End If
                tree字形樹狀結構.AddItem 字形, i
                tree字形樹狀結構.ItemFontName(tree字形樹狀結構.NewIndex) = 顯示字型
                tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = 異體字表.Fields("編號") '檢字表.Fields("編號")
                If Not IsNull(檢字表.Fields("字形")) Then
                    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 字形節點標記
                Else
                    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 構字式節點標記
                End If
                tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
                
                If 異寫型態 = 2 Then
                    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureClosed
                    tree字形樹狀結構.AddItem "", tree字形樹狀結構.NewIndex
                    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
                End If

             'ElseIf 總數 > 1 Or (總數 = 1 And 字頭表.Fields(1) <> 1) Then
             ElseIf 總數 > 1 Then
                tree字形樹狀結構.AddItem 字形, 0
                暫存字體 = 區域字體陣列(檢字表.Fields("字體"))
                If 系統字體 <> "中研院金文" And 暫存字體 = "中研院金文" Then 暫存字體 = "hzkf"
                tree字形樹狀結構.ItemFontName(tree字形樹狀結構.NewIndex) = 轉換顯示字型(暫存字體)
                tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = 檢字表.Fields("編號")
                If Not IsNull(檢字表.Fields("字形")) Then
                    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 字形節點標記
                Else
                    tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 構字式節點標記
                End If
                tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureClosed
                i = tree字形樹狀結構.NewIndex
             Else
                tree字形樹狀結構.AddItem 字形
                tree字形樹狀結構.ItemFontName(0) = 轉換顯示字型(區域字體陣列(檢字表.Fields("字體")))
                tree字形樹狀結構.Image(0) = tree字形樹狀結構.PictureOpen
                tree字形樹狀結構.ItemLngValue(0) = 檢字表.Fields("編號")
                i = 0
                If Not IsNull(檢字表.Fields("字形")) Then
                    tree字形樹狀結構.ItemTag(0) = 字形節點標記
                Else
                    tree字形樹狀結構.ItemTag(0) = 構字式節點標記
                End If
             End If
          End If
          異體字表.MoveNext
       Loop
       tree字形樹狀結構.Expand(i) = True
        
       字頭表.MoveNext
         
       If (tree字形樹狀結構.ListCount + 10) Mod 50 = 0 Then
          tree字形樹狀結構.Redraw = True
          Screen.MousePointer = ccDefault
          DoEvents
          tree字形樹狀結構.Redraw = False
          Screen.MousePointer = ccHourglass
       End If

異寫字處理完畢:

   Loop
Else
   tree字形樹狀結構.AddItem 樹根
   tree字形樹狀結構.ItemFontName(0) = 轉換顯示字型(字體)
   tree字形樹狀結構.ItemLngValue(0) = 編號
   If Not IsNull(檢字表.Fields("字形")) Then
      tree字形樹狀結構.ItemTag(0) = 字形節點標記
   Else
      tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 構字式節點標記
   End If

   '檢字表.Index = "編號"
   '檢字表.Seek "=", 編號
   'If Not 檢字表.NoMatch Then
      'If 檢字表.Fields("連接符號") = 0 Then
         tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
      'Else
         'tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
      'End If
   'End If

End If


If 標記個數 > 0 Then
    tree字形樹狀結構.ViewStyleEx = TLVIEWSEX_MARKPM
Else
    tree字形樹狀結構.ViewStyleEx = TLVIEWSEX_PLUSMIN
End If

tree字形樹狀結構.Redraw = True
異體狀態列 = "載入完畢！ 共 " & 樹根序數 & " 個字形"
mdi漢字字形.txt狀態 = 異體狀態列
'mdi漢字字形.sbar狀態列.Visible = True

Screen.MousePointer = ccDefault

End Sub

Private Sub tree字形樹狀結構_DragOver(Source As Control, x As Single, y As Single, State As Integer)

tree字形樹狀結構.OnDragOver x, y, State

End Sub

Private Sub tree字形樹狀結構_GotFocus()

'tree字形樹狀結構_Click
現用視窗代碼 = 異體字表代碼

End Sub

Private Sub tree字形樹狀結構_LostFocus()

tree字形樹狀結構.ListIndex = -1

End Sub

Private Sub tree字形樹狀結構_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

XCheck = x
YCheck = y

End Sub

Public Sub 載入異寫字一(系統字體 As String, 編號 As Long, i As Long)

Dim SQL陳述式 As String, 字體編號 As Integer

SQL陳述式 = "SELECT * From 異寫字表 Where (異寫編號= " & 編號 & ") and ((上線)>0 and (上線)<6) ORDER BY 編號"
If 中研院金文(系統字體) Then
    Set 異寫字表 = 金文資料庫.OpenRecordset(SQL陳述式)
ElseIf 中研院甲骨文(系統字體) Then
    Set 異寫字表 = 甲骨文資料庫.OpenRecordset(SQL陳述式)
ElseIf 中研院楚系文字(系統字體) Then
    Set 異寫字表 = 楚系文字資料庫.OpenRecordset(SQL陳述式)
End If
                    
異寫字表.MoveLast
異寫字表.MoveFirst
Do Until 異寫字表.EOF
    tree字形樹狀結構.AddItem 異寫字表.Fields("字碼"), i
    字體編號 = 異寫字表.Fields("字體")
    tree字形樹狀結構.ItemFontName(tree字形樹狀結構.NewIndex) = 區域字體陣列(異寫字表.Fields("字體"))
    tree字形樹狀結構.ItemLngValue(tree字形樹狀結構.NewIndex) = 異寫字表.Fields("編號")
    If Not IsNull(檢字表.Fields("字形")) Then
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 字形節點標記
    Else
        tree字形樹狀結構.ItemTag(tree字形樹狀結構.NewIndex) = 構字式節點標記
    End If
    tree字形樹狀結構.Image(tree字形樹狀結構.NewIndex) = tree字形樹狀結構.PictureLeaf
    異寫字表.MoveNext
Loop

End Sub

Public Sub 載入異寫二(i As Long)

Dim 字體 As String

載入異寫字一 tree字形樹狀結構.ItemFontName(i), tree字形樹狀結構.ItemLngValue(i), i

End Sub
