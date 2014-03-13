Attribute VB_Name = "公用變數"
Option Explicit
Public Const mdi漢字字形代碼 = 0
Public Const Big5字根代碼 = 1, Big5及簡化字字根代碼 = 2, 字根代碼 = 3, 小篆獨體字代碼 = 4, 金文字根代碼 = 5, 甲骨文字根代碼 = 6, 楚系簡帛文字字根代碼 = 7, 部件外字代碼 = 8
Public Const 康熙部首代碼 = 9, 說文部首代碼 = 10, 簡牘代碼 = 11, 八卦代碼 = 12, 圖形文字代碼 = 13, 構字符號代碼 = 14
Public Const 字形孳乳代碼 = 15, 出處檢字代碼 = 16, 字形結構代碼 = 17, 異體字表代碼 = 18, 異體字根代碼 = 19
Public Const 字形演變代碼 = 20, 字形索引代碼 = 21
Public Const 字形孳乳_檢字方塊 = 1, 字形孳乳_樹狀結構 = 2
Public Const 出處檢字_檢字方塊 = 1, 出處檢字_樹狀結構 = 2
Public Const mdi漢字字形_編號方塊 = 1, mdi漢字字形_外字集方塊 = 2, mdi漢字字形_字形方塊 = 3
Public Const mdi漢字字形_總筆畫方塊 = 4, mdi漢字字形_部首方塊 = 5, mdi漢字字形_扣除部首筆畫方塊 = 6
Public Const mdi漢字字形_注音方塊 = 7, mdi漢字字形_內碼方塊 = 8, mdi漢字字形_倉頡碼方塊 = 9
Public Const mdi漢字字形_構字式方塊 = 10, mdi漢字字形_冊數方塊 = 11, mdi漢字字形_組字字數方塊 = 12, mdi漢字字形_組字字數含異寫方塊 = 13
Public Const mdi漢字字形_重文方塊 = 14, mdi漢字字形_古漢字方塊 = 15
Public Const 字形節點標記 = 0, 構字式節點標記 = 1, 器名節點標記 = 2, 其他節點標記 = 3
Public Const 字體個數 = 27
Public 複製風格碼 As Boolean, 複製Big5字元 As Boolean, 複製unicode字元 As Boolean
Public Const ccDefault = 0, ccHourglass = 11
Public Const 簡易瀏覽總數 = 4, 預設瀏覽總數 = 4
Public 系統資料庫 As Database, 小篆資料庫 As Database, 金文資料庫 As Database, 甲骨文資料庫 As Database, 楚系文字資料庫 As Database
Public 常用符號及部件 As Recordset, 康熙部首 As Recordset, 說文部首 As Recordset, 常用符號及部件類型 As Recordset
Public 字形表 As Recordset
Public 檢字表 As Recordset, 異體字表 As Recordset, 異寫字根 As Recordset
Public 楷書檢字表 As Recordset, 楷書異體字表 As Recordset, 楷書異寫字根 As Recordset, 楷書字根 As Recordset
Public 小篆檢字表 As Recordset, 小篆異體字表 As Recordset, 小篆異寫字根 As Recordset, 小篆獨體字 As Recordset
Public 金文檢字表 As Recordset, 金文補遺表 As Recordset, 金文異體字表 As Recordset, 金文異寫字表 As Recordset, 金文集成器名 As Recordset, 金文集成引得 As Recordset, 金文詁林 As Recordset, 金文異寫字根 As Recordset, 金文字根 As Recordset
Public 甲骨文檢字表 As Recordset, 甲骨文異體字表 As Recordset, 甲骨文異寫字表 As Recordset, 甲骨文異寫字根 As Recordset, 甲骨文字根 As Recordset
Public 楚系文字檢字表 As Recordset, 楚系文字補遺表 As Recordset, 楚系文字異體字表 As Recordset, 楚系文字異寫字表 As Recordset, 楚系文字異寫字根 As Recordset, 楚系文字字根 As Recordset
Public 現用視窗 As String, 現用視窗代碼 As Integer, 現用控制項代碼 As Integer
Public 現用字體 As String, 系統字體 As String, 顯示字型 As String, 顯示字型大小 As Integer
Public 圖片解析度 As Long, 圖片字型大小 As Long
Public 複製圖片到Word As Boolean, 複製到Word的圖片大小 As Integer
Public 筆畫首筆查詢 As Boolean
Public 組字符號陣列(1 To 14) As String
Public 啟動字形孳乳 As Boolean, 啟動出處檢字 As Boolean, 啟動字形結構 As Boolean, 啟動異體字表 As Boolean, 啟動異體字根 As Boolean, 啟動字形演變 As Boolean, 啟動字形索引 As Boolean, 啟動部件範例 As Boolean
Public 共用視窗(mdi漢字字形代碼 To 字形索引代碼) As String, 共用視窗代碼 As Integer
Public 樹根序數 As Integer
Public 字體陣列(0 To 字體個數) As Variant
Public 欄寬 As Integer
Public 狀態列字數 As Long
Public 系統編號 As Long
Public 視窗代碼(0 To 2) As Boolean
Public 字集_常用字 As String * 250, 字集_五大碼 As String * 250, 字集_簡化字 As String * 250, 字集_漢語大字典 As String * 250, 字集_說文解字 As String * 250, 字集_金文編 As String * 250, 字集_金文編圖形文字 As String * 250, 字集_甲骨類纂 As String * 250, 字集_楚系簡帛文字編 As String * 250, 字集_楷體字 As String * 250
Public 孳乳open As String * 250, 孳乳winstate As String * 250, 孳乳left As String * 250, 孳乳top As String * 250, 孳乳width As String * 250, 孳乳height As String * 250
Public 出處open As String * 250, 出處winstate As String * 250, 出處left As String * 250, 出處top As String * 250, 出處width As String * 250, 出處height As String * 250
Public 結構open As String * 250, 結構winstate As String * 250, 結構left As String * 250, 結構top As String * 250, 結構width As String * 250, 結構height As String * 250
Public 異體open As String * 250, 異體winstate As String * 250, 異體left As String * 250, 異體top As String * 250, 異體width As String * 250, 異體height As String * 250
Public 異根open As String * 250, 異根winstate As String * 250, 異根left As String * 250, 異根top As String * 250, 異根width As String * 250, 異根height As String * 250
Public 部件open As String * 250, 部件winstate As String * 250, 部件left As String * 250, 部件top As String * 250, 部件width As String * 250, 部件height As String * 250
Public 演變open As String * 250, 演變winstate As String * 250, 演變left As String * 250, 演變top As String * 250, 演變width As String * 250, 演變height As String * 250
Public 索引open As String * 250, 索引winstate As String * 250, 索引left As String * 250, 索引top As String * 250, 索引width As String * 250, 索引height As String * 250
Public 初始first As String * 250, 初始部件順序 As String * 250, 初始異寫部件 As String * 250, 初始逐級列出 As String * 250, 初始解形列出 As String * 250, 初始copy As String
Public 初始遠東漢語大字典 As String * 250, 初始建宏漢語大字典 As String * 250, 初始中文大辭典 As String * 250
Public 初始說文解字詁林 As String * 250, 初始中華說文解字 As String * 250
Public 初始金文編 As String * 250, 初始金文詁林 As String * 250, 初始金文器號 As String * 250, 初始金文引得 As String * 250
Public 初始甲骨刻辭類纂 As String * 250, 初始甲骨文字詁林 As String * 250, 初始甲骨文字集釋 As String * 250
Public 初始楚系簡帛文字編 As String * 250, 初始楚系文字出處 As String * 250
Public 初始Unicode As String * 250, 初始Big5 As String * 250
Public 初始甲骨文演變 As String * 250, 初始金文演變 As String * 250, 初始楚系文字演變 As String * 250, 初始小篆演變 As String * 250
Public 初始字頻 As String * 250, 初始風格碼 As String * 250
Public 初始CopyToWord As String * 250, 初始CopyUnicode As String * 250
Public 出處為甲骨文 As Boolean, 出處為金文 As Boolean, 出處為小篆 As Boolean, 出處為楚文字 As Boolean
Public 出處為甲骨文合集 As Boolean, 出處完全匹配 As Boolean
Public 已載入畫面 As Integer
Public 預設瀏覽模式 As Integer, 簡易瀏覽模式 As Boolean, 改變預設瀏覽 As Boolean
Public 孳乳狀態列 As String, 結構狀態列 As String, 部件狀態列 As String, 異體狀態列 As String
Public 欄位別 As String
Public 拖曳字串 As String
Public 暫存目錄 As String, bmpcount As Long, 暫存圖檔 As String, 替代文字 As String

Public WordApp As Word.Application, WordWasNotRunning As Boolean
