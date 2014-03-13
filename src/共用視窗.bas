Attribute VB_Name = "公用視窗"
Option Explicit
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public Sub 計算結束視窗()
Dim i As Integer

部件open = 0
孳乳open = 0
出處open = 0
結構open = 0
演變open = 0
索引open = 0
異體open = 0
異根open = 0

For i = 1 To Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
           Case Big5字根代碼 To 構字符號代碼
                部件open = 1
                部件winstate = frm部件範例.WindowState
                部件left = frm部件範例.Left
                部件top = frm部件範例.Top
                部件width = frm部件範例.Width
                部件height = frm部件範例.Height
           Case 字形孳乳代碼
                孳乳open = 1
                孳乳winstate = frm字形孳乳.WindowState
                孳乳left = frm字形孳乳.Left
                孳乳top = frm字形孳乳.Top
                孳乳width = frm字形孳乳.Width
                孳乳height = frm字形孳乳.Height
           Case 出處檢字代碼
                出處open = 1
                出處winstate = frm出處檢字.WindowState
                出處left = frm出處檢字.Left
                出處top = frm出處檢字.Top
                出處width = frm出處檢字.Width
                出處height = frm出處檢字.Height
           Case 字形結構代碼
                結構open = 1
                結構winstate = frm字形結構.WindowState
                結構left = frm字形結構.Left
                結構top = frm字形結構.Top
                結構width = frm字形結構.Width
                結構height = frm字形結構.Height
           Case 異體字表代碼
                異體open = 1
                異體winstate = frm異體字表.WindowState
                異體left = frm異體字表.Left
                異體top = frm異體字表.Top
                異體width = frm異體字表.Width
                異體height = frm異體字表.Height
           Case 異體字根代碼
                異根open = 1
                異根winstate = frm異體字根.WindowState
                異根left = frm異體字根.Left
                異根top = frm異體字根.Top
                異根width = frm異體字根.Width
                異根height = frm異體字根.Height
           Case 字形演變代碼
                演變open = 1
                演變winstate = frm字形演變.WindowState
                演變left = frm字形演變.Left
                演變top = frm字形演變.Top
                演變width = frm字形演變.Width
                演變height = frm字形演變.Height
           Case 字形索引代碼
                索引open = 1
                索引winstate = frm字形索引.WindowState
                索引left = frm字形索引.Left
                索引top = frm字形索引.Top
                索引width = frm字形索引.Width
                索引height = frm字形索引.Height
    End Select
Next i

End Sub

Public Sub 計算現用視窗()

If Forms.Count = 2 Then
   現用視窗 = "mdi漢字字形"
   現用視窗代碼 = mdi漢字字形代碼
   設定工具列初始狀態
End If

End Sub

Public Sub 切換選取字形工具列狀態(視窗代碼 As Integer, Optional 筆畫 As Integer = 1000, Optional 首筆 As Integer)
Dim i As Integer, show As Boolean

show = False

For i = 1 To Forms.Count - 1
    If (CInt(Forms(i).Tag) >= Big5字根代碼) And (CInt(Forms(i).Tag) <= 說文部首代碼) Then show = True
Next i

If show Then
   mdi漢字字形.cbo筆畫.Enabled = True
   mdi漢字字形.cbo首筆.Enabled = True
   筆畫首筆查詢 = False
   If 筆畫 <> 1000 Then
    mdi漢字字形.cbo筆畫.ListIndex = 筆畫
    mdi漢字字形.cbo首筆.ListIndex = 首筆
   End If
   筆畫首筆查詢 = True
Else
   mdi漢字字形.cbo筆畫.Enabled = False
   mdi漢字字形.cbo首筆.Enabled = False
   筆畫首筆查詢 = False
End If
    
End Sub

Public Sub 設定工具列初始狀態()
mdi漢字字形.cbo筆畫.Enabled = False
mdi漢字字形.cbo首筆.Enabled = False
mdi漢字字形.cbo筆畫.ListIndex = 1
mdi漢字字形.cbo首筆.ListIndex = 0

End Sub
