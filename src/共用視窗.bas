Attribute VB_Name = "���ε���"
Option Explicit
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public Sub �p�⵲������()
Dim i As Integer

����open = 0
�F��open = 0
�X�Bopen = 0
���copen = 0
�t��open = 0
����open = 0
����open = 0
����open = 0

For i = 1 To Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
           Case Big5�r�ڥN�X To �c�r�Ÿ��N�X
                ����open = 1
                ����winstate = frm����d��.WindowState
                ����left = frm����d��.Left
                ����top = frm����d��.Top
                ����width = frm����d��.Width
                ����height = frm����d��.Height
           Case �r�δF�ťN�X
                �F��open = 1
                �F��winstate = frm�r�δF��.WindowState
                �F��left = frm�r�δF��.Left
                �F��top = frm�r�δF��.Top
                �F��width = frm�r�δF��.Width
                �F��height = frm�r�δF��.Height
           Case �X�B�˦r�N�X
                �X�Bopen = 1
                �X�Bwinstate = frm�X�B�˦r.WindowState
                �X�Bleft = frm�X�B�˦r.Left
                �X�Btop = frm�X�B�˦r.Top
                �X�Bwidth = frm�X�B�˦r.Width
                �X�Bheight = frm�X�B�˦r.Height
           Case �r�ε��c�N�X
                ���copen = 1
                ���cwinstate = frm�r�ε��c.WindowState
                ���cleft = frm�r�ε��c.Left
                ���ctop = frm�r�ε��c.Top
                ���cwidth = frm�r�ε��c.Width
                ���cheight = frm�r�ε��c.Height
           Case ����r��N�X
                ����open = 1
                ����winstate = frm����r��.WindowState
                ����left = frm����r��.Left
                ����top = frm����r��.Top
                ����width = frm����r��.Width
                ����height = frm����r��.Height
           Case ����r�ڥN�X
                ����open = 1
                ����winstate = frm����r��.WindowState
                ����left = frm����r��.Left
                ����top = frm����r��.Top
                ����width = frm����r��.Width
                ����height = frm����r��.Height
           Case �r�κt�ܥN�X
                �t��open = 1
                �t��winstate = frm�r�κt��.WindowState
                �t��left = frm�r�κt��.Left
                �t��top = frm�r�κt��.Top
                �t��width = frm�r�κt��.Width
                �t��height = frm�r�κt��.Height
           Case �r�ί��ޥN�X
                ����open = 1
                ����winstate = frm�r�ί���.WindowState
                ����left = frm�r�ί���.Left
                ����top = frm�r�ί���.Top
                ����width = frm�r�ί���.Width
                ����height = frm�r�ί���.Height
    End Select
Next i

End Sub

Public Sub �p��{�ε���()

If Forms.Count = 2 Then
   �{�ε��� = "mdi�~�r�r��"
   �{�ε����N�X = mdi�~�r�r�ΥN�X
   �]�w�u��C��l���A
End If

End Sub

Public Sub ��������r�Τu��C���A(�����N�X As Integer, Optional ���e As Integer = 1000, Optional ���� As Integer)
Dim i As Integer, show As Boolean

show = False

For i = 1 To Forms.Count - 1
    If (CInt(Forms(i).Tag) >= Big5�r�ڥN�X) And (CInt(Forms(i).Tag) <= ���峡���N�X) Then show = True
Next i

If show Then
   mdi�~�r�r��.cbo���e.Enabled = True
   mdi�~�r�r��.cbo����.Enabled = True
   ���e�����d�� = False
   If ���e <> 1000 Then
    mdi�~�r�r��.cbo���e.ListIndex = ���e
    mdi�~�r�r��.cbo����.ListIndex = ����
   End If
   ���e�����d�� = True
Else
   mdi�~�r�r��.cbo���e.Enabled = False
   mdi�~�r�r��.cbo����.Enabled = False
   ���e�����d�� = False
End If
    
End Sub

Public Sub �]�w�u��C��l���A()
mdi�~�r�r��.cbo���e.Enabled = False
mdi�~�r�r��.cbo����.Enabled = False
mdi�~�r�r��.cbo���e.ListIndex = 1
mdi�~�r�r��.cbo����.ListIndex = 0

End Sub
