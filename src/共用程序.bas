Attribute VB_Name = "���ε{��"
Option Explicit
Public Const rtf_version = "\rtf1", character_set = "\ansi\deflang1033"
Public Const �з��� = "{\f0\fscript\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9;}"
Public Const �з���~�r���@ = "{\f1\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'40;}"
Public Const �з���~�r���G = "{\f2\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'47;}"
Public Const �з���~�r���T = "{\f3\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'54;}"
Public Const �з���~�r���| = "{\f4\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a5\'7c;}"
Public Const �з���~�r���� = "{\f5\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'ad;}"
Public Const �з���~�r���� = "{\f6\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'bb;}"
Public Const �з���~�r���C = "{\f7\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'43;}"
Public Const �з���~�r���K = "{\f8\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'4b;}"
Public Const �з���~�r���E = "{\f9\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'45;}"
Public Const �ө��� = "{\f16\fscript\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9;}"
Public Const �ө���~�r���@ = "{\f17\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'40;}"
Public Const �ө���~�r���G = "{\f18\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'47;}"
Public Const �ө���~�r���T = "{\f19\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'54;}"
Public Const �ө���~�r���| = "{\f20\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a5\'7c;}"
Public Const �ө���~�r���� = "{\f21\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'ad;}"
Public Const �ө���~�r���� = "{\f22\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'bb;}"
Public Const �ө���~�r���C = "{\f23\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'43;}"
Public Const �ө���~�r���K = "{\f24\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'4b;}"
Public Const �ө���~�r���E = "{\f25\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'45;}"
Public Const �з���r�� = �з��� & �з���~�r���@ & �з���~�r���G & �з���~�r���T & �з���~�r���| & �з���~�r���� & �з���~�r���� & �з���~�r���C & �з���~�r���K & �з���~�r���E
Public Const �з���r���� = "{\fonttbl" & �з���r�� & "}"
Public Const �ө���r�� = �ө��� & �ө���~�r���@ & �ө���~�r���G & �ө���~�r���T & �ө���~�r���| & �ө���~�r���� & �ө���~�r���� & �ө���~�r���C & �ө���~�r���K & �ө���~�r���E
Public Const �ө���r���� = "{\fonttbl" & �ө���r�� & "}"

Public Function ����~�rSQL���z��(��ƪ� As String, ���e As Integer, ���� As Integer) As String
Dim ��ܶ��� As String
Dim ���� As String

���� = ""
If ��ƪ� = "�c�r�Ÿ�" Then
   ��ܶ��� = "select �r�� from �Ÿ� order by �s�� "
ElseIf ��ƪ� = "�ϧΤ�r" Then
   ��ܶ��� = "select �r�� from �ϧΤ�r order by �s�� "
ElseIf ��ƪ� = "�K��" Then
   ��ܶ��� = "select �r�� from �K�� order by �s�� "
ElseIf ��ƪ� = "²�|" Then
   ��ܶ��� = "select �r�� from ²�| order by �s�� "
ElseIf ��ƪ� = "�d���r�峡��" Then
   ��ܶ��� = "select �r�� from �d������ "
ElseIf ��ƪ� = "����Ѧr����" Then
   ��ܶ��� = "select �r�� from ���峡�� "
ElseIf ��ƪ� = "Big5�r��" Then
   ��ܶ��� = "select �r�� from �r�� where Big5=1 "
ElseIf ��ƪ� = "Big5��²�Ʀr�r��" Then
   ��ܶ��� = "select �r�� from �r�� where (Big5=1 or ²�Ʀr=1) "
ElseIf ��ƪ� = "�r��" Then
   ��ܶ��� = "select �r�� from �r�� "
ElseIf ��ƪ� = "�p�f�W��r" Then
   ��ܶ��� = "select �r�� from �p�f�W��r where �s��<9999999"
ElseIf ��ƪ� = "����r��" Then
   ��ܶ��� = "select �r�� from ����r�� where �s��<9999999"
ElseIf ��ƪ� = "�Ұ���r��" Then
   ��ܶ��� = "select �r�� from �Ұ���r�� where �s��<9999999"
ElseIf ��ƪ� = "���t²����r�r��" Then
   ��ܶ��� = "select �r�� from ���t²����r�r�� where �s��<9999999"
Else
   ��ܶ��� = "select �r�� from ����~�r where �s��<9999999"
End If

If ��ƪ� <> "�c�r�Ÿ�" And ��ƪ� <> "�K��" And ��ƪ� <> "²�|" And ��ƪ� <> "�ϧΤ�r" Then
   If ���e > 0 And ���� > 0 Then
      If ��ƪ� = "�d���r�峡��" Or ��ƪ� = "����Ѧr����" Or ��ƪ� = "�r��" Then
         ���� = "where ���e = " & ���e & " and ���� = " & ���� & " order by �s��"
      Else
         ���� = "and ���e = " & ���e & " and ���� = " & ���� & " order by �s��"
      End If
   ElseIf ���e > 0 Then
      If ��ƪ� = "�d���r�峡��" Or ��ƪ� = "����Ѧr����" Or ��ƪ� = "�r��" Then
         ���� = "where ���e = " & ���e & " order by ����,�s��"
      Else
         ���� = "and ���e = " & ���e & " order by ����,�s��"
      End If
   ElseIf ���� > 0 Then
      If ��ƪ� = "�d���r�峡��" Or ��ƪ� = "����Ѧr����" Or ��ƪ� = "�r��" Then
         ���� = "where ���� = " & ���� & " order by ���e,�s��"
      Else
         ���� = "and ���� = " & ���� & " order by ���e,�s��"
      End If
   Else
      ���� = ""
   End If
End If

����~�rSQL���z�� = ��ܶ��� & ����


End Function

Public Function �ഫ�^���ܾe(�^��X As String) As String

Dim �^��r�� As String, �ܾe�r�� As String
Dim i As Integer
Dim �}�C As Variant

�}�C = Array("��", "��", "��", "��", "��", "��", "�g", "��", "��", "�Q", "�j", "��", "�@", "�}", "�H", "��", "��", "�f", "�r", "��", "�s", "�k", "��", "��", "�R", "��")
For i = 1 To Len(�^��X)
    �^��r�� = Mid(�^��X, i, 1)
    �ܾe�r�� = �}�C(Asc(�^��r��) - 65)
    �ഫ�^���ܾe = �ഫ�^���ܾe & �ܾe�r��
Next i

End Function


Public Function �ഫ�ܾe��^��(�ܾe�X As String) As String
Dim �^��X As String
Dim �^��r�� As String, �ܾe�r�� As String
Dim i As Integer

�^��X = ""
For i = 1 To Len(�ܾe�X)
    �ܾe�r�� = Mid(�ܾe�X, i, 1)
    �^��r�� = �ܾe��^��(�ܾe�r��)
    �^��X = �^��X & �^��r��
Next i
End Function

Public Function �ܾe��^��(�ܾe�r�� As String)
Select Case �ܾe�r��
    Case "��"
        �ܾe��^�� = "A"
    Case "��"
        �ܾe��^�� = "B"
    Case "��"
        �ܾe��^�� = "C"
    Case "��"
        �ܾe��^�� = "D"
    Case "��"
        �ܾe��^�� = "E"
    Case "��"
        �ܾe��^�� = "F"
    Case "�g"
        �ܾe��^�� = "G"
    Case "��"
        �ܾe��^�� = "H"
    Case "��"
        �ܾe��^�� = "I"
    Case "�Q"
        �ܾe��^�� = "J"
    Case "�j"
        �ܾe��^�� = "K"
    Case "��"
        �ܾe��^�� = "L"
    Case "�@"
        �ܾe��^�� = "M"
    Case "�}"
        �ܾe��^�� = "N"
    Case "�H"
        �ܾe��^�� = "O"
    Case "��"
        �ܾe��^�� = "P"
    Case "��"
        �ܾe��^�� = "Q"
    Case "�f"
        �ܾe��^�� = "R"
    Case "�r"
        �ܾe��^�� = "S"
    Case "��"
        �ܾe��^�� = "T"
    Case "�s"
        �ܾe��^�� = "U"
    Case "�k"
        �ܾe��^�� = "V"
    Case "��"
        �ܾe��^�� = "W"
    Case "��"
        �ܾe��^�� = "X"
    Case "�R"
        �ܾe��^�� = "Y"
    Case "��"
        �ܾe��^�� = "Z"
    Case Else
        �ܾe��^�� = �ܾe�r��
 
End Select

End Function

Public Function �M��զr��(���� As Variant, ����� As Variant) As String
Dim i As Integer, j As Integer
Dim �զr�� As String, �Ȧs�զr�� As String

If ���� < 4 Then
   �զr�� = ""
   If ���� <> 0 Then
      i = 1
      Do Until i > Len(�����)
         �զr�� = �զr�� & Mid(�����, i, 1)
         For j = 4 To 11
             If Mid(�����, i, 1) = �զr�Ÿ��}�C(j) Then
                i = i + 1
                �զr�� = �զr�� & Mid(�����, i, 1)
                Exit For
             End If
         Next j
         �զr�� = �զr�� & �զr�Ÿ��}�C(����)
         i = i + 1
      Loop
      �Ȧs�զr�� = Mid(�զr��, 1, Len(�զr��) - 1)
   Else
      �Ȧs�զr�� = �����
   End If
Else
   If �O�_���զr�Ÿ�(Mid(�����, 1, 1), 1, 14) > 0 And Len(�����) <= 2 Then
      �Ȧs�զr�� = �����
   Else
      �Ȧs�զr�� = �զr�Ÿ��}�C(12) & ����� & �զr�Ÿ��}�C(13)
   End If
End If

�M��զr�� = �Ȧs�զr��

End Function

Public Function �M�䭷��X(�r�Y As String, ����� As String, ����X As String) As String

If �r�Y <> "" Then
    �M�䭷��X = "��" & �r�Y & "ơ" & ����X & "��"
Else
    �M�䭷��X = "��" & ����� & "ơ" & ����X & "��"
End If

End Function


Public Function �O�_���զr�Ÿ�(�r�� As String, x As Integer, y As Integer) As Integer
Dim i As Integer

�O�_���զr�Ÿ� = 0
For i = x To y
    If �r�� = �զr�Ÿ��}�C(i) Then
       �O�_���զr�Ÿ� = i
       Exit For
    End If
Next i

End Function

Public Function �O�_���U�Φr��(�r�� As String) As Boolean


�O�_���U�Φr�� = InStr(1, "*?#[]!-", �r��) > 0

End Function

Public Sub �^���ݩ�(�t�Φr�� As String, val�r�� As String, val�s�� As Long)
Dim �Ȧs�� As Recordset
Dim i As Integer, ���� As Integer
Dim �r�� As String
Dim ����� As String, ���� As Integer, �r���� As String
Dim �s�� As Long, �r�� As String

If val�s�� <= 0 Then Exit Sub

�s�� = val�s��
�r�� = val�r��

If �t�Φr�� = "�_�v�j����p�f" Or �t�Φr�� = "�_�v�j���孫��" Then
    �p�f�˦r��.Index = "�s��"
    �p�f�˦r��.Seek "=", �s��
    �s�� = �p�f�˦r��.Fields("���ѽs��")
ElseIf ����|����(�t�Φr��) Then
    ���岧�g�r��.Index = "�s��"
    ���岧�g�r��.Seek "=", �s��
    If Not IsNull(���岧�g�r��.Fields("���ѽs��")) Then
        �s�� = ���岧�g�r��.Fields("���ѽs��")
    Else
        �����˦r��.Index = "�s��"
        �����˦r��.Seek "=", �s��
        �s�� = �����˦r��.Fields("���ѽs��")
    End If
    �˦r��.Index = "�s��"
    �˦r��.Seek "=", �s��
    �r�� = �˦r��.Fields("�r�X")
ElseIf ����|�Ұ���(�t�Φr��) Then
    �Ұ��岧�g�r��.Index = "�s��"
    �Ұ��岧�g�r��.Seek "=", �s��
    If Not IsNull(�Ұ��岧�g�r��.Fields("���ѽs��")) Then
        �s�� = �Ұ��岧�g�r��.Fields("���ѽs��")
    Else
        �Ұ����˦r��.Index = "�s��"
        �Ұ����˦r��.Seek "=", �s��
        �s�� = �Ұ����˦r��.Fields("���ѽs��")
    End If
    �˦r��.Index = "�s��"
    �˦r��.Seek "=", �s��
    �r�� = �˦r��.Fields("�r�X")
ElseIf ����|���t��r(�t�Φr��) Then
    ���t��r���g�r��.Index = "�s��"
    ���t��r���g�r��.Seek "=", �s��
    If Not IsNull(���t��r���g�r��.Fields("���ѽs��")) Then
        �s�� = ���t��r���g�r��.Fields("���ѽs��")
    Else
        ���t��r�˦r��.Index = "�s��"
        ���t��r�˦r��.Seek "=", �s��
        �s�� = ���t��r�˦r��.Fields("���ѽs��")
    End If
    �˦r��.Index = "�s��"
    �˦r��.Seek "=", �s��
    �r�� = �˦r��.Fields("�r�X")
End If

Set �˦r�� = �����˦r��
�˦r��.Index = "�s��"
�˦r��.Seek "=", �s��
�r�� = �˦r��.Fields("�r�X")

���w�ťխ�
'���� = 4
'����� = ""

'If Len(�r��) = 2 And �O�_���զr�Ÿ�(Mid(�r��, 1, 1), 4, 11) > 0 Then
'   ����� = �r��
'   ���� = 5
'   �˦r��.Index = "�c�r��"
'   �˦r��.Seek "=", ����, �����
'ElseIf �s�� <= 0 Then
'   �˦r��.Index = "�r��"
'   �˦r��.Seek "=", �r��
'Else
'   �˦r��.Index = "�s��"
'   �˦r��.Seek "=", �s��
'End If
   
'If Not �˦r��.NoMatch Then
'   If Not IsNull(�˦r��.Fields("�s���Ÿ�")) Then ���� = �˦r��.Fields("�s���Ÿ�")
'   If Not IsNull(�˦r��.Fields("�����")) Then ����� = �˦r��.Fields("�����")
'Else
'   ���� = 0
'   ����� = ""
'End If
   
'If Len(�r��) = 1 Or �O�_���զr�Ÿ�(Left$(�r��, 1), 1, 14) <> 0 Then
    If Not �˦r��.NoMatch Then
'      Do While Not �˦r��.EOF
'         If �˦r��.Fields("�s���Ÿ�") = ���� Or (IsNull(�˦r��.Fields("�s���Ÿ�")) And ���� = 4) Then Exit Do
'         �˦r��.MoveNext
'      Loop
      
'      If �˦r��.EOF Then
'         mdi�~�r�r��.txt�c�r��.Text = �����
'         Exit Sub
'      End If
      
      �t�νs�� = �˦r��.Fields("�s��")
      
      If �˦r��.Fields("�r��") = 0 Then
         �r���� = �{�Φr��
      Else
         �r���� = �r��}�C(�˦r��.Fields("�r��"))
      End If
      mdi�~�r�r��.txt�r��.FontName = �r����
      
      If Not IsNull(�˦r��.Fields("�s��")) Then mdi�~�r�r��.txt�s��.Text = �˦r��.Fields("�s��")
      If Not IsNull(�˦r��.Fields("�r��")) Then mdi�~�r�r��.txt�~�r��.Text = �˦r��.Fields("�r��")

      If Not IsNull(�˦r��.Fields("�r��")) Then
         mdi�~�r�r��.txt�r��.Text = �˦r��.Fields("�r��")
      Else
         If Not IsNull(�˦r��.Fields("�r�X")) Then
            mdi�~�r�r��.txt�r��.Text = �˦r��.Fields("�r�X")
         Else
            mdi�~�r�r��.txt�r��.Text = "��"
         End If
      End If
      
      If Not IsNull(�˦r��.Fields("���e")) Then mdi�~�r�r��.txt�`���e.Text = �˦r��.Fields("���e")
      If Not IsNull(�˦r��.Fields("����")) Then mdi�~�r�r��.txt����.Text = �M�䳡��(�˦r��.Fields("����"))
      If Not IsNull(�˦r��.Fields("�������e")) Then mdi�~�r�r��.txt�����������e.Text = �˦r��.Fields("�������e")
      If Not IsNull(�˦r��.Fields("�`��")) Then mdi�~�r�r��.txt�`��.Text = �M��`��(�˦r��.Fields("�`��"))
      If Not IsNull(�˦r��.Fields("BIG5")) Then mdi�~�r�r��.txt���X.Text = �˦r��.Fields("BIG5")
      If Not IsNull(�˦r��.Fields("�ܾe")) Then mdi�~�r�r��.txt�ܾe�X.Text = �ഫ�^���ܾe(�˦r��.Fields("�ܾe"))
      If Not IsNull(�˦r��.Fields("���F�~�y�j�r��")) Then
         mdi�~�r�r��.txt�U��.Text = �˦r��.Fields("���F�~�y�j�r��")
      Else
         mdi�~�r�r��.txt�U��.Text = ""
      End If
      
      If Not IsNull(�˦r��.Fields("�զr�r��")) Then
         mdi�~�r�r��.txt�զr�r��.Text = �˦r��.Fields("�զr�r��")
      Else
         mdi�~�r�r��.txt�զr�r��.Text = ""
      End If
      
      If Not IsNull(�˦r��.Fields("�զr�r�ƤG")) Then
         mdi�~�r�r��.txt�զr�r�Ƨt���g.Text = �˦r��.Fields("�զr�r�ƤG")
      Else
         mdi�~�r�r��.txt�զr�r�Ƨt���g.Text = ""
      End If
   Else
      mdi�~�r�r��.txt�c�r��.Text = �����
   End If
'Else
'   mdi�~�r�r��.txt�c�r��.Text = �r��
'End If

End Sub

Public Sub ���w�ťխ�()
mdi�~�r�r��.txt�~�r��.Text = ""
mdi�~�r�r��.txt�r��.Text = ""
mdi�~�r�r��.txt����.Text = ""
mdi�~�r�r��.txt�j�~�r.Text = ""
mdi�~�r�r��.txt�`���e.Text = ""
mdi�~�r�r��.txt����.Text = ""
mdi�~�r�r��.txt�����������e.Text = ""
mdi�~�r�r��.txt�`��.Text = ""
mdi�~�r�r��.txt���X.Text = ""
mdi�~�r�r��.txt�ܾe�X.Text = ""
mdi�~�r�r��.txt�c�r��.Text = ""
mdi�~�r�r��.txt�U��.Text = ""
mdi�~�r�r��.txt�զr�r��.Text = ""
mdi�~�r�r��.txt�զr�r�Ƨt���g.Text = ""
�t�νs�� = -999999
End Sub

Public Sub �^���c�r��(�t�Φr�� As String, val�r�� As String, val�s�� As Long)

Dim ����� As String, ���� As Integer
Dim �s�� As Long, �r�� As String, �r�Y As String, ����X As String


�s�� = val�s��
�r�� = val�r��

If �s�� <= 0 Then Exit Sub

�r�Y = ""
����X = ""
�ƻs����X = False

If �t�Φr�� = "�_�v�j����p�f" Or �t�Φr�� = "�_�v�j���孫��" Then
    �p�f�˦r��.Index = "�s��"
    �p�f�˦r��.Seek "=", �s��
    �s�� = �p�f�˦r��.Fields("���ѽs��")
    mdi�~�r�r��.txt���� = �p�f�˦r��.Fields("�r��")
    mdi�~�r�r��.txt�j�~�r = �p�f�˦r��.Fields("�r�X")
    mdi�~�r�r��.txt�j�~�r.FontName = �r��}�C(�p�f�˦r��.Fields("�r��"))
    If �p�f�˦r��.Fields("���X") = 1 Then
        ����X = �p�f�˦r��.Fields("�r��")
    Else
        ����X = �p�f�˦r��.Fields("�r��") & ";" & �p�f�˦r��.Fields("���X")
    End If
    �ƻs����X = True
    
    �˦r��.Index = "�s��"
    �˦r��.Seek "=", �s��
    �r�� = �˦r��.Fields("�r�X")
ElseIf ����|����(�t�Φr��) Then
    ���岧�g�r��.Index = "�s��"
    ���岧�g�r��.Seek "=", �s��
    If Not IsNull(���岧�g�r��.Fields("���ѽs��")) Then
        �s�� = ���岧�g�r��.Fields("���ѽs��")
        mdi�~�r�r��.txt���� = ���岧�g�r��.Fields("�r��")
        mdi�~�r�r��.txt�j�~�r = ���岧�g�r��.Fields("�r�X")
        mdi�~�r�r��.txt�j�~�r.FontName = �r��}�C(���岧�g�r��.Fields("�r��"))
        If mdi�~�r�r��.mnu_��ܭ���X.Checked = True And ((���岧�g�r��.Fields("�W�u") > 0 And ���岧�g�r��.Fields("�W�u") < 6)) And ���岧�g�r��.Fields("�X��") = 0 Then
            If ���岧�g�r��.Fields("���X") = 1 Then
                ����X = "����" & ���岧�g�r��.Fields("����")
            Else
                ����X = "����" & ���岧�g�r��.Fields("����") & ";" & ���岧�g�r��.Fields("���X")
            End If
            �ƻs����X = True
        End If
    Else
        �����˦r��.Index = "�s��"
        �����˦r��.Seek "=", �s��
        �s�� = �����˦r��.Fields("���ѽs��")
        mdi�~�r�r��.txt���� = �����˦r��.Fields("�r��")
        mdi�~�r�r��.txt�j�~�r = �����˦r��.Fields("�r�X")
        mdi�~�r�r��.txt�j�~�r.FontName = �r��}�C(�����˦r��.Fields("�r��"))
    End If
    �˦r��.Index = "�s��"
    �˦r��.Seek "=", �s��
    �r�� = �˦r��.Fields("�r�X")
ElseIf ����|�Ұ���(�t�Φr��) Then
    �Ұ��岧�g�r��.Index = "�s��"
    �Ұ��岧�g�r��.Seek "=", �s��
    If Not IsNull(�Ұ��岧�g�r��.Fields("���ѽs��")) Then
        �s�� = �Ұ��岧�g�r��.Fields("���ѽs��")
        mdi�~�r�r��.txt���� = �Ұ��岧�g�r��.Fields("�r��")
        mdi�~�r�r��.txt�j�~�r = �Ұ��岧�g�r��.Fields("�r�X")
        mdi�~�r�r��.txt�j�~�r.FontName = �r��}�C(�Ұ��岧�g�r��.Fields("�r��"))
        If mdi�~�r�r��.mnu_��ܭ���X.Checked = True And ((�Ұ��岧�g�r��.Fields("�W�u") > 0 And �Ұ��岧�g�r��.Fields("�W�u") < 6)) And �Ұ��岧�g�r��.Fields("�X��") = 0 And Not IsNull(�Ұ��岧�g�r��.Fields("�X�B")) Then
            ����X = �Ұ��岧�g�r��.Fields("�X�B")
            If IsNumeric(Left(����X, 1)) Then ����X = "�X��" & ����X
            If �Ұ��岧�g�r��.Fields("���X") > 1 Then
                ����X = ����X & ";" & �Ұ��岧�g�r��.Fields("���X")
            End If
            �ƻs����X = True
        End If
    Else
        �Ұ����˦r��.Index = "�s��"
        �Ұ����˦r��.Seek "=", �s��
        �s�� = �Ұ����˦r��.Fields("���ѽs��")
        mdi�~�r�r��.txt���� = �Ұ����˦r��.Fields("�r��")
        mdi�~�r�r��.txt�j�~�r = �Ұ����˦r��.Fields("�r�X")
        mdi�~�r�r��.txt�j�~�r.FontName = �r��}�C(�Ұ����˦r��.Fields("�r��"))
    End If
    �˦r��.Index = "�s��"
    �˦r��.Seek "=", �s��
    �r�� = �˦r��.Fields("�r�X")
ElseIf ����|���t��r(�t�Φr��) Then
    ���t��r���g�r��.Index = "�s��"
    ���t��r���g�r��.Seek "=", �s��
    If Not IsNull(���t��r���g�r��.Fields("���ѽs��")) Then
        �s�� = ���t��r���g�r��.Fields("���ѽs��")
        mdi�~�r�r��.txt���� = ���t��r���g�r��.Fields("�r��")
        mdi�~�r�r��.txt�j�~�r = ���t��r���g�r��.Fields("�r�X")
        mdi�~�r�r��.txt�j�~�r.FontName = �r��}�C(���t��r���g�r��.Fields("�r��"))
        If mdi�~�r�r��.mnu_��ܭ���X.Checked = True And ((���t��r���g�r��.Fields("�W�u") > 0 And ���t��r���g�r��.Fields("�W�u") < 6)) And ���t��r���g�r��.Fields("�X��") = 0 And Not IsNull(���t��r���g�r��.Fields("�X�B")) Then
            ����X = ���t��r���g�r��.Fields("�X�B")
            If ���t��r���g�r��.Fields("���X") > 1 Then
                ����X = ����X & ";" & ���t��r���g�r��.Fields("���X")
            End If
            �ƻs����X = True
        End If
    Else
        ���t��r�˦r��.Index = "�s��"
        ���t��r�˦r��.Seek "=", �s��
        �s�� = ���t��r�˦r��.Fields("���ѽs��")
        mdi�~�r�r��.txt���� = ���t��r�˦r��.Fields("�r��")
        mdi�~�r�r��.txt�j�~�r = ���t��r�˦r��.Fields("�r�X")
        mdi�~�r�r��.txt�j�~�r.FontName = �r��}�C(���t��r�˦r��.Fields("�r��"))
    End If
    �˦r��.Index = "�s��"
    �˦r��.Seek "=", �s��
    �r�� = �˦r��.Fields("�r�X")
End If

���� = 4
����� = ""

If Len(�r��) = 2 And �O�_���զr�Ÿ�(Mid(�r��, 1, 1), 4, 11) > 0 Then
   ����� = �r��
   ���� = 5
   �˦r��.Index = "�c�r��"
   �˦r��.Seek "=", ����, �����
ElseIf �s�� = 0 Then
   �˦r��.Index = "�r��"
   �˦r��.Seek "=", �r��
Else
   �˦r��.Index = "�s��"
   �˦r��.Seek "=", �s��
End If
   
If Not �˦r��.NoMatch Then
   If Not �˦r��.EOF Then
      If Not IsNull(�˦r��.Fields("�r��")) Then �r�Y = �˦r��.Fields("�r��")
      If Not IsNull(�˦r��.Fields("�s���Ÿ�")) Then ���� = �˦r��.Fields("�s���Ÿ�")
      If Not IsNull(�˦r��.Fields("�����")) Then ����� = �˦r��.Fields("�����")
   End If
Else
   ���� = 0
   ����� = ""
End If

If Not �˦r��.NoMatch Then
   Do While Not �˦r��.EOF
      If (�˦r��.Fields("�s���Ÿ�") = ����) Or (IsNull(�˦r��.Fields("�s���Ÿ�")) And ���� = 4) Then Exit Do
      �˦r��.MoveNext
   Loop
   �즲�r�� = �˦r��.Fields("�����")
   
   If ����X <> "" Then
        mdi�~�r�r��.txt�c�r��.Text = �M�䭷��X(�r�Y, �����, ����X)
   Else
        mdi�~�r�r��.txt�c�r��.Text = �M��զr��(�˦r��.Fields("�s���Ÿ�"), �˦r��.Fields("�����"))
   End If
End If

End Sub

Public Function �զr�Ÿ�����(�r�� As String) As Integer

�զr�Ÿ����� = 0

Select Case �r��
       Case �զr�Ÿ��}�C(4)
            �զr�Ÿ����� = 2
       Case �զr�Ÿ��}�C(5)
            �զr�Ÿ����� = 2
       Case �զr�Ÿ��}�C(6)
            �զr�Ÿ����� = 3
       Case �զr�Ÿ��}�C(7)
            �զr�Ÿ����� = 3
       Case �զr�Ÿ��}�C(8)
            �զr�Ÿ����� = 3
       Case �զr�Ÿ��}�C(9)
            �զr�Ÿ����� = 4
       Case �զr�Ÿ��}�C(10)
            �զr�Ÿ����� = 4
       Case �զr�Ÿ��}�C(11)
            �զr�Ÿ����� = 4
End Select

End Function

Public Function �M�䳡��(�s�� As Integer) As String
�M�䳡�� = ""
�d������.Index = "�s��"
�d������.Seek "=", �s��
If Not �d������.NoMatch Then
   �M�䳡�� = �d������.Fields("�r��")
End If
End Function

Public Function �M�仡�峡��(�s�� As Integer) As String

���峡��.Index = "�s��"
���峡��.Seek "=", �s��

If Not ���峡��.NoMatch Then
   �M�仡�峡�� = ���峡��.Fields("�r��")
Else
   �M�仡�峡�� = ""
End If

End Function

Public Function �M��`��(�`�� As String) As String

�M��`�� = Left(�`��, Len(�`��) - 1)

Select Case Right$(�`��, 1)
       Case 1
            �M��`�� = �M��`��
       Case 2
            �M��`�� = �M��`�� & "��"
       Case 3
            �M��`�� = �M��`�� & "��"
       Case 4
            �M��`�� = �M��`�� & "��"
       Case 5
            �M��`�� = "��" & �M��`��
End Select

End Function

Public Function �إ�SQL(txt�c�r�� As String, sql As Integer) As String
Dim �Ÿ� As String
Dim �զr�� As String
Dim i As Integer

�إ�SQL = ""
   
�Ÿ� = ""
For i = 1 To Len(txt�c�r��)
    If Mid(txt�c�r��, i, 1) <> "*" Then
       �զr�� = �զr�� + Mid(txt�c�r��, i, 1) + �Ÿ�
    End If
Next i

txt�c�r�� = �զr��
   
If sql = 1 Then
   �Ÿ� = "*"
   �զr�� = �Ÿ�
   For i = 1 To Len(txt�c�r��)
       �զr�� = �զr�� + Mid(txt�c�r��, i, 1) + �Ÿ�
   Next i
End If
�إ�SQL = �զr��

End Function

Public Function �r�ڱƧ�(�r�ڧ� As String) As String
Dim �r�ڪ� As Recordset, �r�ڲ� As String
Dim �r�νs�����|(30) As Integer, �r�ΰ��|(30) As String
Dim Maxlen As Integer, i As Integer, j As Integer
Dim temp1 As Integer, temp2 As String

Set �r�ڪ� = �t�θ�Ʈw.OpenRecordset("�r��")

�r�ڪ�.Index = "�r��"
Maxlen = Len(�r�ڧ�)

'�M��s��
For i = 1 To Maxlen
    �r�ΰ��|(i - 1) = Mid(�r�ڧ�, i, 1)
    �r�ڪ�.Seek "=", Mid(�r�ڧ�, i, 1)
    If Not �r�ڪ�.NoMatch Then
       �r�νs�����|(i - 1) = �r�ڪ�.Fields("�s��")
    Else
       '�䤣��h��-1
       �r�νs�����|(i - 1) = -1
    End If
Next i
    
'�w�j�Ƨ�
For i = 0 To Maxlen - 1
    For j = i + 1 To Maxlen - 1
        If �r�νs�����|(i) > �r�νs�����|(j) Then
           temp1 = �r�νs�����|(i)
           �r�νs�����|(i) = �r�νs�����|(j)
           �r�νs�����|(j) = temp1
           temp2 = �r�ΰ��|(i)
           �r�ΰ��|(i) = �r�ΰ��|(j)
           �r�ΰ��|(j) = temp2
         End If
     Next j
Next i
    
'�N�r�ΰ��|�հ_��
�r�ڲ� = ""
For i = 0 To Maxlen - 1
    �r�ڲ� = �r�ڲ� & �r�ΰ��|(i)
Next i

�r�ڱƧ� = �r�ڲ�

�r�ڪ�.Close

End Function


Public Function �ഫ�r��(�r�� As String) As String

If ��ܦr�� = "�ө���" Then
    If �r�� = "�ө���" Then
        �ഫ�r�� = "�з���"
    Else
        �ഫ�r�� = �r��
    End If
Else
    �ഫ�r�� = �r��
End If
    


End Function



Public Function �ഫ��ܦr��(ByVal �r�� As String) As String

Dim ifound As Integer, rstr As String

If ��ܦr�� = "�ө���" Then
    ifound = InStr(1, �r��, "�з���")
    If ifound > 0 Then
        rstr = �r��
        rstr = Left(rstr, ifound - 1) & "�ө���" & Right(rstr, Len(rstr) - ifound - 2)
        �ഫ��ܦr�� = rstr
    Else
        �ഫ��ܦr�� = �r��
    End If
Else
    �ഫ��ܦr�� = �r��
End If
    
End Function

Public Function ������ܦr��(ByVal �r�� As String) As String

Dim ifound As Integer, rstr As String

If ��ܦr�� = "�ө���" Then
    ifound = InStr(1, �r��, "�з���")
    If ifound > 0 Then
        rstr = �r��
        rstr = Left(rstr, ifound - 1) & "�ө���" & Right(rstr, Len(rstr) - ifound - 2)
        ������ܦr�� = rstr
    Else
        ������ܦr�� = �r��
    End If
Else
    ifound = InStr(1, �r��, "�ө���")
    If ifound > 0 Then
        rstr = �r��
        rstr = Left(rstr, ifound - 1) & "�з���" & Right(rstr, Len(rstr) - ifound - 2)
        ������ܦr�� = rstr
    Else
        ������ܦr�� = �r��
    End If
End If
    
End Function


Public Function �ഫRTF�ʦr(�t�ʦr�r�� As String, ��ܦr�� As String) As String

Dim RTF�r�� As String
Dim �_�l��m As Long, �׵���m As Long, �r����m As Long
Dim �ثe�r�� As String, �W�@�r�� As String, �U�@�r�� As String
Dim �s���Ÿ� As Integer, ����� As String, �c�r�� As String
Dim �r�� As Integer, �r�X As String
Dim i As Integer, wk As String

�r����m = 1
�׵���m = 0

Do While �r����m <= Len(�t�ʦr�r��)

�ثe�r�� = Mid(�t�ʦr�r��, �r����m, 1)

If �ثe�r�� >= "��" And �ثe�r�� <= "��" Then
     
    '�P�_�r�ꤺ�e�κA
   
    If (�ثe�r�� <> "��") Then
   
        If (�ثe�r�� >= "��" And �ثe�r�� <= "��") Then
            �_�l��m = �r����m - 1
        Else
            �_�l��m = �r����m
        End If
      
        Do
            �W�@�r�� = �ثe�r��
            �r����m = �r����m + 1
            If �r����m <= Len(�t�ʦr�r��) Then
                �ثe�r�� = Mid(�t�ʦr�r��, �r����m, 1)
            Else
                �ثe�r�� = Chr(13)
            End If
        Loop Until (�W�@�r�� < "��" Or �W�@�r�� > "��") And (�ثe�r�� < "��" Or �ثe�r�� > "��")
        
        �r����m = �r����m - 1
        �׵���m = �r����m
      
        �c�r�� = Mid(�t�ʦr�r��, �_�l��m, �׵���m - �_�l��m + 1)
        
        ����� = ""
        
        �s���Ÿ� = 5
        For i = 1 To Len(�c�r��)
            wk = Mid(�c�r��, i, 1)
            If (wk < "��" Or wk > "��") Then
                ����� = ����� + wk
            Else
                If wk = "��" Then
                    �s���Ÿ� = 1
                ElseIf wk = "��" Then
                    �s���Ÿ� = 2
                Else
                    �s���Ÿ� = 3
                End If
            End If
        Next i
        
    Else
    
        �_�l��m = �r����m
        ����� = ""
        
        Do
            �r����m = �r����m + 1
            If �r����m <= Len(�t�ʦr�r��) Then
                �ثe�r�� = Mid(�t�ʦr�r��, �r����m, 1)
            Else
                �ثe�r�� = Chr(13)
            End If
        Loop Until (�ثe�r�� = "��") Or (�ثe�r�� = Chr(13))
        
        �׵���m = �r����m
        
        If �ثe�r�� = "��" Then
            �c�r�� = Mid(�t�ʦr�r��, �_�l��m, �׵���m - �_�l��m + 1)
            ����� = Mid(�c�r��, 2, Len(�c�r��) - 2)
            �s���Ÿ� = 4
        End If
        
    End If
   
    �����˦r��.Index = "�c�r��"
    �����˦r��.Seek "=", �s���Ÿ�, �����
    
    If Not �����˦r��.NoMatch Then
    
        �r�� = �����˦r��.Fields("�r��")
        �r�X = �����˦r��.Fields("�r�X")
    
        If ��ܦr�� = "�з���" Then
            RTF�r�� = RTF�r�� & "{\f" & �r�� & �ഫRTF�r��(�r�X) & "}"
        Else
            RTF�r�� = RTF�r�� & "{\f" & �r�� + 16 & �ഫRTF�r��(�r�X) & "}"
        End If
    Else
        RTF�r�� = RTF�r�� & "{\f0\'a1\'b4}" '��
    End If
    
Else
    
    If �r����m < Len(�t�ʦr�r��) Then
        �U�@�r�� = Mid(�t�ʦr�r��, �r����m + 1, 1)
        If (�U�@�r�� >= "��" And �U�@�r�� <= "��") Then GoTo �B�z�U�@�ӯʦr
    End If
    
    If �r����m - �׵���m = 1 Then
        If ��ܦr�� = "�з���" Then
            RTF�r�� = RTF�r�� & "{\f0" & �ഫRTF�r��(�ثe�r��) & "}"
        Else
            RTF�r�� = RTF�r�� & "{\f16" & �ഫRTF�r��(�ثe�r��) & "}"
        End If
    Else
        RTF�r�� = Left(RTF�r��, Len(RTF�r��) - 1) & �ഫRTF�r��(�ثe�r��) & "}"
    End If

End If

�B�z�U�@�ӯʦr:

�r����m = �r����m + 1

Loop

If ��ܦr�� = "�з���" Then
    �ഫRTF�ʦr = "{" & rtf_version & character_set & �з���r���� & RTF�r�� & "}"
Else
    �ഫRTF�ʦr = "{" & rtf_version & character_set & �ө���r���� & RTF�r�� & "}"
End If

End Function

Public Function �ഫRTF�r��(�r�� As String) As String

Dim �r�X As String

�r�X = CStr(Hex(Asc(�r��)))

If Len(�r�X) = 4 Then
    �ഫRTF�r�� = "\'" & LCase(Left(�r�X, 2)) & "\'" & LCase(Right(�r�X, 2))
ElseIf Len(�r�X) = 2 Then
     �ഫRTF�r�� = "\'" & LCase(Left(�r�X, 2))
Else
    �ഫRTF�r�� = ""
End If

End Function

Public Function ����|����(�r��) As Boolean

If InStr(1, �r��, "����|����") > 0 Then
    ����|���� = True
Else
    ����|���� = False
End If

End Function

Public Function ����|�Ұ���(�r��) As Boolean

If InStr(1, �r��, "����|�Ұ���") > 0 Then
    ����|�Ұ��� = True
Else
    ����|�Ұ��� = False
End If

End Function

Public Function ����|���t��r(�r��) As Boolean

If InStr(1, �r��, "����|���t²����r") > 0 Then
    ����|���t��r = True
Else
    ����|���t��r = False
End If

End Function


Public Function ����|����(�r��) As Boolean

If InStr(1, �r��, "�з���") > 0 Then
    ����|���� = True
ElseIf InStr(1, �r��, "�ө���") > 0 Then
    ����|���� = True
Else
    ����|���� = True = False
End If

End Function


Public Sub �X�B��r��(�X�B)

�X�B���Ұ��� = False
�X�B���Ұ���X�� = False
�X�B������ = False
�X�B���p�f = False
�X�B������r = False
�X�B�����ǰt = False

If InStr(1, �X�B, "�X��") > 0 Then
    �X�B���Ұ��� = True
    �X�B���Ұ���X�� = True
    If Len(�X�B) > 2 Then �X�B�����ǰt = True
ElseIf InStr(1, �X�B, "��") > 0 Then
    �X�B���Ұ��� = True
ElseIf InStr(1, �X�B, "�^") > 0 Then
    �X�B���Ұ��� = True
    If Len(�X�B) > 1 Then �X�B�����ǰt = True
ElseIf InStr(1, �X�B, "�h") > 0 Then
    �X�B���Ұ��� = True
    If Len(�X�B) > 1 Then �X�B�����ǰt = True
ElseIf InStr(1, �X�B, "����") > 0 Then
    �X�B������ = True
    If Len(�X�B) > 2 Then �X�B�����ǰt = True
ElseIf InStr(1, �X�B, "����") > 0 Then
    �X�B���p�f = True
Else
    �X�B������r = True
End If

End Sub
