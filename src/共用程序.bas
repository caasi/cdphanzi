Attribute VB_Name = "¤½¥Îµ{§Ç"
Option Explicit
Public Const rtf_version = "\rtf1", character_set = "\ansi\deflang1033"
Public Const ¼Ğ·¢Åé = "{\f0\fscript\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9;}"
Public Const ¼Ğ·¢Åé¥~¦r¶°¤@ = "{\f1\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'40;}"
Public Const ¼Ğ·¢Åé¥~¦r¶°¤G = "{\f2\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'47;}"
Public Const ¼Ğ·¢Åé¥~¦r¶°¤T = "{\f3\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'54;}"
Public Const ¼Ğ·¢Åé¥~¦r¶°¥| = "{\f4\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a5\'7c;}"
Public Const ¼Ğ·¢Åé¥~¦r¶°¤­ = "{\f5\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'ad;}"
Public Const ¼Ğ·¢Åé¥~¦r¶°¤» = "{\f6\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'bb;}"
Public Const ¼Ğ·¢Åé¥~¦r¶°¤C = "{\f7\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'43;}"
Public Const ¼Ğ·¢Åé¥~¦r¶°¤K = "{\f8\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'4b;}"
Public Const ¼Ğ·¢Åé¥~¦r¶°¤E = "{\f9\fmodern\fcharset136\'bc\'d0\'b7\'a2\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'45;}"
Public Const ²Ó©úÅé = "{\f16\fscript\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9;}"
Public Const ²Ó©úÅé¥~¦r¶°¤@ = "{\f17\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'40;}"
Public Const ²Ó©úÅé¥~¦r¶°¤G = "{\f18\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'47;}"
Public Const ²Ó©úÅé¥~¦r¶°¤T = "{\f19\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'54;}"
Public Const ²Ó©úÅé¥~¦r¶°¥| = "{\f20\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a5\'7c;}"
Public Const ²Ó©úÅé¥~¦r¶°¤­ = "{\f21\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'ad;}"
Public Const ²Ó©úÅé¥~¦r¶°¤» = "{\f22\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'bb;}"
Public Const ²Ó©úÅé¥~¦r¶°¤C = "{\f23\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'43;}"
Public Const ²Ó©úÅé¥~¦r¶°¤K = "{\f24\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'4b;}"
Public Const ²Ó©úÅé¥~¦r¶°¤E = "{\f25\fmodern\fcharset136\'b2\'d3\'a9\'fa\'c5\'e9\'a5\'7e\'a6\'72\'b6\'b0\'a4\'45;}"
Public Const ¼Ğ·¢Åé¦r«¬ = ¼Ğ·¢Åé & ¼Ğ·¢Åé¥~¦r¶°¤@ & ¼Ğ·¢Åé¥~¦r¶°¤G & ¼Ğ·¢Åé¥~¦r¶°¤T & ¼Ğ·¢Åé¥~¦r¶°¥| & ¼Ğ·¢Åé¥~¦r¶°¤­ & ¼Ğ·¢Åé¥~¦r¶°¤» & ¼Ğ·¢Åé¥~¦r¶°¤C & ¼Ğ·¢Åé¥~¦r¶°¤K & ¼Ğ·¢Åé¥~¦r¶°¤E
Public Const ¼Ğ·¢Åé¦r«¬ªí = "{\fonttbl" & ¼Ğ·¢Åé¦r«¬ & "}"
Public Const ²Ó©úÅé¦r«¬ = ²Ó©úÅé & ²Ó©úÅé¥~¦r¶°¤@ & ²Ó©úÅé¥~¦r¶°¤G & ²Ó©úÅé¥~¦r¶°¤T & ²Ó©úÅé¥~¦r¶°¥| & ²Ó©úÅé¥~¦r¶°¤­ & ²Ó©úÅé¥~¦r¶°¤» & ²Ó©úÅé¥~¦r¶°¤C & ²Ó©úÅé¥~¦r¶°¤K & ²Ó©úÅé¥~¦r¶°¤E
Public Const ²Ó©úÅé¦r«¬ªí = "{\fonttbl" & ²Ó©úÅé¦r«¬ & "}"

Public Function ³¡¥ó¥~¦rSQL³¯­z¦¡(¸ê®Æªí As String, µ§µe As Integer, ­ºµ§ As Integer) As String
Dim ¿ï¾Ü¶µ¥Ø As String
Dim ±ø¥ó As String

±ø¥ó = ""
If ¸ê®Æªí = "ºc¦r²Å¸¹" Then
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from ²Å¸¹ order by ½s¸¹ "
ElseIf ¸ê®Æªí = "¹Ï§Î¤å¦r" Then
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from ¹Ï§Î¤å¦r order by ½s¸¹ "
ElseIf ¸ê®Æªí = "¤K¨ö" Then
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from ¤K¨ö order by ½s¸¹ "
ElseIf ¸ê®Æªí = "Â²Ã|" Then
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from Â²Ã| order by ½s¸¹ "
ElseIf ¸ê®Æªí = "±dº³¦r¨å³¡­º" Then
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from ±dº³³¡­º "
ElseIf ¸ê®Æªí = "»¡¤å¸Ñ¦r³¡­º" Then
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from »¡¤å³¡­º "
ElseIf ¸ê®Æªí = "Big5¦r®Ú" Then
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from ¦r®Ú where Big5=1 "
ElseIf ¸ê®Æªí = "Big5¤ÎÂ²¤Æ¦r¦r®Ú" Then
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from ¦r®Ú where (Big5=1 or Â²¤Æ¦r=1) "
ElseIf ¸ê®Æªí = "¦r®Ú" Then
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from ¦r®Ú "
ElseIf ¸ê®Æªí = "¤p½f¿WÅé¦r" Then
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from ¤p½f¿WÅé¦r where ½s¸¹<9999999"
ElseIf ¸ê®Æªí = "ª÷¤å¦r®Ú" Then
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from ª÷¤å¦r®Ú where ½s¸¹<9999999"
ElseIf ¸ê®Æªí = "¥Ò°©¤å¦r®Ú" Then
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from ¥Ò°©¤å¦r®Ú where ½s¸¹<9999999"
ElseIf ¸ê®Æªí = "·¡¨tÂ²©­¤å¦r¦r®Ú" Then
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from ·¡¨tÂ²©­¤å¦r¦r®Ú where ½s¸¹<9999999"
Else
   ¿ï¾Ü¶µ¥Ø = "select ¦r§Î from ³¡¥ó¥~¦r where ½s¸¹<9999999"
End If

If ¸ê®Æªí <> "ºc¦r²Å¸¹" And ¸ê®Æªí <> "¤K¨ö" And ¸ê®Æªí <> "Â²Ã|" And ¸ê®Æªí <> "¹Ï§Î¤å¦r" Then
   If µ§µe > 0 And ­ºµ§ > 0 Then
      If ¸ê®Æªí = "±dº³¦r¨å³¡­º" Or ¸ê®Æªí = "»¡¤å¸Ñ¦r³¡­º" Or ¸ê®Æªí = "¦r®Ú" Then
         ±ø¥ó = "where µ§µe = " & µ§µe & " and ­ºµ§ = " & ­ºµ§ & " order by ½s¸¹"
      Else
         ±ø¥ó = "and µ§µe = " & µ§µe & " and ­ºµ§ = " & ­ºµ§ & " order by ½s¸¹"
      End If
   ElseIf µ§µe > 0 Then
      If ¸ê®Æªí = "±dº³¦r¨å³¡­º" Or ¸ê®Æªí = "»¡¤å¸Ñ¦r³¡­º" Or ¸ê®Æªí = "¦r®Ú" Then
         ±ø¥ó = "where µ§µe = " & µ§µe & " order by ­ºµ§,½s¸¹"
      Else
         ±ø¥ó = "and µ§µe = " & µ§µe & " order by ­ºµ§,½s¸¹"
      End If
   ElseIf ­ºµ§ > 0 Then
      If ¸ê®Æªí = "±dº³¦r¨å³¡­º" Or ¸ê®Æªí = "»¡¤å¸Ñ¦r³¡­º" Or ¸ê®Æªí = "¦r®Ú" Then
         ±ø¥ó = "where ­ºµ§ = " & ­ºµ§ & " order by µ§µe,½s¸¹"
      Else
         ±ø¥ó = "and ­ºµ§ = " & ­ºµ§ & " order by µ§µe,½s¸¹"
      End If
   Else
      ±ø¥ó = ""
   End If
End If

³¡¥ó¥~¦rSQL³¯­z¦¡ = ¿ï¾Ü¶µ¥Ø & ±ø¥ó


End Function

Public Function Âà´«­^¤å¨ì­Ü¾e(­^¤å½X As String) As String

Dim ­^¤å¦r¥À As String, ­Ü¾e¦r¥À As String
Dim i As Integer
Dim °}¦C As Variant

°}¦C = Array("¤é", "¤ë", "ª÷", "¤ì", "¤ô", "¤õ", "¤g", "¦Ë", "¤à", "¤Q", "¤j", "¤¤", "¤@", "¤}", "¤H", "¤ß", "¤â", "¤f", "¤r", "¤Ü", "¤s", "¤k", "¥Ğ", "Ãø", "¤R", "­«")
For i = 1 To Len(­^¤å½X)
    ­^¤å¦r¥À = Mid(­^¤å½X, i, 1)
    ­Ü¾e¦r¥À = °}¦C(Asc(­^¤å¦r¥À) - 65)
    Âà´«­^¤å¨ì­Ü¾e = Âà´«­^¤å¨ì­Ü¾e & ­Ü¾e¦r¥À
Next i

End Function


Public Function Âà´«­Ü¾e¨ì­^¤å(­Ü¾e½X As String) As String
Dim ­^¤å½X As String
Dim ­^¤å¦r¥À As String, ­Ü¾e¦r¥À As String
Dim i As Integer

­^¤å½X = ""
For i = 1 To Len(­Ü¾e½X)
    ­Ü¾e¦r¥À = Mid(­Ü¾e½X, i, 1)
    ­^¤å¦r¥À = ­Ü¾eÂà­^¤å(­Ü¾e¦r¥À)
    ­^¤å½X = ­^¤å½X & ­^¤å¦r¥À
Next i
End Function

Public Function ­Ü¾eÂà­^¤å(­Ü¾e¦r¥À As String)
Select Case ­Ü¾e¦r¥À
    Case "¤é"
        ­Ü¾eÂà­^¤å = "A"
    Case "¤ë"
        ­Ü¾eÂà­^¤å = "B"
    Case "ª÷"
        ­Ü¾eÂà­^¤å = "C"
    Case "¤ì"
        ­Ü¾eÂà­^¤å = "D"
    Case "¤ô"
        ­Ü¾eÂà­^¤å = "E"
    Case "¤õ"
        ­Ü¾eÂà­^¤å = "F"
    Case "¤g"
        ­Ü¾eÂà­^¤å = "G"
    Case "¦Ë"
        ­Ü¾eÂà­^¤å = "H"
    Case "¤à"
        ­Ü¾eÂà­^¤å = "I"
    Case "¤Q"
        ­Ü¾eÂà­^¤å = "J"
    Case "¤j"
        ­Ü¾eÂà­^¤å = "K"
    Case "¤¤"
        ­Ü¾eÂà­^¤å = "L"
    Case "¤@"
        ­Ü¾eÂà­^¤å = "M"
    Case "¤}"
        ­Ü¾eÂà­^¤å = "N"
    Case "¤H"
        ­Ü¾eÂà­^¤å = "O"
    Case "¤ß"
        ­Ü¾eÂà­^¤å = "P"
    Case "¤â"
        ­Ü¾eÂà­^¤å = "Q"
    Case "¤f"
        ­Ü¾eÂà­^¤å = "R"
    Case "¤r"
        ­Ü¾eÂà­^¤å = "S"
    Case "¤Ü"
        ­Ü¾eÂà­^¤å = "T"
    Case "¤s"
        ­Ü¾eÂà­^¤å = "U"
    Case "¤k"
        ­Ü¾eÂà­^¤å = "V"
    Case "¥Ğ"
        ­Ü¾eÂà­^¤å = "W"
    Case "Ãø"
        ­Ü¾eÂà­^¤å = "X"
    Case "¤R"
        ­Ü¾eÂà­^¤å = "Y"
    Case "­«"
        ­Ü¾eÂà­^¤å = "Z"
    Case Else
        ­Ü¾eÂà­^¤å = ­Ü¾e¦r¥À
 
End Select

End Function

Public Function ´M§ä²Õ¦r¦¡(¤À¸Ñ As Variant, ³¡¥ó§Ç As Variant) As String
Dim i As Integer, j As Integer
Dim ²Õ¦r¦¡ As String, ¼È¦s²Õ¦r¦¡ As String

If ¤À¸Ñ < 4 Then
   ²Õ¦r¦¡ = ""
   If ¤À¸Ñ <> 0 Then
      i = 1
      Do Until i > Len(³¡¥ó§Ç)
         ²Õ¦r¦¡ = ²Õ¦r¦¡ & Mid(³¡¥ó§Ç, i, 1)
         For j = 4 To 11
             If Mid(³¡¥ó§Ç, i, 1) = ²Õ¦r²Å¸¹°}¦C(j) Then
                i = i + 1
                ²Õ¦r¦¡ = ²Õ¦r¦¡ & Mid(³¡¥ó§Ç, i, 1)
                Exit For
             End If
         Next j
         ²Õ¦r¦¡ = ²Õ¦r¦¡ & ²Õ¦r²Å¸¹°}¦C(¤À¸Ñ)
         i = i + 1
      Loop
      ¼È¦s²Õ¦r¦¡ = Mid(²Õ¦r¦¡, 1, Len(²Õ¦r¦¡) - 1)
   Else
      ¼È¦s²Õ¦r¦¡ = ³¡¥ó§Ç
   End If
Else
   If ¬O§_¬°²Õ¦r²Å¸¹(Mid(³¡¥ó§Ç, 1, 1), 1, 14) > 0 And Len(³¡¥ó§Ç) <= 2 Then
      ¼È¦s²Õ¦r¦¡ = ³¡¥ó§Ç
   Else
      ¼È¦s²Õ¦r¦¡ = ²Õ¦r²Å¸¹°}¦C(12) & ³¡¥ó§Ç & ²Õ¦r²Å¸¹°}¦C(13)
   End If
End If

´M§ä²Õ¦r¦¡ = ¼È¦s²Õ¦r¦¡

End Function

Public Function ´M§ä­·®æ½X(¦rÀY As String, ³¡¥ó§Ç As String, ­·®æ½X As String) As String

If ¦rÀY <> "" Then
    ´M§ä­·®æ½X = "ü" & ¦rÀY & "Æ¡" & ­·®æ½X & "ı"
Else
    ´M§ä­·®æ½X = "ü" & ³¡¥ó§Ç & "Æ¡" & ­·®æ½X & "ı"
End If

End Function


Public Function ¬O§_¬°²Õ¦r²Å¸¹(¦r§Î As String, x As Integer, y As Integer) As Integer
Dim i As Integer

¬O§_¬°²Õ¦r²Å¸¹ = 0
For i = x To y
    If ¦r§Î = ²Õ¦r²Å¸¹°}¦C(i) Then
       ¬O§_¬°²Õ¦r²Å¸¹ = i
       Exit For
    End If
Next i

End Function

Public Function ¬O§_¬°¸U¥Î¦r¤¸(¦r§Î As String) As Boolean


¬O§_¬°¸U¥Î¦r¤¸ = InStr(1, "*?#[]!-", ¦r§Î) > 0

End Function

Public Sub Â^¨úÄİ©Ê(¨t²Î¦rÅé As String, val¦r§Î As String, val½s¸¹ As Long)
Dim ¼È¦sªí As Recordset
Dim i As Integer, ªø«× As Integer
Dim ¦r¦ê As String
Dim ³¡¥ó§Ç As String, ¤À¸Ñ As Integer, ¦r«¬ÀÉ As String
Dim ½s¸¹ As Long, ¦r§Î As String

If val½s¸¹ <= 0 Then Exit Sub

½s¸¹ = val½s¸¹
¦r§Î = val¦r§Î

If ¨t²Î¦rÅé = "¥_®v¤j»¡¤å¤p½f" Or ¨t²Î¦rÅé = "¥_®v¤j»¡¤å­«¤å" Then
    ¤p½fÀË¦rªí.Index = "½s¸¹"
    ¤p½fÀË¦rªí.Seek "=", ½s¸¹
    ½s¸¹ = ¤p½fÀË¦rªí.Fields("·¢®Ñ½s¸¹")
ElseIf ¤¤¬ã°|ª÷¤å(¨t²Î¦rÅé) Then
    ª÷¤å²§¼g¦rªí.Index = "½s¸¹"
    ª÷¤å²§¼g¦rªí.Seek "=", ½s¸¹
    If Not IsNull(ª÷¤å²§¼g¦rªí.Fields("·¢®Ñ½s¸¹")) Then
        ½s¸¹ = ª÷¤å²§¼g¦rªí.Fields("·¢®Ñ½s¸¹")
    Else
        ª÷¤åÀË¦rªí.Index = "½s¸¹"
        ª÷¤åÀË¦rªí.Seek "=", ½s¸¹
        ½s¸¹ = ª÷¤åÀË¦rªí.Fields("·¢®Ñ½s¸¹")
    End If
    ÀË¦rªí.Index = "½s¸¹"
    ÀË¦rªí.Seek "=", ½s¸¹
    ¦r§Î = ÀË¦rªí.Fields("¦r½X")
ElseIf ¤¤¬ã°|¥Ò°©¤å(¨t²Î¦rÅé) Then
    ¥Ò°©¤å²§¼g¦rªí.Index = "½s¸¹"
    ¥Ò°©¤å²§¼g¦rªí.Seek "=", ½s¸¹
    If Not IsNull(¥Ò°©¤å²§¼g¦rªí.Fields("·¢®Ñ½s¸¹")) Then
        ½s¸¹ = ¥Ò°©¤å²§¼g¦rªí.Fields("·¢®Ñ½s¸¹")
    Else
        ¥Ò°©¤åÀË¦rªí.Index = "½s¸¹"
        ¥Ò°©¤åÀË¦rªí.Seek "=", ½s¸¹
        ½s¸¹ = ¥Ò°©¤åÀË¦rªí.Fields("·¢®Ñ½s¸¹")
    End If
    ÀË¦rªí.Index = "½s¸¹"
    ÀË¦rªí.Seek "=", ½s¸¹
    ¦r§Î = ÀË¦rªí.Fields("¦r½X")
ElseIf ¤¤¬ã°|·¡¨t¤å¦r(¨t²Î¦rÅé) Then
    ·¡¨t¤å¦r²§¼g¦rªí.Index = "½s¸¹"
    ·¡¨t¤å¦r²§¼g¦rªí.Seek "=", ½s¸¹
    If Not IsNull(·¡¨t¤å¦r²§¼g¦rªí.Fields("·¢®Ñ½s¸¹")) Then
        ½s¸¹ = ·¡¨t¤å¦r²§¼g¦rªí.Fields("·¢®Ñ½s¸¹")
    Else
        ·¡¨t¤å¦rÀË¦rªí.Index = "½s¸¹"
        ·¡¨t¤å¦rÀË¦rªí.Seek "=", ½s¸¹
        ½s¸¹ = ·¡¨t¤å¦rÀË¦rªí.Fields("·¢®Ñ½s¸¹")
    End If
    ÀË¦rªí.Index = "½s¸¹"
    ÀË¦rªí.Seek "=", ½s¸¹
    ¦r§Î = ÀË¦rªí.Fields("¦r½X")
End If

Set ÀË¦rªí = ·¢®ÑÀË¦rªí
ÀË¦rªí.Index = "½s¸¹"
ÀË¦rªí.Seek "=", ½s¸¹
¦r§Î = ÀË¦rªí.Fields("¦r½X")

µ¹©wªÅ¥Õ­È
'¤À¸Ñ = 4
'³¡¥ó§Ç = ""

'If Len(¦r§Î) = 2 And ¬O§_¬°²Õ¦r²Å¸¹(Mid(¦r§Î, 1, 1), 4, 11) > 0 Then
'   ³¡¥ó§Ç = ¦r§Î
'   ¤À¸Ñ = 5
'   ÀË¦rªí.Index = "ºc¦r¦¡"
'   ÀË¦rªí.Seek "=", ¤À¸Ñ, ³¡¥ó§Ç
'ElseIf ½s¸¹ <= 0 Then
'   ÀË¦rªí.Index = "¦r§Î"
'   ÀË¦rªí.Seek "=", ¦r§Î
'Else
'   ÀË¦rªí.Index = "½s¸¹"
'   ÀË¦rªí.Seek "=", ½s¸¹
'End If
   
'If Not ÀË¦rªí.NoMatch Then
'   If Not IsNull(ÀË¦rªí.Fields("³s±µ²Å¸¹")) Then ¤À¸Ñ = ÀË¦rªí.Fields("³s±µ²Å¸¹")
'   If Not IsNull(ÀË¦rªí.Fields("³¡¥ó§Ç")) Then ³¡¥ó§Ç = ÀË¦rªí.Fields("³¡¥ó§Ç")
'Else
'   ¤À¸Ñ = 0
'   ³¡¥ó§Ç = ""
'End If
   
'If Len(¦r§Î) = 1 Or ¬O§_¬°²Õ¦r²Å¸¹(Left$(¦r§Î, 1), 1, 14) <> 0 Then
    If Not ÀË¦rªí.NoMatch Then
'      Do While Not ÀË¦rªí.EOF
'         If ÀË¦rªí.Fields("³s±µ²Å¸¹") = ¤À¸Ñ Or (IsNull(ÀË¦rªí.Fields("³s±µ²Å¸¹")) And ¤À¸Ñ = 4) Then Exit Do
'         ÀË¦rªí.MoveNext
'      Loop
      
'      If ÀË¦rªí.EOF Then
'         mdiº~¦r¦r§Î.txtºc¦r¦¡.Text = ³¡¥ó§Ç
'         Exit Sub
'      End If
      
      ¨t²Î½s¸¹ = ÀË¦rªí.Fields("½s¸¹")
      
      If ÀË¦rªí.Fields("¦rÅé") = 0 Then
         ¦r«¬ÀÉ = ²{¥Î¦rÅé
      Else
         ¦r«¬ÀÉ = ¦rÅé°}¦C(ÀË¦rªí.Fields("¦rÅé"))
      End If
      mdiº~¦r¦r§Î.txt¦r§Î.FontName = ¦r«¬ÀÉ
      
      If Not IsNull(ÀË¦rªí.Fields("½s¸¹")) Then mdiº~¦r¦r§Î.txt½s¸¹.Text = ÀË¦rªí.Fields("½s¸¹")
      If Not IsNull(ÀË¦rªí.Fields("¦rÅé")) Then mdiº~¦r¦r§Î.txt¥~¦r¶°.Text = ÀË¦rªí.Fields("¦rÅé")

      If Not IsNull(ÀË¦rªí.Fields("¦r§Î")) Then
         mdiº~¦r¦r§Î.txt¦r§Î.Text = ÀË¦rªí.Fields("¦r§Î")
      Else
         If Not IsNull(ÀË¦rªí.Fields("¦r½X")) Then
            mdiº~¦r¦r§Î.txt¦r§Î.Text = ÀË¦rªí.Fields("¦r½X")
         Else
            mdiº~¦r¦r§Î.txt¦r§Î.Text = "¡´"
         End If
      End If
      
      If Not IsNull(ÀË¦rªí.Fields("µ§µe")) Then mdiº~¦r¦r§Î.txtÁ`µ§µe.Text = ÀË¦rªí.Fields("µ§µe")
      If Not IsNull(ÀË¦rªí.Fields("³¡­º")) Then mdiº~¦r¦r§Î.txt³¡­º.Text = ´M§ä³¡­º(ÀË¦rªí.Fields("³¡­º"))
      If Not IsNull(ÀË¦rªí.Fields("³¡­ºµ§µe")) Then mdiº~¦r¦r§Î.txt¦©°£³¡­ºµ§µe.Text = ÀË¦rªí.Fields("³¡­ºµ§µe")
      If Not IsNull(ÀË¦rªí.Fields("ª`­µ")) Then mdiº~¦r¦r§Î.txtª`­µ.Text = ´M§äª`­µ(ÀË¦rªí.Fields("ª`­µ"))
      If Not IsNull(ÀË¦rªí.Fields("BIG5")) Then mdiº~¦r¦r§Î.txt¤º½X.Text = ÀË¦rªí.Fields("BIG5")
      If Not IsNull(ÀË¦rªí.Fields("­Ü¾e")) Then mdiº~¦r¦r§Î.txt­Ü¾e½X.Text = Âà´«­^¤å¨ì­Ü¾e(ÀË¦rªí.Fields("­Ü¾e"))
      If Not IsNull(ÀË¦rªí.Fields("»·ªFº~»y¤j¦r¨å")) Then
         mdiº~¦r¦r§Î.txt¥U¼Æ.Text = ÀË¦rªí.Fields("»·ªFº~»y¤j¦r¨å")
      Else
         mdiº~¦r¦r§Î.txt¥U¼Æ.Text = ""
      End If
      
      If Not IsNull(ÀË¦rªí.Fields("²Õ¦r¦r¼Æ")) Then
         mdiº~¦r¦r§Î.txt²Õ¦r¦r¼Æ.Text = ÀË¦rªí.Fields("²Õ¦r¦r¼Æ")
      Else
         mdiº~¦r¦r§Î.txt²Õ¦r¦r¼Æ.Text = ""
      End If
      
      If Not IsNull(ÀË¦rªí.Fields("²Õ¦r¦r¼Æ¤G")) Then
         mdiº~¦r¦r§Î.txt²Õ¦r¦r¼Æ§t²§¼g.Text = ÀË¦rªí.Fields("²Õ¦r¦r¼Æ¤G")
      Else
         mdiº~¦r¦r§Î.txt²Õ¦r¦r¼Æ§t²§¼g.Text = ""
      End If
   Else
      mdiº~¦r¦r§Î.txtºc¦r¦¡.Text = ³¡¥ó§Ç
   End If
'Else
'   mdiº~¦r¦r§Î.txtºc¦r¦¡.Text = ¦r§Î
'End If

End Sub

Public Sub µ¹©wªÅ¥Õ­È()
mdiº~¦r¦r§Î.txt¥~¦r¶°.Text = ""
mdiº~¦r¦r§Î.txt¦r§Î.Text = ""
mdiº~¦r¦r§Î.txt­«¤å.Text = ""
mdiº~¦r¦r§Î.txt¥jº~¦r.Text = ""
mdiº~¦r¦r§Î.txtÁ`µ§µe.Text = ""
mdiº~¦r¦r§Î.txt³¡­º.Text = ""
mdiº~¦r¦r§Î.txt¦©°£³¡­ºµ§µe.Text = ""
mdiº~¦r¦r§Î.txtª`­µ.Text = ""
mdiº~¦r¦r§Î.txt¤º½X.Text = ""
mdiº~¦r¦r§Î.txt­Ü¾e½X.Text = ""
mdiº~¦r¦r§Î.txtºc¦r¦¡.Text = ""
mdiº~¦r¦r§Î.txt¥U¼Æ.Text = ""
mdiº~¦r¦r§Î.txt²Õ¦r¦r¼Æ.Text = ""
mdiº~¦r¦r§Î.txt²Õ¦r¦r¼Æ§t²§¼g.Text = ""
¨t²Î½s¸¹ = -999999
End Sub

Public Sub Â^¨úºc¦r¦¡(¨t²Î¦rÅé As String, val¦r§Î As String, val½s¸¹ As Long)

Dim ³¡¥ó§Ç As String, ¤À¸Ñ As Integer
Dim ½s¸¹ As Long, ¦r§Î As String, ¦rÀY As String, ­·®æ½X As String


½s¸¹ = val½s¸¹
¦r§Î = val¦r§Î

If ½s¸¹ <= 0 Then Exit Sub

¦rÀY = ""
­·®æ½X = ""
½Æ»s­·®æ½X = False

If ¨t²Î¦rÅé = "¥_®v¤j»¡¤å¤p½f" Or ¨t²Î¦rÅé = "¥_®v¤j»¡¤å­«¤å" Then
    ¤p½fÀË¦rªí.Index = "½s¸¹"
    ¤p½fÀË¦rªí.Seek "=", ½s¸¹
    ½s¸¹ = ¤p½fÀË¦rªí.Fields("·¢®Ñ½s¸¹")
    mdiº~¦r¦r§Î.txt­«¤å = ¤p½fÀË¦rªí.Fields("¦rÅé")
    mdiº~¦r¦r§Î.txt¥jº~¦r = ¤p½fÀË¦rªí.Fields("¦r½X")
    mdiº~¦r¦r§Î.txt¥jº~¦r.FontName = ¦rÅé°}¦C(¤p½fÀË¦rªí.Fields("¦rÅé"))
    If ¤p½fÀË¦rªí.Fields("ªş½X") = 1 Then
        ­·®æ½X = ¤p½fÀË¦rªí.Fields("¦r·½")
    Else
        ­·®æ½X = ¤p½fÀË¦rªí.Fields("¦r·½") & ";" & ¤p½fÀË¦rªí.Fields("ªş½X")
    End If
    ½Æ»s­·®æ½X = True
    
    ÀË¦rªí.Index = "½s¸¹"
    ÀË¦rªí.Seek "=", ½s¸¹
    ¦r§Î = ÀË¦rªí.Fields("¦r½X")
ElseIf ¤¤¬ã°|ª÷¤å(¨t²Î¦rÅé) Then
    ª÷¤å²§¼g¦rªí.Index = "½s¸¹"
    ª÷¤å²§¼g¦rªí.Seek "=", ½s¸¹
    If Not IsNull(ª÷¤å²§¼g¦rªí.Fields("·¢®Ñ½s¸¹")) Then
        ½s¸¹ = ª÷¤å²§¼g¦rªí.Fields("·¢®Ñ½s¸¹")
        mdiº~¦r¦r§Î.txt­«¤å = ª÷¤å²§¼g¦rªí.Fields("¦rÅé")
        mdiº~¦r¦r§Î.txt¥jº~¦r = ª÷¤å²§¼g¦rªí.Fields("¦r½X")
        mdiº~¦r¦r§Î.txt¥jº~¦r.FontName = ¦rÅé°}¦C(ª÷¤å²§¼g¦rªí.Fields("¦rÅé"))
        If mdiº~¦r¦r§Î.mnu_Åã¥Ü­·®æ½X.Checked = True And ((ª÷¤å²§¼g¦rªí.Fields("¤W½u") > 0 And ª÷¤å²§¼g¦rªí.Fields("¤W½u") < 6)) And ª÷¤å²§¼g¦rªí.Fields("¦X¤å") = 0 Then
            If ª÷¤å²§¼g¦rªí.Fields("ªş½X") = 1 Then
                ­·®æ½X = "¶°¦¨" & ª÷¤å²§¼g¦rªí.Fields("¾¹¸¹")
            Else
                ­·®æ½X = "¶°¦¨" & ª÷¤å²§¼g¦rªí.Fields("¾¹¸¹") & ";" & ª÷¤å²§¼g¦rªí.Fields("ªş½X")
            End If
            ½Æ»s­·®æ½X = True
        End If
    Else
        ª÷¤åÀË¦rªí.Index = "½s¸¹"
        ª÷¤åÀË¦rªí.Seek "=", ½s¸¹
        ½s¸¹ = ª÷¤åÀË¦rªí.Fields("·¢®Ñ½s¸¹")
        mdiº~¦r¦r§Î.txt­«¤å = ª÷¤åÀË¦rªí.Fields("¦rÅé")
        mdiº~¦r¦r§Î.txt¥jº~¦r = ª÷¤åÀË¦rªí.Fields("¦r½X")
        mdiº~¦r¦r§Î.txt¥jº~¦r.FontName = ¦rÅé°}¦C(ª÷¤åÀË¦rªí.Fields("¦rÅé"))
    End If
    ÀË¦rªí.Index = "½s¸¹"
    ÀË¦rªí.Seek "=", ½s¸¹
    ¦r§Î = ÀË¦rªí.Fields("¦r½X")
ElseIf ¤¤¬ã°|¥Ò°©¤å(¨t²Î¦rÅé) Then
    ¥Ò°©¤å²§¼g¦rªí.Index = "½s¸¹"
    ¥Ò°©¤å²§¼g¦rªí.Seek "=", ½s¸¹
    If Not IsNull(¥Ò°©¤å²§¼g¦rªí.Fields("·¢®Ñ½s¸¹")) Then
        ½s¸¹ = ¥Ò°©¤å²§¼g¦rªí.Fields("·¢®Ñ½s¸¹")
        mdiº~¦r¦r§Î.txt­«¤å = ¥Ò°©¤å²§¼g¦rªí.Fields("¦rÅé")
        mdiº~¦r¦r§Î.txt¥jº~¦r = ¥Ò°©¤å²§¼g¦rªí.Fields("¦r½X")
        mdiº~¦r¦r§Î.txt¥jº~¦r.FontName = ¦rÅé°}¦C(¥Ò°©¤å²§¼g¦rªí.Fields("¦rÅé"))
        If mdiº~¦r¦r§Î.mnu_Åã¥Ü­·®æ½X.Checked = True And ((¥Ò°©¤å²§¼g¦rªí.Fields("¤W½u") > 0 And ¥Ò°©¤å²§¼g¦rªí.Fields("¤W½u") < 6)) And ¥Ò°©¤å²§¼g¦rªí.Fields("¦X¤å") = 0 And Not IsNull(¥Ò°©¤å²§¼g¦rªí.Fields("¥X³B")) Then
            ­·®æ½X = ¥Ò°©¤å²§¼g¦rªí.Fields("¥X³B")
            If IsNumeric(Left(­·®æ½X, 1)) Then ­·®æ½X = "¦X¶°" & ­·®æ½X
            If ¥Ò°©¤å²§¼g¦rªí.Fields("ªş½X") > 1 Then
                ­·®æ½X = ­·®æ½X & ";" & ¥Ò°©¤å²§¼g¦rªí.Fields("ªş½X")
            End If
            ½Æ»s­·®æ½X = True
        End If
    Else
        ¥Ò°©¤åÀË¦rªí.Index = "½s¸¹"
        ¥Ò°©¤åÀË¦rªí.Seek "=", ½s¸¹
        ½s¸¹ = ¥Ò°©¤åÀË¦rªí.Fields("·¢®Ñ½s¸¹")
        mdiº~¦r¦r§Î.txt­«¤å = ¥Ò°©¤åÀË¦rªí.Fields("¦rÅé")
        mdiº~¦r¦r§Î.txt¥jº~¦r = ¥Ò°©¤åÀË¦rªí.Fields("¦r½X")
        mdiº~¦r¦r§Î.txt¥jº~¦r.FontName = ¦rÅé°}¦C(¥Ò°©¤åÀË¦rªí.Fields("¦rÅé"))
    End If
    ÀË¦rªí.Index = "½s¸¹"
    ÀË¦rªí.Seek "=", ½s¸¹
    ¦r§Î = ÀË¦rªí.Fields("¦r½X")
ElseIf ¤¤¬ã°|·¡¨t¤å¦r(¨t²Î¦rÅé) Then
    ·¡¨t¤å¦r²§¼g¦rªí.Index = "½s¸¹"
    ·¡¨t¤å¦r²§¼g¦rªí.Seek "=", ½s¸¹
    If Not IsNull(·¡¨t¤å¦r²§¼g¦rªí.Fields("·¢®Ñ½s¸¹")) Then
        ½s¸¹ = ·¡¨t¤å¦r²§¼g¦rªí.Fields("·¢®Ñ½s¸¹")
        mdiº~¦r¦r§Î.txt­«¤å = ·¡¨t¤å¦r²§¼g¦rªí.Fields("¦rÅé")
        mdiº~¦r¦r§Î.txt¥jº~¦r = ·¡¨t¤å¦r²§¼g¦rªí.Fields("¦r½X")
        mdiº~¦r¦r§Î.txt¥jº~¦r.FontName = ¦rÅé°}¦C(·¡¨t¤å¦r²§¼g¦rªí.Fields("¦rÅé"))
        If mdiº~¦r¦r§Î.mnu_Åã¥Ü­·®æ½X.Checked = True And ((·¡¨t¤å¦r²§¼g¦rªí.Fields("¤W½u") > 0 And ·¡¨t¤å¦r²§¼g¦rªí.Fields("¤W½u") < 6)) And ·¡¨t¤å¦r²§¼g¦rªí.Fields("¦X¤å") = 0 And Not IsNull(·¡¨t¤å¦r²§¼g¦rªí.Fields("¥X³B")) Then
            ­·®æ½X = ·¡¨t¤å¦r²§¼g¦rªí.Fields("¥X³B")
            If ·¡¨t¤å¦r²§¼g¦rªí.Fields("ªş½X") > 1 Then
                ­·®æ½X = ­·®æ½X & ";" & ·¡¨t¤å¦r²§¼g¦rªí.Fields("ªş½X")
            End If
            ½Æ»s­·®æ½X = True
        End If
    Else
        ·¡¨t¤å¦rÀË¦rªí.Index = "½s¸¹"
        ·¡¨t¤å¦rÀË¦rªí.Seek "=", ½s¸¹
        ½s¸¹ = ·¡¨t¤å¦rÀË¦rªí.Fields("·¢®Ñ½s¸¹")
        mdiº~¦r¦r§Î.txt­«¤å = ·¡¨t¤å¦rÀË¦rªí.Fields("¦rÅé")
        mdiº~¦r¦r§Î.txt¥jº~¦r = ·¡¨t¤å¦rÀË¦rªí.Fields("¦r½X")
        mdiº~¦r¦r§Î.txt¥jº~¦r.FontName = ¦rÅé°}¦C(·¡¨t¤å¦rÀË¦rªí.Fields("¦rÅé"))
    End If
    ÀË¦rªí.Index = "½s¸¹"
    ÀË¦rªí.Seek "=", ½s¸¹
    ¦r§Î = ÀË¦rªí.Fields("¦r½X")
End If

¤À¸Ñ = 4
³¡¥ó§Ç = ""

If Len(¦r§Î) = 2 And ¬O§_¬°²Õ¦r²Å¸¹(Mid(¦r§Î, 1, 1), 4, 11) > 0 Then
   ³¡¥ó§Ç = ¦r§Î
   ¤À¸Ñ = 5
   ÀË¦rªí.Index = "ºc¦r¦¡"
   ÀË¦rªí.Seek "=", ¤À¸Ñ, ³¡¥ó§Ç
ElseIf ½s¸¹ = 0 Then
   ÀË¦rªí.Index = "¦r§Î"
   ÀË¦rªí.Seek "=", ¦r§Î
Else
   ÀË¦rªí.Index = "½s¸¹"
   ÀË¦rªí.Seek "=", ½s¸¹
End If
   
If Not ÀË¦rªí.NoMatch Then
   If Not ÀË¦rªí.EOF Then
      If Not IsNull(ÀË¦rªí.Fields("¦r§Î")) Then ¦rÀY = ÀË¦rªí.Fields("¦r§Î")
      If Not IsNull(ÀË¦rªí.Fields("³s±µ²Å¸¹")) Then ¤À¸Ñ = ÀË¦rªí.Fields("³s±µ²Å¸¹")
      If Not IsNull(ÀË¦rªí.Fields("³¡¥ó§Ç")) Then ³¡¥ó§Ç = ÀË¦rªí.Fields("³¡¥ó§Ç")
   End If
Else
   ¤À¸Ñ = 0
   ³¡¥ó§Ç = ""
End If

If Not ÀË¦rªí.NoMatch Then
   Do While Not ÀË¦rªí.EOF
      If (ÀË¦rªí.Fields("³s±µ²Å¸¹") = ¤À¸Ñ) Or (IsNull(ÀË¦rªí.Fields("³s±µ²Å¸¹")) And ¤À¸Ñ = 4) Then Exit Do
      ÀË¦rªí.MoveNext
   Loop
   ©ì¦²¦r¦ê = ÀË¦rªí.Fields("³¡¥ó§Ç")
   
   If ­·®æ½X <> "" Then
        mdiº~¦r¦r§Î.txtºc¦r¦¡.Text = ´M§ä­·®æ½X(¦rÀY, ³¡¥ó§Ç, ­·®æ½X)
   Else
        mdiº~¦r¦r§Î.txtºc¦r¦¡.Text = ´M§ä²Õ¦r¦¡(ÀË¦rªí.Fields("³s±µ²Å¸¹"), ÀË¦rªí.Fields("³¡¥ó§Ç"))
   End If
End If

End Sub

Public Function ²Õ¦r²Å¸¹ªø«×(¦r§Î As String) As Integer

²Õ¦r²Å¸¹ªø«× = 0

Select Case ¦r§Î
       Case ²Õ¦r²Å¸¹°}¦C(4)
            ²Õ¦r²Å¸¹ªø«× = 2
       Case ²Õ¦r²Å¸¹°}¦C(5)
            ²Õ¦r²Å¸¹ªø«× = 2
       Case ²Õ¦r²Å¸¹°}¦C(6)
            ²Õ¦r²Å¸¹ªø«× = 3
       Case ²Õ¦r²Å¸¹°}¦C(7)
            ²Õ¦r²Å¸¹ªø«× = 3
       Case ²Õ¦r²Å¸¹°}¦C(8)
            ²Õ¦r²Å¸¹ªø«× = 3
       Case ²Õ¦r²Å¸¹°}¦C(9)
            ²Õ¦r²Å¸¹ªø«× = 4
       Case ²Õ¦r²Å¸¹°}¦C(10)
            ²Õ¦r²Å¸¹ªø«× = 4
       Case ²Õ¦r²Å¸¹°}¦C(11)
            ²Õ¦r²Å¸¹ªø«× = 4
End Select

End Function

Public Function ´M§ä³¡­º(½s¸¹ As Integer) As String
´M§ä³¡­º = ""
±dº³³¡­º.Index = "½s¸¹"
±dº³³¡­º.Seek "=", ½s¸¹
If Not ±dº³³¡­º.NoMatch Then
   ´M§ä³¡­º = ±dº³³¡­º.Fields("¦r§Î")
End If
End Function

Public Function ´M§ä»¡¤å³¡­º(½s¸¹ As Integer) As String

»¡¤å³¡­º.Index = "½s¸¹"
»¡¤å³¡­º.Seek "=", ½s¸¹

If Not »¡¤å³¡­º.NoMatch Then
   ´M§ä»¡¤å³¡­º = »¡¤å³¡­º.Fields("¦r§Î")
Else
   ´M§ä»¡¤å³¡­º = ""
End If

End Function

Public Function ´M§äª`­µ(ª`­µ As String) As String

´M§äª`­µ = Left(ª`­µ, Len(ª`­µ) - 1)

Select Case Right$(ª`­µ, 1)
       Case 1
            ´M§äª`­µ = ´M§äª`­µ
       Case 2
            ´M§äª`­µ = ´M§äª`­µ & "£½"
       Case 3
            ´M§äª`­µ = ´M§äª`­µ & "£¾"
       Case 4
            ´M§äª`­µ = ´M§äª`­µ & "£¿"
       Case 5
            ´M§äª`­µ = "£»" & ´M§äª`­µ
End Select

End Function

Public Function «Ø¥ßSQL(txtºc¦r¦¡ As String, sql As Integer) As String
Dim ²Å¸¹ As String
Dim ²Õ¦r¦¡ As String
Dim i As Integer

«Ø¥ßSQL = ""
   
²Å¸¹ = ""
For i = 1 To Len(txtºc¦r¦¡)
    If Mid(txtºc¦r¦¡, i, 1) <> "*" Then
       ²Õ¦r¦¡ = ²Õ¦r¦¡ + Mid(txtºc¦r¦¡, i, 1) + ²Å¸¹
    End If
Next i

txtºc¦r¦¡ = ²Õ¦r¦¡
   
If sql = 1 Then
   ²Å¸¹ = "*"
   ²Õ¦r¦¡ = ²Å¸¹
   For i = 1 To Len(txtºc¦r¦¡)
       ²Õ¦r¦¡ = ²Õ¦r¦¡ + Mid(txtºc¦r¦¡, i, 1) + ²Å¸¹
   Next i
End If
«Ø¥ßSQL = ²Õ¦r¦¡

End Function

Public Function ¦r®Ú±Æ§Ç(¦r®Ú§Ç As String) As String
Dim ¦r®Úªí As Recordset, ¦r®Ú²Õ As String
Dim ¦r§Î½s¸¹°ïÅ|(30) As Integer, ¦r§Î°ïÅ|(30) As String
Dim Maxlen As Integer, i As Integer, j As Integer
Dim temp1 As Integer, temp2 As String

Set ¦r®Úªí = ¨t²Î¸ê®Æ®w.OpenRecordset("¦r®Ú")

¦r®Úªí.Index = "¦r§Î"
Maxlen = Len(¦r®Ú§Ç)

'´M§ä½s¸¹
For i = 1 To Maxlen
    ¦r§Î°ïÅ|(i - 1) = Mid(¦r®Ú§Ç, i, 1)
    ¦r®Úªí.Seek "=", Mid(¦r®Ú§Ç, i, 1)
    If Not ¦r®Úªí.NoMatch Then
       ¦r§Î½s¸¹°ïÅ|(i - 1) = ¦r®Úªí.Fields("½s¸¹")
    Else
       '§ä¤£¨ì«hµ¹-1
       ¦r§Î½s¸¹°ïÅ|(i - 1) = -1
    End If
Next i
    
'ªwªj±Æ§Ç
For i = 0 To Maxlen - 1
    For j = i + 1 To Maxlen - 1
        If ¦r§Î½s¸¹°ïÅ|(i) > ¦r§Î½s¸¹°ïÅ|(j) Then
           temp1 = ¦r§Î½s¸¹°ïÅ|(i)
           ¦r§Î½s¸¹°ïÅ|(i) = ¦r§Î½s¸¹°ïÅ|(j)
           ¦r§Î½s¸¹°ïÅ|(j) = temp1
           temp2 = ¦r§Î°ïÅ|(i)
           ¦r§Î°ïÅ|(i) = ¦r§Î°ïÅ|(j)
           ¦r§Î°ïÅ|(j) = temp2
         End If
     Next j
Next i
    
'±N¦r§Î°ïÅ|²Õ°_¨Ó
¦r®Ú²Õ = ""
For i = 0 To Maxlen - 1
    ¦r®Ú²Õ = ¦r®Ú²Õ & ¦r§Î°ïÅ|(i)
Next i

¦r®Ú±Æ§Ç = ¦r®Ú²Õ

¦r®Úªí.Close

End Function


Public Function Âà´«¦rÅé(¦rÅé As String) As String

If Åã¥Ü¦r«¬ = "²Ó©úÅé" Then
    If ¦rÅé = "²Ó©úÅé" Then
        Âà´«¦rÅé = "¼Ğ·¢Åé"
    Else
        Âà´«¦rÅé = ¦rÅé
    End If
Else
    Âà´«¦rÅé = ¦rÅé
End If
    


End Function



Public Function Âà´«Åã¥Ü¦r«¬(ByVal ¦r«¬ As String) As String

Dim ifound As Integer, rstr As String

If Åã¥Ü¦r«¬ = "²Ó©úÅé" Then
    ifound = InStr(1, ¦r«¬, "¼Ğ·¢Åé")
    If ifound > 0 Then
        rstr = ¦r«¬
        rstr = Left(rstr, ifound - 1) & "²Ó©úÅé" & Right(rstr, Len(rstr) - ifound - 2)
        Âà´«Åã¥Ü¦r«¬ = rstr
    Else
        Âà´«Åã¥Ü¦r«¬ = ¦r«¬
    End If
Else
    Âà´«Åã¥Ü¦r«¬ = ¦r«¬
End If
    
End Function

Public Function ¤Á´«Åã¥Ü¦r«¬(ByVal ¦r«¬ As String) As String

Dim ifound As Integer, rstr As String

If Åã¥Ü¦r«¬ = "²Ó©úÅé" Then
    ifound = InStr(1, ¦r«¬, "¼Ğ·¢Åé")
    If ifound > 0 Then
        rstr = ¦r«¬
        rstr = Left(rstr, ifound - 1) & "²Ó©úÅé" & Right(rstr, Len(rstr) - ifound - 2)
        ¤Á´«Åã¥Ü¦r«¬ = rstr
    Else
        ¤Á´«Åã¥Ü¦r«¬ = ¦r«¬
    End If
Else
    ifound = InStr(1, ¦r«¬, "²Ó©úÅé")
    If ifound > 0 Then
        rstr = ¦r«¬
        rstr = Left(rstr, ifound - 1) & "¼Ğ·¢Åé" & Right(rstr, Len(rstr) - ifound - 2)
        ¤Á´«Åã¥Ü¦r«¬ = rstr
    Else
        ¤Á´«Åã¥Ü¦r«¬ = ¦r«¬
    End If
End If
    
End Function


Public Function Âà´«RTF¯Ê¦r(§t¯Ê¦r¦r¦ê As String, Åã¥Ü¦rÅé As String) As String

Dim RTF¦r¦ê As String
Dim °_©l¦ì¸m As Long, ²×µ²¦ì¸m As Long, ¦r¤¸¦ì¸m As Long
Dim ¥Ø«e¦r¤¸ As String, ¤W¤@¦r¤¸ As String, ¤U¤@¦r¤¸ As String
Dim ³s±µ²Å¸¹ As Integer, ³¡¥ó§Ç As String, ºc¦r¦¡ As String
Dim ¦rÅé As Integer, ¦r½X As String
Dim i As Integer, wk As String

¦r¤¸¦ì¸m = 1
²×µ²¦ì¸m = 0

Do While ¦r¤¸¦ì¸m <= Len(§t¯Ê¦r¦r¦ê)

¥Ø«e¦r¤¸ = Mid(§t¯Ê¦r¦r¦ê, ¦r¤¸¦ì¸m, 1)

If ¥Ø«e¦r¤¸ >= "ñ" And ¥Ø«e¦r¤¸ <= "ş" Then
     
    '§PÂ_¦r¦ê¤º®e§ÎºA
   
    If (¥Ø«e¦r¤¸ <> "ü") Then
   
        If (¥Ø«e¦r¤¸ >= "ñ" And ¥Ø«e¦r¤¸ <= "ó") Then
            °_©l¦ì¸m = ¦r¤¸¦ì¸m - 1
        Else
            °_©l¦ì¸m = ¦r¤¸¦ì¸m
        End If
      
        Do
            ¤W¤@¦r¤¸ = ¥Ø«e¦r¤¸
            ¦r¤¸¦ì¸m = ¦r¤¸¦ì¸m + 1
            If ¦r¤¸¦ì¸m <= Len(§t¯Ê¦r¦r¦ê) Then
                ¥Ø«e¦r¤¸ = Mid(§t¯Ê¦r¦r¦ê, ¦r¤¸¦ì¸m, 1)
            Else
                ¥Ø«e¦r¤¸ = Chr(13)
            End If
        Loop Until (¤W¤@¦r¤¸ < "ñ" Or ¤W¤@¦r¤¸ > "û") And (¥Ø«e¦r¤¸ < "ñ" Or ¥Ø«e¦r¤¸ > "ó")
        
        ¦r¤¸¦ì¸m = ¦r¤¸¦ì¸m - 1
        ²×µ²¦ì¸m = ¦r¤¸¦ì¸m
      
        ºc¦r¦¡ = Mid(§t¯Ê¦r¦r¦ê, °_©l¦ì¸m, ²×µ²¦ì¸m - °_©l¦ì¸m + 1)
        
        ³¡¥ó§Ç = ""
        
        ³s±µ²Å¸¹ = 5
        For i = 1 To Len(ºc¦r¦¡)
            wk = Mid(ºc¦r¦¡, i, 1)
            If (wk < "ñ" Or wk > "ó") Then
                ³¡¥ó§Ç = ³¡¥ó§Ç + wk
            Else
                If wk = "ò" Then
                    ³s±µ²Å¸¹ = 1
                ElseIf wk = "ñ" Then
                    ³s±µ²Å¸¹ = 2
                Else
                    ³s±µ²Å¸¹ = 3
                End If
            End If
        Next i
        
    Else
    
        °_©l¦ì¸m = ¦r¤¸¦ì¸m
        ³¡¥ó§Ç = ""
        
        Do
            ¦r¤¸¦ì¸m = ¦r¤¸¦ì¸m + 1
            If ¦r¤¸¦ì¸m <= Len(§t¯Ê¦r¦r¦ê) Then
                ¥Ø«e¦r¤¸ = Mid(§t¯Ê¦r¦r¦ê, ¦r¤¸¦ì¸m, 1)
            Else
                ¥Ø«e¦r¤¸ = Chr(13)
            End If
        Loop Until (¥Ø«e¦r¤¸ = "ı") Or (¥Ø«e¦r¤¸ = Chr(13))
        
        ²×µ²¦ì¸m = ¦r¤¸¦ì¸m
        
        If ¥Ø«e¦r¤¸ = "ı" Then
            ºc¦r¦¡ = Mid(§t¯Ê¦r¦r¦ê, °_©l¦ì¸m, ²×µ²¦ì¸m - °_©l¦ì¸m + 1)
            ³¡¥ó§Ç = Mid(ºc¦r¦¡, 2, Len(ºc¦r¦¡) - 2)
            ³s±µ²Å¸¹ = 4
        End If
        
    End If
   
    ·¢®ÑÀË¦rªí.Index = "ºc¦r¦¡"
    ·¢®ÑÀË¦rªí.Seek "=", ³s±µ²Å¸¹, ³¡¥ó§Ç
    
    If Not ·¢®ÑÀË¦rªí.NoMatch Then
    
        ¦rÅé = ·¢®ÑÀË¦rªí.Fields("¦rÅé")
        ¦r½X = ·¢®ÑÀË¦rªí.Fields("¦r½X")
    
        If Åã¥Ü¦rÅé = "¼Ğ·¢Åé" Then
            RTF¦r¦ê = RTF¦r¦ê & "{\f" & ¦rÅé & Âà´«RTF¦r¦ê(¦r½X) & "}"
        Else
            RTF¦r¦ê = RTF¦r¦ê & "{\f" & ¦rÅé + 16 & Âà´«RTF¦r¦ê(¦r½X) & "}"
        End If
    Else
        RTF¦r¦ê = RTF¦r¦ê & "{\f0\'a1\'b4}" '¡´
    End If
    
Else
    
    If ¦r¤¸¦ì¸m < Len(§t¯Ê¦r¦r¦ê) Then
        ¤U¤@¦r¤¸ = Mid(§t¯Ê¦r¦r¦ê, ¦r¤¸¦ì¸m + 1, 1)
        If (¤U¤@¦r¤¸ >= "ñ" And ¤U¤@¦r¤¸ <= "ó") Then GoTo ³B²z¤U¤@­Ó¯Ê¦r
    End If
    
    If ¦r¤¸¦ì¸m - ²×µ²¦ì¸m = 1 Then
        If Åã¥Ü¦rÅé = "¼Ğ·¢Åé" Then
            RTF¦r¦ê = RTF¦r¦ê & "{\f0" & Âà´«RTF¦r¦ê(¥Ø«e¦r¤¸) & "}"
        Else
            RTF¦r¦ê = RTF¦r¦ê & "{\f16" & Âà´«RTF¦r¦ê(¥Ø«e¦r¤¸) & "}"
        End If
    Else
        RTF¦r¦ê = Left(RTF¦r¦ê, Len(RTF¦r¦ê) - 1) & Âà´«RTF¦r¦ê(¥Ø«e¦r¤¸) & "}"
    End If

End If

³B²z¤U¤@­Ó¯Ê¦r:

¦r¤¸¦ì¸m = ¦r¤¸¦ì¸m + 1

Loop

If Åã¥Ü¦rÅé = "¼Ğ·¢Åé" Then
    Âà´«RTF¯Ê¦r = "{" & rtf_version & character_set & ¼Ğ·¢Åé¦r«¬ªí & RTF¦r¦ê & "}"
Else
    Âà´«RTF¯Ê¦r = "{" & rtf_version & character_set & ²Ó©úÅé¦r«¬ªí & RTF¦r¦ê & "}"
End If

End Function

Public Function Âà´«RTF¦r¦ê(¦r§Î As String) As String

Dim ¦r½X As String

¦r½X = CStr(Hex(Asc(¦r§Î)))

If Len(¦r½X) = 4 Then
    Âà´«RTF¦r¦ê = "\'" & LCase(Left(¦r½X, 2)) & "\'" & LCase(Right(¦r½X, 2))
ElseIf Len(¦r½X) = 2 Then
     Âà´«RTF¦r¦ê = "\'" & LCase(Left(¦r½X, 2))
Else
    Âà´«RTF¦r¦ê = ""
End If

End Function

Public Function ¤¤¬ã°|ª÷¤å(¦r«¬) As Boolean

If InStr(1, ¦r«¬, "¤¤¬ã°|ª÷¤å") > 0 Then
    ¤¤¬ã°|ª÷¤å = True
Else
    ¤¤¬ã°|ª÷¤å = False
End If

End Function

Public Function ¤¤¬ã°|¥Ò°©¤å(¦r«¬) As Boolean

If InStr(1, ¦r«¬, "¤¤¬ã°|¥Ò°©¤å") > 0 Then
    ¤¤¬ã°|¥Ò°©¤å = True
Else
    ¤¤¬ã°|¥Ò°©¤å = False
End If

End Function

Public Function ¤¤¬ã°|·¡¨t¤å¦r(¦r«¬) As Boolean

If InStr(1, ¦r«¬, "¤¤¬ã°|·¡¨tÂ²©­¤å¦r") > 0 Then
    ¤¤¬ã°|·¡¨t¤å¦r = True
Else
    ¤¤¬ã°|·¡¨t¤å¦r = False
End If

End Function


Public Function ¤¤¬ã°|·¢®Ñ(¦r«¬) As Boolean

If InStr(1, ¦r«¬, "¼Ğ·¢Åé") > 0 Then
    ¤¤¬ã°|·¢®Ñ = True
ElseIf InStr(1, ¦r«¬, "²Ó©úÅé") > 0 Then
    ¤¤¬ã°|·¢®Ñ = True
Else
    ¤¤¬ã°|·¢®Ñ = True = False
End If

End Function


Public Sub ¥X³BÂà¦rÅé(¥X³B)

¥X³B¬°¥Ò°©¤å = False
¥X³B¬°¥Ò°©¤å¦X¶° = False
¥X³B¬°ª÷¤å = False
¥X³B¬°¤p½f = False
¥X³B¬°·¡¤å¦r = False
¥X³B§¹¥ş¤Ç°t = False

If InStr(1, ¥X³B, "¦X¶°") > 0 Then
    ¥X³B¬°¥Ò°©¤å = True
    ¥X³B¬°¥Ò°©¤å¦X¶° = True
    If Len(¥X³B) > 2 Then ¥X³B§¹¥ş¤Ç°t = True
ElseIf InStr(1, ¥X³B, "¤Ù") > 0 Then
    ¥X³B¬°¥Ò°©¤å = True
ElseIf InStr(1, ¥X³B, "­^") > 0 Then
    ¥X³B¬°¥Ò°©¤å = True
    If Len(¥X³B) > 1 Then ¥X³B§¹¥ş¤Ç°t = True
ElseIf InStr(1, ¥X³B, "Ãh") > 0 Then
    ¥X³B¬°¥Ò°©¤å = True
    If Len(¥X³B) > 1 Then ¥X³B§¹¥ş¤Ç°t = True
ElseIf InStr(1, ¥X³B, "¶°¦¨") > 0 Then
    ¥X³B¬°ª÷¤å = True
    If Len(¥X³B) > 2 Then ¥X³B§¹¥ş¤Ç°t = True
ElseIf InStr(1, ¥X³B, "»¡¤å") > 0 Then
    ¥X³B¬°¤p½f = True
Else
    ¥X³B¬°·¡¤å¦r = True
End If

End Sub
