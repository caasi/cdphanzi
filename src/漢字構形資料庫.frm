VERSION 5.00
Begin VB.MDIForm mdiº~¦r¦r§Î 
   BackColor       =   &H8000000C&
   Caption         =   "º~¦rºc§Î¸ê®Æ®w(¤¤¥¡¬ã¨s°|¸ê°T¬ì¾Ç¬ã¨s©Ò)"
   ClientHeight    =   6936
   ClientLeft      =   168
   ClientTop       =   768
   ClientWidth     =   13584
   Icon            =   "º~¦rºc§Î¸ê®Æ®w.frx":0000
   LinkMode        =   1  '¨Ó·½
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox picª¬ºA¦C 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      BorderStyle     =   0  '¨S¦³®Ø½u
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   0
      ScaleHeight     =   468
      ScaleWidth      =   13584
      TabIndex        =   18
      Top             =   6468
      Width           =   13584
      Begin VB.TextBox txtª¬ºA 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "·s²Ó©úÅé"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   108
         TabIndex        =   19
         Top             =   84
         Width           =   4392
      End
   End
   Begin VB.PictureBox picºc¦r²Å¸¹ 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000004&
      ForeColor       =   &H80000004&
      Height          =   492
      Left            =   0
      ScaleHeight     =   468
      ScaleWidth      =   13560
      TabIndex        =   12
      Top             =   0
      Width           =   13584
      Begin VB.ComboBox cbo¹Ï¤ù¤j¤p 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         IntegralHeight  =   0   'False
         ItemData        =   "º~¦rºc§Î¸ê®Æ®w.frx":030A
         Left            =   8016
         List            =   "º~¦rºc§Î¸ê®Æ®w.frx":030C
         TabIndex        =   27
         ToolTipText     =   "¹Ï¤ù¤j¤p(¦r«¬ÂI¼Æ)"
         Top             =   44
         Width           =   855
      End
      Begin VB.ComboBox cbo¸ÑªR«× 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "º~¦rºc§Î¸ê®Æ®w.frx":030E
         Left            =   6924
         List            =   "º~¦rºc§Î¸ê®Æ®w.frx":0318
         TabIndex        =   26
         ToolTipText     =   "¹Ï¤ù¸ÑªR«×(dpi)"
         Top             =   44
         Width           =   1020
      End
      Begin VB.ComboBox cbo¦r«¬¦WºÙ 
         BeginProperty Font 
            Name            =   "¼Ð·¢Åé"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         ItemData        =   "º~¦rºc§Î¸ê®Æ®w.frx":032C
         Left            =   4548
         List            =   "º~¦rºc§Î¸ê®Æ®w.frx":0336
         TabIndex        =   20
         ToolTipText     =   "¦r«¬"
         Top             =   44
         Width           =   1368
      End
      Begin VB.ComboBox cbo¦rÅé¤j¤p 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         IntegralHeight  =   0   'False
         ItemData        =   "º~¦rºc§Î¸ê®Æ®w.frx":034A
         Left            =   6000
         List            =   "º~¦rºc§Î¸ê®Æ®w.frx":037E
         TabIndex        =   17
         ToolTipText     =   "Åã¥Ü¦r«¬¤j¤p"
         Top             =   44
         Width           =   855
      End
      Begin VB.ComboBox cbo­ºµ§ 
         BeginProperty Font 
            Name            =   "¼Ð·¢Åé"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         ItemData        =   "º~¦rºc§Î¸ê®Æ®w.frx":03C0
         Left            =   3456
         List            =   "º~¦rºc§Î¸ê®Æ®w.frx":03D6
         TabIndex        =   16
         ToolTipText     =   "­ºµ§"
         Top             =   44
         Width           =   972
      End
      Begin VB.ComboBox cboµ§µe 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         ItemData        =   "º~¦rºc§Î¸ê®Æ®w.frx":03F4
         Left            =   2376
         List            =   "º~¦rºc§Î¸ê®Æ®w.frx":03F6
         TabIndex        =   15
         ToolTipText     =   "µ§µe"
         Top             =   44
         Width           =   972
      End
      Begin VB.ComboBox cbo²Å¸¹ 
         BeginProperty Font 
            Name            =   "¼Ð·¢Åé"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Left            =   1536
         TabIndex        =   14
         ToolTipText     =   "ºc¦r²Å¸¹"
         Top             =   44
         Width           =   735
      End
      Begin VB.ComboBox cbo²Å¸¹Ãþ«¬ 
         BeginProperty Font 
            Name            =   "¼Ð·¢Åé"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         ItemData        =   "º~¦rºc§Î¸ê®Æ®w.frx":03F8
         Left            =   108
         List            =   "º~¦rºc§Î¸ê®Æ®w.frx":03FA
         TabIndex        =   13
         ToolTipText     =   "ºc¦r²Å¸¹Ãþ«¬"
         Top             =   44
         Width           =   1335
      End
   End
   Begin VB.PictureBox pic¦r§ÎÄÝ©Ê 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      Appearance      =   0  '¥­­±
      BackColor       =   &H80000004&
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000002&
      Height          =   492
      Left            =   0
      ScaleHeight     =   492
      ScaleWidth      =   13584
      TabIndex        =   0
      Top             =   492
      Width           =   13584
      Begin VB.TextBox txt­«¤å 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   1848
         TabIndex        =   25
         ToolTipText     =   "¥jº~¦r¦r«¬"
         Top             =   48
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   2772
         TabIndex        =   24
         ToolTipText     =   "µ§µe"
         Top             =   48
         Width           =   375
      End
      Begin VB.TextBox txt¥jº~¦r 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BeginProperty Font 
            Name            =   "¼Ð·¢Åé"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   2304
         TabIndex        =   23
         ToolTipText     =   "¥jº~¦r"
         Top             =   48
         Width           =   375
      End
      Begin VB.TextBox txt¥~¦r¶° 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   960
         TabIndex        =   22
         ToolTipText     =   "¥~¦r¶°"
         Top             =   44
         Width           =   375
      End
      Begin VB.TextBox txt½s¸¹ 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "½s¸¹"
         Top             =   44
         Width           =   732
      End
      Begin VB.TextBox txt¦r§Î 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BeginProperty Font 
            Name            =   "¼Ð·¢Åé"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   1404
         TabIndex        =   11
         ToolTipText     =   "¿é¤J¦r§Î«á¡A¦A«öEnter¬d¸ß¦r§ÎÄÝ©Ê"
         Top             =   44
         Width           =   375
      End
      Begin VB.TextBox txt¦©°£³¡­ºµ§µe 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   4176
         TabIndex        =   10
         ToolTipText     =   "µ§µe(³¡­º¤£­p)"
         Top             =   48
         Width           =   375
      End
      Begin VB.TextBox txtÁ`µ§µe 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   3228
         TabIndex        =   9
         ToolTipText     =   "µ§µe"
         Top             =   48
         Width           =   375
      End
      Begin VB.TextBox txt³¡­º 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "¼Ð·¢Åé"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   3708
         TabIndex        =   8
         ToolTipText     =   "³¡­º"
         Top             =   48
         Width           =   375
      End
      Begin VB.TextBox txt²Õ¦r¦r¼Æ§t²§¼g 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   13080
         TabIndex        =   7
         ToolTipText     =   "²Õ¦r¦r¼Æ(¥]§t²§Åé¦r®Ú)"
         Top             =   48
         Width           =   612
      End
      Begin VB.TextBox txt²Õ¦r¦r¼Æ 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   12372
         TabIndex        =   6
         ToolTipText     =   "²Õ¦r¦r¼Æ"
         Top             =   48
         Width           =   612
      End
      Begin VB.TextBox txt¥U¼Æ 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   11076
         TabIndex        =   5
         ToolTipText     =   "º~»y¤j¦r¨å¥U­¶¦r"
         Top             =   48
         Width           =   1212
      End
      Begin VB.TextBox txtª`­µ 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "¼Ð·¢Åé"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   4644
         TabIndex        =   4
         ToolTipText     =   "ª`­µ"
         Top             =   48
         Width           =   1095
      End
      Begin VB.TextBox txt¤º½X 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   5832
         TabIndex        =   3
         ToolTipText     =   "Big5"
         Top             =   48
         Width           =   735
      End
      Begin VB.TextBox txt­Ü¾e½X 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "¼Ð·¢Åé"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   6648
         TabIndex        =   2
         ToolTipText     =   "­Ü¾e½X"
         Top             =   48
         Width           =   1572
      End
      Begin VB.TextBox txtºc¦r¦¡ 
         BeginProperty Font 
            Name            =   "¼Ð·¢Åé"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Left            =   8340
         TabIndex        =   1
         Text            =   $"º~¦rºc§Î¸ê®Æ®w.frx":03FC
         ToolTipText     =   "¿é¤Jºc¦r¦¡«á¡A¦A«öEnter¬d¸ß¦r§ÎÄÝ©Ê"
         Top             =   48
         Width           =   2580
      End
   End
   Begin VB.Menu mnu_¦r¶° 
      Caption         =   "¦r¶°"
      Begin VB.Menu mnu_±`¥Î¦r 
         Caption         =   "±`¥Î¦r"
      End
      Begin VB.Menu mnu_Big5 
         Caption         =   "¤­¤j½X(Big5)"
      End
      Begin VB.Menu mnu_Â²¤Æ¦rÁ`ªí 
         Caption         =   "Â²¤Æ¦rÁ`ªí"
      End
      Begin VB.Menu mnu_º~»y¤j¦r¨å 
         Caption         =   "º~»y¤j¦r¨å"
      End
      Begin VB.Menu mnu_line1_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_¤p½f 
         Caption         =   "»¡¤å¸Ñ¦r"
      End
      Begin VB.Menu mnu_line1_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ª÷¤å 
         Caption         =   "ª÷¤å½s"
      End
      Begin VB.Menu mnu_ª÷¤å¹Ï§Î¤å¦r 
         Caption         =   "ª÷¤å½sªþ¿ý¤W(¹Ï§Î¤å¦r)"
      End
      Begin VB.Menu mnu_line1_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_¥Ò°©¤å 
         Caption         =   "®ï¼V¥Ò°©¨èÃãÃþÄ¡"
      End
      Begin VB.Menu mnu_·¡¨t¤å¦r 
         Caption         =   "·¡¨tÂ²©­¤å¦r½s"
      End
      Begin VB.Menu mnu_line1_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_·¢®Ñ 
         Caption         =   "¥H¤W©Ò¦³·¢Åé¦r"
      End
      Begin VB.Menu mnu_line1_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_µ²§ô 
         Caption         =   "µ²§ô"
      End
   End
   Begin VB.Menu mnu_¦r 
      Caption         =   "¦r§Î"
      Begin VB.Menu mnu_¦r§Î´F¨Å 
         Caption         =   "³¡¥óÀË¦r..."
      End
      Begin VB.Menu mnu_¥X³BÀË¦r 
         Caption         =   "¥X³BÀË¦r..."
      End
      Begin VB.Menu mnu_¦r§Îµ²ºc 
         Caption         =   "¦r§Îµ²ºc..."
      End
      Begin VB.Menu mnu_¦r§ÎºtÅÜ 
         Caption         =   "¦r§ÎºtÅÜ..."
      End
      Begin VB.Menu mnu_¦r§Î¯Á¤Þ 
         Caption         =   "¦r§Î¯Á¤Þ..."
      End
      Begin VB.Menu mnu_²§Åé¦rªí 
         Caption         =   "²§Åé¦rªí..."
      End
   End
   Begin VB.Menu mnu_³¡¥ó 
      Caption         =   "³¡¥ó"
      Begin VB.Menu mnu_±dº³¦r¨å³¡­º 
         Caption         =   "±dº³¦r¨å³¡­º..."
      End
      Begin VB.Menu mnu_»¡¤å¸Ñ¦r³¡­º 
         Caption         =   "»¡¤å¸Ñ¦r³¡­º..."
      End
      Begin VB.Menu mnu_line3_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_³¡¥ó¥N½X 
         Caption         =   "Big5¦r®Ú..."
         Index           =   1
      End
      Begin VB.Menu mnu_³¡¥ó¥N½X 
         Caption         =   "Big5¤ÎÂ²¤Æ¦r¦r®Ú..."
         Index           =   2
      End
      Begin VB.Menu mnu_³¡¥ó¥N½X 
         Caption         =   "¦r®Ú..."
         Index           =   3
      End
      Begin VB.Menu mnu_³¡¥ó¥N½X 
         Caption         =   "¤p½f¿WÅé¦r..."
         Index           =   4
      End
      Begin VB.Menu mnu_³¡¥ó¥N½X 
         Caption         =   "ª÷¤å¦r®Ú..."
         Index           =   5
      End
      Begin VB.Menu mnu_³¡¥ó¥N½X 
         Caption         =   "¥Ò°©¤å¦r®Ú..."
         Index           =   6
      End
      Begin VB.Menu mnu_³¡¥ó¥N½X 
         Caption         =   "·¡¨tÂ²©­¤å¦r¦r®Ú..."
         Index           =   7
      End
      Begin VB.Menu mnu_³¡¥ó¥N½X 
         Caption         =   "³¡¥ó¥~¦r..."
         Index           =   8
      End
      Begin VB.Menu mnuline3_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_²§Åé¦r®Ú 
         Caption         =   "²§Åé¦r®Ú..."
      End
   End
   Begin VB.Menu mnu_²Å¸¹ 
      Caption         =   "²Å¸¹"
      Begin VB.Menu mnu_ºc¦r²Å¸¹ 
         Caption         =   "ºc¦r²Å¸¹..."
      End
      Begin VB.Menu mnu_¹Ï§Î¤å¦r 
         Caption         =   "¹Ï§Î¤å¦r(¥¼·¢¤Æ)..."
      End
      Begin VB.Menu mnu_¤K¨ö 
         Caption         =   "¤K¨ö..."
      End
      Begin VB.Menu mnu_Â²Ã| 
         Caption         =   "Â²Ã|..."
      End
   End
   Begin VB.Menu mnu_¦r¼Ë 
      Caption         =   "¦r¼Ë"
      Visible         =   0   'False
      Begin VB.Menu mnu_°Ñ¦Ò¦r¼Ë 
         Caption         =   "°Ñ¦Ò¦r¼Ë..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_¦rÅé­·®æ 
         Caption         =   "¦rÅé­·®æ..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu½s¿è 
      Caption         =   "½s¿è"
      Begin VB.Menu mnu_½Æ»s 
         Caption         =   "½Æ»s"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu_½Æ»s¹Ï¤ù 
         Caption         =   "½Æ»s¹Ï¤ù"
      End
      Begin VB.Menu mnu_½Æ»s¯S®í¹Ï¤ù 
         Caption         =   "½Æ»s¯S®í¹Ï¤ù"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_edit_½Æ»s¨ìWord 
         Caption         =   "½Æ»s¨ìMicrosoft Word"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnu¶K¤W 
         Caption         =   "¶K¤W"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnu_line5_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_¦r«¬ 
         Caption         =   "³]©wÅã¥Ü¦r«¬..."
      End
      Begin VB.Menu mnu_¹Ï¤ù 
         Caption         =   "³]©w½Æ»s¹Ï¤ù..."
      End
   End
   Begin VB.Menu mnu_Tool 
      Caption         =   "¤u¨ã"
      Begin VB.Menu mnu_Tool_ListLikeChar 
         Caption         =   "¦C¥X¬Û¦P¥X³Bªº¦r§Î"
      End
      Begin VB.Menu mnuToolListChar 
         Caption         =   "¦C¥X©Ò¦³¦r§Î"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_¿ï¶µ 
      Caption         =   "¿ï¶µ"
      Begin VB.Menu mnu_³¡¥óÀË¦r¿ï¶µ 
         Caption         =   "³¡¥óÀË¦r"
         Begin VB.Menu mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó 
            Caption         =   "³¡¥óÀË¦r³v¯Å¦C¥X³¡¥ó»P¦r§Î(­­³æ¤@³¡¥ó)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_¦r§Î´F¨Å¥u¦C¦r§Î¤£¦C³¡¥ó 
            Caption         =   "³¡¥óÀË¦r¥u¦C¦r§Î¤£¦C³¡¥ó"
         End
         Begin VB.Menu mnu_line6_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_¦r§Î´F¨Å¿í·Ó¿é¤J³¡¥ó¶¶§Ç 
            Caption         =   "³¡¥óÀË¦r¿í·Ó¿é¤J³¡¥ó¶¶§Ç"
         End
         Begin VB.Menu mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó 
            Caption         =   "³¡¥óÀË¦r¥]§t²§Åé¦r®Ú"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk 
            Caption         =   "³¡¥óÀË¦r±Ä¥ÎSQL Like»yªk"
         End
      End
      Begin VB.Menu mnu_¦r§Îµ²ºc¿ï¶µ 
         Caption         =   "¦r§Îµ²ºc"
         Begin VB.Menu mnu_»¡¤å¸Ñ§Î¶¶§Ç 
            Caption         =   "¦r§Îµ²ºc¨Ì¾Ú»¡¤å¸Ñ§Î¶¶§Ç¦C¥X¤p½f³¡¥ó"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu_¦r§Î¯Á¤Þ¿ï¶µ 
         Caption         =   "¦r§Î¯Á¤Þ"
         Begin VB.Menu mnu_»·ªFº~»y¤j¦r¨å¿ï¶µ 
            Caption         =   "º~»y¤j¦r¨å(»·ªF¹Ï®Ñ¤½¥q)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_«Ø§»º~»y¤j¦r¨å¿ï¶µ 
            Caption         =   "º~»y¤j¦r¨å(«Ø§»¥Xª©ªÀ)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_line6_2 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_¤¤¤å¤jÃã¨å¿ï¶µ 
            Caption         =   "¤¤¤å¤jÃã¨å"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_line6_3 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_»¡¤å¸Ñ¦rµþªL¿ï¶µ 
            Caption         =   "»¡¤å¸Ñ¦rµþªL"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_¤¤µØ»¡¤å¸Ñ¦r¿ï¶µ 
            Caption         =   "»¡¤å¸Ñ¦r(¤¤µØ®Ñ§½)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_line6_4 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_ª÷¤å½s¿ï¶µ 
            Caption         =   "ª÷¤å½s"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_ª÷¤åµþªL¿ï¶µ 
            Caption         =   "ª÷¤åµþªL"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_®ï©Pª÷¤å¶°¦¨¾¹¸¹¿ï¶µ 
            Caption         =   "®ï©Pª÷¤å¶°¦¨¾¹¸¹"
            Checked         =   -1  'True
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_®ï©Pª÷¤å¶°¦¨¤Þ±o¿ï¶µ 
            Caption         =   "®ï©Pª÷¤å¶°¦¨¤Þ±o"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_line6_5 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_¥Ò°©¨èÃãÃþÄ¡¿ï¶µ 
            Caption         =   "®ï¼V¥Ò°©¨èÃãÃþÄ¡"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_¥Ò°©¤å¦rµþªL¿ï¶µ 
            Caption         =   "¥Ò°©¤å¦rµþªL"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_¥Ò°©¤å¦r¶°ÄÀ¿ï¶µ 
            Caption         =   "¥Ò°©¤å¦r¶°ÄÀ"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_line6_6 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_·¡¨tÂ²©­¤å¦r½s¿ï¶µ 
            Caption         =   "·¡¨tÂ²©­¤å¦r½s"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_·¡¨t¤å¦r¥X³B¿ï¶µ 
            Caption         =   "·¡¨t¤å¦r¥X³B"
            Checked         =   -1  'True
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_line6_7 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_Unicode¿ï¶µ 
            Caption         =   "Unicode"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_Big5¿ï¶µ 
            Caption         =   "Big5"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu_¦r§ÎºtÅÜ¿ï¶µ 
         Caption         =   "¦r§ÎºtÅÜ"
         Begin VB.Menu mnu_¥Ò°©¤å¿ï¶µ 
            Caption         =   "¥Ò°©¤å"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_ª÷¤å¿ï¶µ 
            Caption         =   "ª÷¤å"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_·¡¨t¤å¦r¿ï¶µ 
            Caption         =   "·¡¨t¤å¦r"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_¤p½f¿ï¶µ 
            Caption         =   "¤p½f"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu_option_½Æ»s¨ìWord 
         Caption         =   "½Æ»s¨ìMicrosoft Word"
         Begin VB.Menu mnu_½Æ»s¦r§Î¨ìWord 
            Caption         =   "½Æ»s¦¨¦r§Î©Îºc¦r¦¡"
         End
         Begin VB.Menu mnu_½Æ»s¹Ï¤ù¨ìWord 
            Caption         =   "½Æ»s¦¨¹Ï¤ù"
         End
         Begin VB.Menu mnu_½Æ»sUnicode¦r§Î¨ìWord 
            Caption         =   "½Æ»sUnicode¦r§Î"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_´¼¼z«¬½Æ»s¨ìWord 
            Caption         =   "´¼¼z«¬½Æ»s"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu_½Æ»s¿ï¶µ 
         Caption         =   "¨ä¥L"
         Begin VB.Menu mnu_½Æ»s¨ì°Å¶KÃ¯ 
            Caption         =   "¦Û°Ê½Æ»s¿ï¨ú¦r§Î¨ì°Å¶KÃ¯"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_Åã¥Ü­·®æ½X 
            Caption         =   "Åã¥Ü­·®æ½X"
            Checked         =   -1  'True
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnu_Àx¦sµøµ¡³]©w 
         Caption         =   "µ²§ô®ÉÀx¦sµøµ¡³]©w"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_¦P®É³]©w©Ò¦³¶}±Òµøµ¡ªº¦r«¬¤j¤p¤ÎÃC¦â 
         Caption         =   "¦P®É³]©w©Ò¦³¶}±Òµøµ¡ªº¦r«¬¤j¤p¤ÎÃC¦â"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_µøµ¡ 
      Caption         =   "µøµ¡"
      WindowList      =   -1  'True
      Begin VB.Menu mnu_¤ô¥­¨Ã±Æ 
         Caption         =   "¤ô¥­¨Ã±Æ"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_««ª½¨Ã±Æ 
         Caption         =   "««ª½¨Ã±Æ"
      End
      Begin VB.Menu mnu_­«Å|Åã¥Ü 
         Caption         =   "­«Å|Åã¥Ü"
      End
      Begin VB.Menu mnu_·s¼Wµøµ¡ 
         Caption         =   "·s¼Wµøµ¡"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_±Æ¦C¹Ï¥Ü 
         Caption         =   "±Æ¦C¹Ï¥Ü"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Â²©öÂsÄý 
         Caption         =   "¹w³]¶}±Ò(°ò¥»)..."
      End
      Begin VB.Menu mnu_³Ì¨ÎÂsÄý 
         Caption         =   "¹w³]¶}±Ò(¶i¶¥)..."
      End
   End
   Begin VB.Menu mnu_»¡©ú 
      Caption         =   "»¡©ú"
      Begin VB.Menu mnu_»¡©ú¥DÃD 
         Caption         =   "»¡©ú¥DÃD"
      End
      Begin VB.Menu mnu_line8_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_cdp 
         Caption         =   "Ãö©óº~¦rºc§Î¸ê®Æ®w"
      End
   End
End
Attribute VB_Name = "mdiº~¦r¦r§Î"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private µ§µe As Integer, ­ºµ§ As Integer, ª¬ºA¦C As String, ª¬ºA¦C1 As String
Private ªì©lfont As String, ªì©lfontsize As Integer
Private ªì©lleft As Integer, ªì©ltop As Integer, ªì©lwidth As Integer, ªì©lheight As Integer
Private ªì©lsave As String * 250

Private path As String

Private Sub ¸ü¤Jªì©l­È()
Dim nDefault As Long, sDefault As String, lret As Long

ªì©lfirst = GetPrivateProfileInt("Start", "first", nDefault, App.path & "\cdphanzi.ini")

ªì©lleft = GetPrivateProfileInt("Window", "left", nDefault, App.path & "\cdphanzi.ini")
ªì©ltop = GetPrivateProfileInt("Window", "top", nDefault, App.path & "\cdphanzi.ini")
ªì©lwidth = GetPrivateProfileInt("Window", "width", nDefault, App.path & "\cdphanzi.ini")
ªì©lheight = GetPrivateProfileInt("Window", "height", nDefault, App.path & "\cdphanzi.ini")

¦r¶°_±`¥Î¦r = GetPrivateProfileInt("¦r¶°", "±`¥Î¦r", nDefault, App.path & "\cdphanzi.ini")
¦r¶°_¤­¤j½X = GetPrivateProfileInt("¦r¶°", "¤­¤j½X", nDefault, App.path & "\cdphanzi.ini")
¦r¶°_Â²¤Æ¦r = GetPrivateProfileInt("¦r¶°", "Â²¤Æ¦r", nDefault, App.path & "\cdphanzi.ini")
¦r¶°_º~»y¤j¦r¨å = GetPrivateProfileInt("¦r¶°", "º~»y¤j¦r¨å", nDefault, App.path & "\cdphanzi.ini")
¦r¶°_»¡¤å¸Ñ¦r = GetPrivateProfileInt("¦r¶°", "»¡¤å¸Ñ¦r", nDefault, App.path & "\cdphanzi.ini")
¦r¶°_ª÷¤å½s = GetPrivateProfileInt("¦r¶°", "ª÷¤å½s", nDefault, App.path & "\cdphanzi.ini")
¦r¶°_ª÷¤å½s¹Ï§Î¤å¦r = GetPrivateProfileInt("¦r¶°", "ª÷¤å½s¹Ï§Î¤å¦r", nDefault, App.path & "\cdphanzi.ini")
¦r¶°_¥Ò°©ÃþÄ¡ = GetPrivateProfileInt("¦r¶°", "¥Ò°©ÃþÄ¡", nDefault, App.path & "\cdphanzi.ini")
¦r¶°_·¡¨tÂ²©­¤å¦r½s = GetPrivateProfileInt("¦r¶°", "·¡¨tÂ²©­¤å¦r½s", nDefault, App.path & "\cdphanzi.ini")
¦r¶°_·¢Åé¦r = GetPrivateProfileInt("¦r¶°", "·¢Åé¦r", nDefault, App.path & "\cdphanzi.ini")

´F¨Åopen = GetPrivateProfileInt("¦r§Î´F¨Å", "open", nDefault, App.path & "\cdphanzi.ini")
´F¨Åwinstate = GetPrivateProfileInt("¦r§Î´F¨Å", "winstate", nDefault, App.path & "\cdphanzi.ini")
´F¨Åleft = GetPrivateProfileInt("¦r§Î´F¨Å", "left", nDefault, App.path & "\cdphanzi.ini")
´F¨Åtop = GetPrivateProfileInt("¦r§Î´F¨Å", "top", nDefault, App.path & "\cdphanzi.ini")
´F¨Åwidth = GetPrivateProfileInt("¦r§Î´F¨Å", "width", nDefault, App.path & "\cdphanzi.ini")
´F¨Åheight = GetPrivateProfileInt("¦r§Î´F¨Å", "height", nDefault, App.path & "\cdphanzi.ini")

¥X³Bopen = GetPrivateProfileInt("¥X³BÀË¦r", "open", nDefault, App.path & "\cdphanzi.ini")
¥X³Bwinstate = GetPrivateProfileInt("¥X³BÀË¦r", "winstate", nDefault, App.path & "\cdphanzi.ini")
¥X³Bleft = GetPrivateProfileInt("¥X³BÀË¦r", "left", nDefault, App.path & "\cdphanzi.ini")
¥X³Btop = GetPrivateProfileInt("¥X³BÀË¦r", "top", nDefault, App.path & "\cdphanzi.ini")
¥X³Bwidth = GetPrivateProfileInt("¥X³BÀË¦r", "width", nDefault, App.path & "\cdphanzi.ini")
¥X³Bheight = GetPrivateProfileInt("¥X³BÀË¦r", "height", nDefault, App.path & "\cdphanzi.ini")

µ²ºcopen = GetPrivateProfileInt("¦r§Îµ²ºc", "open", nDefault, App.path & "\cdphanzi.ini")
µ²ºcwinstate = GetPrivateProfileInt("¦r§Îµ²ºc", "winstate", nDefault, App.path & "\cdphanzi.ini")
µ²ºcleft = GetPrivateProfileInt("¦r§Îµ²ºc", "left", nDefault, App.path & "\cdphanzi.ini")
µ²ºctop = GetPrivateProfileInt("¦r§Îµ²ºc", "top", nDefault, App.path & "\cdphanzi.ini")
µ²ºcwidth = GetPrivateProfileInt("¦r§Îµ²ºc", "width", nDefault, App.path & "\cdphanzi.ini")
µ²ºcheight = GetPrivateProfileInt("¦r§Îµ²ºc", "height", nDefault, App.path & "\cdphanzi.ini")

²§Åéopen = GetPrivateProfileInt("²§Åé¦rªí", "open", nDefault, App.path & "\cdphanzi.ini")
²§Åéwinstate = GetPrivateProfileInt("²§Åé¦rªí", "winstate", nDefault, App.path & "\cdphanzi.ini")
²§Åéleft = GetPrivateProfileInt("²§Åé¦rªí", "left", nDefault, App.path & "\cdphanzi.ini")
²§Åétop = GetPrivateProfileInt("²§Åé¦rªí", "top", nDefault, App.path & "\cdphanzi.ini")
²§Åéwidth = GetPrivateProfileInt("²§Åé¦rªí", "width", nDefault, App.path & "\cdphanzi.ini")
²§Åéheight = GetPrivateProfileInt("²§Åé¦rªí", "height", nDefault, App.path & "\cdphanzi.ini")

²§®Úopen = GetPrivateProfileInt("²§Åé¦r®Ú", "open", nDefault, App.path & "\cdphanzi.ini")
²§®Úwinstate = GetPrivateProfileInt("²§Åé¦r®Ú", "winstate", nDefault, App.path & "\cdphanzi.ini")
²§®Úleft = GetPrivateProfileInt("²§Åé¦r®Ú", "left", nDefault, App.path & "\cdphanzi.ini")
²§®Útop = GetPrivateProfileInt("²§Åé¦r®Ú", "top", nDefault, App.path & "\cdphanzi.ini")
²§®Úwidth = GetPrivateProfileInt("²§Åé¦r®Ú", "width", nDefault, App.path & "\cdphanzi.ini")
²§®Úheight = GetPrivateProfileInt("²§Åé¦r®Ú", "height", nDefault, App.path & "\cdphanzi.ini")

³¡¥óopen = GetPrivateProfileInt("ºc§Î³¡¥ó", "open", nDefault, App.path & "\cdphanzi.ini")
³¡¥ówinstate = GetPrivateProfileInt("ºc§Î³¡¥ó", "winstate", nDefault, App.path & "\cdphanzi.ini")
³¡¥óleft = GetPrivateProfileInt("ºc§Î³¡¥ó", "left", nDefault, App.path & "\cdphanzi.ini")
³¡¥ótop = GetPrivateProfileInt("ºc§Î³¡¥ó", "top", nDefault, App.path & "\cdphanzi.ini")
³¡¥ówidth = GetPrivateProfileInt("ºc§Î³¡¥ó", "width", nDefault, App.path & "\cdphanzi.ini")
³¡¥óheight = GetPrivateProfileInt("ºc§Î³¡¥ó", "height", nDefault, App.path & "\cdphanzi.ini")

ºtÅÜopen = GetPrivateProfileInt("¦r§ÎºtÅÜ", "open", nDefault, App.path & "\cdphanzi.ini")
ºtÅÜwinstate = GetPrivateProfileInt("¦r§ÎºtÅÜ", "winstate", nDefault, App.path & "\cdphanzi.ini")
ºtÅÜleft = GetPrivateProfileInt("¦r§ÎºtÅÜ", "left", nDefault, App.path & "\cdphanzi.ini")
ºtÅÜtop = GetPrivateProfileInt("¦r§ÎºtÅÜ", "top", nDefault, App.path & "\cdphanzi.ini")
ºtÅÜwidth = GetPrivateProfileInt("¦r§ÎºtÅÜ", "width", nDefault, App.path & "\cdphanzi.ini")
ºtÅÜheight = GetPrivateProfileInt("¦r§ÎºtÅÜ", "height", nDefault, App.path & "\cdphanzi.ini")

¯Á¤Þopen = GetPrivateProfileInt("¦r§Î¯Á¤Þ", "open", nDefault, App.path & "\cdphanzi.ini")
¯Á¤Þwinstate = GetPrivateProfileInt("¦r§Î¯Á¤Þ", "winstate", nDefault, App.path & "\cdphanzi.ini")
¯Á¤Þleft = GetPrivateProfileInt("¦r§Î¯Á¤Þ", "left", nDefault, App.path & "\cdphanzi.ini")
¯Á¤Þtop = GetPrivateProfileInt("¦r§Î¯Á¤Þ", "top", nDefault, App.path & "\cdphanzi.ini")
¯Á¤Þwidth = GetPrivateProfileInt("¦r§Î¯Á¤Þ", "width", nDefault, App.path & "\cdphanzi.ini")
¯Á¤Þheight = GetPrivateProfileInt("¦r§Î¯Á¤Þ", "height", nDefault, App.path & "\cdphanzi.ini")

ªì©lfont = String(256, 0)
lret = GetPrivateProfileString("Font", "fontname", "¼Ð·¢Åé", ªì©lfont, Len(ªì©lfont), App.path & "\cdphanzi.ini")
ªì©lfont = Left(ªì©lfont, InStr(ªì©lfont, Chr(0)) - 1)
ªì©lfontsize = GetPrivateProfileInt("Font", "fontsize", nDefault, App.path & "\cdphanzi.ini")

¹Ï¤ù¸ÑªR«× = GetPrivateProfileInt("Image", "dpi", nDefault, App.path & "\cdphanzi.ini")
¹Ï¤ù¦r«¬¤j¤p = GetPrivateProfileInt("Image", "fontsize", nDefault, App.path & "\cdphanzi.ini")

ªì©l³v¯Å¦C¥X = GetPrivateProfileInt("¿ï¶µ", "³v¯Å¦C¥X", nDefault, App.path & "\cdphanzi.ini")
ªì©l³¡¥ó¶¶§Ç = GetPrivateProfileInt("¿ï¶µ", "³¡¥ó¶¶§Ç", nDefault, App.path & "\cdphanzi.ini")
ªì©l²§¼g³¡¥ó = GetPrivateProfileInt("¿ï¶µ", "²§¼g³¡¥ó", nDefault, App.path & "\cdphanzi.ini")
ªì©l¸Ñ§Î¦C¥X = GetPrivateProfileInt("¿ï¶µ", "¸Ñ§Î¦C¥X", nDefault, App.path & "\cdphanzi.ini")
ªì©l»·ªFº~»y¤j¦r¨å = GetPrivateProfileInt("¿ï¶µ", "»·ªFº~»y¤j¦r¨å", nDefault, App.path & "\cdphanzi.ini")
ªì©l«Ø§»º~»y¤j¦r¨å = GetPrivateProfileInt("¿ï¶µ", "«Ø§»º~»y¤j¦r¨å", nDefault, App.path & "\cdphanzi.ini")
ªì©l¤¤¤å¤jÃã¨å = GetPrivateProfileInt("¿ï¶µ", "¤¤¤å¤jÃã¨å", nDefault, App.path & "\cdphanzi.ini")
ªì©l»¡¤å¸Ñ¦rµþªL = GetPrivateProfileInt("¿ï¶µ", "»¡¤å¸Ñ¦rµþªL", nDefault, App.path & "\cdphanzi.ini")
ªì©l¤¤µØ»¡¤å¸Ñ¦r = GetPrivateProfileInt("¿ï¶µ", "¤¤µØ»¡¤å¸Ñ¦r", nDefault, App.path & "\cdphanzi.ini")
ªì©lª÷¤å½s = GetPrivateProfileInt("¿ï¶µ", "ª÷¤å½s", nDefault, App.path & "\cdphanzi.ini")
ªì©lª÷¤åµþªL = GetPrivateProfileInt("¿ï¶µ", "ª÷¤åµþªL", nDefault, App.path & "\cdphanzi.ini")
'ªì©lª÷¤å¾¹¸¹ = GetPrivateProfileInt("¿ï¶µ", "ª÷¤å¾¹¸¹", nDefault, App.path & "\cdphanzi.ini")
ªì©lª÷¤å¤Þ±o = GetPrivateProfileInt("¿ï¶µ", "ª÷¤å¤Þ±o", nDefault, App.path & "\cdphanzi.ini")
ªì©l¥Ò°©¨èÃãÃþÄ¡ = GetPrivateProfileInt("¿ï¶µ", "¥Ò°©¨èÃãÃþÄ¡", nDefault, App.path & "\cdphanzi.ini")
ªì©l¥Ò°©¤å¦rµþªL = GetPrivateProfileInt("¿ï¶µ", "¥Ò°©¤å¦rµþªL", nDefault, App.path & "\cdphanzi.ini")
ªì©l¥Ò°©¤å¦r¶°ÄÀ = GetPrivateProfileInt("¿ï¶µ", "¥Ò°©¤å¦r¶°ÄÀ", nDefault, App.path & "\cdphanzi.ini")
ªì©l·¡¨tÂ²©­¤å¦r½s = GetPrivateProfileInt("¿ï¶µ", "·¡¨tÂ²©­¤å¦r½s", nDefault, App.path & "\cdphanzi.ini")
'ªì©l·¡¨t¤å¦r¥X³B = GetPrivateProfileInt("¿ï¶µ", "·¡¨t¤å¦r¥X³B", nDefault, App.path & "\cdphanzi.ini")
ªì©lUnicode = GetPrivateProfileInt("¿ï¶µ", "Unicode", nDefault, App.path & "\cdphanzi.ini")
ªì©lBig5 = GetPrivateProfileInt("¿ï¶µ", "Big5", nDefault, App.path & "\cdphanzi.ini")
ªì©l¥Ò°©¤åºtÅÜ = GetPrivateProfileInt("¿ï¶µ", "¥Ò°©¤åºtÅÜ", nDefault, App.path & "\cdphanzi.ini")
ªì©lª÷¤åºtÅÜ = GetPrivateProfileInt("¿ï¶µ", "ª÷¤åºtÅÜ", nDefault, App.path & "\cdphanzi.ini")
ªì©l·¡¨t¤å¦rºtÅÜ = GetPrivateProfileInt("¿ï¶µ", "·¡¨t¤å¦rºtÅÜ", nDefault, App.path & "\cdphanzi.ini")
ªì©l¤p½fºtÅÜ = GetPrivateProfileInt("¿ï¶µ", "¤p½fºtÅÜ", nDefault, App.path & "\cdphanzi.ini")
ªì©lcopy = GetPrivateProfileInt("¿ï¶µ", "copy", nDefault, App.path & "\cdphanzi.ini")
ªì©l¦rÀW = GetPrivateProfileInt("Window", "¦rÀW", nDefault, App.path & "\cdphanzi.ini")
ªì©lsave = GetPrivateProfileInt("¿ï¶µ", "save", nDefault, App.path & "\cdphanzi.ini")
'ªì©l­·®æ½X = GetPrivateProfileInt("¿ï¶µ", "­·®æ½X", nDefault, App.path & "\cdphanzi.ini")

ªì©lCopyToWord = GetPrivateProfileInt("CopyToWord", "CopyMode", nDefault, App.path & "\cdphanzi.ini")
ªì©lCopyUnicode = GetPrivateProfileInt("CopyToWord", "CopyUnicode", nDefault, App.path & "\cdphanzi.ini")

End Sub

Private Sub ¶}±Ò¸ê®Æ®w()

Set ¨t²Î¸ê®Æ®w = OpenDatabase(App.path & "\cdphanzi.mdb")
Set ¤p½f¸ê®Æ®w = OpenDatabase(App.path & "\cdpseal.mdb")
Set ª÷¤å¸ê®Æ®w = OpenDatabase(App.path & "\cdpbronz.mdb")
Set ¥Ò°©¤å¸ê®Æ®w = OpenDatabase(App.path & "\cdpjiagu.mdb")
Set ·¡¨t¤å¦r¸ê®Æ®w = OpenDatabase(App.path & "\cdpchubs.mdb")

¦rÅé¦WºÙ

End Sub

Private Sub Àx¦sµ²§ô­È()
Dim IniEntry As String * 255
Dim fsuccess As Boolean

If mdiº~¦r¦r§Î.mnu_Àx¦sµøµ¡³]©w.Checked = True Then
   ªì©lsave = 1
Else
   ªì©lsave = 0
   Exit Sub
End If

ªì©lfirst = 2

ªì©lleft = mdiº~¦r¦r§Î.Left
ªì©ltop = mdiº~¦r¦r§Î.Top
ªì©lheight = mdiº~¦r¦r§Î.Height
ªì©lwidth = mdiº~¦r¦r§Î.Width

If mnu_±`¥Î¦r.Checked Then
    ¦r¶°_±`¥Î¦r = 1
Else
    ¦r¶°_±`¥Î¦r = 0
End If

If mnu_Big5.Checked Then
    ¦r¶°_¤­¤j½X = 1
Else
    ¦r¶°_¤­¤j½X = 0
End If

If mnu_Â²¤Æ¦rÁ`ªí.Checked Then
    ¦r¶°_Â²¤Æ¦r = 1
Else
    ¦r¶°_Â²¤Æ¦r = 0
End If

If mnu_º~»y¤j¦r¨å.Checked Then
    ¦r¶°_º~»y¤j¦r¨å = 1
Else
    ¦r¶°_º~»y¤j¦r¨å = 0
End If

If mnu_¤p½f.Checked Then
    ¦r¶°_»¡¤å¸Ñ¦r = 1
Else
    ¦r¶°_»¡¤å¸Ñ¦r = 0
End If

If mnu_ª÷¤å.Checked Then
    ¦r¶°_ª÷¤å½s = 1
Else
    ¦r¶°_ª÷¤å½s = 0
End If

If mnu_ª÷¤å¹Ï§Î¤å¦r.Checked Then
    ¦r¶°_ª÷¤å½s¹Ï§Î¤å¦r = 1
Else
    ¦r¶°_ª÷¤å½s¹Ï§Î¤å¦r = 0
End If

If mnu_¥Ò°©¤å.Checked Then
    ¦r¶°_¥Ò°©ÃþÄ¡ = 1
Else
    ¦r¶°_¥Ò°©ÃþÄ¡ = 0
End If

If mnu_·¡¨t¤å¦r.Checked Then
    ¦r¶°_·¡¨tÂ²©­¤å¦r½s = 1
Else
    ¦r¶°_·¡¨tÂ²©­¤å¦r½s = 0
End If

If mnu_·¢®Ñ.Checked Then
    ¦r¶°_·¢Åé¦r = 1
Else
    ¦r¶°_·¢Åé¦r = 0
End If
    
If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Checked = True Then
   ªì©l³v¯Å¦C¥X = 1
Else
   ªì©l³v¯Å¦C¥X = 0
End If

If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å¿í·Ó¿é¤J³¡¥ó¶¶§Ç.Checked = True Then
   ªì©l³¡¥ó¶¶§Ç = 1
Else
   ªì©l³¡¥ó¶¶§Ç = 0
End If

If mdiº~¦r¦r§Î.mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó.Checked = True Then
   ªì©l²§¼g³¡¥ó = 1
Else
   ªì©l²§¼g³¡¥ó = 0
End If

If mdiº~¦r¦r§Î.mnu_»¡¤å¸Ñ§Î¶¶§Ç.Checked = True Then
    ªì©l¸Ñ§Î¦C¥X = 1
Else
    ªì©l¸Ñ§Î¦C¥X = 0
End If

If mnu_»·ªFº~»y¤j¦r¨å¿ï¶µ.Checked = True Then
   ªì©l»·ªFº~»y¤j¦r¨å = 1
Else
   ªì©l»·ªFº~»y¤j¦r¨å = 0
End If

If mnu_«Ø§»º~»y¤j¦r¨å¿ï¶µ.Checked = True Then
   ªì©l«Ø§»º~»y¤j¦r¨å = 1
Else
   ªì©l«Ø§»º~»y¤j¦r¨å = 0
End If

If mnu_¤¤¤å¤jÃã¨å¿ï¶µ.Checked = True Then
   ªì©l¤¤¤å¤jÃã¨å = 1
Else
   ªì©l¤¤¤å¤jÃã¨å = 0
End If

If mnu_»¡¤å¸Ñ¦rµþªL¿ï¶µ.Checked = True Then
   ªì©l»¡¤å¸Ñ¦rµþªL = 1
Else
   ªì©l»¡¤å¸Ñ¦rµþªL = 0
End If

If mnu_¤¤µØ»¡¤å¸Ñ¦r¿ï¶µ.Checked = True Then
   ªì©l¤¤µØ»¡¤å¸Ñ¦r = 1
Else
   ªì©l¤¤µØ»¡¤å¸Ñ¦r = 0
End If

If mnu_ª÷¤å½s¿ï¶µ.Checked = True Then
   ªì©lª÷¤å½s = 1
Else
   ªì©lª÷¤å½s = 0
End If

If mnu_ª÷¤åµþªL¿ï¶µ.Checked = True Then
   ªì©lª÷¤åµþªL = 1
Else
   ªì©lª÷¤åµþªL = 0
End If

If mnu_®ï©Pª÷¤å¶°¦¨¾¹¸¹¿ï¶µ.Checked = True Then
   ªì©lª÷¤å¾¹¸¹ = 1
Else
   ªì©lª÷¤å¾¹¸¹ = 0
End If

If mnu_®ï©Pª÷¤å¶°¦¨¤Þ±o¿ï¶µ.Checked = True Then
   ªì©lª÷¤å¤Þ±o = 1
Else
   ªì©lª÷¤å¤Þ±o = 0
End If

If mnu_¥Ò°©¨èÃãÃþÄ¡¿ï¶µ.Checked = True Then
   ªì©l¥Ò°©¨èÃãÃþÄ¡ = 1
Else
   ªì©l¥Ò°©¨èÃãÃþÄ¡ = 0
End If

If mnu_¥Ò°©¤å¦rµþªL¿ï¶µ.Checked = True Then
   ªì©l¥Ò°©¤å¦rµþªL = 1
Else
   ªì©l¥Ò°©¤å¦rµþªL = 0
End If

If mnu_¥Ò°©¤å¦r¶°ÄÀ¿ï¶µ.Checked = True Then
   ªì©l¥Ò°©¤å¦r¶°ÄÀ = 1
Else
   ªì©l¥Ò°©¤å¦r¶°ÄÀ = 0
End If

If mnu_·¡¨tÂ²©­¤å¦r½s¿ï¶µ.Checked = True Then
   ªì©l·¡¨tÂ²©­¤å¦r½s = 1
Else
   ªì©l·¡¨tÂ²©­¤å¦r½s = 0
End If

If mnu_·¡¨t¤å¦r¥X³B¿ï¶µ.Checked = True Then
   ªì©l·¡¨t¤å¦r¥X³B = 1
Else
   ªì©l·¡¨t¤å¦r¥X³B = 0
End If

If mnu_Unicode¿ï¶µ.Checked = True Then
   ªì©lUnicode = 1
Else
   ªì©lUnicode = 0
End If

If mnu_Big5¿ï¶µ.Checked = True Then
   ªì©lBig5 = 1
Else
   ªì©lBig5 = 0
End If

If mnu_¥Ò°©¤å¿ï¶µ.Checked = True Then
   ªì©l¥Ò°©¤åºtÅÜ = 1
Else
   ªì©l¥Ò°©¤åºtÅÜ = 0
End If

If mnu_ª÷¤å¿ï¶µ.Checked = True Then
   ªì©lª÷¤åºtÅÜ = 1
Else
   ªì©lª÷¤åºtÅÜ = 0
End If

If mnu_·¡¨t¤å¦r¿ï¶µ.Checked = True Then
   ªì©l·¡¨t¤å¦rºtÅÜ = 1
Else
   ªì©l·¡¨t¤å¦rºtÅÜ = 0
End If

If mnu_¤p½f¿ï¶µ.Checked = True Then
   ªì©l¤p½fºtÅÜ = 1
Else
   ªì©l¤p½fºtÅÜ = 0
End If

ªì©lfont = mdiº~¦r¦r§Î.cbo¦r«¬¦WºÙ.Text
ªì©lfontsize = mdiº~¦r¦r§Î.cbo¦rÅé¤j¤p.Text

If mdiº~¦r¦r§Î.mnu_½Æ»s¨ì°Å¶KÃ¯.Checked = True Then
   ªì©lcopy = 1
Else
   ªì©lcopy = 0
End If

If mdiº~¦r¦r§Î.mnu_Åã¥Ü­·®æ½X.Checked = True Then
   ªì©l­·®æ½X = 1
Else
   ªì©l­·®æ½X = 0
End If

If mnu_½Æ»s¦r§Î¨ìWord.Checked = True Then
    ªì©lCopyToWord = 1
ElseIf mnu_½Æ»s¹Ï¤ù¨ìWord.Checked = True Then
    ªì©lCopyToWord = 2
Else
    ªì©lCopyToWord = 3
End If

If mnu_½Æ»sUnicode¦r§Î¨ìWord.Checked = True Then
   ªì©lCopyUnicode = 1
Else
   ªì©lCopyUnicode = 0
End If

fsuccess = WritePrivateProfileString("Start", "first", ªì©lfirst, App.path & "\cdphanzi.ini")

IniEntry = ªì©lleft
fsuccess = WritePrivateProfileString("Window", "left", IniEntry, App.path & "\cdphanzi.ini")
IniEntry = ªì©ltop
fsuccess = WritePrivateProfileString("Window", "top", IniEntry, App.path & "\cdphanzi.ini")
IniEntry = ªì©lwidth
fsuccess = WritePrivateProfileString("Window", "width", IniEntry, App.path & "\cdphanzi.ini")
IniEntry = ªì©lheight
fsuccess = WritePrivateProfileString("Window", "height", IniEntry, App.path & "\cdphanzi.ini")

IniEntry = ªì©lfont
fsuccess = WritePrivateProfileString("Font", "fontname", IniEntry, App.path & "\cdphanzi.ini")
IniEntry = ªì©lfontsize
fsuccess = WritePrivateProfileString("Font", "fontsize", IniEntry, App.path & "\cdphanzi.ini")

IniEntry = ¹Ï¤ù¸ÑªR«×
fsuccess = WritePrivateProfileString("Image", "dpi", IniEntry, App.path & "\cdphanzi.ini")
IniEntry = ¹Ï¤ù¦r«¬¤j¤p
fsuccess = WritePrivateProfileString("Image", "fontsize", IniEntry, App.path & "\cdphanzi.ini")

'If ªì©lsave = 1 Then

fsuccess = WritePrivateProfileString("¦r¶°", "±`¥Î¦r", ¦r¶°_±`¥Î¦r, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r¶°", "¤­¤j½X", ¦r¶°_¤­¤j½X, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r¶°", "Â²¤Æ¦r", ¦r¶°_Â²¤Æ¦r, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r¶°", "º~»y¤j¦r¨å", ¦r¶°_º~»y¤j¦r¨å, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r¶°", "»¡¤å¸Ñ¦r", ¦r¶°_»¡¤å¸Ñ¦r, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r¶°", "ª÷¤å½s", ¦r¶°_ª÷¤å½s, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r¶°", "ª÷¤å½s¹Ï§Î¤å¦r", ¦r¶°_ª÷¤å½s¹Ï§Î¤å¦r, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r¶°", "¥Ò°©ÃþÄ¡", ¦r¶°_¥Ò°©ÃþÄ¡, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r¶°", "·¡¨tÂ²©­¤å¦r½s", ¦r¶°_·¡¨tÂ²©­¤å¦r½s, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r¶°", "·¢Åé¦r", ¦r¶°_·¢Åé¦r, App.path & "\cdphanzi.ini")


fsuccess = WritePrivateProfileString("¦r§Î´F¨Å", "open", ´F¨Åopen, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Î´F¨Å", "winstate", ´F¨Åwinstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Î´F¨Å", "left", ´F¨Åleft, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Î´F¨Å", "top", ´F¨Åtop, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Î´F¨Å", "width", ´F¨Åwidth, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Î´F¨Å", "height", ´F¨Åheight, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("¥X³BÀË¦r", "open", ¥X³Bopen, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¥X³BÀË¦r", "winstate", ¥X³Bwinstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¥X³BÀË¦r", "left", ¥X³Bleft, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¥X³BÀË¦r", "top", ¥X³Btop, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¥X³BÀË¦r", "width", ¥X³Bwidth, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¥X³BÀË¦r", "height", ¥X³Bheight, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("¦r§Îµ²ºc", "open", µ²ºcopen, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Îµ²ºc", "winstate", µ²ºcwinstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Îµ²ºc", "left", µ²ºcleft, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Îµ²ºc", "top", µ²ºctop, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Îµ²ºc", "width", µ²ºcwidth, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Îµ²ºc", "height", µ²ºcheight, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("²§Åé¦rªí", "open", ²§Åéopen, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("²§Åé¦rªí", "winstate", ²§Åéwinstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("²§Åé¦rªí", "left", ²§Åéleft, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("²§Åé¦rªí", "top", ²§Åétop, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("²§Åé¦rªí", "width", ²§Åéwidth, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("²§Åé¦rªí", "height", ²§Åéheight, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("²§Åé¦r®Ú", "open", ²§®Úopen, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("²§Åé¦r®Ú", "winstate", ²§®Úwinstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("²§Åé¦r®Ú", "left", ²§®Úleft, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("²§Åé¦r®Ú", "top", ²§®Útop, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("²§Åé¦r®Ú", "width", ²§®Úwidth, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("²§Åé¦r®Ú", "height", ²§®Úheight, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("ºc§Î³¡¥ó", "open", ³¡¥óopen, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("ºc§Î³¡¥ó", "winstate", ³¡¥ówinstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("ºc§Î³¡¥ó", "left", ³¡¥óleft, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("ºc§Î³¡¥ó", "top", ³¡¥ótop, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("ºc§Î³¡¥ó", "width", ³¡¥ówidth, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("ºc§Î³¡¥ó", "height", ³¡¥óheight, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("¦r§ÎºtÅÜ", "open", ºtÅÜopen, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§ÎºtÅÜ", "winstate", ºtÅÜwinstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§ÎºtÅÜ", "left", ºtÅÜleft, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§ÎºtÅÜ", "top", ºtÅÜtop, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§ÎºtÅÜ", "width", ºtÅÜwidth, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§ÎºtÅÜ", "height", ºtÅÜheight, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("¦r§Î¯Á¤Þ", "open", ¯Á¤Þopen, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Î¯Á¤Þ", "winstate", ¯Á¤Þwinstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Î¯Á¤Þ", "left", ¯Á¤Þleft, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Î¯Á¤Þ", "top", ¯Á¤Þtop, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Î¯Á¤Þ", "width", ¯Á¤Þwidth, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¦r§Î¯Á¤Þ", "height", ¯Á¤Þheight, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("¿ï¶µ", "³v¯Å¦C¥X", ªì©l³v¯Å¦C¥X, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "³¡¥ó¶¶§Ç", ªì©l³¡¥ó¶¶§Ç, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "²§¼g³¡¥ó", ªì©l²§¼g³¡¥ó, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "¸Ñ§Î¦C¥X", ªì©l¸Ñ§Î¦C¥X, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "»·ªFº~»y¤j¦r¨å", ªì©l»·ªFº~»y¤j¦r¨å, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "«Ø§»º~»y¤j¦r¨å", ªì©l«Ø§»º~»y¤j¦r¨å, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "¤¤¤å¤jÃã¨å", ªì©l¤¤¤å¤jÃã¨å, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "»¡¤å¸Ñ¦rµþªL", ªì©l»¡¤å¸Ñ¦rµþªL, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "¤¤µØ»¡¤å¸Ñ¦r", ªì©l¤¤µØ»¡¤å¸Ñ¦r, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "ª÷¤å½s", ªì©lª÷¤å½s, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "ª÷¤åµþªL", ªì©lª÷¤åµþªL, App.path & "\cdphanzi.ini")
'fsuccess = WritePrivateProfileString("¿ï¶µ", "ª÷¤å¾¹¸¹", ªì©lª÷¤å¾¹¸¹, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "ª÷¤å¤Þ±o", ªì©lª÷¤å¤Þ±o, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "¥Ò°©¨èÃãÃþÄ¡", ªì©l¥Ò°©¨èÃãÃþÄ¡, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "¥Ò°©¤å¦rµþªL", ªì©l¥Ò°©¤å¦rµþªL, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "¥Ò°©¤å¦r¶°ÄÀ", ªì©l¥Ò°©¤å¦r¶°ÄÀ, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "·¡¨tÂ²©­¤å¦r½s", ªì©l·¡¨tÂ²©­¤å¦r½s, App.path & "\cdphanzi.ini")
'fsuccess = WritePrivateProfileString("¿ï¶µ", "·¡¨t¤å¦r¥X³B", ªì©l·¡¨t¤å¦r¥X³B, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "Unicode", ªì©lUnicode, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "Big5", ªì©lBig5, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "¥Ò°©¤åºtÅÜ", ªì©l¥Ò°©¤åºtÅÜ, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "ª÷¤åºtÅÜ", ªì©lª÷¤åºtÅÜ, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "·¡¨t¤å¦rºtÅÜ", ªì©l·¡¨t¤å¦rºtÅÜ, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "¤p½fºtÅÜ", ªì©l¤p½fºtÅÜ, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "copy", ªì©lcopy, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "¦rÀW", ªì©l¦rÀW, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("¿ï¶µ", "save", ªì©lsave, App.path & "\cdphanzi.ini")
'fsuccess = WritePrivateProfileString("¿ï¶µ", "­·®æ½X", ªì©l­·®æ½X, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("CopyToWord", "CopyMode", ªì©lCopyToWord, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("CopyToWord", "CopyUnicode", ªì©lCopyUnicode, App.path & "\cdphanzi.ini")

'End If

End Sub

Private Sub ªì©lµe­±³]©w()
Me.show
If ªì©lfirst = 1 Then
   mdiº~¦r¦r§Î.WindowState = 2
Else
   mdiº~¦r¦r§Î.Left = ªì©lleft
   mdiº~¦r¦r§Î.Top = ªì©ltop
   mdiº~¦r¦r§Î.Height = ªì©lheight
   mdiº~¦r¦r§Î.Width = ªì©lwidth
End If

'mdiº~¦r¦r§Î.tbr¦r§ÎÄÝ©Ê.ButtonHeight = 360
'mdiº~¦r¦r§Î.tbr¿ï¨ú¦r§Î.ButtonHeight = 360

cbo¦r«¬¦WºÙ.Text = ªì©lfont
If ªì©lfirst = 1 Then
   cbo¦rÅé¤j¤p.Text = 24
Else
   cbo¦rÅé¤j¤p.Text = ªì©lfontsize
End If

Äæ¼e = CInt(cbo¦rÅé¤j¤p.Text) * 20 + CInt(cbo¦rÅé¤j¤p.Text) * 20 / 3

End Sub

Private Sub ¸ü¤Jµøµ¡()

If ªì©lfirst = 1 Then
   mnu_Àx¦sµøµ¡³]©w.Checked = True
   mnu_½Æ»s¨ì°Å¶KÃ¯.Checked = True
   mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó.Checked = True
   mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Checked = False
   mnu_¦r§Î´F¨Å¥u¦C¦r§Î¤£¦C³¡¥ó.Checked = True

   'mnu_¦r§Î´F¨Å¥u¦C¥X±`¥Î¦r.Checked = False
   'mnu_¦r§Î´F¨Å¥u¦C¥X¹q¸£¥Î¦r.Checked = False
   'mnu_¦r§Î´F¨Å¦C¥X©Ò¦³¦r§Î.Checked = True
   ªì©l¦rÀW = 1
   mnu_·¢®Ñ_Click
   ¹w³]ÂsÄý¤@
Else
   ¤w¸ü¤Jµe­± = 0
   
   ªì©l¦rÀW = 1
   
   If CInt(¦r¶°_±`¥Î¦r) = 1 Then mnu_±`¥Î¦r_Click
   If CInt(¦r¶°_¤­¤j½X) = 1 Then mnu_Big5_Click
   If CInt(¦r¶°_Â²¤Æ¦r) = 1 Then mnu_Â²¤Æ¦rÁ`ªí_Click
   If CInt(¦r¶°_º~»y¤j¦r¨å) = 1 Then mnu_º~»y¤j¦r¨å_Click
   If CInt(¦r¶°_»¡¤å¸Ñ¦r) = 1 Then mnu_¤p½f_Click
   If CInt(¦r¶°_ª÷¤å½s) = 1 Then mnu_ª÷¤å_Click
   If CInt(¦r¶°_ª÷¤å½s¹Ï§Î¤å¦r) = 1 Then mnu_ª÷¤å¹Ï§Î¤å¦r_Click
   If CInt(¦r¶°_¥Ò°©ÃþÄ¡) = 1 Then mnu_¥Ò°©¤å_Click
   If CInt(¦r¶°_·¡¨tÂ²©­¤å¦r½s) = 1 Then mnu_·¡¨t¤å¦r_Click
   If CInt(¦r¶°_·¢Åé¦r) = 1 Then mnu_·¢®Ñ_Click
    
   If ³¡¥óopen = 1 Then
        If ¨t²Î¦rÅé = "¤p½f" Then
            mnu_»¡¤å¸Ñ¦r³¡­º_Click
        Else
            mnu_±dº³¦r¨å³¡­º_Click
        End If
   End If
   
   If µ²ºcopen = 1 Then
      frm¦r§Îµ²ºc.Tag = ¦r§Îµ²ºc¥N½X
      frm¦r§Îµ²ºc.show
   End If
   
   If ²§Åéopen = 1 Then
      frm²§Åé¦rªí.Tag = ²§Åé¦rªí¥N½X
      frm²§Åé¦rªí.show
   End If
   
   If ²§®Úopen = 1 Then
      frm²§Åé¦r®Ú.Tag = ²§Åé¦r®Ú¥N½X
      frm²§Åé¦r®Ú.show
   End If

   If ºtÅÜopen = 1 Then
      frm¦r§ÎºtÅÜ.Tag = ¦r§ÎºtÅÜ¥N½X
      frm¦r§ÎºtÅÜ.show
   End If

   If ¯Á¤Þopen = 1 Then
      frm¦r§Î¯Á¤Þ.Tag = ¦r§Î¯Á¤Þ¥N½X
      frm¦r§Î¯Á¤Þ.show
   End If

   If ¥X³Bopen = 1 Then
      frm¥X³BÀË¦r.Tag = ¥X³BÀË¦r¥N½X
      frm¥X³BÀË¦r.show
   End If
   
   If ´F¨Åopen = 1 Then
      frm¦r§Î´F¨Å.Tag = ¦r§Î´F¨Å¥N½X
      frm¦r§Î´F¨Å.show
   End If
      
   If ªì©l³v¯Å¦C¥X = 1 Then
      mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Checked = True
      mnu_¦r§Î´F¨Å¥u¦C¦r§Î¤£¦C³¡¥ó.Checked = False
   Else
      mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Checked = False
      mnu_¦r§Î´F¨Å¥u¦C¦r§Î¤£¦C³¡¥ó.Checked = True
   End If
   
   If ªì©l³¡¥ó¶¶§Ç = 1 Then
      mnu_¦r§Î´F¨Å¿í·Ó¿é¤J³¡¥ó¶¶§Ç.Checked = True
   Else
      mnu_¦r§Î´F¨Å¿í·Ó¿é¤J³¡¥ó¶¶§Ç.Checked = False
   End If

   If ªì©l²§¼g³¡¥ó = 1 Then
      mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó.Checked = True
   Else
      mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó.Checked = False
   End If

   If ªì©l¸Ñ§Î¦C¥X = 1 Then
      mdiº~¦r¦r§Î.mnu_»¡¤å¸Ñ§Î¶¶§Ç.Checked = True
   Else
      mdiº~¦r¦r§Î.mnu_»¡¤å¸Ñ§Î¶¶§Ç.Checked = False
   End If
   
   If ªì©l»·ªFº~»y¤j¦r¨å = 1 Then
      mnu_»·ªFº~»y¤j¦r¨å¿ï¶µ.Checked = True
   Else
      mnu_»·ªFº~»y¤j¦r¨å¿ï¶µ.Checked = False
   End If

   If ªì©l«Ø§»º~»y¤j¦r¨å = 1 Then
      mnu_«Ø§»º~»y¤j¦r¨å¿ï¶µ.Checked = True
   Else
      mnu_«Ø§»º~»y¤j¦r¨å¿ï¶µ.Checked = False
   End If

   If ªì©l¤¤¤å¤jÃã¨å = 1 Then
      mnu_¤¤¤å¤jÃã¨å¿ï¶µ.Checked = True
   Else
      mnu_¤¤¤å¤jÃã¨å¿ï¶µ.Checked = False
   End If

   If ªì©l»¡¤å¸Ñ¦rµþªL = 1 Then
      mnu_»¡¤å¸Ñ¦rµþªL¿ï¶µ.Checked = True
   Else
      mnu_»¡¤å¸Ñ¦rµþªL¿ï¶µ.Checked = False
   End If

   If ªì©l¤¤µØ»¡¤å¸Ñ¦r = 1 Then
      mnu_¤¤µØ»¡¤å¸Ñ¦r¿ï¶µ.Checked = True
   Else
      mnu_¤¤µØ»¡¤å¸Ñ¦r¿ï¶µ.Checked = False
   End If

   If ªì©lª÷¤å½s = 1 Then
      mnu_ª÷¤å½s¿ï¶µ.Checked = True
   Else
      mnu_ª÷¤å½s¿ï¶µ.Checked = False
   End If

   If ªì©lª÷¤åµþªL = 1 Then
      mnu_ª÷¤åµþªL¿ï¶µ.Checked = True
   Else
      mnu_ª÷¤åµþªL¿ï¶µ.Checked = False
   End If

   'If ªì©lª÷¤å¾¹¸¹ = 1 Then
   '   mnu_®ï©Pª÷¤å¶°¦¨¾¹¸¹¿ï¶µ.Checked = True
   'Else
   '   mnu_®ï©Pª÷¤å¶°¦¨¾¹¸¹¿ï¶µ.Checked = False
   'End If

   If ªì©lª÷¤å¤Þ±o = 1 Then
      mnu_®ï©Pª÷¤å¶°¦¨¤Þ±o¿ï¶µ.Checked = True
   Else
      mnu_®ï©Pª÷¤å¶°¦¨¤Þ±o¿ï¶µ.Checked = False
   End If

   If ªì©l¥Ò°©¨èÃãÃþÄ¡ = 1 Then
      mnu_¥Ò°©¨èÃãÃþÄ¡¿ï¶µ.Checked = True
   Else
      mnu_¥Ò°©¨èÃãÃþÄ¡¿ï¶µ.Checked = False
   End If

   If ªì©l¥Ò°©¤å¦rµþªL = 1 Then
      mnu_¥Ò°©¤å¦rµþªL¿ï¶µ.Checked = True
   Else
      mnu_¥Ò°©¤å¦rµþªL¿ï¶µ.Checked = False
   End If

   If ªì©l¥Ò°©¤å¦r¶°ÄÀ = 1 Then
      mnu_¥Ò°©¤å¦r¶°ÄÀ¿ï¶µ.Checked = True
   Else
      mnu_¥Ò°©¤å¦r¶°ÄÀ¿ï¶µ.Checked = False
   End If

   If ªì©l·¡¨tÂ²©­¤å¦r½s = 1 Then
      mnu_·¡¨tÂ²©­¤å¦r½s¿ï¶µ.Checked = True
   Else
      mnu_·¡¨tÂ²©­¤å¦r½s¿ï¶µ.Checked = False
   End If

   'If ªì©l·¡¨t¤å¦r¥X³B = 1 Then
   '   mnu_·¡¨t¤å¦r¥X³B¿ï¶µ.Checked = True
   'Else
   '   mnu_·¡¨t¤å¦r¥X³B¿ï¶µ.Checked = False
   'End If

   If ªì©lUnicode = 1 Then
      mnu_Unicode¿ï¶µ.Checked = True
   Else
      mnu_Unicode¿ï¶µ.Checked = False
   End If

   If ªì©lBig5 = 1 Then
      mnu_Big5¿ï¶µ.Checked = True
   Else
      mnu_Big5¿ï¶µ.Checked = False
   End If

   If ªì©l¥Ò°©¤åºtÅÜ = 1 Then
      mnu_¥Ò°©¤å¿ï¶µ.Checked = True
   Else
      mnu_¥Ò°©¤å¿ï¶µ.Checked = False
   End If

   If ªì©lª÷¤åºtÅÜ = 1 Then
      mnu_ª÷¤å¿ï¶µ.Checked = True
   Else
      mnu_ª÷¤å¿ï¶µ.Checked = False
   End If

   If ªì©l·¡¨t¤å¦rºtÅÜ = 1 Then
      mnu_·¡¨t¤å¦r¿ï¶µ.Checked = True
   Else
      mnu_·¡¨t¤å¦r¿ï¶µ.Checked = False
   End If

   If ªì©l¤p½fºtÅÜ = 1 Then
      mnu_¤p½f¿ï¶µ.Checked = True
   Else
      mnu_¤p½f¿ï¶µ.Checked = False
   End If
   
   If ªì©lsave = 1 Then
      mnu_Àx¦sµøµ¡³]©w.Checked = True
   Else
      mnu_Àx¦sµøµ¡³]©w.Checked = False
   End If

   If ªì©lcopy = 1 Then
      mnu_½Æ»s¨ì°Å¶KÃ¯.Checked = True
   Else
      mnu_½Æ»s¨ì°Å¶KÃ¯.Checked = False
   End If
   
   'If ªì©l­·®æ½X = 1 Then
   '   mnu_Åã¥Ü­·®æ½X.Checked = True
   'Else
   '   mnu_Åã¥Ü­·®æ½X.Checked = False
   'End If
    If ªì©lCopyToWord = 1 Then
        mnu_½Æ»s¦r§Î¨ìWord.Checked = True
        mnu_½Æ»s¹Ï¤ù¨ìWord.Checked = False
        mnu_´¼¼z«¬½Æ»s¨ìWord.Checked = False
    ElseIf ªì©lCopyToWord = 2 Then
        mnu_½Æ»s¦r§Î¨ìWord.Checked = False
        mnu_½Æ»s¹Ï¤ù¨ìWord.Checked = True
        mnu_´¼¼z«¬½Æ»s¨ìWord.Checked = False
    Else
        mnu_½Æ»s¦r§Î¨ìWord.Checked = False
        mnu_½Æ»s¹Ï¤ù¨ìWord.Checked = False
        mnu_´¼¼z«¬½Æ»s¨ìWord.Checked = True
    End If
   
   ¤w¸ü¤Jµe­± = 1
   
End If

End Sub


Public Sub cbo¦r«¬¦WºÙ_click()

Dim i As Integer
Dim j As Long

If Len(cbo¦r«¬¦WºÙ.Text) = 0 Then cbo¦r«¬¦WºÙ.Text = Åã¥Ü¦r«¬
If Åã¥Ü¦r«¬ = cbo¦r«¬¦WºÙ.Text Then Exit Sub

Åã¥Ü¦r«¬ = cbo¦r«¬¦WºÙ.Text

For i = 1 To Forms.Count - 1

    If (CInt(Forms(i).Tag) >= Big5¦r®Ú¥N½X) And (CInt(Forms(i).Tag) <= ºc¦r²Å¸¹¥N½X) Then
       frm³¡¥ó½d¨Ò.tree¦r§Î¾ðª¬µ²ºc.FontName = Åã¥Ü¦r«¬
       'For j = 0 To frm³¡¥ó½d¨Ò.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
       '    frm³¡¥ó½d¨Ò.tree¦r§Î¾ðª¬µ²ºc.ItemFontSize(j) = Åã¥Ü¦r«¬¤j¤p
       'Next j
    End If
    
    If CInt(Forms(i).Tag) = ¦r§Î´F¨Å¥N½X Then
        frm¦r§Î´F¨Å.tree¦r§Î¾ðª¬µ²ºc.FontName = Åã¥Ü¦r«¬
            For j = 0 To frm¦r§Î´F¨Å.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
                If frm¦r§Î´F¨Å.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) <> Åã¥Ü¦r«¬ Then
                    frm¦r§Î´F¨Å.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) = ¤Á´«Åã¥Ü¦r«¬(frm¦r§Î´F¨Å.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j))
                End If
            Next j
    End If
    
    If CInt(Forms(i).Tag) = ¥X³BÀË¦r¥N½X Then
        frm¥X³BÀË¦r.tree¦r§Î¾ðª¬µ²ºc.FontName = Åã¥Ü¦r«¬
            For j = 0 To frm¥X³BÀË¦r.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
                If frm¥X³BÀË¦r.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) <> Åã¥Ü¦r«¬ Then
                    frm¥X³BÀË¦r.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) = ¤Á´«Åã¥Ü¦r«¬(frm¥X³BÀË¦r.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j))
                End If
            Next j
    End If
    
    If CInt(Forms(i).Tag) = ¦r§Îµ²ºc¥N½X Then
            frm¦r§Îµ²ºc.tree¦r§Î¾ðª¬µ²ºc.FontName = Åã¥Ü¦r«¬
            'frm¦r§Îµ²ºc.tree¦r§Î¾ðª¬µ²ºc.ItemFontSize(0) = Åã¥Ü¦r«¬¤j¤p
            For j = 0 To frm¦r§Îµ²ºc.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
                If frm¦r§Îµ²ºc.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) <> Åã¥Ü¦r«¬ Then
                    frm¦r§Îµ²ºc.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) = ¤Á´«Åã¥Ü¦r«¬(frm¦r§Îµ²ºc.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j))
                End If
            Next j
    End If
    
    If CInt(Forms(i).Tag) = ²§Åé¦rªí¥N½X Then
            frm²§Åé¦rªí.tree¦r§Î¾ðª¬µ²ºc.FontName = Åã¥Ü¦r«¬
            For j = 0 To frm²§Åé¦rªí.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
                If frm²§Åé¦rªí.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) <> Åã¥Ü¦r«¬ Then
                    frm²§Åé¦rªí.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) = ¤Á´«Åã¥Ü¦r«¬(frm²§Åé¦rªí.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j))
                End If
            Next j
    End If

    If CInt(Forms(i).Tag) = ²§Åé¦r®Ú¥N½X Then
            frm²§Åé¦r®Ú.tree¦r§Î¾ðª¬µ²ºc.FontName = Åã¥Ü¦r«¬
            For j = 0 To frm²§Åé¦r®Ú.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
                If frm²§Åé¦r®Ú.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) <> Åã¥Ü¦r«¬ Then
                    frm²§Åé¦r®Ú.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) = ¤Á´«Åã¥Ü¦r«¬(frm²§Åé¦r®Ú.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j))
                End If
            Next j
    End If

    If CInt(Forms(i).Tag) = ¦r§ÎºtÅÜ¥N½X Then
       For j = 0 To frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
           If frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) <> "¥_®v¤j»¡¤å¤p½f" And frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) <> "¥_®v¤j»¡¤å­«¤å" And frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) <> "¤¤¬ã°|ª÷¤å" And frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) <> "¤¤¬ã°|¥Ò°©¤å" And frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) <> "¤¤¬ã°|·¡¨tÂ²©­¤å¦r" Then
              If frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ItemCell(j).RTFStyle = 1 Then
                 frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.List(j) = Âà´«RTF¯Ê¦r(Right(frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ItemTag(j), Len(frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ItemTag(j)) - 1), Åã¥Ü¦r«¬)
              Else
                frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) = ¤Á´«Åã¥Ü¦r«¬(frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j))
              End If
           End If
       Next j
    End If

    If CInt(Forms(i).Tag) = ¦r§Î¯Á¤Þ¥N½X Then
        frm¦r§Î¯Á¤Þ.tree¦r§Î¾ðª¬µ²ºc.FontName = Åã¥Ü¦r«¬
        For j = 0 To frm¦r§Î¯Á¤Þ.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
            If frm¦r§Î¯Á¤Þ.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) <> Åã¥Ü¦r«¬ Then
                frm¦r§Î¯Á¤Þ.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j) = ¤Á´«Åã¥Ü¦r«¬(frm¦r§Î¯Á¤Þ.tree¦r§Î¾ðª¬µ²ºc.ItemFontName(j))
            End If
        Next j
    End If

Next i

End Sub

Private Sub cbo¦r«¬¦WºÙ_KeyPress(KeyAscii As Integer)

Dim ¦r«¬¦WºÙ As String

If KeyAscii = 13 Then
    If Len(cbo¦r«¬¦WºÙ.Text) = 0 Then cbo¦r«¬¦WºÙ.Text = Åã¥Ü¦r«¬
    ¦r«¬¦WºÙ = cbo¦r«¬¦WºÙ.Text
    If ¦r«¬¦WºÙ = "²Ó©úÅé" Or ¦r«¬¦WºÙ = "¼Ð·¢Åé" Then
        cbo¦r«¬¦WºÙ.Text = ¦r«¬¦WºÙ
        cbo¦r«¬¦WºÙ_click
    Else
        cbo¦r«¬¦WºÙ.Text = Åã¥Ü¦r«¬
    End If
End If

End Sub

Private Sub cbo¦r«¬¦WºÙ_LostFocus()

If Len(cbo¦r«¬¦WºÙ.Text) = 0 Then cbo¦r«¬¦WºÙ.Text = Åã¥Ü¦r«¬

End Sub

Private Sub cbo¦rÅé¤j¤p_LostFocus()

If Len(cbo¦rÅé¤j¤p.Text) = 0 Then cbo¦rÅé¤j¤p.Text = Åã¥Ü¦r«¬¤j¤p

End Sub

Private Sub cbo¸ÑªR«×_Click()

Dim ¦rÅé¤j¤p As Integer

¦rÅé¤j¤p = Val(cbo¸ÑªR«×.Text)
If ¦rÅé¤j¤p > 0 Then ¹Ï¤ù¸ÑªR«× = ¦rÅé¤j¤p

End Sub

Private Sub cbo¸ÑªR«×_KeyPress(KeyAscii As Integer)

If Val(cbo¸ÑªR«×.Text) > 0 Then ¹Ï¤ù¸ÑªR«× = Val(cbo¸ÑªR«×.Text)

End Sub

Private Sub cbo¹Ï¤ù¤j¤p_Click()

Dim ¦rÅé¤j¤p As Integer

¦rÅé¤j¤p = Val(cbo¹Ï¤ù¤j¤p.Text)
If ¦rÅé¤j¤p > 0 Then ¹Ï¤ù¦r«¬¤j¤p = ¦rÅé¤j¤p

End Sub

Private Sub cbo¹Ï¤ù¤j¤p_KeyPress(KeyAscii As Integer)

Dim ¦rÅé¤j¤p As Integer

If KeyAscii = 13 Then
   If Len(cbo¹Ï¤ù¤j¤p.Text) = 0 Then cbo¹Ï¤ù¤j¤p.Text = ¹Ï¤ù¦r«¬¤j¤p
   ¦rÅé¤j¤p = Val(cbo¹Ï¤ù¤j¤p.Text)
   If ¦rÅé¤j¤p >= 8 And ¦rÅé¤j¤p <= 1000 Then
      cbo¹Ï¤ù¤j¤p_Click
   ElseIf ¦rÅé¤j¤p < 8 Then
      cbo¹Ï¤ù¤j¤p.Text = 8
      cbo¹Ï¤ù¤j¤p_Click
   ElseIf ¦rÅé¤j¤p > 1000 Then
      cbo¹Ï¤ù¤j¤p.Text = 1000
      cbo¹Ï¤ù¤j¤p_Click
   End If
End If

End Sub

Private Sub MDIForm_Load()

Dim i As Integer
Dim ²Õ¦r²Å¸¹ªí As Recordset
Dim ret As Long, lendata As Long, WinPath As String, path As String

¸ü¤Jªì©l­È
ªì©lµe­±³]©w

If ªì©lfirst = 1 Then
   ªì©lleft = mdiº~¦r¦r§Î.Left
   ªì©ltop = mdiº~¦r¦r§Î.Top
   ªì©lheight = mdiº~¦r¦r§Î.Height
   ªì©lwidth = mdiº~¦r¦r§Î.Width
End If

'Me.show

¨t²Î¦rÅé = "·¢®Ñ"
¶}±Ò¸ê®Æ®w

Set ÀË¦rªí = ¨t²Î¸ê®Æ®w.OpenRecordset("ÀË¦rªí")
Set ±`¥Î²Å¸¹¤Î³¡¥óÃþ«¬ = ¨t²Î¸ê®Æ®w.OpenRecordset("±`¥Î²Å¸¹¤Î³¡¥óÃþ«¬")
Set ²Õ¦r²Å¸¹ªí = ¨t²Î¸ê®Æ®w.OpenRecordset("²Å¸¹")
Set ±dº³³¡­º = ¨t²Î¸ê®Æ®w.OpenRecordset("±dº³³¡­º")
Set »¡¤å³¡­º = ¨t²Î¸ê®Æ®w.OpenRecordset("»¡¤å³¡­º")
Set ²§¼g¦r®Ú = ¨t²Î¸ê®Æ®w.OpenRecordset("²§¼g¦r®Ú")
Set ²§Åé¦rªí = ¨t²Î¸ê®Æ®w.OpenRecordset("²§Åé¦rªí")

Set ·¢®ÑÀË¦rªí = ¨t²Î¸ê®Æ®w.OpenRecordset("ÀË¦rªí")
Set ·¢®Ñ¦r®Ú = ¨t²Î¸ê®Æ®w.OpenRecordset("¦r®Ú")
Set ·¢®Ñ²§¼g¦r®Ú = ¨t²Î¸ê®Æ®w.OpenRecordset("²§¼g¦r®Ú")
Set ·¢®Ñ²§Åé¦rªí = ¨t²Î¸ê®Æ®w.OpenRecordset("²§Åé¦rªí")

Set ¤p½fÀË¦rªí = ¤p½f¸ê®Æ®w.OpenRecordset("ÀË¦rªí")
Set ¤p½f¿WÅé¦r = ¤p½f¸ê®Æ®w.OpenRecordset("¦r®Ú")
Set ¤p½f²§¼g¦r®Ú = ¤p½f¸ê®Æ®w.OpenRecordset("²§¼g¦r®Ú")
Set ¤p½f²§Åé¦rªí = ¤p½f¸ê®Æ®w.OpenRecordset("²§Åé¦rªí")

Set ª÷¤åÀË¦rªí = ª÷¤å¸ê®Æ®w.OpenRecordset("ÀË¦rªí")
Set ª÷¤å¸É¿òªí = ª÷¤å¸ê®Æ®w.OpenRecordset("¸É¿ò")
Set ª÷¤å¦r®Ú = ª÷¤å¸ê®Æ®w.OpenRecordset("¦r®Ú")
Set ª÷¤å²§¼g¦r®Ú = ª÷¤å¸ê®Æ®w.OpenRecordset("²§¼g¦r®Ú")
Set ª÷¤å²§Åé¦rªí = ª÷¤å¸ê®Æ®w.OpenRecordset("²§Åé¦rªí")
Set ª÷¤å²§¼g¦rªí = ª÷¤å¸ê®Æ®w.OpenRecordset("²§¼g¦rªí")
Set ª÷¤å¶°¦¨¾¹¦W = ª÷¤å¸ê®Æ®w.OpenRecordset("¶°¦¨¾¹¦W")
Set ª÷¤å¶°¦¨¤Þ±o = ª÷¤å¸ê®Æ®w.OpenRecordset("¶°¦¨¤Þ±o")
Set ª÷¤åµþªL = ª÷¤å¸ê®Æ®w.OpenRecordset("ª÷¤åµþªL")

Set ¥Ò°©¤åÀË¦rªí = ¥Ò°©¤å¸ê®Æ®w.OpenRecordset("ÀË¦rªí")
Set ¥Ò°©¤å¦r®Ú = ¥Ò°©¤å¸ê®Æ®w.OpenRecordset("¦r®Ú")
Set ¥Ò°©¤å²§¼g¦r®Ú = ¥Ò°©¤å¸ê®Æ®w.OpenRecordset("²§¼g¦r®Ú")
Set ¥Ò°©¤å²§Åé¦rªí = ¥Ò°©¤å¸ê®Æ®w.OpenRecordset("²§Åé¦rªí")
Set ¥Ò°©¤å²§¼g¦rªí = ¥Ò°©¤å¸ê®Æ®w.OpenRecordset("²§¼g¦rªí")

Set ·¡¨t¤å¦rÀË¦rªí = ·¡¨t¤å¦r¸ê®Æ®w.OpenRecordset("ÀË¦rªí")
Set ·¡¨t¤å¦r¸É¿òªí = ·¡¨t¤å¦r¸ê®Æ®w.OpenRecordset("¸É¿ò")
Set ·¡¨t¤å¦r¦r®Ú = ·¡¨t¤å¦r¸ê®Æ®w.OpenRecordset("¦r®Ú")
Set ·¡¨t¤å¦r²§¼g¦r®Ú = ·¡¨t¤å¦r¸ê®Æ®w.OpenRecordset("²§¼g¦r®Ú")
Set ·¡¨t¤å¦r²§Åé¦rªí = ·¡¨t¤å¦r¸ê®Æ®w.OpenRecordset("²§Åé¦rªí")
Set ·¡¨t¤å¦r²§¼g¦rªí = ·¡¨t¤å¦r¸ê®Æ®w.OpenRecordset("²§¼g¦rªí")

²§Åé¦rªí.Index = "½s¸¹"
²§¼g¦r®Ú.Index = "¦r®Ú"
ÀË¦rªí.Index = "¦r§Î"

·¢®Ñ²§Åé¦rªí.Index = "½s¸¹"
·¢®Ñ²§¼g¦r®Ú.Index = "¦r®Ú"
·¢®ÑÀË¦rªí.Index = "¦r§Î"

¤p½f²§Åé¦rªí.Index = "½s¸¹"
¤p½f²§¼g¦r®Ú.Index = "¦r®Ú"
¤p½fÀË¦rªí.Index = "¦r§Î"

ª÷¤å²§Åé¦rªí.Index = "½s¸¹"
ª÷¤å²§¼g¦rªí.Index = "½s¸¹"
ª÷¤å²§¼g¦r®Ú.Index = "¦r®Ú"
ª÷¤åÀË¦rªí.Index = "¦r§Î"
ª÷¤å¸É¿òªí.Index = "·¢®Ñ½s¸¹"
ª÷¤å¶°¦¨¤Þ±o.Index = "·¢®Ñ½s¸¹"
ª÷¤å¶°¦¨¾¹¦W.Index = "¾¹¸¹"
ª÷¤åµþªL.Index = "¦rÀY"

¥Ò°©¤å²§Åé¦rªí.Index = "½s¸¹"
¥Ò°©¤å²§¼g¦rªí.Index = "½s¸¹"
¥Ò°©¤å²§¼g¦r®Ú.Index = "¦r®Ú"
¥Ò°©¤åÀË¦rªí.Index = "¦r§Î"

·¡¨t¤å¦r²§Åé¦rªí.Index = "½s¸¹"
·¡¨t¤å¦r²§¼g¦rªí.Index = "½s¸¹"
·¡¨t¤å¦r²§¼g¦r®Ú.Index = "¦r®Ú"
·¡¨t¤å¦rÀË¦rªí.Index = "¦r§Î"
·¡¨t¤å¦r¸É¿òªí.Index = "·¢®Ñ½s¸¹"

²{¥Î¦rÅé = "¼Ð·¢Åé"
Åã¥Ü¦r«¬ = cbo¦r«¬¦WºÙ
Åã¥Ü¦r«¬¤j¤p = cbo¦rÅé¤j¤p

²Õ¦r²Å¸¹ªí.MoveFirst
For i = 1 To 14
    ²Õ¦r²Å¸¹°}¦C(²Õ¦r²Å¸¹ªí.Fields("½s¸¹")) = ²Õ¦r²Å¸¹ªí.Fields("¦r§Î")
    ²Õ¦r²Å¸¹ªí.MoveNext
Next i
²Õ¦r²Å¸¹ªí.Close

Do Until ±`¥Î²Å¸¹¤Î³¡¥óÃþ«¬.EOF
   cbo²Å¸¹Ãþ«¬.AddItem ±`¥Î²Å¸¹¤Î³¡¥óÃþ«¬.Fields("»¡©ú")
   cbo²Å¸¹Ãþ«¬.ItemData(cbo²Å¸¹Ãþ«¬.NewIndex) = ±`¥Î²Å¸¹¤Î³¡¥óÃþ«¬.Fields("Ãþ«¬")
   ±`¥Î²Å¸¹¤Î³¡¥óÃþ«¬.MoveNext
Loop

cbo²Å¸¹Ãþ«¬.ListIndex = 0
cboµ§µe.AddItem "1-99"

For i = 1 To 99
    cboµ§µe.AddItem i
    cboµ§µe.ItemData(i) = i
Next i


cbo¸ÑªR«×.List(0) = "72"
cbo¸ÑªR«×.List(1) = "120"
cbo¸ÑªR«×.List(2) = "300"
cbo¸ÑªR«×.List(3) = "450"
cbo¸ÑªR«×.List(4) = "600"
cbo¸ÑªR«×.List(5) = "750"
cbo¸ÑªR«×.List(6) = "900"
cbo¸ÑªR«×.List(7) = "1050"
cbo¸ÑªR«×.List(8) = "1200"
cbo¸ÑªR«×.List(9) = "1800"
cbo¸ÑªR«×.List(10) = "2400"
cbo¸ÑªR«×.Text = ¹Ï¤ù¸ÑªR«×

cbo¹Ï¤ù¤j¤p.List(0) = "8"
cbo¹Ï¤ù¤j¤p.List(1) = "9"
cbo¹Ï¤ù¤j¤p.List(2) = "10"
cbo¹Ï¤ù¤j¤p.List(3) = "11"
cbo¹Ï¤ù¤j¤p.List(4) = "12"
cbo¹Ï¤ù¤j¤p.List(5) = "14"
cbo¹Ï¤ù¤j¤p.List(6) = "16"
cbo¹Ï¤ù¤j¤p.List(7) = "18"
cbo¹Ï¤ù¤j¤p.List(8) = "20"
cbo¹Ï¤ù¤j¤p.List(9) = "22"
cbo¹Ï¤ù¤j¤p.List(10) = "24"
cbo¹Ï¤ù¤j¤p.List(11) = "26"
cbo¹Ï¤ù¤j¤p.List(12) = "28"
cbo¹Ï¤ù¤j¤p.List(13) = "36"
cbo¹Ï¤ù¤j¤p.List(14) = "48"
cbo¹Ï¤ù¤j¤p.List(15) = "72"
cbo¹Ï¤ù¤j¤p.Text = ¹Ï¤ù¦r«¬¤j¤p

½Æ»s¹Ï¤ù¨ìWord = False

³]©w¤u¨ã¦Cªì©lª¬ºA

¦@¥Îµøµ¡(mdiº~¦r¦r§Î¥N½X) = "mdiº~¦r¦r§Î"
¦@¥Îµøµ¡(ºc¦r²Å¸¹¥N½X) = "ºc¦r²Å¸¹"
¦@¥Îµøµ¡(Â²Ã|¥N½X) = "Â²Ã|"
¦@¥Îµøµ¡(¤K¨ö¥N½X) = "¤K¨ö"
¦@¥Îµøµ¡(¹Ï§Î¤å¦r¥N½X) = "¹Ï§Î¤å¦r"
¦@¥Îµøµ¡(±dº³³¡­º¥N½X) = "±dº³¦r¨å³¡­º"
¦@¥Îµøµ¡(»¡¤å³¡­º¥N½X) = "»¡¤å¸Ñ¦r³¡­º"
¦@¥Îµøµ¡(Big5¦r®Ú¥N½X) = "Big5¦r®Ú"
¦@¥Îµøµ¡(Big5¤ÎÂ²¤Æ¦r¦r®Ú¥N½X) = "Big5¤ÎÂ²¤Æ¦r¦r®Ú"
¦@¥Îµøµ¡(¦r®Ú¥N½X) = "¦r®Ú"
¦@¥Îµøµ¡(¤p½f¿WÅé¦r¥N½X) = "¤p½f¿WÅé¦r"
¦@¥Îµøµ¡(ª÷¤å¦r®Ú¥N½X) = "ª÷¤å¦r®Ú"
¦@¥Îµøµ¡(¥Ò°©¤å¦r®Ú¥N½X) = "¥Ò°©¤å¦r®Ú"
¦@¥Îµøµ¡(·¡¨tÂ²©­¤å¦r¦r®Ú¥N½X) = "·¡¨tÂ²©­¤å¦r¦r®Ú"
¦@¥Îµøµ¡(³¡¥ó¥~¦r¥N½X) = "³¡¥ó¥~¦r"
¦@¥Îµøµ¡(¦r§Î´F¨Å¥N½X) = "¦r§Î´F¨Å"
¦@¥Îµøµ¡(¥X³BÀË¦r¥N½X) = "¥X³BÀË¦r"
¦@¥Îµøµ¡(¦r§Îµ²ºc¥N½X) = "¦r§Îµ²ºc"
¦@¥Îµøµ¡(¦r§Î¯Á¤Þ¥N½X) = "¦r§Î¯Á¤Þ"
¦@¥Îµøµ¡(²§Åé¦rªí¥N½X) = "²§Åé¦rªí"
¦@¥Îµøµ¡(²§Åé¦r®Ú¥N½X) = "²§Åé¦r®Ú"

²{¥Îµøµ¡ = "mdiº~¦r¦r§Î"
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
Me.Tag = mdiº~¦r¦r§Î¥N½X

µ§µe­ºµ§¬d¸ß = True
'µ§µe = 1
µ§µe = 0
­ºµ§ = 0
±Ò°Ê¦r§Îµ²ºc = False
±Ò°Ê²§Åé¦rªí = False

For i = 0 To 2
    µøµ¡¥N½X(i) = True
Next i

¸ü¤Jµøµ¡


lendata = 255
path = String(lendata, Chr(0))
ret = GetWindowsDirectory(path, lendata)
WinPath = Left(path, InStr(path, Chr(0)) - 1)
¼È¦s¥Ø¿ý = WinPath & "\Temp\CdphanziTempDir"
If Dir(¼È¦s¥Ø¿ý, vbDirectory) = "" Then MkDir ¼È¦s¥Ø¿ý
bmpcount = 0
WordWasNotRunning = True

¹w³]ÂsÄý¼Ò¦¡ = -1
Â²©öÂsÄý¼Ò¦¡ = True

End Sub

Private Sub cbo¦rÅé¤j¤p_Click()
Dim i As Integer
Dim j As Long

If Len(cbo¦rÅé¤j¤p.Text) = 0 Then cbo¦rÅé¤j¤p.Text = Åã¥Ü¦r«¬¤j¤p
If Åã¥Ü¦r«¬¤j¤p = cbo¦rÅé¤j¤p.Text Then Exit Sub

Åã¥Ü¦r«¬¤j¤p = cbo¦rÅé¤j¤p.Text

Äæ¼e = CInt(Åã¥Ü¦r«¬¤j¤p) * 20 + CInt(Åã¥Ü¦r«¬¤j¤p) * 20 / 3

For i = 1 To Forms.Count - 1

    If (CInt(Forms(i).Tag) >= Big5¦r®Ú¥N½X) And (CInt(Forms(i).Tag) <= ºc¦r²Å¸¹¥N½X) Then
       frm³¡¥ó½d¨Ò.tree¦r§Î¾ðª¬µ²ºc.FontSize = Åã¥Ü¦r«¬¤j¤p
       For j = 0 To frm³¡¥ó½d¨Ò.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
           frm³¡¥ó½d¨Ò.tree¦r§Î¾ðª¬µ²ºc.ItemFontSize(j) = Åã¥Ü¦r«¬¤j¤p
       Next j
    End If
    
    If CInt(Forms(i).Tag) = ¦r§Î´F¨Å¥N½X Then
       frm¦r§Î´F¨Å.tree¦r§Î¾ðª¬µ²ºc.FontSize = Åã¥Ü¦r«¬¤j¤p
       For j = 0 To frm¦r§Î´F¨Å.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
           frm¦r§Î´F¨Å.tree¦r§Î¾ðª¬µ²ºc.ItemFontSize(j) = Åã¥Ü¦r«¬¤j¤p
       Next j
    End If
    
    If CInt(Forms(i).Tag) = ¥X³BÀË¦r¥N½X Then
       frm¥X³BÀË¦r.tree¦r§Î¾ðª¬µ²ºc.FontSize = Åã¥Ü¦r«¬¤j¤p
       For j = 0 To frm¥X³BÀË¦r.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
           frm¥X³BÀË¦r.tree¦r§Î¾ðª¬µ²ºc.ItemFontSize(j) = Åã¥Ü¦r«¬¤j¤p
       Next j
    End If
    
    If CInt(Forms(i).Tag) = ¦r§Îµ²ºc¥N½X Then
       frm¦r§Îµ²ºc.tree¦r§Î¾ðª¬µ²ºc.FontSize = Åã¥Ü¦r«¬¤j¤p
       For j = 0 To frm¦r§Îµ²ºc.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
           frm¦r§Îµ²ºc.tree¦r§Î¾ðª¬µ²ºc.ItemFontSize(j) = Åã¥Ü¦r«¬¤j¤p
       Next j
    End If
    
    If CInt(Forms(i).Tag) = ²§Åé¦rªí¥N½X Then
       frm²§Åé¦rªí.tree¦r§Î¾ðª¬µ²ºc.FontSize = Åã¥Ü¦r«¬¤j¤p
       For j = 0 To frm²§Åé¦rªí.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
           frm²§Åé¦rªí.tree¦r§Î¾ðª¬µ²ºc.ItemFontSize(j) = Åã¥Ü¦r«¬¤j¤p
       Next j
    End If

    If CInt(Forms(i).Tag) = ²§Åé¦r®Ú¥N½X Then
       frm²§Åé¦r®Ú.tree¦r§Î¾ðª¬µ²ºc.FontSize = Åã¥Ü¦r«¬¤j¤p
       For j = 0 To frm²§Åé¦r®Ú.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
           frm²§Åé¦r®Ú.tree¦r§Î¾ðª¬µ²ºc.ItemFontSize(j) = Åã¥Ü¦r«¬¤j¤p
       Next j
    End If

    If CInt(Forms(i).Tag) = ¦r§ÎºtÅÜ¥N½X Then
       frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.FontSize = Åã¥Ü¦r«¬¤j¤p
       For j = 0 To frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
          If Len(frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.List(j)) = 1 Then
            frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ItemFontSize(j) = Åã¥Ü¦r«¬¤j¤p
          End If
       Next j
    End If

    If CInt(Forms(i).Tag) = ¦r§Î¯Á¤Þ¥N½X Then
       frm¦r§Î¯Á¤Þ.tree¦r§Î¾ðª¬µ²ºc.FontSize = Åã¥Ü¦r«¬¤j¤p
       For j = 0 To frm¦r§Î¯Á¤Þ.tree¦r§Î¾ðª¬µ²ºc.ListCount - 1
          If Len(frm¦r§Î¯Á¤Þ.tree¦r§Î¾ðª¬µ²ºc.List(j)) = 1 Then
            frm¦r§Î¯Á¤Þ.tree¦r§Î¾ðª¬µ²ºc.ItemFontSize(j) = Åã¥Ü¦r«¬¤j¤p
          End If
       Next j
    End If

Next i

End Sub


Private Sub cbo¦rÅé¤j¤p_KeyPress(KeyAscii As Integer)

Dim ¦rÅé¤j¤p As Integer

If KeyAscii = 13 Then
   If Len(cbo¦rÅé¤j¤p.Text) = 0 Then cbo¦rÅé¤j¤p.Text = Åã¥Ü¦r«¬¤j¤p
   ¦rÅé¤j¤p = Val(cbo¦rÅé¤j¤p.Text)
   If ¦rÅé¤j¤p >= 10 And ¦rÅé¤j¤p <= 1000 Then
      cbo¦rÅé¤j¤p_Click
   ElseIf ¦rÅé¤j¤p < 10 Then
      cbo¦rÅé¤j¤p.Text = 10
      cbo¦rÅé¤j¤p_Click
   ElseIf ¦rÅé¤j¤p > 1000 Then
      cbo¦rÅé¤j¤p.Text = 1000
      cbo¦rÅé¤j¤p_Click
   End If
End If

End Sub

Private Sub cbo²Å¸¹_Click()
Dim ¦r§Î As String, ³¡¥ó§Ç As String
Dim ½s¸¹ As Long, ³s±µ²Å¸¹ As Integer

¦r§Î = cbo²Å¸¹.List(cbo²Å¸¹.ListIndex)
½s¸¹ = 0

ÀË¦rªí.Index = "¦r§Î"
ÀË¦rªí.Seek "=", cbo²Å¸¹.List(cbo²Å¸¹.ListIndex)
If Not ÀË¦rªí.NoMatch Then
   txt¦r§Î.Text = ¦r§Î
   txt¤º½X.Text = ÀË¦rªí.Fields("BIG5")
   txt­Ü¾e½X.Text = Âà´«­^¤å¨ì­Ü¾e(ÀË¦rªí.Fields("­Ü¾e"))
   ³¡¥ó§Ç = ÀË¦rªí.Fields("³¡¥ó§Ç")
   If Not IsNull(ÀË¦rªí.Fields("½s¸¹")) Then
      ½s¸¹ = ÀË¦rªí.Fields("½s¸¹")
   Else
      ½s¸¹ = 0
   End If
   ³s±µ²Å¸¹ = ÀË¦rªí.Fields("³s±µ²Å¸¹")
   Â^¨úÄÝ©Ê "¼Ð·¢Åé", ¦r§Î, ½s¸¹
   Â^¨úºc¦r¦¡ "¼Ð·¢Åé", ¦r§Î, ½s¸¹
   If ±Ò°Ê¦r§Îµ²ºc And (³s±µ²Å¸¹ <> 9) Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¼Ð·¢Åé", ¦r§Î, ½s¸¹
   If ±Ò°Ê²§Åé¦rªí And (³s±µ²Å¸¹ <> 9) Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¼Ð·¢Åé", ¦r§Î, ½s¸¹
   If ±Ò°Ê¦r§ÎºtÅÜ And (³s±µ²Å¸¹ <> 9) Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¼Ð·¢Åé", ¦r§Î, ½s¸¹
   If ±Ò°Ê²§Åé¦r®Ú And (³s±µ²Å¸¹ <> 9) Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¼Ð·¢Åé", ¦r§Î, ½s¸¹
End If

End Sub


Private Sub cbo­ºµ§_Click()
­ºµ§ = cbo­ºµ§.ItemData(cbo­ºµ§.ListIndex)
If Not µ§µe­ºµ§¬d¸ß Then Exit Sub
frm³¡¥ó½d¨Ò.³¡¥ó¬d¸ß µ§µe, ­ºµ§

End Sub

Private Sub cbo²Å¸¹Ãþ«¬_Click()
Dim ¦r§Îªí As Recordset
Dim SQL³¯­z¦¡ As String

SQL³¯­z¦¡ = "SELECT ½s¸¹,¦r§Î From ±`¥Î²Å¸¹¤Î³¡¥ó Where Ãþ«¬ = " & cbo²Å¸¹Ãþ«¬.ItemData(cbo²Å¸¹Ãþ«¬.ListIndex) & " ORDER BY ½s¸¹"
Set ¦r§Îªí = ¨t²Î¸ê®Æ®w.OpenRecordset(SQL³¯­z¦¡)

cbo²Å¸¹.Clear

Do Until ¦r§Îªí.EOF
   cbo²Å¸¹.AddItem ¦r§Îªí.Fields("¦r§Î")
   ¦r§Îªí.MoveNext
Loop
If cbo²Å¸¹.ListCount > 0 Then cbo²Å¸¹.ListIndex = 0
   
End Sub


Private Sub cboµ§µe_Click()
If cboµ§µe.ListIndex = -1 Then cboµ§µe.ListIndex = 0
µ§µe = cboµ§µe.ItemData(cboµ§µe.ListIndex)
If Not µ§µe­ºµ§¬d¸ß Then Exit Sub
frm³¡¥ó½d¨Ò.³¡¥ó¬d¸ß µ§µe, ­ºµ§

End Sub

Private Sub cboµ§µe_KeyPress(KeyAscii As Integer)

Dim µ§µe¼Æ As Integer

If KeyAscii = 13 Then
   µ§µe¼Æ = Val(cboµ§µe.Text)
   If cboµ§µe.Text = "1-99" Then
      cboµ§µe.ListIndex = -1
   ElseIf µ§µe¼Æ >= 1 And µ§µe¼Æ <= 99 Then
      cboµ§µe.ListIndex = µ§µe¼Æ
   End If
End If

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
­pºâµ²§ôµøµ¡
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
²§¼g¦r®Ú.Close
±`¥Î²Å¸¹¤Î³¡¥óÃþ«¬.Close
ÀË¦rªí.Close
±dº³³¡­º.Close
»¡¤å³¡­º.Close
¨t²Î¸ê®Æ®w.Close
¤p½f¸ê®Æ®w.Close
Àx¦sµ²§ô­È

End Sub

Private Sub ¦rÅé¦WºÙ()
Dim ¼È¦sªí As Recordset
Dim i As Integer
Dim ¼È¦s°}¦C As Variant

Set ¼È¦sªí = ¨t²Î¸ê®Æ®w.OpenRecordset("¦rÅé")
¼È¦sªí.MoveFirst

i = 0
Do Until ¼È¦sªí.EOF
   ¦rÅé°}¦C(¼È¦sªí.Fields("½s¸¹")) = ¼È¦sªí.Fields("¦WºÙ")
   i = i + 1
   ¼È¦sªí.MoveNext
Loop

¼È¦sªí.Close

End Sub

Private Sub mnu_Big5_Click()

ªì©l¦rÀW = 2

If mnu_Big5.Checked = False Then
   mnu_±`¥Î¦r.Checked = False
   mnu_Big5.Checked = True
   mnu_Â²¤Æ¦rÁ`ªí.Checked = False
   mnu_º~»y¤j¦r¨å.Checked = False
   mnu_·¢®Ñ.Checked = False
   mnu_¤p½f.Checked = False
   mnu_ª÷¤å.Checked = False
   mnu_ª÷¤å¹Ï§Î¤å¦r.Checked = False
   mnu_¥Ò°©¤å.Checked = False
   mnu_·¡¨t¤å¦r.Checked = False
   If ±Ò°Ê¦r§Î´F¨Å Then frm¦r§Î´F¨Å.txtºc¦r¦¡.FontName = "¼Ð·¢Åé"
   ¨t²Î¦rÅé = "·¢®Ñ"
End If

End Sub

Private Sub mnu_Big5¿ï¶µ_Click()

If mnu_Big5¿ï¶µ.Checked = True Then
   mnu_Big5¿ï¶µ.Checked = False
Else
   mnu_Big5¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_cdp_Click()
frmª©¥».show 1
End Sub

Private Sub mnu_edit_½Æ»s¨ìWord_Click()

If mnu_½Æ»s¦r§Î¨ìWord.Checked = True Then
    mnu_½Æ»s_Click
    WordApp.Selection.Paste
ElseIf mnu_½Æ»s¹Ï¤ù¨ìWord.Checked Then
    ½Æ»s¹Ï¤ù¨ìWord = True
    ½Æ»s¨ìWordªº¹Ï¤ù¤j¤p = WordApp.Selection.font.Size
    mnu_¶K¹Ï¨ìWord_Click
    ½Æ»s¹Ï¤ù¨ìWord = False
Else
    mnu_½Æ»s_Click
    If ½Æ»sBig5¦r¤¸ Then
        WordApp.Selection.Paste
    Else
        ½Æ»s¹Ï¤ù¨ìWord = True
        ½Æ»s¨ìWordªº¹Ï¤ù¤j¤p = WordApp.Selection.font.Size
        mnu_¶K¹Ï¨ìWord_Click
        ½Æ»s¹Ï¤ù¨ìWord = False
    End If
End If

End Sub

Private Sub mnu_Tool_Click()

If InStr(1, txtºc¦r¦¡, "Æ¡") > 0 Then
    mnu_Tool_ListLikeChar.Enabled = True
Else
    mnu_Tool_ListLikeChar.Enabled = False
End If

'If mnu_±`¥Î¦r.Checked = True Then
'    mnuToolListChar.Caption = "¦C¥X©Ò¦³±`¥Î¦r..."
'    mnuToolListChar.Enabled = ±Ò°Ê¦r§Î´F¨Å
'ElseIf mnu_Big5.Checked = True Then
'    mnuToolListChar.Caption = "¦C¥X¤­¤j½X©Ò¦³¦r§Î..."
'    mnuToolListChar.Enabled = ±Ò°Ê¦r§Î´F¨Å
'ElseIf mnu_Â²¤Æ¦rÁ`ªí.Checked = True Then
'    mnuToolListChar.Caption = "¦C¥X©Ò¦³Â²¤Æ¦r..."
'    mnuToolListChar.Enabled = ±Ò°Ê¦r§Î´F¨Å
'ElseIf mnu_º~»y¤j¦r¨å.Checked = True Then
'    mnuToolListChar.Caption = "¦C¥X¡mº~»y¤j¦r¨å¡n©Ò¦³¦r§Î..."
'    mnuToolListChar.Enabled = False
'ElseIf mnu_·¢®Ñ.Checked = True Then
'    mnuToolListChar.Caption = "¦C¥X©Ò¦³·¢®Ñ¦r§Î..."
'    mnuToolListChar.Enabled = False
'ElseIf mnu_¤p½f.Checked = True Then
'    mnuToolListChar.Caption = "¦C¥X¡m»¡¤å¸Ñ¦r¡n©Ò¦³¦r§Î..."
'    mnuToolListChar.Enabled = ±Ò°Ê¦r§Î´F¨Å
'ElseIf mnu_ª÷¤å.Checked = True Then
'    mnuToolListChar.Caption = "¦C¥X¡mª÷¤å½s¡n©Ò¦³¦r§Î..."
'    mnuToolListChar.Enabled = ±Ò°Ê¦r§Î´F¨Å
'ElseIf mnu_ª÷¤å¹Ï§Î¤å¦r.Checked = True Then
'    mnuToolListChar.Caption = "¦C¥X¡mª÷¤å½s¡nªþ¿ý¤W©Ò¦³¦r§Î..."
'    mnuToolListChar.Enabled = ±Ò°Ê¦r§Î´F¨Å
'ElseIf mnu_¥Ò°©¤å.Checked = True Then
'    mnuToolListChar.Caption = "¦C¥X¡m®ï¼V¥Ò°©¨èÃãÃþÄ¡¡n©Ò¦³¦r§Î..."
'    mnuToolListChar.Enabled = ±Ò°Ê¦r§Î´F¨Å
'ElseIf mnu_·¡¨t¤å¦r.Checked = True Then
'    mnuToolListChar.Caption = "¦C¥X¡m·¡¨tÂ²©­¤å¦r½s¡n©Ò¦³¦r§Î..."
'    mnuToolListChar.Enabled = ±Ò°Ê¦r§Î´F¨Å
'End If

End Sub

Private Sub mnu_Tool_ListLikeChar_Click()

Dim tagFont As Integer, tagDuplicate As Integer, tagEnd As Integer

If Not ±Ò°Ê¥X³BÀË¦r Then frm¥X³BÀË¦r.show
frm¥X³BÀË¦r.SetFocus

tagFont = InStr(1, txtºc¦r¦¡, "Æ¡")
tagDuplicate = InStr(1, txtºc¦r¦¡, ";")
tagEnd = InStr(1, txtºc¦r¦¡, "ý")

If tagDuplicate = 0 Then
   frm¥X³BÀË¦r.cbo¥X³B = Mid(txtºc¦r¦¡, tagFont + 1, tagEnd - tagFont - 1)
Else
   frm¥X³BÀË¦r.cbo¥X³B = Mid(txtºc¦r¦¡, tagFont + 1, tagDuplicate - tagFont - 1)
End If

frm¥X³BÀË¦r.cbo¥X³B_KeyPress vbKeyReturn

End Sub

Private Sub mnu_Unicode¿ï¶µ_Click()

If mnu_Unicode¿ï¶µ.Checked = True Then
   mnu_Unicode¿ï¶µ.Checked = False
Else
   mnu_Unicode¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_¤K¨ö_Click()

¦@¥Îµøµ¡¥N½X = ¤K¨ö¥N½X
frm³¡¥ó½d¨Ò.Form_Load
frm³¡¥ó½d¨Ò.show
frm³¡¥ó½d¨Ò.SetFocus

End Sub

Private Sub mnu_¤p½f_Click()

If mnu_¤p½f.Checked = False Then
   mnu_±`¥Î¦r.Checked = False
   mnu_Big5.Checked = False
   mnu_Â²¤Æ¦rÁ`ªí.Checked = False
   mnu_º~»y¤j¦r¨å.Checked = False
   mnu_·¢®Ñ.Checked = False
   mnu_¤p½f.Checked = True
   mnu_ª÷¤å.Checked = False
   mnu_ª÷¤å¹Ï§Î¤å¦r.Checked = False
   mnu_¥Ò°©¤å.Checked = False
   mnu_·¡¨t¤å¦r.Checked = False
   If ±Ò°Ê¦r§Î´F¨Å Then frm¦r§Î´F¨Å.txtºc¦r¦¡.FontName = "¥_®v¤j»¡¤å¤p½f"
   mnu_·¢®Ñ.Checked = False
   ¨t²Î¦rÅé = "¤p½f"
End If

End Sub

Private Sub mnu_¤p½f¿ï¶µ_Click()

If mnu_¤p½f¿ï¶µ.Checked = True Then
   mnu_¤p½f¿ï¶µ.Checked = False
Else
   mnu_¤p½f¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_¤¤¤å¤jÃã¨å¿ï¶µ_Click()

If mnu_¤¤¤å¤jÃã¨å¿ï¶µ.Checked = True Then
   mnu_¤¤¤å¤jÃã¨å¿ï¶µ.Checked = False
Else
   mnu_¤¤¤å¤jÃã¨å¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_¤¤µØ»¡¤å¸Ñ¦r¿ï¶µ_Click()

If mnu_¤¤µØ»¡¤å¸Ñ¦r¿ï¶µ.Checked = True Then
   mnu_¤¤µØ»¡¤å¸Ñ¦r¿ï¶µ.Checked = False
Else
   mnu_¤¤µØ»¡¤å¸Ñ¦r¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_¥X³BÀË¦r_Click()

¦@¥Îµøµ¡¥N½X = ¥X³BÀË¦r¥N½X
frm¥X³BÀË¦r.show
frm¥X³BÀË¦r.SetFocus

End Sub

Private Sub mnu_¥Ò°©¤å_Click()

If mnu_¥Ò°©¤å.Checked = False Then
   mnu_±`¥Î¦r.Checked = False
   mnu_Big5.Checked = False
   mnu_Â²¤Æ¦rÁ`ªí.Checked = False
   mnu_º~»y¤j¦r¨å.Checked = False
   mnu_·¢®Ñ.Checked = False
   mnu_¤p½f.Checked = False
   mnu_ª÷¤å.Checked = False
   mnu_ª÷¤å¹Ï§Î¤å¦r.Checked = False
   mnu_¥Ò°©¤å.Checked = True
   mnu_·¡¨t¤å¦r.Checked = False
   If ±Ò°Ê¦r§Î´F¨Å Then frm¦r§Î´F¨Å.txtºc¦r¦¡.FontName = "¤¤¬ã°|¥Ò°©¤å"
   mnu_·¢®Ñ.Checked = False
   ¨t²Î¦rÅé = "¥Ò°©¤å"
End If

End Sub

Private Sub mnu_ª÷¤å¹Ï§Î¤å¦r_Click()

If mnu_ª÷¤å¹Ï§Î¤å¦r.Checked = False Then
   mnu_±`¥Î¦r.Checked = False
   mnu_Big5.Checked = False
   mnu_Â²¤Æ¦rÁ`ªí.Checked = False
   mnu_º~»y¤j¦r¨å.Checked = False
   mnu_·¢®Ñ.Checked = False
   mnu_¤p½f.Checked = False
   mnu_ª÷¤å.Checked = False
   mnu_ª÷¤å¹Ï§Î¤å¦r.Checked = True
   mnu_¥Ò°©¤å.Checked = False
   mnu_·¡¨t¤å¦r.Checked = False
   If ±Ò°Ê¦r§Î´F¨Å Then frm¦r§Î´F¨Å.txtºc¦r¦¡.FontName = "¤¤¬ã°|ª÷¤å"
   mnu_·¢®Ñ.Checked = False
   ¨t²Î¦rÅé = "ª÷¤å"
End If

End Sub

Private Sub mnu_´¼¼z«¬½Æ»s¨ìWord_Click()

If mnu_´¼¼z«¬½Æ»s¨ìWord.Checked = False Then
   mnu_½Æ»s¦r§Î¨ìWord.Checked = False
   mnu_½Æ»s¹Ï¤ù¨ìWord.Checked = False
   mnu_´¼¼z«¬½Æ»s¨ìWord.Checked = True
End If

End Sub

Private Sub mnu_¶K¹Ï¨ìWord_Click()

On Error GoTo ExitSub

mnu_½Æ»s¹Ï¤ù_Click
If Len(¼È¦s¹ÏÀÉ) > 0 Then
    WordApp.ActiveDocument.InlineShapes.AddPicture ¼È¦s¹ÏÀÉ, False, True, WordApp.Selection.Range
     WordApp.Selection.MoveRight
'                ¿ï¨úºc¦r¦¡.SetRange Start:=¿ï¨úºc¦r¦¡.End, End:=¿ï¨úºc¦r¦¡.End + 1
'                ¿ï¨úºc¦r¦¡.font = ¦r«¬
    WordApp.ActiveDocument.InlineShapes(WordApp.ActiveDocument.InlineShapes.Count).AlternativeText = "¡´¡×" & ´À¥N¤å¦r
    WordApp.Selection.Paragraphs.BaseLineAlignment = wdBaselineAlignCenter
End If

ExitSub:

End Sub

Private Sub mnu_·¡¨t¤å¦r_Click()

If mnu_·¡¨t¤å¦r.Checked = False Then
   mnu_±`¥Î¦r.Checked = False
   mnu_Big5.Checked = False
   mnu_Â²¤Æ¦rÁ`ªí.Checked = False
   mnu_º~»y¤j¦r¨å.Checked = False
   mnu_·¢®Ñ.Checked = False
   mnu_¤p½f.Checked = False
   mnu_ª÷¤å.Checked = False
   mnu_ª÷¤å¹Ï§Î¤å¦r.Checked = False
   mnu_¥Ò°©¤å.Checked = False
   mnu_·¡¨t¤å¦r.Checked = True
   If ±Ò°Ê¦r§Î´F¨Å Then frm¦r§Î´F¨Å.txtºc¦r¦¡.FontName = "¤¤¬ã°|·¡¨tÂ²©­¤å¦r"
   mnu_·¢®Ñ.Checked = False
   ¨t²Î¦rÅé = "·¡¨t¤å¦r"
End If

End Sub

Private Sub mnu_¥Ò°©¤å¦rµþªL¿ï¶µ_Click()

If mnu_¥Ò°©¤å¦rµþªL¿ï¶µ.Checked = True Then
   mnu_¥Ò°©¤å¦rµþªL¿ï¶µ.Checked = False
Else
   mnu_¥Ò°©¤å¦rµþªL¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_¥Ò°©¤å¦r¶°ÄÀ¿ï¶µ_Click()

If mnu_¥Ò°©¤å¦r¶°ÄÀ¿ï¶µ.Checked = True Then
   mnu_¥Ò°©¤å¦r¶°ÄÀ¿ï¶µ.Checked = False
Else
   mnu_¥Ò°©¤å¦r¶°ÄÀ¿ï¶µ.Checked = True
End If

End Sub


Private Sub mnu_·¡¨t¤å¦r¥X³B¿ï¶µ_Click()

If mnu_·¡¨t¤å¦r¥X³B¿ï¶µ.Checked = True Then
   mnu_·¡¨t¤å¦r¥X³B¿ï¶µ.Checked = False
Else
   mnu_·¡¨t¤å¦r¥X³B¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_·¡¨tÂ²©­¤å¦r½s¿ï¶µ_Click()

If mnu_·¡¨tÂ²©­¤å¦r½s¿ï¶µ.Checked = True Then
   mnu_·¡¨tÂ²©­¤å¦r½s¿ï¶µ.Checked = False
Else
   mnu_·¡¨tÂ²©­¤å¦r½s¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_¥Ò°©¤å¿ï¶µ_Click()

If mnu_¥Ò°©¤å¿ï¶µ.Checked = True Then
   mnu_¥Ò°©¤å¿ï¶µ.Checked = False
Else
   mnu_¥Ò°©¤å¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_·¡¨t¤å¦r¿ï¶µ_Click()

If mnu_·¡¨t¤å¦r¿ï¶µ.Checked = True Then
   mnu_·¡¨t¤å¦r¿ï¶µ.Checked = False
Else
   mnu_·¡¨t¤å¦r¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_¥Ò°©¨èÃãÃþÄ¡¿ï¶µ_Click()

If mnu_¥Ò°©¨èÃãÃþÄ¡¿ï¶µ.Checked = True Then
   mnu_¥Ò°©¨èÃãÃþÄ¡¿ï¶µ.Checked = False
Else
   mnu_¥Ò°©¨èÃãÃþÄ¡¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_¦r§Î¯Á¤Þ_Click()

¦@¥Îµøµ¡¥N½X = ¦r§Î¯Á¤Þ¥N½X
frm¦r§Î¯Á¤Þ.show
frm¦r§Î¯Á¤Þ.SetFocus

End Sub

Private Sub mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó_Click()

If mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó.Checked = True Then
   mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó.Checked = False
Else
   mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó.Checked = True
End If

End Sub

Private Sub mnu_¦r§Î´F¨Å¥u¦C¥X±`¥Î¦r_Click()

ªì©l¦rÀW = 5

'If mnu_¦r§Î´F¨Å¥u¦C¥X±`¥Î¦r.Checked = False Then
   'mnu_¦r§Î´F¨Å¥u¦C¥X±`¥Î¦r.Checked = True
   'mnu_¦r§Î´F¨Å¥u¦C¥X¹q¸£¥Î¦r.Checked = False
   'mnu_¦r§Î´F¨Å¦C¥X©Ò¦³¦r§Î.Checked = False
'End If

End Sub


Private Sub mnu_¦r§Î´F¨Å¥u¦C¥X¹q¸£¥Î¦r_Click()

ªì©l¦rÀW = 2

'If mnu_¦r§Î´F¨Å¥u¦C¥X¹q¸£¥Î¦r.Checked = False Then
   'mnu_¦r§Î´F¨Å¥u¦C¥X¹q¸£¥Î¦r.Checked = True
   'mnu_¦r§Î´F¨Å¥u¦C¥X±`¥Î¦r.Checked = False
   'mnu_¦r§Î´F¨Å¦C¥X©Ò¦³¦r§Î.Checked = False
'End If

End Sub

Private Sub mnu_¦r§Î´F¨Å¦C¥X©Ò¦³¦r§Î_Click()

ªì©l¦rÀW = 1

'If mnu_¦r§Î´F¨Å¦C¥X©Ò¦³¦r§Î.Checked = False Then
   'mnu_¦r§Î´F¨Å¦C¥X©Ò¦³¦r§Î.Checked = True
   'mnu_¦r§Î´F¨Å¥u¦C¥X±`¥Î¦r.Checked = False
   'mnu_¦r§Î´F¨Å¥u¦C¥X¹q¸£¥Î¦r.Checked = False
'End If

End Sub

Private Sub mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk_Click()

If mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Checked = True Then
   mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Checked = False
   frm¦r§Î´F¨Å.Caption = "³¡¥óÀË¦r"
Else
   mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Checked = True
   frm¦r§Î´F¨Å.Caption = "³¡¥óÀË¦r(SQL Like)"
End If

End Sub

Private Sub mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó_Click()

ªì©l³v¯Å¦C¥X = 1

If mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Checked = False Then
   mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Checked = True
   mnu_¦r§Î´F¨Å¥u¦C¦r§Î¤£¦C³¡¥ó.Checked = False
End If

End Sub


Private Sub mnu_¦r§Î´F¨Å¥u¦C¦r§Î¤£¦C³¡¥ó_Click()

ªì©l³v¯Å¦C¥X = 0

If mnu_¦r§Î´F¨Å¥u¦C¦r§Î¤£¦C³¡¥ó.Checked = False Then
   mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Checked = False
   mnu_¦r§Î´F¨Å¥u¦C¦r§Î¤£¦C³¡¥ó.Checked = True
End If

End Sub


Private Sub mnu_¦r§Î´F¨Å¿í·Ó¿é¤J³¡¥ó¶¶§Ç_Click()

If mnu_¦r§Î´F¨Å¿í·Ó¿é¤J³¡¥ó¶¶§Ç.Checked = True Then
   mnu_¦r§Î´F¨Å¿í·Ó¿é¤J³¡¥ó¶¶§Ç.Checked = False
Else
   mnu_¦r§Î´F¨Å¿í·Ó¿é¤J³¡¥ó¶¶§Ç.Checked = True
End If

End Sub

Private Sub mnu_¦r§ÎºtÅÜ_Click()

'mnu_¦r§ÎºtÅÜ.Enabled = False
¦@¥Îµøµ¡¥N½X = ¦r§ÎºtÅÜ¥N½X
frm¦r§ÎºtÅÜ.show
frm¦r§ÎºtÅÜ.SetFocus

End Sub

Private Sub mnu_¦r«¬_Click()

frm¦r«¬.show 1
cbo¦r«¬¦WºÙ_click
cbo¦rÅé¤j¤p_Click

End Sub

Private Sub mnu_ª÷¤å_Click()

If mnu_ª÷¤å.Checked = False Then
   mnu_±`¥Î¦r.Checked = False
   mnu_Big5.Checked = False
   mnu_Â²¤Æ¦rÁ`ªí.Checked = False
   mnu_º~»y¤j¦r¨å.Checked = False
   mnu_·¢®Ñ.Checked = False
   mnu_¤p½f.Checked = False
   mnu_ª÷¤å.Checked = True
   mnu_ª÷¤å¹Ï§Î¤å¦r.Checked = False
   mnu_¥Ò°©¤å.Checked = False
   mnu_·¡¨t¤å¦r.Checked = False
   If ±Ò°Ê¦r§Î´F¨Å Then frm¦r§Î´F¨Å.txtºc¦r¦¡.FontName = "¤¤¬ã°|ª÷¤å"
   mnu_·¢®Ñ.Checked = False
   ¨t²Î¦rÅé = "ª÷¤å"
End If

End Sub

Private Sub mnu_ª÷¤åµþªL¿ï¶µ_Click()

If mnu_ª÷¤åµþªL¿ï¶µ.Checked = True Then
   mnu_ª÷¤åµþªL¿ï¶µ.Checked = False
Else
   mnu_ª÷¤åµþªL¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_ª÷¤å½s¿ï¶µ_Click()

If mnu_ª÷¤å½s¿ï¶µ.Checked = True Then
   mnu_ª÷¤å½s¿ï¶µ.Checked = False
Else
   mnu_ª÷¤å½s¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_ª÷¤å¿ï¶µ_Click()

If mnu_ª÷¤å¿ï¶µ.Checked = True Then
   mnu_ª÷¤å¿ï¶µ.Checked = False
Else
   mnu_ª÷¤å¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_«Ø§»º~»y¤j¦r¨å¿ï¶µ_Click()

If mnu_«Ø§»º~»y¤j¦r¨å¿ï¶µ.Checked = True Then
   mnu_«Ø§»º~»y¤j¦r¨å¿ï¶µ.Checked = False
Else
   mnu_«Ø§»º~»y¤j¦r¨å¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_®ï©Pª÷¤å¶°¦¨¾¹¸¹¿ï¶µ_Click()

If mnu_®ï©Pª÷¤å¶°¦¨¾¹¸¹¿ï¶µ.Checked = True Then
   mnu_®ï©Pª÷¤å¶°¦¨¾¹¸¹¿ï¶µ.Checked = False
Else
   mnu_®ï©Pª÷¤å¶°¦¨¾¹¸¹¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_®ï©Pª÷¤å¶°¦¨¤Þ±o¿ï¶µ_Click()

If mnu_®ï©Pª÷¤å¶°¦¨¤Þ±o¿ï¶µ.Checked = True Then
   mnu_®ï©Pª÷¤å¶°¦¨¤Þ±o¿ï¶µ.Checked = False
Else
   mnu_®ï©Pª÷¤å¶°¦¨¤Þ±o¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_±`¥Î¦r_Click()

ªì©l¦rÀW = 5

If mnu_±`¥Î¦r.Checked = False Then
   mnu_±`¥Î¦r.Checked = True
   mnu_Big5.Checked = False
   mnu_Â²¤Æ¦rÁ`ªí.Checked = False
   mnu_º~»y¤j¦r¨å.Checked = False
   mnu_·¢®Ñ.Checked = False
   mnu_¤p½f.Checked = False
   mnu_ª÷¤å.Checked = False
   mnu_ª÷¤å¹Ï§Î¤å¦r.Checked = False
   mnu_¥Ò°©¤å.Checked = False
   mnu_·¡¨t¤å¦r.Checked = False
   If ±Ò°Ê¦r§Î´F¨Å Then frm¦r§Î´F¨Å.txtºc¦r¦¡.FontName = "¼Ð·¢Åé"
   ¨t²Î¦rÅé = "·¢®Ñ"
End If

End Sub

Private Sub mnu_±dº³¦r¨å³¡­º_Click()

¦@¥Îµøµ¡¥N½X = ±dº³³¡­º¥N½X
frm³¡¥ó½d¨Ò.Form_Load
frm³¡¥ó½d¨Ò.show
frm³¡¥ó½d¨Ò.SetFocus

End Sub

Private Sub mnu_²§Åé¦r®Ú_Click()
'mnu_²§Åé¦r®Ú.Enabled = False
¦@¥Îµøµ¡¥N½X = ²§Åé¦r®Ú¥N½X
frm²§Åé¦r®Ú.show
frm²§Åé¦r®Ú.SetFocus

End Sub

Private Sub mnu_¹Ï¤ù_Click()

frm¹Ï¤ù³]©w.show 1

End Sub

Private Sub mnu_ºc¦r²Å¸¹_Click()

¦@¥Îµøµ¡¥N½X = ºc¦r²Å¸¹¥N½X
frm³¡¥ó½d¨Ò.Form_Load
frm³¡¥ó½d¨Ò.show
frm³¡¥ó½d¨Ò.SetFocus

End Sub

Private Sub mnu_¹Ï§Î¤å¦r_Click()

¦@¥Îµøµ¡¥N½X = ¹Ï§Î¤å¦r¥N½X
frm³¡¥ó½d¨Ò.Form_Load
frm³¡¥ó½d¨Ò.show
frm³¡¥ó½d¨Ò.SetFocus

End Sub


Private Sub mnu_³¡¥ó¥N½X_Click(Index As Integer)
Dim i As Integer

'For i = ²Å¸¹¥N½X To ³¡¥ó¥~¦r¥N½X
'    mnu_³¡¥ó¥N½X(i).Checked = False
'Next i

'mnu_³¡¥ó¥N½X(Index).Checked = True
¦@¥Îµøµ¡¥N½X = Index
frm³¡¥ó½d¨Ò.Form_Load
frm³¡¥ó½d¨Ò.show
frm³¡¥ó½d¨Ò.SetFocus

End Sub

Private Sub mnu_µ²§ô_Click()
'­pºâµ²§ôµøµ¡
Unload mdiº~¦r¦r§Î
End

End Sub

Private Sub mnu_¦r§Î´F¨Å_Click()
'mnu_¦r§Î´F¨Å.Enabled = False
¦@¥Îµøµ¡¥N½X = ¦r§Î´F¨Å¥N½X
frm¦r§Î´F¨Å.show
frm¦r§Î´F¨Å.SetFocus

End Sub

Private Sub mnu_¦r§Îµ²ºc_Click()
'mnu_¦r§Îµ²ºc.Enabled = False
¦@¥Îµøµ¡¥N½X = ¦r§Îµ²ºc¥N½X
frm¦r§Îµ²ºc.show
frm¦r§Îµ²ºc.SetFocus

End Sub

Private Sub mnu_²§Åé¦rªí_Click()
'mnu_²§Åé¦rªí.Enabled = False
¦@¥Îµøµ¡¥N½X = ²§Åé¦rªí¥N½X
frm²§Åé¦rªí.show
frm²§Åé¦rªí.SetFocus

End Sub


Private Sub mnu_¤ô¥­¨Ã±Æ_Click()
mdiº~¦r¦r§Î.Arrange 1
End Sub

Private Sub mnu_««ª½¨Ã±Æ_Click()
'frm²§Åé¦rªí.SetFocus
'frm¦r§Îµ²ºc.SetFocus
'frm¦r§Î´F¨Å.SetFocus

mdiº~¦r¦r§Î.Arrange 2
End Sub

Private Sub mnu_­«Å|Åã¥Ü_Click()
'frm²§Åé¦rªí.SetFocus
'frm¦r§Îµ²ºc.SetFocus
'frm¦r§Î´F¨Å.SetFocus

mdiº~¦r¦r§Î.Arrange 0
End Sub

Private Sub mnu_±Æ¦C¹Ï¥Ü_Click()
mdiº~¦r¦r§Î.Arrange 3
End Sub

Private Sub mnu_³Ì¨ÎÂsÄý_click()

frm¹w³]ÂsÄý.show 1

If Not §ïÅÜ¹w³]ÂsÄý Then Exit Sub

Select Case ¹w³]ÂsÄý¼Ò¦¡
    Case 1: ¹w³]ÂsÄý¤@
    Case 2: ¹w³]ÂsÄý¤G
    Case 3: ¹w³]ÂsÄý¤T
    Case 4: ¹w³]ÂsÄý¥|
End Select

End Sub

Private Sub mnu_·¢®Ñ_Click()

ªì©l¦rÀW = 1

If mnu_·¢®Ñ.Checked = False Then
   mnu_±`¥Î¦r.Checked = False
   mnu_Big5.Checked = False
   mnu_Â²¤Æ¦rÁ`ªí.Checked = False
   mnu_º~»y¤j¦r¨å.Checked = False
   mnu_·¢®Ñ.Checked = True
   mnu_¤p½f.Checked = False
   mnu_ª÷¤å.Checked = False
   mnu_ª÷¤å¹Ï§Î¤å¦r.Checked = False
   mnu_¥Ò°©¤å.Checked = False
   mnu_·¡¨t¤å¦r.Checked = False
   If ±Ò°Ê¦r§Î´F¨Å Then frm¦r§Î´F¨Å.txtºc¦r¦¡.FontName = "¼Ð·¢Åé"
   ¨t²Î¦rÅé = "·¢®Ñ"
End If

End Sub

Private Sub mnu_º~»y¤j¦r¨å_Click()

ªì©l¦rÀW = 1

If mnu_º~»y¤j¦r¨å.Checked = False Then
   mnu_±`¥Î¦r.Checked = False
   mnu_Big5.Checked = False
   mnu_Â²¤Æ¦rÁ`ªí.Checked = False
   mnu_º~»y¤j¦r¨å.Checked = True
   mnu_·¢®Ñ.Checked = False
   mnu_¤p½f.Checked = False
   mnu_ª÷¤å.Checked = False
   mnu_ª÷¤å¹Ï§Î¤å¦r.Checked = False
   mnu_¥Ò°©¤å.Checked = False
   mnu_·¡¨t¤å¦r.Checked = False
   If ±Ò°Ê¦r§Î´F¨Å Then frm¦r§Î´F¨Å.txtºc¦r¦¡.FontName = "¼Ð·¢Åé"
   ¨t²Î¦rÅé = "·¢®Ñ"
End If

End Sub

Private Sub mnu_»¡¤å¸Ñ¦r³¡­º_Click()

¦@¥Îµøµ¡¥N½X = »¡¤å³¡­º¥N½X
frm³¡¥ó½d¨Ò.Form_Load
frm³¡¥ó½d¨Ò.show
frm³¡¥ó½d¨Ò.SetFocus

End Sub

Private Sub mnu_»¡¤å¸Ñ¦rµþªL¿ï¶µ_Click()

If mnu_»¡¤å¸Ñ¦rµþªL¿ï¶µ.Checked = True Then
   mnu_»¡¤å¸Ñ¦rµþªL¿ï¶µ.Checked = False
Else
   mnu_»¡¤å¸Ñ¦rµþªL¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_»¡¤å¸Ñ§Î¶¶§Ç_Click()

If mnu_»¡¤å¸Ñ§Î¶¶§Ç.Checked = False Then
   mnu_»¡¤å¸Ñ§Î¶¶§Ç.Checked = True
   ªì©l¸Ñ§Î¦C¥X = 1
Else
   mnu_»¡¤å¸Ñ§Î¶¶§Ç.Checked = False
   ªì©l¸Ñ§Î¦C¥X = 0
End If

End Sub

Private Sub mnu_»·ªFº~»y¤j¦r¨å¿ï¶µ_Click()

If mnu_»·ªFº~»y¤j¦r¨å¿ï¶µ.Checked = True Then
   mnu_»·ªFº~»y¤j¦r¨å¿ï¶µ.Checked = False
Else
   mnu_»·ªFº~»y¤j¦r¨å¿ï¶µ.Checked = True
End If

End Sub

Private Sub mnu_½Æ»s_Click()

Dim ¾ðª¬µ²ºc As TList, ¸`ÂI¼Ð°O As String, ¸`ÂIÃþ§O As Integer, CopyText As String

CopyText = ""
¸`ÂIÃþ§O = -1
½Æ»sBig5¦r¤¸ = False

Select Case ²{¥Îµøµ¡¥N½X

Case mdiº~¦r¦r§Î¥N½X

    Select Case ²{¥Î±±¨î¶µ¥N½X
    
    Case mdiº~¦r¦r§Î_½s¸¹¤è¶ô
        CopyText = txt½s¸¹.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_¥~¦r¶°¤è¶ô
        CopyText = txt¥~¦r¶°.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_¦r§Î¤è¶ô
        CopyText = txt¦r§Î.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_­«¤å¤è¶ô
        CopyText = txt­«¤å.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_¥jº~¦r¤è¶ô
        CopyText = txt¥jº~¦r.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_Á`µ§µe¤è¶ô
        CopyText = txtÁ`µ§µe.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_³¡­º¤è¶ô
        CopyText = txt³¡­º.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_¦©°£³¡­ºµ§µe¤è¶ô
        CopyText = txt¦©°£³¡­ºµ§µe.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_ª`­µ¤è¶ô
        CopyText = txtª`­µ.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_¤º½X¤è¶ô
        CopyText = txt¤º½X.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_­Ü¾e½X¤è¶ô
        CopyText = txt­Ü¾e½X.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_ºc¦r¦¡¤è¶ô
        CopyText = txtºc¦r¦¡.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_¥U¼Æ¤è¶ô
        CopyText = txt¥U¼Æ.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_²Õ¦r¦r¼Æ¤è¶ô
        CopyText = txt²Õ¦r¦r¼Æ.SelText
        GoTo CopyBegin
    
    Case mdiº~¦r¦r§Î_²Õ¦r¦r¼Æ§t²§¼g¤è¶ô
        CopyText = txt²Õ¦r¦r¼Æ§t²§¼g.SelText
        GoTo CopyBegin
    
    Case Else
        Exit Sub
        
    End Select

Case Big5¦r®Ú¥N½X To ºc¦r²Å¸¹¥N½X
    If frm³¡¥ó½d¨Ò.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
        Set ¾ðª¬µ²ºc = frm³¡¥ó½d¨Ò.tree¦r§Î¾ðª¬µ²ºc
    Else
        Exit Sub
    End If


Case ¦r§Î´F¨Å¥N½X
    If ²{¥Î±±¨î¶µ¥N½X = ¦r§Î´F¨Å_ÀË¦r¤è¶ô Then
        CopyText = frm¦r§Î´F¨Å.txtºc¦r¦¡.SelText
        GoTo CopyBegin
    ElseIf ²{¥Î±±¨î¶µ¥N½X = ¦r§Î´F¨Å_¾ðª¬µ²ºc Then
        If frm¦r§Î´F¨Å.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
            Set ¾ðª¬µ²ºc = frm¦r§Î´F¨Å.tree¦r§Î¾ðª¬µ²ºc
        End If
    Else
        Exit Sub
    End If
    
Case ¥X³BÀË¦r¥N½X
    If ²{¥Î±±¨î¶µ¥N½X = ¥X³BÀË¦r_ÀË¦r¤è¶ô Then
        CopyText = frm¥X³BÀË¦r.cbo¥X³B.SelText
        GoTo CopyBegin
    ElseIf ²{¥Î±±¨î¶µ¥N½X = ¥X³BÀË¦r_¾ðª¬µ²ºc Then
        If frm¥X³BÀË¦r.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
            Set ¾ðª¬µ²ºc = frm¥X³BÀË¦r.tree¦r§Î¾ðª¬µ²ºc
        End If
    Else
        Exit Sub
    End If
    
Case ¦r§Îµ²ºc¥N½X
    If frm¦r§Îµ²ºc.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
        Set ¾ðª¬µ²ºc = frm¦r§Îµ²ºc.tree¦r§Î¾ðª¬µ²ºc
    Else
        Exit Sub
    End If
    
Case ¦r§ÎºtÅÜ¥N½X
    If frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
        Set ¾ðª¬µ²ºc = frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc
    Else
        Exit Sub
    End If
    
Case ²§Åé¦rªí¥N½X
    If frm²§Åé¦rªí.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
        Set ¾ðª¬µ²ºc = frm²§Åé¦rªí.tree¦r§Î¾ðª¬µ²ºc
    Else
        Exit Sub
    End If
    
Case ²§Åé¦r®Ú¥N½X
    If frm²§Åé¦r®Ú.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
        Set ¾ðª¬µ²ºc = frm²§Åé¦r®Ú.tree¦r§Î¾ðª¬µ²ºc
    Else
        Exit Sub
    End If
    
Case ¦r§Î¯Á¤Þ¥N½X
    If frm¦r§Î¯Á¤Þ.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
        Set ¾ðª¬µ²ºc = frm¦r§Î¯Á¤Þ.tree¦r§Î¾ðª¬µ²ºc
    Else
        Exit Sub
    End If
    
Case Else
    Exit Sub
End Select

¸`ÂI¼Ð°O = ¾ðª¬µ²ºc.ItemTag(¾ðª¬µ²ºc.ListIndex)
If Len(¸`ÂI¼Ð°O) > 0 Then ¸`ÂIÃþ§O = CInt(Left(¸`ÂI¼Ð°O, 1))

Select Case ¸`ÂIÃþ§O

Case ¦r§Î¸`ÂI¼Ð°O
    If ½Æ»s­·®æ½X Then
        CopyText = txtºc¦r¦¡
    Else
        CopyText = txt¦r§Î
    End If
Case ºc¦r¦¡¸`ÂI¼Ð°O
    CopyText = txtºc¦r¦¡
Case ¾¹¦W¸`ÂI¼Ð°O
    CopyText = Right(¸`ÂI¼Ð°O, Len(¸`ÂI¼Ð°O) - 1)
Case ¨ä¥L¸`ÂI¼Ð°O
    CopyText = ¾ðª¬µ²ºc.List(¾ðª¬µ²ºc.ListIndex)

End Select

On Error GoTo CopyErr
CopyBegin:

If Len(CopyText) > 0 Then
    Clipboard.Clear
    If Len(CopyText) = 1 Then
        ·¢®ÑÀË¦rªí.Index = "¦r§Î"
        ·¢®ÑÀË¦rªí.Seek "=", CopyText
        If Not ·¢®ÑÀË¦rªí.NoMatch Then
            If ·¢®ÑÀË¦rªí.Fields("½s¸¹") <= 13060 Then ½Æ»sBig5¦r¤¸ = True
        End If
    End If
    Clipboard.SetText CopyText
End If

CopyErr:

End Sub

Private Sub mnu_½Æ»sUnicode¦r§Î¨ìWord_Click()

If mnu_½Æ»sUnicode¦r§Î¨ìWord.Checked = True Then
   mnu_½Æ»sUnicode¦r§Î¨ìWord.Checked = False
Else
   mnu_½Æ»sUnicode¦r§Î¨ìWord.Checked = True
End If

End Sub

Private Sub mnu_½Æ»s¦r§Î¨ìWord_Click()

If mnu_½Æ»s¦r§Î¨ìWord.Checked = False Then
   mnu_½Æ»s¦r§Î¨ìWord.Checked = True
   mnu_½Æ»s¹Ï¤ù¨ìWord.Checked = False
   mnu_´¼¼z«¬½Æ»s¨ìWord.Checked = False
End If

End Sub


Private Sub mnu_½Æ»s¯S®í¹Ï¤ù_Click()

Dim ¦r«¬ As CDPFONT, ¦r§Î As String, Success As Boolean

¦r«¬.Name = "cdp000"
¦r«¬.Size = 9
¦r§Î = "¤X"
´À¥N¤å¦r = "„Ð"

bmpcount = bmpcount + 1
¼È¦s¹ÏÀÉ = ¼È¦s¥Ø¿ý & "\" & "~hz" & CStr(bmpcount) & ".bmp"
¦r§ÎÂà¦¨¹Ï¤ù ¦r«¬, ¦r§Î, ¼È¦s¹ÏÀÉ, ¹Ï¤ù¸ÑªR«×, Success

WordApp.ActiveDocument.InlineShapes.AddPicture ¼È¦s¹ÏÀÉ, False, True, WordApp.Selection.Range
WordApp.Selection.MoveRight
WordApp.ActiveDocument.InlineShapes(WordApp.ActiveDocument.InlineShapes.Count).AlternativeText = "¡´¡×" & ´À¥N¤å¦r
WordApp.Selection.Paragraphs.BaseLineAlignment = wdBaselineAlignCenter
    
End Sub

Private Sub mnu_½Æ»s¹Ï¤ù_Click()

Dim ¾ðª¬µ²ºc As TList, ¦r«¬ As CDPFONT, ¦r§Î As String, Success As Boolean

Select Case ²{¥Îµøµ¡¥N½X

Case Big5¦r®Ú¥N½X To ºc¦r²Å¸¹¥N½X
    If frm³¡¥ó½d¨Ò.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
        Set ¾ðª¬µ²ºc = frm³¡¥ó½d¨Ò.tree¦r§Î¾ðª¬µ²ºc
    Else
        Exit Sub
    End If

Case ¦r§Î´F¨Å¥N½X
    If ²{¥Î±±¨î¶µ¥N½X = ¦r§Î´F¨Å_ÀË¦r¤è¶ô Then
        GoTo CopyGlyphErr
    ElseIf ²{¥Î±±¨î¶µ¥N½X = ¦r§Î´F¨Å_¾ðª¬µ²ºc Then
        If frm¦r§Î´F¨Å.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
            Set ¾ðª¬µ²ºc = frm¦r§Î´F¨Å.tree¦r§Î¾ðª¬µ²ºc
        End If
    Else
        Exit Sub
    End If
    
Case ¥X³BÀË¦r¥N½X
    If ²{¥Î±±¨î¶µ¥N½X = ¥X³BÀË¦r_ÀË¦r¤è¶ô Then
        GoTo CopyGlyphErr
    ElseIf ²{¥Î±±¨î¶µ¥N½X = ¥X³BÀË¦r_¾ðª¬µ²ºc Then
        If frm¥X³BÀË¦r.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
            Set ¾ðª¬µ²ºc = frm¥X³BÀË¦r.tree¦r§Î¾ðª¬µ²ºc
        End If
    Else
        Exit Sub
    End If
    
Case ¦r§Îµ²ºc¥N½X
    If frm¦r§Îµ²ºc.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
        Set ¾ðª¬µ²ºc = frm¦r§Îµ²ºc.tree¦r§Î¾ðª¬µ²ºc
    Else
        Exit Sub
    End If
    
Case ¦r§ÎºtÅÜ¥N½X
    If frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
        Set ¾ðª¬µ²ºc = frm¦r§ÎºtÅÜ.tree¦r§Î¾ðª¬µ²ºc
    Else
        Exit Sub
    End If
    
Case ²§Åé¦rªí¥N½X
    If frm²§Åé¦rªí.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
        Set ¾ðª¬µ²ºc = frm²§Åé¦rªí.tree¦r§Î¾ðª¬µ²ºc
    Else
        Exit Sub
    End If
    
Case ²§Åé¦r®Ú¥N½X
    If frm²§Åé¦r®Ú.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
        Set ¾ðª¬µ²ºc = frm²§Åé¦r®Ú.tree¦r§Î¾ðª¬µ²ºc
    Else
        Exit Sub
    End If
    
Case ¦r§Î¯Á¤Þ¥N½X
    If frm¦r§Î¯Á¤Þ.tree¦r§Î¾ðª¬µ²ºc.ListIndex > -1 Then
        Set ¾ðª¬µ²ºc = frm¦r§Î¯Á¤Þ.tree¦r§Î¾ðª¬µ²ºc
    Else
        Exit Sub
    End If
    
Case Else
    Exit Sub
End Select


¦r«¬.Name = ¾ðª¬µ²ºc.ItemFontName(¾ðª¬µ²ºc.ListIndex)
¦r«¬.Size = ¹Ï¤ù¦r«¬¤j¤p
If ½Æ»s¹Ï¤ù¨ìWord And ½Æ»s¨ìWordªº¹Ï¤ù¤j¤p > 0 Then
    ¦r«¬.Size = ½Æ»s¨ìWordªº¹Ï¤ù¤j¤p
End If

¦r«¬.Bold = ¾ðª¬µ²ºc.ItemFontBold(¾ðª¬µ²ºc.ListIndex)
¦r«¬.Italic = ¾ðª¬µ²ºc.ItemFontItalic(¾ðª¬µ²ºc.ListIndex)
¦r«¬.Underline = ¾ðª¬µ²ºc.ItemFontUnder(¾ðª¬µ²ºc.ListIndex)
¦r«¬.StrikeThrough = ¾ðª¬µ²ºc.ItemFontStrike(¾ðª¬µ²ºc.ListIndex)
¦r«¬.color = 0

¦r§Î = ¾ðª¬µ²ºc.List(¾ðª¬µ²ºc.ListIndex)

If Len(¦r§Î) <> 1 Then GoTo CopyGlyphErr

bmpcount = bmpcount + 1
¼È¦s¹ÏÀÉ = ¼È¦s¥Ø¿ý & "\" & "~hz" & CStr(bmpcount) & ".bmp"
¦r§ÎÂà¦¨¹Ï¤ù ¦r«¬, ¦r§Î, ¼È¦s¹ÏÀÉ, ¹Ï¤ù¸ÑªR«×, Success

´À¥N¤å¦r = Clipboard.GetText

Clipboard.Clear
Clipboard.SetData LoadPicture(¼È¦s¹ÏÀÉ), vbCFBitmap

On Error GoTo CopyGlyphErr

CopyGlyphErr:

End Sub

Private Sub mnu_½Æ»s¹Ï¤ù¨ìWord_Click()

If mnu_½Æ»s¹Ï¤ù¨ìWord.Checked = False Then
   mnu_½Æ»s¦r§Î¨ìWord.Checked = False
   mnu_½Æ»s¹Ï¤ù¨ìWord.Checked = True
   mnu_´¼¼z«¬½Æ»s¨ìWord.Checked = False
End If

End Sub

Private Sub mnu_¿ï¶µ_Click()

If ±Ò°Ê¦r§Î´F¨Å Then
    If mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Checked = True Then
        mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Enabled = False
        mnu_¦r§Î´F¨Å¥u¦C¦r§Î¤£¦C³¡¥ó.Enabled = False
    Else
        mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Enabled = True
        mnu_¦r§Î´F¨Å¥u¦C¦r§Î¤£¦C³¡¥ó.Enabled = True
    End If
    If ¨t²Î¦rÅé = "·¢®Ñ" Then
        mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó.Enabled = True
    Else
        mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó.Enabled = False
    End If
    mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Enabled = True
Else
    mnu_¦r§Î´F¨Å³v¯Å¦C¥X³æ¤@³¡¥ó.Enabled = False
    mnu_¦r§Î´F¨Å¥u¦C¦r§Î¤£¦C³¡¥ó.Enabled = False
    mnu_¦r§Î´F¨Å¥]§t²§¼g³¡¥ó.Enabled = False
    mnu_¦r§Î´F¨Å±Ä¥ÎSQL»yªk.Enabled = False
End If

End Sub

Private Sub mnu_Àx¦sµøµ¡³]©w_Click()

If mnu_Àx¦sµøµ¡³]©w.Checked = False Then
   mnu_Àx¦sµøµ¡³]©w.Checked = True
Else
   mnu_Àx¦sµøµ¡³]©w.Checked = False
End If

End Sub

Private Sub mnu_»¡©ú¥DÃD_Click()
Dim istring As String, iret As Integer

Screen.MousePointer = ccHourglass
istring = "winhlp32.exe " & App.path & "\cdphanzi.hlp"
Shell istring, 1
Screen.MousePointer = ccDefault

End Sub

Private Sub mnu_½Æ»s¨ì°Å¶KÃ¯_Click()

If mnu_½Æ»s¨ì°Å¶KÃ¯.Checked = False Then
   mnu_½Æ»s¨ì°Å¶KÃ¯.Checked = True
Else
   mnu_½Æ»s¨ì°Å¶KÃ¯.Checked = False
End If

End Sub

Private Sub mnu_Â²¤Æ¦rÁ`ªí_Click()

ªì©l¦rÀW = 1

If mnu_Â²¤Æ¦rÁ`ªí.Checked = False Then
   mnu_±`¥Î¦r.Checked = False
   mnu_Big5.Checked = False
   mnu_Â²¤Æ¦rÁ`ªí.Checked = True
   mnu_º~»y¤j¦r¨å.Checked = False
   mnu_·¢®Ñ.Checked = False
   mnu_¤p½f.Checked = False
   mnu_ª÷¤å.Checked = False
   mnu_ª÷¤å¹Ï§Î¤å¦r.Checked = False
   mnu_¥Ò°©¤å.Checked = False
   mnu_·¡¨t¤å¦r.Checked = False
   If ±Ò°Ê¦r§Î´F¨Å Then frm¦r§Î´F¨Å.txtºc¦r¦¡.FontName = "¼Ð·¢Åé"
   ¨t²Î¦rÅé = "·¢®Ñ"
End If


End Sub

Private Sub mnu_Â²©öÂsÄý_Click()

frmÂ²©öÂsÄý.show 1

If Not §ïÅÜ¹w³]ÂsÄý Then Exit Sub

Select Case ¹w³]ÂsÄý¼Ò¦¡
    Case 1: Â²©öÂsÄý¤@
    Case 2: Â²©öÂsÄý¤G
    Case 3: Â²©öÂsÄý¤T
    Case 4: Â²©öÂsÄý¥|
End Select

End Sub


Private Sub mnu_Â²Ã|_Click()

¦@¥Îµøµ¡¥N½X = Â²Ã|¥N½X
frm³¡¥ó½d¨Ò.Form_Load
frm³¡¥ó½d¨Ò.show
frm³¡¥ó½d¨Ò.SetFocus

End Sub

Private Sub mnu_Åã¥Ü­·®æ½X_Click()

If mnu_Åã¥Ü­·®æ½X.Checked = False Then
   mnu_Åã¥Ü­·®æ½X.Checked = True
Else
   mnu_Åã¥Ü­·®æ½X.Checked = False
End If

End Sub

Private Sub mnuToolListChar_Click()

frm¦r§Î´F¨Å.¦C¥X¿ï©w¦r¶°¤¤ªº©Ò¦³¦r§Î

End Sub

Private Sub mnu¶K¤W_Click()

Select Case ²{¥Îµøµ¡¥N½X

Case mdiº~¦r¦r§Î¥N½X

    If ²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_½s¸¹¤è¶ô Then
        txt½s¸¹.SelText = Clipboard.GetText
    ElseIf ²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_¦r§Î¤è¶ô Then
        txt¦r§Î.SelText = Clipboard.GetText
    ElseIf ²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_ºc¦r¦¡¤è¶ô Then
        txtºc¦r¦¡.SelText = Clipboard.GetText
    End If

Case ¦r§Î´F¨Å¥N½X
    If ²{¥Î±±¨î¶µ¥N½X = ¦r§Î´F¨Å_ÀË¦r¤è¶ô Then
        frm¦r§Î´F¨Å.txtºc¦r¦¡.SelText = Clipboard.GetText
    End If
    
Case ¥X³BÀË¦r¥N½X
    If ²{¥Î±±¨î¶µ¥N½X = ¥X³BÀË¦r_ÀË¦r¤è¶ô Then
        frm¥X³BÀË¦r.cbo¥X³B.Text = Clipboard.GetText
    End If
    
    
Case Else
    Exit Sub
End Select

End Sub

Private Sub mnu½s¿è_Click()

On Error Resume Next
Set WordApp = GetObject(, "word.application")
If Err.Number <> 0 Then
    WordWasNotRunning = True
    Err.Clear
Else
    WordWasNotRunning = False
End If

mnu_edit_½Æ»s¨ìWord.Enabled = False
If Not WordWasNotRunning Then
    If WordApp.Documents.Count > 0 Then mnu_edit_½Æ»s¨ìWord.Enabled = True
End If

End Sub

Private Sub txt¤º½X_GotFocus()

txt¤º½X.SelStart = 0
txt¤º½X.SelLength = Len(txt¤º½X)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_¤º½X¤è¶ô

End Sub

Private Sub txt¥U¼Æ_GotFocus()

txt¥U¼Æ.SelStart = 0
txt¥U¼Æ.SelLength = Len(txt¥U¼Æ.Text)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_¥U¼Æ¤è¶ô

End Sub


Private Sub txt¥jº~¦r_GotFocus()

txt¥jº~¦r.SelStart = 0
txt¥jº~¦r.SelLength = Len(txt¥jº~¦r.Text)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_¥jº~¦r¤è¶ô

End Sub

Private Sub txt¥~¦r¶°_GotFocus()

txt¥~¦r¶°.SelStart = 0
txt¥~¦r¶°.SelLength = Len(txt¥~¦r¶°.Text)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_¥~¦r¶°¤è¶ô

End Sub

Private Sub txt¥~¦r¶°_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case txt¥~¦r¶°.Text

    Case 0: txt¥~¦r¶°.ToolTipText = "¼Ð·¢Åé(²Ó©úÅé)"
    Case 1: txt¥~¦r¶°.ToolTipText = "¼Ð·¢Åé(²Ó©úÅé)¥~¦r¶°¤@"
    Case 2: txt¥~¦r¶°.ToolTipText = "¼Ð·¢Åé(²Ó©úÅé)¥~¦r¶°¤G"
    Case 3: txt¥~¦r¶°.ToolTipText = "¼Ð·¢Åé(²Ó©úÅé)¥~¦r¶°¤T"
    Case 4: txt¥~¦r¶°.ToolTipText = "¼Ð·¢Åé(²Ó©úÅé)¥~¦r¶°¥|"
    Case 5: txt¥~¦r¶°.ToolTipText = "¼Ð·¢Åé(²Ó©úÅé)¥~¦r¶°¤­"
    Case 6: txt¥~¦r¶°.ToolTipText = "¼Ð·¢Åé(²Ó©úÅé)¥~¦r¶°¤»"
    Case 7: txt¥~¦r¶°.ToolTipText = "¼Ð·¢Åé(²Ó©úÅé)¥~¦r¶°¤C"
    Case 8: txt¥~¦r¶°.ToolTipText = "¼Ð·¢Åé(²Ó©úÅé)¥~¦r¶°¤K"
    Case 9: txt¥~¦r¶°.ToolTipText = "¼Ð·¢Åé(²Ó©úÅé)¥~¦r¶°¤E"

End Select

End Sub

Private Sub txt¦r§Î_Change()

txt¦r§Î.SelStart = 0
txt¦r§Î.SelLength = Len(txt¦r§Î.Text)
txt¦r§Î.FontSize = 12

If mnu_½Æ»s¨ì°Å¶KÃ¯.Checked = True Then
   On Error GoTo Label1
   Clipboard.Clear
   Clipboard.SetText txt¦r§Î.Text
Label1: End If

End Sub

Private Sub txt¦r§Î_GotFocus()

txt¦r§Î.SelStart = 0
txt¦r§Î.SelLength = Len(txt¦r§Î.Text)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_¦r§Î¤è¶ô

End Sub

Private Sub txt¦r§Î_KeyPress(KeyAscii As Integer)
Dim ¦r§Î As String, ½s¸¹ As Long, ¦rÅé As Integer, temp As Integer
Dim ·¢®Ñ½s¸¹ As Long, ¤p½f½s¸¹ As Long, ª÷¤å½s¸¹ As Long, ¥Ò°©¤å½s¸¹ As Long, ·¡¨t¤å¦r½s¸¹ As Long
Dim ¼È¦s²Õ¦r¦¡ As String

mdiº~¦r¦r§Î.txt¦r§Î.FontName = "¼Ð·¢Åé"
If KeyAscii = vbKeyReturn Then
   If Len(Trim(txt¦r§Î.Text)) <> 0 Then
      ¦r§Î = txt¦r§Î.Text
      Set ÀË¦rªí = ·¢®ÑÀË¦rªí
      ÀË¦rªí.Index = "¦r§Î"
      ÀË¦rªí.Seek "=", ¦r§Î
      If ÀË¦rªí.NoMatch Then
         ·¢®Ñ½s¸¹ = -1
      Else
         ·¢®Ñ½s¸¹ = ÀË¦rªí.Fields("½s¸¹")
         If Not IsNull(ÀË¦rªí.Fields("¤p½f½s¸¹")) Then
            ¤p½f½s¸¹ = ÀË¦rªí.Fields("¤p½f½s¸¹")
         Else
            ¤p½f½s¸¹ = -1
         End If
         If Not IsNull(ÀË¦rªí.Fields("ª÷¤å½s¸¹")) Then
            ª÷¤å½s¸¹ = ÀË¦rªí.Fields("ª÷¤å½s¸¹")
         Else
            ª÷¤å½s¸¹ = -1
         End If
         If Not IsNull(ÀË¦rªí.Fields("¥Ò°©¤å½s¸¹")) Then
            ¥Ò°©¤å½s¸¹ = ÀË¦rªí.Fields("¥Ò°©¤å½s¸¹")
         Else
            ¥Ò°©¤å½s¸¹ = -1
         End If
         If Not IsNull(ÀË¦rªí.Fields("·¡¨t¤å¦r½s¸¹")) Then
            ·¡¨t¤å¦r½s¸¹ = ÀË¦rªí.Fields("·¡¨t¤å¦r½s¸¹")
         Else
           ·¡¨t¤å¦r½s¸¹ = -1
         End If
      End If
            
      If ·¢®Ñ½s¸¹ > 0 Then
         Â^¨úÄÝ©Ê "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
         Â^¨úºc¦r¦¡ "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
         If ¨t²Î¦rÅé = "¤p½f" And ¤p½f½s¸¹ > 0 Then
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
         ElseIf ¨t²Î¦rÅé = "ª÷¤å" And ª÷¤å½s¸¹ > 0 Then
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
         ElseIf ¨t²Î¦rÅé = "¥Ò°©¤å" And ¥Ò°©¤å½s¸¹ > 0 Then
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
         ElseIf ¨t²Î¦rÅé = "·¡¨t¤å¦r" And ·¡¨t¤å¦r½s¸¹ > 0 Then
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
         Else
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
         End If
      Else
         '§ä¤£¨ì
         mdiº~¦r¦r§Î.txt²Õ¦r¦r¼Æ.Text = ""
         mdiº~¦r¦r§Î.txt¦r§Î.Text = ""
         mdiº~¦r¦r§Î.txtÁ`µ§µe.Text = ""
         mdiº~¦r¦r§Î.txt³¡­º.Text = ""
         mdiº~¦r¦r§Î.txt¦©°£³¡­ºµ§µe.Text = ""
         mdiº~¦r¦r§Î.txtª`­µ.Text = ""
         mdiº~¦r¦r§Î.txt¤º½X.Text = ""
         mdiº~¦r¦r§Î.txt­Ü¾e½X.Text = ""
         'mdiº~¦r¦r§Î.txtºc¦r¦¡.Text = ""
         mdiº~¦r¦r§Î.txt¥U¼Æ.Text = ""
         
         'µ¹©wªÅ¥Õ­È
         If Len(ª¬ºA¦C1) > 10 Then
            ª¬ºA¦C = "0 ­Ó¦r®Ú" & ª¬ºA¦C1
         Else
            ª¬ºA¦C = "0 ­Ó¦r®Ú"
         End If
         mdiº~¦r¦r§Î.txtª¬ºA = ª¬ºA¦C
      End If
   End If
End If
'mdiº~¦r¦r§Î.txt¦r§Î.FontName = "¼Ð·¢Åé"
'If KeyAscii = vbKeyReturn Then
'   ¦r§Î = txt¦r§Î.Text
'   ÀË¦rªí.Index = "¦r§Î"
'   ÀË¦rªí.Seek "=", ¦r§Î
'   If Not ÀË¦rªí.NoMatch() Then
'      ½s¸¹ = ÀË¦rªí.Fields("½s¸¹")
'      Â^¨úÄÝ©Ê "¼Ð·¢Åé", ¦r§Î, ½s¸¹
'      Â^¨úºc¦r¦¡ "¼Ð·¢Åé", ¦r§Î, ½s¸¹
'      If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¼Ð·¢Åé", mdiº~¦r¦r§Î.txt¦r§Î.Text, ½s¸¹
'      If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¼Ð·¢Åé", mdiº~¦r¦r§Î.txt¦r§Î.Text, ½s¸¹
'      If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¼Ð·¢Åé", mdiº~¦r¦r§Î.txt¦r§Î.Text, ½s¸¹
'      If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¼Ð·¢Åé", mdiº~¦r¦r§Î.txt¦r§Î.Text, ½s¸¹
'      If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¼Ð·¢Åé", mdiº~¦r¦r§Î.txt¦r§Î.Text, ½s¸¹
'   Else
         '§ä¤£¨ì
'         mdiº~¦r¦r§Î.txt²Õ¦r¦r¼Æ.Text = ""
'         'mdiº~¦r¦r§Î.txt¦r§Î.Text = ""
'         mdiº~¦r¦r§Î.txtÁ`µ§µe.Text = ""
'         mdiº~¦r¦r§Î.txt³¡­º.Text = ""
'         mdiº~¦r¦r§Î.txt¦©°£³¡­ºµ§µe.Text = ""
'         mdiº~¦r¦r§Î.txtª`­µ.Text = ""
'         mdiº~¦r¦r§Î.txt¤º½X.Text = ""
'         mdiº~¦r¦r§Î.txt­Ü¾e½X.Text = ""
'         mdiº~¦r¦r§Î.txtºc¦r¦¡.Text = ""
'         mdiº~¦r¦r§Î.txt¥U¼Æ.Text = ""
'
         'µ¹©wªÅ¥Õ­È
'         If Len(ª¬ºA¦C1) > 10 Then
'            ª¬ºA¦C = "0 ­Ó¦r®Ú" & ª¬ºA¦C1
'         Else
'            ª¬ºA¦C = "0 ­Ó¦r®Ú"
'         End If
'         mdiº~¦r¦r§Î.txtª¬ºA = ª¬ºA¦C
'   End If
'End If
End Sub

Private Sub txt¦©°£³¡­ºµ§µe_GotFocus()

txt¦©°£³¡­ºµ§µe.SelStart = 0
txt¦©°£³¡­ºµ§µe.SelLength = Len(txt¦©°£³¡­ºµ§µe.Text)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_¦©°£³¡­ºµ§µe¤è¶ô

End Sub

Private Sub txtª`­µ_GotFocus()

txtª`­µ.SelStart = 0
txtª`­µ.SelLength = Len(txtª`­µ.Text)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_ª`­µ¤è¶ô

End Sub




Private Sub txt­«¤å_GotFocus()

txt­«¤å.SelStart = 0
txt­«¤å.SelLength = Len(txt­«¤å.Text)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_­«¤å¤è¶ô

End Sub

Private Sub txt­«¤å_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If IsNumeric(txt­«¤å.Text) Then txt­«¤å.ToolTipText = ¦rÅé°}¦C(CInt(txt­«¤å.Text))

End Sub

Private Sub txt­Ü¾e½X_GotFocus()

txt­Ü¾e½X.SelStart = 0
txt­Ü¾e½X.SelLength = Len(txt­Ü¾e½X.Text)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_­Ü¾e½X¤è¶ô

End Sub

Private Sub txt²Õ¦r¦r¼Æ§t²§¼g_GotFocus()

²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_²Õ¦r¦r¼Æ§t²§¼g¤è¶ô

End Sub

Private Sub txtºc¦r¦¡_Change()
If mnu_½Æ»s¨ì°Å¶KÃ¯.Checked = True Then
   If Not ÀË¦rªí.NoMatch Then
      If ÀË¦rªí.Fields("¦rÅé") <> 0 Then
         On Error GoTo Lable2
         Clipboard.Clear
         Clipboard.SetText txtºc¦r¦¡.Text
Lable2: End If
   End If
   If ½Æ»s­·®æ½X Then Clipboard.SetText txtºc¦r¦¡.Text
End If

End Sub

'Private Sub txtºc¦r¦¡_DragDrop(Source As Control, X As Single, Y As Single)
'Dim ¦r§Î As String
'Dim ¥ª¥b³¡ As String
'Dim ¥k¥b³¡ As String

'¥ª¥b³¡ = Left(txtºc¦r¦¡, txtºc¦r¦¡.SelStart)'
'¥k¥b³¡ = Right$(txtºc¦r¦¡, Len(txtºc¦r¦¡) - txtºc¦r¦¡.SelStart)
  
'If TypeOf Source Is ListBox Then
'   If Source.ListIndex < 0 Then Exit Sub
'   Source.Drag 2       ' End Dragging
'   txtºc¦r¦¡ = ¥ª¥b³¡ & Source.List(Source.ListIndex) & ¥k¥b³¡
'End If

'If TypeOf Source Is TList Then
'   If Source Is Nothing Then Exit Sub
'   Source.Drag 2       ' End Dragging
'   Screen.MousePointer = 11
'   ¦r§Î = Left(Source, 2)
'   'txtºc¦r¦¡ = ¥ª¥b³¡ & mdiº~¦r¦r§Î.txt¦r§Î.Text & ¥k¥b³¡
'   txtºc¦r¦¡ = ¥ª¥b³¡ & ©ì¦²¦r¦ê & ¥k¥b³¡
'   Screen.MousePointer = 0
'End If
'txtºc¦r¦¡.SetFocus
'txtºc¦r¦¡.SelStart = Len(txtºc¦r¦¡)

'End Sub

Private Sub txtºc¦r¦¡_GotFocus()

txtºc¦r¦¡.SelStart = 0
txtºc¦r¦¡.SelLength = Len(txtºc¦r¦¡.Text)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_ºc¦r¦¡¤è¶ô

End Sub

Private Sub txtºc¦r¦¡_KeyPress(KeyAscii As Integer)
Dim ¦r®Ú§Ç As String
Dim ·¢®Ñ½s¸¹ As Long, ¤p½f½s¸¹ As Long, ª÷¤å½s¸¹ As Long, ¥Ò°©¤å½s¸¹ As Long, ·¡¨t¤å¦r½s¸¹ As Long
Dim ¼È¦s²Õ¦r¦¡ As String

If KeyAscii = vbKeyReturn Then
   If Len(Trim(txtºc¦r¦¡.Text)) <> 0 Then
      ·¢®Ñ½s¸¹ = ¦r§Î¬d¸ß()
      If ·¢®Ñ½s¸¹ > 0 Then
        Set ÀË¦rªí = ·¢®ÑÀË¦rªí
        ÀË¦rªí.Index = "½s¸¹"
        ÀË¦rªí.Seek "=", ·¢®Ñ½s¸¹
        If Not IsNull(ÀË¦rªí.Fields("¤p½f½s¸¹")) Then
            ¤p½f½s¸¹ = ÀË¦rªí.Fields("¤p½f½s¸¹")
        Else
            ¤p½f½s¸¹ = -1
        End If
        If Not IsNull(ÀË¦rªí.Fields("ª÷¤å½s¸¹")) Then
            ª÷¤å½s¸¹ = ÀË¦rªí.Fields("ª÷¤å½s¸¹")
        Else
            ª÷¤å½s¸¹ = -1
        End If
        If Not IsNull(ÀË¦rªí.Fields("¥Ò°©¤å½s¸¹")) Then
            ¥Ò°©¤å½s¸¹ = ÀË¦rªí.Fields("¥Ò°©¤å½s¸¹")
        Else
            ¥Ò°©¤å½s¸¹ = -1
        End If
        If Not IsNull(ÀË¦rªí.Fields("·¡¨t¤å¦r½s¸¹")) Then
            ·¡¨t¤å¦r½s¸¹ = ÀË¦rªí.Fields("·¡¨t¤å¦r½s¸¹")
        Else
            ·¡¨t¤å¦r½s¸¹ = -1
        End If

      End If
      
      If ·¢®Ñ½s¸¹ > 0 Then
         Â^¨úÄÝ©Ê "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
         Â^¨úºc¦r¦¡ "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
         If ¨t²Î¦rÅé = "¤p½f" And ¤p½f½s¸¹ > 0 Then
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
         ElseIf ¨t²Î¦rÅé = "ª÷¤å" And ª÷¤å½s¸¹ > 0 Then
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
         ElseIf ¨t²Î¦rÅé = "¥Ò°©¤å" And ¥Ò°©¤å½s¸¹ > 0 Then
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
         ElseIf ¨t²Î¦rÅé = "·¡¨t¤å¦r" And ·¡¨t¤å¦r½s¸¹ > 0 Then
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
         Else
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
         End If
      Else
         '§ä¤£¨ì
         mdiº~¦r¦r§Î.txt²Õ¦r¦r¼Æ.Text = ""
         mdiº~¦r¦r§Î.txt¦r§Î.Text = ""
         mdiº~¦r¦r§Î.txtÁ`µ§µe.Text = ""
         mdiº~¦r¦r§Î.txt³¡­º.Text = ""
         mdiº~¦r¦r§Î.txt¦©°£³¡­ºµ§µe.Text = ""
         mdiº~¦r¦r§Î.txtª`­µ.Text = ""
         mdiº~¦r¦r§Î.txt¤º½X.Text = ""
         mdiº~¦r¦r§Î.txt­Ü¾e½X.Text = ""
         'mdiº~¦r¦r§Î.txtºc¦r¦¡.Text = ""
         mdiº~¦r¦r§Î.txt¥U¼Æ.Text = ""
         
         'µ¹©wªÅ¥Õ­È
         If Len(ª¬ºA¦C1) > 10 Then
            ª¬ºA¦C = "0 ­Ó¦r®Ú" & ª¬ºA¦C1
         Else
            ª¬ºA¦C = "0 ­Ó¦r®Ú"
         End If
         mdiº~¦r¦r§Î.txtª¬ºA = ª¬ºA¦C
      End If
   End If
End If

End Sub

Private Function ¦r§Î¬d¸ß() As Long
Dim ²Õ¦r¦¡ As String, ¦r®Ú§Ç As String, ¹Bºâ²Å¸¹ As Integer, ºc¦r¦¡ As String, ²§Åéºc¦r¦¡ As String
Dim i As Integer, j As Integer, ²Õ¦r²Å¸¹ As Integer, ­«ÂÐ¦¸¼Æ As Integer, ¤è«K²Å¸¹ As Integer
Dim ­«ÂÐ¦r As String, ¦r®Ú²Õ As String, ¬Û¦ü¦r¦ê As String, ¼È¦s²Õ¦r¦¡ As String
Dim ­ºµ§ºc¦r¦¡ As String, ¤À¸Ñ²Å¸¹ As Boolean

On Error GoTo ¦r§Î¬d¸ß¿ù»~

¹Bºâ²Å¸¹ = 4
­«ÂÐ¦¸¼Æ = 0
²Õ¦r¦¡ = ""
ºc¦r¦¡ = ""
¬Û¦ü¦r¦ê = ""
­ºµ§ºc¦r¦¡ = ""
¤À¸Ñ²Å¸¹ = False

¦r§Î¬d¸ß = -1

If Len(Trim(txtºc¦r¦¡.Text)) = 1 Then
   ÀË¦rªí.Index = "¦r§Î"
   ÀË¦rªí.Seek "=", Trim(txtºc¦r¦¡.Text)
   If Not ÀË¦rªí.NoMatch Then ¦r§Î¬d¸ß = ÀË¦rªí.Fields("½s¸¹")
   Exit Function
End If

i = 1
Do While i <= Len(Trim(txtºc¦r¦¡.Text))
 
   ²Õ¦r²Å¸¹ = ¬O§_¬°²Õ¦r²Å¸¹(Mid(txtºc¦r¦¡.Text, i, 1), 1, 14)

   If ²Õ¦r²Å¸¹ <> 12 And ²Õ¦r²Å¸¹ <> 13 Then
      If ²Õ¦r²Å¸¹ >= 1 And ²Õ¦r²Å¸¹ <= 3 Then
         ¹Bºâ²Å¸¹ = ²Õ¦r²Å¸¹
      ElseIf Len(txtºc¦r¦¡) = 2 Then
         ¹Bºâ²Å¸¹ = 5
         ²Õ¦r¦¡ = txtºc¦r¦¡
         Exit Do
      End If
      If ²Õ¦r²Å¸¹ >= 4 Or ²Õ¦r²Å¸¹ = 0 Then
         ²Õ¦r¦¡ = ²Õ¦r¦¡ & Mid(txtºc¦r¦¡, i, 1)
         If Len(txtºc¦r¦¡) = 1 And ²Õ¦r²Å¸¹ = 0 Then ¹Bºâ²Å¸¹ = 0
      End If
   Else
      ¹Bºâ²Å¸¹ = 4
      ²Õ¦r¦¡ = ²Õ¦r¦¡ & Mid(txtºc¦r¦¡, 2, Len(txtºc¦r¦¡) - 2)
      Exit Do
   End If
   i = i + 1
Loop
    
If ²Õ¦r¦¡ = "" Then ²Õ¦r¦¡ = Trim(txtºc¦r¦¡.Text)

ºc¦r¦¡ = ²Õ¦r¦¡

'¥ý¥Î¤À¸Ñ+³¡¥ó§Ç·j´M

If Len(ºc¦r¦¡) > 1 Or (Len(ºc¦r¦¡) = 1 And ²Õ¦r²Å¸¹ = 0) Then
   ÀË¦rªí.Index = "ºc¦r¦¡"
   ÀË¦rªí.Seek "=", ¹Bºâ²Å¸¹, ºc¦r¦¡
   If Not ÀË¦rªí.NoMatch Then
      ¦r§Î¬d¸ß = ÀË¦rªí.Fields("½s¸¹")
      If ¹Bºâ²Å¸¹ > 0 Then
         If ÀË¦rªí.Fields("¦rÅé") <> 0 Or IsNull(ÀË¦rªí.Fields("¦r§Î")) Then
            ­ºµ§ºc¦r¦¡ = ´M§ä²Õ¦r¦¡(ÀË¦rªí.Fields("³s±µ²Å¸¹"), ÀË¦rªí.Fields("³¡¥ó§Ç"))
         Else
            ­ºµ§ºc¦r¦¡ = ÀË¦rªí.Fields("¦r§Î")
         End If
      End If
      ÀË¦rªí.MoveNext
      Do Until ÀË¦rªí.EOF Or ÀË¦rªí.Fields("³¡¥ó§Ç") <> ºc¦r¦¡
         If ÀË¦rªí.Fields("³s±µ²Å¸¹") = ¹Bºâ²Å¸¹ Then
            ¦r§Î¬d¸ß = ÀË¦rªí.Fields("½s¸¹")
            ¤À¸Ñ²Å¸¹ = True
         Else
            If ÀË¦rªí.Fields("¦rÅé") <> 0 Or IsNull(ÀË¦rªí.Fields("¦r§Î")) Then
               ¼È¦s²Õ¦r¦¡ = ´M§ä²Õ¦r¦¡(ÀË¦rªí.Fields("³s±µ²Å¸¹"), ÀË¦rªí.Fields("³¡¥ó§Ç"))
            Else
               ¼È¦s²Õ¦r¦¡ = ÀË¦rªí.Fields("¦r§Î")
            End If
            ¬Û¦ü¦r¦ê = ¬Û¦ü¦r¦ê & "[" & ¼È¦s²Õ¦r¦¡ & "]"
         End If
         ÀË¦rªí.MoveNext
      Loop
      If ¤À¸Ñ²Å¸¹ = True Then ¬Û¦ü¦r¦ê = ­ºµ§ºc¦r¦¡ & ¬Û¦ü¦r¦ê
      If Len(¬Û¦ü¦r¦ê) > 0 Then ª¬ºA¦C1 = ",¬Û¦ü¦r§Î: " & ¬Û¦ü¦r¦ê
      Exit Function
   End If
Else
   ¦r§Î¬d¸ß = 0
   txt¦r§Î.Text = ºc¦r¦¡
   Exit Function
End If
       
'³¡¥ó§Çnot found,§ï¥Î¦r®Ú§Ç
'¬O²Õ¦r²Å¸¹
i = 1
ºc¦r¦¡ = ""
ÀË¦rªí.Index = "¦r§Î"
Do While i <= Len(²Õ¦r¦¡)
   ¤è«K²Å¸¹ = ¬O§_¬°²Õ¦r²Å¸¹(Mid(²Õ¦r¦¡, i, 1), 4, 11)
   Select Case ¤è«K²Å¸¹
          Case 4, 5
               ­«ÂÐ¦¸¼Æ = 2
          Case 6, 7, 8
               ­«ÂÐ¦¸¼Æ = 3
          Case 9, 10, 11
               ­«ÂÐ¦¸¼Æ = 4
   End Select
   If ¤è«K²Å¸¹ > 0 Then
      i = i + 1
      ÀË¦rªí.Seek "=", Mid(²Õ¦r¦¡, i, 1)
      If Not ÀË¦rªí.NoMatch Then
         '´M§ä¸Ó­«ÂÐ¦rªº¦r§Î,­Y¤À¸Ñ¤£¬°0,«hÄ~Äò©¹¤U§ä¥X©Ò¦³¸Ó¦r§Î¤§¦r®Ú
         If ÀË¦rªí.Fields("³s±µ²Å¸¹") <> 0 Then
            ­«ÂÐ¦r = ¦r®Ú§Ç¬d¸ß(ÀË¦rªí.Fields("¦r®Ú§Ç"))
         Else
            ­«ÂÐ¦r = Mid(²Õ¦r¦¡, i, 1)
         End If
      End If
      For j = 1 To ­«ÂÐ¦¸¼Æ
          ºc¦r¦¡ = ºc¦r¦¡ & ­«ÂÐ¦r
      Next j
   '¤£¬O²Õ¦r²Å¸¹
   Else
       ÀË¦rªí.Index = "¦r§Î"
       ÀË¦rªí.Seek "=", Mid(²Õ¦r¦¡, i, 1)
       If Not ÀË¦rªí.NoMatch Then
          Do Until ÀË¦rªí.EOF Or ÀË¦rªí.Fields("¦r§Î") <> Mid(²Õ¦r¦¡, i, 1)
             If ÀË¦rªí.Fields("¦rÅé") = 0 And ÀË¦rªí.Fields("¦r§Î") = Mid(²Õ¦r¦¡, i, 1) Then
                If ÀË¦rªí.Fields("³s±µ²Å¸¹") <> 0 Then
                   ºc¦r¦¡ = ºc¦r¦¡ & ¦r®Ú§Ç¬d¸ß(ÀË¦rªí.Fields("¦r®Ú§Ç"))
                Else
                   ºc¦r¦¡ = ºc¦r¦¡ & ÀË¦rªí.Fields("¦r®Ú§Ç")
                End If
             End If
             ÀË¦rªí.MoveNext
          Loop
       End If
   End If
   i = i + 1
Loop
    
'¥ý¥Î¦r®Ú§Ç·j´M

ÀË¦rªí.Index = "¦r®Ú§Ç"
ÀË¦rªí.Seek "=", ºc¦r¦¡

If Not ÀË¦rªí.NoMatch Then
   ¦r§Î¬d¸ß = ÀË¦rªí.Fields("½s¸¹")
   If ¹Bºâ²Å¸¹ > 0 Then
      If ÀË¦rªí.Fields("¦rÅé") <> 0 Or IsNull(ÀË¦rªí.Fields("¦r§Î")) Then
         ­ºµ§ºc¦r¦¡ = ´M§ä²Õ¦r¦¡(ÀË¦rªí.Fields("³s±µ²Å¸¹"), ÀË¦rªí.Fields("³¡¥ó§Ç"))
      Else
         ­ºµ§ºc¦r¦¡ = ÀË¦rªí.Fields("¦r§Î")
      End If
   End If
   ÀË¦rªí.MoveNext
   Do Until ÀË¦rªí.EOF Or ÀË¦rªí.Fields("³¡¥ó§Ç") <> ºc¦r¦¡
      If ÀË¦rªí.Fields("³s±µ²Å¸¹") = ¹Bºâ²Å¸¹ Then
         ¦r§Î¬d¸ß = ÀË¦rªí.Fields("½s¸¹")
         ¤À¸Ñ²Å¸¹ = True
      Else
         If ÀË¦rªí.Fields("¦rÅé") <> 0 Or IsNull(ÀË¦rªí.Fields("¦r§Î")) Then
            ¼È¦s²Õ¦r¦¡ = ´M§ä²Õ¦r¦¡(ÀË¦rªí.Fields("³s±µ²Å¸¹"), ÀË¦rªí.Fields("³¡¥ó§Ç"))
         Else
            ¼È¦s²Õ¦r¦¡ = ÀË¦rªí.Fields("¦r§Î")
         End If
         ¬Û¦ü¦r¦ê = ¬Û¦ü¦r¦ê & "[" & ¼È¦s²Õ¦r¦¡ & "]"
      End If
      ÀË¦rªí.MoveNext
   Loop
   If ¤À¸Ñ²Å¸¹ = True Then ¬Û¦ü¦r¦ê = ­ºµ§ºc¦r¦¡ & ¬Û¦ü¦r¦ê
   If Len(¬Û¦ü¦r¦ê) > 0 Then ª¬ºA¦C1 = ",¬Û¦ü¦r§Î: " & ¬Û¦ü¦r¦ê
   Exit Function
End If
   
'­Y¥Î¦r®Ú§Ç·j´M¤£¨ì,«h§ï¥Î¦r®Ú§Ç¤G·j´M
²§Åéºc¦r¦¡ = ""
For i = 1 To Len(ºc¦r¦¡)
    ²§¼g¦r®Ú.Seek "=", Mid(ºc¦r¦¡, i, 1)
    If Not ²§¼g¦r®Ú.NoMatch Then
        ²§Åéºc¦r¦¡ = ²§Åéºc¦r¦¡ & ²§¼g¦r®Ú.Fields("²§¼g")
    Else
        ²§Åéºc¦r¦¡ = ²§Åéºc¦r¦¡ & Mid(ºc¦r¦¡, i, 1)
    End If
Next i

ÀË¦rªí.Index = "¦r®Ú§Ç¤G"
ÀË¦rªí.Seek "=", ²§Åéºc¦r¦¡

If Not ÀË¦rªí.NoMatch Then
   ¦r§Î¬d¸ß = ÀË¦rªí.Fields("½s¸¹")
   If ¹Bºâ²Å¸¹ > 0 Then
      If ÀË¦rªí.Fields("¦rÅé") <> 0 Or IsNull(ÀË¦rªí.Fields("¦r§Î")) Then
         ­ºµ§ºc¦r¦¡ = ´M§ä²Õ¦r¦¡(ÀË¦rªí.Fields("³s±µ²Å¸¹"), ÀË¦rªí.Fields("³¡¥ó§Ç"))
      Else
         ­ºµ§ºc¦r¦¡ = ÀË¦rªí.Fields("¦r§Î")
      End If
   End If
   ÀË¦rªí.MoveNext
   Do Until ÀË¦rªí.EOF Or ÀË¦rªí.Fields("³¡¥ó§Ç") <> ºc¦r¦¡
      If ÀË¦rªí.Fields("³s±µ²Å¸¹") = ¹Bºâ²Å¸¹ Then
         ¦r§Î¬d¸ß = ÀË¦rªí.Fields("½s¸¹")
         ¤À¸Ñ²Å¸¹ = True
      Else
         If ÀË¦rªí.Fields("¦rÅé") <> 0 Or IsNull(ÀË¦rªí.Fields("¦r§Î")) Then
            ¼È¦s²Õ¦r¦¡ = ´M§ä²Õ¦r¦¡(ÀË¦rªí.Fields("³s±µ²Å¸¹"), ÀË¦rªí.Fields("³¡¥ó§Ç"))
         Else
            ¼È¦s²Õ¦r¦¡ = ÀË¦rªí.Fields("¦r§Î")
         End If
         ¬Û¦ü¦r¦ê = ¬Û¦ü¦r¦ê & "[" & ¼È¦s²Õ¦r¦¡ & "]"
      End If
      ÀË¦rªí.MoveNext
   Loop
   If ¤À¸Ñ²Å¸¹ = True Then ¬Û¦ü¦r¦ê = ­ºµ§ºc¦r¦¡ & ¬Û¦ü¦r¦ê
   If Len(¬Û¦ü¦r¦ê) > 0 Then ª¬ºA¦C1 = ",¬Û¦ü¦r§Î: " & ¬Û¦ü¦r¦ê
   Exit Function
End If
   

'­Y¥Î¦r®Ú§Ç·j´M¤£¨ì,«h§ï¥Î¦r®Ú²Õ·j´M

¦r®Ú²Õ = ¦r®Ú±Æ§Ç(ºc¦r¦¡)
ÀË¦rªí.Index = "¦r®Ú²Õ"
ÀË¦rªí.Seek "=", ¦r®Ú²Õ
If Not ÀË¦rªí.NoMatch Then
   ¦r§Î¬d¸ß = ÀË¦rªí.Fields("½s¸¹")
   If ¹Bºâ²Å¸¹ > 0 Then
      If ÀË¦rªí.Fields("¦rÅé") <> 0 Or IsNull(ÀË¦rªí.Fields("¦r§Î")) Then
         ­ºµ§ºc¦r¦¡ = ´M§ä²Õ¦r¦¡(ÀË¦rªí.Fields("³s±µ²Å¸¹"), ÀË¦rªí.Fields("³¡¥ó§Ç"))
      Else
         ­ºµ§ºc¦r¦¡ = ÀË¦rªí.Fields("¦r§Î")
      End If
   End If
   ÀË¦rªí.MoveNext
   Do Until ÀË¦rªí.EOF Or ÀË¦rªí.Fields("³¡¥ó§Ç") <> ºc¦r¦¡
      If ÀË¦rªí.Fields("³s±µ²Å¸¹") = ¹Bºâ²Å¸¹ Then
         ¦r§Î¬d¸ß = ÀË¦rªí.Fields("½s¸¹")
         ¤À¸Ñ²Å¸¹ = True
      Else
         If ÀË¦rªí.Fields("¦rÅé") <> 0 Or IsNull(ÀË¦rªí.Fields("¦r§Î")) Then
            ¼È¦s²Õ¦r¦¡ = ´M§ä²Õ¦r¦¡(ÀË¦rªí.Fields("³s±µ²Å¸¹"), ÀË¦rªí.Fields("³¡¥ó§Ç"))
         Else
            ¼È¦s²Õ¦r¦¡ = ÀË¦rªí.Fields("¦r§Î")
         End If
         ¬Û¦ü¦r¦ê = ¬Û¦ü¦r¦ê & "[" & ¼È¦s²Õ¦r¦¡ & "]"
      End If
      ÀË¦rªí.MoveNext
   Loop
   If ¤À¸Ñ²Å¸¹ = True Then ¬Û¦ü¦r¦ê = ­ºµ§ºc¦r¦¡ & ¬Û¦ü¦r¦ê
   If Len(¬Û¦ü¦r¦ê) > 0 Then ª¬ºA¦C1 = ",¬Û¦ü¦r§Î: " & ¬Û¦ü¦r¦ê
   Exit Function
End If

'­Y¥Î¦r®Ú§Ç·j´M¤£¨ì,«h§ï¥Î¦r®Ú²Õ¤G·j´M

²§Åéºc¦r¦¡ = ""
For i = 1 To Len(¦r®Ú²Õ)
    ²§¼g¦r®Ú.Seek "=", Mid(¦r®Ú²Õ, i, 1)
    If Not ²§¼g¦r®Ú.NoMatch Then
        ²§Åéºc¦r¦¡ = ²§Åéºc¦r¦¡ & ²§¼g¦r®Ú.Fields("²§¼g")
    Else
        ²§Åéºc¦r¦¡ = ²§Åéºc¦r¦¡ & Mid(¦r®Ú²Õ, i, 1)
    End If
Next i

²§Åéºc¦r¦¡ = ¦r®Ú±Æ§Ç(²§Åéºc¦r¦¡)
ÀË¦rªí.Index = "¦r®Ú²Õ¤G"
ÀË¦rªí.Seek "=", ²§Åéºc¦r¦¡
If Not ÀË¦rªí.NoMatch Then
   ¦r§Î¬d¸ß = ÀË¦rªí.Fields("½s¸¹")
   If ¹Bºâ²Å¸¹ > 0 Then
      If ÀË¦rªí.Fields("¦rÅé") <> 0 Or IsNull(ÀË¦rªí.Fields("¦r§Î")) Then
         ­ºµ§ºc¦r¦¡ = ´M§ä²Õ¦r¦¡(ÀË¦rªí.Fields("³s±µ²Å¸¹"), ÀË¦rªí.Fields("³¡¥ó§Ç"))
      Else
         ­ºµ§ºc¦r¦¡ = ÀË¦rªí.Fields("¦r§Î")
      End If
   End If
   ÀË¦rªí.MoveNext
   Do Until ÀË¦rªí.EOF Or ÀË¦rªí.Fields("³¡¥ó§Ç") <> ºc¦r¦¡
      If ÀË¦rªí.Fields("³s±µ²Å¸¹") = ¹Bºâ²Å¸¹ Then
         ¦r§Î¬d¸ß = ÀË¦rªí.Fields("½s¸¹")
         ¤À¸Ñ²Å¸¹ = True
      Else
         If ÀË¦rªí.Fields("¦rÅé") <> 0 Or IsNull(ÀË¦rªí.Fields("¦r§Î")) Then
            ¼È¦s²Õ¦r¦¡ = ´M§ä²Õ¦r¦¡(ÀË¦rªí.Fields("³s±µ²Å¸¹"), ÀË¦rªí.Fields("³¡¥ó§Ç"))
         Else
            ¼È¦s²Õ¦r¦¡ = ÀË¦rªí.Fields("¦r§Î")
         End If
         ¬Û¦ü¦r¦ê = ¬Û¦ü¦r¦ê & "[" & ¼È¦s²Õ¦r¦¡ & "]"
      End If
      ÀË¦rªí.MoveNext
   Loop
   If ¤À¸Ñ²Å¸¹ = True Then ¬Û¦ü¦r¦ê = ­ºµ§ºc¦r¦¡ & ¬Û¦ü¦r¦ê
   If Len(¬Û¦ü¦r¦ê) > 0 Then ª¬ºA¦C1 = ",¬Û¦ü¦r§Î: " & ¬Û¦ü¦r¦ê
   Exit Function
End If

¦r§Î¬d¸ß¿ù»~:

End Function

Private Function ¦r®Ú§Ç¬d¸ß(¦r®Ú§Ç As String) As String
Dim i As Integer, varBookMark As Variant
Dim ²Õ¦r¦¡ As String

varBookMark = ÀË¦rªí.Bookmark
ÀË¦rªí.Index = "¦r§Î"
²Õ¦r¦¡ = ""
For i = 1 To Len(¦r®Ú§Ç)
    ÀË¦rªí.Seek "=", Mid(¦r®Ú§Ç, i, 1)
    If Not ÀË¦rªí.NoMatch Then
       If ÀË¦rªí.Fields("³s±µ²Å¸¹") <> 0 And ÀË¦rªí.Fields("³s±µ²Å¸¹") <> 9 Then
          ²Õ¦r¦¡ = ²Õ¦r¦¡ & ¦r®Ú§Ç¬d¸ß(ÀË¦rªí.Fields("¦r®Ú§Ç"))
       Else
          ²Õ¦r¦¡ = ²Õ¦r¦¡ & ÀË¦rªí.Fields("¦r®Ú§Ç")
       End If
    End If
Next i
¦r®Ú§Ç¬d¸ß = ²Õ¦r¦¡
ÀË¦rªí.Bookmark = varBookMark

End Function


Private Sub txt³¡­º_GotFocus()

txt³¡­º.SelStart = 0
txt³¡­º.SelLength = Len(txt³¡­º.Text)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_³¡­º¤è¶ô

End Sub

Private Sub txt²Õ¦r¦r¼Æ_GotFocus()

txt²Õ¦r¦r¼Æ.SelStart = 0
txt²Õ¦r¦r¼Æ.SelLength = Len(txt²Õ¦r¦r¼Æ.Text)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_²Õ¦r¦r¼Æ¤è¶ô

End Sub

Private Sub txt½s¸¹_GotFocus()

txt½s¸¹.SelStart = 0
txt½s¸¹.SelLength = Len(txt½s¸¹.Text)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_½s¸¹¤è¶ô

End Sub

Private Sub txt½s¸¹_KeyPress(KeyAscii As Integer)

Dim ·¢®Ñ½s¸¹ As Long, ¤p½f½s¸¹ As Long, ª÷¤å½s¸¹ As Long, ¥Ò°©¤å½s¸¹ As Long, ·¡¨t¤å¦r½s¸¹ As Long
Dim ¼È¦s²Õ¦r¦¡ As String

If KeyAscii = vbKeyReturn Then
   If IsNumeric(txt½s¸¹.Text) Then
      ·¢®Ñ½s¸¹ = CLng(txt½s¸¹.Text)
      If ·¢®Ñ½s¸¹ > 0 Then
        Set ÀË¦rªí = ·¢®ÑÀË¦rªí
        ÀË¦rªí.Index = "½s¸¹"
        ÀË¦rªí.Seek "=", ·¢®Ñ½s¸¹
        If ÀË¦rªí.NoMatch Then GoTo §ä¤£¨ì
        If Not IsNull(ÀË¦rªí.Fields("¤p½f½s¸¹")) Then
            ¤p½f½s¸¹ = ÀË¦rªí.Fields("¤p½f½s¸¹")
        Else
            ¤p½f½s¸¹ = -1
        End If
        If Not IsNull(ÀË¦rªí.Fields("ª÷¤å½s¸¹")) Then
            ª÷¤å½s¸¹ = ÀË¦rªí.Fields("ª÷¤å½s¸¹")
        Else
            ª÷¤å½s¸¹ = -1
        End If
        If Not IsNull(ÀË¦rªí.Fields("¥Ò°©¤å½s¸¹")) Then
            ¥Ò°©¤å½s¸¹ = ÀË¦rªí.Fields("¥Ò°©¤å½s¸¹")
        Else
            ¥Ò°©¤å½s¸¹ = -1
        End If
        If Not IsNull(ÀË¦rªí.Fields("·¡¨t¤å¦r½s¸¹")) Then
            ·¡¨t¤å¦r½s¸¹ = ÀË¦rªí.Fields("·¡¨t¤å¦r½s¸¹")
        Else
            ·¡¨t¤å¦r½s¸¹ = -1
        End If

      End If
      
      If ·¢®Ñ½s¸¹ > 0 Then
         Â^¨úÄÝ©Ê "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
         Â^¨úºc¦r¦¡ "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
         If ¨t²Î¦rÅé = "¤p½f" And ¤p½f½s¸¹ > 0 Then
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¥_®v¤j»¡¤å¤p½f", txt¦r§Î.Text, ¤p½f½s¸¹
         ElseIf ¨t²Î¦rÅé = "ª÷¤å" And ª÷¤å½s¸¹ > 0 Then
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¤¤¬ã°|ª÷¤å", txt¦r§Î.Text, ª÷¤å½s¸¹
         ElseIf ¨t²Î¦rÅé = "¥Ò°©¤å" And ¥Ò°©¤å½s¸¹ > 0 Then
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¤¤¬ã°|¥Ò°©¤å", txt¦r§Î.Text, ¥Ò°©¤å½s¸¹
         ElseIf ¨t²Î¦rÅé = "·¡¨t¤å¦r" And ·¡¨t¤å¦r½s¸¹ > 0 Then
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¤¤¬ã°|·¡¨tÂ²©­¤å¦r", txt¦r§Î.Text, ·¡¨t¤å¦r½s¸¹
         Else
            If ±Ò°Ê¦r§Îµ²ºc Then frm¦r§Îµ²ºc.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
            If ±Ò°Ê²§Åé¦rªí Then frm²§Åé¦rªí.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
            If ±Ò°Ê¦r§ÎºtÅÜ Then frm¦r§ÎºtÅÜ.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
            If ±Ò°Ê¦r§Î¯Á¤Þ Then frm¦r§Î¯Á¤Þ.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
            If ±Ò°Ê²§Åé¦r®Ú Then frm²§Åé¦r®Ú.¸ü¤J¦r§Î "¼Ð·¢Åé", txt¦r§Î.Text, ·¢®Ñ½s¸¹
         End If
      Else
§ä¤£¨ì:
         mdiº~¦r¦r§Î.txt²Õ¦r¦r¼Æ.Text = ""
         mdiº~¦r¦r§Î.txt¦r§Î.Text = ""
         mdiº~¦r¦r§Î.txtÁ`µ§µe.Text = ""
         mdiº~¦r¦r§Î.txt³¡­º.Text = ""
         mdiº~¦r¦r§Î.txt¦©°£³¡­ºµ§µe.Text = ""
         mdiº~¦r¦r§Î.txtª`­µ.Text = ""
         mdiº~¦r¦r§Î.txt¤º½X.Text = ""
         mdiº~¦r¦r§Î.txt­Ü¾e½X.Text = ""
         'mdiº~¦r¦r§Î.txtºc¦r¦¡.Text = ""
         mdiº~¦r¦r§Î.txt¥U¼Æ.Text = ""
         
         'µ¹©wªÅ¥Õ­È
         If Len(ª¬ºA¦C1) > 10 Then
            ª¬ºA¦C = "0 ­Ó¦r®Ú" & ª¬ºA¦C1
         Else
            ª¬ºA¦C = "0 ­Ó¦r®Ú"
         End If
         mdiº~¦r¦r§Î.txtª¬ºA = ª¬ºA¦C
      End If
   End If
End If

End Sub

Private Sub txtÁ`µ§µe_GotFocus()

txtÁ`µ§µe.SelStart = 0
txtÁ`µ§µe.SelLength = Len(txtÁ`µ§µe.Text)
²{¥Îµøµ¡¥N½X = mdiº~¦r¦r§Î¥N½X
²{¥Î±±¨î¶µ¥N½X = mdiº~¦r¦r§Î_Á`µ§µe¤è¶ô

End Sub

Public Sub Â²©öÂsÄý¤@()
Dim i As Integer
Dim §¡¼e As Integer
Dim §¡°ª As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case ¦r§Îµ²ºc¥N½X, ²§Åé¦r®Ú¥N½X, ²§Åé¦rªí¥N½X, ¦r§ÎºtÅÜ¥N½X, ¦r§Î¯Á¤Þ¥N½X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

If ¨t²Î¦rÅé = "¤p½f" Or ¨t²Î¦rÅé = "ª÷¤å" Or ¨t²Î¦rÅé = "¥Ò°©¤å" Or ¨t²Î¦rÅé = "·¡¨t¤å¦r" Then
    mnu_»¡¤å¸Ñ¦r³¡­º_Click
Else
    mnu_±dº³¦r¨å³¡­º_Click
End If

frm¦r§Îµ²ºc.Tag = ¦r§Îµ²ºc¥N½X
frm¦r§Îµ²ºc.show
frm¦r§Îµ²ºc.SetFocus

'frm¦r§Î¯Á¤Þ.Tag = ¦r§Î¯Á¤Þ¥N½X
'frm¦r§Î¯Á¤Þ.show
'frm¦r§Î¯Á¤Þ.SetFocus

'frm²§Åé¦r®Ú.Tag = 15
'frm²§Åé¦r®Ú.show
'frm²§Åé¦r®Ú.Visible = False

'frm¦r§ÎºtÅÜ.Tag = 16
'frm¦r§ÎºtÅÜ.show
'frm¦r§ÎºtÅÜ.Visible = False

frm¦r§Î´F¨Å.Tag = ¦r§Î´F¨Å¥N½X
frm¦r§Î´F¨Å.show
frm¦r§Î´F¨Å.SetFocus

mdiº~¦r¦r§Î.Arrange 2

'frm¦r§Îµ²ºc.Height = frm¦r§Î´F¨Å.Height / 2
'frm³¡¥ó½d¨Ò.Height = frm¦r§Îµ²ºc.Height
'frm¦r§ÎºtÅÜ.Height = frm¦r§ÎºtÅÜ.Height

'frm¦r§ÎºtÅÜ.Left = frm¦r§Îµ²ºc.Left
'frm¦r§ÎºtÅÜ.Width = frm¦r§Îµ²ºc.Width
'frm¦r§ÎºtÅÜ.Height = frm¦r§Îµ²ºc.Height - 1
'frm¦r§ÎºtÅÜ.Top = frm¦r§Îµ²ºc.Height

'frm²§Åé¦r®Ú.Left = frm³¡¥ó½d¨Ò.Left
'frm²§Åé¦r®Ú.Width = frm³¡¥ó½d¨Ò.Width
'frm²§Åé¦r®Ú.Height = frm³¡¥ó½d¨Ò.Height - 1
'frm²§Åé¦r®Ú.Top = frm³¡¥ó½d¨Ò.Height

'frm¦r§ÎºtÅÜ.Visible = True
'frm²§Åé¦r®Ú.Visible = True
frm¦r§Î´F¨Å.SetFocus

End Sub

Public Sub Â²©öÂsÄý¤G()

Dim i As Integer
Dim §¡¼e As Integer
Dim §¡°ª As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case Big5¦r®Ú¥N½X To ºc¦r²Å¸¹¥N½X, ²§Åé¦r®Ú¥N½X, ¦r§ÎºtÅÜ¥N½X, ¦r§Î¯Á¤Þ¥N½X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

'mnu_³¡¥ó¥N½X_Click 3
frm¦r§Îµ²ºc.Tag = ¦r§Îµ²ºc¥N½X
frm¦r§Îµ²ºc.show
frm¦r§Îµ²ºc.SetFocus

frm²§Åé¦rªí.Tag = ²§Åé¦rªí¥N½X
frm²§Åé¦rªí.show
frm²§Åé¦rªí.SetFocus

'frm²§Åé¦r®Ú.Tag = 15
'frm²§Åé¦r®Ú.show
'frm²§Åé¦r®Ú.Visible = False

'frm¦r§ÎºtÅÜ.Tag = 16
'frm¦r§ÎºtÅÜ.show
'frm¦r§ÎºtÅÜ.Visible = False

frm¦r§Î´F¨Å.Tag = ¦r§Î´F¨Å¥N½X
frm¦r§Î´F¨Å.show
frm¦r§Î´F¨Å.SetFocus

mdiº~¦r¦r§Î.Arrange 2

'frm¦r§Îµ²ºc.Height = frm¦r§Î´F¨Å.Height / 2
'frm³¡¥ó½d¨Ò.Height = frm¦r§Îµ²ºc.Height
'frm¦r§ÎºtÅÜ.Height = frm¦r§ÎºtÅÜ.Height

'frm¦r§ÎºtÅÜ.Left = frm¦r§Îµ²ºc.Left
'frm¦r§ÎºtÅÜ.Width = frm¦r§Îµ²ºc.Width
'frm¦r§ÎºtÅÜ.Height = frm¦r§Îµ²ºc.Height - 1
'frm¦r§ÎºtÅÜ.Top = frm¦r§Îµ²ºc.Height

'frm²§Åé¦r®Ú.Left = frm³¡¥ó½d¨Ò.Left
'frm²§Åé¦r®Ú.Width = frm³¡¥ó½d¨Ò.Width
'frm²§Åé¦r®Ú.Height = frm³¡¥ó½d¨Ò.Height - 1
'frm²§Åé¦r®Ú.Top = frm³¡¥ó½d¨Ò.Height

'frm¦r§ÎºtÅÜ.Visible = True
'frm²§Åé¦r®Ú.Visible = True
frm¦r§Î´F¨Å.SetFocus

End Sub

Public Sub Â²©öÂsÄý¤T()

Dim i As Integer
Dim §¡¼e As Integer
Dim §¡°ª As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case Big5¦r®Ú¥N½X To ºc¦r²Å¸¹¥N½X, ²§Åé¦r®Ú¥N½X, ¦r§ÎºtÅÜ¥N½X, ¦r§Îµ²ºc¥N½X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

frm¦r§Î¯Á¤Þ.Tag = ¦r§Î¯Á¤Þ¥N½X
frm¦r§Î¯Á¤Þ.show
frm¦r§Î¯Á¤Þ.SetFocus

frm²§Åé¦rªí.Tag = ²§Åé¦rªí¥N½X
frm²§Åé¦rªí.show
frm²§Åé¦rªí.SetFocus

frm¦r§Î´F¨Å.Tag = ¦r§Î´F¨Å¥N½X
frm¦r§Î´F¨Å.show
frm¦r§Î´F¨Å.SetFocus

mdiº~¦r¦r§Î.Arrange 2

frm¦r§Î´F¨Å.SetFocus

End Sub

Public Sub Â²©öÂsÄý¥|()

Dim i As Integer
Dim §¡¼e As Integer
Dim §¡°ª As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case Big5¦r®Ú¥N½X To ºc¦r²Å¸¹¥N½X, ²§Åé¦r®Ú¥N½X, ¦r§ÎºtÅÜ¥N½X, ²§Åé¦rªí¥N½X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

frm¦r§Î¯Á¤Þ.Tag = ¦r§Î¯Á¤Þ¥N½X
frm¦r§Î¯Á¤Þ.show
frm¦r§Î¯Á¤Þ.SetFocus

frm¦r§Îµ²ºc.Tag = ¦r§Îµ²ºc¥N½X
frm¦r§Îµ²ºc.show
frm¦r§Îµ²ºc.SetFocus

frm¦r§Î´F¨Å.Tag = ¦r§Î´F¨Å¥N½X
frm¦r§Î´F¨Å.show
frm¦r§Î´F¨Å.SetFocus

mdiº~¦r¦r§Î.Arrange 2

frm¦r§Î´F¨Å.SetFocus

End Sub
Public Sub ¹w³]ÂsÄý¤@()

Dim i As Integer
Dim §¡¼e As Integer
Dim §¡°ª As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case Big5¦r®Ú¥N½X To ºc¦r²Å¸¹¥N½X, ²§Åé¦r®Ú¥N½X, ¦r§Î¯Á¤Þ¥N½X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

'mnu_³¡¥ó¥N½X_Click 3
frm¦r§Îµ²ºc.Tag = ¦r§Îµ²ºc¥N½X
frm¦r§Îµ²ºc.show
frm¦r§Îµ²ºc.SetFocus

frm²§Åé¦rªí.Tag = ²§Åé¦rªí¥N½X
frm²§Åé¦rªí.show
frm²§Åé¦rªí.SetFocus

'frm²§Åé¦r®Ú.Tag = 15
'frm²§Åé¦r®Ú.show
'frm²§Åé¦r®Ú.Visible = False

frm¦r§ÎºtÅÜ.Tag = ¦r§ÎºtÅÜ¥N½X
frm¦r§ÎºtÅÜ.show
frm¦r§ÎºtÅÜ.Visible = False

frm¦r§Î´F¨Å.Tag = ¦r§Î´F¨Å¥N½X
frm¦r§Î´F¨Å.show
frm¦r§Î´F¨Å.SetFocus

mdiº~¦r¦r§Î.Arrange 2

frm¦r§Îµ²ºc.Height = frm¦r§Î´F¨Å.Height / 2
'frm³¡¥ó½d¨Ò.Height = frm¦r§Îµ²ºc.Height
'frm¦r§ÎºtÅÜ.Height = frm¦r§ÎºtÅÜ.Height

frm¦r§ÎºtÅÜ.Left = frm¦r§Îµ²ºc.Left
frm¦r§ÎºtÅÜ.Width = frm¦r§Îµ²ºc.Width
frm¦r§ÎºtÅÜ.Height = frm¦r§Îµ²ºc.Height - 1
frm¦r§ÎºtÅÜ.Top = frm¦r§Îµ²ºc.Height

'frm²§Åé¦r®Ú.Left = frm³¡¥ó½d¨Ò.Left
'frm²§Åé¦r®Ú.Width = frm³¡¥ó½d¨Ò.Width
'frm²§Åé¦r®Ú.Height = frm³¡¥ó½d¨Ò.Height - 1
'frm²§Åé¦r®Ú.Top = frm³¡¥ó½d¨Ò.Height

frm¦r§ÎºtÅÜ.Visible = True
'frm²§Åé¦r®Ú.Visible = True
frm¦r§Î´F¨Å.SetFocus

End Sub

Public Sub ¹w³]ÂsÄý¤G()

Dim i As Integer
Dim §¡¼e As Integer
Dim §¡°ª As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case Big5¦r®Ú¥N½X To ºc¦r²Å¸¹¥N½X, ²§Åé¦r®Ú¥N½X, ¦r§Î¯Á¤Þ¥N½X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

frm²§Åé¦rªí.Tag = ²§Åé¦rªí¥N½X
frm²§Åé¦rªí.show
frm²§Åé¦rªí.SetFocus

frm¦r§Îµ²ºc.Tag = ¦r§Îµ²ºc¥N½X
frm¦r§Îµ²ºc.show
frm¦r§Îµ²ºc.SetFocus

frm¦r§ÎºtÅÜ.Tag = ¦r§ÎºtÅÜ¥N½X
frm¦r§ÎºtÅÜ.show
frm¦r§ÎºtÅÜ.Visible = False

frm¦r§Î´F¨Å.Tag = ¦r§Î´F¨Å¥N½X
frm¦r§Î´F¨Å.show
frm¦r§Î´F¨Å.SetFocus

mdiº~¦r¦r§Î.Arrange 2

frm²§Åé¦rªí.Height = frm¦r§Î´F¨Å.Height / 2

frm¦r§ÎºtÅÜ.Left = frm²§Åé¦rªí.Left
frm¦r§ÎºtÅÜ.Width = frm²§Åé¦rªí.Width
frm¦r§ÎºtÅÜ.Height = frm²§Åé¦rªí.Height - 1
frm¦r§ÎºtÅÜ.Top = frm²§Åé¦rªí.Height

frm¦r§ÎºtÅÜ.Visible = True
frm¦r§Î´F¨Å.SetFocus

End Sub
Public Sub ¹w³]ÂsÄý¤T()

Dim i As Integer
Dim §¡¼e As Integer
Dim §¡°ª As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case Big5¦r®Ú¥N½X To ºc¦r²Å¸¹¥N½X, ²§Åé¦r®Ú¥N½X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

frm¦r§Î¯Á¤Þ.Tag = ¦r§Î¯Á¤Þ¥N½X
frm¦r§Î¯Á¤Þ.show
frm¦r§Î¯Á¤Þ.SetFocus

frm¦r§Îµ²ºc.Tag = ¦r§Îµ²ºc¥N½X
frm¦r§Îµ²ºc.show
frm¦r§Îµ²ºc.SetFocus

frm²§Åé¦rªí.Tag = ²§Åé¦rªí¥N½X
frm²§Åé¦rªí.show
frm²§Åé¦rªí.Visible = False

frm¦r§ÎºtÅÜ.Tag = ¦r§ÎºtÅÜ¥N½X
frm¦r§ÎºtÅÜ.show
frm¦r§ÎºtÅÜ.Visible = False

frm¦r§Î´F¨Å.Tag = ¦r§Î´F¨Å¥N½X
frm¦r§Î´F¨Å.show
frm¦r§Î´F¨Å.SetFocus

mdiº~¦r¦r§Î.Arrange 2

frm¦r§Îµ²ºc.Height = frm¦r§Î´F¨Å.Height / 2
frm¦r§Î¯Á¤Þ.Height = frm¦r§Îµ²ºc.Height

frm²§Åé¦rªí.Left = frm¦r§Îµ²ºc.Left
frm²§Åé¦rªí.Width = frm¦r§Îµ²ºc.Width
frm²§Åé¦rªí.Height = frm¦r§Îµ²ºc.Height - 1
frm²§Åé¦rªí.Top = frm¦r§Îµ²ºc.Height

frm¦r§ÎºtÅÜ.Left = frm¦r§Î¯Á¤Þ.Left
frm¦r§ÎºtÅÜ.Width = frm¦r§Î¯Á¤Þ.Width
frm¦r§ÎºtÅÜ.Height = frm¦r§Î¯Á¤Þ.Height - 1
frm¦r§ÎºtÅÜ.Top = frm¦r§Î¯Á¤Þ.Height

frm²§Åé¦rªí.Visible = True
frm¦r§ÎºtÅÜ.Visible = True
frm¦r§Î´F¨Å.SetFocus

End Sub

Public Sub ¹w³]ÂsÄý¥|()

Dim i As Integer
Dim §¡¼e As Integer
Dim §¡°ª As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case ¦r§ÎºtÅÜ¥N½X, ¦r§Î¯Á¤Þ¥N½X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

If ¨t²Î¦rÅé = "¤p½f" Then
    mnu_»¡¤å¸Ñ¦r³¡­º_Click
Else
    mnu_±dº³¦r¨å³¡­º_Click
End If

frm¦r§Îµ²ºc.Tag = ¦r§Îµ²ºc¥N½X
frm¦r§Îµ²ºc.show
frm¦r§Îµ²ºc.SetFocus

frm²§Åé¦rªí.Tag = ²§Åé¦rªí¥N½X
frm²§Åé¦rªí.show
frm²§Åé¦rªí.Visible = False

frm²§Åé¦r®Ú.Tag = ²§Åé¦r®Ú¥N½X
frm²§Åé¦r®Ú.show
frm²§Åé¦r®Ú.Visible = False

frm¦r§Î´F¨Å.Tag = ¦r§Î´F¨Å¥N½X
frm¦r§Î´F¨Å.show
frm¦r§Î´F¨Å.SetFocus

mdiº~¦r¦r§Î.Arrange 2

frm¦r§Îµ²ºc.Height = frm¦r§Î´F¨Å.Height / 2
frm³¡¥ó½d¨Ò.Height = frm¦r§Îµ²ºc.Height

frm²§Åé¦rªí.Left = frm¦r§Îµ²ºc.Left
frm²§Åé¦rªí.Width = frm¦r§Îµ²ºc.Width
frm²§Åé¦rªí.Height = frm¦r§Îµ²ºc.Height - 1
frm²§Åé¦rªí.Top = frm¦r§Îµ²ºc.Height

frm²§Åé¦r®Ú.Left = frm³¡¥ó½d¨Ò.Left
frm²§Åé¦r®Ú.Width = frm³¡¥ó½d¨Ò.Width
frm²§Åé¦r®Ú.Height = frm³¡¥ó½d¨Ò.Height - 1
frm²§Åé¦r®Ú.Top = frm³¡¥ó½d¨Ò.Height

frm²§Åé¦rªí.Visible = True
frm²§Åé¦r®Ú.Visible = True
frm¦r§Î´F¨Å.SetFocus

End Sub

