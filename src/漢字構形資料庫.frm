VERSION 5.00
Begin VB.MDIForm mdi�~�r�r�� 
   BackColor       =   &H8000000C&
   Caption         =   "�~�r�c�θ�Ʈw(������s�|��T��Ǭ�s��)"
   ClientHeight    =   6936
   ClientLeft      =   168
   ClientTop       =   768
   ClientWidth     =   13584
   Icon            =   "�~�r�c�θ�Ʈw.frx":0000
   LinkMode        =   1  '�ӷ�
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox pic���A�C 
      Align           =   2  '������U��
      BorderStyle     =   0  '�S���ؽu
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Begin VB.TextBox txt���A 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "�s�ө���"
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
   Begin VB.PictureBox pic�c�r�Ÿ� 
      Align           =   1  '������W��
      Appearance      =   0  '����
      BackColor       =   &H80000004&
      ForeColor       =   &H80000004&
      Height          =   492
      Left            =   0
      ScaleHeight     =   468
      ScaleWidth      =   13560
      TabIndex        =   12
      Top             =   0
      Width           =   13584
      Begin VB.ComboBox cbo�Ϥ��j�p 
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
         ItemData        =   "�~�r�c�θ�Ʈw.frx":030A
         Left            =   8016
         List            =   "�~�r�c�θ�Ʈw.frx":030C
         TabIndex        =   27
         ToolTipText     =   "�Ϥ��j�p(�r���I��)"
         Top             =   44
         Width           =   855
      End
      Begin VB.ComboBox cbo�ѪR�� 
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
         ItemData        =   "�~�r�c�θ�Ʈw.frx":030E
         Left            =   6924
         List            =   "�~�r�c�θ�Ʈw.frx":0318
         TabIndex        =   26
         ToolTipText     =   "�Ϥ��ѪR��(dpi)"
         Top             =   44
         Width           =   1020
      End
      Begin VB.ComboBox cbo�r���W�� 
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         ItemData        =   "�~�r�c�θ�Ʈw.frx":032C
         Left            =   4548
         List            =   "�~�r�c�θ�Ʈw.frx":0336
         TabIndex        =   20
         ToolTipText     =   "�r��"
         Top             =   44
         Width           =   1368
      End
      Begin VB.ComboBox cbo�r��j�p 
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
         ItemData        =   "�~�r�c�θ�Ʈw.frx":034A
         Left            =   6000
         List            =   "�~�r�c�θ�Ʈw.frx":037E
         TabIndex        =   17
         ToolTipText     =   "��ܦr���j�p"
         Top             =   44
         Width           =   855
      End
      Begin VB.ComboBox cbo���� 
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         ItemData        =   "�~�r�c�θ�Ʈw.frx":03C0
         Left            =   3456
         List            =   "�~�r�c�θ�Ʈw.frx":03D6
         TabIndex        =   16
         ToolTipText     =   "����"
         Top             =   44
         Width           =   972
      End
      Begin VB.ComboBox cbo���e 
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
         ItemData        =   "�~�r�c�θ�Ʈw.frx":03F4
         Left            =   2376
         List            =   "�~�r�c�θ�Ʈw.frx":03F6
         TabIndex        =   15
         ToolTipText     =   "���e"
         Top             =   44
         Width           =   972
      End
      Begin VB.ComboBox cbo�Ÿ� 
         BeginProperty Font 
            Name            =   "�з���"
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
         ToolTipText     =   "�c�r�Ÿ�"
         Top             =   44
         Width           =   735
      End
      Begin VB.ComboBox cbo�Ÿ����� 
         BeginProperty Font 
            Name            =   "�з���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         ItemData        =   "�~�r�c�θ�Ʈw.frx":03F8
         Left            =   108
         List            =   "�~�r�c�θ�Ʈw.frx":03FA
         TabIndex        =   13
         ToolTipText     =   "�c�r�Ÿ�����"
         Top             =   44
         Width           =   1335
      End
   End
   Begin VB.PictureBox pic�r���ݩ� 
      Align           =   1  '������W��
      Appearance      =   0  '����
      BackColor       =   &H80000004&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H80000002&
      Height          =   492
      Left            =   0
      ScaleHeight     =   492
      ScaleWidth      =   13584
      TabIndex        =   0
      Top             =   492
      Width           =   13584
      Begin VB.TextBox txt���� 
         Alignment       =   2  '�m�����
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
         ToolTipText     =   "�j�~�r�r��"
         Top             =   48
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  '�m�����
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
         ToolTipText     =   "���e"
         Top             =   48
         Width           =   375
      End
      Begin VB.TextBox txt�j�~�r 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�з���"
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
         ToolTipText     =   "�j�~�r"
         Top             =   48
         Width           =   375
      End
      Begin VB.TextBox txt�~�r�� 
         Alignment       =   2  '�m�����
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
         ToolTipText     =   "�~�r��"
         Top             =   44
         Width           =   375
      End
      Begin VB.TextBox txt�s�� 
         Alignment       =   2  '�m�����
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
         ToolTipText     =   "�s��"
         Top             =   44
         Width           =   732
      End
      Begin VB.TextBox txt�r�� 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�з���"
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
         ToolTipText     =   "��J�r�Ϋ�A�A��Enter�d�ߦr���ݩ�"
         Top             =   44
         Width           =   375
      End
      Begin VB.TextBox txt�����������e 
         Alignment       =   2  '�m�����
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
         ToolTipText     =   "���e(�������p)"
         Top             =   48
         Width           =   375
      End
      Begin VB.TextBox txt�`���e 
         Alignment       =   2  '�m�����
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
         ToolTipText     =   "���e"
         Top             =   48
         Width           =   375
      End
      Begin VB.TextBox txt���� 
         Alignment       =   2  '�m�����
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "�з���"
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
         ToolTipText     =   "����"
         Top             =   48
         Width           =   375
      End
      Begin VB.TextBox txt�զr�r�Ƨt���g 
         Alignment       =   2  '�m�����
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
         ToolTipText     =   "�զr�r��(�]�t����r��)"
         Top             =   48
         Width           =   612
      End
      Begin VB.TextBox txt�զr�r�� 
         Alignment       =   2  '�m�����
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
         ToolTipText     =   "�զr�r��"
         Top             =   48
         Width           =   612
      End
      Begin VB.TextBox txt�U�� 
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
         ToolTipText     =   "�~�y�j�r��U���r"
         Top             =   48
         Width           =   1212
      End
      Begin VB.TextBox txt�`�� 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "�з���"
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
         ToolTipText     =   "�`��"
         Top             =   48
         Width           =   1095
      End
      Begin VB.TextBox txt���X 
         Alignment       =   2  '�m�����
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
      Begin VB.TextBox txt�ܾe�X 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "�з���"
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
         ToolTipText     =   "�ܾe�X"
         Top             =   48
         Width           =   1572
      End
      Begin VB.TextBox txt�c�r�� 
         BeginProperty Font 
            Name            =   "�з���"
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
         Text            =   $"�~�r�c�θ�Ʈw.frx":03FC
         ToolTipText     =   "��J�c�r����A�A��Enter�d�ߦr���ݩ�"
         Top             =   48
         Width           =   2580
      End
   End
   Begin VB.Menu mnu_�r�� 
      Caption         =   "�r��"
      Begin VB.Menu mnu_�`�Φr 
         Caption         =   "�`�Φr"
      End
      Begin VB.Menu mnu_Big5 
         Caption         =   "���j�X(Big5)"
      End
      Begin VB.Menu mnu_²�Ʀr�`�� 
         Caption         =   "²�Ʀr�`��"
      End
      Begin VB.Menu mnu_�~�y�j�r�� 
         Caption         =   "�~�y�j�r��"
      End
      Begin VB.Menu mnu_line1_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_�p�f 
         Caption         =   "����Ѧr"
      End
      Begin VB.Menu mnu_line1_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_���� 
         Caption         =   "����s"
      End
      Begin VB.Menu mnu_����ϧΤ�r 
         Caption         =   "����s�����W(�ϧΤ�r)"
      End
      Begin VB.Menu mnu_line1_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_�Ұ��� 
         Caption         =   "��V�Ұ�������ġ"
      End
      Begin VB.Menu mnu_���t��r 
         Caption         =   "���t²����r�s"
      End
      Begin VB.Menu mnu_line1_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_���� 
         Caption         =   "�H�W�Ҧ�����r"
      End
      Begin VB.Menu mnu_line1_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_���� 
         Caption         =   "����"
      End
   End
   Begin VB.Menu mnu_�r 
      Caption         =   "�r��"
      Begin VB.Menu mnu_�r�δF�� 
         Caption         =   "�����˦r..."
      End
      Begin VB.Menu mnu_�X�B�˦r 
         Caption         =   "�X�B�˦r..."
      End
      Begin VB.Menu mnu_�r�ε��c 
         Caption         =   "�r�ε��c..."
      End
      Begin VB.Menu mnu_�r�κt�� 
         Caption         =   "�r�κt��..."
      End
      Begin VB.Menu mnu_�r�ί��� 
         Caption         =   "�r�ί���..."
      End
      Begin VB.Menu mnu_����r�� 
         Caption         =   "����r��..."
      End
   End
   Begin VB.Menu mnu_���� 
      Caption         =   "����"
      Begin VB.Menu mnu_�d���r�峡�� 
         Caption         =   "�d���r�峡��..."
      End
      Begin VB.Menu mnu_����Ѧr���� 
         Caption         =   "����Ѧr����..."
      End
      Begin VB.Menu mnu_line3_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_����N�X 
         Caption         =   "Big5�r��..."
         Index           =   1
      End
      Begin VB.Menu mnu_����N�X 
         Caption         =   "Big5��²�Ʀr�r��..."
         Index           =   2
      End
      Begin VB.Menu mnu_����N�X 
         Caption         =   "�r��..."
         Index           =   3
      End
      Begin VB.Menu mnu_����N�X 
         Caption         =   "�p�f�W��r..."
         Index           =   4
      End
      Begin VB.Menu mnu_����N�X 
         Caption         =   "����r��..."
         Index           =   5
      End
      Begin VB.Menu mnu_����N�X 
         Caption         =   "�Ұ���r��..."
         Index           =   6
      End
      Begin VB.Menu mnu_����N�X 
         Caption         =   "���t²����r�r��..."
         Index           =   7
      End
      Begin VB.Menu mnu_����N�X 
         Caption         =   "����~�r..."
         Index           =   8
      End
      Begin VB.Menu mnuline3_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_����r�� 
         Caption         =   "����r��..."
      End
   End
   Begin VB.Menu mnu_�Ÿ� 
      Caption         =   "�Ÿ�"
      Begin VB.Menu mnu_�c�r�Ÿ� 
         Caption         =   "�c�r�Ÿ�..."
      End
      Begin VB.Menu mnu_�ϧΤ�r 
         Caption         =   "�ϧΤ�r(������)..."
      End
      Begin VB.Menu mnu_�K�� 
         Caption         =   "�K��..."
      End
      Begin VB.Menu mnu_²�| 
         Caption         =   "²�|..."
      End
   End
   Begin VB.Menu mnu_�r�� 
      Caption         =   "�r��"
      Visible         =   0   'False
      Begin VB.Menu mnu_�ѦҦr�� 
         Caption         =   "�ѦҦr��..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_�r�魷�� 
         Caption         =   "�r�魷��..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu�s�� 
      Caption         =   "�s��"
      Begin VB.Menu mnu_�ƻs 
         Caption         =   "�ƻs"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu_�ƻs�Ϥ� 
         Caption         =   "�ƻs�Ϥ�"
      End
      Begin VB.Menu mnu_�ƻs�S��Ϥ� 
         Caption         =   "�ƻs�S��Ϥ�"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_edit_�ƻs��Word 
         Caption         =   "�ƻs��Microsoft Word"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnu�K�W 
         Caption         =   "�K�W"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnu_line5_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_�r�� 
         Caption         =   "�]�w��ܦr��..."
      End
      Begin VB.Menu mnu_�Ϥ� 
         Caption         =   "�]�w�ƻs�Ϥ�..."
      End
   End
   Begin VB.Menu mnu_Tool 
      Caption         =   "�u��"
      Begin VB.Menu mnu_Tool_ListLikeChar 
         Caption         =   "�C�X�ۦP�X�B���r��"
      End
      Begin VB.Menu mnuToolListChar 
         Caption         =   "�C�X�Ҧ��r��"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_�ﶵ 
      Caption         =   "�ﶵ"
      Begin VB.Menu mnu_�����˦r�ﶵ 
         Caption         =   "�����˦r"
         Begin VB.Menu mnu_�r�δF�ųv�ŦC�X��@���� 
            Caption         =   "�����˦r�v�ŦC�X����P�r��(����@����)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_�r�δF�ťu�C�r�Τ��C���� 
            Caption         =   "�����˦r�u�C�r�Τ��C����"
         End
         Begin VB.Menu mnu_line6_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_�r�δF�ſ�ӿ�J���󶶧� 
            Caption         =   "�����˦r��ӿ�J���󶶧�"
         End
         Begin VB.Menu mnu_�r�δF�ť]�t���g���� 
            Caption         =   "�����˦r�]�t����r��"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_�r�δF�űĥ�SQL�y�k 
            Caption         =   "�����˦r�ĥ�SQL Like�y�k"
         End
      End
      Begin VB.Menu mnu_�r�ε��c�ﶵ 
         Caption         =   "�r�ε��c"
         Begin VB.Menu mnu_����ѧζ��� 
            Caption         =   "�r�ε��c�̾ڻ���ѧζ��ǦC�X�p�f����"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu_�r�ί��޿ﶵ 
         Caption         =   "�r�ί���"
         Begin VB.Menu mnu_���F�~�y�j�r��ﶵ 
            Caption         =   "�~�y�j�r��(���F�ϮѤ��q)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_�ا��~�y�j�r��ﶵ 
            Caption         =   "�~�y�j�r��(�ا��X����)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_line6_2 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_����j���ﶵ 
            Caption         =   "����j���"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_line6_3 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_����Ѧr���L�ﶵ 
            Caption         =   "����Ѧr���L"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_���ػ���Ѧr�ﶵ 
            Caption         =   "����Ѧr(���خѧ�)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_line6_4 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_����s�ﶵ 
            Caption         =   "����s"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_������L�ﶵ 
            Caption         =   "������L"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_��P���嶰�������ﶵ 
            Caption         =   "��P���嶰������"
            Checked         =   -1  'True
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_��P���嶰���ޱo�ﶵ 
            Caption         =   "��P���嶰���ޱo"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_line6_5 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_�Ұ�������ġ�ﶵ 
            Caption         =   "��V�Ұ�������ġ"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_�Ұ���r���L�ﶵ 
            Caption         =   "�Ұ���r���L"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_�Ұ���r�����ﶵ 
            Caption         =   "�Ұ���r����"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_line6_6 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_���t²����r�s�ﶵ 
            Caption         =   "���t²����r�s"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_���t��r�X�B�ﶵ 
            Caption         =   "���t��r�X�B"
            Checked         =   -1  'True
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_line6_7 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_Unicode�ﶵ 
            Caption         =   "Unicode"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_Big5�ﶵ 
            Caption         =   "Big5"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu_�r�κt�ܿﶵ 
         Caption         =   "�r�κt��"
         Begin VB.Menu mnu_�Ұ���ﶵ 
            Caption         =   "�Ұ���"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_����ﶵ 
            Caption         =   "����"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_���t��r�ﶵ 
            Caption         =   "���t��r"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_�p�f�ﶵ 
            Caption         =   "�p�f"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu_option_�ƻs��Word 
         Caption         =   "�ƻs��Microsoft Word"
         Begin VB.Menu mnu_�ƻs�r�Ψ�Word 
            Caption         =   "�ƻs���r�Ωκc�r��"
         End
         Begin VB.Menu mnu_�ƻs�Ϥ���Word 
            Caption         =   "�ƻs���Ϥ�"
         End
         Begin VB.Menu mnu_�ƻsUnicode�r�Ψ�Word 
            Caption         =   "�ƻsUnicode�r��"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_���z���ƻs��Word 
            Caption         =   "���z���ƻs"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu_�ƻs�ﶵ 
         Caption         =   "��L"
         Begin VB.Menu mnu_�ƻs��ŶKï 
            Caption         =   "�۰ʽƻs����r�Ψ�ŶKï"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_��ܭ���X 
            Caption         =   "��ܭ���X"
            Checked         =   -1  'True
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnu_�x�s�����]�w 
         Caption         =   "�������x�s�����]�w"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_�P�ɳ]�w�Ҧ��}�ҵ������r���j�p���C�� 
         Caption         =   "�P�ɳ]�w�Ҧ��}�ҵ������r���j�p���C��"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_���� 
      Caption         =   "����"
      WindowList      =   -1  'True
      Begin VB.Menu mnu_�����ñ� 
         Caption         =   "�����ñ�"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_�����ñ� 
         Caption         =   "�����ñ�"
      End
      Begin VB.Menu mnu_���|��� 
         Caption         =   "���|���"
      End
      Begin VB.Menu mnu_�s�W���� 
         Caption         =   "�s�W����"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_�ƦC�ϥ� 
         Caption         =   "�ƦC�ϥ�"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_²���s�� 
         Caption         =   "�w�]�}��(��)..."
      End
      Begin VB.Menu mnu_�̨��s�� 
         Caption         =   "�w�]�}��(�i��)..."
      End
   End
   Begin VB.Menu mnu_���� 
      Caption         =   "����"
      Begin VB.Menu mnu_�����D�D 
         Caption         =   "�����D�D"
      End
      Begin VB.Menu mnu_line8_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_cdp 
         Caption         =   "����~�r�c�θ�Ʈw"
      End
   End
End
Attribute VB_Name = "mdi�~�r�r��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ���e As Integer, ���� As Integer, ���A�C As String, ���A�C1 As String
Private ��lfont As String, ��lfontsize As Integer
Private ��lleft As Integer, ��ltop As Integer, ��lwidth As Integer, ��lheight As Integer
Private ��lsave As String * 250

Private path As String

Private Sub ���J��l��()
Dim nDefault As Long, sDefault As String, lret As Long

��lfirst = GetPrivateProfileInt("Start", "first", nDefault, App.path & "\cdphanzi.ini")

��lleft = GetPrivateProfileInt("Window", "left", nDefault, App.path & "\cdphanzi.ini")
��ltop = GetPrivateProfileInt("Window", "top", nDefault, App.path & "\cdphanzi.ini")
��lwidth = GetPrivateProfileInt("Window", "width", nDefault, App.path & "\cdphanzi.ini")
��lheight = GetPrivateProfileInt("Window", "height", nDefault, App.path & "\cdphanzi.ini")

�r��_�`�Φr = GetPrivateProfileInt("�r��", "�`�Φr", nDefault, App.path & "\cdphanzi.ini")
�r��_���j�X = GetPrivateProfileInt("�r��", "���j�X", nDefault, App.path & "\cdphanzi.ini")
�r��_²�Ʀr = GetPrivateProfileInt("�r��", "²�Ʀr", nDefault, App.path & "\cdphanzi.ini")
�r��_�~�y�j�r�� = GetPrivateProfileInt("�r��", "�~�y�j�r��", nDefault, App.path & "\cdphanzi.ini")
�r��_����Ѧr = GetPrivateProfileInt("�r��", "����Ѧr", nDefault, App.path & "\cdphanzi.ini")
�r��_����s = GetPrivateProfileInt("�r��", "����s", nDefault, App.path & "\cdphanzi.ini")
�r��_����s�ϧΤ�r = GetPrivateProfileInt("�r��", "����s�ϧΤ�r", nDefault, App.path & "\cdphanzi.ini")
�r��_�Ұ���ġ = GetPrivateProfileInt("�r��", "�Ұ���ġ", nDefault, App.path & "\cdphanzi.ini")
�r��_���t²����r�s = GetPrivateProfileInt("�r��", "���t²����r�s", nDefault, App.path & "\cdphanzi.ini")
�r��_����r = GetPrivateProfileInt("�r��", "����r", nDefault, App.path & "\cdphanzi.ini")

�F��open = GetPrivateProfileInt("�r�δF��", "open", nDefault, App.path & "\cdphanzi.ini")
�F��winstate = GetPrivateProfileInt("�r�δF��", "winstate", nDefault, App.path & "\cdphanzi.ini")
�F��left = GetPrivateProfileInt("�r�δF��", "left", nDefault, App.path & "\cdphanzi.ini")
�F��top = GetPrivateProfileInt("�r�δF��", "top", nDefault, App.path & "\cdphanzi.ini")
�F��width = GetPrivateProfileInt("�r�δF��", "width", nDefault, App.path & "\cdphanzi.ini")
�F��height = GetPrivateProfileInt("�r�δF��", "height", nDefault, App.path & "\cdphanzi.ini")

�X�Bopen = GetPrivateProfileInt("�X�B�˦r", "open", nDefault, App.path & "\cdphanzi.ini")
�X�Bwinstate = GetPrivateProfileInt("�X�B�˦r", "winstate", nDefault, App.path & "\cdphanzi.ini")
�X�Bleft = GetPrivateProfileInt("�X�B�˦r", "left", nDefault, App.path & "\cdphanzi.ini")
�X�Btop = GetPrivateProfileInt("�X�B�˦r", "top", nDefault, App.path & "\cdphanzi.ini")
�X�Bwidth = GetPrivateProfileInt("�X�B�˦r", "width", nDefault, App.path & "\cdphanzi.ini")
�X�Bheight = GetPrivateProfileInt("�X�B�˦r", "height", nDefault, App.path & "\cdphanzi.ini")

���copen = GetPrivateProfileInt("�r�ε��c", "open", nDefault, App.path & "\cdphanzi.ini")
���cwinstate = GetPrivateProfileInt("�r�ε��c", "winstate", nDefault, App.path & "\cdphanzi.ini")
���cleft = GetPrivateProfileInt("�r�ε��c", "left", nDefault, App.path & "\cdphanzi.ini")
���ctop = GetPrivateProfileInt("�r�ε��c", "top", nDefault, App.path & "\cdphanzi.ini")
���cwidth = GetPrivateProfileInt("�r�ε��c", "width", nDefault, App.path & "\cdphanzi.ini")
���cheight = GetPrivateProfileInt("�r�ε��c", "height", nDefault, App.path & "\cdphanzi.ini")

����open = GetPrivateProfileInt("����r��", "open", nDefault, App.path & "\cdphanzi.ini")
����winstate = GetPrivateProfileInt("����r��", "winstate", nDefault, App.path & "\cdphanzi.ini")
����left = GetPrivateProfileInt("����r��", "left", nDefault, App.path & "\cdphanzi.ini")
����top = GetPrivateProfileInt("����r��", "top", nDefault, App.path & "\cdphanzi.ini")
����width = GetPrivateProfileInt("����r��", "width", nDefault, App.path & "\cdphanzi.ini")
����height = GetPrivateProfileInt("����r��", "height", nDefault, App.path & "\cdphanzi.ini")

����open = GetPrivateProfileInt("����r��", "open", nDefault, App.path & "\cdphanzi.ini")
����winstate = GetPrivateProfileInt("����r��", "winstate", nDefault, App.path & "\cdphanzi.ini")
����left = GetPrivateProfileInt("����r��", "left", nDefault, App.path & "\cdphanzi.ini")
����top = GetPrivateProfileInt("����r��", "top", nDefault, App.path & "\cdphanzi.ini")
����width = GetPrivateProfileInt("����r��", "width", nDefault, App.path & "\cdphanzi.ini")
����height = GetPrivateProfileInt("����r��", "height", nDefault, App.path & "\cdphanzi.ini")

����open = GetPrivateProfileInt("�c�γ���", "open", nDefault, App.path & "\cdphanzi.ini")
����winstate = GetPrivateProfileInt("�c�γ���", "winstate", nDefault, App.path & "\cdphanzi.ini")
����left = GetPrivateProfileInt("�c�γ���", "left", nDefault, App.path & "\cdphanzi.ini")
����top = GetPrivateProfileInt("�c�γ���", "top", nDefault, App.path & "\cdphanzi.ini")
����width = GetPrivateProfileInt("�c�γ���", "width", nDefault, App.path & "\cdphanzi.ini")
����height = GetPrivateProfileInt("�c�γ���", "height", nDefault, App.path & "\cdphanzi.ini")

�t��open = GetPrivateProfileInt("�r�κt��", "open", nDefault, App.path & "\cdphanzi.ini")
�t��winstate = GetPrivateProfileInt("�r�κt��", "winstate", nDefault, App.path & "\cdphanzi.ini")
�t��left = GetPrivateProfileInt("�r�κt��", "left", nDefault, App.path & "\cdphanzi.ini")
�t��top = GetPrivateProfileInt("�r�κt��", "top", nDefault, App.path & "\cdphanzi.ini")
�t��width = GetPrivateProfileInt("�r�κt��", "width", nDefault, App.path & "\cdphanzi.ini")
�t��height = GetPrivateProfileInt("�r�κt��", "height", nDefault, App.path & "\cdphanzi.ini")

����open = GetPrivateProfileInt("�r�ί���", "open", nDefault, App.path & "\cdphanzi.ini")
����winstate = GetPrivateProfileInt("�r�ί���", "winstate", nDefault, App.path & "\cdphanzi.ini")
����left = GetPrivateProfileInt("�r�ί���", "left", nDefault, App.path & "\cdphanzi.ini")
����top = GetPrivateProfileInt("�r�ί���", "top", nDefault, App.path & "\cdphanzi.ini")
����width = GetPrivateProfileInt("�r�ί���", "width", nDefault, App.path & "\cdphanzi.ini")
����height = GetPrivateProfileInt("�r�ί���", "height", nDefault, App.path & "\cdphanzi.ini")

��lfont = String(256, 0)
lret = GetPrivateProfileString("Font", "fontname", "�з���", ��lfont, Len(��lfont), App.path & "\cdphanzi.ini")
��lfont = Left(��lfont, InStr(��lfont, Chr(0)) - 1)
��lfontsize = GetPrivateProfileInt("Font", "fontsize", nDefault, App.path & "\cdphanzi.ini")

�Ϥ��ѪR�� = GetPrivateProfileInt("Image", "dpi", nDefault, App.path & "\cdphanzi.ini")
�Ϥ��r���j�p = GetPrivateProfileInt("Image", "fontsize", nDefault, App.path & "\cdphanzi.ini")

��l�v�ŦC�X = GetPrivateProfileInt("�ﶵ", "�v�ŦC�X", nDefault, App.path & "\cdphanzi.ini")
��l���󶶧� = GetPrivateProfileInt("�ﶵ", "���󶶧�", nDefault, App.path & "\cdphanzi.ini")
��l���g���� = GetPrivateProfileInt("�ﶵ", "���g����", nDefault, App.path & "\cdphanzi.ini")
��l�ѧΦC�X = GetPrivateProfileInt("�ﶵ", "�ѧΦC�X", nDefault, App.path & "\cdphanzi.ini")
��l���F�~�y�j�r�� = GetPrivateProfileInt("�ﶵ", "���F�~�y�j�r��", nDefault, App.path & "\cdphanzi.ini")
��l�ا��~�y�j�r�� = GetPrivateProfileInt("�ﶵ", "�ا��~�y�j�r��", nDefault, App.path & "\cdphanzi.ini")
��l����j��� = GetPrivateProfileInt("�ﶵ", "����j���", nDefault, App.path & "\cdphanzi.ini")
��l����Ѧr���L = GetPrivateProfileInt("�ﶵ", "����Ѧr���L", nDefault, App.path & "\cdphanzi.ini")
��l���ػ���Ѧr = GetPrivateProfileInt("�ﶵ", "���ػ���Ѧr", nDefault, App.path & "\cdphanzi.ini")
��l����s = GetPrivateProfileInt("�ﶵ", "����s", nDefault, App.path & "\cdphanzi.ini")
��l������L = GetPrivateProfileInt("�ﶵ", "������L", nDefault, App.path & "\cdphanzi.ini")
'��l���徹�� = GetPrivateProfileInt("�ﶵ", "���徹��", nDefault, App.path & "\cdphanzi.ini")
��l����ޱo = GetPrivateProfileInt("�ﶵ", "����ޱo", nDefault, App.path & "\cdphanzi.ini")
��l�Ұ�������ġ = GetPrivateProfileInt("�ﶵ", "�Ұ�������ġ", nDefault, App.path & "\cdphanzi.ini")
��l�Ұ���r���L = GetPrivateProfileInt("�ﶵ", "�Ұ���r���L", nDefault, App.path & "\cdphanzi.ini")
��l�Ұ���r���� = GetPrivateProfileInt("�ﶵ", "�Ұ���r����", nDefault, App.path & "\cdphanzi.ini")
��l���t²����r�s = GetPrivateProfileInt("�ﶵ", "���t²����r�s", nDefault, App.path & "\cdphanzi.ini")
'��l���t��r�X�B = GetPrivateProfileInt("�ﶵ", "���t��r�X�B", nDefault, App.path & "\cdphanzi.ini")
��lUnicode = GetPrivateProfileInt("�ﶵ", "Unicode", nDefault, App.path & "\cdphanzi.ini")
��lBig5 = GetPrivateProfileInt("�ﶵ", "Big5", nDefault, App.path & "\cdphanzi.ini")
��l�Ұ���t�� = GetPrivateProfileInt("�ﶵ", "�Ұ���t��", nDefault, App.path & "\cdphanzi.ini")
��l����t�� = GetPrivateProfileInt("�ﶵ", "����t��", nDefault, App.path & "\cdphanzi.ini")
��l���t��r�t�� = GetPrivateProfileInt("�ﶵ", "���t��r�t��", nDefault, App.path & "\cdphanzi.ini")
��l�p�f�t�� = GetPrivateProfileInt("�ﶵ", "�p�f�t��", nDefault, App.path & "\cdphanzi.ini")
��lcopy = GetPrivateProfileInt("�ﶵ", "copy", nDefault, App.path & "\cdphanzi.ini")
��l�r�W = GetPrivateProfileInt("Window", "�r�W", nDefault, App.path & "\cdphanzi.ini")
��lsave = GetPrivateProfileInt("�ﶵ", "save", nDefault, App.path & "\cdphanzi.ini")
'��l����X = GetPrivateProfileInt("�ﶵ", "����X", nDefault, App.path & "\cdphanzi.ini")

��lCopyToWord = GetPrivateProfileInt("CopyToWord", "CopyMode", nDefault, App.path & "\cdphanzi.ini")
��lCopyUnicode = GetPrivateProfileInt("CopyToWord", "CopyUnicode", nDefault, App.path & "\cdphanzi.ini")

End Sub

Private Sub �}�Ҹ�Ʈw()

Set �t�θ�Ʈw = OpenDatabase(App.path & "\cdphanzi.mdb")
Set �p�f��Ʈw = OpenDatabase(App.path & "\cdpseal.mdb")
Set �����Ʈw = OpenDatabase(App.path & "\cdpbronz.mdb")
Set �Ұ����Ʈw = OpenDatabase(App.path & "\cdpjiagu.mdb")
Set ���t��r��Ʈw = OpenDatabase(App.path & "\cdpchubs.mdb")

�r��W��

End Sub

Private Sub �x�s������()
Dim IniEntry As String * 255
Dim fsuccess As Boolean

If mdi�~�r�r��.mnu_�x�s�����]�w.Checked = True Then
   ��lsave = 1
Else
   ��lsave = 0
   Exit Sub
End If

��lfirst = 2

��lleft = mdi�~�r�r��.Left
��ltop = mdi�~�r�r��.Top
��lheight = mdi�~�r�r��.Height
��lwidth = mdi�~�r�r��.Width

If mnu_�`�Φr.Checked Then
    �r��_�`�Φr = 1
Else
    �r��_�`�Φr = 0
End If

If mnu_Big5.Checked Then
    �r��_���j�X = 1
Else
    �r��_���j�X = 0
End If

If mnu_²�Ʀr�`��.Checked Then
    �r��_²�Ʀr = 1
Else
    �r��_²�Ʀr = 0
End If

If mnu_�~�y�j�r��.Checked Then
    �r��_�~�y�j�r�� = 1
Else
    �r��_�~�y�j�r�� = 0
End If

If mnu_�p�f.Checked Then
    �r��_����Ѧr = 1
Else
    �r��_����Ѧr = 0
End If

If mnu_����.Checked Then
    �r��_����s = 1
Else
    �r��_����s = 0
End If

If mnu_����ϧΤ�r.Checked Then
    �r��_����s�ϧΤ�r = 1
Else
    �r��_����s�ϧΤ�r = 0
End If

If mnu_�Ұ���.Checked Then
    �r��_�Ұ���ġ = 1
Else
    �r��_�Ұ���ġ = 0
End If

If mnu_���t��r.Checked Then
    �r��_���t²����r�s = 1
Else
    �r��_���t²����r�s = 0
End If

If mnu_����.Checked Then
    �r��_����r = 1
Else
    �r��_����r = 0
End If
    
If mdi�~�r�r��.mnu_�r�δF�ųv�ŦC�X��@����.Checked = True Then
   ��l�v�ŦC�X = 1
Else
   ��l�v�ŦC�X = 0
End If

If mdi�~�r�r��.mnu_�r�δF�ſ�ӿ�J���󶶧�.Checked = True Then
   ��l���󶶧� = 1
Else
   ��l���󶶧� = 0
End If

If mdi�~�r�r��.mnu_�r�δF�ť]�t���g����.Checked = True Then
   ��l���g���� = 1
Else
   ��l���g���� = 0
End If

If mdi�~�r�r��.mnu_����ѧζ���.Checked = True Then
    ��l�ѧΦC�X = 1
Else
    ��l�ѧΦC�X = 0
End If

If mnu_���F�~�y�j�r��ﶵ.Checked = True Then
   ��l���F�~�y�j�r�� = 1
Else
   ��l���F�~�y�j�r�� = 0
End If

If mnu_�ا��~�y�j�r��ﶵ.Checked = True Then
   ��l�ا��~�y�j�r�� = 1
Else
   ��l�ا��~�y�j�r�� = 0
End If

If mnu_����j���ﶵ.Checked = True Then
   ��l����j��� = 1
Else
   ��l����j��� = 0
End If

If mnu_����Ѧr���L�ﶵ.Checked = True Then
   ��l����Ѧr���L = 1
Else
   ��l����Ѧr���L = 0
End If

If mnu_���ػ���Ѧr�ﶵ.Checked = True Then
   ��l���ػ���Ѧr = 1
Else
   ��l���ػ���Ѧr = 0
End If

If mnu_����s�ﶵ.Checked = True Then
   ��l����s = 1
Else
   ��l����s = 0
End If

If mnu_������L�ﶵ.Checked = True Then
   ��l������L = 1
Else
   ��l������L = 0
End If

If mnu_��P���嶰�������ﶵ.Checked = True Then
   ��l���徹�� = 1
Else
   ��l���徹�� = 0
End If

If mnu_��P���嶰���ޱo�ﶵ.Checked = True Then
   ��l����ޱo = 1
Else
   ��l����ޱo = 0
End If

If mnu_�Ұ�������ġ�ﶵ.Checked = True Then
   ��l�Ұ�������ġ = 1
Else
   ��l�Ұ�������ġ = 0
End If

If mnu_�Ұ���r���L�ﶵ.Checked = True Then
   ��l�Ұ���r���L = 1
Else
   ��l�Ұ���r���L = 0
End If

If mnu_�Ұ���r�����ﶵ.Checked = True Then
   ��l�Ұ���r���� = 1
Else
   ��l�Ұ���r���� = 0
End If

If mnu_���t²����r�s�ﶵ.Checked = True Then
   ��l���t²����r�s = 1
Else
   ��l���t²����r�s = 0
End If

If mnu_���t��r�X�B�ﶵ.Checked = True Then
   ��l���t��r�X�B = 1
Else
   ��l���t��r�X�B = 0
End If

If mnu_Unicode�ﶵ.Checked = True Then
   ��lUnicode = 1
Else
   ��lUnicode = 0
End If

If mnu_Big5�ﶵ.Checked = True Then
   ��lBig5 = 1
Else
   ��lBig5 = 0
End If

If mnu_�Ұ���ﶵ.Checked = True Then
   ��l�Ұ���t�� = 1
Else
   ��l�Ұ���t�� = 0
End If

If mnu_����ﶵ.Checked = True Then
   ��l����t�� = 1
Else
   ��l����t�� = 0
End If

If mnu_���t��r�ﶵ.Checked = True Then
   ��l���t��r�t�� = 1
Else
   ��l���t��r�t�� = 0
End If

If mnu_�p�f�ﶵ.Checked = True Then
   ��l�p�f�t�� = 1
Else
   ��l�p�f�t�� = 0
End If

��lfont = mdi�~�r�r��.cbo�r���W��.Text
��lfontsize = mdi�~�r�r��.cbo�r��j�p.Text

If mdi�~�r�r��.mnu_�ƻs��ŶKï.Checked = True Then
   ��lcopy = 1
Else
   ��lcopy = 0
End If

If mdi�~�r�r��.mnu_��ܭ���X.Checked = True Then
   ��l����X = 1
Else
   ��l����X = 0
End If

If mnu_�ƻs�r�Ψ�Word.Checked = True Then
    ��lCopyToWord = 1
ElseIf mnu_�ƻs�Ϥ���Word.Checked = True Then
    ��lCopyToWord = 2
Else
    ��lCopyToWord = 3
End If

If mnu_�ƻsUnicode�r�Ψ�Word.Checked = True Then
   ��lCopyUnicode = 1
Else
   ��lCopyUnicode = 0
End If

fsuccess = WritePrivateProfileString("Start", "first", ��lfirst, App.path & "\cdphanzi.ini")

IniEntry = ��lleft
fsuccess = WritePrivateProfileString("Window", "left", IniEntry, App.path & "\cdphanzi.ini")
IniEntry = ��ltop
fsuccess = WritePrivateProfileString("Window", "top", IniEntry, App.path & "\cdphanzi.ini")
IniEntry = ��lwidth
fsuccess = WritePrivateProfileString("Window", "width", IniEntry, App.path & "\cdphanzi.ini")
IniEntry = ��lheight
fsuccess = WritePrivateProfileString("Window", "height", IniEntry, App.path & "\cdphanzi.ini")

IniEntry = ��lfont
fsuccess = WritePrivateProfileString("Font", "fontname", IniEntry, App.path & "\cdphanzi.ini")
IniEntry = ��lfontsize
fsuccess = WritePrivateProfileString("Font", "fontsize", IniEntry, App.path & "\cdphanzi.ini")

IniEntry = �Ϥ��ѪR��
fsuccess = WritePrivateProfileString("Image", "dpi", IniEntry, App.path & "\cdphanzi.ini")
IniEntry = �Ϥ��r���j�p
fsuccess = WritePrivateProfileString("Image", "fontsize", IniEntry, App.path & "\cdphanzi.ini")

'If ��lsave = 1 Then

fsuccess = WritePrivateProfileString("�r��", "�`�Φr", �r��_�`�Φr, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r��", "���j�X", �r��_���j�X, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r��", "²�Ʀr", �r��_²�Ʀr, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r��", "�~�y�j�r��", �r��_�~�y�j�r��, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r��", "����Ѧr", �r��_����Ѧr, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r��", "����s", �r��_����s, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r��", "����s�ϧΤ�r", �r��_����s�ϧΤ�r, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r��", "�Ұ���ġ", �r��_�Ұ���ġ, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r��", "���t²����r�s", �r��_���t²����r�s, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r��", "����r", �r��_����r, App.path & "\cdphanzi.ini")


fsuccess = WritePrivateProfileString("�r�δF��", "open", �F��open, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�δF��", "winstate", �F��winstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�δF��", "left", �F��left, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�δF��", "top", �F��top, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�δF��", "width", �F��width, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�δF��", "height", �F��height, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("�X�B�˦r", "open", �X�Bopen, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�X�B�˦r", "winstate", �X�Bwinstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�X�B�˦r", "left", �X�Bleft, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�X�B�˦r", "top", �X�Btop, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�X�B�˦r", "width", �X�Bwidth, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�X�B�˦r", "height", �X�Bheight, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("�r�ε��c", "open", ���copen, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�ε��c", "winstate", ���cwinstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�ε��c", "left", ���cleft, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�ε��c", "top", ���ctop, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�ε��c", "width", ���cwidth, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�ε��c", "height", ���cheight, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("����r��", "open", ����open, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("����r��", "winstate", ����winstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("����r��", "left", ����left, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("����r��", "top", ����top, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("����r��", "width", ����width, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("����r��", "height", ����height, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("����r��", "open", ����open, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("����r��", "winstate", ����winstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("����r��", "left", ����left, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("����r��", "top", ����top, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("����r��", "width", ����width, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("����r��", "height", ����height, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("�c�γ���", "open", ����open, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�c�γ���", "winstate", ����winstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�c�γ���", "left", ����left, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�c�γ���", "top", ����top, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�c�γ���", "width", ����width, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�c�γ���", "height", ����height, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("�r�κt��", "open", �t��open, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�κt��", "winstate", �t��winstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�κt��", "left", �t��left, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�κt��", "top", �t��top, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�κt��", "width", �t��width, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�κt��", "height", �t��height, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("�r�ί���", "open", ����open, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�ί���", "winstate", ����winstate, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�ί���", "left", ����left, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�ί���", "top", ����top, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�ί���", "width", ����width, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�r�ί���", "height", ����height, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("�ﶵ", "�v�ŦC�X", ��l�v�ŦC�X, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "���󶶧�", ��l���󶶧�, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "���g����", ��l���g����, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "�ѧΦC�X", ��l�ѧΦC�X, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "���F�~�y�j�r��", ��l���F�~�y�j�r��, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "�ا��~�y�j�r��", ��l�ا��~�y�j�r��, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "����j���", ��l����j���, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "����Ѧr���L", ��l����Ѧr���L, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "���ػ���Ѧr", ��l���ػ���Ѧr, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "����s", ��l����s, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "������L", ��l������L, App.path & "\cdphanzi.ini")
'fsuccess = WritePrivateProfileString("�ﶵ", "���徹��", ��l���徹��, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "����ޱo", ��l����ޱo, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "�Ұ�������ġ", ��l�Ұ�������ġ, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "�Ұ���r���L", ��l�Ұ���r���L, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "�Ұ���r����", ��l�Ұ���r����, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "���t²����r�s", ��l���t²����r�s, App.path & "\cdphanzi.ini")
'fsuccess = WritePrivateProfileString("�ﶵ", "���t��r�X�B", ��l���t��r�X�B, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "Unicode", ��lUnicode, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "Big5", ��lBig5, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "�Ұ���t��", ��l�Ұ���t��, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "����t��", ��l����t��, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "���t��r�t��", ��l���t��r�t��, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "�p�f�t��", ��l�p�f�t��, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "copy", ��lcopy, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "�r�W", ��l�r�W, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("�ﶵ", "save", ��lsave, App.path & "\cdphanzi.ini")
'fsuccess = WritePrivateProfileString("�ﶵ", "����X", ��l����X, App.path & "\cdphanzi.ini")

fsuccess = WritePrivateProfileString("CopyToWord", "CopyMode", ��lCopyToWord, App.path & "\cdphanzi.ini")
fsuccess = WritePrivateProfileString("CopyToWord", "CopyUnicode", ��lCopyUnicode, App.path & "\cdphanzi.ini")

'End If

End Sub

Private Sub ��l�e���]�w()
Me.show
If ��lfirst = 1 Then
   mdi�~�r�r��.WindowState = 2
Else
   mdi�~�r�r��.Left = ��lleft
   mdi�~�r�r��.Top = ��ltop
   mdi�~�r�r��.Height = ��lheight
   mdi�~�r�r��.Width = ��lwidth
End If

'mdi�~�r�r��.tbr�r���ݩ�.ButtonHeight = 360
'mdi�~�r�r��.tbr����r��.ButtonHeight = 360

cbo�r���W��.Text = ��lfont
If ��lfirst = 1 Then
   cbo�r��j�p.Text = 24
Else
   cbo�r��j�p.Text = ��lfontsize
End If

��e = CInt(cbo�r��j�p.Text) * 20 + CInt(cbo�r��j�p.Text) * 20 / 3

End Sub

Private Sub ���J����()

If ��lfirst = 1 Then
   mnu_�x�s�����]�w.Checked = True
   mnu_�ƻs��ŶKï.Checked = True
   mnu_�r�δF�ť]�t���g����.Checked = True
   mnu_�r�δF�ųv�ŦC�X��@����.Checked = False
   mnu_�r�δF�ťu�C�r�Τ��C����.Checked = True

   'mnu_�r�δF�ťu�C�X�`�Φr.Checked = False
   'mnu_�r�δF�ťu�C�X�q���Φr.Checked = False
   'mnu_�r�δF�ŦC�X�Ҧ��r��.Checked = True
   ��l�r�W = 1
   mnu_����_Click
   �w�]�s���@
Else
   �w���J�e�� = 0
   
   ��l�r�W = 1
   
   If CInt(�r��_�`�Φr) = 1 Then mnu_�`�Φr_Click
   If CInt(�r��_���j�X) = 1 Then mnu_Big5_Click
   If CInt(�r��_²�Ʀr) = 1 Then mnu_²�Ʀr�`��_Click
   If CInt(�r��_�~�y�j�r��) = 1 Then mnu_�~�y�j�r��_Click
   If CInt(�r��_����Ѧr) = 1 Then mnu_�p�f_Click
   If CInt(�r��_����s) = 1 Then mnu_����_Click
   If CInt(�r��_����s�ϧΤ�r) = 1 Then mnu_����ϧΤ�r_Click
   If CInt(�r��_�Ұ���ġ) = 1 Then mnu_�Ұ���_Click
   If CInt(�r��_���t²����r�s) = 1 Then mnu_���t��r_Click
   If CInt(�r��_����r) = 1 Then mnu_����_Click
    
   If ����open = 1 Then
        If �t�Φr�� = "�p�f" Then
            mnu_����Ѧr����_Click
        Else
            mnu_�d���r�峡��_Click
        End If
   End If
   
   If ���copen = 1 Then
      frm�r�ε��c.Tag = �r�ε��c�N�X
      frm�r�ε��c.show
   End If
   
   If ����open = 1 Then
      frm����r��.Tag = ����r��N�X
      frm����r��.show
   End If
   
   If ����open = 1 Then
      frm����r��.Tag = ����r�ڥN�X
      frm����r��.show
   End If

   If �t��open = 1 Then
      frm�r�κt��.Tag = �r�κt�ܥN�X
      frm�r�κt��.show
   End If

   If ����open = 1 Then
      frm�r�ί���.Tag = �r�ί��ޥN�X
      frm�r�ί���.show
   End If

   If �X�Bopen = 1 Then
      frm�X�B�˦r.Tag = �X�B�˦r�N�X
      frm�X�B�˦r.show
   End If
   
   If �F��open = 1 Then
      frm�r�δF��.Tag = �r�δF�ťN�X
      frm�r�δF��.show
   End If
      
   If ��l�v�ŦC�X = 1 Then
      mnu_�r�δF�ųv�ŦC�X��@����.Checked = True
      mnu_�r�δF�ťu�C�r�Τ��C����.Checked = False
   Else
      mnu_�r�δF�ųv�ŦC�X��@����.Checked = False
      mnu_�r�δF�ťu�C�r�Τ��C����.Checked = True
   End If
   
   If ��l���󶶧� = 1 Then
      mnu_�r�δF�ſ�ӿ�J���󶶧�.Checked = True
   Else
      mnu_�r�δF�ſ�ӿ�J���󶶧�.Checked = False
   End If

   If ��l���g���� = 1 Then
      mnu_�r�δF�ť]�t���g����.Checked = True
   Else
      mnu_�r�δF�ť]�t���g����.Checked = False
   End If

   If ��l�ѧΦC�X = 1 Then
      mdi�~�r�r��.mnu_����ѧζ���.Checked = True
   Else
      mdi�~�r�r��.mnu_����ѧζ���.Checked = False
   End If
   
   If ��l���F�~�y�j�r�� = 1 Then
      mnu_���F�~�y�j�r��ﶵ.Checked = True
   Else
      mnu_���F�~�y�j�r��ﶵ.Checked = False
   End If

   If ��l�ا��~�y�j�r�� = 1 Then
      mnu_�ا��~�y�j�r��ﶵ.Checked = True
   Else
      mnu_�ا��~�y�j�r��ﶵ.Checked = False
   End If

   If ��l����j��� = 1 Then
      mnu_����j���ﶵ.Checked = True
   Else
      mnu_����j���ﶵ.Checked = False
   End If

   If ��l����Ѧr���L = 1 Then
      mnu_����Ѧr���L�ﶵ.Checked = True
   Else
      mnu_����Ѧr���L�ﶵ.Checked = False
   End If

   If ��l���ػ���Ѧr = 1 Then
      mnu_���ػ���Ѧr�ﶵ.Checked = True
   Else
      mnu_���ػ���Ѧr�ﶵ.Checked = False
   End If

   If ��l����s = 1 Then
      mnu_����s�ﶵ.Checked = True
   Else
      mnu_����s�ﶵ.Checked = False
   End If

   If ��l������L = 1 Then
      mnu_������L�ﶵ.Checked = True
   Else
      mnu_������L�ﶵ.Checked = False
   End If

   'If ��l���徹�� = 1 Then
   '   mnu_��P���嶰�������ﶵ.Checked = True
   'Else
   '   mnu_��P���嶰�������ﶵ.Checked = False
   'End If

   If ��l����ޱo = 1 Then
      mnu_��P���嶰���ޱo�ﶵ.Checked = True
   Else
      mnu_��P���嶰���ޱo�ﶵ.Checked = False
   End If

   If ��l�Ұ�������ġ = 1 Then
      mnu_�Ұ�������ġ�ﶵ.Checked = True
   Else
      mnu_�Ұ�������ġ�ﶵ.Checked = False
   End If

   If ��l�Ұ���r���L = 1 Then
      mnu_�Ұ���r���L�ﶵ.Checked = True
   Else
      mnu_�Ұ���r���L�ﶵ.Checked = False
   End If

   If ��l�Ұ���r���� = 1 Then
      mnu_�Ұ���r�����ﶵ.Checked = True
   Else
      mnu_�Ұ���r�����ﶵ.Checked = False
   End If

   If ��l���t²����r�s = 1 Then
      mnu_���t²����r�s�ﶵ.Checked = True
   Else
      mnu_���t²����r�s�ﶵ.Checked = False
   End If

   'If ��l���t��r�X�B = 1 Then
   '   mnu_���t��r�X�B�ﶵ.Checked = True
   'Else
   '   mnu_���t��r�X�B�ﶵ.Checked = False
   'End If

   If ��lUnicode = 1 Then
      mnu_Unicode�ﶵ.Checked = True
   Else
      mnu_Unicode�ﶵ.Checked = False
   End If

   If ��lBig5 = 1 Then
      mnu_Big5�ﶵ.Checked = True
   Else
      mnu_Big5�ﶵ.Checked = False
   End If

   If ��l�Ұ���t�� = 1 Then
      mnu_�Ұ���ﶵ.Checked = True
   Else
      mnu_�Ұ���ﶵ.Checked = False
   End If

   If ��l����t�� = 1 Then
      mnu_����ﶵ.Checked = True
   Else
      mnu_����ﶵ.Checked = False
   End If

   If ��l���t��r�t�� = 1 Then
      mnu_���t��r�ﶵ.Checked = True
   Else
      mnu_���t��r�ﶵ.Checked = False
   End If

   If ��l�p�f�t�� = 1 Then
      mnu_�p�f�ﶵ.Checked = True
   Else
      mnu_�p�f�ﶵ.Checked = False
   End If
   
   If ��lsave = 1 Then
      mnu_�x�s�����]�w.Checked = True
   Else
      mnu_�x�s�����]�w.Checked = False
   End If

   If ��lcopy = 1 Then
      mnu_�ƻs��ŶKï.Checked = True
   Else
      mnu_�ƻs��ŶKï.Checked = False
   End If
   
   'If ��l����X = 1 Then
   '   mnu_��ܭ���X.Checked = True
   'Else
   '   mnu_��ܭ���X.Checked = False
   'End If
    If ��lCopyToWord = 1 Then
        mnu_�ƻs�r�Ψ�Word.Checked = True
        mnu_�ƻs�Ϥ���Word.Checked = False
        mnu_���z���ƻs��Word.Checked = False
    ElseIf ��lCopyToWord = 2 Then
        mnu_�ƻs�r�Ψ�Word.Checked = False
        mnu_�ƻs�Ϥ���Word.Checked = True
        mnu_���z���ƻs��Word.Checked = False
    Else
        mnu_�ƻs�r�Ψ�Word.Checked = False
        mnu_�ƻs�Ϥ���Word.Checked = False
        mnu_���z���ƻs��Word.Checked = True
    End If
   
   �w���J�e�� = 1
   
End If

End Sub


Public Sub cbo�r���W��_click()

Dim i As Integer
Dim j As Long

If Len(cbo�r���W��.Text) = 0 Then cbo�r���W��.Text = ��ܦr��
If ��ܦr�� = cbo�r���W��.Text Then Exit Sub

��ܦr�� = cbo�r���W��.Text

For i = 1 To Forms.Count - 1

    If (CInt(Forms(i).Tag) >= Big5�r�ڥN�X) And (CInt(Forms(i).Tag) <= �c�r�Ÿ��N�X) Then
       frm����d��.tree�r�ξ𪬵��c.FontName = ��ܦr��
       'For j = 0 To frm����d��.tree�r�ξ𪬵��c.ListCount - 1
       '    frm����d��.tree�r�ξ𪬵��c.ItemFontSize(j) = ��ܦr���j�p
       'Next j
    End If
    
    If CInt(Forms(i).Tag) = �r�δF�ťN�X Then
        frm�r�δF��.tree�r�ξ𪬵��c.FontName = ��ܦr��
            For j = 0 To frm�r�δF��.tree�r�ξ𪬵��c.ListCount - 1
                If frm�r�δF��.tree�r�ξ𪬵��c.ItemFontName(j) <> ��ܦr�� Then
                    frm�r�δF��.tree�r�ξ𪬵��c.ItemFontName(j) = ������ܦr��(frm�r�δF��.tree�r�ξ𪬵��c.ItemFontName(j))
                End If
            Next j
    End If
    
    If CInt(Forms(i).Tag) = �X�B�˦r�N�X Then
        frm�X�B�˦r.tree�r�ξ𪬵��c.FontName = ��ܦr��
            For j = 0 To frm�X�B�˦r.tree�r�ξ𪬵��c.ListCount - 1
                If frm�X�B�˦r.tree�r�ξ𪬵��c.ItemFontName(j) <> ��ܦr�� Then
                    frm�X�B�˦r.tree�r�ξ𪬵��c.ItemFontName(j) = ������ܦr��(frm�X�B�˦r.tree�r�ξ𪬵��c.ItemFontName(j))
                End If
            Next j
    End If
    
    If CInt(Forms(i).Tag) = �r�ε��c�N�X Then
            frm�r�ε��c.tree�r�ξ𪬵��c.FontName = ��ܦr��
            'frm�r�ε��c.tree�r�ξ𪬵��c.ItemFontSize(0) = ��ܦr���j�p
            For j = 0 To frm�r�ε��c.tree�r�ξ𪬵��c.ListCount - 1
                If frm�r�ε��c.tree�r�ξ𪬵��c.ItemFontName(j) <> ��ܦr�� Then
                    frm�r�ε��c.tree�r�ξ𪬵��c.ItemFontName(j) = ������ܦr��(frm�r�ε��c.tree�r�ξ𪬵��c.ItemFontName(j))
                End If
            Next j
    End If
    
    If CInt(Forms(i).Tag) = ����r��N�X Then
            frm����r��.tree�r�ξ𪬵��c.FontName = ��ܦr��
            For j = 0 To frm����r��.tree�r�ξ𪬵��c.ListCount - 1
                If frm����r��.tree�r�ξ𪬵��c.ItemFontName(j) <> ��ܦr�� Then
                    frm����r��.tree�r�ξ𪬵��c.ItemFontName(j) = ������ܦr��(frm����r��.tree�r�ξ𪬵��c.ItemFontName(j))
                End If
            Next j
    End If

    If CInt(Forms(i).Tag) = ����r�ڥN�X Then
            frm����r��.tree�r�ξ𪬵��c.FontName = ��ܦr��
            For j = 0 To frm����r��.tree�r�ξ𪬵��c.ListCount - 1
                If frm����r��.tree�r�ξ𪬵��c.ItemFontName(j) <> ��ܦr�� Then
                    frm����r��.tree�r�ξ𪬵��c.ItemFontName(j) = ������ܦr��(frm����r��.tree�r�ξ𪬵��c.ItemFontName(j))
                End If
            Next j
    End If

    If CInt(Forms(i).Tag) = �r�κt�ܥN�X Then
       For j = 0 To frm�r�κt��.tree�r�ξ𪬵��c.ListCount - 1
           If frm�r�κt��.tree�r�ξ𪬵��c.ItemFontName(j) <> "�_�v�j����p�f" And frm�r�κt��.tree�r�ξ𪬵��c.ItemFontName(j) <> "�_�v�j���孫��" And frm�r�κt��.tree�r�ξ𪬵��c.ItemFontName(j) <> "����|����" And frm�r�κt��.tree�r�ξ𪬵��c.ItemFontName(j) <> "����|�Ұ���" And frm�r�κt��.tree�r�ξ𪬵��c.ItemFontName(j) <> "����|���t²����r" Then
              If frm�r�κt��.tree�r�ξ𪬵��c.ItemCell(j).RTFStyle = 1 Then
                 frm�r�κt��.tree�r�ξ𪬵��c.List(j) = �ഫRTF�ʦr(Right(frm�r�κt��.tree�r�ξ𪬵��c.ItemTag(j), Len(frm�r�κt��.tree�r�ξ𪬵��c.ItemTag(j)) - 1), ��ܦr��)
              Else
                frm�r�κt��.tree�r�ξ𪬵��c.ItemFontName(j) = ������ܦr��(frm�r�κt��.tree�r�ξ𪬵��c.ItemFontName(j))
              End If
           End If
       Next j
    End If

    If CInt(Forms(i).Tag) = �r�ί��ޥN�X Then
        frm�r�ί���.tree�r�ξ𪬵��c.FontName = ��ܦr��
        For j = 0 To frm�r�ί���.tree�r�ξ𪬵��c.ListCount - 1
            If frm�r�ί���.tree�r�ξ𪬵��c.ItemFontName(j) <> ��ܦr�� Then
                frm�r�ί���.tree�r�ξ𪬵��c.ItemFontName(j) = ������ܦr��(frm�r�ί���.tree�r�ξ𪬵��c.ItemFontName(j))
            End If
        Next j
    End If

Next i

End Sub

Private Sub cbo�r���W��_KeyPress(KeyAscii As Integer)

Dim �r���W�� As String

If KeyAscii = 13 Then
    If Len(cbo�r���W��.Text) = 0 Then cbo�r���W��.Text = ��ܦr��
    �r���W�� = cbo�r���W��.Text
    If �r���W�� = "�ө���" Or �r���W�� = "�з���" Then
        cbo�r���W��.Text = �r���W��
        cbo�r���W��_click
    Else
        cbo�r���W��.Text = ��ܦr��
    End If
End If

End Sub

Private Sub cbo�r���W��_LostFocus()

If Len(cbo�r���W��.Text) = 0 Then cbo�r���W��.Text = ��ܦr��

End Sub

Private Sub cbo�r��j�p_LostFocus()

If Len(cbo�r��j�p.Text) = 0 Then cbo�r��j�p.Text = ��ܦr���j�p

End Sub

Private Sub cbo�ѪR��_Click()

Dim �r��j�p As Integer

�r��j�p = Val(cbo�ѪR��.Text)
If �r��j�p > 0 Then �Ϥ��ѪR�� = �r��j�p

End Sub

Private Sub cbo�ѪR��_KeyPress(KeyAscii As Integer)

If Val(cbo�ѪR��.Text) > 0 Then �Ϥ��ѪR�� = Val(cbo�ѪR��.Text)

End Sub

Private Sub cbo�Ϥ��j�p_Click()

Dim �r��j�p As Integer

�r��j�p = Val(cbo�Ϥ��j�p.Text)
If �r��j�p > 0 Then �Ϥ��r���j�p = �r��j�p

End Sub

Private Sub cbo�Ϥ��j�p_KeyPress(KeyAscii As Integer)

Dim �r��j�p As Integer

If KeyAscii = 13 Then
   If Len(cbo�Ϥ��j�p.Text) = 0 Then cbo�Ϥ��j�p.Text = �Ϥ��r���j�p
   �r��j�p = Val(cbo�Ϥ��j�p.Text)
   If �r��j�p >= 8 And �r��j�p <= 1000 Then
      cbo�Ϥ��j�p_Click
   ElseIf �r��j�p < 8 Then
      cbo�Ϥ��j�p.Text = 8
      cbo�Ϥ��j�p_Click
   ElseIf �r��j�p > 1000 Then
      cbo�Ϥ��j�p.Text = 1000
      cbo�Ϥ��j�p_Click
   End If
End If

End Sub

Private Sub MDIForm_Load()

Dim i As Integer
Dim �զr�Ÿ��� As Recordset
Dim ret As Long, lendata As Long, WinPath As String, path As String

���J��l��
��l�e���]�w

If ��lfirst = 1 Then
   ��lleft = mdi�~�r�r��.Left
   ��ltop = mdi�~�r�r��.Top
   ��lheight = mdi�~�r�r��.Height
   ��lwidth = mdi�~�r�r��.Width
End If

'Me.show

�t�Φr�� = "����"
�}�Ҹ�Ʈw

Set �˦r�� = �t�θ�Ʈw.OpenRecordset("�˦r��")
Set �`�βŸ��γ������� = �t�θ�Ʈw.OpenRecordset("�`�βŸ��γ�������")
Set �զr�Ÿ��� = �t�θ�Ʈw.OpenRecordset("�Ÿ�")
Set �d������ = �t�θ�Ʈw.OpenRecordset("�d������")
Set ���峡�� = �t�θ�Ʈw.OpenRecordset("���峡��")
Set ���g�r�� = �t�θ�Ʈw.OpenRecordset("���g�r��")
Set ����r�� = �t�θ�Ʈw.OpenRecordset("����r��")

Set �����˦r�� = �t�θ�Ʈw.OpenRecordset("�˦r��")
Set ���Ѧr�� = �t�θ�Ʈw.OpenRecordset("�r��")
Set ���Ѳ��g�r�� = �t�θ�Ʈw.OpenRecordset("���g�r��")
Set ���Ѳ���r�� = �t�θ�Ʈw.OpenRecordset("����r��")

Set �p�f�˦r�� = �p�f��Ʈw.OpenRecordset("�˦r��")
Set �p�f�W��r = �p�f��Ʈw.OpenRecordset("�r��")
Set �p�f���g�r�� = �p�f��Ʈw.OpenRecordset("���g�r��")
Set �p�f����r�� = �p�f��Ʈw.OpenRecordset("����r��")

Set �����˦r�� = �����Ʈw.OpenRecordset("�˦r��")
Set ����ɿ�� = �����Ʈw.OpenRecordset("�ɿ�")
Set ����r�� = �����Ʈw.OpenRecordset("�r��")
Set ���岧�g�r�� = �����Ʈw.OpenRecordset("���g�r��")
Set ���岧��r�� = �����Ʈw.OpenRecordset("����r��")
Set ���岧�g�r�� = �����Ʈw.OpenRecordset("���g�r��")
Set ���嶰�����W = �����Ʈw.OpenRecordset("�������W")
Set ���嶰���ޱo = �����Ʈw.OpenRecordset("�����ޱo")
Set ������L = �����Ʈw.OpenRecordset("������L")

Set �Ұ����˦r�� = �Ұ����Ʈw.OpenRecordset("�˦r��")
Set �Ұ���r�� = �Ұ����Ʈw.OpenRecordset("�r��")
Set �Ұ��岧�g�r�� = �Ұ����Ʈw.OpenRecordset("���g�r��")
Set �Ұ��岧��r�� = �Ұ����Ʈw.OpenRecordset("����r��")
Set �Ұ��岧�g�r�� = �Ұ����Ʈw.OpenRecordset("���g�r��")

Set ���t��r�˦r�� = ���t��r��Ʈw.OpenRecordset("�˦r��")
Set ���t��r�ɿ�� = ���t��r��Ʈw.OpenRecordset("�ɿ�")
Set ���t��r�r�� = ���t��r��Ʈw.OpenRecordset("�r��")
Set ���t��r���g�r�� = ���t��r��Ʈw.OpenRecordset("���g�r��")
Set ���t��r����r�� = ���t��r��Ʈw.OpenRecordset("����r��")
Set ���t��r���g�r�� = ���t��r��Ʈw.OpenRecordset("���g�r��")

����r��.Index = "�s��"
���g�r��.Index = "�r��"
�˦r��.Index = "�r��"

���Ѳ���r��.Index = "�s��"
���Ѳ��g�r��.Index = "�r��"
�����˦r��.Index = "�r��"

�p�f����r��.Index = "�s��"
�p�f���g�r��.Index = "�r��"
�p�f�˦r��.Index = "�r��"

���岧��r��.Index = "�s��"
���岧�g�r��.Index = "�s��"
���岧�g�r��.Index = "�r��"
�����˦r��.Index = "�r��"
����ɿ��.Index = "���ѽs��"
���嶰���ޱo.Index = "���ѽs��"
���嶰�����W.Index = "����"
������L.Index = "�r�Y"

�Ұ��岧��r��.Index = "�s��"
�Ұ��岧�g�r��.Index = "�s��"
�Ұ��岧�g�r��.Index = "�r��"
�Ұ����˦r��.Index = "�r��"

���t��r����r��.Index = "�s��"
���t��r���g�r��.Index = "�s��"
���t��r���g�r��.Index = "�r��"
���t��r�˦r��.Index = "�r��"
���t��r�ɿ��.Index = "���ѽs��"

�{�Φr�� = "�з���"
��ܦr�� = cbo�r���W��
��ܦr���j�p = cbo�r��j�p

�զr�Ÿ���.MoveFirst
For i = 1 To 14
    �զr�Ÿ��}�C(�զr�Ÿ���.Fields("�s��")) = �զr�Ÿ���.Fields("�r��")
    �զr�Ÿ���.MoveNext
Next i
�զr�Ÿ���.Close

Do Until �`�βŸ��γ�������.EOF
   cbo�Ÿ�����.AddItem �`�βŸ��γ�������.Fields("����")
   cbo�Ÿ�����.ItemData(cbo�Ÿ�����.NewIndex) = �`�βŸ��γ�������.Fields("����")
   �`�βŸ��γ�������.MoveNext
Loop

cbo�Ÿ�����.ListIndex = 0
cbo���e.AddItem "1-99"

For i = 1 To 99
    cbo���e.AddItem i
    cbo���e.ItemData(i) = i
Next i


cbo�ѪR��.List(0) = "72"
cbo�ѪR��.List(1) = "120"
cbo�ѪR��.List(2) = "300"
cbo�ѪR��.List(3) = "450"
cbo�ѪR��.List(4) = "600"
cbo�ѪR��.List(5) = "750"
cbo�ѪR��.List(6) = "900"
cbo�ѪR��.List(7) = "1050"
cbo�ѪR��.List(8) = "1200"
cbo�ѪR��.List(9) = "1800"
cbo�ѪR��.List(10) = "2400"
cbo�ѪR��.Text = �Ϥ��ѪR��

cbo�Ϥ��j�p.List(0) = "8"
cbo�Ϥ��j�p.List(1) = "9"
cbo�Ϥ��j�p.List(2) = "10"
cbo�Ϥ��j�p.List(3) = "11"
cbo�Ϥ��j�p.List(4) = "12"
cbo�Ϥ��j�p.List(5) = "14"
cbo�Ϥ��j�p.List(6) = "16"
cbo�Ϥ��j�p.List(7) = "18"
cbo�Ϥ��j�p.List(8) = "20"
cbo�Ϥ��j�p.List(9) = "22"
cbo�Ϥ��j�p.List(10) = "24"
cbo�Ϥ��j�p.List(11) = "26"
cbo�Ϥ��j�p.List(12) = "28"
cbo�Ϥ��j�p.List(13) = "36"
cbo�Ϥ��j�p.List(14) = "48"
cbo�Ϥ��j�p.List(15) = "72"
cbo�Ϥ��j�p.Text = �Ϥ��r���j�p

�ƻs�Ϥ���Word = False

�]�w�u��C��l���A

�@�ε���(mdi�~�r�r�ΥN�X) = "mdi�~�r�r��"
�@�ε���(�c�r�Ÿ��N�X) = "�c�r�Ÿ�"
�@�ε���(²�|�N�X) = "²�|"
�@�ε���(�K���N�X) = "�K��"
�@�ε���(�ϧΤ�r�N�X) = "�ϧΤ�r"
�@�ε���(�d�������N�X) = "�d���r�峡��"
�@�ε���(���峡���N�X) = "����Ѧr����"
�@�ε���(Big5�r�ڥN�X) = "Big5�r��"
�@�ε���(Big5��²�Ʀr�r�ڥN�X) = "Big5��²�Ʀr�r��"
�@�ε���(�r�ڥN�X) = "�r��"
�@�ε���(�p�f�W��r�N�X) = "�p�f�W��r"
�@�ε���(����r�ڥN�X) = "����r��"
�@�ε���(�Ұ���r�ڥN�X) = "�Ұ���r��"
�@�ε���(���t²����r�r�ڥN�X) = "���t²����r�r��"
�@�ε���(����~�r�N�X) = "����~�r"
�@�ε���(�r�δF�ťN�X) = "�r�δF��"
�@�ε���(�X�B�˦r�N�X) = "�X�B�˦r"
�@�ε���(�r�ε��c�N�X) = "�r�ε��c"
�@�ε���(�r�ί��ޥN�X) = "�r�ί���"
�@�ε���(����r��N�X) = "����r��"
�@�ε���(����r�ڥN�X) = "����r��"

�{�ε��� = "mdi�~�r�r��"
�{�ε����N�X = mdi�~�r�r�ΥN�X
Me.Tag = mdi�~�r�r�ΥN�X

���e�����d�� = True
'���e = 1
���e = 0
���� = 0
�Ұʦr�ε��c = False
�Ұʲ���r�� = False

For i = 0 To 2
    �����N�X(i) = True
Next i

���J����


lendata = 255
path = String(lendata, Chr(0))
ret = GetWindowsDirectory(path, lendata)
WinPath = Left(path, InStr(path, Chr(0)) - 1)
�Ȧs�ؿ� = WinPath & "\Temp\CdphanziTempDir"
If Dir(�Ȧs�ؿ�, vbDirectory) = "" Then MkDir �Ȧs�ؿ�
bmpcount = 0
WordWasNotRunning = True

�w�]�s���Ҧ� = -1
²���s���Ҧ� = True

End Sub

Private Sub cbo�r��j�p_Click()
Dim i As Integer
Dim j As Long

If Len(cbo�r��j�p.Text) = 0 Then cbo�r��j�p.Text = ��ܦr���j�p
If ��ܦr���j�p = cbo�r��j�p.Text Then Exit Sub

��ܦr���j�p = cbo�r��j�p.Text

��e = CInt(��ܦr���j�p) * 20 + CInt(��ܦr���j�p) * 20 / 3

For i = 1 To Forms.Count - 1

    If (CInt(Forms(i).Tag) >= Big5�r�ڥN�X) And (CInt(Forms(i).Tag) <= �c�r�Ÿ��N�X) Then
       frm����d��.tree�r�ξ𪬵��c.FontSize = ��ܦr���j�p
       For j = 0 To frm����d��.tree�r�ξ𪬵��c.ListCount - 1
           frm����d��.tree�r�ξ𪬵��c.ItemFontSize(j) = ��ܦr���j�p
       Next j
    End If
    
    If CInt(Forms(i).Tag) = �r�δF�ťN�X Then
       frm�r�δF��.tree�r�ξ𪬵��c.FontSize = ��ܦr���j�p
       For j = 0 To frm�r�δF��.tree�r�ξ𪬵��c.ListCount - 1
           frm�r�δF��.tree�r�ξ𪬵��c.ItemFontSize(j) = ��ܦr���j�p
       Next j
    End If
    
    If CInt(Forms(i).Tag) = �X�B�˦r�N�X Then
       frm�X�B�˦r.tree�r�ξ𪬵��c.FontSize = ��ܦr���j�p
       For j = 0 To frm�X�B�˦r.tree�r�ξ𪬵��c.ListCount - 1
           frm�X�B�˦r.tree�r�ξ𪬵��c.ItemFontSize(j) = ��ܦr���j�p
       Next j
    End If
    
    If CInt(Forms(i).Tag) = �r�ε��c�N�X Then
       frm�r�ε��c.tree�r�ξ𪬵��c.FontSize = ��ܦr���j�p
       For j = 0 To frm�r�ε��c.tree�r�ξ𪬵��c.ListCount - 1
           frm�r�ε��c.tree�r�ξ𪬵��c.ItemFontSize(j) = ��ܦr���j�p
       Next j
    End If
    
    If CInt(Forms(i).Tag) = ����r��N�X Then
       frm����r��.tree�r�ξ𪬵��c.FontSize = ��ܦr���j�p
       For j = 0 To frm����r��.tree�r�ξ𪬵��c.ListCount - 1
           frm����r��.tree�r�ξ𪬵��c.ItemFontSize(j) = ��ܦr���j�p
       Next j
    End If

    If CInt(Forms(i).Tag) = ����r�ڥN�X Then
       frm����r��.tree�r�ξ𪬵��c.FontSize = ��ܦr���j�p
       For j = 0 To frm����r��.tree�r�ξ𪬵��c.ListCount - 1
           frm����r��.tree�r�ξ𪬵��c.ItemFontSize(j) = ��ܦr���j�p
       Next j
    End If

    If CInt(Forms(i).Tag) = �r�κt�ܥN�X Then
       frm�r�κt��.tree�r�ξ𪬵��c.FontSize = ��ܦr���j�p
       For j = 0 To frm�r�κt��.tree�r�ξ𪬵��c.ListCount - 1
          If Len(frm�r�κt��.tree�r�ξ𪬵��c.List(j)) = 1 Then
            frm�r�κt��.tree�r�ξ𪬵��c.ItemFontSize(j) = ��ܦr���j�p
          End If
       Next j
    End If

    If CInt(Forms(i).Tag) = �r�ί��ޥN�X Then
       frm�r�ί���.tree�r�ξ𪬵��c.FontSize = ��ܦr���j�p
       For j = 0 To frm�r�ί���.tree�r�ξ𪬵��c.ListCount - 1
          If Len(frm�r�ί���.tree�r�ξ𪬵��c.List(j)) = 1 Then
            frm�r�ί���.tree�r�ξ𪬵��c.ItemFontSize(j) = ��ܦr���j�p
          End If
       Next j
    End If

Next i

End Sub


Private Sub cbo�r��j�p_KeyPress(KeyAscii As Integer)

Dim �r��j�p As Integer

If KeyAscii = 13 Then
   If Len(cbo�r��j�p.Text) = 0 Then cbo�r��j�p.Text = ��ܦr���j�p
   �r��j�p = Val(cbo�r��j�p.Text)
   If �r��j�p >= 10 And �r��j�p <= 1000 Then
      cbo�r��j�p_Click
   ElseIf �r��j�p < 10 Then
      cbo�r��j�p.Text = 10
      cbo�r��j�p_Click
   ElseIf �r��j�p > 1000 Then
      cbo�r��j�p.Text = 1000
      cbo�r��j�p_Click
   End If
End If

End Sub

Private Sub cbo�Ÿ�_Click()
Dim �r�� As String, ����� As String
Dim �s�� As Long, �s���Ÿ� As Integer

�r�� = cbo�Ÿ�.List(cbo�Ÿ�.ListIndex)
�s�� = 0

�˦r��.Index = "�r��"
�˦r��.Seek "=", cbo�Ÿ�.List(cbo�Ÿ�.ListIndex)
If Not �˦r��.NoMatch Then
   txt�r��.Text = �r��
   txt���X.Text = �˦r��.Fields("BIG5")
   txt�ܾe�X.Text = �ഫ�^���ܾe(�˦r��.Fields("�ܾe"))
   ����� = �˦r��.Fields("�����")
   If Not IsNull(�˦r��.Fields("�s��")) Then
      �s�� = �˦r��.Fields("�s��")
   Else
      �s�� = 0
   End If
   �s���Ÿ� = �˦r��.Fields("�s���Ÿ�")
   �^���ݩ� "�з���", �r��, �s��
   �^���c�r�� "�з���", �r��, �s��
   If �Ұʦr�ε��c And (�s���Ÿ� <> 9) Then frm�r�ε��c.���J�r�� "�з���", �r��, �s��
   If �Ұʲ���r�� And (�s���Ÿ� <> 9) Then frm����r��.���J�r�� "�з���", �r��, �s��
   If �Ұʦr�κt�� And (�s���Ÿ� <> 9) Then frm�r�κt��.���J�r�� "�з���", �r��, �s��
   If �Ұʲ���r�� And (�s���Ÿ� <> 9) Then frm����r��.���J�r�� "�з���", �r��, �s��
End If

End Sub


Private Sub cbo����_Click()
���� = cbo����.ItemData(cbo����.ListIndex)
If Not ���e�����d�� Then Exit Sub
frm����d��.����d�� ���e, ����

End Sub

Private Sub cbo�Ÿ�����_Click()
Dim �r�Ϊ� As Recordset
Dim SQL���z�� As String

SQL���z�� = "SELECT �s��,�r�� From �`�βŸ��γ��� Where ���� = " & cbo�Ÿ�����.ItemData(cbo�Ÿ�����.ListIndex) & " ORDER BY �s��"
Set �r�Ϊ� = �t�θ�Ʈw.OpenRecordset(SQL���z��)

cbo�Ÿ�.Clear

Do Until �r�Ϊ�.EOF
   cbo�Ÿ�.AddItem �r�Ϊ�.Fields("�r��")
   �r�Ϊ�.MoveNext
Loop
If cbo�Ÿ�.ListCount > 0 Then cbo�Ÿ�.ListIndex = 0
   
End Sub


Private Sub cbo���e_Click()
If cbo���e.ListIndex = -1 Then cbo���e.ListIndex = 0
���e = cbo���e.ItemData(cbo���e.ListIndex)
If Not ���e�����d�� Then Exit Sub
frm����d��.����d�� ���e, ����

End Sub

Private Sub cbo���e_KeyPress(KeyAscii As Integer)

Dim ���e�� As Integer

If KeyAscii = 13 Then
   ���e�� = Val(cbo���e.Text)
   If cbo���e.Text = "1-99" Then
      cbo���e.ListIndex = -1
   ElseIf ���e�� >= 1 And ���e�� <= 99 Then
      cbo���e.ListIndex = ���e��
   End If
End If

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
�p�⵲������
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
���g�r��.Close
�`�βŸ��γ�������.Close
�˦r��.Close
�d������.Close
���峡��.Close
�t�θ�Ʈw.Close
�p�f��Ʈw.Close
�x�s������

End Sub

Private Sub �r��W��()
Dim �Ȧs�� As Recordset
Dim i As Integer
Dim �Ȧs�}�C As Variant

Set �Ȧs�� = �t�θ�Ʈw.OpenRecordset("�r��")
�Ȧs��.MoveFirst

i = 0
Do Until �Ȧs��.EOF
   �r��}�C(�Ȧs��.Fields("�s��")) = �Ȧs��.Fields("�W��")
   i = i + 1
   �Ȧs��.MoveNext
Loop

�Ȧs��.Close

End Sub

Private Sub mnu_Big5_Click()

��l�r�W = 2

If mnu_Big5.Checked = False Then
   mnu_�`�Φr.Checked = False
   mnu_Big5.Checked = True
   mnu_²�Ʀr�`��.Checked = False
   mnu_�~�y�j�r��.Checked = False
   mnu_����.Checked = False
   mnu_�p�f.Checked = False
   mnu_����.Checked = False
   mnu_����ϧΤ�r.Checked = False
   mnu_�Ұ���.Checked = False
   mnu_���t��r.Checked = False
   If �Ұʦr�δF�� Then frm�r�δF��.txt�c�r��.FontName = "�з���"
   �t�Φr�� = "����"
End If

End Sub

Private Sub mnu_Big5�ﶵ_Click()

If mnu_Big5�ﶵ.Checked = True Then
   mnu_Big5�ﶵ.Checked = False
Else
   mnu_Big5�ﶵ.Checked = True
End If

End Sub

Private Sub mnu_cdp_Click()
frm����.show 1
End Sub

Private Sub mnu_edit_�ƻs��Word_Click()

If mnu_�ƻs�r�Ψ�Word.Checked = True Then
    mnu_�ƻs_Click
    WordApp.Selection.Paste
ElseIf mnu_�ƻs�Ϥ���Word.Checked Then
    �ƻs�Ϥ���Word = True
    �ƻs��Word���Ϥ��j�p = WordApp.Selection.font.Size
    mnu_�K�Ϩ�Word_Click
    �ƻs�Ϥ���Word = False
Else
    mnu_�ƻs_Click
    If �ƻsBig5�r�� Then
        WordApp.Selection.Paste
    Else
        �ƻs�Ϥ���Word = True
        �ƻs��Word���Ϥ��j�p = WordApp.Selection.font.Size
        mnu_�K�Ϩ�Word_Click
        �ƻs�Ϥ���Word = False
    End If
End If

End Sub

Private Sub mnu_Tool_Click()

If InStr(1, txt�c�r��, "ơ") > 0 Then
    mnu_Tool_ListLikeChar.Enabled = True
Else
    mnu_Tool_ListLikeChar.Enabled = False
End If

'If mnu_�`�Φr.Checked = True Then
'    mnuToolListChar.Caption = "�C�X�Ҧ��`�Φr..."
'    mnuToolListChar.Enabled = �Ұʦr�δF��
'ElseIf mnu_Big5.Checked = True Then
'    mnuToolListChar.Caption = "�C�X���j�X�Ҧ��r��..."
'    mnuToolListChar.Enabled = �Ұʦr�δF��
'ElseIf mnu_²�Ʀr�`��.Checked = True Then
'    mnuToolListChar.Caption = "�C�X�Ҧ�²�Ʀr..."
'    mnuToolListChar.Enabled = �Ұʦr�δF��
'ElseIf mnu_�~�y�j�r��.Checked = True Then
'    mnuToolListChar.Caption = "�C�X�m�~�y�j�r��n�Ҧ��r��..."
'    mnuToolListChar.Enabled = False
'ElseIf mnu_����.Checked = True Then
'    mnuToolListChar.Caption = "�C�X�Ҧ����Ѧr��..."
'    mnuToolListChar.Enabled = False
'ElseIf mnu_�p�f.Checked = True Then
'    mnuToolListChar.Caption = "�C�X�m����Ѧr�n�Ҧ��r��..."
'    mnuToolListChar.Enabled = �Ұʦr�δF��
'ElseIf mnu_����.Checked = True Then
'    mnuToolListChar.Caption = "�C�X�m����s�n�Ҧ��r��..."
'    mnuToolListChar.Enabled = �Ұʦr�δF��
'ElseIf mnu_����ϧΤ�r.Checked = True Then
'    mnuToolListChar.Caption = "�C�X�m����s�n�����W�Ҧ��r��..."
'    mnuToolListChar.Enabled = �Ұʦr�δF��
'ElseIf mnu_�Ұ���.Checked = True Then
'    mnuToolListChar.Caption = "�C�X�m��V�Ұ�������ġ�n�Ҧ��r��..."
'    mnuToolListChar.Enabled = �Ұʦr�δF��
'ElseIf mnu_���t��r.Checked = True Then
'    mnuToolListChar.Caption = "�C�X�m���t²����r�s�n�Ҧ��r��..."
'    mnuToolListChar.Enabled = �Ұʦr�δF��
'End If

End Sub

Private Sub mnu_Tool_ListLikeChar_Click()

Dim tagFont As Integer, tagDuplicate As Integer, tagEnd As Integer

If Not �ҰʥX�B�˦r Then frm�X�B�˦r.show
frm�X�B�˦r.SetFocus

tagFont = InStr(1, txt�c�r��, "ơ")
tagDuplicate = InStr(1, txt�c�r��, ";")
tagEnd = InStr(1, txt�c�r��, "��")

If tagDuplicate = 0 Then
   frm�X�B�˦r.cbo�X�B = Mid(txt�c�r��, tagFont + 1, tagEnd - tagFont - 1)
Else
   frm�X�B�˦r.cbo�X�B = Mid(txt�c�r��, tagFont + 1, tagDuplicate - tagFont - 1)
End If

frm�X�B�˦r.cbo�X�B_KeyPress vbKeyReturn

End Sub

Private Sub mnu_Unicode�ﶵ_Click()

If mnu_Unicode�ﶵ.Checked = True Then
   mnu_Unicode�ﶵ.Checked = False
Else
   mnu_Unicode�ﶵ.Checked = True
End If

End Sub

Private Sub mnu_�K��_Click()

�@�ε����N�X = �K���N�X
frm����d��.Form_Load
frm����d��.show
frm����d��.SetFocus

End Sub

Private Sub mnu_�p�f_Click()

If mnu_�p�f.Checked = False Then
   mnu_�`�Φr.Checked = False
   mnu_Big5.Checked = False
   mnu_²�Ʀr�`��.Checked = False
   mnu_�~�y�j�r��.Checked = False
   mnu_����.Checked = False
   mnu_�p�f.Checked = True
   mnu_����.Checked = False
   mnu_����ϧΤ�r.Checked = False
   mnu_�Ұ���.Checked = False
   mnu_���t��r.Checked = False
   If �Ұʦr�δF�� Then frm�r�δF��.txt�c�r��.FontName = "�_�v�j����p�f"
   mnu_����.Checked = False
   �t�Φr�� = "�p�f"
End If

End Sub

Private Sub mnu_�p�f�ﶵ_Click()

If mnu_�p�f�ﶵ.Checked = True Then
   mnu_�p�f�ﶵ.Checked = False
Else
   mnu_�p�f�ﶵ.Checked = True
End If

End Sub

Private Sub mnu_����j���ﶵ_Click()

If mnu_����j���ﶵ.Checked = True Then
   mnu_����j���ﶵ.Checked = False
Else
   mnu_����j���ﶵ.Checked = True
End If

End Sub

Private Sub mnu_���ػ���Ѧr�ﶵ_Click()

If mnu_���ػ���Ѧr�ﶵ.Checked = True Then
   mnu_���ػ���Ѧr�ﶵ.Checked = False
Else
   mnu_���ػ���Ѧr�ﶵ.Checked = True
End If

End Sub

Private Sub mnu_�X�B�˦r_Click()

�@�ε����N�X = �X�B�˦r�N�X
frm�X�B�˦r.show
frm�X�B�˦r.SetFocus

End Sub

Private Sub mnu_�Ұ���_Click()

If mnu_�Ұ���.Checked = False Then
   mnu_�`�Φr.Checked = False
   mnu_Big5.Checked = False
   mnu_²�Ʀr�`��.Checked = False
   mnu_�~�y�j�r��.Checked = False
   mnu_����.Checked = False
   mnu_�p�f.Checked = False
   mnu_����.Checked = False
   mnu_����ϧΤ�r.Checked = False
   mnu_�Ұ���.Checked = True
   mnu_���t��r.Checked = False
   If �Ұʦr�δF�� Then frm�r�δF��.txt�c�r��.FontName = "����|�Ұ���"
   mnu_����.Checked = False
   �t�Φr�� = "�Ұ���"
End If

End Sub

Private Sub mnu_����ϧΤ�r_Click()

If mnu_����ϧΤ�r.Checked = False Then
   mnu_�`�Φr.Checked = False
   mnu_Big5.Checked = False
   mnu_²�Ʀr�`��.Checked = False
   mnu_�~�y�j�r��.Checked = False
   mnu_����.Checked = False
   mnu_�p�f.Checked = False
   mnu_����.Checked = False
   mnu_����ϧΤ�r.Checked = True
   mnu_�Ұ���.Checked = False
   mnu_���t��r.Checked = False
   If �Ұʦr�δF�� Then frm�r�δF��.txt�c�r��.FontName = "����|����"
   mnu_����.Checked = False
   �t�Φr�� = "����"
End If

End Sub

Private Sub mnu_���z���ƻs��Word_Click()

If mnu_���z���ƻs��Word.Checked = False Then
   mnu_�ƻs�r�Ψ�Word.Checked = False
   mnu_�ƻs�Ϥ���Word.Checked = False
   mnu_���z���ƻs��Word.Checked = True
End If

End Sub

Private Sub mnu_�K�Ϩ�Word_Click()

On Error GoTo ExitSub

mnu_�ƻs�Ϥ�_Click
If Len(�Ȧs����) > 0 Then
    WordApp.ActiveDocument.InlineShapes.AddPicture �Ȧs����, False, True, WordApp.Selection.Range
     WordApp.Selection.MoveRight
'                ����c�r��.SetRange Start:=����c�r��.End, End:=����c�r��.End + 1
'                ����c�r��.font = �r��
    WordApp.ActiveDocument.InlineShapes(WordApp.ActiveDocument.InlineShapes.Count).AlternativeText = "����" & ���N��r
    WordApp.Selection.Paragraphs.BaseLineAlignment = wdBaselineAlignCenter
End If

ExitSub:

End Sub

Private Sub mnu_���t��r_Click()

If mnu_���t��r.Checked = False Then
   mnu_�`�Φr.Checked = False
   mnu_Big5.Checked = False
   mnu_²�Ʀr�`��.Checked = False
   mnu_�~�y�j�r��.Checked = False
   mnu_����.Checked = False
   mnu_�p�f.Checked = False
   mnu_����.Checked = False
   mnu_����ϧΤ�r.Checked = False
   mnu_�Ұ���.Checked = False
   mnu_���t��r.Checked = True
   If �Ұʦr�δF�� Then frm�r�δF��.txt�c�r��.FontName = "����|���t²����r"
   mnu_����.Checked = False
   �t�Φr�� = "���t��r"
End If

End Sub

Private Sub mnu_�Ұ���r���L�ﶵ_Click()

If mnu_�Ұ���r���L�ﶵ.Checked = True Then
   mnu_�Ұ���r���L�ﶵ.Checked = False
Else
   mnu_�Ұ���r���L�ﶵ.Checked = True
End If

End Sub

Private Sub mnu_�Ұ���r�����ﶵ_Click()

If mnu_�Ұ���r�����ﶵ.Checked = True Then
   mnu_�Ұ���r�����ﶵ.Checked = False
Else
   mnu_�Ұ���r�����ﶵ.Checked = True
End If

End Sub


Private Sub mnu_���t��r�X�B�ﶵ_Click()

If mnu_���t��r�X�B�ﶵ.Checked = True Then
   mnu_���t��r�X�B�ﶵ.Checked = False
Else
   mnu_���t��r�X�B�ﶵ.Checked = True
End If

End Sub

Private Sub mnu_���t²����r�s�ﶵ_Click()

If mnu_���t²����r�s�ﶵ.Checked = True Then
   mnu_���t²����r�s�ﶵ.Checked = False
Else
   mnu_���t²����r�s�ﶵ.Checked = True
End If

End Sub

Private Sub mnu_�Ұ���ﶵ_Click()

If mnu_�Ұ���ﶵ.Checked = True Then
   mnu_�Ұ���ﶵ.Checked = False
Else
   mnu_�Ұ���ﶵ.Checked = True
End If

End Sub

Private Sub mnu_���t��r�ﶵ_Click()

If mnu_���t��r�ﶵ.Checked = True Then
   mnu_���t��r�ﶵ.Checked = False
Else
   mnu_���t��r�ﶵ.Checked = True
End If

End Sub

Private Sub mnu_�Ұ�������ġ�ﶵ_Click()

If mnu_�Ұ�������ġ�ﶵ.Checked = True Then
   mnu_�Ұ�������ġ�ﶵ.Checked = False
Else
   mnu_�Ұ�������ġ�ﶵ.Checked = True
End If

End Sub

Private Sub mnu_�r�ί���_Click()

�@�ε����N�X = �r�ί��ޥN�X
frm�r�ί���.show
frm�r�ί���.SetFocus

End Sub

Private Sub mnu_�r�δF�ť]�t���g����_Click()

If mnu_�r�δF�ť]�t���g����.Checked = True Then
   mnu_�r�δF�ť]�t���g����.Checked = False
Else
   mnu_�r�δF�ť]�t���g����.Checked = True
End If

End Sub

Private Sub mnu_�r�δF�ťu�C�X�`�Φr_Click()

��l�r�W = 5

'If mnu_�r�δF�ťu�C�X�`�Φr.Checked = False Then
   'mnu_�r�δF�ťu�C�X�`�Φr.Checked = True
   'mnu_�r�δF�ťu�C�X�q���Φr.Checked = False
   'mnu_�r�δF�ŦC�X�Ҧ��r��.Checked = False
'End If

End Sub


Private Sub mnu_�r�δF�ťu�C�X�q���Φr_Click()

��l�r�W = 2

'If mnu_�r�δF�ťu�C�X�q���Φr.Checked = False Then
   'mnu_�r�δF�ťu�C�X�q���Φr.Checked = True
   'mnu_�r�δF�ťu�C�X�`�Φr.Checked = False
   'mnu_�r�δF�ŦC�X�Ҧ��r��.Checked = False
'End If

End Sub

Private Sub mnu_�r�δF�ŦC�X�Ҧ��r��_Click()

��l�r�W = 1

'If mnu_�r�δF�ŦC�X�Ҧ��r��.Checked = False Then
   'mnu_�r�δF�ŦC�X�Ҧ��r��.Checked = True
   'mnu_�r�δF�ťu�C�X�`�Φr.Checked = False
   'mnu_�r�δF�ťu�C�X�q���Φr.Checked = False
'End If

End Sub

Private Sub mnu_�r�δF�űĥ�SQL�y�k_Click()

If mnu_�r�δF�űĥ�SQL�y�k.Checked = True Then
   mnu_�r�δF�űĥ�SQL�y�k.Checked = False
   frm�r�δF��.Caption = "�����˦r"
Else
   mnu_�r�δF�űĥ�SQL�y�k.Checked = True
   frm�r�δF��.Caption = "�����˦r(SQL Like)"
End If

End Sub

Private Sub mnu_�r�δF�ųv�ŦC�X��@����_Click()

��l�v�ŦC�X = 1

If mnu_�r�δF�ųv�ŦC�X��@����.Checked = False Then
   mnu_�r�δF�ųv�ŦC�X��@����.Checked = True
   mnu_�r�δF�ťu�C�r�Τ��C����.Checked = False
End If

End Sub


Private Sub mnu_�r�δF�ťu�C�r�Τ��C����_Click()

��l�v�ŦC�X = 0

If mnu_�r�δF�ťu�C�r�Τ��C����.Checked = False Then
   mnu_�r�δF�ųv�ŦC�X��@����.Checked = False
   mnu_�r�δF�ťu�C�r�Τ��C����.Checked = True
End If

End Sub


Private Sub mnu_�r�δF�ſ�ӿ�J���󶶧�_Click()

If mnu_�r�δF�ſ�ӿ�J���󶶧�.Checked = True Then
   mnu_�r�δF�ſ�ӿ�J���󶶧�.Checked = False
Else
   mnu_�r�δF�ſ�ӿ�J���󶶧�.Checked = True
End If

End Sub

Private Sub mnu_�r�κt��_Click()

'mnu_�r�κt��.Enabled = False
�@�ε����N�X = �r�κt�ܥN�X
frm�r�κt��.show
frm�r�κt��.SetFocus

End Sub

Private Sub mnu_�r��_Click()

frm�r��.show 1
cbo�r���W��_click
cbo�r��j�p_Click

End Sub

Private Sub mnu_����_Click()

If mnu_����.Checked = False Then
   mnu_�`�Φr.Checked = False
   mnu_Big5.Checked = False
   mnu_²�Ʀr�`��.Checked = False
   mnu_�~�y�j�r��.Checked = False
   mnu_����.Checked = False
   mnu_�p�f.Checked = False
   mnu_����.Checked = True
   mnu_����ϧΤ�r.Checked = False
   mnu_�Ұ���.Checked = False
   mnu_���t��r.Checked = False
   If �Ұʦr�δF�� Then frm�r�δF��.txt�c�r��.FontName = "����|����"
   mnu_����.Checked = False
   �t�Φr�� = "����"
End If

End Sub

Private Sub mnu_������L�ﶵ_Click()

If mnu_������L�ﶵ.Checked = True Then
   mnu_������L�ﶵ.Checked = False
Else
   mnu_������L�ﶵ.Checked = True
End If

End Sub

Private Sub mnu_����s�ﶵ_Click()

If mnu_����s�ﶵ.Checked = True Then
   mnu_����s�ﶵ.Checked = False
Else
   mnu_����s�ﶵ.Checked = True
End If

End Sub

Private Sub mnu_����ﶵ_Click()

If mnu_����ﶵ.Checked = True Then
   mnu_����ﶵ.Checked = False
Else
   mnu_����ﶵ.Checked = True
End If

End Sub

Private Sub mnu_�ا��~�y�j�r��ﶵ_Click()

If mnu_�ا��~�y�j�r��ﶵ.Checked = True Then
   mnu_�ا��~�y�j�r��ﶵ.Checked = False
Else
   mnu_�ا��~�y�j�r��ﶵ.Checked = True
End If

End Sub

Private Sub mnu_��P���嶰�������ﶵ_Click()

If mnu_��P���嶰�������ﶵ.Checked = True Then
   mnu_��P���嶰�������ﶵ.Checked = False
Else
   mnu_��P���嶰�������ﶵ.Checked = True
End If

End Sub

Private Sub mnu_��P���嶰���ޱo�ﶵ_Click()

If mnu_��P���嶰���ޱo�ﶵ.Checked = True Then
   mnu_��P���嶰���ޱo�ﶵ.Checked = False
Else
   mnu_��P���嶰���ޱo�ﶵ.Checked = True
End If

End Sub

Private Sub mnu_�`�Φr_Click()

��l�r�W = 5

If mnu_�`�Φr.Checked = False Then
   mnu_�`�Φr.Checked = True
   mnu_Big5.Checked = False
   mnu_²�Ʀr�`��.Checked = False
   mnu_�~�y�j�r��.Checked = False
   mnu_����.Checked = False
   mnu_�p�f.Checked = False
   mnu_����.Checked = False
   mnu_����ϧΤ�r.Checked = False
   mnu_�Ұ���.Checked = False
   mnu_���t��r.Checked = False
   If �Ұʦr�δF�� Then frm�r�δF��.txt�c�r��.FontName = "�з���"
   �t�Φr�� = "����"
End If

End Sub

Private Sub mnu_�d���r�峡��_Click()

�@�ε����N�X = �d�������N�X
frm����d��.Form_Load
frm����d��.show
frm����d��.SetFocus

End Sub

Private Sub mnu_����r��_Click()
'mnu_����r��.Enabled = False
�@�ε����N�X = ����r�ڥN�X
frm����r��.show
frm����r��.SetFocus

End Sub

Private Sub mnu_�Ϥ�_Click()

frm�Ϥ��]�w.show 1

End Sub

Private Sub mnu_�c�r�Ÿ�_Click()

�@�ε����N�X = �c�r�Ÿ��N�X
frm����d��.Form_Load
frm����d��.show
frm����d��.SetFocus

End Sub

Private Sub mnu_�ϧΤ�r_Click()

�@�ε����N�X = �ϧΤ�r�N�X
frm����d��.Form_Load
frm����d��.show
frm����d��.SetFocus

End Sub


Private Sub mnu_����N�X_Click(Index As Integer)
Dim i As Integer

'For i = �Ÿ��N�X To ����~�r�N�X
'    mnu_����N�X(i).Checked = False
'Next i

'mnu_����N�X(Index).Checked = True
�@�ε����N�X = Index
frm����d��.Form_Load
frm����d��.show
frm����d��.SetFocus

End Sub

Private Sub mnu_����_Click()
'�p�⵲������
Unload mdi�~�r�r��
End

End Sub

Private Sub mnu_�r�δF��_Click()
'mnu_�r�δF��.Enabled = False
�@�ε����N�X = �r�δF�ťN�X
frm�r�δF��.show
frm�r�δF��.SetFocus

End Sub

Private Sub mnu_�r�ε��c_Click()
'mnu_�r�ε��c.Enabled = False
�@�ε����N�X = �r�ε��c�N�X
frm�r�ε��c.show
frm�r�ε��c.SetFocus

End Sub

Private Sub mnu_����r��_Click()
'mnu_����r��.Enabled = False
�@�ε����N�X = ����r��N�X
frm����r��.show
frm����r��.SetFocus

End Sub


Private Sub mnu_�����ñ�_Click()
mdi�~�r�r��.Arrange 1
End Sub

Private Sub mnu_�����ñ�_Click()
'frm����r��.SetFocus
'frm�r�ε��c.SetFocus
'frm�r�δF��.SetFocus

mdi�~�r�r��.Arrange 2
End Sub

Private Sub mnu_���|���_Click()
'frm����r��.SetFocus
'frm�r�ε��c.SetFocus
'frm�r�δF��.SetFocus

mdi�~�r�r��.Arrange 0
End Sub

Private Sub mnu_�ƦC�ϥ�_Click()
mdi�~�r�r��.Arrange 3
End Sub

Private Sub mnu_�̨��s��_click()

frm�w�]�s��.show 1

If Not ���ܹw�]�s�� Then Exit Sub

Select Case �w�]�s���Ҧ�
    Case 1: �w�]�s���@
    Case 2: �w�]�s���G
    Case 3: �w�]�s���T
    Case 4: �w�]�s���|
End Select

End Sub

Private Sub mnu_����_Click()

��l�r�W = 1

If mnu_����.Checked = False Then
   mnu_�`�Φr.Checked = False
   mnu_Big5.Checked = False
   mnu_²�Ʀr�`��.Checked = False
   mnu_�~�y�j�r��.Checked = False
   mnu_����.Checked = True
   mnu_�p�f.Checked = False
   mnu_����.Checked = False
   mnu_����ϧΤ�r.Checked = False
   mnu_�Ұ���.Checked = False
   mnu_���t��r.Checked = False
   If �Ұʦr�δF�� Then frm�r�δF��.txt�c�r��.FontName = "�з���"
   �t�Φr�� = "����"
End If

End Sub

Private Sub mnu_�~�y�j�r��_Click()

��l�r�W = 1

If mnu_�~�y�j�r��.Checked = False Then
   mnu_�`�Φr.Checked = False
   mnu_Big5.Checked = False
   mnu_²�Ʀr�`��.Checked = False
   mnu_�~�y�j�r��.Checked = True
   mnu_����.Checked = False
   mnu_�p�f.Checked = False
   mnu_����.Checked = False
   mnu_����ϧΤ�r.Checked = False
   mnu_�Ұ���.Checked = False
   mnu_���t��r.Checked = False
   If �Ұʦr�δF�� Then frm�r�δF��.txt�c�r��.FontName = "�з���"
   �t�Φr�� = "����"
End If

End Sub

Private Sub mnu_����Ѧr����_Click()

�@�ε����N�X = ���峡���N�X
frm����d��.Form_Load
frm����d��.show
frm����d��.SetFocus

End Sub

Private Sub mnu_����Ѧr���L�ﶵ_Click()

If mnu_����Ѧr���L�ﶵ.Checked = True Then
   mnu_����Ѧr���L�ﶵ.Checked = False
Else
   mnu_����Ѧr���L�ﶵ.Checked = True
End If

End Sub

Private Sub mnu_����ѧζ���_Click()

If mnu_����ѧζ���.Checked = False Then
   mnu_����ѧζ���.Checked = True
   ��l�ѧΦC�X = 1
Else
   mnu_����ѧζ���.Checked = False
   ��l�ѧΦC�X = 0
End If

End Sub

Private Sub mnu_���F�~�y�j�r��ﶵ_Click()

If mnu_���F�~�y�j�r��ﶵ.Checked = True Then
   mnu_���F�~�y�j�r��ﶵ.Checked = False
Else
   mnu_���F�~�y�j�r��ﶵ.Checked = True
End If

End Sub

Private Sub mnu_�ƻs_Click()

Dim �𪬵��c As TList, �`�I�аO As String, �`�I���O As Integer, CopyText As String

CopyText = ""
�`�I���O = -1
�ƻsBig5�r�� = False

Select Case �{�ε����N�X

Case mdi�~�r�r�ΥN�X

    Select Case �{�α���N�X
    
    Case mdi�~�r�r��_�s�����
        CopyText = txt�s��.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_�~�r�����
        CopyText = txt�~�r��.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_�r�Τ��
        CopyText = txt�r��.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_������
        CopyText = txt����.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_�j�~�r���
        CopyText = txt�j�~�r.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_�`���e���
        CopyText = txt�`���e.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_�������
        CopyText = txt����.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_�����������e���
        CopyText = txt�����������e.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_�`�����
        CopyText = txt�`��.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_���X���
        CopyText = txt���X.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_�ܾe�X���
        CopyText = txt�ܾe�X.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_�c�r�����
        CopyText = txt�c�r��.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_�U�Ƥ��
        CopyText = txt�U��.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_�զr�r�Ƥ��
        CopyText = txt�զr�r��.SelText
        GoTo CopyBegin
    
    Case mdi�~�r�r��_�զr�r�Ƨt���g���
        CopyText = txt�զr�r�Ƨt���g.SelText
        GoTo CopyBegin
    
    Case Else
        Exit Sub
        
    End Select

Case Big5�r�ڥN�X To �c�r�Ÿ��N�X
    If frm����d��.tree�r�ξ𪬵��c.ListIndex > -1 Then
        Set �𪬵��c = frm����d��.tree�r�ξ𪬵��c
    Else
        Exit Sub
    End If


Case �r�δF�ťN�X
    If �{�α���N�X = �r�δF��_�˦r��� Then
        CopyText = frm�r�δF��.txt�c�r��.SelText
        GoTo CopyBegin
    ElseIf �{�α���N�X = �r�δF��_�𪬵��c Then
        If frm�r�δF��.tree�r�ξ𪬵��c.ListIndex > -1 Then
            Set �𪬵��c = frm�r�δF��.tree�r�ξ𪬵��c
        End If
    Else
        Exit Sub
    End If
    
Case �X�B�˦r�N�X
    If �{�α���N�X = �X�B�˦r_�˦r��� Then
        CopyText = frm�X�B�˦r.cbo�X�B.SelText
        GoTo CopyBegin
    ElseIf �{�α���N�X = �X�B�˦r_�𪬵��c Then
        If frm�X�B�˦r.tree�r�ξ𪬵��c.ListIndex > -1 Then
            Set �𪬵��c = frm�X�B�˦r.tree�r�ξ𪬵��c
        End If
    Else
        Exit Sub
    End If
    
Case �r�ε��c�N�X
    If frm�r�ε��c.tree�r�ξ𪬵��c.ListIndex > -1 Then
        Set �𪬵��c = frm�r�ε��c.tree�r�ξ𪬵��c
    Else
        Exit Sub
    End If
    
Case �r�κt�ܥN�X
    If frm�r�κt��.tree�r�ξ𪬵��c.ListIndex > -1 Then
        Set �𪬵��c = frm�r�κt��.tree�r�ξ𪬵��c
    Else
        Exit Sub
    End If
    
Case ����r��N�X
    If frm����r��.tree�r�ξ𪬵��c.ListIndex > -1 Then
        Set �𪬵��c = frm����r��.tree�r�ξ𪬵��c
    Else
        Exit Sub
    End If
    
Case ����r�ڥN�X
    If frm����r��.tree�r�ξ𪬵��c.ListIndex > -1 Then
        Set �𪬵��c = frm����r��.tree�r�ξ𪬵��c
    Else
        Exit Sub
    End If
    
Case �r�ί��ޥN�X
    If frm�r�ί���.tree�r�ξ𪬵��c.ListIndex > -1 Then
        Set �𪬵��c = frm�r�ί���.tree�r�ξ𪬵��c
    Else
        Exit Sub
    End If
    
Case Else
    Exit Sub
End Select

�`�I�аO = �𪬵��c.ItemTag(�𪬵��c.ListIndex)
If Len(�`�I�аO) > 0 Then �`�I���O = CInt(Left(�`�I�аO, 1))

Select Case �`�I���O

Case �r�θ`�I�аO
    If �ƻs����X Then
        CopyText = txt�c�r��
    Else
        CopyText = txt�r��
    End If
Case �c�r���`�I�аO
    CopyText = txt�c�r��
Case ���W�`�I�аO
    CopyText = Right(�`�I�аO, Len(�`�I�аO) - 1)
Case ��L�`�I�аO
    CopyText = �𪬵��c.List(�𪬵��c.ListIndex)

End Select

On Error GoTo CopyErr
CopyBegin:

If Len(CopyText) > 0 Then
    Clipboard.Clear
    If Len(CopyText) = 1 Then
        �����˦r��.Index = "�r��"
        �����˦r��.Seek "=", CopyText
        If Not �����˦r��.NoMatch Then
            If �����˦r��.Fields("�s��") <= 13060 Then �ƻsBig5�r�� = True
        End If
    End If
    Clipboard.SetText CopyText
End If

CopyErr:

End Sub

Private Sub mnu_�ƻsUnicode�r�Ψ�Word_Click()

If mnu_�ƻsUnicode�r�Ψ�Word.Checked = True Then
   mnu_�ƻsUnicode�r�Ψ�Word.Checked = False
Else
   mnu_�ƻsUnicode�r�Ψ�Word.Checked = True
End If

End Sub

Private Sub mnu_�ƻs�r�Ψ�Word_Click()

If mnu_�ƻs�r�Ψ�Word.Checked = False Then
   mnu_�ƻs�r�Ψ�Word.Checked = True
   mnu_�ƻs�Ϥ���Word.Checked = False
   mnu_���z���ƻs��Word.Checked = False
End If

End Sub


Private Sub mnu_�ƻs�S��Ϥ�_Click()

Dim �r�� As CDPFONT, �r�� As String, Success As Boolean

�r��.Name = "cdp000"
�r��.Size = 9
�r�� = "�X"
���N��r = "��"

bmpcount = bmpcount + 1
�Ȧs���� = �Ȧs�ؿ� & "\" & "~hz" & CStr(bmpcount) & ".bmp"
�r���ন�Ϥ� �r��, �r��, �Ȧs����, �Ϥ��ѪR��, Success

WordApp.ActiveDocument.InlineShapes.AddPicture �Ȧs����, False, True, WordApp.Selection.Range
WordApp.Selection.MoveRight
WordApp.ActiveDocument.InlineShapes(WordApp.ActiveDocument.InlineShapes.Count).AlternativeText = "����" & ���N��r
WordApp.Selection.Paragraphs.BaseLineAlignment = wdBaselineAlignCenter
    
End Sub

Private Sub mnu_�ƻs�Ϥ�_Click()

Dim �𪬵��c As TList, �r�� As CDPFONT, �r�� As String, Success As Boolean

Select Case �{�ε����N�X

Case Big5�r�ڥN�X To �c�r�Ÿ��N�X
    If frm����d��.tree�r�ξ𪬵��c.ListIndex > -1 Then
        Set �𪬵��c = frm����d��.tree�r�ξ𪬵��c
    Else
        Exit Sub
    End If

Case �r�δF�ťN�X
    If �{�α���N�X = �r�δF��_�˦r��� Then
        GoTo CopyGlyphErr
    ElseIf �{�α���N�X = �r�δF��_�𪬵��c Then
        If frm�r�δF��.tree�r�ξ𪬵��c.ListIndex > -1 Then
            Set �𪬵��c = frm�r�δF��.tree�r�ξ𪬵��c
        End If
    Else
        Exit Sub
    End If
    
Case �X�B�˦r�N�X
    If �{�α���N�X = �X�B�˦r_�˦r��� Then
        GoTo CopyGlyphErr
    ElseIf �{�α���N�X = �X�B�˦r_�𪬵��c Then
        If frm�X�B�˦r.tree�r�ξ𪬵��c.ListIndex > -1 Then
            Set �𪬵��c = frm�X�B�˦r.tree�r�ξ𪬵��c
        End If
    Else
        Exit Sub
    End If
    
Case �r�ε��c�N�X
    If frm�r�ε��c.tree�r�ξ𪬵��c.ListIndex > -1 Then
        Set �𪬵��c = frm�r�ε��c.tree�r�ξ𪬵��c
    Else
        Exit Sub
    End If
    
Case �r�κt�ܥN�X
    If frm�r�κt��.tree�r�ξ𪬵��c.ListIndex > -1 Then
        Set �𪬵��c = frm�r�κt��.tree�r�ξ𪬵��c
    Else
        Exit Sub
    End If
    
Case ����r��N�X
    If frm����r��.tree�r�ξ𪬵��c.ListIndex > -1 Then
        Set �𪬵��c = frm����r��.tree�r�ξ𪬵��c
    Else
        Exit Sub
    End If
    
Case ����r�ڥN�X
    If frm����r��.tree�r�ξ𪬵��c.ListIndex > -1 Then
        Set �𪬵��c = frm����r��.tree�r�ξ𪬵��c
    Else
        Exit Sub
    End If
    
Case �r�ί��ޥN�X
    If frm�r�ί���.tree�r�ξ𪬵��c.ListIndex > -1 Then
        Set �𪬵��c = frm�r�ί���.tree�r�ξ𪬵��c
    Else
        Exit Sub
    End If
    
Case Else
    Exit Sub
End Select


�r��.Name = �𪬵��c.ItemFontName(�𪬵��c.ListIndex)
�r��.Size = �Ϥ��r���j�p
If �ƻs�Ϥ���Word And �ƻs��Word���Ϥ��j�p > 0 Then
    �r��.Size = �ƻs��Word���Ϥ��j�p
End If

�r��.Bold = �𪬵��c.ItemFontBold(�𪬵��c.ListIndex)
�r��.Italic = �𪬵��c.ItemFontItalic(�𪬵��c.ListIndex)
�r��.Underline = �𪬵��c.ItemFontUnder(�𪬵��c.ListIndex)
�r��.StrikeThrough = �𪬵��c.ItemFontStrike(�𪬵��c.ListIndex)
�r��.color = 0

�r�� = �𪬵��c.List(�𪬵��c.ListIndex)

If Len(�r��) <> 1 Then GoTo CopyGlyphErr

bmpcount = bmpcount + 1
�Ȧs���� = �Ȧs�ؿ� & "\" & "~hz" & CStr(bmpcount) & ".bmp"
�r���ন�Ϥ� �r��, �r��, �Ȧs����, �Ϥ��ѪR��, Success

���N��r = Clipboard.GetText

Clipboard.Clear
Clipboard.SetData LoadPicture(�Ȧs����), vbCFBitmap

On Error GoTo CopyGlyphErr

CopyGlyphErr:

End Sub

Private Sub mnu_�ƻs�Ϥ���Word_Click()

If mnu_�ƻs�Ϥ���Word.Checked = False Then
   mnu_�ƻs�r�Ψ�Word.Checked = False
   mnu_�ƻs�Ϥ���Word.Checked = True
   mnu_���z���ƻs��Word.Checked = False
End If

End Sub

Private Sub mnu_�ﶵ_Click()

If �Ұʦr�δF�� Then
    If mnu_�r�δF�űĥ�SQL�y�k.Checked = True Then
        mnu_�r�δF�ųv�ŦC�X��@����.Enabled = False
        mnu_�r�δF�ťu�C�r�Τ��C����.Enabled = False
    Else
        mnu_�r�δF�ųv�ŦC�X��@����.Enabled = True
        mnu_�r�δF�ťu�C�r�Τ��C����.Enabled = True
    End If
    If �t�Φr�� = "����" Then
        mnu_�r�δF�ť]�t���g����.Enabled = True
    Else
        mnu_�r�δF�ť]�t���g����.Enabled = False
    End If
    mnu_�r�δF�űĥ�SQL�y�k.Enabled = True
Else
    mnu_�r�δF�ųv�ŦC�X��@����.Enabled = False
    mnu_�r�δF�ťu�C�r�Τ��C����.Enabled = False
    mnu_�r�δF�ť]�t���g����.Enabled = False
    mnu_�r�δF�űĥ�SQL�y�k.Enabled = False
End If

End Sub

Private Sub mnu_�x�s�����]�w_Click()

If mnu_�x�s�����]�w.Checked = False Then
   mnu_�x�s�����]�w.Checked = True
Else
   mnu_�x�s�����]�w.Checked = False
End If

End Sub

Private Sub mnu_�����D�D_Click()
Dim istring As String, iret As Integer

Screen.MousePointer = ccHourglass
istring = "winhlp32.exe " & App.path & "\cdphanzi.hlp"
Shell istring, 1
Screen.MousePointer = ccDefault

End Sub

Private Sub mnu_�ƻs��ŶKï_Click()

If mnu_�ƻs��ŶKï.Checked = False Then
   mnu_�ƻs��ŶKï.Checked = True
Else
   mnu_�ƻs��ŶKï.Checked = False
End If

End Sub

Private Sub mnu_²�Ʀr�`��_Click()

��l�r�W = 1

If mnu_²�Ʀr�`��.Checked = False Then
   mnu_�`�Φr.Checked = False
   mnu_Big5.Checked = False
   mnu_²�Ʀr�`��.Checked = True
   mnu_�~�y�j�r��.Checked = False
   mnu_����.Checked = False
   mnu_�p�f.Checked = False
   mnu_����.Checked = False
   mnu_����ϧΤ�r.Checked = False
   mnu_�Ұ���.Checked = False
   mnu_���t��r.Checked = False
   If �Ұʦr�δF�� Then frm�r�δF��.txt�c�r��.FontName = "�з���"
   �t�Φr�� = "����"
End If


End Sub

Private Sub mnu_²���s��_Click()

frm²���s��.show 1

If Not ���ܹw�]�s�� Then Exit Sub

Select Case �w�]�s���Ҧ�
    Case 1: ²���s���@
    Case 2: ²���s���G
    Case 3: ²���s���T
    Case 4: ²���s���|
End Select

End Sub


Private Sub mnu_²�|_Click()

�@�ε����N�X = ²�|�N�X
frm����d��.Form_Load
frm����d��.show
frm����d��.SetFocus

End Sub

Private Sub mnu_��ܭ���X_Click()

If mnu_��ܭ���X.Checked = False Then
   mnu_��ܭ���X.Checked = True
Else
   mnu_��ܭ���X.Checked = False
End If

End Sub

Private Sub mnuToolListChar_Click()

frm�r�δF��.�C�X��w�r�������Ҧ��r��

End Sub

Private Sub mnu�K�W_Click()

Select Case �{�ε����N�X

Case mdi�~�r�r�ΥN�X

    If �{�α���N�X = mdi�~�r�r��_�s����� Then
        txt�s��.SelText = Clipboard.GetText
    ElseIf �{�α���N�X = mdi�~�r�r��_�r�Τ�� Then
        txt�r��.SelText = Clipboard.GetText
    ElseIf �{�α���N�X = mdi�~�r�r��_�c�r����� Then
        txt�c�r��.SelText = Clipboard.GetText
    End If

Case �r�δF�ťN�X
    If �{�α���N�X = �r�δF��_�˦r��� Then
        frm�r�δF��.txt�c�r��.SelText = Clipboard.GetText
    End If
    
Case �X�B�˦r�N�X
    If �{�α���N�X = �X�B�˦r_�˦r��� Then
        frm�X�B�˦r.cbo�X�B.Text = Clipboard.GetText
    End If
    
    
Case Else
    Exit Sub
End Select

End Sub

Private Sub mnu�s��_Click()

On Error Resume Next
Set WordApp = GetObject(, "word.application")
If Err.Number <> 0 Then
    WordWasNotRunning = True
    Err.Clear
Else
    WordWasNotRunning = False
End If

mnu_edit_�ƻs��Word.Enabled = False
If Not WordWasNotRunning Then
    If WordApp.Documents.Count > 0 Then mnu_edit_�ƻs��Word.Enabled = True
End If

End Sub

Private Sub txt���X_GotFocus()

txt���X.SelStart = 0
txt���X.SelLength = Len(txt���X)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_���X���

End Sub

Private Sub txt�U��_GotFocus()

txt�U��.SelStart = 0
txt�U��.SelLength = Len(txt�U��.Text)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_�U�Ƥ��

End Sub


Private Sub txt�j�~�r_GotFocus()

txt�j�~�r.SelStart = 0
txt�j�~�r.SelLength = Len(txt�j�~�r.Text)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_�j�~�r���

End Sub

Private Sub txt�~�r��_GotFocus()

txt�~�r��.SelStart = 0
txt�~�r��.SelLength = Len(txt�~�r��.Text)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_�~�r�����

End Sub

Private Sub txt�~�r��_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case txt�~�r��.Text

    Case 0: txt�~�r��.ToolTipText = "�з���(�ө���)"
    Case 1: txt�~�r��.ToolTipText = "�з���(�ө���)�~�r���@"
    Case 2: txt�~�r��.ToolTipText = "�з���(�ө���)�~�r���G"
    Case 3: txt�~�r��.ToolTipText = "�з���(�ө���)�~�r���T"
    Case 4: txt�~�r��.ToolTipText = "�з���(�ө���)�~�r���|"
    Case 5: txt�~�r��.ToolTipText = "�з���(�ө���)�~�r����"
    Case 6: txt�~�r��.ToolTipText = "�з���(�ө���)�~�r����"
    Case 7: txt�~�r��.ToolTipText = "�з���(�ө���)�~�r���C"
    Case 8: txt�~�r��.ToolTipText = "�з���(�ө���)�~�r���K"
    Case 9: txt�~�r��.ToolTipText = "�з���(�ө���)�~�r���E"

End Select

End Sub

Private Sub txt�r��_Change()

txt�r��.SelStart = 0
txt�r��.SelLength = Len(txt�r��.Text)
txt�r��.FontSize = 12

If mnu_�ƻs��ŶKï.Checked = True Then
   On Error GoTo Label1
   Clipboard.Clear
   Clipboard.SetText txt�r��.Text
Label1: End If

End Sub

Private Sub txt�r��_GotFocus()

txt�r��.SelStart = 0
txt�r��.SelLength = Len(txt�r��.Text)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_�r�Τ��

End Sub

Private Sub txt�r��_KeyPress(KeyAscii As Integer)
Dim �r�� As String, �s�� As Long, �r�� As Integer, temp As Integer
Dim ���ѽs�� As Long, �p�f�s�� As Long, ����s�� As Long, �Ұ���s�� As Long, ���t��r�s�� As Long
Dim �Ȧs�զr�� As String

mdi�~�r�r��.txt�r��.FontName = "�з���"
If KeyAscii = vbKeyReturn Then
   If Len(Trim(txt�r��.Text)) <> 0 Then
      �r�� = txt�r��.Text
      Set �˦r�� = �����˦r��
      �˦r��.Index = "�r��"
      �˦r��.Seek "=", �r��
      If �˦r��.NoMatch Then
         ���ѽs�� = -1
      Else
         ���ѽs�� = �˦r��.Fields("�s��")
         If Not IsNull(�˦r��.Fields("�p�f�s��")) Then
            �p�f�s�� = �˦r��.Fields("�p�f�s��")
         Else
            �p�f�s�� = -1
         End If
         If Not IsNull(�˦r��.Fields("����s��")) Then
            ����s�� = �˦r��.Fields("����s��")
         Else
            ����s�� = -1
         End If
         If Not IsNull(�˦r��.Fields("�Ұ���s��")) Then
            �Ұ���s�� = �˦r��.Fields("�Ұ���s��")
         Else
            �Ұ���s�� = -1
         End If
         If Not IsNull(�˦r��.Fields("���t��r�s��")) Then
            ���t��r�s�� = �˦r��.Fields("���t��r�s��")
         Else
           ���t��r�s�� = -1
         End If
      End If
            
      If ���ѽs�� > 0 Then
         �^���ݩ� "�з���", txt�r��.Text, ���ѽs��
         �^���c�r�� "�з���", txt�r��.Text, ���ѽs��
         If �t�Φr�� = "�p�f" And �p�f�s�� > 0 Then
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
         ElseIf �t�Φr�� = "����" And ����s�� > 0 Then
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "����|����", txt�r��.Text, ����s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|����", txt�r��.Text, ����s��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "����|����", txt�r��.Text, ����s��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "����|����", txt�r��.Text, ����s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|����", txt�r��.Text, ����s��
         ElseIf �t�Φr�� = "�Ұ���" And �Ұ���s�� > 0 Then
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
         ElseIf �t�Φr�� = "���t��r" And ���t��r�s�� > 0 Then
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
         Else
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "�з���", txt�r��.Text, ���ѽs��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "�з���", txt�r��.Text, ���ѽs��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "�з���", txt�r��.Text, ���ѽs��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "�з���", txt�r��.Text, ���ѽs��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "�з���", txt�r��.Text, ���ѽs��
         End If
      Else
         '�䤣��
         mdi�~�r�r��.txt�զr�r��.Text = ""
         mdi�~�r�r��.txt�r��.Text = ""
         mdi�~�r�r��.txt�`���e.Text = ""
         mdi�~�r�r��.txt����.Text = ""
         mdi�~�r�r��.txt�����������e.Text = ""
         mdi�~�r�r��.txt�`��.Text = ""
         mdi�~�r�r��.txt���X.Text = ""
         mdi�~�r�r��.txt�ܾe�X.Text = ""
         'mdi�~�r�r��.txt�c�r��.Text = ""
         mdi�~�r�r��.txt�U��.Text = ""
         
         '���w�ťխ�
         If Len(���A�C1) > 10 Then
            ���A�C = "0 �Ӧr��" & ���A�C1
         Else
            ���A�C = "0 �Ӧr��"
         End If
         mdi�~�r�r��.txt���A = ���A�C
      End If
   End If
End If
'mdi�~�r�r��.txt�r��.FontName = "�з���"
'If KeyAscii = vbKeyReturn Then
'   �r�� = txt�r��.Text
'   �˦r��.Index = "�r��"
'   �˦r��.Seek "=", �r��
'   If Not �˦r��.NoMatch() Then
'      �s�� = �˦r��.Fields("�s��")
'      �^���ݩ� "�з���", �r��, �s��
'      �^���c�r�� "�з���", �r��, �s��
'      If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "�з���", mdi�~�r�r��.txt�r��.Text, �s��
'      If �Ұʲ���r�� Then frm����r��.���J�r�� "�з���", mdi�~�r�r��.txt�r��.Text, �s��
'      If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "�з���", mdi�~�r�r��.txt�r��.Text, �s��
'      If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "�з���", mdi�~�r�r��.txt�r��.Text, �s��
'      If �Ұʲ���r�� Then frm����r��.���J�r�� "�з���", mdi�~�r�r��.txt�r��.Text, �s��
'   Else
         '�䤣��
'         mdi�~�r�r��.txt�զr�r��.Text = ""
'         'mdi�~�r�r��.txt�r��.Text = ""
'         mdi�~�r�r��.txt�`���e.Text = ""
'         mdi�~�r�r��.txt����.Text = ""
'         mdi�~�r�r��.txt�����������e.Text = ""
'         mdi�~�r�r��.txt�`��.Text = ""
'         mdi�~�r�r��.txt���X.Text = ""
'         mdi�~�r�r��.txt�ܾe�X.Text = ""
'         mdi�~�r�r��.txt�c�r��.Text = ""
'         mdi�~�r�r��.txt�U��.Text = ""
'
         '���w�ťխ�
'         If Len(���A�C1) > 10 Then
'            ���A�C = "0 �Ӧr��" & ���A�C1
'         Else
'            ���A�C = "0 �Ӧr��"
'         End If
'         mdi�~�r�r��.txt���A = ���A�C
'   End If
'End If
End Sub

Private Sub txt�����������e_GotFocus()

txt�����������e.SelStart = 0
txt�����������e.SelLength = Len(txt�����������e.Text)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_�����������e���

End Sub

Private Sub txt�`��_GotFocus()

txt�`��.SelStart = 0
txt�`��.SelLength = Len(txt�`��.Text)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_�`�����

End Sub




Private Sub txt����_GotFocus()

txt����.SelStart = 0
txt����.SelLength = Len(txt����.Text)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_������

End Sub

Private Sub txt����_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If IsNumeric(txt����.Text) Then txt����.ToolTipText = �r��}�C(CInt(txt����.Text))

End Sub

Private Sub txt�ܾe�X_GotFocus()

txt�ܾe�X.SelStart = 0
txt�ܾe�X.SelLength = Len(txt�ܾe�X.Text)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_�ܾe�X���

End Sub

Private Sub txt�զr�r�Ƨt���g_GotFocus()

�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_�զr�r�Ƨt���g���

End Sub

Private Sub txt�c�r��_Change()
If mnu_�ƻs��ŶKï.Checked = True Then
   If Not �˦r��.NoMatch Then
      If �˦r��.Fields("�r��") <> 0 Then
         On Error GoTo Lable2
         Clipboard.Clear
         Clipboard.SetText txt�c�r��.Text
Lable2: End If
   End If
   If �ƻs����X Then Clipboard.SetText txt�c�r��.Text
End If

End Sub

'Private Sub txt�c�r��_DragDrop(Source As Control, X As Single, Y As Single)
'Dim �r�� As String
'Dim ���b�� As String
'Dim �k�b�� As String

'���b�� = Left(txt�c�r��, txt�c�r��.SelStart)'
'�k�b�� = Right$(txt�c�r��, Len(txt�c�r��) - txt�c�r��.SelStart)
  
'If TypeOf Source Is ListBox Then
'   If Source.ListIndex < 0 Then Exit Sub
'   Source.Drag 2       ' End Dragging
'   txt�c�r�� = ���b�� & Source.List(Source.ListIndex) & �k�b��
'End If

'If TypeOf Source Is TList Then
'   If Source Is Nothing Then Exit Sub
'   Source.Drag 2       ' End Dragging
'   Screen.MousePointer = 11
'   �r�� = Left(Source, 2)
'   'txt�c�r�� = ���b�� & mdi�~�r�r��.txt�r��.Text & �k�b��
'   txt�c�r�� = ���b�� & �즲�r�� & �k�b��
'   Screen.MousePointer = 0
'End If
'txt�c�r��.SetFocus
'txt�c�r��.SelStart = Len(txt�c�r��)

'End Sub

Private Sub txt�c�r��_GotFocus()

txt�c�r��.SelStart = 0
txt�c�r��.SelLength = Len(txt�c�r��.Text)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_�c�r�����

End Sub

Private Sub txt�c�r��_KeyPress(KeyAscii As Integer)
Dim �r�ڧ� As String
Dim ���ѽs�� As Long, �p�f�s�� As Long, ����s�� As Long, �Ұ���s�� As Long, ���t��r�s�� As Long
Dim �Ȧs�զr�� As String

If KeyAscii = vbKeyReturn Then
   If Len(Trim(txt�c�r��.Text)) <> 0 Then
      ���ѽs�� = �r�άd��()
      If ���ѽs�� > 0 Then
        Set �˦r�� = �����˦r��
        �˦r��.Index = "�s��"
        �˦r��.Seek "=", ���ѽs��
        If Not IsNull(�˦r��.Fields("�p�f�s��")) Then
            �p�f�s�� = �˦r��.Fields("�p�f�s��")
        Else
            �p�f�s�� = -1
        End If
        If Not IsNull(�˦r��.Fields("����s��")) Then
            ����s�� = �˦r��.Fields("����s��")
        Else
            ����s�� = -1
        End If
        If Not IsNull(�˦r��.Fields("�Ұ���s��")) Then
            �Ұ���s�� = �˦r��.Fields("�Ұ���s��")
        Else
            �Ұ���s�� = -1
        End If
        If Not IsNull(�˦r��.Fields("���t��r�s��")) Then
            ���t��r�s�� = �˦r��.Fields("���t��r�s��")
        Else
            ���t��r�s�� = -1
        End If

      End If
      
      If ���ѽs�� > 0 Then
         �^���ݩ� "�з���", txt�r��.Text, ���ѽs��
         �^���c�r�� "�з���", txt�r��.Text, ���ѽs��
         If �t�Φr�� = "�p�f" And �p�f�s�� > 0 Then
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
         ElseIf �t�Φr�� = "����" And ����s�� > 0 Then
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "����|����", txt�r��.Text, ����s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|����", txt�r��.Text, ����s��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "����|����", txt�r��.Text, ����s��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "����|����", txt�r��.Text, ����s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|����", txt�r��.Text, ����s��
         ElseIf �t�Φr�� = "�Ұ���" And �Ұ���s�� > 0 Then
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
         ElseIf �t�Φr�� = "���t��r" And ���t��r�s�� > 0 Then
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
         Else
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "�з���", txt�r��.Text, ���ѽs��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "�з���", txt�r��.Text, ���ѽs��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "�з���", txt�r��.Text, ���ѽs��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "�з���", txt�r��.Text, ���ѽs��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "�з���", txt�r��.Text, ���ѽs��
         End If
      Else
         '�䤣��
         mdi�~�r�r��.txt�զr�r��.Text = ""
         mdi�~�r�r��.txt�r��.Text = ""
         mdi�~�r�r��.txt�`���e.Text = ""
         mdi�~�r�r��.txt����.Text = ""
         mdi�~�r�r��.txt�����������e.Text = ""
         mdi�~�r�r��.txt�`��.Text = ""
         mdi�~�r�r��.txt���X.Text = ""
         mdi�~�r�r��.txt�ܾe�X.Text = ""
         'mdi�~�r�r��.txt�c�r��.Text = ""
         mdi�~�r�r��.txt�U��.Text = ""
         
         '���w�ťխ�
         If Len(���A�C1) > 10 Then
            ���A�C = "0 �Ӧr��" & ���A�C1
         Else
            ���A�C = "0 �Ӧr��"
         End If
         mdi�~�r�r��.txt���A = ���A�C
      End If
   End If
End If

End Sub

Private Function �r�άd��() As Long
Dim �զr�� As String, �r�ڧ� As String, �B��Ÿ� As Integer, �c�r�� As String, ����c�r�� As String
Dim i As Integer, j As Integer, �զr�Ÿ� As Integer, ���Ц��� As Integer, ��K�Ÿ� As Integer
Dim ���Цr As String, �r�ڲ� As String, �ۦ��r�� As String, �Ȧs�զr�� As String
Dim �����c�r�� As String, ���ѲŸ� As Boolean

On Error GoTo �r�άd�߿��~

�B��Ÿ� = 4
���Ц��� = 0
�զr�� = ""
�c�r�� = ""
�ۦ��r�� = ""
�����c�r�� = ""
���ѲŸ� = False

�r�άd�� = -1

If Len(Trim(txt�c�r��.Text)) = 1 Then
   �˦r��.Index = "�r��"
   �˦r��.Seek "=", Trim(txt�c�r��.Text)
   If Not �˦r��.NoMatch Then �r�άd�� = �˦r��.Fields("�s��")
   Exit Function
End If

i = 1
Do While i <= Len(Trim(txt�c�r��.Text))
 
   �զr�Ÿ� = �O�_���զr�Ÿ�(Mid(txt�c�r��.Text, i, 1), 1, 14)

   If �զr�Ÿ� <> 12 And �զr�Ÿ� <> 13 Then
      If �զr�Ÿ� >= 1 And �զr�Ÿ� <= 3 Then
         �B��Ÿ� = �զr�Ÿ�
      ElseIf Len(txt�c�r��) = 2 Then
         �B��Ÿ� = 5
         �զr�� = txt�c�r��
         Exit Do
      End If
      If �զr�Ÿ� >= 4 Or �զr�Ÿ� = 0 Then
         �զr�� = �զr�� & Mid(txt�c�r��, i, 1)
         If Len(txt�c�r��) = 1 And �զr�Ÿ� = 0 Then �B��Ÿ� = 0
      End If
   Else
      �B��Ÿ� = 4
      �զr�� = �զr�� & Mid(txt�c�r��, 2, Len(txt�c�r��) - 2)
      Exit Do
   End If
   i = i + 1
Loop
    
If �զr�� = "" Then �զr�� = Trim(txt�c�r��.Text)

�c�r�� = �զr��

'���Τ���+����Ƿj�M

If Len(�c�r��) > 1 Or (Len(�c�r��) = 1 And �զr�Ÿ� = 0) Then
   �˦r��.Index = "�c�r��"
   �˦r��.Seek "=", �B��Ÿ�, �c�r��
   If Not �˦r��.NoMatch Then
      �r�άd�� = �˦r��.Fields("�s��")
      If �B��Ÿ� > 0 Then
         If �˦r��.Fields("�r��") <> 0 Or IsNull(�˦r��.Fields("�r��")) Then
            �����c�r�� = �M��զr��(�˦r��.Fields("�s���Ÿ�"), �˦r��.Fields("�����"))
         Else
            �����c�r�� = �˦r��.Fields("�r��")
         End If
      End If
      �˦r��.MoveNext
      Do Until �˦r��.EOF Or �˦r��.Fields("�����") <> �c�r��
         If �˦r��.Fields("�s���Ÿ�") = �B��Ÿ� Then
            �r�άd�� = �˦r��.Fields("�s��")
            ���ѲŸ� = True
         Else
            If �˦r��.Fields("�r��") <> 0 Or IsNull(�˦r��.Fields("�r��")) Then
               �Ȧs�զr�� = �M��զr��(�˦r��.Fields("�s���Ÿ�"), �˦r��.Fields("�����"))
            Else
               �Ȧs�զr�� = �˦r��.Fields("�r��")
            End If
            �ۦ��r�� = �ۦ��r�� & "[" & �Ȧs�զr�� & "]"
         End If
         �˦r��.MoveNext
      Loop
      If ���ѲŸ� = True Then �ۦ��r�� = �����c�r�� & �ۦ��r��
      If Len(�ۦ��r��) > 0 Then ���A�C1 = ",�ۦ��r��: " & �ۦ��r��
      Exit Function
   End If
Else
   �r�άd�� = 0
   txt�r��.Text = �c�r��
   Exit Function
End If
       
'�����not found,��Φr�ڧ�
'�O�զr�Ÿ�
i = 1
�c�r�� = ""
�˦r��.Index = "�r��"
Do While i <= Len(�զr��)
   ��K�Ÿ� = �O�_���զr�Ÿ�(Mid(�զr��, i, 1), 4, 11)
   Select Case ��K�Ÿ�
          Case 4, 5
               ���Ц��� = 2
          Case 6, 7, 8
               ���Ц��� = 3
          Case 9, 10, 11
               ���Ц��� = 4
   End Select
   If ��K�Ÿ� > 0 Then
      i = i + 1
      �˦r��.Seek "=", Mid(�զr��, i, 1)
      If Not �˦r��.NoMatch Then
         '�M��ӭ��Цr���r��,�Y���Ѥ���0,�h�~�򩹤U��X�Ҧ��Ӧr�Τ��r��
         If �˦r��.Fields("�s���Ÿ�") <> 0 Then
            ���Цr = �r�ڧǬd��(�˦r��.Fields("�r�ڧ�"))
         Else
            ���Цr = Mid(�զr��, i, 1)
         End If
      End If
      For j = 1 To ���Ц���
          �c�r�� = �c�r�� & ���Цr
      Next j
   '���O�զr�Ÿ�
   Else
       �˦r��.Index = "�r��"
       �˦r��.Seek "=", Mid(�զr��, i, 1)
       If Not �˦r��.NoMatch Then
          Do Until �˦r��.EOF Or �˦r��.Fields("�r��") <> Mid(�զr��, i, 1)
             If �˦r��.Fields("�r��") = 0 And �˦r��.Fields("�r��") = Mid(�զr��, i, 1) Then
                If �˦r��.Fields("�s���Ÿ�") <> 0 Then
                   �c�r�� = �c�r�� & �r�ڧǬd��(�˦r��.Fields("�r�ڧ�"))
                Else
                   �c�r�� = �c�r�� & �˦r��.Fields("�r�ڧ�")
                End If
             End If
             �˦r��.MoveNext
          Loop
       End If
   End If
   i = i + 1
Loop
    
'���Φr�ڧǷj�M

�˦r��.Index = "�r�ڧ�"
�˦r��.Seek "=", �c�r��

If Not �˦r��.NoMatch Then
   �r�άd�� = �˦r��.Fields("�s��")
   If �B��Ÿ� > 0 Then
      If �˦r��.Fields("�r��") <> 0 Or IsNull(�˦r��.Fields("�r��")) Then
         �����c�r�� = �M��զr��(�˦r��.Fields("�s���Ÿ�"), �˦r��.Fields("�����"))
      Else
         �����c�r�� = �˦r��.Fields("�r��")
      End If
   End If
   �˦r��.MoveNext
   Do Until �˦r��.EOF Or �˦r��.Fields("�����") <> �c�r��
      If �˦r��.Fields("�s���Ÿ�") = �B��Ÿ� Then
         �r�άd�� = �˦r��.Fields("�s��")
         ���ѲŸ� = True
      Else
         If �˦r��.Fields("�r��") <> 0 Or IsNull(�˦r��.Fields("�r��")) Then
            �Ȧs�զr�� = �M��զr��(�˦r��.Fields("�s���Ÿ�"), �˦r��.Fields("�����"))
         Else
            �Ȧs�զr�� = �˦r��.Fields("�r��")
         End If
         �ۦ��r�� = �ۦ��r�� & "[" & �Ȧs�զr�� & "]"
      End If
      �˦r��.MoveNext
   Loop
   If ���ѲŸ� = True Then �ۦ��r�� = �����c�r�� & �ۦ��r��
   If Len(�ۦ��r��) > 0 Then ���A�C1 = ",�ۦ��r��: " & �ۦ��r��
   Exit Function
End If
   
'�Y�Φr�ڧǷj�M����,�h��Φr�ڧǤG�j�M
����c�r�� = ""
For i = 1 To Len(�c�r��)
    ���g�r��.Seek "=", Mid(�c�r��, i, 1)
    If Not ���g�r��.NoMatch Then
        ����c�r�� = ����c�r�� & ���g�r��.Fields("���g")
    Else
        ����c�r�� = ����c�r�� & Mid(�c�r��, i, 1)
    End If
Next i

�˦r��.Index = "�r�ڧǤG"
�˦r��.Seek "=", ����c�r��

If Not �˦r��.NoMatch Then
   �r�άd�� = �˦r��.Fields("�s��")
   If �B��Ÿ� > 0 Then
      If �˦r��.Fields("�r��") <> 0 Or IsNull(�˦r��.Fields("�r��")) Then
         �����c�r�� = �M��զr��(�˦r��.Fields("�s���Ÿ�"), �˦r��.Fields("�����"))
      Else
         �����c�r�� = �˦r��.Fields("�r��")
      End If
   End If
   �˦r��.MoveNext
   Do Until �˦r��.EOF Or �˦r��.Fields("�����") <> �c�r��
      If �˦r��.Fields("�s���Ÿ�") = �B��Ÿ� Then
         �r�άd�� = �˦r��.Fields("�s��")
         ���ѲŸ� = True
      Else
         If �˦r��.Fields("�r��") <> 0 Or IsNull(�˦r��.Fields("�r��")) Then
            �Ȧs�զr�� = �M��զr��(�˦r��.Fields("�s���Ÿ�"), �˦r��.Fields("�����"))
         Else
            �Ȧs�զr�� = �˦r��.Fields("�r��")
         End If
         �ۦ��r�� = �ۦ��r�� & "[" & �Ȧs�զr�� & "]"
      End If
      �˦r��.MoveNext
   Loop
   If ���ѲŸ� = True Then �ۦ��r�� = �����c�r�� & �ۦ��r��
   If Len(�ۦ��r��) > 0 Then ���A�C1 = ",�ۦ��r��: " & �ۦ��r��
   Exit Function
End If
   

'�Y�Φr�ڧǷj�M����,�h��Φr�ڲշj�M

�r�ڲ� = �r�ڱƧ�(�c�r��)
�˦r��.Index = "�r�ڲ�"
�˦r��.Seek "=", �r�ڲ�
If Not �˦r��.NoMatch Then
   �r�άd�� = �˦r��.Fields("�s��")
   If �B��Ÿ� > 0 Then
      If �˦r��.Fields("�r��") <> 0 Or IsNull(�˦r��.Fields("�r��")) Then
         �����c�r�� = �M��զr��(�˦r��.Fields("�s���Ÿ�"), �˦r��.Fields("�����"))
      Else
         �����c�r�� = �˦r��.Fields("�r��")
      End If
   End If
   �˦r��.MoveNext
   Do Until �˦r��.EOF Or �˦r��.Fields("�����") <> �c�r��
      If �˦r��.Fields("�s���Ÿ�") = �B��Ÿ� Then
         �r�άd�� = �˦r��.Fields("�s��")
         ���ѲŸ� = True
      Else
         If �˦r��.Fields("�r��") <> 0 Or IsNull(�˦r��.Fields("�r��")) Then
            �Ȧs�զr�� = �M��զr��(�˦r��.Fields("�s���Ÿ�"), �˦r��.Fields("�����"))
         Else
            �Ȧs�զr�� = �˦r��.Fields("�r��")
         End If
         �ۦ��r�� = �ۦ��r�� & "[" & �Ȧs�զr�� & "]"
      End If
      �˦r��.MoveNext
   Loop
   If ���ѲŸ� = True Then �ۦ��r�� = �����c�r�� & �ۦ��r��
   If Len(�ۦ��r��) > 0 Then ���A�C1 = ",�ۦ��r��: " & �ۦ��r��
   Exit Function
End If

'�Y�Φr�ڧǷj�M����,�h��Φr�ڲդG�j�M

����c�r�� = ""
For i = 1 To Len(�r�ڲ�)
    ���g�r��.Seek "=", Mid(�r�ڲ�, i, 1)
    If Not ���g�r��.NoMatch Then
        ����c�r�� = ����c�r�� & ���g�r��.Fields("���g")
    Else
        ����c�r�� = ����c�r�� & Mid(�r�ڲ�, i, 1)
    End If
Next i

����c�r�� = �r�ڱƧ�(����c�r��)
�˦r��.Index = "�r�ڲդG"
�˦r��.Seek "=", ����c�r��
If Not �˦r��.NoMatch Then
   �r�άd�� = �˦r��.Fields("�s��")
   If �B��Ÿ� > 0 Then
      If �˦r��.Fields("�r��") <> 0 Or IsNull(�˦r��.Fields("�r��")) Then
         �����c�r�� = �M��զr��(�˦r��.Fields("�s���Ÿ�"), �˦r��.Fields("�����"))
      Else
         �����c�r�� = �˦r��.Fields("�r��")
      End If
   End If
   �˦r��.MoveNext
   Do Until �˦r��.EOF Or �˦r��.Fields("�����") <> �c�r��
      If �˦r��.Fields("�s���Ÿ�") = �B��Ÿ� Then
         �r�άd�� = �˦r��.Fields("�s��")
         ���ѲŸ� = True
      Else
         If �˦r��.Fields("�r��") <> 0 Or IsNull(�˦r��.Fields("�r��")) Then
            �Ȧs�զr�� = �M��զr��(�˦r��.Fields("�s���Ÿ�"), �˦r��.Fields("�����"))
         Else
            �Ȧs�զr�� = �˦r��.Fields("�r��")
         End If
         �ۦ��r�� = �ۦ��r�� & "[" & �Ȧs�զr�� & "]"
      End If
      �˦r��.MoveNext
   Loop
   If ���ѲŸ� = True Then �ۦ��r�� = �����c�r�� & �ۦ��r��
   If Len(�ۦ��r��) > 0 Then ���A�C1 = ",�ۦ��r��: " & �ۦ��r��
   Exit Function
End If

�r�άd�߿��~:

End Function

Private Function �r�ڧǬd��(�r�ڧ� As String) As String
Dim i As Integer, varBookMark As Variant
Dim �զr�� As String

varBookMark = �˦r��.Bookmark
�˦r��.Index = "�r��"
�զr�� = ""
For i = 1 To Len(�r�ڧ�)
    �˦r��.Seek "=", Mid(�r�ڧ�, i, 1)
    If Not �˦r��.NoMatch Then
       If �˦r��.Fields("�s���Ÿ�") <> 0 And �˦r��.Fields("�s���Ÿ�") <> 9 Then
          �զr�� = �զr�� & �r�ڧǬd��(�˦r��.Fields("�r�ڧ�"))
       Else
          �զr�� = �զr�� & �˦r��.Fields("�r�ڧ�")
       End If
    End If
Next i
�r�ڧǬd�� = �զr��
�˦r��.Bookmark = varBookMark

End Function


Private Sub txt����_GotFocus()

txt����.SelStart = 0
txt����.SelLength = Len(txt����.Text)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_�������

End Sub

Private Sub txt�զr�r��_GotFocus()

txt�զr�r��.SelStart = 0
txt�զr�r��.SelLength = Len(txt�զr�r��.Text)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_�զr�r�Ƥ��

End Sub

Private Sub txt�s��_GotFocus()

txt�s��.SelStart = 0
txt�s��.SelLength = Len(txt�s��.Text)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_�s�����

End Sub

Private Sub txt�s��_KeyPress(KeyAscii As Integer)

Dim ���ѽs�� As Long, �p�f�s�� As Long, ����s�� As Long, �Ұ���s�� As Long, ���t��r�s�� As Long
Dim �Ȧs�զr�� As String

If KeyAscii = vbKeyReturn Then
   If IsNumeric(txt�s��.Text) Then
      ���ѽs�� = CLng(txt�s��.Text)
      If ���ѽs�� > 0 Then
        Set �˦r�� = �����˦r��
        �˦r��.Index = "�s��"
        �˦r��.Seek "=", ���ѽs��
        If �˦r��.NoMatch Then GoTo �䤣��
        If Not IsNull(�˦r��.Fields("�p�f�s��")) Then
            �p�f�s�� = �˦r��.Fields("�p�f�s��")
        Else
            �p�f�s�� = -1
        End If
        If Not IsNull(�˦r��.Fields("����s��")) Then
            ����s�� = �˦r��.Fields("����s��")
        Else
            ����s�� = -1
        End If
        If Not IsNull(�˦r��.Fields("�Ұ���s��")) Then
            �Ұ���s�� = �˦r��.Fields("�Ұ���s��")
        Else
            �Ұ���s�� = -1
        End If
        If Not IsNull(�˦r��.Fields("���t��r�s��")) Then
            ���t��r�s�� = �˦r��.Fields("���t��r�s��")
        Else
            ���t��r�s�� = -1
        End If

      End If
      
      If ���ѽs�� > 0 Then
         �^���ݩ� "�з���", txt�r��.Text, ���ѽs��
         �^���c�r�� "�з���", txt�r��.Text, ���ѽs��
         If �t�Φr�� = "�p�f" And �p�f�s�� > 0 Then
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "�_�v�j����p�f", txt�r��.Text, �p�f�s��
         ElseIf �t�Φr�� = "����" And ����s�� > 0 Then
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "����|����", txt�r��.Text, ����s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|����", txt�r��.Text, ����s��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "����|����", txt�r��.Text, ����s��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "����|����", txt�r��.Text, ����s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|����", txt�r��.Text, ����s��
         ElseIf �t�Φr�� = "�Ұ���" And �Ұ���s�� > 0 Then
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|�Ұ���", txt�r��.Text, �Ұ���s��
         ElseIf �t�Φr�� = "���t��r" And ���t��r�s�� > 0 Then
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "����|���t²����r", txt�r��.Text, ���t��r�s��
         Else
            If �Ұʦr�ε��c Then frm�r�ε��c.���J�r�� "�з���", txt�r��.Text, ���ѽs��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "�з���", txt�r��.Text, ���ѽs��
            If �Ұʦr�κt�� Then frm�r�κt��.���J�r�� "�з���", txt�r��.Text, ���ѽs��
            If �Ұʦr�ί��� Then frm�r�ί���.���J�r�� "�з���", txt�r��.Text, ���ѽs��
            If �Ұʲ���r�� Then frm����r��.���J�r�� "�з���", txt�r��.Text, ���ѽs��
         End If
      Else
�䤣��:
         mdi�~�r�r��.txt�զr�r��.Text = ""
         mdi�~�r�r��.txt�r��.Text = ""
         mdi�~�r�r��.txt�`���e.Text = ""
         mdi�~�r�r��.txt����.Text = ""
         mdi�~�r�r��.txt�����������e.Text = ""
         mdi�~�r�r��.txt�`��.Text = ""
         mdi�~�r�r��.txt���X.Text = ""
         mdi�~�r�r��.txt�ܾe�X.Text = ""
         'mdi�~�r�r��.txt�c�r��.Text = ""
         mdi�~�r�r��.txt�U��.Text = ""
         
         '���w�ťխ�
         If Len(���A�C1) > 10 Then
            ���A�C = "0 �Ӧr��" & ���A�C1
         Else
            ���A�C = "0 �Ӧr��"
         End If
         mdi�~�r�r��.txt���A = ���A�C
      End If
   End If
End If

End Sub

Private Sub txt�`���e_GotFocus()

txt�`���e.SelStart = 0
txt�`���e.SelLength = Len(txt�`���e.Text)
�{�ε����N�X = mdi�~�r�r�ΥN�X
�{�α���N�X = mdi�~�r�r��_�`���e���

End Sub

Public Sub ²���s���@()
Dim i As Integer
Dim ���e As Integer
Dim ���� As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case �r�ε��c�N�X, ����r�ڥN�X, ����r��N�X, �r�κt�ܥN�X, �r�ί��ޥN�X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

If �t�Φr�� = "�p�f" Or �t�Φr�� = "����" Or �t�Φr�� = "�Ұ���" Or �t�Φr�� = "���t��r" Then
    mnu_����Ѧr����_Click
Else
    mnu_�d���r�峡��_Click
End If

frm�r�ε��c.Tag = �r�ε��c�N�X
frm�r�ε��c.show
frm�r�ε��c.SetFocus

'frm�r�ί���.Tag = �r�ί��ޥN�X
'frm�r�ί���.show
'frm�r�ί���.SetFocus

'frm����r��.Tag = 15
'frm����r��.show
'frm����r��.Visible = False

'frm�r�κt��.Tag = 16
'frm�r�κt��.show
'frm�r�κt��.Visible = False

frm�r�δF��.Tag = �r�δF�ťN�X
frm�r�δF��.show
frm�r�δF��.SetFocus

mdi�~�r�r��.Arrange 2

'frm�r�ε��c.Height = frm�r�δF��.Height / 2
'frm����d��.Height = frm�r�ε��c.Height
'frm�r�κt��.Height = frm�r�κt��.Height

'frm�r�κt��.Left = frm�r�ε��c.Left
'frm�r�κt��.Width = frm�r�ε��c.Width
'frm�r�κt��.Height = frm�r�ε��c.Height - 1
'frm�r�κt��.Top = frm�r�ε��c.Height

'frm����r��.Left = frm����d��.Left
'frm����r��.Width = frm����d��.Width
'frm����r��.Height = frm����d��.Height - 1
'frm����r��.Top = frm����d��.Height

'frm�r�κt��.Visible = True
'frm����r��.Visible = True
frm�r�δF��.SetFocus

End Sub

Public Sub ²���s���G()

Dim i As Integer
Dim ���e As Integer
Dim ���� As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case Big5�r�ڥN�X To �c�r�Ÿ��N�X, ����r�ڥN�X, �r�κt�ܥN�X, �r�ί��ޥN�X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

'mnu_����N�X_Click 3
frm�r�ε��c.Tag = �r�ε��c�N�X
frm�r�ε��c.show
frm�r�ε��c.SetFocus

frm����r��.Tag = ����r��N�X
frm����r��.show
frm����r��.SetFocus

'frm����r��.Tag = 15
'frm����r��.show
'frm����r��.Visible = False

'frm�r�κt��.Tag = 16
'frm�r�κt��.show
'frm�r�κt��.Visible = False

frm�r�δF��.Tag = �r�δF�ťN�X
frm�r�δF��.show
frm�r�δF��.SetFocus

mdi�~�r�r��.Arrange 2

'frm�r�ε��c.Height = frm�r�δF��.Height / 2
'frm����d��.Height = frm�r�ε��c.Height
'frm�r�κt��.Height = frm�r�κt��.Height

'frm�r�κt��.Left = frm�r�ε��c.Left
'frm�r�κt��.Width = frm�r�ε��c.Width
'frm�r�κt��.Height = frm�r�ε��c.Height - 1
'frm�r�κt��.Top = frm�r�ε��c.Height

'frm����r��.Left = frm����d��.Left
'frm����r��.Width = frm����d��.Width
'frm����r��.Height = frm����d��.Height - 1
'frm����r��.Top = frm����d��.Height

'frm�r�κt��.Visible = True
'frm����r��.Visible = True
frm�r�δF��.SetFocus

End Sub

Public Sub ²���s���T()

Dim i As Integer
Dim ���e As Integer
Dim ���� As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case Big5�r�ڥN�X To �c�r�Ÿ��N�X, ����r�ڥN�X, �r�κt�ܥN�X, �r�ε��c�N�X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

frm�r�ί���.Tag = �r�ί��ޥN�X
frm�r�ί���.show
frm�r�ί���.SetFocus

frm����r��.Tag = ����r��N�X
frm����r��.show
frm����r��.SetFocus

frm�r�δF��.Tag = �r�δF�ťN�X
frm�r�δF��.show
frm�r�δF��.SetFocus

mdi�~�r�r��.Arrange 2

frm�r�δF��.SetFocus

End Sub

Public Sub ²���s���|()

Dim i As Integer
Dim ���e As Integer
Dim ���� As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case Big5�r�ڥN�X To �c�r�Ÿ��N�X, ����r�ڥN�X, �r�κt�ܥN�X, ����r��N�X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

frm�r�ί���.Tag = �r�ί��ޥN�X
frm�r�ί���.show
frm�r�ί���.SetFocus

frm�r�ε��c.Tag = �r�ε��c�N�X
frm�r�ε��c.show
frm�r�ε��c.SetFocus

frm�r�δF��.Tag = �r�δF�ťN�X
frm�r�δF��.show
frm�r�δF��.SetFocus

mdi�~�r�r��.Arrange 2

frm�r�δF��.SetFocus

End Sub
Public Sub �w�]�s���@()

Dim i As Integer
Dim ���e As Integer
Dim ���� As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case Big5�r�ڥN�X To �c�r�Ÿ��N�X, ����r�ڥN�X, �r�ί��ޥN�X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

'mnu_����N�X_Click 3
frm�r�ε��c.Tag = �r�ε��c�N�X
frm�r�ε��c.show
frm�r�ε��c.SetFocus

frm����r��.Tag = ����r��N�X
frm����r��.show
frm����r��.SetFocus

'frm����r��.Tag = 15
'frm����r��.show
'frm����r��.Visible = False

frm�r�κt��.Tag = �r�κt�ܥN�X
frm�r�κt��.show
frm�r�κt��.Visible = False

frm�r�δF��.Tag = �r�δF�ťN�X
frm�r�δF��.show
frm�r�δF��.SetFocus

mdi�~�r�r��.Arrange 2

frm�r�ε��c.Height = frm�r�δF��.Height / 2
'frm����d��.Height = frm�r�ε��c.Height
'frm�r�κt��.Height = frm�r�κt��.Height

frm�r�κt��.Left = frm�r�ε��c.Left
frm�r�κt��.Width = frm�r�ε��c.Width
frm�r�κt��.Height = frm�r�ε��c.Height - 1
frm�r�κt��.Top = frm�r�ε��c.Height

'frm����r��.Left = frm����d��.Left
'frm����r��.Width = frm����d��.Width
'frm����r��.Height = frm����d��.Height - 1
'frm����r��.Top = frm����d��.Height

frm�r�κt��.Visible = True
'frm����r��.Visible = True
frm�r�δF��.SetFocus

End Sub

Public Sub �w�]�s���G()

Dim i As Integer
Dim ���e As Integer
Dim ���� As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case Big5�r�ڥN�X To �c�r�Ÿ��N�X, ����r�ڥN�X, �r�ί��ޥN�X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

frm����r��.Tag = ����r��N�X
frm����r��.show
frm����r��.SetFocus

frm�r�ε��c.Tag = �r�ε��c�N�X
frm�r�ε��c.show
frm�r�ε��c.SetFocus

frm�r�κt��.Tag = �r�κt�ܥN�X
frm�r�κt��.show
frm�r�κt��.Visible = False

frm�r�δF��.Tag = �r�δF�ťN�X
frm�r�δF��.show
frm�r�δF��.SetFocus

mdi�~�r�r��.Arrange 2

frm����r��.Height = frm�r�δF��.Height / 2

frm�r�κt��.Left = frm����r��.Left
frm�r�κt��.Width = frm����r��.Width
frm�r�κt��.Height = frm����r��.Height - 1
frm�r�κt��.Top = frm����r��.Height

frm�r�κt��.Visible = True
frm�r�δF��.SetFocus

End Sub
Public Sub �w�]�s���T()

Dim i As Integer
Dim ���e As Integer
Dim ���� As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case Big5�r�ڥN�X To �c�r�Ÿ��N�X, ����r�ڥN�X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

frm�r�ί���.Tag = �r�ί��ޥN�X
frm�r�ί���.show
frm�r�ί���.SetFocus

frm�r�ε��c.Tag = �r�ε��c�N�X
frm�r�ε��c.show
frm�r�ε��c.SetFocus

frm����r��.Tag = ����r��N�X
frm����r��.show
frm����r��.Visible = False

frm�r�κt��.Tag = �r�κt�ܥN�X
frm�r�κt��.show
frm�r�κt��.Visible = False

frm�r�δF��.Tag = �r�δF�ťN�X
frm�r�δF��.show
frm�r�δF��.SetFocus

mdi�~�r�r��.Arrange 2

frm�r�ε��c.Height = frm�r�δF��.Height / 2
frm�r�ί���.Height = frm�r�ε��c.Height

frm����r��.Left = frm�r�ε��c.Left
frm����r��.Width = frm�r�ε��c.Width
frm����r��.Height = frm�r�ε��c.Height - 1
frm����r��.Top = frm�r�ε��c.Height

frm�r�κt��.Left = frm�r�ί���.Left
frm�r�κt��.Width = frm�r�ί���.Width
frm�r�κt��.Height = frm�r�ί���.Height - 1
frm�r�κt��.Top = frm�r�ί���.Height

frm����r��.Visible = True
frm�r�κt��.Visible = True
frm�r�δF��.SetFocus

End Sub

Public Sub �w�]�s���|()

Dim i As Integer
Dim ���e As Integer
Dim ���� As Integer

i = 1
Do While i <= Forms.Count - 1
    Select Case CInt(Forms(i).Tag)
        Case �r�κt�ܥN�X, �r�ί��ޥN�X
            Unload Forms(i)
        Case Else
            i = i + 1
    End Select
Loop

If �t�Φr�� = "�p�f" Then
    mnu_����Ѧr����_Click
Else
    mnu_�d���r�峡��_Click
End If

frm�r�ε��c.Tag = �r�ε��c�N�X
frm�r�ε��c.show
frm�r�ε��c.SetFocus

frm����r��.Tag = ����r��N�X
frm����r��.show
frm����r��.Visible = False

frm����r��.Tag = ����r�ڥN�X
frm����r��.show
frm����r��.Visible = False

frm�r�δF��.Tag = �r�δF�ťN�X
frm�r�δF��.show
frm�r�δF��.SetFocus

mdi�~�r�r��.Arrange 2

frm�r�ε��c.Height = frm�r�δF��.Height / 2
frm����d��.Height = frm�r�ε��c.Height

frm����r��.Left = frm�r�ε��c.Left
frm����r��.Width = frm�r�ε��c.Width
frm����r��.Height = frm�r�ε��c.Height - 1
frm����r��.Top = frm�r�ε��c.Height

frm����r��.Left = frm����d��.Left
frm����r��.Width = frm����d��.Width
frm����r��.Height = frm����d��.Height - 1
frm����r��.Top = frm����d��.Height

frm����r��.Visible = True
frm����r��.Visible = True
frm�r�δF��.SetFocus

End Sub

