VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form DGA1030C 
   Caption         =   "�ȴ�����¯��ҵʵ����ѯ���޸�_DGA1030C"
   ClientHeight    =   9495
   ClientLeft      =   -450
   ClientTop       =   1455
   ClientWidth     =   15240
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread ss1 
      Height          =   7530
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1770
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   13282
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ButtonDrawMode  =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   34
      MaxRows         =   20
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "DGA1030C.frx":0000
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   660
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   1164
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox txt_PrcLine 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   20
         Tag             =   "����"
         Top             =   210
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.ComboBox cbo_PrcLine 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "DGA1030C.frx":1140
         Left            =   4170
         List            =   "DGA1030C.frx":1142
         TabIndex        =   19
         Tag             =   "¯����"
         Top             =   180
         Width           =   1635
      End
      Begin VB.TextBox TXT_PLT_NAME 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         MaxLength       =   50
         TabIndex        =   18
         Tag             =   "����"
         Top             =   180
         Width           =   1020
      End
      Begin VB.TextBox txt_Plt 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1305
         MaxLength       =   2
         TabIndex        =   17
         Tag             =   "����"
         Top             =   180
         Width           =   540
      End
      Begin VB.TextBox txt_iType 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   16
         Tag             =   "����"
         Text            =   "1"
         Top             =   210
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.ComboBox cbo_chg_no 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "DGA1030C.frx":1144
         Left            =   7110
         List            =   "DGA1030C.frx":1146
         TabIndex        =   15
         Tag             =   "¯����"
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox TXT_MAT_NO 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10050
         TabIndex        =   14
         Top             =   180
         Width           =   1620
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   90
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   2940
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   5910
         Top             =   180
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "¯����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   8835
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "��ѯ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1020
      Left            =   120
      TabIndex        =   21
      Top             =   750
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   1799
      _Version        =   196609
      BackColor       =   12632319
      Begin VB.TextBox TXT_SHIFT 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   810
         MaxLength       =   1
         TabIndex        =   4
         Top             =   570
         Width           =   705
      End
      Begin VB.TextBox TXT_GROUP 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2325
         MaxLength       =   1
         TabIndex        =   5
         Top             =   570
         Width           =   615
      End
      Begin VB.TextBox TXT_EMP 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7515
         MaxLength       =   8
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   150
         Width           =   1155
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   3795
         Top             =   150
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "�����¶�"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   90
         Top             =   570
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Caption         =   "���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel35 
         Height          =   315
         Left            =   1605
         Top             =   570
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Caption         =   "���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel36 
         Height          =   315
         Left            =   6285
         Top             =   150
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "��ҵ��Ա"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_HeatTemp 
         Height          =   315
         Left            =   5025
         TabIndex        =   1
         Top             =   150
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   1
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   90
         Top             =   150
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "��¯ʱ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sitxEdit TXT_DISCHARGE_TIME 
         Height          =   315
         Left            =   1395
         TabIndex        =   0
         Tag             =   "��¯ʱ��"
         Top             =   150
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __:__:__"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   "____-__-__ __:__:__"
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   5520
         Top             =   465
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "פ��ʱ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   11130
         Top             =   630
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "��ȴ�¶�(��/��)"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   3795
         Top             =   645
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "��¯�¶�"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_DisCharTemp 
         Height          =   315
         Left            =   5025
         TabIndex        =   6
         Top             =   645
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1773
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   1
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_WAT_TEMP 
         Height          =   315
         Left            =   6960
         TabIndex        =   22
         Top             =   465
         Visible         =   0   'False
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   1
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit TXT_ColStaTemp 
         Height          =   315
         Left            =   12840
         TabIndex        =   9
         Top             =   630
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1773
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   1
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit TXT_ColEndTemp 
         Height          =   315
         Left            =   13890
         TabIndex        =   10
         Top             =   630
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1773
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   1
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   6285
         Top             =   630
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_HEAT_RATIO 
         Height          =   315
         Left            =   7515
         TabIndex        =   7
         Top             =   630
         Width           =   690
         _Version        =   262145
         _ExtentX        =   1217
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   1
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   2
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   8820
         Top             =   150
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "��¯�ٶ�"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_SP_CHARGE 
         Height          =   315
         Left            =   10050
         TabIndex        =   2
         Top             =   150
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   1
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   3
         MaxValue        =   999.99
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   8805
         Top             =   630
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "��¯�ٶ�"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_SP_DISCHARGE 
         Height          =   315
         Left            =   10035
         TabIndex        =   8
         Top             =   630
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   1
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   4
         MaxValue        =   9999.99
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Left            =   11130
         Top             =   150
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "�����ٶ�"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_SP_CAL 
         Height          =   315
         Left            =   12360
         TabIndex        =   3
         Top             =   150
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   1
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   3
         MaxValue        =   999.99
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
   End
End
Attribute VB_Name = "DGA1030C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   HTM System
'-- Program Name      �ȴ�����¯��ҵʵ����ѯ���޸�
'-- Program ID        DGA1030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2007.11.20
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String            'Form Type
Public Toolbar_St As String          'Active Form ToolBar Setting
Public sAuthority As String          'Active Form Authority Setting
Public sDateTime As String           'Active Form Time Setting
Public sQuery_Rt As String

Dim pControl1 As New Collection      'Master Primary Key Collection
Dim nControl1 As New Collection      'Master Necessary Collection
Dim mControl1 As New Collection      'Master Maxlength check Collection
Dim iControl1 As New Collection      'Master Insert Collection
Dim rControl1 As New Collection      'Master Refer Collection
Dim cControl1 As New Collection      'Master Copy Collection
Dim aControl1 As New Collection      'Master -> Spread Collection
Dim lControl1 As New Collection      'Master Lock Collection

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection


Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2



Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(txt_PrcLine, "p", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(cbo_chg_no, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_mat_no, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
    'MASTER Collection
     Mc1.Add Item:="DGA1030C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="DGA1030C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"
          
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
 
   'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="DGA1030C.P_REFER1", Key:="P-R"
    sc1.Add Item:="DGA1030C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="DGA1030C.P_SONEROW", Key:="P-O"
    
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(SDB_WAT_TEMP, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(TXT_DISCHARGE_TIME, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(txt_HeatTemp, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(TXT_ColStaTemp, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(TXT_ColEndTemp, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(txt_DisCharTemp, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
               Call Gp_Ms_Collection(TXT_SHIFT, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
               Call Gp_Ms_Collection(TXT_GROUP, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                 Call Gp_Ms_Collection(sUserID, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            
    'MASTER Collection
     Mc2.Add Item:="DGA1030C.P_MODIFY", Key:="P-M"
     Mc2.Add Item:="DGA1030C.P_REFER", Key:="P-R"
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"
     
    Call Gp_Sp_ColHidden(ss1, 1, True)
'    Call Gp_Sp_ColHidden(ss1, 6, True)
'    Call Gp_Sp_ColHidden(ss1, 21, True)
    Call Gp_Sp_ColHidden(ss1, 22, True)
    Call Gp_Sp_ColHidden(ss1, 23, True)
    Call Gp_Sp_ColHidden(ss1, 24, True)
    Call Gp_Sp_ColHidden(ss1, 27, True)
    Call Gp_Sp_ColHidden(ss1, 28, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
     
End Sub
Public Sub Form_Ref()


    Dim iRow As Integer
    Dim iCol As Integer
    Dim I As Integer

On Error GoTo Refer_Err

    If txt_plt <> "C1" And txt_plt <> "C2" And txt_plt <> "C3" Then
         Call Gp_MsgBoxDisplay("ֻ�ܲ�ѯ����ΪC1\C2\C3������")
         Exit Sub
    End If

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1) Then
        If ss1.MaxRows > 0 Then
            ss1.Row = 1
            ss1.Col = 1
            Call Gf_Ms_Refer(M_CN1, Mc1, Nothing, Nothing, False)
            'Call Gp_Ms_ControlLock(Mc1("pControl"), True)
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            
           For iRow = 1 To ss1.MaxRows
    
               ss1.Row = iRow
               ss1.Col = 33
                If ss1.Text = "Y" Then
                  For I = 1 To ss1.MaxCols
                       ss1.Col = I
                       ss1.ForeColor = &HC000&
                  Next
                End If
          
          Next iRow
        End If
    End If
            
    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False                 'Row Cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False                'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False                'Paste
    MDIMain.MenuTool.Buttons(14).Enabled = True                 'Excel
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
End Sub



Private Sub cbo_PrcLine_Click()
If cbo_PrcLine.ListIndex = 0 Then
    txt_PrcLine = "1"
ElseIf cbo_PrcLine.ListIndex = 1 Then
    txt_PrcLine = "2"
ElseIf cbo_PrcLine.ListIndex = 2 Then
    txt_PrcLine = "3"
    
End If
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim iRow As Integer
    Dim iCol As Integer
    Dim I As Integer


    
    If Row < 0 Then Exit Sub
    
If Col = 0 Then
    If Mid(TXT_DISCHARGE_TIME, 1, 1) <> "2" Then
        MsgBox "����ȷ�ϳ�¯ʱ��......!"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If txt_DisCharTemp.Value = 0 Then
       MsgBox "��ȷ�ϳ�¯�¶�....!"
       Screen.MousePointer = vbDefault
       Exit Sub
    End If

    ss1.Row = Row
    ss1.Col = 0
    
    If ss1.Text = "Update" Then
        ss1.Text = Row
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
        Call Gp_Sp_BlockColor(ss1, 1, 2, Row, Row, , &HC0FFFF)
        Call Gp_Sp_BlockColor(ss1, 5, 10, Row, Row, , &HC0FFFF)
        ss1.Col = 1
        ss1.Text = ""
        ss1.Col = 5
        ss1.Text = ""
        ss1.Col = 6
        ss1.Text = ""
        ss1.Col = 7
        ss1.Text = ""
        ss1.Col = 8
        ss1.Text = ""
        ss1.Col = 9
        ss1.Text = ""
        ss1.Col = 10
        ss1.Text = ""
        ss1.Col = 11
        ss1.Text = ""
        ss1.Col = 12
        ss1.Text = ""

    Else
        ss1.Text = "Update"
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)

        ss1.Col = 1
        ss1.Text = "1"
        
        ss1.Col = 6
        If SDB_WAT_TEMP.Value > 0 Then
            ss1.Text = SDB_WAT_TEMP.Value
        End If
        
        ss1.Col = 7
        ss1.Text = TXT_DISCHARGE_TIME
        
        ss1.Col = 8
        If txt_HeatTemp.Value > 0 Then
            ss1.Text = txt_HeatTemp.Value
        End If

        ss1.Col = 9
        If txt_DisCharTemp.Value > 0 Then
            ss1.Text = txt_DisCharTemp.Value
        End If
        
        ss1.Col = 10
        If TXT_ColStaTemp.Value > 0 Then
            ss1.Text = TXT_ColStaTemp.Value
        End If
        
        ss1.Col = 11
        If TXT_ColEndTemp.Value > 0 Then
            ss1.Text = TXT_ColEndTemp.Value
        End If
        
        ss1.Col = 12
        ss1.Text = Trim(TXT_SHIFT.Text)
        ss1.Col = 13
        ss1.Text = Trim(TXT_GROUP.Text)
        ss1.Col = 14
        ss1.Text = sUserID
        ss1.Col = 21
        ss1.Text = Trim(cbo_chg_no.Text)
        
        ss1.Col = 29
        ss1.Text = txt_HEAT_RATIO.Value
        
        ss1.Col = 30
        ss1.Text = txt_SP_CHARGE.Value
        
        ss1.Col = 31
        ss1.Text = txt_SP_CAL.Value
        
        ss1.Col = 32
        ss1.Text = txt_SP_DISCHARGE.Value

    End If
    
     For iRow = 1 To ss1.MaxRows
    
        ss1.Row = iRow
        ss1.Col = 33
          If ss1.Text = "Y" Then
            For I = 1 To ss1.MaxCols
                 ss1.Col = I
                 ss1.ForeColor = &HC000&
            Next
          End If
       
    Next iRow
    
End If

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
'    If Row < 0 Then Exit Sub
'    ss1.Row = Row
'    If Col = 6 Then
'        ss1.Col = 7
'        ss1.Text = Gf_DTSet(M_CN1, , "X")
'    End If
'
'    If Col = 10 Then
'        ss1.Col = 10
'        ss1.Text = Gf_ShiftSet3(M_CN1)
'    End If
'
'    If Col = 11 Then
'        ss1.Col = 11
'        ss1.Text = Gf_GroupSet(M_CN1, Trim(TXT_SHIFT), Gf_DTSet(M_CN1, , "X"))
'    End If
'
'    If Col = 12 Then
'        ss1.Col = 12
'        ss1.Text = sUserID
'    End If
        
End Sub
Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
        MDIMain.MenuTool.Buttons(7).Enabled = False                 'Row Insert
        MDIMain.MenuTool.Buttons(8).Enabled = False                 'Row Delete
        MDIMain.MenuTool.Buttons(9).Enabled = False                 'Row Cancel
        MDIMain.MenuTool.Buttons(11).Enabled = False                'Copy
        MDIMain.MenuTool.Buttons(12).Enabled = False                'Paste
        MDIMain.MenuTool.Buttons(14).Enabled = True                 'Excel
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "DG-System.INI", Me.Name)
    
    cbo_PrcLine.AddItem "һ���ȴ���"
    cbo_PrcLine.AddItem "�����ȴ���"
    cbo_PrcLine.AddItem "�����ȴ���"
    cbo_PrcLine.ListIndex = 0
    
    ''''''ADDED BY GUOLI AT 20080904133500''''
    cbo_chg_no.AddItem "1"
    cbo_chg_no.AddItem "2"
    cbo_chg_no.Text = "1"
    
    TXT_DISCHARGE_TIME = Gf_DTSet(M_CN1, , "X")
    TXT_SHIFT = Gf_ShiftSet(M_CN1)
    TXT_GROUP = Gf_GroupSet(M_CN1, Trim(TXT_SHIFT), Gf_DTSet(M_CN1, , "X"))
    TXT_EMP = sUserID

          
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "DG-System.INI", Me.Name)

    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub
Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()

    Dim SMESG As String

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))

    Call Gf_Sp_Cls(sc1)

    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = False                'Excel
    End With
    
    TXT_DISCHARGE_TIME = Gf_DTSet(M_CN1, , "X")
    TXT_SHIFT = Gf_ShiftSet3(M_CN1)
    TXT_GROUP = Gf_GroupSet(M_CN1, Trim(TXT_SHIFT), Gf_DTSet(M_CN1, , "X"))
    TXT_EMP = sUserID
    
    pControl1(1).SetFocus
    
End Sub

Private Sub TXT_DISCHARGE_TIME_Change()
Dim for_cnt As Integer

    TXT_SHIFT = Gf_ShiftSet(M_CN1, Mid(TXT_DISCHARGE_TIME.RawData, 9, 4))
    TXT_GROUP = Gf_GroupSet(M_CN1, Trim(TXT_SHIFT), Gf_DTSet(M_CN1, , "X"))
    
    For for_cnt = 1 To ss1.MaxRows
        ss1.Row = for_cnt
        ss1.Col = 0
        If ss1.Text = "Input" Or ss1.Text = "Update" Then
            ss1.Col = 7
            ss1.Text = TXT_DISCHARGE_TIME
            ss1.Col = 12
            ss1.Text = TXT_SHIFT
            ss1.Col = 13
            ss1.Text = TXT_GROUP
            ss1.Col = 14
            ss1.Text = sUserID
        End If
    Next
End Sub

Private Sub TXT_DISCHARGE_TIME_DblClick()
    TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1, , "X")
End Sub

Private Sub txt_Plt_Change()
    If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
        TXT_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        TXT_PLT_NAME.Text = ""
    End If
End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub