VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form EGA1140C 
   Caption         =   "ѹƽ��ʵ����ѯ���޸�_EGA1140C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread ss1 
      Height          =   7575
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1650
      Width           =   15210
      _Version        =   393216
      _ExtentX        =   26829
      _ExtentY        =   13361
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
      MaxCols         =   25
      MaxRows         =   1
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "EGA1140C.frx":0000
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   660
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   1164
      _Version        =   196609
      BackColor       =   14737632
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
         Left            =   10035
         TabIndex        =   5
         Top             =   180
         Width           =   1620
      End
      Begin VB.TextBox txt_Prc 
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
         Left            =   14610
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "����ñ��"
         Text            =   "P"
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
         ItemData        =   "EGA1140C.frx":0CB8
         Left            =   1185
         List            =   "EGA1140C.frx":0CBA
         TabIndex        =   3
         Tag             =   "¯����"
         Top             =   180
         Width           =   1410
      End
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
         Left            =   3705
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "����"
         Top             =   390
         Visible         =   0   'False
         Width           =   210
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   120
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "���߱�"
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
         Left            =   3885
         Top             =   180
         Width           =   1170
         _ExtentX        =   2064
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
      Begin InDate.UDate udt_date_fr 
         Height          =   315
         Left            =   5100
         TabIndex        =   6
         Tag             =   "������"
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.74
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         BackColor       =   16777215
         MaxLength       =   10
      End
      Begin InDate.UDate udt_date_to 
         Height          =   315
         Left            =   6570
         TabIndex        =   7
         Tag             =   "������"
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.74
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         BackColor       =   16777215
         MaxLength       =   10
      End
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   8970
         Top             =   180
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         Caption         =   "��ѯ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   945
      Left            =   90
      TabIndex        =   8
      Top             =   690
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   1667
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
         Left            =   13920
         MaxLength       =   1
         TabIndex        =   10
         Top             =   510
         Width           =   900
      End
      Begin VB.ComboBox cbo_HL_NO 
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
         ItemData        =   "EGA1140C.frx":0CBC
         Left            =   1245
         List            =   "EGA1140C.frx":0CBE
         TabIndex        =   9
         Tag             =   "¯����"
         Top             =   525
         Width           =   900
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   9795
         Top             =   135
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "ѹƽ����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   3885
         Top             =   135
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "��ڲ�ѹ����"
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   0
         Left            =   6990
         Top             =   135
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "��б��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   12795
         Top             =   510
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_RETE_PRS_DW_RTO 
         Height          =   315
         Left            =   5355
         TabIndex        =   11
         Top             =   135
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
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
         NumIntDigits    =   3
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit TXT_LEV_SPLT 
         Height          =   315
         Left            =   8220
         TabIndex        =   12
         Top             =   135
         Width           =   930
         _Version        =   262145
         _ExtentX        =   1640
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
      Begin CSTextLibCtl.sidbEdit TXT_PASS_CNT 
         Height          =   315
         Left            =   10905
         TabIndex        =   13
         Top             =   135
         Width           =   930
         _Version        =   262145
         _ExtentX        =   1640
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
         NumIntDigits    =   1
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   3885
         Top             =   525
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "���ڲ�ѹ����"
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
      Begin CSTextLibCtl.sidbEdit PRS_OT_PRS_DW_RTO 
         Height          =   315
         Left            =   5355
         TabIndex        =   14
         Top             =   525
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
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
         NumIntDigits    =   3
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   120
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "ѹƽʱ��"
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
         Left            =   1245
         TabIndex        =   15
         Tag             =   "��¯ʱ��"
         Top             =   135
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   120
         Top             =   525
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "������"
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   1
         Left            =   6990
         Top             =   525
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "ѹƽ�ٶ�"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin CSTextLibCtl.sidbEdit TXT_HL_SPD 
         Height          =   315
         Left            =   8220
         TabIndex        =   16
         Top             =   525
         Width           =   930
         _Version        =   262145
         _ExtentX        =   1640
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
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   4
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   2
         Left            =   9795
         Top             =   510
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "ѹƽ�¶�"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin CSTextLibCtl.sidbEdit TXT_HL_TMP 
         Height          =   315
         Left            =   10905
         TabIndex        =   17
         Top             =   510
         Width           =   930
         _Version        =   262145
         _ExtentX        =   1640
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
   End
End
Attribute VB_Name = "EGA1140C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   HTM System
'-- Program Name      �а��ȴ���ѹƽ��ʵ������
'-- Program ID        EGA1140C
'-- Designer          ZHANGJINBO
'-- Coder             ZHANGJINBO
'-- Date              2011.6.07
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
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
            Call Gp_Ms_Collection(txt_Prc, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_PrcLine, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(udt_date_fr, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(udt_date_to, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="EGA1140C.P_REFER", Key:="P-R"
    sc1.Add Item:="EGA1140C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="EGA1140C.P_SONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Call Gp_Sp_ColHidden(ss1, 21, True)
    Call Gp_Sp_ColHidden(ss1, 22, True)
    Call Gp_Sp_ColHidden(ss1, 23, True)
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub cbo_PrcLine_Click()
    If cbo_PrcLine.ListIndex = 0 Then
        txt_PrcLine = "1"
        cbo_hl_no.Clear
        cbo_hl_no.List(0) = 1
        cbo_hl_no.Text = 1

       
    ElseIf cbo_PrcLine.ListIndex = 1 Then
        txt_PrcLine = "2"
      
        cbo_hl_no.Clear
        cbo_hl_no.List(0) = 1
        cbo_hl_no.Text = 1
        
    End If
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
    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_Sp_Setting(sc1.Item("Spread"), False)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "EG-System.INI", Me.Name)
    
    cbo_PrcLine.AddItem "һ����"
    cbo_PrcLine.AddItem "������"
    cbo_PrcLine.ListIndex = 0
    

    
    TXT_SHIFT = Gf_ShiftSet(M_CN1)
    
    udt_date_fr.RawData = Gf_DTSet(M_CN1, "D", "X")
    udt_date_to.RawData = Gf_DTSet(M_CN1, "D", "X")
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    Call Gp_Sp_ColSet(sc1.Item("Spread"), "EG-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set pColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set iColumn1 = Nothing
    Set aColumn1 = Nothing
    Set lColumn1 = Nothing
    
    Set Mc1 = Nothing
    
    Set sc1 = Nothing
    
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)

End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
        
End Sub

Public Sub Spread_Forzens_Setting()

    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
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

Public Sub Form_Exit()
    Unload Me
End Sub


Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    
    Dim iRow As Integer
    
    
    If ROW < 0 Then Exit Sub
    
    If Mid(TXT_DISCHARGE_TIME, 1, 1) <> "2" Then
        MsgBox "����ȷ����ҵʱ��......!", vbCritical, "ϵͳ��ʾ��Ϣ"
        Exit Sub
    End If
        
    If Col = 0 Then
        
        ss1.ROW = ROW
        ss1.Col = 0
        If ss1.Text = "Update" Then
            ss1.Text = ROW
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW)
            Call Gp_Sp_BlockColor(ss1, 1, 9, ROW, ROW, , &HC0FFFF)
    
            ss1.Col = 2
            ss1.Text = ""
            ss1.Col = 3
            ss1.Text = ""
            ss1.Col = 4
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
            ss1.Text = 21
            ss1.Text = ""
            ss1.Col = 22
            ss1.Text = ""
            ss1.Col = 23
            ss1.Text = ""
    
        Else
            ss1.ROW = ROW
            ss1.Col = 0
            ss1.Text = "Update"
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW, , &HFFFF80)
    
            ss1.Col = 2
            If txt_RETE_PRS_DW_RTO.Value > 0 Then
                ss1.Text = txt_RETE_PRS_DW_RTO.Value
            End If
            ss1.Col = 3
            If PRS_OT_PRS_DW_RTO.Value > 0 Then
                ss1.Text = PRS_OT_PRS_DW_RTO.Value
            End If
            ss1.Col = 4
            If TXT_LEV_SPLT.Value > 0 Then
                ss1.Text = TXT_LEV_SPLT.Value
            End If
            ss1.Col = 5
            If TXT_PASS_CNT.Value > 0 Then
                ss1.Text = TXT_PASS_CNT.Value
            End If
            
            ss1.Col = 6
            ss1.Text = TXT_DISCHARGE_TIME
            
            ss1.Col = 7
            If TXT_HL_SPD.Value > 0 Then
                ss1.Text = TXT_HL_SPD.Value
            End If
            
            ss1.Col = 8
            If TXT_HL_TMP.Value > 0 Then
                ss1.Text = TXT_HL_TMP.Value
            End If
            
            ss1.Col = 9
            ss1.Text = TXT_SHIFT
            ss1.Col = 10
            ss1.Text = sUserID
            ss1.Col = 21
            ss1.Text = txt_Prc.Text
            ss1.Col = 22
            ss1.Text = txt_PrcLine.Text
            ss1.Col = 23
            ss1.Text = Trim(cbo_hl_no.Text)
        End If
            
    End If
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub
Public Sub Form_Ins()
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    
End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call Form_Ref
    
End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
End Sub
Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub TXT_DISCHARGE_TIME_Change()
Dim FOR_CNT As Integer

    TXT_SHIFT = Gf_ShiftSet(M_CN1, Mid(TXT_DISCHARGE_TIME.RawData, 9, 4))
    
    For FOR_CNT = 1 To ss1.MaxRows
        ss1.ROW = FOR_CNT
        ss1.Col = 0
        If ss1.Text = "Input" Or ss1.Text = "Update" Then
            ss1.Col = 6
            ss1.Text = TXT_DISCHARGE_TIME
            ss1.Col = 9
            ss1.Text = TXT_SHIFT
            ss1.Col = 10
            ss1.Text = sUserID
        End If
    Next
End Sub

Private Sub TXT_DISCHARGE_TIME_DblClick()
    TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1, , "X")
End Sub

Private Sub TXT_SHIFT_DblClick()
    TXT_SHIFT = Gf_ShiftSet3(M_CN1)
End Sub


