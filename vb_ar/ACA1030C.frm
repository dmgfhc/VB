VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACA1030C 
   Caption         =   "订单进程详细查询_ACA1030C"
   ClientHeight    =   8595
   ClientLeft      =   150
   ClientTop       =   1935
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin Threed.SSCommand CMD_PRODEND 
      Height          =   420
      Left            =   7740
      TabIndex        =   23
      Top             =   45
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   741
      _Version        =   196609
      ForeColor       =   16711680
      Caption         =   "生产结束"
   End
   Begin VB.TextBox txt_prod_cd 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   4980
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   17
      Tag             =   "订单号"
      Top             =   95
      Width           =   735
   End
   Begin Threed.SSCommand cmd_OS 
      Height          =   420
      Left            =   12150
      TabIndex        =   15
      Top             =   45
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   741
      _Version        =   196609
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "OS再整理"
   End
   Begin Threed.SSCommand cmd_MPState 
      Height          =   420
      Left            =   10665
      TabIndex        =   14
      Top             =   45
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   741
      _Version        =   196609
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "物料进程状态"
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Tag             =   "订单号"
      Top             =   90
      Width           =   645
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1560
      Left            =   30
      TabIndex        =   3
      Top             =   510
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   2752
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_REASON 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   12285
         MultiLine       =   -1  'True
         TabIndex        =   22
         Tag             =   "订单号"
         Top             =   1095
         Width           =   2625
      End
      Begin CSTextLibCtl.sitxEdit txt_sms_duedate 
         Height          =   315
         Left            =   5325
         TabIndex        =   18
         Top             =   120
         Width           =   1335
         _Version        =   262145
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   "____-__-__"
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
         Mask            =   "____-__-__"
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin VB.TextBox Txt_prod_end_fl 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12285
         MaxLength       =   1
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "ORD"
         Top             =   600
         Width           =   315
      End
      Begin VB.TextBox Txt_mill_plt 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2130
         MaxLength       =   2
         TabIndex        =   9
         Tag             =   "ORD"
         Top             =   90
         Width           =   300
      End
      Begin VB.TextBox Txt_sms_plt 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "ORD"
         Top             =   90
         Width           =   300
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   10575
         Top             =   600
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "要生产结束的对象"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   135
         Top             =   1095
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "进程总量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_ord_prc_wgt 
         Height          =   315
         Left            =   1845
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1095
         Width           =   1500
         _Version        =   262145
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
         Text            =   " 0.000"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   7080
         Top             =   600
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "订单欠量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_ord_rem_wgt 
         Height          =   315
         Left            =   8775
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   600
         Width           =   1500
         _Version        =   262145
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
         Text            =   " 0.000"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   3615
         Top             =   1095
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "发货完毕重量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_ship_END_wgt 
         Height          =   315
         Left            =   5340
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1500
         _Version        =   262145
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
         Text            =   " 0.000"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   135
         Top             =   90
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "炼钢/轧钢厂"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   7080
         Top             =   1095
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "轧钢欠重量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit TXT_MILL_REM_WGT 
         Height          =   315
         Left            =   8775
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1095
         Width           =   1500
         _Version        =   262145
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
         Text            =   " 0.000"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   3615
         Top             =   600
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "成品生产结束重量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   3615
         Top             =   120
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "炼钢作业期限"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Left            =   7080
         Top             =   120
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "轧钢作业期限"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   10575
         Top             =   120
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "交货期拖延天数"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit TXT_DEL_DELAY_DAY 
         Height          =   315
         Left            =   12285
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   120
         Width           =   480
         _Version        =   262145
         _ExtentX        =   847
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   135
         Top             =   600
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "订单量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sitxEdit txt_mill_duedate 
         Height          =   315
         Left            =   8775
         TabIndex        =   19
         Top             =   120
         Width           =   1335
         _Version        =   262145
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Text            =   "____-__-__"
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
         Mask            =   "____-__-__"
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sidbEdit TXT_PROD_END_WGT 
         Height          =   315
         Left            =   5325
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   600
         Width           =   1500
         _Version        =   262145
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
         Text            =   " 0.000"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit TXT_TOT_WGT 
         Height          =   315
         Left            =   1845
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   630
         Width           =   1500
         _Version        =   262145
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
         Text            =   " 0.000"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   10575
         Top             =   1095
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "订单返回原因"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "天"
         Height          =   330
         Left            =   12825
         TabIndex        =   10
         Top             =   180
         Width           =   330
      End
   End
   Begin VB.TextBox txt_ord_no 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   310
      Left            =   1845
      MaxLength       =   11
      TabIndex        =   0
      Tag             =   "订单号"
      Top             =   95
      Width           =   1350
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   165
      Top             =   90
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      Caption         =   "订单号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmd_fl_down 
      Height          =   420
      Left            =   13650
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   45
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "订单返回"
   End
   Begin FPSpread.vaSpread SS1 
      Height          =   7020
      Left            =   30
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2145
      Width           =   15105
      _Version        =   393216
      _ExtentX        =   26644
      _ExtentY        =   12382
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   161
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACA1030C.frx":0000
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Left            =   9195
      TabIndex        =   16
      Top             =   45
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   741
      _Version        =   196609
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "订单进程详细"
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   4005
      Top             =   90
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      Caption         =   "产品"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   3135
      X2              =   3315
      Y1              =   210
      Y2              =   210
   End
End
Attribute VB_Name = "ACA1030C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name
'-- Program Name
'-- Program ID        ACA1030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          APPLE
'-- Coder             APPLE
'-- Date              2003.8.4
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
Public WULIAO As String
Public Active_CForm As String       'Form Active

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

Dim ordmark As String
    
    
Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Refer"  ' "PopMaster"

             'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(txt_ord_no, "p", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(Combo1, "p", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_prod_cd, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_TOT_WGT, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(Txt_sms_plt, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(Txt_mill_plt, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_sms_duedate, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_mill_duedate, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(TXT_DEL_DELAY_DAY, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(Txt_prod_end_fl, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_ord_prc_wgt, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_ord_rem_wgt, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_ship_END_wgt, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_MILL_REM_WGT, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_PROD_END_WGT, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
   
    'MASTER Collection
    Mc1.Add Item:="ACA1030C.P_REFER", Key:="P-R"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
     'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub




Private Sub cmd_fl_down_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim sMesg As String
    
    
    Dim adoCmd As adodb.Command
    If TXT_REASON = "" Then
       sMesg = "请输入订单返回原因"
       Call Gp_MsgBoxDisplay(sMesg)
       TXT_REASON.SetFocus
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    

    sQuery = "{call ACA1031P ('" + txt_ord_no + "', '" + Combo1.Text + "','" + TXT_REASON + "','" + sUsername + "',?)}"

    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Exit Sub
    Else
        Call MsgBox("返回完了！", vbInformation, "系统提示信息")
        Call Form_Ref
        
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    ERR.Raise ERR.Number, ERR.Description & sQuery
    
End Sub


Private Sub cmd_OS_Click()
    
    Call Gp_CallACB3050P
End Sub

Private Sub CMD_PRODEND_Click()
On Error GoTo PRODEND_Error

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
   
    Dim adoCmd As adodb.Command
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    If ordmark = "" Then
        Call MsgBox("该订单不符合生产结束或恢复条件！", vbInformation, "系统提示信息")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
              
              
    sQuery = "{call ACA1032P('" + txt_ord_no.Text + "','" + Combo1.Text + "','" + ordmark + "','" + sUserID + "',?)}"

      
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Execute , , adExecuteNoRecords
    
    'OS Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
   
        Call MsgBox(CMD_PRODEND.Caption + "完成！", vbInformation, "系统提示信息")
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

PRODEND_Error:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("PRODEND_Error: " & Error)
    

End Sub

Private Sub Form_Activate()
     
    If Active_CForm <> "" Then
        Call Form_Ref
        Active_CForm = ""
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    'Call Gf_Ms_Refer(M_CN1, Mc1)
    
'    CMD_PRODEND.Visible = False
    txt_prod_cd.Enabled = False
    
    Select Case Mid(sAuthority, 2, 3) 'Insert, Update, Delete

           Case "000"      'No Authority
             CMD_PRODEND.Enabled = False
             cmd_fl_down.Enabled = False
             ULabel9.Enabled = False
             TXT_REASON.Enabled = False
    End Select

    CMD_PRODEND.Visible = False
         

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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
        Combo1.Clear
    End If
    CMD_PRODEND.Visible = False
End Sub

Public Sub Form_Ref()

    Dim sQuery As String
    Dim sMesg As String
     Dim STSCODE As String

    
  '  squery = "{call ACB3050P ('" + txt_ord_no.Text + "','" + Combo1.Text + "',?)}"
  sQuery = "SELECT REC_STS||ORD_STS FROM BP_ORDER_ITEM WHERE ORD_NO = '" + txt_ord_no.Text + "' AND ORD_ITEM = '" + Combo1.Text + "' "
        STSCODE = Gf_CodeFind(M_CN1, sQuery)
        sQuery = ""
        If STSCODE = "2E" Then
            ordmark = "E"
            CMD_PRODEND.Caption = "生产结束"
            CMD_PRODEND.Visible = True
        ElseIf STSCODE = "2F" Then
            ordmark = "R"
            CMD_PRODEND.Caption = "生产恢复"
            CMD_PRODEND.Visible = True
        Else
            ordmark = ""
        End If
  
    
    sQuery = "SELECT "
    
    sQuery = sQuery + " A.CD_NAME,"
    'sQuery = sQuery + " B.PRC,"
    'sQuery = sQuery + "B.PRE_WGT, "
    sQuery = sQuery + "B.TOT_WGT, "
    
    sQuery = sQuery + "B.INS_WGT, "
    
    sQuery = sQuery + "B.WRK_WGT, "
    
    sQuery = sQuery + "B.EST_WGT, "
    
    sQuery = sQuery + "B.REP_WGT, "
    
    sQuery = sQuery + "B.HLD_WGT,"
    
    sQuery = sQuery + "B.END_WGT "
    
    sQuery = sQuery + " FROM  nisco.CP_PRC_DET B ,nisco.zp_cd A "
    'sQuery = sQuery + " FROM  nisco.CP_PRC_DET B "
    sQuery = sQuery + " Where B.ORD_NO = '" + Trim(txt_ord_no.Text) + "'"
    
    sQuery = sQuery + " AND B.ORD_ITEM = '" + Trim(Combo1.Text) + "'"
     sQuery = sQuery + " AND B.PRC = A.CD AND A.CD_MANA_NO= 'C0002' "
    sQuery = sQuery + " ORDER BY B.PRC "
  '  Text1.Text = sQuery
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
    
        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then

            If Not Gf_Only_Display(M_CN1, Proc_Sc("Sc"), sQuery) Then
                ss1.OperationMode = OperationModeNormal
                Exit Sub
                'Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Else
                ss1.OperationMode = OperationModeNormal
            End If
            
        Else
            sMesg = sMesg + " Must input according to length of item"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
    
    Else
        sMesg = sMesg + " Must input necessarily"
        Call Gp_MsgBoxDisplay(sMesg)
    End If
 
   If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       ' Call GP_BACKCOLOR_WHITE(Mc1("rControl"))
        Call Gp_Ms_NeceColor(Mc1("nControl"))
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    End If



End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
    
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()
    
    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

Dim ORD_STS As String
Dim ord_sts_row As String
Dim ord_sts_col  As String
Dim STR_ROW As String                      '每行的第一列的文字

If Col < 2 Or Col > 6 Or Row < 1 Then Exit Sub
ss1.Col = 1
ss1.Row = Row
STR_ROW = ss1.Text

Select Case Col

        Case 2
            ord_sts_col = ""
        Case 3
            ord_sts_col = "A"
        Case 4
            ord_sts_col = "B"
        Case 5
            ord_sts_col = "D"
        Case 6
            ord_sts_col = "C"
        Case Else
             Exit Sub

    End Select
    
Select Case STR_ROW
        Case "倒罐站"
            ord_sts_row = "BA"
        Case "铁水预处理"
            ord_sts_row = "BB"
        Case "转炉"
            ord_sts_row = "BC"
        Case "LF"
            ord_sts_row = "BD"
        Case "VD"
            ord_sts_row = "BE"
        Case "连铸"
            ord_sts_row = "BF"
        Case "QUALITY"
            ord_sts_row = "QA"
        Case "SHIPPING"
            ord_sts_row = "XA"
        Case "加热炉"
            ord_sts_row = "CA"
        Case "定尺剪"
            ord_sts_row = "CG"
        Case "REPAIR"
            ord_sts_row = "DA"
        Case "UST"
            ord_sts_row = "DB"
        Case Else
            Exit Sub
End Select

ORD_STS = ord_sts_row + ord_sts_col

    If Trim(txt_ord_no.Text) <> "" And Trim(Combo1.Text) <> "" Then
        Unload ACA1040C
        Load ACA1040C

        ACA1040C.Text_BB_ORD_NO.Text = Trim(txt_ord_no.Text)
        ACA1040C.Combo_ORD_ITEM.Text = Trim(Combo1.Text)
        ACA1040C.Text_PROC_CD.Text = ORD_STS


        ACA1040C.Active_CForm = "ACA1040C"
        ACA1040C.Show
        ACA1040C.SetFocus
    End If





End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub SELECT_PRC()
         
    If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       ' Call GP_BACKCOLOR_WHITE(Mc1("rControl"))
        Call Gp_Ms_NeceColor(Mc1("nControl"))
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    End If

End Sub

Private Sub cmd_MPState_Click()
    If Trim(txt_ord_no.Text) <> "" And Trim(Combo1.Text) <> "" Then
        Unload ACA1040C
        Load ACA1040C
        
        ACA1040C.Text_BB_ORD_NO.Text = Trim(txt_ord_no.Text)
        ACA1040C.Combo_ORD_ITEM.Text = Trim(Combo1.Text)
    
        
        ACA1040C.Active_CForm = "ACA1040C"
        ACA1040C.Show
        ACA1040C.SetFocus
    End If
End Sub

Private Sub SSCommand1_Click()

  Load ACA1031C
  ACA1031C.txt_ord_no = txt_ord_no
  ACA1031C.txt_ORD_ITEM = Combo1.Text
  ACA1031C.Show
  ACA1031C.Form_Ref
  

End Sub

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String

    If Len(Trim(txt_ord_no.Text)) = txt_ord_no.MaxLength Then

        If Combo1.Text <> "" Then Exit Sub
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(txt_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, Combo1, sQuery)
        
        'If Combo1.ListCount <> 0 Then
        '      Combo1.ListIndex = 0
        'End If
    Else
        Combo1.Clear
    End If
  
End Sub

Private Function JIA(AA As String) As String

    Dim A1 As String
    Dim A2 As String
    Dim A3 As String
    
    If IsNull(AA) Or AA = " " Then
        JIA = ""
        Exit Function
    End If
    
    A1 = Mid(AA, 1, 4) + "-"
    A2 = Mid(AA, 5, 2) + "-"
    A3 = A1 + A2 + Mid(AA, 7, 2)
    
    JIA = A3

End Function
Public Sub Gp_CallACB3050P()

On Error GoTo Gp_CallACB3050P_Error

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As adodb.Command
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call ACB3050P ('" + txt_ord_no.Text + "','" + Combo1.Text + "',?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Execute , , adExecuteNoRecords
    
    'OS Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Gp_CallACB3050P_Error:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Gp_CallACB3050P_Error : " & Error)
    
End Sub

