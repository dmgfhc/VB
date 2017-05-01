VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CKG2070C 
   Caption         =   "轧钢计划查询界面_CKG2070C"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   14790
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread ss1 
      Height          =   7155
      Left            =   120
      TabIndex        =   0
      Top             =   1740
      Width           =   15105
      _Version        =   393216
      _ExtentX        =   26644
      _ExtentY        =   12621
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
      MaxCols         =   30
      MaxRows         =   10
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CKG2070C.frx":0000
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2040
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   3598
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_FL_DATE 
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
         Left            =   6030
         TabIndex        =   40
         Top             =   1320
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox TXT_ALL_COUNT 
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
         Left            =   11460
         MaxLength       =   7
         TabIndex        =   39
         Top             =   1290
         Width           =   795
      End
      Begin VB.TextBox TXT_FL 
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
         Left            =   3870
         MaxLength       =   7
         TabIndex        =   37
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TXT_COUNT 
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
         Left            =   13530
         MaxLength       =   7
         TabIndex        =   33
         Top             =   1290
         Width           =   795
      End
      Begin VB.ComboBox COB_FL 
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
         ItemData        =   "CKG2070C.frx":0DB3
         Left            =   6450
         List            =   "CKG2070C.frx":0DC0
         TabIndex        =   31
         Top             =   930
         Width           =   1530
      End
      Begin VB.TextBox TXT_SIZE 
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
         Left            =   3330
         MaxLength       =   7
         TabIndex        =   30
         Top             =   1410
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "是否定尺"
         Height          =   315
         Left            =   2970
         TabIndex        =   29
         Top             =   1230
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox TXT_MILL_STLGRD 
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
         Left            =   5445
         MaxLength       =   12
         TabIndex        =   7
         Tag             =   "钢种"
         Top             =   540
         Width           =   3165
      End
      Begin VB.TextBox TXT_STLGRD 
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
         Left            =   5445
         MaxLength       =   12
         TabIndex        =   6
         Tag             =   "钢种"
         Top             =   120
         Width           =   1425
      End
      Begin VB.TextBox TXT_STLGRD_NAME 
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
         Left            =   6885
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "钢种(标准号)"
         Top             =   120
         Width           =   1725
      End
      Begin VB.TextBox TXT_SLAB_NO 
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
         Left            =   1365
         TabIndex        =   4
         Top             =   120
         Width           =   1485
      End
      Begin VB.ComboBox CBO_ORD_ITEM 
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
         Left            =   2760
         TabIndex        =   3
         Top             =   510
         Width           =   750
      End
      Begin VB.TextBox TXT_ORD_NO 
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
         Left            =   1365
         MaxLength       =   11
         TabIndex        =   2
         Tag             =   "CD_MANA_NO"
         Top             =   510
         Width           =   1380
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   120
         Top             =   120
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "板坯号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Left            =   4200
         Top             =   120
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "板坯钢种"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   120
         Top             =   510
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "订单号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   4200
         Top             =   540
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "轧制钢种"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   11730
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "厚度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_THK 
         Height          =   315
         Left            =   12765
         TabIndex        =   8
         Top             =   90
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
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
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   ""
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
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   11730
         Top             =   480
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "宽度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_WID 
         Height          =   315
         Left            =   12765
         TabIndex        =   9
         Top             =   480
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
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
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   ""
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
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_THK_TO 
         Height          =   315
         Left            =   13950
         TabIndex        =   10
         Top             =   90
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
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
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
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
         NumDecDigits    =   2
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_WID_TO 
         Height          =   315
         Left            =   13950
         TabIndex        =   11
         Top             =   480
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
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
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
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
         NumDecDigits    =   2
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   11730
         Top             =   870
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "长度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_LEN 
         Height          =   315
         Left            =   12765
         TabIndex        =   12
         Top             =   870
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
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
         FocusSelect     =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   ""
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
         NumDecDigits    =   1
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_LEN_TO 
         Height          =   315
         Left            =   13950
         TabIndex        =   13
         Top             =   870
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
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
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
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
         NumDecDigits    =   1
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   8670
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "板坯厚度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_SLAB_THK 
         Height          =   315
         Left            =   9705
         TabIndex        =   17
         Top             =   90
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
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
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   ""
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
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   8670
         Top             =   480
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "板坯宽度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_SLAB_WID 
         Height          =   315
         Left            =   9705
         TabIndex        =   18
         Top             =   480
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
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
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   ""
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
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_SLAB_THK_TO 
         Height          =   315
         Left            =   10710
         TabIndex        =   19
         Top             =   90
         Width           =   765
         _Version        =   262145
         _ExtentX        =   1349
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
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   ""
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
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_SLAB_WID_TO 
         Height          =   315
         Left            =   10710
         TabIndex        =   20
         Top             =   480
         Width           =   765
         _Version        =   262145
         _ExtentX        =   1349
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
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   ""
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
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   8670
         Top             =   870
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "板坯长度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_SLAB_LEN 
         Height          =   315
         Left            =   9705
         TabIndex        =   21
         Top             =   870
         Width           =   885
         _Version        =   262145
         _ExtentX        =   1561
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
         FocusSelect     =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   ""
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
         NumDecDigits    =   1
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_SLAB_LEN_TO 
         Height          =   315
         Left            =   10710
         TabIndex        =   22
         Top             =   870
         Width           =   765
         _Version        =   262145
         _ExtentX        =   1349
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
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   ""
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
         NumDecDigits    =   1
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   120
         Top             =   930
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "浇料时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.UDate SDT_IN_DATE_FROM 
         Height          =   315
         Left            =   1350
         TabIndex        =   26
         Tag             =   "起始日期"
         Top             =   930
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.74
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         BackColor       =   16777215
      End
      Begin InDate.UDate SDT_IN_DATE_TO 
         Height          =   315
         Left            =   3120
         TabIndex        =   27
         Tag             =   "起始日期"
         Top             =   930
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.74
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         BackColor       =   16777215
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   5010
         Top             =   930
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "是否备料"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSP1 
         Height          =   315
         Left            =   8640
         TabIndex        =   32
         Top             =   1290
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "紧急订单"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   12510
         Top             =   1290
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "备料块数"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption opt_no_fl 
         Height          =   330
         Left            =   2070
         TabIndex        =   34
         Top             =   1320
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "不备料"
      End
      Begin Threed.SSOption opt_fl 
         Height          =   330
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "备料"
      End
      Begin Threed.SSOption opt_cancel_fl 
         Height          =   330
         Left            =   1095
         TabIndex        =   36
         Top             =   1320
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "剔除"
      End
      Begin Threed.SSCommand CMD_CARD 
         Height          =   315
         Left            =   4230
         TabIndex        =   38
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   255
         Caption         =   "打印导出"
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   10440
         Top             =   1290
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "总块数"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   120
         Left            =   2940
         TabIndex        =   28
         Top             =   1080
         Width           =   195
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   120
         Left            =   10590
         TabIndex        =   25
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   120
         Left            =   10590
         TabIndex        =   24
         Top             =   210
         Width           =   195
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   120
         Left            =   10590
         TabIndex        =   23
         Top             =   1005
         Width           =   195
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   120
         Left            =   13800
         TabIndex        =   16
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   120
         Left            =   13800
         TabIndex        =   15
         Top             =   210
         Width           =   195
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   120
         Left            =   13800
         TabIndex        =   14
         Top             =   1005
         Width           =   195
      End
   End
End
Attribute VB_Name = "CKG2070C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   Mill System
'-- Program Name      冷热装指示查询
'-- Program ID        CKG2070C
'-- Document No       Q-00-0010(Specification)
'-- Designer          LiQian
'-- Coder             LiQian
'-- Date              2011.11.21
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
Dim Mode As String

'Public Complete As Boolean           'Move Status Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pColumnl As New Collection      'Spread Primary Key Collection
Dim nColumnl As New Collection      'Spread necessary Column Collection
Dim mColumnl As New Collection      'Spread Maxlength check Column Collection
Dim iColumnl As New Collection      'Spread Insert Column Collection
Dim aColumnl As New Collection      'Master -> Spread Column Collection
Dim lColumnl As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection

Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SPD_ROLL_MANA_NO = 1
Const SPD_ROLL_SLAB_SEQ = 2
Const SPD_ACTUAL_ROLLING_LEN = 3
Const SPD_SEQ_NO = 35
Const SPD_SLAB_EDT_SEQ = 36
Const SPD_FL = 4
Const SPD_URGNT_FL = 28
Const SPD_SLAB_NO = 6
Const SPD_FL1 = 5
Const SPD_CHOSE_FL = 29
Const SPD_LOC = 15
Const SPD_SIZE = 9
Const SPD_STLGRD = 11
Const SPD_MILL_STLGRD = 12
Const SPD_PROD_THK = 17
Const SPD_PROD_WID = 18
Const SPD_ORD_CNT = 28
Const SPD_DATE = 30

Private Sub Form_Define()
        
    Dim i As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

  'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(TXT_SLAB_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_STLGRD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_ORD_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(CBO_ORD_ITEM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
       '20140902 ADD BY LICHAO
   Call Gp_Ms_Collection(SDT_IN_DATE_FROM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDT_IN_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_MILL_STLGRD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_SIZE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_SLAB_THK, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_SLAB_THK_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_SLAB_WID, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_SLAB_WID_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_SLAB_LEN, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_SLAB_LEN_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDB_THK, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDB_THK_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDB_WID, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDB_WID_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDB_LEN, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDB_LEN_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(COB_FL, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)     '剔除
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)     '备料
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumnl, nColumnl, mColumnl, iColumnl, aColumnl, lColumnl)

    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="CKG2080C.P_REFER", Key:="P-R"
    Sc1.Add Item:="CKG2080C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=pColumnl, Key:="pColumn"
    Sc1.Add Item:=nColumnl, Key:="nColumn"
    Sc1.Add Item:=aColumnl, Key:="aColumn"
    Sc1.Add Item:=mColumnl, Key:="mColumn"
    Sc1.Add Item:=iColumnl, Key:="iColumn"
    Sc1.Add Item:=lColumnl, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=Sc1, Key:="Sc"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub Check1_Click()

    If Check1.Value = 1 Then
        TXT_SIZE.Text = "01"
    Else
        TXT_SIZE.Text = ""
    End If

End Sub

Private Sub CMD_CARD_Click()
    Call ExcelPrn_Pile(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End Sub

'Private Sub cmd_input_Click()
'    Dim iDx     As Long
'
'    With ss1
'
'        For iDx = 1 To .MaxRows
'            .ROW = iDx
'            .Col = 0:                .Text = "Update"
'        Next iDx
'
'    End With
'End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = True                 'Excel
    End With

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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    'Call Gp_Sp_ReadOnlySet(Sc1.Item("Spread"))
   
    Call Gf_Sp_Cls(Sc1)
    
    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "G-System.INI", Me.Name)
    
    Call Gp_Sp_ColHidden(ss1, SPD_ROLL_MANA_NO, True)
    Call Gp_Sp_ColHidden(ss1, SPD_ROLL_SLAB_SEQ, True)
    Call Gp_Sp_ColHidden(ss1, SPD_ACTUAL_ROLLING_LEN, True)
    Call Gp_Sp_ColHidden(ss1, SPD_SEQ_NO, True)
    Call Gp_Sp_ColHidden(ss1, SPD_SLAB_EDT_SEQ, True)
    Call Gp_Sp_ColHidden(ss1, SPD_FL1, True)
    Call Gp_Sp_ColHidden(ss1, SPD_CHOSE_FL, True)
'    Call Gp_Sp_ColHidden(ss1, SPD_DATE, True)
    
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Sc1.Item("Spread"), "G-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set iColumnl = Nothing
    Set pColumnl = Nothing
    Set lColumnl = Nothing
    Set nColumnl = Nothing
    Set mColumnl = Nothing
    Set aColumnl = Nothing
    
    Set Mc1 = Nothing

    Set Sc1 = Nothing
  
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Ref()

    Dim lRow As Long
    Dim sFl As Integer
    Dim urgnt As String
    Dim CNT As Integer
    
'     If Not Gp_DateCheck(SDT_PROD_DATE_FROM.Text, "S") Or Not Gp_DateCheck(SDT_PROD_DATE_TO.Text, "S") Then
'        Call Gp_MsgBoxDisplay("请输入生产时间")
'        Exit Sub
'     End If
     If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
     
     If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl")) Then
         ss1.OperationMode = OperationModeNormal
         Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
     End If
    
    Dim iRow    As Integer
    Dim s1num   As Integer

    With ss1

        For iRow = 1 To .MaxRows
            .ROW = iRow
            .Col = 0:                .Text = "Update"

            .ROW = iRow:            .Col = SPD_URGNT_FL:         urgnt = Trim(.Text)
            If urgnt = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, SPD_SLAB_NO, ss1.MaxCols, iRow, iRow, SSP1.BackColor)
            End If
            
            .ROW = iRow:            .Col = SPD_FL1:
            If .Value = 1 Then
                s1num = s1num + 1
            End If
                     
            If opt_fl.Value = True Then
               .ROW = iRow:  .Col = SPD_CHOSE_FL:   .Text = "1"
            ElseIf opt_cancel_fl.Value = True Then
               .ROW = iRow:  .Col = SPD_CHOSE_FL:   .Text = "2"
            ElseIf opt_no_fl.Value = True Then
               .ROW = iRow:  .Col = SPD_CHOSE_FL:   .Text = "3"
            End If

        Next iRow

    End With
    
        ss1.ROW = 1
        ss1.Col = SPD_DATE:                 TXT_FL_DATE.Text = Trim(ss1.Text)

        
    CNT = ss1.MaxRows
    TXT_ALL_COUNT = CNT
    TXT_COUNT = s1num
                
End Sub

Private Sub opt_cancel_fl_Click(Value As Integer)
Dim iRow As Integer
ss1.Col = SPD_CHOSE_FL
    If opt_cancel_fl.Value = True Then
        opt_cancel_fl.ForeColor = &HFF&
        opt_fl.ForeColor = &H808080
        opt_no_fl.ForeColor = &H808080
        TXT_FL.Text = "2"
    Else
        opt_cancel_fl.ForeColor = &H808080
    End If
    
    With ss1
        For iRow = 1 To .MaxRows
            .ROW = iRow:  .Col = SPD_CHOSE_FL:   .Text = TXT_FL.Text
        Next iRow
    End With
    
End Sub

Private Sub opt_fl_Click(Value As Integer)
Dim iRow As Integer
ss1.Col = SPD_CHOSE_FL
    If opt_fl.Value = True Then
        opt_fl.ForeColor = &HFF&
        opt_cancel_fl.ForeColor = &H808080
        opt_no_fl.ForeColor = &H808080
        TXT_FL.Text = "1"
    Else
        opt_fl.ForeColor = &H808080
    End If
    
    With ss1
        For iRow = 1 To .MaxRows
            .ROW = iRow:  .Col = SPD_CHOSE_FL:   .Text = TXT_FL.Text
        Next iRow
    End With
    
End Sub

Private Sub opt_no_fl_Click(Value As Integer)
Dim iRow As Integer
ss1.Col = SPD_CHOSE_FL
    If opt_no_fl.Value = True Then
        opt_no_fl.ForeColor = &HFF&
        opt_cancel_fl.ForeColor = &H808080
        opt_fl.ForeColor = &H808080
        TXT_FL.Text = "3"
    Else
        opt_no_fl.ForeColor = &H808080
    End If
    
    With ss1
        For iRow = 1 To .MaxRows
            .ROW = iRow:  .Col = SPD_CHOSE_FL:   .Text = TXT_FL.Text
        Next iRow
    End With
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
   If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    End If
End Sub

Public Sub Form_Pro()
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    Call Form_Ref
                
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    Dim iRow As Integer
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
    
    If BlockRow < 0 Then Exit Sub
    
    For iRow = BlockRow To BlockRow2
    
        ss1.ROW = iRow
        ss1.Col = 4
        If ss1.Value = 0 Then
            ss1.Value = 1
        Else
            ss1.Value = 0
        End If
        'Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFFFF80)

        
    Next iRow
        
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    If ss1.MaxRows < 1 Then Exit Sub
    
    If ROW <= 0 Then
       
        Call Gp_Sp_Sort1(Proc_Sc("Sc")("Spread"), Col, ROW)
    
        lBlkcol1 = 0
        lBlkcol2 = 0
        lBlkrow1 = 0
        lBlkrow2 = 0
                
    End If

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

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String

    If Len(Trim(TXT_ORD_NO.Text)) = TXT_ORD_NO.MaxLength Then
    
        If CBO_ORD_ITEM.Text <> "" Then Exit Sub
        
        TXT_ORD_NO.Text = StrConv(TXT_ORD_NO.Text, vbUpperCase)
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(TXT_ORD_NO.Text) & "'"
        Call Gf_ComboAdd(M_CN1, CBO_ORD_ITEM, sQuery)

    Else
        CBO_ORD_ITEM.Clear
    End If

End Sub

Private Sub txt_stlgrd_Change()
   If Len(TXT_STLGRD.Text) <> 11 Then TXT_STLGRD_NAME.Text = ""
End Sub

Private Sub txt_stlgrd_DblClick()

    Call txt_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_STLGRD
        DD.rControl.Add Item:=TXT_STLGRD_NAME
        DD.nameType = "1"
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
    End If
        
End Sub


Private Sub TXT_MILL_STLGRD_DblClick()
    Call txt_stdspec_chg_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_MILL_STLGRD

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

'Private Sub Text_PROC_CD_Change()
'
'    If Not TEXT_PROC_CD.Text = "" Then
'        If Len(TEXT_PROC_CD.Text) = TEXT_PROC_CD.MaxLength Then
'            TEXT_PROC_CD.Text = StrConv(TEXT_PROC_CD.Text, vbUpperCase)
'        End If
'    End If
'
'End Sub
'
'Private Sub text_PROC_CD_DblClick()
'
'    Call text_PROC_CD_KeyUp(vbKeyF4, 0)
'
'End Sub
'
'Private Sub text_PROC_CD_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.sKey = "C0004"
'
'        DD.rControl.Add Item:=TEXT_PROC_CD
'        'DD.rControl.Add Item:=Text_PROC_CD_Name
'
'        DD.nameType = "2"
'        'DD.nameType="1" 按中文名称查询
'        'DD.nameType="2" 按英文名称查询
'
'        Call Gf_Common_DD(M_CN1, KeyCode)
'
'        'Call Gf_Customer_DD(M_CN1, KeyCode)
'        ' Gf_Customer_DD() 用于客户代码
'        Exit Sub
'
'    End If
'
'    If Len(Trim(TEXT_PROC_CD.Text)) = TEXT_PROC_CD.MaxLength Then
'       '  Gf_ComnNAME_Find( 连接字符串, DD.sKEy内容 ,DD.nameType)
'       ' Gf_CustNameFind( 连接字符串, 客户代码内容,DD.nameType)
'       ' Text_PROC_CD_Name.Text = Gf_ComnNameFind(M_CN1, "C0004", TEXT_PROC_CD.Text, 2)
'    Else
'       ' Text_PROC_CD_Name.Text = ""
'    End If
'
'End Sub

'Private Sub SDT_PROD_DATE_FROM_GotFocus()
'     If SDT_PROD_DATE_FROM.RawData = "" Then
'        SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
'     End If
'     If SDT_PROD_DATE_TO.RawData = "" Then
'        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
'     End If
'End Sub

'Private Sub SDT_PROD_DATE_TO_GotFocus()
'     If SDT_PROD_DATE_TO.RawData = "" Then
'        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
'     End If
'End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Sc1) Then
       Call Gp_Ms_Cls(Mc1("rControl"))
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
       TXT_SLAB_NO.Text = ""
       TXT_STLGRD.Text = ""
       TXT_STLGRD_NAME.Text = ""
       'TEXT_PROC_CD.Text = ""
       TXT_ORD_NO.Text = ""
       CBO_ORD_ITEM.Text = ""
       TXT_FL_DATE.Text = ""
       
    End If

End Sub

Private Sub Gp_Sp_Sort1(sPname As Variant, Col As Variant, ROW As Variant, Optional CL As Boolean = False, Optional Key_Col As Long = 0)

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim sKey_Value() As String

    With sPname

        If .MaxRows < 1 Then Exit Sub
        
        If ROW <= 0 And Col > 0 Then
        
            If CL And Key_Col <> 0 Then
            
                ReDim sKey_Value(1 To .MaxRows)
                        
                For i = 1 To .MaxRows
                    .ROW = i
                    .Col = 0
                    
                    If .Text <> "" Then
                        j = j + 1
                        .Col = Key_Col
                        sKey_Value(j) = .Text
                        .Col = 0
                        .Text = ""
                        Call Gp_Sp_BlockColor(sPname, 1, .MaxCols, i, i, BLACK, WHITE)
                    End If
                Next i
                
            Else
            
                For i = 1 To .MaxRows
                    .ROW = i
                    .Col = 0
                Next i
                
            End If
        
            .SortBy = SS_SORT_BY_ROW
            
            If .SortKey(1) = Col Then
                If .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING Then
                    .SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
                Else
                    .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
                End If
            Else
                If .SortKey(1) = -1 Then
                    .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
                End If
                .SortKey(1) = Col
                
            End If
            
            .Col = 1: .Col2 = .MaxCols
            .ROW = 0: .Row2 = .MaxRows
            
            .Action = SS_ACTION_SORT
            
            'CLEAR
            If CL And Key_Col <> 0 Then
                For i = 1 To j
                    For k = 1 To .MaxRows
                        .ROW = k
                        .Col = Key_Col
                        If .Text = sKey_Value(i) Then
                            Call Gp_Sp_BlockColor(sPname, 1, .MaxCols, k, k, WHITE, BLUE)
                            .Col = 0
                            .Text = "Select"
                        End If
                    Next k
                Next i
            ElseIf CL And Key_Col = 0 Then
                .Col = 0: .Col2 = 0
                .ROW = 1: .Row2 = .MaxRows
                .BlockMode = True
                .Text = ""
                .BlockMode = False
                Call Gp_Sp_BlockColor(sPname, 1, .MaxCols, 1, .MaxRows, BLACK, WHITE)
            End If
            
        End If
        
    End With
    
End Sub

Public Sub ExcelPrn_Pile(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

On Error GoTo Excel_Error

    Dim ret         As Boolean
    Dim xlApp       As Object
    Dim xlBpp       As Object
    Dim xlBook      As Object
    Dim xlSheet     As Object
    Dim ColIndex    As Integer
    Dim sExlRange   As String
    Dim sExlRange1  As String
    Dim iExlCol     As Integer
    Dim iDate       As String
    
    With sPname
    
        If .MaxRows = 0 Then Exit Sub
        
        If bLkcol1 = 0 Then
           bLkcol1 = 1
        End If
        
        If bLkcol2 = 0 Then
            bLkcol2 = -1
        End If
        
        If bLkrow2 = 0 Then
            bLkrow2 = -1
        End If
        
        Clipboard.Clear
        
        .Col = bLkcol1: .Col2 = bLkcol2
        .ROW = bLkrow1: .Row2 = bLkrow2
        Clipboard.SetText .Clip
        
        'Call Excel
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
    
        xlApp.Visible = True
                        
        xlSheet.Cells.NumberFormatLocal = "G/通用格式"
        
        sExlRange1 = ""
        For ColIndex = 1 To .MaxCols
            .Col = ColIndex
            .ROW = 1

            iExlCol = ColIndex
'            If IsNumeric(.Text) And (Left(.Text, 1) = "0" Or Left(.Text, 1) = "1" Or Left(.Text, 1) = "7") And _
'               (Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14) Then
            If .CellType = SS_CELL_TYPE_EDIT Then
                If ColIndex > 104 Then
                    sExlRange1 = "D"
                    iExlCol = ColIndex - 104
                ElseIf ColIndex > 78 Then
                    sExlRange1 = "C"
                    iExlCol = ColIndex - 78
                ElseIf ColIndex > 52 Then
                    sExlRange1 = "B"
                    iExlCol = ColIndex - 52
                ElseIf ColIndex > 26 Then
                    sExlRange1 = "A"
                    iExlCol = ColIndex - 26
                End If

                sExlRange = sExlRange1 & Chr(iExlCol + 64) & "1:" & sExlRange1 & Chr(iExlCol + 64) & .MaxRows + 5
                If Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14 Then
                     xlSheet.Range(sExlRange).NumberFormat = "@"
                End If
            End If
        Next
       
       xlSheet.Range("A1", "I1").Merge
       xlSheet.Range("G2", "I2").Merge
       
       iDate = TXT_FL_DATE.Text
       ss1.ROW = 0
       ss1.Col = SPD_SLAB_NO:             xlApp.Range("A3").Value = ss1.Text
       ss1.Col = SPD_LOC:                 xlApp.Range("B3").Value = ss1.Text
       ss1.Col = SPD_SIZE:                xlApp.Range("D3").Value = ss1.Text
       ss1.Col = SPD_STLGRD:              xlApp.Range("E3").Value = ss1.Text
       ss1.Col = SPD_MILL_STLGRD:         xlApp.Range("F3").Value = ss1.Text
       ss1.Col = SPD_PROD_THK:            xlApp.Range("G3").Value = ss1.Text
       ss1.Col = SPD_PROD_WID:            xlApp.Range("H3").Value = ss1.Text
       ss1.Col = SPD_ORD_CNT:             xlApp.Range("I3").Value = ss1.Text
              
       xlApp.Range("G2").Value = "备料日期: " & Left(iDate, 4) + "年" + Mid(iDate, 6, 2) + "月" + Mid(iDate, 9, 2) + "日" + Mid(iDate, 12, 2) + "时" + Mid(iDate, 15, 2) + "分" + Mid(iDate, 18, 2) + "秒"
       
        xlApp.Range("A1").Value = "中板厂原料备料单"
'        xlApp.Range("A3").Value = "是否入炉"
        xlApp.Range("C3").Value = "入炉道次"
        
       
        Clipboard.Clear
        ss1.SetSelection SPD_SLAB_NO, 1, SPD_SLAB_NO, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("A4").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SPD_LOC, 1, SPD_LOC, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("B4").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SPD_SIZE, 1, SPD_SIZE, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("D4").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SPD_STLGRD, 1, SPD_STLGRD, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("E4").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SPD_MILL_STLGRD, 1, SPD_MILL_STLGRD, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("F4").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SPD_PROD_THK, 1, SPD_PROD_THK, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("G4").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SPD_PROD_WID, 1, SPD_PROD_WID, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("H4").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SPD_ORD_CNT, 1, SPD_ORD_CNT, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("I4").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
               
              
    
'        xlSheet.Range("A1").Select
'        xlSheet.Paste
'        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
            
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
        
    End With
    
    Exit Sub
    
Excel_Error:
    Call Gp_MsgBoxDisplay("您的机器尚未安装Excel", "W")

End Sub

