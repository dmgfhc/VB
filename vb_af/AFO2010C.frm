VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AFO2010C 
   Caption         =   "炼钢区域监控界面_AFO2010C"
   ClientHeight    =   9270
   ClientLeft      =   420
   ClientTop       =   2325
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_LF3 
      Height          =   285
      Left            =   3960
      TabIndex        =   39
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt_cast3 
      Height          =   285
      Left            =   3600
      TabIndex        =   38
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2025
      Left            =   7095
      TabIndex        =   9
      Top             =   390
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   3572
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16512
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "铁水预处理"
      ShadowStyle     =   1
      Begin FPSpread.vaSpread ss2 
         Height          =   1635
         Left            =   135
         TabIndex        =   30
         Top             =   270
         Width           =   7740
         _Version        =   393216
         _ExtentX        =   13653
         _ExtentY        =   2884
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
         MaxCols         =   6
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "AFO2010C.frx":0000
         UserResize      =   0
      End
      Begin VB.Line Line_CDS_1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   6450
         X2              =   6930
         Y1              =   180
         Y2              =   180
      End
   End
   Begin Threed.SSFrame SSFrame7 
      Height          =   4305
      Left            =   11610
      TabIndex        =   22
      Top             =   2490
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   7594
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16512
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "RH"
      ShadowStyle     =   1
      Begin VB.TextBox txt_rh_htno1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   1065
      End
      Begin VB.TextBox txt_rh_ld1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1335
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   885
      End
      Begin CSTextLibCtl.sitxEdit txt_rh_sta1 
         Height          =   315
         Left            =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1515
         _Version        =   262145
         _ExtentX        =   2672
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
         ForeColor       =   16711680
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         ReadOnly        =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "__-__ __:__"
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
         Mask            =   "__-__ __:__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_rh_end1 
         Height          =   315
         Left            =   1755
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1515
         _Version        =   262145
         _ExtentX        =   2672
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
         ForeColor       =   16711680
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         ReadOnly        =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "__-__ __:__"
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
         Mask            =   "__-__ __:__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_rh_arrv1 
         Height          =   315
         Left            =   240
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1515
         _Version        =   262145
         _ExtentX        =   2672
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
         ForeColor       =   16711680
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         ReadOnly        =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "__-__ __:__"
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
         Mask            =   "__-__ __:__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_rh_dep1 
         Height          =   315
         Left            =   1755
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1515
         _Version        =   262145
         _ExtentX        =   2672
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
         ForeColor       =   16711680
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         ReadOnly        =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "__-__ __:__"
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
         Mask            =   "__-__ __:__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   1335
         Top             =   360
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         Caption         =   "钢包号"
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
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   1275
         Top             =   3300
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "温度"
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
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   240
         Top             =   360
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "炉号"
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   240
         Top             =   1320
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         Caption         =   "开始/结束"
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
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   240
         Top             =   3300
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "钢水量"
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
      Begin InDate.ULabel ULabel54 
         Height          =   315
         Left            =   240
         Top             =   2340
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         Caption         =   "到达/离开"
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
      Begin CSTextLibCtl.sidbEdit txt_rh_wgt1 
         Height          =   315
         Left            =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3660
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
         ForeColor       =   16711680
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         NumIntDigits    =   12
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_rh_temp1 
         Height          =   315
         Left            =   1275
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3660
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   125
         ForeColor       =   16711680
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
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
         NumDecDigits    =   0
         NumIntDigits    =   12
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
   End
   Begin Threed.SSFrame SSFrame6 
      Height          =   2355
      Left            =   60
      TabIndex        =   13
      Top             =   6870
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   4154
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16512
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "连铸"
      ShadowStyle     =   1
      Begin FPSpread.vaSpread ss5 
         Height          =   1605
         Left            =   210
         TabIndex        =   33
         Top             =   300
         Width           =   14910
         _Version        =   393216
         _ExtentX        =   26300
         _ExtentY        =   2831
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
         MaxCols         =   40
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "AFO2010C.frx":043A
         UserResize      =   0
      End
      Begin VB.Line Line_CC_2 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   4770
         X2              =   6750
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line_CC_1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   2130
         X2              =   4110
         Y1              =   240
         Y2              =   240
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2025
      Left            =   60
      TabIndex        =   8
      Top             =   390
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   3572
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16512
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "倒罐站"
      ShadowStyle     =   1
      Begin FPSpread.vaSpread ss1 
         Height          =   1605
         Left            =   210
         TabIndex        =   29
         Top             =   300
         Width           =   6555
         _Version        =   393216
         _ExtentX        =   11562
         _ExtentY        =   2831
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
         MaxCols         =   6
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "AFO2010C.frx":11CB
         UserResize      =   0
      End
   End
   Begin Threed.SSOption Option1 
      Height          =   285
      Left            =   12450
      TabIndex        =   6
      Top             =   90
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "动态进程"
      Value           =   -1
   End
   Begin VB.TextBox txt_LF2 
      Height          =   285
      Left            =   2955
      TabIndex        =   5
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt_cast2 
      Height          =   285
      Left            =   3285
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Timer Timer3 
      Interval        =   25000
      Left            =   1110
      Top             =   30
   End
   Begin VB.TextBox txt_cast 
      Height          =   285
      Left            =   2610
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt_msp 
      Height          =   285
      Left            =   2250
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt_con 
      Height          =   285
      Left            =   1905
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txt_cds 
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   660
      Top             =   30
   End
   Begin Threed.SSOption Option2 
      Height          =   285
      Left            =   13890
      TabIndex        =   7
      Top             =   90
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "静态进程"
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   2115
      Left            =   60
      TabIndex        =   10
      Top             =   2490
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   3731
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16512
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "转炉"
      ShadowStyle     =   1
      Begin FPSpread.vaSpread ss3 
         Height          =   1635
         Left            =   210
         TabIndex        =   31
         Top             =   300
         Width           =   11100
         _Version        =   393216
         _ExtentX        =   19579
         _ExtentY        =   2884
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
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "AFO2010C.frx":1651
         UserResize      =   0
      End
      Begin VB.Line Line_CON_1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   5430
         X2              =   6885
         Y1              =   210
         Y2              =   210
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   2115
      Left            =   60
      TabIndex        =   11
      Top             =   4680
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   3731
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16512
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LF"
      ShadowStyle     =   1
      Begin FPSpread.vaSpread ss4 
         Height          =   1605
         Left            =   210
         TabIndex        =   32
         Top             =   330
         Width           =   11100
         _Version        =   393216
         _ExtentX        =   19579
         _ExtentY        =   2831
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
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "AFO2010C.frx":1B38
         UserResize      =   0
      End
      Begin VB.Line Line_MSP_2 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   7260
         X2              =   7560
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Line Line_MSP_1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   5640
         X2              =   6510
         Y1              =   270
         Y2              =   270
      End
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   225
      Left            =   180
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   397
      _Version        =   196609
      Font3D          =   1
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "VD"
      ShadowStyle     =   1
      Begin VB.TextBox txt_vd_temp1 
         Alignment       =   1  'Right Justify
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
         Left            =   8370
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   " "
         Top             =   750
         Width           =   765
      End
      Begin VB.TextBox txt_vd_wgt1 
         Alignment       =   1  'Right Justify
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
         Left            =   7380
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   " "
         Top             =   750
         Width           =   945
      End
      Begin VB.TextBox txt_vd_ld1 
         Alignment       =   2  'Center
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
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   " "
         Top             =   750
         Width           =   795
      End
      Begin VB.TextBox txt_vd_htno1 
         Alignment       =   2  'Center
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   " "
         Top             =   750
         Width           =   1005
      End
      Begin InDate.ULabel ULabel29 
         Height          =   315
         Left            =   1305
         Top             =   360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         Caption         =   "钢包号"
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
      Begin InDate.ULabel ULabel39 
         Height          =   315
         Left            =   8370
         Top             =   360
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         Caption         =   "温度"
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   240
         Top             =   360
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "炉号"
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
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Left            =   2160
         Top             =   360
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         Caption         =   "开始/结束"
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
      Begin InDate.ULabel ULabel31 
         Height          =   315
         Left            =   7380
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         Caption         =   "钢水量"
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
      Begin InDate.ULabel ULabel44 
         Height          =   315
         Left            =   4770
         Top             =   360
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         Caption         =   "到达/离开"
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
      Begin CSTextLibCtl.sitxEdit txt_vd_sta1 
         Height          =   315
         Left            =   2160
         TabIndex        =   18
         Top             =   750
         Width           =   1275
         _Version        =   262145
         _ExtentX        =   2249
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
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
         ReadOnly        =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __-__-__"
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
         Mask            =   "__-__ __:__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_vd_end1 
         Height          =   315
         Left            =   3435
         TabIndex        =   19
         Top             =   750
         Width           =   1275
         _Version        =   262145
         _ExtentX        =   2249
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
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
         ReadOnly        =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __-__-__"
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
         Mask            =   "__-__ __:__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_vd_arrv1 
         Height          =   315
         Left            =   4770
         TabIndex        =   20
         Top             =   750
         Width           =   1275
         _Version        =   262145
         _ExtentX        =   2249
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
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
         ReadOnly        =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __-__-__"
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
         Mask            =   "__-__ __:__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_vd_dep1 
         Height          =   315
         Left            =   6045
         TabIndex        =   21
         Top             =   750
         Width           =   1275
         _Version        =   262145
         _ExtentX        =   2249
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
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
         ReadOnly        =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __-__-__"
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
         Mask            =   "__-__ __:__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "断开"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   11430
      TabIndex        =   37
      Top             =   150
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "L2 连接"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10230
      TabIndex        =   36
      Top             =   150
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000C000&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   10980
      X2              =   11280
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   11880
      X2              =   12180
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "AFO2010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       STEELMAKING System
'-- Sub_System Name   Common
'-- Program Name      PLANT MONITOR
'-- Program ID        AFO2010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HWANG MAN KI
'-- Coder             HWANG MAN KI
'-- Date              2003.5.19
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
Public sDateTime As String          'Active Form Authority Setting

Public link_cds  As Long
Public link_con  As Long
Public link_msp  As Long
Public link_cast As Long
Public link_lf2  As Long
Public link_lf3  As Long
Public link_cast2 As Long
Public link_cast3 As Long

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

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim pColumn4 As New Collection      'Spread Primary Key Collection
Dim nColumn4 As New Collection      'Spread necessary Column Collection
Dim mColumn4 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn4 As New Collection      'Spread Insert Column Collection
Dim aColumn4 As New Collection      'Master -> Spread Column Collection
Dim lColumn4 As New Collection      'Spread Lock Column Collection

Dim pColumn5 As New Collection      'Spread Primary Key Collection
Dim nColumn5 As New Collection      'Spread necessary Column Collection
Dim mColumn5 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn5 As New Collection      'Spread Insert Column Collection
Dim aColumn5 As New Collection      'Master -> Spread Column Collection
Dim lColumn5 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection

Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Sc4 As New Collection           'Spread Collection
Dim Sc5 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection


Private Sub Form_Define()

    Dim iCol As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_vd_htno1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_vd_ld1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_vd_sta1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_vd_end1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_vd_arrv1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_vd_dep1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_vd_temp1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_vd_wgt1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
    Call Gp_Ms_Collection(txt_rh_htno1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_rh_ld1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_rh_sta1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_rh_end1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_rh_arrv1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_rh_dep1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_rh_temp1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_rh_wgt1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            
    'MASTER Collection
    Mc1.Add Item:="AFO2010C.P_REFER1", Key:="P-R"
    Mc1.Add Item:=pControl1, Key:="pControl"
    Mc1.Add Item:=nControl1, Key:="nControl"
    Mc1.Add Item:=mControl1, Key:="mControl"
    Mc1.Add Item:=iControl1, Key:="iControl"
    Mc1.Add Item:=rControl1, Key:="rControl"
    Mc1.Add Item:=cControl1, Key:="cControl"
    Mc1.Add Item:=aControl1, Key:="aControl"
    Mc1.Add Item:=lControl1, Key:="lControl"
  
      Call Gp_Ms_Collection(txt_cds, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(txt_con, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(txt_msp, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_cast, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(txt_LF2, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(txt_cast2, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(txt_LF3, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(txt_cast3, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    'MASTER Collection
    Mc2.Add Item:="AFO2010C.P_REFER2", Key:="P-R"
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
     
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
    
    For iCol = 1 To ss2.MaxCols
        Call Gp_Sp_Collection(ss2, iCol, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next iCol
    
    For iCol = 1 To ss3.MaxCols
        Call Gp_Sp_Collection(ss3, iCol, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Next iCol
    
    For iCol = 1 To ss4.MaxCols
        Call Gp_Sp_Collection(ss4, iCol, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Next iCol
        
    For iCol = 1 To ss5.MaxCols
        Call Gp_Sp_Collection(ss5, iCol, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Next iCol
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AFO2010C.P_SREFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
     
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AFO2010C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
     
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="AFO2010C.P_SREFER3", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
     
    'Spread_Collection
    Sc4.Add Item:=ss4, Key:="Spread"
    Sc4.Add Item:="AFO2010C.P_SREFER4", Key:="P-R"
    Sc4.Add Item:=pColumn4, Key:="pColumn"
    Sc4.Add Item:=nColumn4, Key:="nColumn"
    Sc4.Add Item:=aColumn4, Key:="aColumn"
    Sc4.Add Item:=mColumn4, Key:="mColumn"
    Sc4.Add Item:=iColumn4, Key:="iColumn"
    Sc4.Add Item:=lColumn4, Key:="lColumn"
     
    'Spread_Collection
    Sc5.Add Item:=ss5, Key:="Spread"
    Sc5.Add Item:="AFO2010C.P_SREFER5", Key:="P-R"
    Sc5.Add Item:=pColumn5, Key:="pColumn"
    Sc5.Add Item:=nColumn5, Key:="nColumn"
    Sc5.Add Item:=aColumn5, Key:="aColumn"
    Sc5.Add Item:=mColumn5, Key:="mColumn"
    Sc5.Add Item:=iColumn5, Key:="iColumn"
    Sc5.Add Item:=lColumn5, Key:="lColumn"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    MDIMain.StatusBar1.Panels(1) = "提示信息："
    
    Call Form_Ref1
    Call Form_Ref2
    
    link_cds = Val(txt_cds.Text)
    link_con = Val(txt_con.Text)
    link_msp = Val(txt_msp.Text)
    link_cast = Val(txt_cast.Text)
    link_lf2 = Val(txt_LF2.Text)
    link_cast2 = Val(txt_cast2.Text)
    link_lf3 = Val(txt_LF3.Text)
    link_cast3 = Val(txt_cast3.Text)
    
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_ControlLock(Mc2("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    
    Call Sp_Setting(ss1)
    Call Sp_Setting(ss2)
    Call Sp_Setting(ss3)
    Call Sp_Setting(ss4)
    Call Sp_Setting(ss5)
    
    Call Gp_Sp_ColGet(ss5, "K-System.INI", Me.Name)
    
    Call Form_Ref1
    Call Form_Ref2
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call Gp_Sp_ColSet(ss5, "K-System.INI", Me.Name)
    
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
    
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
    
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing
    
    Set iColumn4 = Nothing
    Set pColumn4 = Nothing
    Set lColumn4 = Nothing
    Set nColumn4 = Nothing
    Set mColumn4 = Nothing
    Set aColumn4 = Nothing
    
    Set iColumn5 = Nothing
    Set pColumn5 = Nothing
    Set lColumn5 = Nothing
    Set nColumn5 = Nothing
    Set mColumn5 = Nothing
    Set aColumn5 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set Sc4 = Nothing
    Set Sc5 = Nothing
    
    Timer1.Enabled = False
    Timer3.Enabled = False

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Option1.SetFocus
    
End Sub

Public Sub Form_Ref1()

    If Sp_Display(M_CN1, sc1.Item("Spread"), Gf_Sp_MakeQuery(sc1.Item("Spread"), sc1.Item("P-R"), _
                                    "R", sc1.Item("aColumn"), 1), sc1.Item("pColumn"), False) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       
    End If
    
    If Sp_Display(M_CN1, sc2.Item("Spread"), Gf_Sp_MakeQuery(sc2.Item("Spread"), sc2.Item("P-R"), _
                                    "R", sc2.Item("aColumn"), 1), sc2.Item("pColumn"), False) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    If Sp_Display(M_CN1, Sc3.Item("Spread"), Gf_Sp_MakeQuery(Sc3.Item("Spread"), Sc3.Item("P-R"), _
                                    "R", Sc3.Item("aColumn"), 1), Sc3.Item("pColumn"), False) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    If Sp_Display(M_CN1, Sc4.Item("Spread"), Gf_Sp_MakeQuery(Sc4.Item("Spread"), Sc4.Item("P-R"), _
                                    "R", Sc4.Item("aColumn"), 1), Sc4.Item("pColumn"), False) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    If Sp_Display(M_CN1, Sc5.Item("Spread"), Gf_Sp_MakeQuery(Sc5.Item("Spread"), Sc5.Item("P-R"), _
                                    "R", Sc5.Item("aColumn"), 1), Sc5.Item("pColumn"), False) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    If Gf_Ms_Refer(M_CN1, Mc1) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If

End Sub

Public Sub Form_Ref2()

    If Gf_Ms_Refer(M_CN1, Mc2) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
        
'    Call LinkStatusCheck(txt_cds, Line_CDS_1, Line_CDS_2, link_cds1, link_cds2, "CDS")
'
'    Call LinkStatusCheck(txt_con, Line_CON_1, Line_CON_2, link_con1, link_con2, "CON")
'
'    Call LinkStatusCheck(txt_msp, Line_MSP_1, Line_MSP_2, link_msp1, link_msp2, "MSP")
'
'    Call LinkStatusCheck(txt_cast, Line_CC_1, Line_CC_2, link_cast1, link_cast2, "CAST")
End Sub

'Public Sub LinkStatusCheck(oTxtBox As Variant, oLine1 As Variant, oLine2 As Variant, lLink1 As Long, lLink2 As Long, sPrc As String)

'    Dim i      As Integer
'    Dim lLine1 As Long
'    Dim lLine2 As Long
'
'    i = InStr(1, oTxtBox.Text, "-")
'    lLine1 = Val(Left(oTxtBox.Text, i - 1))
'    lLine2 = Val(Mid(oTxtBox.Text, i + 1, Len(oTxtBox.Text)))
'
'    If lLink1 = lLine1 Then
'       oLine1.BorderColor = &HFF00FF
'    Else
'       oLine1.BorderColor = &HC000&
'    End If
'    If lLink2 = lLine2 Then
'       oLine2.BorderColor = &HFF00FF
'    Else
'       oLine2.BorderColor = &HC000&
'    End If
'
'    Select Case sPrc
'        Case "CDS"
'            link_cds1 = lLine1
'            link_cds2 = lLine2
'        Case "CON"
'            link_con1 = lLine1
'            link_con2 = lLine2
'        Case "MSP"
'            link_msp1 = lLine1
'            link_msp2 = lLine2
'        Case "CAST"
'            link_cast1 = lLine1
'            link_cast2 = lLine2
'    End Select
'End Sub

Private Sub Option1_Click(VALUE As Integer)

    Timer1.Enabled = True
    Timer3.Enabled = True

End Sub

Private Sub Option2_Click(VALUE As Integer)

    Timer1.Enabled = False
    Timer3.Enabled = False
    
End Sub

Private Sub Timer1_Timer()

    Call Form_Ref1
    
End Sub

Private Sub Timer3_Timer()

    Call Form_Ref2
    
'    If link_cds = Val(txt_cds.Text) Then
'        Call Gp_Sp_BlockColor(ss2, 0, -1, 1, 3, &HFF00FF)
'        Line_CDS_1.BorderColor = &HFF00FF
'    Else
'        Call Gp_Sp_BlockColor(ss2, 0, -1, 1, 3, &H74A253)
'        Line_CDS_1.BorderColor = &HC000&
'    End If
'
'    If link_con = Val(txt_con.Text) Then
'        Call Gp_Sp_BlockColor(ss3, 0, -1, 1, 3, &HFF00FF)
'        Line_CON_1.BorderColor = &HFF00FF
'    Else
'        Call Gp_Sp_BlockColor(ss3, 0, -1, 1, 3, &H74A253)
'        Line_CON_1.BorderColor = &HC000&
'    End If
'
'    If link_msp = Val(txt_msp.Text) Then
'        Call Gp_Sp_BlockColor(ss4, 0, -1, 1, 1, &HFF00FF)
'        Line_MSP_1.BorderColor = &HFF00FF
'    Else
'        Call Gp_Sp_BlockColor(ss4, 0, -1, 1, 1, &H74A253)
'        Line_MSP_1.BorderColor = &HC000&
'    End If
'
'    If link_lf2 = Val(txt_LF2.Text) Then
'        Call Gp_Sp_BlockColor(ss4, 0, -1, 2, 2, &HFF00FF)
'        Line_MSP_2.BorderColor = &HFF00FF
'    Else
'        Call Gp_Sp_BlockColor(ss4, 0, -1, 2, 2, &H74A253)
'        Line_MSP_2.BorderColor = &HC000&
'    End If
'
'    If link_lf3 = Val(txt_LF3.Text) Then
'        Call Gp_Sp_BlockColor(ss4, 0, -1, 3, 3, &HFF00FF)
'        ss4.BlockMode = False
'    Else
'        Call Gp_Sp_BlockColor(ss4, 0, -1, 3, 3, &H74A253)
'        ss4.BlockMode = False
'    End If
'
'    If link_cast = Val(txt_cast.Text) Then
'        Call Gp_Sp_BlockColor(ss5, 0, -1, 1, 1, &HFF00FF)
'        Line_CC_1.BorderColor = &HFF00FF
'    Else
'        Call Gp_Sp_BlockColor(ss5, 0, -1, 1, 1, &H74A253)
'        Line_CC_1.BorderColor = &HC000&
'    End If
'
'    If link_cast2 = Val(txt_cast2.Text) Then
'        Call Gp_Sp_BlockColor(ss5, 0, -1, 2, 2, &HFF00FF)
'        Line_CC_2.BorderColor = &HFF00FF
'    Else
'        Call Gp_Sp_BlockColor(ss5, 0, -1, 2, 2, &H74A253)
'        Line_CC_2.BorderColor = &HC000&
'    End If
'
'    If link_cast3 = Val(txt_cast3.Text) Then
'        Call Gp_Sp_BlockColor(ss5, 0, -1, 3, 3, &HFF00FF)
'    Else
'        Call Gp_Sp_BlockColor(ss5, 0, -1, 3, 3, &H74A253)
'    End If
'
'    link_cds = Val(txt_cds.Text)
'    link_con = Val(txt_con.Text)
'    link_msp = Val(txt_msp.Text)
'    link_cast = Val(txt_cast.Text)
'    link_lf2 = Val(txt_LF2.Text)
'    link_cast2 = Val(txt_cast2.Text)
'    link_lf3 = Val(txt_LF3.Text)
'    link_cast3 = Val(txt_cast3.Text)
    
End Sub

Private Sub Sp_Setting(ByVal sPname As Variant)

    Dim lCol As Integer

    With sPname
    
        .RowHeight(-1) = 18
        
        If .ColHeaderRows > 1 Then
            .RowHeight(SpreadHeader + (.ColHeaderRows - 2)) = 12
            .RowHeight(SpreadHeader + (.ColHeaderRows - 1)) = 12
        Else
            .RowHeight(0) = 24
        End If
        
        .ColWidth(0) = 6
        
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        '.SelBackColor = &HFFFFFF        ''&HE3F4FF      ''&HFFFF80     '&H808040
        '.SelForeColor = &H80000012
        .ForeColor = &HFF0000
     
        .OperationMode = OperationModeNormal
        .RetainSelBlock = True

        .UserResize = UserResizeNone
        .AllowDragDrop = False
        
        .ProcessTab = True
        .ScrollBarExtMode = True
        .ButtonDrawMode = 1
        .TabStop = True
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10.5
        .BlockMode = False
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        .BlockMode = True
        .FontBold = True
        .BlockMode = False
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = 0
        .BlockMode = True
        .FontBold = True
        .BlockMode = False
        
    End With
    
    Select Case sPname.Name
    
        Case "ss1"
            ss1.MaxRows = 3
            ss1.ColWidth(1) = 9
            ss1.ColWidth(2) = 7
            ss1.ColWidth(3) = 7
            ss1.ColWidth(4) = 6
            ss1.ColWidth(5) = 11.5
            ss1.ColWidth(6) = 7
            ss1.Col = 0
            ss1.Row = 1
            ss1.Text = "PIT 1"
            ss1.Row = 2
            ss1.Text = "PIT 2"
            ss1.Row = 3
            ss1.Text = "PIT 3"
            ss1.Height = 1635
            ss1.Width = 6540
            ss1.Left = 210
            ss1.Top = 300
            ss1.OperationMode = OperationModeRead
            ss1.Col = 0: ss1.Col2 = -1
            ss1.Row = 1: ss1.Row2 = 3
            ss1.BlockMode = True
            ss1.ForeColor = &HFF0000
            ss1.BlockMode = False
        
        Case "ss2"
        
            ss2.MaxRows = 3
            ss2.ColWidth(1) = 9
            ss2.ColWidth(2) = 7
            ss2.ColWidth(3) = 8
            ss2.ColWidth(4) = 7
            ss2.ColWidth(5) = 13
            ss2.ColWidth(6) = 13
            ss2.Col = 0
            ss2.Row = 1
            ss2.Text = "#1"
            ss2.Row = 2
            ss2.Text = "#2"
            ss2.Row = 3
            ss2.Text = "#3"
            ss2.Height = 1635
            ss2.Width = 7650
            ss2.Left = 210
            ss2.Top = 300
            ss2.OperationMode = OperationModeRead
        
        Case "ss3"
        
            ss3.MaxRows = 3
            ss3.ColWidth(1) = 10
            ss3.ColWidth(2) = 8
            ss3.ColWidth(3) = 12.5
            ss3.ColWidth(4) = 12.5
            ss3.ColWidth(5) = 12.5
            ss3.ColWidth(6) = 12.5
            ss3.ColWidth(7) = 9
            ss3.ColWidth(8) = 8
            ss3.Col = 0
            ss3.Row = 1
            ss3.Text = "#1"
            ss3.Row = 2
            ss3.Text = "#2"
            ss3.Row = 3
            ss3.Text = "#3"
            ss3.Height = 1635
            ss3.Width = 11075
            ss3.Left = 210
            ss3.Top = 300
            ss3.OperationMode = OperationModeRead
        
        Case "ss4"
        
            ss4.MaxRows = 3
            ss4.ColWidth(1) = 10
            ss4.ColWidth(2) = 8
            ss4.ColWidth(3) = 12.5
            ss4.ColWidth(4) = 12.5
            ss4.ColWidth(5) = 12.5
            ss4.ColWidth(6) = 12.5
            ss4.ColWidth(7) = 9
            ss4.ColWidth(8) = 8
            ss4.Col = 0
            ss4.Row = 1
            ss4.Text = "#1"
            ss4.Row = 2
            ss4.Text = "#2"
            ss4.Row = 3
            ss4.Text = "#3"
            ss4.Height = 1635
            ss4.Width = 11075
            ss4.Left = 210
            ss4.Top = 300
            ss4.OperationMode = OperationModeRead
        
        Case "ss5"
        
            ss5.MaxRows = 3
            ss5.ColWidth(1) = 10
            ss5.ColWidth(2) = 8
            ss5.ColWidth(3) = 7
            ss5.ColWidth(4) = 12
            ss5.ColWidth(5) = 12
            ss5.ColWidth(6) = 12
            ss5.ColWidth(7) = 12
            ss5.ColWidth(8) = 12
            ss5.ColWidth(9) = 12
            ss5.ColWidth(10) = 8
            
            For lCol = 11 To ss5.MaxCols
                ss5.ColWidth(lCol) = 5.5
            Next lCol
            
            ss5.Col = 0
            ss5.Row = 1
            ss5.Text = "#1"
            ss5.Row = 2
            ss5.Text = "#2"
            ss5.Row = 3
            ss5.Text = "#3"
            ss5.Height = 1935
            ss5.Width = 14750
            ss5.Left = 210
            ss5.Top = 300
            ss5.OperationMode = OperationModeRead
            ss5.ScrollBars = ScrollBarsHorizontal
            ss5.UserResize = UserResizeColumns
            ss5.ColsFrozen = 1
        
    End Select
        
End Sub

Private Function Sp_Display(Conn As ADODB.Connection, sPname As vaSpread, sQuery As String, _
                              Optional lColumn As Variant = Nothing, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim sSpreadClip As String
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Sp_Display = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Display = True
        
        .ReDraw = False
        iCount = 0
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
                
            Sp_Display = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = vbDefault
            Exit Function
        Else
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        For iRowCount = 0 To .MaxRows - 1
        
            .Row = iRowCount + 1
            
            For iColcount = 0 To .MaxCols - 1
            
                .Col = iColcount + 1

                Select Case .CellType

                    Case SS_CELL_TYPE_CHECKBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) <> vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "1" Then
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If

                    Case SS_CELL_TYPE_COMBOBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "" Then
                            .VALUE = 0
                        Else
                            .VALUE = Trim(ArrayRecords(iColcount, iRowCount))
                        End If

                    Case SS_CELL_TYPE_DATE
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Mid(Trim(ArrayRecords(iColcount, iRowCount)), 1, 4) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 5, 2) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 7, 2)
                        End If

                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .VALUE = ""
                        Else
                            .VALUE = Trim(ArrayRecords(iColcount, iRowCount))
                        End If

                    Case Else
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If

                End Select

            Next iColcount
            
        Next iRowCount
            
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Display = False
    Call Gp_MsgBoxDisplay("Sp_Display Error : " & sQuery)
    Screen.MousePointer = vbDefault

End Function

