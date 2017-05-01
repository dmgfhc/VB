VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGA2010C 
   Caption         =   "加热炉作业实绩查询及修改界面_AGA2010C"
   ClientHeight    =   9495
   ClientLeft      =   120
   ClientTop       =   1440
   ClientWidth     =   15120
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin Threed.SSCheck SSC1 
      Height          =   195
      Left            =   450
      TabIndex        =   67
      Top             =   5235
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      _Version        =   196609
      BackColor       =   14804173
   End
   Begin Threed.SSCheck sc3 
      Height          =   315
      Left            =   11370
      TabIndex        =   33
      Top             =   690
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   2
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
      Caption         =   "缺号板坯"
   End
   Begin Threed.SSCheck sc2 
      Height          =   315
      Left            =   5910
      TabIndex        =   0
      Top             =   690
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   2
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
      Caption         =   "板坯出炉"
   End
   Begin Threed.SSCheck sc1 
      Height          =   315
      Left            =   435
      TabIndex        =   1
      Top             =   690
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   2
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
      Caption         =   "板坯装炉"
   End
   Begin Threed.SSFrame sf3 
      Height          =   4275
      Left            =   11115
      TabIndex        =   29
      Top             =   795
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   7541
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_REASON_NAME 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox CHK_EXIT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "出口"
         Height          =   195
         Left            =   2355
         TabIndex        =   38
         Top             =   2130
         Width           =   750
      End
      Begin VB.CheckBox CHK_ENTRY 
         BackColor       =   &H00E0E0E0&
         Caption         =   "入口"
         Height          =   195
         Left            =   2355
         TabIndex        =   37
         Top             =   1860
         Width           =   750
      End
      Begin VB.CheckBox CHK_NONORDER 
         BackColor       =   &H00FFFF80&
         Caption         =   "余材"
         Height          =   240
         Left            =   2445
         TabIndex        =   26
         Top             =   2865
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CheckBox CHK_ORDER 
         BackColor       =   &H00FFFF80&
         Caption         =   "订单"
         Height          =   240
         Left            =   2445
         TabIndex        =   25
         Top             =   2550
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox TXT_COMFRM 
         Height          =   330
         Left            =   1830
         TabIndex        =   36
         Text            =   " "
         Top             =   2550
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox TXT_REJ_LOC 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1620
         TabIndex        =   35
         Text            =   " "
         Top             =   1845
         Width           =   645
      End
      Begin InDate.ULabel ULabel3 
         Height          =   330
         Left            =   300
         Top             =   2550
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         Caption         =   "缺号板坯确定"
         Alignment       =   1
         BackColor       =   16777088
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
      Begin VB.TextBox TXT_REASON_CD 
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
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   24
         Top             =   1080
         Width           =   645
      End
      Begin VB.TextBox TXT_REJ_EMP 
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
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3750
         Width           =   1335
      End
      Begin VB.TextBox TXT_REJ_GROUP 
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
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3750
         Width           =   705
      End
      Begin VB.TextBox TXT_REJ_SHIFT 
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
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3750
         Width           =   705
      End
      Begin InDate.ULabel ULabel26 
         Height          =   315
         Left            =   270
         Top             =   315
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "缺号时间"
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
      Begin InDate.ULabel ULabel32 
         Height          =   315
         Left            =   270
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "缺号代码"
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
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel33 
         Height          =   315
         Left            =   270
         Top             =   1845
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "缺号位置"
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
      Begin CSTextLibCtl.sitxEdit TXT_REJ_OCCR_TIME 
         Height          =   315
         Left            =   1620
         TabIndex        =   23
         Tag             =   "缺号时"
         Top             =   315
         Width           =   2130
         _Version        =   262145
         _ExtentX        =   3757
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
      Begin InDate.ULabel ULabel40 
         Height          =   315
         Left            =   300
         Top             =   3420
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Caption         =   "班次"
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
      Begin InDate.ULabel ULabel41 
         Height          =   315
         Left            =   1005
         Top             =   3420
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Caption         =   "班别"
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
      Begin InDate.ULabel ULabel42 
         Height          =   315
         Left            =   1710
         Top             =   3420
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "作业人员"
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
   End
   Begin Threed.SSFrame sf2 
      Height          =   4275
      Left            =   5625
      TabIndex        =   28
      Top             =   795
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   7541
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.CheckBox CHK_O_UNCHA_IND 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2-再出炉"
         Height          =   240
         Left            =   4200
         TabIndex        =   19
         Tag             =   "出炉/出装炉 再装炉"
         Top             =   1200
         Width           =   1020
      End
      Begin VB.CheckBox CHK_O_CHA_IND 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1-出炉"
         Height          =   240
         Left            =   3255
         TabIndex        =   18
         Tag             =   "出炉/出装炉 装炉"
         Top             =   1200
         Width           =   870
      End
      Begin VB.TextBox TXT_DIS_UNDIS_IND 
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
         Left            =   4620
         TabIndex        =   34
         Text            =   " "
         Top             =   825
         Width           =   570
      End
      Begin VB.TextBox TXT_DISCHARGE_EMP 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3750
         Width           =   1335
      End
      Begin VB.TextBox TXT_DISCHARGE_GROUP 
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
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3750
         Width           =   705
      End
      Begin VB.TextBox TXT_DISCHARGE_SHIFT 
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
         Left            =   285
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   3750
         Width           =   705
      End
      Begin CSTextLibCtl.sidbEdit SDB_EXP_TEMP 
         Height          =   315
         Left            =   1635
         TabIndex        =   8
         Tag             =   "温度"
         Top             =   825
         Width           =   915
         _Version        =   262145
         _ExtentX        =   1614
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_PRE_TOP_SLAB_TEMP 
         Height          =   315
         Left            =   1635
         TabIndex        =   9
         Tag             =   "预热区上表面温度"
         Top             =   1935
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_HT_TOP_SLAB_TEMP 
         Height          =   315
         Left            =   1635
         TabIndex        =   12
         Tag             =   "加热区上表面温度"
         Top             =   2340
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_SOK_HOT_SLAB_TEMP 
         Height          =   315
         Left            =   1635
         TabIndex        =   15
         Tag             =   "均热区上表面温度"
         Top             =   2745
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_PRE_BOT_SLAB_TEMP 
         Height          =   315
         Left            =   2820
         TabIndex        =   10
         Tag             =   "预热区 下表面温度"
         Top             =   1935
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_HT_BOT_SLAB_TEMP 
         Height          =   315
         Left            =   2820
         TabIndex        =   13
         Tag             =   "加热区 下表面温度"
         Top             =   2340
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_SOK_BOT_SLAB_TEMP 
         Height          =   315
         Left            =   2820
         TabIndex        =   16
         Tag             =   "均热区 下表面温度"
         Top             =   2745
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_PRE_ZONE_TIME 
         Height          =   315
         Left            =   4005
         TabIndex        =   11
         Tag             =   "预热区 驻留时间"
         Top             =   1935
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_HT_ZONE_TIME 
         Height          =   315
         Left            =   4005
         TabIndex        =   14
         Tag             =   "加热区 驻留时间"
         Top             =   2340
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_SOK_ZONE_TIME 
         Height          =   315
         Left            =   4005
         TabIndex        =   17
         Tag             =   "均热区 驻留时间"
         Top             =   2745
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   285
         Top             =   1935
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "预热区"
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
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   285
         Top             =   2340
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "加热区"
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
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   285
         Top             =   2745
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "均热区"
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
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   1635
         Top             =   1575
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "TOP 温度"
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   2820
         Top             =   1575
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "BOT 温度"
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
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   4005
         Top             =   1575
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "驻留时间"
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Left            =   3255
         Top             =   825
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "出炉/再出炉"
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
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   285
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "出炉时间"
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
      Begin InDate.ULabel ULabel25 
         Height          =   315
         Left            =   285
         Top             =   825
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "温度(℃)"
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
      Begin CSTextLibCtl.sitxEdit TXT_DISCHARGE_TIME 
         Height          =   315
         Left            =   1635
         TabIndex        =   7
         Tag             =   "出炉时间"
         Top             =   360
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
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
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   285
         Top             =   3420
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Caption         =   "班次"
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
      Begin InDate.ULabel ULabel38 
         Height          =   315
         Left            =   990
         Top             =   3420
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         Caption         =   "班别"
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
         Left            =   1680
         Top             =   3420
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "作业人员"
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
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   285
         Top             =   1185
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "温度均匀性"
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
      Begin CSTextLibCtl.sidbEdit SDB_PDT_UNI_TEMP 
         Height          =   315
         Left            =   1635
         TabIndex        =   68
         Tag             =   "温度均匀性"
         Top             =   1185
         Width           =   915
         _Version        =   262145
         _ExtentX        =   1614
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2550
         TabIndex        =   70
         Top             =   855
         Width           =   255
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   720
      Left            =   165
      TabIndex        =   27
      Top             =   105
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   1270
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.ComboBox TXT_PROC_LINE 
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
         ItemData        =   "AGA2010C.frx":0000
         Left            =   7830
         List            =   "AGA2010C.frx":000A
         TabIndex        =   74
         Tag             =   "工厂代码"
         Top             =   165
         Width           =   600
      End
      Begin VB.TextBox TXT_RHF_CH_NUM_REF 
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
         Left            =   4820
         TabIndex        =   72
         Top             =   165
         Width           =   465
      End
      Begin VB.TextBox TXT_UPD_EMP 
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
         Left            =   13305
         MaxLength       =   7
         TabIndex        =   6
         Tag             =   "作业人员"
         Top             =   165
         Width           =   1230
      End
      Begin VB.ComboBox CBO_SLAB_NO 
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
         Left            =   1650
         TabIndex        =   2
         Tag             =   "板坯号"
         Top             =   165
         Width           =   1665
      End
      Begin VB.ComboBox CBO_PLT 
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
         ItemData        =   "AGA2010C.frx":0014
         Left            =   7125
         List            =   "AGA2010C.frx":001E
         TabIndex        =   3
         Tag             =   "工厂代码"
         Top             =   165
         Width           =   720
      End
      Begin VB.ComboBox CBO_SHIFT 
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
         ItemData        =   "AGA2010C.frx":002A
         Left            =   9525
         List            =   "AGA2010C.frx":0037
         TabIndex        =   4
         Tag             =   "班次"
         Top             =   165
         Width           =   720
      End
      Begin VB.ComboBox CBO_GROUP 
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
         ItemData        =   "AGA2010C.frx":0044
         Left            =   11370
         List            =   "AGA2010C.frx":0054
         TabIndex        =   5
         Tag             =   "班别"
         Top             =   165
         Width           =   720
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   285
         Top             =   165
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "板坯号"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   5730
         Top             =   165
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "工厂/炉座号"
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
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   8655
         Top             =   165
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "班次"
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
         Left            =   10500
         Top             =   165
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "班别"
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
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   12315
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         Caption         =   "作业人员"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   3750
         Top             =   165
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "装炉次数"
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
   End
   Begin InDate.ULabel ULabel31 
      Height          =   315
      Left            =   11250
      Top             =   4575
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Caption         =   "重量"
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
      Left            =   11250
      Top             =   4095
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Caption         =   "长度"
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
   Begin InDate.ULabel ULabel29 
      Height          =   315
      Left            =   11250
      Top             =   3615
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Caption         =   "尾部宽度"
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
   Begin InDate.ULabel ULabel28 
      Height          =   315
      Left            =   11250
      Top             =   3150
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Caption         =   "头部宽度"
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
   Begin InDate.ULabel ULabel27 
      Height          =   315
      Left            =   11250
      Top             =   2655
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Caption         =   "厚度"
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
   Begin CSTextLibCtl.sidbEdit SDB_REJ_SLAB_LEN 
      Height          =   315
      Left            =   12780
      TabIndex        =   39
      Top             =   4095
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
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
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit SDB_REJ_SLAB_WGT 
      Height          =   315
      Left            =   12780
      TabIndex        =   40
      Top             =   4575
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
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
      NumIntDigits    =   4
      ShowZero        =   0   'False
      MaxValue        =   9999.999
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit SDB_TAIL_SLAB_WID 
      Height          =   315
      Left            =   12780
      TabIndex        =   41
      Top             =   3615
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
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
   Begin CSTextLibCtl.sidbEdit SDB_HEAD_SLAB_WID 
      Height          =   315
      Left            =   12780
      TabIndex        =   42
      Top             =   3150
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
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
   Begin CSTextLibCtl.sidbEdit SDB_SLAB_REJ_THK 
      Height          =   315
      Left            =   12780
      TabIndex        =   43
      Top             =   2655
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
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
   Begin Threed.SSFrame sf1 
      Height          =   4275
      Left            =   165
      TabIndex        =   49
      Top             =   795
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   7541
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_RHF_CH_NUM 
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TXT_SLAB_SIZE 
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
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   " "
         Top             =   2745
         Width           =   2805
      End
      Begin VB.TextBox TXT_CH_SHIFT 
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
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   3750
         Width           =   705
      End
      Begin VB.TextBox TXT_CH_GROUP 
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
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   3750
         Width           =   705
      End
      Begin VB.TextBox TXT_CH_EMP 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   3750
         Width           =   1335
      End
      Begin VB.TextBox TXT_CD 
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
         Left            =   3570
         TabIndex        =   57
         Text            =   "CA"
         Top             =   3540
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CheckBox CHK_RHF_ROW_A 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0-单排料"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2850
         TabIndex        =   56
         Top             =   1935
         Width           =   1200
      End
      Begin VB.CheckBox CHK_RHF_ROW_L 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1-双排左料"
         Height          =   225
         Left            =   2850
         TabIndex        =   55
         Top             =   2175
         Width           =   1200
      End
      Begin VB.CheckBox CHK_RHF_ROW_R 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2-双排右料"
         Height          =   240
         Left            =   2850
         TabIndex        =   54
         Top             =   2415
         Width           =   1200
      End
      Begin VB.TextBox TXT_RHF_CH_ROW 
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
         Left            =   4215
         TabIndex        =   53
         Text            =   " "
         Top             =   1575
         Width           =   915
      End
      Begin VB.TextBox TXT_CHA_UNCHA_IND 
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
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   " "
         Top             =   1575
         Width           =   855
      End
      Begin VB.CheckBox CHK_CHA_IND 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1-装炉"
         Height          =   240
         Left            =   270
         TabIndex        =   51
         Tag             =   "装炉"
         Top             =   1935
         Width           =   840
      End
      Begin VB.CheckBox CHK_UNCHA_IND 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2-再装炉"
         Height          =   240
         Left            =   270
         TabIndex        =   50
         Tag             =   "装炉/再装炉 再装炉"
         Top             =   2175
         Width           =   1020
      End
      Begin CSTextLibCtl.sitxEdit TXT_RHF_CH_TIME 
         Height          =   315
         Left            =   1635
         TabIndex        =   61
         Tag             =   "装炉时间"
         Top             =   360
         Width           =   2130
         _Version        =   262145
         _ExtentX        =   3757
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
      Begin CSTextLibCtl.sidbEdit SDB_CHARGE_TEMP 
         Height          =   315
         Left            =   1635
         TabIndex        =   62
         Tag             =   "温度"
         Top             =   1185
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_RHF_SLAB_WGT 
         Height          =   315
         Left            =   4215
         TabIndex        =   63
         Tag             =   "重量"
         Top             =   1185
         Width           =   915
         _Version        =   262145
         _ExtentX        =   1614
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999.999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   270
         Top             =   360
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "装炉时间"
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
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   2850
         Top             =   1575
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "布料方式"
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
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   270
         Top             =   1185
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "温度(℃)"
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
         Left            =   2850
         Top             =   1185
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "重量(ton)"
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
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   270
         Top             =   1575
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "装炉/再装炉"
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
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   270
         Top             =   3420
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Caption         =   "班次"
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
      Begin InDate.ULabel ULabel35 
         Height          =   315
         Left            =   975
         Top             =   3420
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Caption         =   "班别"
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
      Begin InDate.ULabel ULabel36 
         Height          =   315
         Left            =   1680
         Top             =   3420
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "作业人员"
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
      Begin InDate.ULabel ULabel43 
         Height          =   315
         Left            =   270
         Top             =   2745
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "板坯规格"
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
      Begin CSTextLibCtl.sitxEdit TXT_RHF_CH_TIME_UPD 
         Height          =   315
         Left            =   1635
         TabIndex        =   69
         Tag             =   "装炉时间"
         Top             =   720
         Width           =   2130
         _Version        =   262145
         _ExtentX        =   3757
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
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   270
         Top             =   720
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "装炉时间修正"
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
         Left            =   3960
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "装炉次数"
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
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4530
         TabIndex        =   64
         Top             =   2760
         Width           =   285
      End
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   4305
      Left            =   165
      TabIndex        =   66
      Top             =   5055
      Width           =   14970
      _Version        =   393216
      _ExtentX        =   26405
      _ExtentY        =   7594
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
      MaxCols         =   10
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AGA2010C.frx":0064
   End
   Begin VB.Label Label51 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ton"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13995
      TabIndex        =   48
      Top             =   4605
      Width           =   390
   End
   Begin VB.Label Label49 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13995
      TabIndex        =   47
      Top             =   3645
      Width           =   255
   End
   Begin VB.Label Label47 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13995
      TabIndex        =   46
      Top             =   4125
      Width           =   255
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13995
      TabIndex        =   45
      Top             =   2655
      Width           =   255
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13995
      TabIndex        =   44
      Top             =   3180
      Width           =   255
   End
End
Attribute VB_Name = "AGA2010C"
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
'-- Program Name      加热炉作业实绩查询及修改界面
'-- Program ID        AGA2010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
'-- Date              2003.7.23
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
Public sDateTime As String          'Active Form Time Setting
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

Dim pControl3 As New Collection      'Master Primary Key Collection
Dim nControl3 As New Collection      'Master Necessary Collection
Dim mControl3 As New Collection      'Master Maxlength check Collection
Dim iControl3 As New Collection      'Master Insert Collection
Dim rControl3 As New Collection      'Master Refer Collection
Dim cControl3 As New Collection      'Master Copy Collection
Dim aControl3 As New Collection      'Master -> Spread Collection
Dim lControl3 As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection
 
Dim Proc_Sc As New Collection       'Spread Struc Collection
 
Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim sControl  As New Collection      'Master Clear Key Collection
Dim MC        As New Collection      'Master Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection           'Master Collection
Dim sc4 As New Collection           'Spread Collection
 
Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master" '"Msheet"

'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  BELOW EDIT  -------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(CBO_SLAB_NO, "p", "n", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(TXT_RHF_CH_NUM_REF, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(TXT_PROC_LINE, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(CBO_PLT, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            Call Gp_Ms_Collection(cbo_group, " ", " ", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(TXT_UPD_EMP, " ", "n", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(TXT_RHF_CH_TIME, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
  Call Gp_Ms_Collection(TXT_RHF_CH_TIME_UPD, " ", " ", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(TXT_RHF_CH_NUM, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(TXT_RHF_CH_ROW, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(SDB_CHARGE_TEMP, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(TXT_SLAB_SIZE, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(SDB_RHF_SLAB_WGT, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(TXT_CHA_UNCHA_IND, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(TXT_CH_SHIFT, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(TXT_CH_GROUP, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
           Call Gp_Ms_Collection(TXT_CH_EMP, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)

          Call Gp_Ms_Collection(CBO_SLAB_NO, "p", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   Call Gp_Ms_Collection(TXT_RHF_CH_NUM_REF, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(TXT_PROC_LINE, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
              Call Gp_Ms_Collection(CBO_PLT, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(cbo_group, " ", " ", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(TXT_UPD_EMP, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   Call Gp_Ms_Collection(TXT_DISCHARGE_TIME, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(SDB_EXP_TEMP, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(SDB_PDT_UNI_TEMP, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
Call Gp_Ms_Collection(SDB_PRE_TOP_SLAB_TEMP, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
Call Gp_Ms_Collection(SDB_PRE_BOT_SLAB_TEMP, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(SDB_PRE_ZONE_TIME, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
 Call Gp_Ms_Collection(SDB_HT_TOP_SLAB_TEMP, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
 Call Gp_Ms_Collection(SDB_HT_BOT_SLAB_TEMP, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(SDB_HT_ZONE_TIME, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
Call Gp_Ms_Collection(SDB_SOK_HOT_SLAB_TEMP, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
Call Gp_Ms_Collection(SDB_SOK_BOT_SLAB_TEMP, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(SDB_SOK_ZONE_TIME, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(TXT_DIS_UNDIS_IND, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  Call Gp_Ms_Collection(TXT_DISCHARGE_SHIFT, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  Call Gp_Ms_Collection(TXT_DISCHARGE_GROUP, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(TXT_DISCHARGE_EMP, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)

          Call Gp_Ms_Collection(CBO_SLAB_NO, "p", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
    '    Call Gp_Ms_Collection(TXT_PROC_LINE, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(CBO_PLT, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
            Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
            Call Gp_Ms_Collection(cbo_group, " ", " ", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
          Call Gp_Ms_Collection(TXT_UPD_EMP, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
    Call Gp_Ms_Collection(TXT_REJ_OCCR_TIME, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     Call Gp_Ms_Collection(SDB_SLAB_REJ_THK, " ", " ", " ", "i", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
    Call Gp_Ms_Collection(SDB_HEAD_SLAB_WID, " ", " ", " ", "i", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
    Call Gp_Ms_Collection(SDB_TAIL_SLAB_WID, " ", " ", " ", "i", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     Call Gp_Ms_Collection(SDB_REJ_SLAB_LEN, " ", " ", " ", "i", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     Call Gp_Ms_Collection(SDB_REJ_SLAB_WGT, " ", " ", " ", "i", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
        Call Gp_Ms_Collection(TXT_REASON_CD, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
          Call Gp_Ms_Collection(TXT_REJ_LOC, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
           Call Gp_Ms_Collection(TXT_COMFRM, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
        Call Gp_Ms_Collection(TXT_REJ_SHIFT, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
        Call Gp_Ms_Collection(TXT_REJ_GROUP, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
          Call Gp_Ms_Collection(TXT_REJ_EMP, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              ' Call Gp_Ms_Collection(TXT_CD, " ", " ", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
               
     Call Gp_Clear_Collection(CHK_RHF_ROW_A, "s", sControl)
     Call Gp_Clear_Collection(CHK_RHF_ROW_L, "s", sControl)
     Call Gp_Clear_Collection(CHK_RHF_ROW_R, "s", sControl)
     Call Gp_Clear_Collection(CHK_CHA_IND, "s", sControl)
     Call Gp_Clear_Collection(CHK_UNCHA_IND, "s", sControl)
     Call Gp_Clear_Collection(CHK_O_CHA_IND, "s", sControl)
     Call Gp_Clear_Collection(CHK_O_UNCHA_IND, "s", sControl)
     Call Gp_Clear_Collection(CHK_ENTRY, "s", sControl)
     Call Gp_Clear_Collection(CHK_EXIT, "s", sControl)
     Call Gp_Clear_Collection(CHK_ORDER, "s", sControl)
     Call Gp_Clear_Collection(CHK_NONORDER, "s", sControl)
     
     MC.Add Item:=sControl, Key:="sControl"
            
    'MASTER Collection
     Mc1.Add Item:="AGA2010C.P_MODIFY1", Key:="P-M"
     Mc1.Add Item:="AGA2010C.P_REFER1", Key:="P-R"
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"
     
     Mc2.Add Item:="AGA2010C.P_MODIFY2", Key:="P-M"
     Mc2.Add Item:="AGA2010C.P_REFER2", Key:="P-R"
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"
     
     
     Mc3.Add Item:="AGA2010C.P_MODIFY3", Key:="P-M"
     Mc3.Add Item:="AGA2010C.P_REFER3", Key:="P-R"
     Mc3.Add Item:=pControl3, Key:="pControl"
     Mc3.Add Item:=nControl3, Key:="nControl"
     Mc3.Add Item:=mControl3, Key:="mControl"
     Mc3.Add Item:=iControl3, Key:="iControl"
     Mc3.Add Item:=rControl3, Key:="rControl"
     Mc3.Add Item:=cControl3, Key:="cControl"
     Mc3.Add Item:=aControl3, Key:="aControl"
     Mc3.Add Item:=lControl3, Key:="lControl"
     
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
     'Spread_Collection
    sc4.Add Item:=ss1, Key:="Spread"
    sc4.Add Item:="AGA2010C.P_SREFER", Key:="P-R"
    sc4.Add Item:=pColumn1, Key:="pColumn"
    sc4.Add Item:=nColumn1, Key:="nColumn"
    sc4.Add Item:=aColumn1, Key:="aColumn"
    sc4.Add Item:=mColumn1, Key:="mColumn"
    sc4.Add Item:=iColumn1, Key:="iColumn"
    sc4.Add Item:=lColumn1, Key:="lColumn"
    sc4.Add Item:=1, Key:="First"
    sc4.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc4, Key:="Sc"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
     
End Sub

Private Sub CBO_SLAB_NO_Change()
   Dim SMESG As String
      If Len(CBO_SLAB_NO.Text) > 10 Then
      SMESG = "板坯号长度不能超过10位，请确认板坯号 ！！！"
      Call Gp_MsgBoxDisplay(SMESG)
   End If
End Sub

Private Sub CBO_SLAB_NO_Click()
    CBO_SLAB_NO.Text = Mid(CBO_SLAB_NO.Text, 1, 10)
'    Call Form_Ref
End Sub

Private Sub CHK_CHA_IND_Click()

    If CHK_CHA_IND.Value = ssCBUnchecked Then
        If CHK_UNCHA_IND.Value = ssCBUnchecked Then
'           CHK_CHA_IND.Value = ssCBChecked
           TXT_CHA_UNCHA_IND.Text = ""
           CHK_CHA_IND.ForeColor = &H80000012
           CHK_UNCHA_IND.ForeColor = &H80000012
        End If
        Exit Sub
    End If
    
    TXT_CHA_UNCHA_IND.Text = "1"
    
    CHK_CHA_IND.ForeColor = &HFF&
    CHK_CHA_IND.Value = ssCBChecked
    
    CHK_UNCHA_IND.ForeColor = &H808080
    CHK_UNCHA_IND.Value = ssCBUnchecked
    
End Sub

Private Sub CHK_ENTRY_Click()

    If CHK_ENTRY.Value = ssCBUnchecked Then
        If CHK_EXIT.Value = ssCBUnchecked Then
          ' CHK_ENTRY.Value = ssCBChecked
           TXT_REJ_LOC.Text = ""
           CHK_ENTRY.ForeColor = &H80000012
           CHK_EXIT.ForeColor = &H80000012
        End If
        Exit Sub
    End If
    
    TXT_REJ_LOC.Text = "1"
    
    CHK_ENTRY.ForeColor = &HFF&
    CHK_ENTRY.Value = ssCBChecked
    
    CHK_EXIT.ForeColor = &H808080
    CHK_EXIT.Value = ssCBUnchecked
 
End Sub

Private Sub CHK_EXIT_Click()

    If CHK_EXIT.Value = ssCBUnchecked Then
        If CHK_ENTRY.Value = ssCBUnchecked Then
         '  CHK_EXIT.Value = ssCBChecked
           TXT_REJ_LOC.Text = ""
           CHK_EXIT.ForeColor = &H80000012
           CHK_ENTRY.ForeColor = &H80000012
        End If
        Exit Sub
    End If
    
    TXT_REJ_LOC.Text = "2"
    
    CHK_EXIT.ForeColor = &HFF&
    CHK_EXIT.Value = ssCBChecked
    
    CHK_ENTRY.ForeColor = &H808080
    CHK_ENTRY.Value = ssCBUnchecked
   
End Sub

Private Sub CHK_NONORDER_Click()

    If CHK_NONORDER.Value = ssCBUnchecked Then
        If CHK_ORDER.Value = ssCBUnchecked Then
          ' CHK_NONORDER.Value = ssCBChecked
           TXT_COMFRM.Text = ""
           CHK_NONORDER.ForeColor = &H80000012
           CHK_ORDER.ForeColor = &H80000012
        End If
        Exit Sub
    End If
    
    TXT_COMFRM.Text = "2"
    
    CHK_NONORDER.ForeColor = &HFF&
    CHK_NONORDER.Value = ssCBChecked
    
    CHK_ORDER.ForeColor = &H808080
    CHK_ORDER.Value = ssCBUnchecked
   
End Sub

Private Sub CHK_O_CHA_IND_Click()

    If CHK_O_CHA_IND.Value = ssCBUnchecked Then
        If CHK_O_UNCHA_IND.Value = ssCBUnchecked Then
          ' CHK_O_CHA_IND.Value = ssCBChecked
           TXT_DIS_UNDIS_IND.Text = ""
           CHK_O_CHA_IND.ForeColor = &H80000012
           CHK_O_UNCHA_IND.ForeColor = &H80000012
        End If
        Exit Sub
    End If
    
    TXT_DIS_UNDIS_IND.Text = "1"
    
    CHK_O_CHA_IND.ForeColor = &HFF&
    CHK_O_CHA_IND.Value = ssCBChecked
    
    CHK_O_UNCHA_IND.ForeColor = &H808080
    CHK_O_UNCHA_IND.Value = ssCBUnchecked
   
End Sub

Private Sub CHK_O_UNCHA_IND_Click()

    If CHK_O_UNCHA_IND.Value = ssCBUnchecked Then
        If CHK_O_CHA_IND.Value = ssCBUnchecked Then
          ' CHK_O_UNCHA_IND.Value = ssCBChecked
           TXT_DIS_UNDIS_IND.Text = ""
           CHK_O_UNCHA_IND.ForeColor = &H80000012
           CHK_O_CHA_IND.ForeColor = &H80000012
        End If
        Exit Sub
    End If
    
    TXT_DIS_UNDIS_IND.Text = "2"
    
    CHK_O_UNCHA_IND.ForeColor = &HFF&
    CHK_O_UNCHA_IND.Value = ssCBChecked
    
    CHK_O_CHA_IND.ForeColor = &H808080
    CHK_O_CHA_IND.Value = ssCBUnchecked
   
End Sub

Private Sub CHK_ORDER_Click()

    If CHK_ORDER.Value = ssCBUnchecked Then
        If CHK_NONORDER.Value = ssCBUnchecked Then
          ' CHK_ORDER.Value = ssCBChecked
           TXT_COMFRM.Text = ""
           CHK_ORDER.ForeColor = &H80000012
           CHK_NONORDER.ForeColor = &H80000012
        End If
        Exit Sub
    End If
    
    TXT_COMFRM.Text = "1"
    
    CHK_ORDER.ForeColor = &HFF&
    CHK_ORDER.Value = ssCBChecked
    
    CHK_NONORDER.ForeColor = &H808080
    CHK_NONORDER.Value = ssCBUnchecked
   
End Sub


Private Sub CHK_RHF_ROW_A_Click()

    If CHK_RHF_ROW_A.Value = ssCBUnchecked Then
        If CHK_RHF_ROW_L.Value = ssCBUnchecked And CHK_RHF_ROW_R.Value = ssCBUnchecked Then
         '  CHK_RHF_ROW_A.Value = ssCBChecked
           TXT_RHF_CH_ROW.Text = ""
           CHK_RHF_ROW_A.ForeColor = &H80000012
           CHK_RHF_ROW_L.ForeColor = &H80000012
           CHK_RHF_ROW_R.ForeColor = &H80000012
        End If
        Exit Sub
    End If
    
    TXT_RHF_CH_ROW.Text = "0"
    
    CHK_RHF_ROW_A.ForeColor = &HFF&
    CHK_RHF_ROW_A.Value = ssCBChecked
    
    CHK_RHF_ROW_L.ForeColor = &H808080
    CHK_RHF_ROW_L.Value = ssCBUnchecked
       
    CHK_RHF_ROW_R.ForeColor = &H808080
    CHK_RHF_ROW_R.Value = ssCBUnchecked
        
End Sub

Private Sub CHK_RHF_ROW_L_Click()

    If CHK_RHF_ROW_L.Value = ssCBUnchecked Then
        If CHK_RHF_ROW_A.Value = ssCBUnchecked And CHK_RHF_ROW_R.Value = ssCBUnchecked Then
          ' CHK_RHF_ROW_L.Value = ssCBChecked
           TXT_RHF_CH_ROW.Text = ""
            CHK_RHF_ROW_L.ForeColor = &H80000012
            CHK_RHF_ROW_A.ForeColor = &H80000012
            CHK_RHF_ROW_R.ForeColor = &H80000012
        End If
        Exit Sub
    End If
    
    TXT_RHF_CH_ROW.Text = "1"
    
    CHK_RHF_ROW_L.ForeColor = &HFF&
    CHK_RHF_ROW_L.Value = ssCBChecked
    
    CHK_RHF_ROW_A.ForeColor = &H808080
    CHK_RHF_ROW_A.Value = ssCBUnchecked
       
    CHK_RHF_ROW_R.ForeColor = &H808080
    CHK_RHF_ROW_R.Value = ssCBUnchecked
         
End Sub

Private Sub CHK_RHF_ROW_R_Click()

    If CHK_RHF_ROW_R.Value = ssCBUnchecked Then
        If CHK_RHF_ROW_A.Value = ssCBUnchecked And CHK_RHF_ROW_L.Value = ssCBUnchecked Then
           ' CHK_RHF_ROW_R.Value = ssCBChecked
           TXT_RHF_CH_ROW.Text = ""
           CHK_RHF_ROW_R.ForeColor = &H80000012
           CHK_RHF_ROW_A.ForeColor = &H80000012
           CHK_RHF_ROW_L.ForeColor = &H80000012
        End If
        Exit Sub
    End If
    
    TXT_RHF_CH_ROW.Text = "2"
    
    CHK_RHF_ROW_R.ForeColor = &HFF&
    CHK_RHF_ROW_R.Value = ssCBChecked
    
    CHK_RHF_ROW_A.ForeColor = &H808080
    CHK_RHF_ROW_A.Value = ssCBUnchecked
       
    CHK_RHF_ROW_L.ForeColor = &H808080
    CHK_RHF_ROW_L.Value = ssCBUnchecked
   
End Sub

Private Sub CHK_UNCHA_IND_Click()

    If CHK_UNCHA_IND.Value = ssCBUnchecked Then
        If CHK_CHA_IND.Value = ssCBUnchecked Then
          ' CHK_UNCHA_IND.Value = ssCBChecked
            TXT_CHA_UNCHA_IND.Text = ""
            CHK_UNCHA_IND.ForeColor = &H80000012
            CHK_CHA_IND.ForeColor = &H80000012
        End If
        Exit Sub
    End If
    
    TXT_CHA_UNCHA_IND.Text = "2"
    
    CHK_UNCHA_IND.ForeColor = &HFF&
    CHK_UNCHA_IND.Value = ssCBChecked
    
    CHK_CHA_IND.ForeColor = &H808080
    CHK_CHA_IND.Value = ssCBUnchecked
    
End Sub

Private Sub Label5_Click()

End Sub

Private Sub SSCheck1_Click(Value As Integer)

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

        If SSC1 = -1 Then
           SSC1.Value = ssCBUnchecked
        End If
        
        ss1.Row = Row
        ss1.Col = 1
        CBO_SLAB_NO.Text = ss1.Text
        Call Form_Ref
        
End Sub

Private Sub TXT_REASON_CD_Change()
    If TXT_REASON_CD = "" Then
       TXT_REASON_NAME = ""
    End If
End Sub

Private Sub TXT_REASON_CD_DblClick()
    Call TXT_REASON_CD_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_REASON_CD_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "G0001"
        DD.rControl.Add Item:=TXT_REASON_CD
        DD.rControl.Add Item:=TXT_REASON_NAME
    
        DD.nameType = "1"
    
        Call Gf_Common_DD(M_CN1, KeyCode)
    
        Exit Sub
    
    End If

End Sub

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
    
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_Cls(Mc3("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_ControlLock(Mc2("lControl"), True)
    Call Gp_Ms_ControlLock(Mc3("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    Call Gp_Ms_NeceColor(Mc3("nControl"))
    
    Call Gp_Sp_Setting(sc4.Item("Spread"))
    
    Call Gf_Sp_Cls(sc4)
    
    Call Gp_Sp_ColGet(sc4.Item("Spread"), "G-System.INI", Me.Name)
    Call Gf_Mill_ComboAdd(M_CN1, CBO_SLAB_NO, "CA")
    
    If CBO_SLAB_NO.ListCount <> 0 Then
       CBO_SLAB_NO.ListIndex = 0
    End If
    
    TXT_PROC_LINE = "1"
    CBO_PLT.ListIndex = 0
    TXT_UPD_EMP = sUserID '+ ":" + sUsername
      
    sc1.ForeColor = &HFF&
    sc2.ForeColor = &H808080
    sc3.ForeColor = &H808080
    sc1.Value = ssCBChecked
    sc2.Value = ssCBUnchecked
    sc3.Value = ssCBUnchecked
    sf1.Enabled = True
    sf2.Enabled = False
    sf3.Enabled = False
       
    If CBO_SHIFT.Text <> "1" Or CBO_SHIFT.Text <> "2" Or CBO_SHIFT.Text <> "3" Then
       If sShiftSet = "" Or sShiftSet = "0" Then
          sShiftSet = Gf_ShiftSet(M_CN1)
       End If
       CBO_SHIFT.Text = sShiftSet
    End If
       
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc4.Item("Spread"), "G-System.INI", Me.Name)

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
    
    Set pControl3 = Nothing
    Set nControl3 = Nothing
    Set iControl3 = Nothing
    Set rControl3 = Nothing
    Set cControl3 = Nothing
    Set aControl3 = Nothing
    Set lControl3 = Nothing
    Set mControl3 = Nothing
    
    Set sControl = Nothing
    Set MC = Nothing
    
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set sc4 = Nothing
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
    Call Gp_Ms_Cls(Mc3("rControl"))
    Call Gf_Sp_Cls(sc4)

    Call Gp_SSCheck_Cls(MC("sControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    Call Gp_Ms_ControlLock(Mc2("pControl"), False)
    
    sc1.ForeColor = &HFF&
    sc1.Value = ssCBChecked
    sc2.ForeColor = &H808080
    sc2.Value = ssCBUnchecked
    sc3.ForeColor = &H808080
    sc3.Value = ssCBUnchecked
    sf1.Enabled = True
    sf2.Enabled = False
    sf3.Enabled = False
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = False                'Excel
    End With
    
    TXT_RHF_CH_NUM_REF = ""
    TXT_RHF_CH_TIME_UPD = ""
    TXT_PROC_LINE = "1"
    CBO_PLT.ListIndex = 0
    
    If Trim(CBO_SHIFT.Text) = "" Then
       If sShiftSet = "" Or sShiftSet = "0" Then
          sShiftSet = Gf_ShiftSet(M_CN1)
       End If
       CBO_SHIFT.Text = sShiftSet
    End If
    
    TXT_UPD_EMP = sUserID
    
    Call Gf_Mill_ComboAdd(M_CN1, CBO_SLAB_NO, "CA")

    pControl1(1).SetFocus
    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       ' Call Gp_Ms_ControlLock(Mc1("pControl"), False)
     End If

End Sub

Public Sub Form_Ref()

    Dim sKey As String
       
    Call Gp_SSCheck_Cls(MC("sControl"))
    
    If Gf_Ms_Refer(M_CN1, Mc1, , , False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    End If
    
    If Gf_Ms_Refer(M_CN1, Mc2, , , False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    End If
    
    If Gf_Ms_Refer(M_CN1, Mc3, , , False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    End If
    
    If SSC1 = -1 Then
        If Gf_Sp_Refer(M_CN1, sc4, Mc1, , , False) Then
           Call ss1.SetActiveCell(1, ss1.MaxRows)
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        End If
    End If
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = True                 'Excel
    End With
    
    If Trim(TXT_PROC_LINE.Text) = "" Then
       TXT_PROC_LINE.Text = "1"
    End If
     
End Sub

Public Sub Form_Pro()

    Dim SMESG As String
    Dim sLoc As String
    Dim Temp_no As String
    
    Temp_no = CBO_SLAB_NO.Text
    TXT_UPD_EMP.Text = sUserID
    
    If sc1 = -1 Then
       If Not Gp_DateCheck(TXT_RHF_CH_TIME) Then
            SMESG = " 请正确输入装炉时间 ！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
       End If
       If TXT_RHF_CH_TIME_UPD.RawData <> "" Then
            If Not Gp_DateCheck(TXT_RHF_CH_TIME_UPD) Then
                 SMESG = " 请正确输入装炉时间修正 ！"
                 Call Gp_MsgBoxDisplay(SMESG)
                 Exit Sub
            End If
       End If
       If Gf_Mc_Authority(sAuthority, Mc1) Then
            If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
               Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
               CBO_SLAB_NO.Enabled = True
            End If
       End If
    ElseIf sc2 = -1 Then
        If Not Gp_DateCheck(TXT_DISCHARGE_TIME) Then
            SMESG = " 请正确输入出炉时间 ！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
        End If
               
        If Gf_Mc_Authority(sAuthority, Mc2) Then
           If Gf_Ms_Process(M_CN1, Mc2, sAuthority) Then
              Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
              CBO_SLAB_NO.Enabled = True
           End If
        End If
    ElseIf sc3 = -1 Then
    
        If Not Gp_DateCheck(TXT_REJ_OCCR_TIME) Then
            SMESG = " 请正确输入缺号时间 ！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
        End If
       
        If Trim(TXT_REJ_LOC.Text) = "1" Then
           sLoc = "入口"
        Else
           sLoc = "出口"
        End If
    
        SMESG = " 确定此板坯在加热炉 （ " + sLoc + " ）处缺号 ？ "
        
        If Gp_MsgBox(SMESG, "C") = 6 Then
            If Gf_Mc_Authority(sAuthority, Mc3) Then
               If Gf_Ms_Process(M_CN1, Mc3, sAuthority) Then
                  Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
                  CBO_SLAB_NO.Enabled = True
               End If
            End If
        End If
    End If
    
    TXT_RHF_CH_TIME_UPD.RawData = ""
    TXT_RHF_CH_NUM_REF = ""
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = True                 'Excel
    End With
    
    CBO_SLAB_NO.Text = Temp_no
    
'    TXT_PROC_LINE = "1"
   
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub


Private Sub TXT_RHF_CH_TIME_DblClick()

    TXT_RHF_CH_TIME.RawData = Gf_DTSet(M_CN1)
   
End Sub

Private Sub TXT_DISCHARGE_TIME_DblClick()

    TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1)
End Sub

Private Sub TXT_REJ_OCCR_TIME_DblClick()

     TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1)

End Sub

Private Sub sc1_Click(Value As Integer)
    
    If sc1.Value = ssCBUnchecked Then
       If sc2.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked Then
          sc1.Value = ssCBChecked
                   
       End If
        Exit Sub
    End If
  
    sc1.ForeColor = &HFF&
    sc1.Value = ssCBChecked

    sc2.ForeColor = &H808080
    sc2.Value = ssCBUnchecked
    sc3.ForeColor = &H808080
    sc3.Value = ssCBUnchecked
    sf1.Enabled = True
    sf2.Enabled = False
    sf3.Enabled = False
 
   
End Sub

Private Sub sc2_Click(Value As Integer)
   
    If sc2.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked Then
          sc2.Value = ssCBChecked
                   
       End If
        Exit Sub
    End If
  
  '  If sc2.Value = -1 Then    '-1: ssCBChecked
    sc2.ForeColor = &HFF&
    sc1.ForeColor = &H808080
    sc1.Value = ssCBUnchecked
    sc3.ForeColor = &H808080
    sc3.Value = ssCBUnchecked
    sf2.Enabled = True
    sf1.Enabled = False
    sf3.Enabled = False
        
'    Else
'        sc2.Value = ssCBUnchecked
'        sc1.Value = ssCBChecked
  '  End If

End Sub

Private Sub sc3_Click(Value As Integer)
   
    If sc3.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And sc2.Value = ssCBUnchecked Then
          sc3.Value = ssCBChecked
                   
       End If
        Exit Sub
    End If
  
  '  If sc2.Value = -1 Then    '-1: ssCBChecked
    sc3.ForeColor = &HFF&
    sc1.ForeColor = &H808080
    sc1.Value = ssCBUnchecked
    sc2.ForeColor = &H808080
    sc2.Value = ssCBUnchecked
    sf3.Enabled = True
    sf1.Enabled = False
    sf2.Enabled = False
'    Else
'        sc2.Value = ssCBUnchecked
'        sc1.Value = ssCBChecked
  '  End If

End Sub

Private Sub TXT_RHF_CH_TIME_UPD_DblClick()

    TXT_RHF_CH_TIME_UPD.RawData = Gf_DTSet(M_CN1)
End Sub
