VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form CEC0000C 
   Caption         =   "标准板坯设计_CEC0000C"
   ClientHeight    =   7545
   ClientLeft      =   735
   ClientTop       =   2235
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   15300
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "CEC0000C.frx":0000
      Left            =   14520
      List            =   "CEC0000C.frx":0002
      TabIndex        =   47
      Top             =   840
      Width           =   630
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "CEC0000C.frx":0004
      Left            =   14520
      List            =   "CEC0000C.frx":0006
      TabIndex        =   46
      Top             =   450
      Width           =   630
   End
   Begin VB.TextBox TXT_CUST_CD 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   13110
      MaxLength       =   11
      TabIndex        =   44
      Tag             =   "产品"
      Top             =   60
      Width           =   1035
   End
   Begin VB.TextBox txt_stdgrd 
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
      Left            =   11370
      MaxLength       =   11
      TabIndex        =   29
      Top             =   450
      Width           =   1995
   End
   Begin VB.TextBox txt_stdgrd_name 
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
      Left            =   11370
      TabIndex        =   28
      Top             =   840
      Width           =   1995
   End
   Begin VB.TextBox txt_size_knd_name 
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
      TabIndex        =   27
      Tag             =   "钢种"
      Top             =   450
      Width           =   1425
   End
   Begin VB.TextBox txt_size_knd 
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
      Left            =   4515
      MaxLength       =   2
      TabIndex        =   26
      Tag             =   "钢种"
      Top             =   450
      Width           =   465
   End
   Begin VB.TextBox txt_ord_knd 
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
      Left            =   1245
      MaxLength       =   1
      TabIndex        =   25
      Tag             =   "订单种类"
      Top             =   450
      Width           =   465
   End
   Begin VB.TextBox txt_ord_knd_nm 
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
      Left            =   1710
      TabIndex        =   24
      Tag             =   "订单种类"
      Top             =   450
      Width           =   1455
   End
   Begin VB.TextBox txt_stdspec 
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
      Left            =   7815
      MaxLength       =   30
      TabIndex        =   23
      Top             =   450
      Width           =   2205
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1185
      Left            =   15240
      TabIndex        =   17
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2090
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
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
      Begin VB.TextBox txt_des_fl 
         Height          =   270
         Left            =   900
         TabIndex        =   18
         Top             =   30
         Visible         =   0   'False
         Width           =   255
      End
      Begin Threed.SSOption opt_des_all 
         Height          =   390
         Left            =   300
         TabIndex        =   19
         Top             =   45
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   688
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   "全部"
         Value           =   -1
      End
      Begin Threed.SSOption opt_des_not 
         Height          =   390
         Left            =   300
         TabIndex        =   20
         Top             =   390
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   688
         _Version        =   196609
         Font3D          =   1
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
         Caption         =   "待设计"
      End
      Begin Threed.SSOption opt_des_com 
         Height          =   390
         Left            =   300
         TabIndex        =   21
         Top             =   750
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   688
         _Version        =   196609
         Font3D          =   1
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
         Caption         =   "设计完成"
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   495
      Left            =   90
      TabIndex        =   15
      Top             =   1710
      Width           =   16725
      _ExtentX        =   29501
      _ExtentY        =   873
      _Version        =   196609
      BackColor       =   14737918
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSCheck chk_all 
         Height          =   345
         Left            =   360
         TabIndex        =   16
         Top             =   75
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   609
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BackColor       =   14737918
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "设计板坯批次选择"
      End
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   5820
         Top             =   90
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         Caption         =   "设计板坯宽度标准"
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
         ForeColor       =   255
      End
      Begin CSTextLibCtl.sidbEdit sdb_slab_wid 
         Height          =   315
         Left            =   7650
         TabIndex        =   7
         Top             =   90
         Width           =   1365
         _Version        =   262145
         _ExtentX        =   2408
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
         FocusSelect     =   -1  'True
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_prod_len_chg 
         Height          =   315
         Left            =   11760
         TabIndex        =   8
         Tag             =   "产品设计长度"
         Top             =   90
         Width           =   1365
         _Version        =   262145
         _ExtentX        =   2408
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
         FocusSelect     =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.0"
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   9930
         Top             =   90
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         Caption         =   "产品设计长度"
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
         ForeColor       =   255
      End
      Begin Threed.SSCommand cmd_change 
         Height          =   375
         Left            =   13230
         TabIndex        =   22
         Top             =   60
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   12583104
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "变更产品设计长度"
         BevelWidth      =   3
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   2490
         Top             =   90
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         Caption         =   "设计板坯厚度标准"
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
         ForeColor       =   255
      End
      Begin CSTextLibCtl.sidbEdit sdb_slab_thk 
         Height          =   315
         Left            =   4320
         TabIndex        =   6
         Top             =   90
         Width           =   1365
         _Version        =   262145
         _ExtentX        =   2408
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
         FocusSelect     =   -1  'True
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
         Undo            =   0
         Data            =   0
      End
   End
   Begin VB.ComboBox cbo_ord_item 
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
      Left            =   2550
      TabIndex        =   1
      Top             =   70
      Width           =   630
   End
   Begin VB.TextBox txt_ord_no 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1245
      MaxLength       =   11
      TabIndex        =   0
      Tag             =   "产品"
      Top             =   75
      Width           =   1305
   End
   Begin VB.TextBox txt_plt_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   17370
      MaxLength       =   50
      TabIndex        =   14
      Tag             =   "工厂"
      Top             =   1080
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox txt_plt 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   17190
      MaxLength       =   2
      TabIndex        =   13
      Tag             =   "工厂"
      Top             =   1080
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txt_prc_line 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   17910
      MaxLength       =   1
      TabIndex        =   9
      Tag             =   "机号"
      Top             =   1080
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txt_prod_cd_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   18690
      MaxLength       =   40
      TabIndex        =   11
      Tag             =   "产品"
      Top             =   1080
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txt_prod_cd 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   18450
      MaxLength       =   2
      TabIndex        =   10
      Tag             =   "产品"
      Top             =   1080
      Visible         =   0   'False
      Width           =   210
   End
   Begin InDate.UDate txt_del_fr 
      Height          =   315
      Left            =   4515
      TabIndex        =   2
      Top             =   60
      Width           =   1440
      _ExtentX        =   2540
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
      Left            =   18210
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
      _ExtentX        =   318
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   3360
      Top             =   60
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "交货期"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   6915
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2250
      Width           =   16770
      _Version        =   393216
      _ExtentX        =   29580
      _ExtentY        =   12197
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ButtonDrawMode  =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   50
      MaxRows         =   2
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "CEC0000C.frx":0008
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   17700
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   556
      Caption         =   "机号"
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
      Left            =   16890
      Top             =   1080
      Visible         =   0   'False
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   556
      Caption         =   "工厂"
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   90
      Top             =   75
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "订单号"
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
      ForeColor       =   0
   End
   Begin InDate.UDate udt_release_date_to 
      Height          =   315
      Left            =   10185
      TabIndex        =   5
      Top             =   60
      Width           =   1440
      _ExtentX        =   2540
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
   Begin InDate.UDate udt_release_date_fr 
      Height          =   315
      Left            =   8745
      TabIndex        =   4
      Top             =   60
      Width           =   1440
      _ExtentX        =   2540
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
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Left            =   7590
      Top             =   60
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "投入日期"
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
   Begin InDate.UDate txt_del_to 
      Height          =   315
      Left            =   5970
      TabIndex        =   3
      Top             =   60
      Width           =   1440
      _ExtentX        =   2540
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
   Begin CSTextLibCtl.sidbEdit sdb_prod_thk_fr 
      Height          =   315
      Left            =   1245
      TabIndex        =   30
      Top             =   840
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
      FocusSelect     =   -1  'True
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
      NumDecDigits    =   2
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   90
      Top             =   840
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "产品厚度"
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   3360
      Top             =   840
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "产品宽度"
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
      Left            =   6660
      Top             =   840
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "产品长度"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   10200
      Top             =   450
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "钢种"
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
      ForeColor       =   16711680
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_thk_to 
      Height          =   315
      Left            =   2190
      TabIndex        =   31
      Top             =   840
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
      FocusSelect     =   -1  'True
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
      NumDecDigits    =   2
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_len_fr 
      Height          =   315
      Left            =   7815
      TabIndex        =   32
      Top             =   840
      Width           =   1095
      _Version        =   262145
      _ExtentX        =   1931
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
      FocusSelect     =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   " 0.0"
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
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_len_to 
      Height          =   315
      Left            =   8910
      TabIndex        =   33
      Top             =   840
      Width           =   1095
      _Version        =   262145
      _ExtentX        =   1931
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
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
      Text            =   " 0.0"
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
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_wid_fr 
      Height          =   315
      Left            =   4515
      TabIndex        =   34
      Top             =   840
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
      FocusSelect     =   -1  'True
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
      NumDecDigits    =   2
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_wid_to 
      Height          =   315
      Left            =   5460
      TabIndex        =   35
      Top             =   840
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
      FocusSelect     =   -1  'True
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
      NumDecDigits    =   2
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_thk_fr 
      Height          =   315
      Left            =   1245
      TabIndex        =   36
      Top             =   1230
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
      FocusSelect     =   -1  'True
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   90
      Top             =   1230
      Width           =   1140
      _ExtentX        =   2011
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Left            =   3360
      Top             =   1230
      Width           =   1140
      _ExtentX        =   2011
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
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   6660
      Top             =   1230
      Width           =   1140
      _ExtentX        =   2011
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_thk_to 
      Height          =   315
      Left            =   2190
      TabIndex        =   37
      Top             =   1230
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
      FocusSelect     =   -1  'True
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
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_len_fr 
      Height          =   315
      Left            =   7815
      TabIndex        =   38
      Top             =   1230
      Width           =   1095
      _Version        =   262145
      _ExtentX        =   1931
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
      FocusSelect     =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   7
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_len_to 
      Height          =   315
      Left            =   8910
      TabIndex        =   39
      Top             =   1230
      Width           =   1095
      _Version        =   262145
      _ExtentX        =   1931
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
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
      NumIntDigits    =   7
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_wid_fr 
      Height          =   315
      Left            =   4515
      TabIndex        =   40
      Top             =   1230
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
      FocusSelect     =   -1  'True
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
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_wid_to 
      Height          =   315
      Left            =   5460
      TabIndex        =   41
      Top             =   1230
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
      FocusSelect     =   -1  'True
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel15 
      Height          =   315
      Left            =   3360
      Top             =   450
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "定尺区分"
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
      ForeColor       =   16711680
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_wgt_fr 
      Height          =   315
      Left            =   11370
      TabIndex        =   42
      Top             =   1230
      Width           =   1005
      _Version        =   262145
      _ExtentX        =   1773
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
      FocusSelect     =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   7
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_wgt_to 
      Height          =   315
      Left            =   12375
      TabIndex        =   43
      Top             =   1230
      Width           =   1005
      _Version        =   262145
      _ExtentX        =   1773
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
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
      NumIntDigits    =   7
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   10200
      Top             =   1230
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "产品重量"
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
   Begin InDate.ULabel ULabel19 
      Height          =   315
      Left            =   90
      Top             =   450
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "订单种类"
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel20 
      Height          =   315
      Left            =   6660
      Top             =   450
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "标准号"
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel21 
      Height          =   315
      Left            =   10200
      Top             =   840
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "钢种说明"
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel22 
      Height          =   315
      Left            =   11820
      Top             =   60
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "客户代码"
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
      ForeColor       =   16711680
   End
   Begin Threed.SSCheck chk_key 
      Height          =   345
      Left            =   14280
      TabIndex        =   45
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   196609
      Font3D          =   1
      BackColor       =   12632319
      BackStyle       =   1
      Caption         =   "重点订单"
   End
   Begin InDate.ULabel ULabel23 
      Height          =   315
      Left            =   13440
      Top             =   450
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "热处理对象"
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
   Begin InDate.ULabel ULabel24 
      Height          =   315
      Left            =   13440
      Top             =   840
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "是否真空"
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
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   75
      X2              =   15240
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   75
      X2              =   15240
      Y1              =   1665
      Y2              =   1665
   End
End
Attribute VB_Name = "CEC0000C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       DAILY SCHEDULE
'-- Sub_System Name
'-- Program Name      SLAB DESIGN
'-- Program ID        CEC0000C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM SUNG HO
'-- Coder             KIM SUNG HO
'-- Date              2010.09.21
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

Public Design_Change As Boolean
Dim Active_Row As Long

Dim iCount As Integer

Const SPD_ORD_NO = 1
Const SPD_ORD_ITEM = 2
Const SPD_SIZE_KND = 12
Const SPD_LEN = 15
Const SPD_ORDWGT = 26
Const SPD_ORDCNT = 27
Const SPD_ORDREMWGT = 30
Const SPD_ORDREMCNT = 31
Const SPD_USERID = 46
'Const SPD_USERNAME = 43


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
             Call Gp_Ms_Collection(txt_des_fl, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_plt_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_prc_line, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_prod_cd, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_prod_cd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(cbo_ord_item, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_stdgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_stdgrd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_del_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_del_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udt_release_date_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udt_release_date_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_prod_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_prod_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_prod_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_prod_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_prod_len_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_prod_len_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_prod_wgt_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_prod_wgt_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_slab_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_slab_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_slab_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_slab_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_slab_len_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_slab_len_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_slab_thk, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_slab_wid, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_prod_len_chg, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_size_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_size_knd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ord_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ord_knd_nm, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(TXT_CUST_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(chk_key, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(Combo1, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(Combo2, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             
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
     Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss1, 2, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
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
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CEC0000C.P_REFER", Key:="P-R"
    sc1.Add Item:="CEC0000C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="CEC0000C.P_ONEROW", Key:="P-O"
    
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, 22, True)
    Call Gp_Sp_ColHidden(ss1, 39, True)
    Call Gp_Sp_ColHidden(ss1, 40, True)
    Call Gp_Sp_ColHidden(ss1, 41, True)
    Call Gp_Sp_ColHidden(ss1, 42, True)
    Call Gp_Sp_ColHidden(ss1, 43, True)
    Call Gp_Sp_ColHidden(ss1, 44, True)
    Call Gp_Sp_ColHidden(ss1, 45, True)
    Call Gp_Sp_ColHidden(ss1, 46, True)
    
End Sub

Private Sub chk_all_Click(Value As Integer)

    Dim lRow As Long
    
    For lRow = 1 To ss1.MaxRows - 1
    
        ss1.Row = lRow
        ss1.Col = 0
        
        If chk_all.Value Then
            If ss1.Text <> "设计" Then
                'ss1.Col = 16    'SLAB_WID
                'If ss1.Value = 0 Then
                    ss1.Col = 0:    ss1.Text = "设计"
                    ss1.Col = SPD_USERID:   ss1.Text = sUserID
                    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lRow, lRow, , &HFFFF80)
                'End If
            End If
        Else
            ss1.Col = 0:     ss1.Text = ""
            ss1.Col = SPD_USERID:    ss1.Text = ""
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lRow, lRow)
            ss1.Col = 12
            If Trim(ss1.Text) <> "定尺" Then
                Call Gp_Sp_CellColor(ss1, SPD_LEN, lRow, , &HC0FFFF)
                ss1.Col = SPD_LEN:    ss1.Lock = False
            Else
                ss1.Col = SPD_LEN:    ss1.Lock = True
            End If
        End If
        
    Next lRow
    
End Sub

Private Sub cmd_change_Click()

    Dim iRow As Integer
    Dim dMax As Double
    Dim dMin As Double
    
    If ss1.MaxRows < 1 Then Exit Sub
    If sdb_prod_len_chg.Value = 0 Then
        Call Gp_MsgBoxDisplay(sdb_prod_len_chg.Tag + "必须输入", "", "错误提示")
        Exit Sub
    End If
    
    For iRow = 1 To ss1.MaxRows - 1
        
        ss1.Row = iRow
        ss1.Col = 12
        If Trim(ss1.Text) <> "定尺" Then
            
            ss1.Col = 13:    dMin = ss1.Value
            ss1.Col = 14:    dMax = ss1.Value
            
            If dMin <= sdb_prod_len_chg.Value And dMax >= sdb_prod_len_chg.Value Then
                ss1.Col = 0:     ss1.Text = "设计"
                ss1.Col = SPD_LEN:    ss1.Value = sdb_prod_len_chg.Value
                ss1.Col = SPD_USERID:    ss1.Text = sUserID
                'ss1.Col = SPD_USERNAME:    ss1.Value = sUserName
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFFFF80)
            End If
            
        End If
            
    Next iRow
    
End Sub

Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuTool_ReSet
   
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
    
    Combo1.AddItem " "
    Combo1.AddItem "A"
    Combo1.AddItem "N"
    Combo1.AddItem "Q"
    Combo1.AddItem "T"
    
    Combo2.AddItem " "
    Combo2.AddItem "Y"
    Combo2.AddItem "N"
    
    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "E-System.INI", Me.Name)
    
    txt_plt.Text = "C3"
    Call txt_plt_KeyUp(0, 0)
    
    txt_prod_cd.Text = "PP"
    txt_des_fl.Text = "1"
    opt_des_all.Value = True
    
    txt_prc_line.Text = ""
    Design_Change = False
    
    txt_del_fr.RawData = ""
    txt_del_to.RawData = ""
    udt_release_date_fr.Text = Mid(DateAdd("M", -1, udt_release_date_to.Text), 1, 8) & "20"
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "E-System.INI", Me.Name)
    
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
    
    Dim iRow As Long
    Dim sQuery As String
    
    ss1.ReDraw = False
    
    For iRow = ss1.SelBlockRow To ss1.SelBlockRow2
        ss1.Row = iRow
        ss1.Col = 0
        
        If ss1.Text = "设计" Then
            sQuery = Gf_Sp_MakeQuery(Proc_Sc("Sc")("Spread"), sc1.Item("P-O"), "O", sc1.Item("pColumn"), iRow)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, Proc_Sc("Sc")("Spread"), iRow)
            ss1.Col = 0:    ss1.Text = ""
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow)
            
            ss1.Col = 12
            If Trim(ss1.Text) <> "定尺" Then
                Call Gp_Sp_CellColor(ss1, SPD_LEN, iRow, , &HC0FFFF)
                ss1.Col = SPD_LEN:    ss1.Lock = False
            End If
        End If
        
    Next iRow
    
    ss1.ReDraw = True
    
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(7).SetFocus
        
        txt_plt.Text = "C3"
        Call txt_plt_KeyUp(0, 0)
        
        cbo_ord_item.Clear
        
        txt_prod_cd.Text = "PP"
        txt_des_fl.Text = "1"
        opt_des_all.Value = True
        chk_all.Value = False
        txt_del_fr.RawData = ""
        txt_del_to.RawData = ""
        udt_release_date_fr.RawData = Mid(udt_release_date_to.RawData, 1, 6) & "01"
    End If

End Sub

Public Sub Form_Ref()

    Dim lRow            As Long
    Dim iCount          As Integer
    Dim iOrd_Wgt        As Double
    Dim iOrd_RemWgt     As Double
    Dim iOrd_Cnt        As Double
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    chk_all.Value = ssCBUnchecked
    sdb_slab_thk.Value = 0
    sdb_slab_wid.Value = 0
    sdb_prod_len_chg.Value = 0
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss1.OperationMode = OperationModeNormal
        Call MenuTool_ReSet
        chk_all.Value = False
        
        For lRow = 1 To ss1.MaxRows
            ss1.Row = lRow
            ss1.Col = 12
            If Trim(ss1.Text) <> "定尺" Then
                ss1.Col = SPD_LEN:    ss1.Lock = False
                Call Gp_Sp_CellColor(ss1, SPD_LEN, lRow, , &HC0FFFF)
            Else
                ss1.Col = SPD_LEN:    ss1.Lock = True
                Call Gp_Sp_CellColor(ss1, SPD_LEN, lRow, , vbWhite)
            End If
        Next lRow
        
    End If
    
    iOrd_Wgt = 0
    iOrd_RemWgt = 0
    iOrd_Cnt = 0
        
    With ss1
        If .MaxRows = 0 Then
            Exit Sub
        End If
        .MaxRows = .MaxRows + 1
        For iCount = 1 To .MaxRows - 1
            .Row = iCount

            .Col = SPD_ORDWGT
             iOrd_Wgt = iOrd_Wgt + .Value

            .Col = SPD_ORDREMWGT
             iOrd_RemWgt = iOrd_RemWgt + .Value

'            .Col = SPD_SIZE_KND
'             If .Text = "单定尺" Then
'                 .Col = SPD_ORDCNT
'                 .Value = 0
'                 .Col = SPD_ORDREMCNT
'                 .Value = 0
'             End If

        Next iCount
        
        .Row = .MaxRows
        .Col = SPD_ORD_NO
        .Text = "汇总"
        .Col = SPD_ORDWGT
        .Value = iOrd_Wgt
        .Col = SPD_ORDREMWGT
        .Value = iOrd_RemWgt
'            Call Gp_Sp_ColLock(ss1, SPD_LEN, True)
        Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HE6E6FF)
        .Col = SPD_LEN: .Col2 = SPD_LEN
        .Row = .MaxRows: .Row2 = .MaxRows
        .Lock = True

    End With
    
    '重点订单红色标记 2013-11-20  by  CaoLei
    Call SS1_CHANGE_COLOR
    
    Call ss1.SetActiveCell(1, ss1.MaxRows)
    
End Sub

Private Sub SS1_CHANGE_COLOR()

    With ss1

        If .MaxRows <= 0 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .Row = iCount

             '重点订单红色标记 2013-11-20  by  CaoLei
            ss1.Row = .Row:          ss1.Col = 48
            If ss1.Text = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, 1, 2, .Row, .Row, &HFF&)
                 Call Gp_Sp_BlockColor(ss1, 48, 48, .Row, .Row, &HFF&)
            End If

        Next iCount

    End With

End Sub

Public Sub Form_Pro()

    Dim lRow            As Long
    Dim iCount          As Integer
    Dim iOrd_Wgt        As Double
    Dim iOrd_RemWgt     As Double
    Dim iOrd_Cnt        As Double
    
    If ss1.MaxRows <= 0 Then Exit Sub
    
'    If sdb_slab_wid.Value = 0 Then
'        Call Gp_MsgBoxDisplay("设计板坯宽度标准必须输入", , Me.Caption)
'        Exit Sub
'    End If
'
    If Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        chk_all.Value = False
        
        For lRow = 1 To ss1.MaxRows
            ss1.Row = lRow
            ss1.Col = 12
            If Trim(ss1.Text) <> "定尺" Then
                ss1.Col = SPD_LEN:    ss1.Lock = False
                Call Gp_Sp_CellColor(ss1, SPD_LEN, lRow, , &HC0FFFF)
            Else
                ss1.Col = SPD_LEN:    ss1.Lock = True
                Call Gp_Sp_CellColor(ss1, SPD_LEN, lRow, , vbWhite)
            End If
        Next lRow
        
        iOrd_Wgt = 0
        iOrd_RemWgt = 0
        iOrd_Cnt = 0
            
        With ss1
        
            If .MaxRows = 0 Then
                Exit Sub
            End If
            
            .MaxRows = .MaxRows + 1
            
            For iCount = 1 To .MaxRows - 1
                .Row = iCount
    
                .Col = SPD_ORDWGT
                 iOrd_Wgt = iOrd_Wgt + .Value
    
                .Col = SPD_ORDREMWGT
                 iOrd_RemWgt = iOrd_RemWgt + .Value
    
    '            .Col = SPD_SIZE_KND
    '             If .Text = "单定尺" Then
    '                 .Col = SPD_ORDCNT
    '                 .Value = 0
    '                 .Col = SPD_ORDREMCNT
    '                 .Value = 0
    '             End If
    
            Next iCount
            
            .Row = .MaxRows
            .Col = SPD_ORD_NO
            .Text = "汇总"
            .Col = SPD_ORDWGT
            .Value = iOrd_Wgt
            .Col = SPD_ORDREMWGT
            .Value = iOrd_RemWgt
    '            Call Gp_Sp_ColLock(ss1, SPD_LEN, True)
            Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HE6E6FF)
            
            .Col = SPD_LEN: .Col2 = SPD_LEN
            .Row = .MaxRows: .Row2 = .MaxRows
            .Lock = True
    
        End With
        
        Call ss1.SetActiveCell(1, ss1.MaxRows)
    
    End If
    
End Sub

Public Sub Form_Ins()
    
End Sub

Public Sub Spread_Cpy()

End Sub

Public Sub Spread_Pst()

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

End Sub

Private Sub opt_des_all_Click(Value As Integer)

    txt_des_fl.Text = "1"
    opt_des_all.ForeColor = &HFF&
    opt_des_not.ForeColor = &H808080
    opt_des_com.ForeColor = &H808080
    
End Sub

Private Sub opt_des_com_Click(Value As Integer)

    txt_des_fl.Text = "3"
    opt_des_all.ForeColor = &H808080
    opt_des_not.ForeColor = &H808080
    opt_des_com.ForeColor = &HFF&
    
End Sub

Private Sub opt_des_not_Click(Value As Integer)

    txt_des_fl.Text = "2"
    opt_des_all.ForeColor = &H808080
    opt_des_not.ForeColor = &HFF&
    opt_des_com.ForeColor = &H808080

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If Row <= 0 Then Exit Sub
    
    If Col >= 17 And Col <= 20 Then Exit Sub
        
    ss1.Row = Row
    ss1.Col = 1
    If ss1.Text = "汇总" Then Exit Sub
    
    ss1.Col = 0
    
    If ss1.Text = "" Then
        ss1.Col = 0:    ss1.Text = "设计"
        ss1.Col = SPD_USERID:    ss1.Text = sUserID
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
    Else
        ss1.Col = 0:    ss1.Text = ""
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
        
        ss1.Col = 12
        If Trim(ss1.Text) <> "定尺" Then
            Call Gp_Sp_CellColor(ss1, SPD_LEN, Row, , &HC0FFFF)
        End If
        
        Call SS1_CHANGE_COLOR
        
    End If
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim sQuery As String
    
    If Row <= 0 Then Exit Sub
    
    If Col < 17 Or Col > 23 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 1
    If ss1.Text = "汇总" Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 17
    
    'SLAB_WID
    'If ss1.Value <> 0 Then
        
        Active_Row = Row
        Design_Change = False
        
        Unload CEC0010C
    
        ss1.Row = Row
        ss1.Col = 1
        CEC0010C.txt_ord_no.Text = Trim(ss1.Text)
        
        ss1.Row = Row
        ss1.Col = 2
        CEC0010C.txt_ord_item.Text = Trim(ss1.Text)
        
        CEC0010C.Active_CForm = "CEC0010C"
        
        CEC0010C.Show 1
        
        If Design_Change Then
            sQuery = Gf_Sp_MakeQuery(Proc_Sc("Sc")("Spread"), sc1.Item("P-O"), "O", sc1.Item("pColumn"), Active_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, Proc_Sc("Sc")("Spread"), Active_Row)
            Design_Change = False
        End If
    
    'End If
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    ss1.Row = Row
    ss1.Col = 1
    If ss1.Text = "汇总" Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), SPD_USERID)
        ss1.Col = 0
        ss1.Row = Row
        
        If ss1.Text = "Update" Then
            ss1.Col = 0:    ss1.Text = "设计"
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
        End If
        
    End If

End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    Dim iCol As Integer
    Dim iRow As Integer
    Dim dMin As Double
    Dim dMax As Double
    Dim cValue As Double
    Dim sQuery As String
   
    If Row < 0 Or Row = 0 Or Row = ss1.MaxRows Then Exit Sub
    
    With ss1
            
            If .CellTag = "False" Then Exit Sub
            
            .Row = Row
                  
            Select Case Col
            
                Case SPD_LEN      'Design Product Len
                
                    .Col = Col
                    cValue = .Value
                    
                    .Col = Col - 2
                    If .Value = "" Then
                        dMin = 0
                    Else
                        dMin = .Value
                    End If
                    
                    .Col = Col - 1
                    If .Value = "" Then
                        dMax = 0
                    Else
                        dMax = .Value
                    End If
                                    
                    If cValue > dMax Or cValue < dMin Then
                    
                        .Col = Col
                        .Row = Row
                        .CellTag = "False"
                     
                        Call Gp_MsgBoxDisplay("已超出最大/最小值...!!")
                      
                        .Col = Col
                        .Row = Row
                        .CellTag = ""
                        
                        .Value = 0
                        .TabStop = True
                        .SetFocus
                        .SetActiveCell Col, Row
                        .Action = SS_ACTION_ACTIVE_CELL
                        .EditMode = True
                        .TabStop = False
                        
                        sQuery = Gf_Sp_MakeQuery(Proc_Sc("Sc")("Spread"), sc1.Item("P-O"), "O", sc1.Item("pColumn"), Row)
                        Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, Proc_Sc("Sc")("Spread"), Row)
                        .Col = 0:    .Text = ""
                        Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, Row, Row)
                        
                        .Col = 12
                        If Trim(.Text) <> "定尺" Then
                            Call Gp_Sp_CellColor(ss1, SPD_LEN, Row, , &HC0FFFF)
                            .Col = SPD_LEN:    .Lock = False
                        End If
        
                    End If
               
            End Select
            
   End With

End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub txt_cust_cd_DblClick()
    Call txt_cust_cd_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_CUST_CD

        DD.nameType = "1"
        Call Gf_Customer_DD(M_CN1, KeyCode)

    End If
    
End Sub

Private Sub txt_ord_knd_DblClick()

    Call txt_ord_knd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_ord_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0009"

        DD.rControl.Add Item:=txt_ord_knd
        DD.rControl.Add Item:=txt_ord_knd_nm

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_ord_knd.Text)) = txt_ord_knd.MaxLength Then
            txt_ord_knd_nm.Text = Gf_ComnNameFind(M_CN1, "B0009", txt_ord_knd.Text, 2)
            Exit Sub
        Else
            txt_ord_knd_nm.Text = ""
        End If
        
    End If
    
End Sub

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String

    If Len(Trim(txt_ord_no.Text)) = txt_ord_no.MaxLength Then
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(txt_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, cbo_ord_item, sQuery)
    Else
        cbo_ord_item.Clear
    End If
    
End Sub

Private Sub txt_plt_DblClick()

    Call txt_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=txt_plt_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
        txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        txt_plt_name.Text = ""
    End If

End Sub

Private Sub txt_prdo_cd_DblClick()

    Call txt_prod_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "B0005"
        
        DD.rControl.Add Item:=txt_prod_cd
        DD.rControl.Add Item:=txt_prod_cd_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else

        If Len(Trim(txt_prod_cd.Text)) = txt_prod_cd.MaxLength Then
            txt_prod_cd_name.Text = Gf_ComnNameFind(M_CN1, "B0005", Trim(txt_prod_cd.Text), 2)
        Else
            txt_prod_cd_name.Text = ""
        End If
    End If
End Sub

Private Sub txt_size_knd_DblClick()
    
    Call txt_size_knd_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_size_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"

        DD.rControl.Add Item:=txt_size_knd
        DD.rControl.Add Item:=txt_size_knd_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_size_knd.Text)) = txt_size_knd.MaxLength Then
            txt_size_knd_name.Text = Gf_ComnNameFind(M_CN1, "B0043", txt_size_knd.Text, 2)
            Exit Sub
        Else
            txt_size_knd_name.Text = ""
        End If
        
    End If
    
End Sub

Private Sub TxT_stdgrd_DblClick()

    Call TxT_stdgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TxT_stdgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        
        DD.nameType = "1"
        DD.sWitch = "MS"
        
        DD.rControl.Add Item:=txt_stdgrd
        DD.rControl.Add Item:=txt_stdgrd_name
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_stdgrd.Text)) = txt_stdgrd.MaxLength Then
            txt_stdgrd_name.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stdgrd.Text))
        Else
            txt_stdgrd_name.Text = ""
        End If
        
    End If
    
End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

Private Function Sp_Process(Conn As ADODB.Connection, Sc As Collection, Optional MC As Collection, _
                            Optional RefChek As Boolean = False) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iCount, iProcessCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim adoCmd As ADODB.Command

    Sp_Process = True
    
    Screen.MousePointer = vbHourglass
    Sc.Item("Spread").ReDraw = False
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Sp_Process = False: Exit Function
    End If
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = Conn
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = Sc.Item("P-M")
    
    Conn.BeginTrans
    
    'Create Parameter (Input) iType + iColumn
    For iCount = 0 To 6
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        ss1.Row = iCount
        ss1.Col = 0
        
        If ss1.Text = "设计" Then
        
            adoCmd.Parameters(0).Value = "U"
            ss1.Col = SPD_ORD_NO    'ORD_NO
            adoCmd.Parameters(1).Value = ss1.Text
            ss1.Col = SPD_ORD_ITEM   'ORD_ITEM
            adoCmd.Parameters(2).Value = ss1.Text
            ss1.Col = SPD_LEN   'PROD_LEN
            adoCmd.Parameters(3).Value = ss1.Value
            'SLAB_THK
            adoCmd.Parameters(4).Value = sdb_slab_thk.Value
            'SLAB_WID
            adoCmd.Parameters(5).Value = sdb_slab_wid.Value
            ss1.Col = SPD_USERID   'EMP_ID
            adoCmd.Parameters(6).Value = sUserID
            
            adoCmd.Execute
                
            'Error Check
            If adoCmd("Error") <> "0" Then
            
                ret_Result_ErrCode = adoCmd("Error")
                ret_Result_ErrMsg = adoCmd("Messg")
        
                sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                
                Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                Screen.MousePointer = vbDefault
                Set adoCmd = Nothing
                
                Conn.RollbackTrans
                Sp_Process = False
                Exit Function
        
             End If
             
             iProcessCount = iProcessCount + 1
        
        End If
        
    Next iCount
    
    Conn.CommitTrans
    
    ' 0 Column Space
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
        
            Case "设计"
                Call Gp_Sp_SendData(Sc.Item("Spread"), "", 0, iCount)
                
        End Select
        
    Next iCount
    
    Sc.Item("Spread").ReDraw = True
    
    If iProcessCount > 0 Then
        If Not MC Is Nothing Then
            If RefChek = False Then Call Gf_Sp_Display(Conn, Sc.Item("Spread"), _
                                                    Gf_Ms_MakeQuery(Sc.Item("P-R"), "R", MC("pControl")), Sc.Item("pColumn"), False)
                                                    
        Else
            If RefChek = False Then Call Gf_Sp_Display(Conn, Sc.Item("Spread"), _
                           Gf_Sp_MakeQuery(Sc.Item("Spread"), Sc.Item("P-R"), "R", Sc.Item("aColumn"), 1), Sc.Item("pColumn"), False)
        End If
        
        MDIMain.StatusBar1.Panels(1) = "提示信息 ：成功处理了" & iProcessCount & "条记录"
        
    End If
            
    If iProcessCount > 0 Then
        If Not MC Is Nothing Then
            Call Gp_Ms_ControlLock(MC.Item("lControl"), True)
        End If
    Else
        Sp_Process = False
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Function

SpreadPro_Error:
    
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Sp_Process = False
    Call Gp_MsgBoxDisplay("Gf_Sp_Process Error : " & Error)
    Screen.MousePointer = vbDefault

End Function

Private Sub txt_stdspec_DblClick()

    Call txt_stdspec_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stdspec_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec

        Call Gf_StdSPEC_DD_Y(M_CN1, KeyCode)
        
    End If
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_StdSPEC_DD_Y
'   2.Name         : StdSPEC Code Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 06 .20
'   7.Modify Date  :
'   8.Comment      : StdSPEC Code Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_StdSPEC_DD_Y(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    Dim iCount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "T"        'StdSPEC Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "            SELECT StdSPEC ""标准代号"", StdSPEC_YY ""发布年度"", STDSPEC_CHR_CD ""标准特性代码"", "
        DD.sQuery = DD.sQuery + "       Gf_ComnNameFind('Q0025',STDSPEC_CHR_CD) ""标准特性名称"", "
        DD.sQuery = DD.sQuery + "       STDSPEC_NAME_ENG ""标准英文名"", STDSPEC_NAME_CHN ""标准中文名"" FROM  NISCO.QP_STD_HEAD "
        DD.sWhere = "             WHERE StdSPEC like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + "   AND STDSPEC_CHR_CD IN ('Y','2') "
            
        If DD.rControl.Count > 1 Then
            DD.sWhere = DD.sWhere + " AND NVL(StdSPEC_YY,'0')   like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  StdSPEC  ASC "
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "            SELECT StdSPEC ""标准代号"", StdSPEC_YY ""发布年度"", STDSPEC_CHR_CD ""标准特性代码"", "
        DD.sQuery = DD.sQuery + "       Gf_ComnNameFind('Q0025',STDSPEC_CHR_CD) ""标准特性名称"", "
        DD.sQuery = DD.sQuery + "       STDSPEC_NAME_ENG ""标准英文名"", STDSPEC_NAME_CHN ""标准中文名"" FROM  NISCO.QP_STD_HEAD "
        DD.sWhere = "             WHERE StdSPEC like '" & Trim(DD.sPname.Text) & "%' "
        DD.sWhere = DD.sWhere + "   AND STDSPEC_CHR_CD IN ('Y','2') "
            
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            DD.sWhere = DD.sWhere + " AND NVL(StdSPEC_YY,'0')   like '" & Trim(DD.sPname.Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  StdSPEC  ASC "
   
    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
        If DD.sWitch = "SP" Then
            
            DD.sPname.Col = DD.rControl.Item(1)
            sNew_Code = DD.sPname.Text
            
            If DD.rControl.Count > 1 Then
                DD.sPname.Col = DD.rControl.Item(2)
                sNew_Name = DD.sPname.Text
            End If
            
            DD.sPname.TabStop = True
            DD.sPname.SetFocus
            DD.sPname.SetActiveCell DD.rControl.Item(1), DD.sPname.ActiveRow
            DD.sPname.Action = SS_ACTION_ACTIVE_CELL
            DD.sPname.EditMode = True
            DD.sPname.TabStop = False
            
            If DD.sSelect Then
                If sOld_Code <> sNew_Code Then Call Gp_Sp_UpdateMake(DD.sPname, False)
            End If
            
        End If
    
    End If
    
    DD.sWitch = ""
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function
