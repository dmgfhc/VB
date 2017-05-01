VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB4020C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "转库作业实绩录入_ACB4020C"
   ClientHeight    =   10560
   ClientLeft      =   75
   ClientTop       =   1470
   ClientWidth     =   15165
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_loc 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11310
      MaxLength       =   7
      TabIndex        =   28
      Top             =   1020
      Width           =   1935
   End
   Begin VB.TextBox Text_size_knd_name_IN 
      Enabled         =   0   'False
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
      Left            =   6480
      TabIndex        =   26
      Tag             =   "钢种"
      Top             =   660
      Width           =   930
   End
   Begin VB.ComboBox cbo_prod_grd 
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
      Left            =   11310
      TabIndex        =   25
      Top             =   660
      Width           =   1965
   End
   Begin VB.TextBox txt_Trim_NAME 
      Enabled         =   0   'False
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
      Left            =   9105
      TabIndex        =   24
      Tag             =   "钢种"
      Top             =   660
      Width           =   930
   End
   Begin VB.TextBox txt_Trim_fl 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8685
      MaxLength       =   1
      TabIndex        =   23
      Top             =   660
      Width           =   420
   End
   Begin VB.TextBox txt_SizeKnd_s 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6135
      MaxLength       =   2
      TabIndex        =   22
      Top             =   660
      Width           =   345
   End
   Begin VB.TextBox txt_EndUse_s 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4305
      TabIndex        =   21
      Top             =   660
      Width           =   510
   End
   Begin VB.TextBox txt_prod_grd_s 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   20
      Top             =   9300
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt_Sale_dept 
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
      Left            =   9360
      MaxLength       =   3
      TabIndex        =   16
      Tag             =   "部门代码"
      Top             =   165
      Width           =   450
   End
   Begin VB.TextBox txt_Sale_dept_name 
      Enabled         =   0   'False
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
      Left            =   9825
      TabIndex        =   15
      Tag             =   "工 厂"
      Top             =   165
      Width           =   1785
   End
   Begin VB.TextBox text_cur_inv 
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
      Left            =   3600
      TabIndex        =   14
      Top             =   165
      Width           =   1200
   End
   Begin VB.TextBox text_cur_inv_code 
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
      Left            =   3210
      MaxLength       =   2
      TabIndex        =   13
      Top             =   165
      Width           =   375
   End
   Begin VB.TextBox txt_PLT 
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
      Left            =   6345
      MaxLength       =   2
      TabIndex        =   12
      Tag             =   "工 厂"
      Top             =   165
      Width           =   375
   End
   Begin VB.TextBox txt_PLT_NAME 
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
      Left            =   6735
      TabIndex        =   11
      Tag             =   "工 厂"
      Top             =   165
      Width           =   1050
   End
   Begin VB.TextBox Text_PROD_CD 
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
      Left            =   1245
      MaxLength       =   2
      TabIndex        =   10
      Tag             =   "产品"
      Top             =   165
      Width           =   375
   End
   Begin VB.TextBox txt_car_no 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13140
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "车辆号"
      Top             =   165
      Width           =   1395
   End
   Begin VB.TextBox txt_cust_cd_s 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2325
      TabIndex        =   7
      Top             =   9300
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt_stdspec_s 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1155
      TabIndex        =   3
      Top             =   660
      Width           =   1935
   End
   Begin VB.TextBox txt_stlgrd_s 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      TabIndex        =   2
      Top             =   660
      Width           =   1800
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7740
      Left            =   105
      TabIndex        =   0
      Top             =   1470
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   13653
      _Version        =   196609
      SplitterBarWidth=   3
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACB4020C.frx":0000
      Begin FPSpread.vaSpread ss2 
         Height          =   4380
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   3360
         Width           =   15015
         _Version        =   393216
         _ExtentX        =   26485
         _ExtentY        =   7726
         _StockProps     =   64
         ButtonDrawMode  =   4
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   24
         MaxRows         =   20
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACB4020C.frx":0052
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   3315
         Left            =   0
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   15015
         _Version        =   393216
         _ExtentX        =   26485
         _ExtentY        =   5847
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
         MaxCols         =   23
         MaxRows         =   21
         Protect         =   0   'False
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "ACB4020C.frx":0BB0
      End
   End
   Begin CSTextLibCtl.sidbEdit txt_len_min_s 
      Height          =   330
      Left            =   7500
      TabIndex        =   4
      Top             =   1005
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      NumIntDigits    =   0
      ShowZero        =   0   'False
      MaxValue        =   99999999
      MinValue        =   -99999999
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_wid_min_s 
      Height          =   330
      Left            =   4305
      TabIndex        =   5
      Top             =   1005
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      NumIntDigits    =   9
      ShowZero        =   0   'False
      MaxValue        =   99999999
      MinValue        =   -99999999
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_thk_min_s 
      Height          =   330
      Left            =   1155
      TabIndex        =   6
      Top             =   1005
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   225
      Top             =   165
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "产品"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   5145
      Tag             =   "移 送 工 厂"
      Top             =   165
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "目标库"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   2010
      Top             =   165
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "起始库"
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   11940
      Top             =   165
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "车辆号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   8160
      Tag             =   "移 送 工 厂"
      Top             =   165
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "部门代码"
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
   Begin CSTextLibCtl.sidbEdit txt_len_max_s 
      Height          =   330
      Left            =   8475
      TabIndex        =   17
      Top             =   1005
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      NumIntDigits    =   0
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_wid_max_s 
      Height          =   330
      Left            =   5295
      TabIndex        =   18
      Top             =   1005
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      NumIntDigits    =   9
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_thk_max_s 
      Height          =   330
      Left            =   2130
      TabIndex        =   19
      Top             =   1005
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel15 
      Height          =   315
      Left            =   5160
      Top             =   660
      Width           =   960
      _ExtentX        =   1693
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
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   10335
      Top             =   660
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   "等级"
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   195
      Top             =   660
      Width           =   960
      _ExtentX        =   1693
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
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Left            =   3315
      Top             =   660
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   "用途"
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
   Begin InDate.ULabel ULabel23 
      Height          =   315
      Left            =   7695
      Top             =   660
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   "切边"
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
   Begin Threed.SSCommand cmd_Clear 
      Height          =   330
      Left            =   13980
      TabIndex        =   27
      Top             =   645
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      _Version        =   196609
      ForeColor       =   16711680
      Caption         =   "条件删除"
   End
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   180
      Top             =   1005
      Width           =   960
      _ExtentX        =   1693
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   3315
      Top             =   1005
      Width           =   960
      _ExtentX        =   1693
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Left            =   6510
      Top             =   1005
      Width           =   960
      _ExtentX        =   1693
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   10305
      Top             =   1020
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   "堆放位置"
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
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      X1              =   135
      X2              =   15105
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   135
      X2              =   15105
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   135
      X2              =   15105
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   135
      X2              =   15105
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "ACB4020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name   Common
'-- Program Name      Insert Moving result
'-- Program ID        ACB4020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim S.H
'-- Coder             Kim S.H
'-- Date              2005.8.24
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

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iCurr_Row As Integer            'SS1 Current Row

Dim crxApplication As New CRAXDRT.Application

Public Report As CRAXDRT.Report

Dim crxDatabaseTable As CRAXDRT.DatabaseTable
Dim crxSubreport As CRAXDRT.Report
Dim CPProperties As CRAXDRT.ConnectionProperties


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
       
         Call Gp_Ms_Collection(Text_PROD_CD, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(text_cur_inv_code, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_sale_dept, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
   'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB4020C.P_SREFER1", Key:="P-R"
    sc1.Add Item:="ACB4020C.P_MODIFY2", Key:="P-M"
'    Sc1.Add Item:="ACB4020C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(Text_PROD_CD, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(txt_stdspec_s, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(txt_stlgrd_s, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(txt_prod_grd_s, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(txt_enduse_s, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(txt_cust_cd_s, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(txt_sizeKnd_s, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(txt_thk_min_s, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(txt_thk_max_s, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(txt_wid_min_s, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(txt_wid_max_s, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(txt_len_min_s, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(txt_len_max_s, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(txt_trim_fl, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
             Call Gp_Ms_Collection(txt_loc, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    'MASTER Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 23, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 24, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)

    'Spread_Collection
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc2.Add Item:="ACB4020C.P_MODIFY", Key:="P-M"
    Sc2.Add Item:="ACB4020C.P_SREFER2", Key:="P-R"
    Sc2.Add Item:="ACB4020C.P_MODIFY1", Key:="P-L"
    Sc2.Add Item:=pColumn2, Key:="pColumn"
    Sc2.Add Item:=nColumn2, Key:="nColumn"
    Sc2.Add Item:=aColumn2, Key:="aColumn"
    Sc2.Add Item:=mColumn2, Key:="mColumn"
    Sc2.Add Item:=iColumn2, Key:="iColumn"
    Sc2.Add Item:=lColumn2, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc2, Key:="Sc"
    Call Gp_Sp_ColHidden(ss1, 21, True)
    Call Gp_Sp_ColHidden(ss1, 22, True)
    Call Gp_Sp_ColHidden(ss1, 23, True)
    
    Me.KeyPreview = True

End Sub

Private Sub cmd_Clear_Click()
    Call Gp_Ms_Cls(Mc2("rControl"))
    cbo_prod_grd.Text = ""
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
    
    Call Form_Define
    Call AC_ComboAdd(M_CN1, cbo_prod_grd, "Q0034")
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(Sc2.Item("Spread"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(Sc2)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc2.Item("Spread"), "C-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    Text_PROD_CD.Text = "PP"
    text_cur_inv_code.Text = "00"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc2.Item("Spread"), "C-System.INI", Me.Name)
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
     
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
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set Sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()
    If Gf_Sp_Cls(sc1) Then
        If Gf_Sp_Cls(Proc_Sc("Sc")) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gp_Ms_Cls(Mc2("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            
            Call MenuTool_ReSet
            
            ss1.MaxRows = 0
            ss2.MaxRows = 0
            
            Text_PROD_CD.Text = "PP"
            text_cur_inv_code.Text = "00"
            cbo_prod_grd.Text = ""
            iCurr_Row = 0
        End If
    End If
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Call Gp_Sp_Excel(Me, Sc2.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

On Error Resume Next

    Dim iRow, iCol As Integer
    Dim dIst_Wgt, dEnd_Wgt As Double
      
    If Gf_Sp_ProceExist(Sc2.Item("Spread")) Then Exit Sub
            
    ss2.MaxRows = 0
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
           
'        Call ss1_Click(1, 1)
    
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet

    End If

End Sub

Public Sub Form_Pro()

    Dim sQuery      As String
    Dim sErrMessg   As String
    Dim TransNo     As String
    Dim iRow        As Integer
    Dim iCol        As Integer
    Dim iCount      As Integer
    Dim sTemp       As String
    Dim dTempInt    As Double
    Dim intLastRow  As Integer
    On Error GoTo Process_Exec_ERROR
    
    iCount = 0
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        ss1.Col = 0
        If ss1.Value = "Delete" Then
            iCount = iCount + 1
            Exit For
        End If
    Next iRow
    
    If iCount > 0 Then
        If Gf_Sp_Process(M_CN1, sc1, Mc1) Then
            Call Form_Ref
        End If
        Exit Sub
    End If
    
    If Trim(txt_car_no.Text) = "" Then
        Call Gp_MsgBoxDisplay("输入车辆号...")
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    M_CN1.BeginTrans
    
    ss1.Row = iCurr_Row
    ss1.Col = 1
    Call MoveTransNoEdit(ss1.Text, TransNo)
            
    iCount = 0
    For iRow = 1 To ss2.MaxRows
        ss2.Row = iRow
        ss2.Col = 1
        If ss2.Value = "1" Then
            iCount = iCount + 1
            intLastRow = iRow
            ss2.Col = 3
            ss2.Text = TransNo
            
            ss2.Col = 23
            ss2.Text = txt_car_no.Text
            
            sErrMessg = ""
            Call Sp_Process(iRow, sErrMessg)
                        
            'Error Check
            If Trim(sErrMessg) <> "" Then
                Call Gp_Sp_RowColor(ss2, iRow, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                M_CN1.RollbackTrans
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
             
        End If
    Next iRow

    For iRow = 1 To ss2.MaxRows
        ss2.Row = iRow
        ss2.Col = 0
        ss2.Text = ""
    Next iRow
    
    
    If iCount > 0 Then
        Call Sp_Process(intLastRow, sErrMessg, True)
    End If
    M_CN1.CommitTrans
    If iCount > 0 Then
    
        Call Form_Ref
'        sQuery = Gf_Sp_MakeQuery(ss1, Sc1.Item("P-O"), "O", Sc1.Item("pColumn"), iCurr_Row)
'        Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, Sc1.Item("Spread"), iCurr_Row)
'        Call ss1_Click(1, iCurr_Row)
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
    End If
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Process_Exec_ERROR:

    M_CN1.RollbackTrans
    Call Gp_MsgBoxDisplay(Error & sErrMessg)
    Screen.MousePointer = vbDefault
End Sub

Public Sub MoveTransNoEdit(MoveIspNo As String, TransNo As String)
    
    Dim SQL    As String
    Dim sDate  As String
    Dim AdoRs  As New ADODB.Recordset
    
    
    SQL = "SELECT  TO_CHAR(SYSDATE,'YYMMDD') FROM  DUAL " & vbCrLf
    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    If AdoRs.EOF = False Then
        sDate = AdoRs(0).Value & ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    SQL = " SELECT    SUBSTR(MAX(MV_LST_NO),1,11) || LPAD(TO_NUMBER(SUBSTR(MAX(MV_LST_NO),12,2)) + 1,2,'0')"
    SQL = SQL & "     FROM  CP_MOVE_SLT                                     " & vbCrLf
    SQL = SQL & "    WHERE  SUBSTR(MV_LST_NO,1,11) = '" & Left(MoveIspNo, 5) & sDate & "'" & vbCrLf
    
    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly

    If AdoRs.EOF Or AdoRs.BOF Then
        TransNo = Left(MoveIspNo, 5) & sDate & "01"
    Else
        TransNo = AdoRs.Fields(0) & ""
    End If
    
    If TransNo = "" Then
        TransNo = Left(MoveIspNo, 5) & sDate & "01"
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
End Sub

Private Sub Sp_Process(iRow As Integer, sErrMessg As String, Optional bLast As Boolean = False)

    Dim iCount      As Integer
    Dim iCol        As Integer
    Dim sTemp       As String
    Dim dTempInt    As Double

    Dim adoCmd As ADODB.Command

    On Error GoTo Process_Exec_ERROR

    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command

    Set adoCmd.ActiveConnection = M_CN1
    adoCmd.CommandType = adCmdStoredProc
    If bLast Then
       adoCmd.CommandText = Sc2.Item("P-L")
    Else
        adoCmd.CommandText = Sc2.Item("P-M")
    End If
    'Create Parameter (Input) iType + iColumn
    For iCount = 0 To Sc2.Item("iColumn").Count
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount

    adoCmd.Parameters(0).Value = "U"
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)

    Sc2.Item("Spread").Row = iRow

'Parameters Setting
    For iCol = 1 To Sc2.Item("iColumn").Count

        Sc2.Item("Spread").Col = Sc2.Item("iColumn").Item(iCol)
        Select Case Sc2.Item("Spread").CellType

            Case SS_CELL_TYPE_NUMBER
                If Trim(Sc2.Item("Spread").Text) = "" Then
                    adoCmd.Parameters(iCol).Value = 0
                Else
                    dTempInt = Sc2.Item("Spread").Text
                    adoCmd.Parameters(iCol).Value = Trim(Str(dTempInt))
                End If
                
            Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If Trim(Sc2.Item("Spread").Value) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(Str(Sc2.Item("Spread").Value))
                        End If
                        
            Case Else
                sTemp = Replace(Sc2.Item("Spread").Text, "'", "''")
                adoCmd.Parameters(iCol).Value = Trim(sTemp)

        End Select

    Next iCol

    adoCmd.Execute

    If adoCmd("Error") <> "0" Then
        sErrMessg = adoCmd("Messg")
    End If

    Set adoCmd = Nothing
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    sErrMessg = Error
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

Public Sub Spread_Del()

    Call Gp_Sp_Del(sc1)

End Sub

Public Sub Spread_Can()

    On Error Resume Next

    Dim sQuery As String
    Dim I As Integer
    Dim iRow, BR1, BR2 As Long

    With sc1
        
        .Item("Spread").ReDraw = False
        
        If .Item("Spread").MaxRows < 1 Or .Item("Spread").SelBlockRow < 1 Then
            Exit Sub
        End If
    
        BR1 = .Item("Spread").SelBlockRow
        BR2 = .Item("Spread").SelBlockRow2
        
        For iRow = .Item("Spread").SelBlockRow To BR2
            
            Select Case Trim(Gf_Sp_RcvData(.Item("Spread"), 0, iRow))
                
                Case "Delete"
                    Call Gp_Sp_SendData(.Item("Spread"), "", 0, iRow)
                    Call Gp_Sp_RowColor(.Item("Spread"), iRow)
                    
                    For I% = 1 To sc1!iColumn.Count
                        Call Gp_Sp_CellColor(.Item("Spread"), sc1!iColumn(I%), iRow, , &HC0FFFF)
                    Next I%
                Case Else
                    'sQuery = Gf_Sp_MakeQuery(.Item("Spread"), .Item("P-O"), "O", .Item("icolumn"), iRow)
                    'Call Gp_Sp_OneRowDisplay(Conn, sQuery, .Item("Spread"), iRow)
            End Select
            
            If iRow = BR2 Then
                Exit For
            End If

        Next iRow
        
        .Item("Spread").ReDraw = True
        
    End With
          
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


Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ss1.MaxRows > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim iRow        As Integer
    Dim dMoveWgt    As Double
    Dim dMoveCmdWgt As Double
    Dim dMoveSelWgt As Double
    Dim dMoveSelCnt As Double
    Dim sMoveNo     As String
    
    ss1.Row = iCurr_Row
        
    ss2.Row = Row
    ss2.Col = 12:   dMoveWgt = Val(ss2.Text & "")
    
    ss2.Col = 1
    If ss2.Text = "1" Then
        ss2.Col = 0:    ss2.Text = "Update"
        ss1.Col = 1
        ss2.Col = 2:    ss2.Text = ss1.Text
        ss2.Col = 24:   ss2.Text = sUserID
        
        ss1.Col = 10:   dMoveCmdWgt = Val(ss1.Value & "")
        ss1.Col = 11:   dMoveCmdWgt = dMoveCmdWgt - Val(ss1.Value & "")
        ss1.Col = 12:   dMoveSelWgt = Val(ss1.Value & "")
        ss1.Col = 13:   dMoveSelCnt = Val(ss1.Value & "")
        
        If dMoveSelWgt - dMoveCmdWgt >= 0 Then
            ss2.Col = 0:    ss2.Text = ""
            ss2.Col = 1:    ss2.Text = "0"
            ss2.Col = 3:    ss2.Text = ""
        Else
            ss1.Col = 12:  ss1.Text = dMoveSelWgt + dMoveWgt
            ss1.Col = 13:  ss1.Text = dMoveSelCnt + 1
            ss2.Col = 7:   ss2.Text = Format(Now, "YYYY-MM-DD")
            ss2.Col = 8:   ss2.Text = Format(Now, "HH:MM:SS")
        End If
        Call Gp_Sp_RowColor(ss2, Row, &HFF0000)
        Call Gp_Sp_BlockColor(ss2, 4, 4, Row, Row, &HFF0000, &HC0FFFF)
        Call Gp_Sp_BlockColor(ss2, 6, 8, Row, Row, &HFF0000, &HC0FFFF)
    Else
        ss2.Col = 0:     ss2.Text = ""
        ss2.Col = 2:     ss2.Text = ""
        ss2.Col = 7:     ss2.Text = ""
        ss2.Col = 8:     ss2.Text = ""
        ss2.Col = 24:    ss2.Text = ""
        Call Gp_Sp_RowColor(ss2, Row)
        'Call Gp_Sp_BlockColor(ss2, 4, 7, Row, Row, , &HC0FFFF)
        Call Gp_Sp_BlockColor(ss2, 4, 4, Row, Row, &HFF0000, &HC0FFFF)
        Call Gp_Sp_BlockColor(ss2, 6, 8, Row, Row, &HFF0000, &HC0FFFF)
        ss1.Col = 12:   ss1.Text = Val(ss1.Value & "") - dMoveWgt
        If Val(ss1.Text & "") < 0 Then ss1.Text = 0
        ss1.Col = 13:   ss1.Text = Val(ss1.Value & "") - 1
        If Val(ss1.Text & "") < 0 Then ss1.Text = 0
    End If
        
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Col > 1 Then Call Gp_Sp_Sort(Sc2.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    If Col <> 7 And Col <> 8 Then Exit Sub
    
    ss2.Row = Row
    ss2.Col = Col
    
    If Col = 7 Then
        ss2.Text = Format(Now, "YYYY-MM-DD")
    Else
        ss2.Text = Format(Now, "HH:MM:SS")
    End If
End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Row < 1 Or Col = 1 Then Exit Sub

    ss2.Row = Row
    ss2.Col = 1
    
    If ss2.Text <> "1" Then
        ss2.Text = "1"
    End If
End Sub

Private Sub ss2_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ss2.MaxRows > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub

    If Gf_Sp_ProceExist(Sc2.Item("Spread")) Then Exit Sub
        
    ss1.Row = Row
    iCurr_Row = Row
    
    ss1.Col = 14
    txt_cust_cd_s.Text = Trim(ss1.Text)
    
    
    ss1.Col = 16
    If Trim(txt_stdspec_s.Text) = "" Then txt_stdspec_s.Text = Trim(ss1.Text)
    
    ss1.Col = 18
    If Trim(cbo_prod_grd.Text) = "" Then cbo_prod_grd.Text = Trim(ss1.Text)
    
    ss1.Col = 19
    If Trim(txt_enduse_s.Text) = "" Then txt_enduse_s.Text = Left(Trim(ss1.Text), 3)
    
    ss1.Col = 21
    If Trim(txt_sizeKnd_s.Text) = "" Then txt_sizeKnd_s.Text = Trim(ss1.Text)
    
    ss1.Col = 22
    If Trim(txt_stlgrd_s.Text) = "" Then txt_stlgrd_s.Text = Trim(ss1.Text)
    
    ss1.Col = 23
    If Trim(txt_trim_fl.Text) = "" Then txt_trim_fl.Text = Trim(ss1.Text)
        
    ss1.Col = 2
    If Val(txt_thk_min_s.Text & "") = 0 Then txt_thk_min_s.Text = Val(ss1.Value & "")
    
    ss1.Col = 3
    If Val(txt_thk_max_s.Text & "") = 0 Then txt_thk_max_s.Text = Val(ss1.Value & "")
    
    ss1.Col = 4
    If Val(txt_wid_min_s.Text & "") = 0 Then txt_wid_min_s.Text = Val(ss1.Value & "")
    
    ss1.Col = 5
    If Val(txt_wid_max_s.Text & "") = 0 Then txt_wid_max_s.Text = Val(ss1.Value & "")
    
    ss1.Col = 6
    If Val(txt_len_min_s.Text & "") = 0 Then txt_len_min_s.Text = Val(ss1.Value & "")
    
    ss1.Col = 7
    If Val(txt_len_max_s.Text & "") = 0 Then txt_len_max_s.Text = Val(ss1.Value & "")
        
    Call Gf_Sp_Refer(M_CN1, Sc2, Mc2, Mc2("nControl"), Mc2("mControl"), False)
    ss2.OperationMode = OperationModeNormal
    
End Sub

Private Sub Text_PROD_CD_Change()
    Select Case Text_PROD_CD.Text
        Case "S", "s", "SL"
            Text_PROD_CD.Text = "SL"
        Case "P", "p", "PP"
            Text_PROD_CD.Text = "PP"
        Case "H", "h", "HC"
            Text_PROD_CD.Text = "HC"
        Case "", "**"
            Text_PROD_CD.Text = ""
        Case Else
            Text_PROD_CD.Text = ""
            Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
        End Select
    Call Gp_Ms_Cls(Mc2("rControl"))
    cbo_prod_grd.Clear
    Select Case Text_PROD_CD.Text
        Case "S", "s", "SL"
            Call Gp_Sp_ColHidden(ss1, 8, True)
            Call Gp_Sp_ColHidden(ss1, 16, True)
            Call Gp_Sp_ColHidden(ss1, 17, False)
            
            Call Gp_Sp_ColHidden(ss2, 13, True)
            Call Gp_Sp_ColHidden(ss2, 20, True)
            
            cbo_prod_grd.AddItem "0:合格"
            cbo_prod_grd.AddItem "1:表面不合格"
            cbo_prod_grd.AddItem "2:内部缺陷"
            cbo_prod_grd.AddItem "3:内外缺陷"
            cbo_prod_grd.AddItem "4:操作员变更"
            cbo_prod_grd.AddItem "5:长度不合格"
            
            txt_stlgrd_s.Visible = True
            txt_stdspec_s.Visible = False
        Case Else
            Call Gp_Sp_ColHidden(ss1, 8, False)
            Call Gp_Sp_ColHidden(ss1, 16, False)
            Call Gp_Sp_ColHidden(ss1, 17, True)
            
            Call Gp_Sp_ColHidden(ss2, 13, False)
            Call Gp_Sp_ColHidden(ss2, 20, False)
            Call AC_ComboAdd(M_CN1, cbo_prod_grd, "Q0034")
            
            txt_stlgrd_s.Visible = False
            txt_stdspec_s.Visible = True
    End Select
End Sub

Private Sub Text_PROD_CD_DblClick()

    Call Text_PROD_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=Text_PROD_CD
    
        DD.nameType = "2"
    
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If

End Sub

Private Sub text_cur_inv_code_DblClick()

    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_cur_inv_code_Change()
    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
        text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
    Else
      text_cur_inv.Text = ""
    End If
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=text_cur_inv_code
        DD.rControl.Add Item:=text_cur_inv
        

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If
End Sub

Private Sub txt_EndUse_s_Change()
'Gf_UsageNameFind(Conn As ADODB.Connection, Prod_Knd As String, Code As String)
End Sub

Private Sub txt_EndUse_s_DblClick()

    Call txt_EndUse_s_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_EndUse_s_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        If Text_PROD_CD.Text = "SL" Then
            DD.sKey = "S"
        Else
            DD.sKey = "P"
        End If
        
        DD.rControl.Add Item:=txt_enduse_s
        
        Call Gf_Usage_DD(M_CN1, KeyCode)
    End If
    
End Sub
Private Sub txt_plt_DblClick()

    Call txt_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_PLT_Change()
    If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
          txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_plt.Text, 2)
          Exit Sub
    Else
          txt_plt_name.Text = ""
    End If
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
        DD.rControl.Add Item:=txt_plt
       ' DD.rControl.Add Item:=txt_PLT_NAME
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
    
End Sub

Private Sub txt_Sale_dept_DblClick()

    Call txt_Sale_dept_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_Sale_dept_Change()
    If Len(Trim(txt_sale_dept.Text)) = txt_sale_dept.MaxLength Then
        txt_sale_dept_name.Text = Gf_ComnNameFind(M_CN1, "Z0002", txt_sale_dept.Text, 2)
    Else
      txt_sale_dept_name.Text = ""
    End If
End Sub

Private Sub txt_Sale_dept_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Z0002"

        DD.rControl.Add Item:=txt_sale_dept

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If
End Sub

Private Sub txt_car_no_DblClick()

    Call txt_car_no_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_car_no_KeyUp(KeyCode As Integer, Shift As Integer)

    If ULabel10.Caption <> "车辆号" Then Exit Sub
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_car_no
  '      DD.rControl.Add Item:=txt_fac_name

        DD.nameType = "2"

        Call Gf_CAR_NO_DD(M_CN1, KeyCode)

    End If

End Sub

Public Function Gf_CAR_NO_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String

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

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
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
    
    DD.DataDicType = "A"        'Apply Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "SELECT CAR_NO, CAR_KND,CAR_WGT_MAX,CAR_WGT_AVE,CAR_CMP_CD,Gf_Comnnamefind('H0002',CAR_CMP_CD) AS CAR_CMP_NAME FROM  HP_CAR_IMF "
    '    DD.sQuery = DD.sQuery + " WHERE "
        DD.sWhere = " WHERE CAR_NO like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        
    Else
    

    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
    End If
    
    DD.sWitch = ""
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing
    
End Function

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
'        .Buttons(8).Enabled = False                 'Row Delete
'        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Spread Copy
        .Buttons(12).Enabled = False                'Paste
    End With

End Sub

Private Sub cbo_prod_grd_Click()
    If Trim(cbo_prod_grd.Text) <> "" Then
        txt_prod_grd_s.Text = Left(cbo_prod_grd.Text, 1)
    Else
        txt_prod_grd_s.Text = ""
    End If
End Sub

Private Sub cbo_prod_grd_Change()
    If Trim(cbo_prod_grd.Text) <> "" Then
        txt_prod_grd_s.Text = Left(cbo_prod_grd.Text, 1)
    Else
        txt_prod_grd_s.Text = ""
    End If
End Sub

Private Sub txt_SizeKnd_s_Change()
    If Len(Trim(txt_sizeKnd_s.Text)) = txt_sizeKnd_s.MaxLength Then
        text_size_knd_name_in.Text = Gf_ComnNameFind(M_CN1, "B0043", txt_sizeKnd_s.Text, 2)
        Exit Sub
    Else
        text_size_knd_name_in.Text = ""
    End If
End Sub

Private Sub txt_SizeKnd_s_DblClick()

    Call txt_SizeKnd_s_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_SizeKnd_s_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"

        DD.rControl.Add Item:=txt_sizeKnd_s

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub txt_stdspec_s_DblClick()

    Call txt_stdspec_s_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_trim_fl_Change()
    If Len(Trim(txt_trim_fl.Text)) = txt_trim_fl.MaxLength Then
        txt_trim_name.Text = Gf_ComnNameFind(M_CN1, "B0021", txt_trim_fl.Text, 2)
        txt_trim_fl.Text = Trim(txt_trim_fl.Text)
        Exit Sub
    Else
        txt_trim_name.Text = ""
        txt_trim_fl.Text = ""
    End If

End Sub

Private Sub txt_trim_fl_DblClick()

    Call txt_trim_fl_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_trim_fl_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0021"

        DD.rControl.Add Item:=txt_trim_fl

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If

End Sub

Private Sub txt_stdspec_s_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec_s

        Call Gf_StdSPEC_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub txt_stlgrd_s_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stlgrd_s
        
        DD.nameType = "1"
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
    End If
End Sub

Private Function AC_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, sPrc As String, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim sQuery As String
    Dim intCount As Integer
    intCount = 1
    Dim AdoRs As ADODB.Recordset
    
    If Trim(sPrc) = "" Then
        AC_ComboAdd = False: Exit Function
    End If
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then AC_ComboAdd = False: Exit Function
    End If
    
    sQuery = "SELECT CD_NAME FROM ZP_CD Where CD_MANA_NO = '" + Trim(sPrc) + "'"

    If ClsChk Then
        Cbo.Clear
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
                If intCount = 6 Then intCount = 7
                Cbo.AddItem Trim(Str(intCount)) + ":" + AdoRs.Fields(0)
                intCount = intCount + 1
            End If
            AdoRs.MoveNext
            
        Wend
        AC_ComboAdd = True
    Else
        AC_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    AC_ComboAdd = False

End Function


