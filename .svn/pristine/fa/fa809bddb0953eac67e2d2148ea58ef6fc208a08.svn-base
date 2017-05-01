VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form CEC1010C 
   Caption         =   "订单查询/选定_CEC1010C"
   ClientHeight    =   8220
   ClientLeft      =   600
   ClientTop       =   2640
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chk_five_wgt 
      Caption         =   "5%欠量"
      Height          =   375
      Left            =   14280
      TabIndex        =   37
      Top             =   840
      Visible         =   0   'False
      Width           =   855
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
      Left            =   12125
      MaxLength       =   11
      TabIndex        =   36
      Tag             =   "产品"
      Top             =   485
      Width           =   1035
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
      Left            =   5370
      TabIndex        =   9
      Tag             =   "钢种"
      Top             =   485
      Width           =   1485
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
      Left            =   4895
      MaxLength       =   2
      TabIndex        =   8
      Tag             =   "钢种"
      Top             =   485
      Width           =   465
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
      Left            =   13470
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1725
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
      Left            =   1435
      MaxLength       =   1
      TabIndex        =   6
      Tag             =   "订单种类"
      Top             =   485
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
      Left            =   1920
      TabIndex        =   7
      Tag             =   "订单种类"
      Top             =   485
      Width           =   1395
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
      Left            =   8430
      MaxLength       =   30
      TabIndex        =   10
      Top             =   485
      Width           =   2130
   End
   Begin VB.TextBox txt_stlgrd_grp 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13335
      MaxLength       =   11
      TabIndex        =   35
      Tag             =   "钢种组"
      Top             =   15
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txt_ord_item 
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
      Left            =   2850
      MaxLength       =   2
      TabIndex        =   1
      Top             =   110
      Width           =   465
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
      Left            =   1435
      MaxLength       =   11
      TabIndex        =   0
      Tag             =   "产品"
      Top             =   110
      Width           =   1395
   End
   Begin VB.TextBox Txt_urgnt_fl_name 
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
      Left            =   12600
      MaxLength       =   80
      TabIndex        =   30
      Tag             =   "紧急订单"
      Top             =   30
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.TextBox Txt_urgnt_fl 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   12420
      MaxLength       =   2
      TabIndex        =   29
      Tag             =   "紧急订单"
      Top             =   30
      Visible         =   0   'False
      Width           =   165
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Left            =   13680
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   120
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "订单编制"
      BevelWidth      =   3
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
      Height          =   285
      Left            =   13080
      MaxLength       =   40
      TabIndex        =   28
      Tag             =   "产品"
      Top             =   30
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
      Height          =   285
      Left            =   12870
      MaxLength       =   2
      TabIndex        =   26
      Tag             =   "产品"
      Top             =   30
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox TxT_stdgrd 
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
      Left            =   12125
      MaxLength       =   11
      TabIndex        =   11
      Top             =   1260
      Width           =   1335
   End
   Begin InDate.UDate txt_del_fr 
      Height          =   315
      Left            =   4895
      TabIndex        =   2
      Top             =   105
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
   Begin CSTextLibCtl.sidbEdit txt_prod_thk_from 
      Height          =   315
      Left            =   1440
      TabIndex        =   12
      Top             =   870
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   150
      Top             =   870
      Width           =   1260
      _ExtentX        =   2223
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
      Left            =   3600
      Top             =   870
      Width           =   1260
      _ExtentX        =   2223
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
      Left            =   7140
      Top             =   870
      Width           =   1260
      _ExtentX        =   2223
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   3600
      Top             =   105
      Width           =   1260
      _ExtentX        =   2223
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   10830
      Top             =   1260
      Width           =   1260
      _ExtentX        =   2223
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
   Begin FPSpread.vaSpread ss1 
      Height          =   7440
      Left            =   60
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1740
      Width           =   15150
      _Version        =   393216
      _ExtentX        =   26723
      _ExtentY        =   13123
      _StockProps     =   64
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
      MaxCols         =   36
      MaxRows         =   2
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "CEC1010C.frx":0000
   End
   Begin CSTextLibCtl.sidbEdit txt_prod_thk_to 
      Height          =   315
      Left            =   2370
      TabIndex        =   13
      Top             =   870
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_prod_len_from 
      Height          =   315
      Left            =   8430
      TabIndex        =   16
      Top             =   870
      Width           =   1065
      _Version        =   262145
      _ExtentX        =   1879
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
   Begin CSTextLibCtl.sidbEdit txt_prod_len_to 
      Height          =   315
      Left            =   9495
      TabIndex        =   17
      Top             =   870
      Width           =   1065
      _Version        =   262145
      _ExtentX        =   1879
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
   Begin CSTextLibCtl.sidbEdit txt_prod_wid_from 
      Height          =   315
      Left            =   4895
      TabIndex        =   14
      Top             =   870
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
   Begin CSTextLibCtl.sidbEdit txt_prod_wid_to 
      Height          =   315
      Left            =   5880
      TabIndex        =   15
      Top             =   870
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
   Begin Threed.SSCommand SSCommand3 
      Height          =   360
      Left            =   16380
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   330
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
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
   End
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   150
      Top             =   105
      Width           =   1260
      _ExtentX        =   2223
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
   Begin Threed.SSCommand cmd_Thk_Modify 
      Height          =   270
      Left            =   16170
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   476
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "厚度变更"
   End
   Begin InDate.UDate txt_del_to 
      Height          =   315
      Left            =   6345
      TabIndex        =   3
      Top             =   105
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
   Begin InDate.ULabel ULabel20 
      Height          =   315
      Left            =   7140
      Top             =   480
      Width           =   1260
      _ExtentX        =   2223
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
   Begin InDate.UDate udt_release_date_fr 
      Height          =   315
      Left            =   9480
      TabIndex        =   4
      Top             =   105
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
   Begin InDate.UDate udt_release_date_to 
      Height          =   315
      Left            =   10920
      TabIndex        =   5
      Top             =   105
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
      Left            =   8190
      Top             =   105
      Width           =   1260
      _ExtentX        =   2223
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
   Begin InDate.ULabel ULabel19 
      Height          =   315
      Left            =   150
      Top             =   480
      Width           =   1260
      _ExtentX        =   2223
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
   Begin InDate.ULabel ULabel15 
      Height          =   315
      Left            =   3600
      Top             =   480
      Width           =   1260
      _ExtentX        =   2223
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
   Begin CSTextLibCtl.sidbEdit txt_slab_thk_from 
      Height          =   315
      Left            =   1440
      TabIndex        =   18
      Top             =   1260
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   150
      Top             =   1260
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "板坯厚度标准"
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   3600
      Top             =   1260
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "板坯宽度标准"
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
   Begin CSTextLibCtl.sidbEdit txt_slab_thk_to 
      Height          =   315
      Left            =   2370
      TabIndex        =   19
      Top             =   1260
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
   Begin CSTextLibCtl.sidbEdit txt_slab_wid_from 
      Height          =   315
      Left            =   4895
      TabIndex        =   20
      Top             =   1260
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
   Begin CSTextLibCtl.sidbEdit txt_slab_wid_to 
      Height          =   315
      Left            =   5880
      TabIndex        =   21
      Top             =   1260
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
   Begin CSTextLibCtl.sidbEdit txt_prod_wgt_from 
      Height          =   315
      Left            =   12125
      TabIndex        =   24
      Top             =   870
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
   Begin CSTextLibCtl.sidbEdit txt_prod_wgt_to 
      Height          =   315
      Left            =   13140
      TabIndex        =   25
      Top             =   870
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
      Left            =   10830
      Top             =   870
      Width           =   1260
      _ExtentX        =   2223
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   7140
      Top             =   1260
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "板坯长度标准"
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
   Begin CSTextLibCtl.sidbEdit txt_slab_len_from 
      Height          =   315
      Left            =   8430
      TabIndex        =   22
      Top             =   1260
      Width           =   1065
      _Version        =   262145
      _ExtentX        =   1879
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
      NumIntDigits    =   7
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_slab_len_to 
      Height          =   315
      Left            =   9495
      TabIndex        =   23
      Top             =   1260
      Width           =   1065
      _Version        =   262145
      _ExtentX        =   1879
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
      NumIntDigits    =   7
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   10830
      Top             =   485
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
      Height          =   315
      Left            =   12480
      TabIndex        =   38
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   1
      BackColor       =   12632319
      BackStyle       =   1
      Caption         =   "重点订单"
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   90
      X2              =   15150
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   15120
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "CEC1010C"
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
'-- Program Name
'-- Program ID        CEC1010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM.SUNG.HO
'-- Coder             KIM.SUNG.HO
'-- Date              2010.9.14
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

Dim iCount As Integer

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_prod_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_prod_cd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(Txt_urgnt_fl, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Txt_urgnt_fl_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udt_release_date_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udt_release_date_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_del_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_del_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_stlgrd_grp, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_stdgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_stdgrd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_prod_thk_from, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_prod_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_prod_len_from, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_prod_len_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_prod_wid_from, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_prod_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_prod_wgt_from, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_prod_wgt_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ord_item, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ord_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ord_knd_nm, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_size_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_size_knd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_slab_thk_from, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_slab_thk_to, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_slab_wid_from, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_slab_wid_to, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_slab_len_from, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_slab_len_to, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'          Call Gp_Ms_Collection(chk_five_wgt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(chk_key, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  BELOW EDIT ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", "n", "m", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss1, 2, "p", "n", "m", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
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
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    'Spread_Collection
    
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CEC1010C.P_SREFER", Key:="P-R"
    sc1.Add Item:="CEC1010C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="CEC1010C.P_SONEROW", Key:="P-O"
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  EDIT  End      ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    'Sum Column Count
    iSumCnt = 4
    
    'Sum Column Setting
    iSumCol.Add Item:=26
    iSumCol.Add Item:=27
    iSumCol.Add Item:=28
    iSumCol.Add Item:=29
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, 4, True)
    Call Gp_Sp_ColHidden(ss1, 15, True)
    Call Gp_Sp_ColHidden(ss1, 23, True)
    Call Gp_Sp_ColHidden(ss1, 24, True)
    
End Sub

Private Sub cmd_Thk_Modify_Click()

'    On Error GoTo Thk_Modify_ERROR
'
'    Dim strOrd_No As String
'    Dim strOrd_Item As String
'    Dim ret_Result_Err As String
'    Dim intRow As Integer
'    Dim intTrue As Integer
'    Dim intFalse As Integer
'
'    Exit Sub
'
'    If ss1.MaxRows < 1 Then Exit Sub
'    If Not Gf_MessConfirm("您确定修改要设计的板坯厚度吗？", "Q") Then Exit Sub
'    intRow = 1: intTrue = 0: intFalse = 0
'    Screen.MousePointer = vbHourglass
'    With ss1
'        .Col = 0
'        For intRow = 1 To .MaxRows
'            .Col = 0: .Row = intRow
'            If .Text = "Update" Then
'                .Col = 1: strOrd_No = Trim(.Text)
'                .Col = 2: strOrd_Item = Trim(.Text)
'                If Thk_Change(strOrd_No, strOrd_Item, False) Then
'                    intTrue = intTrue + 1
'                Else
'                    intFalse = intFalse + 1
'                End If
'                .Col = 0: .Text = ""
'            End If
'        Next intRow
'    End With
'
'    'Process Error Check
'    If intFalse > 0 Then
'        ret_Result_Err = "Error Mesg : " & "有" & Str(intFalse) & "记录板坯厚度修改失败!"
'        Call Gp_MsgBoxDisplay(ret_Result_Err)
'    Else
'        Call Form_Ref
'    End If
'
'    Screen.MousePointer = vbDefault
'    Exit Sub
'
'Thk_Modify_ERROR:
'
'    Screen.MousePointer = vbDefault
'    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)

End Sub

Private Sub Form_Activate()
    
    If Active_CForm <> "" Then
        Call txt_prod_cd_KeyUp(0, 0)
        Call Form_Ref
        Active_CForm = ""
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    If Mid(sAuthority, 4, 1) <> "1" Then
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
    Else
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
    End If
    
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

    'UPDATE AUTHORITY
    If Mid(sAuthority, 3, 1) <> "1" Then
        SSCommand1.Enabled = False
        SSCommand3.Enabled = False
    End If
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    If Mid(sAuthority, 4, 1) <> "1" Then
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
    Else
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = True
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
    End If
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "E-System.INI", Me.Name)
    
    txt_del_fr.Text = ""
    txt_del_to.Text = ""
    
    udt_release_date_fr.Text = Mid(DateAdd("M", -1, udt_release_date_to.Text), 1, 8) & "20"
    
    Active_CForm = ""
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

    Dim iRow As Integer
    
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    
    With ss1
        For iRow = 1 To .MaxRows - 1
            .Row = iRow
            .Col = 16
            If Trim(.Text) <> "定尺" Then
                .Col = 13:    .Lock = False
            Else
                .Col = 13:    .Lock = True
                Call Gp_Sp_BlockColor(ss1, 13, 13, iRow, iRow, BLACK, WHITE)
            End If
        Next iRow
    End With
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
    
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        If Mid(sAuthority, 4, 1) <> "1" Then
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        Else
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        End If
        
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        
        txt_del_fr.Text = ""
        txt_del_to.Text = ""
        txt_prod_cd_name.Text = ""
        Txt_urgnt_fl_name.Text = ""
        udt_release_date_fr.Text = Mid(DateAdd("M", -1, udt_release_date_to.Text), 1, 8) & "20"
    End If

End Sub

Public Sub Form_Ref()

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call Sp_Total
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss1.OperationMode = OperationModeNormal
        If Mid(sAuthority, 4, 1) <> "1" Then   'DELETE
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        Else
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        End If
        
        '重点订单红色标记 2013-11-16  by  CaoLei
         Call SS1_CHANGE_COLOR
               
    End If
    
End Sub


Private Sub SS1_CHANGE_COLOR()

    With ss1

        If .MaxRows <= 0 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .Row = iCount

             '重点订单红色标记 2013-11-16  by  CaoLei
            ss1.Row = .Row:          ss1.Col = 34
            If ss1.Text = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, 1, 2, .Row, .Row, &HFF&)
                 Call Gp_Sp_BlockColor(ss1, 34, 34, .Row, .Row, &HFF&)
            End If

        Next iCount

    End With

End Sub


Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call Sp_Total
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        If Mid(sAuthority, 4, 1) <> "1" Then
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        Else
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        End If
    End If
    
End Sub

Public Sub Form_Ins()
    
    'Call Gp_Sp_Ins(Proc_Sc("Sc"))
    'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)

End Sub

Public Sub Spread_Cpy()

    'Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    'Call Gp_Sp_Paste(Proc_Sc("Sc"))
    'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
    
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

    Dim SMESG As String
    Dim iRow As Integer
    
    If lBlkrow1 = 0 Or lBlkrow2 = 0 Then Exit Sub
    
    For iRow = lBlkrow1 To lBlkrow2
    
       ss1.Row = iRow
       ss1.Col = 28
       
       If ss1.Value <> 0 Then
           SMESG = "设计已经完成不能删除！"
           Call Gp_MsgBoxDisplay(SMESG)
           Exit Sub
       End If
       
    Next iRow
    
    Call Gp_Sp_Del(Proc_Sc("SC"))
    
    For iRow = 1 To ss1.MaxRows
    
        ss1.Row = iRow
        ss1.Col = 1
        
        If Trim(ss1.Text) = "合   计" Then
            ss1.Col = 0
            ss1.Text = ""
        End If
        
    Next iRow
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    Dim i As Integer
    Dim intRow As Integer
    Dim intRow2 As Integer
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
    
'    If BlockCol = -1 And BlockCol2 = -1 Then Exit Sub
'
'    If BlockCol > 2 Or BlockCol2 > 2 Then Exit Sub
'    intRow = IIf(BlockRow < BlockRow2, BlockRow, BlockRow2)
'    intRow2 = IIf(BlockRow > BlockRow2, BlockRow, BlockRow2)
'    intRow2 = IIf(intRow2 >= ss1.MaxRows, ss1.MaxRows - 1, intRow2)
    
'    With ss1
'        For I = intRow To intRow2
'            .Row = I: .Col = 10
'            If .Text <> "SL" Then
'                .Col = 30
'                If .Text <> "C3" Then
'                    .Col = 0: .Text = "Update"
'                End If
'            End If
'        Next
'    End With

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
'    If Row < 1 Then Exit Sub
'    If ss1.MaxRows < 1 Or Col > 0 Or Row = ss1.MaxRows Then Exit Sub
    
'    ss1.Row = Row
'    ss1.Col = 10
'    If Trim(ss1.Text) <> "SL" Then
'        ss1.Col = 30
'        If Trim(ss1.Text) <> "C1" Then
'            ss1.Col = 0
'            If ss1.Text <> "Update" Then
'                ss1.Text = "Update"
'            Else
'                ss1.Text = ""
'            End If
'        End If
'    Else
'        Gp_MsgBoxDisplay ("提示 : 不能选择修改板坯产品的厚度！")
'    End If
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    Dim dValue As Double
    Dim sString As String
    
    If ChangeMade = False Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 30
    
    If ss1.Text = "C1" Then Exit Sub
    
    'ORD_HCR_FL  SLAB_WID   DESIGN_CNF_WGT
    If Col = 19 Or Col = 24 Or Col = 27 Then
'        ss1.Col = 7
'        If ss1.Text = "SL" Then
'            ss1.Col = 0
'            ss1.Text = "Update"
'            Call Spread_Can
'            Exit Sub
'        End If
        
        ss1.Col = 28
        If ss1.Value > 0 Then
            ss1.Col = 0
            ss1.Text = "Update"
            Call Spread_Can
            Exit Sub
        End If
    End If
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        'Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

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

Private Sub SSCommand1_Click()

    Call Gp_Process_Exec("1")
    
End Sub

Private Sub SSCommand3_Click()

    Call Gp_Process_Exec("Y")
    
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

Private Sub txt_prod_cd_DblClick()

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

        If Len(Trim(txt_stdgrd.Text)) >= 10 Then
            txt_stdgrd_name.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stdgrd.Text))
        Else
            txt_stdgrd_name.Text = ""
        End If
            
    End If
    
End Sub

Private Sub txt_stlgrd_grp_DblClick()

    Call txt_stlgrd_grp_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_stlgrd_grp_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Q0048"
        DD.rControl.Add Item:=txt_stlgrd_grp
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If

End Sub

Private Sub Txt_urgnt_fl_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "B0022"
        
        DD.rControl.Add Item:=Txt_urgnt_fl
        DD.rControl.Add Item:=Txt_urgnt_fl_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else

        If Len(Trim(Txt_urgnt_fl.Text)) = Txt_urgnt_fl.MaxLength Then
            Txt_urgnt_fl_name.Text = Gf_ComnNameFind(M_CN1, "B0022", Trim(Txt_urgnt_fl.Text), 2)
        Else
            Txt_urgnt_fl_name.Text = ""
        End If
    
    End If
    
End Sub

Public Sub Gp_Process_Exec(P_MODE As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
        
    If txt_prod_thk_to.Value = 0 Then
       txt_prod_thk_to.Value = 999.99
    End If
        
    If txt_prod_wid_to.Value = 0 Then
       txt_prod_wid_to.Value = 9999
    End If
        
    If txt_prod_len_to.Value = 0 Then
       txt_prod_len_to.Value = 9999999
    End If
    
    If txt_prod_wgt_to.Value = 0 Then
       txt_prod_wgt_to.Value = 999
    End If
    
    If txt_slab_thk_to.Value = 0 Then
       txt_slab_thk_to.Value = 9999.99
    End If
        
    If txt_slab_wid_to.Value = 0 Then
       txt_slab_wid_to.Value = 9999
    End If
        
    If txt_slab_len_to.Value = 0 Then
       txt_slab_len_to.Value = 9999999
    End If
        
        
'      Call Gp_MsgBoxDisplay(chk_five_wgt.Value)
    sQuery = "{call CEC1020P ('C3','" & Trim(txt_ord_no.Text) + "','" + Trim(txt_ord_item.Text) + "','" & _
                                        Trim(txt_prod_cd.Text) + "','" + Trim(TXT_CUST_CD.Text) + "','" + Trim(txt_stlgrd_grp.Text) + "','" & _
                                        Trim(txt_stdgrd.Text) + "','" + Trim(txt_stdspec.Text) + "','" & _
                                        Trim(txt_del_fr.RawData) + "','" + Trim(txt_del_to.RawData) + "'," & _
                                        txt_prod_thk_from.Value & "," & txt_prod_thk_to.Value & "," & txt_prod_wid_from.Value & "," & _
                                        txt_prod_wid_to.Value & "," & txt_prod_len_from.Value & "," & txt_prod_len_to.Value & "," & _
                                        txt_prod_wgt_from.Value & "," & txt_prod_wgt_to.Value & ",'" & _
                                        Trim(udt_release_date_fr.RawData) + "','" + Trim(udt_release_date_to.RawData) + "','" & _
                                        Trim(txt_ord_knd.Text) + "','" + Trim(txt_size_knd.Text) + "'," & _
                                        txt_slab_thk_from.Value & "," & txt_slab_thk_to.Value & "," & txt_slab_wid_from.Value & "," & txt_slab_wid_to.Value & "," & _
                                        txt_slab_len_from.Value & "," & txt_slab_len_to.Value & ", '" & _
                                        sUserID & "'," & chk_key.Value & ",?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Form_Ref
        
    '重点订单红色标记 2013-11-16  by  CaoLei
        Call SS1_CHANGE_COLOR
        
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Public Sub Sp_Total()
    
    Dim j As Integer
    Dim iBas As Integer
    Dim iCot As Integer
    Dim iRow As Integer
    
    Dim sCol_a As String
    Dim sCol_b As String
    
    With ss1
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1
        
        Call Gp_Sp_BlockLock(ss1, 1, .MaxCols, .MaxRows, .MaxRows, True)
        Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HE6E6FF)
        
        For j = 1 To iSumCnt
            .Col = j
            If .ColHidden = False Then
                .Text = "合   计"
                j = iSumCnt
            End If
        Next j
        
        For j = 1 To iSumCnt
            .Col = iSumCol(j)
            
            If iSumCol(j) <= 26 Then
                sCol_a = Chr(iSumCol(j) + 64)
                .Formula = "sum(" + sCol_a + "1:" + sCol_a & .MaxRows - 1 & ")"
            Else
                iCot = Int(((iSumCol(j) - 1) / 26))
                iBas = 26 * iCot
                sCol_a = Chr((iSumCol(j) - iBas) + 64)
                sCol_b = Chr(iCot + 64)
                .Formula = "sum(" + sCol_b + sCol_a + "1:" + sCol_b + sCol_a & .MaxRows - 1 & ")"
            End If
        Next j
        
        For iRow = 1 To .MaxRows - 1
            .Row = iRow
            .Col = 16
            If Trim(.Text) <> "定尺" Then
                .Col = 13:    .Lock = False
            Else
                .Col = 13:    .Lock = True
                Call Gp_Sp_BlockColor(ss1, 13, 13, iRow, iRow, BLACK, WHITE)
            End If
        Next iRow
        
    End With
        
End Sub

Private Sub ss1_EditChange(ByVal Col As Long, ByVal Row As Long)

    Dim d_Thk       As Double
    Dim d_Wth       As Double
    Dim d_Lth       As Double
    Dim d_Wgt       As Double
    Dim sEnduse     As String
    Dim sStdspec    As String
    Dim sStlgrd     As String
    
    If Col <> 14 Then Exit Sub

    With ss1
        .Row = Row
        .Col = 5
        sEnduse = Trim(.Text)
        .Col = 7
        sStlgrd = Trim(.Text)
        .Col = 9
        sStdspec = Trim(.Text)
        .Col = 11
        d_Thk = Val(.Value & "")
        .Col = 12
        d_Wth = Val(.Value & "")
        .Col = 13
        d_Lth = Val(.Value & "")
        
        .Col = 10
        If .Text = "PP" Then
            .Col = 14
            .Value = Cal_Plate_Wgt("PP", "WGT", sEnduse, sStdspec, d_Thk, d_Wth, d_Lth, d_Wgt)
        Else
            .Col = 14
            .Value = Cal_Plate_Wgt("SL", "WGT", "", sStlgrd, d_Thk, d_Wth, d_Lth, d_Wgt)
        End If
    End With

End Sub

Private Function Cal_Plate_Wgt(sProd As String, sMode As String, sEnduse As String, sStdspec As String, dThk As Double, dWid As Double, dLen As Double, dWgt As Double) As Double
    
    Dim sQuery As String
    Dim RS  As New ADODB.Recordset
    
    Cal_Plate_Wgt = 0
    
    If sProd = "PP" Then
        sQuery = "           SELECT  Gf_Cal_Plate_Wgt('" & sMode & "'" & vbCrLf
        sQuery = sQuery & "                          ,'" & sEnduse & "'" & vbCrLf
        sQuery = sQuery & "                          ,'" & sStdspec & "'" & vbCrLf
        sQuery = sQuery & "                          ," & dThk & vbCrLf
        sQuery = sQuery & "                          ," & dWid & vbCrLf
        sQuery = sQuery & "                          ," & dLen & vbCrLf
        sQuery = sQuery & "                          ," & dWgt & ")" & vbCrLf
        sQuery = sQuery & "    FROM  DUAL " & vbCrLf
    Else
        sQuery = "           SELECT  GF_JP_WGT('" & sMode & "'" & vbCrLf
        sQuery = sQuery & "                          ,'" & sStdspec & "'" & vbCrLf
        sQuery = sQuery & "                          ," & dThk & vbCrLf
        sQuery = sQuery & "                          ," & dWid & vbCrLf
        sQuery = sQuery & "                          ," & dLen & vbCrLf
        sQuery = sQuery & "                          ," & dWgt & ")" & vbCrLf
        sQuery = sQuery & "    FROM  DUAL " & vbCrLf
    End If
    
    RS.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        Cal_Plate_Wgt = Val(RS(0).Value & "")
    End If
    
    RS.Close
    Set RS = Nothing
     
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
