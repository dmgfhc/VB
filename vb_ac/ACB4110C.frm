VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form ACB4110C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "精整作业对象查询_ACB4110C"
   ClientHeight    =   9225
   ClientLeft      =   285
   ClientTop       =   1815
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_lot_no 
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
      Left            =   11190
      TabIndex        =   37
      Tag             =   "轧批号"
      Top             =   1170
      Width           =   1875
   End
   Begin VB.TextBox txt_mat_no 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1125
      MaxLength       =   15
      TabIndex        =   5
      Tag             =   "物料号"
      Top             =   450
      Width           =   1635
   End
   Begin VB.TextBox txt_next_plan_htm 
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
      Height          =   310
      Left            =   11190
      MaxLength       =   1
      TabIndex        =   32
      Top             =   810
      Width           =   420
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   1125
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "生产厂"
      Top             =   105
      Width           =   420
   End
   Begin VB.ComboBox cbo_ust_fl 
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
      Left            =   4395
      TabIndex        =   15
      Top             =   1170
      Width           =   750
   End
   Begin VB.TextBox txt_trim_fl 
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
      Left            =   1125
      MaxLength       =   1
      TabIndex        =   13
      Tag             =   "钢种"
      Top             =   1175
      Width           =   420
   End
   Begin VB.TextBox txt_trim_nm 
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
      Height          =   310
      Left            =   1545
      TabIndex        =   14
      Tag             =   "钢种"
      Top             =   1175
      Width           =   1080
   End
   Begin VB.TextBox txt_surf_grd_nm 
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
      Height          =   310
      Left            =   4725
      TabIndex        =   11
      Tag             =   "钢种"
      Top             =   815
      Width           =   1590
   End
   Begin VB.TextBox txt_surf_grd 
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
      Left            =   4395
      MaxLength       =   1
      TabIndex        =   10
      Top             =   810
      Width           =   345
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
      Left            =   7680
      MaxLength       =   2
      TabIndex        =   16
      Tag             =   "钢种"
      Top             =   1185
      Width           =   345
   End
   Begin VB.TextBox txt_size_knd_nm 
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
      Height          =   310
      Left            =   8025
      TabIndex        =   17
      Tag             =   "钢种"
      Top             =   1185
      Width           =   1740
   End
   Begin VB.TextBox txt_cur_inv_nm 
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
      Left            =   4875
      TabIndex        =   2
      Top             =   105
      Width           =   1590
   End
   Begin VB.TextBox txt_cur_inv 
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
      Left            =   4395
      MaxLength       =   2
      TabIndex        =   1
      Top             =   105
      Width           =   450
   End
   Begin VB.TextBox txt_cust_cd 
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
      Left            =   7680
      MaxLength       =   6
      TabIndex        =   8
      Top             =   465
      Width           =   1260
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
      Left            =   5715
      TabIndex        =   7
      Top             =   465
      Width           =   750
   End
   Begin VB.TextBox txt_rec_sts 
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
      Left            =   7680
      MaxLength       =   1
      TabIndex        =   3
      Tag             =   "CD_MANA_NO"
      Text            =   "2"
      Top             =   105
      Width           =   315
   End
   Begin VB.TextBox txt_proc_cd 
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
      Left            =   9165
      MaxLength       =   3
      TabIndex        =   4
      Tag             =   "CD_MANA_NO"
      Top             =   105
      Visible         =   0   'False
      Width           =   600
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
      Height          =   315
      Left            =   7665
      MaxLength       =   20
      TabIndex        =   12
      Tag             =   "钢种(标准号)"
      Top             =   825
      Width           =   2100
   End
   Begin VB.TextBox txt_loc 
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
      Left            =   1125
      MaxLength       =   7
      TabIndex        =   9
      Tag             =   "CD_MANA_NO"
      Top             =   810
      Width           =   1260
   End
   Begin VB.TextBox txt_ord_no 
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
      Left            =   4395
      MaxLength       =   11
      TabIndex        =   6
      Tag             =   "CD_MANA_NO"
      Top             =   465
      Width           =   1320
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   3390
      Top             =   465
      Width           =   990
      _ExtentX        =   1746
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   6660
      Top             =   105
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "信息状态"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   6660
      Top             =   825
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "标准号"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   8160
      Top             =   105
      Visible         =   0   'False
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "物料状态"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   120
      Top             =   810
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "堆放位置"
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
   Begin CSTextLibCtl.sidbEdit sdb_thk_fr 
      Height          =   315
      Left            =   1125
      TabIndex        =   18
      Top             =   1545
      Width           =   1020
      _Version        =   262145
      _ExtentX        =   1799
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   120
      Top             =   1545
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "厚度"
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   3390
      Top             =   1545
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "宽度"
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   6660
      Top             =   1545
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "长度"
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
   Begin CSTextLibCtl.sidbEdit sdb_wid_fr 
      Height          =   315
      Left            =   4395
      TabIndex        =   20
      Top             =   1545
      Width           =   1020
      _Version        =   262145
      _ExtentX        =   1799
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumIntDigits    =   4
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_len_fr 
      Height          =   315
      Left            =   7680
      TabIndex        =   22
      Top             =   1545
      Width           =   1020
      _Version        =   262145
      _ExtentX        =   1799
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumIntDigits    =   7
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   6660
      Top             =   465
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "客户"
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
   Begin CSTextLibCtl.sidbEdit sdb_thk_to 
      Height          =   315
      Left            =   2145
      TabIndex        =   19
      Top             =   1545
      Width           =   1020
      _Version        =   262145
      _ExtentX        =   1799
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
   Begin CSTextLibCtl.sidbEdit sdb_wid_to 
      Height          =   315
      Left            =   5430
      TabIndex        =   21
      Top             =   1545
      Width           =   1020
      _Version        =   262145
      _ExtentX        =   1799
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumIntDigits    =   4
      MaxValue        =   9999
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_len_to 
      Height          =   315
      Left            =   8700
      TabIndex        =   23
      Top             =   1545
      Width           =   1020
      _Version        =   262145
      _ExtentX        =   1799
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumIntDigits    =   7
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Left            =   3390
      Top             =   105
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "堆放仓库"
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
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   6660
      Top             =   1185
      Width           =   990
      _ExtentX        =   1746
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
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   3390
      Top             =   810
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "外观等级"
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
      Left            =   120
      Top             =   1170
      Width           =   990
      _ExtentX        =   1746
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
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   3390
      Top             =   1170
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "探伤是否"
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   120
      Top             =   105
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "生产厂"
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
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Left            =   9960
      Top             =   805
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Caption         =   "热处理对象"
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
   Begin Threed.SSCheck chk_htm_shot_blast 
      Height          =   285
      Left            =   14010
      TabIndex        =   30
      Top             =   135
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
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
      Caption         =   "抛丸指示"
   End
   Begin Threed.SSCheck chk_grid_fl 
      Height          =   285
      Left            =   9960
      TabIndex        =   24
      Top             =   135
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
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
      Caption         =   "修磨指示"
   End
   Begin Threed.SSCheck chk_grid_rslt 
      Height          =   285
      Left            =   9960
      TabIndex        =   25
      Top             =   495
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
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
      Caption         =   "修磨实绩"
   End
   Begin Threed.SSCheck chk_cl_fl 
      Height          =   285
      Left            =   11220
      TabIndex        =   26
      Top             =   135
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
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
      Caption         =   "冷矫直指示"
   End
   Begin Threed.SSCheck chk_cl_rslt 
      Height          =   285
      Left            =   11220
      TabIndex        =   27
      Top             =   495
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
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
      Caption         =   "冷矫直实绩"
   End
   Begin Threed.SSCheck chk_gas_fl 
      Height          =   285
      Left            =   12600
      TabIndex        =   28
      Top             =   135
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
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
      Caption         =   "气切割指示"
   End
   Begin Threed.SSCheck chk_gas_rslt 
      Height          =   285
      Left            =   12600
      TabIndex        =   29
      Top             =   495
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
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
      Caption         =   "气切割实绩"
   End
   Begin Threed.SSCheck chk_htm_shot_blast_rlt 
      Height          =   285
      Left            =   14010
      TabIndex        =   31
      Top             =   480
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
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
      Caption         =   "抛丸实绩"
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   7215
      Left            =   120
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2040
      Width           =   15255
      _Version        =   393216
      _ExtentX        =   26908
      _ExtentY        =   12726
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
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
      MaxCols         =   55
      MaxRows         =   2
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "ACB4110C.frx":0000
   End
   Begin Threed.SSOption opt_all 
      Height          =   255
      Left            =   9960
      TabIndex        =   33
      Top             =   1605
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
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
   Begin Threed.SSOption opt_ord 
      Height          =   255
      Left            =   10800
      TabIndex        =   34
      Top             =   1605
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
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
      Caption         =   "订单"
   End
   Begin Threed.SSOption opt_nonord 
      Height          =   255
      Left            =   11610
      TabIndex        =   35
      Top             =   1605
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
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
      Caption         =   "余材"
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   120
      Top             =   465
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "物料号"
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
      Left            =   9960
      Top             =   1170
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Caption         =   "轧批号"
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
   Begin Threed.SSCommand SSCommand2 
      Height          =   315
      Left            =   12480
      TabIndex        =   38
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   255
      Caption         =   "离线切割"
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   15290
      Y1              =   1905
      Y2              =   1905
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00404040&
      X1              =   120
      X2              =   15290
      Y1              =   1935
      Y2              =   1935
   End
End
Attribute VB_Name = "ACB4110C"
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
'-- Program ID        ACB4110C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2008.01.18
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
Public AIMNO As String

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iCount As Integer
Const SS1_URGNT_FL = 47
Const SS1_MAT_NO = 2 '物料号
Const SS1_PLT = 53 '工厂号
Const SS1_PROC_CD = 3 '物料状态
Const SS1_STL_GRD = 6 '钢种
Const SS1_STDSPEC_ORG_KND = 7 '执行标准
Const SS1_PAINTNUM = 8 '标识次数
Const SS1_STD_SPEC = 5 '标准
Const SS1_PAINT_FL = 54 '冲印指示
Const SS1_VESSEL_NO = 49 '加喷内容
Const SS1_COLOR_STROKE = 52 '色标及备注
Const SS1_THICK = 9 '厚度
Const SS1_WID = 10 '宽度
Const SS1_LEN = 11 '长度
Const SS1_ORD_THK = 13 '订单厚度
Const SS1_ORD_WID = 14 '订单宽度
Const SS1_ORD_LEN = 15 '订单长度
Const SS1_CUST_CD = 25 '客户
Const SS1_THK_TOL_MIN = 38 '厚度公差最小值
Const SS1_THK_TOL_MAX = 39 '厚度公差最大值
Const SS1_WID_TOL_MIN = 40 '宽度公差最小值
Const SS1_WID_TOL_MAX = 41 '宽度公差最大值
Const SS1_LEN_TOL_MIN = 42 '长度公差最小值
Const SS1_LEN_TOL_MAX = 43 '长度公差最大值
Const SS1_SIZE_KIND = 16 '定尺区分
Const SS1_TRIM_FL = 17 '切边
Const SS1_UST_FL = 18 '探伤是否
Const SS1_CL_FL = 30 '冷矫直指示/实绩
Const SS1_GAS_FL = 31 '气切割指示/实绩
Const SS1_HTM_SHOT_BLAST = 32 '抛丸指示/实绩
Const SS1_HTM_METH = 33 '热处理指示/实绩
Const SS1_CUR_INV = 34 '堆放仓库
Const SS1_LOC = 35 '堆放位置
Const SS1_DEFECT = 37 '缺陷
Const SS1_ORD_REMARK = 36 '订单备注
Const SS1_SIDEMARK = 50 '侧喷加喷
Const SS1_SEALMEMO = 51 '冲印加喷
Const SS1_CE = 48 '认证标识
Const SS1_CUT_PLAN = 55 '切割计划

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
                   Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '1
               Call Gp_Ms_Collection(txt_CUR_INV, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '2
            Call Gp_Ms_Collection(txt_cur_inv_nm, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '3
               Call Gp_Ms_Collection(txt_REC_STS, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '4
               Call Gp_Ms_Collection(txt_PROC_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '5
                Call Gp_Ms_Collection(txt_mat_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '6
                   Call Gp_Ms_Collection(txt_Loc, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '7
                Call Gp_Ms_Collection(TXT_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '8
              Call Gp_Ms_Collection(CBO_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '9
               Call Gp_Ms_Collection(txt_cust_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '10
              Call Gp_Ms_Collection(txt_SURF_GRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '11
           Call Gp_Ms_Collection(txt_surf_grd_nm, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '12
               Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '13
               Call Gp_Ms_Collection(txt_TRIM_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '14
               Call Gp_Ms_Collection(txt_trim_nm, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '15
                Call Gp_Ms_Collection(cbo_ust_fl, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '16
              Call Gp_Ms_Collection(txt_size_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '17
           Call Gp_Ms_Collection(txt_size_knd_nm, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '18
                Call Gp_Ms_Collection(sdb_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '19
                Call Gp_Ms_Collection(SDB_THK_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '20
                Call Gp_Ms_Collection(sdb_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '21
                Call Gp_Ms_Collection(SDB_WID_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '22
                Call Gp_Ms_Collection(sdb_len_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '23
                Call Gp_Ms_Collection(SDB_LEN_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '24
               Call Gp_Ms_Collection(chk_grid_fl, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '25
             Call Gp_Ms_Collection(chk_grid_rslt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '26
                 Call Gp_Ms_Collection(chk_cl_fl, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '27
               Call Gp_Ms_Collection(chk_cl_rslt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '28
                Call Gp_Ms_Collection(chk_gas_fl, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '29
              Call Gp_Ms_Collection(chk_gas_rslt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '30
        Call Gp_Ms_Collection(chk_htm_shot_blast, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '31
    Call Gp_Ms_Collection(chk_htm_shot_blast_rlt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '32
         Call Gp_Ms_Collection(txt_next_plan_htm, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '33
                   Call Gp_Ms_Collection(Opt_all, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '34
                   Call Gp_Ms_Collection(opt_ord, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '35
                Call Gp_Ms_Collection(opt_nonord, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '36
                Call Gp_Ms_Collection(TXT_LOT_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '37
            
    'MASTER Collection
    Mc1.Add Item:="ACB4110C.P_SREFER", Key:="P-R"
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
    
    'Sum Column Count
    iSumCnt = 1
    
    'Sum Column Setting
    iSumCol.Add Item:=7
    
    cbo_ust_fl.AddItem " "
    cbo_ust_fl.AddItem "Y"
    cbo_ust_fl.AddItem "N"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub opt_nonord_Click(Value As Integer)

    If opt_nonord.Value Then
        opt_nonord.ForeColor = &HFF&
        Opt_all.ForeColor = &H0&
        opt_ord.ForeColor = &H0&
    End If
    
End Sub

Private Sub opt_ord_Click(Value As Integer)

    If opt_ord.Value Then
        opt_ord.ForeColor = &HFF&
        Opt_all.ForeColor = &H0&
        opt_nonord.ForeColor = &H0&
    End If

End Sub

Private Sub opt_all_Click(Value As Integer)

    If Opt_all.Value Then
        Opt_all.ForeColor = &HFF&
        opt_ord.ForeColor = &H0&
        opt_nonord.ForeColor = &H0&
    End If

End Sub

Private Sub sdb_len_fr_Change()
    If sdb_len_fr.Value > 0 And SDB_LEN_TO.Value < sdb_len_fr.Value Then
        SDB_LEN_TO.Value = sdb_len_fr.Value
    End If
End Sub

Private Sub sdb_thk_fr_Change()
    If sdb_thk_fr.Value > 0 And SDB_THK_TO.Value < sdb_thk_fr.Value Then
        SDB_THK_TO.Value = sdb_thk_fr.Value
    End If
End Sub

Private Sub sdb_wid_fr_Change()
    If sdb_wid_fr.Value > 0 And SDB_WID_TO.Value < sdb_wid_fr.Value Then
        SDB_WID_TO.Value = sdb_wid_fr.Value
    End If
End Sub

Private Sub SSCommand2_Click()

Call ExcelPrn

End Sub

Private Sub ExcelPrn()
    Dim I               As Integer
    Dim k               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    Dim sDate           As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\ACB4110C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
'    sDate = SDT_PROD_DATE_FR.Text
'
'    If SDT_PROD_DATE_FR.Text <> SDT_PROD_DATE_TO.Text Then
'        xlApp.Range("D2").Value = "日期: " & Left(sDate, 4) + "年" + Mid(sDate, 6, 2) + "月" + Mid(sDate, 9, 2) + "日 - " + Mid(SDT_PROD_DATE_TO.Text, 9, 2) + "日"
'    Else
'        xlApp.Range("D2").Value = "日期: " & Left(sDate, 4) + "年" + Mid(sDate, 6, 2) + "月" + Mid(sDate, 9, 2) + "日"
'    End If
    
'    For i = 1 To ss1.MaxRows
'    k = i + 3
'    ss1.Row = i
'    ss1.Col = SS1_MAT_NO:           xlApp.Range("B" + CStr(k)).Value = ss1.Text        '物料号
'    ss1.Col = SS1_PLT:              xlApp.Range("C" + CStr(k)).Value = ss1.Text         '工厂号
'    ss1.Col = SS1_PROC_CD:          xlApp.Range("D" + CStr(k)).Value = ss1.Text         '物料状态
'    ss1.Col = SS1_STL_GRD:          xlApp.Range("E" + CStr(k)).Value = ss1.Text         '钢种
'    ss1.Col = SS1_STD_SPEC:         xlApp.Range("F" + CStr(k)).Value = ss1.Text         '执行标准
'    ss1.Col = SS1_PAINT_FL:         xlApp.Range("G" + CStr(k)).Value = ss1.Text         '冲印指示
'    ss1.Col = SS1_VESSEL_NO:        xlApp.Range("H" + CStr(k)).Value = ss1.Text         '加喷内容
'    ss1.Col = SS1_COLOR_STROKE:     xlApp.Range("I" + CStr(k)).Value = ss1.Text         '色标及备注
'    ss1.Col = SS1_THICK:            xlApp.Range("J" + CStr(k)).Value = ss1.Text         '厚度
'    ss1.Col = SS1_WID:              xlApp.Range("K" + CStr(k)).Value = ss1.Text         '宽度
'    ss1.Col = SS1_LEN:              xlApp.Range("L" + CStr(k)).Value = ss1.Text         '长度
'    ss1.Col = SS1_THK_TOL_MIN:      xlApp.Range("M" + CStr(k)).Value = ss1.Text         '厚度公差最小值
'    ss1.Col = SS1_THK_TOL_MAX:      xlApp.Range("N" + CStr(k)).Value = ss1.Text         '厚度公差最大值
'    ss1.Col = SS1_WID_TOL_MIN:      xlApp.Range("O" + CStr(k)).Value = ss1.Text         '宽度公差最小值
'    ss1.Col = SS1_WID_TOL_MAX:      xlApp.Range("P" + CStr(k)).Value = ss1.Text         '宽度公差最大值
'    ss1.Col = SS1_LEN_TOL_MIN:      xlApp.Range("Q" + CStr(k)).Value = ss1.Text         '长度公差最小值
'    ss1.Col = SS1_LEN_TOL_MAX:      xlApp.Range("R" + CStr(k)).Value = ss1.Text         '长度公差最大值
'    ss1.Col = SS1_SIZE_KIND:        xlApp.Range("S" + CStr(k)).Value = ss1.Text         '定尺区分
'    ss1.Col = SS1_TRIM_FL:          xlApp.Range("T" + CStr(k)).Value = ss1.Text         '切边
'    ss1.Col = SS1_UST_FL:           xlApp.Range("U" + CStr(k)).Value = ss1.Text         '探伤是否
'    ss1.Col = SS1_CL_FL:            xlApp.Range("V" + CStr(k)).Value = ss1.Text         '冷矫直指示/实绩
'    ss1.Col = SS1_GAS_FL:           xlApp.Range("W" + CStr(k)).Value = ss1.Text         '气切割指示/实绩
'    ss1.Col = SS1_HTM_SHOT_BLAST:   xlApp.Range("X" + CStr(k)).Value = ss1.Text         '抛丸指示/实绩
'    ss1.Col = SS1_HTM_METH:         xlApp.Range("Y" + CStr(k)).Value = ss1.Text         '热处理指示/实绩
'    ss1.Col = SS1_CUR_INV:          xlApp.Range("Z" + CStr(k)).Value = ss1.Text         '堆放仓库
'    ss1.Col = SS1_LOC:              xlApp.Range("AA" + CStr(k)).Value = ss1.Text         '堆放位置
'    ss1.Col = SS1_DEFECT:           xlApp.Range("AB" + CStr(k)).Value = ss1.Text         '缺陷
'    ss1.Col = SS1_ORD_REMARK:       xlApp.Range("AC" + CStr(k)).Value = ss1.Text         '订单备注
'    Next i

    
    
Clipboard.Clear
ss1.SetSelection SS1_MAT_NO, 1, SS1_MAT_NO, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("B" + CStr(5), "B" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
    
ss1.SetSelection SS1_STL_GRD, 1, SS1_STL_GRD, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("C" + CStr(5), "C" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
                                                                                
ss1.SetSelection SS1_STDSPEC_ORG_KND, 1, SS1_STDSPEC_ORG_KND, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("D" + CStr(5), "D" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_PAINTNUM, 1, SS1_PAINTNUM, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("E" + CStr(5), "E" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_THICK, 1, SS1_THICK, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("F" + CStr(5), "F" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_WID, 1, SS1_WID, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("G" + CStr(5), "G" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_LEN, 1, SS1_LEN, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("H" + CStr(5), "H" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_SIZE_KIND, 1, SS1_SIZE_KIND, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("I" + CStr(5), "I" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_UST_FL, 1, SS1_UST_FL, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("J" + CStr(5), "J" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_CUST_CD, 1, SS1_CUST_CD, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("K" + CStr(5), "K" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_CL_FL, 1, SS1_CL_FL, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("L" + CStr(5), "L" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_GAS_FL, 1, SS1_GAS_FL, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("M" + CStr(5), "M" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_HTM_METH, 1, SS1_HTM_METH, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("N" + CStr(5), "N" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_ORD_REMARK, 1, SS1_ORD_REMARK, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("O" + CStr(5), "O" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_DEFECT, 1, SS1_DEFECT, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("P" + CStr(5), "P" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear

ss1.SetSelection SS1_THK_TOL_MIN, 1, SS1_THK_TOL_MIN, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("Q" + CStr(5), "Q" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear

ss1.SetSelection SS1_THK_TOL_MAX, 1, SS1_THK_TOL_MAX, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("R" + CStr(5), "R" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_WID_TOL_MIN, 1, SS1_WID_TOL_MIN, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("S" + CStr(5), "S" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_WID_TOL_MAX, 1, SS1_WID_TOL_MAX, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("T" + CStr(5), "T" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_LEN_TOL_MIN, 1, SS1_LEN_TOL_MIN, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("U" + CStr(5), "U" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_LEN_TOL_MAX, 1, SS1_LEN_TOL_MAX, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("V" + CStr(5), "V" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_CE, 1, SS1_CE, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("W" + CStr(5), "W" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_VESSEL_NO, 1, SS1_VESSEL_NO, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("X" + CStr(5), "X" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_SIDEMARK, 1, SS1_SIDEMARK, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("Y" + CStr(5), "Y" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_SEALMEMO, 1, SS1_SEALMEMO, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("Z" + CStr(5), "Z" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
ss1.SetSelection SS1_PAINT_FL, 1, SS1_PAINT_FL, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("AA" + CStr(5), "AA" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear

ss1.SetSelection SS1_CUT_PLAN, 1, SS1_CUT_PLAN, ss1.MaxRows
ss1.ClipboardCopy
xlApp.Range("AB" + CStr(5), "AB" + CStr(ss1.MaxRows + 3)).Select
xlApp.ActiveSheet.Paste
Clipboard.Clear
                                                                                
   
'
'    xlApp.Range("I2").Select
'    xlApp.ActiveSheet.Paste
    
'    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
    ss1.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
'     xlApp.Application.Visible = False
'     xlSheet.Close False
'     xlApp.Quit
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
'    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub txt_cur_inv_DblClick()

    Call txt_cur_inv_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    If App.Title = "AC" Then
        txt_plt.Text = "C1"
        txt_CUR_INV.Text = "00"
    ElseIf App.Title = "BG" Then
        txt_plt.Text = "C1"
        txt_CUR_INV.Text = "00"
    ElseIf App.Title = "CG" Then
        txt_plt.Text = "C3"
        txt_CUR_INV.Text = "ZB"
    ElseIf App.Title = "DG" Then
        txt_plt.Text = "C1"
        txt_CUR_INV.Text = "00"
    ElseIf App.Title = "DE" Then
        txt_plt.Text = "C1"
        txt_CUR_INV.Text = "00"
    End If
    
    Call txt_cur_inv_KeyUp(0, 0)
    Opt_all.Value = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    Set iSumCol = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        CBO_ORD_ITEM.Clear
    End If
    
    If App.Title = "AC" Then
        txt_plt.Text = "C1"
        txt_CUR_INV.Text = "00"
    ElseIf App.Title = "BG" Then
        txt_plt.Text = "C1"
        txt_CUR_INV.Text = "00"
    ElseIf App.Title = "CG" Then
        txt_plt.Text = "C3"
        txt_CUR_INV.Text = "ZB"
    ElseIf App.Title = "DG" Then
        txt_plt.Text = "C1"
        txt_CUR_INV.Text = "00"
    End If
    
    Call txt_cur_inv_KeyUp(0, 0)
    Opt_all.Value = True
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()
 Dim I, j  As Integer
 Dim Back_Wgt As Double
 
 Dim iord As Integer
 
    Dim sQuery As String
     
    If SDB_THK_TO.Value = 0 Then SDB_THK_TO.Value = 9999.99
    If SDB_WID_TO.Value = 0 Then SDB_WID_TO.Value = 9999
    If SDB_LEN_TO.Value = 0 Then SDB_LEN_TO.Value = 9999999
   
    If txt_REC_STS.Text = "" Then txt_REC_STS.Text = "2"
    
    sQuery = Gf_Ms_MakeQuery(Mc1("P-R"), "R", Mc1("pControl"))
 
    If Gf_Total_Display(M_CN1, Proc_Sc("Sc"), sQuery, 0, iSumCnt, iSumCol) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss1.OperationMode = OperationModeNormal
    End If
    
      With ss1
             For I = 1 To .MaxRows
                .Row = I
                .Col = 12
                 Back_Wgt = Back_Wgt + Val(.Text)
             Next I

             .Col = 12
             .Text = str(Back_Wgt)
             
             '紧急订单  李超 20121123
            For iord = 1 To .MaxRows
                .Row = iord
                ss1.Row = .Row:     ss1.Col = SS1_URGNT_FL
                If ss1.Text = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, SS1_MAT_NO, SS1_MAT_NO, .Row, .Row, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_URGNT_FL, SS1_URGNT_FL, .Row, .Row, &HC000&)
                End If
            Next iord
        End With

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

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub txt_cur_inv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
           DD.sWitch = "MS"
           DD.sKey = "C0013"
    
           DD.rControl.Add Item:=txt_CUR_INV
           DD.rControl.Add Item:=txt_cur_inv_nm
           
           DD.nameType = "2"
           Call Gf_Common_DD(M_CN1, KeyCode)
    
    Else
    
        If Len(Trim(txt_CUR_INV.Text)) = txt_CUR_INV.MaxLength Then
            txt_cur_inv_nm.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_CUR_INV.Text, 2)
        Else
            txt_cur_inv_nm.Text = ""
        End If
    End If
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

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

Private Sub TXT_PROC_CD_DblClick()

    Call TXT_PROC_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TXT_PROC_CD_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "C0004"

        DD.rControl.Add Item:=txt_PROC_CD
   
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

    End If
    
End Sub

Private Sub txt_rec_sts_DblClick()

    Call txt_rec_sts_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_rec_sts_KeyUp(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "Z0005"

        DD.rControl.Add Item:=txt_REC_STS
   
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

    End If

End Sub

Private Sub txt_size_knd_DblClick()

    Call txt_size_knd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_cust_cd_DblClick()

    Call txt_cust_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_next_plan_htm_DblClick()

    Call txt_next_plan_htm_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_next_plan_htm_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Q0073"
        
        DD.rControl.Add Item:=txt_next_plan_htm
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
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
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

    End If

End Sub

Private Sub txt_size_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"

        DD.rControl.Add Item:=txt_size_knd
        DD.rControl.Add Item:=txt_size_knd_nm

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_size_knd.Text)) = txt_size_knd.MaxLength Then
            txt_size_knd_nm.Text = Gf_ComnNameFind(M_CN1, "B0043", txt_size_knd.Text, 2)
        Else
            txt_size_knd_nm.Text = ""
        End If
        
    End If
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
 
    If Row > 0 And Col > 0 Then
    
        If ss1.MaxRows = Row Then Exit Sub
    
        Unload ACB1030C
    
        ss1.Col = 1
        ss1.Row = Row
        AIMNO = Trim(ss1.Text)
        'BASE = Trim(Text_PROD_CD.Text)
        'STR1 = Trim(sQuery)
        ACB1030C.FORM_A = "ACB4110C"
        ACB1030C.Show
        
    End If

End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_cust_cd

        DD.nameType = "1"
        Call Gf_Customer_DD(M_CN1, KeyCode)

    End If

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

Private Sub txt_ord_no_LostFocus()

    If TXT_ORD_NO.Text <> "" Then
       If (Len(TXT_ORD_NO.Text) < TXT_ORD_NO.MaxLength) Then
          Call Gp_MsgBoxDisplay("订单号输入未完成！")
          CBO_ORD_ITEM.Text = ""
          TXT_ORD_NO.SetFocus
       End If
    End If

End Sub

Private Sub txt_stdspec_DblClick()

    Call txt_stdspec_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stdspec_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

    End If
    
End Sub

Private Sub txt_surf_grd_DblClick()

    Call txt_surf_grd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_surf_grd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Q0049"
        
        DD.rControl.Add Item:=txt_SURF_GRD
        DD.rControl.Add Item:=txt_surf_grd_nm
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_SURF_GRD.Text)) = txt_SURF_GRD.MaxLength Then
            txt_surf_grd_nm.Text = Gf_ComnNameFind(M_CN1, "Q0049", txt_SURF_GRD.Text, 2)
        Else
            txt_surf_grd_nm.Text = ""
        End If
        
    End If

End Sub

Private Sub txt_trim_fl_DblClick()

    Call txt_trim_fl_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_trim_fl_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "B0021"
        
        DD.rControl.Add Item:=txt_TRIM_FL
        DD.rControl.Add Item:=txt_trim_nm
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_TRIM_FL.Text)) = txt_TRIM_FL.MaxLength Then
            txt_trim_nm.Text = Gf_ComnNameFind(M_CN1, "B0021", txt_TRIM_FL.Text, 2)
        Else
            txt_trim_nm.Text = ""
        End If
        
    End If

End Sub
