VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACB1010C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "物料库存总计查询_ACB1010C"
   ClientHeight    =   9240
   ClientLeft      =   300
   ClientTop       =   1425
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   15300
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_next_plan_htm 
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
      ItemData        =   "ACB1010C.frx":0000
      Left            =   11340
      List            =   "ACB1010C.frx":0002
      TabIndex        =   40
      Top             =   90
      Width           =   750
   End
   Begin VB.TextBox Text_LOC 
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
      Left            =   10905
      MaxLength       =   7
      TabIndex        =   7
      Tag             =   "CD_MANA_NO"
      Top             =   465
      Width           =   1515
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
      Left            =   15990
      MaxLength       =   1
      TabIndex        =   12
      Top             =   180
      Visible         =   0   'False
      Width           =   570
   End
   Begin Threed.SSCheck Chk_NonOrd_Product 
      Height          =   285
      Left            =   7710
      TabIndex        =   4
      Top             =   90
      Width           =   720
      _ExtentX        =   1270
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
      Caption         =   "余材"
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   2715
      MaxLength       =   2
      TabIndex        =   32
      Tag             =   "生产厂"
      Top             =   90
      Width           =   450
   End
   Begin VB.ComboBox cbo_input_fl 
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
      Left            =   13695
      TabIndex        =   6
      Top             =   465
      Width           =   750
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
      Left            =   13695
      TabIndex        =   11
      Top             =   90
      Width           =   750
   End
   Begin VB.TextBox txt_TRIM_NAME 
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
      Left            =   14130
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "钢种"
      Top             =   840
      Width           =   1080
   End
   Begin VB.TextBox txt_TRIM_FL 
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
      Left            =   13695
      MaxLength       =   1
      TabIndex        =   20
      Tag             =   "钢种"
      Top             =   840
      Width           =   420
   End
   Begin VB.TextBox txt_prod_grd 
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
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   9
      Top             =   465
      Width           =   525
   End
   Begin VB.TextBox txt_prod_grd_name 
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
      Left            =   4995
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "钢种"
      Top             =   465
      Width           =   1395
   End
   Begin VB.TextBox Text_size_knd 
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
      Left            =   10905
      MaxLength       =   2
      TabIndex        =   19
      Tag             =   "钢种"
      Top             =   840
      Width           =   420
   End
   Begin VB.TextBox Text_size_knd_name 
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
      Left            =   11340
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   "钢种"
      Top             =   840
      Width           =   1080
   End
   Begin VB.TextBox text_cur_inv_code 
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
      Left            =   7650
      MaxLength       =   2
      TabIndex        =   10
      Top             =   465
      Width           =   495
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
      Left            =   8175
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   465
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox Text_STLGRD 
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
      Left            =   1215
      MaxLength       =   20
      TabIndex        =   8
      Tag             =   "CD_MANA_NO"
      Top             =   465
      Width           =   1965
   End
   Begin VB.TextBox Text_STLGRD_name 
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
      Left            =   12900
      TabIndex        =   24
      Top             =   9315
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox Text_REC_STS_name 
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
      Left            =   12810
      TabIndex        =   23
      Top             =   9450
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox Text_PROC_CD_mate 
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
      Left            =   12405
      TabIndex        =   22
      Top             =   9315
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox Text_PROD_CD 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Left            =   1215
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "SL"
      Top             =   90
      Width           =   405
   End
   Begin VB.TextBox Text_PROD_CD_mate 
      Height          =   270
      Left            =   12315
      TabIndex        =   21
      Top             =   9405
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox Text_PROC_CD 
      Alignment       =   2  'Center
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
      Left            =   4440
      MaxLength       =   3
      TabIndex        =   1
      Top             =   90
      Width           =   525
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   105
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   3360
      Top             =   90
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "进程状态"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   105
      Top             =   465
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "钢种"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   105
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   3360
      Top             =   840
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "宽度"
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
      Left            =   6570
      Top             =   840
      Width           =   1065
      _ExtentX        =   1879
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
   Begin CSTextLibCtl.sidbEdit sidbEdit_size_Athk 
      Height          =   315
      Left            =   1215
      TabIndex        =   13
      Top             =   840
      Width           =   885
      _Version        =   262145
      _ExtentX        =   1561
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
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sidbEdit_size_Bthk 
      Height          =   315
      Left            =   2280
      TabIndex        =   14
      Top             =   840
      Width           =   885
      _Version        =   262145
      _ExtentX        =   1561
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
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sidbEdit_size_Awid 
      Height          =   315
      Left            =   4440
      TabIndex        =   15
      Top             =   840
      Width           =   885
      _Version        =   262145
      _ExtentX        =   1561
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
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sidbEdit_size_Bwid 
      Height          =   315
      Left            =   5505
      TabIndex        =   16
      Top             =   840
      Width           =   885
      _Version        =   262145
      _ExtentX        =   1561
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
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sidbEdit_size_Alen 
      Height          =   315
      Left            =   7650
      TabIndex        =   17
      Top             =   840
      Width           =   885
      _Version        =   262145
      _ExtentX        =   1561
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
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sidbEdit_size_Blen 
      Height          =   315
      Left            =   8730
      TabIndex        =   18
      Top             =   840
      Width           =   885
      _Version        =   262145
      _ExtentX        =   1561
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
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   6570
      Top             =   465
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "堆放仓库"
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
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   9825
      Top             =   840
      Width           =   1045
      _ExtentX        =   1852
      _ExtentY        =   556
      Caption         =   "定尺"
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
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   3360
      Top             =   465
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "等级"
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
   Begin InDate.ULabel ULabel23 
      Height          =   315
      Left            =   12605
      Top             =   840
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "切边"
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
   Begin InDate.ULabel ULabel15 
      Height          =   315
      Left            =   12605
      Top             =   465
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "入库是否"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   12605
      Top             =   90
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "探伤是否"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   1710
      Top             =   90
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
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7965
      Left            =   60
      TabIndex        =   33
      Top             =   1210
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   14049
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "ACB1010C.frx":0004
      Begin Threed.SSPanel SSPanel1 
         Height          =   495
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   873
         _Version        =   196609
         BackColor       =   14737918
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   8970
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "数量合计"
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
         Begin CSTextLibCtl.sidbEdit Text_TOT_SHEETS 
            Height          =   315
            Left            =   10110
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   90
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   16711680
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
            ReadOnly        =   -1  'True
            Insert          =   0   'False
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
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   11790
            Top             =   90
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "重量合计"
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
         Begin CSTextLibCtl.sidbEdit Text_TOT_WGT 
            Height          =   315
            Left            =   12930
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   90
            Width           =   1635
            _Version        =   262145
            _ExtentX        =   2884
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   255
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
            ReadOnly        =   -1  'True
            Insert          =   0   'False
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
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E1FE&
            Caption         =   "吨"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   14655
            TabIndex        =   39
            Top             =   135
            Width           =   195
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E1FE&
            Caption         =   "张"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   11160
            TabIndex        =   37
            Top             =   135
            Width           =   195
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   7440
         Left            =   0
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   525
         Width           =   15165
         _Version        =   393216
         _ExtentX        =   26749
         _ExtentY        =   13123
         _StockProps     =   64
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
         MaxRows         =   1
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACB1010C.frx":0056
      End
   End
   Begin Threed.SSCheck Chk_Ord_Product 
      Height          =   285
      Left            =   6705
      TabIndex        =   3
      Top             =   90
      Width           =   900
      _ExtentX        =   1588
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
      Caption         =   "订单材"
   End
   Begin Threed.SSCheck Chk_ship_product 
      Height          =   285
      Left            =   5295
      TabIndex        =   2
      Top             =   90
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
      Caption         =   "可发货产品"
   End
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   9825
      Top             =   90
      Width           =   1485
      _ExtentX        =   2619
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
   Begin Threed.SSCheck chk_htm_shot_blast 
      Height          =   285
      Left            =   8505
      TabIndex        =   5
      Top             =   90
      Width           =   1140
      _ExtentX        =   2011
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
      Caption         =   "抛丸对象"
   End
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   9825
      Top             =   465
      Width           =   1045
      _ExtentX        =   1852
      _ExtentY        =   556
      Caption         =   "垛位号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   180
      Left            =   8580
      TabIndex        =   27
      Top             =   960
      Width           =   90
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   180
      Left            =   5370
      TabIndex        =   26
      Top             =   960
      Width           =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   180
      Left            =   2145
      TabIndex        =   25
      Top             =   960
      Width           =   90
   End
End
Attribute VB_Name = "ACB1010C"
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
'-- Program ID        ACB1010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Yang Zhibin
'-- Date              2003.9.8
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

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim SumCnt   As Integer
Dim SumCol   As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2   cbo_ust_fl

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(Text_PROD_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(Text_PROC_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(Text_STLGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(sidbEdit_size_Athk, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(sidbEdit_size_Bthk, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(sidbEdit_size_Awid, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(sidbEdit_size_Bwid, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(sidbEdit_size_Alen, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(sidbEdit_size_Blen, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_plt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(Text_LOC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
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
    
    'Duplicate Count
    iDupCnt = 1
    
    'Sum Column Count
    SumCnt = 18
    
   ' Sum Column Setting
    SumCol.Add Item:=5
    SumCol.Add Item:=6
    SumCol.Add Item:=7
    SumCol.Add Item:=8
    SumCol.Add Item:=9
    SumCol.Add Item:=10
    SumCol.Add Item:=11
    SumCol.Add Item:=12
    SumCol.Add Item:=13
    SumCol.Add Item:=14
    SumCol.Add Item:=15
    SumCol.Add Item:=16
    SumCol.Add Item:=17
    SumCol.Add Item:=18
    SumCol.Add Item:=19
    SumCol.Add Item:=20
    SumCol.Add Item:=21
    SumCol.Add Item:=22
        
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
'
'    With ss1
'
'       .Col = 3
'       .Row = -1
'       .TypeVAlign = 2
'
'        .Col = 4
'       .Row = -1
'       .TypeVAlign = 2
'
'        .Col = 7
'       .Row = -1
'       .TypeVAlign = 2
'
'
'        .Col = 9
'       .Row = -1
'       .TypeVAlign = 2
'
'       .Col = 2
'        .TypeHAlign = TypeHAlignRight
'       .Col = 8
'       .CellType = CellTypeNumber
'       .TypeNumberDecPlaces = 3
'       .TypeNumberShowSep = True
'       .TypeNumberSeparator = ","
'       .TypeHAlign = TypeHAlignRight
'    End With
    

End Sub



Private Sub chk_htm_shot_blast_Click(Value As Integer)

    If chk_htm_shot_blast Then
        Text_PROC_CD.Text = "DZB"
    End If
    
End Sub

Private Sub Chk_NonOrd_Product_Click(Value As Integer)
    If Chk_NonOrd_Product.Value = True Then Chk_Ord_Product.Value = False
End Sub

Private Sub Chk_Ord_Product_Click(Value As Integer)
    If Chk_Ord_Product.Value = True Then Chk_NonOrd_Product.Value = False
End Sub

Private Sub Chk_ship_product_Click(Value As Integer)
    If Chk_ship_product.Value = True Then
        Text_PROC_CD.Text = ""
        Text_PROC_CD_mate.Text = ""
    End If
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

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
        
    cbo_ust_fl.AddItem " "
    cbo_ust_fl.AddItem "Y"
    cbo_ust_fl.AddItem "N"
    
    cbo_input_fl.AddItem " "
    cbo_input_fl.AddItem "Y"
    cbo_input_fl.AddItem "N"
    
    cbo_next_plan_htm.AddItem ""
    cbo_next_plan_htm.AddItem "N"
    cbo_next_plan_htm.AddItem "Q"
    cbo_next_plan_htm.AddItem "T"
    
    Call Form_Define
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
'    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    If App.Title = "DE" Then
        text_cur_inv_code.Text = "00"
        Text_PROD_CD.Text = "PP"
    End If
    
    Call text_cur_inv_code_KeyUp(0, 0)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
  
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    Set SumCol = Nothing
    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
    
    Text_TOT_SHEETS.Value = 0
    Text_TOT_WGT.Value = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    Dim sQuery      As String
    Dim sMesg       As String
    Dim sTable      As String
    Dim iDR, iDc    As Integer
    Dim TotalWeight As Single
    Dim TotalSheets As Single
    
    Dim minSIZEthk  As Single
    Dim maxSIZEthk  As Single
    Dim minSIZEwid  As Single
    Dim maxSIZEwid  As Single
    Dim minSIZElen  As Single
    Dim maxSIZElen  As Single
    
    TotalWeight = 0
    TotalSheets = 0
            
    If sidbEdit_size_Athk.Value = 0 Then
        minSIZEthk = 0
    Else
        minSIZEthk = sidbEdit_size_Athk.Value
    End If
    
    If sidbEdit_size_Bthk.Value = 0 Then
        maxSIZEthk = 9999.99
    Else
        maxSIZEthk = sidbEdit_size_Bthk.Value
    End If
      
    If sidbEdit_size_Awid.Value = 0 Then
        minSIZEwid = 0
    Else
        minSIZEwid = sidbEdit_size_Awid.Value
    End If
    
    If sidbEdit_size_Bwid.Value = 0 Then
        maxSIZEwid = 9999.99
    Else
        maxSIZEwid = sidbEdit_size_Bwid.Value
    End If
      
    If sidbEdit_size_Alen.Value = 0 Then
        minSIZElen = 0
    Else
        minSIZElen = sidbEdit_size_Alen.Value
    End If
     
    If sidbEdit_size_Blen.Value = 0 Then
        maxSIZElen = 9999999.9
    Else
        maxSIZElen = sidbEdit_size_Blen.Value
    End If
      
    ss1.ReDraw = False
    
    If Chk_Ord_Product.Value = True Or Chk_NonOrd_Product.Value = True Then
        Call Gp_Sp_ColHidden(ss1, 7, True)
        Call Gp_Sp_ColHidden(ss1, 8, True)
        Call Gp_Sp_ColHidden(ss1, 9, True)
        Call Gp_Sp_ColHidden(ss1, 10, True)
    Else
        Call Gp_Sp_ColHidden(ss1, 7, False)
        Call Gp_Sp_ColHidden(ss1, 8, False)
        Call Gp_Sp_ColHidden(ss1, 9, False)
        Call Gp_Sp_ColHidden(ss1, 10, False)
    End If
    
    Select Case Text_PROD_CD.Text
        Case "SL"
           sTable = "FP_SLAB"
        Case "PP"
           sTable = "GP_PLATE"
        Case "HC"
           sTable = "GP_COIL"
        Case Else
            Call MsgBox("产品分类代码为空" & Chr(10) & "或不规范!请重试。", vbExclamation + vbOKOnly, "警告")
            Text_PROD_CD.Text = ""
            Text_PROD_CD.SetFocus
    End Select
  
    If maxSIZEthk >= minSIZEthk Then
        If maxSIZEwid >= minSIZEwid Then
            If maxSIZElen >= minSIZElen Then
                If Text_PROD_CD.Text = "SL" Then
                    sQuery = "Select Gf_Stlgrd_Detail(STLGRD),"
                    sQuery = sQuery + "NVL(THK,0),NVL(TRUNC(WID/10)*10,0),NVL(TRUNC(LEN /100)* 100,0),"
'                    sQuery = sQuery + "COUNT(CASE when REC_STS <> '1'  THEN 1 END),SUM(CASE when REC_STS <> '1'  THEN WGT ELSE 0 END),"
'                    sQuery = sQuery + "COUNT(CASE when SUBSTR(PROC_CD,1,1) = 'C' THEN 1 END),SUM(CASE when SUBSTR(PROC_CD,1,1) = 'C' THEN WGT ELSE 0 END),"
                    sQuery = sQuery + "DECODE(COUNT(CASE when REC_STS <> '1'  THEN 1 END),0,NULL,COUNT(CASE when REC_STS        <> '1'  THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when REC_STS   <> '1'  THEN WGT ELSE 0 END),0,NULL,SUM(CASE when REC_STS <> '1'  THEN WGT ELSE 0 END)),"
                    
                    sQuery = sQuery + "DECODE(COUNT(CASE when ORD_FL = 1 AND REC_STS = '2' THEN 1 END),0,NULL,       COUNT(CASE when  ORD_FL = 1  AND REC_STS = '2' THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when   ORD_FL = 1 AND REC_STS = '2' THEN WGT ELSE 0 END),0,NULL,SUM(CASE when  ORD_FL = 1  AND REC_STS = '2' THEN WGT ELSE 0 END)),"
                    sQuery = sQuery + "DECODE(COUNT(CASE when ORD_FL = 2 AND REC_STS = '2' THEN 1 END),0,NULL,       COUNT(CASE when  ORD_FL = 2  AND REC_STS = '2' THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when   ORD_FL = 2 AND REC_STS = '2' THEN WGT ELSE 0 END),0,NULL,SUM(CASE when  ORD_FL = 2  AND REC_STS = '2' THEN WGT ELSE 0 END)),"
                    sQuery = sQuery + "NULL,  NULL,"    'CGD
                    sQuery = sQuery + "NULL,  NULL,"    'DAB
                    sQuery = sQuery + "NULL,  NULL,"    'QAB
                    sQuery = sQuery + "NULL,  NULL,"    'QAE
                    sQuery = sQuery + "DECODE(COUNT(CASE when SUBSTR(PROC_CD,1,1) = 'X' THEN 1 END),0,NULL,COUNT(CASE when SUBSTR(PROC_CD,1,1)        = 'X'  THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when SUBSTR(PROC_CD,1,1)   = 'X' THEN WGT ELSE 0 END),0,NULL,SUM(CASE when SUBSTR(PROC_CD,1,1) = 'X'  THEN WGT ELSE 0 END)),"
                    
                    sQuery = sQuery + "DECODE(COUNT(CASE when SUBSTR(PROC_CD,1,1) = 'C' THEN 1 END),0,NULL,COUNT(CASE when SUBSTR(PROC_CD,1,1)        = 'C'  THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when SUBSTR(PROC_CD,1,1)   = 'C' THEN WGT ELSE 0 END),0,NULL,SUM(CASE when SUBSTR(PROC_CD,1,1) = 'C'  THEN WGT ELSE 0 END)),"
                    sQuery = sQuery + "STLGRD"
                Else
                    sQuery = "Select APLY_STDSPEC,"
                    sQuery = sQuery + "NVL(THK,0),"
                    If Text_PROD_CD.Text = "HC" Then
                        sQuery = sQuery + "ROUND(NVL(WID,0)/ 10) * 10,0,"
                    Else
                        sQuery = sQuery + "NVL(WID,0),NVL(LEN,0),"
                    End If
                    sQuery = sQuery + "DECODE(COUNT(CASE when REC_STS <> '1'  THEN 1 END),0,NULL,COUNT(CASE when REC_STS        <> '1'  THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when REC_STS   <> '1'  THEN WGT ELSE 0 END),0,NULL,SUM(CASE when REC_STS <> '1'  THEN WGT ELSE 0 END)),"
                    
                    sQuery = sQuery + "DECODE(COUNT(CASE when ORD_FL = 1 AND REC_STS = '2' THEN 1 END),0,NULL,       COUNT(CASE when  ORD_FL = 1  AND REC_STS = '2' THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when   ORD_FL = 1 AND REC_STS = '2' THEN WGT ELSE 0 END),0,NULL,SUM(CASE when  ORD_FL = 1  AND REC_STS = '2' THEN WGT ELSE 0 END)),"
                    sQuery = sQuery + "DECODE(COUNT(CASE when ORD_FL = 2 AND REC_STS = '2' THEN 1 END),0,NULL,       COUNT(CASE when  ORD_FL = 2  AND REC_STS = '2' THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when   ORD_FL = 2 AND REC_STS = '2' THEN WGT ELSE 0 END),0,NULL,SUM(CASE when  ORD_FL = 2  AND REC_STS = '2' THEN WGT ELSE 0 END)),"
                    
                    sQuery = sQuery + "DECODE(COUNT(CASE when PROC_CD = 'CGD' THEN 1 END),0,NULL,COUNT(CASE when PROC_CD        = 'CGD' THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when PROC_CD   = 'CGD' THEN WGT ELSE 0 END),0,NULL,SUM(CASE when PROC_CD = 'CGD' THEN WGT ELSE 0 END)),"
                     
                    sQuery = sQuery + "DECODE(COUNT(CASE when SUBSTR(PROC_CD,1,1) = 'D' THEN 1 END),0,NULL,COUNT(CASE when SUBSTR(PROC_CD,1,1)        = 'D'  THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when SUBSTR(PROC_CD,1,1)   = 'D' THEN WGT ELSE 0 END),0,NULL,SUM(CASE when SUBSTR(PROC_CD,1,1) = 'D'  THEN WGT ELSE 0 END)),"
                    
                    sQuery = sQuery + "DECODE(COUNT(CASE when PROC_CD = 'QAB' THEN 1 END),0,NULL,COUNT(CASE when PROC_CD        = 'QAB'  THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when PROC_CD   = 'QAB' THEN WGT ELSE 0 END),0,NULL,SUM(CASE when PROC_CD = 'QAB'  THEN WGT ELSE 0 END)),"
                    
                    sQuery = sQuery + "DECODE(COUNT(CASE when PROC_CD = 'QAE' THEN 1 END),0,NULL,COUNT(CASE when PROC_CD        = 'QAE'  THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when PROC_CD   = 'QAE' THEN WGT ELSE 0 END),0,NULL,SUM(CASE when PROC_CD = 'QAE'  THEN WGT ELSE 0 END)),"
                   
                    sQuery = sQuery + "DECODE(COUNT(CASE when SUBSTR(PROC_CD,1,1) = 'X' THEN 1 END),0,NULL,COUNT(CASE when SUBSTR(PROC_CD,1,1)        = 'X'  THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when SUBSTR(PROC_CD,1,1)   = 'X' THEN WGT ELSE 0 END),0,NULL,SUM(CASE when SUBSTR(PROC_CD,1,1) = 'X'  THEN WGT ELSE 0 END)),"
                     
                    sQuery = sQuery + "DECODE(COUNT(CASE when REC_STS = '1'   THEN 1 END),0,NULL,COUNT(CASE when REC_STS        = '1'   THEN 1 END)),"
                    sQuery = sQuery + "DECODE(SUM(CASE when REC_STS   = '1'   THEN WGT ELSE 0 END),0,NULL,SUM(CASE when REC_STS = '1'   THEN WGT ELSE 0 END)),"
                                     
                    sQuery = sQuery + "APLY_STDSPEC"
                End If
                sQuery = sQuery + "  From  " & sTable
                sQuery = sQuery + "   Where  REC_STS IN ('1','2') "
                
                If Chk_ship_product.Value = True Then
                    sQuery = sQuery + "   AND NVL(PROC_CD,' ')   Like 'X%' "
                End If
                
                If Chk_Ord_Product.Value = True Then
                    sQuery = sQuery + "   AND NVL(ORD_FL,' ')    =  '1' "
                End If
                
                If Chk_NonOrd_Product.Value = True Then
                    sQuery = sQuery + "   AND NVL(ORD_FL,' ')    =  '2' "
                End If
                
                sQuery = sQuery + "       AND NVL(PROC_CD,' ')   Like '" + Trim(Text_PROC_CD.Text) + "%' "
                
                If Text_PROD_CD.Text = "SL" Then
                    sQuery = sQuery + "   AND NVL(STLGRD,' ') Like '" + Trim(Text_STLGRD.Text) + "%' "
                    sQuery = sQuery + "   AND NVL(THK,0) >= " + Str$(minSIZEthk)
                    sQuery = sQuery + "   AND NVL(THK,0) <= " + Str$(maxSIZEthk)
                    sQuery = sQuery + "   AND NVL(TRUNC(WID/10)*10,0) >= " + Str$(minSIZEwid)
                    sQuery = sQuery + "   AND NVL(TRUNC(WID/10)*10,0) <= " + Str$(maxSIZEwid)
                    sQuery = sQuery + "   AND NVL(TRUNC(LEN /100)* 100,0) >= " + Str$(minSIZElen)
                    sQuery = sQuery + "   AND NVL(TRUNC(LEN /100)* 100,0) <= " + Str$(maxSIZElen)
                                        
                Else
                    sQuery = sQuery + "   AND NVL(APLY_STDSPEC,' ') Like '" + Trim(Text_STLGRD.Text) + "%' "
                    sQuery = sQuery + "   AND NVL(THK,0) >= " + Str$(minSIZEthk)
                    sQuery = sQuery + "   AND NVL(THK,0) <= " + Str$(maxSIZEthk)
                    sQuery = sQuery + "   AND NVL(WID,0) >= " + Str$(minSIZEwid)
                    sQuery = sQuery + "   AND NVL(WID,0) <= " + Str$(maxSIZEwid)
                    sQuery = sQuery + "   AND NVL(LEN,0) >= " + Str$(minSIZElen)
                    sQuery = sQuery + "   AND NVL(LEN,0) <= " + Str$(maxSIZElen)
                End If
                
                
                sQuery = sQuery + " AND CUR_INV           LIKE '" + Trim(text_cur_inv_code.Text) + "%'"
                sQuery = sQuery + " AND NVL(LOC,' ')      LIKE '" + Trim(Text_LOC.Text) + "%' "
                sQuery = sQuery + " AND NVL(SIZE_KND,' ') LIKE '" + Trim(Text_size_knd.Text) + "%'"
                sQuery = sQuery + " AND NVL(PROD_GRD,' ') LIKE '" + Trim(txt_prod_grd.Text) + "%'"
                sQuery = sQuery + " AND PLT               LIKE '" + Trim(txt_plt.Text) + "%'"
                
                If UCase(Trim(cbo_input_fl.Text)) = "Y" Then
                       sQuery = sQuery + "   AND BED_PILE_DATE IS NOT NULL "
                ElseIf UCase(Trim(cbo_input_fl.Text)) = "N" Then
                       sQuery = sQuery + "   AND BED_PILE_DATE IS NULL "
                End If
                
                If Text_PROD_CD.Text = "PP" Then
                    If UCase(Trim(cbo_ust_fl.Text)) = "Y" Then
                           sQuery = sQuery + "   AND (  (SUBSTR(PROC_CD,1,1) <= 'D' AND NVL(UST_FL,'X') <> 'X') "
                           sQuery = sQuery + "       OR (GF_USTPLATE_CHK(PLATE_NO) <> 'X')) "
                    ElseIf UCase(Trim(cbo_ust_fl.Text)) = "N" Then
                           sQuery = sQuery + "   AND (   NVL(UST_FL,'X') = 'X' "
                           sQuery = sQuery + "       OR (SUBSTR(PROC_CD,1,1) > 'D' AND GF_USTPLATE_CHK(PLATE_NO) = 'X')) "
                    End If

                    sQuery = sQuery + " AND NVL(TRIM_FL,'N')  LIKE '" + Trim(txt_TRIM_FL.Text) + "%'"
                    sQuery = sQuery + " AND PROD_CD           =    'PP' "
                    
                    If chk_htm_shot_blast Or cbo_next_plan_htm.Text <> "" Then
                        sQuery = sQuery + "   AND  PROC_CD                       =        'DZB' "
                    End If
                    
                    If chk_htm_shot_blast Then
                        sQuery = sQuery + "   AND  NVL(HTM_SHOT_BLAST,'NN')     <>       'NN'  "
                    End If
                    
                    If cbo_next_plan_htm.Text <> "" Then
                        sQuery = sQuery + "   AND  NVL(HTM_METH1,' ')||NVL(HTM_RLT_METH1,'X')||NVL(HTM_METH2,' ')||NVL(HTM_RLT_METH2,'X')||NVL(HTM_METH3,' '||NVL(HTM_RLT_METH3,'X'))  LIKE '%" + Trim(cbo_next_plan_htm.Text) + "X%' "
                    End If

                End If
                
                If Text_PROD_CD.Text = "SL" Then
                    sQuery = sQuery + "   Group By STLGRD,THK, TRUNC(WID/10)*10, TRUNC(LEN /100)* 100 "
                    sQuery = sQuery + "   Order By STLGRD,THK, TRUNC(WID/10)*10, TRUNC(LEN /100)* 100 "
                ElseIf Text_PROD_CD.Text = "HC" Then
                    sQuery = sQuery + "   Group By APLY_STDSPEC,THK,ROUND(NVL(WID,0)/ 10) * 10 "
                    sQuery = sQuery + "   Order By 1,2,3 "
                Else
                    sQuery = sQuery + "   Group By APLY_STDSPEC,THK,WID,LEN "
                    sQuery = sQuery + "   Order By 1,THK,WID,LEN "
                End If
                        
                sMesg = Gf_Ms_NeceCheck(nControl)
                If sMesg = "OK" Then
                
                    sMesg = Gf_Ms_NeceCheck2(mControl)
                    If sMesg = "OK" Then
                    
                        If Gf_Multi_Stotal_Display(M_CN1, Proc_Sc("Sc"), sQuery, 1, 1, SumCnt, SumCol) Then
'                        If Gf_Total_Display(M_CN1, Proc_Sc("Sc"), sQuery, iDupCnt, localize_SumCnt, SumCol) Then
                            
                            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                        End If
                
                    Else
                        sMesg = sMesg + " Must input according to length of item"
                        Call Gp_MsgBoxDisplay(sMesg)
                    End If
                
                 Else
                    sMesg = sMesg + " Must input necessarily"
                    Call Gp_MsgBoxDisplay(sMesg)
                 End If

                 With ss1
                     If .MaxRows = 0 Then
                        Text_TOT_SHEETS.Text = "0"
                        Text_TOT_WGT.Value = 0
                     Else
                        .ROW = .MaxRows
                        .Col = 5: Text_TOT_SHEETS.Text = Val(.Value & "")
                        .Col = 6: Text_TOT_WGT.Text = Val(.Value & "")
                     End If
                 End With
                 
            Else
                Call MsgBox("长度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
            End If
        Else
            Call MsgBox("宽度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
        End If
    Else
        Call MsgBox("厚度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
        
    End If
    
    ss1.ReDraw = True

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

Private Sub sidbEdit_size_Alen_Change()
    If sidbEdit_size_Alen.Value > 0 And sidbEdit_size_Blen.Value < sidbEdit_size_Alen.Value Then
        sidbEdit_size_Blen.Value = sidbEdit_size_Alen.Value
    End If
End Sub

Private Sub sidbEdit_size_Athk_Change()
    If sidbEdit_size_Athk.Value > 0 And sidbEdit_size_Bthk.Value < sidbEdit_size_Athk.Value Then
        sidbEdit_size_Bthk.Value = sidbEdit_size_Athk.Value
    End If
End Sub

Private Sub sidbEdit_size_Awid_Change()
    If sidbEdit_size_Awid.Value > 0 And sidbEdit_size_Bwid.Value < sidbEdit_size_Awid.Value Then
        sidbEdit_size_Bwid.Value = sidbEdit_size_Awid.Value
    End If
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)

    Dim iRowCount As Long
    Dim MaxRow    As Long
    Dim iRow      As Integer
    Dim grd       As String
    Dim sStlgrd   As String
  
    If ss1.MaxRows < 1 Or ROW = 0 Or ss1.MaxRows = ROW Then Exit Sub
        
        ss1.ROW = ss1.ActiveRow
        ss1.Col = 23
        sStlgrd = Trim(ss1.Text)
        
        If sStlgrd = "" Then Exit Sub
                
        Unload ACB1020C
        ss1.ROW = ss1.ActiveRow
        ss1.Col = 2
        
        ACB1020C.sdb_thk_fr = Trim(ss1.Value)
        ACB1020C.sdb_thk_to = Trim(ss1.Value)
               
        ss1.Col = 3
        ACB1020C.sdb_wid_fr = Trim(ss1.Value)
        If Text_PROD_CD.Text = "SL" Then
            ACB1020C.SDB_WID_TO = Val(ss1.Value & "") + 9
        ElseIf Text_PROD_CD.Text = "HC" Then
            ACB1020C.sdb_wid_fr = Val(ss1.Value & "") - 5
            ACB1020C.SDB_WID_TO = Val(ss1.Value & "") + 5
        Else
            ACB1020C.SDB_WID_TO = Trim(ss1.Value)
        End If
        
        ss1.Col = 4
        ACB1020C.sdb_len_fr = Trim(ss1.Value)
        If Text_PROD_CD.Text = "SL" Then
            ACB1020C.SDB_LEN_TO = Val(ss1.Value & "") + 99
        Else
            ACB1020C.SDB_LEN_TO = Trim(ss1.Value)
        End If
                
        'ACB1020C.dtp_ins_date_PROD_DATE1.Text = "20040101"
        
        ACB1020C.text_cur_inv_code = text_cur_inv_code.Text
        ACB1020C.text_cur_inv = text_cur_inv.Text
        ACB1020C.CBO_PROD_CD.Text = Text_PROD_CD.Text
        ACB1020C.Text_PROC_CD.Text = Text_PROC_CD.Text
        ACB1020C.Text_size_knd = Text_size_knd.Text
        ACB1020C.txt_TRIM_FL = txt_TRIM_FL.Text
        ACB1020C.txt_prod_grd = txt_prod_grd.Text
        ACB1020C.DTP_PROD_FR = ""
        ACB1020C.DTP_PROD_TO = ""
        ACB1020C.TXT_BED_PILE_DATE = cbo_input_fl.Text
        ACB1020C.TXT_UST_FL = cbo_ust_fl.Text
        ACB1020C.CBO_PLT = txt_plt.Text
        'ACB1020C.txt_rec_sts = "2"

        ACB1020C.Text_STLGRD = sStlgrd
        
'        ACB1020C.cbo_next_plan_htm.Text = cbo_next_plan_htm.Text
        
        If chk_htm_shot_blast.Value = True Then ACB1020C.chk_htm_shot_blast.Value = True
        If Chk_Ord_Product.Value = True Then ACB1020C.Option_ORD_FL_Y.Value = 1
        If Chk_NonOrd_Product.Value = True Then ACB1020C.Option_ORD_FL_N.Value = 1
        
        ACB1020C.Refer_Fl = "N"
        
        ACB1020C.Show
        ACB1020C.Form_Ref

End Sub

Private Sub ss1_LostFocus()
'
'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub text_cur_inv_code_Change()

    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
          text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
          Exit Sub
    Else
          text_cur_inv.Text = ""
    End If
End Sub

Private Sub text_cur_inv_code_DblClick()

    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
    
        DD.rControl.Add Item:=text_cur_inv_code
        DD.rControl.Add Item:=text_cur_inv
        
    
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
       
        If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
            text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
            Exit Sub
        Else
            text_cur_inv.Text = ""
        End If
    End If
End Sub

Private Sub Text_PROC_CD_Change()

    If Not Text_PROC_CD.Text = "" Then
        If Len(Text_PROC_CD.Text) = Text_PROC_CD.MaxLength Then
            Text_PROC_CD.Text = StrConv(Text_PROC_CD.Text, vbUpperCase)
            
'            Select Case Text_PROC_CD.Text
'               Case "BAA", "BAB", "BAC", "BAD", "BAE", "BAF"
'               Case "BBA", "BBB", "BBC", "BBD", "BBE", "BBF"
'               Case "BCA", "BCB", "BCC", "BCD", "BCE", "BCF"
'               Case "BDA", "BDB", "BDC", "BDD", "BDE", "BDF"
'               Case "BEA", "BEB", "BEC", "BED", "BEE", "BEF"
'               Case "CAA", "CAB", "CAC", "CAD", "CAE", "CAF"
'               Case "CBA", "CBB", "CBC", "CBD", "CBE", "CBF"
'               Case "CGA", "CGB", "CGC", "CGD", "CGE", "CGF"
'               Case "DAA", "DAB", "DAC", "DAD", "DAE", "DAF"
'               Case "DBA", "DBB", "DBC", "DBD", "DBE", "DBF"
'               Case "QAA", "QAB", "QAC", "QAD", "QAE", "QAF"
'               Case "XAA", "XAB", "XAC", "XAD", "XAE", "XAF"
'               Case ""
'                Case Else
'                    Call MsgBox("进程代码不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
'                    Text_PROC_CD.Text = ""
'                    'Text_PROC_CD_Name.Text = ""
'            End Select
        End If
        Chk_ship_product.Value = ssCBUnchecked
        If Text_PROC_CD.Text <> "DZB" Then
            chk_htm_shot_blast.Value = False
            cbo_next_plan_htm.Text = ""
        End If
        
    End If

End Sub

Private Sub Text_PROC_CD_DblClick()

    Call Text_PROC_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_PROC_CD_LostFocus()

    If Text_PROC_CD.Text <> "" Then
        If (Len(Text_PROC_CD.Text) < Text_PROC_CD.MaxLength) Then
            Call Gp_MsgBoxDisplay("进程代码输入未完成！")
            'Text_PROD_CD.Text = ""
            Text_PROC_CD.SetFocus
        End If
    End If
    
End Sub

Private Sub Text_PROD_CD_Change()

    ULabel2.Caption = "钢种"
    chk_htm_shot_blast.Visible = False
    ULabel10.Visible = False
    cbo_next_plan_htm.Visible = False
            
    Select Case Text_PROD_CD.Text
        Case "S", "s", "SL"
            Text_PROD_CD.Text = "SL"
        Case "P", "p", "PP"
            Text_PROD_CD.Text = "PP"
            ULabel2.Caption = "标准号"
            chk_htm_shot_blast.Visible = True
            ULabel10.Visible = True
            cbo_next_plan_htm.Visible = True
        Case "H", "h", "HC"
            Text_PROD_CD.Text = "HC"
            ULabel2.Caption = "标准号"
        Case ""
            Text_PROD_CD.Text = ""
        Case Else
            Text_PROD_CD.Text = ""
            Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
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
        DD.rControl.Add Item:=Text_PROD_CD_mate

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

    If Len(Trim(Text_PROD_CD.Text)) = Text_PROD_CD.MaxLength Then
        Text_PROD_CD_mate.Text = Gf_ComnNameFind(M_CN1, "B0005", Text_PROD_CD.Text, 2)
    Else
        Text_PROD_CD_mate.Text = ""
    End If
    
End Sub

Private Sub Text_PROC_CD_KeyUp(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "C0004"

        DD.rControl.Add Item:=Text_PROC_CD
        DD.rControl.Add Item:=Text_PROC_CD_mate
   
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(Text_PROC_CD.Text)) = Text_PROC_CD.MaxLength Then
        Text_PROC_CD_mate.Text = Gf_ComnNameFind(M_CN1, "C0004", Text_PROC_CD.Text, 2)
    Else
        Text_PROC_CD_mate.Text = ""
    End If
    
End Sub

Private Sub Text_size_knd_DblClick()

    Call Text_size_knd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_stlgrd_DblClick()

    Call text_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
       
        If Text_PROD_CD.Text = "SL" Then
            DD.sWitch = "MS"
            DD.rControl.Add Item:=Text_STLGRD
            
            DD.nameType = "1"
            Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        Else
            DD.sWitch = "MS"
            DD.rControl.Add Item:=Text_STLGRD
    
            Call Gf_StdSPEC_DD2(M_CN1, KeyCode)
        End If
        
    End If

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
        
        If txt_next_plan_htm.Text <> "" Then
            Text_PROC_CD.Text = "DZB"
        End If
        
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

Private Sub txt_prod_grd_Change()
    If Len(Trim(txt_prod_grd.Text)) = txt_prod_grd.MaxLength Then
        txt_prod_grd_name.Text = Gf_ComnNameFind(M_CN1, "Q0034", txt_prod_grd.Text, 1)
        Exit Sub
    Else
        txt_prod_grd_name.Text = ""
    End If
End Sub

Private Sub txt_prod_grd_DblClick()

    Call txt_prod_grd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_prod_grd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0034"

        DD.rControl.Add Item:=txt_prod_grd

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub Text_size_knd_Change()
    If Len(Trim(Text_size_knd.Text)) = Text_size_knd.MaxLength Then
        Text_size_knd_name.Text = Gf_ComnNameFind(M_CN1, "B0043", Text_size_knd.Text, 2)
        Exit Sub
    Else
        Text_size_knd_name.Text = ""
    End If
End Sub

Private Sub Text_size_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"

        DD.rControl.Add Item:=Text_size_knd

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub txt_trim_fl_Change()
    If Len(Trim(txt_TRIM_FL.Text)) = txt_TRIM_FL.MaxLength Then
        txt_TRIM_NAME.Text = Gf_ComnNameFind(M_CN1, "B0021", txt_TRIM_FL.Text, 2)
        txt_TRIM_FL.Text = Trim(txt_TRIM_FL.Text)
        Exit Sub
    Else
        txt_TRIM_NAME.Text = ""
        txt_TRIM_FL.Text = ""
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
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If

End Sub

