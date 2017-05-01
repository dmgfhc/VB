VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AFH2020C 
   Caption         =   "板坯切割实绩修改及查询界面_AFH2020C"
   ClientHeight    =   9225
   ClientLeft      =   375
   ClientTop       =   2355
   ClientWidth     =   15225
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_act_stlgrd 
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
      Left            =   3315
      MaxLength       =   11
      TabIndex        =   32
      Top             =   1710
      Width           =   1290
   End
   Begin VB.CommandButton cmd_down 
      Caption         =   "▼"
      Height          =   255
      Left            =   3390
      TabIndex        =   30
      Top             =   620
      Width           =   375
   End
   Begin VB.CommandButton cmd_up 
      Caption         =   "▲"
      Height          =   255
      Left            =   3390
      TabIndex        =   31
      Top             =   340
      Width           =   375
   End
   Begin VB.TextBox txt_element_content_o 
      Alignment       =   1  'Right Justify
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
      Left            =   12225
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   8760
      Width           =   870
   End
   Begin VB.TextBox txt_element_content_n 
      Alignment       =   1  'Right Justify
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
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   8760
      Width           =   870
   End
   Begin VB.TextBox txt_element_content_h 
      Alignment       =   1  'Right Justify
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
      Left            =   9675
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   8760
      Width           =   870
   End
   Begin VB.TextBox txt_element_content_al 
      Alignment       =   1  'Right Justify
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
      Left            =   8385
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   8760
      Width           =   870
   End
   Begin VB.TextBox txt_element_content_s 
      Alignment       =   1  'Right Justify
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
      Left            =   7095
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   8760
      Width           =   870
   End
   Begin VB.TextBox txt_element_content_p 
      Alignment       =   1  'Right Justify
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
      Left            =   5805
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   8760
      Width           =   870
   End
   Begin VB.TextBox txt_element_content_si 
      Alignment       =   1  'Right Justify
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
      Left            =   4515
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   8760
      Width           =   870
   End
   Begin VB.TextBox txt_element_content_mn 
      Alignment       =   1  'Right Justify
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
      Left            =   3210
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   8760
      Width           =   870
   End
   Begin VB.TextBox txt_element_content_c 
      Alignment       =   1  'Right Justify
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
      Left            =   1875
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   8760
      Width           =   870
   End
   Begin VB.TextBox txt_act_stlgrd_c 
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
      TabIndex        =   20
      Top             =   1710
      Width           =   3390
   End
   Begin VB.TextBox txt_stlgrd_n 
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
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1320
      Width           =   3390
   End
   Begin VB.TextBox txt_emp_cd 
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
      Height          =   315
      Left            =   13575
      Locked          =   -1  'True
      TabIndex        =   18
      Tag             =   "作业人员"
      Top             =   450
      Width           =   1095
   End
   Begin Threed.SSOption Option2 
      Height          =   285
      Left            =   1515
      TabIndex        =   17
      Top             =   90
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "静态显示"
      Value           =   -1
   End
   Begin Threed.SSOption Option1 
      Height          =   285
      Left            =   180
      TabIndex        =   16
      Top             =   90
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "动态显示"
   End
   Begin Threed.SSCommand cmd_work_order 
      Height          =   555
      Left            =   120
      TabIndex        =   15
      Top             =   1470
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   979
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "查询指示"
   End
   Begin VB.TextBox txt_cast_no 
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
      Height          =   315
      Left            =   3315
      TabIndex        =   9
      Top             =   930
      Width           =   735
   End
   Begin VB.TextBox txt_heat_seq 
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
      Height          =   315
      Left            =   5775
      TabIndex        =   8
      Top             =   930
      Width           =   735
   End
   Begin VB.TextBox txt_stlgrd 
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
      Left            =   3315
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   1300
   End
   Begin VB.ComboBox cbo_shift 
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
      IMEMode         =   1  'ON
      ItemData        =   "AFH2020C.frx":0000
      Left            =   8235
      List            =   "AFH2020C.frx":0002
      TabIndex        =   6
      Top             =   450
      Width           =   735
   End
   Begin VB.ComboBox cbo_group 
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
      ItemData        =   "AFH2020C.frx":0004
      Left            =   10695
      List            =   "AFH2020C.frx":0006
      TabIndex        =   5
      Top             =   450
      Width           =   735
   End
   Begin VB.ComboBox cbo_prc_line 
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
      ItemData        =   "AFH2020C.frx":0008
      Left            =   5775
      List            =   "AFH2020C.frx":000A
      TabIndex        =   4
      Tag             =   "连铸机号"
      Top             =   450
      Width           =   700
   End
   Begin VB.ComboBox cbo_heat_no 
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
      Left            =   1755
      TabIndex        =   3
      Tag             =   "炉号"
      Top             =   450
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2820
      Top             =   0
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
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   14205
      TabIndex        =   2
      Text            =   "B1"
      Top             =   180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_prc 
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
      Height          =   315
      Left            =   14250
      TabIndex        =   1
      Text            =   "BF"
      Top             =   90
      Visible         =   0   'False
      Width           =   735
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   6615
      Left            =   90
      TabIndex        =   0
      Top             =   2100
      Width           =   15045
      _Version        =   393216
      _ExtentX        =   26538
      _ExtentY        =   11668
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   37
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AFH2020C.frx":000C
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   180
      Top             =   450
      Width           =   1530
      _ExtentX        =   2699
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   4185
      Top             =   450
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "连铸机号"
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
      Left            =   6645
      Top             =   450
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "班次"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   9105
      Top             =   450
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "班别"
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
      Left            =   11985
      Top             =   450
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "作业人员"
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
      Left            =   1755
      Top             =   930
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "连浇号"
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
      Left            =   4185
      Top             =   930
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "连浇号内炉序号"
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
      Left            =   1755
      Top             =   1320
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "目标钢种号"
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
      Left            =   6645
      Top             =   930
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "冷却水总用量"
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
      Left            =   9105
      Top             =   930
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "金属收得率"
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
      Left            =   11985
      Top             =   930
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "钢铁料消耗"
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
   Begin CSTextLibCtl.sidbEdit txt_cal2 
      Height          =   315
      Left            =   13575
      TabIndex        =   10
      Top             =   930
      Width           =   1065
      _Version        =   262145
      _ExtentX        =   1879
      _ExtentY        =   556
      _StockProps     =   125
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
   Begin CSTextLibCtl.sidbEdit txt_cal1 
      Height          =   315
      Left            =   10695
      TabIndex        =   11
      Top             =   930
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   125
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
      NumIntDigits    =   2
      ShowZero        =   0   'False
      MaxValue        =   9999
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_water_flow 
      Height          =   315
      Left            =   8235
      TabIndex        =   12
      Top             =   930
      Width           =   705
      _Version        =   262145
      _ExtentX        =   1244
      _ExtentY        =   556
      _StockProps     =   125
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   1755
      Top             =   1710
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "实际钢种号"
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
   Begin InDate.ULabel ULabel29 
      Height          =   315
      Left            =   120
      Top             =   8760
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "中包成分(%)"
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
      Left            =   6705
      Top             =   8760
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "S"
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
   Begin InDate.ULabel ULabel33 
      Height          =   315
      Left            =   5415
      Top             =   8760
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "P"
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
      Left            =   4125
      Top             =   8760
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Si"
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
      Left            =   2820
      Top             =   8760
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Mn"
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
      Left            =   1485
      Top             =   8760
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "C"
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
   Begin InDate.ULabel ULabel45 
      Height          =   315
      Left            =   8010
      Top             =   8760
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Al"
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
   Begin InDate.ULabel ULabel46 
      Height          =   315
      Left            =   9300
      Top             =   8760
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "[H]"
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
   Begin InDate.ULabel ULabel47 
      Height          =   315
      Left            =   10545
      Top             =   8760
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "[N]"
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
   Begin InDate.ULabel ULabel48 
      Height          =   315
      Left            =   11850
      Top             =   8760
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "[O]"
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
   Begin Threed.SSPanel SSPpdt 
      Height          =   345
      Left            =   14235
      TabIndex        =   33
      Top             =   1665
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   609
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "指示"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPsend 
      Height          =   345
      Left            =   13410
      TabIndex        =   34
      Top             =   1665
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   609
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   16761087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "实际"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   11745
      TabIndex        =   14
      Top             =   1005
      Width           =   195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kg/t"
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
      Left            =   14685
      TabIndex        =   13
      Top             =   1020
      Width           =   420
   End
End
Attribute VB_Name = "AFH2020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name      SLABCUT
'-- Program ID        Afh2020c
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2003.8.8
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType     As String       'Form Type
Public Toolbar_St   As String       'Active Form ToolBar Setting
Public sAuthority   As String       'Active Form Authority Setting
Public QueryYN      As Boolean
Public sQuery_Rt    As String

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

Dim sBef_ccm_line As Boolean

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Hsheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
               Call Gp_Ms_Collection(cbo_heat_no, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(cbo_shift, " ", " ", " ", " ", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(cbo_prc_line, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(cbo_group, " ", " ", " ", " ", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_emp_cd, " ", " ", " ", " ", "r", "a", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_cast_no, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_heat_seq, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_stlgrd, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_water_flow, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_cal1, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_cal2, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_stlgrd_n, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_act_stlgrd, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_act_stlgrd_c, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_element_content_c, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_element_content_mn, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_element_content_si, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_element_content_p, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_element_content_s, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_element_content_al, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_element_content_h, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_element_content_n, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_element_content_o, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
    'MASTER Collection
    'Mc1.Add Item:="Afh2020c.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="AFH2020C.P_REFER", Key:="P-R"
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
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     'add by guoli at 200702131526 for 刘汝营
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     ''''''''''''''''''''''''''''''''''''''''
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    'ADDTIONAL QUALITY_L2
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 34, "p", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AFH2020C.P_SMODIFY", Key:="P-M"
    sc1.Add Item:="AFH2020C.P_SREFER", Key:="P-R"
    sc1.Add Item:="AFH2020C.P_SONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=2, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, 7, True)
    Call Gp_Sp_ColHidden(ss1, 18, True)
    Call Gp_Sp_ColHidden(ss1, 19, True)
    Call Gp_Sp_ColHidden(ss1, 20, True)
    Call Gp_Sp_ColHidden(ss1, 25, True)
    Call Gp_Sp_ColHidden(ss1, 26, True)
    Call Gp_Sp_ColHidden(ss1, 27, True)
'    Call Gp_Sp_ColHidden(ss1, 28, True)
    Call Gp_Sp_ColHidden(ss1, 34, True)
    Call Gp_Sp_ColHidden(ss1, 35, True)
    Call Gp_Sp_ColHidden(ss1, 36, True)
    Call Gp_Sp_ColHidden(ss1, 37, True)
End Sub

Private Sub cbo_group_Change()

    If cbo_group.Text <> "" Then
       If cbo_group.Text <> "A" And cbo_group.Text <> "B" And cbo_group.Text <> "C" And cbo_group.Text <> "D" Then
          MsgBox "您输入了不正确的数据！", vbCritical, "错误提示"
       End If
    End If

End Sub

Private Sub cbo_prc_line_Click()

    Dim sHeat_No As String
    Dim sCcm_line As String
    Dim sDynamic As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    If sBef_ccm_line Then Exit Sub
        
    sHeat_No = cbo_heat_no.Text
    sCcm_line = cbo_prc_line.Text
    
    Call Gf_Sp_Cls(Proc_Sc("SC"))
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
    cbo_heat_no.Enabled = True
    sBef_ccm_line = True
    cbo_prc_line.ListIndex = Val(sCcm_line) - 1
    sBef_ccm_line = False
    
    Call HeatNo_ComboAdd(sCcm_line)
    
    If cbo_heat_no.ListCount <> 0 Then
        cbo_heat_no.ListIndex = 0
    End If

    sDynamic = Gf_CodeFind(M_CN1, "SELECT GF_SYSTEM_RUN('SC" & sCcm_line & "') FROM DUAL ")
    
    If sDynamic = "Y" Then
        Call MenuTool_ReSet_Y
    Else
        Call MenuTool_ReSet_N
    End If
    
End Sub

Private Sub cbo_SHIFT_Change()

    If cbo_shift.Text <> "" Then
        If cbo_shift.Text <> "1" And cbo_shift.Text <> "2" And cbo_shift.Text <> "3" Then
           MsgBox "您输入了不正确的数据！", vbCritical, "错误提示"
        End If
    End If

End Sub

Private Sub cmd_down_Click()

    Dim sHeat_No As String
    
    If Trim(cbo_heat_no.Text) = "" Then
        Exit Sub
    End If
    
    sHeat_No = Format(cbo_heat_no.Text - 1, "00000000")
    
    Call Form_Cls
    
    cbo_heat_no.Text = sHeat_No
    
    If cbo_heat_no.Text <> "" Then
        Call Form_Ref
    End If
    
End Sub

Private Sub cmd_up_Click()

    Dim sHeat_No As String
    
    If Trim(cbo_heat_no.Text) = "" Then
        Exit Sub
    End If
    
    sHeat_No = Format(cbo_heat_no.Text + 1, "00000000")
    
    Call Form_Cls
    
    cbo_heat_no.Text = sHeat_No
    
    If cbo_heat_no.Text <> "" Then
        Call Form_Ref
    End If
    
End Sub

Private Sub cmd_work_order_Click()

    Dim sQuery As String
    Dim sMesg As String
    Dim sCcm_line As String
    Dim sDynamic As String
    
    Timer1.Enabled = False
    Option1.Visible = False
    Option2.Visible = False
    QueryYN = True
    
    Call Gp_Sp_ColHidden(ss1, 37, False)
    
    sCcm_line = Gf_CodeFind(M_CN1, "SELECT CCM_PRC_LINE FROM NISCO.EP_CHARGE_IDX WHERE HEAT_MANA_NO = '" & cbo_heat_no.Text & "' ")
    sDynamic = Gf_CodeFind(M_CN1, "SELECT GF_SYSTEM_RUN('SC" & sCcm_line & "') FROM DUAL ")
    
    sQuery = "           SELECT  '',SLAB_NO,CCM_PRC_LINE,NULL,NULL,NULL,NULL,NULL,substr(CCM_CUT_PRE_TME,1,4) || '-' || substr(CCM_CUT_PRE_TME,5,2) || '-' || SUBSTR(CCM_CUT_PRE_TME,7,2) || ' '||"
    sQuery = sQuery & "          SUBSTR(CCM_CUT_PRE_TME,9,2) || ':' || SUBSTR(CCM_CUT_PRE_TME,11,2) || ':' || SUBSTR(CCM_CUT_PRE_TME,13,2),SLAB_THK,SLAB_WID,SLAB_LEN,SLAB_WID_TOP,"
    sQuery = sQuery & "          SLAB_WID_BOT,HCR_FL,NULL,NULL,NULL,0,NULL,NULL,SLAB_WGT,SAMPLE_LENGTH,SAMPLE_CUT,'1' AA,SLAB_LEN_MIN,SLAB_LEN_MAX,NULL,NULL,NULL,NULL,NULL,NULL,LOCK_HEAT_NO,PLAN_NAME,SLAB_IN_PLAN,ORD_FL"
    sQuery = sQuery & "    FROM  NISCO.EP_SLAB_INS  "
    
'    If sDynamic = "Y" Then
'        sQuery = sQuery & "   WHERE  LOCK_HEAT_NO = '" & cbo_heat_no.Text & "'"
'        sQuery = sQuery & "   ORDER  BY PLAN_NAME, SLAB_IN_PLAN "
'    Else
        sQuery = sQuery & "   WHERE  SLAB_NO LIKE '" & cbo_heat_no.Text & "%'"
        sQuery = sQuery & "   ORDER  BY SLAB_NO "
'    End If
    
    sMesg = Gf_Ms_NeceCheck(pControl)
    If sMesg = "OK" Then
    
        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then

            If Gf_Only_Display(M_CN1, Proc_Sc("Sc"), sQuery) Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                
                ss1.OperationMode = OperationModeRead
                
                If sDynamic = "Y" Then
                    Call MenuTool_ReSet_Y
                Else
                    Call MenuTool_ReSet_N
                End If
                
            End If
            
        Else
            sMesg = sMesg + "长度不正确"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
    
    Else
        sMesg = sMesg + "必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
    End If
    
End Sub

Private Sub Form_Activate()
     
    Dim sCcm_line As String
    Dim sDynamic As String
    
    sCcm_line = Gf_CodeFind(M_CN1, "SELECT CCM_PRC_LINE FROM NISCO.EP_CHARGE_IDX WHERE HEAT_MANA_NO = '" & cbo_heat_no.Text & "' ")
    sDynamic = Gf_CodeFind(M_CN1, "SELECT GF_SYSTEM_RUN('SC" & sCcm_line & "') FROM DUAL ")
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    If sDynamic = "Y" Then
        Call MenuTool_ReSet_Y
    Else
        Call MenuTool_ReSet_N
    End If
    
    MDIMain.StatusBar1.Panels(1) = "提示信息："

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    cbo_shift.AddItem "1"
    cbo_shift.AddItem "2"
    cbo_shift.AddItem "3"

    cbo_group.AddItem "A"
    cbo_group.AddItem "B"
    cbo_group.AddItem "C"
    cbo_group.AddItem "D"

    cbo_prc_line.AddItem "1"
    cbo_prc_line.AddItem "2"
    cbo_prc_line.AddItem "3"
    
    Screen.MousePointer = vbHourglass
        
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet_Y
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    cbo_prc_line.ListIndex = 0
    
    If cbo_heat_no <> "" Then
       Call Form_Ref
    End If
    
    txt_emp_cd = sUserID
    txt_emp_cd.ForeColor = &H80000011
    
    If cbo_shift.Text = "1" Or cbo_shift.Text = "2" Or cbo_shift.Text = "3" Then
       Exit Sub
    End If
    
    sQuery_Rt = Gf_CodeFind(M_CN1, "SELECT SHIFT  FROM GP_SHIFTSTD WHERE TRANCD = 'BF21' ")
    If sQuery_Rt = "" Then
       cbo_shift.ListIndex = 0
    Else
       cbo_shift.ListIndex = Val(sQuery_Rt) - 1
    End If
    
    sQuery_Rt = Gf_CodeFind(M_CN1, "SELECT GROUP_CD  FROM GP_SHIFTSTD WHERE TRANCD = 'BF21' ")
    
    Call HeatNo_ComboAdd(Trim(cbo_prc_line.Text))
    
    If cbo_heat_no.ListCount <> 0 Then
        cbo_heat_no.ListIndex = 0
    End If
    
    Select Case sQuery_Rt
        Case "B"
            cbo_group.ListIndex = 1
        Case "C"
            cbo_group.ListIndex = 2
        Case "D"
            cbo_group.ListIndex = 3
        Case Else
            cbo_group.ListIndex = 0
   End Select
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
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

    Dim i As Integer
    Dim sAct_Slab_No As String
    
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    
    With ss1
              
        For i = 1 To .MaxRows
            .Row = i
            
            .Col = 0
            If .Text <> "Input" Then
                .Col = 1
                If .Text = "" Then
                    If .Row = 1 Then
                        sAct_Slab_No = cbo_heat_no.Text & "01"
                    Else
                        sAct_Slab_No = Mid(sAct_Slab_No, 1, 8) & Right("0" & Trim(STR(Val(Mid(sAct_Slab_No, 9, 2)) + 1)), 2)
                    End If
                Else
                    sAct_Slab_No = .Text
                End If
                 
                .Text = sAct_Slab_No
            End If
            
            .Col = 25
            If .Text = "0" Then
                Call Gp_Sp_RowColor(ss1, i, , &HFFC0FF)
            Else
                Call Gp_Sp_RowColor(ss1, i, , &HC0FFC0)
            End If
        Next
              
    End With
    
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet_Y
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
    End If
    
    cbo_prc_line.Text = "1"
    
    cbo_heat_no.Text = ""
    cbo_heat_no.Enabled = True
    cbo_heat_no.SetFocus
    Timer1.Enabled = False
    txt_stlgrd.Enabled = False
    
    txt_emp_cd = sUserID
    txt_emp_cd.ForeColor = &H80000011
    
End Sub

Public Sub Form_Ref()

    Dim i, j As Integer
    Dim sAct_Slab_No As String
    Dim sCcm_line As String
    Dim sDynamic As String
    
    Call Gp_Sp_ColHidden(ss1, 37, True)
    
    cbo_heat_no.Text = Mid(cbo_heat_no.Text, 1, 8)
    QueryYN = False
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("mControl")) Then
    
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, , , False) Then
        
            sAct_Slab_No = cbo_heat_no.Text
            
            With ss1
              
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = 1
                    
                    If .Text = "" Then
                        sAct_Slab_No = Mid(sAct_Slab_No, 1, 8) & Right("0" & Trim(STR(Val(Mid(sAct_Slab_No, 9, 2)) + 1)), 2)
                    Else
                        sAct_Slab_No = .Text
                    End If
                     
                    .Text = sAct_Slab_No
                    
                    .Col = 25
                    If .Text = "0" Then
                        Call Gp_Sp_RowColor(ss1, i, , &HFFC0FF)
                    Else
                        Call Gp_Sp_RowColor(ss1, i, , &HC0FFC0)
                    End If
                Next
              
            End With
           
            Option1.Visible = True
            Option2.Visible = True
            
            If Option1.VALUE = True Then
                Timer1.Enabled = True
            ElseIf Option2.VALUE = True Then
                Timer1.Enabled = False
            End If
         
            ss1.OperationMode = OperationModeNormal
            
            Call Gp_Ms_ControlLock(Mc1("pControl"), True)
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            
            sCcm_line = Gf_CodeFind(M_CN1, "SELECT CCM_PRC_LINE FROM NISCO.EP_CHARGE_IDX WHERE HEAT_MANA_NO = '" & cbo_heat_no.Text & "' ")
            sDynamic = Gf_CodeFind(M_CN1, "SELECT GF_SYSTEM_RUN('SC" & sCcm_line & "') FROM DUAL ")

            If sDynamic = "Y" Then
                'Call Gp_Sp_ColHidden(ss1, 28, True)
                Call MenuTool_ReSet_Y
            Else
                'Call Gp_Sp_ColHidden(ss1, 28, False)
                Call MenuTool_ReSet_N
            End If
            
            txt_stlgrd.Enabled = True
            txt_stlgrd.Locked = True
            txt_stlgrd.ForeColor = &H80000011
            
        End If
        
    End If
    
    If txt_emp_cd = "" Then
        txt_emp_cd = sUserID
        txt_emp_cd.ForeColor = &H80000011
    End If
                
End Sub

Public Sub Form_Pro()

    Dim i, j, iCol, iRow As Integer
    Dim sMesg As String
    Dim fLen, fLenMin, fLenMax As Double
    
    If Gf_Mc_Authority(sAuthority, Mc1, Proc_Sc("Sc")) Then
      
       With ss1
            For iRow = 1 To .MaxRows
                .Row = iRow
                .Col = 12
                fLen = CDbl(IIf(.Text = "", 0, .Text))
                
                .Col = 26
                If .Text <> "" Then
                    fLenMin = CDbl(IIf(.Text = "", 0, .Text))
                End If
                
                .Col = 27
                If .Text <> "" Then
                    fLenMax = CDbl(IIf(.Text = "", 0, .Text))
                End If
                
                .Col = 0
                If .Text = "Update" Or .Text = "Input" Then
                    If fLen <> "" And (fLen < fLenMin Or fLen > fLenMax) Then
                      .Col = 11
                      .BackColor = vbYellow
                       If Gf_MessConfirm("您确定输入了正确的板坯长度值吗？", "W", "") Then
                          .Col = 25
                          .Text = 5
                           If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
                              Call Form_Ref
                           Else
                              Exit Sub
                           End If
                      
                       Else
                           Exit Sub
                       End If
                   
                       Exit For
                    Else
                   
                       If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
                           Call Form_Ref
                       Else
                           Exit Sub
                       End If
                    End If
                ElseIf .Text = "Delete" Then
                    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
                       Call Form_Ref
                    Else
                       Exit Sub
                    End If
                End If
            Next iRow
       End With
    End If
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
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

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub Option1_Click(VALUE As Integer)

    Timer1.Enabled = True
    
End Sub

Private Sub Option2_Click(VALUE As Integer)

    Timer1.Enabled = False
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If QueryYN = False Then
         
         If Col = 9 Then
            ss1.Col = 9
            ss1.Row = Row
            If Row <> 0 Then
            ss1.Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
            End If
         End If
         
         ss1.Col = 0
         ss1.Row = Row
         If ss1.Text = "" Then
            ss1.Text = "Update"
         End If
         
         ss1.Col = 16
         ss1.Row = Row
        If ss1.Text = "" Then
          ss1.Col = 1
          ss1.Row = Row
          ss1.Lock = False
        End If
         
    End If
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim fLen, fLenMin As Double

    If Gf_Sc_Authority(sAuthority, "U") Then Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    
'    With ss1
'      .Row = .ActiveRow
'      If .Col = 11 Then
'            fLen = Val(.Text)
'            .Col = 24
'            fLenMin = Val(.Text)
'            If fLen <> 0 And fLenMin <> 0 Then
'                  If (fLen - fLenMin <= 80) And (fLen - fLenMin >= 0) Then
'                    .Col = 27
'                    .Text = "Y"
'                  Else
'                    .Col = 27
'                    .Text = "N"
'                  End If
'            End If
'      End If
'    End With
'是否定尺在AFH3010P中判断
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub SS1_KeyUp(KeyCode As Integer, Shift As Integer)

    If ss1.ActiveCol = 16 Then
    
        If QueryYN = False Then
        
           With ss1
             .Row = .ActiveRow
             .Col = 0
             If .Text = "Update" Or .Text = "Input" Or .Text = "Delete" Then
                .Row = .ActiveRow
                .Col = 16
                 If .Text <> "H" And .Text <> "S" And .Text <> "C" And .Text <> "" Then
                     MsgBox "您输入了不正确的数据！", vbCritical, "错误提示"
                     ss1.SetSelection 15, .ActiveRow, 15, .ActiveRow
                     Exit Sub
                 End If
             End If
                 
           End With
        
        End If
        
    End If
    
End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub Timer1_Timer()

    Dim i As Integer
    
    With ss1
        .Col = 28
        For i = 1 To .MaxRows
            .Row = i
            If .Text = "1" Then
               Exit For
            End If
        Next
        
        If i = .MaxRows Then
           Timer1.Enabled = False
        Else
           Timer1.Enabled = True
           Call Form_Ref
        End If
        
    End With
    
End Sub

Private Sub HeatNo_ComboAdd(sPrcLine As String)

On Error GoTo HeatNo_ComboAdd_Error
    
    Dim AdoRs  As ADODB.Recordset
    Dim sQuery As String
    Dim sDynamic As String
    
    'Db Connection Check
    If M_CN1 Is Nothing Then
        If GF_DbConnect = False Then Exit Sub
    End If
    
    cbo_heat_no.Clear
    
    sDynamic = Gf_CodeFind(M_CN1, "SELECT GF_SYSTEM_RUN('SC" & sPrcLine & "') FROM DUAL ")
     
    sQuery = "          SELECT  C.HEAT_MANA_NO "
    sQuery = sQuery & "   FROM  (SELECT A.HEAT_MANA_NO "
    sQuery = sQuery & "            FROM EP_CHARGE_IDX A "
    sQuery = sQuery & "           WHERE A.PRC_STS         IN  ('B')  "
    sQuery = sQuery & "             AND A.CCM_PRC_LINE    =   '" & Trim(sPrcLine) & "' "
    sQuery = sQuery & "             AND A.BOF_RSLT        =   'Y' "
    
    If sDynamic = "Y" Then
        sQuery = sQuery & "         AND A.L2_CCM_LOCK     =   'Y' "
        sQuery = sQuery & "       ORDER BY A.PLAN_NAME, PLAN_NAME_SEQ) C "
    
    Else
        sQuery = sQuery & "       ORDER BY A.HEAT_MANA_NO) C "
    
    End If
    
    sQuery = sQuery & "   WHERE ROWNUM <= 15 "
    
    Set AdoRs = New ADODB.Recordset
     
    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
     
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
                cbo_heat_no.AddItem AdoRs.Fields(0)
            End If
            
            AdoRs.MoveNext
            
        Wend
        
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Sub

HeatNo_ComboAdd_Error:

    Set AdoRs = Nothing

End Sub

Private Sub MenuTool_ReSet_Y()

    With MDIMain.MenuTool
        .Buttons(5).Enabled = False                  'Delete
        .Buttons(7).Enabled = True                   'Row Insert
        .Buttons(8).Enabled = True                   'Row Delete
        .Buttons(11).Enabled = True                 'Spread Copy
        .Buttons(12).Enabled = True                 'Spread Paste
    End With

End Sub

Private Sub MenuTool_ReSet_N()

    With MDIMain.MenuTool
        .Buttons(5).Enabled = True                   'Delete
        .Buttons(7).Enabled = True                   'Row Insert
        .Buttons(8).Enabled = True                   'Row Delete
        .Buttons(11).Enabled = True                  'Spread Copy
        .Buttons(12).Enabled = True                  'Spread Paste
    End With

End Sub
