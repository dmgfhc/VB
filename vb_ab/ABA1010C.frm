VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ABA1010C 
   Caption         =   "接受订单及录入_ABA1010C"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
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
      Left            =   9600
      MaxLength       =   40
      TabIndex        =   45
      Tag             =   "产品"
      Top             =   600
      Width           =   1950
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
      Left            =   9135
      MaxLength       =   2
      TabIndex        =   44
      Tag             =   "产品"
      Top             =   600
      Width           =   465
   End
   Begin VB.TextBox txt_reg_time 
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
      Left            =   9135
      MaxLength       =   8
      TabIndex        =   43
      Tag             =   "Emp_Time"
      Top             =   4230
      Width           =   1185
   End
   Begin VB.TextBox txt_emp_id 
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
      Left            =   10800
      MaxLength       =   7
      TabIndex        =   42
      Tag             =   "Emp_ID"
      Top             =   3870
      Visible         =   0   'False
      Width           =   1395
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
      Left            =   1695
      MaxLength       =   11
      TabIndex        =   41
      Tag             =   "订单号"
      Top             =   120
      Width           =   1530
   End
   Begin VB.TextBox txt_dept_cd 
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
      Left            =   9135
      MaxLength       =   3
      TabIndex        =   39
      Tag             =   "部门"
      Top             =   1320
      Width           =   465
   End
   Begin VB.TextBox txt_dept_cd_name 
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
      Left            =   9600
      MaxLength       =   40
      TabIndex        =   38
      Tag             =   "Dept_Cd_Name"
      Top             =   1320
      Width           =   1950
   End
   Begin VB.TextBox txt_mod_fl 
      Enabled         =   0   'False
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
      Left            =   4890
      MaxLength       =   1
      TabIndex        =   37
      Tag             =   "mod_fl"
      Top             =   3495
      Width           =   465
   End
   Begin VB.TextBox txt_mod_fl_name 
      Enabled         =   0   'False
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
      Left            =   5355
      MaxLength       =   40
      TabIndex        =   36
      Tag             =   "Mod_Fl_name"
      Top             =   3495
      Width           =   1950
   End
   Begin VB.TextBox txt_mod_time 
      Enabled         =   0   'False
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
      Left            =   4890
      MaxLength       =   8
      TabIndex        =   35
      Tag             =   "Mod_Time"
      Top             =   4215
      Width           =   1185
   End
   Begin VB.TextBox txt_ord_st 
      Enabled         =   0   'False
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
      Left            =   1695
      MaxLength       =   1
      TabIndex        =   34
      Tag             =   "Ord_St"
      Top             =   1320
      Width           =   465
   End
   Begin VB.TextBox txt_ord_st_name 
      Enabled         =   0   'False
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
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   33
      Tag             =   "Ord_St_name"
      Top             =   1320
      Width           =   1950
   End
   Begin VB.TextBox txt_can_fl 
      Enabled         =   0   'False
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
      Left            =   12585
      MaxLength       =   1
      TabIndex        =   32
      Tag             =   "Can_Fl"
      Top             =   3495
      Width           =   465
   End
   Begin VB.TextBox txt_can_fl_name 
      Enabled         =   0   'False
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
      Left            =   13050
      MaxLength       =   40
      TabIndex        =   31
      Tag             =   "Can_Fl_Name"
      Top             =   3495
      Width           =   1980
   End
   Begin VB.TextBox txt_can_time 
      Enabled         =   0   'False
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
      Left            =   12585
      MaxLength       =   8
      TabIndex        =   30
      Tag             =   "Can_Time"
      Top             =   4215
      Width           =   1185
   End
   Begin VB.TextBox txt_ord_knd 
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
      Left            =   1695
      MaxLength       =   1
      TabIndex        =   29
      Tag             =   "订单种类"
      Top             =   960
      Width           =   465
   End
   Begin VB.TextBox txt_ord_knd_name 
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
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   28
      Tag             =   "ord_knd_name"
      Top             =   960
      Width           =   1950
   End
   Begin VB.TextBox txt_prod_dgr 
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
      Left            =   9135
      MaxLength       =   1
      TabIndex        =   27
      Tag             =   "产品等级"
      Top             =   960
      Width           =   465
   End
   Begin VB.TextBox txt_prod_dgr_name 
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
      Left            =   9600
      MaxLength       =   40
      TabIndex        =   26
      Tag             =   "产品等级"
      Top             =   960
      Width           =   1950
   End
   Begin VB.TextBox txt_sale_way 
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
      Left            =   9135
      MaxLength       =   2
      TabIndex        =   25
      Tag             =   "销售公司"
      Top             =   2040
      Width           =   465
   End
   Begin VB.TextBox txt_sale_way_name 
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
      Left            =   9600
      MaxLength       =   40
      TabIndex        =   24
      Tag             =   "sale_way_name"
      Top             =   2040
      Width           =   1950
   End
   Begin VB.TextBox txt_dome_fl 
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
      Left            =   1695
      MaxLength       =   1
      TabIndex        =   23
      Tag             =   "订单分类"
      Top             =   2775
      Width           =   465
   End
   Begin VB.TextBox txt_dome_fl_name 
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
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   22
      Tag             =   "dome_fl_name"
      Top             =   2775
      Width           =   1950
   End
   Begin VB.TextBox txt_cust_cd 
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
      Left            =   1695
      MaxLength       =   6
      TabIndex        =   21
      Tag             =   "客户"
      Top             =   600
      Width           =   870
   End
   Begin VB.TextBox txt_pono 
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
      Left            =   1695
      MaxLength       =   30
      TabIndex        =   20
      Tag             =   "PoNo"
      Top             =   3135
      Width           =   4980
   End
   Begin VB.TextBox txt_cont_end_time 
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
      Left            =   1695
      MaxLength       =   8
      TabIndex        =   19
      Tag             =   "Cont_End_TIme"
      Top             =   4230
      Width           =   1185
   End
   Begin VB.TextBox txt_cust_cd_name 
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
      Left            =   2565
      MaxLength       =   40
      TabIndex        =   16
      Tag             =   "客户"
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox txt_emp_nm 
      Height          =   315
      Left            =   9135
      TabIndex        =   15
      Top             =   3495
      Width           =   1440
   End
   Begin VB.TextBox TXT_XSXZ_NM 
      Enabled         =   0   'False
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
      Left            =   9600
      MaxLength       =   40
      TabIndex        =   14
      Tag             =   "Can_Fl_Name"
      Top             =   1680
      Width           =   1950
   End
   Begin VB.TextBox TXT_XSXZ 
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
      Left            =   9135
      MaxLength       =   1
      TabIndex        =   13
      Tag             =   "销售性质"
      Top             =   1680
      Width           =   465
   End
   Begin VB.TextBox TXT_JGXZ_NM 
      Enabled         =   0   'False
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
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   12
      Tag             =   "Mod_Fl_name"
      Top             =   1695
      Width           =   1950
   End
   Begin VB.TextBox TXT_JGXZ 
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
      Left            =   1695
      MaxLength       =   1
      TabIndex        =   11
      Tag             =   "价格性质"
      Top             =   1695
      Width           =   465
   End
   Begin VB.TextBox TXT_THFS_NM 
      Enabled         =   0   'False
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
      Left            =   9600
      MaxLength       =   40
      TabIndex        =   10
      Tag             =   "Can_Fl_Name"
      Top             =   2400
      Width           =   1950
   End
   Begin VB.TextBox TXT_THFS 
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
      Left            =   9135
      MaxLength       =   1
      TabIndex        =   9
      Tag             =   "提货方式"
      Top             =   2400
      Width           =   465
   End
   Begin VB.TextBox TXT_FPLX_NM 
      Enabled         =   0   'False
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
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   8
      Tag             =   "Mod_Fl_name"
      Top             =   2415
      Width           =   1950
   End
   Begin VB.TextBox TXT_FPLX 
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
      Left            =   1695
      MaxLength       =   1
      TabIndex        =   7
      Tag             =   "发票类型"
      Top             =   2415
      Width           =   465
   End
   Begin VB.TextBox TXT_JSFS_NM 
      Enabled         =   0   'False
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
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   6
      Tag             =   "Mod_Fl_name"
      Top             =   2055
      Width           =   1950
   End
   Begin VB.TextBox TXT_JSFS 
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
      Left            =   1695
      MaxLength       =   2
      TabIndex        =   5
      Tag             =   "结算方式"
      Top             =   2055
      Width           =   465
   End
   Begin VB.TextBox TXT_YSFS_NM 
      Enabled         =   0   'False
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
      Left            =   9600
      MaxLength       =   40
      TabIndex        =   4
      Tag             =   "Mod_Fl_name"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox TXT_YSFS 
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
      Left            =   9135
      MaxLength       =   1
      TabIndex        =   3
      Tag             =   "运输方式"
      Top             =   2760
      Width           =   465
   End
   Begin VB.TextBox TXT_DEST_NM 
      Enabled         =   0   'False
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
      Left            =   9960
      MaxLength       =   40
      TabIndex        =   2
      Tag             =   "Mod_Fl_name"
      Top             =   3135
      Width           =   5070
   End
   Begin VB.TextBox TXT_DEST 
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
      Left            =   9135
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "目的地"
      Top             =   3135
      Width           =   825
   End
   Begin VB.TextBox txt_ship_no 
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
      Left            =   12585
      MaxLength       =   30
      TabIndex        =   0
      Tag             =   "ship_No"
      Top             =   2745
      Width           =   2445
   End
   Begin Threed.SSCommand SCmd_COPY 
      Height          =   375
      Left            =   12525
      TabIndex        =   17
      Top             =   75
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "订单复制"
   End
   Begin InDate.ULabel ULabel26 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1500
      _ExtentX        =   2646
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
      ForeColor       =   16711680
   End
   Begin InDate.UDate dtp_reg_date 
      Height          =   315
      Left            =   9135
      TabIndex        =   18
      Tag             =   "Emp_Date"
      Top             =   3855
      Width           =   1455
      _ExtentX        =   2566
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
   Begin FPSpread.vaSpread ss1 
      Height          =   4575
      Left            =   120
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4710
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   8070
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
      MaxCols         =   45
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ABA1010C.frx":0000
   End
   Begin CSTextLibCtl.sidbEdit sdb_num_item 
      Height          =   315
      Left            =   6825
      TabIndex        =   46
      Tag             =   "Num_Item"
      Top             =   120
      Width           =   465
      _Version        =   262145
      _ExtentX        =   820
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
      Enabled         =   0   'False
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
      NumIntDigits    =   2
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_tot_wgt 
      Height          =   315
      Left            =   9135
      TabIndex        =   47
      Tag             =   "Tot_Wgt"
      Top             =   120
      Width           =   1335
      _Version        =   262145
      _ExtentX        =   2355
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.UDate dtp_mod_date 
      Height          =   315
      Left            =   4890
      TabIndex        =   48
      Top             =   3855
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Enabled         =   0   'False
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
   Begin InDate.UDate dtp_can_date 
      Height          =   315
      Left            =   12585
      TabIndex        =   49
      Tag             =   "Can_Date"
      Top             =   3855
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Enabled         =   0   'False
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
   Begin InDate.UDate dtp_cont_date 
      Height          =   315
      Left            =   1695
      TabIndex        =   50
      Tag             =   "Cont_Date"
      Top             =   3495
      Width           =   1455
      _ExtentX        =   2566
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
   Begin InDate.UDate dtp_cont_end_date 
      Height          =   315
      Left            =   1695
      TabIndex        =   51
      Tag             =   "Cont_End_Date"
      Top             =   3840
      Width           =   1455
      _ExtentX        =   2566
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
   Begin InDate.ULabel ULabel27 
      Height          =   315
      Left            =   120
      Top             =   600
      Width           =   1500
      _ExtentX        =   2646
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
   Begin InDate.ULabel ULabel28 
      Height          =   315
      Left            =   120
      Top             =   960
      Width           =   1500
      _ExtentX        =   2646
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
   Begin InDate.ULabel ULabel29 
      Height          =   315
      Left            =   120
      Top             =   1320
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "订单状态"
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
   Begin InDate.ULabel ULabel30 
      Height          =   315
      Left            =   3315
      Top             =   3495
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "修改分类"
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
   Begin InDate.ULabel ULabel31 
      Height          =   315
      Left            =   3315
      Top             =   3855
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "修改日期"
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
   Begin InDate.ULabel ULabel32 
      Height          =   315
      Left            =   3315
      Top             =   4215
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "修改时间"
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
   Begin InDate.ULabel ULabel33 
      Height          =   315
      Left            =   7545
      Top             =   2055
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "销售公司"
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
   Begin InDate.ULabel ULabel34 
      Height          =   315
      Left            =   120
      Top             =   3135
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "客户合同号"
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
   Begin InDate.ULabel ULabel35 
      Height          =   315
      Left            =   120
      Top             =   3480
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "合同签订日期"
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
   Begin InDate.ULabel ULabel36 
      Height          =   315
      Left            =   120
      Top             =   3840
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "合同截止日期"
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
      Left            =   7545
      Top             =   3495
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "输入人员"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   5250
      Top             =   120
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "序列数"
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
      Left            =   7545
      Top             =   120
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "重量"
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
      Left            =   7545
      Top             =   585
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "产品"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   7545
      Top             =   1320
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "部门"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   7545
      Top             =   960
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "产品等级"
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
      Left            =   10995
      Top             =   3495
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "取消分类"
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
      Left            =   10995
      Top             =   3855
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "取消日期"
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
      Left            =   120
      Top             =   2775
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "订单分类"
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
      Left            =   120
      Top             =   4230
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "合同截止时间"
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
      Left            =   7545
      Top             =   3855
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "输入日期"
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
      Left            =   7545
      Top             =   4230
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "输入时间"
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
      Left            =   10995
      Top             =   4215
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "取消时间"
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
      Left            =   120
      Top             =   1695
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "价格性质"
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
      Left            =   120
      Top             =   2070
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "结算方式"
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
      Left            =   7545
      Top             =   1680
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "销售性质"
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
      Left            =   7545
      Top             =   3135
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "目的地"
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   120
      Top             =   2430
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "发票类型"
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
   Begin InDate.ULabel ULabel19 
      Height          =   315
      Left            =   7545
      Top             =   2415
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "提货方式"
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
      Left            =   7545
      Top             =   2775
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "运输方式"
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
   Begin InDate.ULabel ULabel01 
      Height          =   315
      Index           =   35
      Left            =   10995
      Top             =   2760
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "船号"
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
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   15
      X2              =   15015
      Y1              =   510
      Y2              =   510
   End
End
Attribute VB_Name = "ABA1010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name
'-- Sub_System Name
'-- Program Name      Order
'-- Program ID        ABA1010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.6.2
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

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Hsheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
              Call Gp_Ms_Collection(txt_ord_no, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_emp_id, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_emp_nm, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(dtp_reg_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_reg_time, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_prod_cd, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_prod_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_dept_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_dept_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_mod_fl, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_mod_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(dtp_mod_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_mod_time, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_ord_st, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ord_st_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_can_fl, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_can_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(dtp_can_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_can_time, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ord_knd, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ord_knd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_prod_dgr, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_prod_dgr_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_sale_way, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_sale_way_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_dome_fl, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_dome_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_cust_cd, " ", "n", "m", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_cust_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_num_item, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_tot_wgt, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_pono, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(dtp_cont_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(dtp_cont_end_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_cont_end_time, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              
                Call Gp_Ms_Collection(TXT_JGXZ, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_JGXZ_NM, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(TXT_JSFS, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_JSFS_NM, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(TXT_FPLX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_FPLX_NM, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(TXT_XSXZ, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_XSXZ_NM, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(TXT_THFS, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_THFS_NM, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(TXT_YSFS, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_YSFS_NM, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(TXT_DEST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_DEST_NM, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ship_no, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    'MASTER Collection
    Mc1.Add Item:="ABA1010C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="ABA1010C.P_REFER", Key:="P-R"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ABA1010C.P_SREFER", Key:="P-R"
'    Sc1.Add Item:="ABA1010C.P_SMODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    
    MDIMain.MenuTool.Buttons(11).ButtonMenus(1).Enabled = False 'All Copy
    MDIMain.MenuTool.Buttons(11).ButtonMenus(2).Enabled = True  'Master Copy
    MDIMain.MenuTool.Buttons(11).ButtonMenus(3).Enabled = False 'Spread Copy

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "B-System.INI", Me.Name)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    
    MDIMain.MenuTool.Buttons(11).ButtonMenus(1).Enabled = False 'All Copy
    MDIMain.MenuTool.Buttons(11).ButtonMenus(2).Enabled = True  'Master Copy
    MDIMain.MenuTool.Buttons(11).ButtonMenus(3).Enabled = False 'Spread Copy

    Screen.MousePointer = vbDefault
    
    dtp_cont_end_date.RawData = ""
    dtp_cont_date.RawData = ""
    dtp_mod_date.RawData = ""
    dtp_can_date.RawData = ""
    dtp_reg_date.Text = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "B-System.INI", Me.Name)
    
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

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
      
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Active_Spread.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub


Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        txt_cust_cd.Enabled = True
        txt_prod_cd.Enabled = True
        txt_prod_dgr.Enabled = True
        txt_ord_knd.Enabled = True
        txt_dept_cd.Enabled = True
        txt_sale_way.Enabled = True
        txt_pono.Enabled = True
        dtp_cont_date.Enabled = True
        txt_ship_no.Enabled = True
        TXT_JGXZ.Enabled = True
        TXT_JSFS.Enabled = True
        TXT_FPLX.Enabled = True
        TXT_XSXZ.Enabled = True
        TXT_THFS.Enabled = True
        TXT_YSFS.Enabled = True
        TXT_DEST.Enabled = True
        
        pControl(1).SetFocus
        MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
        MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
        MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
        
        MDIMain.MenuTool.Buttons(11).ButtonMenus(1).Enabled = False 'All Copy
        MDIMain.MenuTool.Buttons(11).ButtonMenus(2).Enabled = True  'Master Copy
        MDIMain.MenuTool.Buttons(11).ButtonMenus(3).Enabled = False 'Spread Copy
    End If
    
    dtp_cont_end_date.RawData = ""
    dtp_cont_date.RawData = ""
    dtp_mod_date.RawData = ""
    dtp_can_date.RawData = ""
    dtp_reg_date.RawData = ""
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

    If Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl")) Then
    
        Call Gf_Sp_Display(M_CN1, Proc_Sc("Sc").Item("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", Mc1("pControl")), Proc_Sc("Sc").Item("pColumn"), False)
        Call Gp_Sp_EvenRowBackcolor(Proc_Sc("Sc").Item("Spread"))
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        txt_cust_cd.Enabled = False
        txt_cust_cd_name.Enabled = False
        txt_prod_cd.Enabled = False
        txt_prod_dgr.Enabled = False
        txt_ord_knd.Enabled = False
        txt_dept_cd.Enabled = False
        txt_sale_way.Enabled = False
        txt_pono.Enabled = False
        dtp_cont_date.Enabled = False
        txt_ship_no.Enabled = False
        TXT_JGXZ.Enabled = False
        TXT_JSFS.Enabled = False
        TXT_FPLX.Enabled = False
        TXT_XSXZ.Enabled = False
        TXT_THFS.Enabled = False
        TXT_YSFS.Enabled = False
        TXT_DEST.Enabled = False
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
        MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
        MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
        
        MDIMain.MenuTool.Buttons(11).ButtonMenus(1).Enabled = False 'All Copy
        MDIMain.MenuTool.Buttons(11).ButtonMenus(2).Enabled = True  'Master Copy
        MDIMain.MenuTool.Buttons(11).ButtonMenus(3).Enabled = False 'Spread Copy
        
    End If
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()
    
    Dim sMesg As String
    Dim sQuery As String
    
    If txt_ord_knd.Text <> "P" Then
        If TXT_JGXZ.Text = "" Then
           Call Gp_MsgBoxDisplay("请输入价格性质")
           Exit Sub
        ElseIf TXT_JSFS.Text = "" Then
           Call Gp_MsgBoxDisplay("请输入结算方式")
           Exit Sub
        ElseIf TXT_FPLX.Text = "" Then
           Call Gp_MsgBoxDisplay("请输入发票类型")
           Exit Sub
        ElseIf TXT_XSXZ.Text = "" Then
           Call Gp_MsgBoxDisplay("请输入销售性质")
           Exit Sub
        ElseIf TXT_THFS.Text = "" Then
           Call Gp_MsgBoxDisplay("请输入退货方式")
           Exit Sub
        ElseIf TXT_YSFS.Text = "" Then
           Call Gp_MsgBoxDisplay("请输入运输方式")
           Exit Sub
        ElseIf TXT_DEST.Text = "" Then
           Call Gp_MsgBoxDisplay("请输入目的地")
           Exit Sub
        End If
    End If
    
    If txt_ord_knd.Text = "P" And txt_cust_cd.Text <> "NG0001" Then
       txt_cust_cd.Text = ""
       txt_cust_cd_name.Text = ""
    End If

    If txt_sale_way.Text <> "GD" And txt_sale_way.Text <> "GE" And txt_sale_way.Text <> "GF" And txt_sale_way.Text <> "GH" And txt_sale_way.Text <> "GO" Then
       Call Gp_MsgBoxDisplay("该销售公司已禁用，请重选选择销售公司")
       Exit Sub
    End If
    
    If Gf_Mc_Authority(sAuthority, Mc1, Proc_Sc("Sc")) Then
        txt_emp_id.Text = sUserID
        
        If Mc1.Item("pControl")(1).Enabled Then
            'Order_No Make
            sQuery = "{call ABA1010C.P_ORD_NO ( '" + txt_cust_cd.Text + "' )}"
            txt_ord_no.Text = Gf_CodeFind(M_CN1, sQuery)
        Else
           If txt_ord_st.Text <> "2" Then
                sMesg = "不能修改该状态的订单"
                Call Gp_MsgBoxDisplay(sMesg)
                Exit Sub
            End If
        End If
        
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
            Call Gf_Sp_Display(M_CN1, Proc_Sc("Sc").Item("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", Mc1("pControl")), Proc_Sc("Sc").Item("pColumn"), False)
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            txt_cust_cd.Enabled = False
            txt_prod_cd.Enabled = False
            txt_prod_dgr.Enabled = False
            
            MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
            MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
            MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
            
            MDIMain.MenuTool.Buttons(11).ButtonMenus(1).Enabled = False 'All Copy
            MDIMain.MenuTool.Buttons(11).ButtonMenus(2).Enabled = True  'Master Copy
            MDIMain.MenuTool.Buttons(11).ButtonMenus(3).Enabled = False 'Spread Copy
        End If
            
    End If
                
'     Call txt_ord_st_KeyUp(0, 0)

    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))

End Sub

Public Sub Form_Del()

    Dim sMesg As String
    
    If txt_ord_st.Text <> "2" Then
        sMesg = "不能删除该状态的订单"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
        
    If Not Gf_Ms_AllDel(M_CN1, Proc_Sc("Sc"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
        txt_cust_cd.Enabled = True
        txt_prod_cd.Enabled = True
        txt_prod_dgr.Enabled = True
        
        MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
        MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
        MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
        
        MDIMain.MenuTool.Buttons(11).ButtonMenus(1).Enabled = False 'All Copy
        MDIMain.MenuTool.Buttons(11).ButtonMenus(2).Enabled = True  'Master Copy
        MDIMain.MenuTool.Buttons(11).ButtonMenus(3).Enabled = False 'Spread Copy
    End If
    
End Sub

Public Sub Form_Cpy()
    
    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Form_Pst()
    
    If Gf_Ms_FormPaste(Mc1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        txt_cust_cd.Enabled = True
        txt_prod_cd.Enabled = True
        txt_prod_dgr.Enabled = True
        txt_ord_no.Text = ""
        
        MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
        MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
        MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
        
        MDIMain.MenuTool.Buttons(11).ButtonMenus(1).Enabled = False 'All Copy
        MDIMain.MenuTool.Buttons(11).ButtonMenus(2).Enabled = True  'Master Copy
        MDIMain.MenuTool.Buttons(11).ButtonMenus(3).Enabled = False 'Spread Copy
    End If
    
End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        txt_cust_cd.Enabled = True
        txt_prod_cd.Enabled = True
        txt_prod_dgr.Enabled = True
        txt_ord_no.Text = ""
    End If
    
End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("Sc"))

End Sub

Private Sub SCmd_COPY_Click()

 Dim iRow As Integer
 
 If txt_ord_no.Enabled = True Then
 
    Exit Sub
    
 End If
 
 Load OrderCopy
 
 OrderCopy.ord_no_fr.Caption = ABA1010C.txt_ord_no.Text
 OrderCopy.txt_emp_id.Text = sUserID
 If ABA1010C.txt_ord_knd.Text = "P" Then
    OrderCopy.txt_cust_cd.Text = ABA1010C.txt_cust_cd.Text
    OrderCopy.txt_cust_cd_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
    OrderCopy.txt_cust_cd.Enabled = False
    OrderCopy.txt_cust_cd_name.Enabled = False
 End If
 
 OrderCopy.ss1.MaxRows = ss1.MaxRows
 
 For iRow = 1 To sdb_num_item.Value
     OrderCopy.ss1.Col = 2
     OrderCopy.ss1.Row = iRow
     If iRow < 10 Then
        OrderCopy.ss1.Text = "0" + LTrim(str(iRow))
     Else
         OrderCopy.ss1.Text = iRow
     End If
 Next iRow
 
 OrderCopy.Show 1

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

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    'Spread --> Control Value Move
    
    If Mc1("pControl").Item(1).Enabled Then Exit Sub
    
    With ss1

        If Row <> 0 Then
        
            Load ABA1020C
    
            .Row = Row
    
            .Col = 1: ABA1020C.txt_ord_no.Text = .Text
            .Col = 2: ABA1020C.txt_ORD_ITEM.Text = .Text
            
        End If

        ABA1020C.Dis_sw = True
        ABA1020C.Show 1
        
    End With

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)

End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
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

Private Sub txt_cust_cd_Change()

     If txt_cust_cd.Text = "NG0001" Then
       txt_ord_knd.Text = "P"
       txt_ord_knd.Enabled = False
       txt_ord_knd_name.Text = Gf_ComnNameFind(M_CN1, "B0009", Trim(txt_ord_knd.Text), 2)
    Else
       If txt_ord_knd.Text = "P" Then
          txt_ord_knd.Text = ""
          txt_ord_knd_name = ""
          txt_ord_knd.Enabled = True
       End If
    End If
    
    
    'Dim sQuery As String
    'Dim AdoRs As adodb.Recordset
    'If Len(txt_cust_cd.Text) <> 0 Then
    '    sQuery = "select dome_fl from bp_cust_cd  where cust_cd='" + txt_cust_cd.Text + "'"
    '    Set AdoRs = New adodb.Recordset
    '    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    '
    '        If VarType(AdoRs.Fields("dome_fl").Value) = vbNull Then
    '           txt_dome_fl.Text = ""
    '           txt_dome_fl_name.Text = ""
    '        Else
    '           txt_dome_fl.Text = AdoRs.Fields("dome_fl").Value
    '           Call txt_dome_fl_KeyUp(0, 0)
    '        End If
    '   AdoRs.Close
    '   Set AdoRs = Nothing
    'End If
    
End Sub

Private Sub txt_cust_cd_DblClick()

Call txt_cust_cd_KeyUp(vbKeyF4, 0)
End Sub


Private Sub txt_cust_cd_LostFocus()

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    If Len(txt_cust_cd.Text) <> 0 Then
    
        sQuery = "select dome_fl from bp_cust_cd  where cust_cd='" + txt_cust_cd.Text + "'"
        Set AdoRs = New ADODB.Recordset
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
        If Not AdoRs.BOF And Not AdoRs.EOF Then
        
            If VarType(AdoRs.Fields("dome_fl").Value) = vbNull Then
    '         If AdoRs.BOF = True Or AdoRs.EOF = True Then
               txt_dome_fl.Text = ""
               txt_dome_fl_name.Text = ""
            Else
               txt_dome_fl.Text = AdoRs.Fields("dome_fl").Value
               Call txt_dome_fl_KeyUp(0, 0)
            End If
            AdoRs.Close
            Set AdoRs = Nothing
            
       Else
       
            Call Gp_MsgBoxDisplay("客户代码不存在.......")
            
            txt_cust_cd.Text = ""
            txt_dome_fl.Text = ""
            If txt_cust_cd.Enabled = True Then
                txt_cust_cd.SetFocus
            End If
            
       End If
       
    End If
    
End Sub


Private Sub txt_dept_cd_DblClick()
Call txt_dept_cd_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_dept_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Z0002"
        DD.rControl.Add Item:=txt_dept_cd
        DD.rControl.Add Item:=txt_dept_cd_name
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(txt_dept_cd)) = txt_dept_cd.MaxLength Then
        txt_dept_cd_name.Text = Gf_ComnNameFind(M_CN1, "Z0002", Trim(txt_dept_cd.Text), 2)
    Else
        txt_dept_cd_name.Text = ""
    End If
    
End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_cust_cd
        DD.rControl.Add Item:=txt_cust_cd_name

        DD.nameType = "1"

        Call Gf_Customer_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_cust_cd)) = txt_cust_cd.MaxLength Then
        txt_cust_cd_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
    Else
        txt_cust_cd_name.Text = ""
    End If

End Sub

Private Sub TXT_DEST_DblClick()
Call txt_dest_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_dome_fl_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0002"
        DD.rControl.Add Item:=txt_dome_fl
        DD.rControl.Add Item:=txt_dome_fl_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_dome_fl)) = txt_dome_fl.MaxLength Then
        txt_dome_fl_name.Text = Gf_ComnNameFind(M_CN1, "B0002", Trim(txt_dome_fl.Text), 2)
    Else
        txt_dome_fl_name.Text = ""
    End If

End Sub

Private Sub TXT_FPLX_DblClick()
Call TXT_FPLX_KeyUp(vbKeyF4, 0)

End Sub

Private Sub TXT_JGXZ_DblClick()
Call TXT_JGXZ_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_JSFS_DblClick()
Call TXT_JSFS_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_mod_fl_DblClick()
Call txt_mod_fl_KeyUp(vbKeyF4, 0)

End Sub


Private Sub txt_mod_fl_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0006"
        DD.rControl.Add Item:=txt_mod_fl
        DD.rControl.Add Item:=txt_mod_fl_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_mod_fl)) = txt_mod_fl.MaxLength Then
        txt_mod_fl_name.Text = Gf_ComnNameFind(M_CN1, "B0006", Trim(txt_mod_fl.Text), 2)
    Else
        txt_mod_fl_name.Text = ""
    End If

End Sub


Private Sub txt_ord_knd_Change()

    If txt_ord_knd.Text = "A" Or txt_ord_knd.Text = "P" Or txt_ord_knd.Text = "T" Then
    
       txt_prod_dgr.Text = "1"
       Call txt_prod_dgr_KeyUp(0, 0)
       txt_prod_dgr.Locked = True
       
       If txt_ord_knd.Text = "P" Then
          txt_cust_cd.Text = "NG0001"
          txt_cust_cd_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
       End If
       
    Else
    
       If txt_prod_dgr.Locked = True Then
          txt_prod_dgr.Text = ""
          txt_prod_dgr_name.Text = ""
       End If
       txt_prod_dgr.Locked = False
       
    End If
    
       
    If txt_ord_knd.Text = "P" Then
       TXT_JGXZ.Enabled = False
       TXT_JSFS.Enabled = False
       TXT_FPLX.Enabled = False
       TXT_XSXZ.Enabled = False
       TXT_THFS.Enabled = False
       TXT_YSFS.Enabled = False
       TXT_DEST.Enabled = False
       TXT_JGXZ.Text = ""
       TXT_JSFS.Text = ""
       TXT_FPLX.Text = ""
       TXT_XSXZ.Text = ""
       TXT_THFS.Text = ""
       TXT_YSFS.Text = ""
       TXT_DEST.Text = ""
       TXT_JGXZ_NM.Text = ""
       TXT_JSFS_NM.Text = ""
       TXT_FPLX_NM.Text = ""
       TXT_XSXZ_NM.Text = ""
       TXT_THFS_NM.Text = ""
       TXT_YSFS_NM.Text = ""
       TXT_DEST_NM.Text = ""
       TXT_JGXZ.BackColor = &H80000005
       TXT_JSFS.BackColor = &H80000005
       TXT_FPLX.BackColor = &H80000005
       TXT_XSXZ.BackColor = &H80000005
       TXT_THFS.BackColor = &H80000005
       TXT_YSFS.BackColor = &H80000005
       TXT_DEST.BackColor = &H80000005
    Else
       TXT_JGXZ.Enabled = True
       TXT_JSFS.Enabled = True
       TXT_FPLX.Enabled = True
       TXT_XSXZ.Enabled = True
       TXT_THFS.Enabled = True
       TXT_YSFS.Enabled = True
       TXT_DEST.Enabled = True
       TXT_JGXZ.BackColor = &HC0FFFF
       TXT_JSFS.BackColor = &HC0FFFF
       TXT_FPLX.BackColor = &HC0FFFF
       TXT_XSXZ.BackColor = &HC0FFFF
       TXT_THFS.BackColor = &HC0FFFF
       TXT_YSFS.BackColor = &HC0FFFF
       TXT_DEST.BackColor = &HC0FFFF
       TXT_JGXZ.Text = "1"
       Call TXT_JGXZ_KeyUp(0, 0)
       TXT_FPLX.Text = "1"
       Call TXT_FPLX_KeyUp(0, 0)
    
    End If
    
End Sub

Private Sub txt_ord_knd_DblClick()
Call txt_ord_knd_KeyUp(vbKeyF4, 0)
End Sub


Private Sub txt_ord_no_DblClick()
Call txt_ord_no_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_ord_st_DblClick()
Call txt_ord_st_KeyUp(vbKeyF4, 0)
End Sub


Private Sub txt_ord_st_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0007"
        DD.rControl.Add Item:=txt_ord_st
        DD.rControl.Add Item:=txt_ord_st_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_ord_st)) = txt_ord_st.MaxLength Then
        txt_ord_st_name.Text = Gf_ComnNameFind(M_CN1, "B0007", Trim(txt_ord_st.Text), 2)
    Else
        txt_ord_st_name.Text = ""
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

        Exit Sub

    End If

    If Len(Trim(txt_prod_cd)) = txt_prod_cd.MaxLength Then
        txt_prod_cd_name.Text = Gf_ComnNameFind(M_CN1, "B0005", Trim(txt_prod_cd.Text), 2)
    Else
        txt_prod_cd_name.Text = ""
    End If

End Sub

Private Sub txt_ord_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0009"
        DD.rControl.Add Item:=txt_ord_knd
        DD.rControl.Add Item:=txt_ord_knd_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_ord_knd)) = txt_ord_knd.MaxLength Then
        txt_ord_knd_name.Text = Gf_ComnNameFind(M_CN1, "B0009", Trim(txt_ord_knd.Text), 2)
    Else
        txt_ord_knd_name.Text = ""
    End If

End Sub


Private Sub txt_prod_dgr_DblClick()
Call txt_prod_dgr_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_prod_dgr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0034"
        DD.rControl.Add Item:=txt_prod_dgr
        DD.rControl.Add Item:=txt_prod_dgr_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_prod_dgr)) = txt_prod_dgr.MaxLength Then
        txt_prod_dgr_name.Text = Gf_ComnNameFind(M_CN1, "Q0034", Trim(txt_prod_dgr.Text), 2)
    Else
        txt_prod_dgr_name.Text = ""
    End If

End Sub

Private Sub TXT_SALE_WAY_DblClick()
Call txt_sale_way_KeyUp(vbKeyF4, 0)
End Sub


Private Sub txt_sale_way_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0010"
        DD.rControl.Add Item:=txt_sale_way
        DD.rControl.Add Item:=txt_sale_way_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_sale_way)) = txt_sale_way.MaxLength Then
        txt_sale_way_name.Text = Gf_ComnNameFind(M_CN1, "B0010", Trim(txt_sale_way.Text), 2)
    Else
        txt_sale_way_name.Text = ""
    End If

End Sub

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
       
       Load ABX1040C
       
       ABX1040C.txt_form_nm.Text = "ABA1010C"
       
       ABX1040C.Show 1
       
    End If
    
End Sub

Private Sub txt_ord_no_KeyPress(KeyAscii As Integer)

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub TXT_JGXZ_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0048"
        DD.rControl.Add Item:=TXT_JGXZ
        DD.rControl.Add Item:=TXT_JGXZ_NM

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(TXT_JGXZ)) = TXT_JGXZ.MaxLength Then
        TXT_JGXZ_NM.Text = Gf_ComnNameFind(M_CN1, "B0048", Trim(TXT_JGXZ.Text), 2)
    Else
        TXT_JGXZ_NM.Text = ""
    End If

End Sub
Private Sub TXT_JSFS_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0045"
        DD.rControl.Add Item:=TXT_JSFS
        DD.rControl.Add Item:=TXT_JSFS_NM

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(TXT_JSFS)) = TXT_JSFS.MaxLength Then
        TXT_JSFS_NM.Text = Gf_ComnNameFind(M_CN1, "B0045", Trim(TXT_JSFS.Text), 2)
    Else
        TXT_JSFS_NM.Text = ""
    End If

End Sub
Private Sub TXT_FPLX_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0044"
        DD.rControl.Add Item:=TXT_FPLX
        DD.rControl.Add Item:=TXT_FPLX_NM

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(TXT_FPLX)) = TXT_FPLX.MaxLength Then
        TXT_FPLX_NM.Text = Gf_ComnNameFind(M_CN1, "B0044", Trim(TXT_FPLX.Text), 2)
    Else
        TXT_FPLX_NM.Text = ""
    End If

End Sub

Private Sub TXT_THFS_DblClick()
Call TXT_THFS_KeyUp(vbKeyF4, 0)

End Sub

Private Sub TXT_XSXZ_DblClick()
Call TXT_XSXZ_KeyUp(vbKeyF4, 0)

End Sub

Private Sub TXT_XSXZ_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0047"
        DD.rControl.Add Item:=TXT_XSXZ
        DD.rControl.Add Item:=TXT_XSXZ_NM

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(TXT_XSXZ)) = TXT_XSXZ.MaxLength Then
        TXT_XSXZ_NM.Text = Gf_ComnNameFind(M_CN1, "B0047", Trim(TXT_XSXZ.Text), 2)
    Else
        TXT_XSXZ_NM.Text = ""
    End If

End Sub
Private Sub TXT_THFS_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0046"
        DD.rControl.Add Item:=TXT_THFS
        DD.rControl.Add Item:=TXT_THFS_NM

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(TXT_THFS)) = TXT_THFS.MaxLength Then
        TXT_THFS_NM.Text = Gf_ComnNameFind(M_CN1, "B0046", Trim(TXT_THFS.Text), 2)
    Else
        TXT_THFS_NM.Text = ""
    End If

End Sub

Private Sub TXT_YSFS_DblClick()
Call TXT_YSFS_KeyUp(vbKeyF4, 0)

End Sub

Private Sub TXT_YSFS_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0020"
        DD.rControl.Add Item:=TXT_YSFS
        DD.rControl.Add Item:=TXT_YSFS_NM

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(TXT_YSFS)) = TXT_YSFS.MaxLength Then
        TXT_YSFS_NM.Text = Gf_ComnNameFind(M_CN1, "B0020", Trim(TXT_YSFS.Text), 2)
    Else
        TXT_YSFS_NM.Text = ""
    End If

End Sub
Private Sub txt_dest_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyF4 Then

            DD.sWitch = "MS"
            DD.rControl.Add Item:=TXT_DEST
            DD.rControl.Add Item:=TXT_DEST_NM
            DD.nameType = "1"

            Call Gf_Destination_DD(M_CN1, KeyCode)

            Exit Sub

    End If

    If Len(Trim(TXT_DEST)) = TXT_DEST.MaxLength Then
        TXT_DEST_NM.Text = Gf_DestNameFind(M_CN1, Trim(TXT_DEST.Text), 1)
    Else
        TXT_DEST_NM.Text = ""
    End If
        
End Sub


