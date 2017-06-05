VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2040C 
   Caption         =   "探伤实绩查询及修改界面_AGC2040C"
   ClientHeight    =   10500
   ClientLeft      =   390
   ClientTop       =   1935
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10500
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_REASON_FL 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   1
      Left            =   6480
      TabIndex        =   80
      Top             =   9630
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox TXT_REASON_NAME 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   7365
      TabIndex        =   79
      Top             =   9630
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.CheckBox CHK_NEXT_PRC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "热处理"
      Height          =   240
      Index           =   1
      Left            =   3210
      TabIndex        =   78
      Tag             =   "T"
      Top             =   10050
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox TXT_NEXT_PROC 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   77
      Tag             =   "后道工序"
      Top             =   10020
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CheckBox CHK_NEXT_PRC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "钢板库"
      Height          =   240
      Index           =   0
      Left            =   1905
      TabIndex        =   76
      Tag             =   "P"
      Top             =   10050
      Visible         =   0   'False
      Width           =   900
   End
   Begin Threed.SSFrame sf1 
      Height          =   4245
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   7488
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
      Caption         =   "探伤实绩"
      Begin VB.TextBox txt_woo_rsn 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1605
         MaxLength       =   2
         TabIndex        =   86
         Tag             =   "余材原因"
         Top             =   3855
         Width           =   885
      End
      Begin VB.TextBox TXT_UST_PLAN_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2505
         TabIndex        =   85
         Top             =   1695
         Width           =   2895
      End
      Begin VB.TextBox TXT_UST_FL 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1605
         TabIndex        =   84
         Top             =   1695
         Width           =   885
      End
      Begin VB.CheckBox CHK_PRD_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "待判"
         Height          =   240
         Index           =   5
         Left            =   3855
         TabIndex        =   75
         Tag             =   "4"
         Top             =   3195
         Width           =   1230
      End
      Begin VB.TextBox TXT_LOC 
         Height          =   315
         Left            =   12675
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   73
         Top             =   1365
         Width           =   1755
      End
      Begin Threed.SSCommand Cmd_Edit 
         Height          =   300
         Left            =   11460
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1365
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   196609
         Font3D          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "更新垛位"
      End
      Begin VB.TextBox TXT_ADDR 
         Height          =   315
         Index           =   2
         Left            =   13815
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   67
         Top             =   1035
         Width           =   600
      End
      Begin VB.TextBox TXT_ADDR 
         Height          =   315
         Index           =   1
         Left            =   13200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   66
         Top             =   1035
         Width           =   600
      End
      Begin VB.TextBox txt_Scrap_code 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   300
         Left            =   6885
         MaxLength       =   1
         TabIndex        =   69
         Tag             =   "原因"
         Top             =   3795
         Width           =   885
      End
      Begin VB.TextBox txt_Scrap_name 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   3795
         Width           =   1995
      End
      Begin VB.TextBox TXT_ADDR 
         Height          =   315
         Index           =   0
         Left            =   12675
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   65
         Top             =   1035
         Width           =   510
      End
      Begin VB.TextBox TXT_STLGRD 
         Height          =   285
         Left            =   11955
         TabIndex        =   64
         Top             =   -45
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox TXT_APLY_ENDUSE_CD 
         Height          =   285
         Left            =   11475
         TabIndex        =   63
         Top             =   -60
         Visible         =   0   'False
         Width           =   390
      End
      Begin InDate.ULabel ULabel22 
         Height          =   345
         Index           =   2
         Left            =   5700
         Top             =   3345
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   609
         Caption         =   "(改判)"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
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
      Begin VB.TextBox TXT_UST_GRADE 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1605
         TabIndex        =   62
         Top             =   2355
         Width           =   885
      End
      Begin VB.TextBox TXT_UST_STAND_NO 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1605
         TabIndex        =   61
         Top             =   2025
         Width           =   885
      End
      Begin VB.TextBox TXT_UST_GRADE_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2505
         TabIndex        =   60
         Top             =   2355
         Width           =   2895
      End
      Begin VB.TextBox TXT_REMARK 
         Height          =   630
         Left            =   11475
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   59
         Top             =   2025
         Width           =   2970
      End
      Begin VB.TextBox TXT_ADD_THK 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   12675
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   58
         Top             =   375
         Width           =   1755
      End
      Begin VB.TextBox TXT_REASON_NAME 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   13200
         TabIndex        =   53
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TXT_REASON_FL 
         Height          =   315
         Index           =   2
         Left            =   12675
         TabIndex        =   52
         Top             =   720
         Width           =   510
      End
      Begin VB.TextBox TXT_UST_STAND_NAME 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2505
         TabIndex        =   51
         Top             =   2025
         Width           =   2895
      End
      Begin VB.TextBox TXT_UST_PREC 
         Height          =   315
         Left            =   1605
         TabIndex        =   50
         Top             =   1365
         Width           =   3795
      End
      Begin VB.TextBox TXT_UST_METHOD 
         Height          =   315
         Left            =   1605
         TabIndex        =   49
         Top             =   1035
         Width           =   3795
      End
      Begin VB.TextBox TXT_UST_HEAD 
         Height          =   315
         Left            =   1605
         TabIndex        =   48
         Top             =   705
         Width           =   3795
      End
      Begin VB.TextBox TXT_KIND_NO 
         Height          =   315
         Left            =   1605
         TabIndex        =   47
         Top             =   375
         Width           =   3795
      End
      Begin VB.CheckBox CHK_PRD_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "正品"
         Height          =   240
         Index           =   0
         Left            =   2580
         TabIndex        =   45
         Tag             =   "1"
         Top             =   3195
         Width           =   1230
      End
      Begin VB.CheckBox CHK_PRD_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "改判"
         Height          =   240
         Index           =   1
         Left            =   2580
         TabIndex        =   44
         Tag             =   "2"
         Top             =   3525
         Width           =   1230
      End
      Begin VB.CheckBox CHK_PRD_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "协议"
         Height          =   240
         Index           =   2
         Left            =   2580
         TabIndex        =   43
         Tag             =   "3"
         Top             =   3870
         Width           =   1230
      End
      Begin VB.CheckBox CHK_PRD_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "次品"
         Height          =   240
         Index           =   3
         Left            =   3855
         TabIndex        =   42
         Tag             =   "5"
         Top             =   3525
         Width           =   1230
      End
      Begin VB.CheckBox CHK_PRD_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "废钢"
         Height          =   240
         Index           =   4
         Left            =   3855
         TabIndex        =   41
         Tag             =   "7"
         Top             =   3870
         Width           =   1230
      End
      Begin VB.TextBox TXT_PRD_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   3180
         Width           =   885
      End
      Begin VB.TextBox TXT_INSP_MAN 
         Height          =   330
         Left            =   12690
         MaxLength       =   7
         TabIndex        =   39
         Tag             =   "探伤人员"
         Top             =   2745
         Width           =   930
      End
      Begin VB.TextBox TXT_UST_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   " "
         Top             =   2775
         Width           =   885
      End
      Begin VB.CheckBox CHK_UST_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "合格"
         Height          =   240
         Index           =   0
         Left            =   2580
         TabIndex        =   37
         Tag             =   "Y"
         Top             =   2835
         Width           =   1230
      End
      Begin VB.CheckBox CHK_UST_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "不合格"
         Height          =   240
         Index           =   1
         Left            =   3855
         TabIndex        =   36
         Tag             =   "N"
         Top             =   2835
         Width           =   1230
      End
      Begin VB.TextBox TXT_PROC_FLAG 
         Height          =   285
         Left            =   5490
         TabIndex        =   35
         Top             =   150
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.TextBox txt_stdspec 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6885
         Locked          =   -1  'True
         TabIndex        =   34
         Tag             =   "标准代码"
         Top             =   2745
         Width           =   2205
      End
      Begin VB.TextBox txt_stdspec_name 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9090
         MaxLength       =   40
         TabIndex        =   33
         Tag             =   "STDSPEC"
         Top             =   2745
         Width           =   2145
      End
      Begin VB.TextBox txt_stdspec_chg 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6885
         MaxLength       =   18
         TabIndex        =   32
         Tag             =   "标准代码"
         Top             =   3075
         Width           =   2205
      End
      Begin VB.TextBox txt_stdspec_name_chg 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9090
         MaxLength       =   40
         TabIndex        =   31
         Tag             =   "STDSPEC"
         Top             =   3075
         Width           =   2145
      End
      Begin VB.TextBox txt_stdspec_yy 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10845
         MaxLength       =   40
         TabIndex        =   30
         Tag             =   "STDSPEC"
         Top             =   -45
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox TXT_INSP_WID_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   7830
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2355
         Width           =   1080
      End
      Begin VB.TextBox TXT_INSP_LEN_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   8925
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2355
         Width           =   1245
      End
      Begin VB.TextBox TXT_INSP_THK_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6885
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2355
         Width           =   930
      End
      Begin VB.TextBox TXT_INSP_WGT_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   10185
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2355
         Width           =   1035
      End
      Begin VB.TextBox TXT_REASON_FL 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   8205
         TabIndex        =   4
         Top             =   3435
         Width           =   885
      End
      Begin VB.TextBox TXT_REASON_NAME 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   9090
         TabIndex        =   0
         Top             =   3435
         Width           =   3555
      End
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Index           =   0
         Left            =   6885
         Top             =   3435
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "改判原因"
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   195
         Top             =   705
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "探头"
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
         Left            =   195
         Top             =   375
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "仪器型号"
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
         Index           =   2
         Left            =   11475
         Top             =   720
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "返剪原因"
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
         Left            =   195
         Top             =   1035
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "探伤方式"
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   195
         Top             =   1365
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "探伤灵敏度"
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
         Left            =   195
         Top             =   2025
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "探伤标准"
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
      Begin InDate.ULabel ULabel28 
         Height          =   315
         Left            =   7830
         Top             =   375
         Width           =   1080
         _ExtentX        =   1905
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
      Begin InDate.ULabel ULabel29 
         Height          =   315
         Left            =   6885
         Top             =   375
         Width           =   930
         _ExtentX        =   1640
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
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Left            =   8925
         Top             =   375
         Width           =   1245
         _ExtentX        =   2196
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
      Begin InDate.ULabel ULabel33 
         Height          =   315
         Left            =   5655
         Top             =   2355
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "判定结果"
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
      Begin CSTextLibCtl.sidbEdit SDB_WGT_ORD 
         Height          =   315
         Left            =   10185
         TabIndex        =   14
         Top             =   1365
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_WGT 
         Height          =   315
         Left            =   10185
         TabIndex        =   15
         Top             =   705
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_WID_MX 
         Height          =   315
         Left            =   7830
         TabIndex        =   16
         Top             =   1695
         Width           =   1080
         _Version        =   262145
         _ExtentX        =   1905
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   14737632
         BackColor       =   14737632
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LEN_MX 
         Height          =   315
         Left            =   8925
         TabIndex        =   17
         Top             =   1695
         Width           =   1245
         _Version        =   262145
         _ExtentX        =   2196
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_WID_MN 
         Height          =   315
         Left            =   7830
         TabIndex        =   18
         Top             =   2025
         Width           =   1080
         _Version        =   262145
         _ExtentX        =   1905
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_THK_MN 
         Height          =   315
         Left            =   6885
         TabIndex        =   19
         Top             =   2025
         Width           =   930
         _Version        =   262145
         _ExtentX        =   1640
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LEN_MN 
         Height          =   315
         Left            =   8925
         TabIndex        =   20
         Top             =   2025
         Width           =   1245
         _Version        =   262145
         _ExtentX        =   2196
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_PWGT_MN 
         Height          =   315
         Left            =   10185
         TabIndex        =   21
         Top             =   2025
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_WID 
         Height          =   315
         Left            =   7830
         TabIndex        =   22
         Top             =   705
         Width           =   1080
         _Version        =   262145
         _ExtentX        =   1905
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
      Begin CSTextLibCtl.sidbEdit SDB_THK 
         Height          =   315
         Left            =   6885
         TabIndex        =   23
         Top             =   705
         Width           =   930
         _Version        =   262145
         _ExtentX        =   1640
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
      Begin CSTextLibCtl.sidbEdit SDB_LEN 
         Height          =   315
         Left            =   8925
         TabIndex        =   24
         Top             =   705
         Width           =   1245
         _Version        =   262145
         _ExtentX        =   2196
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
      Begin InDate.ULabel ULabel38 
         Height          =   315
         Left            =   5655
         Top             =   2025
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "下公差"
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
         Left            =   5655
         Top             =   705
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "改判/返剪"
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_THK_MX 
         Height          =   315
         Left            =   6885
         TabIndex        =   25
         Top             =   1695
         Width           =   930
         _Version        =   262145
         _ExtentX        =   1640
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_PWGT_MX 
         Height          =   315
         Left            =   10185
         TabIndex        =   26
         Top             =   1695
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   5655
         Top             =   1695
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "上公差"
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
         Left            =   10185
         Top             =   375
         Width           =   1035
         _ExtentX        =   1826
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
      Begin CSTextLibCtl.sidbEdit SDB_ORD_WID 
         Height          =   315
         Left            =   7830
         TabIndex        =   27
         Top             =   1365
         Width           =   1080
         _Version        =   262145
         _ExtentX        =   1905
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
      Begin CSTextLibCtl.sidbEdit SDB_ORD_THK 
         Height          =   315
         Left            =   6885
         TabIndex        =   28
         Top             =   1365
         Width           =   930
         _Version        =   262145
         _ExtentX        =   1640
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
      Begin CSTextLibCtl.sidbEdit SDB_ORD_LEN 
         Height          =   315
         Left            =   8925
         TabIndex        =   29
         Top             =   1365
         Width           =   1245
         _Version        =   262145
         _ExtentX        =   2196
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
      Begin InDate.ULabel ULabel45 
         Height          =   315
         Left            =   5655
         Top             =   1365
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "订单"
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
         Left            =   11475
         Top             =   2745
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "探伤人员"
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
         Left            =   11475
         Top             =   3075
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "探伤时间"
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
      Begin CSTextLibCtl.sitxEdit TXT_INSP_OCCR_TIME 
         Height          =   315
         Left            =   12690
         TabIndex        =   46
         Top             =   3075
         Width           =   2160
         _Version        =   262145
         _ExtentX        =   3810
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
      Begin InDate.ULabel ULabel36 
         Height          =   360
         Left            =   195
         Top             =   2775
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   635
         Caption         =   "探伤判定结果"
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
      Begin CSTextLibCtl.sidbEdit SDB_WGT_ORG 
         Height          =   315
         Left            =   10185
         TabIndex        =   54
         Top             =   1035
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_WID_ORG 
         Height          =   315
         Left            =   7830
         TabIndex        =   55
         Top             =   1035
         Width           =   1080
         _Version        =   262145
         _ExtentX        =   1905
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
      Begin CSTextLibCtl.sidbEdit SDB_THK_ORG 
         Height          =   315
         Left            =   6885
         TabIndex        =   56
         Top             =   1035
         Width           =   930
         _Version        =   262145
         _ExtentX        =   1640
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
      Begin CSTextLibCtl.sidbEdit SDB_LEN_ORG 
         Height          =   315
         Left            =   8925
         TabIndex        =   57
         Top             =   1035
         Width           =   1245
         _Version        =   262145
         _ExtentX        =   2196
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
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   5655
         Top             =   1035
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "实绩"
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   5655
         Top             =   375
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "尺寸"
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
         Height          =   330
         Left            =   11475
         Top             =   375
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         Caption         =   "厚度附加值"
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
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   11475
         Top             =   1695
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         Caption         =   "备注"
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
         Left            =   195
         Top             =   2355
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "检查标准等级"
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
      Begin InDate.ULabel ULabel22 
         Height          =   1005
         Index           =   1
         Left            =   5640
         Top             =   2745
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1773
         Caption         =   "标准号"
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
         Index           =   3
         Left            =   11475
         Top             =   1035
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "垛位号"
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
         Height          =   300
         Left            =   5640
         Top             =   3795
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         Caption         =   "判废原因"
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
      Begin Threed.SSCommand Cmd_Edit_Date 
         Height          =   300
         Left            =   13620
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   2745
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   196609
         Font3D          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "更新时间"
      End
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   195
         Top             =   1695
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "计划探伤标准"
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
      Begin InDate.ULabel ULabel22 
         Height          =   615
         Index           =   0
         Left            =   195
         Top             =   3180
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   1085
         Caption         =   "最终等级判定"
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
      Begin InDate.ULabel ULabel77 
         Height          =   315
         Left            =   195
         Top             =   3855
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "余材原因"
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
   End
   Begin Threed.SSFrame Single 
      Height          =   570
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   1005
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox CBO_EQPM 
         Height          =   315
         ItemData        =   "AGC2040C.frx":0000
         Left            =   12495
         List            =   "AGC2040C.frx":000A
         TabIndex        =   87
         Tag             =   "班次"
         Text            =   "NDT"
         Top             =   120
         Width           =   1155
      End
      Begin VB.ComboBox CBO_SHIFT 
         Height          =   315
         ItemData        =   "AGC2040C.frx":0017
         Left            =   10215
         List            =   "AGC2040C.frx":0024
         TabIndex        =   71
         Tag             =   "班次"
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox TXT_PLATE_NO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1710
         MaxLength       =   14
         TabIndex        =   3
         Tag             =   "钢板号"
         Top             =   105
         Width           =   2070
      End
      Begin InDate.ULabel ULabel16 
         Height          =   360
         Left            =   345
         Top             =   105
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         Caption         =   "钢板号"
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
         Left            =   4210
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "生产日期"
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
         Left            =   9180
         Top             =   120
         Width           =   1005
         _ExtentX        =   1773
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
      Begin InDate.UDate SDT_PROD_DATE 
         Height          =   315
         Left            =   5580
         TabIndex        =   82
         Tag             =   "起始日期"
         Top             =   120
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
      Begin InDate.UDate SDT_PROD_TO_DATE 
         Height          =   315
         Left            =   7350
         TabIndex        =   83
         Tag             =   "起始日期"
         Top             =   120
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
      Begin InDate.ULabel UST_CD 
         Height          =   315
         Left            =   11460
         Top             =   120
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "设备"
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
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   120
         Left            =   7125
         TabIndex        =   81
         Top             =   240
         Width           =   195
      End
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   4365
      Left            =   90
      TabIndex        =   70
      Top             =   600
      Width           =   15075
      _Version        =   393216
      _ExtentX        =   26591
      _ExtentY        =   7699
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   27
      MaxRows         =   5
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "AGC2040C.frx":0031
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   690
      Left            =   7755
      TabIndex        =   5
      Top             =   1665
      Visible         =   0   'False
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   1217
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   540
         Top             =   180
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "钢板分板数"
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
      Begin Threed.SSCommand cmd_divide 
         Height          =   360
         Left            =   2505
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   150
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   635
         _Version        =   196609
         Font3D          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "分板"
      End
      Begin CSTextLibCtl.sidbEdit SDB_DIVIDE_CNT 
         Height          =   315
         Left            =   1890
         TabIndex        =   7
         Top             =   180
         Width           =   450
         _Version        =   262145
         _ExtentX        =   794
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         MaxValue        =   9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin Threed.SSCommand cmd_divide_ok 
         Height          =   360
         Left            =   4035
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   150
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   635
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "确认分板"
      End
      Begin Threed.SSCommand cmd_divide_delete 
         Height          =   360
         Left            =   5235
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   150
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   635
         _Version        =   196609
         Font3D          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "取消分板"
      End
   End
   Begin InDate.ULabel ULabel14 
      Height          =   330
      Left            =   90
      Top             =   10020
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   582
      Caption         =   "后道工序"
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
      Index           =   1
      Left            =   5160
      Top             =   9630
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "原因2"
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
End
Attribute VB_Name = "AGC2040C"
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
'-- Program Name      UST实绩查询及修改界面
'-- Program ID        AGC2040C
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
Public sQuery_load As String        'Active Form sQuery Setting
Public sQuery_Rt As String          'Active Form sQuery Setting

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

Dim sControl  As New Collection      'Master Clear Key Collection
Dim MC        As New Collection      'Master Collection
Dim Mc1       As New Collection      'Master Collection

Dim sc1       As New Collection      'Spread Collection
Dim Proc_Sc   As New Collection      'Spread Struc Collection

Dim sCheck As String
Dim sQuery      As String

Const SS1_URGNT_FL = 23                    '紧急订单绿色标记 2012-11-08  by  LiQian

Private Sub Form_Define()
    Dim iIndex As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDT_PROD_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDT_PROD_TO_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_PROC_FLAG, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_APLY_ENDUSE_CD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_STLGRD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                                                                                                                                                
          Call Gp_Ms_Collection(TXT_KIND_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_UST_HEAD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TXT_UST_METHOD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_UST_PREC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_UST_STAND_NO, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_UST_GRADE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_THK, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_THK_MX, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_THK_MN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_WID, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_WID_MX, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_WID_MN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_LEN, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LEN_MX, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LEN_MN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_WGT, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_PWGT_MX, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_PWGT_MN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_THK_ORG, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_WID_ORG, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_LEN_ORG, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_WGT_ORG, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_ORD_THK, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_ORD_WID, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_ORD_LEN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_WGT_ORD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_ust_GRD, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_PRD_GRD, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_INSP_MAN, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(TXT_INSP_OCCR_TIME, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_ADD_THK, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(TXT_LOC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_REMARK, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_STDSPEC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_stdspec_chg, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_REASON_FL(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_REASON_FL(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_REASON_FL(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_ADDR(0), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_ADDR(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_ADDR(2), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Scrap_code, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Scrap_name, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_NEXT_PROC, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_UST_FL, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_woo_rsn, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '余材代码 20110608  ADD BY LiQian 如有发生降余材，提示必须输入余材代码，将余材代码做记录
             Call Gp_Ms_Collection(CBO_EQPM, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
     Call Gp_Clear_Collection(CHK_UST_GRD(0), "s", sControl)
     Call Gp_Clear_Collection(CHK_UST_GRD(1), "s", sControl)
     Call Gp_Clear_Collection(CHK_PRD_GRD(0), "s", sControl)
     Call Gp_Clear_Collection(CHK_PRD_GRD(1), "s", sControl)
     Call Gp_Clear_Collection(CHK_PRD_GRD(2), "s", sControl)
     Call Gp_Clear_Collection(CHK_PRD_GRD(3), "s", sControl)
     Call Gp_Clear_Collection(CHK_PRD_GRD(4), "s", sControl)
     Call Gp_Clear_Collection(CHK_PRD_GRD(5), "s", sControl)
     Call Gp_Clear_Collection(CHK_NEXT_PRC(0), "s", sControl)
     Call Gp_Clear_Collection(CHK_NEXT_PRC(1), "s", sControl)
     
    MC.Add Item:=sControl, Key:="sControl"
    
    'MASTER Collection
    Mc1.Add Item:="AGC2040C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="AGC2040C.P_REFER", Key:="P-R"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
      
    'Spread_Collection
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
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
'
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2040C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
        
    Call Gp_Sp_ColHidden(ss1, 4, True)
    Call Gp_Sp_ColHidden(ss1, 5, True)
    Call Gp_Sp_ColHidden(ss1, 6, True)
    
    Call Gp_Sp_ColHidden(ss1, 15, True)
    Call Gp_Sp_ColHidden(ss1, 16, True)
    Call Gp_Sp_ColHidden(ss1, 17, True)
    Call Gp_Sp_ColHidden(ss1, 18, True)
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub CHK_NEXT_PRC_Click(Index As Integer)
    Dim iCount      As Integer
    Dim iIndexStr   As Integer
    
    If sCheck <> "" Then Exit Sub

    iCount = 0
    sCheck = "**"
            
    If CHK_NEXT_PRC(Index).Value = ssCBUnchecked Then
        For iIndexStr = 0 To 1
            If CHK_NEXT_PRC(iIndexStr).Value = ssCBChecked Then
               iCount = iCount + 1
            End If
        Next iIndexStr
        If iCount = 0 Then
            txt_NEXT_PROC.Text = ""
            CHK_NEXT_PRC(Index).ForeColor = &H808080
            sCheck = ""
            Exit Sub
        End If
    Else
        For iIndexStr = 0 To 1
            CHK_NEXT_PRC(iIndexStr).ForeColor = &H808080
            CHK_NEXT_PRC(iIndexStr).Value = ssCBUnchecked
        Next iIndexStr
    End If
    
    CHK_NEXT_PRC(Index).ForeColor = &HFF&
    CHK_NEXT_PRC(Index).Value = ssCBChecked
    
    txt_NEXT_PROC.Text = CHK_NEXT_PRC(Index).Tag
        
    sCheck = ""
    
End Sub

Private Sub Cmd_Edit_Click()
    Dim sQuery      As String
    Dim sLoc        As String
    Dim sComments   As String
    Dim sDate       As String
    Dim lSeq        As Long
    Dim iROW        As Integer
    
    Dim sMesg       As String
    
    On Error GoTo UPDATE_ERROR

    Screen.MousePointer = vbHourglass
    
    If sUserID <> "1CJM001" And sUserID <> "1CJD001" And sUserID <> "1CJM005" Then '刘素梅，姜利，李钟玮
        sMesg = " 您没有权限更新垛位号 ！"
        Call Gp_MsgBoxDisplay(sMesg)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    M_CN1.BeginTrans
    
    sLoc = Trim(TXT_LOC.Text)
    sComments = Trim(TXT_REMARK.Text)
    
    sQuery = "         UPDATE  GP_USTRESULT                                      " & vbCrLf
    sQuery = sQuery & "   SET  UST_LOC       = '" & sLoc & "'                    " & vbCrLf
    sQuery = sQuery & "       ,UST_REMARTS   = '" & sComments & "'               " & vbCrLf
    sQuery = sQuery & " WHERE  PLATE_NO      = '" & Trim(TXT_PLATE_NO.Text) & "' " & vbCrLf

    M_CN1.Execute sQuery
        
    M_CN1.CommitTrans

    Screen.MousePointer = vbDefault
    
    Exit Sub

UPDATE_ERROR:

    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay(Err.Description & sQuery)
    
    M_CN1.RollbackTrans
End Sub

Private Sub Cmd_Edit_Date_Click()

    Dim sQuery               As String
    Dim sUST_END_DATE        As String
    Dim sShift               As String
    Dim sGroup_cd            As String
    
    Dim sMesg                As String
    
    On Error GoTo UPDATE_ERROR

    Screen.MousePointer = vbHourglass
    
    If sUserID <> "1CJM001" And sUserID <> "1CJD001" Then '刘素梅，姜利
        sMesg = " 您没有权限更新垛位号 ！"
        Call Gp_MsgBoxDisplay(sMesg)
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    M_CN1.BeginTrans

    sUST_END_DATE = TXT_INSP_OCCR_TIME.RawData
    
    sQuery = "         UPDATE  GP_USTRESULT                                      " & vbCrLf
    sQuery = sQuery & "   SET  UST_END_DATE       = '" & sUST_END_DATE & "'      " & vbCrLf
    sQuery = sQuery & "       ,SHIFT              = Gf_Shiftset('" & sUST_END_DATE & "')             " & vbCrLf
    sQuery = sQuery & "       ,GROUP_CD           = Gf_Groupset('C1',Gf_Shiftset('" & sUST_END_DATE & "'),'" & Mid(sUST_END_DATE, 1, 8) & "')           " & vbCrLf
    sQuery = sQuery & " WHERE  PLATE_NO           = '" & Trim(TXT_PLATE_NO.Text) & "' " & vbCrLf

    M_CN1.Execute sQuery

    M_CN1.CommitTrans

    Screen.MousePointer = vbDefault
    
    Exit Sub

UPDATE_ERROR:

    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay(Err.Description & sQuery)
    
    M_CN1.RollbackTrans
End Sub




Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        If Len(TXT_PLATE_NO.Text) >= 8 Then
           Call Form_Ref
        End If
'        KeyAscii = 0
'        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    
    Call Gp_Sp_ColHidden(ss1, 4, True)
    Call Gp_Sp_ColHidden(ss1, 5, True)
    Call Gp_Sp_ColHidden(ss1, 6, True)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)

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
    
    Set sControl = Nothing
    Set MC = Nothing

    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    Dim iCount As Integer
    
    If Gf_Sp_Cls(sc1) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_SSCheck_Cls(MC("sControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)

'        TXT_INSP_MAN = sUserID
        SDB_DIVIDE_CNT.Value = 0
        TXT_INSP_THK_GRD.Text = ""
        TXT_INSP_WID_GRD.Text = ""
        TXT_INSP_LEN_GRD.Text = ""
        TXT_INSP_WGT_GRD.Text = ""
        
        For iCount = 0 To 1
            CHK_NEXT_PRC(iCount).Value = 0
        Next iCount
        
        ss1.BlockMode = True
        ss1.Row = -1
        ss1.Col = -1
        ss1.BackColor = &HFFFFFF
        ss1.BlockMode = False
    End If
End Sub

Public Sub Form_Ref()

    Dim iAddr As String
    Dim iAddr1 As String
    Dim iAddr2 As String
    
    Dim iSTAND_NO As String
    Dim iDATETIME As String
    Dim iTXT_REMARK As String
    
    Dim iROW As Integer
    Dim iCol As Integer
    Dim sUrgnt_Fl As String
    
    iAddr = TXT_ADDR(0).Text
    iAddr1 = TXT_ADDR(1).Text
    iAddr2 = TXT_ADDR(2).Text
    
    iSTAND_NO = TXT_UST_STAND_NO.Text
    iDATETIME = TXT_INSP_OCCR_TIME.RawData
    iTXT_REMARK = TXT_REMARK.Text
    
    Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, , , True)
    ss1.OperationMode = OperationModeNormal
    
    '紧急订单绿色显示 add by liqian 2012-11-08
     With ss1
          For iROW = 1 To .MaxRows
             .Row = iROW:
              .Col = SS1_URGNT_FL:    sUrgnt_Fl = Trim(.Text)
            
              If sUrgnt_Fl = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iROW, iROW, &HC000&)
              End If
          Next iROW
    End With
    
    If Len(TXT_PLATE_NO.Text) = 14 Then
        If Gf_Ms_Refer(M_CN1, Mc1, , , True) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Display_Data_Edit
        End If
    Else
        Call ss1_DblClick(1, 1)
    End If
    
    If txt_NEXT_PROC.Text = "" Or txt_NEXT_PROC.Text = "U" Then
       CHK_NEXT_PRC(0).Value = 1
       txt_NEXT_PROC.Text = "P"
    End If
        
    If Len(iAddr) = 3 And Len(iAddr1) = 4 And Len(iAddr2) > 0 Then
'       TXT_ADDR(0).Text = iAddr
'       TXT_ADDR(1).Text = iAddr1
'       TXT_ADDR(2).Text = Format(Val(iAddr2) + 1, "000")
       TXT_LOC = iAddr & iAddr1 & Format(Val(iAddr2) + 1, "000")
    End If
    
    If TXT_UST_STAND_NO.Text = "" Then TXT_UST_STAND_NO.Text = iSTAND_NO
    If TXT_INSP_OCCR_TIME.RawData = "" Then TXT_INSP_OCCR_TIME.RawData = iDATETIME
    TXT_REMARK.Text = iTXT_REMARK
            
    Call Gp_Sp_ColHidden(ss1, 4, True)
    Call Gp_Sp_ColHidden(ss1, 5, True)
    Call Gp_Sp_ColHidden(ss1, 6, True)
    
End Sub

Public Sub Form_Pro()

    Dim sMesg   As String
    Dim iCount  As Integer
    
    Dim iAddr As String
    Dim iAddr1 As String
    Dim iAddr2 As String
    
    iAddr = TXT_ADDR(0).Text
    iAddr1 = TXT_ADDR(1).Text
    iAddr2 = TXT_ADDR(2).Text
    
    If txt_stdspec_chg.Text <> "" And Trim(TXT_REASON_FL(0).Text) = "" And Trim(TXT_REASON_FL(1).Text) = "" Then
        sMesg = " 请输入改判原因 ！"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
'    If CHK_PRD_GRD(1).Value <> ssCBChecked Then
        If SDB_WGT_ORG.Value > 0 And SDB_WGT.Value <> SDB_WGT_ORG.Value And Trim(TXT_REASON_FL(2).Text) = "" Then
            sMesg = " 请输入返剪原因 ！"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        End If
'    End If
    
    If Not Gp_DateCheck(TXT_INSP_OCCR_TIME) Then
        sMesg = " 请正确输入检查时间 ！"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    If CHK_PRD_GRD(4).Value = ssCBChecked Then
        If txt_Scrap_code.Text = "" Then
            sMesg = " 请正确输入废钢原因 ！"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        End If
    End If
    
    If CBO_EQPM.Text <> "NDT" And CBO_EQPM.Text <> "NK" Then
            sMesg = " 请正确选择设备"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
    End If
        
    If Gf_Mc_Authority(sAuthority, Mc1) Then
    
       TXT_INSP_MAN = sUserID
       
       If TXT_UST_STAND_NO.Text <> TXT_UST_FL.Text Then
            sMesg = " 探伤标准与计划探伤标准不一致，您确认保存么？"
            If Gf_MessConfirm(sMesg, "Q") Then
                If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
                    Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
                    TXT_PLATE_NO.Enabled = True
                End If
            Else
                Exit Sub
            End If
       Else
                If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
                    Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
                    TXT_PLATE_NO.Enabled = True
                End If
       End If

    End If
    
End Sub

Private Sub SDB_THK_Change()
    Call PRD_WEIGHT_CALC
End Sub
    
Private Sub SDB_WID_Change()
    Call PRD_WEIGHT_CALC
End Sub

Private Sub SDB_LEN_Change()
    Call PRD_WEIGHT_CALC
End Sub

Private Sub PRD_WEIGHT_CALC()

    Dim dThk        As Double
    Dim dWid        As Double
    Dim dLen        As Double
    Dim sQuery      As String
    Dim RS As New ADODB.Recordset
    
    dThk = Val(Format(SDB_THK.Text, "####0.##") & "")
    dWid = Val(Format(SDB_WID.Text, "###0") & "")
    dLen = Val(Format(SDB_LEN.Text, "###0.##") & "")
    
    If dThk > 0 And dWid > 0 And dLen > 0 Then
        SDB_WGT.Text = Cal_Plate_Wgt("WGT", dThk, dWid, dLen)
        TXT_ADD_THK.Text = Cal_Plate_Wgt("VAT", dThk, dWid, dLen)
    End If
        
    Call Size_Grade_Edit
End Sub



Private Sub SDT_PROD_DATE_DblClick()
     SDT_PROD_DATE.RawData = Gf_DTSet(M_CN1, "D")
End Sub

Private Sub SDT_PROD_DATE_GotFocus()
     If SDT_PROD_DATE.RawData = "" Then
        SDT_PROD_DATE.RawData = Gf_DTSet(M_CN1, "D")
     End If
     If SDT_PROD_TO_DATE.RawData = "" Then
        SDT_PROD_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub SDT_PROD_TO_DATE_GotFocus()
     If SDT_PROD_TO_DATE.RawData = "" Then
        SDT_PROD_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub


Private Sub TXT_STDSPEC_Change()
    Dim RS  As New ADODB.Recordset

    If Trim(TXT_STDSPEC.Text) = "" Then Exit Sub
    
    sQuery = "SELECT  Gf_Stdspec_Name_Chn('" & Trim(TXT_STDSPEC.Text) & "')" & vbCrLf
    sQuery = sQuery & "       FROM  DUAL " & vbCrLf
    RS.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        txt_stdspec_name.Text = RS(0).Value & ""
    End If
    
    RS.Close
    Set RS = Nothing
End Sub



Private Sub txt_stdspec_chg_DblClick()
  DD.sWitch = "MS"
         DD.DataDicType = "C"
         DD.rControl.Add Item:=txt_stdspec_chg
        
         Call Pf_Common_DD(M_CN1, vbKeyF4)
         
         Exit Sub
End Sub

Private Function Pf_Common_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sSelect = False
        DD.sWhere = ""
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
        DD.sSelect = False
        DD.sWhere = ""
        DD.sKey = ""
        
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "HC"        'Common Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    DD.sQuery = "SELECT CD_SHORT_NAME ""标准代号"", CD_NAME ""标准中文名"" FROM ZP_CD WHERE CD_MANA_NO = 'G0035'"
    
    Call Gf_DD_Display(Conn, DD.sQuery, False)
    
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function


Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        txt_stdspec_yy.Text = ""
        DD.rControl.Add Item:=txt_stdspec_chg
        DD.rControl.Add Item:=txt_stdspec_yy
        DD.rControl.Add Item:=txt_stdspec_name_chg

        Call Gf_StdSPEC_DD2(M_CN1, vbKeyF4)

        Exit Sub
    End If

End Sub

Private Sub TXT_UST_FL_Change()
    TXT_UST_PLAN_NAME.Text = Gf_ComnNameFind(M_CN1, "Q0046", TXT_UST_FL.Text, 1)
End Sub

Private Sub TXT_UST_STAND_NO_Change()
    TXT_UST_STAND_NAME.Text = Gf_ComnNameFind(M_CN1, "Q0046", TXT_UST_STAND_NO.Text, 1)
End Sub

Private Sub TXT_UST_STAND_NO_dblClick()

    Call TXT_UST_STAND_NO_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TXT_UST_STAND_NO_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "Q0046"
        DD.rControl.Add Item:=TXT_UST_STAND_NO
    
        DD.nameType = "2"
        
        Call Gf_Mill_Common_DD(M_CN1, vbKeyF4)
    End If

End Sub

Private Sub TXT_UST_GRADE_Change()
    TXT_UST_GRADE_NAME.Text = Gf_ComnNameFind(M_CN1, "Q0053", TXT_UST_GRADE.Text, 1)
End Sub

Private Sub TXT_UST_GRADE_DblClick()

    Call TXT_UST_GRADE_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TXT_UST_GRADE_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Q0053"
        DD.rControl.Add Item:=TXT_UST_GRADE
    
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, vbKeyF4)
    
    End If

End Sub

Private Sub TXT_INSP_MAN_DblClick()
    TXT_INSP_MAN.Text = sUserID
End Sub

Private Sub TXT_INSP_OCCR_TIME_DblClick()
    TXT_INSP_OCCR_TIME.RawData = Gf_DTSet(M_CN1)
End Sub

Private Sub TXT_REASON_FL_Change(Index As Integer)
    TXT_REASON_NAME(Index).Text = Gf_ComnNameFind(M_CN1, "G0002", TXT_REASON_FL(Index).Text, 1)
End Sub

Private Sub TXT_REASON_FL_DblClick(Index As Integer)

    TXT_REASON_FL(Index).Text = "DB"
    Call TXT_REASON_FL_KeyUp(Index, vbKeyF4, 0)
    
End Sub

Private Sub TXT_REASON_FL_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
    DD.sWitch = "MS"
    DD.sKey = "G0002"
    DD.rControl.Add Item:=TXT_REASON_FL(Index)

    DD.nameType = "2"
    
    Call Gf_Common_DD(M_CN1, vbKeyF4)
    
    End If

End Sub
Private Sub txt_Scrap_code_Change()
    
    If Len(Trim(txt_Scrap_code)) = txt_Scrap_code.MaxLength Then
        txt_Scrap_name.Text = Gf_ComnNameFind(M_CN1, "G0017", Trim(txt_Scrap_code.Text), 1)
    Else
        txt_Scrap_name.Text = ""
    End If
    
End Sub

Private Sub txt_Scrap_code_DblClick()

    Call txt_Scrap_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_Scrap_code_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
            
        DD.sWitch = "MS"
        DD.sKey = "G0017"
        DD.rControl.Add Item:=txt_Scrap_code
        DD.rControl.Add Item:=txt_Scrap_name
        
        DD.nameType = "1"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    End If

End Sub

'Private Sub TXT_ADDR_DblClick()
'
'    Call txt_addr_KeyUp(vbKeyF4, 0)
'
'End Sub
'
'Private Sub txt_addr_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.rControl.Add Item:=TXT_ADDR
'
'        DD.nameType = "2"
'
'        Call CAR_NO_DD(M_CN1, KeyCode)
'
'    End If
'
'End Sub

Public Function CAR_NO_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String

    DD.DataDicType = "A"        'Apply Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    DD.sQuery = "SELECT YARD_ADDR||LPAD(BED_SEQ,3,'0') YARD_LOCATION, PLATE_NO FROM  GP_PLATEYARD "
    DD.sWhere = " WHERE YARD_ADDR like '" & Trim(DD.rControl.Item(1).Text) & "%' AND SUBSTR(YARD_ADDR,1,2) = 'P4' ORDER BY YARD_ADDR, BED_SEQ"
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
    End If
    
    DD.sWitch = ""
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing
    
End Function

Private Sub CHK_UST_GRD_Click(Index As Integer)
    Dim iNext       As Integer
    
    If sCheck <> "" Then Exit Sub

    sCheck = "**"
    
    If Index = 0 Then
        iNext = 1
    Else
        iNext = 0
    End If
    
    If CHK_UST_GRD(Index).Value = ssCBUnchecked Then
        If CHK_UST_GRD(iNext).Value = ssCBUnchecked Then
            txt_ust_GRD.Text = ""
            CHK_UST_GRD(Index).ForeColor = &H808080
            sCheck = ""
            Exit Sub
        End If
    Else
        CHK_UST_GRD(iNext).Value = ssCBUnchecked
    End If
    
    CHK_UST_GRD(Index).ForeColor = &HFF&
    CHK_UST_GRD(Index).Value = ssCBChecked
                
    CHK_UST_GRD(iNext).ForeColor = &H808080
    CHK_UST_GRD(iNext).Value = ssCBUnchecked

    txt_ust_GRD.Text = CHK_UST_GRD(Index).Tag
    sCheck = ""
    
End Sub

Private Sub CHK_PRD_GRD_Click(Index As Integer)
    Dim iCount      As Integer
    Dim iIndexStr   As Integer
    
    If sCheck <> "" Then Exit Sub

    iCount = 0
    sCheck = "**"
    
    If CHK_PRD_GRD(Index).Value = ssCBUnchecked Then
        For iIndexStr = 0 To 5
            If CHK_PRD_GRD(iIndexStr).Value = ssCBChecked Then
               iCount = iCount + 1
            End If
        Next iIndexStr
        If iCount = 0 Then
            TXT_PRD_GRD.Text = ""
            CHK_PRD_GRD(Index).ForeColor = &H808080
            sCheck = ""
            Exit Sub
        End If
    Else
        For iIndexStr = 0 To 5
            CHK_PRD_GRD(iIndexStr).ForeColor = &H808080
            CHK_PRD_GRD(iIndexStr).Value = ssCBUnchecked
        Next iIndexStr
    End If
    
    CHK_PRD_GRD(Index).ForeColor = &HFF&
    CHK_PRD_GRD(Index).Value = ssCBChecked
    
    TXT_PRD_GRD.Text = CHK_PRD_GRD(Index).Tag
                 
    txt_stdspec_chg.Text = ""
    txt_stdspec_name_chg.Text = ""
    If CHK_PRD_GRD(1).Value = ssCBChecked Or CHK_PRD_GRD(2).Value = ssCBChecked Then
        txt_stdspec_chg.Enabled = True
    Else
        txt_stdspec_chg.Enabled = False
    End If
    
    If CHK_PRD_GRD(4).Value = ssCBChecked Then
        txt_Scrap_code.Enabled = True
    Else
        txt_Scrap_code.Enabled = False
    End If
    
    sCheck = ""
        
End Sub

Private Sub Display_Data_Edit()

    Dim iIndexChk   As Integer
    Dim iIndexStr   As Integer
    
    sCheck = "**"
            
    For iIndexChk = 0 To 1
        If txt_ust_GRD.Text = CHK_UST_GRD(iIndexChk).Tag Then
            CHK_UST_GRD(iIndexChk).ForeColor = &HFF&
            CHK_UST_GRD(iIndexChk).Value = CHECKED
        Else
            CHK_UST_GRD(iIndexChk).ForeColor = &H808080
            CHK_UST_GRD(iIndexChk).Value = UNCHECKED
        End If
    Next iIndexChk

    For iIndexChk = 0 To 5
        If TXT_PRD_GRD.Text = CHK_PRD_GRD(iIndexChk).Tag Then
            CHK_PRD_GRD(iIndexChk).ForeColor = &HFF&
            CHK_PRD_GRD(iIndexChk).Value = CHECKED
        Else
            CHK_PRD_GRD(iIndexChk).ForeColor = &H808080
            CHK_PRD_GRD(iIndexChk).Value = UNCHECKED
        End If
    Next iIndexChk
        
    sCheck = ""
    
    If Left(TXT_PROC_FLAG.Text, 1) <> "Q" Then
        If TXT_INSP_MAN.Text = "" Then TXT_INSP_MAN.Text = sUserID
        If TXT_INSP_OCCR_TIME.RawData = "" Then TXT_INSP_OCCR_TIME.RawData = Gf_DTSet(M_CN1)
        'TXT_INSP_OCCR_TIME.RawData = Gf_DTSet(M_CN1)
    End If
    
    Call Size_Grade_Edit
    
End Sub

Private Sub Size_Grade_Edit()
    Dim sGradeFlag As String
    
    sGradeFlag = ""

    ' THICK GRAND CHECK
    If Val(SDB_THK & "") >= Val(SDB_ORD_THK & "") + Val(SDB_INSP_THK_MN & "") And _
       Val(SDB_THK & "") <= Val(SDB_ORD_THK & "") + Val(SDB_INSP_THK_MX & "") Then
        TXT_INSP_THK_GRD = "Y"
        SDB_THK.ForeColor = &H80000012
    Else
        TXT_INSP_THK_GRD = "N"
        SDB_THK.ForeColor = &HFF&
        sGradeFlag = "N"
    End If
    
    ' WIDTH GRAND CHECK
    If Val(SDB_WID & "") >= Val(SDB_ORD_WID & "") + Val(SDB_INSP_WID_MN & "") And _
       Val(SDB_WID & "") <= Val(SDB_ORD_WID & "") + Val(SDB_INSP_WID_MX & "") Then
        TXT_INSP_WID_GRD = "Y"
        SDB_WID.ForeColor = &H80000012
    Else
        TXT_INSP_WID_GRD = "N"
        SDB_WID.ForeColor = &HFF&
        sGradeFlag = "N"
    End If
        
    ' LENGTH GRAND CHECK
    If Val(SDB_LEN & "") >= Val(SDB_ORD_LEN & "") + Val(SDB_INSP_LEN_MN & "") And _
       Val(SDB_LEN & "") <= Val(SDB_ORD_LEN & "") + Val(SDB_INSP_LEN_MX & "") Then
        TXT_INSP_LEN_GRD = "Y"
        SDB_LEN.ForeColor = &H80000012
    Else
        TXT_INSP_LEN_GRD = "N"
        SDB_LEN.ForeColor = &HFF&
        sGradeFlag = "N"
    End If
    
    ' WEIGHT GRAND CHECK
    If Val(SDB_WGT & "") >= Val(SDB_WGT_ORD & "") + Val(SDB_PWGT_MN & "") And _
       Val(SDB_WGT & "") <= Val(SDB_WGT_ORD & "") + Val(SDB_PWGT_MX & "") Then
        TXT_INSP_WGT_GRD = "Y"
        SDB_WGT.ForeColor = &H80000012
    Else
        TXT_INSP_WGT_GRD = "N"
        SDB_WGT.ForeColor = &HFF&
        sGradeFlag = "N"
    End If
    
    If txt_ust_GRD = "" Then
        CHK_UST_GRD(0).Value = CHECKED
        Call CHK_UST_GRD_Click(0)
    End If
    
'    If TXT_PRD_GRD = "" Then
'        If sGradeFlag = "N" Then
'            CHK_PRD_GRD(1).Value = CHECKED
'            Call CHK_PRD_GRD_Click(1)
'    '        CHK_PRD_GRD(0).Enabled = False
'        Else
'            CHK_PRD_GRD(0).Value = CHECKED
'            Call CHK_PRD_GRD_Click(0)
'    '        CHK_PRD_GRD(0).Enabled = True
'        End If
'    End If
    
End Sub
  
Private Function Cal_Plate_Wgt(sMode As String, dThk As Double, dWid As Double, dLen As Double) As Double

    Dim RS  As New ADODB.Recordset
    
    Cal_Plate_Wgt = 0
    
    sQuery = "SELECT  Gf_Cal_Plate_Wgt('" & sMode & "'" & vbCrLf
    sQuery = sQuery & "             ,'" & Trim(txt_APLY_ENDUSE_CD.Text) & "'" & vbCrLf
    sQuery = sQuery & "             ,'" & Trim(TXT_STLGRD.Text) & "'" & vbCrLf
    sQuery = sQuery & "             ," & dThk & vbCrLf
    sQuery = sQuery & "             ," & dWid & vbCrLf
    sQuery = sQuery & "             ," & dLen & vbCrLf
    sQuery = sQuery & "             ,0 )" & vbCrLf
    sQuery = sQuery & "       FROM  DUAL " & vbCrLf
    RS.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        Cal_Plate_Wgt = Val(RS(0).Value & "")
    End If
    
    RS.Close
    Set RS = Nothing
     
End Function

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       ' Call Gp_Ms_ControlLock(Mc1("pControl"), False)
     End If

End Sub

Public Sub Form_Del()

    TXT_PLATE_NO.Enabled = False
    SDT_PROD_DATE.Enabled = False
    SDT_PROD_TO_DATE.Enabled = False
    CBO_SHIFT.Enabled = False

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
End Sub

'Private Sub ss1_EditChange(ByVal Col As Long, ByVal Row As Long)
'    Dim dThk        As Double
'    Dim dWid        As Double
'    Dim dLen        As Double
'    Dim dWidSum     As Double
'    Dim dLenSum     As Double
'
'    Dim iIdr        As Integer
'    Dim RS          As New adodb.Recordset
'
'    If Col <> 4 And Col <> 5 Then Exit Sub
'
'    ss1.Row = Row
'    dThk = Val(SDB_THK.Value & "")
'    ss1.Col = 4:  dWid = Val(ss1.Text & "")
'    ss1.Col = 5:  dLen = Val(ss1.Text & "")
'
'    ss1.Col = 6
'    ss1.Text = Cal_Plate_Wgt("WGT", dThk, dWid, dLen)
'
'    For iIdr = 1 To ss1.MaxRows - 1
'        ss1.Row = iIdr
'        ss1.Col = 4
'        dWidSum = dWidSum + Val(ss1.Text & "")
'        ss1.Col = 5
'        dLenSum = dLenSum + Val(ss1.Text & "")
'    Next iIdr
'
'    ss1.Row = ss1.MaxRows
'    ss1.Col = 4
'    'dWid = Val(ss1.Text & "") 'SDB_WID.Value - dWidSum
'    dWid = SDB_WID.Value - dWidSum
'    ss1.Col = 5
'    'dLen = Val(ss1.Text & "") 'SDB_LEN.Value - dLenSum
'    dLen = SDB_LEN.Value - dLenSum
'    ss1.Text = dLen
'
'    ss1.Col = 6
'    ss1.Text = Cal_Plate_Wgt("WGT", dThk, dWid, dLen)
'End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim iAddr As String
    Dim iAddr1 As String
    Dim iAddr2 As String
    Dim iTXT_REMARK As String

    If Row < 1 Or SDB_DIVIDE_CNT.Value > 0 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 1
    TXT_PLATE_NO.Text = ss1.Text
    iTXT_REMARK = TXT_REMARK.Text
    
    If Len(TXT_PLATE_NO.Text) = 14 Then
        Call Gp_SSCheck_Cls(MC("sControl"))
        If Gf_Ms_Refer(M_CN1, Mc1, , , True) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Display_Data_Edit
            If txt_NEXT_PROC.Text = "" Or txt_NEXT_PROC.Text = "U" Then
               CHK_NEXT_PRC(0).Value = 1
               txt_NEXT_PROC.Text = "P"
            End If
            iAddr = TXT_ADDR(0).Text
            iAddr1 = TXT_ADDR(1).Text
            iAddr2 = TXT_ADDR(2).Text
            If Len(iAddr) = 3 And Len(iAddr1) = 4 And Len(iAddr2) > 0 Then
               TXT_LOC = iAddr & iAddr1 & Format(Val(iAddr2) + 1, "000")
            End If
            TXT_REMARK.Text = iTXT_REMARK
        End If
    End If
    
End Sub

Private Sub txt_woo_rsn_DblClick()

    Call txt_woo_rsn_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_woo_rsn_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0008"
        DD.rControl.Add Item:=txt_woo_rsn
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

End Sub
