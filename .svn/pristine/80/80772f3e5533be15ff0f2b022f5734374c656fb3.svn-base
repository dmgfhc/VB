VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSTEXT32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB1030C 
   Caption         =   "物料状况详细查询_ACB1030C"
   ClientHeight    =   8595
   ClientLeft      =   195
   ClientTop       =   1590
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_RECV_DATE 
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
      Left            =   7785
      MaxLength       =   19
      TabIndex        =   131
      Tag             =   "到库日期"
      Top             =   7080
      Width           =   1785
   End
   Begin VB.TextBox txt_MOVE_DATE 
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
      Left            =   4560
      MaxLength       =   19
      TabIndex        =   130
      Tag             =   "转库日期"
      Top             =   7080
      Width           =   1785
   End
   Begin VB.TextBox txt_CUR_INV 
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
      Left            =   1425
      MaxLength       =   10
      TabIndex        =   129
      Tag             =   "机号"
      Top             =   6345
      Width           =   1785
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
      Left            =   11310
      TabIndex        =   128
      Tag             =   "钢种"
      Top             =   6345
      Width           =   1170
   End
   Begin VB.TextBox Text_size_knd 
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
      Left            =   10950
      MaxLength       =   2
      TabIndex        =   127
      Tag             =   "钢种"
      Top             =   6345
      Width           =   345
   End
   Begin VB.TextBox txt_STDSPEC_CHG_FL 
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
      Left            =   7785
      MaxLength       =   1
      TabIndex        =   126
      Tag             =   "改板区分"
      Top             =   3365
      Width           =   450
   End
   Begin VB.TextBox txt_BEF_APLY_STDSPEC 
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
      Left            =   7785
      MaxLength       =   18
      TabIndex        =   125
      Tag             =   "原始标准号"
      Top             =   3725
      Width           =   1740
   End
   Begin VB.TextBox txt_DSC_DATE 
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
      Left            =   4560
      MaxLength       =   10
      TabIndex        =   124
      Tag             =   "机号"
      Top             =   6345
      Width           =   1260
   End
   Begin VB.TextBox txt_IN_SHEET_NO 
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
      Left            =   7785
      MaxLength       =   10
      TabIndex        =   123
      Tag             =   "机号"
      Top             =   5975
      Width           =   1335
   End
   Begin VB.TextBox TXT_CUST_CD 
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
      Left            =   4560
      MaxLength       =   50
      TabIndex        =   122
      Tag             =   "机号"
      Top             =   1865
      Width           =   1740
   End
   Begin VB.TextBox T18A1 
      Height          =   270
      Left            =   13680
      TabIndex        =   121
      Top             =   6840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt_MAT_NO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1425
      MaxLength       =   14
      TabIndex        =   120
      Tag             =   "物料号"
      Top             =   735
      Width           =   1875
   End
   Begin VB.TextBox txt_INS_DATE 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
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
      Left            =   13890
      MaxLength       =   10
      TabIndex        =   119
      Tag             =   "机号"
      Top             =   3725
      Width           =   1260
   End
   Begin VB.TextBox txt_WID_GRP 
      Alignment       =   2  'Center
      Height          =   310
      Left            =   1890
      MaxLength       =   1
      TabIndex        =   118
      Top             =   4470
      Width           =   465
   End
   Begin VB.TextBox txt_THK_GRP 
      Alignment       =   2  'Center
      Height          =   310
      Left            =   1425
      MaxLength       =   1
      TabIndex        =   117
      Text            =   " "
      Top             =   4470
      Width           =   465
   End
   Begin VB.TextBox txt_REC_STS 
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
      Left            =   1425
      TabIndex        =   108
      Tag             =   "机号"
      Top             =   1110
      Width           =   1365
   End
   Begin VB.TextBox txt_PROC_CD 
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
      Left            =   1425
      TabIndex        =   107
      Tag             =   "机号"
      Top             =   1860
      Width           =   1365
   End
   Begin VB.TextBox txt_BEF_PROC_CD 
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
      Left            =   1425
      TabIndex        =   106
      Tag             =   "机号"
      Top             =   2235
      Width           =   1365
   End
   Begin VB.TextBox T8 
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
      Left            =   1425
      MaxLength       =   2
      TabIndex        =   105
      Tag             =   "机号"
      Top             =   2610
      Width           =   1365
   End
   Begin VB.TextBox txt_STLGRD 
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
      Left            =   1425
      TabIndex        =   104
      Tag             =   "机号"
      Top             =   2985
      Width           =   1365
   End
   Begin VB.TextBox txt_LOC 
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
      Left            =   1425
      MaxLength       =   10
      TabIndex        =   103
      Tag             =   "机号"
      Top             =   6720
      Width           =   1365
   End
   Begin VB.TextBox txt_BED_PILE_DATE 
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
      Left            =   1425
      MaxLength       =   19
      TabIndex        =   102
      Tag             =   "机号"
      Top             =   7080
      Width           =   1785
   End
   Begin VB.TextBox txt_END_RES 
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
      Left            =   1425
      TabIndex        =   101
      Tag             =   "机号"
      Top             =   1485
      Width           =   1365
   End
   Begin VB.TextBox txt_woo_rsn 
      BackColor       =   &H00C0FFFF&
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
      Left            =   10800
      MaxLength       =   2
      TabIndex        =   100
      Tag             =   "余材原因"
      Top             =   60
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox txt_DSC_TIME 
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
      Left            =   4560
      MaxLength       =   8
      TabIndex        =   99
      Tag             =   "机号"
      Top             =   6725
      Width           =   1260
   End
   Begin VB.TextBox txt_UPD_DATE 
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
      Left            =   13890
      MaxLength       =   10
      TabIndex        =   94
      Tag             =   "机号"
      Top             =   4845
      Width           =   1260
   End
   Begin VB.TextBox txt_PRC 
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
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   81
      Tag             =   "机号"
      Top             =   10200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt_INS_EMP_CD 
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
      Left            =   13890
      MaxLength       =   7
      TabIndex        =   80
      Tag             =   "机号"
      Top             =   4095
      Width           =   1260
   End
   Begin VB.TextBox txt_INS_PGMID 
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
      Left            =   13890
      MaxLength       =   14
      TabIndex        =   79
      Tag             =   "机号"
      Top             =   4470
      Width           =   1260
   End
   Begin VB.TextBox txt_UPD_EMP_CD 
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
      Left            =   13890
      MaxLength       =   7
      TabIndex        =   78
      Tag             =   "机号"
      Top             =   5220
      Width           =   1260
   End
   Begin VB.TextBox txt_UPD_PGM 
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
      Left            =   13890
      MaxLength       =   14
      TabIndex        =   77
      Tag             =   "机号"
      Top             =   5595
      Width           =   1260
   End
   Begin VB.TextBox txt_SHP_EMP 
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
      Left            =   13890
      MaxLength       =   7
      TabIndex        =   76
      Tag             =   "机号"
      Top             =   740
      Width           =   1260
   End
   Begin VB.TextBox txt_HOUSING_DATE 
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
      Left            =   7785
      TabIndex        =   75
      Tag             =   "机号"
      Top             =   6350
      Width           =   1335
   End
   Begin VB.TextBox txt_HOUSING_TIME 
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
      Left            =   7785
      TabIndex        =   74
      Tag             =   "机号"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox txt_OCCR_CD 
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
      Left            =   7785
      MaxLength       =   1
      TabIndex        =   73
      Tag             =   "机号"
      Top             =   1485
      Width           =   450
   End
   Begin VB.ComboBox CBO_NO 
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
      Left            =   3090
      TabIndex        =   1
      Top             =   90
      Width           =   1905
   End
   Begin Threed.SSFrame S2 
      Height          =   1650
      Left            =   0
      TabIndex        =   62
      Top             =   7710
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   2910
      _Version        =   196609
      Caption         =   "特殊信息"
      Begin VB.TextBox txt_ORG_COIL_NO 
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
         Left            =   10800
         MaxLength       =   12
         TabIndex        =   87
         Tag             =   "机号"
         Top             =   720
         Width           =   1380
      End
      Begin VB.TextBox txt_ACT_SMP_FL 
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
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   86
         Tag             =   "机号"
         Top             =   1200
         Width           =   540
      End
      Begin VB.TextBox txt_SURF_GRD_UPD_DATE 
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
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   85
         Tag             =   "机号"
         Top             =   1200
         Width           =   1500
      End
      Begin VB.TextBox txt_COIL_MARKING 
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
         Left            =   7800
         MaxLength       =   1
         TabIndex        =   65
         Tag             =   "机号"
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox txt_COIL_BAND_YN 
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
         Left            =   10800
         MaxLength       =   1
         TabIndex        =   64
         Tag             =   "机号"
         Top             =   240
         Width           =   540
      End
      Begin VB.TextBox txt_COIL_DC_CNT 
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
         Left            =   7800
         MaxLength       =   2
         TabIndex        =   63
         Tag             =   "机号"
         Top             =   720
         Width           =   555
      End
      Begin InDate.ULabel ULabel83 
         Height          =   315
         Left            =   120
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "内径"
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
      Begin InDate.ULabel ULabel84 
         Height          =   315
         Left            =   3120
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "外径"
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
      Begin InDate.ULabel ULabel85 
         Height          =   315
         Left            =   6480
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "标识代码"
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
      Begin InDate.ULabel ULabel86 
         Height          =   315
         Left            =   9480
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "打包代码"
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
      Begin InDate.ULabel ULabel89 
         Height          =   315
         Left            =   6480
         Top             =   720
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "卷取咬入次数"
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
      Begin CSTextLibCtl.sidbEdit txt_INDIA 
         Height          =   315
         Left            =   1440
         TabIndex        =   66
         Top             =   240
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
      Begin CSTextLibCtl.sidbEdit txt_OUTDIA 
         Height          =   315
         Left            =   4440
         TabIndex        =   67
         Top             =   240
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
      Begin CSTextLibCtl.sidbEdit txt_COIL_HD_LEN 
         Height          =   315
         Left            =   1440
         TabIndex        =   68
         Top             =   720
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
      Begin InDate.ULabel ULabel87 
         Height          =   315
         Left            =   120
         Top             =   720
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "头部长度"
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
      Begin InDate.ULabel ULabel88 
         Height          =   315
         Left            =   3120
         Top             =   720
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "尾部长度"
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
      Begin InDate.ULabel ULabel67 
         Height          =   315
         Left            =   6480
         Top             =   1200
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
         Caption         =   "外观等级修改日期"
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
      Begin InDate.ULabel ULabel76 
         Height          =   315
         Left            =   120
         Top             =   1200
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "实际试样代码"
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
      Begin InDate.ULabel ULabel78 
         Height          =   315
         Left            =   3120
         Top             =   1200
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "实际试样长度"
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
      Begin InDate.ULabel ULabel79 
         Height          =   315
         Left            =   9480
         Top             =   720
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "原始钢卷号"
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
      Begin CSTextLibCtl.sidbEdit txt_ACT_SMP_LEN 
         Height          =   315
         Left            =   4440
         TabIndex        =   96
         Top             =   1200
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
      Begin CSTextLibCtl.sidbEdit txt_COIL_TAIL_LEN 
         Height          =   315
         Left            =   4440
         TabIndex        =   97
         Top             =   720
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
   End
   Begin VB.TextBox txt_SHP_IST_CAN_FL 
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
      Left            =   13890
      MaxLength       =   1
      TabIndex        =   55
      Tag             =   "机号"
      Top             =   3360
      Width           =   465
   End
   Begin VB.TextBox txt_SHP_IST_CAN_TIME 
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
      Left            =   13890
      MaxLength       =   8
      TabIndex        =   54
      Tag             =   "机号"
      Top             =   2990
      Width           =   1260
   End
   Begin VB.TextBox txt_SHP_IST_CAN_DATE 
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
      Left            =   13890
      MaxLength       =   10
      TabIndex        =   53
      Tag             =   "机号"
      Top             =   2610
      Width           =   1260
   End
   Begin VB.TextBox txt_CERT_RPT_TIME 
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
      Left            =   13890
      MaxLength       =   8
      TabIndex        =   52
      Tag             =   "机号"
      Top             =   2240
      Width           =   1260
   End
   Begin VB.TextBox txt_CERT_RPT_DATE 
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
      Left            =   13890
      MaxLength       =   10
      TabIndex        =   51
      Tag             =   "机号"
      Top             =   1860
      Width           =   1260
   End
   Begin VB.TextBox txt_CERT_RPT_FL 
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
      Left            =   13890
      MaxLength       =   1
      TabIndex        =   50
      Tag             =   "机号"
      Top             =   1490
      Width           =   465
   End
   Begin VB.TextBox txt_DEST_DETAIL 
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
      Left            =   10950
      MaxLength       =   80
      TabIndex        =   49
      Tag             =   "机号"
      Top             =   5970
      Width           =   4215
   End
   Begin VB.TextBox txt_TRAIN_LINE_NAME 
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
      Left            =   13890
      MaxLength       =   12
      TabIndex        =   48
      Tag             =   "机号"
      Top             =   1115
      Width           =   1260
   End
   Begin VB.TextBox txt_TRNS_CMPY_CD 
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
      Left            =   6975
      MaxLength       =   6
      TabIndex        =   47
      Tag             =   "机号"
      Top             =   8880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox txt_CAR_NO 
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
      Left            =   10950
      TabIndex        =   46
      Tag             =   "机号"
      Top             =   5220
      Width           =   1410
   End
   Begin VB.TextBox txt_TRNS_NO 
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
      Left            =   10950
      TabIndex        =   45
      Tag             =   "机号"
      Top             =   4845
      Width           =   1410
   End
   Begin VB.TextBox txt_OUT_SHEET_NO 
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
      Left            =   10950
      TabIndex        =   44
      Tag             =   "机号"
      Top             =   2990
      Width           =   1410
   End
   Begin VB.TextBox txt_OUT_CAR_NO 
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
      Left            =   10950
      TabIndex        =   43
      Tag             =   "机号"
      Top             =   2615
      Width           =   1410
   End
   Begin VB.TextBox txt_OUT_PLT_TIME 
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
      Left            =   10950
      MaxLength       =   8
      TabIndex        =   42
      Tag             =   "机号"
      Top             =   2240
      Width           =   1410
   End
   Begin VB.TextBox txt_OUT_PLT_DATE 
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
      Left            =   10950
      MaxLength       =   10
      TabIndex        =   41
      Tag             =   "机号"
      Top             =   1865
      Width           =   1410
   End
   Begin VB.TextBox txt_OUT_PLT 
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
      Left            =   10950
      MaxLength       =   2
      TabIndex        =   40
      Tag             =   "机号"
      Top             =   1490
      Width           =   465
   End
   Begin VB.TextBox txt_OUT_PLT_CD 
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
      Left            =   10950
      TabIndex        =   39
      Tag             =   "机号"
      Top             =   1115
      Width           =   1410
   End
   Begin VB.TextBox txt_IN_PLT_CO 
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
      Left            =   10950
      TabIndex        =   38
      Tag             =   "机号"
      Top             =   5595
      Width           =   1410
   End
   Begin VB.TextBox txt_SHP_TIME 
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
      Left            =   10950
      TabIndex        =   37
      Tag             =   "机号"
      Top             =   4470
      Width           =   1410
   End
   Begin VB.TextBox txt_SHP_DATE 
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
      Left            =   10950
      TabIndex        =   36
      Tag             =   "机号"
      Top             =   4095
      Width           =   1410
   End
   Begin VB.TextBox txt_SHP_IST_DATE 
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
      Left            =   10950
      TabIndex        =   35
      Tag             =   "机号"
      Top             =   3725
      Width           =   1410
   End
   Begin VB.TextBox txt_SHP_IST_NO 
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
      Left            =   10950
      TabIndex        =   34
      Tag             =   "机号"
      Top             =   3365
      Width           =   1410
   End
   Begin VB.TextBox txt_IN_CAR_NO 
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
      Left            =   7785
      MaxLength       =   10
      TabIndex        =   33
      Tag             =   "机号"
      Top             =   5600
      Width           =   1335
   End
   Begin VB.TextBox txt_IN_PLT_TIME 
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
      Left            =   7785
      MaxLength       =   8
      TabIndex        =   32
      Tag             =   "机号"
      Top             =   5225
      Width           =   1335
   End
   Begin VB.TextBox txt_IN_PLT_DATE 
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
      Left            =   7785
      MaxLength       =   10
      TabIndex        =   31
      Tag             =   "机号"
      Top             =   4850
      Width           =   1335
   End
   Begin VB.TextBox txt_IN_PLT 
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
      Left            =   7785
      MaxLength       =   2
      TabIndex        =   30
      Tag             =   "机号"
      Top             =   4475
      Width           =   450
   End
   Begin VB.TextBox txt_IN_PLT_CD 
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
      Left            =   7785
      TabIndex        =   29
      Tag             =   "机号"
      Top             =   4100
      Width           =   1335
   End
   Begin VB.TextBox txt_APLY_ENDUSE_CD 
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
      Left            =   4560
      TabIndex        =   28
      Tag             =   "机号"
      Top             =   4100
      Width           =   1740
   End
   Begin VB.TextBox txt_APLY_STDSPEC 
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
      Left            =   4560
      MaxLength       =   18
      TabIndex        =   27
      Tag             =   "机号"
      Top             =   3725
      Width           =   1740
   End
   Begin VB.TextBox txt_QUALITY_UPD_GRD 
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
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   26
      Tag             =   "机号"
      Top             =   3365
      Width           =   465
   End
   Begin VB.TextBox txt_QUALITY_GRD 
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
      Left            =   5040
      MaxLength       =   1
      TabIndex        =   25
      Tag             =   "机号"
      Top             =   2990
      Width           =   465
   End
   Begin VB.TextBox txt_SURF_GRD 
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
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   24
      Tag             =   "机号"
      Top             =   2990
      Width           =   465
   End
   Begin VB.TextBox txt_PROD_GRD 
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
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   23
      Tag             =   "机号"
      Top             =   2615
      Width           =   465
   End
   Begin VB.TextBox txt_INSP_EMP 
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
      Left            =   4560
      MaxLength       =   7
      TabIndex        =   22
      Tag             =   "机号"
      Top             =   5975
      Width           =   1260
   End
   Begin VB.TextBox txt_SMP_NO 
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
      Left            =   4560
      MaxLength       =   14
      TabIndex        =   21
      Tag             =   "机号"
      Top             =   5600
      Width           =   1605
   End
   Begin VB.TextBox txt_SMP_LOC 
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
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   19
      Tag             =   "机号"
      Top             =   4850
      Width           =   465
   End
   Begin VB.TextBox txt_SMP_FL 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
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
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   18
      Tag             =   "机号"
      Top             =   4475
      Width           =   465
   End
   Begin VB.TextBox txt_GROUP_CD 
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
      Left            =   7785
      MaxLength       =   1
      TabIndex        =   17
      Tag             =   "机号"
      Top             =   2990
      Width           =   450
   End
   Begin VB.TextBox txt_SHIFT 
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
      Left            =   7785
      MaxLength       =   1
      TabIndex        =   16
      Tag             =   "机号"
      Top             =   2615
      Width           =   450
   End
   Begin VB.TextBox txt_PROD_TIME 
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
      Left            =   7785
      MaxLength       =   8
      TabIndex        =   15
      Tag             =   "机号"
      Top             =   2240
      Width           =   1335
   End
   Begin VB.TextBox txt_PROD_DATE 
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
      Left            =   7785
      MaxLength       =   10
      TabIndex        =   14
      Tag             =   "机号"
      Top             =   1865
      Width           =   1335
   End
   Begin VB.TextBox txt_DEL_TO_DATE 
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
      Left            =   10950
      TabIndex        =   13
      Tag             =   "机号"
      Top             =   740
      Width           =   1410
   End
   Begin VB.TextBox txt_ENDUSE_CD 
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
      Left            =   9975
      MaxLength       =   4
      TabIndex        =   12
      Tag             =   "机号"
      Top             =   10215
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txt_ORG_ORD_ITEM 
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
      Left            =   5850
      MaxLength       =   2
      TabIndex        =   11
      Tag             =   "机号"
      Top             =   2240
      Width           =   465
   End
   Begin VB.TextBox txt_ORG_ORD_NO 
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
      Left            =   4560
      MaxLength       =   11
      TabIndex        =   10
      Tag             =   "机号"
      Top             =   2240
      Width           =   1305
   End
   Begin VB.TextBox txt_ORD_ITEM 
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
      Left            =   5850
      MaxLength       =   2
      TabIndex        =   9
      Tag             =   "机号"
      Top             =   1490
      Width           =   465
   End
   Begin VB.TextBox txt_ORD_NO 
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
      Left            =   4560
      MaxLength       =   11
      TabIndex        =   8
      Tag             =   "机号"
      Top             =   1490
      Width           =   1305
   End
   Begin VB.TextBox T19A 
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
      Left            =   4560
      TabIndex        =   7
      Tag             =   "机号"
      Top             =   1115
      Width           =   1740
   End
   Begin VB.TextBox t18A 
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
      Left            =   4560
      TabIndex        =   6
      Tag             =   "机号"
      Top             =   740
      Width           =   1740
   End
   Begin VB.TextBox txt_OVER_FL 
      CausesValidation=   0   'False
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
      Left            =   3390
      MaxLength       =   1
      TabIndex        =   5
      Tag             =   "机号"
      Top             =   10200
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt_PRC_LINE 
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
      Left            =   7785
      MaxLength       =   1
      TabIndex        =   3
      Tag             =   "机号"
      Top             =   1115
      Width           =   450
   End
   Begin VB.TextBox txt_plt 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
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
      Left            =   7785
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "机号"
      Top             =   740
      Width           =   450
   End
   Begin VB.TextBox txt_no 
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
      Left            =   1425
      MaxLength       =   15
      TabIndex        =   0
      Tag             =   "物料号"
      Top             =   80
      Width           =   1635
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   75
      Top             =   75
      Width           =   1320
      _ExtentX        =   2328
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   6450
      Top             =   735
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "工厂"
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
      Left            =   6450
      Top             =   1110
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "机号"
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
      Left            =   5895
      Top             =   10200
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      Caption         =   "计算重量"
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
   Begin CSTextLibCtl.sidbEdit txt_CAL_WGT 
      Height          =   315
      Left            =   7110
      TabIndex        =   4
      Top             =   10185
      Visible         =   0   'False
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
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.000"
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   2175
      Top             =   10200
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      Caption         =   "过量生产"
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
      Left            =   3330
      Top             =   735
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "订单/余材"
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
   Begin InDate.ULabel ULabel22 
      Height          =   315
      Left            =   3330
      Top             =   2235
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "原始订单号"
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
   Begin InDate.ULabel ULabel24 
      Height          =   315
      Left            =   6450
      Top             =   7080
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "到库日期"
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
   Begin InDate.ULabel ULabel25 
      Height          =   315
      Left            =   9615
      Top             =   735
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "交货日期"
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
   Begin InDate.ULabel ULabel26 
      Height          =   315
      Left            =   6450
      Top             =   1860
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "生产日期"
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
   Begin InDate.ULabel ULabel27 
      Height          =   315
      Left            =   6450
      Top             =   2235
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "生产时间"
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
   Begin InDate.ULabel ULabel28 
      Height          =   315
      Left            =   6450
      Top             =   2610
      Width           =   1320
      _ExtentX        =   2328
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
   Begin InDate.ULabel ULabel32 
      Height          =   315
      Left            =   3330
      Top             =   4470
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "取样"
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
   Begin InDate.ULabel ULab 
      Height          =   315
      Left            =   3330
      Top             =   4845
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "取样部位"
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
      Left            =   3330
      Top             =   5220
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "取样长度"
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
   Begin CSTextLibCtl.sidbEdit txt_SMP_LEN 
      Height          =   315
      Left            =   4560
      TabIndex        =   20
      Top             =   5220
      Width           =   1260
      _Version        =   262145
      _ExtentX        =   2222
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
   Begin InDate.ULabel ULabel34 
      Height          =   315
      Left            =   3330
      Top             =   5595
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "取样号"
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
   Begin InDate.ULabel ULabel38 
      Height          =   315
      Left            =   3330
      Top             =   2610
      Width           =   1200
      _ExtentX        =   2117
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
   End
   Begin InDate.ULabel ULabel39 
      Height          =   315
      Left            =   3330
      Top             =   2985
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "外观/材质"
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
   Begin InDate.ULabel ULabel41 
      Height          =   315
      Left            =   3330
      Top             =   3360
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "调后等级"
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
   Begin InDate.ULabel ULabel42 
      Height          =   315
      Left            =   3330
      Top             =   3720
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "调后标准号"
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
   Begin InDate.ULabel ULabel43 
      Height          =   315
      Left            =   3330
      Top             =   4095
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "订单用途"
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
   Begin InDate.ULabel ULabel44 
      Height          =   315
      Left            =   6450
      Top             =   4095
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "入库状态"
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
   Begin InDate.ULabel ULabel45 
      Height          =   315
      Left            =   6450
      Top             =   4470
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "入库工厂"
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
   Begin InDate.ULabel ULabel46 
      Height          =   315
      Left            =   6450
      Top             =   4845
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "购买日期"
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
   Begin InDate.ULabel ULabel47 
      Height          =   315
      Left            =   6450
      Top             =   5220
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "购买时间"
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
   Begin InDate.ULabel ULabel48 
      Height          =   315
      Left            =   6450
      Top             =   5595
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "入库车辆号"
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
   Begin InDate.ULabel ULabel59 
      Height          =   315
      Left            =   9615
      Top             =   3360
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "发货指示号"
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
   Begin InDate.ULabel ULabel60 
      Height          =   315
      Left            =   9615
      Top             =   3720
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "发货指示日期"
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
   Begin InDate.ULabel ULabel61 
      Height          =   315
      Left            =   75
      Top             =   735
      Width           =   1320
      _ExtentX        =   2328
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
   Begin InDate.ULabel ULabel62 
      Height          =   315
      Left            =   9615
      Top             =   4095
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "发货日期"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel63 
      Height          =   315
      Left            =   9615
      Top             =   4470
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "发货时间"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel50 
      Height          =   315
      Left            =   9615
      Top             =   5595
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "供货商"
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
   Begin InDate.ULabel ULabel51 
      Height          =   315
      Left            =   9615
      Top             =   1110
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "出库状态"
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
   Begin InDate.ULabel ULabel52 
      Height          =   315
      Left            =   9615
      Top             =   1485
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "出库工厂"
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
   Begin InDate.ULabel ULabel53 
      Height          =   315
      Left            =   9615
      Top             =   1860
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "出库日期"
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
      Left            =   9615
      Top             =   2235
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "出库时间"
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
   Begin InDate.ULabel ULabel55 
      Height          =   315
      Left            =   9615
      Top             =   2610
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "出库车辆号"
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
   Begin InDate.ULabel ULabel56 
      Height          =   315
      Left            =   9615
      Top             =   2985
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "出库提货单号"
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
   Begin InDate.ULabel ULabel64 
      Height          =   315
      Left            =   9615
      Top             =   4845
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "提货单号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel65 
      Height          =   315
      Left            =   9615
      Top             =   5220
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "车辆号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel60sf 
      Height          =   315
      Left            =   5775
      Top             =   8895
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "运输公司"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel61sf 
      Height          =   315
      Left            =   12555
      Top             =   1110
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "专线名称"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
      Left            =   9615
      Top             =   5970
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "详细目的地"
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
   Begin InDate.ULabel ULabel63sf 
      Height          =   315
      Left            =   12555
      Top             =   1485
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "质保书发放"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel64f 
      Height          =   315
      Left            =   12555
      Top             =   1860
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "发放日期"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel65f 
      Height          =   315
      Left            =   12555
      Top             =   2235
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "发放时间"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel67f 
      Height          =   315
      Left            =   12555
      Top             =   2610
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "取消发货日期"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel68f 
      Height          =   315
      Left            =   12555
      Top             =   2985
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "取消发货时间"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel76f 
      Height          =   315
      Left            =   6450
      Top             =   1485
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "信息来源"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel57 
      Height          =   315
      Left            =   6450
      Top             =   6345
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "入库日期"
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
   Begin InDate.ULabel ULabel58 
      Height          =   315
      Left            =   6450
      Top             =   6720
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "入库时间"
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
   Begin InDate.ULabel ULabel91sf 
      Height          =   315
      Left            =   12555
      Top             =   735
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "发货人员"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel93f 
      Height          =   315
      Left            =   12555
      Top             =   4095
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "录入人员"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel94f 
      Height          =   315
      Left            =   12555
      Top             =   4470
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "录入程序"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel95f 
      Height          =   315
      Left            =   12555
      Top             =   4845
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "修改日期"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel96f 
      Height          =   315
      Left            =   12555
      Top             =   5220
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "修改人员"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel97f 
      Height          =   315
      Left            =   12555
      Top             =   5595
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "修改程序"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel82A 
      Height          =   315
      Left            =   4020
      Top             =   10200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "作业状态"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
      Left            =   6450
      Top             =   2985
      Width           =   1320
      _ExtentX        =   2328
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
   Begin Threed.SSCommand cmd_fl_down 
      Height          =   420
      Left            =   11640
      TabIndex        =   98
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
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
      Caption         =   "余材降级"
   End
   Begin InDate.ULabel ULabel35 
      Height          =   315
      Left            =   3330
      Top             =   5970
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "检查人员"
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
      Left            =   3330
      Top             =   6345
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "综判日期"
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
      Left            =   3330
      Top             =   6720
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "综判时间"
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
   Begin InDate.ULabel ULabel77 
      Height          =   315
      Left            =   9480
      Top             =   60
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "余材原因"
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
      Left            =   75
      Top             =   1860
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "当前进程"
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
      Left            =   75
      Top             =   2235
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "前进程"
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
      Left            =   75
      Top             =   2610
      Width           =   1320
      _ExtentX        =   2328
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
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   75
      Top             =   2985
      Width           =   1320
      _ExtentX        =   2328
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
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   75
      Top             =   3360
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "厚度/宽度"
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
   Begin CSTextLibCtl.sidbEdit txt_THK 
      Height          =   315
      Left            =   1425
      TabIndex        =   109
      Top             =   3360
      Width           =   720
      _Version        =   262145
      _ExtentX        =   1270
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
      Modified        =   -1  'True
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
   Begin CSTextLibCtl.sidbEdit txt_WID 
      Height          =   315
      Left            =   2145
      TabIndex        =   110
      Top             =   3360
      Width           =   660
      _Version        =   262145
      _ExtentX        =   1164
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
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_LEN 
      Height          =   315
      Left            =   1425
      TabIndex        =   111
      Top             =   3720
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
   Begin CSTextLibCtl.sidbEdit txt_WGT 
      Height          =   315
      Left            =   1425
      TabIndex        =   112
      Top             =   4095
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel30 
      Height          =   315
      Left            =   75
      Top             =   6720
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "垛位号"
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
      Left            =   75
      Top             =   3720
      Width           =   1320
      _ExtentX        =   2328
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Left            =   75
      Top             =   4095
      Width           =   1320
      _ExtentX        =   2328
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
   Begin InDate.ULabel ULabel98 
      Height          =   315
      Left            =   75
      Top             =   1485
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "产品结束处理"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULABEL_ORD_THK 
      Height          =   315
      Left            =   75
      Top             =   4845
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "订单厚度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel99 
      Height          =   315
      Left            =   75
      Top             =   5220
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "订单宽度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel101 
      Height          =   315
      Left            =   75
      Top             =   5595
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "订单长度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel102 
      Height          =   315
      Left            =   75
      Top             =   5970
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "订单重量"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin CSTextLibCtl.sidbEdit txt_ORD_WGT 
      Height          =   315
      Left            =   1425
      TabIndex        =   113
      Top             =   5970
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ORD_THK 
      Height          =   315
      Left            =   1425
      TabIndex        =   114
      Top             =   4845
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
   Begin CSTextLibCtl.sidbEdit txt_ORD_WID 
      Height          =   315
      Left            =   1425
      TabIndex        =   115
      Top             =   5220
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
   Begin CSTextLibCtl.sidbEdit txt_ORD_LEN 
      Height          =   315
      Left            =   1425
      TabIndex        =   116
      Top             =   5595
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
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   75
      Top             =   4470
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "厚/宽度组"
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
   Begin InDate.ULabel ULabel69f 
      Height          =   315
      Left            =   12555
      Top             =   3360
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "取消发货指示"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   1
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
   Begin InDate.ULabel ULabel92f 
      Height          =   315
      Left            =   12555
      Top             =   3720
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "录入日期"
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
   Begin Threed.SSFrame S3 
      Height          =   1245
      Left            =   0
      TabIndex        =   69
      Top             =   7710
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   2196
      _Version        =   196609
      Font3D          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "特殊信息"
      Begin VB.TextBox txt_ACT_SMP_FL_PLATE 
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
         Left            =   13680
         MaxLength       =   1
         TabIndex        =   93
         Tag             =   "机号"
         Top             =   240
         Width           =   540
      End
      Begin VB.TextBox txt_NEXT_PROC 
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
         Left            =   10755
         MaxLength       =   1
         TabIndex        =   92
         Tag             =   "机号"
         Top             =   720
         Width           =   540
      End
      Begin VB.TextBox txt_ORG_PLATE 
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
         Left            =   10755
         MaxLength       =   14
         TabIndex        =   91
         Tag             =   "机号"
         Top             =   240
         Width           =   540
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
         Left            =   7800
         MaxLength       =   1
         TabIndex        =   90
         Tag             =   "机号"
         Top             =   720
         Width           =   660
      End
      Begin VB.TextBox txt_UST_FL 
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
         Left            =   4380
         MaxLength       =   4
         TabIndex        =   89
         Tag             =   "机号"
         Top             =   720
         Width           =   570
      End
      Begin VB.TextBox txt_SF_ORNOT_PLATE 
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
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   88
         Tag             =   "机号"
         Top             =   720
         Width           =   570
      End
      Begin VB.TextBox txt_PILE_NO 
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
         Left            =   7800
         MaxLength       =   6
         TabIndex        =   72
         Tag             =   "机号"
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txt_PLATE_SEC 
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
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   71
         Tag             =   "机号"
         Top             =   240
         Width           =   570
      End
      Begin VB.TextBox txt_CR_CD 
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
         Left            =   4380
         MaxLength       =   1
         TabIndex        =   70
         Tag             =   "机号"
         Top             =   240
         Width           =   570
      End
      Begin InDate.ULabel ULabel81 
         Height          =   315
         Left            =   120
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "母板/钢板"
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
      Begin InDate.ULabel ULabel82 
         Height          =   315
         Left            =   3150
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "是否控轧"
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
      Begin InDate.ULabel ULabel75 
         Height          =   315
         Left            =   6480
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "堆垛号"
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
      Begin InDate.ULabel ULabel80 
         Height          =   315
         Left            =   120
         Top             =   720
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "是否修磨"
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
      Begin InDate.ULabel ULabel90 
         Height          =   315
         Left            =   3150
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "UST代码"
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
      Begin InDate.ULabel ULabel91 
         Height          =   315
         Left            =   6480
         Top             =   720
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "切边代码"
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
      Begin InDate.ULabel ULabel92 
         Height          =   315
         Left            =   9450
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "原始钢板号"
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
      Begin InDate.ULabel ULabel93 
         Height          =   315
         Left            =   9450
         Top             =   720
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "后续工序"
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
      Begin InDate.ULabel ULabel94 
         Height          =   315
         Left            =   12360
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "实际试样代码"
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
      Begin InDate.ULabel ULabel95 
         Height          =   315
         Left            =   12360
         Top             =   720
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "实际试样长度"
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
      Begin CSTextLibCtl.sidbEdit txt_ACT_SMP_LEN_PLATE 
         Height          =   315
         Left            =   13680
         TabIndex        =   95
         Top             =   720
         Width           =   1140
         _Version        =   262145
         _ExtentX        =   2011
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
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
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
         NumDecDigits    =   1
         NumIntDigits    =   7
         Undo            =   0
         Data            =   0
      End
   End
   Begin Threed.SSFrame S1 
      Height          =   1335
      Left            =   0
      TabIndex        =   56
      Top             =   7710
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   2355
      _Version        =   196609
      Caption         =   "特殊信息"
      Begin VB.TextBox txt_quality_id 
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
         Left            =   7800
         MaxLength       =   1
         TabIndex        =   84
         Tag             =   "机号"
         Top             =   840
         Width           =   540
      End
      Begin VB.TextBox txt_MOTHER_SLAB 
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
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   83
         Tag             =   "机号"
         Top             =   840
         Width           =   1515
      End
      Begin VB.TextBox txt_SCR_ORNOT 
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
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   82
         Tag             =   "机号"
         Top             =   840
         Width           =   660
      End
      Begin VB.TextBox txt_SLAB_RHF_IN_DATE 
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
         Left            =   10800
         MaxLength       =   20
         TabIndex        =   61
         Tag             =   "机号"
         Top             =   840
         Width           =   2385
      End
      Begin VB.TextBox txt_SF_ORNOT 
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
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   60
         Tag             =   "机号"
         Top             =   240
         Width           =   660
      End
      Begin VB.TextBox txt_RHF_REJ_ORNOT 
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
         Left            =   4440
         MaxLength       =   1
         TabIndex        =   59
         Tag             =   "机号"
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txt_AIM_HCR_KND 
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
         Left            =   7800
         MaxLength       =   1
         TabIndex        =   58
         Tag             =   "机号"
         Top             =   240
         Width           =   540
      End
      Begin VB.TextBox txt_HCR_KND 
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
         Left            =   10680
         MaxLength       =   1
         TabIndex        =   57
         Tag             =   "机号"
         Top             =   240
         Width           =   540
      End
      Begin InDate.ULabel ULabel70 
         Height          =   315
         Left            =   9480
         Top             =   840
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "板坯装炉时间"
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
      Begin InDate.ULabel ULabel71 
         Height          =   315
         Left            =   3120
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "炉内是否缺号"
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
      Begin InDate.ULabel ULabel72 
         Height          =   315
         Left            =   6480
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "目标板坯去向"
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
      Begin InDate.ULabel ULabel73 
         Height          =   315
         Left            =   9480
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "实际板坯去向"
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
      Begin InDate.ULabel ULabel66 
         Height          =   315
         Left            =   120
         Top             =   840
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "废钢代码"
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
      Begin InDate.ULabel MOTHER_SLAB 
         Height          =   315
         Left            =   3120
         Top             =   840
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "母板坯号"
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
      Begin InDate.ULabel ULabel68 
         Height          =   315
         Left            =   6480
         Top             =   840
         Width           =   1320
         _ExtentX        =   2328
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
      Begin InDate.ULabel ULabel66f 
         Height          =   315
         Left            =   120
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "是否修磨"
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
   End
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Left            =   3330
      Top             =   1110
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "余材原因"
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
      Left            =   3330
      Top             =   1485
      Width           =   1200
      _ExtentX        =   2117
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
   End
   Begin InDate.ULabel ULabel40 
      Height          =   315
      Left            =   3330
      Top             =   1860
      Width           =   1200
      _ExtentX        =   2117
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
   End
   Begin InDate.ULabel ULabel49 
      Height          =   315
      Left            =   6450
      Top             =   5970
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "入库提货单号"
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
   Begin InDate.ULabel ULabel15 
      Height          =   315
      Left            =   6450
      Top             =   3360
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "改板区分"
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
   Begin InDate.ULabel ULabel21 
      Height          =   315
      Left            =   6450
      Top             =   3720
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "原始标准号"
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
      Left            =   75
      Top             =   1110
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "物料状态"
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
      Left            =   75
      Top             =   7080
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "堆垛日期"
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
   Begin InDate.ULabel ULabel23 
      Height          =   315
      Left            =   9615
      Top             =   6345
      Width           =   1320
      _ExtentX        =   2328
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
   End
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   75
      Top             =   6345
      Width           =   1320
      _ExtentX        =   2328
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
   End
   Begin InDate.ULabel ULabel69 
      Height          =   315
      Left            =   3330
      Top             =   7080
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "转库日期"
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
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   15120
      Y1              =   7575
      Y2              =   7575
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   15135
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Line Line2 
      X1              =   15435
      X2              =   15390
      Y1              =   135
      Y2              =   8190
   End
End
Attribute VB_Name = "ACB1030C"
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
'-- Program ID        ACB1030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          APPLE
'-- Coder             APPLE
'-- Date              2003.8.14
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
Public sqlstring As String

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
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2
Public FORM_A As String
Dim WULIAO  As String


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(txt_no, "p", "n", "", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    S1.BackColor = &HE0E0E0
    S2.BackColor = &HE0E0E0
    S3.BackColor = &HE0E0E0
        
End Sub

Private Sub CBO_NO_Click()
    txt_no.Text = Trim(CBO_NO.Text)
    Call Form_Ref
End Sub

Private Sub cmd_fl_down_Click()

'On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command
        
    If Trim(txt_mat_no.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_mat_no.Tag + " Must input necessarily")
        Exit Sub
    End If
    
    If Trim(txt_woo_rsn.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_woo_rsn.Tag + " Must input necessarily")
        Exit Sub
    End If
    
    If Len(Trim(txt_woo_rsn.Text)) <> txt_woo_rsn.MaxLength Then
        Call Gp_MsgBoxDisplay(txt_woo_rsn.Tag + " Must input according to length of item")
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
     
    'COIL
    If WULIAO = "HC" Then
        sQuery = "{call ACE2010P ('" + txt_mat_no.Text + "','" + txt_woo_rsn.Text + "','Y','" + sUserID + "',?)}"
    'PLATE
    ElseIf WULIAO = "PP" Then
        sQuery = "{call ACE2020P ('" + txt_mat_no.Text + "','" + txt_woo_rsn.Text + "','Y','" + sUserID + "',?)}"
    Else
    'SLAB
        sQuery = "{call ACE2030P ('" + txt_mat_no.Text + "','" + txt_woo_rsn.Text + "','Y','" + sUserID + "',?)}"
    End If
                
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
        Call Gp_MsgBoxDisplay("余材降级完了..!!", "I")
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

Private Sub Form_Activate()
    
    Dim I As Integer
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    For I = 3 To 16
        MDIMain.MenuTool.Buttons(I).Enabled = False
    Next I
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN And Trim(WULIAO) <> "" Then
'        KeyAscii = 0
'        SendKeys "{TAB}"
        Call Form_Ref
    End If
    
End Sub

Private Sub Form_Load()


    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Dim I As Integer
    
    For I = 3 To 16
        MDIMain.MenuTool.Buttons(I).Enabled = False
    Next I
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Screen.MousePointer = vbDefault
    
    S1.Visible = False
    S2.Visible = False
    S3.Visible = False
    
    If FORM_A = "ACB1020C" Then
        If ACB1020C.AIMNO <> "" Then
            txt_no.Text = Trim(ACB1020C.AIMNO)
            Call Form_Ref
        End If
    End If
    
    If FORM_A = "ACB1025C" Then
        If ACB1025C.AIMNO <> "" Then
            txt_no.Text = Trim(ACB1025C.AIMNO)
            Call Form_Ref
        End If
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
     Dim I As Integer
    
    For I = 3 To 16
        MDIMain.MenuTool.Buttons(I).Enabled = False
    Next I
    
    CBO_NO.Clear
    ACB1020C.AIMNO = ""
    ACB1020C.STR1 = ""
    ACB1020C.BASE = ""
    
End Sub

Public Sub Form_Cls()
    
    Dim I As Integer
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
    For I = 3 To 16
        MDIMain.MenuTool.Buttons(I).Enabled = False
    Next I
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    If CBO_NO.ListCount = 0 Then
        rControl(1).SetFocus
        txt_no.Text = ""
    Else
        txt_no.Text = CBO_NO.Text
    End If
    
    Call ALL_ITEMCLS

End Sub

Public Sub Form_Ref()

    Dim sQuery As String
    Dim sMesg As String
        
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
    
        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then

        Else
            sMesg = sMesg + " Must input according to length of item"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        End If
    
    Else
        sMesg = sMesg + " Must input necessarily"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
        
    End If
    
    Call ALL_ITEMCLS
    Call txt_no_Change

    Call SELECT_PRC

End Sub

Public Sub Form_Pro()

    Dim I As Integer
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        For I = 3 To 16
            MDIMain.MenuTool.Buttons(I).Enabled = False
        Next I
        
    End If
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

'-------------------------------------I MADE THE IMPORMENT PROGRAM --------------------------------------------

'------------------------------------------------*******---------------------------------------
'--------------------------------------------------***-----------------------------------------

Private Sub SELECT_PRC()

    Dim select_no   As String
    Dim sQuery1     As String
    Dim sMesg       As String
    Dim sCurInv     As String
    
    select_no = txt_no.Text

    If WULIAO = "SL" Then
        sQuery1 = " SELECT * FROM FP_SLAB WHERE SLAB_NO    = '" + select_no + "'"
        S1.Visible = True
    ElseIf WULIAO = "HC" Then
        sQuery1 = " SELECT * FROM GP_COIL WHERE COIL_NO    =  '" + select_no + "'"
        S2.Visible = True
    ElseIf WULIAO = "PP" Then
        sQuery1 = " SELECT * FROM GP_PLATE WHERE PLATE_NO  =  '" + select_no + "'"
        S3.Visible = True
    Else
        Call MsgBox("物料号" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
        Exit Sub
    End If
    
    Dim AdoRs As ADODB.Recordset
    
    Set AdoRs = New ADODB.Recordset
       
    AdoRs.Open sQuery1, M_CN1, adOpenKeyset
    If AdoRs.RecordCount = 0 Then
        MDIMain.StatusBar1.Panels(1) = "提示信息: 没有资料"
'        sMesg = " there are no any data,please choose the condition again"
'        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
  '  Text4.Text = squery1
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        txt_mat_no.Text = Trim(AdoRs.Fields(0) & "")
        txt_OCCR_CD.Text = Trim(AdoRs.Fields(1) & "")
        txt_PLT.Text = Trim(AdoRs.Fields(2) & "")
        txt_PRC_line.Text = Trim(AdoRs.Fields(3) & "")
        txt_REC_STS.Text = Gf_ComnNameFind(M_CN1, "Z0005", Trim(AdoRs.Fields(4) & ""), 2)
        txt_PROC_CD.Text = Trim(AdoRs.Fields(5) & "")
        txt_BEF_PROC_CD.Text = Trim(AdoRs.Fields(6) & "")
        T8.Text = Trim(AdoRs.Fields(7) & "")
        txt_stlgrd.Text = Trim(AdoRs.Fields(8) & "")
        txt_THK.RawData = Trim(AdoRs.Fields(9) & "")
        txt_WID.RawData = Trim(AdoRs.Fields(10) & "")
        txt_LEN.RawData = Trim(AdoRs.Fields(11) & "")
        txt_WGT.RawData = Trim(AdoRs.Fields(12) & "")
        txt_CAL_WGT.RawData = Trim(AdoRs.Fields(13) & "")
        txt_THK_GRP.Text = Trim(AdoRs.Fields(14) & "")
        txt_WID_GRP.Text = Trim(AdoRs.Fields(15) & "")
        txt_OVER_FL.Text = Trim(AdoRs.Fields(16) & "")
        t18A.Text = Trim(AdoRs.Fields(17) & "")
        T18A1.Text = Trim(AdoRs.Fields(17) & "")
        T19A.Text = Trim(AdoRs.Fields(18) & "")
        txt_woo_rsn.Text = Trim(AdoRs.Fields(18) & "")
        txt_ord_no.Text = Trim(AdoRs.Fields(19) & "")
        TXT_ORD_ITEM.Text = Trim(AdoRs.Fields(20) & "")
        txt_ORG_ORD_NO.Text = Trim(AdoRs.Fields(21) & "")
        txt_ORG_ORD_ITEM.Text = Trim(AdoRs.Fields(22) & "")
        txt_enduse_cd.Text = Trim(AdoRs.Fields(23) & "")
        txt_del_to_date.Text = DATESET(AdoRs.Fields(24) & "")
        txt_PROD_DATE.Text = DATESET(AdoRs.Fields(25) & "")
        txt_PROD_TIME.Text = TIMESET(AdoRs.Fields(26) & "")
        txt_SHIFT.Text = Trim(AdoRs.Fields(27) & "")
        txt_GROUP_CD.Text = Trim(AdoRs.Fields(28) & "")
        txt_LOC.Text = Trim(AdoRs.Fields(29) & "")
        txt_BED_PILE_DATE.Text = ALLSET(AdoRs.Fields(30) & "")
        txt_SMP_FL.Text = Trim(AdoRs.Fields(31) & "")
        txt_SMP_LOC.Text = Trim(AdoRs.Fields(32) & "")
        txt_SMP_LEN.RawData = Trim(AdoRs.Fields(33) & "")
        txt_SMP_NO.Text = Trim(AdoRs.Fields(34) & "")
        txt_INSP_EMP.Text = Trim(AdoRs.Fields(35) & "")
        txt_DSC_DATE.Text = DATESET(AdoRs.Fields(36) & "")
        txt_DSC_TIME.Text = TIMESET(AdoRs.Fields(37) & "")
        txt_prod_grd.Text = Trim(AdoRs.Fields(38) & "")
        txt_SURF_GRD.Text = Trim(AdoRs.Fields(39) & "")
        txt_QUALITY_GRD.Text = Trim(AdoRs.Fields(40) & "")
        txt_QUALITY_UPD_GRD.Text = Trim(AdoRs.Fields(41) & "")
        txt_APLY_STDSPEC.Text = Trim(AdoRs.Fields(42) & "")
        txt_APLY_ENDUSE_CD.Text = Trim(AdoRs.Fields(43) & "")
        txt_IN_PLT_CD.Text = Trim(AdoRs.Fields(44) & "")
        txt_IN_PLT.Text = Trim(AdoRs.Fields(45) & "")
        txt_IN_PLT_DATE.Text = DATESET(AdoRs.Fields(46) & "")
        txt_IN_PLT_TIME.Text = TIMESET(AdoRs.Fields(47) & "")
        txt_IN_CAR_NO.Text = Trim(AdoRs.Fields(48) & "")
        txt_IN_SHEET_NO.Text = Trim(AdoRs.Fields(49) & "")
        txt_IN_PLT_CO.Text = Trim(AdoRs.Fields(50) & "")
        txt_out_plt_cd.Text = Trim(AdoRs.Fields(51) & "")
        txt_OUT_PLT.Text = Trim(AdoRs.Fields(52) & "")
        txt_OUT_PLT_DATE.Text = DATESET(AdoRs.Fields(53) & "")
        txt_OUT_PLT_TIME.Text = TIMESET(AdoRs.Fields(54) & "")
        txt_OUT_CAR_NO.Text = Trim(AdoRs.Fields(55) & "")
        txt_OUT_SHEET_NO.Text = Trim(AdoRs.Fields(56) & "")
        txt_HOUSING_DATE.Text = DATESET(AdoRs.Fields(57) & "")
        txt_HOUSING_TIME.Text = TIMESET(AdoRs.Fields(58) & "")
        txt_SHP_IST_NO.Text = Trim(AdoRs.Fields(59) & "")
        txt_SHP_IST_DATE.Text = DATESET(AdoRs.Fields(60) & "")
        txt_SHP_DATE.Text = DATESET(AdoRs.Fields(62) & "")
        txt_SHP_TIME.Text = TIMESET(AdoRs.Fields(63) & "")
        txt_TRNS_NO.Text = Trim(AdoRs.Fields(64) & "")
        txt_car_no.Text = Trim(AdoRs.Fields(65) & "")
        txt_TRNS_CMPY_CD.Text = Trim(AdoRs.Fields(66) & "")
        txt_SHP_EMP.Text = Trim(AdoRs.Fields(67) & "")
        txt_TRAIN_LINE_NAME.Text = Trim(AdoRs.Fields(68) & "")
        txt_dest_detail.Text = Trim(AdoRs.Fields(69) & "")
        txt_CERT_RPT_FL.Text = Trim(AdoRs.Fields(70) & "")
        txt_CERT_RPT_DATE.Text = DATESET(AdoRs.Fields(71) & "")
        txt_CERT_RPT_TIME.Text = TIMESET(AdoRs.Fields(72) & "")
        txt_SHP_IST_CAN_DATE.Text = DATESET(AdoRs.Fields(73) & "")
        txt_SHP_IST_CAN_TIME.Text = TIMESET(AdoRs.Fields(74) & "")
        txt_SHP_IST_CAN_FL.Text = Trim(AdoRs.Fields(75) & "")
        txt_INS_DATE.Text = DATESET(AdoRs.Fields(76) & "")
        txt_INS_EMP_CD.Text = Trim(AdoRs.Fields(77) & "")
        txt_INS_PGMID.Text = Trim(AdoRs.Fields(78) & "")
        txt_UPD_DATE.Text = DATESET(AdoRs.Fields(79) & "")
        txt_UPD_EMP_CD.Text = Trim(AdoRs.Fields(80) & "")
        txt_UPD_PGM.Text = Trim(AdoRs.Fields(81) & "")
        
        
        If WULIAO = "SL" Then
        
            txt_SF_ORNOT.Text = Trim(AdoRs.Fields(82) & "")
            txt_SLAB_RHF_IN_DATE.Text = ALLSET(AdoRs.Fields(83) & "")
            txt_RHF_REJ_ORNOT.Text = Trim(AdoRs.Fields(84) & "")
            txt_AIM_HCR_KND.Text = Trim(AdoRs.Fields(85) & "")
            txt_HCR_KND.Text = Trim(AdoRs.Fields(86) & "")
            txt_PRC.Text = Trim(AdoRs.Fields(87) & "")
            txt_SCR_ORNOT.Text = Trim(AdoRs.Fields(88) & "")
            txt_END_RES.Text = Trim(AdoRs.Fields(89) & "")
            txt_MOTHER_SLAB.Text = Trim(AdoRs.Fields(90) & "")
            txt_ORD_THK.Text = Trim(AdoRs.Fields(91) & "")
            txt_ORD_WID.Text = Trim(AdoRs.Fields(92) & "")
            txt_ORD_LEN.Text = Trim(AdoRs.Fields(93) & "")
            txt_ORD_WGT.Text = Trim(AdoRs.Fields(94) & "")
            txt_quality_id.Text = Trim(AdoRs.Fields(95) & "")
            TXT_CUST_CD.Text = Trim(AdoRs.Fields(99) & "")
            txt_CUR_INV.Text = Gf_ComnNameFind(M_CN1, "C0013", Trim(AdoRs.Fields(104) & ""), 2)
            sCurInv = Trim(AdoRs.Fields(104) & "")
            Text_size_knd.Text = Trim(AdoRs.Fields(105) & "")
                    
        ElseIf WULIAO = "HC" Then
        
            txt_INDIA.RawData = Trim(AdoRs.Fields(82) & "")
            txt_OUTDIA.RawData = Trim(AdoRs.Fields(83) & "")
            txt_COIL_MARKING.Text = Trim(AdoRs.Fields(84) & "")
            txt_COIL_BAND_YN.Text = Trim(AdoRs.Fields(85) & "")
            txt_COIL_HD_LEN.Text = Trim(AdoRs.Fields(86) & "")
            txt_COIL_TAIL_LEN.RawData = Trim(AdoRs.Fields(87) & "")
            txt_COIL_DC_CNT.Text = Trim(AdoRs.Fields(88) & "")
            txt_PRC.Text = Trim(AdoRs.Fields(89) & "")
            txt_SURF_GRD_UPD_DATE.Text = DATESET(AdoRs.Fields(90) & "")
            txt_END_RES.Text = Trim(AdoRs.Fields(91) & "")
            txt_ORD_THK.Text = Trim(AdoRs.Fields(92) & "")
            txt_ORD_WID.Text = Trim(AdoRs.Fields(93) & "")
            txt_ORD_LEN.Text = Trim(AdoRs.Fields(94) & "")
            txt_ORD_WGT.Text = Trim(AdoRs.Fields(95) & "")
            txt_ACT_SMP_FL.Text = Trim(AdoRs.Fields(97) & "")
            txt_ACT_SMP_LEN.Text = Trim(AdoRs.Fields(98) & "")
            txt_ORG_COIL_NO.Text = Trim(AdoRs.Fields(99) & "")
            TXT_CUST_CD.Text = Trim(AdoRs.Fields(100) & "")
            txt_CUR_INV.Text = Gf_ComnNameFind(M_CN1, "C0013", Trim(AdoRs.Fields(101) & ""), 2)
            sCurInv = Trim(AdoRs.Fields(101) & "")
            Text_size_knd.Text = Trim(AdoRs.Fields(102) & "")
            txt_BEF_APLY_STDSPEC.Text = Trim(AdoRs.Fields(105) & "")
            txt_STDSPEC_CHG_FL.Text = Trim(AdoRs.Fields(106) & "")
            
        ElseIf WULIAO = "PP" Then
        
            txt_PLATE_SEC.Text = Trim(AdoRs.Fields(82) & "")
            txt_CR_CD.Text = Trim(AdoRs.Fields(83) & "")
            txt_PILE_NO.Text = Trim(AdoRs.Fields(84) & "")
            txt_PRC.Text = Trim(AdoRs.Fields(85) & "")
            txt_SF_ORNOT_PLATE.Text = Trim(AdoRs.Fields(86) & "")
            txt_UST_FL.Text = Trim(AdoRs.Fields(87) & "")
            txt_TRIM_FL.Text = Trim(AdoRs.Fields(88) & "")
            txt_END_RES.Text = Trim(AdoRs.Fields(89) & "")
            txt_ORG_PLATE.Text = Trim(AdoRs.Fields(90) & "")
            txt_ORD_THK.Text = Trim(AdoRs.Fields(91) & "")
            txt_ORD_WID.Text = Trim(AdoRs.Fields(92) & "")
            txt_ORD_LEN.Text = Trim(AdoRs.Fields(93) & "")
            txt_ORD_WGT.Text = Trim(AdoRs.Fields(94) & "")
            txt_NEXT_PROC.Text = Trim(AdoRs.Fields(95) & "")
            txt_ACT_SMP_FL_PLATE.Text = Trim(AdoRs.Fields(96) & "")
            txt_ACT_SMP_LEN_PLATE.Text = Trim(AdoRs.Fields(97) & "")
            TXT_CUST_CD.Text = Trim(AdoRs.Fields(98) & "")
            txt_CUR_INV.Text = Gf_ComnNameFind(M_CN1, "C0013", Trim(AdoRs.Fields(99) & ""), 2)
            sCurInv = Trim(AdoRs.Fields(99) & "")
            txt_BEF_APLY_STDSPEC.Text = Trim(AdoRs.Fields(100) & "")
            txt_STDSPEC_CHG_FL.Text = Trim(AdoRs.Fields(101) & "")
            Text_size_knd.Text = Trim(AdoRs.Fields(102) & "")
        End If
                
        '代码转化成中文
        
        Call CONVERT_A("C0011", txt_END_RES, txt_END_RES)   '83
        Call CONVERT_A("C0004", txt_PROC_CD, txt_PROC_CD)   '4
        Call CONVERT_A("C0004", txt_BEF_PROC_CD, txt_BEF_PROC_CD)   '5
        Call CONVERT_A("B0005", T8, T8)
        Call CONVERT_A("C0006", t18A, t18A)
        Call CONVERT_A("C0008", T19A, T19A)
        Call CONVERT_C(Left(WULIAO, 1), Trim(txt_enduse_cd))  '43
        Call CONVERT_A("F0015", txt_IN_PLT_CD, txt_IN_PLT_CD)  '44
        Call CONVERT_A("C0011", txt_out_plt_cd, txt_out_plt_cd)  '51
        Call CONVERT_B(txt_SHP_EMP, txt_SHP_EMP)         '67
        Call CONVERT_B(txt_INS_EMP_CD, txt_INS_EMP_CD)  '77
        Call CONVERT_D(TXT_CUST_CD, TXT_CUST_CD)
    
    End If
  
    AdoRs.Close
    
    MDIMain.StatusBar1.Panels(1) = "提示信息: 资料已被查询"
    
    If WULIAO = "SL" Then
        txt_no.Text = Left(txt_no.Text, 8)
        txt_no.SelStart = 9
    Else
        txt_no.Text = Left(txt_no.Text, 10)
        txt_no.SelStart = 11
    End If
    
    If sCurInv = "00" Then
        Set AdoRs = Nothing
        Exit Sub
    End If
  
    sQuery1 = "SELECT     MOVE_DATE||MOVE_TIME, RCV_DATE, TO_INV"
    sQuery1 = sQuery1 + " FROM CP_MOVE_SLT"
    sQuery1 = sQuery1 + " WHERE MAT_NO = '" + Trim(txt_mat_no) + "'"
    If sCurInv <> "ZZ" Then
        sQuery1 = sQuery1 + "   AND TO_INV = '" + sCurInv + "'"
    End If
    
    sQuery1 = sQuery1 + " UNION "
    sQuery1 = sQuery1 + "SELECT  MOVE_DATE||MOVE_TIME, MOVE_DATE||MOVE_TIME, TO_PLT"
    sQuery1 = sQuery1 + " FROM CP_MOVE_INS"
    sQuery1 = sQuery1 + " WHERE MAT_NO = '" + Trim(txt_mat_no) + "'"
    If sCurInv <> "ZZ" Then
        sQuery1 = sQuery1 + "AND TO_PLT = '" + sCurInv + "'"
    End If
    AdoRs.Open sQuery1, M_CN1, adOpenKeyset
    
    If AdoRs.BOF Or AdoRs.EOF Then
        Exit Sub
    End If
 
    txt_MOVE_DATE.Text = ALLSET(Trim(AdoRs.Fields(0) & ""))
    txt_RECV_DATE.Text = ALLSET(Trim(AdoRs.Fields(1) & ""))
    
    If sCurInv = "ZZ" Then
        txt_CUR_INV.Text = txt_CUR_INV.Text & "(" & Gf_ComnNameFind(M_CN1, "C0013", Trim(AdoRs.Fields(2) & ""), 2) & ")"
    End If
    AdoRs.Close
    Set AdoRs = Nothing
    
End Sub

Private Function CONVERT_A(Cd_Mana_No As String, CD As String, textbox_name As TextBox)


    Dim AdoRs As ADODB.Recordset
    Dim sQuery5 As String
    
    Set AdoRs = New ADODB.Recordset
  
    If IsNull(CD) Or CD = "" Then
        Exit Function
    End If
  
    sQuery5 = "SELECT "
    sQuery5 = sQuery5 + "CD_SHORT_NAME"
    sQuery5 = sQuery5 + " FROM zp_cd "
    sQuery5 = sQuery5 + " WHERE cd_mana_no = '" + Trim(Cd_Mana_No) + "'"
    sQuery5 = sQuery5 + "AND CD='" + Trim(CD) + "'"
    AdoRs.Open sQuery5, M_CN1, adOpenKeyset
    
    If AdoRs.BOF Or AdoRs.EOF Then
        Exit Function
    End If
 
    If IsNull(AdoRs.Fields(0)) Or AdoRs.Fields(0) = "" Then
        textbox_name.Text = ""
    Else
        textbox_name.Text = Trim(AdoRs.Fields(0))
    End If
    
    AdoRs.Close

End Function

Private Function CONVERT_B(emp_id As String, textbox_name As TextBox)


    Dim AdoRs As ADODB.Recordset
    Dim sQuery5 As String
    
    Set AdoRs = New ADODB.Recordset
    
    If IsNull(emp_id) Or emp_id = "" Then
        Exit Function
    End If
  
    sQuery5 = "SELECT "
    sQuery5 = sQuery5 + "emp_name"
    sQuery5 = sQuery5 + " FROM ZP_EMPLOYEE"
    sQuery5 = sQuery5 + " WHERE emp_id = '" + Trim(emp_id) + "'"
    AdoRs.Open sQuery5, M_CN1, adOpenKeyset
    
    If AdoRs.BOF Or AdoRs.EOF Then
        Exit Function
    End If
 
    If IsNull(AdoRs.Fields(0)) Then
        textbox_name.Text = ""
    Else
        textbox_name.Text = Trim(AdoRs.Fields(0))
    End If
    AdoRs.Close

End Function

Private Function CONVERT_C(Prod_Knd As String, ENDUSE_CD As String)
    
    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    Set AdoRs = New ADODB.Recordset
    
    If IsNull(Prod_Knd) Or Prod_Knd = "" Then
        Exit Function
    End If
    
    sQuery = "SELECT "
    sQuery = sQuery + " ENDUSE_NAME"
    sQuery = sQuery + " FROM QP_ORD_USAGE"
    sQuery = sQuery + " WHERE PROD_KND='" + Prod_Knd + "'"
    sQuery = sQuery + " AND  ENDUSE_CD='" + ENDUSE_CD + "'"
        
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If AdoRs.BOF Or AdoRs.EOF Then
        Exit Function
    End If
    
    txt_APLY_ENDUSE_CD = Trim(AdoRs.Fields(0))
    
    AdoRs.Close

End Function

Private Function CONVERT_D(CUST_CD As String, TXT_CUST_CD As TextBox)
    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    Set AdoRs = New ADODB.Recordset
    
    If IsNull(TXT_CUST_CD) Or TXT_CUST_CD = "" Then
        Exit Function
    End If
      
    sQuery = "SELECT "
    sQuery = sQuery + " CUST_NM"
    sQuery = sQuery + " FROM BP_CUST_CD"
    sQuery = sQuery + " WHERE CUST_CD='" + CUST_CD + "'"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If AdoRs.BOF Or AdoRs.EOF Then
        Exit Function
    End If
     
    TXT_CUST_CD = Trim(AdoRs.Fields(0))
    
    AdoRs.Close
End Function


Private Function DATESET(AA As String) As String

    Dim AA1 As String
    Dim AA2 As String
    Dim AA3 As String
    
    If Trim(AA) = "" Then Exit Function
    
    AA1 = Mid(AA, 1, 4) + "-"
    AA2 = Mid(AA, 5, 2) + "-"
    AA3 = AA1 + AA2 + Mid(AA, 7, 2)
    
    DATESET = AA3

End Function

Private Function TIMESET(AA As String) As String

    Dim BA1 As String
    Dim BA2 As String
    Dim BA3 As String
    
    If Trim(AA) = "" Then Exit Function
    
    BA1 = Mid(AA, 1, 2) + ":"
    BA2 = Mid(AA, 3, 2) + ":"
    BA3 = BA1 + BA2 + Mid(AA, 5, 2)
    
    TIMESET = BA3

End Function

Private Function ALLSET(AA As String) As String

    Dim CA1 As String
    Dim CA2 As String
    Dim CA3 As String
    Dim CA4 As String
    Dim CA5 As String
    Dim CA6 As String
    
    If Trim(AA) = "" Then Exit Function
    
    CA1 = Mid(AA, 1, 4) + "-"
    CA2 = Mid(AA, 5, 2) + "-"
    CA3 = Mid(AA, 7, 2) + " "
    CA4 = Mid(AA, 9, 2) + ":"
    CA5 = Mid(AA, 11, 2) + ":"
    CA6 = CA1 + CA2 + CA3 + CA4 + CA5 + Mid(AA, 13, 2)
    ALLSET = CA6
    
End Function

Private Sub ALL_ITEMCLS()

    txt_OCCR_CD.Text = ""
    txt_PLT.Text = ""
    txt_PRC_line.Text = ""
    txt_REC_STS.Text = ""
    txt_PROC_CD.Text = ""
    txt_BEF_PROC_CD.Text = ""
    T8.Text = ""
    txt_stlgrd.Text = ""
    txt_THK.RawData = ""
    '----------------------
    txt_WID.RawData = ""
    txt_LEN.RawData = ""
    txt_WGT.RawData = ""
    txt_CAL_WGT.RawData = ""
    
    txt_THK_GRP = ""
    txt_WID_GRP = ""
    
    txt_OVER_FL.Text = ""
    t18A.Text = ""
    txt_woo_rsn.Text = ""
    txt_ord_no.Text = ""
    '----------------------
    TXT_ORD_ITEM.Text = ""
    txt_ORG_ORD_NO.Text = ""
    txt_ORG_ORD_ITEM.Text = ""
    txt_enduse_cd.Text = ""
    txt_del_to_date.Text = ""
    txt_PROD_DATE.Text = ""
    txt_PROD_TIME.Text = ""
    txt_SHIFT.Text = ""
    txt_GROUP_CD.Text = ""
    '----------------------
    txt_LOC.Text = ""
    txt_BED_PILE_DATE.Text = ""
    txt_SMP_FL.Text = ""
    txt_SMP_LOC.Text = ""
    txt_SMP_LEN.RawData = ""
    txt_SMP_NO.Text = ""
    
    txt_INSP_EMP.Text = ""
    
    txt_DSC_DATE.Text = ""
    txt_DSC_TIME.Text = ""
    txt_prod_grd.Text = ""
    '---------------------
    txt_SURF_GRD.Text = ""
    txt_QUALITY_GRD.Text = ""
    txt_QUALITY_UPD_GRD.Text = ""
    txt_APLY_STDSPEC.Text = ""
    txt_APLY_ENDUSE_CD.Text = ""
    txt_IN_PLT_CD.Text = ""
    txt_IN_PLT.Text = ""
    txt_IN_PLT_DATE.Text = ""
    txt_IN_PLT_TIME.Text = ""
    txt_IN_CAR_NO.Text = ""
    '---------------------
    txt_IN_SHEET_NO.Text = ""
    txt_IN_PLT_CO.Text = ""
    txt_out_plt_cd.Text = ""
    txt_OUT_PLT.Text = ""
    txt_OUT_PLT_DATE.Text = ""
    txt_OUT_PLT_TIME.Text = ""
    txt_OUT_CAR_NO.Text = ""
    txt_OUT_SHEET_NO.Text = ""
    txt_SHP_IST_NO.Text = ""
    txt_SHP_IST_DATE.Text = ""
    '---------------------
    txt_mat_no.Text = ""
    txt_SHP_DATE.Text = ""
    txt_SHP_TIME.Text = ""
    txt_TRNS_NO.Text = ""
    txt_car_no.Text = ""
    txt_TRNS_CMPY_CD.Text = ""
    txt_TRAIN_LINE_NAME.Text = ""
    txt_dest_detail.Text = ""
    txt_CERT_RPT_FL.Text = ""
    txt_CERT_RPT_DATE.Text = ""
    '---------------------
    txt_CERT_RPT_TIME.Text = ""
    txt_SHP_IST_CAN_DATE.Text = ""
    txt_SHP_IST_CAN_TIME.Text = ""
    txt_SHP_IST_CAN_FL.Text = ""
    '---------------------
    txt_RHF_REJ_ORNOT.Text = ""
    txt_SF_ORNOT.Text = ""
    txt_AIM_HCR_KND.Text = ""
    txt_HCR_KND.Text = ""
    txt_SCR_ORNOT.Text = ""
    txt_MOTHER_SLAB.Text = ""
    txt_quality_id.Text = ""
    '---------------------
    txt_INDIA.RawData = ""
    txt_OUTDIA.RawData = ""
    txt_COIL_MARKING.Text = ""
    txt_COIL_BAND_YN.Text = ""
    txt_COIL_HD_LEN.RawData = ""
    txt_COIL_TAIL_LEN.RawData = ""
    txt_COIL_DC_CNT.Text = ""
    '----------------------
    txt_PLATE_SEC.Text = ""
    txt_CR_CD.Text = ""
    txt_PILE_NO.Text = ""
    '----------------------
    
    txt_BEF_APLY_STDSPEC.Text = ""
    txt_STDSPEC_CHG_FL.Text = ""
    Text_size_knd.Text = ""
    txt_CUR_INV.Text = ""
    txt_MOVE_DATE.Text = ""
    txt_RECV_DATE.Text = ""
    
    '-----------dzr
    
    txt_HOUSING_DATE = ""
    txt_HOUSING_TIME = ""
    txt_INS_EMP_CD = ""
    txt_INS_PGMID = ""
    txt_UPD_DATE = ""
    txt_INS_DATE = " "
    txt_UPD_EMP_CD = ""
    txt_UPD_PGM = ""
    txt_ORD_THK = " "
    txt_ORD_WID = " "
    txt_ORD_LEN = " "
    txt_ORD_WGT = " "
    txt_SHP_EMP = ""
    txt_PRC = ""
    txt_END_RES = ""
    T19A = ""
    TXT_CUST_CD = ""
    
    '-------------------------------pp
    txt_SF_ORNOT_PLATE = ""
    txt_UST_FL = ""
    txt_TRIM_FL = ""
    txt_NEXT_PROC = ""
    txt_ACT_SMP_FL_PLATE = ""
    txt_ORG_PLATE = ""
    txt_ACT_SMP_LEN_PLATE = ""
    
    '------------------------------hc
    txt_INDIA = ""
    txt_OUTDIA = ""
    txt_COIL_HD_LEN = ""
    txt_COIL_TAIL_LEN = ""
    txt_ORG_COIL_NO = ""
    txt_SURF_GRD_UPD_DATE = ""
    txt_ACT_SMP_FL = ""
    txt_ACT_SMP_LEN = ""
    
    'sl
    txt_SLAB_RHF_IN_DATE = ""
End Sub

Private Sub T18_Change()

    If Trim(t18A.Text) = "使用材质" Then
        cmd_fl_down.Visible = True
        ULabel77.Visible = True
        txt_woo_rsn.Visible = True
    Else
        cmd_fl_down.Visible = False
        ULabel77.Visible = False
        txt_woo_rsn.Visible = False
    End If

End Sub

Private Sub t18A1_Change()

 '   If Trim(T18A1.Text) = "1" Then
 '       cmd_fl_down.Visible = True
 '       ULabel77.Visible = True
 '       txt_woo_rsn.Visible = True
 '   Else
 '       cmd_fl_down.Visible = False
 '       ULabel77.Visible = False
 '       txt_woo_rsn.Visible = False
 '   End If
 
   If Trim(T18A1.Text) = "1" And Trim(txt_REC_STS) = "2" Then
        cmd_fl_down.Visible = True
        ULabel77.Visible = True
        txt_woo_rsn.Visible = True
   Else
       cmd_fl_down.Visible = False
       ULabel77.Visible = False
        txt_woo_rsn.Visible = False
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

Private Sub txt_no_Change()
    Dim SQL As String
    
    If Len(txt_no.Text) = 10 Then
        WULIAO = "SL"
    ElseIf Len(txt_no.Text) = 14 And Right(txt_no.Text, 2) = "00" Then
        WULIAO = "HC"
    ElseIf Len(txt_no.Text) = 14 Then
        WULIAO = "PP"
    Else
        WULIAO = ""
        Exit Sub
    End If
        
    If Left(txt_no, 10) = Left(txt_mat_no, 10) Then Exit Sub
    
    SQL = "SELECT COIL_NO FROM GP_COIL WHERE COIL_NO LIKE '" & Left(txt_no.Text, 10) & "%'"
    SQL = SQL & " UNION "
    SQL = SQL & "SELECT PLATE_NO FROM GP_PLATE WHERE PLATE_NO LIKE '" & Left(txt_no.Text, 10) & "%'"
    
    Call Gf_ComboAdd(M_CN1, CBO_NO, SQL)
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

