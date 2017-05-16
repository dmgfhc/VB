VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0222C 
   Caption         =   "加热规程_AQA0222C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox P_txt_THK_MIN 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13290
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox P_txt_THK_MAX 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14130
      TabIndex        =   9
      Top             =   1455
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox P_txt_WID_MAX 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14145
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox P_txt_WID_MIN 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13305
      TabIndex        =   0
      Top             =   1785
      Visible         =   0   'False
      Width           =   855
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   585
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   1032
      _Version        =   196609
      Begin VB.TextBox P_txt_PLT 
         Height          =   315
         Left            =   1860
         TabIndex        =   37
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmd_ListView_THK 
         Caption         =   "<"
         Height          =   315
         Left            =   10605
         TabIndex        =   4
         Top             =   165
         Width           =   435
      End
      Begin VB.TextBox P_txt_RHF_STLGRD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5010
         TabIndex        =   3
         Top             =   180
         Width           =   1455
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   315
         Left            =   9240
         TabIndex        =   5
         Top             =   165
         Width           =   1350
         _Version        =   393216
         _ExtentX        =   2381
         _ExtentY        =   556
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   2
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "AQA0222C.frx":0000
         UserResize      =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   0
         Left            =   3180
         Top             =   180
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         Caption         =   "钢种分类编号"
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Index           =   0
         Left            =   7440
         Top             =   150
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "厚度组"
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
         Left            =   90
         Top             =   180
         Width           =   1755
         _ExtentX        =   3096
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
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1140
      Left            =   0
      TabIndex        =   6
      Top             =   1395
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   2011
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSFrame SSFrame1 
         Height          =   1110
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   1958
         _Version        =   196609
         Begin VB.TextBox txt_REHEAT1_TMP_MAX 
            Height          =   315
            Left            =   3000
            TabIndex        =   40
            Top             =   90
            Width           =   575
         End
         Begin VB.TextBox txt_REHEAT1_TMP_MIN 
            Height          =   315
            Left            =   2400
            TabIndex        =   39
            Top             =   90
            Width           =   575
         End
         Begin VB.TextBox txt_PRE_ZONE_TMP_MAX 
            Height          =   300
            Left            =   2970
            TabIndex        =   36
            Top             =   690
            Width           =   575
         End
         Begin VB.TextBox txt_PRE_ZONE_TMP_MIN 
            Height          =   300
            Left            =   2400
            TabIndex        =   35
            Top             =   690
            Width           =   575
         End
         Begin VB.TextBox txt_PRE_ZONE_TMP_TGT 
            Height          =   300
            Left            =   1830
            TabIndex        =   34
            Top             =   690
            Width           =   575
         End
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   60
            Top             =   675
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "预热段温度"
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
         Begin VB.TextBox txt_FUR_R_TIME_MIN 
            Height          =   300
            Left            =   11940
            TabIndex        =   33
            Top             =   690
            Width           =   575
         End
         Begin VB.TextBox txt_SOK_ZONE_TIME_MIN 
            Height          =   300
            Left            =   7200
            TabIndex        =   32
            Top             =   690
            Width           =   575
         End
         Begin VB.TextBox txt_SOK_SLAB_TEMP_MIN 
            Height          =   315
            Left            =   11940
            TabIndex        =   31
            Top             =   90
            Width           =   575
         End
         Begin VB.TextBox txt_REHEAT2_TMP_MIN 
            Height          =   315
            Left            =   7170
            TabIndex        =   30
            Top             =   90
            Width           =   575
         End
         Begin VB.TextBox txt_FUR_R_TIME_MAX 
            Height          =   300
            Left            =   12510
            TabIndex        =   29
            Top             =   690
            Width           =   575
         End
         Begin VB.TextBox txt_SOK_ZONE_TIME_MAX 
            Height          =   300
            Left            =   7770
            TabIndex        =   28
            Top             =   690
            Width           =   575
         End
         Begin VB.TextBox txt_SOK_SLAB_TEMP_MAX 
            Height          =   315
            Left            =   12510
            TabIndex        =   27
            Top             =   90
            Width           =   575
         End
         Begin VB.TextBox txt_SOK_SLAB_TEMP_TGT 
            Height          =   315
            Left            =   11340
            TabIndex        =   26
            Top             =   90
            Width           =   575
         End
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   9570
            Top             =   675
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "总加热时间"
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
            Left            =   4830
            Top             =   675
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "均热时间"
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
            Left            =   9570
            Top             =   90
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "均热段温度"
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
         Begin VB.TextBox txt_REHEAT2_TMP_MAX 
            Height          =   315
            Left            =   7740
            TabIndex        =   25
            Top             =   90
            Width           =   575
         End
         Begin VB.TextBox txt_REHEAT2_TMP_TGT 
            Height          =   315
            Left            =   6600
            TabIndex        =   24
            Top             =   90
            Width           =   575
         End
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   4830
            Top             =   90
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "加热段二阶段温度"
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
         Begin VB.TextBox txt_REHEAT1_TMP_TGT 
            Height          =   315
            Left            =   1830
            MaxLength       =   4
            TabIndex        =   8
            Top             =   90
            Width           =   575
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   14
            Left            =   60
            Top             =   90
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "加热段一阶段温度"
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10770
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0222C.frx":03A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   5730
      Left            =   15
      TabIndex        =   11
      Top             =   3465
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   10107
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
      MaxCols         =   31
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQA0222C.frx":06FA
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   885
      Left            =   0
      TabIndex        =   12
      Top             =   2565
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   1561
      _Version        =   196609
      Begin VB.TextBox txt_RHF_STLGRD_ED 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1845
         MaxLength       =   80
         TabIndex        =   17
         Top             =   60
         Width           =   8115
      End
      Begin VB.TextBox txt_UPD_EMP 
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
         Height          =   300
         Left            =   11085
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txt_UPD_DATE 
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
         Height          =   300
         Left            =   7995
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   465
         Width           =   1215
      End
      Begin VB.TextBox txt_INS_EMP 
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
         Height          =   300
         Left            =   4935
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   465
         Width           =   1215
      End
      Begin VB.TextBox txt_INS_DATE 
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
         Height          =   300
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   465
         Width           =   1215
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Index           =   11
         Left            =   75
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "规范编辑号"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   16
         Left            =   60
         Top             =   465
         Width           =   1755
         _ExtentX        =   3096
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   17
         Left            =   3150
         Top             =   465
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "录入人"
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
         Index           =   18
         Left            =   6210
         Top             =   465
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "修改日期"
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
         Index           =   19
         Left            =   9300
         Top             =   480
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "修改人"
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
   Begin Threed.SSFrame SSFrame4 
      Height          =   795
      Left            =   0
      TabIndex        =   18
      Top             =   585
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   1402
      _Version        =   196609
      Begin VB.TextBox txt_PLT 
         Height          =   315
         Left            =   9210
         TabIndex        =   38
         Top             =   450
         Width           =   855
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   7440
         Top             =   450
         Width           =   1755
         _ExtentX        =   3096
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
      End
      Begin VB.TextBox txt_RHF_STLGRD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1860
         MaxLength       =   6
         TabIndex        =   23
         Top             =   45
         Width           =   1635
      End
      Begin VB.TextBox txt_THK_MIN 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9210
         TabIndex        =   22
         Top             =   45
         Width           =   855
      End
      Begin VB.TextBox txt_THK_MAX 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10080
         TabIndex        =   21
         Top             =   45
         Width           =   855
      End
      Begin VB.TextBox txt_HCR_KND 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1860
         MaxLength       =   1
         TabIndex        =   20
         Top             =   450
         Width           =   405
      End
      Begin VB.TextBox txt_HCR_NAME 
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
         Height          =   300
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   19
         Top             =   450
         Width           =   1380
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   4
         Left            =   90
         Top             =   450
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "HCR 分类"
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
         Index           =   20
         Left            =   90
         Top             =   45
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "钢种分类编号"
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Index           =   1
         Left            =   7440
         Top             =   45
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "厚度组"
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
End
Attribute VB_Name = "AQA0222C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量标准管理
'-- Program Name      加热规范输入
'-- Program ID        AQA0222C
'-- Document No       Q-00-0010(Specification)
'-- Designer          SUN BIN
'-- Coder             SUN BIN
'-- Date              2009.3.24
'-- Description       加热规范输入
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
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim lCopyRow As Long                'Copy Row
Dim btChk_THK   As Boolean          '厚度组(ss3)是否显示选择
Dim btChk_WID   As Boolean


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(P_txt_RHF_STLGRD, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(P_txt_THK_MIN, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(P_txt_THK_MAX, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'        Call Gp_Ms_Collection(P_txt_WID_MIN, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'        Call Gp_Ms_Collection(P_txt_WID_MAX, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(P_txt_PLT, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_HCR_KND, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, "p", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, "p", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
'    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
'
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQA0222C.P_SREFER", Key:="P-R"
    Sc1.Add Item:="AQA0222C.P_ONEROW", Key:="P-O"
    Sc1.Add Item:="AQA0222C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
         
End Sub
'Private Function Change_MLT_PLT(ByVal iC1Val As Integer, ByVal iC2Val As Integer, Optional iRow As Long = 0) As String
'Dim sOLD_ML_PLT_CD  As String
'Dim sNEW_ML_PLT_CD  As String
'Dim sLAB            As String
'Dim s_C1_CD         As String
'Dim s_C2_CD         As String
'
'    If iRow > 0 Then
'        With ss1
'            .Row = iRow
'            .Col = 1
'            sLAB = .Text
'            .Col = 43
'            sOLD_ML_PLT_CD = .Text
'        End With
'    End If
'
'    If sOLD_ML_PLT_CD = "**" And (sLAB <> "Input" Or sLAB <> "Update") Then
'        Change_MLT_PLT = sOLD_ML_PLT_CD
'        Exit Function
'    End If
'
'    If iC1Val = 1 Then
'        s_C1_CD = "C1"
'    Else
'        s_C1_CD = "NO"
'    End If
'
'    If iC2Val = 1 Then
'        s_C2_CD = "C3"
'    Else
'        s_C2_CD = "NO"
'    End If
'
'    If s_C1_CD = "C1" And s_C2_CD = "C3" Then
'        sNEW_ML_PLT_CD = "**"
'    ElseIf s_C1_CD = "C1" And s_C2_CD = "NO" Then
'        sNEW_ML_PLT_CD = "C1"
'    ElseIf s_C1_CD = "NO" And s_C2_CD = "C3" Then
'        sNEW_ML_PLT_CD = "C3"
'    Else
'        sNEW_ML_PLT_CD = sOLD_ML_PLT_CD
'
'    End If
'
'    Change_MLT_PLT = sNEW_ML_PLT_CD
'
'End Function

Private Sub cmd_ListView_THK_Click()
Dim sQuery As String
    
    sQuery = " Select THK_MIN,THK_MAX From QP_RHF_STD Where RHF_STLGRD = "
    btChk_THK = Not btChk_THK

    If btChk_THK = False Then
            
        With ss3
            
            .MaxCols = 2
            .MaxRows = 1
            .Height = 313
        
            btChk_THK = False
    
        End With

    Else
         
       If P_txt_RHF_STLGRD.Text = "" Or Trim(P_txt_RHF_STLGRD.Text) = "" Then
            Call MsgBox("请输入钢种分类编号", vbOKOnly, "系统信息")
          Exit Sub
       End If
           
        sQuery = sQuery + "'" + P_txt_RHF_STLGRD.Text + "' AND"
        
        Call GS_Combo_SS_ADD(sQuery, ss3)
        
        Call GS_ssBackColorSet(ss3)
    
    End If
    
        If Gf_GetCellNullCheck(ss3, 1, 1) <> "" And Gf_GetCellNullCheck(ss3, 1, 2) <> "" Then
            P_txt_THK_MIN.Text = Gf_GetCellNullCheck(ss3, 1, 1)
            P_txt_THK_MAX.Text = Gf_GetCellNullCheck(ss3, 1, 2)
        End If

End Sub

'Private Sub cmd_ListView_WID_Click()
'Dim sQuery As String
'
'    sQuery = " Select WID_MIN,WID_MAX From QP_RHF_STD Where RHF_STLGRD = "
'    btChk_WID = Not btChk_WID
'
'    If btChk_THK = False Then
'
'        With ss4
'
'            .MaxCols = 2
'            .MaxRows = 1
'            .Height = 313
'
'            btChk_THK = False
'
'        End With
'
'    Else
'
'     If P_txt_RHF_STLGRD.Text = "" Or Trim(P_txt_RHF_STLGRD.Text) = "" Then
'            Call MsgBox("请输入钢种分类编号", vbOKOnly, "系统信息")
'          Exit Sub
'       End If
'
'        sQuery = sQuery + "'" + P_txt_RHF_STLGRD.Text + "' AND"
'
'        Call GS_Combo_SS_ADD(sQuery, ss4)
'
'        Call GS_ssBackColorSet(ss4)
'
'    End If
'
''        If Gf_GetCellNullCheck(ss4, 1, 1) <> "" And Gf_GetCellNullCheck(ss4, 1, 2) <> "" Then
''            P_txt_WID_MIN.Text = Gf_GetCellNullCheck(ss4, 1, 1)
''            P_txt_WID_MAX.Text = Gf_GetCellNullCheck(ss4, 1, 2)
''        End If
'End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
'    If Not (KeyCode = vbKeyF4) Then Exit Sub
    
    Select Case Me.ActiveControl.Name
'
'        Case "P_txt_RHF_STLGRD"                            '轧钢规程编号
'            sCode = "RHF_STLGRD"
'
'        Case "txt_RHF_STLGRD"                              '轧钢规程编号
'            sCode = "RHF_STLGRD"

'        Case "txt_STLGRD"               '钢种
'            sCode = "STLGRD"
'            Set oCodeName = txt_STLGRD_DETAIL

        Case "txt_HCR_KND"              'HCR 分类
            sCode = "C0005"
            Set oCodeName = txt_HCR_NAME
            
        Case "txt_PLT"              '工厂
             sCode = "C0001"
      
        Case "P_txt_PLT"            '工厂
             sCode = "C0001"

    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
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
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)
       
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
'    Call GP_ROW_BACKCOLOR(ss1)
    
    'Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
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
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Ins()
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 26)
    Call Spread_to_Master(ss1, ss1.ActiveRow)
    txt_RHF_STLGRD.SetFocus
End Sub

Public Sub Form_Pro()
    Dim iMaxrow As Long
    Dim iRow As Long
    Dim icount As Long

    iRow = ss1.Row
    iMaxrow = ss1.MaxRows
    
    ss1.ReDraw = False
    
    For icount = 1 To ss1.MaxRows
         ss1.Col = 0
         ss1.Row = icount
        Select Case ss1.Text
            Case "Input", "Update"
             If Sp_AllUse_NecessaryCheck(icount) = False Then          '必须输入项检查
                Call Spread_to_Master(ss1, icount)
                Exit Sub
             End If                                             '最大值,最小值,目标值检查
             If Sp_subMinMaxValueCheck(icount) = False Then
                Call Spread_to_Master(ss1, icount)
                Exit Sub
             End If
        End Select
    Next icount
    
    ss1.ReDraw = True
     
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        txt_ins_emp.Text = sUserID
         If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            Call Gp_Goto_Row(ss1, iMaxrow, iRow)
            Call Spread_to_Master(ss1, iRow)
         End If
    End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_AllDel(M_CN1, Proc_Sc("Sc"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub
Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call MS_Cls
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        ss3.MaxRows = 1
        ss3.Height = 255
        btChk_THK = False
        Call GP_SET_CELL_VALUE(ss3, 1, 2, "")
'        Call GP_SET_CELL_VALUE(ss4, 1, 2, "")
'        txt_THK_MIN.Text = ""
'        txt_THK_MAX.Text = ""
        'rControl(1).SetFocus
        P_txt_RHF_STLGRD.SetFocus
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
            
            If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
                Call Spread_to_Master(ss1, 1)
                Call Gp_Ms_ControlLock(Mc1("pControl"), True)
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call GP_SELECT_ROW(ss1, 1)
                ss3.MaxRows = 1
                ss3.Height = 255
'                ss4.MaxRows = 1
'                ss4.Height = 255
                btChk_THK = False
'                btChk_WID = False
                Exit Sub
            End If
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

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

Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)
    If Gf_Sc_Authority(sAuthority, "U") Then

        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), 0)
        
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 26)
        
        Call Spread_to_Master(ss1, ss1.ActiveRow)

    End If
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub



Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 26)
    End If
End Sub

Private Sub ss1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    Call Spread_to_Master(ss1, NewRow)
'    Call GP_SetRowHeaderClear(ss1, NewRow)
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

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

Private Sub Spread_to_Master(ByVal sp As vaSpread, ByVal iRow As Long)
    Dim RowLabel As String

        With sp
        
            If iRow > 0 Then
                .Row = iRow
               
                .Col = 0: RowLabel = .Text
                .Col = 1: txt_RHF_STLGRD.Text = .Text
                .Col = 2: txt_THK_MIN.Text = .Text
                .Col = 3: txt_THK_MAX.Text = .Text
'                .Col = 4: txt_WID_MIN.Text = .Text
'                .Col = 5: txt_WID_MAX.Text = .Text
                .Col = 6: txt_HCR_KND.Text = .Text
                .Col = 7: txt_HCR_NAME.Text = .Text
                
                .Col = 8: txt_PRE_ZONE_TMP_TGT.Text = .Text
                .Col = 9: txt_PRE_ZONE_TMP_MIN.Text = .Text
                .Col = 10: txt_PRE_ZONE_TMP_MAX.Text = .Text
                .Col = 11: txt_REHEAT1_TMP_TGT.Text = .Text
                .Col = 12: txt_REHEAT1_TMP_MIN = .Text
                .Col = 13: txt_REHEAT1_TMP_MAX.Text = .Text
                .Col = 14: txt_REHEAT2_TMP_TGT.Text = .Text
                .Col = 15: txt_REHEAT2_TMP_MIN.Text = .Text
                .Col = 16: txt_REHEAT2_TMP_MAX.Text = .Text
                .Col = 17: txt_SOK_SLAB_TEMP_TGT.Text = .Text
                .Col = 18: txt_SOK_SLAB_TEMP_MIN.Text = .Text
                .Col = 19: txt_SOK_SLAB_TEMP_MAX.Text = .Text
                .Col = 20: txt_SOK_ZONE_TIME_MIN.Text = .Text
                .Col = 21: txt_SOK_ZONE_TIME_MAX.Text = .Text
                .Col = 22: txt_FUR_R_TIME_MIN.Text = .Text
                .Col = 23: txt_FUR_R_TIME_MAX.Text = .Text
                             
                .Col = 24: txt_RHF_STLGRD_ED.Text = .Text
                .Col = 25: txt_INS_DATE.Text = .Text
                .Col = 27: txt_ins_emp.Text = .Text
                .Col = 28: txt_UPD_DATE.Text = .Text
                .Col = 30: txt_UPD_EMP.Text = .Text
                .Col = 31: txt_PLT.Text = .Text
                
                If RowLabel = "Input" Then
                    txt_RHF_STLGRD.Locked = False
                    txt_THK_MIN.Locked = False
                    txt_THK_MAX.Locked = False
'                    txt_WID_MIN.Locked = False
'                    txt_WID_MAX.Locked = False
                Else
                    txt_RHF_STLGRD.Locked = True
                    txt_THK_MIN.Locked = True
                    txt_THK_MAX.Locked = True
'                    txt_WID_MIN.Locked = True
'                    txt_WID_MAX.Locked = True
                End If
            Else
                Exit Sub
            End If
        
        End With

End Sub

Public Sub Spread_Can()

    Call GP_SELECT_ROW(ss1, ss1.Row)
    Call GP_ROW_CANCEL(Proc_Sc("Sc"))
'    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    Call Spread_to_Master(ss1, ss1.ActiveRow)
    Call Gp_Ms_ControlLock(Mc1("pControl"), True)
      
End Sub

Public Sub Spread_Del()
    
    Call GP_SET_CELL_VALUE(ss1, ss1.Row, 0, "Delete")
    'Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Public Sub Spread_Cpy()

    lCopyRow = ss1.ActiveRow
    'Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()
    
    Call GP_ROW_PASTE(Proc_Sc("Sc"), lCopyRow)
    'Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 26)
    Call Spread_to_Master(ss1, ss1.ActiveRow)
    txt_RHF_STLGRD.SetFocus
    
End Sub

Public Sub Ms_To_SP(ByVal sp As vaSpread, ByVal iRow As Long, ByVal iCol As Long, ByVal vName As String)
    Dim old_Value As Variant
    Dim iValue As Variant
    
    If (vName <> "0") And (vName <> "1") Then
        If TypeName(Me.Controls(vName)) = "TextBox" Then
            iValue = Me.Controls(vName).Text
        End If
        
        If TypeName(Me.Controls(vName)) = "sidbEdit" Then
                iValue = Me.Controls(vName).Value
        End If
    Else
        iValue = vName
    End If
    
    With sp
        If iCol = 1 Or iCol = 2 Or iCol = 3 Or iCol = 4 Or iCol = 5 Then
            .Row = iRow
            .Col = 0
            If (.Text = "Input") Then
                .Col = iCol
                .Value = iValue
                .Text = iValue
            Else
                Exit Sub
            End If
        Else
            .Row = iRow
            .Col = iCol
            old_Value = .Value
            .Value = iValue
            .Text = iValue
            If old_Value <> .Value Then
                .Col = 0
                    If (.Text = "Input") Or (.Text = "Update") Then
                        .Text = .Text
                    Else
                        .Text = "Update"
                    End If
                    .Col = iCol
            Else
                Exit Sub
            End If
        End If
    End With
End Sub



Private Sub ss3_DblClick(ByVal Col As Long, ByVal Row As Long)
    With ss3
    
        If Gf_GetCellNullCheck(ss3, Row, 1) <> "" And Gf_GetCellNullCheck(ss3, Row, 2) <> "" Then
            Call GP_SET_CELL_VALUE(ss3, 1, 1, Gf_GetCellNullCheck(ss3, Row, 1))
            Call GP_SET_CELL_VALUE(ss3, 1, 2, Gf_GetCellNullCheck(ss3, Row, 2))
        End If
        
        .MaxRows = 1
        .Height = 313
        
        P_txt_THK_MIN.Text = Gf_GetCellNullCheck(ss3, 1, 1)
        P_txt_THK_MAX.Text = Gf_GetCellNullCheck(ss3, 1, 2)
        
        btChk_THK = False
    
    End With


End Sub


Private Sub txt_RHF_STLGRD_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 1, txt_RHF_STLGRD.Name)
    End If
End Sub


Private Sub txt_RHF_STLGRD_ED_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 24, txt_RHF_STLGRD_ED.Name)
    End If
End Sub

Private Sub txt_THK_MIN_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 2, txt_THK_MIN.Name)
    End If
End Sub

Private Sub txt_THK_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 3, txt_THK_MAX.Name)
    End If
End Sub


'Private Sub txt_WID_MAX_Change()
'    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
'        Call Ms_To_SP(ss1, ss1.Row, 5, txt_WID_MAX.Name)
'    End If
'End Sub
'
'Private Sub txt_WID_MIN_Change()
'    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
'        Call Ms_To_SP(ss1, ss1.Row, 4, txt_WID_MIN.Name)
'    End If
'End Sub

Private Sub txt_HCR_KND_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 6, txt_HCR_KND.Name)
    End If
End Sub

Private Sub txt_HCR_NAME_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 7, txt_HCR_NAME.Name)
    End If
End Sub

Private Sub txt_PRE_ZONE_TMP_TGT_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 8, txt_PRE_ZONE_TMP_TGT.Name)
    End If
End Sub
Private Sub txt_PRE_ZONE_TMP_MIN_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 9, txt_PRE_ZONE_TMP_MIN.Name)
    End If
End Sub
Private Sub txt_PRE_ZONE_TMP_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 10, txt_PRE_ZONE_TMP_MAX.Name)
    End If
End Sub

Private Sub txt_REHEAT1_TMP_TGT_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 11, txt_REHEAT1_TMP_TGT.Name)
    End If
End Sub
Private Sub txt_REHEAT1_TMP_MIN_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 12, txt_REHEAT1_TMP_MIN.Name)
    End If
End Sub
Private Sub txt_REHEAT1_TMP_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 13, txt_REHEAT1_TMP_MAX.Name)
    End If
End Sub
Private Sub txt_REHEAT2_TMP_TGT_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 14, txt_REHEAT2_TMP_TGT.Name)
    End If
End Sub
Private Sub txt_REHEAT2_TMP_MIN_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 15, txt_REHEAT2_TMP_MIN.Name)
    End If
End Sub
Private Sub txt_REHEAT2_TMP_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 16, txt_REHEAT2_TMP_MAX.Name)
    End If
End Sub
Private Sub txt_SOK_SLAB_TEMP_TGT_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 17, txt_SOK_SLAB_TEMP_TGT.Name)
    End If
End Sub
Private Sub txt_SOK_SLAB_TEMP_MIN_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 18, txt_SOK_SLAB_TEMP_MIN.Name)
    End If
End Sub
Private Sub txt_SOK_SLAB_TEMP_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 19, txt_SOK_SLAB_TEMP_MAX.Name)
    End If
End Sub
Private Sub txt_SOK_ZONE_TIME_MIN_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 20, txt_SOK_ZONE_TIME_MIN.Name)
    End If
End Sub
Private Sub txt_SOK_ZONE_TIME_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 21, txt_SOK_ZONE_TIME_MAX.Name)
    End If
End Sub
Private Sub txt_FUR_R_TIME_MIN_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 22, txt_FUR_R_TIME_MIN.Name)
    End If
End Sub
Private Sub txt_FUR_R_TIME_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 23, txt_FUR_R_TIME_MAX.Name)
    End If
End Sub
Private Sub TXT_PLT_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 31, txt_PLT.Name)
    End If
End Sub


'Private Function txt_KeyPress(KeyAscii As Integer) As Integer
'
'        Select Case KeyAscii
'
'               Case Is <= 32
'                    txt_KeyPress = KeyAscii
'               Case 48 To 57
'                    txt_KeyPress = KeyAscii
'               Case 46
'                    txt_KeyPress = KeyAscii
'               Case 45
'                    txt_KeyPress = KeyAscii
'               Case Else
'                    txt_KeyPress = 0
'        End Select
'
'
'End Function

Private Sub MS_Cls()
    Dim i As Integer
    For i = 0 To Me.COUNT - 1
        If TypeName(Me.Controls(i)) = "TextBox" Then
            Me.Controls(i).Text = ""
        ElseIf TypeName(Me.Controls(i)) = "sidbEdit" Then
            Me.Controls(i).Text = ""
        ElseIf TypeName(Me.Controls(i)) = "CheckBox" Then
            Me.Controls(i).Value = 0
        End If
        
    Next i
End Sub


'下限值 , 上限值,目标值 Check
Private Function Sp_subMinMaxValueCheck(iRow As Long) As Boolean
    
'厚度组
    If Gf_Sp_subValueCheck(Sc1, iRow, 2, 3, ULabel3(1).Caption, txt_THK_MIN) = False Then Exit Function
'宽度组
    If Gf_Sp_subValueCheck(Sc1, iRow, 4, 5, ULabel3(2).Caption, txt_THK_MIN) = False Then Exit Function
'加热段一阶段温度
    If GF_MIN_MAX_TARGET_CHECK(txt_REHEAT1_TMP_MIN, txt_REHEAT1_TMP_MAX, txt_REHEAT1_TMP_TGT) = False Then Exit Function
'加热段二阶段温度
    If GF_MIN_MAX_TARGET_CHECK(txt_REHEAT2_TMP_MIN, txt_REHEAT2_TMP_MAX, txt_REHEAT2_TMP_TGT) = False Then Exit Function
'均热段温度
    If GF_MIN_MAX_TARGET_CHECK(txt_SOK_SLAB_TEMP_MIN, txt_SOK_SLAB_TEMP_MAX, txt_SOK_SLAB_TEMP_TGT) = False Then Exit Function
'预热段温度
    If GF_MIN_MAX_TARGET_CHECK(txt_PRE_ZONE_TMP_MIN, txt_PRE_ZONE_TMP_MAX, txt_PRE_ZONE_TMP_TGT) = False Then Exit Function
    
    Sp_subMinMaxValueCheck = True

End Function

'必须输入项目检查
Private Function Sp_AllUse_NecessaryCheck(iRow As Long) As Boolean
Dim sPLT_CD As String

        With ss1
            .Row = iRow
            .Col = 33
            sPLT_CD = .Text
        End With

    Sp_AllUse_NecessaryCheck = True
End Function

'Private Function Sp_Item_NecessaryCheck(iRow As Long) As Boolean
'''预热段温度
'    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 10, "预热段目标温度", txt_REHEAT1_TMP_TGT, True) = False Then Exit Function
''加热一阶段温度
'    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 11, "预热段目标温度", txt_REHEAT1_TMP_TGT, True) = False Then Exit Function
''加热二阶段温度
'    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 14, "预热段目标温度", txt_REHEAT1_TMP_TGT, True) = False Then Exit Function
'
'    Sp_C1_Item_NecessaryCheck = True
'
'End Function

    
'End Function



