VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQB0020C 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "质量设计现状详细查询 "
   ClientHeight    =   9975
   ClientLeft      =   1440
   ClientTop       =   2355
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   14550
   Begin Threed.SSPanel SSPanel5 
      Align           =   1  'Align Top
      Height          =   3135
      Left            =   0
      TabIndex        =   48
      Top             =   7755
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   5530
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_METH3_NAME_Q 
         Height          =   315
         Left            =   9450
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   1800
         Width           =   2235
      End
      Begin VB.TextBox txt_METH2_NAME_Q 
         Height          =   315
         Left            =   6210
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   1800
         Width           =   2235
      End
      Begin VB.TextBox txt_METH1_NAME_Q 
         Height          =   315
         Left            =   2970
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   1800
         Width           =   2235
      End
      Begin VB.TextBox txt_HTM_SHOT_BLAST_NAME_Q 
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
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1395
         Width           =   7395
      End
      Begin VB.TextBox txt_HTM_SHOT_BLAST_Q 
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
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   1395
         Width           =   585
      End
      Begin VB.TextBox txt_COND3_Q 
         Height          =   315
         Left            =   8865
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   1800
         Width           =   570
      End
      Begin VB.TextBox txt_METH3_Q 
         Height          =   310
         Left            =   8505
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   1800
         Width           =   350
      End
      Begin VB.TextBox txt_COND2_Q 
         Height          =   315
         Left            =   5625
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   1800
         Width           =   570
      End
      Begin VB.TextBox txt_METH2_Q 
         Height          =   310
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   1800
         Width           =   350
      End
      Begin VB.TextBox txt_COND1_Q 
         Height          =   315
         Left            =   2385
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   1800
         Width           =   570
      End
      Begin VB.TextBox txt_METH1_Q 
         Height          =   310
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   1800
         Width           =   350
      End
      Begin VB.TextBox txt_CUST_SPEC_NO 
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
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   570
         Width           =   2055
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   570
         Width           =   1380
      End
      Begin VB.TextBox txt_STLGRD_GRP 
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
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   570
         Width           =   1065
      End
      Begin VB.TextBox txt_DESIGN_DATE 
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   150
         Width           =   1605
      End
      Begin VB.TextBox txt_MLT_STD_NO 
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
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   990
         Width           =   1605
      End
      Begin VB.TextBox txt_DESIGN_STS 
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
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   150
         Width           =   285
      End
      Begin VB.TextBox txt_MILL_STD_NO 
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
         Left            =   9990
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   990
         Width           =   1605
      End
      Begin VB.TextBox txt_DEV_STD_CD 
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
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   990
         Width           =   2055
      End
      Begin VB.TextBox txt_DESIGN_STS_NAME 
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
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   150
         Width           =   1755
      End
      Begin VB.TextBox txt_STLGRD_DETAIL 
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
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   570
         Width           =   3285
      End
      Begin VB.TextBox txt_Nisco_Quality_No 
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
         Left            =   9990
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   150
         Width           =   2025
      End
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   24
         Left            =   4320
         Top             =   570
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "钢种/钢种组"
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
         Index           =   27
         Left            =   90
         Top             =   150
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "质量设计状态"
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
         Index           =   28
         Left            =   4320
         Top             =   990
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "炼钢/连铸规程编号"
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
         Index           =   29
         Left            =   4320
         Top             =   150
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "质量设计日期"
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
         Index           =   13
         Left            =   90
         Top             =   570
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "客户特殊要求编号"
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
         Index           =   14
         Left            =   90
         Top             =   990
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "代表性交付条件标准"
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
         Index           =   34
         Left            =   8070
         Top             =   990
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "轧钢规程编号"
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
         Index           =   35
         Left            =   8070
         Top             =   150
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "企标材质编号"
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
         Index           =   43
         Left            =   90
         Top             =   1800
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "热处理方法/条件"
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
         Index           =   44
         Left            =   90
         Top             =   1395
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "抛丸代码"
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
   Begin Threed.SSPanel SSPanel4 
      Align           =   1  'Align Top
      Height          =   2775
      Left            =   0
      TabIndex        =   39
      Top             =   4980
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   4895
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_CE_QS 
         Height          =   310
         Left            =   10050
         TabIndex        =   88
         Top             =   140
         Width           =   1545
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   8100
         Top             =   135
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         Caption         =   "CE/QS认证"
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
      Begin VB.TextBox txt_UST_NAME 
         Height          =   315
         Left            =   6975
         TabIndex        =   75
         Top             =   1260
         Width           =   3075
      End
      Begin VB.TextBox txt_HTM_SHOT_BLAST_NAME 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1980
         Width           =   7395
      End
      Begin VB.TextBox txt_HTM_SHOT_BLAST 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1980
         Width           =   585
      End
      Begin VB.TextBox txt_CFM_MILL_PLT_NAME 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1575
         Width           =   1680
      End
      Begin VB.TextBox txt_CFM_MILL_PLT 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   1575
         Width           =   585
      End
      Begin VB.TextBox txt_CFM_SMS_PLT_NAME 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1635
      End
      Begin VB.TextBox txt_CFM_SMS_PLT 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1215
         Width           =   585
      End
      Begin VB.TextBox txt_CUST_REQ_CODE 
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
         Left            =   6270
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1620
         Width           =   3765
      End
      Begin VB.TextBox txt_MATR_FL 
         Height          =   310
         Left            =   10350
         Locked          =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   510
         Width           =   585
      End
      Begin VB.TextBox txt_COLOR_STROKE 
         Height          =   310
         Left            =   6270
         Locked          =   -1  'True
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   862
         Width           =   5415
      End
      Begin VB.TextBox txt_METH3_NAME 
         Height          =   310
         Left            =   8790
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   2340
         Width           =   2000
      End
      Begin VB.TextBox txt_METH2_NAME 
         Height          =   310
         Left            =   5595
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   2340
         Width           =   2000
      End
      Begin VB.TextBox txt_METH1_NAME 
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   2340
         Width           =   2000
      End
      Begin VB.TextBox txt_METH3 
         Height          =   310
         Left            =   8430
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   2340
         Width           =   350
      End
      Begin VB.TextBox txt_METH2 
         Height          =   310
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   2340
         Width           =   350
      End
      Begin VB.TextBox txt_METH1 
         Height          =   310
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   2340
         Width           =   350
      End
      Begin VB.TextBox txt_TRIM_FL 
         Height          =   310
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   870
         Width           =   585
      End
      Begin VB.TextBox txt_STAMP 
         Height          =   310
         Left            =   6570
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   510
         Width           =   585
      End
      Begin VB.TextBox txt_UST_FL 
         Height          =   315
         Left            =   6270
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1260
         Width           =   675
      End
      Begin VB.TextBox txt_INDIA 
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
         Left            =   6270
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   105
         Width           =   1545
      End
      Begin VB.TextBox txt_INSP_CD 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   120
         Width           =   585
      End
      Begin VB.TextBox txt_INSP_NAME 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   120
         Width           =   1635
      End
      Begin VB.TextBox txt_PACK_WAY_NAME 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   495
         Width           =   1635
      End
      Begin VB.TextBox txt_PACK_WAY 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   495
         Width           =   585
      End
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   19
         Left            =   4350
         Top             =   105
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "内径"
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
         Index           =   12
         Left            =   4350
         Top             =   510
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "喷印"
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
         Index           =   21
         Left            =   4350
         Top             =   1260
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "超声波探伤(UST)"
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
         Index           =   22
         Left            =   90
         Top             =   120
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "检查机关"
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
         Index           =   7
         Left            =   90
         Top             =   862
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "是否切边"
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
         Index           =   33
         Left            =   90
         Top             =   491
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "包装方式"
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
         Index           =   25
         Left            =   8100
         Top             =   510
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   556
         Caption         =   "是否保证性能"
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
         Index           =   37
         Left            =   4350
         Top             =   862
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "色标"
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
         Index           =   38
         Left            =   4350
         Top             =   1620
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "客户要求代码"
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
         Index           =   39
         Left            =   90
         Top             =   1200
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "投入炼钢厂"
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
         Index           =   40
         Left            =   75
         Top             =   1575
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "投入轧钢厂"
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
         Index           =   41
         Left            =   90
         Top             =   1975
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "抛丸代码"
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
         Index           =   42
         Left            =   90
         Top             =   2340
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "热处理方法"
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
      Begin VB.Image Img_MATR_FL 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         Stretch         =   -1  'True
         Top             =   510
         Width           =   285
      End
      Begin VB.Image Img_TRIM_FL 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2040
         Stretch         =   -1  'True
         Top             =   870
         Width           =   285
      End
      Begin VB.Image Img_STAMP 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6270
         Stretch         =   -1  'True
         Top             =   510
         Width           =   285
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Align           =   1  'Align Top
      Height          =   2355
      Left            =   0
      TabIndex        =   20
      Top             =   2625
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   4154
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_DEPT_CD 
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
         Left            =   10020
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   840
         Width           =   585
      End
      Begin VB.TextBox txt_DEPT_NAME 
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
         Left            =   10620
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   840
         Width           =   1035
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
         Left            =   10020
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   465
         Width           =   1605
      End
      Begin VB.TextBox txt_DEST_CD 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   465
         Width           =   1365
      End
      Begin VB.TextBox txt_DEST_NAME 
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
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   465
         Width           =   4485
      End
      Begin VB.TextBox txt_PONO 
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
         Left            =   10020
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   90
         Width           =   1605
      End
      Begin VB.TextBox txt_CUST_NAME 
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
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   90
         Width           =   4485
      End
      Begin VB.TextBox txt_CUST_CD 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   90
         Width           =   1365
      End
      Begin VB.TextBox txt_ORD_KND 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1965
         Width           =   285
      End
      Begin VB.TextBox txt_ORD_KND_NAME 
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
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1965
         Width           =   1095
      End
      Begin VB.TextBox txt_ORD_CUST_CD 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   840
         Width           =   1365
      End
      Begin VB.TextBox txt_ORD_CUST_NAME 
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
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   840
         Width           =   4485
      End
      Begin VB.TextBox txt_END_CUST_CD 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1365
      End
      Begin VB.TextBox txt_END_CUST_NAME 
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
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1215
         Width           =   4485
      End
      Begin VB.TextBox txt_MOD_FL 
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
         Left            =   10020
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1215
         Width           =   585
      End
      Begin VB.TextBox txt_PLN_ORD_ITEM 
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
         Left            =   10020
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1590
         Width           =   1635
      End
      Begin VB.TextBox txt_PLN_ORD 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1590
         Width           =   1395
      End
      Begin VB.TextBox txt_MOD_FL_NAME 
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
         Left            =   10620
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1035
      End
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   17
         Left            =   8100
         Top             =   1965
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "库存销售"
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
         Index           =   31
         Left            =   8100
         Top             =   1590
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "原始计划订单序列号"
         Alignment       =   0
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
         Index           =   23
         Left            =   8100
         Top             =   1215
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "订单修改分类"
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
         Index           =   1
         Left            =   120
         Top             =   1965
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "订单种类"
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
         Index           =   10
         Left            =   120
         Top             =   840
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "订单客户"
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
         Index           =   11
         Left            =   120
         Top             =   1215
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "最终用户"
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
         Index           =   30
         Left            =   120
         Top             =   1590
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "原始计划订单号"
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
         Index           =   8
         Left            =   8100
         Top             =   465
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "交货日期"
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
         Index           =   16
         Left            =   120
         Top             =   465
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "目的地"
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
         Index           =   4
         Left            =   8100
         Top             =   840
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "销售部门"
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
         Index           =   9
         Left            =   120
         Top             =   90
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "客户"
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
         Index           =   26
         Left            =   8100
         Top             =   90
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "客户合同号"
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
      Begin VB.Image Img_Stock 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10020
         Stretch         =   -1  'True
         Top             =   1965
         Width           =   285
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  'Align Top
      Height          =   1545
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   2725
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_TOT_WGT 
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
         Height          =   310
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   435
         Width           =   1995
      End
      Begin VB.TextBox txt_ENDUSE_NAME 
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
         Left            =   7140
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   445
         Width           =   1365
      End
      Begin VB.TextBox txt_PROD_CD 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   90
         Width           =   405
      End
      Begin VB.TextBox txt_PROD_NAME 
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
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   90
         Width           =   1590
      End
      Begin VB.TextBox txt_ORD_THK 
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
         Height          =   310
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   800
         Width           =   670
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
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   445
         Width           =   615
      End
      Begin VB.TextBox txt_STDSPEC 
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
         Left            =   10920
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   90
         Width           =   2715
      End
      Begin VB.TextBox txt_ORD_WID 
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
         Height          =   310
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   800
         Width           =   670
      End
      Begin VB.TextBox txt_ORD_LEN 
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
         Height          =   310
         Left            =   7890
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   800
         Width           =   630
      End
      Begin VB.TextBox txt_PROD_DGR 
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
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   90
         Width           =   615
      End
      Begin VB.TextBox txt_ORD_SIZE 
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
         Height          =   310
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   800
         Width           =   1995
      End
      Begin VB.TextBox txt_PROD_DGR_NAME 
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
         Left            =   7140
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   90
         Width           =   1365
      End
      Begin VB.TextBox txt_THK_TGT 
         Height          =   310
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   435
         Width           =   2715
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   9000
         Top             =   435
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "轧制目标厚度"
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
         Index           =   32
         Left            =   4590
         Top             =   90
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "产品等级"
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
         Index           =   20
         Left            =   4590
         Top             =   450
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "订单用途"
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
         Index           =   18
         Left            =   9000
         Top             =   90
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "标准号/年度"
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
         Index           =   3
         Left            =   150
         Top             =   90
         Width           =   1845
         _ExtentX        =   3254
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
      End
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   5
         Left            =   150
         Top             =   435
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "订单产品重量"
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
         Index           =   36
         Left            =   150
         Top             =   800
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "订单尺寸"
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
         Index           =   6
         Left            =   4590
         Top             =   795
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "厚／宽／长"
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
         Index           =   46
         Left            =   120
         Top             =   1200
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "LP板宽度"
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
         Index           =   47
         Left            =   9000
         Top             =   810
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "LP板厚度1/2/3"
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
         Index           =   48
         Left            =   4560
         Top             =   1200
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "LP板长度1/2/3/4/5"
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
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_THK1 
         Height          =   315
         Left            =   10920
         TabIndex        =   90
         Tag             =   "订单宽度"
         Top             =   810
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   6
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_THK3 
         Height          =   315
         Left            =   12750
         TabIndex        =   91
         Tag             =   "订单宽度"
         Top             =   810
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   6
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_THK2 
         Height          =   315
         Left            =   11835
         TabIndex        =   92
         Tag             =   "订单宽度"
         Top             =   810
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         NumIntDigits    =   6
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN1 
         Height          =   315
         Left            =   6480
         TabIndex        =   93
         Tag             =   "Tot_Wgt"
         Top             =   1200
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
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN2 
         Height          =   315
         Left            =   7740
         TabIndex        =   94
         Tag             =   "Tot_Wgt"
         Top             =   1200
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
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN3 
         Height          =   315
         Left            =   9000
         TabIndex        =   95
         Tag             =   "Tot_Wgt"
         Top             =   1200
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
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN4 
         Height          =   315
         Left            =   10260
         TabIndex        =   96
         Tag             =   "Tot_Wgt"
         Top             =   1200
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
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN5 
         Height          =   315
         Left            =   11540
         TabIndex        =   97
         Tag             =   "Tot_Wgt"
         Top             =   1200
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
      Begin CSTextLibCtl.sidbEdit txt_ORD_LP_WID 
         Height          =   315
         Left            =   2040
         TabIndex        =   98
         Tag             =   "订单宽度"
         Top             =   1200
         Width           =   1995
         _Version        =   262145
         _ExtentX        =   3519
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
         NumIntDigits    =   6
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   615
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   820
      _Version        =   196609
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_ORD_STS 
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
         Left            =   7110
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   90
         Width           =   285
      End
      Begin VB.TextBox txt_ORD_STS_NAME 
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
         Left            =   7410
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   90
         Width           =   1725
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
         Left            =   2055
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   4
         Top             =   90
         Width           =   2250
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
         Left            =   4515
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   90
         Width           =   615
      End
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   2
         Left            =   150
         Top             =   90
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "订单号/序列号"
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
         Index           =   15
         Left            =   5220
         Top             =   90
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "订单状态"
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
         Index           =   0
         Left            =   4320
         Top             =   90
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   556
         Caption         =   "-"
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
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14490
      TabIndex        =   0
      Top             =   0
      Width           =   14550
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   600
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15420
         _ExtentX        =   27199
         _ExtentY        =   1058
         BandCount       =   1
         _CBWidth        =   15420
         _CBHeight       =   600
         _Version        =   "6.7.8988"
         Child1          =   "MenuTool"
         MinHeight1      =   540
         Width1          =   15360
         NewRow1         =   0   'False
         BandStyle1      =   1
         Begin MSComctlLib.Toolbar MenuTool 
            Height          =   540
            Left            =   30
            TabIndex        =   76
            Top             =   30
            Width           =   15360
            _ExtentX        =   27093
            _ExtentY        =   953
            ButtonWidth     =   1244
            ButtonHeight    =   953
            AllowCustomize  =   0   'False
            Style           =   1
            ImageList       =   "ImageList1"
            DisabledImageList=   "ImageList2"
            HotImageList    =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   9
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Clear"
                  Object.ToolTipText     =   "空界面"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Save"
                  Object.ToolTipText     =   "保存"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Delete"
                  Object.ToolTipText     =   "删除"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Copy"
                  Object.ToolTipText     =   "复制"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Paste"
                  Object.ToolTipText     =   "粘贴"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Line3"
                  Style           =   4
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Exit"
                  Object.ToolTipText     =   "退出"
                  ImageIndex      =   6
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8730
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   30
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":04B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":07BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":0ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":0CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":0DAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":109E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":1550
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9330
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   30
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":1664
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":1964
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":1BB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":1C99
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":1EA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":1FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQB0020C.frx":2215
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   45
      Left            =   6780
      Top             =   5085
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      Caption         =   "内径"
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
Attribute VB_Name = "AQB0020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量设计
'-- Program Name      质量设计现状详细查询
'-- Program ID        AQB0020C (Master-AQB0010C)
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       质量设计现状详细查询
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
Dim Old_ORD_NO As String            'Save The Previous ORD_NO For Next Use

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "PopMaster"

                 'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary )", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
                     Call Gp_Ms_Collection(txt_ORD_NO, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   Call Gp_Ms_Collection(txt_ORD_ITEM, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_ORD_STS, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ORD_STS_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'------------------------------------------------------------------------------------------------------------
                    Call Gp_Ms_Collection(txt_PROD_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_PROD_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   Call Gp_Ms_Collection(txt_PROD_DGR, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_PROD_DGR_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_STDSPEC, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_ENDUSE_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ENDUSE_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   Call Gp_Ms_Collection(txt_ORD_SIZE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_ORD_THK, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_ORD_WID, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_ORD_LEN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_THK_TGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_TOT_WGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
'------------------------------------------------------------------------------------------------------------
                    Call Gp_Ms_Collection(txt_CUST_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_CUST_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                       Call Gp_Ms_Collection(txt_PONO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_DEST_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_DEST_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_DEL_TO_DATE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ORD_CUST_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_ORD_CUST_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_DEPT_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_DEPT_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_END_CUST_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_END_CUST_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     Call Gp_Ms_Collection(txt_MOD_FL, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_MOD_FL_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_PLN_ORD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_PLN_ORD_ITEM, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_ORD_KND, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ORD_KND_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'------------------------------------------------------------------------------------------------------------
                    Call Gp_Ms_Collection(txt_INSP_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_INSP_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                      Call Gp_Ms_Collection(txt_INDIA, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   Call Gp_Ms_Collection(txt_PACK_WAY, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_PACK_WAY_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                      Call Gp_Ms_Collection(txt_STAMP, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_TRIM_FL, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_COLOR_STROKE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_MATR_FL, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     Call Gp_Ms_Collection(txt_UST_FL, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   Call Gp_Ms_Collection(txt_UST_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_CFM_SMS_PLT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_CFM_SMS_PLT_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_CUST_REQ_CODE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HTM_SHOT_BLAST, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_HTM_SHOT_BLAST_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_CFM_MILL_PLT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_CFM_MILL_PLT_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                      Call Gp_Ms_Collection(txt_METH1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                      Call Gp_Ms_Collection(txt_METH2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                      Call Gp_Ms_Collection(txt_METH3, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_METH1_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_METH2_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_METH3_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                      Call Gp_Ms_Collection(txt_CE_QS, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
'------------------------------------------------------------------------------------------------------------
                 Call Gp_Ms_Collection(txt_Design_STS, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_DESIGN_STS_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_DESIGN_DATE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_Nisco_Quality_No, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_CUST_SPEC_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     Call Gp_Ms_Collection(txt_stlgrd, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_STLGRD_GRP, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_STLGRD_Detail, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_DEV_STD_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(txt_MLT_STD_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_MILL_STD_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           
           Call Gp_Ms_Collection(txt_HTM_SHOT_BLAST_Q, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_HTM_SHOT_BLAST_NAME_Q, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_METH1_Q, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_METH2_Q, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_METH3_Q, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_COND1_Q, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_COND2_Q, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_COND3_Q, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_METH1_NAME_Q, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_METH2_NAME_Q, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_METH3_NAME_Q, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               
               'LP板尺寸 刘翔 2012.11.15
                 Call Gp_Ms_Collection(txt_ORD_LP_WID, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ORD_LP_THK1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ORD_LP_THK2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ORD_LP_THK3, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ORD_LP_LEN1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ORD_LP_LEN2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ORD_LP_LEN3, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ORD_LP_LEN4, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ORD_LP_LEN5, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    
    'MASTER Collection
     Mc1.Add Item:="AQB0020C.P_REFER", Key:="P-R"
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

End Sub



Private Sub Form_Activate()

    If Mc1("pControl").Item(1).Text = "" Then
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        pControl(1).SetFocus
    Else
     Call Form_Ref
     Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        If Me.ActiveControl.Name = "txt_ORD_ITEM" Then
            If Len(Trim(txt_ORD_NO.Text)) > 0 And Len(Trim(txt_ORD_ITEM.Text)) > 0 Then
                Call Form_Ref
            End If
        End If
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name, True)

    Call Popup_Menu_Setting

    Call Form_Define

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_ControlLock(Mc1("pControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_FormCenter(Me)

    Screen.MousePointer = vbDefault

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

'    Call AQB0010C.Form_Ref

End Sub

Public Sub Form_Ref()
 On Error GoTo Refer_Err

    If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        Call Check_Box_Change
    End If

    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Cls()
    
    Old_ORD_NO = Trim(txt_ORD_NO.Text)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    Call Check_Box_Change
    
    MenuTool.Buttons(4).Enabled = False    'save
    MenuTool.Buttons(5).Enabled = False    'delete
    MenuTool.Buttons(7).Enabled = False    'Copy
    MenuTool.Buttons(8).Enabled = True    'Paste

    pControl(1).SetFocus

End Sub

Private Sub MenuTool_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "Clear"              'Clear
            Call Form_Cls
        Case "Refer"               'refer
            Call Form_Ref
        Case "Copy"               'Copy
            Call Form_Copy
        Case "Paste"              'Paste
            Call Form_Paste
        Case "Exit"               'Exit
            Call Form_Exit
    End Select

End Sub


Public Sub Popup_Menu_Setting()

    Select Case Mid(sAuthority, 2, 3)

        Case "000"      'No Authority
            MenuTool.Buttons(4).Enabled = False                     'Save
            MenuTool.Buttons(5).Enabled = False                     'Delete
            MenuTool.Buttons(7).Enabled = False                     'Copy
            MenuTool.Buttons(8).Enabled = False                     'Paste

        Case "001"      'Delete Authority
            MenuTool.Buttons(4).Enabled = False                     'Save
            MenuTool.Buttons(7).Enabled = False                     'Copy
            MenuTool.Buttons(8).Enabled = False                     'Paste

        Case "010"      'Update Authority
            MenuTool.Buttons(5).Enabled = False                     'Delete
            MenuTool.Buttons(7).Enabled = False                     'Copy
            MenuTool.Buttons(8).Enabled = False                     'Paste

        Case "011"      'Update, Delete Authority
            MenuTool.Buttons(7).Enabled = False                     'Copy
            MenuTool.Buttons(8).Enabled = False                     'Paste

        Case "100"      'Insert Authority
            MenuTool.Buttons(5).Enabled = False                     'Delete

        Case "101"      'Insert, Delete Authority

        Case "110"      'Insert, Update Authority
            MenuTool.Buttons(5).Enabled = False                     'Delete

        Case "111"      'Insert, Update, Delete Authority

    End Select

End Sub

Private Sub Check_Box_Change()

    If txt_TRIM_FL.Text = "Y" Then
        Img_TRIM_FL.Picture = ImageList1.ListImages.Item(8).Picture
    Else
        Img_TRIM_FL.Picture = Nothing
    End If
                    
    If txt_STAMP.Text = "Y" Then
        Img_STAMP.Picture = ImageList1.ListImages.Item(8).Picture
    Else
        Img_STAMP.Picture = Nothing
    End If
                    
                    
    If txt_ORD_KND.Text = "S" Then
        Img_Stock.Picture = ImageList1.ListImages.Item(8).Picture
    Else
        Img_Stock.Picture = Nothing
    End If
    
    If txt_MATR_FL.Text = "Y" Then
        Img_MATR_FL.Picture = ImageList1.ListImages.Item(8).Picture
    Else
        Img_MATR_FL.Picture = Nothing
    End If

End Sub

Private Sub Form_Copy()
    
    Old_ORD_NO = Trim(txt_ORD_NO.Text)

End Sub

Private Sub Form_Paste()
    txt_ORD_NO.Text = ""
    txt_ORD_NO.Text = Old_ORD_NO

End Sub


Private Sub txt_CUST_CD_Change()
    If Len(Trim(txt_CUST_CD.Text)) = 0 Then txt_CUST_NAME.Text = ""
End Sub


Private Sub txt_DEPT_CD_Change()
    If Len(Trim(txt_DEPT_CD.Text)) = 0 Then txt_DEPT_NAME.Text = ""
End Sub

Private Sub txt_DESIGN_STS_Change()
    If Len(Trim(txt_Design_STS.Text)) = 0 Then txt_DESIGN_STS_NAME.Text = ""
End Sub

Private Sub txt_DEST_CD_Change()
    If Len(Trim(txt_DEST_CD.Text)) = 0 Then txt_DEST_NAME.Text = ""
End Sub

Private Sub txt_END_CUST_CD_Change()
    If Len(Trim(txt_END_CUST_CD.Text)) = 0 Then txt_END_CUST_NAME.Text = ""
End Sub

Private Sub txt_ENDUSE_CD_Change()
    If Len(Trim(txt_ENDUSE_CD.Text)) = 0 Then txt_ENDUSE_NAME.Text = ""
End Sub

Private Sub txt_INSP_CD_Change()
    If Len(Trim(txt_INSP_CD.Text)) = 0 Then txt_INSP_NAME.Text = ""
End Sub

Private Sub txt_MOD_FL_Change()
    If Len(Trim(txt_MOD_FL.Text)) = 0 Then txt_MOD_FL_NAME.Text = ""
End Sub

Private Sub txt_ORD_CUST_CD_Change()
    If Len(Trim(txt_ORD_CUST_CD.Text)) = 0 Then txt_ORD_CUST_NAME.Text = ""
End Sub

Private Sub txt_ORD_KND_Change()
    If Len(Trim(txt_ORD_KND.Text)) = 0 Then txt_ORD_KND_NAME.Text = ""
End Sub

Private Sub txt_ORD_STS_Change()
    If Len(Trim(txt_ORD_STS.Text)) = 0 Then txt_ORD_STS_NAME.Text = ""
End Sub

Private Sub txt_PACK_WAY_Change()
    If Len(Trim(txt_PACK_WAY.Text)) = 0 Then txt_PACK_WAY_NAME.Text = ""
End Sub

Private Sub txt_PROD_CD_Change()
    If Len(Trim(txt_PROD_CD.Text)) = 0 Then txt_PROD_NAME.Text = ""
End Sub

Private Sub txt_PROD_DGR_Change()
    If Len(Trim(txt_PROD_DGR.Text)) = 0 Then txt_PROD_DGR_NAME.Text = ""
End Sub

