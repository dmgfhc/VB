VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AKM2030C 
   Caption         =   "返送钢水实绩修改及查询界面_AKM2030C"
   ClientHeight    =   9270
   ClientLeft      =   330
   ClientTop       =   2145
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame Frame4 
      Height          =   3375
      Left            =   7800
      TabIndex        =   11
      Top             =   1980
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin InDate.ULabel ULabel32 
         Height          =   315
         Left            =   510
         Top             =   390
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "废钢量"
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
         Left            =   510
         Top             =   1335
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "处理时间"
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
      Begin CSTextLibCtl.sitxEdit txt_RET_SCR_TIME 
         Height          =   315
         Left            =   1950
         TabIndex        =   23
         Top             =   1335
         Width           =   2070
         _Version        =   262145
         _ExtentX        =   3651
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
      End
      Begin CSTextLibCtl.sidbEdit txt_RET_SCR_WGT 
         Height          =   315
         Left            =   1950
         TabIndex        =   24
         Top             =   390
         Width           =   1125
         _Version        =   262145
         _ExtentX        =   1984
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
         NumIntDigits    =   12
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "t"
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
         Left            =   3120
         TabIndex        =   25
         Top             =   465
         Width           =   375
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   885
      Left            =   150
      TabIndex        =   5
      Top             =   210
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   1561
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.ComboBox cbo_RET_LD_NO 
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
         Left            =   4785
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cbo_RET_HEAT_NO 
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
         Left            =   1455
         TabIndex        =   9
         Tag             =   "炉号"
         Top             =   240
         Width           =   1485
      End
      Begin VB.ComboBox cbo_GROUP 
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
         ItemData        =   "AKM2030C.frx":0000
         Left            =   9930
         List            =   "AKM2030C.frx":0013
         TabIndex        =   8
         Text            =   "cbo_GROUP"
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cbo_SHIFT 
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
         IMEMode         =   1  'ON
         ItemData        =   "AKM2030C.frx":0029
         Left            =   7415
         List            =   "AKM2030C.frx":0039
         TabIndex        =   7
         Text            =   "cbo_SHIFT"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox cbo_EMP_CD 
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
         Left            =   12390
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   6
         Top             =   240
         Width           =   1065
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   180
         Top             =   240
         Width           =   1230
         _ExtentX        =   2170
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
         Left            =   3525
         Top             =   240
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "钢包号"
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
         Left            =   11130
         Top             =   240
         Width           =   1230
         _ExtentX        =   2170
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   6150
         Top             =   240
         Width           =   1230
         _ExtentX        =   2170
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   8670
         Top             =   240
         Width           =   1230
         _ExtentX        =   2170
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
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   13260
      TabIndex        =   2
      Text            =   "1"
      Top             =   1140
      Visible         =   0   'False
      Width           =   1590
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   12090
      Top             =   6120
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Caption         =   "累计废钢量"
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
      Left            =   4680
      Top             =   6120
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Caption         =   "累计回炉量"
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
   Begin Threed.SSCheck Chk_ss2 
      Height          =   315
      Left            =   7830
      TabIndex        =   0
      Top             =   1590
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
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
      Caption         =   "废钢处理"
   End
   Begin Threed.SSCheck Chk_ss1 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   1590
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
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
      Caption         =   "返送处理"
      Value           =   1
   End
   Begin Threed.SSPanel txt_RET_SCR_TOT 
      Height          =   315
      Left            =   13770
      TabIndex        =   3
      Top             =   6120
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   2
      PictureMaskColor=   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel txt_RET_STEEL_TOT 
      Height          =   315
      Left            =   6360
      TabIndex        =   4
      Top             =   6120
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   2
      PictureMaskColor=   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSFrame Frame3 
      Height          =   3375
      Left            =   150
      TabIndex        =   12
      Top             =   1980
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   5953
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.ComboBox cbo_LD_OPEN_YN 
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
         ItemData        =   "AKM2030C.frx":004B
         Left            =   5445
         List            =   "AKM2030C.frx":004D
         TabIndex        =   19
         Top             =   2280
         Width           =   585
      End
      Begin VB.TextBox txt_RET_PLT_CD 
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
         Left            =   5445
         MaxLength       =   2
         TabIndex        =   18
         Top             =   390
         Width           =   480
      End
      Begin VB.TextBox txt_RET_DEST_CD 
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
         Left            =   5445
         MaxLength       =   2
         TabIndex        =   17
         Top             =   1335
         Width           =   480
      End
      Begin VB.TextBox txt_RET_LD_RES 
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
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   16
         Top             =   2280
         Width           =   345
      End
      Begin VB.TextBox txt_ret_res_name 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000003&
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1845
      End
      Begin VB.TextBox txt_RET_PRC_NAME 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000003&
         Height          =   315
         Left            =   5925
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   390
         Width           =   1260
      End
      Begin VB.TextBox txt_RET_DEST_NAME 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000003&
         Height          =   315
         Left            =   5925
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1335
         Width           =   1260
      End
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   165
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "返送时间"
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
         Left            =   150
         Top             =   2280
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "返送原因"
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Left            =   4035
         Top             =   1335
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "返送目的工位"
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
         Left            =   4035
         Top             =   2280
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "钢包是否开浇"
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
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   4035
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "返送始发工位"
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
      Begin CSTextLibCtl.sitxEdit txt_RET_TIME 
         Height          =   315
         Left            =   1560
         TabIndex        =   20
         Top             =   390
         Width           =   1920
         _Version        =   262145
         _ExtentX        =   3387
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
         Text            =   "____-__-__ __:__"
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
         Mask            =   "____-__-__ __:__"
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   150
         Top             =   1335
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "返送钢水量"
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
      Begin CSTextLibCtl.sidbEdit txt_RET_STEEL_WGT 
         Height          =   315
         Left            =   1560
         TabIndex        =   21
         Top             =   1335
         Width           =   1125
         _Version        =   262145
         _ExtentX        =   1984
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
         NumIntDigits    =   12
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2760
         TabIndex        =   22
         Top             =   1380
         Width           =   195
      End
   End
End
Attribute VB_Name = "AKM2030C"
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
'-- Program Name      RETURN STEEL
'-- Program ID        AFM2030C
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2003.8.27
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
Public sDateTime As String              'Active Form Authority Setting
Public sQuery As String

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection

Private Sub Form_Define()
   'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")

    Call Gp_Ms_Collection(cbo_RET_HEAT_NO, "p", "n", "m", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(cbo_RET_LD_NO, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(Text1, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(CBO_SHIFT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(cbo_GROUP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(cbo_EMP_CD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      'Call Gp_Ms_Collection(dtp_WORK_DATE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
       Call Gp_Ms_Collection(txt_RET_TIME, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_RET_STEEL_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_RET_LD_RES, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_RET_PLT_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_RET_DEST_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(cbo_LD_OPEN_YN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_RET_SCR_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_RET_SCR_TIME, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_RET_SCR_TOT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_RET_STEEL_TOT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
     Mc1.Add Item:="AFM2030C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AFM2030C.P_REFER", Key:="P-R"
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


Private Sub txt_RET_DEST_CD_Change()
    If Len(Trim(txt_RET_DEST_CD)) = txt_RET_DEST_CD.MaxLength Then
        txt_RET_DEST_NAME.Text = Gf_ComnNameFind(M_CN1, "C0002", Trim(txt_RET_DEST_CD.Text), 2)
    Else
        txt_RET_DEST_NAME.Text = ""
    End If
End Sub

Private Sub txt_RET_DEST_CD_DblClick()

    Call txt_RET_DEST_CD_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_RET_LD_RES_Change()
    If Len(Trim(txt_RET_LD_RES)) = txt_RET_LD_RES.MaxLength Then
        txt_ret_res_name.Text = Gf_ComnNameFind(M_CN1, "F0003", Trim(txt_RET_LD_RES.Text), 2)
    Else
        txt_ret_res_name.Text = ""
    End If
End Sub

Private Sub txt_RET_LD_RES_DblClick()

    Call txt_RET_LD_RES_KeyUP(vbKeyF4, 0)

End Sub

Private Sub txt_RET_LD_RES_KeyUP(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
         DD.sWitch = "MS"
         DD.sKey = "F0003"
         DD.rControl.Add Item:=txt_RET_LD_RES
         DD.rControl.Add Item:=txt_ret_res_name
         DD.nameType = "2"
         Call Gf_Common_DD(M_CN1, KeyCode)
    End If
    
End Sub

Private Sub Form_Activate()

   Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
   
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    Call Gf_ComboAdd(M_CN1, cbo_RET_HEAT_NO, "select RET_HEAT_NO from FP_RTSTEEL")
    Call Gf_ComboAdd(M_CN1, cbo_RET_LD_NO, "select CD from ZP_CD WHERE CD_MANA_NO='F0004' AND CD LIKE 'S%'")

    sAuthority = Gf_Pgm_Authority(Me.Name)
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Screen.MousePointer = vbDefault
    cbo_LD_OPEN_YN.AddItem "Y"
    cbo_LD_OPEN_YN.AddItem "N"
    cbo_EMP_CD.Text = sUserID
    
    Frame3.Enabled = True
    Frame4.Enabled = False
    Frame3.ShadowStyle = ssRaisedShadow
    Frame4.ShadowStyle = ssInsetShadow
        
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

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    pControl(1).SetFocus
    cbo_EMP_CD.Text = sUserID
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)

End Sub

Public Sub Form_Ref()

    If Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        If cbo_EMP_CD.Text = "" Then
           cbo_EMP_CD.Text = sUserID
        End If
    End If

End Sub

Public Sub Form_Pro()
cbo_EMP_CD.Text = sUserID
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
           'Call Gp_Process_Exec
        End If
    End If

End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub

Private Sub Chk_ss1_Click(Value As Integer)

    If Chk_ss1.Value = ssCBUnchecked Then
       If Chk_ss2.Value = ssCBUnchecked Then
          Chk_ss1.Value = ssCBChecked
       End If
       Exit Sub
    End If

    If Chk_ss1.Value = -1 Then
        Chk_ss1.ForeColor = &HFF&
        Chk_ss2.ForeColor = &H808080
        Chk_ss2.Value = ssCBUnchecked
        Frame3.Enabled = True
        Frame4.Enabled = False
        Frame3.ShadowStyle = ssRaisedShadow
        Frame4.ShadowStyle = ssInsetShadow
        Text1.Text = "1"
        txt_RET_DEST_CD.SetFocus
    Else
        Chk_ss1.Value = ssCBUnchecked
        Chk_ss2.Value = ssCBChecked
    End If
    
End Sub

Private Sub Chk_ss2_Click(Value As Integer)

    If Chk_ss2.Value = ssCBUnchecked Then
       If Chk_ss1.Value = ssCBUnchecked Then
          Chk_ss2.Value = ssCBChecked
       End If
       Exit Sub
    End If

    If Chk_ss2.Value = -1 Then
        Chk_ss2.ForeColor = &HFF&
        Chk_ss1.ForeColor = &H808080
        Chk_ss1.Value = ssCBUnchecked
        Frame3.Enabled = False
        Frame4.Enabled = True
        Frame3.ShadowStyle = ssInsetShadow
        Frame4.ShadowStyle = ssRaisedShadow
        Text1.Text = "2"
        txt_RET_SCR_WGT.SetFocus
    Else
        Chk_ss2.Value = ssCBUnchecked
        Chk_ss1.Value = ssCBChecked
    End If
    
End Sub

Private Sub txt_RET_DEST_CD_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "C0002"
        DD.rControl.Add Item:=txt_RET_DEST_CD
        DD.rControl.Add Item:=txt_RET_DEST_NAME
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
             
    End If
    
End Sub

Private Sub txt_RET_PLT_CD_Change()
    If Len(Trim(txt_RET_PLT_CD)) = txt_RET_PLT_CD.MaxLength Then
        txt_RET_PRC_NAME.Text = Gf_ComnNameFind(M_CN1, "C0002", Trim(txt_RET_PLT_CD.Text), 2)
    Else
        txt_RET_PRC_NAME.Text = ""
    End If
End Sub

Private Sub txt_RET_PLT_CD_DblClick()

    Call txt_RET_PLT_CD_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_RET_PLT_CD_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
      DD.sWitch = "MS"
      DD.sKey = "C0002"
      DD.rControl.Add Item:=txt_RET_PLT_CD
      DD.rControl.Add Item:=txt_RET_PRC_NAME

      DD.nameType = "2"
     
      Call Gf_Common_DD(M_CN1, KeyCode)
     
    End If
    
End Sub

Private Sub txt_RET_SCR_TIME_DblClick()

    txt_RET_SCR_TIME.RawData = Format(Now, "YYYYMMDDHHMMSS")
        
End Sub

Private Sub txt_RET_TIME_DblClick()

    txt_RET_TIME.RawData = Format(Now, "YYYYMMDDHHMMSS")
        
End Sub
