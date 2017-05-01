VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AFM2030C 
   Caption         =   "返送钢水实绩修改及查询界面_AFM2030C"
   ClientHeight    =   9225
   ClientLeft      =   375
   ClientTop       =   2115
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame Frame3 
      Height          =   3435
      Left            =   180
      TabIndex        =   11
      Top             =   1800
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6059
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
         ItemData        =   "AFM2030C.frx":0000
         Left            =   6135
         List            =   "AFM2030C.frx":0002
         TabIndex        =   19
         Top             =   2490
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
         Left            =   6135
         MaxLength       =   2
         TabIndex        =   18
         Top             =   600
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
         Left            =   6135
         MaxLength       =   2
         TabIndex        =   17
         Top             =   1545
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
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   16
         Top             =   2490
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
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2490
         Width           =   2295
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
         Left            =   6615
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
         Width           =   1920
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
         Left            =   6615
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1545
         Width           =   1920
      End
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   255
         Top             =   600
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
         Left            =   255
         Top             =   2490
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
         Left            =   4725
         Top             =   1545
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
         Left            =   4725
         Top             =   2490
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
         Left            =   4725
         Top             =   600
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
         Left            =   1680
         TabIndex        =   20
         Top             =   600
         Width           =   1845
         _Version        =   262145
         _ExtentX        =   3254
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   255
         Top             =   1545
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
         Left            =   1680
         TabIndex        =   21
         Top             =   1545
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
         Left            =   2850
         TabIndex        =   22
         Top             =   1590
         Width           =   195
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   795
      Left            =   180
      TabIndex        =   5
      Top             =   180
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   1402
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
         Left            =   5145
         TabIndex        =   10
         Top             =   210
         Width           =   885
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
         Left            =   1815
         TabIndex        =   9
         Tag             =   "炉号"
         Top             =   210
         Width           =   1425
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
         ItemData        =   "AFM2030C.frx":0004
         Left            =   10035
         List            =   "AFM2030C.frx":0017
         TabIndex        =   8
         Tag             =   "班别"
         Text            =   "cbo_GROUP"
         Top             =   210
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
         ItemData        =   "AFM2030C.frx":002D
         Left            =   7785
         List            =   "AFM2030C.frx":003D
         TabIndex        =   7
         Tag             =   "班次"
         Text            =   "cbo_SHIFT"
         Top             =   210
         Width           =   735
      End
      Begin VB.TextBox cbo_EMP_CD 
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
         Left            =   12315
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   6
         Top             =   210
         Width           =   1095
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   540
         Top             =   210
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
         Left            =   3885
         Top             =   210
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
         Left            =   11010
         Top             =   210
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
         Left            =   6510
         Top             =   210
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
         Left            =   8760
         Top             =   210
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
      Left            =   690
      TabIndex        =   2
      Text            =   "1"
      Top             =   840
      Visible         =   0   'False
      Width           =   1590
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   11235
      Top             =   6375
      Width           =   2115
      _ExtentX        =   3731
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
      ForeColor       =   255
   End
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   4755
      Top             =   6375
      Width           =   2415
      _ExtentX        =   4260
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
      ForeColor       =   255
   End
   Begin Threed.SSCheck Chk_ss2 
      Height          =   315
      Left            =   9000
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   14737632
      Enabled         =   0   'False
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
      Left            =   300
      TabIndex        =   1
      Top             =   1440
      Width           =   1170
      _ExtentX        =   2064
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
      Left            =   13395
      TabIndex        =   3
      Top             =   6375
      Width           =   1620
      _ExtentX        =   2858
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
      Left            =   7215
      TabIndex        =   4
      Top             =   6375
      Width           =   1620
      _ExtentX        =   2858
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
   Begin Threed.SSFrame Frame4 
      Height          =   3435
      Left            =   8910
      TabIndex        =   12
      Top             =   1800
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   6059
      _Version        =   196609
      BackColor       =   14737632
      Begin InDate.ULabel ULabel32 
         Height          =   315
         Left            =   690
         Top             =   630
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
         Left            =   690
         Top             =   1560
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
         Left            =   2130
         TabIndex        =   23
         Top             =   1560
         Width           =   2190
         _Version        =   262145
         _ExtentX        =   3863
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_RET_SCR_WGT 
         Height          =   315
         Left            =   2130
         TabIndex        =   24
         Top             =   630
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
         Left            =   3300
         TabIndex        =   25
         Top             =   675
         Width           =   375
      End
   End
End
Attribute VB_Name = "AFM2030C"
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
          Call Gp_Ms_Collection(cbo_shift, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(cbo_group, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(cbo_emp_cd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      'Call Gp_Ms_Collection(dtp_WORK_DATE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
       Call Gp_Ms_Collection(txt_RET_TIME, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_ret_steel_wgt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_RET_LD_RES, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_RET_PLT_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_RET_DEST_CD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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

Private Sub cbo_RET_HEAT_NO_Change()
   If Len(cbo_RET_HEAT_NO.Text) = 8 Then
      cbo_RET_LD_NO.Text = Gf_FloatFind(M_CN1, "SELECT LADLE_NO FROM FP_CHARGE WHERE HEAT_NO = '" + cbo_RET_HEAT_NO.Text + "'")
   Else
      cbo_RET_LD_NO.Text = ""
   End If
End Sub

Private Sub cbo_RET_HEAT_NO_Click()
   If Len(cbo_RET_HEAT_NO.Text) = 8 Then
      cbo_RET_LD_NO.Text = Gf_FloatFind(M_CN1, "SELECT LADLE_NO FROM FP_CHARGE WHERE HEAT_NO = '" + cbo_RET_HEAT_NO.Text + "'")
   Else
      cbo_RET_LD_NO.Text = ""
   End If
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
    
    Frame4.Enabled = False

    Screen.MousePointer = vbDefault
    cbo_LD_OPEN_YN.AddItem "Y"
    cbo_LD_OPEN_YN.AddItem "N"
    cbo_emp_cd.Text = sUserID
    
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
    cbo_emp_cd.Text = sUserID
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
    End If

End Sub

Public Sub Form_Pro()
If Len(Trim(cbo_RET_HEAT_NO)) <> 8 Then
   MsgBox "炉号不正确！", vbCritical, "系统提示信息"
   Exit Sub
End If

    If Gf_Mc_Authority(sAuthority, Mc1) Then
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
           'Call Gp_Process_Exec
        End If
    End If

End Sub

Public Sub Form_Del()
Dim HEAT_NO As String
    If Gf_Ms_Del(M_CN1, Mc1) Then
       HEAT_NO = cbo_RET_HEAT_NO.Text
       Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
       Call Form_Cls
       cbo_RET_HEAT_NO.Text = HEAT_NO
    End If

End Sub

Private Sub Chk_ss1_Click(VALUE As Integer)

    If Chk_ss1.VALUE = ssCBUnchecked Then
       If Chk_ss2.VALUE = ssCBUnchecked Then
          Chk_ss1.VALUE = ssCBChecked
       End If
       Exit Sub
    End If

    If Chk_ss1.VALUE = -1 Then
        Chk_ss1.ForeColor = &HFF&
        Chk_ss2.ForeColor = &H808080
        Chk_ss2.VALUE = ssCBUnchecked
        Frame3.Enabled = True
        Frame4.Enabled = False
        Frame3.ShadowStyle = ssRaisedShadow
        Frame4.ShadowStyle = ssInsetShadow
        Text1.Text = "1"
        txt_RET_TIME.SetFocus
    Else
        Chk_ss1.VALUE = ssCBUnchecked
        Chk_ss2.VALUE = ssCBChecked
    End If
    
End Sub

Private Sub Chk_ss2_Click(VALUE As Integer)

    If Chk_ss2.VALUE = ssCBUnchecked Then
       If Chk_ss1.VALUE = ssCBUnchecked Then
          Chk_ss2.VALUE = ssCBChecked
       End If
       Exit Sub
    End If

    If Chk_ss2.VALUE = -1 Then
        Chk_ss2.ForeColor = &HFF&
        Chk_ss1.ForeColor = &H808080
        Chk_ss1.VALUE = ssCBUnchecked
        Frame3.Enabled = False
        Frame4.Enabled = True
        Frame3.ShadowStyle = ssInsetShadow
        Frame4.ShadowStyle = ssRaisedShadow
        Text1.Text = "2"
        txt_RET_SCR_WGT.SetFocus
    Else
        Chk_ss2.VALUE = ssCBUnchecked
        Chk_ss1.VALUE = ssCBChecked
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
