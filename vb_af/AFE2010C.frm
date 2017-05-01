VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AFE2010C 
   Caption         =   "LF实绩修改及查询界面_AFE2010C"
   ClientHeight    =   9225
   ClientLeft      =   90
   ClientTop       =   2250
   ClientWidth     =   15225
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame2 
      Height          =   1095
      Left            =   210
      TabIndex        =   6
      Top             =   1080
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   1931
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.ComboBox cbo_shift 
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
         ItemData        =   "AFE2010C.frx":0000
         Left            =   7815
         List            =   "AFE2010C.frx":0002
         TabIndex        =   23
         Tag             =   "班次"
         Top             =   180
         Width           =   600
      End
      Begin VB.ComboBox cbo_group_cd 
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
         ItemData        =   "AFE2010C.frx":0004
         Left            =   10035
         List            =   "AFE2010C.frx":0006
         TabIndex        =   22
         Tag             =   "班别"
         Top             =   180
         Width           =   630
      End
      Begin VB.TextBox txt_act_steel_grd 
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
         Left            =   1450
         MaxLength       =   11
         TabIndex        =   21
         Top             =   600
         Width           =   1290
      End
      Begin VB.TextBox txt_dir_steel_grd 
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
         Left            =   1450
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   1290
      End
      Begin VB.TextBox cbo_ref_proc 
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
         Left            =   12240
         TabIndex        =   19
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txt_emp_cd 
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
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   900
      End
      Begin VB.TextBox txt_mlt_prod_cd 
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
         Left            =   7815
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   1440
      End
      Begin VB.TextBox txt_stlgrd_c 
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
         Left            =   2745
         TabIndex        =   16
         Top             =   600
         Width           =   3090
      End
      Begin VB.TextBox txt_stlgrd_n 
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
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   3090
      End
      Begin InDate.ULabel ULabel84 
         Height          =   315
         Left            =   6570
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
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
      Begin InDate.ULabel ULabel85 
         Height          =   315
         Left            =   8790
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
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
      Begin InDate.ULabel ULabel86 
         Height          =   315
         Left            =   10980
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
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
         Left            =   10980
         Top             =   600
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "上道工序"
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
         Left            =   210
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   210
         Top             =   600
         Width           =   1200
         _ExtentX        =   2117
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Left            =   6570
         Top             =   600
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "工艺路线"
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   915
      Left            =   210
      TabIndex        =   5
      Top             =   120
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   1614
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.ComboBox cbo_heat_no 
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
         ItemData        =   "AFE2010C.frx":0008
         Left            =   1455
         List            =   "AFE2010C.frx":000A
         TabIndex        =   14
         Tag             =   "炉号"
         Top             =   300
         Width           =   1425
      End
      Begin VB.ComboBox cbo_ld_id 
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
         Left            =   10275
         TabIndex        =   13
         Tag             =   "钢包号"
         Top             =   300
         Width           =   750
      End
      Begin VB.ComboBox cbo_re_cd 
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
         ItemData        =   "AFE2010C.frx":000C
         Left            =   5400
         List            =   "AFE2010C.frx":000E
         TabIndex        =   12
         Tag             =   "再处理"
         Top             =   300
         Width           =   660
      End
      Begin VB.ComboBox cbo_prc_line 
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
         ItemData        =   "AFE2010C.frx":0010
         Left            =   7830
         List            =   "AFE2010C.frx":0012
         TabIndex        =   11
         Tag             =   "LF"
         Top             =   300
         Width           =   615
      End
      Begin VB.CommandButton cmd_down 
         Caption         =   "▼"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   450
         Width           =   375
      End
      Begin VB.CommandButton cmd_up 
         Caption         =   "▲"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   150
         Width           =   375
      End
      Begin InDate.ULabel ULabel82 
         Height          =   315
         Left            =   9030
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
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
      Begin InDate.ULabel ULabel83 
         Height          =   315
         Left            =   210
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   4170
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "再处理"
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
         Left            =   6585
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "LF"
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
   End
   Begin VB.TextBox txt_oper 
      Height          =   270
      Left            =   9930
      TabIndex        =   4
      Text            =   "1"
      Top             =   885
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txt_plt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10290
      TabIndex        =   1
      Text            =   "B1"
      Top             =   870
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txt_proc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11730
      TabIndex        =   0
      Text            =   "BD"
      Top             =   855
      Visible         =   0   'False
      Width           =   1185
   End
   Begin Threed.SSCheck Chk_ss1 
      Height          =   330
      Left            =   315
      TabIndex        =   2
      Top             =   2355
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1.进程"
      Value           =   1
   End
   Begin Threed.SSCheck Chk_ss2 
      Height          =   330
      Left            =   315
      TabIndex        =   3
      Top             =   4230
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2.实绩"
   End
   Begin Threed.SSFrame Frame3 
      Height          =   1215
      Left            =   210
      TabIndex        =   7
      Top             =   2730
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   2143
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   210
         Top             =   210
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "钢包到达时间"
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
         Left            =   5460
         Top             =   210
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "开始时间"
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
         Left            =   5460
         Top             =   705
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "结束时间"
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
         Left            =   210
         Top             =   705
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "钢包离开时间"
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
      Begin CSTextLibCtl.sitxEdit txt_end_date 
         Height          =   315
         Left            =   6825
         TabIndex        =   24
         Top             =   705
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __:__:__"
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
      Begin CSTextLibCtl.sitxEdit txt_start_date 
         Height          =   315
         Left            =   6825
         TabIndex        =   25
         Top             =   210
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __:__:__"
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
      Begin CSTextLibCtl.sitxEdit txt_ld_arrv_date 
         Height          =   315
         Left            =   1575
         TabIndex        =   26
         Top             =   210
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __:__:__"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
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
      Begin CSTextLibCtl.sitxEdit txt_ld_dep_date 
         Height          =   315
         Left            =   1575
         TabIndex        =   27
         Top             =   705
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __:__:__"
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
   End
   Begin Threed.SSFrame Frame4 
      Height          =   4425
      Left            =   210
      TabIndex        =   8
      Top             =   4560
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   7805
      _Version        =   196609
      BackColor       =   14737632
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
         Left            =   12315
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   3765
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
         Left            =   11010
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   3765
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
         Left            =   9765
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   3765
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
         Left            =   8475
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   3765
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
         Left            =   7185
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   3765
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
         Left            =   5895
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   3765
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
         Left            =   4605
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3765
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
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   3765
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
         Left            =   1965
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3765
         Width           =   870
      End
      Begin VB.TextBox txt_wire_cd1 
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
         Left            =   1575
         TabIndex        =   31
         Top             =   2400
         Width           =   825
      End
      Begin VB.TextBox txt_wire_cd2 
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
         Left            =   8205
         TabIndex        =   30
         Top             =   2400
         Width           =   825
      End
      Begin VB.TextBox txt_wire_cd3 
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
         Left            =   1575
         TabIndex        =   29
         Top             =   2760
         Width           =   825
      End
      Begin VB.TextBox txt_wire_cd4 
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
         Left            =   8205
         TabIndex        =   28
         Top             =   2760
         Width           =   825
      End
      Begin CSTextLibCtl.sidbEdit txt_arrv_steel_wgt 
         Height          =   315
         Left            =   1575
         TabIndex        =   32
         Top             =   720
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         NumIntDigits    =   12
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   210
         Top             =   720
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "开始时钢水量"
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
         Left            =   210
         Top             =   1500
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "开始时渣量"
         Alignment       =   1
         BackColor       =   16761024
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
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   210
         Top             =   1110
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "结束时钢水量"
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
         Left            =   210
         Top             =   1890
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "结束时渣量"
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
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   3510
         Top             =   720
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "空包重量"
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
         Left            =   3510
         Top             =   1110
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "开始时温度"
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
         Left            =   3510
         Top             =   1500
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "结束时温度"
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
         Left            =   3510
         Top             =   1890
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "钢包包龄"
         Alignment       =   1
         BackColor       =   16761024
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
         Left            =   6840
         Top             =   720
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "氮气用量"
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
         Left            =   6840
         Top             =   1110
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "氩气用量"
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
         Left            =   210
         Top             =   2400
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "喂丝种类 1"
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
         Left            =   3510
         Top             =   2400
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "喂丝重量 1"
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
         Left            =   6840
         Top             =   1500
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "电极消耗"
         Alignment       =   1
         BackColor       =   16761024
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
         Left            =   10260
         Top             =   720
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "通电时间"
         Alignment       =   1
         BackColor       =   16761024
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
         Left            =   10260
         Top             =   1110
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "通电电压"
         Alignment       =   1
         BackColor       =   16761024
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
      Begin InDate.ULabel ULabel50 
         Height          =   315
         Left            =   10260
         Top             =   1500
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "通电量"
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
      Begin InDate.ULabel ULabel29 
         Height          =   315
         Left            =   210
         Top             =   3765
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "LF成分(%)"
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
         Left            =   6795
         Top             =   3765
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
         Left            =   5505
         Top             =   3765
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
         Left            =   4215
         Top             =   3765
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
         Left            =   2910
         Top             =   3765
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
         Left            =   1575
         Top             =   3765
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
         Left            =   8100
         Top             =   3765
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
         Left            =   9390
         Top             =   3765
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
         Left            =   10635
         Top             =   3765
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
         Left            =   11940
         Top             =   3765
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
      Begin CSTextLibCtl.sidbEdit txt_dep_steel_wgt 
         Height          =   315
         Left            =   1575
         TabIndex        =   42
         Top             =   1110
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         NumIntDigits    =   12
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_sta_slag_wgt 
         Height          =   315
         Left            =   1575
         TabIndex        =   43
         Top             =   1500
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         NumIntDigits    =   12
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_end_slag_wgt 
         Height          =   315
         Left            =   1575
         TabIndex        =   44
         Top             =   1890
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         NumIntDigits    =   12
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ld_empty_wgt 
         Height          =   315
         Left            =   4875
         TabIndex        =   45
         Top             =   720
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         NumIntDigits    =   12
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ld_arrv_temp 
         Height          =   315
         Left            =   4875
         TabIndex        =   46
         Top             =   1110
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ld_dep_temp 
         Height          =   315
         Left            =   4875
         TabIndex        =   47
         Top             =   1500
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ld_life 
         Height          =   315
         Left            =   4875
         TabIndex        =   48
         Top             =   1890
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_n2_gas_consp 
         Height          =   315
         Left            =   8205
         TabIndex        =   49
         Top             =   720
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         NumIntDigits    =   6
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ar_gas_consp 
         Height          =   315
         Left            =   8205
         TabIndex        =   50
         Top             =   1110
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         NumIntDigits    =   6
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_elect_consp 
         Height          =   315
         Left            =   8205
         TabIndex        =   51
         Top             =   1500
         Visible         =   0   'False
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         NumIntDigits    =   6
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_pow_on_ts 
         Height          =   315
         Left            =   11625
         TabIndex        =   52
         Top             =   720
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_elect_press 
         Height          =   315
         Left            =   11625
         TabIndex        =   53
         Top             =   1110
         Visible         =   0   'False
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         NumIntDigits    =   6
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_elect_energy 
         Height          =   315
         Left            =   11625
         TabIndex        =   54
         Top             =   1500
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         NumIntDigits    =   6
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_wire_lth1 
         Height          =   315
         Left            =   4875
         TabIndex        =   55
         Top             =   2400
         Width           =   795
         _Version        =   262145
         _ExtentX        =   1402
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   210
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "实绩发生时间"
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
      Begin CSTextLibCtl.sitxEdit txt_OCCR_DATE 
         Height          =   315
         Left            =   1575
         TabIndex        =   56
         Top             =   240
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __:__:__"
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   6840
         Top             =   2400
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "喂丝种类 2"
         Alignment       =   1
         BackColor       =   16761024
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
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   10260
         Top             =   2400
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "喂丝重量 2"
         Alignment       =   1
         BackColor       =   16761024
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
      Begin CSTextLibCtl.sidbEdit txt_wire_lth2 
         Height          =   315
         Left            =   11625
         TabIndex        =   57
         Top             =   2400
         Width           =   795
         _Version        =   262145
         _ExtentX        =   1402
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   210
         Top             =   2760
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "喂丝种类 3"
         Alignment       =   1
         BackColor       =   16761024
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
         Left            =   3510
         Top             =   2760
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "喂丝重量 3"
         Alignment       =   1
         BackColor       =   16761024
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
      Begin CSTextLibCtl.sidbEdit txt_wire_lth3 
         Height          =   315
         Left            =   4875
         TabIndex        =   58
         Top             =   2760
         Width           =   795
         _Version        =   262145
         _ExtentX        =   1402
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   6840
         Top             =   2760
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "喂丝种类 4"
         Alignment       =   1
         BackColor       =   16761024
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
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   10260
         Top             =   2760
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "喂丝重量 4"
         Alignment       =   1
         BackColor       =   16761024
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
      Begin CSTextLibCtl.sidbEdit txt_wire_lth4 
         Height          =   315
         Left            =   11625
         TabIndex        =   59
         Top             =   2760
         Width           =   795
         _Version        =   262145
         _ExtentX        =   1402
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_upd_date 
         Height          =   315
         Left            =   1575
         TabIndex        =   60
         Top             =   3180
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __:__:__"
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
         Mask            =   "____-__-__ __:__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Left            =   210
         Top             =   3180
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "上次修改时间"
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2685
         TabIndex        =   77
         Top             =   780
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5940
         TabIndex        =   76
         Top             =   750
         Width           =   180
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2685
         TabIndex        =   75
         Top             =   1950
         Width           =   180
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2685
         TabIndex        =   74
         Top             =   1545
         Width           =   180
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2685
         TabIndex        =   73
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Kw/h"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   12705
         TabIndex        =   72
         Top             =   1545
         Width           =   420
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg/t"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9270
         TabIndex        =   71
         Top             =   1575
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5940
         TabIndex        =   70
         Top             =   1560
         Width           =   210
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5940
         TabIndex        =   69
         Top             =   1170
         Width           =   210
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "m3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9285
         TabIndex        =   68
         Top             =   750
         Width           =   210
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "m3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   9300
         TabIndex        =   67
         Top             =   1170
         Width           =   210
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12720
         TabIndex        =   66
         Top             =   735
         Width           =   375
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12765
         TabIndex        =   65
         Top             =   1125
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5685
         TabIndex        =   64
         Top             =   2400
         Width           =   345
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12465
         TabIndex        =   63
         Top             =   2445
         Width           =   345
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5685
         TabIndex        =   62
         Top             =   2760
         Width           =   345
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12465
         TabIndex        =   61
         Top             =   2760
         Width           =   345
      End
   End
End
Attribute VB_Name = "AFE2010C"
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
'-- Program Name      LF
'-- Program ID        AFE2010C
'-- Document No
'-- Designer          Nisco
'-- Coder             Nisco
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
Public sDateTime As String              'Active Form Authority Setting
Public sQuery_Rt As String
Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection
Dim pControl1 As New Collection      'Master Primary Key Collection
Dim nControl1 As New Collection      'Master Necessary Collection
Dim mControl1 As New Collection      'Master Maxlength check Collection
Dim iControl1 As New Collection      'Master Insert Collection
Dim rControl1 As New Collection      'Master Refer Collection
Dim cControl1 As New Collection      'Master Copy Collection
Dim aControl1 As New Collection      'Master -> Spread Collection
Dim lControl1 As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection


Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"              'form类型
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       
           Call Gp_Ms_Collection(cbo_heat_no, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(cbo_re_cd, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(cbo_prc_line, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(cbo_ld_id, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(cbo_shift, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(cbo_group_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_emp_cd, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(cbo_ref_proc, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_dir_steel_grd, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_stlgrd_n, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_act_steel_grd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_stlgrd_c, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_mlt_prod_cd, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ld_arrv_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_start_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_end_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ld_dep_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_OCCR_DATE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_arrv_steel_wgt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_dep_steel_wgt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_sta_slag_wgt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_end_slag_wgt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ld_empty_wgt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ld_arrv_temp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ld_dep_temp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ld_life, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_n2_gas_consp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ar_gas_consp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_pow_on_ts, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_elect_energy, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_wire_cd1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_wire_lth1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_wire_cd2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_wire_lth2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_wire_cd3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_wire_lth3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_wire_cd4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_wire_lth4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_element_content_c, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_element_content_mn, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_element_content_si, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_element_content_p, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_element_content_s, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_element_content_al, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_element_content_h, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_element_content_n, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_element_content_o, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_upd_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_plt, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_proc, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_oper, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
    'MASTER Collection
     Mc1.Add Item:="AFE2010C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AFE2010C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
  
           Call Gp_Ms_Collection(cbo_heat_no, "p", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(cbo_prc_line, " ", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             Call Gp_Ms_Collection(cbo_ld_id, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_dir_steel_grd, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_stlgrd_n, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_act_steel_grd, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_stlgrd_c, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_mlt_prod_cd, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_ld_arrv_date, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(txt_start_date, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_end_date, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_ld_dep_date, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_proc, "p", " ", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              
     Mc2.Add Item:="AFE2010C.P_REFER1", Key:="P-R"
     Mc2.Add Item:=pControl1, Key:="pControl"
     Mc2.Add Item:=nControl1, Key:="nControl"
     Mc2.Add Item:=mControl1, Key:="mControl"
     
     Mc2.Add Item:=iControl1, Key:="iControl"
     Mc2.Add Item:=rControl1, Key:="rControl"
     Mc2.Add Item:=cControl1, Key:="cControl"
     Mc2.Add Item:=aControl1, Key:="aControl"
     Mc2.Add Item:=lControl1, Key:="lControl"
  
     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0

End Sub

Private Sub cbo_HEAT_NO_Change()
   If Len(cbo_heat_no.Text) = 8 Then
      cbo_ld_id.Text = Gf_FloatFind(M_CN1, "SELECT LADLE_NO FROM FP_CHARGE WHERE HEAT_NO = '" + cbo_heat_no.Text + "'")
      If cbo_ld_id.Text = "0" Then
         cbo_ld_id.Text = ""
      End If
   Else
      cbo_ld_id.Text = ""
   End If
End Sub

Private Sub cbo_HEAT_NO_Click()
   If Len(cbo_heat_no.Text) = 8 Then
      'cbo_LD_ID.Text = Gf_FloatFind(M_CN1, "SELECT LADLE_NO FROM FP_CHARGE WHERE HEAT_NO = '" + cbo_HEAT_NO.Text + "'")
      If cbo_ld_id.Text = "0" Then
         cbo_ld_id.Text = ""
      End If
   Else
      cbo_ld_id.Text = ""
   End If
End Sub

Private Sub cbo_prc_line_Change()
'    Dim AdoRs  As ADODB.Recordset
'    Dim sQuery As String
'   'Ado Setting
'
'    If cbo_heat_no.Enabled = True Then
'
'
'        M_CN1.CursorLocation = adUseServer
'        Set AdoRs = New ADODB.Recordset
'        'Set adoCmd = New adodb.Command
'
'        'adoCmd.CommandType = adCmdText
'        'Set adoCmd.ActiveConnection = M_CN1
'
'        sQuery = "         SELECT HEAT_MANA_NO"
'        sQuery = sQuery & "  FROM EP_CHARGE_INS "
'        sQuery = sQuery & " WHERE MLT_PROC_CD LIKE '%BD" & cbo_prc_line.Text & "%'"
'        sQuery = sQuery & "   AND PRC_STS IN('A','B')"
'        sQuery = sQuery & "   AND ROWNUM <= 10 "
'        sQuery = sQuery & "   ORDER BY HEAT_EDT_SEQ"
'
'        'Ado Execute
'        AdoRs.Open sQuery, M_CN1, adOpenKeyset
'        'AdoRs.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
'        cbo_heat_no.Clear
'        If Not AdoRs.BOF And Not AdoRs.EOF Then
'            While Not AdoRs.EOF
'
'                If VarType(AdoRs.Fields(0)) <> vbNull Then
'                   'If AdoRs.Fields(1) > "0" Then
'                      cbo_heat_no.AddItem AdoRs.Fields(0)
'                   'Else
'                      'cbo_HEAT_NO.AddItem AdoRs.Fields(0)
'                   'End If
'                End If
'                AdoRs.MoveNext
'            Wend
'        End If
'
'        AdoRs.Close
'        Set AdoRs = Nothing
'
'
'        'Call Gf_HeatNo_ComboAdd(M_CN1, cbo_HEAT_NO, "FP_LFRSLT", "DEP_STEEL_WGT", Trim(cbo_prc_line.Text))
'        If cbo_heat_no.ListCount <> 0 And Trim(cbo_heat_no.Text) = "" Then
'           cbo_heat_no.ListIndex = 0
'           cbo_re_cd.Text = "1"
'        End If
'    End If
End Sub

Private Sub cbo_prc_line_Click()

    Call cbo_prc_line_Change
    
End Sub

Private Sub cbo_ref_proc_DblClick()

    Call cbo_ref_proc_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub cbo_ref_proc_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "C0002"
        DD.rControl.Add Item:=cbo_ref_proc
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
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
        Chk_ss2.ForeColor = &H0&
        Chk_ss2.VALUE = ssCBUnchecked
        Frame3.Enabled = True
        Frame4.Enabled = False
        Frame3.ShadowStyle = ssRaisedShadow
        Frame4.ShadowStyle = ssInsetShadow
        txt_oper = "1"
        txt_ld_arrv_date.SetFocus
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
        Chk_ss1.ForeColor = &H0&
        Chk_ss2.ForeColor = &HFF&
        Chk_ss1.VALUE = ssCBUnchecked
        Frame3.Enabled = False
        Frame4.Enabled = True
        Frame3.ShadowStyle = ssInsetShadow
        Frame4.ShadowStyle = ssRaisedShadow
        txt_oper = "2"
        txt_OCCR_DATE.SetFocus
    Else
        Chk_ss2.VALUE = ssCBUnchecked
        Chk_ss1.VALUE = ssCBChecked
    End If

End Sub

Private Sub cmd_up_Click()
  Dim V_HEAT_SEQ_C As String
  Dim V_HEAT_SEQ_N, V_LEN As Integer
  Dim V_HEAT_NO As String
  
'  If Trim(cbo_heat_no.Text) = "" Then
'     Exit Sub
'  End If
'
'  V_HEAT_SEQ_C = Mid(cbo_heat_no, 4, 5)
'  V_HEAT_SEQ_N = Val(V_HEAT_SEQ_C)
'  V_LEN = Len(Trim(V_HEAT_SEQ_N + 1))
'  If V_LEN = 1 Then
'     V_HEAT_SEQ_C = "0000" + Trim(STR(V_HEAT_SEQ_N + 1))
'  ElseIf V_LEN = 2 Then
'     V_HEAT_SEQ_C = "000" + Trim(STR(V_HEAT_SEQ_N + 1))
'  ElseIf V_LEN = 3 Then
'     V_HEAT_SEQ_C = "00" + Trim(STR(V_HEAT_SEQ_N + 1))
'  ElseIf V_LEN = 4 Then
'     V_HEAT_SEQ_C = "0" + Trim(STR(V_HEAT_SEQ_N + 1))
'  End If
'
'  cbo_heat_no = Mid(cbo_heat_no, 1, 3) + V_HEAT_SEQ_C
'  V_HEAT_NO = cbo_heat_no
  If Trim(cbo_heat_no.Text) = "" Or Mid(cbo_heat_no, 4, 5) = "99999" Then
     Exit Sub
  End If
    
  cbo_heat_no = Mid(cbo_heat_no, 1, 3) + Format(Val(Mid(cbo_heat_no, 4, 5)) + 1, "00000")
  V_HEAT_NO = cbo_heat_no
    
  Call Form_Cls
  cbo_heat_no = V_HEAT_NO
  cbo_re_cd.Text = "1"
  Call Form_Ref
End Sub

Private Sub cmd_down_Click()
   Dim V_HEAT_SEQ_C As String
   Dim V_HEAT_SEQ_N, V_LEN As Integer
   Dim V_HEAT_NO As String
  
'  If Trim(cbo_heat_no.Text) = "" Then
'     Exit Sub
'  End If
'
'  V_HEAT_SEQ_C = Mid(cbo_heat_no, 4, 5)
'  V_HEAT_SEQ_N = Val(V_HEAT_SEQ_C)
'  V_LEN = Len(Trim(STR(V_HEAT_SEQ_N - 1)))
'  If V_LEN = 1 Then
'     V_HEAT_SEQ_C = "0000" + Trim(STR(V_HEAT_SEQ_N - 1))
'  ElseIf V_LEN = 2 Then
'     V_HEAT_SEQ_C = "000" + Trim(STR(V_HEAT_SEQ_N - 1))
'  ElseIf V_LEN = 3 Then
'     V_HEAT_SEQ_C = "00" + Trim(STR(V_HEAT_SEQ_N - 1))
'  ElseIf V_LEN = 4 Then
'     V_HEAT_SEQ_C = "0" + Trim(STR(V_HEAT_SEQ_N - 1))
'  End If
'
'  cbo_heat_no = Mid(cbo_heat_no, 1, 3) + V_HEAT_SEQ_C
'  V_HEAT_NO = cbo_heat_no
  If Trim(cbo_heat_no.Text) = "" Or Mid(cbo_heat_no, 4, 5) = "00001" Then
     Exit Sub
  End If
    
  cbo_heat_no = Mid(cbo_heat_no, 1, 3) + Format(Val(Mid(cbo_heat_no, 4, 5)) - 1, "00000")
  V_HEAT_NO = cbo_heat_no

  Call Form_Cls
  cbo_heat_no = V_HEAT_NO
  cbo_re_cd.Text = "1"
  Call Form_Ref
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

    cbo_shift.AddItem "1"
    cbo_shift.AddItem "2"
    cbo_shift.AddItem "3"

    cbo_group_cd.AddItem "A"
    cbo_group_cd.AddItem "B"
    cbo_group_cd.AddItem "C"
    cbo_group_cd.AddItem "D"

    cbo_prc_line.AddItem "1"
    cbo_prc_line.AddItem "2"
    cbo_prc_line.AddItem "3"
    
    cbo_re_cd.AddItem "1"
    cbo_re_cd.AddItem "2"
    cbo_re_cd.AddItem "3"
    
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    cbo_prc_line.Text = "1"
        
    Chk_ss1.VALUE = ssCBChecked
    Chk_ss1.ForeColor = &HFF&
    Frame4.Enabled = False
    
    Call Gf_ComboAdd(M_CN1, cbo_ld_id, "SELECT CD  FROM ZP_CD WHERE CD_MANA_NO='F0004' AND CD LIKE  'S%' ")
'    txt_act_steel_grd.Enabled = False
    Screen.MousePointer = vbDefault
    
    txt_emp_cd = sUserID
    txt_emp_cd.ForeColor = &H80000011
    
        Call Gf_HeatNo_ComboAdd(M_CN1, cbo_heat_no, "FP_LFRSLT", "DEP_STEEL_WGT", Trim(cbo_prc_line.Text))
        If cbo_heat_no.ListCount <> 0 And Trim(cbo_heat_no.Text) = "" Then
           cbo_heat_no.ListIndex = 0
           cbo_re_cd.Text = "1"
        End If
    
    If cbo_heat_no <> "" Then
       Call Form_Ref
    End If
    If cbo_shift.Text = "1" Or cbo_shift.Text = "2" Or cbo_shift.Text = "3" Then
       Exit Sub
    End If
    
    sQuery_Rt = Gf_CodeFind(M_CN1, "SELECT SHIFT  FROM GP_SHIFTSTD WHERE TRANCD = 'BD21' ")
    If sQuery_Rt = "" Then
       cbo_shift.ListIndex = 0
    Else
       cbo_shift.ListIndex = Val(sQuery_Rt) - 1
    End If
    
    sQuery_Rt = Gf_CodeFind(M_CN1, "SELECT GROUP_CD  FROM GP_SHIFTSTD WHERE TRANCD = 'BD21' ")
    
    Select Case sQuery_Rt
        Case "B"
            cbo_group_cd.ListIndex = 1
        Case "C"
            cbo_group_cd.ListIndex = 2
        Case "D"
            cbo_group_cd.ListIndex = 3
        Case Else
            cbo_group_cd.ListIndex = 0
    End Select

        
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
    
    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    Call Gp_Ms_ControlLock(Mc2("pControl"), False)
    pControl(1).SetFocus
    
    'cbo_prc_line.Text = "1"
    
    Chk_ss1.VALUE = ssCBChecked
    Chk_ss1.ForeColor = &HFF&
    Frame4.Enabled = False
    txt_oper = "1"
    
    txt_dir_steel_grd.Enabled = False
    txt_act_steel_grd.Enabled = False
    
    txt_emp_cd = sUserID
    txt_emp_cd.ForeColor = &H80000011

End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
 '   Call Gf_Ms_Copy(Mc2)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
  '  If Gf_Ms_Paste(M_CN1, Mc2) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
End Sub

Public Sub Form_Ref()
Dim Temp_Heat_No, Temp_Treat_No As String

    If cbo_re_cd.Text = "" Then
       MsgBox "再处理必须输入！", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
    If cbo_heat_no.Text = "" Then
       MsgBox "炉号必须输入！", vbCritical, "系统提示信息"
       Exit Sub
    End If
    Temp_Heat_No = Mid(cbo_heat_no.Text, 1, 8)
    Temp_Treat_No = cbo_re_cd.Text
    Call Gp_Ms_Cls(Mc1("rControl"))
    cbo_heat_no.Text = Temp_Heat_No
    cbo_re_cd.Text = Temp_Treat_No
    If Gf_Ms_Refer(M_CN1, Mc2, Nothing, Nothing, False) Or Gf_Ms_Refer(M_CN1, Mc1, Nothing, Nothing, False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        Call Gp_Ms_ControlLock(Mc2("pControl"), True)
        
            txt_dir_steel_grd.Enabled = True
            txt_dir_steel_grd.Locked = True
            txt_dir_steel_grd.ForeColor = &H80000011
            
            txt_act_steel_grd.Enabled = True
            cbo_re_cd.Enabled = True
'    Else
'       MsgBox "无相关记录", vbInformation, "系统提示信息"
    End If
    
    If txt_emp_cd = "" Then
       txt_emp_cd = sUserID
       txt_emp_cd.ForeColor = &H80000011
    End If
            
            
End Sub

Public Sub Form_Pro()
   
    Dim sMesg As String
    
    If cbo_prc_line.Text <> "1" And cbo_prc_line.Text <> "2" And cbo_prc_line.Text <> "3" Then
       MsgBox "   LF stattion ！Please Input  ", vbCritical, "错误提示"
       Exit Sub
    End If
    If cbo_shift = "" Then
       MsgBox "   班次 ！Please Input  ", vbCritical, "错误提示"
       Exit Sub
    End If
    If cbo_group_cd = "" Then
       MsgBox "   班别 ！Please Input  ", vbCritical, "错误提示"
       Exit Sub
    End If
    If txt_emp_cd = "" Then
       MsgBox "   作业人员 ！Please Input  ", vbCritical, "错误提示"
       Exit Sub
    End If
'    If cbo_ld_id.Text = "" Then
'       MsgBox "   钢包号 ！Please Input  ", vbCritical, "错误提示"
'       Exit Sub
'    End If
    
    If Len(Trim(cbo_heat_no.Text)) <> 8 Then
        sMesg = sMesg + " 炉号必须是8位"
        Call Gp_MsgBoxDisplay(sMesg)
       Exit Sub
    Else
        If Gf_Mc_Authority(sAuthority, Mc1) Then
            If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            txt_dir_steel_grd.Enabled = True
            txt_dir_steel_grd.Locked = True
            txt_dir_steel_grd.ForeColor = &H80000011
        End If
    End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    If Not Gf_Ms_Del(M_CN1, Mc2) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
End Sub

'Private Sub Text2_Change()
'  If KeyCode = vbKeyF4 Then
'         DD.sWitch = "MS"
'         DD.sKey = "C0002"
'         DD.rControl.Add Item:=cbo_ref_proc
'
'         DD.nameType = "2"
'
'         Call Gf_Common_DD(M_CN1, KeyCode)
'
'         Exit Sub
'
'       End If
'End Sub


Private Sub txt_act_steel_grd_Change()
    Dim sQuery As String
    
    If Len(Trim(txt_act_steel_grd.Text)) >= 10 Then
      sQuery = "SELECT STEEL_GRD_DETAIL FROM qp_nisco_chmc WHERE STLGRD = '" + txt_act_steel_grd.Text + "' AND NVL(STLGRD_FL,'X') <> 'H' "
      txt_stlgrd_c.Text = Gf_CodeFind(M_CN1, sQuery)
    Else
      txt_stlgrd_c.Text = ""
    End If
End Sub

Private Sub txt_act_steel_grd_DblClick()

    Call txt_act_steel_grd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_act_steel_grd_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sQuery As String
    
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_act_steel_grd
        DD.rControl.Add Item:=txt_stlgrd_c
        
        Call Pf_Common_DD(M_CN1, KeyCode)
    Else
        If Len(Trim(txt_act_steel_grd.Text)) >= 10 Then
          sQuery = "SELECT STEEL_GRD_DETAIL FROM qp_nisco_chmc WHERE STLGRD = '" + txt_act_steel_grd.Text + "' AND NVL(STLGRD_FL,'X') <> 'H' "
          txt_stlgrd_c.Text = Gf_CodeFind(M_CN1, sQuery)
        Else
          txt_stlgrd_c.Text = ""
        End If
    End If
End Sub

Private Sub txt_end_date_DblClick()
    txt_end_date.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_ld_arrv_date_DblClick()
    txt_ld_arrv_date.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_ld_dep_date_DblClick()
    txt_ld_dep_date.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_OCCR_DATE_DblClick()
    txt_OCCR_DATE.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_start_date_DblClick()
    txt_start_date.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_wire_cd1_DblClick()

    Call txt_wire_cd1_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_wire_cd1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "F0008"
        DD.rControl.Add Item:=txt_wire_cd1
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub txt_wire_cd2_DblClick()

    Call txt_wire_cd2_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_wire_cd2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "F0008"
        DD.rControl.Add Item:=txt_wire_cd2
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub txt_wire_cd3_DblClick()

    Call txt_wire_cd3_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_wire_cd3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "F0008"
        DD.rControl.Add Item:=txt_wire_cd3
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub txt_wire_cd4_DblClick()

    Call txt_wire_cd4_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_wire_cd4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "F0008"
        DD.rControl.Add Item:=txt_wire_cd4
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
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
    
    DD.DataDicType = "S"        'Common Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    DD.sQuery = "SELECT STLGRD ""钢种代码"", STEEL_GRD_DETAIL ""钢种名称"" FROM qp_nisco_chmc "
    
    If DD.rControl.Count > 1 Then
        DD.sWhere = " WHERE NVL(STLGRD_FL,'X') <> 'H'  "
        DD.sWhere = DD.sWhere + "   AND STLGRD           like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + "   AND STEEL_GRD_DETAIL like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        
    End If
    
    Call Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False)
    
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function
