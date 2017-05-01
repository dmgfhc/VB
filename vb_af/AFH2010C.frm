VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AFH2010C 
   Caption         =   "连铸实绩修改及查询界面_AFH2010C"
   ClientHeight    =   9225
   ClientLeft      =   375
   ClientTop       =   2175
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame Frame2 
      Height          =   1095
      Left            =   150
      TabIndex        =   6
      Top             =   1080
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   1931
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
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
         Height          =   315
         Left            =   11025
         TabIndex        =   25
         Top             =   180
         Width           =   1365
      End
      Begin VB.TextBox txt_CAST_HEAT_SEQ 
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
         Left            =   8790
         TabIndex        =   24
         Top             =   180
         Width           =   735
      End
      Begin VB.TextBox txt_CAST_HEAT_CNT 
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
         Left            =   6120
         TabIndex        =   23
         Top             =   180
         Width           =   735
      End
      Begin VB.TextBox txt_CAST_MANA_NO 
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
         Left            =   3780
         TabIndex        =   22
         Top             =   180
         Width           =   780
      End
      Begin VB.ComboBox cbo_LD_ID 
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
         TabIndex        =   21
         Top             =   180
         Width           =   735
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
         Left            =   12390
         TabIndex        =   20
         Top             =   180
         Width           =   2325
      End
      Begin VB.TextBox txt_stlgrd_r 
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
         Left            =   11025
         TabIndex        =   19
         Top             =   600
         Width           =   1365
      End
      Begin VB.TextBox txt_stlgrd_c 
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
         Left            =   12390
         TabIndex        =   18
         Top             =   600
         Width           =   2325
      End
      Begin VB.TextBox txt_rl_cast_no 
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
         Left            =   3780
         TabIndex        =   17
         Tag             =   "实际连浇号"
         Top             =   600
         Width           =   660
      End
      Begin VB.TextBox txt_rl_heat_seq 
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
         Left            =   6120
         TabIndex        =   16
         Tag             =   "实际连浇序号"
         Top             =   600
         Width           =   465
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   2550
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
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
         Left            =   4890
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "连浇炉数"
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
         Left            =   7260
         Top             =   180
         Width           =   1500
         _ExtentX        =   2646
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
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   9795
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
      End
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   210
         Top             =   180
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
      Begin InDate.ULabel ULabel43 
         Height          =   315
         Left            =   2550
         Top             =   600
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "实际连浇号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel44 
         Height          =   315
         Left            =   4890
         Top             =   600
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "实际连浇序号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel41 
         Height          =   315
         Left            =   9795
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
      End
   End
   Begin Threed.SSFrame Frame1 
      Height          =   915
      Left            =   150
      TabIndex        =   5
      Top             =   120
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   1614
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.ComboBox cbo_HEAT_NO 
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
         TabIndex        =   15
         Tag             =   "炉号"
         Top             =   300
         Width           =   1425
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
         ItemData        =   "AFH2010C.frx":0000
         Left            =   5400
         List            =   "AFH2010C.frx":0002
         TabIndex        =   14
         Tag             =   "连铸机号"
         Top             =   300
         Width           =   660
      End
      Begin VB.ComboBox cbo_SHIFT 
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
         ItemData        =   "AFH2010C.frx":0004
         Left            =   7830
         List            =   "AFH2010C.frx":0006
         TabIndex        =   13
         Tag             =   "班次"
         Top             =   300
         Width           =   735
      End
      Begin VB.ComboBox cbo_GROUP_CD 
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
         ItemData        =   "AFH2010C.frx":0008
         Left            =   10275
         List            =   "AFH2010C.frx":000A
         TabIndex        =   12
         Tag             =   "班别"
         Top             =   300
         Width           =   735
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
         Left            =   12735
         Locked          =   -1  'True
         TabIndex        =   11
         Tag             =   "作业人员"
         Top             =   300
         Width           =   1095
      End
      Begin VB.CommandButton cmd_up 
         Caption         =   "▲"
         Height          =   255
         Left            =   2955
         TabIndex        =   10
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmd_down 
         Caption         =   "▼"
         Height          =   255
         Left            =   2955
         TabIndex        =   9
         Top             =   450
         Width           =   300
      End
      Begin InDate.ULabel ULabel1 
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   4170
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
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
         Left            =   6585
         Top             =   300
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   9030
         Top             =   300
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
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   11475
         Top             =   300
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
   End
   Begin VB.TextBox txt_oper 
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
      Left            =   2580
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin Threed.SSCheck Chk_ss2 
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Top             =   3915
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "2.实绩"
   End
   Begin Threed.SSCheck Chk_ss1 
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   2340
      Width           =   1065
      _ExtentX        =   1879
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
      Caption         =   "1.进程"
      Value           =   1
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
      Left            =   5970
      TabIndex        =   2
      Text            =   "B1"
      Top             =   990
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
      Left            =   6960
      TabIndex        =   3
      Text            =   "BF"
      Top             =   990
      Visible         =   0   'False
      Width           =   735
   End
   Begin Threed.SSFrame Frame3 
      Height          =   1065
      Left            =   150
      TabIndex        =   7
      Top             =   2730
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   1879
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin CSTextLibCtl.sitxEdit txt_OCCR_DATE1 
         Height          =   315
         Left            =   2100
         TabIndex        =   26
         Top             =   150
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
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
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   270
         Top             =   150
         Width           =   1800
         _ExtentX        =   3175
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
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   10455
         Top             =   585
         Width           =   1800
         _ExtentX        =   3175
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
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   10455
         Top             =   150
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "钢包关闭时间"
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
         Left            =   270
         Top             =   585
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "钢包开浇时间"
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
         Left            =   5265
         Top             =   150
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "浇铸开始时间"
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
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   5265
         Top             =   585
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "中包浇铸结束时间"
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
      Begin CSTextLibCtl.sitxEdit txt_OCCR_DATE2 
         Height          =   315
         Left            =   2100
         TabIndex        =   27
         Top             =   585
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
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
      Begin CSTextLibCtl.sitxEdit txt_OCCR_DATE3 
         Height          =   315
         Left            =   7110
         TabIndex        =   28
         Top             =   150
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
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
      Begin CSTextLibCtl.sitxEdit txt_OCCR_DATE4 
         Height          =   315
         Left            =   7110
         TabIndex        =   29
         Top             =   585
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
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
      Begin CSTextLibCtl.sitxEdit txt_OCCR_DATE5 
         Height          =   315
         Left            =   12300
         TabIndex        =   30
         Top             =   150
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
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
      Begin CSTextLibCtl.sitxEdit txt_OCCR_DATE6 
         Height          =   315
         Left            =   12300
         TabIndex        =   31
         Top             =   585
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
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
   End
   Begin Threed.SSFrame Frame4 
      Height          =   4815
      Left            =   150
      TabIndex        =   8
      Top             =   4260
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   8493
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.ComboBox cbo_TD_PWD_TYPE 
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
         Left            =   13665
         TabIndex        =   47
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txt_ELEMENT_VAL_S 
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
         Height          =   315
         Left            =   7050
         TabIndex        =   46
         Top             =   4350
         Width           =   735
      End
      Begin VB.TextBox txt_ELEMENT_VAL_P 
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
         Height          =   315
         Left            =   5850
         TabIndex        =   45
         Top             =   4350
         Width           =   735
      End
      Begin VB.TextBox txt_ELEMENT_VAL_Si 
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
         Height          =   315
         Left            =   4680
         TabIndex        =   44
         Top             =   4350
         Width           =   735
      End
      Begin VB.TextBox txt_ELEMENT_VAL_Mn 
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
         Height          =   315
         Left            =   3480
         TabIndex        =   43
         Top             =   4350
         Width           =   735
      End
      Begin VB.TextBox txt_ELEMENT_VAL_C 
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
         Height          =   315
         Left            =   2280
         TabIndex        =   42
         Top             =   4350
         Width           =   735
      End
      Begin VB.TextBox txt_ELEMENT_VAL_O 
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
         Height          =   315
         Left            =   12150
         TabIndex        =   41
         Top             =   4350
         Width           =   735
      End
      Begin VB.TextBox txt_ELEMENT_VAL_N 
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
         Height          =   315
         Left            =   10860
         TabIndex        =   40
         Top             =   4350
         Width           =   735
      End
      Begin VB.TextBox txt_ELEMENT_VAL_H 
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
         Height          =   315
         Left            =   9570
         TabIndex        =   39
         Top             =   4350
         Width           =   735
      End
      Begin VB.TextBox txt_ELEMENT_VAL_Al 
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
         Height          =   315
         Left            =   8250
         TabIndex        =   38
         Top             =   4350
         Width           =   735
      End
      Begin VB.ComboBox cbo_MD_PWD_TYPE 
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
         ItemData        =   "AFH2010C.frx":000C
         Left            =   13665
         List            =   "AFH2010C.frx":000E
         TabIndex        =   37
         Top             =   1650
         Width           =   855
      End
      Begin VB.ComboBox cbo_td_no 
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
         ItemData        =   "AFH2010C.frx":0010
         Left            =   2340
         List            =   "AFH2010C.frx":0032
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cbo_td_car 
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
         ItemData        =   "AFH2010C.frx":005E
         Left            =   2340
         List            =   "AFH2010C.frx":0080
         TabIndex        =   33
         Top             =   660
         Width           =   1215
      End
      Begin VB.ComboBox cbo_md_no 
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
         ItemData        =   "AFH2010C.frx":00AC
         Left            =   5905
         List            =   "AFH2010C.frx":00CB
         TabIndex        =   32
         Top             =   660
         Width           =   1155
      End
      Begin CSTextLibCtl.sidbEdit txt_LD_ARRV_WGT 
         Height          =   315
         Left            =   2340
         TabIndex        =   35
         Top             =   1080
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
      Begin CSTextLibCtl.sidbEdit txt_SLAB_CNT 
         Height          =   315
         Left            =   2340
         TabIndex        =   36
         Top             =   3930
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   210
         Top             =   3930
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "板坯总数"
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
         Top             =   1080
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "钢包到达重量(ton)"
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
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   3780
         Top             =   1080
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "钢包剩余钢水量(ton)"
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
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   7365
         Top             =   1080
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "中包 剩余钢水重量"
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
         Left            =   210
         Top             =   2025
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "中包钢水平均温度"
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
         Left            =   3780
         Top             =   3930
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "板坯总长度(mm)"
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
         Left            =   210
         Top             =   2520
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "浇铸平均速度(m/min)"
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
         Left            =   3780
         Top             =   2520
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器浇铸时间"
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
         Left            =   7365
         Top             =   3930
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "切割结束时间"
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
         Left            =   11550
         Top             =   240
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器平均液位(mm)"
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
         Left            =   11550
         Top             =   660
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器最低液位(mm)"
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
         Left            =   11550
         Top             =   1080
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器最高液位(mm)"
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
         Left            =   11550
         Top             =   1650
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器保护渣代码"
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
      Begin InDate.ULabel ULabel35 
         Height          =   315
         Left            =   11550
         Top             =   2025
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器保护渣用量(kg)"
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
         Left            =   11550
         Top             =   2520
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "中包覆盖剂代码"
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
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   11550
         Top             =   3930
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "中包覆盖剂用量(kg)"
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
         Left            =   210
         Top             =   4350
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         Caption         =   "中包钢水成分(%)"
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
         Left            =   1860
         Top             =   4350
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         Caption         =   "C"
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
         Left            =   3060
         Top             =   4350
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         Caption         =   "Mn"
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
         Left            =   4260
         Top             =   4350
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         Caption         =   "Si"
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
      Begin InDate.ULabel ULabel50 
         Height          =   315
         Left            =   5460
         Top             =   4350
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         Caption         =   "P"
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
         Left            =   6630
         Top             =   4350
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         Caption         =   "S"
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
         Left            =   7830
         Top             =   4350
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         Caption         =   "Al"
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
         Left            =   11640
         Top             =   4350
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Caption         =   "[O]"
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
         Left            =   9030
         Top             =   4350
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Caption         =   "[H]"
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
         Left            =   10335
         Top             =   4350
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Caption         =   "[N]"
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
      Begin CSTextLibCtl.sitxEdit txt_MD_IN_DATE 
         Height          =   315
         Left            =   5905
         TabIndex        =   48
         Top             =   2520
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
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
      Begin CSTextLibCtl.sitxEdit txt_LAST_CUT_DATE 
         Height          =   315
         Left            =   9480
         TabIndex        =   49
         Top             =   3930
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
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
      Begin CSTextLibCtl.sidbEdit txt_LD_DEP_WGT 
         Height          =   315
         Left            =   5905
         TabIndex        =   50
         Top             =   1080
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
      Begin CSTextLibCtl.sidbEdit txt_TD_REM_WGT 
         Height          =   315
         Left            =   9480
         TabIndex        =   51
         Top             =   1080
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
      Begin CSTextLibCtl.sidbEdit txt_TD_AVE_TEMP 
         Height          =   315
         Left            =   2340
         TabIndex        =   52
         Top             =   2025
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_TOT_LEN 
         Height          =   315
         Left            =   5905
         TabIndex        =   53
         Top             =   3930
         Width           =   1335
         _Version        =   262145
         _ExtentX        =   2355
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_AVE_CAST_SPEED 
         Height          =   315
         Left            =   2340
         TabIndex        =   54
         Top             =   2520
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_AVE_MOLD_LEVEL 
         Height          =   315
         Left            =   13665
         TabIndex        =   55
         Top             =   240
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_MIN_MOLD_LEVEL 
         Height          =   315
         Left            =   13665
         TabIndex        =   56
         Top             =   660
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   125
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
         DataProperty    =   2
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_MAX_MOLD_LEVEL 
         Height          =   315
         Left            =   13665
         TabIndex        =   57
         Top             =   1080
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_MD_PWD_WGT 
         Height          =   315
         Left            =   13665
         TabIndex        =   58
         Top             =   2025
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
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
         NumIntDigits    =   5
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_TD_PWD_WGT 
         Height          =   315
         Left            =   13665
         TabIndex        =   59
         Top             =   3930
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
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
         NumIntDigits    =   5
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   210
         Top             =   240
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "中包包号"
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
         Left            =   210
         Top             =   660
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "中包号车"
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
         Left            =   3780
         Top             =   240
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "中包包龄"
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
      Begin CSTextLibCtl.sidbEdit sde_td_life 
         Height          =   315
         Left            =   5905
         TabIndex        =   60
         Top             =   240
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel28 
         Height          =   315
         Left            =   210
         Top             =   1650
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "中包 一次温度(℃)"
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
      Begin CSTextLibCtl.sidbEdit sde_td_temp1 
         Height          =   315
         Left            =   2340
         TabIndex        =   61
         Top             =   1650
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel29 
         Height          =   315
         Left            =   3780
         Top             =   1650
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "中包 二次温度(℃)"
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
      Begin CSTextLibCtl.sidbEdit sde_td_temp2 
         Height          =   315
         Left            =   5905
         TabIndex        =   62
         Top             =   1650
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel38 
         Height          =   315
         Left            =   7365
         Top             =   1650
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "中包 三次温度(℃)"
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
      Begin CSTextLibCtl.sidbEdit sde_td_temp3 
         Height          =   315
         Left            =   9480
         TabIndex        =   63
         Top             =   1650
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel39 
         Height          =   315
         Left            =   3780
         Top             =   660
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器号"
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
      Begin CSTextLibCtl.sitxEdit txt_upd_date 
         Height          =   315
         Left            =   9480
         TabIndex        =   64
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
      Begin InDate.ULabel ULabel45 
         Height          =   315
         Left            =   7365
         Top             =   240
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "上次修改时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel40 
         Height          =   315
         Left            =   3780
         Top             =   2025
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "液相线温度"
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
      Begin CSTextLibCtl.sidbEdit txt_stl_lquid_temp 
         Height          =   315
         Left            =   5905
         TabIndex        =   65
         Top             =   2025
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel42 
         Height          =   315
         Left            =   7365
         Top             =   2025
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "平均过热度"
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
      Begin CSTextLibCtl.sidbEdit txt_stl_avg_temp 
         Height          =   315
         Left            =   9480
         TabIndex        =   66
         Top             =   2025
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel56 
         Height          =   315
         Left            =   210
         Top             =   2970
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器左侧窄面流量"
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
      Begin InDate.ULabel ULabel57 
         Height          =   315
         Left            =   210
         Top             =   3390
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器左侧窄面温差"
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
         Left            =   7365
         Top             =   2970
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器活动侧宽面流量"
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
         Left            =   11550
         Top             =   2970
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器固定侧宽面流量"
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
         Left            =   3780
         Top             =   2970
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器右侧窄面流量"
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
         Left            =   3780
         Top             =   3390
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器右侧窄面温差"
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
         Left            =   7365
         Top             =   3390
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器活动侧宽面温差"
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
      Begin InDate.ULabel ULabel63 
         Height          =   315
         Left            =   11550
         Top             =   3390
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         Caption         =   "结晶器固定侧宽面温差"
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
      Begin CSTextLibCtl.sidbEdit sdb_modl_narrow_flow_l 
         Height          =   315
         Left            =   2340
         TabIndex        =   67
         Top             =   2970
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_modl_narrow_dif_temp_l 
         Height          =   315
         Left            =   2340
         TabIndex        =   68
         Top             =   3390
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_mold_wide_flow_l 
         Height          =   315
         Left            =   9480
         TabIndex        =   69
         Top             =   2970
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_mold_wide_flow_f 
         Height          =   315
         Left            =   13665
         TabIndex        =   70
         Top             =   2970
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_modl_narrow_flow_r 
         Height          =   315
         Left            =   5905
         TabIndex        =   71
         Top             =   2970
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_modl_narrow_dif_temp_r 
         Height          =   315
         Left            =   5905
         TabIndex        =   72
         Top             =   3390
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_mold_wide_dif_temp_l 
         Height          =   315
         Left            =   9480
         TabIndex        =   73
         Top             =   3390
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_mold_wide_dif_temp_f 
         Height          =   315
         Left            =   13665
         TabIndex        =   74
         Top             =   3390
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
   End
End
Attribute VB_Name = "AFH2010C"
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
'-- Program Name      CAST
'-- Program ID        AFH2010C
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2003.7.25
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

Public sDateTime As String          'Active Form Authority Setting
Public sYear As String              'Active Form Authority Setting
Public sMonth As String             'Active Form Authority Setting
Public sDay As String               'Active Form Authority Setting
Public sOur As String               'Active Form Authority Setting
Public sMin As String               'Active Form Authority Setting
Public sSec As String               'Active Form Authority Setting
Public sQuery As String
Public QueryYN As Boolean
Public sQuery_Rt As String

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pControl1 As New Collection     'Master Primary Key Collection
Dim nControl1 As New Collection     'Master Necessary Collection
Dim mControl1 As New Collection     'Master Maxlength check Collection
Dim iControl1 As New Collection     'Master Insert Collection
Dim rControl1 As New Collection     'Master Refer Collection
Dim cControl1 As New Collection     'Master Copy Collection
Dim aControl1 As New Collection     'Master -> Spread Collection
Dim lControl1 As New Collection     'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Master"
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
                      Call Gp_Ms_Collection(txt_oper, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   Call Gp_Ms_Collection(cbo_heat_no, "p", "n", "m", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(cbo_PRC_LINE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                       Call Gp_Ms_Collection(txt_plt, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                       Call Gp_Ms_Collection(txt_PRC, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     Call Gp_Ms_Collection(cbo_shift, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(cbo_group_cd, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_emp_cd, " ", "n", " ", "i", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     Call Gp_Ms_Collection(cbo_ld_id, " ", " ", " ", "i", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_CAST_MANA_NO, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_CAST_HEAT_CNT, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_CAST_HEAT_SEQ, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_rl_cast_no, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_rl_heat_seq, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(txt_stlgrd, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_stlgrd_n, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_stlgrd_r, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_stlgrd_c, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
                Call Gp_Ms_Collection(txt_OCCR_DATE1, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_OCCR_DATE2, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_OCCR_DATE3, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_OCCR_DATE4, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_OCCR_DATE5, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_OCCR_DATE6, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

                  Call Gp_Ms_Collection(txt_slab_cnt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_LD_ARRV_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_LD_DEP_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_TD_REM_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_TD_AVE_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_stl_lquid_temp, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_stl_avg_temp, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   Call Gp_Ms_Collection(txt_TOT_LEN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_AVE_CAST_SPEED, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_MD_IN_DATE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_LAST_CUT_DATE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_AVE_MOLD_LEVEL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_MIN_MOLD_LEVEL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_MAX_MOLD_LEVEL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(cbo_MD_PWD_TYPE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_MD_PWD_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(cbo_TD_PWD_TYPE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_TD_PWD_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     Call Gp_Ms_Collection(cbo_td_no, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   Call Gp_Ms_Collection(sde_td_life, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                    Call Gp_Ms_Collection(cbo_td_car, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(sde_td_temp1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(sde_td_temp2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(sde_td_temp3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     Call Gp_Ms_Collection(cbo_md_no, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                
        Call Gp_Ms_Collection(sdb_modl_narrow_flow_l, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_modl_narrow_dif_temp_l, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_modl_narrow_flow_r, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_modl_narrow_dif_temp_r, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_mold_wide_flow_l, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_mold_wide_dif_temp_f, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_mold_wide_flow_f, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_mold_wide_dif_temp_f, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

             Call Gp_Ms_Collection(txt_ELEMENT_VAL_C, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ELEMENT_VAL_Mn, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ELEMENT_VAL_Si, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ELEMENT_VAL_P, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ELEMENT_VAL_S, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ELEMENT_VAL_Al, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ELEMENT_VAL_H, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ELEMENT_VAL_N, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ELEMENT_VAL_O, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_upd_date, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
    'MASTER Collection
    Mc1.Add Item:="AFH2010C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="AFH2010C.P_REFER", Key:="P-R"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(cbo_heat_no, "p", "n", "m", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(cbo_PRC_LINE, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(cbo_shift, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(cbo_group_cd, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(txt_emp_cd, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(cbo_ld_id, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_OCCR_DATE1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_OCCR_DATE2, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_OCCR_DATE3, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_OCCR_DATE4, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_OCCR_DATE5, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_OCCR_DATE6, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      
    Mc2.Add Item:="AFH2010C.P_REFER1", Key:="P-R"
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

Private Sub cbo_GROUP_CD_Change()

    If cbo_group_cd.Text <> "" Then
        If cbo_group_cd.Text <> "A" And cbo_group_cd.Text <> "B" And cbo_group_cd.Text <> "C" And cbo_group_cd.Text <> "D" Then
            Call Gp_MsgBoxDisplay("班别输入不正确！", "", "错误提示")
        End If
    End If
    
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
        cbo_ld_id.Text = Gf_FloatFind(M_CN1, "SELECT LADLE_NO FROM FP_CHARGE WHERE HEAT_NO = '" + cbo_heat_no.Text + "'")
        If cbo_ld_id.Text = "0" Then
            cbo_ld_id.Text = ""
        End If
   Else
        cbo_ld_id.Text = ""
   End If
   
End Sub

Private Sub CBO_MD_PWD_TYPE_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "F0006"
        DD.rControl.Add Item:=cbo_MD_PWD_TYPE
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If
    
End Sub

Private Sub cbo_prc_line_Change()

    Dim str1 As String

    'If cbo_HEAT_NO.Enabled = True Then
    '    Call Gf_HeatNo_ComboAdd(M_CN1, cbo_HEAT_NO, "FP_CASTRSLT", "LD_ARRV_WGT", Trim(cbo_prc_line.Text))
    '    If cbo_HEAT_NO.ListCount <> 0 And Trim(cbo_HEAT_NO.Text) = "" Then
    '       cbo_HEAT_NO.ListIndex = 0
    '    End If
    'End If

End Sub

Private Sub cbo_prc_line_Click()
    
'    Call cbo_prc_line_Change
'
'    If Trim(cbo_prc_line.Text) = "2" Then
'        str1 = cbo_heat_no.Text
'        cbo_heat_no.Clear
'        cbo_heat_no.Text = str1
'        ' Call Gf_ComboAdd(M_CN1, cbo_HEAT_NO, "SELECT B.GOODS_ID  FROM FP_TRACKIDX A, FP_TRACKDATA B WHERE  A.LAST_SEQ >= B.SEQ_NO  AND B.FACT_CD = '" & txt_plt.Text & "'AND B.PRC = '" & txt_prc.Text & "'AND B.PRC_LINE='2' AND A.FACT_CD=B.FACT_CD  AND A.PRC=B.PRC AND A.PRC_LINE=B.PRC_LINE ")
'    ElseIf Trim(cbo_prc_line.Text) = "1" Then
'        str1 = cbo_heat_no.Text
'        cbo_heat_no.Clear
'        cbo_heat_no.Text = str1
'        ' Call Gf_ComboAdd(M_CN1, cbo_HEAT_NO, "SELECT B.GOODS_ID  FROM FP_TRACKIDX A, FP_TRACKDATA B WHERE  A.LAST_SEQ >= B.SEQ_NO  AND B.FACT_CD = '" & txt_plt.Text & "'AND B.PRC = '" & txt_prc.Text & "'AND B.PRC_LINE='1' AND A.FACT_CD=B.FACT_CD  AND A.PRC=B.PRC AND A.PRC_LINE=B.PRC_LINE ")
'    ElseIf Trim(cbo_prc_line.Text) = "" Then
'        str1 = cbo_heat_no.Text
'        cbo_heat_no.Clear
'        cbo_heat_no.Text = str1
'    End If

End Sub

Private Sub cbo_SHIFT_Change()

    If cbo_shift.Text <> "" Then
        If cbo_shift.Text <> "1" And cbo_shift.Text <> "2" And cbo_shift.Text <> "3" Then
            Call Gp_MsgBoxDisplay("班次输入不正确！", "", "错误提示")
        End If
    End If
    
End Sub

Private Sub cbo_TD_PWD_TYPE_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "F0005"
        DD.rControl.Add Item:=cbo_TD_PWD_TYPE
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If
    
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
        Chk_ss2.ForeColor = &H0&
        Chk_ss2.Value = ssCBUnchecked
        Frame3.Enabled = True
        Frame4.Enabled = False
        Frame3.ShadowStyle = ssRaisedShadow
        Frame4.ShadowStyle = ssInsetShadow
        txt_oper.Text = "1"
        txt_OCCR_DATE1.SetFocus
    Else
        Chk_ss1.Value = ssCBUnchecked
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
        Chk_ss1.ForeColor = &H0&
        Chk_ss1.Value = ssCBUnchecked
        Frame3.Enabled = False
        Frame4.Enabled = True
        Frame3.ShadowStyle = ssInsetShadow
        Frame4.ShadowStyle = ssRaisedShadow
        txt_oper.Text = "2"
        txt_slab_cnt.SetFocus
    Else
        Chk_ss2.Value = ssCBUnchecked
    End If
    
End Sub

Private Sub cmd_up_Click()

    Dim V_HEAT_NO As String
    
    If Trim(cbo_heat_no.Text) = "" Then Exit Sub
    
    V_HEAT_NO = Format(cbo_heat_no + 1, "00000000")
    Call Form_Cls
    
    cbo_heat_no = V_HEAT_NO
    Call Form_Ref
    
End Sub

Private Sub cmd_down_Click()

    Dim V_HEAT_NO As String
    
    If Trim(cbo_heat_no.Text) = "" Then Exit Sub
    
    V_HEAT_NO = Format(cbo_heat_no - 1, "00000000")
    Call Form_Cls
    
    cbo_heat_no = V_HEAT_NO
    Call Form_Ref
    
End Sub

Private Sub Form_Activate()

   Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
   
End Sub

Private Sub Form_Load()

    cbo_shift.AddItem "1"
    cbo_shift.AddItem "2"
    cbo_shift.AddItem "3"

    cbo_group_cd.AddItem "A"
    cbo_group_cd.AddItem "B"
    cbo_group_cd.AddItem "C"
    cbo_group_cd.AddItem "D"
    
    cbo_PRC_LINE.AddItem "1"
    cbo_PRC_LINE.AddItem "2"
    cbo_PRC_LINE.AddItem "3"
    
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    cbo_PRC_LINE.Text = "1"
    
    Call Gf_ComboAdd(M_CN1, cbo_ld_id, "SELECT CD  FROM ZP_CD WHERE CD_MANA_NO='F0004' AND CD LIKE  'S%' ")
    Call Gf_HeatNo_ComboAdd(M_CN1, cbo_heat_no, "FP_CASTRSLT", "LD_ARRV_WGT", Trim(cbo_PRC_LINE.Text))
    
    If cbo_heat_no.ListCount <> 0 Then
        cbo_heat_no.ListIndex = 0
    End If
        
    txt_emp_cd = sUserID
    txt_emp_cd.ForeColor = &H80000011
    
    If cbo_heat_no <> "" Then Call Form_Ref
    
    If cbo_shift.Text = "1" Or cbo_shift.Text = "2" Or cbo_shift.Text = "3" Then Exit Sub
    
    sQuery_Rt = Gf_CodeFind(M_CN1, "SELECT SHIFT  FROM GP_SHIFTSTD WHERE TRANCD = 'BF21' ")
    If sQuery_Rt = "" Then
        cbo_shift.ListIndex = 0
    Else
        cbo_shift.ListIndex = Val(sQuery_Rt) - 1
    End If
    
    sQuery_Rt = Gf_CodeFind(M_CN1, "SELECT GROUP_CD  FROM GP_SHIFTSTD WHERE TRANCD = 'BF21' ")
    
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
    
    If Chk_ss1.Value = -1 Then
        txt_oper.Text = "1"
    ElseIf Chk_ss2.Value = -1 Then
        txt_oper.Text = "2"
    End If
    
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
    
    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
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
    Chk_ss1.Value = ssCBChecked
    Chk_ss1.ForeColor = &HFF&
    Frame4.Enabled = False
    
    txt_emp_cd = sUserID
    txt_emp_cd.ForeColor = &H80000011
    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
End Sub

Public Sub Form_Ref()
    
    cbo_heat_no.Text = Mid(cbo_heat_no.Text, 1, 8)
  
    If Trim(cbo_heat_no.Text) = "" Then
        Call Gp_MsgBoxDisplay("炉号必须输入", "", "错误提示")
    ElseIf Len(Trim(cbo_heat_no.Text)) <> 8 Then
        Call Gp_MsgBoxDisplay("炉号长度应为8位", "", "错误提示")
    Else
    
        QueryYN = True
            
        If Gf_Ms_Refer(M_CN1, Mc1, Nothing, Nothing, False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("pControl"), True)
            txt_stlgrd.Enabled = True
            txt_stlgrd.Locked = True
            txt_stlgrd.ForeColor = &H80000011
            QueryYN = False
        End If
        
        If Gf_Ms_Refer(M_CN1, Mc2, Nothing, Nothing, False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc2("pControl"), True)
            QueryYN = False
        End If
        
    End If
    
    If txt_emp_cd = "" Then
        txt_emp_cd = sUserID
        txt_emp_cd.ForeColor = &H80000011
    End If
    
End Sub

Public Sub Form_Pro()
    
    Dim msg As String
    
    cbo_heat_no.Text = Mid(cbo_heat_no.Text, 1, 8)
    
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            msg = MDIMain.StatusBar1.Panels(1)
            txt_stlgrd.Enabled = True
            txt_stlgrd.Locked = True
            txt_stlgrd.ForeColor = &H80000011

            Call Gf_Ms_Refer(M_CN1, Mc2, Nothing, Nothing, False)
            MDIMain.StatusBar1.Panels(1) = msg
        End If
    End If
     
End Sub

Public Sub Form_Del()
    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
End Sub

Private Sub txt_LAST_CUT_DATE_DblClick()
    txt_LAST_CUT_DATE.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_MD_IN_DATE_DblClick()
    txt_MD_IN_DATE.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_OCCR_DATE1_DblClick()
    txt_OCCR_DATE1.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_OCCR_DATE2_DblClick()
    txt_OCCR_DATE2.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_OCCR_DATE3_DblClick()
    txt_OCCR_DATE3.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_OCCR_DATE4_DblClick()
    txt_OCCR_DATE4.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_OCCR_DATE5_DblClick()
    txt_OCCR_DATE5.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_OCCR_DATE6_DblClick()
    txt_OCCR_DATE6.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_STLGRD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim sQuery As String
    
    sQuery = "SELECT STEEL_GRD_DETAIL FROM QP_STLGRD_INF WHERE STLGRD = '" + txt_stlgrd.Text + "'"
    txt_stlgrd.ToolTipText = Gf_FloatFind(M_CN1, sQuery)
    
End Sub

