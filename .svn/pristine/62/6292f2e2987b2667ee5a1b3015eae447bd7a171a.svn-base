VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AFF2020C 
   Caption         =   "RH实绩修改及查询界面_AFF2020C"
   ClientHeight    =   9225
   ClientLeft      =   300
   ClientTop       =   2145
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame Frame2 
      Height          =   1095
      Left            =   210
      TabIndex        =   7
      Top             =   1080
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   1931
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
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
         TabIndex        =   25
         Top             =   180
         Width           =   1290
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
         TabIndex        =   24
         Top             =   600
         Width           =   1290
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
         ItemData        =   "AFF2020C.frx":0000
         Left            =   10035
         List            =   "AFF2020C.frx":0002
         TabIndex        =   23
         Tag             =   "班别"
         Top             =   180
         Width           =   630
      End
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
         ItemData        =   "AFF2020C.frx":0004
         Left            =   7815
         List            =   "AFF2020C.frx":0006
         TabIndex        =   22
         Tag             =   "班次"
         Top             =   180
         Width           =   630
      End
      Begin VB.TextBox cbo_ref_proc 
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
         TabIndex        =   21
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
         TabIndex        =   20
         Tag             =   "作业人员"
         Top             =   180
         Width           =   885
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
         TabIndex        =   19
         Top             =   600
         Width           =   1440
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
         TabIndex        =   18
         Top             =   180
         Width           =   3090
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
         TabIndex        =   17
         Top             =   600
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
      Begin InDate.ULabel ULabel54 
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
      Begin InDate.ULabel ULabel55 
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
      Begin InDate.ULabel ULabel56 
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
   Begin Threed.SSFrame Frame1 
      Height          =   915
      Left            =   210
      TabIndex        =   6
      Top             =   120
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   1614
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
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
         Left            =   12705
         TabIndex        =   16
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
         ItemData        =   "AFF2020C.frx":0008
         Left            =   5400
         List            =   "AFF2020C.frx":000A
         TabIndex        =   15
         Tag             =   "再处理"
         Top             =   300
         Width           =   675
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
         ItemData        =   "AFF2020C.frx":000C
         Left            =   7830
         List            =   "AFF2020C.frx":000E
         TabIndex        =   14
         Tag             =   "VD"
         Top             =   300
         Width           =   615
      End
      Begin VB.CommandButton cbo_Low 
         Caption         =   "▼"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   435
         Width           =   375
      End
      Begin VB.CommandButton cbo_Up 
         Caption         =   "▲"
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   150
         Width           =   375
      End
      Begin VB.ComboBox cbo_HEAT_NO 
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
         Left            =   1455
         TabIndex        =   11
         Tag             =   "炉号"
         Top             =   300
         Width           =   1425
      End
      Begin VB.ComboBox cbo_RH_ID 
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
         ItemData        =   "AFF2020C.frx":0010
         Left            =   10255
         List            =   "AFF2020C.frx":0012
         TabIndex        =   10
         Tag             =   "VD"
         Top             =   300
         Width           =   615
      End
      Begin InDate.ULabel ULabel82 
         Height          =   315
         Left            =   11460
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
      Begin InDate.ULabel ULabel13 
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
      Begin InDate.ULabel ULabel53 
         Height          =   315
         Left            =   6585
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "RH机号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Left            =   9015
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "RH坑号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Left            =   10020
      TabIndex        =   4
      Text            =   "1"
      Top             =   915
      Visible         =   0   'False
      Width           =   255
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
      Left            =   11775
      TabIndex        =   3
      Text            =   "BH"
      Top             =   900
      Visible         =   0   'False
      Width           =   1185
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
      Left            =   10335
      TabIndex        =   2
      Text            =   "B1"
      Top             =   900
      Visible         =   0   'False
      Width           =   1275
   End
   Begin Threed.SSCheck Chk_ss1 
      Height          =   330
      Left            =   270
      TabIndex        =   0
      Top             =   2355
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   582
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
   Begin Threed.SSCheck Chk_ss2 
      Height          =   330
      Left            =   270
      TabIndex        =   1
      Top             =   4200
      Width           =   1020
      _ExtentX        =   1799
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
      TabIndex        =   8
      Top             =   2730
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   2143
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin InDate.ULabel ULabel49 
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
      Begin InDate.ULabel ULabel50 
         Height          =   315
         Left            =   6885
         Top             =   210
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "抽真空开始"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   6885
         Top             =   705
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "抽真空结束"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sitxEdit txt_star_vac 
         Height          =   315
         Left            =   8235
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
      Begin CSTextLibCtl.sitxEdit txt_end_vac 
         Height          =   315
         Left            =   8235
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
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_ld_sta_date 
         Height          =   315
         Left            =   1560
         TabIndex        =   28
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
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel52 
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
      Begin CSTextLibCtl.sitxEdit txt_ld_end_date 
         Height          =   315
         Left            =   1560
         TabIndex        =   29
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
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   3555
         Top             =   210
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "开始作业"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sitxEdit txt_start_date 
         Height          =   315
         Left            =   4905
         TabIndex        =   30
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
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   3555
         Top             =   705
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "结束作业"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   4905
         TabIndex        =   31
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
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   10215
         Top             =   210
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "吹氧开始"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   10215
         Top             =   705
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "吹氧结束"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sitxEdit txt_star_oxy 
         Height          =   315
         Left            =   11565
         TabIndex        =   32
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
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_end_oxy 
         Height          =   315
         Left            =   11565
         TabIndex        =   33
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
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
   End
   Begin Threed.SSFrame Frame4 
      Height          =   4425
      Left            =   210
      TabIndex        =   9
      Top             =   4560
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   7805
      _Version        =   196609
      BackColor       =   14737632
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
         TabIndex        =   78
         Top             =   2400
         Width           =   825
      End
      Begin VB.TextBox txt_element_content_c 
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
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   3285
         Width           =   870
      End
      Begin VB.TextBox txt_element_content_mn 
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   3285
         Width           =   870
      End
      Begin VB.TextBox txt_element_content_si 
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
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   3285
         Width           =   870
      End
      Begin VB.TextBox txt_element_content_p 
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
         Left            =   5850
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   3285
         Width           =   870
      End
      Begin VB.TextBox txt_element_content_s 
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
         Left            =   7155
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   3285
         Width           =   870
      End
      Begin VB.TextBox txt_element_content_al 
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
         TabIndex        =   41
         Top             =   3285
         Width           =   870
      End
      Begin VB.TextBox txt_element_content_h 
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
         Left            =   9780
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   3285
         Width           =   870
      End
      Begin VB.TextBox txt_element_content_n 
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
         Left            =   11085
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   3285
         Width           =   870
      End
      Begin VB.TextBox txt_element_content_o 
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
         Left            =   12405
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   3285
         Width           =   870
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
         Left            =   8840
         TabIndex        =   36
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
         Left            =   8840
         TabIndex        =   35
         Top             =   2040
         Width           =   825
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
         TabIndex        =   34
         Top             =   2040
         Width           =   825
      End
      Begin CSTextLibCtl.sidbEdit txt_arrv_steel_wgt 
         Height          =   315
         Left            =   1575
         TabIndex        =   37
         Top             =   705
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   210
         Top             =   705
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   210
         Top             =   1590
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "开始时渣量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   4050
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "处理重量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   4050
         Top             =   705
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   4050
         Top             =   1110
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
      Begin InDate.ULabel ULabel57 
         Height          =   315
         Left            =   7470
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "真空度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   4050
         Top             =   1590
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "抽真空时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   7470
         Top             =   705
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "喷吹代码"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   210
         Top             =   3285
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "RH成分(%)"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   6780
         Top             =   3285
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
         Left            =   5475
         Top             =   3285
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
         Left            =   4185
         Top             =   3285
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
         Left            =   2850
         Top             =   3285
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
         Left            =   1575
         Top             =   3285
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
         Left            =   8100
         Top             =   3285
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel70 
         Height          =   315
         Left            =   9405
         Top             =   3285
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
         Left            =   10710
         Top             =   3285
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
         Left            =   12030
         Top             =   3285
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
            Size            =   9.76
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
         TabIndex        =   47
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
      Begin CSTextLibCtl.sidbEdit txt_slag_wgt 
         Height          =   315
         Left            =   1575
         TabIndex        =   48
         Top             =   1590
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
      Begin CSTextLibCtl.sidbEdit txt_stl_wgt 
         Height          =   315
         Left            =   5415
         TabIndex        =   49
         Top             =   240
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
      Begin CSTextLibCtl.sidbEdit txt_ld_str_temp 
         Height          =   315
         Left            =   5415
         TabIndex        =   50
         Top             =   705
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
      Begin CSTextLibCtl.sidbEdit txt_ld_end_temp 
         Height          =   315
         Left            =   5415
         TabIndex        =   51
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
      Begin CSTextLibCtl.sidbEdit txt_vacuum_dur 
         Height          =   315
         Left            =   5415
         TabIndex        =   52
         Top             =   1590
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
      Begin CSTextLibCtl.sidbEdit txt_BB_YN 
         Height          =   315
         Left            =   8840
         TabIndex        =   53
         Top             =   705
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
         NumIntDigits    =   5
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_vacuum_press 
         Height          =   315
         Left            =   8840
         TabIndex        =   54
         Top             =   240
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
            Charset         =   134
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_BUB_GAS_CONSP_A 
         Height          =   315
         Left            =   8840
         TabIndex        =   55
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
      Begin CSTextLibCtl.sidbEdit txt_BUB_GAS_CONSP_N 
         Height          =   315
         Left            =   8840
         TabIndex        =   56
         Top             =   1590
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
      Begin InDate.ULabel ULabel14 
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
         TabIndex        =   57
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
      Begin InDate.ULabel ULabel27 
         Height          =   315
         Left            =   210
         Top             =   2040
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
         Left            =   4050
         Top             =   2040
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
      Begin CSTextLibCtl.sidbEdit txt_wire_lth1 
         Height          =   315
         Left            =   5415
         TabIndex        =   58
         Top             =   2040
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
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   7470
         Top             =   2040
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
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   10770
         Top             =   2040
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
         Left            =   12150
         TabIndex        =   59
         Top             =   2040
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
         Left            =   210
         Top             =   2400
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   4050
         Top             =   2400
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_wire_lth3 
         Height          =   315
         Left            =   5415
         TabIndex        =   60
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
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   7470
         Top             =   2400
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
         Left            =   10770
         Top             =   2400
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_wire_lth4 
         Height          =   315
         Left            =   12150
         TabIndex        =   61
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
      Begin CSTextLibCtl.sidbEdit txt_O_CONSP 
         Height          =   315
         Left            =   12150
         TabIndex        =   62
         Top             =   1590
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
            Charset         =   134
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
      Begin CSTextLibCtl.sidbEdit txt_BUB_DURATION_A 
         Height          =   315
         Left            =   12150
         TabIndex        =   63
         Top             =   240
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
      Begin CSTextLibCtl.sidbEdit txt_BUB_DURATION_N 
         Height          =   315
         Left            =   12150
         TabIndex        =   64
         Top             =   705
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
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   7470
         Top             =   1110
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
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   7470
         Top             =   1590
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
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   10770
         Top             =   1590
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "氧气用量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   10770
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "氮气吹入时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   10770
         Top             =   705
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "氩气吹入时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sidbEdit txt_BUB_TOTAL_DUR 
         Height          =   315
         Left            =   12150
         TabIndex        =   65
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
            Charset         =   134
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
      Begin InDate.ULabel ULabel29 
         Height          =   315
         Left            =   10770
         Top             =   1110
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "总搅拌时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   210
         Top             =   2760
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "真空槽槽龄"
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
      Begin CSTextLibCtl.sidbEdit txt_groove 
         Height          =   315
         Left            =   1575
         TabIndex        =   79
         Top             =   2760
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
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
         Undo            =   0
         Data            =   0
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
         Height          =   300
         Left            =   9945
         TabIndex        =   5
         Top             =   1590
         Width           =   330
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
         Left            =   9945
         TabIndex        =   77
         Top             =   1230
         Width           =   330
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
         Left            =   6480
         TabIndex        =   76
         Top             =   780
         Width           =   210
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
         Left            =   6480
         TabIndex        =   75
         Top             =   1215
         Width           =   210
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "mbar"
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
         Left            =   9915
         TabIndex        =   74
         Top             =   360
         Width           =   405
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
         Left            =   2655
         TabIndex        =   73
         Top             =   1140
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
         Left            =   2655
         TabIndex        =   72
         Top             =   1635
         Width           =   180
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
         Left            =   6525
         TabIndex        =   71
         Top             =   345
         Width           =   180
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
         Left            =   2670
         TabIndex        =   70
         Top             =   750
         Width           =   75
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
         Left            =   13005
         TabIndex        =   69
         Top             =   2460
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
         Left            =   6255
         TabIndex        =   68
         Top             =   2460
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
         Left            =   13005
         TabIndex        =   67
         Top             =   2115
         Width           =   345
      End
      Begin VB.Label Label9 
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
         Left            =   6240
         TabIndex        =   66
         Top             =   2130
         Width           =   345
      End
   End
End
Attribute VB_Name = "AFF2020C"
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
'-- Program Name      RH
'-- Program ID        AFF2020C
'-- Document No
'-- Designer          SHIN.C.S
'-- Coder             SHIN.C.S
'-- Date              2006.11.25
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
Public sQuery_Rt As String          'Active Form Authority Setting

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
     FormType = "Master"              'form类型
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       
           Call Gp_Ms_Collection(cbo_HEAT_NO, "p", "n", "m", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(cbo_re_cd, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(cbo_prc_line, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(cbo_RH_ID, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
       
       Call Gp_Ms_Collection(txt_ld_sta_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_start_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_end_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ld_end_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
          Call Gp_Ms_Collection(txt_star_vac, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_end_vac, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_star_oxy, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_end_oxy, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
         Call Gp_Ms_Collection(txt_OCCR_DATE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_arrv_steel_wgt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_dep_steel_wgt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_slag_wgt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ld_str_temp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ld_end_temp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_vacuum_dur, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_vacuum_press, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
             Call Gp_Ms_Collection(txt_BB_YN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_BUB_GAS_CONSP_A, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_BUB_GAS_CONSP_N, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_BUB_DURATION_A, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_BUB_DURATION_N, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_BUB_TOTAL_DUR, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_O_CONSP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
          Call Gp_Ms_Collection(txt_wire_cd1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_wire_lth1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_wire_cd2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_wire_lth2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_wire_cd3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_wire_lth3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_wire_cd4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_wire_lth4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_groove, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
 Call Gp_Ms_Collection(txt_element_content_c, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_element_content_mn, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_element_content_si, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_element_content_p, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_element_content_s, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_element_content_al, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_element_content_h, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_element_content_n, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_element_content_o, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_plt, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_proc, "p", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_oper, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           
    'MASTER Collection
     Mc1.Add Item:="AFF2020C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AFF2020C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
  
           Call Gp_Ms_Collection(cbo_HEAT_NO, "p", "n", "m", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(cbo_prc_line, " ", "n", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             Call Gp_Ms_Collection(cbo_ld_id, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(cbo_ref_proc, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_dir_steel_grd, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_stlgrd_n, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_act_steel_grd, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_stlgrd_c, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_mlt_prod_cd, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_ld_sta_date, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_start_date, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_end_date, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ld_end_date, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
          Call Gp_Ms_Collection(txt_star_vac, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_end_vac, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_star_oxy, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_end_oxy, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
              Call Gp_Ms_Collection(txt_proc, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      
      
     Mc2.Add Item:="AFF2020C.P_REFER1", Key:="P-R"
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
   If Len(cbo_HEAT_NO.Text) = 8 Then
      cbo_ld_id.Text = Gf_FloatFind(M_CN1, "SELECT LADLE_NO FROM FP_CHARGE WHERE HEAT_NO = '" + cbo_HEAT_NO.Text + "'")
      If cbo_ld_id.Text = "0" Then
         cbo_ld_id.Text = ""
      End If
   Else
      cbo_ld_id.Text = ""
   End If
End Sub

Private Sub cbo_HEAT_NO_Click()
   If Len(cbo_HEAT_NO.Text) = 8 Then
      cbo_ld_id.Text = Gf_FloatFind(M_CN1, "SELECT LADLE_NO FROM FP_CHARGE WHERE HEAT_NO = '" + cbo_HEAT_NO.Text + "'")
      If cbo_ld_id.Text = "0" Then
         cbo_ld_id.Text = ""
      End If
   Else
      cbo_ld_id.Text = ""
   End If
End Sub

Private Sub cbo_prc_line_Change()
'    Call Gf_HeatNo_ComboAdd(M_CN1, cbo_HEAT_NO, "FP_RHRSLT", "DEP_STEEL_WGT", Trim(cbo_prc_line.Text))
'    If cbo_HEAT_NO.ListCount <> 0 And Trim(cbo_HEAT_NO.Text) = "" Then
'       cbo_HEAT_NO.ListIndex = 0
'       cbo_re_cd.Text = "1"
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
        
         Exit Sub
        
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
            txt_ld_sta_date.SetFocus

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

Private Sub cbo_up_Click()
  Dim V_HEAT_NO As String
  
  If Trim(cbo_HEAT_NO.Text) = "" Then
     Exit Sub
  End If

  V_HEAT_NO = Format(cbo_HEAT_NO + 1, "00000000")
  Call Form_Cls
  
  cbo_HEAT_NO = V_HEAT_NO
  cbo_re_cd.Text = "1"
  
  Call Form_Ref
  
End Sub

Private Sub cbo_Low_Click()
   Dim V_HEAT_NO As String
  
  If Trim(cbo_HEAT_NO.Text) = "" Then
     Exit Sub
  End If
  
  V_HEAT_NO = Format(cbo_HEAT_NO - 1, "00000000")
  Call Form_Cls
  
  cbo_HEAT_NO = V_HEAT_NO
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
     
     Dim A As Date
     
    cbo_shift.AddItem "1"
    cbo_shift.AddItem "2"
    cbo_shift.AddItem "3"

    cbo_group_cd.AddItem "A"
    cbo_group_cd.AddItem "B"
    cbo_group_cd.AddItem "C"
    cbo_group_cd.AddItem "D"
    
    cbo_prc_line.AddItem "1"
    cbo_prc_line.AddItem "2"
    
    cbo_RH_ID.AddItem "1"
    cbo_RH_ID.AddItem "2"
    cbo_RH_ID.AddItem "3"
    cbo_RH_ID.AddItem "4"
 
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gf_ComboAdd(M_CN1, cbo_ld_id, "SELECT CD  FROM ZP_CD WHERE CD_MANA_NO = 'F0004' AND CD LIKE  'S%' ")
   
    cbo_re_cd.AddItem "1"
    cbo_re_cd.AddItem "2"
    cbo_re_cd.AddItem "3"
    
    cbo_prc_line.Text = "1"
    
    Chk_ss1.VALUE = ssCBChecked
    Chk_ss1.ForeColor = &HFF&
    Frame4.Enabled = False
    txt_oper.Text = "1"
    
    txt_emp_cd = sUserID
    txt_emp_cd.ForeColor = &H80000011

    Screen.MousePointer = vbDefault

    Call Gf_RHHeatNo_ComboAdd(M_CN1, cbo_HEAT_NO, "FP_CHARGE", "DEP_STEEL_WGT", Trim(cbo_prc_line.Text))
    If cbo_HEAT_NO.ListCount <> 0 And Trim(cbo_HEAT_NO.Text) = "" Then
       cbo_HEAT_NO.ListIndex = 0
       cbo_re_cd.Text = "1"
    End If
    
    If cbo_HEAT_NO <> "" Then
       Call Form_Ref
    End If
    
    If cbo_shift.Text = "1" Or cbo_shift.Text = "2" Or cbo_shift.Text = "3" Then
       Exit Sub
    End If
    
    sQuery_Rt = Gf_CodeFind(M_CN1, "SELECT SHIFT  FROM GP_SHIFTSTD WHERE TRANCD = 'BH21' ")
    If sQuery_Rt = "" Then
       cbo_shift.ListIndex = 0
    Else
       cbo_shift.ListIndex = Val(sQuery_Rt) - 1
    End If
    
    sQuery_Rt = Gf_CodeFind(M_CN1, "SELECT GROUP_CD  FROM GP_SHIFTSTD WHERE TRANCD = 'BH21' ")
    
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

    cbo_prc_line.Text = "1"
    
    cbo_RH_ID.Text = "1"
    
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
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
End Sub

Public Sub Form_Ref()

    Dim Temp_Heat_No, Temp_Treat_No As String
    
    cbo_HEAT_NO.Text = Mid(cbo_HEAT_NO.Text, 1, 8)
    
    If cbo_re_cd.Text = "" Then
       MsgBox "再处理必须输入！", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
    If cbo_HEAT_NO.Text = "" Then
       MsgBox "炉号必须输入！", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
    Temp_Heat_No = Mid(cbo_HEAT_NO.Text, 1, 8)
    Temp_Treat_No = cbo_re_cd.Text
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    cbo_HEAT_NO.Text = Temp_Heat_No
    cbo_re_cd.Text = Temp_Treat_No
    If Gf_Ms_Refer(M_CN1, Mc1, Nothing, Nothing, False) Or Gf_Ms_Refer(M_CN1, Mc2, Nothing, Nothing, False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        Call Gp_Ms_ControlLock(Mc2("pControl"), True)
        
        txt_dir_steel_grd.Enabled = True
        txt_dir_steel_grd.Locked = True
        txt_dir_steel_grd.ForeColor = &H80000011
        
        txt_act_steel_grd.Enabled = True
        cbo_re_cd.Enabled = True
    Else
       MsgBox "无相关记录", vbInformation, "系统提示信息"
    End If
    
    If txt_emp_cd = "" Then
       txt_emp_cd = sUserID
       txt_emp_cd.ForeColor = &H80000011
    End If
End Sub

Public Sub Form_Pro()

    Dim sMesg As String
    
    cbo_HEAT_NO.Text = Mid(cbo_HEAT_NO.Text, 1, 8)
    If cbo_prc_line.Text <> "1" And cbo_prc_line.Text <> "2" And cbo_prc_line.Text <> "3" Then
       MsgBox "   RH stattion ！Please Input  ", vbCritical, "错误提示"
       Exit Sub
    End If
    
    If Chk_ss2 = True Then
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
            
        If txt_act_steel_grd = "" Then
           MsgBox "实际钢种号必须输入！", vbCritical, "系统提示信息"
           Exit Sub
        End If
                
    End If
    
'    If cbo_LD_ID.Text = "" Then
'       MsgBox "   钢包号 ！Please Input  ", vbCritical, "错误提示"
'       Exit Sub
'    End If
    
    txt_emp_cd = sUserID
    
    If Len(Trim(cbo_HEAT_NO.Text)) <> 8 Then
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
    
    txt_emp_cd = sUserID
    txt_act_steel_grd.Enabled = True
    
End Sub


Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    If Not Gf_Ms_Del(M_CN1, Mc2) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
End Sub

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

Private Sub txt_dir_steel_grd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sQuery As String
    
    sQuery = "SELECT STEEL_GRD_DETAIL FROM QP_STLGRD_INF WHERE STLGRD = '" + txt_dir_steel_grd.Text + "'"
    txt_dir_steel_grd.ToolTipText = Gf_CodeFind(M_CN1, sQuery)
End Sub

Private Sub txt_end_date_DblClick()
    txt_end_date.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_end_oxy_DblClick()
    txt_end_oxy.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_end_vac_DblClick()
    txt_end_vac.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_ld_end_date_DblClick()
    txt_ld_end_date.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_ld_sta_date_DblClick()
    txt_ld_sta_date.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_OCCR_DATE_DblClick()
    txt_OCCR_DATE.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_star_oxy_DblClick()
    txt_star_oxy.RawData = Format(Now, "YYYYMMDDHHMM")
End Sub

Private Sub txt_star_vac_DblClick()
    txt_star_vac.RawData = Format(Now, "YYYYMMDDHHMM")
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

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_RHHeatNo_ComboAdd
'   2.Name         :
'   3.Input  Value : Conn Connection, Cbo Variant, sTableName String, sColId String, sPrcLine String, {ClsChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : SHIN.C.S
'   6.Create Date  : 2006. 01 .24
'   7.Modify Date  :
'   8.Comment      : Add Heat No in combo box
'---------------------------------------------------------------------------------------
Public Function Gf_RHHeatNo_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, _
                                   sTableName As String, sColId As String, sPrcLine As String, _
                                   Optional ClsChk As Boolean = True) As Boolean

On Error GoTo Gf_RHHeatNo_ComboAdd_Error
    
    Dim AdoRs As ADODB.Recordset
    Dim sQuery As String
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_RHHeatNo_ComboAdd = False: Exit Function
    End If
    
    If ClsChk Then
        Cbo.Clear
    End If
    
     
    sQuery = "           SELECT HEAT_MANA_NO"
    sQuery = sQuery & "    FROM EP_CHARGE_INS"
    sQuery = sQuery & "   WHERE PRC_STS IN('A','B')"
    sQuery = sQuery & "     AND MLT_PROC_CD LIKE '%BH2%'"
    sQuery = sQuery & "     AND HEAT_MANA_NO NOT IN(SELECT HEAT_NO FROM NISCO.FP_RHRSLT)   "
    sQuery = sQuery & "     AND ROWNUM <= 10"
    sQuery = sQuery & "   ORDER BY HEAT_MANA_NO ASC"
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
                Cbo.AddItem AdoRs.Fields(0)
            End If
            AdoRs.MoveNext
            
        Wend
        Gf_RHHeatNo_ComboAdd = True
    Else
        Gf_RHHeatNo_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function
Gf_RHHeatNo_ComboAdd_Error:

    Set AdoRs = Nothing
    Gf_RHHeatNo_ComboAdd = False

End Function

