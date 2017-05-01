VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AFC2020C 
   Caption         =   "CAS实绩修改及查询界面_AFC2020C"
   ClientHeight    =   9225
   ClientLeft      =   270
   ClientTop       =   2115
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame2 
      Height          =   795
      Left            =   210
      TabIndex        =   3
      Top             =   1020
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   1402
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.TextBox txt_act_steel_grd 
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
         Left            =   7410
         MaxLength       =   11
         TabIndex        =   16
         Top             =   210
         Width           =   1275
      End
      Begin VB.TextBox txt_dir_steel_grd 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   210
         Width           =   1275
      End
      Begin VB.TextBox txt_mlt_prod_cd 
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
         Left            =   13065
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   210
         Width           =   1440
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
         Left            =   3075
         TabIndex        =   13
         Top             =   210
         Width           =   2535
      End
      Begin VB.TextBox txt_stlgrd_s 
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
         Left            =   8685
         TabIndex        =   12
         Top             =   210
         Width           =   2535
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   360
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "目标钢种号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   5955
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "实际钢种号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel32 
         Height          =   315
         Left            =   11610
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "工艺路线"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   795
      Left            =   210
      TabIndex        =   2
      Top             =   180
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   1402
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.ComboBox cbo_group_cd 
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
         ItemData        =   "AFC2020C.frx":0000
         Left            =   10560
         List            =   "AFC2020C.frx":0002
         TabIndex        =   11
         Tag             =   "班别"
         Top             =   240
         Width           =   720
      End
      Begin VB.ComboBox cbo_shift 
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
         ItemData        =   "AFC2020C.frx":0004
         Left            =   8115
         List            =   "AFC2020C.frx":0006
         TabIndex        =   10
         Tag             =   "班次"
         Top             =   240
         Width           =   720
      End
      Begin VB.ComboBox cbo_heat_no 
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
         ItemData        =   "AFC2020C.frx":0008
         Left            =   1785
         List            =   "AFC2020C.frx":000A
         TabIndex        =   9
         Tag             =   "炉号"
         Top             =   240
         Width           =   1380
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
         ItemData        =   "AFC2020C.frx":000C
         Left            =   5655
         List            =   "AFC2020C.frx":000E
         TabIndex        =   8
         Tag             =   "炉座号"
         Top             =   240
         Width           =   705
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
         Height          =   315
         Left            =   13050
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   7
         Tag             =   "作业人员"
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cbo_down 
         Caption         =   "▼"
         Height          =   255
         Left            =   3300
         TabIndex        =   6
         Top             =   405
         Width           =   375
      End
      Begin VB.CommandButton cbo_up 
         Caption         =   "▲"
         Height          =   255
         Left            =   3300
         TabIndex        =   5
         Top             =   120
         Width           =   375
      End
      Begin InDate.ULabel ULabel63 
         Height          =   315
         Left            =   360
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "炉号"
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
      Begin InDate.ULabel ULabel64 
         Height          =   315
         Left            =   6675
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
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
      Begin InDate.ULabel ULabel69 
         Height          =   315
         Left            =   9105
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "班别"
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
      Begin InDate.ULabel ULabel70 
         Height          =   315
         Left            =   11610
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "作业人员"
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
         Left            =   4200
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "炉座号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox txt_proc 
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
      Left            =   12975
      TabIndex        =   1
      Text            =   "BG"
      Top             =   -15
      Visible         =   0   'False
      Width           =   675
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
      Height          =   315
      Left            =   13680
      TabIndex        =   0
      Text            =   "B1"
      Top             =   0
      Visible         =   0   'False
      Width           =   675
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   7125
      Left            =   210
      TabIndex        =   4
      Top             =   1950
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   12568
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.ComboBox cbo_ld_id 
         Height          =   315
         ItemData        =   "AFC2020C.frx":0010
         Left            =   2760
         List            =   "AFC2020C.frx":0012
         TabIndex        =   21
         Top             =   2490
         Width           =   1065
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
         Left            =   2760
         TabIndex        =   20
         Top             =   4095
         Width           =   945
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
         Left            =   2760
         TabIndex        =   19
         Top             =   3645
         Width           =   945
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
         Left            =   2760
         TabIndex        =   18
         Top             =   4545
         Width           =   945
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
         Left            =   2760
         TabIndex        =   17
         Top             =   5010
         Width           =   945
      End
      Begin InDate.ULabel ULabel43 
         Height          =   315
         Left            =   9180
         Top             =   2490
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "氧气用量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel33 
         Height          =   315
         Left            =   1290
         Top             =   2490
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "钢包号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel39 
         Height          =   315
         Left            =   5190
         Top             =   2490
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "开始温度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel40 
         Height          =   315
         Left            =   5190
         Top             =   2970
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "结束温度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel45 
         Height          =   315
         Left            =   9180
         Top             =   2970
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "氩气用量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   1290
         Top             =   2970
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "处理量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel47 
         Height          =   315
         Left            =   9180
         Top             =   3645
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "氩气平均压力"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel49 
         Height          =   315
         Left            =   9180
         Top             =   4095
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "氩气最小压力"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_steel_net_wgt 
         Height          =   315
         Left            =   2760
         TabIndex        =   22
         Top             =   2970
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
      Begin CSTextLibCtl.sidbEdit txt_cas_end_temp 
         Height          =   315
         Left            =   6645
         TabIndex        =   23
         Top             =   2970
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
      Begin CSTextLibCtl.sidbEdit txt_cas_str_temp 
         Height          =   315
         Left            =   6645
         TabIndex        =   24
         Top             =   2490
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
      Begin CSTextLibCtl.sidbEdit txt_avg_ar_press 
         Height          =   315
         Left            =   10635
         TabIndex        =   25
         Top             =   3645
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
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
      Begin CSTextLibCtl.sidbEdit txt_cas_ar_usage 
         Height          =   315
         Left            =   10635
         TabIndex        =   26
         Top             =   2970
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
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
      Begin CSTextLibCtl.sidbEdit txt_cas_o_using 
         Height          =   315
         Left            =   10635
         TabIndex        =   27
         Top             =   2490
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
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
         RawData         =   "0.0000"
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
         NumDecDigits    =   4
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_min_ar_press 
         Height          =   315
         Left            =   10635
         TabIndex        =   28
         Top             =   4095
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
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
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   1290
         Top             =   570
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "实绩发生时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sitxEdit txt_OCCR_DATE 
         Height          =   315
         Left            =   2760
         TabIndex        =   29
         Top             =   570
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
         Left            =   1290
         Top             =   3645
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "喂丝种类 1"
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
         Left            =   5190
         Top             =   3645
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "喂丝重量 1"
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
      Begin CSTextLibCtl.sidbEdit txt_wire_lth1 
         Height          =   315
         Left            =   6645
         TabIndex        =   30
         Top             =   3645
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
      Begin InDate.ULabel ULabel29 
         Height          =   315
         Left            =   1290
         Top             =   4095
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "喂丝种类 2"
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
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Left            =   5190
         Top             =   4095
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "喂丝重量 2"
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
      Begin CSTextLibCtl.sidbEdit txt_wire_lth2 
         Height          =   315
         Left            =   6645
         TabIndex        =   31
         Top             =   4095
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
         Left            =   10635
         TabIndex        =   32
         Top             =   5010
         Width           =   1770
         _Version        =   262145
         _ExtentX        =   3122
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
      Begin InDate.ULabel ULabel44 
         Height          =   315
         Left            =   9180
         Top             =   5010
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "上次修改时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   1290
         Top             =   4545
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "喂丝种类 3"
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   5190
         Top             =   4545
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "喂丝重量 3"
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
      Begin CSTextLibCtl.sidbEdit txt_wire_lth3 
         Height          =   315
         Left            =   6645
         TabIndex        =   33
         Top             =   4545
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
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   1290
         Top             =   5010
         Width           =   1410
         _ExtentX        =   2487
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   5190
         Top             =   5010
         Width           =   1410
         _ExtentX        =   2487
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
         Left            =   6645
         TabIndex        =   34
         Top             =   5010
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
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   9180
         Top             =   4545
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "氩气最大压力"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_max_ar_press 
         Height          =   315
         Left            =   10635
         TabIndex        =   35
         Top             =   4545
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
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
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   1290
         Top             =   1050
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "开始时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel25 
         Height          =   315
         Left            =   1290
         Top             =   1545
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "结束时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sitxEdit txt_occr_date_1 
         Height          =   315
         Left            =   2760
         TabIndex        =   36
         Top             =   1050
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
      Begin CSTextLibCtl.sitxEdit txt_occr_date_2 
         Height          =   315
         Left            =   2760
         TabIndex        =   37
         Top             =   1545
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
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "℃"
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
         Left            =   7695
         TabIndex        =   49
         Top             =   2535
         Width           =   180
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "m3"
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
         Left            =   11685
         TabIndex        =   48
         Top             =   2535
         Width           =   195
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   3840
         TabIndex        =   47
         Top             =   3030
         Width           =   180
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7695
         TabIndex        =   46
         Top             =   3030
         Width           =   180
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "m3"
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
         Left            =   11685
         TabIndex        =   45
         Top             =   3000
         Width           =   195
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Mpa"
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
         Left            =   11685
         TabIndex        =   44
         Top             =   4125
         Width           =   360
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
         Left            =   7470
         TabIndex        =   43
         Top             =   4125
         Width           =   345
      End
      Begin VB.Label Label18 
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
         Left            =   7470
         TabIndex        =   42
         Top             =   3720
         Width           =   345
      End
      Begin VB.Label Label1 
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
         Left            =   7470
         TabIndex        =   41
         Top             =   4620
         Width           =   345
      End
      Begin VB.Label Label2 
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
         Left            =   7470
         TabIndex        =   40
         Top             =   5070
         Width           =   345
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Mpa"
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
         Left            =   11685
         TabIndex        =   39
         Top             =   4620
         Width           =   360
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Mpa"
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
         Left            =   11685
         TabIndex        =   38
         Top             =   3720
         Width           =   360
      End
   End
End
Attribute VB_Name = "AFC2020C"
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
'-- Program Name      CAS
'-- Program ID        AFC2020C
'-- Document No
'-- Designer          KSH
'-- Coder             KSH
'-- Date              2006.2.28
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
Public sQuery_Rt As String              'Active Form Authority Setting
       
Dim pControl1 As New Collection      'Master Primary Key Collection
Dim nControl1 As New Collection      'Master Necessary Collection
Dim mControl1 As New Collection      'Master Maxlength check Collection
Dim iControl1 As New Collection      'Master Insert Collection
Dim rControl1 As New Collection      'Master Refer Collection
Dim cControl1 As New Collection      'Master Copy Collection
Dim aControl1 As New Collection      'Master -> Spread Collection
Dim lControl1 As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"              'form类型
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")

           Call Gp_Ms_Collection(cbo_heat_no, "p", "n", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(cbo_prc_line, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             Call Gp_Ms_Collection(cbo_shift, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(cbo_group_cd, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            Call Gp_Ms_Collection(txt_emp_cd, " ", " ", " ", "i", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_dir_steel_grd, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_stlgrd_n, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_act_steel_grd, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_stlgrd_s, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_mlt_prod_cd, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_OCCR_DATE, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_occr_date_1, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_occr_date_2, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             Call Gp_Ms_Collection(cbo_ld_id, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_steel_net_wgt, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_cas_str_temp, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_cas_end_temp, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_cas_o_using, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_cas_ar_usage, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_avg_ar_press, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_min_ar_press, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_max_ar_press, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_wire_cd1, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_wire_lth1, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_wire_cd2, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_wire_lth2, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_wire_cd3, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_wire_lth3, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_wire_cd4, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_wire_lth4, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_upd_date, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
               
               Call Gp_Ms_Collection(txt_plt, " ", " ", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_proc, " ", " ", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     
'MASTER Collection
     Mc1.Add Item:="AFC2020C.P_REFER1", Key:="P-R"
     Mc1.Add Item:="AFC2020C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"
      

     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0

End Sub

Private Sub cbo_HEAT_NO_Change()
    If Len(Trim(cbo_heat_no.Text)) = 8 Then
       cbo_prc_line.Text = Mid(cbo_heat_no, 3, 1)
    End If
End Sub

Private Sub cbo_prc_line_Click()
    Dim sQuery As String
    
    sQuery = "SELECT A.HEAT_MANA_NO,B.STEEL_NET_WGT "
    sQuery = sQuery & "  FROM EP_CHARGE_INS A, FP_CONRSLT B  "
    sQuery = sQuery & " WHERE A.PRC_STS       IN ('A','B')   "
    sQuery = sQuery & "   AND A.PRC_LINE      = '" & cbo_prc_line.Text & "'"
    sQuery = sQuery & "   AND A.HEAT_MANA_NO  = B.HEAT_NO(+)  "
    sQuery = sQuery & "   AND ROWNUM <= 15                   "
    sQuery = sQuery & " ORDER BY A.HEAT_MANA_NO "

    Call Gf_ComboAdd2(M_CN1, cbo_heat_no, sQuery)
    If cbo_heat_no.ListCount <> 0 And Trim(cbo_heat_no.Text) = "" Then
       cbo_heat_no.ListIndex = 0
    End If
    
End Sub

Private Sub cbo_up_Click()
  Dim V_HEAT_NO As String
  
  If Trim(cbo_heat_no.Text) = "" Then
     Exit Sub
  End If

  V_HEAT_NO = Mid(cbo_heat_no, 1, 3) + Format(Val(Mid(cbo_heat_no, 4, 5)) + 1, "00000")
  Call Form_Cls
  
  cbo_heat_no = V_HEAT_NO
  cbo_prc_line.Text = Mid(cbo_heat_no, 3, 1)
  Call Form_Ref
  
End Sub

Private Sub cbo_down_Click()
   Dim V_HEAT_NO As String
  
  If Trim(cbo_heat_no.Text) = "" Then
     Exit Sub
  End If
  
  V_HEAT_NO = Mid(cbo_heat_no, 1, 3) + Format(Val(Mid(cbo_heat_no, 4, 5)) - 1, "00000")
  Call Form_Cls
  
  cbo_heat_no = V_HEAT_NO
  cbo_prc_line.Text = Mid(cbo_heat_no, 3, 1)
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
    Dim sQuery  As String
    
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
    
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    cbo_prc_line.Text = "1"
    
    Call cbo_prc_line_Click
  
    Call Gf_ComboAdd(M_CN1, cbo_ld_id, "SELECT CD  FROM ZP_CD WHERE CD_MANA_NO='F0004' AND CD LIKE  'S%' ")

    txt_act_steel_grd.Enabled = False
    
    txt_emp_cd = sUserID
    txt_emp_cd.ForeColor = &H80000011
    
    If cbo_heat_no <> "" Then
       Call Form_Ref
    End If
    
    Screen.MousePointer = vbDefault
    If cbo_shift.Text = "1" Or cbo_shift.Text = "2" Or cbo_shift.Text = "3" Then
       Exit Sub
    End If
    
    sQuery_Rt = Gf_CodeFind(M_CN1, "SELECT SHIFT  FROM GP_SHIFTSTD WHERE TRANCD = 'BC21' ")
    If sQuery_Rt = "" Then
       cbo_shift.ListIndex = 0
    Else
       cbo_shift.ListIndex = Val(sQuery_Rt) - 1
    End If
    
    sQuery_Rt = Gf_CodeFind(M_CN1, "SELECT GROUP_CD  FROM GP_SHIFTSTD WHERE TRANCD = 'BC21' ")

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
    
    Set pControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
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
    pControl1(1).SetFocus
     
    cbo_prc_line.Text = "1"
    Call cbo_prc_line_Click

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
    
    Dim Scr_wgt, Hm_wgt, Steel_wgt As Integer
    cbo_heat_no.Text = Mid(cbo_heat_no.Text, 1, 8)
    
    If Trim(cbo_heat_no.Text) = "" Then
       Call Gp_MsgBoxDisplay("炉号必须输入", "", "错误提示")
    ElseIf Len(Trim(cbo_heat_no.Text)) <> 8 Then
       Call Gp_MsgBoxDisplay("炉号长度应为8位", "", "错误提示")
    Else
    
        If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Nothing, False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("pControl"), True)
            txt_act_steel_grd.Enabled = True
        End If
        If cbo_prc_line = "" Then
           cbo_prc_line.Text = Mid(cbo_heat_no, 3, 1)
        End If
    End If
             
    If txt_emp_cd = "" Then
       txt_emp_cd = sUserID
       txt_emp_cd.ForeColor = &H80000011
    End If
End Sub

Public Sub Form_Pro()
    
    Dim sMesg As String
    cbo_heat_no.Text = Mid(cbo_heat_no.Text, 1, 8)
    
    If txt_act_steel_grd.Text = "" Then
       MsgBox "实际钢种号必须输入", vbCritical, "错误提示"
       Exit Sub
    End If
    
    If cbo_prc_line.Text <> "1" And cbo_prc_line.Text <> "2" And cbo_prc_line.Text <> "3" Then
       MsgBox "炉座号必须输入", vbCritical, "错误提示"
       Exit Sub
    End If
    
    If txt_OCCR_DATE.RawData = "" Then
       MsgBox "实绩发生时间必须输入", vbCritical, "错误提示"
       Exit Sub
    End If
    If txt_occr_date_1.RawData = "" Then
       MsgBox "开始时间必须输入", vbCritical, "错误提示"
       Exit Sub
    End If
    If txt_occr_date_2.RawData = "" Then
       MsgBox "结束时间必须输入", vbCritical, "错误提示"
       Exit Sub
    End If
    If txt_emp_cd = "" Then
       MsgBox "作业人员必须输入", vbCritical, "错误提示"
       Exit Sub
    End If
    If cbo_ld_id.Text = "" Then
       MsgBox "钢包号必须输入", vbCritical, "错误提示"
       Exit Sub
    End If
    
    If Len(Trim(cbo_heat_no.Text)) <> 8 Then
        sMesg = sMesg + " 炉号必须是8位"
        Call Gp_MsgBoxDisplay(sMesg)
       Exit Sub
    Else
           
        If Gf_Mc_Authority(sAuthority, Mc1) Then
            If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
               Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            End If
       End If
    End If
    
    
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
End Sub

Private Sub sf1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub txt_act_steel_grd_Change()

    Dim sQuery As String

    If Len(Trim(txt_act_steel_grd.Text)) >= 10 Then
        sQuery = "SELECT STEEL_GRD_DETAIL FROM qp_nisco_chmc WHERE STLGRD = '" + txt_act_steel_grd.Text + "' AND NVL(STLGRD_FL,'X') <> 'H' "
        txt_stlgrd_s.Text = Gf_CodeFind(M_CN1, sQuery)
    Else
        txt_stlgrd_s.Text = ""
    End If
    
End Sub

Private Sub txt_act_steel_grd_DblClick()

    Call txt_act_steel_grd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_act_steel_grd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
         DD.sWitch = "MS"
         DD.rControl.Add Item:=txt_act_steel_grd
         DD.rControl.Add Item:=txt_stlgrd_s
        
         Call Pf_Common_DD(M_CN1, KeyCode)
    Else
         Dim sQuery As String
         If Len(Trim(txt_act_steel_grd.Text)) >= 10 Then
            sQuery = "SELECT STEEL_GRD_DETAIL FROM qp_nisco_chmc WHERE STLGRD = '" + txt_act_steel_grd.Text + "' AND NVL(STLGRD_FL,'X') <> 'H' "
            txt_stlgrd_s.Text = Gf_CodeFind(M_CN1, sQuery)
         Else
            txt_stlgrd_s.Text = ""
         End If
        
    End If
    
End Sub


Private Sub txt_occr_date_1_DblClick()
         
    txt_occr_date_1.RawData = Format(Now, "YYYYMMDDHHMM")
          
End Sub


Private Sub txt_occr_date_2_DblClick()
         
    txt_occr_date_2.RawData = Format(Now, "YYYYMMDDHHMM")
          
End Sub


Private Sub txt_OCCR_DATE_DblClick()
         
    txt_OCCR_DATE.RawData = Format(Now, "YYYYMMDDHHMM")
          
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
        
         Exit Sub
        
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
        
         Exit Sub
        
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
        
         Exit Sub
        
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
        
         Exit Sub
        
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



