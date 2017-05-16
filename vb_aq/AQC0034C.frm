VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AQC0034C 
   Caption         =   "产品检验实绩录入（力学）_AQC0034C"
   ClientHeight    =   11970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11970
   ScaleWidth      =   14475
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   11970
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14475
      _ExtentX        =   25532
      _ExtentY        =   21114
      _Version        =   196609
      AutoSize        =   1
      SplitterBarAppearance=   1
      BorderStyle     =   0
      PaneTree        =   "AQC0034C.frx":0000
      Begin Threed.SSPanel SSPanel3 
         Height          =   705
         Left            =   0
         TabIndex        =   24
         Top             =   600
         Width           =   14475
         _ExtentX        =   25532
         _ExtentY        =   1244
         _Version        =   196609
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   2
            Left            =   90
            Top             =   0
            Width           =   1950
            _ExtentX        =   3440
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
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   3
            Left            =   2040
            Top             =   0
            Width           =   1440
            _ExtentX        =   2540
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
            Index           =   4
            Left            =   4770
            Top             =   0
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   556
            Caption         =   "标准号"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            Index           =   5
            Left            =   8040
            Top             =   0
            Width           =   1950
            _ExtentX        =   3440
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
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   6
            Left            =   9990
            Top             =   0
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   556
            Caption         =   "序列号"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            Index           =   7
            Left            =   10590
            Top             =   0
            Width           =   1950
            _ExtentX        =   3440
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
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   10
            Left            =   12540
            Top             =   0
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "订单厚度"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            Index           =   11
            Left            =   13740
            Top             =   0
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "订单宽度"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            Index           =   12
            Left            =   3480
            Top             =   0
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "取样日期"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.ULabel lbl_STLGRD 
            Height          =   345
            Left            =   90
            Top             =   300
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   609
            Caption         =   ""
            Alignment       =   1
            BackColor       =   15529975
            BackgroundStyle =   1
            BorderStyle     =   1
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
         Begin InDate.ULabel lbl_HEAT_NO 
            Height          =   345
            Left            =   2040
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            Caption         =   ""
            Alignment       =   1
            BackColor       =   15529975
            BackgroundStyle =   1
            BorderStyle     =   1
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
         Begin InDate.ULabel lbl_STDSPEC 
            Height          =   345
            Left            =   4800
            Top             =   300
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   609
            Caption         =   ""
            Alignment       =   1
            BackColor       =   15529975
            BackgroundStyle =   1
            BorderStyle     =   1
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
         Begin InDate.ULabel lbl_ORD_NO 
            Height          =   345
            Left            =   8040
            Top             =   300
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   609
            Caption         =   ""
            Alignment       =   1
            BackColor       =   15529975
            BackgroundStyle =   1
            BorderStyle     =   1
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
         Begin InDate.ULabel lbl_ORD_ITEM 
            Height          =   345
            Left            =   9990
            Top             =   300
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   609
            Caption         =   ""
            Alignment       =   1
            BackColor       =   15529975
            BackgroundStyle =   1
            BorderStyle     =   1
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
         Begin InDate.ULabel lbl_ENDUSE_CD 
            Height          =   345
            Left            =   10620
            Top             =   300
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            Caption         =   ""
            Alignment       =   1
            BackColor       =   15529975
            BackgroundStyle =   1
            BorderStyle     =   1
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
         Begin InDate.ULabel lbl_ORD_THK 
            Height          =   345
            Left            =   12540
            Top             =   300
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            Caption         =   ""
            BackColor       =   15529975
            BackgroundStyle =   1
            BorderStyle     =   1
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
         Begin InDate.ULabel lbl_ORD_WID 
            Height          =   345
            Left            =   13740
            Top             =   300
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   609
            Caption         =   ""
            BackColor       =   15529975
            BackgroundStyle =   1
            BorderStyle     =   1
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
         Begin InDate.ULabel lbl_Cut_DD 
            Height          =   345
            Left            =   3480
            Top             =   300
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   609
            Caption         =   ""
            Alignment       =   1
            BackColor       =   15529975
            BackgroundStyle =   1
            BorderStyle     =   1
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
            Index           =   61
            Left            =   6780
            Top             =   0
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            Caption         =   "发布年度"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.ULabel lbl_STD_YY 
            Height          =   345
            Left            =   6780
            Top             =   300
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   609
            Caption         =   ""
            Alignment       =   1
            BackColor       =   15529975
            BackgroundStyle =   1
            BorderStyle     =   1
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   525
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   14475
         _ExtentX        =   25532
         _ExtentY        =   926
         _Version        =   196609
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txt_UPD_EMP 
            Height          =   315
            Left            =   13440
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   120
            Width           =   1395
         End
         Begin VB.TextBox txt_INPUT_EMP 
            Height          =   315
            Left            =   10710
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   120
            Width           =   1395
         End
         Begin Threed.SSRibbon SSRibbon_SMP_TYPE_KND 
            Height          =   345
            Left            =   6720
            TabIndex        =   33
            Top             =   120
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   609
            _Version        =   196609
            Font3D          =   5
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "现在录入常规样"
         End
         Begin VB.TextBox txt_SMP_NO 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   1470
            MaxLength       =   14
            TabIndex        =   1
            Tag             =   "99"
            Top             =   120
            Width           =   2655
         End
         Begin VB.TextBox txt_SMP_CUT_LOC 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   5640
            MaxLength       =   1
            TabIndex        =   2
            Tag             =   "取样位置"
            Top             =   120
            Width           =   435
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   0
            Left            =   90
            Top             =   120
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "试样编号"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            Index           =   1
            Left            =   4260
            Top             =   120
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "取样位置"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txt_smp_loc_p 
            Height          =   345
            Left            =   7920
            TabIndex        =   37
            Tag             =   "取样位置"
            Top             =   120
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txt_INS_EMP 
            Height          =   375
            Left            =   6360
            TabIndex        =   35
            Tag             =   "INS_EMP"
            Top             =   210
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txt_smp_no_p 
            Height          =   315
            Left            =   7050
            TabIndex        =   36
            Tag             =   "99"
            Top             =   60
            Visible         =   0   'False
            Width           =   1245
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   15
            Left            =   9390
            Top             =   120
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "录入人员"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            Index           =   31
            Left            =   12120
            Top             =   120
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "修改人员"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Height          =   10590
         Left            =   0
         TabIndex        =   22
         Top             =   1380
         Width           =   14475
         _ExtentX        =   25532
         _ExtentY        =   18680
         _Version        =   196609
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin TabDlg.SSTab SSTab1 
            Height          =   8055
            Left            =   0
            TabIndex        =   25
            Top             =   120
            Width           =   14835
            _ExtentX        =   26167
            _ExtentY        =   14208
            _Version        =   393216
            Tabs            =   4
            Tab             =   1
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "拉伸/高温拉伸"
            TabPicture(0)   =   "AQC0034C.frx":0072
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "SSPanel5"
            Tab(0).Control(1)=   "SSPanel6(0)"
            Tab(0).Control(2)=   "SSPanel4(0)"
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "追加拉伸/追加高温拉伸"
            TabPicture(1)   =   "AQC0034C.frx":008E
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "SSPanel6(1)"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "ULabel1(36)"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "SSPanel9(1)"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "SSPanel4(1)"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).ControlCount=   4
            TabCaption(2)   =   "冲击/时效冲击"
            TabPicture(2)   =   "AQC0034C.frx":00AA
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SSPanel11"
            Tab(2).Control(1)=   "SSPanel10(0)"
            Tab(2).ControlCount=   2
            TabCaption(3)   =   "配置化材质项目"
            TabPicture(3)   =   "AQC0034C.frx":00C6
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Ss3"
            Tab(3).ControlCount=   1
            Begin Threed.SSPanel SSPanel4 
               Height          =   3495
               Index           =   0
               Left            =   -67800
               TabIndex        =   26
               Top             =   360
               Width           =   7515
               _ExtentX        =   13256
               _ExtentY        =   6165
               _Version        =   196609
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin Threed.SSPanel SSPanel8 
                  Height          =   3105
                  Index           =   0
                  Left            =   0
                  TabIndex        =   28
                  Top             =   240
                  Width           =   7455
                  _ExtentX        =   13150
                  _ExtentY        =   5477
                  _Version        =   196609
                  BevelOuter      =   1
                  RoundedCorners  =   0   'False
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_YP_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   2400
                     TabIndex        =   3
                     Tag             =   "8"
                     Top             =   60
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_TS_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   2400
                     TabIndex        =   4
                     Tag             =   "9"
                     Top             =   540
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_RST_1 
                     Height          =   315
                     Index           =   0
                     Left            =   2400
                     TabIndex        =   5
                     Tag             =   "10"
                     Top             =   915
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_EL_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   2400
                     TabIndex        =   6
                     Tag             =   "11"
                     Top             =   1320
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   2
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_SNPP_EL_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   2400
                     TabIndex        =   7
                     Tag             =   "12"
                     Top             =   1800
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   2400
                     TabIndex        =   8
                     Tag             =   "13"
                     Top             =   2160
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ul_H_YP 
                     Height          =   315
                     Index           =   0
                     Left            =   6360
                     Top             =   60
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   17
                     Left            =   60
                     Top             =   60
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "屈服强度   YP   MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   18
                     Left            =   60
                     Top             =   540
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "抗拉强度   TS   MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   19
                     Left            =   60
                     Top             =   915
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "断面收缩率   RA    %"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   20
                     Left            =   60
                     Top             =   1320
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "断后伸长率   EL    %"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   21
                     Left            =   60
                     Top             =   1800
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "规定非比例伸长应力MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   22
                     Left            =   60
                     Top             =   2160
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "规定残余伸长应力  MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ul_H_TS 
                     Height          =   315
                     Index           =   0
                     Left            =   6360
                     Top             =   540
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_H_RA 
                     Height          =   315
                     Index           =   0
                     Left            =   6360
                     Top             =   915
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_H_EL 
                     Height          =   315
                     Index           =   0
                     Left            =   6360
                     Top             =   1320
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_H_SNPP_EL 
                     Height          =   315
                     Index           =   0
                     Left            =   6360
                     Top             =   1800
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_H_SP_EL 
                     Height          =   315
                     Index           =   0
                     Left            =   6360
                     Top             =   2160
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_RST_A 
                     Height          =   315
                     Index           =   0
                     Left            =   4950
                     TabIndex        =   143
                     Tag             =   "10"
                     Top             =   915
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_RST_3 
                     Height          =   315
                     Index           =   0
                     Left            =   4110
                     TabIndex        =   144
                     Tag             =   "10"
                     Top             =   915
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_RST_2 
                     Height          =   315
                     Index           =   0
                     Left            =   3270
                     TabIndex        =   145
                     Tag             =   "10"
                     Top             =   915
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   50
                     Left            =   60
                     Top             =   2640
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "均匀变形伸长率UEL %"
                     Alignment       =   0
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
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                     Height          =   315
                     Index           =   2
                     Left            =   2400
                     TabIndex        =   169
                     Tag             =   "54"
                     Top             =   2640
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   2
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ul_H_SP_EL 
                     Height          =   315
                     Index           =   2
                     Left            =   6360
                     Top             =   2640
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin VB.Line Line3 
                     Index           =   11
                     X1              =   0
                     X2              =   7440
                     Y1              =   3000
                     Y2              =   3000
                  End
                  Begin VB.Line Line3 
                     Index           =   10
                     X1              =   0
                     X2              =   7440
                     Y1              =   1680
                     Y2              =   1680
                  End
                  Begin VB.Line Line4 
                     Index           =   0
                     X1              =   6240
                     X2              =   6240
                     Y1              =   -360
                     Y2              =   3000
                  End
                  Begin VB.Line Line3 
                     Index           =   3
                     X1              =   0
                     X2              =   7560
                     Y1              =   2520
                     Y2              =   2520
                  End
                  Begin VB.Line Line3 
                     Index           =   2
                     X1              =   0
                     X2              =   7440
                     Y1              =   1260
                     Y2              =   1260
                  End
                  Begin VB.Line Line3 
                     Index           =   1
                     X1              =   0
                     X2              =   7440
                     Y1              =   870
                     Y2              =   870
                  End
                  Begin VB.Line Line3 
                     Index           =   0
                     X1              =   30
                     X2              =   7470
                     Y1              =   480
                     Y2              =   480
                  End
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   195
                  Index           =   9
                  Left            =   30
                  Top             =   30
                  Width           =   7440
                  _ExtentX        =   13123
                  _ExtentY        =   344
                  Caption         =   "高温拉伸试验"
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
            End
            Begin Threed.SSPanel SSPanel11 
               Height          =   7575
               Left            =   -67680
               TabIndex        =   39
               Top             =   480
               Width           =   7335
               _ExtentX        =   12938
               _ExtentY        =   13361
               _Version        =   196609
               BevelInner      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.TextBox TXT_A_TIM_IMPACT_SIZE_CD 
                  Height          =   315
                  Left            =   1290
                  MaxLength       =   1
                  TabIndex        =   51
                  Tag             =   "27"
                  Top             =   5760
                  Width           =   405
               End
               Begin VB.TextBox TXT_TIM_IMPACT_SIZE_CD 
                  Height          =   315
                  Left            =   1290
                  MaxLength       =   1
                  TabIndex        =   50
                  Tag             =   "26"
                  Top             =   2010
                  Width           =   405
               End
               Begin VB.ComboBox Cob_A_TIM_IMPACT_SIZE 
                  Height          =   300
                  ItemData        =   "AQC0034C.frx":00E2
                  Left            =   1710
                  List            =   "AQC0034C.frx":00F2
                  Locked          =   -1  'True
                  TabIndex        =   49
                  Tag             =   "27"
                  Top             =   5760
                  Width           =   1935
               End
               Begin VB.ComboBox Cob_TIM_IMPACT_SIZE 
                  Height          =   300
                  ItemData        =   "AQC0034C.frx":0116
                  Left            =   1710
                  List            =   "AQC0034C.frx":0126
                  Locked          =   -1  'True
                  TabIndex        =   48
                  Tag             =   "26"
                  Top             =   2010
                  Width           =   1905
               End
               Begin VB.TextBox txt_A_TIM_IMPACT_DIR_NAME 
                  Enabled         =   0   'False
                  Height          =   300
                  Left            =   5460
                  TabIndex        =   47
                  Top             =   5430
                  Width           =   1695
               End
               Begin VB.TextBox txt_A_TIM_IMPACT_KND_NAME 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1710
                  TabIndex        =   46
                  Top             =   5430
                  Width           =   1905
               End
               Begin VB.TextBox txt_A_TIM_IMPACT_KND 
                  Height          =   315
                  Left            =   1290
                  MaxLength       =   1
                  TabIndex        =   45
                  Tag             =   "26"
                  Top             =   5430
                  Width           =   405
               End
               Begin VB.TextBox txt_A_TIM_IMPACT_DIR 
                  Height          =   300
                  Left            =   5040
                  MaxLength       =   1
                  TabIndex        =   44
                  Tag             =   "26"
                  Top             =   5430
                  Width           =   405
               End
               Begin VB.TextBox txt_TIM_IMPACT_DIR 
                  Height          =   315
                  Left            =   5040
                  MaxLength       =   1
                  TabIndex        =   43
                  Tag             =   "26"
                  Top             =   1620
                  Width           =   405
               End
               Begin VB.TextBox txt_TIM_IMPACT_KND 
                  Height          =   315
                  Left            =   1290
                  MaxLength       =   1
                  TabIndex        =   42
                  Tag             =   "26"
                  Top             =   1620
                  Width           =   405
               End
               Begin VB.TextBox txt_TIM_IMPACT_KND_NAME 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1710
                  TabIndex        =   41
                  Top             =   1620
                  Width           =   1875
               End
               Begin VB.TextBox txt_TIM_IMPACT_DIR_NAME 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   5460
                  TabIndex        =   40
                  Top             =   1620
                  Width           =   1695
               End
               Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST1 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   52
                  Tag             =   "26"
                  Top             =   3420
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST2 
                  Height          =   315
                  Left            =   930
                  TabIndex        =   53
                  Tag             =   "26"
                  Top             =   3420
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST3 
                  Height          =   315
                  Left            =   1740
                  TabIndex        =   54
                  Tag             =   "26"
                  Top             =   3420
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST4 
                  Height          =   315
                  Left            =   2550
                  TabIndex        =   55
                  Tag             =   "26"
                  Top             =   3420
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST5 
                  Height          =   315
                  Left            =   3360
                  TabIndex        =   56
                  Tag             =   "26"
                  Top             =   3420
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST6 
                  Height          =   315
                  Left            =   4170
                  TabIndex        =   57
                  Tag             =   "26"
                  Top             =   3420
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RST_AVE 
                  Height          =   315
                  Left            =   4980
                  TabIndex        =   58
                  Tag             =   "26"
                  Top             =   3420
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RATE_RST 
                  Height          =   315
                  Left            =   5790
                  TabIndex        =   59
                  Tag             =   "26"
                  Top             =   3420
                  Width           =   1350
                  _Version        =   262145
                  _ExtentX        =   2381
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin InDate.ULabel ULabel4 
                  Height          =   300
                  Index           =   4
                  Left            =   60
                  Top             =   1230
                  Width           =   7200
                  _ExtentX        =   12700
                  _ExtentY        =   529
                  Caption         =   "时效冲击试验"
                  Alignment       =   1
                  BackColor       =   16761024
                  BackgroundStyle =   1
                  BorderEffect    =   0
                  BorderStyle     =   1
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
                  ForeColor       =   0
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   27
                  Left            =   3840
                  Top             =   1620
                  Width           =   1140
                  _ExtentX        =   2011
                  _ExtentY        =   556
                  Caption         =   "试样方向"
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
                  Index           =   28
                  Left            =   120
                  Top             =   1620
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   556
                  Caption         =   "类别"
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
                  Index           =   29
                  Left            =   120
                  Top             =   2760
                  Width           =   5640
                  _ExtentX        =   9948
                  _ExtentY        =   556
                  Caption         =   "试验实绩（J）"
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
                  Index           =   30
                  Left            =   120
                  Top             =   2430
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "下限"
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
               Begin InDate.ULabel ul_TIM_MIN 
                  Height          =   315
                  Left            =   930
                  Top             =   2430
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   32
                  Left            =   1740
                  Top             =   2430
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   556
                  Caption         =   "最小下限"
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
               Begin InDate.ULabel ul_TIM_MIN_MIN 
                  Height          =   315
                  Left            =   3360
                  Top             =   2430
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   35
                  Left            =   4170
                  Top             =   2430
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
               Begin InDate.ULabel ul_TIM_AVE 
                  Height          =   315
                  Left            =   4980
                  Top             =   2430
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   645
                  Index           =   37
                  Left            =   5790
                  Top             =   2760
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   1138
                  Caption         =   "断面纤维率(%)"
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
               Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST1 
                  Height          =   315
                  Left            =   150
                  TabIndex        =   60
                  Tag             =   "27"
                  Top             =   7170
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST2 
                  Height          =   315
                  Left            =   960
                  TabIndex        =   61
                  Tag             =   "27"
                  Top             =   7170
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST3 
                  Height          =   315
                  Left            =   1770
                  TabIndex        =   62
                  Tag             =   "27"
                  Top             =   7170
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST4 
                  Height          =   315
                  Left            =   2580
                  TabIndex        =   63
                  Tag             =   "27"
                  Top             =   7170
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST5 
                  Height          =   315
                  Left            =   3390
                  TabIndex        =   64
                  Tag             =   "27"
                  Top             =   7170
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST6 
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   65
                  Tag             =   "27"
                  Top             =   7170
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RST_AVE 
                  Height          =   315
                  Left            =   5010
                  TabIndex        =   66
                  Tag             =   "27"
                  Top             =   7170
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RATE_RST 
                  Height          =   315
                  Left            =   5820
                  TabIndex        =   67
                  Tag             =   "27"
                  Top             =   7170
                  Width           =   1350
                  _Version        =   262145
                  _ExtentX        =   2381
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin InDate.ULabel ULabel4 
                  Height          =   300
                  Index           =   5
                  Left            =   60
                  Top             =   5010
                  Width           =   7200
                  _ExtentX        =   12700
                  _ExtentY        =   529
                  Caption         =   "追加时效冲击试验"
                  Alignment       =   1
                  BackColor       =   16761024
                  BackgroundStyle =   1
                  BorderEffect    =   0
                  BorderStyle     =   1
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
                  ForeColor       =   0
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   49
                  Left            =   3840
                  Top             =   5430
                  Width           =   1140
                  _ExtentX        =   2011
                  _ExtentY        =   556
                  Caption         =   "试样方向"
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
                  Index           =   50
                  Left            =   120
                  Top             =   5430
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   556
                  Caption         =   "类别"
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
                  Index           =   51
                  Left            =   150
                  Top             =   6510
                  Width           =   5640
                  _ExtentX        =   9948
                  _ExtentY        =   556
                  Caption         =   "试验实绩（J）"
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
                  Index           =   52
                  Left            =   150
                  Top             =   6180
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "下限"
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
               Begin InDate.ULabel ul_A_TIM_MIN 
                  Height          =   315
                  Left            =   960
                  Top             =   6180
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   54
                  Left            =   1770
                  Top             =   6180
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   556
                  Caption         =   "最小下限"
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
               Begin InDate.ULabel ul_A_TIM_MIN_MIN 
                  Height          =   315
                  Left            =   3390
                  Top             =   6180
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ul_A_TIM_AVE 
                  Height          =   315
                  Left            =   5010
                  Top             =   6180
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   58
                  Left            =   4200
                  Top             =   6180
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
                  Height          =   645
                  Index           =   59
                  Left            =   5820
                  Top             =   6510
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   1138
                  Caption         =   "断面纤维率(%)"
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
                  Index           =   63
                  Left            =   120
                  Top             =   2010
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   556
                  Caption         =   "试片尺寸"
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
                  Index           =   65
                  Left            =   120
                  Top             =   5760
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   556
                  Caption         =   "试片尺寸"
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
                  Index           =   107
                  Left            =   120
                  Top             =   3090
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "1"
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
                  Index           =   108
                  Left            =   930
                  Top             =   3090
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "2"
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
                  Index           =   109
                  Left            =   1740
                  Top             =   3090
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "3"
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
                  Index           =   110
                  Left            =   2550
                  Top             =   3090
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "4"
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
                  Index           =   111
                  Left            =   3360
                  Top             =   3090
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "5"
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
                  Index           =   112
                  Left            =   4170
                  Top             =   3090
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "6"
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
                  Index           =   113
                  Left            =   4980
                  Top             =   3090
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
               Begin InDate.ULabel ul_TIM_RATE 
                  Height          =   315
                  Left            =   5790
                  Top             =   2430
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   115
                  Left            =   150
                  Top             =   6840
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "1"
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
                  Index           =   116
                  Left            =   960
                  Top             =   6840
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "2"
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
                  Index           =   117
                  Left            =   1770
                  Top             =   6840
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "3"
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
                  Index           =   118
                  Left            =   2580
                  Top             =   6840
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "4"
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
                  Index           =   119
                  Left            =   3390
                  Top             =   6840
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "5"
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
                  Index           =   120
                  Left            =   4200
                  Top             =   6840
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "6"
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
                  Index           =   121
                  Left            =   5010
                  Top             =   6840
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
               Begin InDate.ULabel ul_A_TIM_RATE 
                  Height          =   315
                  Left            =   5820
                  Top             =   6180
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   53
                  Left            =   120
                  Top             =   120
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   556
                  Caption         =   "侧向膨胀值(%)"
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
                  Index           =   55
                  Left            =   1560
                  Top             =   480
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "1"
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
                  Index           =   56
                  Left            =   2370
                  Top             =   480
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "2"
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
                  Index           =   57
                  Left            =   3180
                  Top             =   480
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "3"
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
                  Index           =   68
                  Left            =   3990
                  Top             =   480
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "4"
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
                  Index           =   69
                  Left            =   4800
                  Top             =   480
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "5"
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
                  Index           =   70
                  Left            =   5610
                  Top             =   480
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "6"
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
                  Index           =   71
                  Left            =   6420
                  Top             =   480
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   155
                  Tag             =   "24"
                  Top             =   840
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   1
                  Left            =   2400
                  TabIndex        =   156
                  Tag             =   "24"
                  Top             =   840
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   3
                  Left            =   3960
                  TabIndex        =   157
                  Tag             =   "24"
                  Top             =   840
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   4
                  Left            =   4800
                  TabIndex        =   158
                  Tag             =   "24"
                  Top             =   840
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   5
                  Left            =   5640
                  TabIndex        =   159
                  Tag             =   "24"
                  Top             =   840
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   6
                  Left            =   6480
                  TabIndex        =   160
                  Tag             =   "24"
                  Top             =   840
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   80
                  Left            =   1560
                  Top             =   120
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "下限"
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
                  Index           =   82
                  Left            =   5640
                  Top             =   120
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
                  Index           =   83
                  Left            =   2400
                  Top             =   120
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  Index           =   84
                  Left            =   6480
                  Top             =   120
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   2
                  Left            =   3180
                  TabIndex        =   161
                  Tag             =   "24"
                  Top             =   840
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   85
                  Left            =   0
                  Top             =   3960
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   556
                  Caption         =   "侧向膨胀值(%)"
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
                  Index           =   87
                  Left            =   1440
                  Top             =   4320
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "1"
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
                  Index           =   89
                  Left            =   2250
                  Top             =   4320
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "2"
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
                  Index           =   90
                  Left            =   3060
                  Top             =   4320
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "3"
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
                  Index           =   92
                  Left            =   3870
                  Top             =   4320
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "4"
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
                  Index           =   101
                  Left            =   4680
                  Top             =   4320
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "5"
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
                  Index           =   103
                  Left            =   5490
                  Top             =   4320
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "6"
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
                  Index           =   104
                  Left            =   6300
                  Top             =   4320
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
                  Index           =   105
                  Left            =   1440
                  Top             =   3960
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "下限"
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
                  Index           =   106
                  Left            =   5520
                  Top             =   3960
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
                  Index           =   114
                  Left            =   2280
                  Top             =   3960
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  Index           =   122
                  Left            =   6360
                  Top             =   3960
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   7
                  Left            =   1440
                  TabIndex        =   162
                  Tag             =   "24"
                  Top             =   4680
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   8
                  Left            =   2250
                  TabIndex        =   163
                  Tag             =   "24"
                  Top             =   4680
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   9
                  Left            =   3060
                  TabIndex        =   164
                  Tag             =   "24"
                  Top             =   4680
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   10
                  Left            =   3870
                  TabIndex        =   165
                  Tag             =   "24"
                  Top             =   4680
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   11
                  Left            =   4680
                  TabIndex        =   166
                  Tag             =   "24"
                  Top             =   4680
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   12
                  Left            =   5490
                  TabIndex        =   167
                  Tag             =   "24"
                  Top             =   4680
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_EXPAIN_RST 
                  Height          =   315
                  Index           =   13
                  Left            =   6300
                  TabIndex        =   168
                  Tag             =   "24"
                  Top             =   4680
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin VB.Shape Shape3 
                  Height          =   2205
                  Left            =   60
                  Top             =   5340
                  Width           =   7185
               End
               Begin VB.Shape Shape2 
                  Height          =   2205
                  Left            =   60
                  Top             =   1560
                  Width           =   7185
               End
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   7575
               Index           =   0
               Left            =   -75000
               TabIndex        =   68
               Top             =   705
               Width           =   7425
               _ExtentX        =   13097
               _ExtentY        =   13361
               _Version        =   196609
               BevelInner      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.TextBox TXT_A_IMPACT_SIZE_CD 
                  Height          =   315
                  Left            =   1320
                  MaxLength       =   1
                  TabIndex        =   80
                  Tag             =   "25"
                  Top             =   4740
                  Width           =   405
               End
               Begin VB.TextBox TXT_IMPACT_SIZE_CD 
                  Height          =   315
                  Left            =   1350
                  MaxLength       =   1
                  TabIndex        =   79
                  Tag             =   "24"
                  Top             =   840
                  Width           =   405
               End
               Begin VB.ComboBox Cob_A_IMPACT_SIZE 
                  Height          =   300
                  ItemData        =   "AQC0034C.frx":014A
                  Left            =   1770
                  List            =   "AQC0034C.frx":015A
                  Locked          =   -1  'True
                  TabIndex        =   78
                  Tag             =   "25"
                  Top             =   4740
                  Width           =   1845
               End
               Begin VB.ComboBox Cob_IMPACT_SIZE 
                  Height          =   300
                  ItemData        =   "AQC0034C.frx":017E
                  Left            =   1770
                  List            =   "AQC0034C.frx":018E
                  Locked          =   -1  'True
                  TabIndex        =   77
                  Tag             =   "24"
                  Top             =   840
                  Width           =   1845
               End
               Begin VB.TextBox txt_A_IMPACT_DIR 
                  Height          =   300
                  Left            =   5070
                  MaxLength       =   1
                  TabIndex        =   76
                  Tag             =   "25"
                  Top             =   4320
                  Width           =   405
               End
               Begin VB.TextBox txt_A_IMPACT_KND 
                  Height          =   300
                  Left            =   1350
                  MaxLength       =   1
                  TabIndex        =   75
                  Tag             =   "25"
                  Top             =   4320
                  Width           =   405
               End
               Begin VB.TextBox txt_A_IMPACT_KND_NAME 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1770
                  TabIndex        =   74
                  Top             =   4320
                  Width           =   1845
               End
               Begin VB.TextBox txt_A_IMPACT_DIR_NAME 
                  Enabled         =   0   'False
                  Height          =   300
                  Left            =   5490
                  TabIndex        =   73
                  Top             =   4320
                  Width           =   1695
               End
               Begin VB.TextBox txt_IMPACT_DIR_NAME 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   5490
                  TabIndex        =   72
                  Top             =   420
                  Width           =   1695
               End
               Begin VB.TextBox txt_IMPACT_KND_NAME 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1770
                  TabIndex        =   71
                  Top             =   420
                  Width           =   1845
               End
               Begin VB.TextBox txt_IMPACT_KND 
                  Height          =   315
                  Left            =   1350
                  MaxLength       =   1
                  TabIndex        =   70
                  Tag             =   "24"
                  Top             =   420
                  Width           =   405
               End
               Begin VB.TextBox txt_IMPACT_DIR 
                  Height          =   315
                  Left            =   5070
                  MaxLength       =   1
                  TabIndex        =   69
                  Tag             =   "24"
                  Top             =   420
                  Width           =   405
               End
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST1 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   81
                  Tag             =   "24"
                  Top             =   1905
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST2 
                  Height          =   315
                  Left            =   2370
                  TabIndex        =   82
                  Tag             =   "24"
                  Top             =   1905
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST3 
                  Height          =   315
                  Left            =   3180
                  TabIndex        =   83
                  Tag             =   "24"
                  Top             =   1905
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST4 
                  Height          =   315
                  Left            =   3990
                  TabIndex        =   84
                  Tag             =   "24"
                  Top             =   1905
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST5 
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   85
                  Tag             =   "24"
                  Top             =   1905
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST6 
                  Height          =   315
                  Left            =   5610
                  TabIndex        =   86
                  Tag             =   "24"
                  Top             =   1905
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RST_AVE 
                  Height          =   315
                  Left            =   6420
                  TabIndex        =   87
                  Tag             =   "24"
                  Top             =   1905
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin InDate.ULabel ULabel4 
                  Height          =   300
                  Index           =   1
                  Left            =   120
                  Top             =   0
                  Width           =   7200
                  _ExtentX        =   12700
                  _ExtentY        =   529
                  Caption         =   "冲击试验"
                  Alignment       =   1
                  BackColor       =   16761024
                  BackgroundStyle =   1
                  BorderEffect    =   0
                  BorderStyle     =   1
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
                  ForeColor       =   0
               End
               Begin InDate.ULabel ULabel4 
                  Height          =   300
                  Index           =   2
                  Left            =   120
                  Top             =   3960
                  Width           =   7200
                  _ExtentX        =   12700
                  _ExtentY        =   529
                  Caption         =   "追加冲击试验"
                  Alignment       =   1
                  BackColor       =   16761024
                  BackgroundStyle =   1
                  BorderEffect    =   0
                  BorderStyle     =   1
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
                  ForeColor       =   0
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   16
                  Left            =   3870
                  Top             =   420
                  Width           =   1140
                  _ExtentX        =   2011
                  _ExtentY        =   556
                  Caption         =   "试样方向"
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
                  Index           =   17
                  Left            =   195
                  Top             =   420
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   556
                  Caption         =   "类别"
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
                  Index           =   18
                  Left            =   195
                  Top             =   1575
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   556
                  Caption         =   "试验实绩（J）"
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
                  Left            =   1560
                  Top             =   1575
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "1"
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
                  Index           =   20
                  Left            =   2370
                  Top             =   1575
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "2"
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
                  Index           =   21
                  Left            =   3180
                  Top             =   1575
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "3"
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
                  Index           =   22
                  Left            =   3990
                  Top             =   1575
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "4"
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
                  Index           =   23
                  Left            =   4800
                  Top             =   1575
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "5"
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
                  Index           =   24
                  Left            =   5610
                  Top             =   1575
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "6"
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
                  Index           =   25
                  Left            =   6420
                  Top             =   1575
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST1 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   88
                  Tag             =   "25"
                  Top             =   5850
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST2 
                  Height          =   315
                  Left            =   2370
                  TabIndex        =   89
                  Tag             =   "25"
                  Top             =   5850
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST3 
                  Height          =   315
                  Left            =   3180
                  TabIndex        =   90
                  Tag             =   "25"
                  Top             =   5850
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST4 
                  Height          =   315
                  Left            =   3990
                  TabIndex        =   91
                  Tag             =   "25"
                  Top             =   5850
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST5 
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   92
                  Tag             =   "25"
                  Top             =   5850
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST6 
                  Height          =   315
                  Left            =   5610
                  TabIndex        =   93
                  Tag             =   "25"
                  Top             =   5850
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RST_AVE 
                  Height          =   315
                  Left            =   6420
                  TabIndex        =   94
                  Tag             =   "25"
                  Top             =   5850
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   300
                  Index           =   38
                  Left            =   3870
                  Top             =   4320
                  Width           =   1140
                  _ExtentX        =   2011
                  _ExtentY        =   529
                  Caption         =   "试样方向"
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
                  Height          =   300
                  Index           =   39
                  Left            =   195
                  Top             =   4320
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   529
                  Caption         =   "类别"
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
                  Index           =   40
                  Left            =   195
                  Top             =   5520
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  Caption         =   "试验实绩（J）"
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
                  Index           =   41
                  Left            =   1560
                  Top             =   5520
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "1"
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
                  Index           =   42
                  Left            =   2370
                  Top             =   5520
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "2"
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
                  Index           =   43
                  Left            =   3180
                  Top             =   5520
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "3"
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
                  Index           =   44
                  Left            =   3990
                  Top             =   5520
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "4"
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
                  Index           =   45
                  Left            =   4800
                  Top             =   5520
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "5"
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
                  Index           =   46
                  Left            =   5610
                  Top             =   5520
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "6"
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
                  Index           =   47
                  Left            =   6420
                  Top             =   5520
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
                  Index           =   48
                  Left            =   195
                  Top             =   6600
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  Caption         =   "断面纤维率(%)"
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
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_RST1 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   95
                  Tag             =   "24"
                  Top             =   3210
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_RST2 
                  Height          =   315
                  Left            =   2370
                  TabIndex        =   96
                  Tag             =   "24"
                  Top             =   3210
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_RST3 
                  Height          =   315
                  Left            =   3180
                  TabIndex        =   97
                  Tag             =   "24"
                  Top             =   3210
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_RST4 
                  Height          =   315
                  Left            =   3990
                  TabIndex        =   98
                  Tag             =   "24"
                  Top             =   3210
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_RST5 
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   99
                  Tag             =   "24"
                  Top             =   3210
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_RST6 
                  Height          =   315
                  Left            =   5610
                  TabIndex        =   100
                  Tag             =   "24"
                  Top             =   3210
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_AVE_RST 
                  Height          =   315
                  Left            =   6420
                  TabIndex        =   101
                  Tag             =   "24"
                  Top             =   3210
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_RST1 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   102
                  Tag             =   "25"
                  Top             =   6930
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_RST2 
                  Height          =   315
                  Left            =   2370
                  TabIndex        =   103
                  Tag             =   "25"
                  Top             =   6930
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_RST3 
                  Height          =   315
                  Left            =   3180
                  TabIndex        =   104
                  Tag             =   "25"
                  Top             =   6930
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_RST4 
                  Height          =   315
                  Left            =   3990
                  TabIndex        =   105
                  Tag             =   "25"
                  Top             =   6930
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_RST5 
                  Height          =   315
                  Left            =   4815
                  TabIndex        =   106
                  Tag             =   "25"
                  Top             =   6930
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_RST6 
                  Height          =   315
                  Left            =   5610
                  TabIndex        =   107
                  Tag             =   "25"
                  Top             =   6930
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_AVE_RST 
                  Height          =   315
                  Left            =   6420
                  TabIndex        =   108
                  Tag             =   "25"
                  Top             =   6930
                  Width           =   780
                  _Version        =   262145
                  _ExtentX        =   1376
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
                  NumDecDigits    =   2
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   62
                  Left            =   195
                  Top             =   840
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   556
                  Caption         =   "试片尺寸"
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
                  Height          =   300
                  Index           =   64
                  Left            =   195
                  Top             =   4740
                  Width           =   1110
                  _ExtentX        =   1958
                  _ExtentY        =   529
                  Caption         =   "试片尺寸"
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
                  Index           =   60
                  Left            =   1560
                  Top             =   1245
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "下限"
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
               Begin InDate.ULabel ul_IMPACT_MIN 
                  Height          =   315
                  Left            =   2370
                  Top             =   1245
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   67
                  Left            =   3180
                  Top             =   1245
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   556
                  Caption         =   "最小下限"
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
               Begin InDate.ULabel ul_IMPACT_MIN_MIN 
                  Height          =   315
                  Left            =   4800
                  Top             =   1245
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   66
                  Left            =   5610
                  Top             =   1245
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
               Begin InDate.ULabel ul_IMPACT_AVE 
                  Height          =   315
                  Left            =   6420
                  Top             =   1245
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   72
                  Left            =   1560
                  Top             =   2880
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "1"
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
                  Index           =   73
                  Left            =   2370
                  Top             =   2880
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "2"
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
                  Index           =   74
                  Left            =   3180
                  Top             =   2880
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "3"
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
                  Index           =   75
                  Left            =   3990
                  Top             =   2880
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "4"
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
                  Index           =   76
                  Left            =   4800
                  Top             =   2880
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "5"
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
                  Index           =   77
                  Left            =   5610
                  Top             =   2880
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "6"
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
                  Index           =   78
                  Left            =   6420
                  Top             =   2880
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
                  Index           =   86
                  Left            =   1560
                  Top             =   5160
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "下限"
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
               Begin InDate.ULabel ul_IMPACT_A_MIN 
                  Height          =   315
                  Left            =   2370
                  Top             =   5160
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   88
                  Left            =   3180
                  Top             =   5160
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   556
                  Caption         =   "最小下限"
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
               Begin InDate.ULabel ul_IMPACT_A_MIN_MIN 
                  Height          =   315
                  Left            =   4800
                  Top             =   5160
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   91
                  Left            =   5610
                  Top             =   5160
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
               Begin InDate.ULabel ul_IMPACT_A_AVE 
                  Height          =   315
                  Left            =   6420
                  Top             =   5160
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   93
                  Left            =   1560
                  Top             =   6600
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "1"
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
                  Index           =   94
                  Left            =   2370
                  Top             =   6600
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "2"
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
                  Index           =   95
                  Left            =   3180
                  Top             =   6600
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "3"
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
                  Index           =   96
                  Left            =   3990
                  Top             =   6600
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "4"
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
                  Index           =   97
                  Left            =   4800
                  Top             =   6600
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "5"
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
                  Index           =   98
                  Left            =   5610
                  Top             =   6600
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "6"
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
                  Index           =   99
                  Left            =   6420
                  Top             =   6600
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "平均值"
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
                  Index           =   100
                  Left            =   1560
                  Top             =   6240
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "下限"
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
               Begin InDate.ULabel ul_IMPACT_A_RATE_MIN 
                  Height          =   315
                  Left            =   2400
                  Top             =   6240
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   102
                  Left            =   3180
                  Top             =   6240
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "上限"
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
               Begin InDate.ULabel ul_IMPACT_A_RATE_MAX 
                  Height          =   315
                  Left            =   3990
                  Top             =   6240
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   26
                  Left            =   195
                  Top             =   2880
                  Width           =   1350
                  _ExtentX        =   2381
                  _ExtentY        =   556
                  Caption         =   "断面纤维率(%)"
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
                  Index           =   79
                  Left            =   1560
                  Top             =   2520
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "下限"
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
               Begin InDate.ULabel ul_IMPACT_RATE_MIN 
                  Height          =   315
                  Left            =   2370
                  Top             =   2520
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   81
                  Left            =   3180
                  Top             =   2520
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   "上限"
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
               Begin InDate.ULabel ul_IMPACT_RATE_MAX 
                  Height          =   315
                  Left            =   3990
                  Top             =   2520
                  Width           =   780
                  _ExtentX        =   1376
                  _ExtentY        =   556
                  Caption         =   ""
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
                  ForeColor       =   255
               End
               Begin VB.Shape Shape1 
                  Height          =   3765
                  Index           =   2
                  Left            =   120
                  Top             =   0
                  Width           =   7185
               End
               Begin VB.Shape Shape1 
                  Height          =   3645
                  Index           =   1
                  Left            =   120
                  Top             =   3900
                  Width           =   7185
               End
               Begin VB.Line Line7 
                  BorderStyle     =   5  'Dash-Dot-Dot
                  Index           =   0
                  X1              =   120
                  X2              =   7305
                  Y1              =   2490
                  Y2              =   2490
               End
               Begin VB.Line Line7 
                  BorderStyle     =   5  'Dash-Dot-Dot
                  Index           =   1
                  X1              =   120
                  X2              =   7290
                  Y1              =   1920
                  Y2              =   1920
               End
               Begin VB.Line Line7 
                  BorderStyle     =   5  'Dash-Dot-Dot
                  Index           =   2
                  X1              =   120
                  X2              =   7290
                  Y1              =   5100
                  Y2              =   5100
               End
               Begin VB.Line Line7 
                  BorderStyle     =   5  'Dash-Dot-Dot
                  Index           =   3
                  X1              =   120
                  X2              =   7290
                  Y1              =   6210
                  Y2              =   6210
               End
               Begin VB.Shape Shape1 
                  Height          =   3765
                  Index           =   0
                  Left            =   120
                  Top             =   120
                  Width           =   7185
               End
            End
            Begin Threed.SSPanel SSPanel6 
               Height          =   4335
               Index           =   0
               Left            =   -74880
               TabIndex        =   109
               Top             =   360
               Width           =   14595
               _ExtentX        =   25744
               _ExtentY        =   7646
               _Version        =   196609
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin InDate.ULabel ULabel1 
                  Height          =   195
                  Index           =   8
                  Left            =   30
                  Top             =   30
                  Width           =   7080
                  _ExtentX        =   12488
                  _ExtentY        =   344
                  Caption         =   "拉伸试验"
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
               Begin CSTextLibCtl.sidbEdit ZRA_RST_4 
                  Height          =   315
                  Index           =   2
                  Left            =   4920
                  TabIndex        =   206
                  Tag             =   "51"
                  Top             =   3600
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumDecDigits    =   1
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit ZRA_RST_5 
                  Height          =   315
                  Index           =   2
                  Left            =   5760
                  TabIndex        =   207
                  Tag             =   "51"
                  Top             =   3600
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumDecDigits    =   1
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit ZRA_RST_6 
                  Height          =   315
                  Index           =   2
                  Left            =   6600
                  TabIndex        =   208
                  Tag             =   "51"
                  Top             =   3600
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumDecDigits    =   1
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin Threed.SSPanel SSPanel7 
                  Height          =   4275
                  Index           =   0
                  Left            =   0
                  TabIndex        =   110
                  Top             =   240
                  Width           =   14535
                  _ExtentX        =   25638
                  _ExtentY        =   7541
                  _Version        =   196609
                  BevelOuter      =   1
                  RoundedCorners  =   0   'False
                  FloodShowPct    =   -1  'True
                  Begin VB.TextBox txt_HARD_NAME 
                     Height          =   315
                     Index           =   0
                     Left            =   11880
                     Locked          =   -1  'True
                     TabIndex        =   233
                     Top             =   3360
                     Width           =   765
                  End
                  Begin VB.TextBox txt_HARD_TYP 
                     Height          =   315
                     Index           =   0
                     Left            =   11280
                     TabIndex        =   232
                     Tag             =   "14"
                     Top             =   3360
                     Width           =   600
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   0
                     Left            =   60
                     Top             =   60
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "屈服强度   YP   MPa"
                     Alignment       =   0
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
                  Begin CSTextLibCtl.sidbEdit sdb_YP_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   2400
                     TabIndex        =   111
                     Tag             =   "1"
                     Top             =   60
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_TS_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   2400
                     TabIndex        =   112
                     Tag             =   "2"
                     Top             =   840
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_EL_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   2400
                     TabIndex        =   113
                     Tag             =   "4"
                     Top             =   1230
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_SG_EL_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   2400
                     TabIndex        =   114
                     Tag             =   "7"
                     Top             =   492
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_YR_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   2400
                     TabIndex        =   115
                     Tag             =   "35"
                     Top             =   2160
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_SNPP_EL_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   2400
                     TabIndex        =   116
                     Tag             =   "5"
                     Top             =   2520
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_SP_EL_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   2400
                     TabIndex        =   117
                     Tag             =   "6"
                     Top             =   2880
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   1
                     Left            =   60
                     Top             =   492
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "规定总伸长应力   MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   2
                     Left            =   60
                     Top             =   840
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "抗拉强度   TS   MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   3
                     Left            =   60
                     Top             =   1230
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "断后伸长率   EL   %"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   5
                     Left            =   60
                     Top             =   2160
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "屈强比   Y.S/T.S   %"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   6
                     Left            =   60
                     Top             =   2520
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "规定非比例伸长应力MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   7
                     Left            =   60
                     Top             =   2880
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "规定残余伸长应力  MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ul_YP 
                     Height          =   315
                     Index           =   0
                     Left            =   5970
                     Top             =   60
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_SG_EL 
                     Height          =   315
                     Index           =   0
                     Left            =   5970
                     Top             =   480
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_TS 
                     Height          =   315
                     Index           =   0
                     Left            =   5970
                     Top             =   840
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_EL 
                     Height          =   315
                     Index           =   0
                     Left            =   5970
                     Top             =   1230
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_RA 
                     Height          =   315
                     Index           =   0
                     Left            =   5970
                     Top             =   1665
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_YR 
                     Height          =   315
                     Index           =   0
                     Left            =   5970
                     Top             =   2160
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_SNPP_EL 
                     Height          =   315
                     Index           =   0
                     Left            =   5970
                     Top             =   2520
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_SP_EL 
                     Height          =   315
                     Index           =   0
                     Left            =   5970
                     Top             =   2880
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   33
                     Left            =   60
                     Top             =   3360
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "厚度方向断面收缩率    RA   %"
                     Alignment       =   0
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
                  Begin CSTextLibCtl.sidbEdit ZRA_RST_1 
                     Height          =   315
                     Index           =   2
                     Left            =   2400
                     TabIndex        =   150
                     Tag             =   "51"
                     Top             =   3360
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit ZRA_RST_2 
                     Height          =   315
                     Index           =   2
                     Left            =   3240
                     TabIndex        =   151
                     Tag             =   "51"
                     Top             =   3360
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ul_RA 
                     Height          =   315
                     Index           =   2
                     Left            =   8280
                     Top             =   3360
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin CSTextLibCtl.sidbEdit ZRA_RST_3 
                     Height          =   315
                     Index           =   2
                     Left            =   4080
                     TabIndex        =   152
                     Tag             =   "51"
                     Top             =   3360
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit ZRA_RST_AVE 
                     Height          =   315
                     Index           =   2
                     Left            =   7440
                     TabIndex        =   153
                     Tag             =   "51"
                     Top             =   3360
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   34
                     Left            =   60
                     Top             =   3720
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "厚度方向抗拉强度   TS   MPa"
                     Alignment       =   0
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
                  Begin CSTextLibCtl.sidbEdit TS_RST_1 
                     Height          =   315
                     Index           =   2
                     Left            =   2400
                     TabIndex        =   209
                     Tag             =   "61"
                     Top             =   3720
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit TS_RST_2 
                     Height          =   315
                     Index           =   2
                     Left            =   3240
                     TabIndex        =   210
                     Tag             =   "61"
                     Top             =   3720
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit TS_RST_3 
                     Height          =   315
                     Index           =   2
                     Left            =   4080
                     TabIndex        =   211
                     Tag             =   "61"
                     Top             =   3720
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ul_TS 
                     Height          =   315
                     Index           =   2
                     Left            =   8280
                     Top             =   3720
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_RA_RST_AVE 
                     Height          =   315
                     Index           =   0
                     Left            =   4920
                     TabIndex        =   212
                     Tag             =   "3"
                     Top             =   1680
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_RA_RST_3 
                     Height          =   315
                     Index           =   2
                     Left            =   4080
                     TabIndex        =   213
                     Tag             =   "3"
                     Top             =   1680
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_RA_RST_2 
                     Height          =   315
                     Index           =   2
                     Left            =   3240
                     TabIndex        =   214
                     Tag             =   "3"
                     Top             =   1680
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_RA_RST_1 
                     Height          =   315
                     Index           =   2
                     Left            =   2400
                     TabIndex        =   215
                     Tag             =   "3"
                     Top             =   1680
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   62
                     Left            =   60
                     Top             =   1680
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "断面收缩率    RA   %"
                     Alignment       =   0
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
                  Begin CSTextLibCtl.sidbEdit TS_RST_4 
                     Height          =   315
                     Index           =   2
                     Left            =   4920
                     TabIndex        =   226
                     Tag             =   "61"
                     Top             =   3720
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit TS_RST_5 
                     Height          =   315
                     Index           =   2
                     Left            =   5760
                     TabIndex        =   227
                     Tag             =   "61"
                     Top             =   3720
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit TS_RST_6 
                     Height          =   315
                     Index           =   2
                     Left            =   6600
                     TabIndex        =   228
                     Tag             =   "61"
                     Top             =   3720
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumDecDigits    =   1
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   49
                     Left            =   9600
                     Top             =   3360
                     Width           =   840
                     _ExtentX        =   1482
                     _ExtentY        =   556
                     Caption         =   "硬度试验"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   36
                     Left            =   10440
                     Top             =   3360
                     Width           =   840
                     _ExtentX        =   1482
                     _ExtentY        =   556
                     Caption         =   "硬度类型"
                     Alignment       =   0
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
                  Begin CSTextLibCtl.sidbEdit sdb_HARD_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   10440
                     TabIndex        =   234
                     Tag             =   "14"
                     Top             =   3720
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   37
                     Left            =   9600
                     Top             =   3720
                     Width           =   840
                     _ExtentX        =   1482
                     _ExtentY        =   556
                     Caption         =   "硬度值"
                     Alignment       =   0
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
                     ForeColor       =   0
                  End
                  Begin InDate.ULabel ul_HARD 
                     Height          =   315
                     Index           =   0
                     Left            =   13440
                     Top             =   3720
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin VB.Line Line5 
                     Index           =   2
                     X1              =   9360
                     X2              =   9360
                     Y1              =   2040
                     Y2              =   4440
                  End
                  Begin VB.Line Line12 
                     X1              =   120
                     X2              =   14640
                     Y1              =   3240
                     Y2              =   3240
                  End
                  Begin VB.Line Line1 
                     Index           =   0
                     X1              =   5880
                     X2              =   5880
                     Y1              =   30
                     Y2              =   3240
                  End
                  Begin VB.Line Line2 
                     Index           =   0
                     X1              =   0
                     X2              =   7110
                     Y1              =   420
                     Y2              =   420
                  End
                  Begin VB.Line Line2 
                     Index           =   3
                     X1              =   0
                     X2              =   7110
                     Y1              =   1590
                     Y2              =   1590
                  End
                  Begin VB.Line Line2 
                     Index           =   5
                     X1              =   0
                     X2              =   7110
                     Y1              =   2040
                     Y2              =   2040
                  End
               End
            End
            Begin Threed.SSPanel SSPanel4 
               Height          =   3615
               Index           =   1
               Left            =   7200
               TabIndex        =   131
               Top             =   360
               Width           =   7515
               _ExtentX        =   13256
               _ExtentY        =   6376
               _Version        =   196609
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin Threed.SSPanel SSPanel8 
                  Height          =   3225
                  Index           =   1
                  Left            =   0
                  TabIndex        =   132
                  Top             =   360
                  Width           =   7455
                  _ExtentX        =   13150
                  _ExtentY        =   5689
                  _Version        =   196609
                  BevelOuter      =   1
                  RoundedCorners  =   0   'False
                  FloodShowPct    =   -1  'True
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_YP_RST 
                     Height          =   315
                     Index           =   1
                     Left            =   2400
                     TabIndex        =   133
                     Tag             =   "43"
                     Top             =   60
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   4
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_TS_RST 
                     Height          =   315
                     Index           =   1
                     Left            =   2400
                     TabIndex        =   134
                     Tag             =   "44"
                     Top             =   600
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   4
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_EL_RST 
                     Height          =   315
                     Index           =   1
                     Left            =   2400
                     TabIndex        =   135
                     Tag             =   "46"
                     Top             =   1560
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   2
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_SNPP_EL_RST 
                     Height          =   315
                     Index           =   1
                     Left            =   2400
                     TabIndex        =   136
                     Tag             =   "47"
                     Top             =   1920
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   4
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                     Height          =   315
                     Index           =   1
                     Left            =   2400
                     TabIndex        =   137
                     Tag             =   "48"
                     Top             =   2280
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   4
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ul_H_YP 
                     Height          =   315
                     Index           =   1
                     Left            =   6360
                     Top             =   60
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   25
                     Left            =   60
                     Top             =   60
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "屈服强度   YP   MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   26
                     Left            =   60
                     Top             =   600
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "抗拉强度   TS   MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   27
                     Left            =   60
                     Top             =   1080
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "断面收缩率   RA    %"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   28
                     Left            =   60
                     Top             =   1560
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "断后伸长率   EL    %"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   29
                     Left            =   60
                     Top             =   1920
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "规定非比例伸长应力MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   30
                     Left            =   60
                     Top             =   2280
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "规定残余伸长应力  MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ul_H_TS 
                     Height          =   315
                     Index           =   1
                     Left            =   6360
                     Top             =   600
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_H_RA 
                     Height          =   315
                     Index           =   1
                     Left            =   6360
                     Top             =   1080
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_H_EL 
                     Height          =   315
                     Index           =   1
                     Left            =   6360
                     Top             =   1560
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_H_SNPP_EL 
                     Height          =   315
                     Index           =   1
                     Left            =   6360
                     Top             =   1920
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_H_SP_EL 
                     Height          =   315
                     Index           =   1
                     Left            =   6360
                     Top             =   2280
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_RST_1 
                     Height          =   315
                     Index           =   1
                     Left            =   2400
                     TabIndex        =   146
                     Tag             =   "45"
                     Top             =   1080
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_RST_A 
                     Height          =   315
                     Index           =   1
                     Left            =   4920
                     TabIndex        =   147
                     Tag             =   "45"
                     Top             =   1080
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_RST_3 
                     Height          =   315
                     Index           =   1
                     Left            =   4080
                     TabIndex        =   148
                     Tag             =   "45"
                     Top             =   1080
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_RST_2 
                     Height          =   315
                     Index           =   1
                     Left            =   3240
                     TabIndex        =   149
                     Tag             =   "45"
                     Top             =   1080
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   51
                     Left            =   60
                     Top             =   2760
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "均匀变形伸长率UEL %"
                     Alignment       =   0
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
                  Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                     Height          =   315
                     Index           =   3
                     Left            =   2400
                     TabIndex        =   170
                     Tag             =   "55"
                     Top             =   2760
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   2
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ul_H_SP_EL 
                     Height          =   315
                     Index           =   3
                     Left            =   6360
                     Top             =   2760
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin VB.Line Line3 
                     Index           =   9
                     X1              =   0
                     X2              =   7440
                     Y1              =   480
                     Y2              =   480
                  End
                  Begin VB.Line Line3 
                     Index           =   8
                     X1              =   0
                     X2              =   7440
                     Y1              =   960
                     Y2              =   960
                  End
                  Begin VB.Line Line3 
                     Index           =   7
                     X1              =   0
                     X2              =   7440
                     Y1              =   1440
                     Y2              =   1440
                  End
                  Begin VB.Line Line3 
                     Index           =   5
                     X1              =   0
                     X2              =   7440
                     Y1              =   2640
                     Y2              =   2640
                  End
                  Begin VB.Line Line4 
                     Index           =   1
                     X1              =   6240
                     X2              =   6240
                     Y1              =   30
                     Y2              =   3480
                  End
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   34
                  Left            =   30
                  Top             =   30
                  Width           =   7440
                  _ExtentX        =   13123
                  _ExtentY        =   556
                  Caption         =   "追加高温拉伸试验"
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
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   3135
               Index           =   1
               Left            =   120
               TabIndex        =   138
               Top             =   5280
               Width           =   14625
               _ExtentX        =   25797
               _ExtentY        =   5530
               _Version        =   196609
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.TextBox txt_HARD_NAME 
                  Height          =   315
                  Index           =   1
                  Left            =   4140
                  Locked          =   -1  'True
                  TabIndex        =   141
                  Tag             =   "49"
                  Top             =   60
                  Width           =   1965
               End
               Begin VB.TextBox txt_HARD_TYP 
                  Height          =   315
                  Index           =   1
                  Left            =   3300
                  TabIndex        =   140
                  Tag             =   "49"
                  Top             =   60
                  Width           =   840
               End
               Begin VB.TextBox txt_BEND_RST 
                  Height          =   315
                  Index           =   1
                  Left            =   2340
                  TabIndex        =   139
                  Tag             =   "50"
                  Top             =   462
                  Width           =   840
               End
               Begin CSTextLibCtl.sidbEdit sdb_HARD_RST 
                  Height          =   315
                  Index           =   1
                  Left            =   7170
                  TabIndex        =   142
                  Tag             =   "49"
                  Top             =   60
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin InDate.ULabel ul_HARD 
                  Height          =   315
                  Index           =   1
                  Left            =   13500
                  Top             =   60
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   556
                  Caption         =   ""
                  Alignment       =   0
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
                  ForeColor       =   192
               End
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   31
                  Left            =   2340
                  Top             =   60
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   556
                  Caption         =   "硬度类型"
                  Alignment       =   0
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
                  ForeColor       =   16711680
               End
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   32
                  Left            =   6240
                  Top             =   60
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   556
                  Caption         =   "硬度值"
                  Alignment       =   0
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
                  ForeColor       =   0
               End
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   52
                  Left            =   0
                  Top             =   465
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   556
                  Caption         =   "追加弯曲试验"
                  Alignment       =   0
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
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   58
                  Left            =   0
                  Top             =   60
                  Width           =   2160
                  _ExtentX        =   3810
                  _ExtentY        =   556
                  Caption         =   "追加硬度试验"
                  Alignment       =   0
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
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   61
                  Left            =   3300
                  Top             =   480
                  Width           =   1650
                  _ExtentX        =   2910
                  _ExtentY        =   556
                  Caption         =   "Y-合格；N-不合格"
                  Alignment       =   0
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
                  ForeColor       =   0
               End
               Begin InDate.ULabel ul_BEND 
                  Height          =   315
                  Index           =   1
                  Left            =   13500
                  Top             =   462
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   556
                  Caption         =   ""
                  Alignment       =   0
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
                  ForeColor       =   192
               End
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   53
                  Left            =   0
                  Top             =   840
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   556
                  Caption         =   "应力比项目1"
                  Alignment       =   0
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
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   54
                  Left            =   0
                  Top             =   2280
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   556
                  Caption         =   "应力比项目5"
                  Alignment       =   0
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
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   55
                  Left            =   0
                  Top             =   1920
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   556
                  Caption         =   "应力比项目4"
                  Alignment       =   0
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
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   56
                  Left            =   0
                  Top             =   1560
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   556
                  Caption         =   "应力比项目3"
                  Alignment       =   0
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
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   57
                  Left            =   0
                  Top             =   1200
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   556
                  Caption         =   "应力比项目2"
                  Alignment       =   0
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
               Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                  Height          =   315
                  Index           =   4
                  Left            =   2340
                  TabIndex        =   171
                  Tag             =   "56"
                  Top             =   840
                  Width           =   1200
                  _Version        =   262145
                  _ExtentX        =   2117
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
                  NumIntDigits    =   2
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                  Height          =   315
                  Index           =   5
                  Left            =   2340
                  TabIndex        =   172
                  Tag             =   "57"
                  Top             =   1200
                  Width           =   1200
                  _Version        =   262145
                  _ExtentX        =   2117
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
                  NumIntDigits    =   2
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                  Height          =   315
                  Index           =   6
                  Left            =   2340
                  TabIndex        =   173
                  Tag             =   "58"
                  Top             =   1560
                  Width           =   1200
                  _Version        =   262145
                  _ExtentX        =   2117
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
                  NumIntDigits    =   2
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                  Height          =   315
                  Index           =   7
                  Left            =   2340
                  TabIndex        =   174
                  Tag             =   "59"
                  Top             =   1920
                  Width           =   1200
                  _Version        =   262145
                  _ExtentX        =   2117
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
                  NumIntDigits    =   2
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                  Height          =   315
                  Index           =   8
                  Left            =   2340
                  TabIndex        =   175
                  Tag             =   "60"
                  Top             =   2280
                  Width           =   1200
                  _Version        =   262145
                  _ExtentX        =   2117
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
                  NumIntDigits    =   2
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin InDate.ULabel ul_H_SP_EL 
                  Height          =   315
                  Index           =   4
                  Left            =   6000
                  Top             =   840
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   556
                  Caption         =   ""
                  Alignment       =   0
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
                  ForeColor       =   192
               End
               Begin InDate.ULabel ul_H_SP_EL 
                  Height          =   315
                  Index           =   5
                  Left            =   6000
                  Top             =   1200
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   556
                  Caption         =   ""
                  Alignment       =   0
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
                  ForeColor       =   192
               End
               Begin InDate.ULabel ul_H_SP_EL 
                  Height          =   315
                  Index           =   6
                  Left            =   6000
                  Top             =   1560
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   556
                  Caption         =   ""
                  Alignment       =   0
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
                  ForeColor       =   192
               End
               Begin InDate.ULabel ul_H_SP_EL 
                  Height          =   315
                  Index           =   7
                  Left            =   6000
                  Top             =   1920
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   556
                  Caption         =   ""
                  Alignment       =   0
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
                  ForeColor       =   192
               End
               Begin InDate.ULabel ul_H_SP_EL 
                  Height          =   315
                  Index           =   8
                  Left            =   6000
                  Top             =   2280
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   556
                  Caption         =   ""
                  Alignment       =   0
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
                  ForeColor       =   192
               End
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   59
                  Left            =   7200
                  Top             =   960
                  Width           =   1290
                  _ExtentX        =   2275
                  _ExtentY        =   556
                  Caption         =   "应力值1-5"
                  Alignment       =   0
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
               Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                  Height          =   315
                  Index           =   9
                  Left            =   8760
                  TabIndex        =   176
                  Tag             =   "0"
                  Top             =   960
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   4
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                  Height          =   315
                  Index           =   10
                  Left            =   9720
                  TabIndex        =   177
                  Tag             =   "0"
                  Top             =   960
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   4
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                  Height          =   315
                  Index           =   11
                  Left            =   10680
                  TabIndex        =   178
                  Tag             =   "0"
                  Top             =   960
                  Width           =   960
                  _Version        =   262145
                  _ExtentX        =   1693
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
                  NumIntDigits    =   4
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                  Height          =   315
                  Index           =   12
                  Left            =   11760
                  TabIndex        =   179
                  Tag             =   "0"
                  Top             =   960
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   4
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_RST 
                  Height          =   315
                  Index           =   13
                  Left            =   12720
                  TabIndex        =   180
                  Tag             =   "0"
                  Top             =   960
                  Width           =   960
                  _Version        =   262145
                  _ExtentX        =   1693
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
                  NumIntDigits    =   4
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   65
                  Left            =   7200
                  Top             =   1680
                  Width           =   1245
                  _ExtentX        =   2196
                  _ExtentY        =   556
                  Caption         =   "  断口  %"
                  Alignment       =   0
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
               Begin InDate.ULabel ul_METCH_FRACT_DSC_CD 
                  Height          =   315
                  Index           =   9
                  Left            =   13380
                  Top             =   1680
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   556
                  Caption         =   ""
                  Alignment       =   0
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
                  ForeColor       =   192
               End
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   66
                  Left            =   9720
                  Top             =   1680
                  Width           =   1650
                  _ExtentX        =   2910
                  _ExtentY        =   556
                  Caption         =   "Y-合格；N-不合格"
                  Alignment       =   0
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
                  ForeColor       =   0
               End
               Begin CSTextLibCtl.sidbEdit sdb_METCH_FRACT_RSLT 
                  Height          =   315
                  Index           =   14
                  Left            =   8760
                  TabIndex        =   235
                  Tag             =   "62"
                  Top             =   1680
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin VB.Line Line1 
                  Index           =   3
                  X1              =   5880
                  X2              =   5880
                  Y1              =   840
                  Y2              =   2760
               End
               Begin VB.Line Line1 
                  Index           =   2
                  X1              =   7080
                  X2              =   7080
                  Y1              =   840
                  Y2              =   2760
               End
               Begin VB.Line Line2 
                  Index           =   1
                  X1              =   -120
                  X2              =   14520
                  Y1              =   2760
                  Y2              =   2760
               End
               Begin VB.Line Line5 
                  Index           =   3
                  X1              =   13410
                  X2              =   13410
                  Y1              =   0
                  Y2              =   810
               End
               Begin VB.Line Line2 
                  Index           =   27
                  X1              =   0
                  X2              =   14640
                  Y1              =   420
                  Y2              =   420
               End
               Begin VB.Line Line2 
                  Index           =   26
                  X1              =   0
                  X2              =   14640
                  Y1              =   810
                  Y2              =   810
               End
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   36
               Left            =   0
               Top             =   5040
               Width           =   14610
               _ExtentX        =   25770
               _ExtentY        =   556
               Caption         =   "其它试验"
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
            Begin FPSpread.vaSpread Ss3 
               Height          =   7545
               Left            =   -74955
               TabIndex        =   181
               Top             =   360
               Width           =   14700
               _Version        =   393216
               _ExtentX        =   25929
               _ExtentY        =   13309
               _StockProps     =   64
               AllowDragDrop   =   -1  'True
               AllowMultiBlocks=   -1  'True
               AllowUserFormulas=   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   13
               MaxRows         =   1
               Protect         =   0   'False
               RetainSelBlock  =   0   'False
               SpreadDesigner  =   "AQC0034C.frx":01B2
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   3735
               Left            =   -74880
               TabIndex        =   27
               Top             =   4320
               Width           =   14685
               _ExtentX        =   25903
               _ExtentY        =   6588
               _Version        =   196609
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin Threed.SSPanel SSPanel9 
                  Height          =   3555
                  Index           =   0
                  Left            =   0
                  TabIndex        =   29
                  Top             =   0
                  Width           =   14505
                  _ExtentX        =   25585
                  _ExtentY        =   6271
                  _Version        =   196609
                  BevelOuter      =   1
                  RoundedCorners  =   0   'False
                  FloodShowPct    =   -1  'True
                  Begin VB.TextBox txt_NDT_RST 
                     Height          =   315
                     Left            =   8040
                     TabIndex        =   154
                     Tag             =   "53"
                     Top             =   3240
                     Width           =   840
                  End
                  Begin VB.TextBox txt_BEND_RST 
                     Height          =   315
                     Index           =   0
                     Left            =   2340
                     TabIndex        =   9
                     Tag             =   "15"
                     Top             =   462
                     Width           =   840
                  End
                  Begin VB.TextBox txt_RPT_BEND_RST 
                     Height          =   315
                     Left            =   2340
                     MaxLength       =   2
                     TabIndex        =   10
                     Tag             =   "18"
                     Top             =   864
                     Width           =   840
                  End
                  Begin VB.TextBox txt_FOAT_RST 
                     Height          =   315
                     Left            =   2340
                     TabIndex        =   11
                     Tag             =   "19"
                     Top             =   1266
                     Width           =   840
                  End
                  Begin VB.TextBox txt_WLD_HARD_TYP 
                     Height          =   315
                     Left            =   4290
                     TabIndex        =   12
                     Tag             =   "16"
                     Top             =   1668
                     Width           =   840
                  End
                  Begin VB.TextBox txt_WLD_BEND_RST 
                     Height          =   315
                     Left            =   10200
                     TabIndex        =   14
                     Tag             =   "17"
                     Top             =   1668
                     Visible         =   0   'False
                     Width           =   270
                  End
                  Begin VB.TextBox txt_WLD_HARD_NAME 
                     Height          =   315
                     Left            =   5130
                     TabIndex        =   32
                     Top             =   1665
                     Width           =   1965
                  End
                  Begin CSTextLibCtl.sivbLB sivbLB1 
                     Height          =   300
                     Index           =   23
                     Left            =   2280
                     TabIndex        =   30
                     Top             =   2160
                     Width           =   480
                     _Version        =   262145
                     _ExtentX        =   847
                     _ExtentY        =   529
                     _StockProps     =   111
                     Caption         =   "  CSR                                                      %"
                     ForeColor       =   -2147483640
                     BackColor       =   14804173
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Caption         =   "  CSR                                                      %"
                     BorderStyle     =   0
                     BorderEffect    =   2
                     ChiselText      =   2
                  End
                  Begin CSTextLibCtl.sivbLB sivbLB1 
                     Height          =   300
                     Index           =   24
                     Left            =   2280
                     TabIndex        =   31
                     Top             =   2520
                     Width           =   480
                     _Version        =   262145
                     _ExtentX        =   847
                     _ExtentY        =   529
                     _StockProps     =   111
                     Caption         =   "  CLR                                                         %"
                     ForeColor       =   -2147483640
                     BackColor       =   14804173
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Caption         =   "  CLR                                                         %"
                     BorderStyle     =   0
                     BorderEffect    =   2
                     ChiselText      =   2
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_WLD_HARD_RST 
                     Height          =   315
                     Left            =   8130
                     TabIndex        =   13
                     Tag             =   "16"
                     Top             =   1665
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CSR 
                     Height          =   315
                     Index           =   0
                     Left            =   2760
                     TabIndex        =   15
                     Tag             =   "21"
                     Top             =   2160
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   "0.00"
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CLR 
                     Height          =   315
                     Index           =   0
                     Left            =   2760
                     TabIndex        =   16
                     Tag             =   "21"
                     Top             =   2520
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CTR 
                     Height          =   315
                     Index           =   0
                     Left            =   2760
                     TabIndex        =   17
                     Tag             =   "21"
                     Top             =   2880
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_SSCC_YP_RST 
                     Height          =   315
                     Left            =   12360
                     TabIndex        =   18
                     Tag             =   "22"
                     Top             =   2835
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_DWTT_YP_RST1 
                     Height          =   315
                     Left            =   2340
                     TabIndex        =   19
                     Tag             =   "23"
                     Top             =   3240
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_DWTT_YP_RST2 
                     Height          =   315
                     Left            =   3300
                     TabIndex        =   20
                     Tag             =   "23"
                     Top             =   3240
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_DWTT_YP_RST3 
                     Height          =   315
                     Left            =   4290
                     TabIndex        =   21
                     Tag             =   "23"
                     Top             =   3240
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ul_BEND 
                     Height          =   315
                     Index           =   0
                     Left            =   13500
                     Top             =   462
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_RPT_BEND 
                     Height          =   315
                     Left            =   13500
                     Top             =   864
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_FOAT 
                     Height          =   315
                     Left            =   13500
                     Top             =   1266
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_WLD_BEND 
                     Height          =   315
                     Left            =   13500
                     Top             =   1668
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_CWR 
                     Height          =   315
                     Left            =   13500
                     Top             =   2070
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_SSCC_YP 
                     Height          =   315
                     Left            =   13500
                     Top             =   2835
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_DWTT 
                     Height          =   315
                     Left            =   13500
                     Top             =   3240
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   38
                     Left            =   3300
                     Top             =   1668
                     Width           =   840
                     _ExtentX        =   1482
                     _ExtentY        =   556
                     Caption         =   "硬度类型"
                     Alignment       =   0
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
                     ForeColor       =   12582912
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   39
                     Left            =   7200
                     Top             =   1668
                     Width           =   840
                     _ExtentX        =   1482
                     _ExtentY        =   556
                     Caption         =   "硬度值"
                     Alignment       =   0
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
                     ForeColor       =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   40
                     Left            =   9270
                     Top             =   1668
                     Width           =   840
                     _ExtentX        =   1482
                     _ExtentY        =   556
                     Caption         =   "焊缝弯曲"
                     Alignment       =   0
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
                     ForeColor       =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   41
                     Left            =   2340
                     Top             =   1668
                     Width           =   840
                     _ExtentX        =   1482
                     _ExtentY        =   556
                     Caption         =   "焊接硬度"
                     Alignment       =   0
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
                     ForeColor       =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   42
                     Left            =   60
                     Top             =   3240
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "重力撕裂试验 DWTT  %"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   43
                     Left            =   60
                     Top             =   462
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "弯曲试验"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   44
                     Left            =   60
                     Top             =   864
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "反复弯曲          次"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   45
                     Left            =   60
                     Top             =   1266
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "锻平试验"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   46
                     Left            =   60
                     Top             =   1668
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "焊接试验"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   1155
                     Index           =   47
                     Left            =   60
                     Top             =   2070
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   2037
                     Caption         =   "抗氢裂能力（HIC）"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   48
                     Left            =   10800
                     Top             =   2835
                     Width           =   1530
                     _ExtentX        =   2699
                     _ExtentY        =   556
                     Caption         =   "硫化物腐蚀裂纹%"
                     Alignment       =   0
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
                  Begin InDate.ULabel ul_CLR 
                     Height          =   315
                     Left            =   12270
                     Top             =   2070
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_CSR 
                     Height          =   315
                     Left            =   11040
                     Top             =   2040
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_WLD_HARD 
                     Height          =   315
                     Left            =   12270
                     Top             =   1680
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   9
                     Left            =   3300
                     Top             =   1260
                     Width           =   1650
                     _ExtentX        =   2910
                     _ExtentY        =   556
                     Caption         =   "Y-合格；N-不合格"
                     Alignment       =   0
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
                     ForeColor       =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   10
                     Left            =   10500
                     Top             =   1680
                     Width           =   1650
                     _ExtentX        =   2910
                     _ExtentY        =   556
                     Caption         =   "Y-合格；N-不合格"
                     Alignment       =   0
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
                     ForeColor       =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   8
                     Left            =   3300
                     Top             =   480
                     Width           =   1650
                     _ExtentX        =   2910
                     _ExtentY        =   556
                     Caption         =   "Y-合格；N-不合格"
                     Alignment       =   0
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
                     ForeColor       =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   35
                     Left            =   5280
                     Top             =   3240
                     Width           =   2490
                     _ExtentX        =   4392
                     _ExtentY        =   556
                     Caption         =   "NDT重力撕裂试验 DWTT  %"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   60
                     Left            =   2280
                     Top             =   2880
                     Width           =   480
                     _ExtentX        =   847
                     _ExtentY        =   556
                     Caption         =   " CTR"
                     Alignment       =   0
                     BackColor       =   14804173
                     BackgroundStyle =   1
                     ChiselText      =   2
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "幼圆"
                        Size            =   9
                        Charset         =   134
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CSR 
                     Height          =   315
                     Index           =   1
                     Left            =   3600
                     TabIndex        =   182
                     Tag             =   "21"
                     Top             =   2160
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   "0.00"
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CSR 
                     Height          =   315
                     Index           =   2
                     Left            =   4440
                     TabIndex        =   183
                     Tag             =   "21"
                     Top             =   2160
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   "0.00"
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CSR 
                     Height          =   315
                     Index           =   3
                     Left            =   5280
                     TabIndex        =   184
                     Tag             =   "21"
                     Top             =   2160
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   "0.00"
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CSR 
                     Height          =   315
                     Index           =   4
                     Left            =   6120
                     TabIndex        =   185
                     Tag             =   "21"
                     Top             =   2160
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   "0.00"
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CSR 
                     Height          =   315
                     Index           =   5
                     Left            =   6960
                     TabIndex        =   186
                     Tag             =   "21"
                     Top             =   2160
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   "0.00"
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CSR 
                     Height          =   315
                     Index           =   6
                     Left            =   7800
                     TabIndex        =   187
                     Tag             =   "21"
                     Top             =   2160
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   "0.00"
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CSR 
                     Height          =   315
                     Index           =   7
                     Left            =   8640
                     TabIndex        =   188
                     Tag             =   "21"
                     Top             =   2160
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   "0.00"
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CSR 
                     Height          =   315
                     Index           =   8
                     Left            =   9480
                     TabIndex        =   189
                     Tag             =   "21"
                     Top             =   2160
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     Modified        =   -1  'True
                     HideSelection   =   -1  'True
                     RawData         =   "0.00"
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CLR 
                     Height          =   315
                     Index           =   1
                     Left            =   3600
                     TabIndex        =   190
                     Tag             =   "21"
                     Top             =   2520
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CLR 
                     Height          =   315
                     Index           =   2
                     Left            =   4440
                     TabIndex        =   191
                     Tag             =   "21"
                     Top             =   2520
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CLR 
                     Height          =   315
                     Index           =   3
                     Left            =   5280
                     TabIndex        =   192
                     Tag             =   "21"
                     Top             =   2520
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CLR 
                     Height          =   315
                     Index           =   4
                     Left            =   6120
                     TabIndex        =   193
                     Tag             =   "21"
                     Top             =   2520
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CLR 
                     Height          =   315
                     Index           =   5
                     Left            =   6960
                     TabIndex        =   194
                     Tag             =   "21"
                     Top             =   2520
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CLR 
                     Height          =   315
                     Index           =   6
                     Left            =   7800
                     TabIndex        =   195
                     Tag             =   "21"
                     Top             =   2520
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CLR 
                     Height          =   315
                     Index           =   7
                     Left            =   8640
                     TabIndex        =   196
                     Tag             =   "21"
                     Top             =   2520
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CLR 
                     Height          =   315
                     Index           =   8
                     Left            =   9480
                     TabIndex        =   197
                     Tag             =   "21"
                     Top             =   2520
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CTR 
                     Height          =   315
                     Index           =   1
                     Left            =   3600
                     TabIndex        =   198
                     Tag             =   "21"
                     Top             =   2880
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CTR 
                     Height          =   315
                     Index           =   2
                     Left            =   4440
                     TabIndex        =   199
                     Tag             =   "21"
                     Top             =   2880
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CTR 
                     Height          =   315
                     Index           =   3
                     Left            =   5280
                     TabIndex        =   200
                     Tag             =   "21"
                     Top             =   2880
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CTR 
                     Height          =   315
                     Index           =   4
                     Left            =   6120
                     TabIndex        =   201
                     Tag             =   "21"
                     Top             =   2880
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CTR 
                     Height          =   315
                     Index           =   5
                     Left            =   6960
                     TabIndex        =   202
                     Tag             =   "21"
                     Top             =   2880
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CTR 
                     Height          =   315
                     Index           =   6
                     Left            =   7800
                     TabIndex        =   203
                     Tag             =   "21"
                     Top             =   2880
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CTR 
                     Height          =   315
                     Index           =   7
                     Left            =   8640
                     TabIndex        =   204
                     Tag             =   "21"
                     Top             =   2880
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_HIC_CTR 
                     Height          =   315
                     Index           =   8
                     Left            =   9480
                     TabIndex        =   205
                     Tag             =   "21"
                     Top             =   2880
                     Width           =   720
                     _Version        =   262145
                     _ExtentX        =   1270
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
                     NumDecDigits    =   2
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin VB.Line Line11 
                     BorderStyle     =   3  'Dot
                     Index           =   0
                     X1              =   12180
                     X2              =   12180
                     Y1              =   2430
                     Y2              =   2040
                  End
                  Begin VB.Line Line10 
                     Index           =   0
                     X1              =   10680
                     X2              =   10680
                     Y1              =   2760
                     Y2              =   3150
                  End
                  Begin VB.Line Line9 
                     Index           =   0
                     X1              =   12180
                     X2              =   12180
                     Y1              =   1620
                     Y2              =   2040
                  End
                  Begin VB.Line Line8 
                     BorderStyle     =   3  'Dot
                     Index           =   0
                     X1              =   13410
                     X2              =   13410
                     Y1              =   1620
                     Y2              =   2430
                  End
                  Begin VB.Line Line5 
                     Index           =   1
                     X1              =   13410
                     X2              =   13410
                     Y1              =   2790
                     Y2              =   3600
                  End
                  Begin VB.Line Line6 
                     BorderStyle     =   3  'Dot
                     Index           =   1
                     X1              =   3240
                     X2              =   3240
                     Y1              =   1620
                     Y2              =   2040
                  End
                  Begin VB.Line Line6 
                     BorderStyle     =   3  'Dot
                     Index           =   0
                     X1              =   9120
                     X2              =   9120
                     Y1              =   1620
                     Y2              =   2040
                  End
                  Begin VB.Line Line2 
                     Index           =   13
                     X1              =   0
                     X2              =   14640
                     Y1              =   3210
                     Y2              =   3210
                  End
                  Begin VB.Line Line2 
                     Index           =   12
                     X1              =   10680
                     X2              =   14640
                     Y1              =   2790
                     Y2              =   2790
                  End
                  Begin VB.Line Line2 
                     Index           =   11
                     X1              =   0
                     X2              =   14640
                     Y1              =   2040
                     Y2              =   2040
                  End
                  Begin VB.Line Line2 
                     Index           =   10
                     X1              =   0
                     X2              =   14640
                     Y1              =   1620
                     Y2              =   1620
                  End
                  Begin VB.Line Line2 
                     Index           =   9
                     X1              =   0
                     X2              =   14640
                     Y1              =   1230
                     Y2              =   1230
                  End
                  Begin VB.Line Line2 
                     Index           =   8
                     X1              =   0
                     X2              =   14640
                     Y1              =   810
                     Y2              =   810
                  End
                  Begin VB.Line Line2 
                     Index           =   7
                     X1              =   0
                     X2              =   14640
                     Y1              =   420
                     Y2              =   420
                  End
                  Begin VB.Line Line5 
                     Index           =   0
                     X1              =   13410
                     X2              =   13410
                     Y1              =   120
                     Y2              =   1740
                  End
               End
            End
            Begin Threed.SSPanel SSPanel6 
               Height          =   4920
               Index           =   1
               Left            =   30
               TabIndex        =   118
               Top             =   360
               Width           =   14715
               _ExtentX        =   25956
               _ExtentY        =   8678
               _Version        =   196609
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin Threed.SSPanel SSPanel7 
                  Height          =   3465
                  Index           =   1
                  Left            =   0
                  TabIndex        =   119
                  Top             =   360
                  Width           =   7095
                  _ExtentX        =   12515
                  _ExtentY        =   6112
                  _Version        =   196609
                  BevelOuter      =   1
                  RoundedCorners  =   0   'False
                  FloodShowPct    =   -1  'True
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   11
                     Left            =   60
                     Top             =   60
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "屈服强度   YP   MPa"
                     Alignment       =   0
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
                  Begin CSTextLibCtl.sidbEdit sdb_YP_RST 
                     Height          =   315
                     Index           =   1
                     Left            =   2340
                     TabIndex        =   120
                     Tag             =   "36"
                     Top             =   60
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   4
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_TS_RST 
                     Height          =   315
                     Index           =   1
                     Left            =   2340
                     TabIndex        =   121
                     Tag             =   "37"
                     Top             =   924
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   4
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_EL_RST 
                     Height          =   315
                     Index           =   1
                     Left            =   2340
                     TabIndex        =   122
                     Tag             =   "39"
                     Top             =   1356
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_SG_EL_RST 
                     Height          =   315
                     Index           =   1
                     Left            =   2340
                     TabIndex        =   123
                     Tag             =   "42"
                     Top             =   492
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   4
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_YR_RST 
                     Height          =   315
                     Index           =   1
                     Left            =   2340
                     TabIndex        =   124
                     Tag             =   "35"
                     Top             =   2220
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_SNPP_EL_RST 
                     Height          =   315
                     Index           =   1
                     Left            =   2340
                     TabIndex        =   125
                     Tag             =   "40"
                     Top             =   2652
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   4
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_SP_EL_RST 
                     Height          =   315
                     Index           =   1
                     Left            =   2340
                     TabIndex        =   126
                     Tag             =   "41"
                     Top             =   3090
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   4
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_RA_RST_1 
                     Height          =   315
                     Index           =   1
                     Left            =   2340
                     TabIndex        =   127
                     Tag             =   "38"
                     Top             =   1788
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_RA_RST_2 
                     Height          =   315
                     Index           =   1
                     Left            =   3195
                     TabIndex        =   128
                     Tag             =   "38"
                     Top             =   1785
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_RA_RST_3 
                     Height          =   315
                     Index           =   1
                     Left            =   4035
                     TabIndex        =   129
                     Tag             =   "38"
                     Top             =   1785
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin CSTextLibCtl.sidbEdit sdb_RA_RST_AVE 
                     Height          =   315
                     Index           =   1
                     Left            =   4890
                     TabIndex        =   130
                     Tag             =   "38"
                     Top             =   1785
                     Width           =   840
                     _Version        =   262145
                     _ExtentX        =   1482
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
                     NumIntDigits    =   3
                     ShowZero        =   0   'False
                     Undo            =   0
                     Data            =   0
                  End
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   12
                     Left            =   60
                     Top             =   492
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "规定总伸长应力   MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   13
                     Left            =   60
                     Top             =   924
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "抗拉强度   TS   MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   14
                     Left            =   60
                     Top             =   1356
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "断后伸长率   EL   %"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   15
                     Left            =   60
                     Top             =   1788
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "断面收缩率    RA   %"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   16
                     Left            =   60
                     Top             =   2220
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "屈强比   Y.S/T.S   %"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   23
                     Left            =   60
                     Top             =   2652
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "规定非比例伸长应力MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ULabel2 
                     Height          =   315
                     Index           =   24
                     Left            =   60
                     Top             =   3090
                     Width           =   2130
                     _ExtentX        =   3757
                     _ExtentY        =   556
                     Caption         =   "规定残余伸长应力  MPa"
                     Alignment       =   0
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
                  Begin InDate.ULabel ul_YP 
                     Height          =   315
                     Index           =   1
                     Left            =   5970
                     Top             =   60
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_SG_EL 
                     Height          =   315
                     Index           =   1
                     Left            =   5970
                     Top             =   492
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_TS 
                     Height          =   315
                     Index           =   1
                     Left            =   5970
                     Top             =   924
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_EL 
                     Height          =   315
                     Index           =   1
                     Left            =   5970
                     Top             =   1356
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_RA 
                     Height          =   315
                     Index           =   1
                     Left            =   5970
                     Top             =   1788
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_YR 
                     Height          =   315
                     Index           =   1
                     Left            =   5970
                     Top             =   2220
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_SNPP_EL 
                     Height          =   315
                     Index           =   1
                     Left            =   5970
                     Top             =   2652
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin InDate.ULabel ul_SP_EL 
                     Height          =   315
                     Index           =   1
                     Left            =   5970
                     Top             =   3090
                     Width           =   1050
                     _ExtentX        =   1852
                     _ExtentY        =   556
                     Caption         =   ""
                     Alignment       =   0
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
                     ForeColor       =   192
                  End
                  Begin VB.Line Line2 
                     Index           =   20
                     X1              =   0
                     X2              =   7110
                     Y1              =   3030
                     Y2              =   3030
                  End
                  Begin VB.Line Line2 
                     Index           =   19
                     X1              =   0
                     X2              =   7110
                     Y1              =   2580
                     Y2              =   2580
                  End
                  Begin VB.Line Line2 
                     Index           =   18
                     X1              =   0
                     X2              =   7110
                     Y1              =   2160
                     Y2              =   2160
                  End
                  Begin VB.Line Line2 
                     Index           =   17
                     X1              =   0
                     X2              =   7110
                     Y1              =   1710
                     Y2              =   1710
                  End
                  Begin VB.Line Line2 
                     Index           =   16
                     X1              =   0
                     X2              =   7110
                     Y1              =   1290
                     Y2              =   1290
                  End
                  Begin VB.Line Line2 
                     Index           =   15
                     X1              =   0
                     X2              =   7110
                     Y1              =   870
                     Y2              =   870
                  End
                  Begin VB.Line Line2 
                     Index           =   14
                     X1              =   0
                     X2              =   7110
                     Y1              =   420
                     Y2              =   420
                  End
                  Begin VB.Line Line1 
                     Index           =   1
                     X1              =   5880
                     X2              =   5880
                     Y1              =   30
                     Y2              =   3480
                  End
               End
               Begin InDate.ULabel ULabel1 
                  Height          =   315
                  Index           =   33
                  Left            =   30
                  Top             =   30
                  Width           =   7080
                  _ExtentX        =   12488
                  _ExtentY        =   556
                  Caption         =   "追加拉伸试验"
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
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   63
                  Left            =   60
                  Top             =   3840
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   556
                  Caption         =   "追加厚度方向断面收缩率    RA   %"
                  Alignment       =   0
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
               Begin CSTextLibCtl.sidbEdit Z_ZRA_RST_1 
                  Height          =   315
                  Index           =   2
                  Left            =   2400
                  TabIndex        =   216
                  Tag             =   "62"
                  Top             =   3840
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit Z_ZRA_RST_2 
                  Height          =   315
                  Index           =   2
                  Left            =   3240
                  TabIndex        =   217
                  Tag             =   "62"
                  Top             =   3840
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin InDate.ULabel Z_ul_ZRA 
                  Height          =   315
                  Index           =   2
                  Left            =   13560
                  Top             =   3720
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   556
                  Caption         =   ""
                  Alignment       =   0
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
                  ForeColor       =   192
               End
               Begin CSTextLibCtl.sidbEdit Z_ZRA_RST_3 
                  Height          =   315
                  Index           =   2
                  Left            =   4080
                  TabIndex        =   218
                  Tag             =   "62"
                  Top             =   3840
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit Z_ZRA_RST_AVE 
                  Height          =   315
                  Index           =   2
                  Left            =   7440
                  TabIndex        =   219
                  Tag             =   "62"
                  Top             =   3840
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   64
                  Left            =   60
                  Top             =   4320
                  Width           =   2130
                  _ExtentX        =   3757
                  _ExtentY        =   556
                  Caption         =   "追加厚度方向抗拉强度   TS   MPa"
                  Alignment       =   0
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
               Begin CSTextLibCtl.sidbEdit Z_TS_RST_1 
                  Height          =   315
                  Index           =   2
                  Left            =   2400
                  TabIndex        =   220
                  Tag             =   "63"
                  Top             =   4320
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit Z_TS_RST_2 
                  Height          =   315
                  Index           =   2
                  Left            =   3240
                  TabIndex        =   221
                  Tag             =   "63"
                  Top             =   4320
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit Z_TS_RST_3 
                  Height          =   315
                  Index           =   2
                  Left            =   4080
                  TabIndex        =   222
                  Tag             =   "63"
                  Top             =   4320
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin InDate.ULabel Z_ul_TS 
                  Height          =   315
                  Index           =   2
                  Left            =   13560
                  Top             =   4320
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   556
                  Caption         =   ""
                  Alignment       =   0
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
                  ForeColor       =   192
               End
               Begin CSTextLibCtl.sidbEdit Z_ZRA_RST_4 
                  Height          =   315
                  Index           =   2
                  Left            =   4920
                  TabIndex        =   223
                  Tag             =   "62"
                  Top             =   3840
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit Z_ZRA_RST_5 
                  Height          =   315
                  Index           =   2
                  Left            =   5760
                  TabIndex        =   224
                  Tag             =   "62"
                  Top             =   3840
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit Z_ZRA_RST_6 
                  Height          =   315
                  Index           =   2
                  Left            =   6600
                  TabIndex        =   225
                  Tag             =   "62"
                  Top             =   3840
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit Z_TS_RST_4 
                  Height          =   315
                  Index           =   2
                  Left            =   4920
                  TabIndex        =   229
                  Tag             =   "61"
                  Top             =   4320
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit Z_TS_RST_5 
                  Height          =   315
                  Index           =   2
                  Left            =   5760
                  TabIndex        =   230
                  Tag             =   "61"
                  Top             =   4320
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin CSTextLibCtl.sidbEdit Z_TS_RST_6 
                  Height          =   315
                  Index           =   2
                  Left            =   6600
                  TabIndex        =   231
                  Tag             =   "61"
                  Top             =   4320
                  Width           =   840
                  _Version        =   262145
                  _ExtentX        =   1482
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
                  NumIntDigits    =   3
                  ShowZero        =   0   'False
                  Undo            =   0
                  Data            =   0
               End
               Begin VB.Line Line17 
                  X1              =   13440
                  X2              =   13440
                  Y1              =   3600
                  Y2              =   4800
               End
               Begin VB.Line Line16 
                  X1              =   0
                  X2              =   14760
                  Y1              =   4200
                  Y2              =   4200
               End
               Begin VB.Line Line15 
                  X1              =   13200
                  X2              =   13200
                  Y1              =   0
                  Y2              =   960
               End
               Begin VB.Line Line14 
                  X1              =   0
                  X2              =   14520
                  Y1              =   480
                  Y2              =   480
               End
            End
         End
      End
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   14
      Left            =   0
      Top             =   0
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "取样位置"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Index           =   4
      Left            =   0
      Top             =   0
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   556
      Caption         =   "厚度方向抗拉强度   TS   MPa"
      Alignment       =   0
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
End
Attribute VB_Name = "AQC0034C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   判定管理
'-- Program Name      材质试验实绩输入-力学组
'-- Program ID        AQC0034C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Li Qing Yu
'-- Coder             Li Qing Yu
'-- Date              2006.12.01
'-- Description       材质试验实绩输入
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'
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

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

Dim pControl3 As New Collection      'Master Primary Key Collection
Dim nControl3 As New Collection      'Master Necessary Collection
Dim mControl3 As New Collection      'Master Maxlength check Collection
Dim iControl3 As New Collection      'Master Insert Collection
Dim rControl3 As New Collection      'Master Refer Collection
Dim cControl3 As New Collection      'Master Copy Collection
Dim aControl3 As New Collection      'Master -> Spread Collection
Dim lControl3 As New Collection      'Master Lock Collection

Dim pColumn As New Collection      'Spread Primary Key Collection
Dim nColumn As New Collection      'Spread necessary Column Collection
Dim mColumn As New Collection      'Spread Maxlength check Column Collection
Dim iColumn As New Collection      'Spread Insert Column Collection
Dim aColumn As New Collection      'Master -> Spread Column Collection
Dim lColumn As New Collection      'Spread Lock Column Collection


Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection
Dim sc3 As New Collection


Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim sOldAuthority As String         'Save First Load Authority
Dim bExpo_SMP   As Boolean          'This sampling is Expo sampling when value is true


Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'TOP and STAND
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
          Call Gp_Ms_Collection(txt_smp_no_p, "p", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_STLGRD, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(lbl_HEAT_NO, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_Cut_DD, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(lbl_STDSPEC, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_STD_YY, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(lbl_ORD_NO, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(lbl_ORD_ITEM, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(lbl_ENDUSE_CD, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(lbl_ORD_THK, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(lbl_ORD_WID, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
              Call Gp_Ms_Collection(ul_YP(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(ul_SG_EL(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
              Call Gp_Ms_Collection(ul_TS(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
              Call Gp_Ms_Collection(ul_EL(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
              Call Gp_Ms_Collection(ul_RA(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
               'louyannan 201011121
               Call Gp_Ms_Collection(ul_RA(2), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2) 'louyannan 201011121
              
              Call Gp_Ms_Collection(ul_YR(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(ul_SNPP_EL(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(ul_SP_EL(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(ul_H_YP(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(ul_H_TS(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(ul_H_RA(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            'louyanna 20101121
             'Call Gp_Ms_Collection(ul_H_RA(2), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
             
            Call Gp_Ms_Collection(ul_H_EL(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(ul_H_SNPP_EL(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(ul_H_SP_EL(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(ul_HARD(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(ul_BEND(0), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(ul_RPT_BEND, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
               Call Gp_Ms_Collection(ul_FOAT, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(ul_WLD_HARD, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(ul_WLD_BEND, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                Call Gp_Ms_Collection(ul_CSR, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                Call Gp_Ms_Collection(ul_CLR, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                Call Gp_Ms_Collection(ul_CWR, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(ul_SSCC_YP, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
               Call Gp_Ms_Collection(ul_DWTT, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(ul_IMPACT_MIN, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(ul_IMPACT_MIN_MIN, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(ul_IMPACT_AVE, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(ul_IMPACT_RATE_MIN, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(ul_IMPACT_RATE_MAX, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         'louyanna 20101121 侧膨胀下限
          Call Gp_Ms_Collection(ULabel1(83), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(ULabel1(84), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       
       Call Gp_Ms_Collection(ul_IMPACT_A_MIN, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   Call Gp_Ms_Collection(ul_IMPACT_A_MIN_MIN, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(ul_IMPACT_A_AVE, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  Call Gp_Ms_Collection(ul_IMPACT_A_RATE_MIN, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  Call Gp_Ms_Collection(ul_IMPACT_A_RATE_MAX, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            
             'louyanna 20101121 追加侧膨胀下限
          Call Gp_Ms_Collection(ULabel1(114), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(ULabel1(122), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            
            Call Gp_Ms_Collection(ul_TIM_MIN, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(ul_TIM_MIN_MIN, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(ul_TIM_AVE, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(ul_TIM_RATE, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(ul_A_TIM_MIN, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(ul_A_TIM_MIN_MIN, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(ul_A_TIM_AVE, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(ul_A_TIM_RATE, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
'20090803 SUN BIN START
              Call Gp_Ms_Collection(ul_YP(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(ul_SG_EL(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
              Call Gp_Ms_Collection(ul_TS(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
              Call Gp_Ms_Collection(ul_EL(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
              Call Gp_Ms_Collection(ul_RA(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
              Call Gp_Ms_Collection(ul_YR(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(ul_SNPP_EL(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(ul_SP_EL(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(ul_H_YP(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(ul_H_TS(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(ul_H_RA(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(ul_H_EL(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(ul_H_SNPP_EL(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(ul_H_SP_EL(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(ul_HARD(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(ul_BEND(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
'edit by gengxueyu for kangda on 20110211 uel普通+追加+应力比1-5 标准
         Call Gp_Ms_Collection(ul_H_SP_EL(2), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(ul_H_SP_EL(3), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(ul_H_SP_EL(4), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(ul_H_SP_EL(5), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(ul_H_SP_EL(6), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(ul_H_SP_EL(7), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(ul_H_SP_EL(8), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         
'20090803 SUN BIN END
         
'2016-11-29 LJN START
         Call Gp_Ms_Collection(ul_TS(2), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(Z_ul_ZRA(2), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(Z_ul_TS(2), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         
'2016-11-29  LJN  END
    'MASTER2 Collection
     Mc2.Add Item:="AQC0034C.P_REFER_HEAD", Key:="P-R"
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'试样号&取样位置
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
              Call Gp_Ms_Collection(txt_smp_no_p, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_smp_loc_p, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'拉伸／高温拉伸／其它 - TAB 1
'----------------------------------------------------------------------------------------------------------------------------------------------------------------


             Call Gp_Ms_Collection(sdb_YP_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_SG_EL_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_SNPP_EL_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_SP_EL_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_TS_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_EL_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

          Call Gp_Ms_Collection(sdb_RA_RST_1(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_RA_RST_2(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_RA_RST_3(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_RA_RST_AVE(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             
            'louyannan 20101121 start
            
           Call Gp_Ms_Collection(ZRA_RST_1(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(ZRA_RST_2(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(ZRA_RST_3(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(ZRA_RST_AVE(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              'louyannan 20101121 end
              
             Call Gp_Ms_Collection(sdb_YR_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_HGT_YP_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_HGT_TS_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_RA_RST_1(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_RA_RST_2(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_RA_RST_3(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_RA_RST_A(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
          'louyannan 20101121 start
       'Call Gp_Ms_Collection(sdb_HGT_RA_RST_1(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       'Call Gp_Ms_Collection(sdb_HGT_RA_RST_2(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       'Call Gp_Ms_Collection(sdb_HGT_RA_RST_3(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       'Call Gp_Ms_Collection(sdb_HGT_RA_RST_A(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         'louyannan 20101121 end
         
         
         
         Call Gp_Ms_Collection(sdb_HGT_EL_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_HGT_SNPP_EL_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_HARD_TYP(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_HARD_NAME(0), " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_HARD_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_BEND_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_WLD_HARD_TYP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_WLD_HARD_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_WLD_HARD_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_WLD_BEND_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_RPT_BEND_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_FOAT_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CSR(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CLR(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CTR(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CSR(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CLR(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CTR(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CSR(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CLR(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CTR(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CSR(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CLR(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CTR(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CSR(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CLR(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CTR(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CSR(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CLR(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CTR(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CSR(6), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CLR(6), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CTR(6), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CSR(7), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CLR(7), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CTR(7), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CSR(8), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CLR(8), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HIC_CTR(8), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

           Call Gp_Ms_Collection(sdb_SSCC_YP_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_DWTT_YP_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_DWTT_YP_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_DWTT_YP_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
          'louyannan 20101201
          Call Gp_Ms_Collection(txt_NDT_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'冲击／时效冲击 - TAB 2
'----------------------------------------------------------------------------------------------------------------------------------------------------------------

            Call Gp_Ms_Collection(txt_IMPACT_KND, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_IMPACT_KND_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_IMPACT_DIR, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_IMPACT_DIR_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(Cob_IMPACT_SIZE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_IMPACT_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_IMPACT_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_IMPACT_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_IMPACT_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_IMPACT_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_IMPACT_RST6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_IMPACT_RST_AVE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_IMPACT_RATE_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_IMPACT_RATE_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_IMPACT_RATE_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_IMPACT_RATE_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_IMPACT_RATE_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_IMPACT_RATE_RST6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_IMPACT_RATE_AVE_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

 'louyannan 20101121 start


  Call Gp_Ms_Collection(sdb_EXPAIN_RST(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_EXPAIN_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_EXPAIN_RST(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_EXPAIN_RST(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_EXPAIN_RST(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_EXPAIN_RST(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_EXPAIN_RST(6), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 'louyannan 20101121 end



          Call Gp_Ms_Collection(txt_A_IMPACT_KND, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_A_IMPACT_KND_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_A_IMPACT_DIR, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_A_IMPACT_DIR_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(Cob_A_IMPACT_SIZE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_A_IMPACT_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_A_IMPACT_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_A_IMPACT_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_A_IMPACT_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_A_IMPACT_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_A_IMPACT_RST6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_A_IMPACT_RST_AVE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'louyannan 20101121 start

     Call Gp_Ms_Collection(sdb_EXPAIN_RST(7), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_EXPAIN_RST(8), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_EXPAIN_RST(9), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_EXPAIN_RST(10), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_EXPAIN_RST(11), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_EXPAIN_RST(12), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_EXPAIN_RST(13), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    'louyannan 20101121 end
    
    Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_RST6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_AVE_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

        Call Gp_Ms_Collection(txt_TIM_IMPACT_KND, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_TIM_IMPACT_KND_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_TIM_IMPACT_DIR, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_TIM_IMPACT_DIR_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(Cob_TIM_IMPACT_SIZE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_TIM_IMPACT_RST_AVE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_TIM_IMPACT_RATE_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

      Call Gp_Ms_Collection(txt_A_TIM_IMPACT_KND, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_A_TIM_IMPACT_KND_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_A_TIM_IMPACT_DIR, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_A_TIM_IMPACT_DIR_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(Cob_A_TIM_IMPACT_SIZE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RST_AVE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RATE_RST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'20090803 SUN BIN START
             Call Gp_Ms_Collection(sdb_YP_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_SG_EL_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_SNPP_EL_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_SP_EL_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_TS_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_EL_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_RA_RST_1(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_RA_RST_2(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_RA_RST_3(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_RA_RST_AVE(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_YR_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_HGT_YP_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_HGT_TS_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_RA_RST_1(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_RA_RST_2(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_RA_RST_3(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_RA_RST_A(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_HGT_EL_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_HGT_SNPP_EL_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_HARD_TYP(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_HARD_NAME(1), " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_HARD_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_BEND_RST(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'20090803 SUN BIN END
'----------------------------------------------------------- Master End ------------------------------------------------------------------------------------
                   'Call Gp_Ms_Collection(txt_KND, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_INS_EMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_INPUT_EMP, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_UPD_EMP, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'edit by gengxueyu for kangda 20110211 uel 普通UEL+追加UEL+应力比1-5+应力值1-5
       Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(6), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(7), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(8), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(9), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(10), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(11), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(12), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_HGT_SP_EL_RST(13), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
'2016-11-29  ljn  start
       Call Gp_Ms_Collection(ZRA_RST_4(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(ZRA_RST_5(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(ZRA_RST_6(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TS_RST_1(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TS_RST_2(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TS_RST_3(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TS_RST_4(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TS_RST_5(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TS_RST_6(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Z_ZRA_RST_1(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Z_ZRA_RST_2(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Z_ZRA_RST_3(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Z_ZRA_RST_4(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Z_ZRA_RST_5(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Z_ZRA_RST_6(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Z_ZRA_RST_AVE(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Z_TS_RST_1(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Z_TS_RST_2(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Z_TS_RST_3(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Z_TS_RST_4(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Z_TS_RST_5(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Z_TS_RST_6(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'2016-11-29  ljn  end
'2017-01-22  ljn  start
       Call Gp_Ms_Collection(sdb_METCH_FRACT_RSLT(14), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(ul_METCH_FRACT_DSC_CD(9), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'2016-01-22  ljn  end
      
    'MASTER Collection
     Mc1.Add Item:="AQC0034C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AQC0034C.P_REFER", Key:="P-R"
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
'----配置化材质项目    刘翔   2012.11.20  --------------------------------------------------------------------------------------------------------------------

Private Sub Form_Define1()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"
     
 Call Gp_Ms_Collection(txt_smp_no_p, "p", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
Call Gp_Ms_Collection(txt_smp_loc_p, "p", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
             
     Mc3.Add Item:=pControl3, Key:="pControl"
     Mc3.Add Item:=nControl3, Key:="nControl"
     Mc3.Add Item:=mControl3, Key:="mControl"
     Mc3.Add Item:=iControl3, Key:="iControl"
     Mc3.Add Item:=rControl3, Key:="rControl"
     Mc3.Add Item:=cControl3, Key:="cControl"
     Mc3.Add Item:=aControl3, Key:="aControl"
     Mc3.Add Item:=lControl3, Key:="lControl"
     
     Call Gp_Sp_Collection(ss3, 1, "p", "n", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss3, 2, "p", "n", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss3, 3, "p", "n", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss3, 5, "p", "n", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)

         'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="AQC0034C.P_SMODIFY", Key:="P-M"
    sc3.Add Item:="AQC0034C.P_SREFER", Key:="P-R"
    sc3.Add Item:=pColumn, Key:="pColumn"
    sc3.Add Item:=nColumn, Key:="nColumn"
    sc3.Add Item:=aColumn, Key:="aColumn"
    sc3.Add Item:=mColumn, Key:="mColumn"
    sc3.Add Item:=iColumn, Key:="iColumn"
    sc3.Add Item:=lColumn, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc3, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss3, 3, True)
    Call Gp_Sp_ColHidden(ss3, 5, True)
    Call Gp_Sp_ColHidden(ss3, 12, True)
     
End Sub

Private Sub Cob_A_IMPACT_SIZE_Change()

    Call Impact_Size_Cob_Select(Cob_A_IMPACT_SIZE, TXT_A_IMPACT_SIZE_CD)

End Sub

Private Sub Cob_A_TIM_IMPACT_SIZE_Change()

    Call Impact_Size_Cob_Select(Cob_A_TIM_IMPACT_SIZE, TXT_A_TIM_IMPACT_SIZE_CD)

End Sub

Private Sub Cob_IMPACT_SIZE_Change()

    Call Impact_Size_Cob_Select(Cob_IMPACT_SIZE, TXT_IMPACT_SIZE_CD)

End Sub

Private Sub Cob_TIM_IMPACT_SIZE_Change()

    Call Impact_Size_Cob_Select(Cob_TIM_IMPACT_SIZE, TXT_TIM_IMPACT_SIZE_CD)

End Sub


'
''---------------------------------------------------------------------------------------------------------------------------------------------
''--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
''---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String

    Select Case Me.ActiveControl.Name

        Case "txt_HARD_TYP"         '硬度试验
            sCode = "Q0010"
           
           If Me.ActiveControl.Text = txt_HARD_TYP(0) Then
               Set oCodeName = txt_HARD_NAME(0)
            ElseIf Me.ActiveControl.Text = txt_HARD_TYP(1) Then
               Set oCodeName = txt_HARD_NAME(1)
            End If
            
        Case "txt_WLD_HARD_TYP"     '焊接硬度
            sCode = "Q0011"

        Case "txt_JOMINY_TYP"      '淬透性试验
            sCode = "Q0012"

        Case "txt_IMPACT_KND"      '冲击试验
            sCode = "Q0008"
            Set oCodeName = txt_IMPACT_KND_NAME

        Case "txt_IMPACT_DIR"      '冲击试验
            sCode = "Q0009"
            Set oCodeName = txt_IMPACT_DIR_NAME

        Case "txt_A_IMPACT_KND"      '追加冲击试验
            sCode = "Q0008"
            Set oCodeName = txt_A_IMPACT_KND_NAME

        Case "txt_A_IMPACT_DIR"      '追加冲击试验
            sCode = "Q0009"
            Set oCodeName = txt_A_IMPACT_DIR_NAME

        Case "txt_TIM_IMPACT_KND"      '时效冲击试验
            sCode = "Q0008"
            Set oCodeName = txt_TIM_IMPACT_KND_NAME

        Case "txt_TIM_IMPACT_DIR"      '时效冲击试验
            sCode = "Q0009"
            Set oCodeName = txt_TIM_IMPACT_DIR_NAME

        Case "txt_A_TIM_IMPACT_KND"      '追加时效冲击试验
            sCode = "Q0008"
            Set oCodeName = txt_A_TIM_IMPACT_KND_NAME

        Case "txt_A_TIM_IMPACT_DIR"      '追加时效冲击试验
            sCode = "Q0009"
            Set oCodeName = txt_A_TIM_IMPACT_DIR_NAME


        Case Else
            Exit Sub

    End Select

    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)

    Set oCodeName = Nothing
Err_Track:
End Sub
'

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub
'
Private Sub Form_KeyPress(KeyAscii As Integer)


    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    ElseIf KeyAscii = 19 Or KeyAscii = 10 Then
        KeyAscii = 0
        Call Form_Pro
    End If


End Sub
'
Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)
    sOldAuthority = sAuthority
    
    Call Form_Define
    Call Form_Define1

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    
    
    Call Gp_Ms_Cls(Mc3("rControl"))
    
    Call Gp_Ms_NeceColor(Mc3("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Z-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing

    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
    Set pControl3 = Nothing
    Set nControl3 = Nothing
    Set iControl3 = Nothing
    Set rControl3 = Nothing
    Set cControl3 = Nothing
    Set aControl3 = Nothing
    Set lControl3 = Nothing
    Set mControl3 = Nothing
    
    Set iColumn = Nothing
    Set pColumn = Nothing
    Set lColumn = Nothing
    Set nColumn = Nothing
    Set mColumn = Nothing
    Set aColumn = Nothing
    

    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set sc3 = Nothing
    Set Proc_Sc = Nothing


    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub
'
Public Sub Form_Exit()

    Unload Me

End Sub
'
Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc3("rControl"))
        Call Gp_Ms_ControlLock(Mc3("lControl"), False)
    End If


    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    lbl_STLGRD.Caption = ""
    lbl_HEAT_NO.Caption = ""
    lbl_Cut_DD.Caption = ""
    lbl_STDSPEC.Caption = ""
    lbl_STD_YY.Caption = ""
    lbl_ORD_NO.Caption = ""
    lbl_ORD_ITEM.Caption = ""
    lbl_ENDUSE_CD.Caption = ""
    lbl_ORD_WID.Caption = ""
    lbl_ORD_THK.Caption = ""
    


End Sub
'
Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)

End Sub
'
Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)

End Sub
'
Public Sub Form_Ref()
    Dim sMesg As String
    Dim sSMP_NO As String
    Dim sPROD_CD As String
    'sAuthority = "1111" 'louyannan 20101121
    
        sSMP_NO = Trim(txt_SMP_NO.Text)
        
        sPROD_CD = SMP_PROD_Check(sSMP_NO)
        
        If sPROD_CD = "ER" Then Exit Sub
                
        Call Form_Cls

        If SSRibbon_SMP_TYPE_KND.Value = True Then
            If bExpo_SMP = False Then
                Call MsgBox("现在只能输入作普样！", vbOKOnly, "系统提示")
                Exit Sub
            End If
        Else
            If bExpo_SMP = True Then
                Call MsgBox("现在只能输入非作普样！", vbOKOnly, "系统提示")
                Exit Sub
            End If
        End If
        
        If Gf_Ms_Refer(M_CN1, Mc2, Mc1("nControl"), Mc1("mControl")) Then
            If sAuthority = "1000" Or sAuthority = "0000" Then
                Call MsgBox("你没有当前试样号：" + sSMP_NO + " 操作权限！", vbOKOnly, "系统提示")
            End If
            Call Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl"), False)
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("pControl"), True)

        End If

    Call subItemLock(txt_SMP_NO.Text)
    Call subCODENAMElOCK


    If Val(lbl_ORD_THK.Caption) >= 12 Then
       If TXT_IMPACT_SIZE_CD.Text = "" Then
            TXT_IMPACT_SIZE_CD.Text = "3"
       End If
    End If
    
     '--------------------配置化材质项目  刘翔  2012.11.20-----------------------------------------------------
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc3, Mc3("nControl"), Mc3("lControl"), False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Exit Sub
    End If

End Sub



Private Sub ss3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 12)
    End If
    
End Sub

Private Sub txt_BEND_RST_Change(Index As Integer)

Select Case Me.ActiveControl.Name
   Case "txt_BEND_RST"
        If Me.ActiveControl.Text = txt_BEND_RST(0) Then
            Call TXT_INPUT_VALUE_CHECK(txt_BEND_RST(0))
        ElseIf Me.ActiveControl.Text = txt_BEND_RST(1) Then
            Call TXT_INPUT_VALUE_CHECK(txt_BEND_RST(1))
        End If
     Case Else
            Exit Sub
        
    End Select
  
End Sub

Private Sub txt_SMP_CUT_LOC_Change()
Dim sPROD_CD As String

    sPROD_CD = SMP_PROD_Check(Trim(txt_smp_no_p.Text))
    
    If sPROD_CD = "ER" Then
        Exit Sub
    Else
        txt_smp_loc_p.Text = Trim(txt_smp_cut_loc.Text)
    End If
End Sub

Public Sub Form_Pro()

    If Gf_Mc_Authority(sAuthority, Mc1) Then

        txt_INS_EMP.Text = sUserID
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
            If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc3) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
        End If
    End If
    

End Sub
'
Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub
'
Private Sub subItemLock(ByVal sSMP_NO As String)
    Dim sQuery          As String
    Dim arrayRecord     As Variant
    Dim AdoRs           As adodb.Recordset

 On Error GoTo Error_Rtn
    Set AdoRs = New adodb.Recordset
    
    sQuery = "{call AQC0034C.P_MART_ITEM_SELECT('" + sSMP_NO + "')}"

    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not (AdoRs.BOF And AdoRs.EOF) Then
        arrayRecord = AdoRs.GetRows
    Else
        GoTo Error_Rtn
    End If

    AdoRs.Close

    Call subControlLock(arrayRecord, False, Mc1("iControl"))

    Set AdoRs = Nothing
    Set arrayRecord = Nothing

Error_Rtn:

    Set AdoRs = Nothing
    Set arrayRecord = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub subControlLock(ByVal vARRAY As Variant, ByVal bAllLock As Boolean, ByVal iCtrl As Collection)
    Dim icount       As Integer
    Dim iarrCOUNT    As Integer

    If bAllLock Then
        For icount = 1 To iCtrl.count
            iCtrl.Item(icount).Visible = False
        Next
    Else

        For icount = 1 To iCtrl.count
            If iCtrl.Item(icount).Tag <> 99 And iCtrl.Item(icount).Tag <> "INS_EMP" Then

             For iarrCOUNT = 0 To UBound(vARRAY, 1)


                        If Val(iCtrl.Item(icount).Tag) = Val(vARRAY(iarrCOUNT, 0)) Then
                            iCtrl.Item(icount).Visible = True
                            Exit For
                        Else
                            iCtrl.Item(icount).Visible = False
                        End If

                    Next
                   

            End If
        Next
        If txt_BEND_RST(0).Visible = False Then
           txt_BEND_RST(0).Text = ""
        Else
           If txt_BEND_RST(0).Text = "" Then
              txt_BEND_RST(0).Text = "Y"
           End If
        End If
         If txt_BEND_RST(1).Visible = False Then
            txt_BEND_RST(1).Text = ""
        Else
'           If txt_BEND_RST(1).Text = "" Then
'              txt_BEND_RST(1).Text = "Y"
'           End If
        End If
        If txt_WLD_BEND_RST.Visible = False Then
           txt_WLD_BEND_RST.Text = ""
        Else
           If txt_WLD_BEND_RST.Text = "" Then
              txt_WLD_BEND_RST.Text = "Y"
           End If
        End If
        If txt_FOAT_RST.Visible = False Then
           txt_FOAT_RST.Text = ""
        Else
           If txt_FOAT_RST.Text = "" Then
              txt_FOAT_RST.Text = "Y"
           End If
        End If
            
    End If


End Sub
'
'
Private Sub subCODENAMElOCK()
       
   txt_HARD_NAME(0).Visible = txt_HARD_TYP(0).Visible
   txt_HARD_NAME(1).Visible = txt_HARD_TYP(1).Visible
   txt_WLD_HARD_NAME.Visible = txt_WLD_HARD_TYP.Visible
   
   txt_IMPACT_KND_NAME.Visible = txt_IMPACT_KND.Visible
   txt_IMPACT_DIR_NAME.Visible = txt_IMPACT_DIR.Visible
   TXT_IMPACT_SIZE_CD.Visible = Cob_IMPACT_SIZE.Visible
   txt_TIM_IMPACT_KND_NAME.Visible = txt_TIM_IMPACT_KND.Visible
   txt_TIM_IMPACT_DIR_NAME.Visible = txt_TIM_IMPACT_DIR.Visible
   TXT_TIM_IMPACT_SIZE_CD.Visible = Cob_TIM_IMPACT_SIZE.Visible
   txt_A_IMPACT_KND_NAME.Visible = txt_A_IMPACT_KND.Visible
   txt_A_IMPACT_DIR_NAME.Visible = txt_A_IMPACT_DIR.Visible
   TXT_A_IMPACT_SIZE_CD.Visible = Cob_A_IMPACT_SIZE.Visible
   txt_A_TIM_IMPACT_KND_NAME.Visible = txt_A_TIM_IMPACT_KND.Visible
   txt_A_TIM_IMPACT_DIR_NAME.Visible = txt_A_TIM_IMPACT_DIR.Visible
   TXT_A_TIM_IMPACT_SIZE_CD.Visible = Cob_A_TIM_IMPACT_SIZE.Visible

End Sub
'

Private Sub TXT_A_IMPACT_SIZE_CD_Change()

   Call Impact_Size_Text_Select(Cob_A_IMPACT_SIZE, TXT_A_IMPACT_SIZE_CD)

End Sub

Private Sub TXT_A_TIM_IMPACT_SIZE_CD_Change()

    Call Impact_Size_Text_Select(Cob_A_TIM_IMPACT_SIZE, TXT_A_TIM_IMPACT_SIZE_CD)

End Sub

Private Sub TXT_IMPACT_SIZE_CD_Change()

    Call Impact_Size_Text_Select(Cob_IMPACT_SIZE, TXT_IMPACT_SIZE_CD)

End Sub
'
Private Sub txt_SMP_CUT_LOC_LostFocus()

    Call Form_Ref
    
End Sub

Private Sub Impact_Size_Cob_Select(oCob As ComboBox, oEdit As TextBox)

     Select Case Trim(oCob.Text)
            Case "5*10*55"
                oEdit.Text = "1"
            Case "7.5*10*55"
                oEdit.Text = "2"
            Case "10*10*55"
                oEdit.Text = "3"
            Case Else
                oEdit.Text = ""
     End Select

End Sub

Private Sub Impact_Size_Text_Select(oCob As ComboBox, oEdit As TextBox)

     Select Case Trim(oEdit.Text)
            Case "1"
                oCob.ListIndex = 1
            Case "2"
                oCob.ListIndex = 2
            Case "3"
                oCob.ListIndex = 3
            Case Else
                oCob.ListIndex = 0
     End Select

End Sub
'
Private Sub txt_SMP_NO_Change()
Dim sPROD_CD As String
    
    txt_smp_no_p.Text = txt_SMP_NO.Text
    
    sPROD_CD = SMP_PROD_Check(Trim(txt_smp_no_p.Text))
    
    If sPROD_CD = "ER" Then
        txt_smp_cut_loc.Text = ""
    Else
        txt_smp_cut_loc.Text = Find_SMP_LOC(Trim(txt_smp_no_p.Text))
        sAuthority = Ship_Input_AUTH(Trim(txt_smp_no_p.Text), sUserID, sOldAuthority)
        bExpo_SMP = Expo_SMP_Check(Trim(txt_smp_no_p.Text))
    End If

End Sub
'
Private Sub TXT_TIM_IMPACT_SIZE_CD_Change()
    
    Call Impact_Size_Text_Select(Cob_TIM_IMPACT_SIZE, TXT_TIM_IMPACT_SIZE_CD)
    
End Sub
'

Private Sub SSRibbon_SMP_TYPE_KND_Click(Value As Integer)
    
    If Value = True Then
        SSRibbon_SMP_TYPE_KND.Caption = "录入状态：作普样录入"
        SSRibbon_SMP_TYPE_KND.ForeColor = &HFF0000
    Else
        SSRibbon_SMP_TYPE_KND.Caption = "录入状态： 常规样录入"
        SSRibbon_SMP_TYPE_KND.ForeColor = &HFF
    End If
    txt_SMP_NO.SetFocus
    
End Sub

Private Sub TXT_INPUT_VALUE_CHECK(ByVal oControl As Object)
    Dim s_INPUT_VALUE As String
    
    If TypeOf oControl Is TextBox Then
        s_INPUT_VALUE = UCase(Trim(oControl.Text))
        
        If s_INPUT_VALUE <> "" Or Len(s_INPUT_VALUE) > 0 Then
            If s_INPUT_VALUE <> "Y" And s_INPUT_VALUE <> "N" Then
                Call MsgBox("录入错误！只能录入[Y/N]！", vbOKOnly, "系统提示")
                oControl.Text = ""
                oControl.SetFocus
            Else
                oControl.Text = s_INPUT_VALUE
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
End Sub

