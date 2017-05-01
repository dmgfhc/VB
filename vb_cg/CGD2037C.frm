VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGD2037C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "上/下线实绩处理界面_CGD2037C"
   ClientHeight    =   9840
   ClientLeft      =   1725
   ClientTop       =   1815
   ClientWidth     =   15195
   FillStyle       =   2  'Horizontal Line
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   15195
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9285
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   16378
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      Locked          =   -1  'True
      PaneTree        =   "CGD2037C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   7725
         Left            =   0
         TabIndex        =   1
         Top             =   1560
         Width           =   15225
         _Version        =   393216
         _ExtentX        =   26855
         _ExtentY        =   13626
         _StockProps     =   64
         ColsFrozen      =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   19
         MaxRows         =   10
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGD2037C.frx":0052
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1500
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   2646
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.ComboBox CBO_GROUP 
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
            ItemData        =   "CGD2037C.frx":0A87
            Left            =   7245
            List            =   "CGD2037C.frx":0A89
            TabIndex        =   23
            Tag             =   "班别"
            Top             =   120
            Width           =   855
         End
         Begin VB.ComboBox CBO_SHIFT 
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
            ItemData        =   "CGD2037C.frx":0A8B
            Left            =   6390
            List            =   "CGD2037C.frx":0A8D
            TabIndex        =   22
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox TXT_MAT_NO 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1470
            MaxLength       =   14
            TabIndex        =   21
            Top             =   540
            Width           =   2160
         End
         Begin VB.TextBox TXT_SEQ 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6390
            MaxLength       =   12
            TabIndex        =   20
            Top             =   540
            Width           =   870
         End
         Begin VB.TextBox txt_line 
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
            Left            =   9600
            MaxLength       =   1
            TabIndex        =   19
            Tag             =   "CD_MANA_NO"
            Text            =   "1"
            Top             =   120
            Width           =   480
         End
         Begin VB.TextBox txt_onoff 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4020
            TabIndex        =   7
            Tag             =   "上/下线代码"
            Text            =   "ON"
            Top             =   2550
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.TextBox txt_plt 
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
            Left            =   1830
            MaxLength       =   2
            TabIndex        =   6
            Top             =   2550
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.TextBox txt_plt_dec 
            Enabled         =   0   'False
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
            Left            =   2550
            MaxLength       =   11
            TabIndex        =   5
            Top             =   2550
            Visible         =   0   'False
            Width           =   1260
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
            Height          =   285
            Left            =   4950
            TabIndex        =   4
            Tag             =   "上/下线代码"
            Text            =   "CG1"
            Top             =   2550
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.TextBox txt_m_r 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5880
            TabIndex        =   3
            Tag             =   "操作 / 查询"
            Text            =   "M"
            Top             =   2550
            Visible         =   0   'False
            Width           =   840
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   435
            Left            =   4320
            TabIndex        =   8
            Top             =   960
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   767
            _Version        =   196609
            BackColor       =   14737632
            Begin Threed.SSOption opt_OnPosition 
               Height          =   375
               Index           =   2
               Left            =   4470
               TabIndex        =   9
               Tag             =   "CI1"
               Top             =   30
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   661
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "4# 剪"
            End
            Begin Threed.SSOption opt_OnPosition 
               Height          =   375
               Index           =   3
               Left            =   6495
               TabIndex        =   10
               Tag             =   "CH3"
               Top             =   30
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   661
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "2# 线划线"
            End
            Begin Threed.SSOption opt_OnPosition 
               Height          =   375
               Index           =   4
               Left            =   8520
               TabIndex        =   11
               Tag             =   "CI2"
               Top             =   30
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   661
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "2# 线定尺剪"
            End
            Begin Threed.SSOption opt_OnPosition 
               Height          =   375
               Index           =   5
               Left            =   9930
               TabIndex        =   12
               Tag             =   "CI1"
               Top             =   30
               Visible         =   0   'False
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "预留"
            End
            Begin Threed.SSOption opt_OnPosition 
               Height          =   375
               Index           =   6
               Left            =   10710
               TabIndex        =   13
               Tag             =   "CI2"
               Top             =   30
               Visible         =   0   'False
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "预留"
            End
            Begin Threed.SSOption opt_OnPosition 
               Height          =   375
               Index           =   0
               Left            =   420
               TabIndex        =   14
               Tag             =   "CH1"
               Top             =   30
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   661
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "1# 剪"
               Value           =   -1
            End
            Begin Threed.SSOption opt_OnPosition 
               Height          =   375
               Index           =   1
               Left            =   2445
               TabIndex        =   15
               Tag             =   "CH2"
               Top             =   30
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   661
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "3# 剪"
            End
         End
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   510
            Top             =   2550
            Visible         =   0   'False
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "生产工厂"
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
         Begin InDate.ULabel ULabel1 
            Height          =   435
            Left            =   150
            Top             =   960
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   767
            Caption         =   "上/下线工位"
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
            ForeColor       =   0
         End
         Begin Threed.SSFrame SSFrame3 
            Height          =   435
            Left            =   1680
            TabIndex        =   16
            Top             =   960
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   767
            _Version        =   196609
            BackColor       =   14737632
            Begin Threed.SSOption opt_online 
               Height          =   255
               Left            =   1470
               TabIndex        =   17
               Top             =   90
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "上线"
            End
            Begin Threed.SSOption opt_offline 
               Height          =   255
               Left            =   420
               TabIndex        =   18
               Top             =   90
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "下线"
               Value           =   -1
            End
         End
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   150
            Top             =   540
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            Caption         =   "查询号"
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
            Left            =   5070
            Top             =   540
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            Caption         =   "分段号"
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
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   5070
            Top             =   120
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            Caption         =   "班次/别"
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
         Begin InDate.UDate SDT_PROD_DATE_FROM 
            Height          =   315
            Left            =   1470
            TabIndex        =   24
            Tag             =   "起始日期"
            Top             =   120
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.74
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            BackColor       =   16777215
         End
         Begin InDate.UDate SDT_PROD_DATE_TO 
            Height          =   315
            Left            =   3225
            TabIndex        =   25
            Tag             =   "起始日期"
            Top             =   120
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.74
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            BackColor       =   16777215
         End
         Begin InDate.ULabel ULabel27 
            Height          =   315
            Left            =   150
            Top             =   120
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            Caption         =   "生产日期"
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
         Begin Threed.SSFrame SSFrame5 
            Height          =   315
            Left            =   10110
            TabIndex        =   26
            Top             =   120
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line1 
               Height          =   255
               Left            =   300
               TabIndex        =   27
               Top             =   30
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "# 1"
               Value           =   -1
            End
            Begin Threed.SSOption opt_line2 
               Height          =   255
               Left            =   1320
               TabIndex        =   28
               Top             =   30
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "# 2"
            End
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   8550
            Top             =   120
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            Caption         =   "剪切线    "
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
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   3045
            TabIndex        =   29
            Top             =   240
            Width           =   195
         End
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   405
      Left            =   15510
      TabIndex        =   30
      Top             =   1050
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   714
      _Version        =   196609
      BackColor       =   14737632
      Begin Threed.SSOption opt_p_m 
         Height          =   255
         Left            =   720
         TabIndex        =   31
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "操作"
         Value           =   -1
      End
      Begin Threed.SSOption opt_p_r 
         Height          =   255
         Left            =   2070
         TabIndex        =   32
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "查询"
      End
   End
End
Attribute VB_Name = "CGD2037C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   Mill System
'-- Program Name      精整区域上、下线实绩处理界面
'-- Program ID        CGD2037C
'-- Document No       Q-00-0010(Specification)
'-- Designer          杨猛
'-- Coder             杨猛
'-- Date              2011.02.21
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE        EDITOR       DESCRIPTION
'-- 1.01  2011.02.21  杨猛         精整区域上、下线实绩处理界面
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting
Public sDateTime As String          'Active Form Time Setting
Public sQuery_Rt As String          'Active Form sQuery Setting
       
Dim pControl1 As New Collection      'Master Primary Key Collection
Dim nControl1 As New Collection      'Master Necessary Collection
Dim mControl1 As New Collection      'Master Maxlength check Collection
Dim iControl1 As New Collection      'Master Insert Collection
Dim rControl1 As New Collection      'Master Refer Collection
Dim cControl1 As New Collection      'Master Copy Collection
Dim aControl1 As New Collection      'Master -> Spread Collection
Dim lControl1 As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collectionn

Dim sc1 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SPD_PLATE_NO = 1
Const SPD_OCCR_TIME = 2
Const SPD_LOT_NO = 3
Const SPD_SEQ = 4
Const SPD_LINE_DATE = 9
Const SPD_OFF_PRC = 10
Const SPD_OFF_DATE = 11
Const SPD_OFF_REASON = 12
Const SPD_OFF_LOC = 13
Const SPD_OFF_EMP = 14
Const SPD_ON_PRC = 15
Const SPD_ON_DATE = 16
Const SPD_ON_REASON = 17
Const SPD_ON_EMP = 18
Const SPD_ONOFF_CD = 19



Private Sub Form_Define()
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

    ' Call Master_Collection("CONTROL1_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
             
            Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
               Call Gp_Ms_Collection(TXT_SEQ, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_line, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             Call Gp_Ms_Collection(cbo_group, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
               Call Gp_Ms_Collection(txt_m_r, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             Call Gp_Ms_Collection(txt_onoff, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
               Call Gp_Ms_Collection(txt_prc, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             
    'MASTER Collection
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"
   
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
     'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGD2037C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="CGD2037C.P_REFER", Key:="P-R"
    sc1.Add Item:="CGD2037C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, SPD_OCCR_TIME, True)
    Call Gp_Sp_ColHidden(ss1, SPD_LINE_DATE, True)
    Call Gp_Sp_ColHidden(ss1, SPD_OFF_REASON, True)
    Call Gp_Sp_ColHidden(ss1, SPD_OFF_LOC, True)
'    Call Gp_Sp_ColHidden(ss1, SPD_ON_PRC, True)
'    Call Gp_Sp_ColHidden(ss1, SPD_ON_DATE, True)
    Call Gp_Sp_ColHidden(ss1, SPD_ON_REASON, True)
'    Call Gp_Sp_ColHidden(ss1, SPD_ON_EMP, True)
    Call Gp_Sp_ColHidden(ss1, SPD_ONOFF_CD, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
     
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

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Mill_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    
    opt_offline.Value = True
    opt_OnPosition(0).Value = True
    
    CBO_SHIFT.AddItem "1"
    CBO_SHIFT.AddItem "2"
    CBO_SHIFT.AddItem "3"
    
    cbo_group.AddItem "A"
    cbo_group.AddItem "B"
    cbo_group.AddItem "C"
    cbo_group.AddItem "D"

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "CG-System.INI", Me.Name)

    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set Mc1 = Nothing
    
    Set sc1 = Nothing

    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gf_Sp_Cls(sc1)
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = False                'Excel
    End With

End Sub



Public Sub Form_Ref()

    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
End Sub

Public Sub Form_Pro()
    
    If Gf_Sp_Process(M_CN1, sc1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If

End Sub

Public Sub Form_Del()

'    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))

End Sub

Private Sub opt_line1_Click(Value As Integer)
   txt_line.Text = "1"
   If ss1.MaxRows > 0 Then Call Form_Ref
End Sub

Private Sub opt_line2_Click(Value As Integer)
   txt_line.Text = "2"
   If ss1.MaxRows > 0 Then Call Form_Ref
End Sub

Private Sub opt_offline_Click(Value As Integer)

    If opt_offline Then
        txt_onoff = "F"
        opt_offline.ForeColor = &HFF&
        opt_online.ForeColor = &H80000012
                   
        If ss1.MaxRows > 0 Then
           Call Form_Ref
        End If
           
'        Call Gf_Sp_Cls(sc1)
        
'        Call Gp_Sp_ColHidden(ss1, SPD_ON_PRC, True)
'        Call Gp_Sp_ColHidden(ss1, SPD_ON_DATE, True)
'        Call Gp_Sp_ColHidden(ss1, SPD_ON_REASON, True)
'        Call Gp_Sp_ColHidden(ss1, SPD_ON_EMP, True)
'        Call Gp_Sp_ColHidden(ss1, SPD_OFF_PRC, False)
'        Call Gp_Sp_ColHidden(ss1, SPD_OFF_DATE, False)
'        Call Gp_Sp_ColHidden(ss1, SPD_OFF_REASON, False)
'        Call Gp_Sp_ColHidden(ss1, SPD_OFF_LOC, False)
'        Call Gp_Sp_ColHidden(ss1, SPD_OFF_EMP, False)

    End If
End Sub

Private Sub opt_online_Click(Value As Integer)
    If opt_online Then
        txt_onoff = "N"
        opt_online.ForeColor = &HFF&
        opt_offline.ForeColor = &H80000012
        If ss1.MaxRows > 0 Then
           Call Form_Ref
        End If
'        Call Gf_Sp_Cls(sc1)
        
'        Call Gp_Sp_ColHidden(ss1, SPD_ON_PRC, False)
'        Call Gp_Sp_ColHidden(ss1, SPD_ON_DATE, False)
'        Call Gp_Sp_ColHidden(ss1, SPD_ON_REASON, False)
'        Call Gp_Sp_ColHidden(ss1, SPD_ON_EMP, False)
'        Call Gp_Sp_ColHidden(ss1, SPD_OFF_PRC, False)
'        Call Gp_Sp_ColHidden(ss1, SPD_OFF_DATE, False)
'        Call Gp_Sp_ColHidden(ss1, SPD_OFF_REASON, False)
'        Call Gp_Sp_ColHidden(ss1, SPD_OFF_LOC, False)
'        Call Gp_Sp_ColHidden(ss1, SPD_OFF_EMP, False)
    End If
End Sub

Private Sub opt_OnPosition_Click(Index As Integer, Value As Integer)
    
    txt_prc = opt_OnPosition(Index).Tag
    If opt_OnPosition(0).Value = True Then
       opt_OnPosition(0).ForeColor = &HFF&       'red
       opt_OnPosition(1).ForeColor = &H80000012  'black
       opt_OnPosition(2).ForeColor = &H80000012  'black
       opt_OnPosition(3).ForeColor = &H80000012  'black
       opt_OnPosition(4).ForeColor = &H80000012  'black
       opt_OnPosition(5).ForeColor = &H80000012  'black
       opt_OnPosition(6).ForeColor = &H80000012  'black
    ElseIf opt_OnPosition(1).Value = True Then
       opt_OnPosition(0).ForeColor = &H80000012       'red
       opt_OnPosition(1).ForeColor = &HFF&  'black
       opt_OnPosition(2).ForeColor = &H80000012  'black
       opt_OnPosition(3).ForeColor = &H80000012  'black
       opt_OnPosition(4).ForeColor = &H80000012  'black
       opt_OnPosition(5).ForeColor = &H80000012  'black
       opt_OnPosition(6).ForeColor = &H80000012  'black
    ElseIf opt_OnPosition(2).Value = True Then
       opt_OnPosition(0).ForeColor = &H80000012       'red
       opt_OnPosition(1).ForeColor = &H80000012  'black
       opt_OnPosition(2).ForeColor = &HFF&  'black
       opt_OnPosition(3).ForeColor = &H80000012  'black
       opt_OnPosition(4).ForeColor = &H80000012  'black
       opt_OnPosition(5).ForeColor = &H80000012  'black
       opt_OnPosition(6).ForeColor = &H80000012  'black
    ElseIf opt_OnPosition(3).Value = True Then
       opt_OnPosition(0).ForeColor = &H80000012       'red
       opt_OnPosition(1).ForeColor = &H80000012  'black
       opt_OnPosition(2).ForeColor = &H80000012  'black
       opt_OnPosition(3).ForeColor = &HFF&  'black
       opt_OnPosition(4).ForeColor = &H80000012  'black
       opt_OnPosition(5).ForeColor = &H80000012  'black
       opt_OnPosition(6).ForeColor = &H80000012  'black
    ElseIf opt_OnPosition(4).Value = True Then
       opt_OnPosition(0).ForeColor = &H80000012       'red
       opt_OnPosition(1).ForeColor = &H80000012  'black
       opt_OnPosition(2).ForeColor = &H80000012  'black
       opt_OnPosition(3).ForeColor = &H80000012  'black
       opt_OnPosition(4).ForeColor = &HFF&  'black
       opt_OnPosition(5).ForeColor = &H80000012  'black
       opt_OnPosition(6).ForeColor = &H80000012  'black
    ElseIf opt_OnPosition(5).Value = True Then
       opt_OnPosition(0).ForeColor = &H80000012       'red
       opt_OnPosition(1).ForeColor = &H80000012  'black
       opt_OnPosition(2).ForeColor = &H80000012  'black
       opt_OnPosition(3).ForeColor = &H80000012  'black
       opt_OnPosition(4).ForeColor = &H80000012  'black
       opt_OnPosition(5).ForeColor = &HFF&  'black
       opt_OnPosition(6).ForeColor = &H80000012  'black
    ElseIf opt_OnPosition(6).Value = True Then
       opt_OnPosition(0).ForeColor = &H80000012       'red
       opt_OnPosition(1).ForeColor = &H80000012  'black
       opt_OnPosition(2).ForeColor = &H80000012  'black
       opt_OnPosition(3).ForeColor = &H80000012  'black
       opt_OnPosition(4).ForeColor = &H80000012  'black
       opt_OnPosition(5).ForeColor = &H80000012  'black
       opt_OnPosition(6).ForeColor = &HFF&  'black
    End If
End Sub

Private Sub opt_p_m_Click(Value As Integer)
    If opt_p_m Then
        txt_m_r = "M"
        opt_p_m.ForeColor = &HFF&
        opt_p_r.ForeColor = &H80000012
        Call Gf_Sp_Cls(sc1)
    End If
End Sub
Private Sub opt_p_r_Click(Value As Integer)
    If opt_p_r Then
        txt_m_r = "R"
        opt_p_r.ForeColor = &HFF&
        opt_p_m.ForeColor = &H80000012
        Call Gf_Sp_Cls(sc1)
    End If
End Sub

Private Sub SDT_PROD_DATE_FROM_GotFocus()
     If SDT_PROD_DATE_FROM.RawData = "" Then
        SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
     End If
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub SDT_PROD_DATE_TO_GotFocus()
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    
    If ROW < 1 Then Exit Sub
    If ss1.MaxRows < 1 Then Exit Sub
    If Col = 0 Then
       Call ss1_row_Click(Col, ROW)
       If txt_onoff.Text = "F" Then
          ss1.ROW = ROW:  ss1.Col = SPD_OFF_PRC:   ss1.Text = txt_prc.Text
          ss1.ROW = ROW:  ss1.Col = SPD_OFF_EMP:   ss1.Text = sUserID
       Else
          ss1.ROW = ROW:  ss1.Col = SPD_ON_PRC:    ss1.Text = txt_prc.Text
          ss1.ROW = ROW:  ss1.Col = SPD_ON_EMP:    ss1.Text = sUserID
       End If
       ss1.ROW = ROW:  ss1.Col = SPD_ONOFF_CD:  ss1.Text = txt_onoff.Text
    End If
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim iCol As Long
    Dim iRow As Long
    Dim iMode As Integer
    
    Dim iRowNum As Long
    
    iCol = Col
    iRow = ROW
    iMode = Mode

    If ROW <= 0 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") Then
    
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), iMode)
        
        ss1.ROW = iRow
        If txt_onoff = "F" Then
           ss1.Col = SPD_OFF_PRC:           ss1.Text = txt_prc.Text
           ss1.Col = SPD_OFF_EMP:           ss1.Text = sUserID
        Else
           ss1.Col = SPD_ON_PRC:            ss1.Text = txt_prc.Text
           ss1.Col = SPD_ON_EMP:            ss1.Text = sUserID
        End If
        
        ss1.Col = SPD_ONOFF_CD:            ss1.Text = txt_onoff.Text
        
    End If

End Sub

Private Sub txt_plt_DblClick()
    Call txt_plt_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=txt_plt_dec

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

    If Len(Trim(txt_plt)) = txt_plt.MaxLength Then
        txt_plt_dec.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        txt_plt_dec.Text = ""
    End If
End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim sTemp_Code As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss1.ActiveCol
    
        Case SPD_ON_REASON
        
             If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss1
                
                DD.sWitch = "SP"
                DD.sKey = "G0031"
                DD.rControl.Add Item:=SPD_ON_REASON
                
                DD.nameType = "2"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
              End If
       
        Case SPD_OFF_REASON
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss1
                
                DD.sWitch = "SP"
                DD.sKey = "G0031"
                DD.rControl.Add Item:=SPD_OFF_REASON
'                DD.rControl.Add Item:=6
                
                DD.nameType = "2"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
            
            End If
          
    End Select

End Sub

Private Sub ss1_row_Click(ByVal Col As Long, ByVal ROW As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If ROW < 1 Or ROW = ss1.MaxRows Then Exit Sub
    If ss1.MaxRows < 1 Then Exit Sub
    
    ss1.ROW = ROW
    ss1.Col = 0
    
    ss1.ReDraw = False
    If ss1.Text <> "Update" Then
        
        ss1.Text = "Update"
        
        Call Gp_Sp_BlockColor(ss1, 1, -1, ROW, ROW, , &HFFFF80)
    Else
       
        ss1.Text = ""
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW)
       
    End If
    ss1.ReDraw = True
    
End Sub

