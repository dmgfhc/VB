VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2011C 
   Caption         =   "冷床实绩查询及修改界面_AGC2011C"
   ClientHeight    =   9210
   ClientLeft      =   450
   ClientTop       =   2340
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8325
      Left            =   60
      TabIndex        =   7
      Top             =   840
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   14684
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "AGC2011C.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   7665
         Left            =   0
         TabIndex        =   24
         Top             =   660
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   13520
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "上冷床"
         TabPicture(0)   =   "AGC2011C.frx":0052
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSPanel7"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "下冷床"
         TabPicture(1)   =   "AGC2011C.frx":006E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSPanel8"
         Tab(1).ControlCount=   1
         Begin Threed.SSPanel SSPanel7 
            Height          =   7215
            Left            =   60
            TabIndex        =   25
            Top             =   360
            Width           =   15165
            _ExtentX        =   26749
            _ExtentY        =   12726
            _Version        =   196609
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin FPSpread.vaSpread ss1 
               Height          =   7215
               Left            =   0
               TabIndex        =   28
               Top             =   0
               Width           =   15135
               _Version        =   393216
               _ExtentX        =   26696
               _ExtentY        =   12726
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   16
               MaxRows         =   2
               RetainSelBlock  =   0   'False
               SpreadDesigner  =   "AGC2011C.frx":008A
            End
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   7215
            Left            =   -74940
            TabIndex        =   26
            Top             =   360
            Width           =   15165
            _ExtentX        =   26749
            _ExtentY        =   12726
            _Version        =   196609
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin FPSpread.vaSpread ss2 
               Height          =   7215
               Left            =   0
               TabIndex        =   27
               Top             =   0
               Width           =   15135
               _Version        =   393216
               _ExtentX        =   26696
               _ExtentY        =   12726
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   13
               MaxRows         =   2
               RetainSelBlock  =   0   'False
               SpreadDesigner  =   "AGC2011C.frx":0AAB
            End
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   630
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   1111
         _Version        =   196609
         BackColor       =   14737918
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txt_bed_indic 
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
            Height          =   315
            Left            =   840
            TabIndex        =   11
            Tag             =   "冷床"
            Text            =   " "
            Top             =   450
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox txt_pos 
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
            Height          =   315
            Left            =   1320
            TabIndex        =   10
            Tag             =   "布料方式"
            Text            =   " "
            Top             =   450
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txt_bed_fl 
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
            Height          =   315
            Left            =   360
            TabIndex        =   9
            Tag             =   "上/下冷床"
            Text            =   " "
            Top             =   450
            Visible         =   0   'False
            Width           =   420
         End
         Begin InDate.ULabel ULabel28 
            Height          =   315
            Left            =   4620
            Top             =   150
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            Caption         =   "布料方式"
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   375
            Left            =   1920
            TabIndex        =   12
            Top             =   150
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   661
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSOption opt_bed1 
               Height          =   285
               Left            =   30
               TabIndex        =   13
               Top             =   30
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   14737918
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "一号"
            End
            Begin Threed.SSOption opt_bed2 
               Height          =   285
               Left            =   870
               TabIndex        =   14
               Top             =   30
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               BackColor       =   14737918
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "二号"
            End
            Begin Threed.SSOption opt_bed3 
               Height          =   285
               Left            =   1740
               TabIndex        =   15
               Top             =   30
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "三号"
               Value           =   -1
            End
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   375
            Left            =   6210
            TabIndex        =   16
            Top             =   150
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   661
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSOption opt_pos_l 
               Height          =   285
               Left            =   30
               TabIndex        =   17
               Top             =   30
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "左料"
            End
            Begin Threed.SSOption opt_pos_r 
               Height          =   285
               Left            =   870
               TabIndex        =   18
               Top             =   30
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "右料"
            End
            Begin Threed.SSOption opt_pos_c 
               Height          =   285
               Left            =   1740
               TabIndex        =   19
               Top             =   30
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "双排"
               Value           =   -1
            End
         End
         Begin InDate.ULabel ULabel26 
            Height          =   315
            Left            =   360
            Top             =   150
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            Caption         =   "冷床"
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
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   8910
            Top             =   150
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            Caption         =   "冷床时间"
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
         Begin InDate.ULabel ULabel13 
            Height          =   315
            Left            =   12780
            Top             =   150
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            Caption         =   "冷床温度"
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
         Begin CSTextLibCtl.sitxEdit stx_cb_time 
            Height          =   315
            Left            =   10425
            TabIndex        =   20
            Top             =   150
            Width           =   2175
            _Version        =   262145
            _ExtentX        =   3836
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
            ValidateMask    =   0   'False
         End
         Begin CSTextLibCtl.sidbEdit sdb_cd_temp 
            Height          =   315
            Left            =   14310
            TabIndex        =   21
            Top             =   150
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
            NumIntDigits    =   4
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   315
            Left            =   1980
            TabIndex        =   22
            Top             =   150
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin CSTextLibCtl.sitxEdit stx_occr_date 
            Height          =   315
            Left            =   1830
            TabIndex        =   23
            Top             =   450
            Visible         =   0   'False
            Width           =   2175
            _Version        =   262145
            _ExtentX        =   3836
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
            Enabled         =   0   'False
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
            ValidateMask    =   0   'False
         End
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15300
      _ExtentX        =   26988
      _ExtentY        =   1296
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_onoff 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   12150
         MaxLength       =   1
         TabIndex        =   29
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txt_emp_cd 
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
         Left            =   14760
         MaxLength       =   7
         TabIndex        =   6
         Tag             =   "作业人员"
         Top             =   600
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.ComboBox cbo_group 
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
         ItemData        =   "AGC2011C.frx":12AB
         Left            =   14190
         List            =   "AGC2011C.frx":12BB
         TabIndex        =   5
         Tag             =   "班别"
         Top             =   600
         Visible         =   0   'False
         Width           =   585
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
         ItemData        =   "AGC2011C.frx":12CB
         Left            =   14130
         List            =   "AGC2011C.frx":12D8
         TabIndex        =   4
         Tag             =   "班次"
         Top             =   210
         Width           =   735
      End
      Begin VB.ComboBox cbo_plt 
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
         ItemData        =   "AGC2011C.frx":12E5
         Left            =   13065
         List            =   "AGC2011C.frx":12EF
         TabIndex        =   3
         Tag             =   "工厂"
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txt_plate_no 
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
         Left            =   5535
         MaxLength       =   14
         TabIndex        =   2
         Tag             =   "母板号"
         Top             =   210
         Width           =   1635
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
         Height          =   315
         Left            =   12630
         TabIndex        =   1
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   435
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   4020
         Top             =   210
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         Caption         =   "母板号"
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
         Left            =   7710
         Top             =   210
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         Caption         =   "生产日期"
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
      Begin Threed.SSOption opt_on 
         Height          =   285
         Left            =   1950
         TabIndex        =   30
         Top             =   240
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
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
         Caption         =   "在线"
         Value           =   -1
      End
      Begin Threed.SSOption opt_off 
         Height          =   285
         Left            =   2820
         TabIndex        =   31
         Top             =   240
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
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
         Caption         =   "离线"
      End
      Begin InDate.UDate udt_date_fr 
         Height          =   315
         Left            =   9225
         TabIndex        =   32
         Tag             =   "INS_DATE"
         Top             =   210
         Width           =   1440
         _ExtentX        =   2540
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
         MaxLength       =   10
      End
      Begin InDate.UDate udt_date_to 
         Height          =   315
         Left            =   10680
         TabIndex        =   33
         Tag             =   "INS_DATE"
         Top             =   210
         Width           =   1440
         _ExtentX        =   2540
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
         MaxLength       =   10
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   360
         Top             =   210
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         Caption         =   "在/离线"
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
      Begin InDate.ULabel ULabel31 
         Height          =   315
         Left            =   12600
         Top             =   210
         Width           =   1485
         _ExtentX        =   2619
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
   End
End
Attribute VB_Name = "AGC2011C"
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
'-- Program Name      冷床实绩查询及修改界面
'-- Program ID        AGC2011C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM SUNG HO
'-- Coder             KIM SUNG HO
'-- Date              2010.7.12
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

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection
 
Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection

Dim opt_chk As Boolean
Dim lMain_Row As Integer

Private Sub Form_Define()

    Dim iCol As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(txt_plate_no, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(CBO_PLT, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_onoff, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(udt_date_fr, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(udt_date_to, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_PRC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(cbo_group, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TXT_EMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(stx_occr_date, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_bed_fl, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_bed_indic, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_pos, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(stx_cb_time, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_cd_temp, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
    'MASTER Collection
    Mc1.Add Item:="AGC2011C.P_REFER", Key:="P-R"
    Mc1.Add Item:="AGC2011C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
      
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
   
     'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2011C.P_SREFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss2.MaxCols
        Call Gp_Sp_Collection(ss2, iCol, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next iCol
   
     'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AGC2011C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss2, 7, True)
    Call Gp_Sp_ColHidden(ss2, ss2.MaxCols, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    If txt_plate_no.Text <> "" Then
        Call MenuTool_ReSet(True)
    Else
        Call MenuTool_ReSet(False)
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Dim sQuery As String
    
    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet(False)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    txt_onoff.Text = "I"
    TXT_EMP_CD.Text = sUserID
    txt_bed_fl.Text = "1"
    opt_pos_c.Value = True
    txt_bed_indic.Text = "30"
    txt_pos.Text = "1"
    TXT_PRC.Text = "CE"
    CBO_PLT.Text = "C1"
    CBO_SHIFT.Text = ""
    opt_chk = True
    udt_date_fr.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
    udt_date_to.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
    SSTab1.Tab = 0
    ULabel12.Caption = "上冷床时间"
    ULabel13.Caption = "上冷床温度"
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Gp_Sp_ReadOnlySet(ss1)
    Call Gp_Sp_ReadOnlySet(ss2)
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "G-System.INI", Me.Name)

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "G-System.INI", Me.Name)

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
    
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
        
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call MenuTool_ReSet(False)
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
    opt_on.Value = True
    TXT_EMP_CD.Text = sUserID
    CBO_PLT.Text = "C1"
    TXT_PRC.Text = "CE"
    CBO_SHIFT.Text = ""
    opt_bed3.Value = True
    opt_pos_l.Value = True
    udt_date_fr.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
    udt_date_to.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
    lMain_Row = 0
    
    If opt_chk Then
        SSTab1.Tab = 0
        Call Gf_Sp_Cls(sc1)
        Call Gf_Sp_Cls(sc2)
    End If
        
End Sub

Public Sub Form_Exc()

    If SSTab1.Tab = 0 Then
        Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Else
        Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If
    
End Sub

Public Sub Master_Cpy()

End Sub

Public Sub Master_Pst()

End Sub

Public Sub Form_Ref()

    Dim sPlate_No As String
    
    If SSTab1.Tab = 0 Then
        
        If Gf_Sp_Refer(M_CN1, sc1, Mc1) Then
            ss1.OperationMode = OperationModeNormal
            Call Gp_Sp_EvenRowBackcolor(ss1)
        End If
    
    Else
        
        If Gf_Sp_Refer(M_CN1, sc2, Mc1) Then
            ss2.OperationMode = OperationModeNormal
            Call Gp_Sp_EvenRowBackcolor(ss2)
        End If
    
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call MenuTool_ReSet(False)
    lMain_Row = 0
    
End Sub

Private Sub opt_bed1_Click(Value As Integer)

'    If opt_bed1.Value Then
'        opt_bed1.ForeColor = &HFF&
'        opt_bed2.ForeColor = &H80000012
'        opt_bed3.ForeColor = &H80000012
'        txt_bed_indic.Text = "10"
'
'        If SSTab1.Tab = 1 And opt_chk Then
'
'            txt_plate_no.Text = ""
'            opt_pos_l.Value = True
'            stx_occr_date.RawData = ""
'
'            If Gf_Sp_Refer(M_CN1, sc2, Mc1, , , False) Then
'                ss2.OperationMode = OperationModeNormal
'                Call Gp_Sp_EvenRowBackcolor(ss2)
'            End If
'        End If
'
'    End If
    
End Sub

Private Sub opt_bed2_Click(Value As Integer)

'    If opt_bed2.Value Then
'        opt_bed1.ForeColor = &H80000012
'        opt_bed2.ForeColor = &HFF&
'        opt_bed3.ForeColor = &H80000012
'        txt_bed_indic.Text = "20"
'
'        If SSTab1.Tab = 1 And opt_chk Then
'
'            txt_plate_no.Text = ""
'            opt_pos_l.Value = True
'            stx_occr_date.RawData = ""
'
'            If Gf_Sp_Refer(M_CN1, sc2, Mc1, , , False) Then
'                ss2.OperationMode = OperationModeNormal
'                Call Gp_Sp_EvenRowBackcolor(ss2)
'            End If
'        End If
'
'    End If

End Sub

Private Sub opt_bed3_Click(Value As Integer)

    If opt_bed3.Value Then
        opt_bed1.ForeColor = &H80000012
        opt_bed2.ForeColor = &H80000012
        opt_bed3.ForeColor = &HFF&
        txt_bed_indic.Text = "30"

        If SSTab1.Tab = 1 And opt_chk Then

            txt_plate_no.Text = ""
            opt_pos_l.Value = True
            stx_occr_date.RawData = ""

            If Gf_Sp_Refer(M_CN1, sc2, Mc1, , , False) Then
                ss2.OperationMode = OperationModeNormal
                Call Gp_Sp_EvenRowBackcolor(ss2)
            End If
        End If

    End If

End Sub

Private Sub opt_off_Click(Value As Integer)

    If opt_off.Value Then
        opt_off.ForeColor = &HFF&
        opt_on.ForeColor = &H80000012
        txt_onoff.Text = "O"
    End If

End Sub

Private Sub opt_on_Click(Value As Integer)

    If opt_on.Value Then
        opt_on.ForeColor = &HFF&
        opt_off.ForeColor = &H80000012
        txt_onoff.Text = "I"
    End If
    
End Sub

Private Sub opt_pos_l_Click(Value As Integer)

    If opt_pos_l.Value Then
        opt_pos_l.ForeColor = &HFF&
        opt_pos_r.ForeColor = &H80000012
        opt_pos_c.ForeColor = &H80000012
        txt_pos.Text = "1"
    End If
    
End Sub

Private Sub opt_pos_r_Click(Value As Integer)

    If opt_pos_r.Value Then
        opt_pos_l.ForeColor = &H80000012
        opt_pos_r.ForeColor = &HFF&
        opt_pos_c.ForeColor = &H80000012
        txt_pos.Text = "2"
    End If

End Sub

Private Sub opt_pos_c_Click(Value As Integer)

    If opt_pos_c.Value Then
        opt_pos_l.ForeColor = &H80000012
        opt_pos_r.ForeColor = &H80000012
        opt_pos_c.ForeColor = &HFF&
        txt_pos.Text = "3"
    End If

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If ss1.MaxRows < 1 Or Row = 0 Then Exit Sub
    
    If lMain_Row <> 0 Then
    
        ss1.Row = lMain_Row
        ss1.Col = 0
        ss1.Text = ""
        
        If lMain_Row Mod 2 <> 0 Then
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lMain_Row, lMain_Row, , &HF2F2F2)
        Else
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lMain_Row, lMain_Row, , &HFFFFFF)
        End If
        
    End If
    
    lMain_Row = Row
    ss1.Row = Row
    ss1.Col = 0
    ss1.Text = "选择"
    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.Row, ss1.Row, , CYAN)
    
    ss1.Col = 1
    txt_plate_no.Text = ss1.Text
    
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If ss2.MaxRows < 1 Or Row = 0 Then Exit Sub
    
    If lMain_Row <> 0 Then
    
        ss2.Row = lMain_Row
        ss2.Col = 0
        ss2.Text = ""
        
        If lMain_Row Mod 2 <> 0 Then
            Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, lMain_Row, lMain_Row, , &HF2F2F2)
        Else
            Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, lMain_Row, lMain_Row, , &HFFFFFF)
        End If
        
    End If
    
    lMain_Row = Row
    ss2.Row = Row
    ss2.Col = 0
    ss2.Text = "选择"
    Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, ss2.Row, ss2.Row, , CYAN)
    
    ss2.Col = 1
    txt_plate_no.Text = ss2.Text
        
    opt_chk = False
        
    ss2.Col = 5
    If ss2.Text = "一号" Then
        opt_bed1.Value = True
    ElseIf ss2.Text = "二号" Then
        opt_bed2.Value = True
    ElseIf ss2.Text = "三号" Or ss2.Text = "" Then
        opt_bed3.Value = True
    End If
    
    opt_chk = True
    
    ss2.Col = 6
    If ss2.Text = "左料" Or ss2.Text = "" Then
        opt_pos_l.Value = True
    ElseIf ss2.Text = "右料" Then
        opt_pos_r.Value = True
    ElseIf ss2.Text = "双排" Then
        opt_pos_c.Value = True
    End If
        
    ss2.Col = 7
    If ss2.Text <> "" Then
        stx_occr_date.RawData = ss2.Value
        Call MenuTool_ReSet(True)
    Else
        stx_occr_date.RawData = ""
        Call MenuTool_ReSet(False)
    End If
        
    
End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Public Sub Form_Pro()

    Dim sPlt As String
    Dim sPrc As String
    Dim sShift As String
    Dim sGroup As String
    Dim sEmp_Cd As String
    Dim Bed_Fl As String
    Dim Bed_indic As String
    
    If txt_plate_no.Text = "" Or Len(txt_plate_no.Text) <> 12 Then Exit Sub
    
    txt_plate_no.Enabled = False
    udt_date_fr.Enabled = False
    udt_date_to.Enabled = False
    
    If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
    
        Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
        Call MenuTool_ReSet(False)
        
        sPlt = CBO_PLT.Text
        sPrc = TXT_PRC.Text
        sShift = CBO_SHIFT.Text
        sGroup = cbo_group.Text
        sEmp_Cd = TXT_EMP_CD.Text
        Bed_Fl = txt_bed_fl.Text
        Bed_indic = txt_bed_indic.Text
        txt_plate_no.Text = ""
        
        If SSTab1.Tab = 0 Then
            Call Gf_Sp_Refer(M_CN1, sc1, Mc1, , , False)
            ss1.OperationMode = OperationModeNormal
            Call Gp_Sp_EvenRowBackcolor(ss1)
            Call Gp_Ms_Cls(Mc1("rControl"))
        Else
            Call Gf_Sp_Refer(M_CN1, sc2, Mc1, , , False)
            ss2.OperationMode = OperationModeNormal
            Call Gp_Sp_EvenRowBackcolor(ss2)
            Call Gp_Ms_Cls(Mc1("rControl"))
        End If
        
        lMain_Row = 0
        'txt_plate_no.Text = ""
        CBO_PLT.Text = sPlt
        TXT_PRC.Text = sPrc
        CBO_SHIFT.Text = sShift
        cbo_group.Text = sGroup
        TXT_EMP_CD.Text = sEmp_Cd
        txt_bed_fl.Text = Bed_Fl
        txt_bed_indic.Text = Bed_indic
        opt_bed3.Value = True
        opt_pos_l.Value = True
        
    End If
    
    txt_plate_no.Enabled = True
    udt_date_fr.Enabled = True
    udt_date_to.Enabled = True
 
End Sub

Public Sub Form_Del()

    Dim sPlt As String
    Dim sPrc As String
    Dim sShift As String
    Dim sGroup As String
    Dim sEmp_Cd As String
    Dim Bed_Fl As String
    Dim Bed_indic As String
    
    If stx_occr_date.RawData = "" Then Exit Sub
    
    txt_plate_no.Enabled = False
    udt_date_fr.Enabled = False
    udt_date_to.Enabled = False
    CBO_SHIFT.Enabled = False
    CBO_PLT.Enabled = False
    txt_onoff.Enabled = False
    txt_bed_fl.Enabled = False
    txt_bed_indic.Enabled = False
    
    If Gf_Ms_Del(M_CN1, Mc1) Then
        
        Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
        Call MenuTool_ReSet(False)
    
        sPlt = CBO_PLT.Text
        sPrc = TXT_PRC.Text
        sShift = CBO_SHIFT.Text
        sGroup = cbo_group.Text
        sEmp_Cd = TXT_EMP_CD.Text
        Bed_Fl = txt_bed_fl.Text
        Bed_indic = txt_bed_indic.Text
        txt_plate_no.Text = ""
        
        If SSTab1.Tab = 0 Then
        
            Call Gf_Sp_Refer(M_CN1, sc1, Mc1, , , False)
            ss1.OperationMode = OperationModeNormal
            Call Gp_Sp_EvenRowBackcolor(ss1)
            Call Gp_Ms_Cls(Mc1("rControl"))
            opt_bed1.Value = True
            
        Else
        
            Call Gf_Sp_Refer(M_CN1, sc2, Mc1, , , False)
            ss2.OperationMode = OperationModeNormal
            Call Gp_Sp_EvenRowBackcolor(ss2)
            
            Call Gp_Ms_Cls(Mc1("rControl"))
            
            opt_chk = False
            
            If Bed_indic = "10" Then
                opt_bed1.Value = True
            ElseIf Bed_indic = "20" Then
                opt_bed2.Value = True
            Else
                opt_bed3.Value = True
            End If
            
            opt_chk = True
            
        End If
        
        lMain_Row = 0
        CBO_PLT.Text = sPlt
        TXT_PRC.Text = sPrc
        CBO_SHIFT.Text = sShift
        cbo_group.Text = sGroup
        TXT_EMP_CD.Text = sEmp_Cd
        txt_bed_fl.Text = Bed_Fl
        opt_pos_l.Value = True
        
    End If
    
    txt_plate_no.Enabled = True
    udt_date_fr.Enabled = True
    udt_date_to.Enabled = True
    CBO_SHIFT.Enabled = True
    CBO_PLT.Enabled = True
    txt_onoff.Enabled = True
    txt_bed_fl.Enabled = True
    txt_bed_indic.Enabled = True
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

    Dim sOnoff As String
    Dim sDate_Fr As String
    Dim sDate_To As String
    
    If SSTab1.Tab = 0 Then
        
        If lMain_Row <> 0 Then
            If lMain_Row Mod 2 <> 0 Then
                Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, lMain_Row, lMain_Row, , &HF2F2F2)
            Else
                Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, lMain_Row, lMain_Row, , &HFFFFFF)
            End If
            
            ss2.Row = lMain_Row
            ss2.Col = 0
            ss2.Text = ""
        End If
        
        sOnoff = txt_onoff.Text
        sDate_Fr = udt_date_fr.RawData
        sDate_To = udt_date_to.RawData
        
        opt_chk = False
        Call Form_Cls
        opt_chk = True
        
        If sOnoff = "I" Then
            opt_on.Value = True
        Else
            opt_off.Value = True
        End If
        
        udt_date_fr.RawData = sDate_Fr
        udt_date_to.RawData = sDate_To
        
        txt_bed_fl.Text = "1"
        ULabel12.Caption = "上冷床时间"
        ULabel13.Caption = "上冷床温度"
        
        If Gf_Sp_Refer(M_CN1, sc1, Mc1, , , False) Then
            ss1.OperationMode = OperationModeNormal
            Call Gp_Sp_EvenRowBackcolor(ss1)
        End If
        
        Call Gf_Sp_Cls(sc2)
        Call MenuTool_ReSet(False)
        
    Else
        
        If lMain_Row <> 0 Then
            If lMain_Row Mod 2 <> 0 Then
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lMain_Row, lMain_Row, , &HF2F2F2)
            Else
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lMain_Row, lMain_Row, , &HFFFFFF)
            End If
            
            ss1.Row = lMain_Row
            ss1.Col = 0
            ss1.Text = ""
        End If
        
        sOnoff = txt_onoff.Text
        sDate_Fr = udt_date_fr.RawData
        sDate_To = udt_date_to.RawData
        
        opt_chk = False
        Call Form_Cls
        opt_chk = True
        
        If sOnoff = "I" Then
            opt_on.Value = True
        Else
            opt_off.Value = True
        End If
        
        udt_date_fr.RawData = sDate_Fr
        udt_date_to.RawData = sDate_To
        
        txt_bed_fl.Text = "2"
        ULabel12.Caption = "下冷床时间"
        ULabel13.Caption = "下冷床温度"
        
        If Gf_Sp_Refer(M_CN1, sc2, Mc1, , , False) Then
            ss2.OperationMode = OperationModeNormal
            Call Gp_Sp_EvenRowBackcolor(ss2)
        End If
        
        Call Gf_Sp_Cls(sc1)
        Call MenuTool_ReSet(False)
        
    End If

End Sub

Private Sub stx_cb_time_DblClick()

    stx_cb_time.RawData = Gf_DTSet(M_CN1, "S")

End Sub

Private Sub MenuTool_ReSet(bDel_Fl As Boolean)

    With MDIMain.MenuTool
    
        If bDel_Fl Then
            .Buttons(5).Enabled = True                  'Delete
        Else
            .Buttons(5).Enabled = False                  'Delete
        End If
        
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Delete
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = True                 'Excel
            
    End With

End Sub
