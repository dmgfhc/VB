VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0220C 
   Caption         =   "轧钢生产规范输入 - AQA0220C"
   ClientHeight    =   9420
   ClientLeft      =   165
   ClientTop       =   -375
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox P_txt_WID_MIN 
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
      Left            =   13335
      TabIndex        =   81
      Top             =   4785
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox P_txt_WID_MAX 
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
      Left            =   14175
      TabIndex        =   80
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   585
      Left            =   30
      TabIndex        =   72
      Top             =   30
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   1032
      _Version        =   196609
      Begin VB.CommandButton cmd_ListView 
         Caption         =   "<"
         Height          =   315
         Left            =   6960
         TabIndex        =   78
         Top             =   150
         Width           =   435
      End
      Begin VB.CommandButton cmd_ListView_WID 
         Caption         =   "<"
         Height          =   315
         Left            =   14625
         TabIndex        =   77
         Top             =   165
         Width           =   435
      End
      Begin VB.TextBox P_txt_MILL_STD_NO 
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
         Left            =   1920
         TabIndex        =   75
         Top             =   150
         Width           =   1635
      End
      Begin VB.CommandButton cmd_ListView_THK 
         Caption         =   "<"
         Height          =   315
         Left            =   10815
         TabIndex        =   74
         Top             =   165
         Width           =   435
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   315
         Left            =   9450
         TabIndex        =   73
         Top             =   165
         Width           =   1350
         _Version        =   393216
         _ExtentX        =   2381
         _ExtentY        =   556
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   2
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "AQA0220C.frx":0000
         UserResize      =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   0
         Left            =   90
         Top             =   150
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "规范编号"
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Index           =   0
         Left            =   7650
         Top             =   150
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "厚度组"
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
      Begin FPSpread.vaSpread ss4 
         Height          =   315
         Left            =   13260
         TabIndex        =   76
         Top             =   165
         Width           =   1350
         _Version        =   393216
         _ExtentX        =   2381
         _ExtentY        =   556
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   2
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "AQA0220C.frx":0366
         UserResize      =   0
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Index           =   3
         Left            =   11460
         Top             =   150
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "宽度组"
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
      Begin FPSpread.vaSpread ss2 
         Height          =   315
         Left            =   5760
         TabIndex        =   79
         Top             =   150
         Width           =   1170
         _Version        =   393216
         _ExtentX        =   2064
         _ExtentY        =   556
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   1
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "AQA0220C.frx":06CC
         UserResize      =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   0
         Left            =   3960
         Top             =   150
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "开始执行日期"
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
      Height          =   2895
      Left            =   60
      TabIndex        =   21
      Top             =   1410
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   5106
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSFrame SSFrame1 
         Height          =   2865
         Left            =   30
         TabIndex        =   22
         Top             =   0
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   5054
         _Version        =   196609
         Begin VB.CheckBox Check_C1 
            BackColor       =   &H00E1E4CD&
            Height          =   225
            Left            =   150
            TabIndex        =   83
            Top             =   60
            Width           =   255
         End
         Begin VB.TextBox txt_MILL_TMP_TGT 
            Height          =   300
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   40
            Top             =   2145
            Width           =   575
         End
         Begin VB.TextBox txt_CHG_TMP_DEF_SC 
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
            Height          =   300
            Left            =   5730
            MaxLength       =   4
            TabIndex        =   39
            Top             =   390
            Width           =   1305
         End
         Begin VB.TextBox txt_COOL_TMP_RATE 
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
            Height          =   300
            Left            =   5730
            TabIndex        =   38
            Top             =   1785
            Width           =   915
         End
         Begin VB.TextBox txt_MILL_RATET2 
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
            Height          =   300
            Left            =   2970
            MaxLength       =   3
            TabIndex        =   37
            Top             =   1800
            Width           =   645
         End
         Begin VB.TextBox txt_MILL_RATET1 
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
            Height          =   300
            Left            =   2970
            MaxLength       =   3
            TabIndex        =   36
            Top             =   1440
            Width           =   645
         End
         Begin VB.TextBox txt_MILL_TMPT2 
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
            Height          =   300
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   35
            Top             =   1800
            Width           =   1035
         End
         Begin VB.TextBox txt_MILL_TMPT1 
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
            Height          =   300
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   34
            Top             =   1440
            Width           =   1035
         End
         Begin VB.TextBox txt_MILL_TIME 
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
            Height          =   300
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   33
            Top             =   405
            Width           =   855
         End
         Begin VB.TextBox txt_CHG_TMP_DEF_TAPE 
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
            Height          =   300
            Left            =   5730
            MaxLength       =   4
            TabIndex        =   32
            Top             =   735
            Width           =   1305
         End
         Begin VB.TextBox txt_CHG_TMP_TGT 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   31
            Top             =   750
            Width           =   855
         End
         Begin VB.PictureBox Pic_HOT_USE 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6210
            ScaleHeight     =   240
            ScaleWidth      =   285
            TabIndex        =   30
            Top             =   2490
            Width           =   345
         End
         Begin VB.TextBox txt_HOT_USE 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   5730
            MaxLength       =   1
            TabIndex        =   29
            Top             =   2490
            Width           =   465
         End
         Begin VB.TextBox txt_COOL_CTL_TYP 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   28
            Top             =   2490
            Width           =   465
         End
         Begin VB.TextBox txt_COOL_CTL_NAME 
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
            Height          =   300
            Left            =   2385
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   2490
            Width           =   1260
         End
         Begin VB.TextBox txt_COOL_WAY 
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
            Left            =   5730
            MaxLength       =   1
            TabIndex        =   26
            Top             =   1095
            Width           =   465
         End
         Begin VB.TextBox txt_COOL_WAY_NAME 
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
            Height          =   300
            Left            =   6210
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   1095
            Width           =   1215
         End
         Begin VB.TextBox txt_CR_CD 
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
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   24
            Top             =   1095
            Width           =   465
         End
         Begin VB.TextBox txt_CR_NAME 
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
            Height          =   300
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1095
            Width           =   1215
         End
         Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_TGT 
            Height          =   300
            Left            =   5730
            TabIndex        =   41
            Top             =   1440
            Width           =   540
            _Version        =   262145
            _ExtentX        =   952
            _ExtentY        =   529
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoScroll      =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_MAX 
            Height          =   300
            Left            =   6870
            TabIndex        =   42
            Top             =   1440
            Width           =   540
            _Version        =   262145
            _ExtentX        =   952
            _ExtentY        =   529
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoScroll      =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   3
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_MIN 
            Height          =   300
            Left            =   6300
            TabIndex        =   43
            Top             =   1440
            Width           =   540
            _Version        =   262145
            _ExtentX        =   952
            _ExtentY        =   529
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoScroll      =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   3
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txt_MILL_TMP_MAX 
            Height          =   300
            Left            =   3060
            TabIndex        =   44
            Top             =   2145
            Width           =   570
            _Version        =   262145
            _ExtentX        =   1014
            _ExtentY        =   529
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoScroll      =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txt_MILL_TMP_MIN 
            Height          =   300
            Left            =   2505
            TabIndex        =   45
            Top             =   2145
            Width           =   570
            _Version        =   262145
            _ExtentX        =   1014
            _ExtentY        =   529
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
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
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtThousands    =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   2
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   5
            Left            =   120
            Top             =   405
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "轧制间隔（S）"
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
            Left            =   120
            Top             =   750
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "平均出炉温度"
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
            Index           =   8
            Left            =   120
            Top             =   1440
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "一阶段温度/厚度比"
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
            Left            =   120
            Top             =   1800
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "二阶段温度/厚度比"
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
            Index           =   13
            Left            =   3945
            Top             =   750
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "板坯头尾温差"
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
            Index           =   14
            Left            =   120
            Top             =   2145
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "终轧目标温度/误差"
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
            Left            =   3945
            Top             =   405
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "板坯表面/中心温差"
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
            Left            =   3945
            Top             =   1440
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "冷却目标温度/误差"
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
            Index           =   9
            Left            =   3945
            Top             =   1800
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "冷却速率"
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
            Left            =   120
            Top             =   2490
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "控制冷却"
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
         Begin CSTextLibCtl.sidbEdit txt_COOL_BED_TMP_TGT 
            Height          =   300
            Left            =   5730
            TabIndex        =   46
            Top             =   2145
            Width           =   735
            _Version        =   262145
            _ExtentX        =   1296
            _ExtentY        =   529
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoScroll      =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   21
            Left            =   3945
            Top             =   2145
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "冷床目标温度"
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
            Index           =   15
            Left            =   3945
            Top             =   2490
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "使用热矫"
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
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   2
            Left            =   120
            Top             =   1095
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "控制轧制"
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
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   3
            Left            =   3945
            Top             =   1095
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "冷却方法"
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
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   22
            Left            =   30
            Top             =   0
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   556
            Caption         =   "中厚板卷轧制"
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
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2865
         Left            =   7620
         TabIndex        =   47
         Top             =   0
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   5054
         _Version        =   196609
         Begin VB.CheckBox Check_C2 
            BackColor       =   &H00E1E4CD&
            Height          =   225
            Left            =   120
            TabIndex        =   84
            Top             =   60
            Width           =   255
         End
         Begin VB.TextBox txt_CR_NAME_Z 
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
            Height          =   300
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   1065
            Width           =   1215
         End
         Begin VB.TextBox txt_CR_CD_Z 
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
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   64
            Top             =   1065
            Width           =   465
         End
         Begin VB.TextBox txt_COOL_WAY_NAME_Z 
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
            Height          =   300
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   1065
            Width           =   1215
         End
         Begin VB.TextBox txt_COOL_WAY_Z 
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
            Left            =   5760
            MaxLength       =   1
            TabIndex        =   62
            Top             =   1065
            Width           =   465
         End
         Begin VB.TextBox txt_COOL_CTL_NAME_Z 
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
            Height          =   300
            Left            =   2385
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   2460
            Width           =   1260
         End
         Begin VB.TextBox txt_COOL_CTL_TYP_Z 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1920
            MaxLength       =   1
            TabIndex        =   60
            Top             =   2460
            Width           =   465
         End
         Begin VB.TextBox txt_HOT_USE_Z 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   5760
            MaxLength       =   1
            TabIndex        =   59
            Top             =   2460
            Width           =   465
         End
         Begin VB.PictureBox Pic_HOT_USE_Z 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6240
            ScaleHeight     =   240
            ScaleWidth      =   285
            TabIndex        =   58
            Top             =   2460
            Width           =   345
         End
         Begin VB.TextBox txt_CHG_TMP_TGT_Z 
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
            Height          =   300
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   57
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txt_CHG_TMP_DEF_TAPE_Z 
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
            Height          =   300
            Left            =   5760
            MaxLength       =   4
            TabIndex        =   56
            Top             =   705
            Width           =   1305
         End
         Begin VB.TextBox txt_MILL_TIME_Z 
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
            Height          =   300
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   55
            Top             =   375
            Width           =   855
         End
         Begin VB.TextBox txt_MILL_TMPT1_Z 
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
            Height          =   300
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   54
            Top             =   1410
            Width           =   1035
         End
         Begin VB.TextBox txt_MILL_TMPT2_Z 
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
            Height          =   300
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   53
            Top             =   1770
            Width           =   1035
         End
         Begin VB.TextBox txt_MILL_RATET1_Z 
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
            Height          =   300
            Left            =   2970
            MaxLength       =   3
            TabIndex        =   52
            Top             =   1410
            Width           =   645
         End
         Begin VB.TextBox txt_MILL_RATET2_Z 
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
            Height          =   300
            Left            =   2970
            MaxLength       =   3
            TabIndex        =   51
            Top             =   1770
            Width           =   645
         End
         Begin VB.TextBox txt_COOL_TMP_RATE_Z 
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
            Height          =   300
            Left            =   5760
            TabIndex        =   50
            Top             =   1755
            Width           =   915
         End
         Begin VB.TextBox txt_CHG_TMP_DEF_SC_Z 
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
            Height          =   300
            Left            =   5760
            MaxLength       =   4
            TabIndex        =   49
            Top             =   360
            Width           =   1305
         End
         Begin VB.TextBox txt_MILL_TMP_TGT_Z 
            Height          =   300
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   48
            Top             =   2115
            Width           =   570
         End
         Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_TGT_Z 
            Height          =   300
            Left            =   5760
            TabIndex        =   66
            Top             =   1410
            Width           =   540
            _Version        =   262145
            _ExtentX        =   952
            _ExtentY        =   529
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoScroll      =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_MAX_Z 
            Height          =   300
            Left            =   6880
            TabIndex        =   67
            Top             =   1410
            Width           =   540
            _Version        =   262145
            _ExtentX        =   952
            _ExtentY        =   529
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoScroll      =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   3
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_MIN_Z 
            Height          =   300
            Left            =   6330
            TabIndex        =   68
            Top             =   1410
            Width           =   540
            _Version        =   262145
            _ExtentX        =   952
            _ExtentY        =   529
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoScroll      =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   3
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txt_MILL_TMP_MAX_Z 
            Height          =   300
            Left            =   3075
            TabIndex        =   69
            Top             =   2115
            Width           =   570
            _Version        =   262145
            _ExtentX        =   1005
            _ExtentY        =   529
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoScroll      =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txt_MILL_TMP_MIN_Z 
            Height          =   300
            Left            =   2505
            TabIndex        =   70
            Top             =   2115
            Width           =   570
            _Version        =   262145
            _ExtentX        =   1005
            _ExtentY        =   529
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            Justification   =   2
            BorderStyle     =   0
            FmtThousands    =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   2
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   23
            Left            =   120
            Top             =   375
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "轧制间隔（S）"
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
            Index           =   24
            Left            =   120
            Top             =   720
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "平均出炉温度"
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
            Index           =   25
            Left            =   120
            Top             =   1410
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "一阶段温度/厚度比"
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
            Index           =   26
            Left            =   120
            Top             =   1770
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "二阶段温度/厚度比"
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
            Index           =   27
            Left            =   3975
            Top             =   720
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "板坯头尾温差"
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
            Top             =   2115
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "终轧目标温度/误差"
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
            Index           =   29
            Left            =   3975
            Top             =   375
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "板坯表面/中心温差"
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
            Left            =   3975
            Top             =   1410
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "冷却目标温度/误差"
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
            Index           =   31
            Left            =   3975
            Top             =   1770
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "冷却速率"
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
            Index           =   32
            Left            =   120
            Top             =   2460
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "控制冷却"
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
         Begin CSTextLibCtl.sidbEdit txt_COOL_BED_TMP_TGT_Z 
            Height          =   300
            Left            =   5760
            TabIndex        =   71
            Top             =   2115
            Width           =   735
            _Version        =   262145
            _ExtentX        =   1296
            _ExtentY        =   529
            _StockProps     =   125
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoScroll      =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   ""
            StartText.x     =   3
            StartText.y     =   2
            FirstVisPos     =   0
            HiAnchor        =   0
            HiNew           =   0
            CaretHeight     =   15
            CurNumDataChars =   0
            MaxDataChars    =   0
            FirstDataPos    =   0
            CurPos          =   0
            MaxLen          =   0
            DataReadOnly    =   0   'False
            Mask            =   ""
            BorderStyle     =   0
            FmtControl      =   1
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   33
            Left            =   3975
            Top             =   2115
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "冷床目标温度"
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
            Index           =   34
            Left            =   3975
            Top             =   2460
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "使用热矫"
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
            Height          =   315
            Index           =   35
            Left            =   120
            Top             =   1065
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "控制轧制"
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
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   36
            Left            =   3975
            Top             =   1065
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Caption         =   "冷却方法"
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
            Height          =   315
            Index           =   37
            Left            =   30
            Top             =   0
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   556
            Caption         =   "中板轧制"
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
      End
   End
   Begin VB.TextBox P_txt_THK_MAX 
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
      Left            =   14160
      TabIndex        =   1
      Top             =   4455
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox P_txt_THK_MIN 
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
      Left            =   13320
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox P_txt_APP_DATE 
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
      Left            =   13470
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10800
      Top             =   7110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AQA0220C.frx":09E3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   4065
      Left            =   45
      TabIndex        =   3
      Top             =   5160
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   7170
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   66
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQA0220C.frx":0D35
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   885
      Left            =   30
      TabIndex        =   4
      Top             =   4290
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   1561
      _Version        =   196609
      Begin VB.TextBox txt_INS_DATE 
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
         Height          =   300
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   465
         Width           =   1215
      End
      Begin VB.TextBox txt_INS_EMP 
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
         Height          =   300
         Left            =   4965
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   465
         Width           =   1215
      End
      Begin VB.TextBox txt_UPD_DATE 
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
         Height          =   300
         Left            =   8025
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   465
         Width           =   1215
      End
      Begin VB.TextBox txt_UPD_EMP 
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
         Height          =   300
         Left            =   11115
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txt_MILL_STD_EDT_NO 
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
         Left            =   1875
         MaxLength       =   80
         TabIndex        =   5
         Top             =   60
         Width           =   8115
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Index           =   11
         Left            =   45
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "规范编辑号"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   16
         Left            =   30
         Top             =   465
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "录入日期"
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
         Index           =   17
         Left            =   3150
         Top             =   465
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "录入人"
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
         Index           =   18
         Left            =   6210
         Top             =   465
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "修改日期"
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
         Index           =   19
         Left            =   9300
         Top             =   480
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "修改人"
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
   Begin Threed.SSFrame SSFrame4 
      Height          =   795
      Left            =   30
      TabIndex        =   10
      Top             =   600
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   1402
      _Version        =   196609
      Begin VB.TextBox txt_C2 
         Height          =   270
         Left            =   13995
         TabIndex        =   86
         Top             =   450
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txt_C1 
         Height          =   285
         Left            =   13260
         TabIndex        =   85
         Top             =   450
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txt_MLT_PLT 
         Height          =   285
         Left            =   12300
         TabIndex        =   82
         Top             =   465
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox txt_WID_MAX 
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
         Height          =   300
         Left            =   13980
         TabIndex        =   20
         Top             =   60
         Width           =   855
      End
      Begin VB.TextBox txt_WID_MIN 
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
         Height          =   300
         Left            =   13110
         TabIndex        =   19
         Top             =   60
         Width           =   855
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
         Height          =   300
         Left            =   1920
         MaxLength       =   11
         TabIndex        =   18
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox txt_STLGRD_DETAIL 
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
         Height          =   300
         Left            =   3240
         TabIndex        =   17
         Top             =   450
         Width           =   3675
      End
      Begin VB.TextBox txt_HCR_NAME 
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
         Height          =   300
         Left            =   9660
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   16
         Top             =   450
         Width           =   1380
      End
      Begin VB.TextBox txt_HCR_KND 
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
         Left            =   9240
         MaxLength       =   1
         TabIndex        =   15
         Top             =   450
         Width           =   405
      End
      Begin VB.TextBox txt_THK_MAX 
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
         Height          =   300
         Left            =   10110
         TabIndex        =   14
         Top             =   60
         Width           =   855
      End
      Begin VB.TextBox txt_THK_MIN 
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
         Height          =   300
         Left            =   9240
         TabIndex        =   13
         Top             =   60
         Width           =   855
      End
      Begin VB.TextBox txt_APP_DATE 
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
         Left            =   5745
         TabIndex        =   12
         Top             =   60
         Width           =   1170
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
         Height          =   300
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   11
         Top             =   60
         Width           =   1635
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   1
         Left            =   90
         Top             =   450
         Width           =   1755
         _ExtentX        =   3096
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
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   4
         Left            =   7440
         Top             =   450
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "HCR 分类"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   20
         Left            =   90
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "规范编号"
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Index           =   1
         Left            =   7440
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "厚度组"
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
         Index           =   1
         Left            =   3945
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "开始执行日期"
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
         Index           =   2
         Left            =   11310
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         Caption         =   "宽度组"
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
End
Attribute VB_Name = "AQA0220C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量标准管理
'-- Program Name      轧钢生产规范输入
'-- Program ID        AQA0200C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       轧钢生产规范输入
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

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim lCopyRow As Long                'Copy Row
Dim btChk_Year  As Boolean          '年度（ss2）是否显示选择
Dim btChk_THK   As Boolean          '厚度组(ss3)是否显示选择
Dim btChk_WID   As Boolean


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(P_txt_MILL_STD_NO, "p", " ", " ", "r", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(P_txt_APP_DATE, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(P_txt_THK_MIN, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(P_txt_THK_MAX, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(P_txt_WID_MIN, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(P_txt_WID_MAX, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 53, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 58, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 59, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 60, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 61, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 62, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 63, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 64, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 65, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 66, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQA0220C.P_SREFER", Key:="P-R"
    Sc1.Add Item:="AQA0220C.P_ONEROW", Key:="P-O"
    Sc1.Add Item:="AQA0220C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
         
End Sub
Private Function Change_MLT_PLT(ByVal iC1Val As Integer, ByVal iC2Val As Integer, Optional iRow As Long = 0) As String
Dim sOLD_ML_PLT_CD  As String
Dim sNEW_ML_PLT_CD  As String
Dim sLAB            As String
Dim s_C1_CD         As String
Dim s_C2_CD         As String

    If iRow > 0 Then
        With ss1
            .Row = iRow
            .Col = 1
            sLAB = .Text
            .Col = 66
            sOLD_ML_PLT_CD = .Text
        End With
    End If
    
    If sOLD_ML_PLT_CD = "**" And (sLAB <> "Input" Or sLAB <> "Update") Then
        Change_MLT_PLT = sOLD_ML_PLT_CD
        Exit Function
    End If
    
    If iC1Val = 1 Then
        s_C1_CD = "C1"
    Else
        s_C1_CD = "NO"
    End If
    
    If iC2Val = 1 Then
        s_C2_CD = "C3"
    Else
        s_C2_CD = "NO"
    End If
    
    If s_C1_CD = "C1" And s_C2_CD = "C3" Then
        sNEW_ML_PLT_CD = "**"
    ElseIf s_C1_CD = "C1" And s_C2_CD = "NO" Then
        sNEW_ML_PLT_CD = "C1"
    ElseIf s_C1_CD = "NO" And s_C2_CD = "C3" Then
        sNEW_ML_PLT_CD = "C3"
    Else
        sNEW_ML_PLT_CD = sOLD_ML_PLT_CD
        
    End If
    
    Change_MLT_PLT = sNEW_ML_PLT_CD
    
End Function

Private Sub Check_C1_Click()
    txt_MLT_PLT.Text = Change_MLT_PLT(Check_C1.Value, Check_C2.Value, ss1.ActiveRow)
    Select Case Check_C1.Value
        Case 0
            txt_C1.Text = "NO"
        Case 1
            txt_C1.Text = "C1"
    End Select
End Sub

Private Sub Check_C2_Click()
    txt_MLT_PLT.Text = Change_MLT_PLT(Check_C1.Value, Check_C2.Value, ss1.ActiveRow)
    Select Case Check_C2.Value
        Case 0
            txt_C2.Text = "NO"
        Case 1
            txt_C2.Text = "C3"
    End Select
End Sub

Private Sub cmd_ListView_Click()
Dim sQuery As String
    
    sQuery = " Select Distinct APP_DATE From QP_ROLL_STD Where MILL_STD_NO = "
    btChk_Year = Not btChk_Year

    If btChk_Year = False Then
            
        With ss2
            
            .MaxCols = 1
            .MaxRows = 1
            .Height = 313
        
            btChk_Year = False
    
        End With

    Else
         
       If P_txt_MILL_STD_NO.Text = "" Or Trim(P_txt_MILL_STD_NO.Text) = "" Then
          Exit Sub
       End If
           
        sQuery = sQuery + "'" + P_txt_MILL_STD_NO.Text + "'"
        
        Call GS_Combo_SS_ADD(sQuery, ss2)
        
        Call GS_ssBackColorSet(ss2)
    
    End If
    
        If Gf_GetCellNullCheck(ss2, 1, 1) <> "" Then
            P_txt_APP_DATE.Text = Gf_GetCellNullCheck(ss2, 1, 1)
        End If

End Sub

Private Sub cmd_ListView_THK_Click()
Dim sQuery As String
    
    sQuery = " Select THK_MIN,THK_MAX From QP_MILL_STD Where MILL_STD_NO = "
    btChk_THK = Not btChk_THK

    If btChk_THK = False Then
            
        With ss3
            
            .MaxCols = 2
            .MaxRows = 1
            .Height = 313
        
            btChk_THK = False
    
        End With

    Else
         
       If P_txt_MILL_STD_NO.Text = "" Or Trim(P_txt_MILL_STD_NO.Text) = "" Then
            Call MsgBox("请输入轧钢规程编号", vbOKOnly, "系统信息")
          Exit Sub
       ElseIf P_txt_APP_DATE.Text = "" Or Trim(P_txt_APP_DATE.Text) = "" Then
            Call MsgBox("请输入开始执行日期", vbOKOnly, "系统信息")
          Exit Sub
       End If
           
        sQuery = sQuery + "'" + P_txt_MILL_STD_NO.Text + "' AND"
        sQuery = sQuery + " APP_DATE = '" + P_txt_APP_DATE.Text + "'"
        
        Call GS_Combo_SS_ADD(sQuery, ss3)
        
        Call GS_ssBackColorSet(ss3)
    
    End If
    
        If Gf_GetCellNullCheck(ss3, 1, 1) <> "" And Gf_GetCellNullCheck(ss3, 1, 2) <> "" Then
            P_txt_THK_MIN.Text = Gf_GetCellNullCheck(ss3, 1, 1)
            P_txt_THK_MAX.Text = Gf_GetCellNullCheck(ss3, 1, 2)
        End If

End Sub

Private Sub cmd_ListView_WID_Click()
Dim sQuery As String
    
    sQuery = " Select WID_MIN,WID_MAX From QP_MILL_STD Where MILL_STD_NO = "
    btChk_WID = Not btChk_WID

    If btChk_THK = False Then
            
        With ss4
            
            .MaxCols = 2
            .MaxRows = 1
            .Height = 313
        
            btChk_THK = False
    
        End With

    Else
         
       If P_txt_MILL_STD_NO.Text = "" Or Trim(P_txt_MILL_STD_NO.Text) = "" Then
            Call MsgBox("请输入轧钢规程编号", vbOKOnly, "系统信息")
          Exit Sub
       ElseIf P_txt_APP_DATE.Text = "" Or Trim(P_txt_APP_DATE.Text) = "" Then
            Call MsgBox("请输入开始执行日期", vbOKOnly, "系统信息")
          Exit Sub
       End If
           
        sQuery = sQuery + "'" + P_txt_MILL_STD_NO.Text + "' AND"
        sQuery = sQuery + " APP_DATE = '" + P_txt_APP_DATE.Text + "'"
        
        Call GS_Combo_SS_ADD(sQuery, ss4)
        
        Call GS_ssBackColorSet(ss4)
    
    End If
    
        If Gf_GetCellNullCheck(ss4, 1, 1) <> "" And Gf_GetCellNullCheck(ss4, 1, 2) <> "" Then
            P_txt_WID_MIN.Text = Gf_GetCellNullCheck(ss4, 1, 1)
            P_txt_WID_MAX.Text = Gf_GetCellNullCheck(ss4, 1, 2)
        End If
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
'    If Not (KeyCode = vbKeyF4) Then Exit Sub
    
    Select Case Me.ActiveControl.Name
    
        Case "P_txt_MILL_STD_NO"                            '轧钢规程编号
            sCode = "MILL_STD_NO"
            'Set oCodeName = P_txt_APP_DATE
            
        Case "txt_MILL_STD_NO"                              '轧钢规程编号
            sCode = "MILL_STD_NO"
            Set oCodeName = txt_APP_DATE
                        
        Case "txt_STLGRD"               '钢种
            sCode = "STLGRD"
            Set oCodeName = txt_STLGRD_DETAIL
            
        Case "txt_HCR_KND"              'HCR 分类
            sCode = "C0005"
            Set oCodeName = txt_HCR_NAME
            
        Case "txt_CR_CD"                '控制轧制
            sCode = "Q0035"
            Set oCodeName = txt_CR_NAME
            
        Case "txt_COOL_WAY"             '冷却方法
            sCode = "Q0036"
            Set oCodeName = txt_COOL_WAY_NAME
            
        Case "txt_COOL_CTL_TYP"         '控制冷却代码
            sCode = "Q0037"
            Set oCodeName = txt_COOL_CTL_NAME
            
        Case "txt_HOT_USE"         '使用热矫代码
            sCode = "Q0038"
            'Set oCodeName = txt_COOL_CTL_NAME
            
'HYS ADD START
        Case "txt_CR_CD_Z"                '控制轧制
            sCode = "Q0035"
            Set oCodeName = txt_CR_NAME_Z
            
        Case "txt_COOL_WAY_Z"             '冷却方法
            sCode = "Q0036"
            Set oCodeName = txt_COOL_WAY_NAME_Z
            
        Case "txt_COOL_CTL_TYP_Z"         '控制冷却代码
            sCode = "Q0037"
            Set oCodeName = txt_COOL_CTL_NAME
            
        Case "txt_HOT_USE_Z"         '使用热矫代码
            sCode = "Q0038"
            'Set oCodeName = txt_COOL_CTL_NAME
'HYS ADD END

    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
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
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)
       
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
'    Call GP_ROW_BACKCOLOR(ss1)
    
    'Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
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

    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Ins()
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 61)
    Call Spread_to_Master(ss1, ss1.ActiveRow)
    txt_MILL_STD_NO.SetFocus
End Sub

Public Sub Form_Pro()
    Dim iMaxrow As Long
    Dim iRow As Long
    Dim icount As Long

    iRow = ss1.Row
    iMaxrow = ss1.MaxRows
    
    ss1.ReDraw = False
    
    For icount = 1 To ss1.MaxRows
         ss1.Col = 0
         ss1.Row = icount
        Select Case ss1.Text
            Case "Input", "Update"
             If Sp_AllUse_NecessaryCheck(icount) = False Then          '必须输入项检查
                Call Spread_to_Master(ss1, icount)
                Exit Sub
             End If                                             '最大值,最小值,目标值检查
             If Sp_subMinMaxValueCheck(icount) = False Then
                Call Spread_to_Master(ss1, icount)
                Exit Sub
             End If
        End Select
    Next icount
    
    ss1.ReDraw = True
     
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        txt_INS_EMP.Text = sUserID
         If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            Call Gp_Goto_Row(ss1, iMaxrow, iRow)
            Call Spread_to_Master(ss1, iRow)
         End If
    End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_AllDel(M_CN1, Proc_Sc("Sc"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub
Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call MS_Cls
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        ss2.MaxRows = 1
        ss2.Height = 255
        ss3.MaxRows = 1
        ss3.Height = 255
        btChk_Year = False
        btChk_THK = False
        Call GP_SET_CELL_VALUE(ss2, 1, 1, "")
        Call GP_SET_CELL_VALUE(ss3, 1, 2, "")
        Call GP_SET_CELL_VALUE(ss4, 1, 2, "")
'        txt_THK_MIN.Text = ""
'        txt_THK_MAX.Text = ""
        'rControl(1).SetFocus
        P_txt_MILL_STD_NO.SetFocus
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
            
            If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
                Call Spread_to_Master(ss1, 1)
                Call Gp_Ms_ControlLock(Mc1("pControl"), True)
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call GP_SELECT_ROW(ss1, 1)
                ss2.MaxRows = 1
                ss2.Height = 255
                ss3.MaxRows = 1
                ss3.Height = 255
                ss4.MaxRows = 1
                ss4.Height = 255
                btChk_Year = False
                btChk_THK = False
                btChk_WID = False
                Exit Sub
            End If
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub



Private Sub P_txt_APP_DATE_KeyPress(KeyAscii As Integer)
   ' KeyAscii = txt_KeyPress(KeyAscii)
End Sub


Private Sub P_txt_THK_MAX_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub P_txt_THK_MIN_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub Pic_HOT_USE_Click()
    If Pic_HOT_USE.Picture.Width <= 0 Then
       txt_HOT_USE.Text = "Y"
    Else
       txt_HOT_USE.Text = "N"
    End If
End Sub

Private Sub Pic_HOT_USE_Z_Click()
    If Pic_HOT_USE_Z.Picture.Width <= 0 Then
       txt_HOT_USE_Z.Text = "Y"
    Else
       txt_HOT_USE_Z.Text = "N"
    End If
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)
    If Gf_Sc_Authority(sAuthority, "U") Then

        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), 0)
        
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 61)

    End If
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub



Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 61)
    End If
End Sub

Private Sub ss1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    Call Spread_to_Master(ss1, NewRow)
'    Call GP_SetRowHeaderClear(ss1, NewRow)
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub Spread_to_Master(ByVal sp As vaSpread, ByVal iRow As Long)
    Dim RowLabel As String

        With sp
        
            If iRow > 0 Then
                .Row = iRow
               
                .Col = 0: RowLabel = .Text
                .Col = 1: txt_MILL_STD_NO.Text = .Text
                .Col = 2: txt_APP_DATE.Text = .Text
                .Col = 3: txt_THK_MIN.Text = .Text
                .Col = 4: txt_THK_MAX.Text = .Text
                .Col = 5: txt_WID_MIN.Text = .Text
                .Col = 6: txt_WID_MAX.Text = .Text
                .Col = 7: txt_STLGRD.Text = .Text
                .Col = 8: txt_STLGRD_DETAIL.Text = .Text
                .Col = 9: txt_HCR_KND.Text = .Text
                .Col = 10: txt_HCR_NAME.Text = .Text
                
                .Col = 11: txt_CHG_TMP_TGT.Text = .Text
                .Col = 12: txt_CHG_TMP_DEF_SC = .Text
                .Col = 13: txt_CHG_TMP_DEF_TAPE.Text = .Text
                .Col = 14: txt_CR_CD.Text = .Text
                .Col = 15: txt_CR_NAME.Text = .Text
                .Col = 16: txt_MILL_TIME.Text = .Text
                .Col = 17: txt_MILL_TMPT1.Text = .Text
                .Col = 18: txt_MILL_RATET1.Text = .Text
                .Col = 19: txt_MILL_TMPT2.Text = .Text
                .Col = 20: txt_MILL_RATET2.Text = .Text
                .Col = 21: txt_MILL_TMP_MIN.Text = .Text
                .Col = 22: txt_MILL_TMP_MAX.Text = .Text
                .Col = 23: txt_MILL_TMP_TGT.Text = .Text
                .Col = 24: txt_COOL_WAY.Text = .Text
                .Col = 25: txt_COOL_WAY_NAME.Text = .Text
                .Col = 26: txt_COOL_TMP_MIN.Text = .Text
                .Col = 27: txt_COOL_TMP_MAX.Text = .Text
                .Col = 28: txt_COOL_TMP_TGT.Text = .Text
                .Col = 29: txt_COOL_TMP_RATE.Text = .Text
                .Col = 30: txt_COOL_BED_TMP_TGT.Text = .Text
                .Col = 31: txt_COOL_CTL_TYP.Text = .Text
                .Col = 32: txt_COOL_CTL_NAME.Text = .Text
                .Col = 33: txt_HOT_USE.Text = .Text
                
                .Col = 35: txt_CHG_TMP_TGT_Z.Text = .Text
                .Col = 36: txt_CHG_TMP_DEF_SC_Z.Text = .Text
                .Col = 37: txt_CHG_TMP_DEF_TAPE_Z.Text = .Text
                .Col = 38: txt_CR_CD_Z.Text = .Text
                .Col = 39: txt_CR_NAME_Z.Text = .Text
                .Col = 40: txt_MILL_TIME_Z.Text = .Text
                .Col = 41: txt_MILL_TMPT1_Z.Text = .Text
                .Col = 42: txt_MILL_RATET1_Z.Text = .Text
                .Col = 43: txt_MILL_TMPT2_Z.Text = .Text
                .Col = 44: txt_MILL_RATET2_Z.Text = .Text
                .Col = 45: txt_MILL_TMP_MIN_Z.Text = .Text
                .Col = 46: txt_MILL_TMP_MAX_Z.Text = .Text
                .Col = 47: txt_MILL_TMP_TGT_Z.Text = .Text
                .Col = 48: txt_COOL_WAY_Z.Text = .Text
                .Col = 49: txt_COOL_WAY_NAME_Z.Text = .Text
                .Col = 50: txt_COOL_TMP_MIN_Z.Text = .Text
                .Col = 51: txt_COOL_TMP_MAX_Z.Text = .Text
                .Col = 52: txt_COOL_TMP_TGT_Z.Text = .Text
                .Col = 53: txt_COOL_TMP_RATE_Z.Text = .Text
                .Col = 54: txt_COOL_BED_TMP_TGT_Z.Text = .Text
                .Col = 55: txt_COOL_CTL_TYP_Z.Text = .Text
                .Col = 56: txt_COOL_CTL_NAME_Z.Text = .Text
                .Col = 57: txt_HOT_USE_Z.Text = .Text
                
                .Col = 59: txt_MILL_STD_EDT_NO.Text = .Text
                .Col = 60: txt_INS_DATE.Text = .Text
                .Col = 62: txt_INS_EMP.Text = .Text
                .Col = 63: txt_UPD_DATE.Text = .Text
                .Col = 65: txt_UPD_EMP.Text = .Text
                .Col = 66: txt_MLT_PLT.Text = .Text
                
                If RowLabel = "Input" Then
                    txt_MILL_STD_NO.Locked = False
                    txt_THK_MIN.Locked = False
                    txt_THK_MAX.Locked = False
                    txt_WID_MIN.Locked = False
                    txt_WID_MAX.Locked = False
                    txt_APP_DATE.Locked = False
                Else
                    txt_MILL_STD_NO.Locked = True
                    txt_THK_MIN.Locked = True
                    txt_THK_MAX.Locked = True
                    txt_WID_MIN.Locked = True
                    txt_WID_MAX.Locked = True
                    txt_APP_DATE.Locked = True
                End If
            Else
                Exit Sub
            End If
        
        End With

End Sub

Public Sub Spread_Can()

    Call GP_SELECT_ROW(ss1, ss1.Row)
    Call GP_ROW_CANCEL(Proc_Sc("Sc"))
'    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    Call Spread_to_Master(ss1, ss1.ActiveRow)
    Call Gp_Ms_ControlLock(Mc1("pControl"), True)
      
End Sub

Public Sub Spread_Del()
    
    Call GP_SET_CELL_VALUE(ss1, ss1.Row, 0, "Delete")
    'Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Public Sub Spread_Cpy()

    lCopyRow = ss1.ActiveRow
    'Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()
    
    Call GP_ROW_PASTE(Proc_Sc("Sc"), lCopyRow)
    'Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 61)
    Call Spread_to_Master(ss1, ss1.ActiveRow)
    txt_MILL_STD_NO.SetFocus
    
End Sub

Public Sub Ms_To_SP(ByVal sp As vaSpread, ByVal iRow As Long, ByVal iCol As Long, ByVal vName As String)
    Dim old_Value As Variant
    Dim iValue As Variant
    
    If (vName <> "0") And (vName <> "1") Then
        If TypeName(Me.Controls(vName)) = "TextBox" Then
            iValue = Me.Controls(vName).Text
        End If
        
        If TypeName(Me.Controls(vName)) = "sidbEdit" Then
                iValue = Me.Controls(vName).Value
        End If
    Else
        iValue = vName
    End If
    
    With sp
        If iCol = 1 Or iCol = 2 Or iCol = 3 Or iCol = 4 Or iCol = 5 Or iCol = 6 Then
            .Row = iRow
            .Col = 0
            If (.Text = "Input") Then
                .Col = iCol
                .Value = iValue
                .Text = iValue
            Else
                Exit Sub
            End If
        Else
            .Row = iRow
            .Col = iCol
            old_Value = .Value
            .Value = iValue
            .Text = iValue
            If old_Value <> .Value Then
                .Col = 0
                    If (.Text = "Input") Or (.Text = "Update") Then
                        .Text = .Text
                    Else
                        .Text = "Update"
                    End If
                    .Col = iCol
            Else
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
    With ss2
    
        If Gf_GetCellNullCheck(ss2, Row, 1) <> "" Then
            Call GP_SET_CELL_VALUE(ss2, 1, 1, Gf_GetCellNullCheck(ss2, Row, 1))
        End If
        
        .MaxRows = 1
        .Height = 313
        
        P_txt_APP_DATE.Text = Gf_GetCellNullCheck(ss2, 1, 1)
        
        btChk_Year = False
    
    End With

End Sub

Private Sub ss3_DblClick(ByVal Col As Long, ByVal Row As Long)
    With ss3
    
        If Gf_GetCellNullCheck(ss3, Row, 1) <> "" And Gf_GetCellNullCheck(ss3, Row, 2) <> "" Then
            Call GP_SET_CELL_VALUE(ss3, 1, 1, Gf_GetCellNullCheck(ss3, Row, 1))
            Call GP_SET_CELL_VALUE(ss3, 1, 2, Gf_GetCellNullCheck(ss3, Row, 2))
        End If
        
        .MaxRows = 1
        .Height = 313
        
        P_txt_THK_MIN.Text = Gf_GetCellNullCheck(ss3, 1, 1)
        P_txt_THK_MAX.Text = Gf_GetCellNullCheck(ss3, 1, 2)
        
        btChk_THK = False
    
    End With


End Sub



Private Sub txt_APP_DATE_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 2, txt_APP_DATE.Name)
    End If
End Sub

Private Sub txt_APP_DATE_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_C1_Change()
Dim sOLD_PLT_CD As String
Dim sOLD_C1 As String
Dim sOLD_C2 As String
Dim sNEW_C1 As String

    sOLD_PLT_CD = Trim(txt_MLT_PLT.Text)
    
    Select Case sOLD_PLT_CD
        Case "C1"
            sOLD_C1 = "C1"
            sOLD_C2 = "NO"
        Case "C3"
            sOLD_C1 = "NO"
            sOLD_C2 = "C3"
        Case "**"
            sOLD_C1 = "C1"
            sOLD_C2 = "C3"
    End Select
    
    sNEW_C1 = txt_C1.Text
    
    If sNEW_C1 = sOLD_C1 Then
        Exit Sub
    Else
        Select Case sOLD_C2
            Case "NO"
                If sNEW_C1 = "C1" Then
                    txt_MLT_PLT.Text = "C1"
                ElseIf sNEW_C1 = "NO" Then
                     txt_MLT_PLT.Text = "C3"
                     
'                    txt_MLT_PLT.Text = sOLD_PLT_CD
'                    Select Case Trim(sOLD_PLT_CD)
'                        Case "C1"
'                            Check_C1.Value = 1
'                            Check_C2.Value = 0
'                        Case "C3"
'                            Check_C1.Value = 0
'                            Check_C2.Value = 1
'                        Case "**"
'                            Check_C1.Value = 1
'                            Check_C2.Value = 1
'                    End Select
                End If
            Case "C3"
                If sNEW_C1 = "C1" Then
                    txt_MLT_PLT.Text = "**"
                ElseIf sNEW_C1 = "NO" Then
                    txt_MLT_PLT.Text = "C3"
                End If
        End Select
    End If
End Sub

Private Sub txt_C2_Change()
Dim sOLD_PLT_CD As String
Dim sOLD_C1 As String
Dim sOLD_C2 As String
Dim sNEW_C2 As String
Dim sNEW_PLT_CD As String

    sOLD_PLT_CD = Trim(txt_MLT_PLT.Text)
    
    Select Case sOLD_PLT_CD
        Case "C1"
            sOLD_C1 = "C1"
            sOLD_C2 = "NO"
        Case "C3"
            sOLD_C1 = "NO"
            sOLD_C2 = "C3"
        Case "**"
            sOLD_C1 = "C1"
            sOLD_C2 = "C3"
    End Select
    
    sNEW_C2 = txt_C2.Text
    
    If sNEW_C2 = sOLD_C2 Then
        Exit Sub
    Else
        Select Case sOLD_C1
            Case "NO"
                 If sNEW_C2 = "C3" Then
                    txt_MLT_PLT.Text = "C3"
                 ElseIf sNEW_C2 = "NO" Then
                    txt_MLT_PLT.Text = "C1"

'                    txt_MLT_PLT.Text = sOLD_PLT_CD
'                    Select Case Trim(sOLD_PLT_CD)
'                        Case "C1"
'                             Check_C1.Value = 1
'                             Check_C2.Value = 0
'                        Case "C3"
'                             Check_C1.Value = 0
'                             Check_C2.Value = 1
'                        Case "**"
'                             Check_C1.Value = 1
'                             Check_C2.Value = 1
'                    End Select
                 End If
            Case "C1"
                 If sNEW_C2 = "C3" Then
                    txt_MLT_PLT.Text = "**"
                 ElseIf sNEW_C2 = "NO" Then
                    txt_MLT_PLT.Text = "C1"
                 End If
        End Select
    End If

End Sub

Private Sub txt_CHG_TMP_DEF_SC_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 12, txt_CHG_TMP_DEF_SC.Name)
    End If
End Sub

Private Sub txt_CHG_TMP_DEF_SC_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_CHG_TMP_DEF_SC_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 36, txt_CHG_TMP_DEF_SC_Z.Name)
    End If

End Sub

Private Sub txt_CHG_TMP_DEF_TAPE_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 13, txt_CHG_TMP_DEF_TAPE.Name)
    End If
End Sub

Private Sub txt_CHG_TMP_DEF_TAPE_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_CHG_TMP_DEF_TAPE_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 37, txt_CHG_TMP_DEF_TAPE_Z.Name)
    End If
End Sub

Private Sub txt_CHG_TMP_TGT_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 11, txt_CHG_TMP_TGT.Name)
    End If
End Sub

Private Sub txt_CHG_TMP_TGT_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_CHG_TMP_TGT_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 35, txt_CHG_TMP_TGT_Z.Name)
    End If

End Sub

Private Sub txt_COOL_BED_TMP_TGT_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 30, txt_COOL_BED_TMP_TGT.Name)
    End If
End Sub

Private Sub txt_COOL_BED_TMP_TGT_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 54, txt_COOL_BED_TMP_TGT_Z.Name)
    End If

End Sub

Private Sub txt_COOL_CTL_NAME_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 32, txt_COOL_CTL_NAME.Name)
    End If
End Sub

Private Sub txt_COOL_CTL_NAME_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 56, txt_COOL_CTL_NAME_Z.Name)
    End If

End Sub

Private Sub txt_COOL_CTL_TYP_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 31, txt_COOL_CTL_TYP.Name)
    End If
End Sub

Private Sub txt_COOL_CTL_TYP_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 55, txt_COOL_CTL_TYP_Z.Name)
    End If

End Sub

Private Sub txt_COOL_TMP_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 27, txt_COOL_TMP_MAX.Name)
    End If
End Sub

Private Sub txt_COOL_TMP_MAX_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_COOL_TMP_MAX_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 51, txt_COOL_TMP_MAX_Z.Name)
    End If

End Sub

Private Sub txt_COOL_TMP_MIN_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 26, txt_COOL_TMP_MIN.Name)
    End If
End Sub

Private Sub txt_COOL_TMP_MIN_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_COOL_TMP_MIN_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 50, txt_COOL_TMP_MIN_Z.Name)
    End If

End Sub

Private Sub txt_COOL_TMP_RATE_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 29, txt_COOL_TMP_RATE.Name)
    End If
End Sub

Private Sub txt_COOL_TMP_RATE_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_COOL_TMP_RATE_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 53, txt_COOL_TMP_RATE_Z.Name)
    End If

End Sub

Private Sub txt_COOL_TMP_TGT_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 28, txt_COOL_TMP_TGT.Name)
    End If
End Sub

Private Sub txt_COOL_TMP_TGT_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_COOL_TMP_TGT_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 52, txt_COOL_TMP_TGT_Z.Name)
    End If

End Sub

Private Sub txt_COOL_WAY_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 24, txt_COOL_WAY.Name)
    End If
End Sub

Private Sub txt_COOL_WAY_NAME_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 25, txt_COOL_WAY_NAME.Name)
    End If
End Sub

Private Sub txt_COOL_WAY_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 48, txt_COOL_WAY_Z.Name)
    End If

End Sub

Private Sub txt_COOL_WAY_NAME_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 49, txt_COOL_WAY_NAME_Z.Name)
    End If

End Sub


Private Sub txt_CR_CD_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 14, txt_CR_CD.Name)
    End If
End Sub


Private Sub txt_CR_CD_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 38, txt_CR_CD_Z.Name)
    End If
End Sub

Private Sub txt_CR_NAME_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 15, txt_CR_NAME.Name)
    End If
End Sub

Private Sub txt_CR_NAME_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 39, txt_CR_NAME_Z.Name)
    End If

End Sub

Private Sub txt_HCR_KND_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 9, txt_HCR_KND.Name)
    End If
End Sub


Private Sub txt_HCR_NAME_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 10, txt_HCR_NAME.Name)
    End If
End Sub

Private Sub txt_HOT_USE_Change()
    Dim hot_use_flag As String
    
    If txt_HOT_USE.Text = "Y" Or txt_HOT_USE.Text = "y" Then
        Pic_HOT_USE.Picture = ImageList1.ListImages.Item(1).Picture
        hot_use_flag = "1"
        
    Else
        Pic_HOT_USE.Picture = Nothing
        hot_use_flag = "0"
    End If
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 33, txt_HOT_USE.Name)
        Call Ms_To_SP(ss1, ss1.Row, 34, hot_use_flag)
    End If

End Sub

Private Sub txt_HOT_USE_Z_Change()
    Dim hot_use_flag As String
    
    If txt_HOT_USE_Z.Text = "Y" Or txt_HOT_USE_Z.Text = "y" Then
        Pic_HOT_USE_Z.Picture = ImageList1.ListImages.Item(1).Picture
        hot_use_flag = "1"
        
    Else
        Pic_HOT_USE_Z.Picture = Nothing
        hot_use_flag = "0"
    End If
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 57, txt_HOT_USE_Z.Name)
        Call Ms_To_SP(ss1, ss1.Row, 58, hot_use_flag)
    End If
End Sub
Private Sub txt_MILL_RATET1_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 18, txt_MILL_RATET1.Name)
    End If
End Sub
Private Sub txt_MILL_RATET1_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub
Private Sub txt_MILL_RATET1_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 42, txt_MILL_RATET1_Z.Name)
    End If

End Sub
Private Sub txt_MILL_RATET2_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 20, txt_MILL_RATET2.Name)
    End If
End Sub

Private Sub txt_MILL_RATET2_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub
Private Sub txt_MILL_RATET2_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 44, txt_MILL_RATET2_Z.Name)
    End If

End Sub

Private Sub txt_MILL_STD_NO_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 1, txt_MILL_STD_NO.Name)
    End If
End Sub

Private Sub txt_MILL_TIME_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 16, txt_MILL_TIME.Name)
    End If
End Sub

Private Sub txt_MILL_TIME_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_MILL_TIME_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 40, txt_MILL_TIME_Z.Name)
    End If

End Sub

Private Sub txt_MILL_TMP_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 22, txt_MILL_TMP_MAX.Name)
    End If
End Sub

Private Sub txt_MILL_TMP_MAX_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_MILL_TMP_MAX_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 46, txt_MILL_TMP_MAX_Z.Name)
    End If

End Sub

Private Sub txt_MILL_TMP_MIN_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 21, txt_MILL_TMP_MIN.Name)
    End If
End Sub

Private Sub txt_MILL_TMP_MIN_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_MILL_TMP_MIN_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 45, txt_MILL_TMP_MIN_Z.Name)
    End If

End Sub

Private Sub txt_MILL_TMP_TGT_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 23, txt_MILL_TMP_TGT.Name)
    End If
End Sub

Private Sub txt_MILL_TMP_TGT_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_MILL_TMP_TGT_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 47, txt_MILL_TMP_TGT_Z.Name)
    End If

End Sub

Private Sub txt_MILL_TMPT1_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 17, txt_MILL_TMPT1.Name)
    End If
End Sub

Private Sub txt_MILL_TMPT1_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_MILL_TMPT1_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 41, txt_MILL_TMPT1_Z.Name)
    End If

End Sub

Private Sub txt_MILL_TMPT2_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 19, txt_MILL_TMPT2.Name)
    End If
End Sub

Private Sub txt_MILL_TMPT2_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_MILL_TMPT2_Z_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 43, txt_MILL_TMPT2_Z.Name)
    End If

End Sub

Private Sub txt_MLT_PLT_Change()

    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 66, txt_MLT_PLT.Name)
    End If
    
    Select Case Trim(txt_MLT_PLT.Text)
        Case "C1"
            Check_C1.Value = 1
            Check_C2.Value = 0
        Case "C3"
            Check_C1.Value = 0
            Check_C2.Value = 1
        Case "**"
            Check_C1.Value = 1
            Check_C2.Value = 1
    End Select
    
    If txt_MLT_PLT = "C1" Then
' C1 COLOR SET : YELLOW
           txt_CHG_TMP_TGT.BackColor = &HC0FFFF
           txt_CHG_TMP_DEF_SC.BackColor = &HC0FFFF
           txt_CHG_TMP_DEF_TAPE.BackColor = &HC0FFFF
           txt_CR_CD.BackColor = &HC0FFFF
           txt_COOL_WAY.BackColor = &HC0FFFF
           txt_COOL_CTL_TYP.BackColor = &HC0FFFF
           txt_HOT_USE.BackColor = &HC0FFFF
          
' C3 COLOR CLEAR
           txt_CHG_TMP_TGT_Z.BackColor = &H80000005
           txt_CHG_TMP_DEF_SC_Z.BackColor = &H80000005
           txt_CHG_TMP_DEF_TAPE_Z.BackColor = &H80000005
           txt_CR_CD_Z.BackColor = &H80000005
           txt_COOL_WAY_Z.BackColor = &H80000005
           txt_COOL_CTL_TYP_Z.BackColor = &H80000005
           txt_HOT_USE_Z.BackColor = &H80000005
    
    ElseIf txt_MLT_PLT = "C3" Then
' C3 COLOR SET : YELLOW
           txt_CHG_TMP_TGT_Z.BackColor = &HC0FFFF
           txt_CHG_TMP_DEF_SC_Z.BackColor = &HC0FFFF
           txt_CHG_TMP_DEF_TAPE_Z.BackColor = &HC0FFFF
           txt_CR_CD_Z.BackColor = &HC0FFFF
           txt_COOL_WAY_Z.BackColor = &HC0FFFF
           txt_COOL_CTL_TYP_Z.BackColor = &HC0FFFF
           txt_HOT_USE_Z.BackColor = &HC0FFFF
         
' C1 COLOR CLEAR : WHITE
           txt_CHG_TMP_TGT.BackColor = &H80000005
           txt_CHG_TMP_DEF_SC.BackColor = &H80000005
           txt_CHG_TMP_DEF_TAPE.BackColor = &H80000005
           txt_CR_CD.BackColor = &H80000005
           txt_COOL_WAY.BackColor = &H80000005
           txt_COOL_CTL_TYP.BackColor = &H80000005
           txt_HOT_USE.BackColor = &H80000005
         
    ElseIf txt_MLT_PLT = "**" Then
' C1 COLOR SET : YELLOW
           txt_CHG_TMP_TGT.BackColor = &HC0FFFF
           txt_CHG_TMP_DEF_SC.BackColor = &HC0FFFF
           txt_CHG_TMP_DEF_TAPE.BackColor = &HC0FFFF
           txt_CR_CD.BackColor = &HC0FFFF
           txt_COOL_WAY.BackColor = &HC0FFFF
           txt_COOL_CTL_TYP.BackColor = &HC0FFFF
           txt_HOT_USE.BackColor = &HC0FFFF
' C3 COLOR SET : YELLOW
           txt_CHG_TMP_TGT_Z.BackColor = &HC0FFFF
           txt_CHG_TMP_DEF_SC_Z.BackColor = &HC0FFFF
           txt_CHG_TMP_DEF_TAPE_Z.BackColor = &HC0FFFF
           txt_CR_CD_Z.BackColor = &HC0FFFF
           txt_COOL_WAY_Z.BackColor = &HC0FFFF
           txt_COOL_CTL_TYP_Z.BackColor = &HC0FFFF
           txt_HOT_USE_Z.BackColor = &HC0FFFF
    End If

End Sub

Private Sub txt_STLGRD_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 7, txt_STLGRD.Name)
    End If
End Sub


Private Sub txt_MILL_STD_EDT_NO_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 59, txt_MILL_STD_EDT_NO.Name)
    End If
End Sub


Private Sub txt_STLGRD_DETAIL_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 8, txt_STLGRD_DETAIL.Name)
    End If
End Sub

Private Sub txt_THK_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 4, txt_THK_MAX.Name)
    End If
End Sub

Private Sub txt_THK_MAX_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub txt_THK_MIN_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 3, txt_THK_MIN.Name)
    End If
End Sub

Private Function txt_KeyPress(KeyAscii As Integer) As Integer

        Select Case KeyAscii
               
               Case Is <= 32
                    txt_KeyPress = KeyAscii
               Case 48 To 57
                    txt_KeyPress = KeyAscii
               Case 46
                    txt_KeyPress = KeyAscii
               Case 45
                    txt_KeyPress = KeyAscii
               Case Else
                    txt_KeyPress = 0
        End Select

    
End Function

Private Sub MS_Cls()
    Dim i As Integer
    For i = 0 To Me.COUNT - 1
        If TypeName(Me.Controls(i)) = "TextBox" Then
            Me.Controls(i).Text = ""
        ElseIf TypeName(Me.Controls(i)) = "sidbEdit" Then
            Me.Controls(i).Text = ""
        ElseIf TypeName(Me.Controls(i)) = "CheckBox" Then
            Me.Controls(i).Value = 0
        End If
        
    Next i
End Sub

Private Sub txt_THK_MIN_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

'下限值 , 上限值,目标值 Check
Private Function Sp_subMinMaxValueCheck(iRow As Long) As Boolean
    
'厚度组
    If Gf_Sp_subValueCheck(Sc1, iRow, 3, 4, ULabel3(1).Caption, txt_THK_MIN) = False Then Exit Function
'宽度组
    If Gf_Sp_subValueCheck(Sc1, iRow, 5, 6, ULabel3(2).Caption, txt_THK_MIN) = False Then Exit Function
'终轧温度
    If Gf_Sp_subValueCheck(Sc1, iRow, 21, 22, ULabel1(22).Caption + ULabel1(14).Caption + "误差", txt_MILL_TMP_MIN) = False Then Exit Function
    If Gf_Sp_subValueCheck(Sc1, iRow, 45, 46, ULabel1(37).Caption + ULabel1(28).Caption + "误差", txt_MILL_TMP_MIN_Z) = False Then Exit Function
'冷却温度
    If Gf_Sp_subValueCheck(Sc1, iRow, 26, 27, ULabel1(22).Caption + ULabel1(6).Caption + "误差", txt_COOL_TMP_MIN) = False Then Exit Function
    If Gf_Sp_subValueCheck(Sc1, iRow, 50, 51, ULabel1(37).Caption + ULabel1(30).Caption + "误差", txt_COOL_TMP_MIN) = False Then Exit Function

    Sp_subMinMaxValueCheck = True

End Function

'必须输入项目检查
Private Function Sp_AllUse_NecessaryCheck(iRow As Long) As Boolean
Dim sPLT_CD As String
        With ss1
            .Row = iRow
            .Col = 66
            sPLT_CD = .Text
        End With

'------------------------------------------------------ 共同项目 ---------------------------------------------------------

'钢种
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 7, ULabel1(1).Caption, txt_STLGRD) = False Then Exit Function

'铸坯去向
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 9, ULabel1(4).Caption, txt_HCR_KND) = False Then Exit Function
'生产工厂路线
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 66, "请选择轧钢厂！") = False Then Exit Function
    Select Case sPLT_CD
           Case "C1"
            If Sp_C1_Item_NecessaryCheck(iRow) = False Then Exit Function
           Case "C3"
            If Sp_C2_Item_NecessaryCheck(iRow) = False Then Exit Function
           Case "**"
            If Sp_C1_Item_NecessaryCheck(iRow) = False Then Exit Function
            If Sp_C2_Item_NecessaryCheck(iRow) = False Then Exit Function
    End Select
    
    Sp_AllUse_NecessaryCheck = True
End Function

Private Function Sp_C1_Item_NecessaryCheck(iRow As Long) As Boolean
'平均出炉温差
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 11, "平均出炉温差", txt_CHG_TMP_TGT, True) = False Then Exit Function
    
'板坯表面/中心温差
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 12, "板坯表面/中心温差", txt_CHG_TMP_DEF_SC, True) = False Then Exit Function
    
'板坯头尾温差
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 13, "板坯头尾温差", txt_CHG_TMP_DEF_TAPE, True) = False Then Exit Function
    
'控轧代码
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 14, "控轧代码", txt_CR_CD, True) = False Then Exit Function
    
'T1温度/压下率&终轧温度
    With ss1
        .Row = iRow
        .Col = 14
        If .Text = "Y" Or .Text = "y" Then
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 17, "T1温度", txt_MILL_TMPT1, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 18, "T1温度/压下率", txt_MILL_RATET1, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 21, "终轧温度误差下限", txt_MILL_TMP_MIN, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 22, "终轧温度误差上限", txt_MILL_TMP_MAX, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 23, "终轧温度目标值", txt_MILL_TMP_TGT, True) = False Then Exit Function
        End If
    End With

'冷却方法代码
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 24, "冷却方法代码", txt_COOL_WAY, True) = False Then Exit Function
    
'冷却温度/冷却速率
    With ss1
        .Row = iRow
        .Col = 24
        If .Text = "W" Or .Text = "w" Then
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 26, "冷却温度误差下限", txt_COOL_TMP_MIN, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 27, "冷却温度误差上限", txt_COOL_TMP_MAX, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 26, "冷却温度目标值", txt_COOL_TMP_TGT, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 27, "冷却速率", txt_COOL_TMP_RATE, True) = False Then Exit Function
        End If
    End With
'控冷代码
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 31, "控冷代码", txt_COOL_CTL_TYP, True) = False Then Exit Function
'热矫直代码
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 33, "热矫直代码", txt_HOT_USE, True) = False Then Exit Function
    
    Sp_C1_Item_NecessaryCheck = True
    
End Function

Private Function Sp_C2_Item_NecessaryCheck(iRow As Long) As Boolean
'平均出炉温差
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 35, "平均出炉温差", txt_CHG_TMP_TGT_Z, True) = False Then Exit Function
    
'板坯表面/中心温差
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 36, "板坯表面/中心温差", txt_CHG_TMP_DEF_SC_Z, True) = False Then Exit Function
    
'板坯头尾温差
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 37, "板坯头尾温差", txt_CHG_TMP_DEF_TAPE_Z, True) = False Then Exit Function
    
'控轧代码
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 38, "控轧代码", txt_CR_CD_Z, True) = False Then Exit Function
    
'T1温度/压下率&终轧温度
    With ss1
        .Row = iRow
        .Col = 38
        If .Text = "Y" Or .Text = "y" Then
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 41, "T1温度", txt_MILL_TMPT1_Z, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 42, "T1温度/压下率", txt_MILL_RATET1_Z, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 45, "终轧温度误差下限", txt_MILL_TMP_MIN_Z, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 46, "终轧温度误差上限", txt_MILL_TMP_MAX_Z, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 47, "终轧温度目标值", txt_MILL_TMP_TGT_Z, True) = False Then Exit Function
        End If
    End With

'冷却方法代码
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 48, "冷却方法代码", txt_COOL_WAY_Z, True) = False Then Exit Function
    
'冷却温度/冷却速率
    With ss1
        .Row = iRow
        .Col = 48
        If .Text = "W" Or .Text = "w" Then
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 50, "冷却温度误差下限", txt_COOL_TMP_MIN_Z, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 51, "冷却温度误差上限", txt_COOL_TMP_MAX_Z, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 52, "冷却温度目标值", txt_COOL_TMP_TGT_Z, True) = False Then Exit Function
            If GF_Sp_Necessary_Value_Check(Sc1, iRow, 53, "冷却速率", txt_COOL_TMP_RATE_Z, True) = False Then Exit Function
        End If
    End With
'控冷代码
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 55, "控冷代码", txt_COOL_CTL_TYP_Z, True) = False Then Exit Function
'热矫直代码
    If GF_Sp_Necessary_Value_Check(Sc1, iRow, 57, "热矫直代码", txt_HOT_USE_Z, True) = False Then Exit Function
    
    Sp_C2_Item_NecessaryCheck = True
    
End Function

Private Sub txt_WID_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 6, txt_WID_MAX.Name)
    End If
End Sub

Private Sub txt_WID_MIN_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) Then
        Call Ms_To_SP(ss1, ss1.Row, 5, txt_WID_MIN.Name)
    End If
End Sub
