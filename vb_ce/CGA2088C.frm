VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form CGA2088C 
   Caption         =   "中板厂外板坯切割作业界面_CGA2088C"
   ClientHeight    =   9225
   ClientLeft      =   675
   ClientTop       =   2235
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15315
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7905
      Left            =   60
      TabIndex        =   21
      Top             =   1260
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   13944
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "CGA2088C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   3375
         Left            =   0
         TabIndex        =   22
         Top             =   4530
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   5953
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   14737632
         PaneTree        =   "CGA2088C.frx":0052
         Begin Threed.SSPanel SSPanel1 
            Height          =   540
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   953
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.ComboBox cbo_cutcnt 
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
               ItemData        =   "CGA2088C.frx":00A4
               Left            =   1515
               List            =   "CGA2088C.frx":00A6
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Tag             =   "连铸机号"
               Top             =   120
               Width           =   705
            End
            Begin VB.TextBox TXT_SLABNO 
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
               Left            =   12360
               MaxLength       =   10
               TabIndex        =   25
               Top             =   120
               Visible         =   0   'False
               Width           =   2040
            End
            Begin InDate.ULabel ULabel4 
               Height          =   315
               Left            =   180
               Top             =   120
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
               Caption         =   "切割块数"
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
               Left            =   2340
               Top             =   120
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
               Caption         =   "总长度"
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
            Begin CSTextLibCtl.sidbEdit txt_total_len 
               Height          =   315
               Left            =   3660
               TabIndex        =   27
               Top             =   120
               Width           =   915
               _Version        =   262145
               _ExtentX        =   1614
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   255
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderEffect    =   2
               DataProperty    =   1
               ReadOnly        =   -1  'True
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
               MaxValue        =   20
               MinValue        =   10
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel11 
               Height          =   315
               Left            =   4770
               Top             =   120
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
               Caption         =   "总重量"
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
            Begin CSTextLibCtl.sidbEdit txt_total_wgt 
               Height          =   315
               Left            =   6090
               TabIndex        =   28
               Top             =   120
               Width           =   915
               _Version        =   262145
               _ExtentX        =   1614
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   255
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderEffect    =   2
               DataProperty    =   1
               ReadOnly        =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   "0.000"
               Text            =   " 0.000"
               StartText.x     =   3
               StartText.y     =   3
               FirstVisPos     =   0
               HiAnchor        =   0
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
               NumIntDigits    =   4
               MaxValue        =   20
               MinValue        =   10
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel12 
               Height          =   315
               Left            =   7200
               Top             =   120
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
               Caption         =   "废钢重量"
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
            Begin CSTextLibCtl.sidbEdit txt_scrap_wgt 
               Height          =   315
               Left            =   8520
               TabIndex        =   29
               Top             =   120
               Width           =   915
               _Version        =   262145
               _ExtentX        =   1614
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   255
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderEffect    =   2
               DataProperty    =   1
               ReadOnly        =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   "0.000"
               Text            =   " 0.000"
               StartText.x     =   3
               StartText.y     =   3
               FirstVisPos     =   0
               HiAnchor        =   0
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
               NumIntDigits    =   4
               MaxValue        =   20
               MinValue        =   10
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_tmCalMo 
               Height          =   315
               Left            =   10950
               TabIndex        =   30
               Top             =   120
               Visible         =   0   'False
               Width           =   915
               _Version        =   262145
               _ExtentX        =   1614
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   255
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.76
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderEffect    =   2
               DataProperty    =   1
               ReadOnly        =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   "0.000"
               Text            =   " 0.000"
               StartText.x     =   3
               StartText.y     =   3
               FirstVisPos     =   0
               HiAnchor        =   0
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
               NumIntDigits    =   4
               MaxValue        =   20
               MinValue        =   10
               Undo            =   0
               Data            =   0
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   2805
            Left            =   0
            TabIndex        =   31
            Top             =   570
            Width           =   15195
            _Version        =   393216
            _ExtentX        =   26802
            _ExtentY        =   4948
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   15
            MaxRows         =   2
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CGA2088C.frx":00A8
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   4470
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   15195
         _Version        =   393216
         _ExtentX        =   26802
         _ExtentY        =   7885
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   21
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGA2088C.frx":0953
      End
   End
   Begin Threed.SSOption opt_prc_status1 
      Height          =   285
      Left            =   1260
      TabIndex        =   19
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
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
      Caption         =   "板坯切割"
      Value           =   -1
   End
   Begin VB.TextBox txt_cur_name 
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
      Left            =   10560
      MaxLength       =   11
      TabIndex        =   18
      Text            =   "新库"
      Top             =   480
      Width           =   2400
   End
   Begin VB.TextBox txt_cur_inv 
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
      Left            =   10080
      MaxLength       =   2
      TabIndex        =   17
      Text            =   "XK"
      Top             =   480
      Width           =   480
   End
   Begin VB.TextBox txt_act_stlgrd 
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
      Left            =   5100
      MaxLength       =   11
      TabIndex        =   7
      Top             =   480
      Width           =   1500
   End
   Begin VB.TextBox txt_MOSLAB 
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
      Left            =   5100
      MaxLength       =   10
      TabIndex        =   6
      Top             =   90
      Width           =   1350
   End
   Begin VB.TextBox txt_act_stlgrd_dec 
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
      Left            =   6600
      TabIndex        =   5
      Top             =   480
      Width           =   2010
   End
   Begin VB.TextBox txt_Status 
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
      Left            =   3420
      MaxLength       =   11
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   270
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
      Left            =   1260
      MaxLength       =   2
      TabIndex        =   3
      Top             =   480
      Width           =   540
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
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   2
      Top             =   480
      Width           =   1260
   End
   Begin VB.TextBox txt_tmpPLT 
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
      Left            =   14100
      MaxLength       =   20
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txt_IST_DATE 
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
      Left            =   14115
      MaxLength       =   20
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   870
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   3960
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   3960
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "母板坯号"
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
      Left            =   120
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "处理分类"
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
      Left            =   120
      Top             =   870
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "厚度"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   3960
      Top             =   870
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "宽度"
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   8955
      Top             =   870
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "长度"
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   120
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "生产工厂"
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
   Begin CSTextLibCtl.sidbEdit txt_wid 
      Height          =   315
      Left            =   5100
      TabIndex        =   8
      Top             =   870
      Width           =   990
      _Version        =   262145
      _ExtentX        =   1746
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
      DataProperty    =   1
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   " 0.00"
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
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
      NumIntDigits    =   4
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_len 
      Height          =   315
      Left            =   10080
      TabIndex        =   9
      Top             =   870
      Width           =   1080
      _Version        =   262145
      _ExtentX        =   1905
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
      DataProperty    =   1
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   " 0.0"
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
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
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_thk 
      Height          =   315
      Left            =   1260
      TabIndex        =   10
      Top             =   870
      Width           =   870
      _Version        =   262145
      _ExtentX        =   1535
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
      DataProperty    =   1
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
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_thk_to 
      Height          =   315
      Left            =   2130
      TabIndex        =   11
      Top             =   870
      Width           =   870
      _Version        =   262145
      _ExtentX        =   1535
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
      DataProperty    =   1
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
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_wid_to 
      Height          =   315
      Left            =   6090
      TabIndex        =   12
      Top             =   870
      Width           =   990
      _Version        =   262145
      _ExtentX        =   1746
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
      DataProperty    =   1
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   " 0.00"
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
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
      NumIntDigits    =   4
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_len_to 
      Height          =   315
      Left            =   11160
      TabIndex        =   13
      Top             =   870
      Width           =   1080
      _Version        =   262145
      _ExtentX        =   1905
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
      DataProperty    =   1
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   " 0.0"
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
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
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   375
      Left            =   13920
      TabIndex        =   14
      Top             =   810
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   14737632
      BackStyle       =   1
      ActiveColors    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "指示取消"
   End
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   8955
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "指示日期"
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
   Begin InDate.UDate U_FROM_DATE 
      Height          =   315
      Left            =   10080
      TabIndex        =   15
      Tag             =   "起始日期"
      Top             =   90
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
   End
   Begin InDate.UDate U_TO_DATE 
      Height          =   315
      Left            =   11520
      TabIndex        =   16
      Tag             =   "起始日期"
      Top             =   75
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
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   8955
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "堆放仓库"
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
   Begin Threed.SSOption opt_prc_status2 
      Height          =   285
      Left            =   2430
      TabIndex        =   20
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
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
      Caption         =   "查询与修改"
   End
End
Attribute VB_Name = "CGA2088C"
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
'-- Program Name      中板厂外板坯切割作业界面
'-- Program ID        CGA2088c
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2009.8.7
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String            'Form Type
Public Toolbar_St As String          'Active Form ToolBar Setting
Public sAuthority As String          'Active Form Authority Setting
Public sDateTime As String           'Active Form Authority Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pControl2 As New Collection       'Master Primary Key Collection
Dim nControl2 As New Collection       'Master Necessary Collection
Dim mControl2 As New Collection       'Master Maxlength check Collection
Dim iControl2 As New Collection       'Master Insert Collection
Dim rControl2 As New Collection       'Master Refer Collection
Dim cControl2 As New Collection       'Master Copy Collection
Dim aControl2 As New Collection       'Master -> Spread Collection
Dim lControl2 As New Collection       'Master Lock Collection

Dim pColumn As New Collection        'Spread Primary Key Collection
Dim nColumn As New Collection        'Spread necessary Column Collection
Dim mColumn As New Collection        'Spread Maxlength check Column Collection
Dim iColumn As New Collection        'Spread Insert Column Collection
Dim aColumn As New Collection        'Master -> Spread Column Collection
Dim lColumn As New Collection        'Spread Lock Column Collection

Dim pColumn1 As New Collection       'Spread Primary Key Collection
Dim nColumn1 As New Collection       'Spread necessary Column Collection
Dim mColumn1 As New Collection       'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection       'Spread Insert Column Collection
Dim aColumn1 As New Collection       'Master -> Spread Column Collection
Dim lColumn1 As New Collection       'Spread Lock Column Collection

Dim pColumn2 As New Collection       'Spread Primary Key Collection
Dim nColumn2 As New Collection       'Spread necessary Column Collection
Dim mColumn2 As New Collection       'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection       'Spread Insert Column Collection
Dim aColumn2 As New Collection       'Master -> Spread Column Collection
Dim lColumn2 As New Collection       'Spread Lock Column Collection

Dim Mc1 As New Collection            'Master Collection
Dim Mc2 As New Collection

Dim sc1 As New Collection            'Spread Collection
Dim sc2 As New Collection            'Spread Collection
Dim Proc_Sc As New Collection        'Spread Struc Collection

Dim lBlkcol1 As Long                 'To Excel Block Col1
Dim lBlkcol2 As Long                 'To Excel Block Col2
Dim lBlkrow1 As Long                 'To Excel Block Row1
Dim lBlkrow2 As Long                 'To Excel Block Row2

'DOTHER SLAB LENGTH,WGT CUALUCATE


Public cSlabno As String              'dother slab no
Public cSlabthk As Double
Public cSlabwid As Double
Public cSlabLen As Double              'Mother Slab Length
Public cSlabWgt As Double              'Mother Slab Wgt
Public cSlabCalWgt As Double           'Mother Slab Cal Wgt
Public cStlgrd As String
Public cOrdno As String
Public cProddate As String
Public cLoc As String
Public cRcvDate As String
Public tmWgt As Double
Public tmpSlabNo As String
Public NEWSLABNO As String
Public cfLen, cfWgt, cfCalWgt As Double
Public addSlabNo As String
Public lCurrRow As Long
Public SCRAP_NO As String


Dim sQuery As String

'Public Sub Form_Ins()
'    If ss2.SelBlockRow2 = ss2.MaxRows Then
'       ss2.Row = ss2.MaxRows
'       ss2.Col = 0
'       If ss2.Text <> "Delete" Then
'            Call Gp_Sp_Ins(Proc_Sc("Sc2"))
'
'            With ss1
'                .Row = .ActiveRow
'                .Col = 8
'                .Text = sUserID
'            End With
'
'            Call INS_WGT_CAL
'        End If
'    End If
'
'End Sub
Public Sub Spread_Del()
Dim i%
       For i = 1 To ss2.MaxRows
           ss2.Row = i
           ss2.Col = 0
           If UCase(ss2.Text) = "" Then
              ss2.Text = "Delete"
           End If
       Next i

End Sub

Public Sub Spread_Can()

    If ss2.SelBlockRow2 = ss2.MaxRows Then
       ss2.Row = ss2.MaxRows
       ss2.Col = 0
       If ss2.Text = "Input" Then
            ss2.MaxRows = ss2.MaxRows - 1
            addSlabNo = Mid(addSlabNo, 1, 8) & CStr(CInt(Mid(addSlabNo, 9, 2)) - 1)
            Call CANCEL_WGT_CAL
       End If
    End If
End Sub

Public Sub WGT_CAL()
Dim tmThk As Double
Dim tmWid As Double
Dim tmLen As Double
Dim tempWgt As Double
Dim tot_cal_total As Double
Dim cal_wgt As Double
Dim tmp_rat As Double
Dim tmTotalLen As Double
Dim tmpLen As Double
Dim sub_wgt As Double
Dim sub_len As Double
Dim tmCalCut As Double
Dim tmCalMo As Double
Dim tmCalCutOne As Double



Dim i As Integer

    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    tmCalMo = cSlabthk * cSlabwid * cSlabLen
    
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 0
        If ss2.Text <> "Delete" Then
            ss2.Row = i
            ss2.Col = 2
            tmThk = ss2.Value
            ss2.Col = 3
            tmWid = ss2.Value
            ss2.Col = 4
            tmLen = ss2.Value
            tmTotalLen = tmTotalLen + ss2.Value
            
            tmCalCut = tmCalCut + (tmThk * tmWid * tmLen)
        End If
    Next i
        
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 0
        If ss2.Text <> "Delete" Then
            ss2.Row = i
            ss2.Col = 2
            tmThk = ss2.Value
            ss2.Col = 3
            tmWid = ss2.Value
            ss2.Col = 4
            tmLen = ss2.Value
            
            tmCalCutOne = tmThk * tmWid * tmLen
            
            ss2.Col = 5
            If tmCalCut <= tmCalMo Then
                tempWgt = tempWgt + Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
                sub_wgt = sub_wgt - Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
                ss2.Value = Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
            Else
                tempWgt = tempWgt + Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
                sub_wgt = sub_wgt - Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
                ss2.Value = Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
            End If
            
            ss2.Col = 6
            ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
        End If
    Next i
    
    If tmCalCut = tmCalMo Then
        sub_len = cSlabLen
        sub_wgt = cSlabWgt
        For i = 1 To ss2.MaxRows
            ss2.Row = i
            If i < ss2.MaxRows Then
               ss2.Col = 5
               sub_wgt = sub_wgt - ss2.Value
            End If
        Next i
        ss2.Row = ss2.MaxRows

        ss2.Col = 5
        ss2.Text = sub_wgt
    End If
    
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 0
        If ss2.Text <> "Delete" Then
            ss2.Row = i
            ss2.Col = 4
            tmTotalLen = tmTotalLen + ss2.Value
            
            ss2.Col = 5
            tempWgt = tempWgt + ss2.Value
        End If

    Next i
    
    For i = 1 To ss2.MaxRows
         ss2.Row = i
         ss2.Col = 0
         If UCase(ss2.Text) = "" Then
            ss2.Text = "Update"
         End If
    Next i
    
    
    If tmTotalLen = cSlabLen Then
       txt_total_len.ForeColor = &H0&
    Else
       txt_total_len.ForeColor = &HFF&
    End If
    txt_total_len = tmTotalLen
    
    txt_total_wgt = tempWgt
    If CDbl(txt_total_wgt) - cSlabWgt = 0 Then
       txt_total_wgt.ForeColor = &H0&
    Else
       txt_total_wgt.ForeColor = &HFF&
    End If
    
    txt_scrap_wgt = Format(cSlabWgt - tempWgt, "###0.000")
    If cSlabWgt - tempWgt = 0 Then
       txt_scrap_wgt.ForeColor = &H0&
    Else
       txt_scrap_wgt.ForeColor = &HFF&
    End If

       
       
End Sub
Public Sub INS_WGT_CAL()
Dim tmThk As Double
Dim tmWid As Double
Dim tmLen As Double
Dim tempWgt As Double
Dim tot_cal_total As Double
Dim cal_wgt As Double
Dim tmp_rat As Double
Dim tmTotalLen As Double
Dim tmpLen As Double
Dim sub_wgt As Double
Dim sub_len As Double
Dim S1 As String
Dim S2 As Double
Dim S3 As Double
Dim S4 As Double
Dim S5 As Double
Dim S6 As Double
Dim S7 As String
Dim S8 As String
Dim S9 As String
Dim S10 As String
Dim S11 As String
Dim S12 As String

Dim i, delete_cnt As Integer
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&

    delete_cnt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 0
        If UCase(ss2.Text) <> "DELETE" Then
           delete_cnt = delete_cnt + 1
        End If
    Next i

    cfLen = Format(cSlabLen / ss2.MaxRows, "####0")
    cfWgt = Round(cSlabWgt / ss2.MaxRows, 3)
    cfCalWgt = Round(cSlabCalWgt / ss2.MaxRows, 3)
        
    ' DATA COPY
    ss2.Row = ss2.MaxRows - 1
    
    ss2.Col = 1
    S1 = ss2.Value
    
    ss2.Col = 2
    S2 = ss2.Value
    
    ss2.Col = 3
    S3 = ss2.Value
    
    ss2.Col = 4
    S4 = ss2.Value
    
    ss2.Col = 5
    S5 = ss2.Value
    
    ss2.Col = 6
    S6 = ss2.Value
    
    ss2.Col = 7
    S7 = ss2.Text
    
    ss2.Col = 8
    S8 = ss2.Text
    
    ss2.Col = 9
    S9 = ss2.Text
    
    ss2.Col = 10
    S10 = ss2.Text
    
    ss2.Col = 11
    S11 = ss2.Text
    
    ss2.Col = 12
    S12 = ss2.Text

    ' DATA PAST
    ss2.Row = ss2.MaxRows
    ss2.Col = 1
    ss2.Text = addSlabNo
    addSlabNo = Mid(addSlabNo, 1, 8) & CStr(CInt(Mid(addSlabNo, 9, 2)) + 1)
    ss2.Col = 2
    ss2.Text = S2
    tmThk = S2
    ss2.Col = 3
    ss2.Text = S3
    tmWid = S3
    ss2.Col = 4
    ss2.Text = S4
    tmLen = S4
    ss2.Col = 5
    ss2.Text = S5
    ss2.Col = 6
    ss2.Text = S6
    ss2.Col = 7
    ss2.Text = S7
    ss2.Col = 8
    ss2.Text = S8
    ss2.Col = 9
    ss2.Text = S9
    ss2.Col = 10
    ss2.Text = S10
    ss2.Col = 11
    ss2.Text = S11
    ss2.Col = 12
    ss2.Text = S12
    
    tmp_rat = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
         ss2.Row = i
         ss2.Col = 4
         ss2.Text = cSlabLen / ss2.MaxRows
         ss2.Col = 5
         ss2.Text = Round(cSlabWgt * ((cSlabLen / ss2.MaxRows) / cSlabLen), 3)
         tempWgt = tempWgt + Round(cSlabWgt * ((cSlabLen / ss2.MaxRows) / cSlabLen), 3)
         sub_wgt = sub_wgt - Round(cSlabWgt * ((cSlabLen / ss2.MaxRows) / cSlabLen), 3)
         
         ss2.Col = 6
         ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
    Next i
    
    sub_len = cSlabLen
    sub_wgt = cSlabWgt
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        If i <> ss2.MaxRows Then
           ss2.Col = 4
           sub_len = sub_len - ss2.Value
           
           ss2.Col = 5
           sub_wgt = sub_wgt - ss2.Value
        End If
    Next i
    ss2.Row = ss2.MaxRows
    ss2.Col = 4
    ss2.Text = sub_len
    
    ss2.Col = 5
    ss2.Text = sub_wgt
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 4
        tmTotalLen = tmTotalLen + ss2.Value
        
        ss2.Col = 5
        tempWgt = tempWgt + ss2.Value

    Next i
    
    If tmTotalLen = cSlabLen Then
       txt_total_len.ForeColor = &H0&
    Else
       txt_total_len.ForeColor = &HFF&
    End If
    txt_total_len = tmTotalLen
    
    txt_total_wgt = tempWgt
    If CDbl(txt_total_wgt) = cSlabWgt Then
       txt_total_wgt.ForeColor = &H0&
    Else
       txt_total_wgt.ForeColor = &HFF&
    End If

    
    txt_scrap_wgt = Format(cSlabWgt - tempWgt, "###0.000")
    If cSlabWgt - tempWgt = 0 Then
       txt_scrap_wgt.ForeColor = &H0&
    Else
       txt_scrap_wgt.ForeColor = &HFF&
    End If


    
    For i = 1 To ss2.MaxRows
         ss2.Row = i
         ss2.Col = 0
         If UCase(ss2.Text) = "" Then
            ss2.Text = "Update"
         End If
    Next i
        
End Sub
Public Sub DEL_WGT_CAL()
Dim tmThk As Double
Dim tmWid As Double
Dim tmLen As Double
Dim tempWgt As Double
Dim tot_cal_total As Double
Dim cal_wgt As Double
Dim tmp_rat As Double
Dim tmTotalLen As Double
Dim tmpLen As Double
Dim sub_wgt As Double
Dim sub_len As Double


Dim i, delete_cnt As Integer
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&

    delete_cnt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 0
        If UCase(ss2.Text) <> "DELETE" Then
           delete_cnt = delete_cnt + 1
        End If
    Next i

    cfLen = Format(cSlabLen / delete_cnt, "####0")
    cfWgt = Round(cSlabWgt / delete_cnt, 3)
    cfCalWgt = Round(cSlabCalWgt / delete_cnt, 3)
    
        
    tempWgt = 0
    For i = 1 To delete_cnt
         ss2.Row = i
         
         ss2.Col = 2
         tmThk = ss2.Value
         
         ss2.Col = 3
         tmWid = ss2.Value
         
         ss2.Col = 4
         ss2.Text = cfLen
         tmLen = cfLen
         
         ss2.Col = 4
         If ss2.Row = ss2.MaxRows Then
            ss2.Text = cSlabLen - tmTotalLen
            tmTotalLen = tmTotalLen + ss2.Text
         Else
            ss2.Text = cfLen
            tmLen = cfLen
            tmTotalLen = tmTotalLen + cfLen
         End If
         
         ss2.Col = 5
         ss2.Text = cfWgt
         tmWgt = tmWgt + ss2.Value
         
         ss2.Col = 6
         ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
    Next i
    
    sub_len = cSlabLen
    sub_wgt = cSlabWgt
    For i = 1 To delete_cnt
        ss2.Row = i
        If i <> delete_cnt Then
           ss2.Col = 4
           sub_len = sub_len - ss2.Value
           
           ss2.Col = 5
           sub_wgt = sub_wgt - ss2.Value
        End If
    Next i
    ss2.Row = delete_cnt
    ss2.Col = 4
    ss2.Text = sub_len
    
    ss2.Col = 5
    ss2.Text = sub_wgt
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To delete_cnt
        ss2.Row = i
        ss2.Col = 4
        tmTotalLen = tmTotalLen + ss2.Value
        
        ss2.Col = 5
        tempWgt = tempWgt + ss2.Value

    Next i
    
    For i = 1 To ss2.MaxRows
         ss2.Row = i
         ss2.Col = 0
         If UCase(ss2.Text) = "" Then
            ss2.Text = "Update"
         End If
    Next i
    
    If tmTotalLen = cSlabLen Then
       txt_total_len.ForeColor = &H0&
    Else
       txt_total_len.ForeColor = &HFF&
    End If
    txt_total_len = tmTotalLen
    
    txt_total_wgt = tempWgt
    If CDbl(txt_total_wgt) = cSlabWgt Then
       txt_total_wgt.ForeColor = &H0&
    Else
       txt_total_wgt.ForeColor = &HFF&
    End If

    txt_scrap_wgt = Format(cSlabWgt - tempWgt, "###0.000")
    If cSlabWgt - tempWgt = 0 Then
       txt_scrap_wgt.ForeColor = &H0&
    Else
       txt_scrap_wgt.ForeColor = &HFF&
    End If

       

End Sub

Public Sub CANCEL_WGT_CAL()
Dim tmThk As Double
Dim tmWid As Double
Dim tmLen As Double
Dim tempWgt As Double
Dim tot_cal_total As Double
Dim cal_wgt As Double
Dim tmp_rat As Double
Dim tmTotalLen As Double
Dim tmpLen As Double
Dim sub_wgt As Double
Dim sub_len As Double
Dim i As Integer

    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    tmTotalLen = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        
        ss2.Col = 4
        ss2.Text = cSlabLen / ss2.MaxRows
        tmLen = cSlabLen / ss2.MaxRows
        tmTotalLen = tmTotalLen + tmLen
        
    Next i
    
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 2
        tmThk = ss2.Value
        
        ss2.Col = 3
        tmWid = ss2.Value
        
        ss2.Col = 4
        tmLen = ss2.Value
        
        ss2.Col = 5
        ss2.Text = Round(cSlabWgt * (tmLen / tmTotalLen), 3)
        
        ss2.Col = 6
        ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
    Next i
    
    sub_len = cSlabLen
    sub_wgt = cSlabWgt
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        If i <> ss2.MaxRows Then
           ss2.Col = 4
           sub_len = sub_len - ss2.Value
           
           ss2.Col = 5
           sub_wgt = sub_wgt - ss2.Value
        End If
    Next i
    
    ss2.Row = ss2.MaxRows
    ss2.Col = 4
    ss2.Text = sub_len
    
    ss2.Col = 5
    ss2.Text = sub_wgt
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 4
        tmTotalLen = tmTotalLen + ss2.Value
        
        ss2.Col = 5
        tempWgt = tempWgt + ss2.Value

    Next i
    If tmTotalLen = cSlabLen Then
       txt_total_len.ForeColor = &H0&
    Else
       txt_total_len.ForeColor = &HFF&
    End If
    txt_total_len = tmTotalLen
    
    txt_total_wgt = tempWgt
    If CDbl(txt_total_wgt) = cSlabWgt Then
       txt_total_wgt.ForeColor = &H0&
    Else
       txt_total_wgt.ForeColor = &HFF&
    End If

    txt_scrap_wgt = Format(cSlabWgt - tempWgt, "###0.000")
    If cSlabWgt - tempWgt = 0 Then
       txt_scrap_wgt.ForeColor = &H0&
    Else
       txt_scrap_wgt.ForeColor = &HFF&
    End If
    
End Sub

Public Sub LENMODIFY_WGT_CAL(ByVal Col As Long, ByVal Row As Long)
Dim tmThk As Double
Dim tmWid As Double
Dim tmLen As Double
Dim tempWgt As Double
Dim tot_cal_total As Double
Dim cal_wgt As Double
Dim tmp_rat As Double
Dim tmTotalLen As Double
Dim tmpLen As Double
Dim sub_wgt As Double
Dim sub_len As Double
Dim i As Integer

Dim tmCalCut As Double
Dim tmCalMo As Double
Dim tmCalCutOne As Double

    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    tmCalMo = cSlabthk * cSlabwid * cSlabLen
    
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 0
        If UCase(ss2.Text) <> "DELETE" Then
            ss2.Col = 4
            tmTotalLen = tmTotalLen + ss2.Value
        End If
    Next i
    
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 0
        If UCase(ss2.Text) <> "DELETE" Then
            ss2.Row = i
            ss2.Col = 2
            tmThk = ss2.Value
            ss2.Col = 3
            tmWid = ss2.Value
            ss2.Col = 4
            tmLen = ss2.Value
            
            tmCalCut = tmCalCut + (tmThk * tmWid * tmLen)
        End If
    Next i
        
    tmp_rat = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 2
        tmThk = ss2.Value
        ss2.Col = 3
        tmWid = ss2.Value
        ss2.Col = 4
        tmLen = ss2.Value
        
        tmCalCutOne = tmThk * tmWid * tmLen
        
        ss2.Col = 5
        If tmCalCut <= tmCalMo Then
            tempWgt = tempWgt + Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
            sub_wgt = sub_wgt - Round((cSlabWgt * (cfLen / cSlabLen)), 3)
            ss2.Value = Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
        Else
            tempWgt = tempWgt + Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
            sub_wgt = sub_wgt - Round((cSlabWgt * (cfLen / tmTotalLen)), 3)
            ss2.Value = Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
        End If
        
        ss2.Col = 6
        ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
    Next i
    
    
    If tmCalCut >= tmCalMo Then
        sub_len = cSlabLen
        sub_wgt = cSlabWgt
        For i = 1 To ss2.MaxRows
            ss2.Row = i
            If i <> ss2.MaxRows Then
               ss2.Col = 5
               sub_wgt = sub_wgt - ss2.Value
            End If
        Next i
        ss2.Row = ss2.MaxRows

        ss2.Col = 5
        ss2.Text = sub_wgt
    End If
    
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 4
        tmTotalLen = tmTotalLen + ss2.Value
        
        ss2.Col = 5
        tempWgt = tempWgt + ss2.Value

    Next i
    
    If tmTotalLen = cSlabLen Then
       txt_total_len.ForeColor = &H0&
    Else
       txt_total_len.ForeColor = &HFF&
    End If
    txt_total_len = tmTotalLen
    
    txt_total_wgt = tempWgt
    If CDbl(txt_total_wgt) = cSlabWgt Then
       txt_total_wgt.ForeColor = &H0&
    Else
       txt_total_wgt.ForeColor = &HFF&
    End If

    txt_scrap_wgt = Format(cSlabWgt - tempWgt, "###0.000")
    If cSlabWgt - tempWgt = 0 Then
       txt_scrap_wgt.ForeColor = &H0&
    Else
       txt_scrap_wgt.ForeColor = &HFF&
    End If
    
       
       

End Sub

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(txt_Status, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_act_stlgrd, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_MOSLAB, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_cur_inv, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_plt, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_thk, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_thk_to, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_wid, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_wid_to, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_len, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_len_to, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(U_FROM_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(U_TO_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGA2088C.P_REFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    
    Call Gp_Ms_Collection(txt_tmpPLT, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  Call Gp_Ms_Collection(txt_IST_DATE, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(TXT_SLABNO, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    'MASTER Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, "p", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CGA2088C.P_MODIFY1", Key:="P-M"
    sc2.Add Item:="CGA2088C.P_REFER1", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    Call Gp_Sp_ColHidden(ss2, 6, True)
    'Call Gp_Sp_ColHidden(ss2, 10, True)
    Call Gp_Sp_ColHidden(ss2, 12, True)
    Call Gp_Sp_ColHidden(ss2, 13, True)

    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub


Private Sub cbo_cutcnt_Click()
Dim i, j As Integer

Dim tmThk, tmWid, tmLen As Double
Dim tmTotalLen As Double
Dim tmpLen As Double

    
Dim tempWgt As Double
Dim tot_cal_total As Double
Dim cal_wgt As Double
Dim sub_wgt As Double
Dim tmp_rat As Double

    If TXT_SLABNO.Text = "" Then Exit Sub
    
    If cbo_cutcnt.ListIndex = 0 Then Exit Sub
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    ss2.MaxRows = 0
    ss2.MaxRows = CInt(cbo_cutcnt)
    
    sQuery = "          SELECT MAX(SLAB_NO) "
    sQuery = sQuery & "   FROM NISCO.FP_SLAB "
    sQuery = sQuery & "  WHERE SLAB_NO LIKE '" & Mid(cSlabno, 1, 8) & "%'"
    
    tmpSlabNo = Gf_CodeFind(M_CN1, sQuery)
    If CInt(Mid(tmpSlabNo, 9, 2)) < 30 Then
       tmpSlabNo = Mid(tmpSlabNo, 1, 8) & "30"
    End If
    

    cfLen = Format(cSlabLen / cbo_cutcnt, "####0")
    cfWgt = Round(cSlabWgt / cbo_cutcnt, 3)
    cfCalWgt = Round(cSlabCalWgt / cbo_cutcnt, 3)
    
    For i = 1 To cbo_cutcnt
        ss2.Row = i
        ss2.Col = 1
        
        NEWSLABNO = Mid(tmpSlabNo, 1, 4) & Mid(tmpSlabNo, 5, 6) + i
        If Len(Mid(NEWSLABNO, 5, 6)) = 5 Then
           NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "0" & Mid(NEWSLABNO, 5, 5)
        ElseIf Len(Mid(NEWSLABNO, 5, 6)) = 4 Then
           NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "00" & Mid(NEWSLABNO, 5, 5)
        ElseIf Len(Mid(NEWSLABNO, 5, 6)) = 3 Then
           NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "000" & Mid(NEWSLABNO, 5, 5)
        End If
        
        ss2.Text = NEWSLABNO
    
        ss2.Col = 2
        ss2.Text = cSlabthk
        tmThk = cSlabthk
    
        ss2.Col = 3
        ss2.Text = cSlabwid
        tmWid = cSlabwid
    
        ss2.Col = 4
        If ss2.Row = ss2.MaxRows Then
            ss2.Text = cSlabLen - tmTotalLen
            tmTotalLen = tmTotalLen + ss2.Text
        Else
            ss2.Text = cfLen
            tmLen = cfLen
            tmTotalLen = tmTotalLen + cfLen
        End If
        
        ss2.Col = 5
        ss2.Text = cfWgt
        tmWgt = tmWgt + ss2.Value
    
        ss2.Col = 6
        ss2.Text = cfCalWgt
    
        ss2.Col = 7
        ss2.Text = Format(Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD') FROM DUAL"), "YYYY-MM-DD")
    
        ss2.Col = 8
        ss2.Text = Format(Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'HH24:MI') FROM DUAL"), "HH:MM")
    
        
        ss2.Col = 9
        ss2.Text = cLoc
        
        ss2.Col = 10
        ss2.Text = sUserID
        
        ss2.Col = 11
        ss2.Text = TXT_SLABNO
        
        ss2.Col = 12
        If i = cbo_cutcnt Then
            ss2.Text = "Y"
        Else
            ss2.Text = ""
        End If
        
        ss2.Col = 0
        ss2.Row = i
        ss2.Text = "Input"
    
    Next i
    SCRAP_NO = TXT_SLABNO
    
    Call WGT_CAL
    
    MDIMain.MenuTool.Buttons(1).Enabled = True                 'Save
    MDIMain.MenuTool.Buttons(2).Enabled = True                 'Delete
    MDIMain.MenuTool.Buttons(4).Enabled = True                 'Separator
    MDIMain.MenuTool.Buttons(14).Enabled = True                  'Row Delete
    
    MDIMain.MenuTool.Buttons(5).Enabled = False                 'Save
    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Row Insert
    MDIMain.MenuTool.Buttons(9).Enabled = False                 'Separator
    MDIMain.MenuTool.Buttons(11).Enabled = False                 'Row Insert
    MDIMain.MenuTool.Buttons(12).Enabled = False                 'Row Delete
    MDIMain.MenuTool.Buttons(15).Enabled = False                 'Row Delete
    
End Sub

Private Sub Cmd_Cancel_Click()
    Dim OutParam(2, 4) As Variant
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
    
    On Error Resume Next

    Screen.MousePointer = vbHourglass

        
    'Return loaction1 Parameter
    OutParam(1, 1) = "arg_loaction1"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 10

    'Return loaction2 Parameter
    OutParam(2, 1) = "arg_loaction2"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 10
    
    sQuery = "{call CGA2088C.P_ORDCANCEL('" & Trim(TXT_SLABNO.Text) & "','" & sUserID & "',?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If Trim(adoCmd("arg_loaction2")) <> "" Then
        Call Gp_MsgBoxDisplay("实绩处理失败，请确认=> " & adoCmd("arg_loaction2"))
    End If
    
    Set adoCmd = Nothing
    
    Call Form_Ref
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
   


    MDIMain.MenuTool.Buttons(14).Enabled = True                 'Save
'    MDIMain.MenuTool.Buttons(2).Enabled = True                 'Delete
'    MDIMain.MenuTool.Buttons(4).Enabled = True                 'Separator
'    MDIMain.MenuTool.Buttons(14).Enabled = True                  'Row Delete
'
'    MDIMain.MenuTool.Buttons(5).Enabled = False                 'Save
'    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Delete
'    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Row Insert
'    MDIMain.MenuTool.Buttons(9).Enabled = False                 'Separator
'    MDIMain.MenuTool.Buttons(11).Enabled = False                 'Row Insert
'    MDIMain.MenuTool.Buttons(12).Enabled = False                 'Row Delete
'    MDIMain.MenuTool.Buttons(15).Enabled = False                 'Row Delete

   
End Sub

Private Sub Form_Load()

    Dim sQuery As String
    
    Dim i, j As Integer
    
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Screen.MousePointer = vbDefault
    
    cbo_cutcnt.AddItem "0"
    cbo_cutcnt.AddItem "1"
    cbo_cutcnt.AddItem "2"
    cbo_cutcnt.AddItem "3"
    cbo_cutcnt.AddItem "4"
    cbo_cutcnt.AddItem "5"
    cbo_cutcnt.AddItem "6"
    cbo_cutcnt.AddItem "7"
    cbo_cutcnt.AddItem "8"
    cbo_cutcnt.AddItem "9"
    cbo_cutcnt.AddItem "10"
    
    Call opt_prc_status1_click(True)
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"))

    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)

    Call Gp_Spl_SizeGet(SSSplitter1, "CG-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "CG-System.INI", Me.Name)
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    txt_plt = "B1"
    Call txt_plt_KeyUp(0, 0)
    
    txt_thk = 150
    txt_thk_to = 320
    txt_wid = 1000
    txt_wid_to = 4000
    txt_len = 1000
    txt_len_to = 99999
    
    txt_cur_inv.Text = "XK"
    Call txt_cur_inv_KeyUp(0, 0)
    
    U_FROM_DATE.RawData = Format(Now, "YYYYMM") + "01"
    U_TO_DATE.RawData = Format(Now, "YYYYMMDD")
    
    If Mid(sAuthority, 1, 3) = "111" Then
       cmd_Cancel.Enabled = True
    Else
       cmd_Cancel.Enabled = False
    End If
    
'    MDIMain.MenuTool.Buttons(1).Enabled = True                 'Save
'    MDIMain.MenuTool.Buttons(2).Enabled = True                 'Delete
'    MDIMain.MenuTool.Buttons(4).Enabled = True                 'Separator
'    MDIMain.MenuTool.Buttons(8).Enabled = True                 'Row Insert
'    MDIMain.MenuTool.Buttons(14).Enabled = True                  'Row Delete
'
'    MDIMain.MenuTool.Buttons(5).Enabled = False                 'Save
'    MDIMain.MenuTool.Buttons(7).Enabled = True                 'Delete
'    MDIMain.MenuTool.Buttons(9).Enabled = True                 'Separator
'    MDIMain.MenuTool.Buttons(11).Enabled = False                 'Row Insert
'    MDIMain.MenuTool.Buttons(12).Enabled = False                 'Row Delete
'    MDIMain.MenuTool.Buttons(15).Enabled = False                 'Row Delete

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call Gp_Spl_SizeSet(SSSplitter1, "CG-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "CG-System.INI", Me.Name)
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
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
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Form_Cls()

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    txt_act_stlgrd_dec = ""
    TXT_SLABNO.Text = ""
    cbo_cutcnt.ListIndex = 0
    txt_total_len.Value = 0
    txt_total_wgt.Value = 0
    txt_scrap_wgt.Value = 0
    
    opt_prc_status1.Value = True
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&

End Sub

Public Sub Form_Pro()

    Dim iRow As Integer
    Dim sErrMessg As String

'    If txt_total_wgt.Value > txt_tmCalMo.Value Then
'       MsgBox "请确认算重量超过本来重量...!"
'       Exit Sub
'    End If
           
    If opt_prc_status1.Value = True Then
       Call Gf_Sp_Process(M_CN1, Proc_Sc("Sc2"), Mc1, True)
       Call Form_Ref
    End If
    If opt_prc_status2.Value = True Then
       M_CN1.BeginTrans
        
       For iRow = 1 To ss2.MaxRows
           ss2.Col = 0
           ss2.Row = iRow
           If ss2.Text = "Update" Or ss2.Text = "Delete" Then
               Call Sp_Process(sc2, iRow, sErrMessg)
           
               If Trim(sErrMessg) <> "" Then
                   Call Gp_MsgBoxDisplay(sErrMessg)
                   M_CN1.RollbackTrans
                   Screen.MousePointer = vbDefault
                   Exit Sub
               End If
           End If
       Next iRow
        
       M_CN1.CommitTrans
       
       Call Pro_ACB3020P
       
       If ss2.MaxRows > 1 Then
          Call ss1_Click(2, ss1.ActiveRow)
       Else
          ss2.MaxRows = 0
          Call Form_Ref
       End If

    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    
    MDIMain.MenuTool.Buttons(1).Enabled = True                   'Save
    MDIMain.MenuTool.Buttons(2).Enabled = True                   'Delete
    MDIMain.MenuTool.Buttons(4).Enabled = True                   'Separator
    MDIMain.MenuTool.Buttons(8).Enabled = True                   'Row Insert
    MDIMain.MenuTool.Buttons(14).Enabled = True                  'Row Delete
    
    MDIMain.MenuTool.Buttons(5).Enabled = False                  'Save
    MDIMain.MenuTool.Buttons(7).Enabled = True                   'Delete
    MDIMain.MenuTool.Buttons(9).Enabled = True                   'Separator
    MDIMain.MenuTool.Buttons(11).Enabled = False                 'Row Insert
    MDIMain.MenuTool.Buttons(12).Enabled = False                 'Row Delete
    MDIMain.MenuTool.Buttons(15).Enabled = False                 'Row Delete

End Sub

Public Sub Form_Ref()

    Dim ForCnt As Integer
    Dim tmWgt As Long
    Dim tmLen As Long

    If Not Gf_Sp_Cls(sc2) Then Exit Sub
    
    If Len(Trim(txt_MOSLAB)) <> 0 Then
        If Len(Trim(txt_MOSLAB)) < 8 Then
           MsgBox "请确认炉号", vbCritical, "系统提示信息"
           txt_MOSLAB.SetFocus
           Exit Sub
        End If
    End If
    
    If Len(Trim(txt_MOSLAB)) <> 8 Then
        If Len(Trim(txt_plt)) = 0 Then
            MsgBox "请确认工厂代码", vbCritical, "系统提示信息"
            txt_plt.SetFocus
            Exit Sub
        End If
    End If
    
    If txt_len.Value <= 0 Then
        MsgBox "请确认长度", vbCritical, "系统提示信息"
        txt_len.SetFocus
        Exit Sub
    End If
    
    If Trim(txt_cur_inv.Text) = "00" Or Trim(txt_cur_inv.Text) = "A2" Then
       MsgBox "请确认堆放仓库", vbCritical, "系统提示信息"
       txt_cur_inv.SetFocus
       Exit Sub
    End If
    
    TXT_SLABNO.Text = ""
    cbo_cutcnt.ListIndex = 0
    txt_total_len.Value = 0
    txt_total_wgt.Value = 0
    txt_scrap_wgt.Value = 0
    
    Call Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"))
    ss1.OperationMode = OperationModeNormal
    
End Sub

Private Sub opt_prc_status1_click(Value As Integer)

    If opt_prc_status1.Value = True Then
       cmd_Cancel.Visible = True
    End If
    
    If opt_prc_status1.Tag <> "" Then
       opt_prc_status1.Tag = ""
       Exit Sub
    End If
    
    If Gf_Sp_Cls(sc2) = False Then
        opt_prc_status2.Tag = "A"
        opt_prc_status2.Value = True
        txt_Status.Text = "2"
        Exit Sub
    End If
    
    'opt_prc_status1.Value = True
    
    opt_prc_status1.ForeColor = &HFF&
    opt_prc_status2.ForeColor = &H80000011
    
    Call Gf_Sp_Cls(sc1)
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    txt_Status.Text = "1"
    
    txt_act_stlgrd_dec = ""
    txt_MOSLAB.Text = ""
    TXT_SLABNO.Text = ""
    
    'cbo_cutcnt.Enabled = True
    
    cbo_cutcnt.Clear
    cbo_cutcnt.AddItem "0"
    cbo_cutcnt.AddItem "1"
    cbo_cutcnt.AddItem "2"
    cbo_cutcnt.AddItem "3"
    cbo_cutcnt.AddItem "4"
    cbo_cutcnt.AddItem "5"
    cbo_cutcnt.AddItem "6"
    cbo_cutcnt.AddItem "7"
    cbo_cutcnt.AddItem "8"
    cbo_cutcnt.AddItem "9"
    cbo_cutcnt.AddItem "10"
    
    cbo_cutcnt.ListIndex = 0
    txt_total_len.Value = 0
    txt_total_wgt.Value = 0
    txt_scrap_wgt.Value = 0
    
    txt_plt = "B1"
    Call txt_plt_KeyUp(0, 0)
    
    txt_thk = 150
    txt_thk_to = 320
    txt_wid = 1000
    txt_wid_to = 4000
    txt_len = 1000
    txt_len_to = 99999
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False                 'Separator

End Sub

Private Sub opt_prc_status2_Click(Value As Integer)

    If opt_prc_status2.Value = True Then
        cmd_Cancel.Visible = False
    End If
    
    If opt_prc_status2.Tag <> "" Then
       opt_prc_status2.Tag = ""
       Exit Sub
    End If
    
    If Gf_Sp_Cls(sc2) = False Then
        opt_prc_status1.Tag = "A"
        opt_prc_status1.Value = True
        txt_Status.Text = "1"
        Exit Sub
    End If
    
    'opt_prc_status2.Value = True
    
    opt_prc_status2.ForeColor = &HFF&
    opt_prc_status1.ForeColor = &H80000011
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    txt_Status.Text = "2"
    
    txt_act_stlgrd_dec = ""
    txt_MOSLAB.Text = ""
    TXT_SLABNO.Text = ""
    
    cbo_cutcnt.Enabled = False
    cbo_cutcnt.ListIndex = 0
    txt_total_len.Value = 0
    txt_total_wgt.Value = 0
    txt_scrap_wgt.Value = 0
    
    txt_plt = "B1"
    Call txt_plt_KeyUp(0, 0)
    
    txt_thk = 150
    txt_thk_to = 320
    txt_wid = 1000
    txt_wid_to = 4000
    txt_len = 1000
    txt_len_to = 99999
    
    U_FROM_DATE.RawData = Format(Now, "YYYYMM") + "01"
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    MDIMain.MenuTool.Buttons(7).Enabled = True                 'Delete
    MDIMain.MenuTool.Buttons(8).Enabled = True                 'Delete
    MDIMain.MenuTool.Buttons(9).Enabled = True                 'Separator
    
    'Call Form_Cls
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    'Dim cSlabLen As Long
    
    Dim iRow1, iRow2, iCol   As Integer
    Dim sColor, sHeat, sTemp As String
    Dim sChgPrcLine          As String
    Dim sL2SendFL            As String
    Dim i                    As Integer
    Dim ForCnt               As Integer
    Dim tmLen                As Double
    Dim tmWgt                As Double
        
    Dim tmThk As Double
    Dim tmWid As Double
    
    Dim tempWgt As Double
    Dim tot_cal_total As Double
    Dim cal_wgt As Double
    Dim tmp_rat As Double
    Dim tmTotalLen As Double
    Dim tmpLen As Double
    Dim sub_wgt As Double
    Dim sub_len As Double
    Dim tmCalCut As Double
    Dim tmCalMo As Double
    Dim tmCalCutOne As Double

    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, "&H00000000", "&HFFFF80")
      
    For i = 1 To ss1.MaxRows
        If i <> Row Then
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, i, i)
        End If
    Next
    
    If Row <> 0 Then

        ss1.Col = 1
        ss1.Row = Row
        TXT_SLABNO = ss1.Text
        SCRAP_NO = ss1.Text
        cSlabno = ss1.Text
        ss1.Col = 7
        cSlabLen = ss1.Value
        ss1.Col = 8
        cSlabWgt = ss1.Value
        ss1.Col = 16
        txt_tmpPLT = ss1.Value
        ss1.Col = 17
        txt_IST_DATE = ss1.Value
    End If

    sQuery = "          SELECT MAX(SLAB_NO) "
    sQuery = sQuery & "   FROM NISCO.FP_SLAB "
    sQuery = sQuery & "  WHERE SLAB_NO LIKE '" & Mid(cSlabno, 1, 8) & "%'"
    
    tmpSlabNo = Gf_CodeFind(M_CN1, sQuery)
    If CInt(Mid(tmpSlabNo, 9, 2)) < 30 Or CInt(Mid(tmpSlabNo, 9, 2)) >= 97 Then  'modified by guoli at 20080418
       tmpSlabNo = Mid(tmpSlabNo, 1, 8) & "30"
    End If
    
    ss1.Row = Row
    ss1.Col = 1

    lBlkrow1 = Row
    lBlkrow2 = Row
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").Row = 0
    sc2.Item("Spread").Text = ""

    If Row = 0 Then Exit Sub

    If Row = 0 Then Call Gp_Sp_Sort(sc1.Item("Spread"), Col, Row)
    
    If opt_prc_status2 Then
        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2, Nothing, Mc2("mControl"), False)
        For i = 1 To ss2.MaxRows
             ss2.Row = i
             ss2.Col = 15
             If Trim(ss2.Text) = "订单材" Then
                Call Gp_Sp_BlockLock(ss2, 3, 4, i, i)
             End If
        Next i
        Exit Sub
    End If
    

    Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2, Nothing, Mc2("mControl"), False)


    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 1
        
        NEWSLABNO = Mid(tmpSlabNo, 1, 8) & Mid(tmpSlabNo, 9, 2) + i
        If Len(Mid(NEWSLABNO, 5, 6)) = 5 Then
           NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "0" & Mid(NEWSLABNO, 5, 5)
        ElseIf Len(Mid(NEWSLABNO, 5, 6)) = 4 Then
           NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "00" & Mid(NEWSLABNO, 5, 5)
        ElseIf Len(Mid(NEWSLABNO, 5, 6)) = 3 Then
           NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "000" & Mid(NEWSLABNO, 5, 5)
        End If
        
        ss2.Text = NEWSLABNO
        
        ss2.Col = 2
        tmThk = ss2.Value
        
        ss2.Col = 3
        tmWid = ss2.Value
        
        ss2.Col = 4
        tmLen = ss2.Value
            
        ss2.Col = 6
        ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
        
        ss2.Col = 11
        ss2.Text = sUserID
        
        ss2.Col = 12
        ss2.Text = TXT_SLABNO
        
        ss2.Col = 13
        If i = ss2.MaxRows Then
            ss2.Text = "Y"
        Else
            ss2.Text = ""
        End If
    Next

    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 0
        If ss2.Text <> "Delete" Then
            ss2.Row = i
            ss2.Col = 4
            tmTotalLen = tmTotalLen + ss2.Value
            
            ss2.Col = 5
            tempWgt = tempWgt + ss2.Value
        End If

    Next i
    
    If txt_Status = "1" Then
        For i = 1 To ss2.MaxRows
             ss2.Row = i
             ss2.Col = 0
             If UCase(ss2.Text) = "" Then
                ss2.Text = "Input"
             End If
             ss2.Col = 15
             If Trim(ss2.Text) = "订单材" Then
                Call Gp_Sp_BlockLock(ss2, 3, 4, i, i)
             End If
             
        Next i
    End If
    
    If tmTotalLen = cSlabLen Then
       txt_total_len.ForeColor = &H0&
    Else
       txt_total_len.ForeColor = &HFF&
    End If
    txt_total_len = tmTotalLen
    
    txt_total_wgt = tempWgt
    If CDbl(txt_total_wgt) - cSlabWgt = 0 Then
       txt_total_wgt.ForeColor = &H0&
    Else
       txt_total_wgt.ForeColor = &HFF&
    End If
    
    txt_scrap_wgt = Format(cSlabWgt - tempWgt, "###0.000")
    If cSlabWgt - tempWgt = 0 Then
       txt_scrap_wgt.ForeColor = &H0&
    Else
       txt_scrap_wgt.ForeColor = &HFF&
    End If
    
End Sub

Public Sub Form_Exc()
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If
    
End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim tmThk As Double
    Dim tmWid As Double
    Dim tmLen As Double
    Dim tempWgt As Double
    Dim i As Integer

    If Gf_Sc_Authority(sAuthority, "U") Then Call Gp_Sp_UpdateMake(Proc_Sc("SC2")("Spread"), Mode)
    
    
    If Col <> 2 And Col <> 3 And Col <> 4 Then Exit Sub
    
    If ChangeMade Then
        Call WGT_CAL
    End If
    
End Sub

Private Sub txt_act_stlgrd_Change()

    If Len(Trim(txt_act_stlgrd.Text)) = 0 Then txt_act_stlgrd_dec.Text = ""
    
End Sub

Private Sub txt_act_stlgrd_DblClick()

    Call txt_act_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_act_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        'txt_act_stlgrd.Text = ""
        DD.rControl.Add Item:=txt_act_stlgrd
        DD.rControl.Add Item:=txt_act_stlgrd_dec

        Call Gf_Stlgrd_DD(M_CN1, vbKeyF4)

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

    Else

        If Len(Trim(txt_plt)) = txt_plt.MaxLength Then
            txt_plt_dec.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
        Else
            txt_plt_dec.Text = ""
        End If
    
    End If

End Sub

Private Sub txt_cur_inv_Change()

    If Len(Trim(txt_cur_inv.Text)) = txt_cur_inv.MaxLength Then
        txt_cur_name.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv.Text, 2)
    Else
        txt_cur_name.Text = ""
    End If
    
End Sub

Private Sub txt_cur_inv_DblClick()

    Call txt_cur_inv_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_cur_inv_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=txt_cur_inv
        DD.rControl.Add Item:=txt_cur_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
     
        If Len(Trim(txt_cur_inv.Text)) = txt_cur_inv.MaxLength Then
            txt_cur_name.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv.Text, 2)
        Else
            txt_cur_name.Text = ""
        End If
        
    End If
    
End Sub

Private Sub Sp_Process(Sc As Collection, iRow As Integer, sErrMessg As String)

    Dim iCol     As Integer
    Dim sTemp       As String
    Dim dTempInt    As Double

    Dim adoCmd As ADODB.Command
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command

    Set adoCmd.ActiveConnection = M_CN1

    With Sc.Item("Spread")
        .Row = iRow
        .Col = 0
        If .Text = "Input" Or .Text = "Update" Or ss2.Text = "Delete" Then
        
            adoCmd.CommandType = adCmdStoredProc
            adoCmd.CommandText = Sc.Item("P-M")
            
            'Create Parameter (Input) iType + iColumn
            For iCol = 0 To Sc.Item("iColumn").Count
                adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
            Next iCol
            
            If .Text = "Input" Then
               adoCmd.Parameters(0).Value = "I"
            ElseIf .Text = "Update" Then
               adoCmd.Parameters(0).Value = "U"
            ElseIf .Text = "Delete" Then
               adoCmd.Parameters(0).Value = "D"
            End If
            
            'Ceate Parameter (Output)
            adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
            adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
               
            'Parameters Setting
            For iCol = 1 To Sc.Item("iColumn").Count

                Sc.Item("Spread").Col = Sc.Item("iColumn").Item(iCol)
                Select Case Sc.Item("Spread").CellType

                       Case SS_CELL_TYPE_NUMBER
                            If Trim(Sc.Item("Spread").Text) = "" Then
                                adoCmd.Parameters(iCol).Value = 0
                            Else
                                dTempInt = Sc.Item("Spread").Text
                                adoCmd.Parameters(iCol).Value = Trim(Str(dTempInt))
                            End If
                
                       Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                            If Trim(Sc.Item("Spread").Value) = "" Then
                                adoCmd.Parameters(iCol).Value = ""
                            Else
                                adoCmd.Parameters(iCol).Value = Trim(Str(Sc.Item("Spread").Value))
                            End If
                            
                       Case SS_CELL_TYPE_DATE
                            If Trim(Sc.Item("Spread").Text) = "" Then
                                adoCmd.Parameters(iCol).Value = ""
                            Else
                                adoCmd.Parameters(iCol).Value = Mid(Trim(Sc.Item("Spread").Text), 1, 4) & _
                                                                Mid(Trim(Sc.Item("Spread").Text), 6, 2) & _
                                                                Mid(Trim(Sc.Item("Spread").Text), 9, 2)
                            End If
                        
                       Case Else
                            sTemp = Replace(Sc.Item("Spread").Text, "'", "''")
                            adoCmd.Parameters(iCol).Value = Trim(sTemp)

                End Select
            Next iCol
            
            adoCmd.Execute

            'Error Check
            If adoCmd("Error") <> "0" Then
               sErrMessg = adoCmd("Error") & ":" & adoCmd("Messg")
            End If
            
        End If
        
    End With
    
End Sub

Private Sub Pro_ACB3020P()

    Dim adoCmd As ADODB.Command
    Dim sQuery As String
    
    Set adoCmd = Nothing
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandType = adCmdText
    
    'Ceate Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter(Str(1), adVariant, adParamOutput)
    
    sQuery = "{call ACB3020P (?)}"
    
    adoCmd.CommandText = sQuery
    adoCmd.Execute , , adExecuteNoRecords
        
    Set adoCmd = Nothing
End Sub

