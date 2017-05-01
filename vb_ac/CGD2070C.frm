VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGD2070C 
   Caption         =   "录入精整作业指示_CGD2070C"
   ClientHeight    =   9225
   ClientLeft      =   315
   ClientTop       =   2610
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1_PLATE_NO 
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
      Left            =   240
      MaxLength       =   14
      TabIndex        =   40
      Top             =   480
      Visible         =   0   'False
      Width           =   1875
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9165
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   16166
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "CGD2070C.frx":0000
      Begin Threed.SSFrame SSFrame2 
         Height          =   1335
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "查询区域"
         Top             =   0
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   2355
         _Version        =   196609
         BackColor       =   14737632
         ShadowStyle     =   1
         Begin VB.ComboBox cbo_prod_cd 
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
            ItemData        =   "CGD2070C.frx":0072
            Left            =   13920
            List            =   "CGD2070C.frx":007F
            TabIndex        =   39
            Tag             =   "产品"
            Top             =   900
            Width           =   1155
         End
         Begin VB.TextBox TXT_LOC 
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
            MaxLength       =   10
            TabIndex        =   29
            Tag             =   "标准号"
            Top             =   900
            Width           =   1245
         End
         Begin VB.TextBox txt_lot_no 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   3405
            MaxLength       =   14
            TabIndex        =   11
            Tag             =   "轧批号"
            Top             =   900
            Width           =   1875
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
            ItemData        =   "CGD2070C.frx":008F
            Left            =   11265
            List            =   "CGD2070C.frx":0091
            TabIndex        =   10
            Tag             =   "班次"
            Top             =   120
            Width           =   735
         End
         Begin VB.TextBox txt_stdspec_chg 
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
            Left            =   6825
            MaxLength       =   18
            TabIndex        =   9
            Tag             =   "标准号"
            Top             =   510
            Width           =   2895
         End
         Begin VB.TextBox txt_cur_inv 
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
            Left            =   4800
            TabIndex        =   7
            Top             =   810
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.TextBox txt_plt 
            Alignment       =   2  'Center
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   3405
            MaxLength       =   2
            TabIndex        =   6
            Tag             =   "生产工厂"
            Top             =   120
            Width           =   570
         End
         Begin VB.TextBox txt_plt_name 
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   3990
            TabIndex        =   5
            Tag             =   "机号"
            Top             =   120
            Width           =   1290
         End
         Begin VB.TextBox txt_PrcLine 
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
            Left            =   60
            MaxLength       =   11
            TabIndex        =   4
            Top             =   60
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.TextBox TXT_PLATE_NO 
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
            Left            =   3405
            MaxLength       =   14
            TabIndex        =   3
            Top             =   510
            Width           =   1875
         End
         Begin Threed.SSOption opt_LineFlag 
            Height          =   405
            Index           =   2
            Left            =   450
            TabIndex        =   12
            Top             =   750
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   714
            _Version        =   196609
            Font3D          =   1
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "指示取消"
         End
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   2190
            Top             =   510
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "查询号"
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
            Left            =   10050
            Top             =   510
            Width           =   1185
            _ExtentX        =   2090
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
         End
         Begin CSTextLibCtl.sidbEdit SDB_THK 
            Height          =   315
            Left            =   11265
            TabIndex        =   13
            Top             =   510
            Width           =   1005
            _Version        =   262145
            _ExtentX        =   1773
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
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   10050
            Top             =   900
            Width           =   1185
            _ExtentX        =   2090
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
         End
         Begin CSTextLibCtl.sidbEdit SDB_WID 
            Height          =   315
            Left            =   11265
            TabIndex        =   14
            Top             =   900
            Width           =   1005
            _Version        =   262145
            _ExtentX        =   1773
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
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSOption opt_LineFlag 
            Height          =   405
            Index           =   0
            Left            =   450
            TabIndex        =   15
            Top             =   210
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   714
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   255
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "精整指示"
            Value           =   -1
         End
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   5610
            Top             =   900
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "当前库"
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
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Index           =   0
            Left            =   2190
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
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
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   5610
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "生产时间"
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
         Begin InDate.UDate SDT_PROD_DATE_FROM 
            Height          =   315
            Left            =   6825
            TabIndex        =   16
            Tag             =   "起始日期"
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
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
            Left            =   8280
            TabIndex        =   17
            Tag             =   "起始日期"
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
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
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   1
            Left            =   5610
            Top             =   510
            Width           =   1185
            _ExtentX        =   2090
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
            ForeColor       =   16711680
         End
         Begin InDate.ULabel ULabel19 
            Height          =   315
            Left            =   2190
            Top             =   900
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "轧批号"
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
         Begin VB.TextBox txt_cur_inv_code 
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
            Left            =   6825
            MaxLength       =   2
            TabIndex        =   8
            Tag             =   "当前库"
            Top             =   900
            Width           =   660
         End
         Begin CSTextLibCtl.sidbEdit SDB_THK_TO 
            Height          =   315
            Left            =   12540
            TabIndex        =   33
            Top             =   510
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_WID_TO 
            Height          =   315
            Left            =   12540
            TabIndex        =   35
            Top             =   900
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   7560
            Top             =   900
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            Caption         =   "垛位"
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
            Index           =   0
            Left            =   10050
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
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
         Begin Threed.SSPanel SSPpdt 
            Height          =   315
            Left            =   12540
            TabIndex        =   38
            Top             =   120
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   255
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "当月以前交货订单"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   13920
            Top             =   510
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            Caption         =   "产品"
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
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   12360
            TabIndex        =   36
            Top             =   1020
            Width           =   195
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   12360
            TabIndex        =   34
            Top             =   630
            Width           =   195
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1215
         Left            =   0
         TabIndex        =   18
         ToolTipText     =   "精整作业指示区域"
         Top             =   1365
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   2143
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737918
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowStyle     =   1
         Begin VB.TextBox TXT_SPEC_PROC 
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
            Left            =   1560
            TabIndex        =   37
            Text            =   " "
            Top             =   780
            Width           =   1005
         End
         Begin VB.TextBox txt_REMARKS 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   9750
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   120
            Width           =   5085
         End
         Begin VB.TextBox txt_HTM_METH3 
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
            Left            =   6870
            TabIndex        =   26
            Top             =   450
            Width           =   585
         End
         Begin VB.TextBox txt_HTM_COND3 
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
            Left            =   7470
            TabIndex        =   25
            Top             =   450
            Width           =   795
         End
         Begin VB.TextBox txt_HTM_METH2 
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
            Left            =   5460
            TabIndex        =   24
            Top             =   450
            Width           =   585
         End
         Begin VB.TextBox txt_HTM_COND2 
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
            Left            =   6060
            TabIndex        =   23
            Top             =   450
            Width           =   795
         End
         Begin VB.TextBox txt_HTM_METH1 
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
            Left            =   4050
            TabIndex        =   22
            Top             =   450
            Width           =   585
         End
         Begin VB.TextBox txt_HTM_COND1 
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
            Left            =   4650
            TabIndex        =   21
            Top             =   450
            Width           =   795
         End
         Begin VB.TextBox txt_SB 
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
            Left            =   1560
            TabIndex        =   20
            Text            =   " "
            Top             =   450
            Width           =   1005
         End
         Begin VB.TextBox txt_UST 
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
            Left            =   1560
            TabIndex        =   19
            Text            =   " "
            Top             =   120
            Width           =   1005
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   0
            Left            =   330
            Top             =   120
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "探伤方法"
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
            Index           =   2
            Left            =   330
            Top             =   450
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "抛丸"
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
            Height          =   645
            Index           =   3
            Left            =   2820
            Top             =   120
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1138
            Caption         =   "热处理指示"
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
            ForeColor       =   64
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   4050
            Top             =   120
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            Caption         =   "方法/条件一"
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
            Left            =   5460
            Top             =   120
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            Caption         =   "方法/条件二"
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
            Left            =   6870
            Top             =   120
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            Caption         =   "方法/条件三"
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
            Height          =   975
            Index           =   5
            Left            =   8520
            Top             =   120
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1720
            Caption         =   "备注"
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
            ForeColor       =   64
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   4
            Left            =   2820
            Top             =   780
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "热处理取消"
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
            ForeColor       =   64
         End
         Begin Threed.SSCheck chk_can 
            Height          =   270
            Index           =   0
            Left            =   4050
            TabIndex        =   30
            Top             =   810
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   476
            _Version        =   196609
            Font3D          =   2
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
            Caption         =   " 取消一"
         End
         Begin Threed.SSCheck chk_can 
            Height          =   270
            Index           =   1
            Left            =   5460
            TabIndex        =   31
            Top             =   810
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   476
            _Version        =   196609
            Font3D          =   2
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
            Caption         =   " 取消二"
         End
         Begin Threed.SSCheck chk_can 
            Height          =   270
            Index           =   2
            Left            =   6870
            TabIndex        =   32
            Top             =   810
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   476
            _Version        =   196609
            Font3D          =   2
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
            Caption         =   " 取消三"
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   6
            Left            =   330
            Top             =   780
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "特殊工序"
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
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   6555
         Left            =   0
         TabIndex        =   28
         Top             =   2610
         Width           =   15285
         _Version        =   393216
         _ExtentX        =   26961
         _ExtentY        =   11562
         _StockProps     =   64
         ColsFrozen      =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   52
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGD2070C.frx":0093
      End
   End
   Begin Threed.SSOption opt_LineFlag 
      Height          =   405
      Index           =   1
      Left            =   510
      TabIndex        =   0
      Top             =   11880
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "精整保留"
   End
End
Attribute VB_Name = "CGD2070C"
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
'-- Program Name      钢板实绩查询界面
'-- Program ID        AGC2200C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
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
Public sDateTime As String          'Active Form Time Setting
Public sQuery_load As String        'Active Form sQuery Setting

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

Dim pColumn  As New Collection      'Spread Primary Key Collection
Dim nColumn  As New Collection      'Spread necessary Column Collection
Dim mColumn  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn  As New Collection      'Spread Insert Column Collection
Dim aColumn  As New Collection      'Master -> Spread Column Collection
Dim lColumn  As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_CONF_TIME = 1
Const SS1_LINE = 2
Const SS1_PLATE_NO = 3
Const SS1_PROD_CD = 4
Const SS1_PROC_CD = 5
Const SS1_SPEC_FL = 6
Const SS1_SPEC_NAME = 7
Const SS1_DEL_DATE_TO = 9
Const SS1_GAS_FL = 13
Const SS1_GRID_FL = 16       '15 -> 16
Const SS1_CL_FL = 18         '17 -> 18
Const SS1_UST_FL = 20        '19 -> 20
Const SS1_UST_M = 21         '20 -> 21
Const SS1_SB_FL = 25         '22 -> 23
Const SS1_SB_M = 26          '23 -> 24
Const SS1_HTM_FL = 28        '25 -> 26
Const SS1_HTM_M1 = 31        '28 -> 29
Const SS1_HTM_C1 = 32        '29 -> 30
Const SS1_HTM_M2 = 33        '30 -> 31
Const SS1_HTM_C2 = 34        '31 -> 32
Const SS1_HTM_M3 = 35        '32 -> 33
Const SS1_HTM_C3 = 36        '33 -> 34
Const SS1_REMARK = 38        '35 -> 36
Const SS1_USERID = 46        '43 -> 44
Const SPD_URGNT_FL = 52

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(txt_PrcLine, "p", " ", " ", " ", " ", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(cbo_shift, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_cur_inv_code, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_stdspec_chg, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(SDB_THK, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(SDB_WID, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_lot_no, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(TXT_LOC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDB_THK_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDB_WID_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(cbo_prod_cd, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
           
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"

     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", "a", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGD2070C.P_REFER", Key:="P-R"
    sc1.Add Item:="CGD2070C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="CGD2070C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
'          Call Gp_Ms_Collection(txt_UST, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'           Call Gp_Ms_Collection(txt_SB, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'    Call Gp_Ms_Collection(txt_HTM_METH1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'    Call Gp_Ms_Collection(txt_HTM_COND1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'    Call Gp_Ms_Collection(txt_HTM_METH2, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'    Call Gp_Ms_Collection(txt_HTM_COND2, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'    Call Gp_Ms_Collection(txt_HTM_METH3, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'    Call Gp_Ms_Collection(txt_HTM_COND3, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'    Call Gp_Ms_Collection(txt_HTM_COND3, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'
    Call Gp_Ms_Collection(Text1_PLATE_NO, "P", " ", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_REMARKS, "", " ", " ", "i", " r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
    
    Mc2.Add Item:="CGD2070C.P_MODIFY1", Key:="P-M"
    Mc2.Add Item:="CGD2070C.P_REFER1", Key:="P-R"
    Mc2.Add Item:=pControl1, Key:="pControl"
    Mc2.Add Item:=nControl1, Key:="nControl"
    Mc2.Add Item:=mControl1, Key:="mControl"
    Mc2.Add Item:=iControl1, Key:="iControl"
    Mc2.Add Item:=rControl1, Key:="rControl"
    Mc2.Add Item:=cControl1, Key:="cControl"
    Mc2.Add Item:=aControl1, Key:="aControl"
    Mc2.Add Item:=lControl1, Key:="lControl"

    Call Gp_Sp_ColHidden(ss1, SS1_CONF_TIME, True)
    Call Gp_Sp_ColHidden(ss1, SS1_LINE, True)
    Call Gp_Sp_ColHidden(ss1, SS1_USERID, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
     
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Private Sub chk_can_Click(Index As Integer, Value As Integer)
    If chk_can(Index).Value = -1 Then
       chk_can(Index).ForeColor = &HFF&
    Else
       chk_can(Index).ForeColor = &H808080
    End If
End Sub

Private Sub opt_LineFlag_Click(Index As Integer, Value As Integer)

    If opt_LineFlag(0).Value = True Then
       txt_PrcLine = "1"
       opt_LineFlag(0).ForeColor = &HFF&       'red
       opt_LineFlag(1).ForeColor = &H80000012  'black
       opt_LineFlag(2).ForeColor = &H80000012  'black
        'Call Form_Ref

    ElseIf opt_LineFlag(1).Value = True Then
       txt_PrcLine = "2"
       opt_LineFlag(0).ForeColor = &H80000012       'black
       opt_LineFlag(1).ForeColor = &HFF&  'red
       opt_LineFlag(2).ForeColor = &H80000012  'black
        'Call Form_Ref
    
    Else
       txt_PrcLine = "3"
       opt_LineFlag(0).ForeColor = &H80000012       'black
       opt_LineFlag(1).ForeColor = &H80000012       'black
       opt_LineFlag(2).ForeColor = &HFF&  'red
        'Call Form_Ref

    End If
    
End Sub



'Private Sub TXT_SPEC_PROC_Change()
'    If Len(Trim(TXT_SPEC_PROC.Text)) = TXT_SPEC_PROC.MaxLength Then
'        TXT_SPEC_PROC_NAME.Text = Gf_ComnNameFind(M_CN1, "G0046", TXT_SPEC_PROC.Text, 2)
'    Else
'        TXT_SPEC_PROC_NAME.Text = ""
'    End If
'End Sub

Private Sub TXT_SPEC_PROC_DblClick()
    Call TXT_SPEC_PROC_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_SPEC_PROC_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sSpec_proc As String
    sSpec_proc = TXT_SPEC_PROC.Text

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0046"

        DD.rControl.Add Item:=TXT_SPEC_PROC

        DD.nameType = "2"
        TXT_SPEC_PROC.Text = ""
        Call Gf_Common_DD(M_CN1, KeyCode)
        If TXT_SPEC_PROC.Text = "" Then
            TXT_SPEC_PROC.Text = sSpec_proc
        End If
        
    End If
End Sub

Private Sub txt_cur_inv_code_DblClick()
    Call txt_cur_inv_code_KeyUp(vbKeyF4, 0)
End Sub

Public Sub txt_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
    
        DD.rControl.Add Item:=txt_cur_inv_code
        DD.rControl.Add Item:=txt_cur_inv
    
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
       
        If Len(Trim(txt_cur_inv_code.Text)) = txt_cur_inv_code.MaxLength Then
            txt_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv_code.Text, 2)
        Else
            txt_cur_inv.Text = ""
        End If
    End If
    
End Sub

Private Sub txt_HTM_COND1_Change()

    If Len(Trim(txt_HTM_COND1.Text)) = 4 Then
       If Trim(txt_HTM_METH1) <> Mid(Trim(txt_HTM_COND1), 1, 1) Then
          MsgBox "热处理方法与热处理条件不一样"
          txt_HTM_COND1 = ""
       End If
    End If
    
End Sub

Private Sub txt_HTM_COND1_DblClick()
    Call txt_HTM_COND1_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_HTM_COND1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = txt_HTM_METH1.Text
        
        DD.rControl.Add Item:=txt_HTM_COND1
        'DD.rControl.Add Item:=TXT_HTM_MET1_nm

        DD.nameType = "2"

        Call Gf_HEAT_COND_DD(M_CN1, KeyCode)

    End If
    
End Sub

Private Sub txt_HTM_COND2_DblClick()
    Call txt_HTM_COND2_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_HTM_COND2_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = txt_HTM_METH2.Text
        DD.rControl.Add Item:=txt_HTM_COND2
        'DD.rControl.Add Item:=TXT_HTM_MET1_nm

        DD.nameType = "2"

        Call Gf_HEAT_COND_DD(M_CN1, KeyCode)

    End If
    
End Sub

Private Sub txt_HTM_COND3_DblClick()
    Call txt_HTM_COND3_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_HTM_COND3_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = txt_HTM_METH3.Text
        DD.rControl.Add Item:=txt_HTM_COND3
        'DD.rControl.Add Item:=TXT_HTM_MET1_nm

        DD.nameType = "2"

        Call Gf_HEAT_COND_DD(M_CN1, KeyCode)

    End If
    
End Sub

Private Sub txt_HTM_METH1_DblClick()
    Call txt_HTM_METH1_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_HTM_METH1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0073"
        DD.rControl.Add Item:=txt_HTM_METH1
        'DD.rControl.Add Item:=TXT_HTM_MET1_nm

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

    End If

End Sub

Private Sub txt_HTM_METH2_DblClick()
    Call txt_HTM_METH2_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_HTM_METH2_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0073"
        DD.rControl.Add Item:=txt_HTM_METH2
        'DD.rControl.Add Item:=TXT_HTM_MET1_nm

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

    End If

End Sub

Private Sub txt_HTM_METH3_DblClick()
    Call txt_HTM_METH3_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_HTM_METH3_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0073"
        DD.rControl.Add Item:=txt_HTM_METH3
        'DD.rControl.Add Item:=TXT_HTM_MET1_nm

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

    End If

End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuTool_ReSet

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        If Len(TXT_PLATE_NO.Text) >= 8 Then
           Call Form_Ref
        End If
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "AC-System.INI", Me.Name)
    
    If App.Title = "AC" Then
        txt_plt.Text = "C1"
        txt_cur_inv_code.Text = "00"
    ElseIf App.Title = "BG" Then
        txt_plt.Text = "C1"
        txt_cur_inv_code.Text = "00"
    ElseIf App.Title = "CG" Then
        txt_plt.Text = "C3"
        txt_cur_inv_code.Text = "ZB"
'    ElseIf App.Title = "DG" Then
'        txt_plt.Text = "C1"
'        txt_cur_inv_code.Text = "WD"
    ElseIf App.Title = "DE" Then
        txt_plt.Text = "C1"
        txt_cur_inv_code.Text = "00"
    End If
    
    Call txt_plt_KeyUp(0, 0)
    Call txt_cur_inv_code_KeyUp(0, 0)
    
    SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
    SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
    cbo_shift.AddItem ""
    cbo_shift.AddItem "1"
    cbo_shift.AddItem "2"
    cbo_shift.AddItem "3"
    
    opt_LineFlag(0).Value = True
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "AC-System.INI", Me.Name)
    
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
    
    Set iColumn = Nothing
    Set pColumn = Nothing
    Set lColumn = Nothing
    Set nColumn = Nothing
    Set mColumn = Nothing
    Set aColumn = Nothing

    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
       Call Gp_Ms_Cls(Mc1("rControl"))
       Call Gp_Ms_Cls(Mc2("rControl"))
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call MenuTool_ReSet
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If
    
    If App.Title = "AC" Then
        txt_plt.Text = "C1"
        txt_cur_inv_code.Text = "00"
    ElseIf App.Title = "BG" Then
        txt_plt.Text = "C1"
        txt_cur_inv_code.Text = "00"
    ElseIf App.Title = "CG" Then
        txt_plt.Text = "C3"
        txt_cur_inv_code.Text = "ZB"
    ElseIf App.Title = "DG" Then
        txt_plt.Text = "C1"
        txt_cur_inv_code.Text = "WD"
    ElseIf App.Title = "DE" Then
        txt_plt.Text = "C1"
        txt_cur_inv_code.Text = "00"
    End If
    
    Call txt_plt_KeyUp(0, 0)
    Call txt_cur_inv_code_KeyUp(0, 0)
    
    Text1_PLATE_NO = ""

End Sub

Public Sub Form_Ref()
    
    Dim SMESG As String
    
    Dim iRow As Integer
    Dim sCurDate As String
    Dim sDel_To_Date As String
    Dim sURGNT As String
    
    sCurDate = Format(Now, "YYYYMM")

    If opt_LineFlag(0).Value <> True And opt_LineFlag(2).Value <> True Then
        MsgBox "请选择精整等待或精整保留...!"
         Exit Sub
    End If
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), , False) Then
       
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
    End If
    
    '超交货期用红色显示 add by liqian 2012-07-23
    With ss1
        For iRow = 1 To .MaxRows
            .Row = iRow:             .Col = SS1_DEL_DATE_TO
             sDel_To_Date = Mid(.Value, 1, 6)
             If sDel_To_Date < sCurDate Then
                  Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, &HFF&)
             End If
             
              '是否紧急订单警示
            .Row = iRow:             .Col = SPD_URGNT_FL
             sURGNT = .Text
              If sURGNT = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, &H80FF80)
              End If
        Next iRow
        ss1.Row = 1
        ss1.Col = SS1_PLATE_NO:   Text1_PLATE_NO.Text = ss1.Text
        If Gf_Ms_Refer(M_CN1, Mc2, , , False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc2("pControl"), True)
        End If
    End With
    
    

End Sub

Public Sub Form_Pro()

    Dim iCount As Integer

    For iCount = 1 To ss1.MaxRows

        Select Case Trim(Gf_Sp_RcvData(ss1, 0, iCount))

            Case "Update"

            ss1.Col = SS1_USERID
            ss1.Text = sUserID

        End Select

    Next iCount
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
    End If
    
'
'    If Gf_Mc_Authority(sAuthority, Mc1) Then
'                       If Gf_Ms_Process(M_CN1, Mc2, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
''                       CBO_SLAB_NO.Enabled = True
'                    End If
    
    
          If Gf_Mc_Authority(sAuthority, Mc2) Then
        If Gf_Ms_Process(M_CN1, Mc2, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            End If
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    End If

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim ForCnt As Long
    Dim chkFL As String
    Dim chkOrd As String
    Dim ORGCOL As Integer

    If Row < 1 Then Exit Sub
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    ORGCOL = Col
    
    If txt_PrcLine = "1" Or txt_PrcLine = "3" Then
    
        ss1.Row = Row:      ss1.Col = SS1_USERID:     ss1.Text = sUserID

        
        If ORGCOL = SS1_GAS_FL Or ORGCOL = SS1_GRID_FL Or ORGCOL = SS1_CL_FL Then
            ss1.Row = Row
            ss1.Col = SS1_LINE:            ss1.Text = txt_PrcLine
        End If
        
        If ORGCOL = SS1_SPEC_FL Then
            If txt_PrcLine = "1" Then
                If Len(Trim(TXT_SPEC_PROC)) <> 1 Then
                   MsgBox "请确认特殊工序", vbOKOnly + vbQuestion, "提示"
                   Exit Sub
                End If
                ss1.Col = SS1_LINE:                ss1.Text = txt_PrcLine
                ss1.Col = SS1_SPEC_NAME:           ss1.Text = TXT_SPEC_PROC.Text
            End If
        End If
        
        If ORGCOL = SS1_UST_FL Then
            If txt_PrcLine = "1" Then
                If Len(Trim(txt_UST)) <> 4 Then
                   MsgBox "请确认探伤方法", vbOKOnly + vbQuestion, "提示"
                   Exit Sub
                End If
                ss1.Col = SS1_LINE:                ss1.Text = txt_PrcLine
                ss1.Col = SS1_UST_M:               ss1.Text = txt_UST
            End If
        End If
    
        If ORGCOL = SS1_SB_FL Then
            If txt_PrcLine = "1" Then
                If Trim(txt_SB) = "" Then
                   MsgBox "请确认抛丸方法", vbOKOnly + vbQuestion, "提示"
                   Exit Sub
                End If
                ss1.Col = SS1_LINE:                ss1.Text = txt_PrcLine
                ss1.Col = SS1_SB_M:                ss1.Text = txt_SB
            End If
        End If
    
        If ORGCOL = SS1_HTM_FL Then
        
            If txt_PrcLine = "1" Then
            
                If Trim(txt_HTM_METH1) = "" And txt_HTM_METH2 = "" And txt_HTM_METH3 = "" Then
                   MsgBox "请确认热处理方法", vbOKOnly + vbQuestion, "提示"
                   Exit Sub
                End If
                
                If Trim(txt_HTM_METH1) = "" And txt_HTM_METH2 <> "" Then
                   MsgBox "请确认热处理方法一", vbOKOnly + vbQuestion, "提示"
                   Exit Sub
                End If
                
                If Trim(txt_HTM_METH2) = "" And txt_HTM_METH3 <> "" Then
                   MsgBox "请确认热处理方法二", vbOKOnly + vbQuestion, "提示"
                   Exit Sub
                End If
                
                If Trim(txt_HTM_METH1) <> "" And txt_HTM_COND1 = "" Then
                   MsgBox "请确认热处理条件一", vbOKOnly + vbQuestion, "提示"
                   Exit Sub
                End If
                
                If Trim(txt_HTM_METH2) <> "" And txt_HTM_COND2 = "" Then
                   MsgBox "请确认热处理条件二", vbOKOnly + vbQuestion, "提示"
                   Exit Sub
                End If
                
                If Trim(txt_HTM_METH3) <> "" And txt_HTM_COND3 = "" Then
                   MsgBox "请确认热处理条件三", vbOKOnly + vbQuestion, "提示"
                   Exit Sub
                End If
            
                ss1.Col = SS1_LINE:               ss1.Text = txt_PrcLine
                ss1.Col = SS1_HTM_M1:             ss1.Text = txt_HTM_METH1
                ss1.Col = SS1_HTM_C1:             ss1.Text = txt_HTM_COND1
                ss1.Col = SS1_HTM_M2:             ss1.Text = txt_HTM_METH2
                ss1.Col = SS1_HTM_C2:             ss1.Text = txt_HTM_COND2
                ss1.Col = SS1_HTM_M3:             ss1.Text = txt_HTM_METH3
                ss1.Col = SS1_HTM_C3:             ss1.Text = txt_HTM_COND3
               
            Else
            
               If chk_can(0).Value = 0 And chk_can(1).Value = 0 And chk_can(2).Value = 0 Then
                   MsgBox "请选择取消方法", vbOKOnly + vbQuestion, "提示"
                   Exit Sub
               End If
               
               If chk_can(0).Value = -1 And chk_can(1).Value = 0 And chk_can(2).Value = -1 Then
                   MsgBox "请按顺序选择取消方法", vbOKOnly + vbQuestion, "提示"
                   Exit Sub
               End If
               
               If chk_can(0).Value = -1 Then
                  ss1.Col = SS1_HTM_M1:             ss1.Text = ""
                  ss1.Col = SS1_HTM_C1:             ss1.Text = ""
               End If
               
               If chk_can(1).Value = -1 Then
                  ss1.Col = SS1_HTM_M2:             ss1.Text = ""
                  ss1.Col = SS1_HTM_C2:             ss1.Text = ""
               End If
               
               If chk_can(2).Value = -1 Then
                  ss1.Col = SS1_HTM_M3:             ss1.Text = ""
                  ss1.Col = SS1_HTM_C3:             ss1.Text = ""
               End If
               
            End If
            
        End If
        
        If ORGCOL = SS1_REMARK Then
        
'            If txt_PrcLine = "1" Then
'                If Trim(txt_REMARKS) = "" Then
'                   MsgBox "请确认备注内容", vbOKOnly + vbQuestion, "提示"
'                   Exit Sub
'                End If
'            End If
            
            ss1.Col = 0:            ss1.Text = "Update"
            ss1.Col = SS1_LINE:     ss1.Text = txt_PrcLine
            ss1.Col = SS1_REMARK:   ss1.Text = ss1.Text        'ss1.Text = txt_REMARKS
            
        End If
    End If
    
    Text1_PLATE_NO.Text = ""
    
    txt_REMARKS.Text = ""
    
    
    ss1.Col = SS1_PLATE_NO:   Text1_PLATE_NO.Text = ss1.Text        'ss1.Text = txt_REMARKS
    
    
        If Gf_Ms_Refer(M_CN1, Mc2, , , False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc2("pControl"), True)
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
        DD.rControl.Add Item:=txt_plt_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

    Else

        If Len(Trim(txt_plt)) = txt_plt.MaxLength Then
            txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
        Else
            txt_plt_name.Text = ""
        End If
    
    End If
    
End Sub

Private Sub txt_SB_DblClick()
    Call txt_SB_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_SB_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0074"
        DD.rControl.Add Item:=txt_SB
        'DD.rControl.Add Item:=TXT_shot_blast_nm

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

    End If

End Sub

Private Sub txt_UST_DblClick()
    Call txt_UST_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_UST_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "Q0046"
        DD.rControl.Add Item:=txt_UST
    
        DD.nameType = "2"
        
        Call Gf_Mill_Common_DD(M_CN1, vbKeyF4)
    End If

End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

Private Sub SDT_PROD_DATE_FROM_GotFocus()
    SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
    SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
End Sub

Private Sub SDT_PROD_DATE_TO_GotFocus()
     SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
End Sub

Private Sub txt_stdspec_chg_DblClick()
    Call txt_stdspec_chg_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec_chg

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

    End If
    
End Sub

