VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form ACB5070C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "产品交接信息查询及修改_ACB5070C"
   ClientHeight    =   9255
   ClientLeft      =   225
   ClientTop       =   2250
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15330
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_CHK 
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
      Left            =   16530
      MaxLength       =   7
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt_htm_ord_name 
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
      Left            =   11460
      MaxLength       =   7
      TabIndex        =   5
      Top             =   9390
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txt_htm_ord_cd 
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
      Left            =   10950
      MaxLength       =   7
      TabIndex        =   4
      Top             =   9390
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txt_prod_grd_s 
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
      Left            =   13605
      TabIndex        =   3
      Top             =   9390
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox text_cur_inv 
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
      Left            =   8040
      TabIndex        =   2
      Top             =   9210
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox txt_cust_cd_s 
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
      Left            =   12750
      TabIndex        =   1
      Top             =   9390
      Visible         =   0   'False
      Width           =   795
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9195
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   16219
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "ACB5070C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   6975
         Left            =   0
         TabIndex        =   35
         Top             =   2220
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   12303
         _Version        =   196609
         SplitterBarWidth=   4
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   16761087
         PaneTree        =   "ACB5070C.frx":0072
         Begin FPSpread.vaSpread ss3 
            Height          =   2520
            Left            =   0
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   0
            Width           =   6855
            _Version        =   393216
            _ExtentX        =   12091
            _ExtentY        =   4445
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ButtonDrawMode  =   4
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   12
            MaxRows         =   1
            Protect         =   0   'False
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "ACB5070C.frx":00E4
         End
         Begin FPSpread.vaSpread ss1 
            Height          =   4395
            Left            =   0
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   2580
            Width           =   6855
            _Version        =   393216
            _ExtentX        =   12091
            _ExtentY        =   7752
            _StockProps     =   64
            ButtonDrawMode  =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   28
            MaxRows         =   2
            ProcessTab      =   -1  'True
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "ACB5070C.frx":07C5
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   6975
            Left            =   6915
            TabIndex        =   38
            Top             =   0
            Width           =   8280
            _Version        =   393216
            _ExtentX        =   14605
            _ExtentY        =   12303
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
            MaxCols         =   13
            MaxRows         =   2
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "ACB5070C.frx":136B
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   900
         Left            =   0
         TabIndex        =   7
         Top             =   1290
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1588
         _Version        =   196609
         BackColor       =   14737632
         ShadowStyle     =   1
         Begin VB.TextBox text_cur_inv_name 
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
            Left            =   7320
            TabIndex        =   14
            Tag             =   "起始库"
            Top             =   90
            Width           =   1230
         End
         Begin VB.TextBox txt_to_inv_name 
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
            Left            =   7320
            TabIndex        =   13
            Tag             =   "目标库"
            Top             =   450
            Width           =   1230
         End
         Begin VB.TextBox text_cur_inv_code 
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
            MaxLength       =   2
            TabIndex        =   12
            Tag             =   "起始库代码"
            Top             =   90
            Width           =   435
         End
         Begin VB.TextBox txt_to_inv 
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
            MaxLength       =   2
            TabIndex        =   11
            Tag             =   "目标库代码"
            Top             =   450
            Width           =   435
         End
         Begin VB.TextBox txt_mv_lst_no 
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
            Left            =   2670
            MaxLength       =   15
            TabIndex        =   10
            Tag             =   "移拨码单号"
            Top             =   450
            Width           =   2070
         End
         Begin VB.TextBox txt_t_addr 
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
            Left            =   10500
            MaxLength       =   7
            TabIndex        =   9
            Top             =   450
            Width           =   1230
         End
         Begin Threed.SSCheck SSCHK 
            Height          =   405
            Left            =   180
            TabIndex        =   8
            Top             =   45
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   714
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
            Caption         =   "转库查询"
         End
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   11970
            Top             =   450
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "已选择量"
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
            ForeColor       =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_plate_num 
            Height          =   315
            Left            =   13170
            TabIndex        =   15
            Top             =   450
            Width           =   645
            _Version        =   262145
            _ExtentX        =   1138
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   255
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            NumIntDigits    =   7
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_plate_wgt 
            Height          =   315
            Left            =   13845
            TabIndex        =   16
            Top             =   450
            Width           =   1230
            _Version        =   262145
            _ExtentX        =   2170
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   16711680
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0.000"
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel13 
            Height          =   315
            Left            =   9300
            Top             =   450
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "目标垛位"
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
            Left            =   1410
            Top             =   90
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "转库日期"
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
         Begin InDate.UDate udate_in_plt_date_a 
            Height          =   315
            Left            =   2670
            TabIndex        =   17
            Tag             =   "交接期"
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
            MaxLength       =   10
         End
         Begin InDate.UDate udate_in_plt_date_b 
            Height          =   315
            Left            =   4125
            TabIndex        =   18
            Tag             =   "交接期"
            Top             =   90
            Width           =   1410
            _ExtentX        =   2487
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
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   1410
            Top             =   450
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "移拨码单号"
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
            Left            =   5760
            Top             =   90
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            Caption         =   "起始库"
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
         Begin InDate.ULabel ULabel10 
            Height          =   315
            Left            =   5760
            Tag             =   "目标库"
            Top             =   450
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            Caption         =   "目标库"
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
      Begin Threed.SSFrame SSFrame1 
         Height          =   1260
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   2223
         _Version        =   196609
         BackColor       =   14737632
         ShadowStyle     =   1
         Begin VB.ComboBox CBO_CUR_INV 
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
            ItemData        =   "ACB5070C.frx":1A8D
            Left            =   8235
            List            =   "ACB5070C.frx":1A9D
            TabIndex        =   24
            Tag             =   "库"
            Top             =   480
            Width           =   765
         End
         Begin VB.TextBox txt_f_addr 
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
            Left            =   9015
            MaxLength       =   10
            TabIndex        =   23
            Top             =   480
            Width           =   1275
         End
         Begin VB.ComboBox CBO_PLT 
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
            Height          =   315
            ItemData        =   "ACB5070C.frx":1AB1
            Left            =   8235
            List            =   "ACB5070C.frx":1ABE
            TabIndex        =   22
            Tag             =   "生产厂"
            Text            =   "C1"
            Top             =   90
            Width           =   765
         End
         Begin VB.TextBox txt_mat_no 
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
            Left            =   3435
            TabIndex        =   21
            Tag             =   "轧批号"
            Top             =   870
            Width           =   2070
         End
         Begin VB.TextBox txt_stdspec_s 
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
            Left            =   3435
            TabIndex        =   20
            Top             =   480
            Width           =   2870
         End
         Begin CSTextLibCtl.sidbEdit txt_len_min_s 
            Height          =   315
            Left            =   12210
            TabIndex        =   25
            Top             =   870
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   99999999
            MinValue        =   -99999999
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txt_wid_min_s 
            Height          =   315
            Left            =   12210
            TabIndex        =   26
            Top             =   480
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
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
            NumIntDigits    =   9
            ShowZero        =   0   'False
            MaxValue        =   99999999
            MinValue        =   -99999999
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txt_thk_min_s 
            Height          =   315
            Left            =   12210
            TabIndex        =   27
            Top             =   90
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
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
         Begin CSTextLibCtl.sidbEdit txt_len_max_s 
            Height          =   315
            Left            =   13335
            TabIndex        =   28
            Top             =   870
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txt_wid_max_s 
            Height          =   315
            Left            =   13335
            TabIndex        =   29
            Top             =   480
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
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
            NumIntDigits    =   9
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txt_thk_max_s 
            Height          =   315
            Left            =   13335
            TabIndex        =   30
            Top             =   90
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
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
         Begin InDate.ULabel ULabel17 
            Height          =   315
            Left            =   2175
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
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
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Left            =   10950
            Top             =   90
            Width           =   1230
            _ExtentX        =   2170
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   10950
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
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
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   10950
            Top             =   870
            Width           =   1230
            _ExtentX        =   2170
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
         End
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   2175
            Top             =   90
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "生产日期"
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
         Begin InDate.UDate udt_prod_date_fr 
            Height          =   315
            Left            =   3435
            TabIndex        =   31
            Tag             =   "交接期"
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
            MaxLength       =   10
         End
         Begin InDate.UDate udt_prod_date_to 
            Height          =   315
            Left            =   4890
            TabIndex        =   32
            Tag             =   "交接期"
            Top             =   90
            Width           =   1410
            _ExtentX        =   2487
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
         Begin InDate.ULabel ULabel19 
            Height          =   315
            Left            =   2175
            Top             =   870
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "物料号"
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
         Begin Threed.SSOption opt_del_bed 
            Height          =   315
            Left            =   540
            TabIndex        =   33
            Top             =   720
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
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
            Caption         =   "退入库"
         End
         Begin Threed.SSOption opt_wait 
            Height          =   315
            Left            =   540
            TabIndex        =   34
            Top             =   240
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
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
            Caption         =   "指定货位"
            Value           =   -1
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   6975
            Top             =   90
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "生产厂"
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
            ForeColor       =   16711680
         End
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   6975
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "库/垛位"
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
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   9780
      Top             =   9390
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
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
      ForeColor       =   16711680
   End
End
Attribute VB_Name = "ACB5070C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name
'-- Program Name
'-- Program ID        ACB5070C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim.Sung.Ho
'-- Coder             Kim.Sung.Ho
'-- Date              2007.12.1
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

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_PLATE_NO = 1
Const SS1_USERID = 27
Const SS2_PLATE_NO = 2
Const SS2_USERID = 13
Const SS3_MV_LST_NO = 1
Const SS3_TO_INV = 4

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
       
         Call Gp_Ms_Collection(TXT_CHK, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_PLT, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(udt_prod_date_fr, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(udt_prod_date_to, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_mat_no, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_stdspec_s, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(CBO_CUR_INV, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(text_cur_inv, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_f_addr, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_t_addr, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_thk_min_s, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_thk_max_s, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_wid_min_s, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_wid_max_s, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_len_min_s, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_len_max_s, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_mv_lst_no, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_plate_num, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_plate_wgt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             
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
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
 
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB5070C.P_SMODIFY", Key:="P-M"
    sc1.Add Item:="ACB5070C.P_SREFER", Key:="P-R"
    sc1.Add Item:="ACB5070C.P_SONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_Collection(ss2, 1, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ACB5070C.P_SMODIFY2", Key:="P-M"
    sc2.Add Item:="ACB5070C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    ' control part   Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, 24, True)

End Sub

Private Sub CBO_CUR_INV_Click()
    text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", CBO_CUR_INV.Text, 2)
End Sub

Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuTool_ReSet
    
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
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc3.Item("Spread"), False)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)

    Call Gp_Spl_SizeGet(SSSplitter2, "C-System.INI", Me.Name, "W")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "C-System.INI", Me.Name)
    
    Call Gp_Sp_ColHidden(ss1, 2, True)
    Call Gp_Sp_ColHidden(ss2, 3, True)

    udt_prod_date_fr.RawData = Format(Date, "YYYYMMDD")
    udt_prod_date_to.RawData = Format(Date, "YYYYMMDD")
    udate_in_plt_date_a.RawData = Format(Date, "YYYYMMDD")
    udate_in_plt_date_b.RawData = Format(Date, "YYYYMMDD")
    txt_to_inv.Text = "WG"
    CBO_PLT.Text = "C1"
    CBO_CUR_INV.Text = "00"
    txt_t_addr = "P2A0101"
    Call CBO_CUR_INV_KeyUp(0, 0)
    opt_wait.Value = True
    
    SSSplitter1.Panes(0).LockHeight = True
    SSSplitter1.Panes(1).LockHeight = True
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter2, "C-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc3.Item("Spread"), "C-System.INI", Me.Name)
    
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
    
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) And Gf_Sp_Cls(sc2) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        CBO_CUR_INV.Text = "WG"
        Call CBO_CUR_INV_KeyUp(0, 0)
        sdb_plate_num.Value = 0
        sdb_plate_wgt.Value = 0
        opt_del_bed.Value = True
        opt_del_bed.Enabled = True
        opt_wait.Enabled = True
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    Dim iRow  As Integer
    Dim sRow As Integer
    Dim tRow As Integer
    Dim SMESG As String
    
    Dim I As Integer
 
    Dim iCol As Integer

'    If Gf_Sp_ProceExist(sc1.Item("Spread")) Then Exit Sub
    If SSCHK.Value = -1 Then
    
        If (udate_in_plt_date_a.RawData <> "" And udate_in_plt_date_b.RawData <> "") _
           And (text_cur_inv_name.Text <> "" Or txt_to_inv_name.Text <> "") Then
           Call Mv_lst_no_Ref
           ss2.OperationMode = OperationModeNormal
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           Call MenuTool_ReSet
        End If
        
        Exit Sub
        
    End If
            
    Call Sub_Ref
    
     For iRow = 1 To ss1.MaxRows
    
               ss1.Row = iRow
               ss1.Col = 28
                If ss1.Text = "Y" Then
                  For I = 1 To ss1.MaxCols
                       ss1.Col = I
                       ss1.ForeColor = &HC000&
                  Next
                End If
        
         Next iRow
    

End Sub

Private Sub Mv_lst_no_Ref()

    Dim sQuery      As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    ss1.MaxRows = 0
    ss2.MaxRows = 0
    
    sQuery = "          Select   MV_LST_NO,"
    sQuery = sQuery & "          FR_INV,"
    sQuery = sQuery & "          Gf_ComnNameFind('C0013',FR_INV),"
    sQuery = sQuery & "          TO_INV,"
    sQuery = sQuery & "          Gf_ComnNameFind('C0013',TO_INV),"
    sQuery = sQuery & "          MOVE_CAR_NO,"
    sQuery = sQuery & "          COUNT(*),"
    sQuery = sQuery & "          SUM(WGT),"
    sQuery = sQuery & "          DECODE(MAX(MOVE_DATE),NULL,NULL,MAX(SUBSTR(MOVE_DATE,1,4)||'-'||SUBSTR(MOVE_DATE,5,2)||'-'||SUBSTR(MOVE_DATE,7,2))),"
    sQuery = sQuery & "          GF_EMPNAMEFIND(MAX(MOVE_EMP)),"
    sQuery = sQuery & "          DECODE(MAX(RCV_DATE),NULL,NULL,MAX(SUBSTR(RCV_DATE,1,4)||'-'||SUBSTR(RCV_DATE,5,2)||'-'||SUBSTR(RCV_DATE,7,2))),"
    sQuery = sQuery & "          GF_EMPNAMEFIND(MAX(RCV_EMP))"
    sQuery = sQuery & "   FROM   CP_MOVE_SLT "
    sQuery = sQuery & "  WHERE   PROD_CD = 'PP' "
    sQuery = sQuery & "    AND   FR_INV  LIKE '" & Trim(text_cur_inv_code.Text) + "%' "
    sQuery = sQuery & "    AND   TO_INV  LIKE '" & Trim(txt_to_inv.Text) & "%' "
    
    If Trim(txt_mv_lst_no.Text) <> "" Then
        sQuery = sQuery & "   AND MV_LST_NO  Like '" & Trim(txt_mv_lst_no.Text) & "%' "
    End If
    
    If IsDate(udate_in_plt_date_a.Text) Then
        sQuery = sQuery & "   AND MOVE_DATE >= '" & udate_in_plt_date_a.RawData & "'"
    End If
    
    If IsDate(udate_in_plt_date_b.Text) Then
        sQuery = sQuery & "   AND MOVE_DATE <= '" & udate_in_plt_date_b.RawData & "'"
    End If
    
    sQuery = sQuery & "   Group By MV_LST_NO,PROD_CD,FR_INV,TO_INV,MOVE_CAR_NO"
    sQuery = sQuery & "   Order By MV_LST_NO DESC"
                    
    If Gf_Only_Display(M_CN1, Sc3, sQuery) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss3.OperationMode = OperationModeNormal
    End If

End Sub

Public Sub Form_Pro()

    Dim iRow As Integer
    
    If TXT_CHK.Text = "D" Then
    
        If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call MenuTool_ReSet
        End If
    Else
        If Gf_Mill_Process(M_CN1, sc2, Mc1, , "P") Then
            Call Form_Ref
        End If
    
    End If
    

    
End Sub

Public Sub Spread_ColumnsSort()
    Spread_ColSort.Show 1
End Sub

Public Sub Spread_Forzens_Setting()
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
End Sub

Public Sub Spread_Forzens_Cancel()
    Me.ActiveControl.ColsFrozen = 0
End Sub

Public Sub Spread_Del()

'    Call Gp_Sp_Del(sc1)
    
End Sub

Public Sub Spread_Can()

    Call Form_Ref
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub opt_del_bed_Click(Value As Integer)

    opt_del_bed.ForeColor = &HFF&
    opt_wait.ForeColor = &H808080
    TXT_CHK.Text = "D"
    
End Sub

Private Sub opt_wait_Click(Value As Integer)

    opt_del_bed.ForeColor = &H808080
    opt_wait.ForeColor = &HFF&
    TXT_CHK.Text = "W"
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ss1.MaxRows > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim plate_no As String
    Dim iCnt As Integer
    Dim iPlate_cnt As Integer
    Dim iPlate_wgt As Double
    
    Dim tRow  As Integer
    Dim delete As String

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
   If Row <= 0 Then Exit Sub
    
   If TXT_CHK.Text = "D" Then
        ss1.Row = Row
        ss1.Col = 0
         
        If ss1.Text = "" Then
             ss1.Col = 0
             If opt_del_bed Then
                 ss1.Text = "Delete"
             Else
                 ss1.Text = "Update"
             End If
             ss1.Col = SS1_USERID:   ss1.Text = sUserID
             
             Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
        Else
             ss1.Col = 0:              ss1.Text = ""
             ss1.Col = SS1_USERID:     ss1.Text = ""
             Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
        End If
   End If
   
   If TXT_CHK.Text = "W" Then
    
    delete = ""

    If ss1.MaxRows < 1 Then Exit Sub
    
    If Col = 0 Then
    
        iPlate_cnt = 0
        iPlate_wgt = 0
        
            ss1.Row = Row
            ss1.Col = 0
            If ss1.Text = "Delete" Or ss1.Text = "Input" Or ss1.Text = "Update" Then
                delete = "Y"
            End If
            
            ss1.Col = SS1_PLATE_NO
            plate_no = Trim(ss1.Text)
        
            If ss2.MaxRows = 0 Or plate_no = "" Then
               Exit Sub
            End If
            
            If delete = "Y" Then
                With ss2
                    
                    For iCnt = .MaxRows To 1 Step -1
                       .Col = 0:    .Row = iCnt
                        If Trim(.Text) = "Input" Then
                           .Col = SS2_PLATE_NO
                            If .Text = plate_no Then
                               .Text = "":                .BackColor = &HC0FFFF
                               .Col = 0:                  .Text = ""
                                Exit For
                            End If
                        End If
                    Next iCnt
                     
                End With

                With ss1
                       .Col = 0:                          .Text = ""
                       .Col = SS1_PLATE_NO:               .BackColor = &HC0FFFF
                End With
                Exit Sub
            End If
        
            ss1.Row = Row
            ss1.Col = 0:                       ss1.Text = "Delete"
            ss1.Col = SS1_PLATE_NO:            ss1.BackColor = &HFFC0FF
                            
            With ss2
                
                tRow = .ActiveRow
                .Row = tRow:                .Col = SS2_PLATE_NO
            
                If Len(.Text) = 14 Then
                
                     For iCnt = 1 To .MaxRows     '.MaxRows To 1 Step -1
                        .Col = SS2_PLATE_NO:      .Row = iCnt
                         If Trim(.Text) <> "" Then
                            .Row = iCnt - 1
                            If .Row < 1 Then
                               Call Gp_MsgBoxDisplay("此垛位无可用垛层位置")
                            End If
                            .Text = plate_no:     .BackColor = &HFFC0FF
                            .Col = 0:             .Text = "Input"
                            .Col = SS2_USERID:    .Text = sUserID:           .BackColor = &HFFC0FF
                             Exit Sub
                         End If
                     Next iCnt
                     
                Else
                
                    .Col = SS2_PLATE_NO
                    .Row = tRow
                     If Trim(.Text) = "" Then
                        .Text = plate_no:                .BackColor = &HFFC0FF
                        .Col = 0:                        .Text = "Input"
                        .Col = SS2_USERID:               .Text = sUserID
                         If tRow > 1 Then
                         Call .SetActiveCell(1, tRow - 1)
                         End If
                         Exit Sub
                     End If
                     
                End If
                 
            End With
    
    End If
   
   End If

End Sub

Private Sub CBO_CUR_INV_DblClick()

    Call CBO_CUR_INV_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub CBO_CUR_INV_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=CBO_CUR_INV
        DD.rControl.Add Item:=text_cur_inv
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
     
        If Len(Trim(CBO_CUR_INV.Text)) = 2 Then
            text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", CBO_CUR_INV.Text, 2)
        Else
          text_cur_inv.Text = ""
        End If
        
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



Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

'    ss2.Row = Row + 1
'    ss2.Col = SS2_PLATE_NO
'    If Trim(ss2.Text) = "" Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") Then
    
        Call Gp_Sp_UpdateMake(ss2, Mode)
        
        ss2.Col = SS2_USERID:        ss2.Text = sUserID
     
    End If
End Sub

Private Sub ss3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)

    If ss3.MaxRows < 1 Then Exit Sub
    If Row < 1 Then Exit Sub
    If Col = 1 Then
       ss2.MaxRows = 0
       Call Gp_Sp_BlockColor(ss3, SS3_MV_LST_NO, SS3_MV_LST_NO, 1, ss3.MaxRows, , &HC0FFFF)
       Call Gp_Sp_BlockColor(ss3, SS3_MV_LST_NO, SS3_MV_LST_NO, Row, Row, , &HFFC0FF)
       ss3.Row = Row:     ss3.Col = Col:            txt_mv_lst_no.Text = ss3.Text
       ss3.Row = Row:     ss3.Col = SS3_TO_INV:     CBO_CUR_INV.Text = ss3.Text
       If Len(txt_mv_lst_no.Text) = 15 Then
          Call Sub_Ref
       End If
    End If
    
End Sub

Private Sub ss3_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub
Private Sub ss3_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If ss3.MaxRows > 0 Then
        Set Active_Spread = Me.ss3
        PopupMenu MDIMain.PopUp_Spread
    End If
End Sub

Private Sub SSCHK_Click(Value As Integer)
    ss1.MaxRows = 0
    ss2.MaxRows = 0
    ss3.MaxRows = 0
End Sub

Private Sub text_cur_inv_code_DblClick()

    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_cur_inv_code_Change()
    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
        text_cur_inv_name.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
    Else
      text_cur_inv_name.Text = ""
    End If
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
         DD.sWitch = "MS"
         DD.sKey = "C0013"
    
         DD.rControl.Add Item:=text_cur_inv_code
    
         DD.nameType = "2"
         Call Gf_Common_DD(M_CN1, KeyCode)
    
    End If

End Sub

Private Sub txt_stdspec_s_DblClick()

    Call txt_stdspec_s_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stdspec_s_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec_s

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)
    End If
    
End Sub
Private Sub txt_to_inv_DblClick()

    Call txt_to_inv_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_to_inv_Change()
    If Len(Trim(txt_to_inv.Text)) = txt_to_inv.MaxLength Then
        txt_to_inv_name.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_to_inv.Text, 2)
    Else
        txt_to_inv_name.Text = ""
    End If
End Sub

Private Sub txt_to_inv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
        DD.rControl.Add Item:=txt_to_inv
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    Else
        If Len(Trim(txt_to_inv.Text)) = txt_to_inv.MaxLength Then
            txt_to_inv_name.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_to_inv.Text, 2)
        Else
            txt_to_inv_name.Text = ""
        End If
    
    End If
    
End Sub


Private Function AC_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, sPrc As String, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim sQuery As String
    Dim intCount As Integer
    Dim AdoRs As ADODB.Recordset
    
    If Trim(sPrc) = "" Then
        AC_ComboAdd = False: Exit Function
    End If
    
    intCount = 1
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then AC_ComboAdd = False: Exit Function
    End If
    
    sQuery = "SELECT CD_NAME FROM ZP_CD Where CD_MANA_NO = '" + Trim(sPrc) + "'"

    If ClsChk Then
        Cbo.Clear
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
                If intCount = 6 Then intCount = 7
                Cbo.AddItem Trim(str(intCount)) + ":" + AdoRs.Fields(0)
                intCount = intCount + 1
            End If
            AdoRs.MoveNext
            
        Wend
        AC_ComboAdd = True
    Else
        AC_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    AC_ComboAdd = False

End Function

Private Sub Sub_Ref()

    Dim iRow  As Integer
    Dim sRow As Integer
    Dim tRow As Integer
    Dim sMsg As String
            
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        ss1.OperationMode = OperationModeNormal
        If text_cur_inv.Text <> "" And Len(txt_t_addr.Text) = 7 Then
           Call Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False)
           ss2.OperationMode = OperationModeNormal
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           Call MenuTool_ReSet
        End If
    End If
    
    With ss2
    
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = SS2_PLATE_NO
             If Trim(.Text) <> "" Then
                sRow = iRow
                Exit For
             End If
             sRow = .MaxRows
         Next iRow
         
         tRow = sRow + 15
         If tRow > .MaxRows Then
            tRow = .MaxRows
         End If
         
         Call .SetActiveCell(1, tRow)
         
    End With

End Sub

