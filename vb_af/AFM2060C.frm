VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AFM2060C 
   Caption         =   "板坯切割作业界面_AFM2060C"
   ClientHeight    =   9225
   ClientLeft      =   180
   ClientTop       =   2070
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox txt_plan_mill_plt 
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
      ItemData        =   "AFM2060C.frx":0000
      Left            =   12120
      List            =   "AFM2060C.frx":0010
      TabIndex        =   30
      Top             =   540
      Width           =   780
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
      Left            =   8160
      MaxLength       =   2
      TabIndex        =   28
      Top             =   540
      Width           =   525
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
      Left            =   8670
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   540
      Width           =   1740
   End
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   630
      Left            =   13740
      TabIndex        =   9
      Top             =   630
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1111
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   14737632
      ActiveColors    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "指示取消"
   End
   Begin Threed.SSOption opt_prc_status1 
      Height          =   375
      Left            =   1320
      TabIndex        =   23
      Top             =   120
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
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
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7845
      Left            =   60
      TabIndex        =   12
      Top             =   1320
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   13838
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AFM2060C.frx":0022
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   3195
         Left            =   0
         TabIndex        =   14
         Top             =   4650
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   5636
         _Version        =   196609
         SplitterBarWidth=   3
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   16761087
         PaneTree        =   "AFM2060C.frx":0074
         Begin Threed.SSPanel SSPanel1 
            Height          =   555
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   979
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox TXT_INGOT_FL 
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
               Left            =   13050
               MaxLength       =   10
               TabIndex        =   29
               Top             =   120
               Visible         =   0   'False
               Width           =   570
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
               Left            =   12030
               MaxLength       =   10
               TabIndex        =   17
               Top             =   120
               Visible         =   0   'False
               Width           =   900
            End
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
               ItemData        =   "AFM2060C.frx":00C6
               Left            =   1605
               List            =   "AFM2060C.frx":00C8
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Tag             =   "连铸机号"
               Top             =   120
               Width           =   705
            End
            Begin InDate.ULabel ULabel4 
               Height          =   315
               Left            =   270
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
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel10 
               Height          =   315
               Left            =   2940
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
                  Size            =   9.75
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
               Left            =   4280
               TabIndex        =   18
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
               Left            =   5490
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
                  Size            =   9.75
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
               Left            =   6830
               TabIndex        =   19
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
               Left            =   8040
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
                  Size            =   9.75
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
               Left            =   9360
               TabIndex        =   20
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
               TabIndex        =   21
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
            Height          =   2595
            Left            =   0
            TabIndex        =   22
            Top             =   600
            Width           =   15105
            _Version        =   393216
            _ExtentX        =   26644
            _ExtentY        =   4577
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
            MaxCols         =   18
            MaxRows         =   2
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AFM2060C.frx":00CA
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   4590
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   15105
         _Version        =   393216
         _ExtentX        =   26644
         _ExtentY        =   8096
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
         MaxCols         =   32
         MaxRows         =   2
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AFM2060C.frx":0A98
      End
   End
   Begin VB.TextBox txt_ord_no 
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
      Left            =   8160
      MaxLength       =   11
      TabIndex        =   11
      Tag             =   "CD_MANA_NO"
      Top             =   120
      Width           =   1530
   End
   Begin VB.ComboBox cbo_ord_item 
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
      Left            =   9690
      TabIndex        =   10
      Top             =   120
      Width           =   720
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
      Left            =   14280
      MaxLength       =   20
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   870
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
      Left            =   13620
      MaxLength       =   20
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   870
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
      Left            =   1830
      MaxLength       =   11
      TabIndex        =   6
      Top             =   540
      Width           =   1620
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   1290
      MaxLength       =   2
      TabIndex        =   0
      Top             =   540
      Width           =   525
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
      Height          =   255
      Left            =   3480
      MaxLength       =   11
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   195
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
      Height          =   345
      Left            =   11925
      TabIndex        =   4
      Top             =   930
      Width           =   1650
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
      Left            =   5160
      MaxLength       =   10
      TabIndex        =   2
      Top             =   120
      Width           =   1440
   End
   Begin VB.TextBox txt_LOC 
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
      Left            =   5160
      MaxLength       =   11
      TabIndex        =   3
      Top             =   540
      Width           =   1440
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
      Height          =   345
      Left            =   10515
      MaxLength       =   11
      TabIndex        =   1
      Top             =   930
      Width           =   1410
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   120
      Top             =   150
      Width           =   1125
      _ExtentX        =   1984
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
      Top             =   930
      Width           =   1125
      _ExtentX        =   1984
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
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   6990
      Top             =   120
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "订单号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   3990
      Top             =   540
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "垛位号"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   3990
      Top             =   120
      Width           =   1125
      _ExtentX        =   1984
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   3180
      Top             =   930
      Width           =   1125
      _ExtentX        =   1984
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
   Begin Threed.SSOption opt_prc_status2 
      Height          =   375
      Left            =   2550
      TabIndex        =   24
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
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
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   10770
      Top             =   120
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "生成日期"
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
   Begin InDate.UDate SDT_PROD_DATE_FROM 
      Height          =   315
      Left            =   11940
      TabIndex        =   25
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
      Left            =   13395
      TabIndex        =   26
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
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   10770
      Top             =   540
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   $"AFM2060C.frx":1BDB
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   120
      Top             =   540
      Width           =   1125
      _ExtentX        =   1984
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
   Begin InDate.ULabel ULabel15 
      Height          =   315
      Left            =   6990
      Top             =   540
      Width           =   1125
      _ExtentX        =   1984
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
   Begin CSTextLibCtl.sidbEdit txt_len 
      Height          =   315
      Left            =   7365
      TabIndex        =   31
      Top             =   930
      Width           =   795
      _Version        =   262145
      _ExtentX        =   1402
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
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_wid_to 
      Height          =   315
      Left            =   5115
      TabIndex        =   32
      Top             =   930
      Width           =   795
      _Version        =   262145
      _ExtentX        =   1402
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
      NumIntDigits    =   4
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_len_to 
      Height          =   315
      Left            =   8160
      TabIndex        =   33
      Top             =   930
      Width           =   795
      _Version        =   262145
      _ExtentX        =   1402
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
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_wid 
      Height          =   315
      Left            =   4350
      TabIndex        =   34
      Top             =   930
      Width           =   795
      _Version        =   262145
      _ExtentX        =   1402
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
      NumIntDigits    =   4
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   9240
      Top             =   930
      Width           =   1215
      _ExtentX        =   2143
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   6180
      Top             =   930
      Width           =   1125
      _ExtentX        =   1984
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
   Begin CSTextLibCtl.sidbEdit txt_thk_to 
      Height          =   315
      Left            =   2085
      TabIndex        =   35
      Top             =   930
      Width           =   795
      _Version        =   262145
      _ExtentX        =   1402
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
   Begin CSTextLibCtl.sidbEdit txt_thk 
      Height          =   315
      Left            =   1290
      TabIndex        =   36
      Top             =   930
      Width           =   795
      _Version        =   262145
      _ExtentX        =   1402
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
End
Attribute VB_Name = "AFM2060C"
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
'-- Program Name      板坯切割作业界面
'-- Program ID        AFM2060C
'-- Designer          SHIN.C.S
'-- Coder             SHIN.C.S
'-- Date              2007.7.25
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

Const SS1_EMP_CD = 17                   'EMP_CD
Const SS2_SLAB_NO = 1                   'SLAB_NO
Const SS2_INGOT_FL = 2                  'INGOT_FL
Const SS2_SLAB_THK = 3                  'THK
Const SS2_SLAB_WID = 4                  'WID
Const SS2_SLAB_LEN = 5                  'LEN
Const SS2_SLAB_WGT = 6                  'SLAB_WGT
Const SS2_CAL_WGT = 7                   'CAL_WGT
Const SS2_SLAB_DATE = 8                 'DATE
Const SS2_SLAB_TIME = 9                 'TIME
Const SS2_CHG_GRD_RES = 10              'CHG_GRD_RES
Const SS2_SLAB_LOC = 11                 'LOC
Const SS2_EMP_CD = 12                   'INS_EMP_CD
Const SS2_MOSLAB_NO = 13                'MOSLAB_NO
Const SS2_LAST_STATUS = 14              'LAST_STATUS
Const SS2_ORD_FL = 16                   'ORD_FL
Const SS2_SLAB_CNT = 17                 'SLAB_CNT




Public Sub Form_Ins()

    If ss2.SelBlockRow2 = ss2.MaxRows Then
       ss2.Row = ss2.MaxRows
       ss2.Col = 0
       If ss2.Text <> "Delete" Then
            Call Gp_Sp_Ins(Proc_Sc("Sc2"))
            
            With ss1
                .Row = .ActiveRow
                .Col = SS1_EMP_CD       '17
                .Text = sUserID
            End With
            
            Call INS_WGT_CAL
        End If
    End If

End Sub

Public Sub Spread_Del()

Dim i As Integer

       
       For i = 1 To ss2.MaxRows
           ss2.Row = i
           ss2.Col = 0
           If UCase(ss2.Text) = "" Then
              ss2.Text = "Delete"
           End If
             '20140116
           ss2.Col = SS2_LAST_STATUS   '14
            If i = ss2.MaxRows Then
                ss2.Text = "Y"
            Else
                ss2.Text = ""
            End If
        '20140116
       Next i
       


End Sub

Public Sub Spread_Can()

    If ss2.SelBlockRow2 = ss2.MaxRows Then
       ss2.Row = ss2.MaxRows
       ss2.Col = 0
       If ss2.Text = "Input" Then
            ss2.MaxRows = ss2.MaxRows - 1
            If ss2.MaxRows > 0 Then
                Call CANCEL_WGT_CAL
            End If
       End If
    End If
    
End Sub

Public Sub WGT_CAL(Row)

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
    Dim tot_slabwgt As Double
    Dim tmWid1 As Double
    
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
            ss2.Col = SS2_SLAB_THK   '3
            tmThk = ss2.VALUE
            ss2.Col = SS2_SLAB_WID   '4
            tmWid = ss2.VALUE
            ss2.Col = SS2_SLAB_LEN   '5
            tmLen = ss2.VALUE
            If ss2.Row = 1 Or tmWid = tmWid1 Then
               tmTotalLen = tmTotalLen + ss2.VALUE
               tmWid1 = tmWid
            End If
            tmCalCut = tmCalCut + (tmThk * tmWid * tmLen)
        End If
    Next i
        
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 0
        If ss2.Text <> "Delete" Then
            ss2.Row = i
            ss2.Col = SS2_SLAB_THK   '3
            tmThk = ss2.VALUE
            ss2.Col = SS2_SLAB_WID   '4
            tmWid = ss2.VALUE
            ss2.Col = SS2_SLAB_LEN   '5
            tmLen = ss2.VALUE
            
            tmCalCutOne = tmThk * tmWid * tmLen
            
            ss2.Col = SS2_SLAB_WGT    '6
            If tmCalCut <= tmCalMo Then
                tempWgt = tempWgt + Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
                sub_wgt = sub_wgt - Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
                'ss2.Value = Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
            Else
                tempWgt = tempWgt + Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
                sub_wgt = sub_wgt - Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
                'ss2.Value = Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
            End If
            If i < ss2.MaxRows Then
               ss2.VALUE = Val(Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('WGT',''," & tmThk & "," & tmWid & "," & tmLen & ",0) FROM DUAL"))
               tot_slabwgt = tot_slabwgt + ss2.VALUE
            Else
               ss2.VALUE = cSlabWgt - tot_slabwgt
            End If
            
            ss2.Col = SS2_CAL_WGT      '7
            ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
        End If
    Next i
    
    If tmCalCut = tmCalMo Then
        sub_len = cSlabLen
        sub_wgt = cSlabWgt
        For i = 1 To ss2.MaxRows
            ss2.Row = i
            If i < ss2.MaxRows Then
               ss2.Col = SS2_SLAB_WGT    '6
               sub_wgt = sub_wgt - ss2.VALUE
            End If
        Next i
        ss2.Row = ss2.MaxRows

        ss2.Col = SS2_SLAB_WGT       '6
        ss2.Text = sub_wgt
    End If
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 0
        If ss2.Text <> "Delete" Then
            ss2.Row = i
            ss2.Col = SS2_SLAB_WID   '4
            tmWid = ss2.VALUE
            ss2.Col = SS2_SLAB_LEN   '5
            If ss2.Row = 1 Or tmWid1 = tmWid Then
               tmTotalLen = tmTotalLen + ss2.VALUE
               tmWid1 = tmWid
            End If
            
            ss2.Col = SS2_SLAB_WGT    '6
            tempWgt = tempWgt + ss2.VALUE
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
    Dim S13 As String

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
    
    ss2.Col = SS2_SLAB_NO     '1
    S1 = ss2.VALUE
    
    ss2.Col = SS2_INGOT_FL    '2
    S2 = ss2.Text
    
    ss2.Col = SS2_SLAB_THK    '3
    S3 = ss2.VALUE
    
    ss2.Col = SS2_SLAB_WID    '4
    S4 = ss2.VALUE
    
    ss2.Col = SS2_SLAB_LEN     '5
    S5 = ss2.VALUE
    
    ss2.Col = SS2_SLAB_WGT     '6
    S6 = ss2.VALUE
    
    ss2.Col = SS2_CAL_WGT      '7
    S7 = ss2.VALUE
    
    ss2.Col = SS2_SLAB_DATE    '8
    S8 = ss2.Text
    
    ss2.Col = SS2_SLAB_TIME    '9
    S9 = ss2.Text
    
    ss2.Col = SS2_CHG_GRD_RES  '10
    S10 = ss2.Text
    
    ss2.Col = SS2_SLAB_LOC     '11
    S11 = ss2.Text
    
    ss2.Col = SS2_EMP_CD       '12
    S12 = ss2.Text
    
    ss2.Col = SS2_MOSLAB_NO    '13
    S13 = ss2.Text

    ' DATA PAST
    ss2.Row = ss2.MaxRows
    ss2.Col = SS2_SLAB_NO       '1
    S1 = Mid(S1, 1, 8) & CStr(CInt(Mid(S1, 9, 2)) + 1)
    ss2.Text = S1
    ss2.Col = SS2_INGOT_FL      '2
    ss2.Text = S2
    
    ss2.Col = SS2_SLAB_THK      '3
    ss2.Text = S3
    tmThk = S3
    ss2.Col = SS2_SLAB_WID      '4
    ss2.Text = S4
    tmWid = S4
    ss2.Col = SS2_SLAB_LEN      '5
    ss2.Text = S5
    tmLen = S5
    ss2.Col = SS2_SLAB_WGT      '6
    ss2.Text = S6
    ss2.Col = SS2_CAL_WGT       '7
    ss2.Text = S7
    ss2.Col = SS2_SLAB_DATE     '8
    ss2.Text = S8
    ss2.Col = SS2_SLAB_TIME     '9
    ss2.Text = S9
    ss2.Col = SS2_CHG_GRD_RES   '10
    ss2.Text = S10
    ss2.Col = SS2_SLAB_LOC      '11
    ss2.Text = S11
    ss2.Col = SS2_EMP_CD        '12
    ss2.Text = S12
    ss2.Col = SS2_MOSLAB_NO     '13
    ss2.Text = S13
    
    tmp_rat = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
         ss2.Row = i

         ss2.Col = SS2_SLAB_WGT      '6
         If i <> ss2.MaxRows Then
            ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
            tempWgt = tempWgt + ss2.VALUE
         Else
            ss2.Text = cSlabWgt - tempWgt
         End If
         tempWgt = tempWgt + ss2.VALUE
         sub_wgt = sub_wgt - ss2.VALUE
         
         ss2.Col = SS2_CAL_WGT       '7
         ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
    Next i
    
    sub_len = cSlabLen
    sub_wgt = cSlabWgt
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        If i <> ss2.MaxRows Then
           ss2.Col = SS2_SLAB_LEN      '5
           sub_len = sub_len - ss2.VALUE
           
           ss2.Col = SS2_SLAB_WGT      '6
           sub_wgt = sub_wgt - ss2.VALUE
        End If
    Next i
    ss2.Row = ss2.MaxRows
    ss2.Col = SS2_SLAB_LEN      '5
    ss2.Text = sub_len
    
    ss2.Col = SS2_SLAB_WGT      '6
    ss2.Text = sub_wgt
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = SS2_SLAB_LEN      '5
        tmTotalLen = tmTotalLen + ss2.VALUE
        
        ss2.Col = SS2_SLAB_WGT      '6
        tempWgt = tempWgt + ss2.VALUE

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
         
         ss2.Col = SS2_SLAB_THK      '3
         tmThk = ss2.VALUE
         
         ss2.Col = SS2_SLAB_WID      '4
         tmWid = ss2.VALUE
         
         ss2.Col = SS2_SLAB_LEN      '5
         ss2.Text = cfLen
         tmLen = cfLen
         
         ss2.Col = SS2_SLAB_LEN      '5
         If ss2.Row = ss2.MaxRows Then
            ss2.Text = cSlabLen - tmTotalLen
            tmTotalLen = tmTotalLen + ss2.Text
         Else
            ss2.Text = cfLen
            tmLen = cfLen
            tmTotalLen = tmTotalLen + cfLen
         End If
         
         ss2.Col = SS2_SLAB_WGT      '6
         ss2.Text = cfWgt
         tmWgt = tmWgt + ss2.VALUE
         
         ss2.Col = SS2_CAL_WGT       '7
         ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
    Next i
    
    sub_len = cSlabLen
    sub_wgt = cSlabWgt
    For i = 1 To delete_cnt
        ss2.Row = i
        If i <> delete_cnt Then
           ss2.Col = SS2_SLAB_LEN      '5
           sub_len = sub_len - ss2.VALUE
           
           ss2.Col = SS2_SLAB_WGT      '6
           sub_wgt = sub_wgt - ss2.VALUE
        End If
    Next i
    ss2.Row = delete_cnt
    ss2.Col = SS2_SLAB_LEN      '5
    ss2.Text = sub_len
    
    ss2.Col = SS2_SLAB_WGT      '6
    ss2.Text = sub_wgt
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To delete_cnt
        ss2.Row = i
        ss2.Col = SS2_SLAB_LEN      '5
        tmTotalLen = tmTotalLen + ss2.VALUE
        
        ss2.Col = SS2_SLAB_WGT      '6
        tempWgt = tempWgt + ss2.VALUE

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
        ss2.Col = SS2_SLAB_THK      '3
        tmThk = ss2.VALUE
        
        ss2.Col = SS2_SLAB_WID      '4
        tmWid = ss2.VALUE
        
        ss2.Col = SS2_SLAB_LEN      '5
        tmLen = ss2.VALUE
        
        ss2.Col = SS2_SLAB_WGT      '6
        If i <> ss2.MaxRows Then
           ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
           tempWgt = tempWgt + ss2.VALUE
        Else
           ss2.Text = cSlabWgt - tempWgt
        End If
        
        ss2.Col = SS2_CAL_WGT       '7
        ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
    Next i
    
    sub_len = cSlabLen
    sub_wgt = cSlabWgt
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        If i <> ss2.MaxRows Then
           ss2.Col = SS2_SLAB_LEN      '5
           sub_len = sub_len - ss2.VALUE
           
           ss2.Col = SS2_SLAB_WGT      '6
           sub_wgt = sub_wgt - ss2.VALUE
        End If
    Next i
    
    ss2.Row = ss2.MaxRows
    ss2.Col = SS2_SLAB_LEN      '5
    ss2.Text = sub_len
    
    ss2.Col = SS2_SLAB_WGT      '6
    ss2.Text = sub_wgt
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = SS2_SLAB_LEN      '5
        tmTotalLen = tmTotalLen + ss2.VALUE
        
        ss2.Col = SS2_SLAB_WGT      '6
        tempWgt = tempWgt + ss2.VALUE

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
            ss2.Col = SS2_SLAB_LEN      '5
            tmTotalLen = tmTotalLen + ss2.VALUE
        End If
    Next i
    
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = 0
        If UCase(ss2.Text) <> "DELETE" Then
            ss2.Row = i
            ss2.Col = SS2_SLAB_THK      '3
            tmThk = ss2.VALUE
            ss2.Col = SS2_SLAB_WID      '4
            tmWid = ss2.VALUE
            ss2.Col = SS2_SLAB_LEN      '5
            tmLen = ss2.VALUE
            
            tmCalCut = tmCalCut + (tmThk * tmWid * tmLen)
        End If
    Next i
        
    tmp_rat = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = SS2_SLAB_THK      '3
        tmThk = ss2.VALUE
        ss2.Col = SS2_SLAB_WID      '4
        tmWid = ss2.VALUE
        ss2.Col = SS2_SLAB_LEN      '5
        tmLen = ss2.VALUE
        
        tmCalCutOne = tmThk * tmWid * tmLen
        
        ss2.Col = SS2_SLAB_WGT      '6
        If tmCalCut <= tmCalMo Then
            tempWgt = tempWgt + Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
            sub_wgt = sub_wgt - Round((cSlabWgt * (cfLen / cSlabLen)), 3)
            ss2.VALUE = Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
        Else
            tempWgt = tempWgt + Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
            sub_wgt = sub_wgt - Round((cSlabWgt * (cfLen / tmTotalLen)), 3)
            ss2.VALUE = Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
        End If
        
        ss2.Col = SS2_CAL_WGT       '7
        ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
    Next i
    
    
    If tmCalCut >= tmCalMo Then
        sub_len = cSlabLen
        sub_wgt = cSlabWgt
        For i = 1 To ss2.MaxRows
            ss2.Row = i
            If i <> ss2.MaxRows Then
               ss2.Col = SS2_SLAB_WGT      '6
               sub_wgt = sub_wgt - ss2.VALUE
            End If
        Next i
        ss2.Row = ss2.MaxRows

        ss2.Col = SS2_SLAB_WGT      '6
        ss2.Text = sub_wgt
    End If
    
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.Row = i
        ss2.Col = SS2_SLAB_LEN      '5
        tmTotalLen = tmTotalLen + ss2.VALUE
        
        ss2.Col = SS2_SLAB_WGT      '6
        tempWgt = tempWgt + ss2.VALUE

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
            Call Gp_Ms_Collection(txt_Status, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_act_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_MOSLAB, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_LOC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_plt_dec, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_thk, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_wid, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_len, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_len_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(cbo_ord_item, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_plan_mill_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    
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
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AFM2060C.P_REFER", Key:="P-R"
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
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AFM2060C.P_MODIFY1", Key:="P-M"
    sc2.Add Item:="AFM2060C.P_REFER1", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    'Call Gp_Sp_ColHidden(ss2, 6, True)
    Call Gp_Sp_ColHidden(ss2, SS2_SLAB_LOC, True)    '11
    Call Gp_Sp_ColHidden(ss2, SS2_MOSLAB_NO, True)   '13
    Call Gp_Sp_ColHidden(ss2, SS2_LAST_STATUS, True) '14
    Call Gp_Sp_ColHidden(ss2, SS2_SLAB_CNT, True)    '17
    
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
        
        ss2.Col = SS2_INGOT_FL     '2
        ss2.Text = TXT_INGOT_FL
    
        ss2.Col = SS2_SLAB_THK     '3
        ss2.Text = cSlabthk
        tmThk = cSlabthk
     
        ss2.Col = SS2_SLAB_WID     '4
        ss2.Text = cSlabwid
        tmWid = cSlabwid
    
        ss2.Col = SS2_SLAB_LEN     '5
        If ss2.Row = ss2.MaxRows Then
            ss2.Text = cSlabLen - tmTotalLen
            tmTotalLen = tmTotalLen + ss2.Text
        Else
            ss2.Text = cfLen
            tmLen = cfLen
            tmTotalLen = tmTotalLen + cfLen
        End If
        
        ss2.Col = SS2_SLAB_WGT      '6
        ss2.Text = cfWgt
        tmWgt = tmWgt + ss2.VALUE
    
        ss2.Col = SS2_CAL_WGT       '7
        ss2.Text = cfCalWgt
    
        ss2.Col = SS2_SLAB_DATE     '8
        ss2.Text = Format(Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD') FROM DUAL"), "YYYY-MM-DD")
    
        ss2.Col = SS2_SLAB_TIME     '9
        ss2.Text = Format(Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'HH24:MI') FROM DUAL"), "HH:MM")
        
        ss2.Col = SS2_SLAB_LOC     '11
        ss2.Text = cLoc
        
        ss2.Col = SS2_EMP_CD       '12
        ss2.Text = sUserID
        
        ss2.Col = SS2_MOSLAB_NO    '13
        ss2.Text = TXT_SLABNO
        
        ss2.Col = SS2_LAST_STATUS  '14
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
    
    'Call WGT_CAL(row)
    
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
    
    sQuery = "{call AFM2060C.P_ORDCANCEL('" & Trim(TXT_SLABNO.Text) & "','" & sUserID & "',?,?)}"
    
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

'    MDIMain.MenuTool.Buttons(1).Enabled = True                 'Save
'    MDIMain.MenuTool.Buttons(2).Enabled = True                 'Delete
'    MDIMain.MenuTool.Buttons(4).Enabled = True                 'Separator
    MDIMain.MenuTool.Buttons(14).Enabled = True                  'Row Delete
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
    
    'Call opt_prc_status1_click(1)
    opt_prc_status1.VALUE = True
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"))

    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "F-System.INI", Me.Name)
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    txt_plt = "B1"
    Call txt_PLT_KeyUp(0, 0)
    
    txt_thk = 150
    txt_thk_to = 999
    txt_wid = 1800
    txt_wid_to = 9999
    txt_len = 2600
    txt_len_to = 99999
    
    SDT_PROD_DATE_FROM.RawData = Format(Now, "YYYYMM") + "01"
    SDT_PROD_DATE_TO.RawData = Format(Now, "YYYYMMDD")
   
    text_cur_inv_code.Text = "00"
    
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
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "F-System.INI", Me.Name)
    
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
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    txt_act_stlgrd_dec = ""
    TXT_SLABNO.Text = ""
    cbo_cutcnt.ListIndex = 0
    txt_total_len.VALUE = 0
    txt_total_wgt.VALUE = 0
    txt_scrap_wgt.VALUE = 0
    
    Call opt_prc_status1_click(True)
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&

End Sub

Public Sub Form_Pro()
Dim iRow As Integer
Dim sErrMessg As String
Dim sMes As String


'删除分切实绩保存时，弹出对话框提示

    sMes = "你确定要删除分切实绩吗？"
    ss2.Row = ss2.ActiveRow
    ss2.Col = 0
    If ss2.Text = "Delete" Then
    
        If Not Gf_MessConfirm(sMes, "Q") Then Exit Sub

    End If

'modified by guoli at 20090729090900 for 单笔执行时，其中一笔执行失败，另一笔执行成功，导致成功的这笔资料错误
    
    
    If txt_Status.Text = "1" Then
       'ADDED BY GUOLI AT 20100604 FOR 避免重量不守恒保存成功
       If txt_scrap_wgt.VALUE <> 0 Then
          MsgBox "板坯分切前后重量不守恒，请确认分切数据!", vbCritical, "系统提示信息"
          Exit Sub
       End If
        
        If Gf_Sp_Process(M_CN1, Proc_Sc("Sc2"), Mc1, True) Then
           Call Form_Ref
        End If
    ElseIf txt_Status.Text = "2" Then
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
                Else
                    ss2.Row = iRow
                    ss2.Col = 0
                    ss2.Text = ""
                End If
            End If
        Next iRow
        
        M_CN1.CommitTrans
        
        Call Pro_ACB3020P
        
        Call Form_Ref
    End If
    
'    Call Gf_Sp_Process(M_CN1, Proc_Sc("Sc2"), Mc1, True)
''    Call Scrap_Pro   '''COMMENT BY GUOLI AT 20081026
'
'    If opt_prc_status1.Value = True Then
'         Call Form_Ref
'    End If
'    If opt_prc_status2.Value = True Then
'         If ss2.MaxRows < 1 Then
'            Call Form_Ref
'         End If
'    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    
'    MDIMain.MenuTool.Buttons(1).Enabled = True                   'Save
'    MDIMain.MenuTool.Buttons(2).Enabled = True                   'Delete
'    MDIMain.MenuTool.Buttons(4).Enabled = True                   'Separator
'    MDIMain.MenuTool.Buttons(8).Enabled = True                   'Row Insert
'    MDIMain.MenuTool.Buttons(14).Enabled = True                  'Row Delete
'
'    MDIMain.MenuTool.Buttons(5).Enabled = False                  'Save
'    MDIMain.MenuTool.Buttons(7).Enabled = True                   'Delete
'    MDIMain.MenuTool.Buttons(9).Enabled = True                   'Separator
'    MDIMain.MenuTool.Buttons(11).Enabled = False                 'Row Insert
'    MDIMain.MenuTool.Buttons(12).Enabled = False                 'Row Delete
'    MDIMain.MenuTool.Buttons(15).Enabled = False                 'Row Delete

End Sub

Public Sub Form_Ref()

    Dim ForCnt As Integer
    Dim tmWgt As Long
    Dim tmLen As Long
    Dim lRow As Long
    Dim iRow As Integer
    Dim i As Integer
    Dim TIME As String
    

    If Not Gf_Sp_Cls(sc2) Then Exit Sub
    
    If Len(Trim(txt_MOSLAB)) <> 0 Then
        If Len(Trim(txt_MOSLAB)) < 8 Then
           MsgBox "请确认母板坯号"
           txt_MOSLAB.SetFocus
           Exit Sub
        End If
    End If
    
    If Len(Trim(txt_MOSLAB)) <> 8 Then
        If Len(Trim(txt_plt)) = 0 Then
            MsgBox "请确认工厂代码"
            txt_plt.SetFocus
            Exit Sub
        End If
    End If
    
    If txt_len.VALUE <= 0 Then
        MsgBox "请确认从长度"
        txt_len.SetFocus
        Exit Sub
    End If
    
    TXT_SLABNO.Text = ""
    txt_total_len.VALUE = 0
    txt_total_wgt.VALUE = 0
    txt_scrap_wgt.VALUE = 0
    
    Call Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"))
     
    ss1.OperationMode = OperationModeNormal
   
     
    If opt_prc_status2.VALUE = True Then
       Call Gp_Sp_ColLock(ss2, SS2_SLAB_WID, True)     '4
       Call Gp_Sp_ColLock(ss2, SS2_SLAB_LEN, True)     '5
       Call Gp_Sp_ColLock(ss2, SS2_SLAB_DATE, True)    '8
       Call Gp_Sp_ColLock(ss2, SS2_SLAB_TIME, True)    '9
       Call Gp_Sp_ColLock(ss2, SS2_CHG_GRD_RES, True)  '10
    Else
       Call Gp_Sp_ColLock(ss2, SS2_SLAB_WID, False)    '4
       Call Gp_Sp_ColLock(ss2, SS2_SLAB_LEN, False)    '5
       Call Gp_Sp_ColLock(ss2, SS2_SLAB_DATE, False)   '8
       Call Gp_Sp_ColLock(ss2, SS2_SLAB_TIME, False)   '9
       Call Gp_Sp_ColLock(ss2, SS2_CHG_GRD_RES, False) '10
    End If
        
    If opt_prc_status1.VALUE = True Then
        
    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Save
    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False                'Separator
  
  End If
        
        
   TIME = Format(Now, "YYYY-MM-DD")


     For iRow = 1 To ss1.MaxRows

      ss1.Row = iRow
      ss1.Col = 25
        If Mid(ss1.Text, 1, 10) < TIME Then
          For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.ForeColor = &HFF&
          Next
        End If

        If ss1.Text = "" Then
           Exit For
        End If

    Next iRow

        
End Sub

Private Sub opt_prc_status1_click(VALUE As Integer)

    If opt_prc_status1.VALUE = True Then
       cmd_Cancel.Visible = True
    End If
    
    If opt_prc_status1.Tag <> "" Then
       opt_prc_status1.Tag = ""
       Exit Sub
    End If
    
    If Gf_Sp_Cls(sc2) = False Then
        opt_prc_status2.Tag = "A"
        opt_prc_status2.VALUE = True
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
    
'    cbo_cutcnt.Enabled = True
    
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
    
    txt_thk = 150
    txt_thk_to = 999
    txt_wid = 1800
    txt_wid_to = 9999
    txt_len = 2600
    txt_len_to = 99999
    
    cbo_cutcnt.ListIndex = 0
    txt_total_len.VALUE = 0
    txt_total_wgt.VALUE = 0
    txt_scrap_wgt.VALUE = 0
    
    txt_plt = "B1"
    Call txt_PLT_KeyUp(0, 0)
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False                 'Separator

End Sub

Private Sub opt_prc_status2_Click(VALUE As Integer)

    If opt_prc_status2.VALUE = True Then
        cmd_Cancel.Visible = False
    End If
    
    If opt_prc_status2.Tag <> "" Then
       opt_prc_status2.Tag = ""
       Exit Sub
    End If
    
    If Gf_Sp_Cls(sc2) = False Then
        opt_prc_status1.Tag = "A"
        opt_prc_status1.VALUE = True
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
    
    txt_thk = 150
    txt_thk_to = 999
    txt_wid = 1800
    txt_wid_to = 9999
    txt_len = 2600
    txt_len_to = 99999
    
    cbo_cutcnt.Enabled = False
    cbo_cutcnt.ListIndex = 0
    txt_total_len.VALUE = 0
    txt_total_wgt.VALUE = 0
    txt_scrap_wgt.VALUE = 0
    
    txt_plt = "B1"
    Call txt_PLT_KeyUp(0, 0)
    
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
    Dim TIME As String
    Dim iRow As Integer

    With ss1
      .Row = Row
      .Row2 = Row
      .Col = 1
      .Col2 = ss1.MaxCols
      .BlockMode = True
      .BackColor = &HFFFF80
      .BlockMode = False
   End With

        
    For i = 1 To ss1.MaxRows
         If i <> Row Then
              With ss1
                  .Row = i
                  .Row2 = i
                  .Col = 1
                  .Col2 = ss1.MaxCols
                  .BlockMode = True
                  .BackColor = &HFFFFFF
                  .BlockMode = False
              End With
          End If
     Next

    
    
    If Row > 0 Then

        ss1.Col = 1
        ss1.Row = Row
        TXT_SLABNO = ss1.Text
        SCRAP_NO = ss1.Text
        cSlabno = ss1.Text
       
        ss1.Col = 3
        TXT_INGOT_FL = ss1.Text
        
        ss1.Col = 9
        cSlabLen = ss1.VALUE
        ss1.Col = 10
        cSlabWgt = ss1.VALUE
        ss1.Col = 20
        txt_tmpPLT = ss1.VALUE
        ss1.Col = 21
        txt_IST_DATE = ss1.VALUE
        
    End If

    sQuery = "          SELECT MAX(SLAB_NO) "
    sQuery = sQuery & "   FROM NISCO.FP_SLAB "
    sQuery = sQuery & "  WHERE SLAB_NO LIKE '" & Mid(cSlabno, 1, 8) & "%'"
    
    tmpSlabNo = Gf_CodeFind(M_CN1, sQuery)
    If CInt(Mid(tmpSlabNo, 9, 2)) < 30 Then
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
             ss2.Col = SS2_ORD_FL     '16
             If Trim(ss2.Text) = "订单材" Then
                Call Gp_Sp_BlockLock(ss2, SS2_SLAB_WID, SS2_SLAB_LEN, i, i)   '4,5
             End If
        Next i
        Exit Sub
    End If
    
    Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2, Nothing, Mc2("mControl"), False)

    For i = 1 To ss2.MaxRows
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
        
        ss2.Col = SS2_INGOT_FL     '2
        ss2.Text = TXT_INGOT_FL
        
        ss2.Col = SS2_SLAB_THK     '3
        tmThk = ss2.VALUE
        
        ss2.Col = SS2_SLAB_WID     '4
        tmWid = ss2.VALUE
        
        ss2.Col = SS2_SLAB_LEN     '5
        tmLen = ss2.VALUE
            
        ss2.Col = SS2_CAL_WGT      '7
        ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
        
        ss2.Col = SS2_EMP_CD        '12
        ss2.Text = sUserID
        
        ss2.Col = SS2_MOSLAB_NO     '13
        ss2.Text = TXT_SLABNO
        
        ss2.Col = SS2_LAST_STATUS   '14
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
            ss2.Col = SS2_SLAB_LEN     '5
            tmTotalLen = tmTotalLen + ss2.VALUE
            
            ss2.Col = SS2_SLAB_WGT     '6
            tempWgt = tempWgt + ss2.VALUE
        End If

    Next i
    
    If txt_Status = "1" Then
        For i = 1 To ss2.MaxRows
             ss2.Row = i
             ss2.Col = 0
             If UCase(ss2.Text) = "" Then
                ss2.Text = "Input"
             End If
             ss2.Col = SS2_ORD_FL     '16
             If Trim(ss2.Text) = "订单材" Then
                Call Gp_Sp_BlockLock(ss2, SS2_SLAB_WID, SS2_SLAB_LEN, i, i)   '4,5
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

'    If Gf_Sc_Authority(sAuthority, "U") Then Call Gp_Sp_UpdateMake(Proc_Sc("SC2")("Spread"), Mode)

    If Col <> 4 And Col <> 5 Then Exit Sub

    If ChangeMade Then
        Call WGT_CAL(Row)
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

        Exit Sub
    End If
    
End Sub

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String

    If Len(Trim(txt_ord_no.Text)) = txt_ord_no.MaxLength Then
    
        If cbo_ord_item.Text <> "" Then Exit Sub
        
        txt_ord_no.Text = StrConv(txt_ord_no.Text, vbUpperCase)
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(txt_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, cbo_ord_item, sQuery)
        
        'If Combo_ORD_ITEM.ListCount <> 0 Then
        '      Combo_ORD_ITEM.ListIndex = 0
        'End If
    Else
        cbo_ord_item.Clear
    End If
    
End Sub

Private Sub txt_plt_DblClick()

    Call txt_PLT_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_PLT_KeyUp(KeyCode As Integer, Shift As Integer)

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
               adoCmd.Parameters(0).VALUE = "I"
            ElseIf .Text = "Update" Then
               adoCmd.Parameters(0).VALUE = "U"
            ElseIf .Text = "Delete" Then
               adoCmd.Parameters(0).VALUE = "D"
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
                                adoCmd.Parameters(iCol).VALUE = 0
                            Else
                                dTempInt = Sc.Item("Spread").Text
                                adoCmd.Parameters(iCol).VALUE = Trim(STR(dTempInt))
                            End If
                
                       Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                            If Trim(Sc.Item("Spread").VALUE) = "" Then
                                adoCmd.Parameters(iCol).VALUE = ""
                            Else
                                adoCmd.Parameters(iCol).VALUE = Trim(STR(Sc.Item("Spread").VALUE))
                            End If
                            
                       Case SS_CELL_TYPE_DATE
                            If Trim(Sc.Item("Spread").Text) = "" Then
                                adoCmd.Parameters(iCol).VALUE = ""
                            Else
                                adoCmd.Parameters(iCol).VALUE = Mid(Trim(Sc.Item("Spread").Text), 1, 4) & _
                                                                Mid(Trim(Sc.Item("Spread").Text), 6, 2) & _
                                                                Mid(Trim(Sc.Item("Spread").Text), 9, 2)
                            End If
                        
                       Case Else
                            sTemp = Replace(Sc.Item("Spread").Text, "'", "''")
                            adoCmd.Parameters(iCol).VALUE = Trim(sTemp)

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
    adoCmd.Parameters.Append adoCmd.CreateParameter(STR(1), adVariant, adParamOutput)
    
    sQuery = "{call ACB3020P (?)}"
    
    adoCmd.CommandText = sQuery
    adoCmd.Execute , , adExecuteNoRecords
        
    Set adoCmd = Nothing
End Sub

Private Sub text_cur_inv_code_Change()

    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
          text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
          Exit Sub
    Else
          text_cur_inv.Text = ""
    End If
End Sub

Private Sub text_cur_inv_code_DblClick()

    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
    
        DD.rControl.Add Item:=text_cur_inv_code
        DD.rControl.Add Item:=text_cur_inv
        
    
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
       
        If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
            text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
            Exit Sub
        Else
            text_cur_inv.Text = ""
        End If
    End If
End Sub
