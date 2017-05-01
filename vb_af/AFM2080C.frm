VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AFM2080C 
   Caption         =   "板坯焊接作业界面_AFM2080C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
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
      Left            =   6465
      TabIndex        =   30
      Top             =   450
      Width           =   1530
   End
   Begin VB.TextBox txt_SLAB 
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
      Left            =   9480
      MaxLength       =   10
      TabIndex        =   25
      Top             =   60
      Width           =   1770
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
      Left            =   5070
      MaxLength       =   11
      TabIndex        =   24
      Top             =   450
      Width           =   1410
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
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   22
      Top             =   450
      Width           =   420
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
      Left            =   3240
      MaxLength       =   11
      TabIndex        =   11
      Top             =   450
      Visible         =   0   'False
      Width           =   195
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
      Left            =   1620
      MaxLength       =   11
      TabIndex        =   10
      Top             =   450
      Width           =   1410
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
      Left            =   14460
      MaxLength       =   20
      TabIndex        =   9
      Top             =   5970
      Visible         =   0   'False
      Width           =   600
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
      Left            =   9960
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   450
      Width           =   1320
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
      Left            =   9465
      MaxLength       =   2
      TabIndex        =   0
      Top             =   450
      Width           =   495
   End
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   630
      Left            =   13740
      TabIndex        =   2
      Top             =   510
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
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7845
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   13838
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AFM2080C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   3195
         Left            =   0
         TabIndex        =   4
         Top             =   4650
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   5636
         _Version        =   196609
         SplitterBarWidth=   3
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   16761087
         PaneTree        =   "AFM2080C.frx":0052
         Begin Threed.SSPanel SSPanel1 
            Height          =   555
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   15105
            _ExtentX        =   26644
            _ExtentY        =   979
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
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
               Left            =   13770
               MaxLength       =   20
               TabIndex        =   29
               Top             =   120
               Visible         =   0   'False
               Width           =   870
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
               Left            =   13380
               MaxLength       =   10
               TabIndex        =   6
               Top             =   120
               Visible         =   0   'False
               Width           =   990
            End
            Begin InDate.ULabel ULabel10 
               Height          =   315
               Left            =   2400
               Top             =   120
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
               Caption         =   "总厚度"
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
            Begin InDate.ULabel ULabel11 
               Height          =   315
               Left            =   10980
               Top             =   120
               Width           =   1185
               _ExtentX        =   2090
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
            Begin CSTextLibCtl.sidbEdit txt_total_thk 
               Height          =   315
               Left            =   3600
               TabIndex        =   18
               Top             =   120
               Width           =   1185
               _Version        =   262145
               _ExtentX        =   2090
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
            Begin CSTextLibCtl.sidbEdit txt_total_wgt 
               Height          =   315
               Left            =   12180
               TabIndex        =   19
               Top             =   120
               Width           =   1185
               _Version        =   262145
               _ExtentX        =   2090
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
               NumIntDigits    =   4
               MaxValue        =   20
               MinValue        =   10
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel4 
               Height          =   315
               Left            =   150
               Top             =   120
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   556
               Caption         =   "焊接块数"
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
            Begin CSTextLibCtl.sidbEdit txt_cnt 
               Height          =   315
               Left            =   1200
               TabIndex        =   20
               Top             =   120
               Width           =   735
               _Version        =   262145
               _ExtentX        =   1296
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
            Begin InDate.ULabel ULabel12 
               Height          =   315
               Left            =   5250
               Top             =   120
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
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_total_wid 
               Height          =   315
               Left            =   6450
               TabIndex        =   27
               Top             =   120
               Width           =   1185
               _Version        =   262145
               _ExtentX        =   2090
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
            Begin InDate.ULabel ULabel16 
               Height          =   315
               Left            =   8100
               Top             =   120
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   556
               Caption         =   "长度"
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
               Left            =   9300
               TabIndex        =   28
               Top             =   120
               Width           =   1185
               _Version        =   262145
               _ExtentX        =   2090
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
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   2595
            Left            =   0
            TabIndex        =   7
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
            MaxCols         =   7
            MaxRows         =   2
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AFM2080C.frx":00A4
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   4590
         Left            =   0
         TabIndex        =   8
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
         MaxCols         =   20
         MaxRows         =   2
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AFM2080C.frx":0612
      End
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   8340
      Top             =   840
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_len 
      Height          =   315
      Left            =   9465
      TabIndex        =   12
      Top             =   840
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
   Begin CSTextLibCtl.sidbEdit txt_thk_to 
      Height          =   315
      Left            =   2115
      TabIndex        =   13
      Top             =   840
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
      Left            =   5985
      TabIndex        =   14
      Top             =   840
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
      Left            =   10380
      TabIndex        =   15
      Top             =   840
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
   Begin Threed.SSOption opt_prc_status2 
      Height          =   375
      Left            =   2370
      TabIndex        =   16
      Top             =   30
      Width           =   1335
      _ExtentX        =   2355
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
   Begin InDate.UDate SDT_PROD_DATE_FROM 
      Height          =   315
      Left            =   5070
      TabIndex        =   17
      Tag             =   "起始日期"
      Top             =   60
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
   Begin CSTextLibCtl.sidbEdit txt_thk 
      Height          =   315
      Left            =   1200
      TabIndex        =   21
      Top             =   840
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
   Begin Threed.SSOption opt_prc_status1 
      Height          =   375
      Left            =   1200
      TabIndex        =   23
      Top             =   30
      Width           =   1155
      _ExtentX        =   2037
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
      Caption         =   "板坯焊接"
   End
   Begin CSTextLibCtl.sidbEdit txt_wid 
      Height          =   315
      Left            =   5070
      TabIndex        =   26
      Top             =   840
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   60
      Top             =   60
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
      Left            =   60
      Top             =   840
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   60
      Top             =   450
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   8340
      Top             =   60
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "板坯号"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   3930
      Top             =   840
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
   Begin InDate.ULabel ULabel15 
      Height          =   315
      Left            =   8340
      Top             =   450
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
         Size            =   9.75
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
      Left            =   3930
      Top             =   450
      Width           =   1125
      _ExtentX        =   1984
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
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   3930
      Top             =   60
      Width           =   1125
      _ExtentX        =   1984
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
   Begin InDate.UDate SDT_PROD_DATE_TO 
      Height          =   315
      Left            =   6525
      TabIndex        =   31
      Tag             =   "起始日期"
      Top             =   60
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
End
Attribute VB_Name = "AFM2080C"
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
'-- Program Name      板坯焊接作业界面
'-- Program ID        AFM2080C
'-- Designer          ZHANG JIN BO
'-- Coder             ZHANG JIN BO
'-- Date              2013.2.16
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

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_Status, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_act_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_SLAB, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_plt_dec, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_thk, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_wid, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_len, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_len_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    
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
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
 
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AFM2080C.P_REFER", Key:="P-R"
    sc1.Add Item:="AFM2080C.P_MODIFY", Key:="P-M"
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
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AFM2080C.P_REFER1", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"

  
    Call Gp_Sp_ColHidden(ss1, 19, True)
    Call Gp_Sp_ColHidden(ss1, 20, True)
  
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
   
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
    
    
    'Call opt_prc_status1_click(1)
    opt_prc_status1.VALUE = True
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"))

    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "F-System.INI", Me.Name)
    

    txt_cnt.ForeColor = &H0&
    txt_total_thk.ForeColor = &H0&
    txt_total_wid.ForeColor = &H0&
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&

    
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
    
     Call Gp_Sp_ColHidden(ss1, 2, True)
   
   
   
     text_cur_inv_code.Text = "00"
     
'    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Save
'    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Delete
'    MDIMain.MenuTool.Buttons(9).Enabled = False                'Separator
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
       
    txt_cnt.VALUE = 0
    txt_total_thk.VALUE = 0
    txt_total_wid.VALUE = 0
    txt_total_len.VALUE = 0
    txt_total_wgt.VALUE = 0
   
    Call opt_prc_status1_click(True)
    
    txt_cnt.ForeColor = &H0&
    txt_total_thk.ForeColor = &H0&
    txt_total_wid.ForeColor = &H0&
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&

   

End Sub

Public Sub Form_Pro()
Dim iRow As Integer
Dim sErrMessg As String
Dim sMes As String

    sMes = "你确定要删除焊接实绩吗？"
    ss1.Row = ss1.ActiveRow
    ss1.Col = 0
    If ss1.Text = "Delete" Then

        If Not Gf_MessConfirm(sMes, "Q") Then Exit Sub

    End If
    
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1, True) Then
       Call Form_Ref
    End If

    
    Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    

End Sub

Public Sub Form_Ref()

    Dim ForCnt As Integer
    Dim tmWgt As Long
    Dim tmLen As Long
    

    If Not Gf_Sp_Cls(sc2) Then Exit Sub
    
    If Len(Trim(txt_SLAB)) <> 0 Then
        If Len(Trim(txt_SLAB)) < 8 Then
           MsgBox "请确认母板坯号"
           txt_SLAB.SetFocus
           Exit Sub
        End If
    End If
    
    If Len(Trim(txt_SLAB)) <> 8 Then
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
    txt_cnt.VALUE = 0
    txt_total_thk.VALUE = 0
    txt_total_wid.VALUE = 0
    txt_total_len.VALUE = 0
    txt_total_wgt.VALUE = 0
    
    Call Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"))
    ss1.OperationMode = OperationModeNormal
  

        
    If opt_prc_status1.VALUE = True Then
        
    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Save
    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False                 'Separator
  
  End If
        
        
        
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


    opt_prc_status1.ForeColor = &HFF&
    opt_prc_status2.ForeColor = &H80000011

    Call Gf_Sp_Cls(sc1)
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    txt_Status.Text = "1"

    txt_act_stlgrd_dec = ""
    txt_SLAB.Text = ""
    TXT_SLABNO.Text = ""

    txt_thk = 150
    txt_thk_to = 999
    txt_wid = 1800
    txt_wid_to = 9999
    txt_len = 2600
    txt_len_to = 99999
    
    txt_cnt.VALUE = 0
    txt_total_thk.VALUE = 0
    txt_total_wid.VALUE = 0
    txt_total_len.VALUE = 0
    txt_total_wgt.VALUE = 0

    txt_plt = "B1"
    Call txt_PLT_KeyUp(0, 0)
    
    txt_cnt.ForeColor = &H0&
    txt_total_thk.ForeColor = &H0&
    txt_total_wid.ForeColor = &H0&
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
   

    Call Gp_Ms_ControlLock(Mc1("pControl"), False)

    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(9).Enabled = True                  'Separator
    
    Call Gp_Sp_ColHidden(ss1, 2, True)
    Call Gp_Sp_ColHidden(ss1, 19, True)
    Call Gp_Sp_ColHidden(ss1, 20, True)
    
    SDT_PROD_DATE_FROM.RawData = Format(Now, "YYYYMM") + "01"
    SDT_PROD_DATE_TO.RawData = Format(Now, "YYYYMMDD")

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

    opt_prc_status2.ForeColor = &HFF&
    opt_prc_status1.ForeColor = &H80000011

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    txt_Status.Text = "2"

    txt_act_stlgrd_dec = ""
    txt_SLAB.Text = ""
    TXT_SLABNO.Text = ""

    txt_thk = 150
    txt_thk_to = 999
    txt_wid = 1800
    txt_wid_to = 9999
    txt_len = 2600
    txt_len_to = 99999

    
    txt_cnt.VALUE = 0
    txt_total_thk.VALUE = 0
    txt_total_wid.VALUE = 0
    txt_total_len.VALUE = 0
    txt_total_wgt.VALUE = 0
 
    txt_plt = "B1"
    Call txt_PLT_KeyUp(0, 0)
   
    txt_cnt.ForeColor = &H0&
    txt_total_thk.ForeColor = &H0&
    txt_total_wid.ForeColor = &H0&
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
   

    Call Gp_Ms_ControlLock(Mc1("pControl"), False)

    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(9).Enabled = True                  'Separator

    Call Gp_Sp_ColHidden(ss1, 2, False)
    Call Gp_Sp_ColHidden(ss1, 19, False)
    Call Gp_Sp_ColHidden(ss1, 20, False)
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

     Dim i  As Integer
     Dim iRow As Integer
     Dim ForCnt As Integer

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
        ss1.Col = 3
        txt_cnt = ss1.VALUE
        ss1.Col = 4
        txt_tmpPLT = ss1.VALUE
        ss1.Col = 6
        txt_total_thk = ss1.VALUE
        ss1.Col = 7
        txt_total_wid = ss1.VALUE
        ss1.Col = 8
        txt_total_len = ss1.VALUE
        ss1.Col = 9
        txt_total_wgt = ss1.VALUE
        ss1.Col = 17
        txt_IST_DATE = ss1.VALUE
        
        
        
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
    
     Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2, Nothing, Mc2("mControl"), False)
    
    If txt_Status = "1" Then
    
    
        For i = 1 To ss1.MaxRows
              
             If ss1.MaxRows = 1 Then
                 
                 ss1.Row = i
                 ss1.Col = 0
                   If ss1.Text = "Input" Then
                      ss1.Text = ""
                      ss1.Col = 20
                      ss1.Text = ""
        
                   Else
                        ss1.Text = "Input"
                        ss1.Col = 20
                        ss1.Text = sUserID
                   End If
            
             Else
                
                 ss1.Row = i
                 ss1.Col = 0
                    If ss1.Row = Row Then
                       ss1.Text = "Input"
                       ss1.Col = 20
                       ss1.Text = sUserID
                    Else
                       ss1.Col = 0
                       ss1.Text = ""
                       ss1.Col = 20
                       ss1.Text = ""
                   End If
             End If

        Next i
       
  End If
    
   If txt_Status = "2" Then
    
    
        For i = 1 To ss1.MaxRows
              
             If ss1.MaxRows = 1 Then
                 
                 ss1.Row = i
                 ss1.Col = 0
                   If ss1.Text = "Delete" Then
                      ss1.Text = ""
                   Else
                        ss1.Text = "Delete"
                      
                   End If
            
             Else
                
                 ss1.Row = i
                 ss1.Col = 0
                    If ss1.Row = Row Then
                       ss1.Text = "Delete"
                     
                    Else
                       ss1.Col = 0
                       ss1.Text = ""
                 
                   End If
             End If

        Next i
       
  End If
    
           
End Sub
Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))

End Sub
Public Sub Spread_Del()

  Dim i  As Integer
  Dim iRow As Integer

  Call Gp_Sp_Del(Proc_Sc("SC"))


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

    If Col <> 3 And Col <> 4 Then Exit Sub

'    If ChangeMade Then
'        Call WGT_CAL(Row)
'    End If

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

    sQuery = "{call AFM2080C.P_ORDCANCEL('" & Trim(TXT_SLABNO.Text) & "','" & sUserID & "',?,?)}"

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
