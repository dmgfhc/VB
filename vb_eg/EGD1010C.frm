VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form EGD1010C 
   Caption         =   "热处理作业对象指示_EGD1010C"
   ClientHeight    =   7620
   ClientLeft      =   180
   ClientTop       =   2805
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
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
      Left            =   13440
      TabIndex        =   21
      Tag             =   "轧批号"
      Top             =   90
      Width           =   1815
   End
   Begin VB.TextBox txt_loc 
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
      Left            =   9315
      MaxLength       =   10
      TabIndex        =   14
      Top             =   90
      Width           =   1170
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
      Left            =   1380
      MaxLength       =   2
      TabIndex        =   13
      Tag             =   "堆放仓库"
      Top             =   90
      Width           =   465
   End
   Begin VB.TextBox txt_cur_inv_name 
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
      Left            =   1845
      MaxLength       =   50
      TabIndex        =   12
      Tag             =   "工厂"
      Top             =   90
      Width           =   2085
   End
   Begin VB.TextBox txt_stdspec 
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
      Left            =   5700
      TabIndex        =   11
      Top             =   90
      Width           =   1965
   End
   Begin SSSplitter.SSSplitter SSSplitter2 
      Height          =   8340
      Left            =   90
      TabIndex        =   0
      Top             =   840
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   14711
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "EGD1010C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   3885
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15180
         _Version        =   393216
         _ExtentX        =   26776
         _ExtentY        =   6853
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
         MaxCols         =   0
         MaxRows         =   20
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         SpreadDesigner  =   "EGD1010C.frx":0052
      End
      Begin SSSplitter.SSSplitter SSSplitter3 
         Height          =   4395
         Left            =   0
         TabIndex        =   2
         Top             =   3945
         Width           =   15180
         _ExtentX        =   26776
         _ExtentY        =   7752
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   14737632
         PaneTree        =   "EGD1010C.frx":041A
         Begin Threed.SSPanel SSPanel1 
            Height          =   570
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   15180
            _ExtentX        =   26776
            _ExtentY        =   1005
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_cond_name 
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
               Left            =   5610
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   120
               Width           =   5085
            End
            Begin VB.TextBox txt_cond 
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
               Left            =   4950
               Locked          =   -1  'True
               MaxLength       =   4
               TabIndex        =   8
               Top             =   120
               Width           =   645
            End
            Begin VB.TextBox txt_meth 
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
               Left            =   3180
               Locked          =   -1  'True
               MaxLength       =   2
               TabIndex        =   5
               Top             =   120
               Width           =   315
            End
            Begin InDate.ULabel ULabel5 
               Height          =   315
               Left            =   10890
               Top             =   120
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               Caption         =   "热处理作业线"
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
            Begin Threed.SSOption opt_htn_plt1 
               Height          =   285
               Left            =   12330
               TabIndex        =   6
               Top             =   150
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   -2147483641
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
               Caption         =   "1号热处理线"
            End
            Begin Threed.SSOption opt_htn_plt2 
               Height          =   285
               Left            =   13770
               TabIndex        =   7
               Top             =   150
               Width           =   1380
               _ExtentX        =   2434
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
               Caption         =   "2号热处理线"
               Value           =   -1
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Left            =   1800
               Top             =   120
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               Caption         =   "热处理方法"
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
            Begin InDate.ULabel ULabel7 
               Height          =   315
               Left            =   3600
               Top             =   120
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               Caption         =   "热处理条件"
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
            Begin Threed.SSCheck chk_sel 
               Height          =   345
               Left            =   90
               TabIndex        =   10
               Top             =   120
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   609
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "批次取消/选择"
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   3795
            Left            =   0
            TabIndex        =   4
            Top             =   600
            Width           =   15180
            _Version        =   393216
            _ExtentX        =   26776
            _ExtentY        =   6694
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
            MaxCols         =   27
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "EGD1010C.frx":046C
         End
      End
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_thk_fr 
      Height          =   315
      Left            =   1380
      TabIndex        =   15
      Top             =   450
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   0
      Left            =   90
      Top             =   450
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "产品厚度"
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
      Left            =   4440
      Top             =   450
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "产品宽度"
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
      Index           =   1
      Left            =   8040
      Top             =   450
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "产品长度"
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
   Begin CSTextLibCtl.sidbEdit sdb_prod_thk_to 
      Height          =   315
      Left            =   2370
      TabIndex        =   16
      Top             =   450
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_len_fr 
      Height          =   315
      Left            =   9315
      TabIndex        =   17
      Top             =   450
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      NumIntDigits    =   7
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_len_to 
      Height          =   315
      Left            =   10590
      TabIndex        =   18
      Top             =   450
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
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
      NumIntDigits    =   7
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_wid_fr 
      Height          =   315
      Left            =   5700
      TabIndex        =   19
      Top             =   450
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_wid_to 
      Height          =   315
      Left            =   6690
      TabIndex        =   20
      Top             =   450
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   8040
      Top             =   90
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "物料位置"
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
      Left            =   90
      Top             =   90
      Width           =   1245
      _ExtentX        =   2196
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   4440
      Top             =   90
      Width           =   1245
      _ExtentX        =   2196
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
      Left            =   12165
      Top             =   90
      Width           =   1245
      _ExtentX        =   2196
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
End
Attribute VB_Name = "EGD1010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name
'-- Sub_System Name   HTM System
'-- Program Name      HEAT PROCESS WORK INSTRUCTION
'-- Program ID        EGD1010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2010.07.20
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE        EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting
'Public sForm_Fl As String           'Active Form Authority Setting

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

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_cur_inv, "p", "n", "m", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_cur_inv_name, " ", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_Loc, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_prod_thk_fr, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_prod_thk_to, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_prod_wid_fr, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_prod_wid_to, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_prod_len_fr, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_prod_len_to, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_LOT_NO, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(txt_meth, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(txt_cond, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(txt_cond_name, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(txt_Loc, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(sdb_prod_thk_fr, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(sdb_prod_thk_to, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(sdb_prod_wid_fr, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(sdb_prod_wid_to, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(sdb_prod_len_fr, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(sdb_prod_len_to, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(txt_cur_inv, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(TXT_LOT_NO, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
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
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="EGD1010C.P_REFER1", Key:="P-R"
    
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, False)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, False)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="EGD1010C.P_REFER2", Key:="P-R"
    sc2.Add Item:="EGD1010C.P_ONEROW2", Key:="P-O"
    sc2.Add Item:="EGD1010C.P_MODIFY2", Key:="P-M"
    
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").ROW = 0
    sc2.Item("Spread").Text = "◎"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss2, 20, True)
    Call Gp_Sp_ColHidden(ss2, 21, True)
    Call Gp_Sp_ColHidden(ss2, 22, True)
    Call Gp_Sp_ColHidden(ss2, 23, True)
    
    Call Gp_Sp_ColHidden(ss2, ss2.MaxCols - 1, True)
    Call Gp_Sp_ColHidden(ss2, ss2.MaxCols, True)
        
End Sub

Private Sub chk_sel_Click(Value As Integer)

    Dim iRow As Integer
    
    If chk_sel Then
        For iRow = 1 To ss2.MaxRows
            ss2.ROW = iRow
            ss2.Col = 0
            If ss2.Text = "" Then
                ss2.Text = "Update"
                ss2.Col = 2
                If opt_htn_plt1 Then
                    ss2.Text = "1"
                Else
                    ss2.Text = "2"
                End If
                ss2.Col = ss2.MaxCols - 1
                ss2.Text = txt_meth.Text
                ss2.Col = ss2.MaxCols
                ss2.Text = sUserID
                Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, iRow, iRow, , &HFFFF80)
            End If
        Next iRow
    Else
        For iRow = 1 To ss2.MaxRows
            ss2.ROW = iRow
            ss2.Col = 0
            ss2.Text = ""
            ss2.Col = 2
            ss2.Text = ""
            ss2.Col = ss2.MaxCols - 1
            ss2.Text = ""
            ss2.Col = ss2.MaxCols
            ss2.Text = ""
            Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, iRow, iRow)
        Next iRow

    End If
    
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
    Call MenuTool_ReSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Sp_Setting
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
        
    Call Gp_Spl_SizeGet(SSSplitter2, "EG-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "EG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "EG-System.INI", Me.Name)
    
    txt_cur_inv.Text = "WG"
    txt_cur_inv_name.Text = Gf_ComnNameFind(M_CN1, "C0013", Trim(txt_cur_inv.Text), "1")
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter2, "EG-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "EG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "EG-System.INI", Me.Name)
    
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

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, sc2)
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc2) Then
        Call Gf_Sp_Cls(sc1)
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
'        opt_htn_plt1.Value = True
'        opt_htn_plt2.Value = False
'        opt_htn_plt1.ForeColor = &HFF&
'        opt_htn_plt2.ForeColor = &H80000012
'        opt_htn_plt1.Enabled = True
'        opt_htn_plt2.Enabled = True
        chk_sel.Value = ssCBUnchecked
    End If
    
End Sub

Public Sub Form_Ref()

    Dim sQuery1 As String   'Header Display
    
    If txt_cur_inv.Text <> "00" And txt_cur_inv.Text <> "WD" And txt_cur_inv.Text <> "WG" Then
        Call Gp_MsgBoxDisplay("只堆放仓库 --> '00' / 'WD' / 'WG' ", "I")
        Exit Sub
    End If
    
    If Not Gf_Sp_Cls(sc2) Then Exit Sub
    
    'Header Display
    sQuery1 = "          SELECT  CD "
    sQuery1 = sQuery1 + "  FROM  ZP_CD "
    sQuery1 = sQuery1 + " WHERE  CD_MANA_NO = 'Q0073' "
    sQuery1 = sQuery1 + " ORDER  BY CD "
    
    'Header Display
    Call Sp_Header_Refer1(ss1, sQuery1)      'Header Display
            
    If Sp_Display(M_CN1, sc1.Item("Spread"), Gf_Ms_MakeQuery(sc1.Item("P-R"), "R", Mc1("pControl")), _
                                    sc1.Item("pColumn"), False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss1.OperationMode = OperationModeNormal
        Call Gp_Ms_ControlLock(Mc1("lControl"), True)
        Call Gp_Ms_Cls(Mc2("rControl"))
        
'        If txt_cur_inv.Text = "00" Then
'            opt_htn_plt1.Value = True
'            opt_htn_plt2.Value = False
'            opt_htn_plt1.ForeColor = &HFF&
'            opt_htn_plt2.ForeColor = &H80000012
'            opt_htn_plt2.Enabled = False
'        Else
'            opt_htn_plt1.Value = False
'            opt_htn_plt2.Value = True
'            opt_htn_plt1.ForeColor = &H80000012
'            opt_htn_plt2.ForeColor = &HFF&
'            opt_htn_plt1.Enabled = False
'        End If
        chk_sel.Value = ssCBUnchecked
    End If
            
End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, sc2, Mc2) Then
        Call Form_Ref
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

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, sc2("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()

End Sub

Private Sub opt_htn_plt1_Click(Value As Integer)

    If opt_htn_plt1 Then
        opt_htn_plt1.ForeColor = &HFF&
        opt_htn_plt2.ForeColor = &H80000012
    End If
    
End Sub

Private Sub opt_htn_plt2_Click(Value As Integer)

     If opt_htn_plt2 Then
        opt_htn_plt2.ForeColor = &HFF&
        opt_htn_plt1.ForeColor = &H80000012
    End If
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    If ROW < 1 Or Col < 1 Then Exit Sub
    
    ss1.Col = Col
    ss1.ROW = ROW
    
    If ss1.Text = "" Then Exit Sub
    
    If Col = 1 Then
        txt_meth.Text = "SB"
    Else
        ss1.Col = Col
        ss1.ROW = SpreadHeader + (ss1.ColHeaderRows - 1)
        txt_meth.Text = ss1.Text
    End If
    
    ss1.ROW = ROW
    ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 2)
    txt_cond.Text = ss1.Text
    ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 1)
    txt_cond_name.Text = ss1.Text
    
    ss1.ROW = ROW
    
    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, , , False)
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call MenuTool_ReSet
    ss2.OperationMode = OperationModeNormal
    
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)
    
    Dim lCol As Integer
    Dim lrCol As Integer
    Dim sFind As Boolean
    
    Call Gp_Sp_Sort(sc2.Item("Spread"), Col, ROW)
    
    lBlkcol1 = Col
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If ROW < 1 Then Exit Sub
    
    ss2.Col = 0
    ss2.ROW = ROW
    
    If ss2.Text = "" Then
        ss2.Text = "Update"
        ss2.Col = 2
        If opt_htn_plt1 Then
            ss2.Text = "1"
        End If
        If opt_htn_plt2 And txt_meth = "SB" Then
            ss2.Text = "1"
        End If
       
        If opt_htn_plt2 And txt_meth <> "SB" Then
            ss2.Text = "2"
        End If
        
        ss2.Col = ss2.MaxCols - 1
        ss2.Text = txt_meth.Text
        ss2.Col = ss2.MaxCols
        ss2.Text = sUserID
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, ROW, ROW, , &HFFFF80)
    Else
        ss2.Text = ""
        ss2.Col = 2
        ss2.Text = ""
        ss2.Col = ss2.MaxCols - 1
        ss2.Text = ""
        ss2.Col = ss2.MaxCols
        ss2.Text = ""
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, ROW, ROW)
    End If

End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If ROW > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub txt_cur_inv_Change()
If Len(Trim(txt_cur_inv.Text)) = txt_cur_inv.MaxLength Then
   txt_cur_inv_name.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv.Text, 2)
End If
End Sub

Private Sub txt_stdspec_DblClick()

    Call txt_stdspec_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stdspec_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
       
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)
        
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

Private Function Sp_Display(Conn As ADODB.Connection, sPname As vaSpread, sQuery As String, _
                              Optional lColumn As Variant = Nothing, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim icount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Sp_Display = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Display = True
        
        .ReDraw = False
        .MaxRows = 0: icount = 0
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
                
            Sp_Display = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        .MaxRows = UBound(ArrayRecords, 2) + 2
        
        .ROW = .MaxRows
        .RowHidden = True
        
        For iRowCount = 0 To .MaxRows - 2
        
            .ROW = iRowCount + 1
            
            If Trim(ArrayRecords(0, iRowCount)) = "SB" Then
                .Col = SpreadHeader + (.RowHeaderCols - 2)
                .Text = Trim(ArrayRecords(1, iRowCount))
                .Col = SpreadHeader + (.RowHeaderCols - 1)
                .Text = Trim(ArrayRecords(2, iRowCount))
                .Col = 1
                .Text = Trim(ArrayRecords(3, iRowCount))
                .ROW = .MaxRows
                .Text = "Y"
            End If
            
            For iColcount = 2 To .MaxCols
                .Col = iColcount
                .ROW = SpreadHeader + (.ColHeaderRows - 1)
                If .Text = Trim(ArrayRecords(0, iRowCount)) Then
                    .ROW = iRowCount + 1
                    .Col = SpreadHeader + (.RowHeaderCols - 2)
                    .Text = Trim(ArrayRecords(1, iRowCount))
                    .Col = SpreadHeader + (.RowHeaderCols - 1)
                    .Text = Trim(ArrayRecords(2, iRowCount))
                    .Col = iColcount
                    .Text = Trim(ArrayRecords(3, iRowCount))
                    .ROW = .MaxRows
                    .Text = "Y"
                End If
            Next iColcount
            
        Next iRowCount
        
        For iColcount = 1 To .MaxCols
            .ROW = .MaxRows
            .Col = iColcount
            If .Text <> "Y" Then Call Gp_Sp_ColHidden(ss1, iColcount, True)
        Next iColcount
        
        .BlockMode = True
        .ROW = SpreadHeader + (.ColHeaderRows - 2)
        .Col = 1
        .Row2 = SpreadHeader + (.ColHeaderRows - 2)
        .Col2 = .MaxCols
        .RowMerge = MergeAlways
        '.ColMerge = MergeAlways
        .BlockMode = False
            
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Display = False
    Call Gp_MsgBoxDisplay("Sp_Display Error : " & sQuery)
    Screen.MousePointer = vbDefault
    
End Function

Private Function Sp_Header_Refer1(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay1_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim iColCnt As Integer
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Header_Refer1 = True
        
        .ReDraw = False
        .MaxRows = 0:  .MaxCols = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Header_Refer1 = False
            '.ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing
        
        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = (UBound(ArrayRecords, 2) + 2)
            
            .Col = 1
            .ROW = SpreadHeader + (.ColHeaderRows - 2)
            .Text = "热处理方法"
            .ROW = SpreadHeader + (.ColHeaderRows - 1)
            .Text = "抛丸"
            .ColWidth(1) = 14
            
            For iCol = 1 To .MaxCols - 1
            
                .Col = iCol + 1
                
                .ROW = SpreadHeader + (.ColHeaderRows - 2)
                .Text = "热处理方法"
                .ROW = SpreadHeader + (.ColHeaderRows - 1)
                If VarType(ArrayRecords(0, iCol - 1)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCol - 1))
                End If
                
                .ColWidth(iCol + 1) = 13
                
            Next iCol
            
            .Col = 1: .Col2 = -1
            .ROW = 1: .Row2 = -1
            .BlockMode = True
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            .BlockMode = False
            
        End If
        
        .BlockMode = True
        .ROW = SpreadHeader + (.ColHeaderRows - 2)
        .Col = 2
        .Row2 = SpreadHeader + (.ColHeaderRows - 2)
        .Col2 = .MaxCols
        .RowMerge = MergeAlways
        '.ColMerge = MergeAlways
        .BlockMode = False
        
        .ReDraw = True
        .Refresh
        
        Screen.MousePointer = vbDefault
        
    End With
        
    Exit Function

SpreadDisplay1_Error:
    
    Set AdoRs = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer1 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay1_Error : " & Error)
    
End Function

Public Sub Sp_Setting()

    ss1.ColWidth(SpreadHeader + (ss1.RowHeaderCols - 2)) = 5
    ss1.ColWidth(SpreadHeader + (ss1.RowHeaderCols - 1)) = 40

End Sub

