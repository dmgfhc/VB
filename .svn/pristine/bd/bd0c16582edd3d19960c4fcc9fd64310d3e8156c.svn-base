VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AEB2060C 
   Caption         =   "板坯设计结果修改_AEB2060C"
   ClientHeight    =   9225
   ClientLeft      =   210
   ClientTop       =   2265
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_prod_cd_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   9030
      MaxLength       =   40
      TabIndex        =   4
      Tag             =   "产品"
      Top             =   80
      Width           =   1860
   End
   Begin VB.TextBox txt_prod_cd 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   8565
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "产品"
      Top             =   80
      Width           =   465
   End
   Begin VB.TextBox txt_hcr_fl 
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
      Height          =   310
      Left            =   8565
      MaxLength       =   1
      TabIndex        =   7
      Tag             =   "HCR/CCR"
      Top             =   470
      Width           =   420
   End
   Begin VB.TextBox txt_hcr_fl_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   9180
      MaxLength       =   50
      TabIndex        =   8
      Tag             =   "PLT"
      Top             =   470
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.ComboBox cbo_order_cnt 
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
      ItemData        =   "AEB2060C.frx":0000
      Left            =   5685
      List            =   "AEB2060C.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "多订单板坯"
      Top             =   465
      Width           =   1410
   End
   Begin VB.TextBox txt_stlgrd 
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
      Left            =   1485
      MaxLength       =   11
      TabIndex        =   6
      Tag             =   "钢种"
      Top             =   465
      Width           =   1275
   End
   Begin VB.TextBox txt_prc_line 
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
      Height          =   310
      Left            =   5685
      MaxLength       =   1
      TabIndex        =   2
      Tag             =   "连铸机号"
      Top             =   80
      Width           =   420
   End
   Begin VB.TextBox txt_plt_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1950
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   80
      Width           =   2190
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
      Height          =   310
      Left            =   1485
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   80
      Width           =   465
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   7305
      Top             =   465
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "HCR/CCR"
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
      Left            =   225
      Top             =   465
      Width           =   1230
      _ExtentX        =   2170
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   4425
      Top             =   465
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "多订单板坯"
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
      Left            =   225
      Top             =   75
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "工厂"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   4425
      Top             =   75
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "连铸机号"
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   7305
      Top             =   75
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "产品"
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
   Begin Threed.SSCommand cmd_slab 
      Height          =   495
      Left            =   13830
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   150
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   873
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   12583104
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "板坯设计"
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmd_charge 
      Height          =   405
      Left            =   15720
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1650
      Visible         =   0   'False
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "炉次编制"
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   11475
      Top             =   465
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "错误数"
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
   Begin CSTextLibCtl.sidbEdit sdb_err_cnt 
      Height          =   315
      Left            =   12645
      TabIndex        =   11
      Top             =   465
      Width           =   825
      _Version        =   262145
      _ExtentX        =   1455
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
      DataProperty    =   2
      ReadOnly        =   -1  'True
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin Threed.SSCommand cmd_add 
      Height          =   405
      Left            =   15720
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16576
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "追加余材板坯"
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Index           =   0
      Left            =   11475
      Top             =   75
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "总件数"
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
   Begin CSTextLibCtl.sidbEdit sdb_tot_cnt 
      Height          =   315
      Left            =   12645
      TabIndex        =   13
      Top             =   75
      Width           =   825
      _Version        =   262145
      _ExtentX        =   1455
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
      DataProperty    =   2
      ReadOnly        =   -1  'True
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8340
      Left            =   60
      TabIndex        =   14
      Top             =   840
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   14711
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "AEB2060C.frx":0004
      Begin Threed.SSPanel SSPanel1 
         Height          =   915
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   15120
         _ExtentX        =   26670
         _ExtentY        =   1614
         _Version        =   196609
         BackColor       =   14737918
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txt_stlgrdR_nm 
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
            Left            =   2850
            TabIndex        =   17
            Top             =   90
            Width           =   2175
         End
         Begin VB.TextBox txt_stlgrdR 
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
            Left            =   1425
            MaxLength       =   11
            TabIndex        =   16
            Top             =   90
            Width           =   1425
         End
         Begin Threed.SSOption opt_hcr 
            Height          =   255
            Left            =   11850
            TabIndex        =   18
            Top             =   300
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   255
            BackColor       =   14737918
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "HCR"
            Value           =   -1
         End
         Begin CSTextLibCtl.sidbEdit sdb_slab_len_fr 
            Height          =   315
            Left            =   8925
            TabIndex        =   19
            Top             =   480
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            MaxValue        =   9999999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSCommand cmd_change 
            Height          =   495
            Left            =   13770
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   180
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   873
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   8421376
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "批次变更"
            BevelWidth      =   3
         End
         Begin CSTextLibCtl.sidbEdit sdb_slab_len_to 
            Height          =   315
            Left            =   10020
            TabIndex        =   21
            Top             =   480
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            MaxValue        =   9999999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_slab_thk_fr 
            Height          =   315
            Left            =   1425
            TabIndex        =   22
            Top             =   480
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            MaxValue        =   99999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_slab_thk_to 
            Height          =   315
            Left            =   2520
            TabIndex        =   23
            Top             =   480
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            MaxValue        =   9999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_slab_wid_fr 
            Height          =   315
            Left            =   5195
            TabIndex        =   24
            Top             =   480
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            MaxValue        =   9999999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_slab_wid_to 
            Height          =   315
            Left            =   6285
            TabIndex        =   25
            Top             =   480
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            MaxValue        =   9999999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Index           =   0
            Left            =   7680
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "板坯长度"
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
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Index           =   1
            Left            =   160
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "板坯厚度"
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
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Index           =   2
            Left            =   3930
            Top             =   480
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "板坯宽度"
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
         Begin Threed.SSOption opt_ccr 
            Height          =   255
            Left            =   12750
            TabIndex        =   26
            Top             =   300
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   196609
            Font3D          =   1
            BackColor       =   14737918
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "CCR"
         End
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   160
            Top             =   90
            Width           =   1230
            _ExtentX        =   2170
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
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   7395
         Left            =   0
         TabIndex        =   27
         Top             =   945
         Width           =   15120
         _Version        =   393216
         _ExtentX        =   26670
         _ExtentY        =   13044
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
         MaxCols         =   48
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AEB2060C.frx":0056
      End
   End
End
Attribute VB_Name = "AEB2060C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       DAILY SCHEDULE
'-- Sub_System Name
'-- Program Name
'-- Program ID        AEB2060C
'-- Document No       Q-00-0010(Specification)
'-- Designer          jianing
'-- Coder             jianing
'-- Date              2003.6.19
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
Public P_SALB_EDT_NO  As String
Public Complete As Boolean

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




Const SS1_SLAB_EDT_SEQ = 1
Const SS1_STLGRD = 10
Const SS1_HEAT_EDT_SEQ = 3
Const SS1_HEAT_SLAB_SEQ = 5

Const SS1_SMS_PLT = 6
Const SS1_MILL_PLT = 7
Const SS1_SMS_PRC_LINE = 8
Const SS1_CCM_PRC_LINE = 9
'Const SS1_STLGRD = 10

Const SS1_THK = 13
Const SS1_WID = 14
Const SS1_LEN = 15
Const SS1_WGT = 16
Const SS1_PLATE_WGT = 17
Const SS1_SHOUDE = 18
Const SS1_PLATECNT = 26
Const SS1_HCR = 27
Const SS1_ORD_FL = 28
Const SS1_FL_CNT = 29
Const SS1_REQ_NO = 46



Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_prc_line, "p", "n", "m", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_prod_cd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_prod_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(cbo_order_cnt, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      
      Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_hcr_fl, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_hcr_fl_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_err_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_tot_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  BELOW EDIT ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(SS1, 1, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     
     Call Gp_Sp_Collection(SS1, 6, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     
    Call Gp_Sp_Collection(SS1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 13, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 14, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 27, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '余材数
    Call Gp_Sp_Collection(SS1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 38, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 39, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 40, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 41, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 42, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 43, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 44, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 45, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 46, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  'useid
    Call Gp_Sp_Collection(SS1, 47, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 48, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)


    'Spread_Collection
    
    Sc1.Add Item:=SS1, Key:="Spread"
    Sc1.Add Item:="AEB2060C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AEB2060C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AEB2060C.P_ONEROW", Key:="P-O"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  EDIT  End      ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=SS1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"

    cbo_order_cnt.AddItem "Type"
    cbo_order_cnt.AddItem "Y"
    cbo_order_cnt.AddItem "N"
         
    cbo_order_cnt.ListIndex = 0
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Sc1.Item("Spread").Col = 0
    Sc1.Item("Spread").Row = 0
    Sc1.Item("Spread").Text = "◎"
    
    Call Gp_Sp_ColHidden(SS1, 1, True)
    Call Gp_Sp_ColHidden(SS1, 42, True)
    
    Call Gp_Sp_ColHidden(SS1, SS1_SMS_PLT, True)
    Call Gp_Sp_ColHidden(SS1, SS1_MILL_PLT, True)
    Call Gp_Sp_ColHidden(SS1, SS1_SMS_PRC_LINE, True)
    Call Gp_Sp_ColHidden(SS1, SS1_CCM_PRC_LINE, True)
    Call Gp_Sp_ColHidden(SS1, SS1_STLGRD, True)

End Sub

Private Sub cmd_add_Click()

    Complete = False

    If SS1.MaxRows = 0 Then Exit Sub
    If SS1.ActiveRow <= 0 Then Exit Sub
    
    SS1.Col = 4
    SS1.Row = SS1.ActiveRow
    
    If SS1.Text = "" Then Exit Sub
    
'    ss1.Col = 4
'    If ss1.Text = "" Then Exit Sub
    
'    If Val(ss1.Text) = 0 Then Exit Sub
    
    Load Slab_Add
    
    SS1.Col = SS1_SLAB_EDT_SEQ
    Slab_Add.P_SLAB_EDT_SEQ = SS1.Text               'SLAB_EDT_SEQ
    
    SS1.Col = SS1_STLGRD
    Slab_Add.P_STLGRD = SS1.Text                     'STLGRD
    
    SS1.Col = SS1_HEAT_EDT_SEQ
    Slab_Add.sdb_heat_edt_seq.Value = SS1.Value      'HEAT_EDT_SEQ
    
    SS1.Col = SS1_HEAT_SLAB_SEQ
    Slab_Add.txt_heat_slab_seq.Text = SS1.Text       'HEAT_SLAB_SEQ
    
    SS1.Col = SS1_THK
    Slab_Add.sdb_thk.Value = SS1.Value               'THK
    SS1.Col = SS1_WID
    Slab_Add.sdb_wid.Value = SS1.Value               'WID
    SS1.Col = SS1_LEN
    Slab_Add.sdb_len.Value = SS1.Value               'LEN
    SS1.Col = SS1_WGT
    Slab_Add.sdb_wgt.Value = SS1.Value               'WGT
    
    Slab_Add.Show 1
    
    If Complete Then
        Call Form_Ref
    End If

End Sub

Private Sub cmd_change_Click()

    Dim lRow As Integer
    Dim sStlgrd As String
    Dim dThk As Double
    Dim dWid As Double
    Dim dLen As Double
    
    For lRow = 1 To SS1.MaxRows
        
        SS1.Row = lRow
        SS1.Col = SS1_STLGRD
        sStlgrd = SS1.Text
        
        SS1.Col = SS1_REQ_NO    'C3   Req_no
        
        If sStlgrd <> "" And SS1.Text = "" Then
        
            SS1.Col = SS1_THK
            dThk = SS1.Value
            SS1.Col = SS1_WID
            dWid = SS1.Value
            SS1.Col = SS1_LEN
            dLen = SS1.Value
            
            If sdb_slab_thk_fr.Value <= dThk And sdb_slab_thk_to.Value >= dThk Then
            
                If sdb_slab_wid_fr.Value <= dWid And sdb_slab_wid_to.Value >= dWid Then
                
                    If sdb_slab_len_fr.Value <= dLen And sdb_slab_len_to.Value >= dLen Then
                    
                        If txt_stlgrdR_nm.Text <> "" Then
                        
                            If txt_stlgrdR_nm.Text = sStlgrd Then
                                
                                SS1.Col = SS1_HCR
                                If opt_hcr Then
                                    SS1.Text = "H"
                                Else
                                    SS1.Text = "C"
                                End If
                                
                                SS1.Col = 0
                                SS1.Text = "Update"
                            End If
                            
                        Else
                        
                            SS1.Col = SS1_HCR
                            If opt_hcr Then
                                SS1.Text = "H"
                            Else
                                SS1.Text = "C"
                            End If
                            
                            SS1.Col = 0
                            SS1.Text = "Update"
                            
                        End If
                        
                    End If
                    
                End If
                
            End If
    
        End If
        
    Next lRow
    
End Sub

Private Sub cmd_charge_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As adodb.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call AEB3000P ('" + txt_plt.Text + "','','','" + sUserID + "',?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Gp_MsgBoxDisplay("炉次编制完了..!!", "I")
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub cmd_slab_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As adodb.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call AEB2020P ('" + txt_plt.Text + "','" + txt_prc_line.Text + "','R','" + sUserID + "',?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Gp_MsgBoxDisplay("板坯设计完了..!!", "I")
'        txt_prod_cd.Text = "PP"
'        Call txt_prod_cd_KeyUp(0, 0)
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
'    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False
    
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
    
    'UPDATE AUTHORITY
    If Mid(sAuthority, 3, 1) <> "1" Then
        cmd_slab.Enabled = False
        cmd_charge.Enabled = False
    End If

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
'    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False
  
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "E-System.INI", Me.Name)
    
    If App.Title = "AE" Then
        txt_plt.Text = "B1"
    ElseIf App.Title = "BE" Then
        txt_plt.Text = "B1"
    End If
    
    Call txt_plt_KeyUp(0, 0)
    txt_prc_line.Text = "1"
    
'    txt_prod_cd.Text = "PP"
'    Call txt_prod_cd_KeyUp(0, 0)
    
    txt_hcr_fl_name.Text = ""
    Screen.MousePointer = vbDefault
    
   

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "E-System.INI", Me.Name)
    
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

Public Sub Spread_Can()

    SS1.Row = SS1.ActiveRow
    
    'If ss1.Text <> "" Then Exit Sub
    'ss1.Col = 0: ss1.Text = ""
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
    
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        
'        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
      
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
        
        If App.Title = "AE" Then
            txt_plt.Text = "B1"
        ElseIf App.Title = "BE" Then
            txt_plt.Text = "B1"
        End If
    
        Call txt_plt_KeyUp(0, 0)
        txt_prc_line.Text = "1"
        txt_hcr_fl_name.Text = ""
        
'        txt_prod_cd.Text = "PP"
'        Call txt_prod_cd_KeyUp(0, 0)
    
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err
    Dim iRow        As Integer
    Dim sQuery      As String
    Dim dSlabWgt    As Double
    Dim dPlateWgt   As Double
    
    sdb_err_cnt.Value = 0
    sdb_tot_cnt.Value = 0
    
    sQuery = "SELECT COUNT(*) FROM NISCO.EP_SLAB_EDT "
    sQuery = sQuery + " WHERE PROD_CD         LIKE '" + txt_prod_cd.Text + "%'"
    sQuery = sQuery + "   AND SMS_PLT         LIKE '" + txt_plt.Text + "%'"
    sQuery = sQuery + "   AND SMS_CCM_LINE    LIKE '" + txt_prc_line.Text + "%'"
    sQuery = sQuery + "   AND HCR_FL          LIKE '" + txt_hcr_fl.Text + "%'"
    sQuery = sQuery + "   AND STLGRD          LIKE '" + txt_stlgrd.Text + "%'"
    sQuery = sQuery + "   AND FL = 'E' "
    
    If cbo_order_cnt.ListIndex = 1 Then
        sQuery = sQuery + "   AND ORD_CNT > 1 "
    ElseIf cbo_order_cnt.ListIndex = 2 Then
        sQuery = sQuery + "   AND ORD_CNT = 1 "
    End If
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        SS1.OperationMode = OperationModeNormal
        
       
         
'        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
        sdb_err_cnt.Value = Gf_FloatFind(M_CN1, sQuery)
        sdb_tot_cnt.Value = SS1.MaxRows
'------------------------------------------------------------------------------------------------
        
         Call Spread_Color_Setting(SS1)
         Call SS1_CHANGE_COLOR
        
 '--------------------------------ss1,ord_fl=1,slab_len color--white,lock--------------
        dSlabWgt = 0
        dPlateWgt = 0

        With SS1
            For iRow = 1 To .MaxRows
                .Row = iRow
                .Col = SS1_PLATECNT
                If .Value = 1 Then
'                    .Protect = True
'                    .BlockMode = False
                    .Col = SS1_LEN
                    .Lock = True
'                    .BackColor = &H80000005
                End If

                .Col = SS1_WGT
                If Val(.Text & "") > 0 Then
                    .Col = SS1_WGT:      dSlabWgt = dSlabWgt + Val(.Text & "")
                    .Col = SS1_PLATE_WGT:      dPlateWgt = dPlateWgt + Val(.Text & "")
                End If
            Next
 '----------------------------------------change end-----------------------------------
            If dSlabWgt > 0 Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 2:       .Text = "合   计"
                .Col = SS1_LEN:       .Lock = True:       .BackColor = &H80000005
                .Col = SS1_WGT:      .Text = dSlabWgt
                .Col = SS1_PLATE_WGT:      .Text = dPlateWgt
                .Col = SS1_SHOUDE:      .Text = Format(dPlateWgt * 100 / dSlabWgt, "###.0")
                Call Gp_Sp_BlockColor(SS1, 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HE6E6FF)
            End If

        End With

        SS1.SetFocus
       
    End If
    
           
            
    Exit Sub

Refer_Err:

   

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
    End If
    
    Call Form_Ref
    
End Sub

Public Sub Form_Ins()
    
   Call Gp_Sp_Copy(Proc_Sc("Sc"))
   Call Gp_Sp_Paste(Proc_Sc("Sc"))
        
        SS1.Row = SS1.ActiveRow
        SS1.Col = 47
        SS1.Text = sUserID
                
        SS1.Col = 1
        SS1.Row = SS1.ActiveRow
        SS1.Text = "0"
End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 46)
    
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

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Spread_Del()
    
    SS1.Col = 4
    SS1.Row = SS1.ActiveRow
    
    'If ss1.Text <> "" Then Exit Sub
    'ss1.Col = 0: ss1.Text = ""
    Call Gp_Sp_Del(Proc_Sc("SC"))
    
    SS1.Row = SS1.MaxRows
    SS1.Col = 0: SS1.Text = ""

End Sub

Private Sub opt_ccr_Click(Value As Integer)

    If opt_ccr Then
        opt_ccr.ForeColor = &HFF&
        opt_hcr.ForeColor = &H80000012
    End If
    
End Sub

Private Sub opt_hcr_Click(Value As Integer)

    If opt_hcr Then
        opt_hcr.ForeColor = &HFF&
        opt_ccr.ForeColor = &H80000012
    End If
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Col = SS1_LEN Or Col = SS1_HCR Then Exit Sub
    If Row <= 0 Or Col <= 0 Then Exit Sub
    
    Complete = False
    
    SS1.Col = 0
    SS1.Row = Row
    If SS1.Text = "Input" Then Exit Sub
    
    SS1.Col = 1
    SS1.Row = Row
    If SS1.Text = "" Then Exit Sub
    
    
    Unload AEB2080C
    Load AEB2080C
    AEB2080C.sdb_slab_edt_seq.Value = SS1.Value
    AEB2080C.txt_slab_edt_fl.Text = Gf_FloatFind(M_CN1, "SELECT SLAB_EDT_FL FROM NISCO.EP_SLAB_EDT WHERE SLAB_EDT_SEQ = " & SS1.Value)
    SS1.Col = SS1_LEN
    AEB2080C.sdb_slab_len.Value = SS1.Value
    SS1.Col = SS1_WGT
    AEB2080C.sdb_slab_wgt.Value = SS1.Value
    AEB2080C.Show 1
    
    If Complete Then
        Call Gp_Sp_OneRowDisplay(M_CN1, Gf_Sp_MakeQuery(Proc_Sc("SC")("Spread"), Proc_Sc("SC")("P-O"), "O", Proc_Sc("SC")("pColumn"), Row), Proc_Sc("SC")("Spread"), Row)
    End If
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Row < 0 Or Row = 0 Then Exit Sub
    
    SS1.Col = 4
    SS1.Row = Row
    'If ss1.Text <> "" Then Exit Sub     'HEAT_MANA_NO IS NOT NULL  EXIT SUB
    
    SS1.Col = SS1_ORD_FL
    'If ss1.Text <> "2" Then Exit Sub    'ORD_FL <> '2'  EXIT SUB
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 29)
        
        'Call Gp_Sp_RowColor(ss1, Row, , &HFFC0FF)
        
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        'Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim TTT
    Dim DATA1
    Dim I

    Dim sTemp_Code As String

    If SS1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If
    
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.SS1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub txt_hcr_fl_DblClick()

    Call Txt_hcr_fl_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Txt_hcr_fl_KeyUp(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyF4 Then
        
        DD.sWitch = "MS"
        DD.sKey = "C0005"
        DD.rControl.Add Item:=txt_hcr_fl
        DD.rControl.Add Item:=txt_hcr_fl_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_hcr_fl.Text)) = txt_hcr_fl.MaxLength Then
            txt_hcr_fl_name.Text = Gf_ComnNameFind(M_CN1, "C0005", Trim(txt_hcr_fl.Text), 2)
        Else
            txt_hcr_fl_name.Text = ""
        End If
    
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

        If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
            txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
        Else
            txt_plt_name.Text = ""
        End If
    
    End If

End Sub

Private Sub txt_prod_cd_DblClick()

    Call txt_prod_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_prod_cd
        DD.rControl.Add Item:=txt_prod_cd_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else

        If Len(Trim(txt_prod_cd.Text)) = txt_prod_cd.MaxLength Then
            txt_prod_cd_name.Text = Gf_ComnNameFind(M_CN1, "B0005", Trim(txt_prod_cd.Text), 2)
        Else
            txt_prod_cd_name.Text = ""
        End If
    
    End If
    
End Sub

Private Sub txt_stlgrd_DblClick()

    Call txt_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
           
        DD.nameType = "1"
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stlgrd
        
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
            
    End If
    
End Sub

Public Sub Spread_Color_Setting(oSpr As vaSpread)

    Dim iRow   As Integer
    Dim sTemp  As String
    Dim iMode As Integer
    Dim intBlock As Integer
    Dim intStartRow As Integer
    
    intBlock = 1
    iMode = 0
    With oSpr
        
        For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 2: sTemp = .Text
            .Row = iRow + 1
            If sTemp = .Text Then
                intBlock = intBlock + 1
            Else
                If iMode = 1 Then
                    intStartRow = .Row - intBlock
                    Call Gp_Sp_BlockColor(oSpr, 1, 14, intStartRow, iRow, , &HFFC0FF)
                    Call Gp_Sp_BlockColor(oSpr, 16, 28, intStartRow, iRow, , &HFFC0FF)
                    Call Gp_Sp_BlockColor(oSpr, 30, .MaxCols, intStartRow, iRow, , &HFFC0FF)
                    iMode = 0
                Else
                    'Call Gp_Sp_BlockColor(oSpr, 1, .MaxCols, .Row - intBlock, iRow, , &HFFC0FF)
                    iMode = 1
                End If
                intBlock = 1
                
            End If
        Next iRow
    End With
    
End Sub

Private Sub txt_stlgrdR_DblClick()

    Call txt_stlgrdR_KeyUp(vbKeyF4, 0)
    
End Sub
Private Sub SS1_CHANGE_COLOR()

Dim iCount As Integer


    With SS1
      
        If .MaxRows <= 0 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .Row = iCount
            
            '有余材
            SS1.Row = .Row:       SS1.Col = SS1_FL_CNT
            If SS1.Value > 0 Then

'                 Call Gp_Sp_RowColor(ss1, .Row, , &HFF&)
                 Call Gp_Sp_BlockColor(SS1, 1, SS1_FL_CNT, .Row, .Row, &HFF&)

            
            End If
   
        Next iCount

    End With
    
End Sub

Private Sub txt_stlgrdR_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        
        DD.nameType = "1"
        DD.sWitch = "MS"
        
        DD.rControl.Add Item:=txt_stlgrdR
        DD.rControl.Add Item:=txt_stlgrdR_nm
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_stlgrdR.Text)) >= 10 Then
            txt_stlgrdR_nm.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stlgrdR.Text))
        Else
            txt_stlgrdR_nm.Text = ""
        End If
        
    End If
    
End Sub
