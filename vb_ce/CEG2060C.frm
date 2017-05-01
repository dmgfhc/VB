VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form CEG2060C 
   Caption         =   "HMI板坯设计_CEG2060C"
   ClientHeight    =   9285
   ClientLeft      =   1305
   ClientTop       =   1680
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   15345
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   5355
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   9446
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "CEG2060C.frx":0000
      Begin Threed.SSPanel SSPanel6 
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   873
         _Version        =   196609
         BackColor       =   14737918
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txt_stdgrd 
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
            Left            =   4065
            MaxLength       =   11
            TabIndex        =   8
            Top             =   90
            Width           =   1275
         End
         Begin VB.TextBox txt_ord_no 
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
            Left            =   1005
            MaxLength       =   11
            TabIndex        =   7
            Tag             =   "产品"
            Top             =   90
            Width           =   1305
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
            Left            =   2310
            TabIndex        =   6
            Top             =   90
            Width           =   660
         End
         Begin CSTextLibCtl.sidbEdit sdb_thk_fr 
            Height          =   315
            Left            =   6450
            TabIndex        =   9
            Top             =   90
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Index           =   1
            Left            =   5520
            Top             =   90
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            Caption         =   "产品厚度"
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
         Begin InDate.ULabel ULabel35 
            Height          =   315
            Left            =   8760
            Top             =   90
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            Caption         =   "产品宽度"
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
         Begin InDate.ULabel ULabel36 
            Height          =   315
            Left            =   12000
            Top             =   90
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            Caption         =   "产品长度"
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
         Begin InDate.ULabel ULabel37 
            Height          =   315
            Left            =   3120
            Top             =   90
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            Caption         =   "钢种"
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
         Begin CSTextLibCtl.sidbEdit sdb_thk_to 
            Height          =   315
            Left            =   7545
            TabIndex        =   10
            Top             =   90
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            RawData         =   "9999.99"
            Text            =   " 9,999.99"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            Undo            =   0
            Data            =   9999.99
         End
         Begin CSTextLibCtl.sidbEdit sdb_len_fr 
            Height          =   315
            Left            =   12945
            TabIndex        =   11
            Top             =   90
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_len_to 
            Height          =   315
            Left            =   14040
            TabIndex        =   12
            Top             =   90
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "99999.9"
            Text            =   " 99,999.9"
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   99999.9
         End
         Begin CSTextLibCtl.sidbEdit sdb_wid_fr 
            Height          =   315
            Left            =   9705
            TabIndex        =   13
            Top             =   90
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_wid_to 
            Height          =   315
            Left            =   10800
            TabIndex        =   14
            Top             =   90
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            RawData         =   "9999.99"
            Text            =   " 9,999.99"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            Undo            =   0
            Data            =   9999.99
         End
         Begin InDate.ULabel ULabel38 
            Height          =   315
            Left            =   60
            Top             =   90
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            Caption         =   "订单号"
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
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   4830
         Left            =   0
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   525
         Width           =   15225
         _Version        =   393216
         _ExtentX        =   26855
         _ExtentY        =   8520
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ButtonDrawMode  =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   28
         MaxRows         =   1
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CEG2060C.frx":0052
      End
   End
   Begin VB.TextBox txt_stlgrd1 
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
      Left            =   9885
      MaxLength       =   11
      TabIndex        =   0
      Tag             =   "钢种"
      Top             =   10245
      Visible         =   0   'False
      Width           =   1275
   End
   Begin CSTextLibCtl.sidbEdit sdb_thk1 
      Height          =   315
      Left            =   11235
      TabIndex        =   1
      Tag             =   "产品厚度（MIN）"
      Top             =   10245
      Visible         =   0   'False
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
      NumDecDigits    =   2
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_wid1 
      Height          =   315
      Left            =   12585
      TabIndex        =   2
      Tag             =   "产品宽度（MIN）"
      Top             =   10245
      Visible         =   0   'False
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
      NumDecDigits    =   2
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_len1 
      Height          =   315
      Left            =   13935
      TabIndex        =   3
      Tag             =   "产品长度（MIN）"
      Top             =   10245
      Visible         =   0   'False
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
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.0"
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
      NumDecDigits    =   1
      NumIntDigits    =   7
      Undo            =   0
      Data            =   0
   End
   Begin SSSplitter.SSSplitter SSSplitter3 
      Height          =   3705
      Left            =   60
      TabIndex        =   15
      Top             =   5490
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   6535
      _Version        =   196609
      SplitterBarWidth=   3
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   12632319
      PaneTree        =   "CEG2060C.frx":108B
      Begin Threed.SSPanel SSPanel1 
         Height          =   810
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   1429
         _Version        =   196609
         BackColor       =   14737918
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txt_ord_no4 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   310
            Left            =   930
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   22
            Top             =   420
            Width           =   1650
         End
         Begin VB.TextBox txt_ord_no5 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   310
            Left            =   6015
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   21
            Top             =   420
            Width           =   1650
         End
         Begin VB.TextBox txt_ord_no6 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   310
            Left            =   11085
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   20
            Top             =   420
            Width           =   1650
         End
         Begin VB.TextBox txt_ord_no1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   310
            Left            =   930
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   19
            Top             =   60
            Width           =   1650
         End
         Begin VB.TextBox txt_ord_no2 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   310
            Left            =   6015
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   18
            Top             =   60
            Width           =   1650
         End
         Begin VB.TextBox txt_ord_no3 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   310
            Left            =   11085
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   17
            Top             =   60
            Width           =   1650
         End
         Begin InDate.ULabel ULabel14 
            Height          =   315
            Left            =   120
            Top             =   60
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "订单号"
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
            Left            =   5205
            Top             =   60
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "订单号"
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
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   10275
            Top             =   60
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "订单号"
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
         Begin CSTextLibCtl.sidbEdit sdb_ord11_cnt 
            Height          =   315
            Left            =   3450
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   60
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel13 
            Height          =   315
            Left            =   2640
            Top             =   60
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "张数"
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
         Begin CSTextLibCtl.sidbEdit sdb_ord21_cnt 
            Height          =   315
            Left            =   8535
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   60
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel15 
            Height          =   315
            Left            =   7725
            Top             =   60
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "张数"
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
         Begin CSTextLibCtl.sidbEdit sdb_ord31_cnt 
            Height          =   315
            Left            =   13605
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   60
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel21 
            Height          =   315
            Left            =   12795
            Top             =   60
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "张数"
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
         Begin Threed.SSCommand cmd_ord1 
            Height          =   330
            Left            =   4395
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   60
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "适用"
            BevelWidth      =   3
         End
         Begin Threed.SSCommand cmd_ord2 
            Height          =   330
            Left            =   9480
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   60
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "适用"
            BevelWidth      =   3
         End
         Begin Threed.SSCommand cmd_ord3 
            Height          =   330
            Left            =   14550
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   60
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "适用"
            BevelWidth      =   3
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord32_cnt 
            Height          =   315
            Left            =   14055
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   60
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord22_cnt 
            Height          =   315
            Left            =   8985
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   60
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord12_cnt 
            Height          =   315
            Left            =   3900
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   60
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord1_len 
            Height          =   315
            Left            =   930
            TabIndex        =   32
            Top             =   30
            Visible         =   0   'False
            Width           =   1410
            _Version        =   262145
            _ExtentX        =   2487
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord2_len 
            Height          =   315
            Left            =   6015
            TabIndex        =   33
            Top             =   30
            Visible         =   0   'False
            Width           =   1410
            _Version        =   262145
            _ExtentX        =   2487
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord3_len 
            Height          =   315
            Left            =   11085
            TabIndex        =   34
            Top             =   30
            Visible         =   0   'False
            Width           =   1410
            _Version        =   262145
            _ExtentX        =   2487
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel28 
            Height          =   315
            Left            =   5205
            Top             =   420
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "订单号"
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
         Begin InDate.ULabel ULabel29 
            Height          =   315
            Left            =   10275
            Top             =   420
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "订单号"
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
         Begin CSTextLibCtl.sidbEdit sdb_ord41_cnt 
            Height          =   315
            Left            =   3450
            TabIndex        =   35
            Top             =   420
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel30 
            Height          =   315
            Left            =   2640
            Top             =   420
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "张数"
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
         Begin CSTextLibCtl.sidbEdit sdb_ord51_cnt 
            Height          =   315
            Left            =   8535
            TabIndex        =   36
            Top             =   420
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel31 
            Height          =   315
            Left            =   7725
            Top             =   420
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "张数"
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
         Begin CSTextLibCtl.sidbEdit sdb_ord61_cnt 
            Height          =   315
            Left            =   13605
            TabIndex        =   37
            Top             =   420
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel32 
            Height          =   315
            Left            =   12795
            Top             =   420
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "张数"
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
         Begin Threed.SSCommand cmd_ord4 
            Height          =   330
            Left            =   4395
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   420
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "适用"
            BevelWidth      =   3
         End
         Begin Threed.SSCommand cmd_ord5 
            Height          =   330
            Left            =   9480
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   420
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "适用"
         End
         Begin Threed.SSCommand cmd_ord6 
            Height          =   330
            Left            =   14550
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   420
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "适用"
            BevelWidth      =   3
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord62_cnt 
            Height          =   315
            Left            =   14055
            TabIndex        =   41
            Top             =   420
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord42_cnt 
            Height          =   315
            Left            =   3900
            TabIndex        =   42
            Top             =   420
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel33 
            Height          =   315
            Left            =   120
            Top             =   405
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "订单号"
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
         Begin CSTextLibCtl.sidbEdit sdb_ord6_len 
            Height          =   315
            Left            =   11085
            TabIndex        =   43
            Top             =   375
            Visible         =   0   'False
            Width           =   1410
            _Version        =   262145
            _ExtentX        =   2487
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord5_len 
            Height          =   315
            Left            =   6015
            TabIndex        =   44
            Top             =   375
            Visible         =   0   'False
            Width           =   1410
            _Version        =   262145
            _ExtentX        =   2487
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord4_len 
            Height          =   315
            Left            =   930
            TabIndex        =   45
            Top             =   405
            Visible         =   0   'False
            Width           =   1410
            _Version        =   262145
            _ExtentX        =   2487
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord52_cnt 
            Height          =   315
            Left            =   8985
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   420
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   0
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1215
         Left            =   0
         TabIndex        =   47
         Top             =   855
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   2143
         _Version        =   196609
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin InDate.ULabel lbl_mplate 
            Height          =   150
            Index           =   0
            Left            =   675
            Top             =   180
            Visible         =   0   'False
            Width           =   105
            _ExtentX        =   185
            _ExtentY        =   265
            Caption         =   ""
            Alignment       =   1
            BackColor       =   12632256
            BackgroundStyle =   1
            BorderEffect    =   0
            BorderStyle     =   1
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
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   45
            Top             =   45
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            Caption         =   "母板"
            Alignment       =   1
            BackColor       =   16761024
            BackgroundStyle =   1
            BorderEffect    =   0
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
         Begin Threed.SSCommand cmd_mplate_init 
            Height          =   315
            Left            =   13260
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   420
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "初始化"
            BevelWidth      =   3
         End
         Begin Threed.SSCommand cmd_mplate_design 
            Height          =   315
            Left            =   13260
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   60
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   255
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "设计探讨"
            BevelWidth      =   3
         End
         Begin Threed.SSCommand cmd_mplate_del 
            Height          =   315
            Left            =   14205
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   60
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   32896
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "删除"
            BevelWidth      =   3
         End
         Begin Threed.SSCommand cmd_mplate_complete 
            Height          =   315
            Left            =   14205
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   420
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   12583104
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "设计确定"
            BevelWidth      =   3
         End
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   630
            Top             =   810
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "母板厚度"
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   5955
            Top             =   810
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "母板长度"
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
         Begin CSTextLibCtl.sidbEdit sdb_asroll_len 
            Height          =   315
            Left            =   7350
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   810
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_asroll_thk 
            Height          =   315
            Left            =   2010
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   810
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   3285
            Top             =   810
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "母板宽度"
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
         Begin CSTextLibCtl.sidbEdit sdb_asroll_wid 
            Height          =   315
            Left            =   4665
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   810
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   8670
            Top             =   810
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "探讨母板长度"
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
         Begin CSTextLibCtl.sidbEdit sdb_plate_len 
            Height          =   315
            Left            =   10065
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   810
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00000000&
            Height          =   675
            Left            =   630
            Shape           =   4  'Rounded Rectangle
            Top             =   60
            Width           =   12525
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "50(M)"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   450
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1590
         Left            =   0
         TabIndex        =   57
         Top             =   2115
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   2805
         _Version        =   196609
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin InDate.ULabel ULabel23 
            Height          =   315
            Left            =   45
            Top             =   45
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            Caption         =   "轧件"
            Alignment       =   1
            BackColor       =   12640511
            BackgroundStyle =   1
            BorderEffect    =   0
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
         Begin Threed.SSCommand cmd_slab_init 
            Height          =   315
            Left            =   13260
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   390
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "初始化"
            BevelWidth      =   3
         End
         Begin Threed.SSCommand cmd_slab_design 
            Height          =   315
            Left            =   13260
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   45
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   255
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "设计探讨"
            BevelWidth      =   3
         End
         Begin InDate.ULabel lbl_slab 
            Height          =   150
            Index           =   0
            Left            =   675
            Top             =   180
            Visible         =   0   'False
            Width           =   105
            _ExtentX        =   185
            _ExtentY        =   265
            Caption         =   ""
            Alignment       =   1
            BackColor       =   16744576
            BackgroundStyle =   1
            BorderEffect    =   0
            BorderStyle     =   1
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
         Begin Threed.SSCommand cmd_slab_del 
            Height          =   315
            Left            =   14205
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   45
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   32896
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "删除"
            BevelWidth      =   3
         End
         Begin Threed.SSCommand cmd_slab_complete 
            Height          =   315
            Left            =   14205
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   390
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   12583104
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "设计确定"
            BevelWidth      =   3
         End
         Begin InDate.ULabel ULabel19 
            Height          =   315
            Left            =   630
            Top             =   810
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "轧件厚度"
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
         Begin InDate.ULabel ULabel20 
            Height          =   315
            Left            =   5955
            Top             =   810
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "轧件长度"
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
         Begin CSTextLibCtl.sidbEdit sdb_slab_len 
            Height          =   315
            Left            =   7350
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   810
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_slab_thk 
            Height          =   315
            Left            =   2010
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   810
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Left            =   3285
            Top             =   810
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "轧件宽度"
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
         Begin CSTextLibCtl.sidbEdit sdb_slab_wid 
            Height          =   315
            Left            =   4665
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   810
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel24 
            Height          =   315
            Left            =   11310
            Top             =   810
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            Caption         =   "母板长度"
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
         Begin CSTextLibCtl.sidbEdit sdb_asroll_prod_len 
            Height          =   315
            Left            =   12405
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   810
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel40 
            Height          =   315
            Left            =   8670
            Top             =   1200
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "板坯重量"
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
         Begin CSTextLibCtl.sidbEdit sdb_slab_wgt1 
            Height          =   315
            Left            =   10065
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel34 
            Height          =   315
            Left            =   11310
            Top             =   1200
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            Caption         =   "成材率"
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
         Begin CSTextLibCtl.sidbEdit sdb_slab_ratio 
            Height          =   315
            Left            =   12390
            TabIndex        =   69
            Top             =   1200
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSCommand cmd_design_modify 
            Height          =   315
            Left            =   14205
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   750
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   16576
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "设计现状"
            BevelWidth      =   3
         End
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   630
            Top             =   1200
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "板坯厚度"
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
            Left            =   5955
            Top             =   1200
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "板坯长度"
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
         Begin CSTextLibCtl.sidbEdit sdb_slab_len1 
            Height          =   315
            Left            =   7350
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_slab_thk1 
            Height          =   315
            Left            =   2010
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   3285
            Top             =   1200
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "板坯宽度"
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
         Begin CSTextLibCtl.sidbEdit sdb_slab_wid1 
            Height          =   315
            Left            =   4665
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
         Begin CSTextLibCtl.sidbEdit sdb_slab_lenq 
            Height          =   315
            Left            =   120
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   1170
            Visible         =   0   'False
            Width           =   570
            _Version        =   262145
            _ExtentX        =   1005
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   8670
            Top             =   810
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "轧件重量"
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
         Begin CSTextLibCtl.sidbEdit sdb_slab_wgt 
            Height          =   315
            Left            =   10065
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   810
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   16711680
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
            NumIntDigits    =   7
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_slab_cnt 
            Height          =   315
            Left            =   14670
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   1200
            Width           =   435
            _Version        =   262145
            _ExtentX        =   767
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   13860
            Top             =   1200
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            Caption         =   "板坯数"
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
         Begin VB.Label Label5 
            BackColor       =   &H00C0E0FF&
            Height          =   240
            Left            =   90
            TabIndex        =   66
            Top             =   510
            Width           =   525
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00000000&
            Height          =   675
            Left            =   630
            Shape           =   4  'Rounded Rectangle
            Top             =   90
            Width           =   12525
         End
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   60
      X2              =   15240
      Y1              =   5460
      Y2              =   5460
   End
End
Attribute VB_Name = "CEG2060C"
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
'-- Program ID        CEG20600C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2008.5.5
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

Public Complete As Boolean          'Plate Delete Setting

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

Dim pContro2 As New Collection      'Master Primary Key Collection
Dim nContro2 As New Collection      'Master Necessary Collection
Dim mContro2 As New Collection      'Master Maxlength check Collection
Dim iContro2 As New Collection      'Master Insert Collection
Dim rContro2 As New Collection      'Master Refer Collection
Dim cContro2 As New Collection      'Master Copy Collection
Dim aContro2 As New Collection      'Master -> Spread Collection
Dim lContro2 As New Collection      'Master Lock Collection

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

Dim oRd_cnt As Integer              'Select Order Count
Dim iMplate_cnt As Integer          'Mplate Design Count
Dim iSlab_cnt As Integer            'Slab Design Count
Dim iLastSlab_cnt As Integer        'Last Slab Complte Count
Dim iSlab_Complete As Integer       'Slab Complete Count
Dim iSlab_Row As Integer            'Slab Row
Dim iOrd1_Curr_Row As Integer       'Select Order1 Row
Dim iOrd2_Curr_Row As Integer       'Select Order2 Row
Dim iOrd3_Curr_Row As Integer       'Select Order3 Row
Dim iOrd4_Curr_Row As Integer       'Select Order1 Row
Dim iOrd5_Curr_Row As Integer       'Select Order2 Row
Dim iOrd6_Curr_Row As Integer       'Select Order3 Row

Dim lCool_max As Long               'COOLING BED LENGTH MAX SIZE
Dim lAsroll_max As Long             'ASROLL LENGTH MAX SIZE

Dim lMain_row As Long               'Main Row(Order no1)
Dim lSlab_left As Long              'Slab Left Position
Dim lMplate_left As Long            'Mplate Left Position
Dim iSLAB_EDT_SEQ As Long           'SLAB_EDT_SEQ Value

Dim vCR_CD As Variant               'First Slab CR_CD
Dim vSTLGRD As Variant              'First Slab STLGR
Dim vUST_FL As Variant              'First Slab UST_FL
Dim vPROD_WID As Variant            'First Slab PROD_WID
Dim vPROD_THK As Variant            'First Slab PROD_THK
Dim vENDUSE_CD As Variant           'First Slab ENDUSE_CD
Dim vORD_HCR_FL As Variant          'First Slab ORD_HCR_F
Dim vORD_TRIM_FL As Variant         'First Slab ORD_TRIM_FL

Dim sHTM_METH As String             'First Plate HTM_METH

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()

    Dim lSpread_Row As Integer
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(txt_ord_no1, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(txt_ord_no2, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(txt_ord_no3, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(txt_ord_no4, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(txt_ord_no5, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(txt_ord_no6, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_ord11_cnt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_ord12_cnt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_ord21_cnt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_ord22_cnt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_ord31_cnt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_ord32_cnt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_ord41_cnt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_ord42_cnt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_ord51_cnt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_ord52_cnt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_ord61_cnt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_ord62_cnt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(sdb_ord1_len, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(sdb_ord2_len, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(sdb_ord3_len, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(sdb_ord4_len, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(sdb_ord5_len, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(sdb_ord6_len, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(sdb_asroll_thk, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(sdb_asroll_wid, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(sdb_asroll_len, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
  Call Gp_Ms_Collection(sdb_asroll_prod_len, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_plate_len, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(sdb_slab_thk, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(sdb_slab_wid, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(sdb_slab_len, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_slab_thk1, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_slab_wid1, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_slab_len1, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(sdb_slab_wgt1, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(sdb_slab_ratio, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
  
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
   'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
   Call Gp_Ms_Collection(cbo_ord_item, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
     Call Gp_Ms_Collection(txt_stdgrd, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
     Call Gp_Ms_Collection(sdb_thk_fr, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
     Call Gp_Ms_Collection(sdb_thk_to, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
     Call Gp_Ms_Collection(sdb_wid_fr, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
     Call Gp_Ms_Collection(sdb_wid_to, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
     Call Gp_Ms_Collection(sdb_len_fr, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
     Call Gp_Ms_Collection(sdb_len_to, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
   
    'MASTER Collection
    Mc2.Add Item:=pContro2, Key:="pControl"
    Mc2.Add Item:=nContro2, Key:="nControl"
    Mc2.Add Item:=mContro2, Key:="mControl"
    Mc2.Add Item:=iContro2, Key:="iControl"
    Mc2.Add Item:=rContro2, Key:="rControl"
    Mc2.Add Item:=cContro2, Key:="cControl"
    Mc2.Add Item:=aContro2, Key:="aControl"
    Mc2.Add Item:=lContro2, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, "p", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    For lSpread_Row = 3 To 28
        Call Gp_Sp_Collection(ss2, lSpread_Row, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next lSpread_Row
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CEG2060C.P_REFER2", Key:="P-R"
    sc2.Add Item:="CEG2060C.P_ONEROW2", Key:="P-O"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Proc_Sc.Add Item:=sc2, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss2, 27, True)
    
    iMplate_cnt = 0
    iSlab_cnt = 0
    
End Sub

Private Sub cmd_design_modify_Click()

    Dim P_SLAB_NO As String
    
    Complete = False
    Load CEG2061C

    P_SLAB_NO = "9999999999"

    CEG2061C.txt_slab_no.Text = P_SLAB_NO
    CEG2061C.Show 1

    If Complete Then
        Call cmd_slab_design_Click
        Call cmd_slab_del_Click
        cmd_slab_complete.Enabled = False
    End If
    
End Sub

Private Sub cmd_mplate_complete_Click()
    
    Dim sSeq As String
    Dim sQuery As String
    
    If sdb_plate_len.Value = 0 Then Exit Sub
    
    'SLAB SIZE
    'Call Slab_Size
    
    If sdb_plate_len.Value + sdb_slab_len.Value >= sdb_slab_lenq.Value Then
        Call Gp_MsgBoxDisplay("轧件长度 > " & sdb_slab_lenq.Text & "(mm)")
        Exit Sub
    End If
    
    If iSlab_cnt = 0 Then
        ss2.Row = lMain_row
        
        ss2.Col = 5
        vSTLGRD = ss2.Text
        ss2.Col = 4
        vENDUSE_CD = ss2.Text
        ss2.Col = 8
        vPROD_THK = ss2.Value
        ss2.Col = 11
        vPROD_WID = ss2.Value
        'ss2.Col = 18
        'vORD_HCR_FL = ss2.Text
        ss2.Col = 21
        vCR_CD = ss2.Text
        ss2.Col = 22
        vORD_TRIM_FL = ss2.Text
        ss2.Col = 23
        vUST_FL = ss2.Text
    End If
    
    iSlab_cnt = iSlab_cnt + 1
    cmd_slab_del.Enabled = True
    cmd_slab_design.Enabled = True
    cmd_slab_complete.Enabled = False
    
    iSlab_Complete = 0
    sdb_slab_wgt1.Value = 0
    sdb_asroll_prod_len.Value = 0
    
    If iSlab_cnt < 10 Then
        sSeq = "0" & iSlab_cnt
    Else
        sSeq = Trim(Str(iSlab_cnt))
    End If
    
    sdb_slab_len.Value = sdb_slab_len.Value + sdb_plate_len.Value
    
    Load lbl_slab(iSlab_cnt)
    lbl_slab(iSlab_cnt).Tag = Str(sdb_plate_len.Value)
    lbl_slab(iSlab_cnt).Caption = sSeq
    lbl_slab(iSlab_cnt).Top = 180
    lbl_slab(iSlab_cnt).Height = 500
    lbl_slab(iSlab_cnt).Width = (Shape4.Width / sdb_slab_lenq.Value) * sdb_plate_len.Value
        
    If iSlab_cnt = 1 Then
        lbl_slab(iSlab_cnt).Left = Shape4.Left
        lbl_slab(iSlab_cnt).Visible = True
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='00', SEQ='00'
        Call Slab_Block_Seq_Create("I")
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ=sSeq, SEQ= '00' ADD 1
        Call Slab_Seq_Create(sSeq, "I")
        
    Else
    
        If lbl_slab(iSlab_cnt - 1).Caption <> "删除" Then
            lbl_slab(iSlab_cnt).Left = lbl_slab(iSlab_cnt - 1).Left + lbl_slab(iSlab_cnt - 1).Width
        Else
            lbl_slab(iSlab_cnt).Left = lbl_slab(iSlab_cnt - 1).Left + lbl_slab(iSlab_cnt - 1).Width - 30
        End If
        
        lbl_slab(iSlab_cnt).Visible = True
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ=sSeq, SEQ= '00' ADD 1
        Call Slab_Seq_Create(sSeq, "I")
        
    End If
    
End Sub

Private Sub cmd_mplate_del_Click()

    Dim sSeq As String
    
    Dim iCount As Integer
    Dim iVisible_Cnt As Integer
    
    If iMplate_cnt = 0 Then Exit Sub
    
    For iCount = 1 To iMplate_cnt
        
        If lbl_mplate(iCount).Caption = "删除" Then
            
            If lbl_mplate(iCount).Visible Then
            
                lbl_mplate(iCount).Width = 0
                lbl_mplate(iCount).Visible = False
                
                If lbl_mplate(iCount).Tag = "ord1" Then
                    sdb_asroll_len.Value = sdb_asroll_len.Value - sdb_ord1_len.Value
                ElseIf lbl_mplate(iCount).Tag = "ord2" Then
                    sdb_asroll_len.Value = sdb_asroll_len.Value - sdb_ord2_len.Value
                ElseIf lbl_mplate(iCount).Tag = "ord3" Then
                    sdb_asroll_len.Value = sdb_asroll_len.Value - sdb_ord3_len.Value
                ElseIf lbl_mplate(iCount).Tag = "ord4" Then
                    sdb_asroll_len.Value = sdb_asroll_len.Value - sdb_ord4_len.Value
                ElseIf lbl_mplate(iCount).Tag = "ord5" Then
                    sdb_asroll_len.Value = sdb_asroll_len.Value - sdb_ord5_len.Value
                Else
                    sdb_asroll_len.Value = sdb_asroll_len.Value - sdb_ord6_len.Value
                End If
                
                If iCount < 10 Then
                    sSeq = "0" & iCount
                Else
                    sSeq = Str(iCount)
                End If
                
                'EP_SLAB_EDT3_D UPDATE  BLOCK_SEQ='01', SEQ      --> LEN = 0
                If lbl_mplate(iCount).Tag = "ord1" Then
                    Call Plate_Seq_Create(iOrd1_Curr_Row, sSeq, "U")
                ElseIf lbl_mplate(iCount).Tag = "ord2" Then
                    Call Plate_Seq_Create(iOrd2_Curr_Row, sSeq, "U")
                ElseIf lbl_mplate(iCount).Tag = "ord3" Then
                    Call Plate_Seq_Create(iOrd3_Curr_Row, sSeq, "U")
                ElseIf lbl_mplate(iCount).Tag = "ord4" Then
                    Call Plate_Seq_Create(iOrd4_Curr_Row, sSeq, "U")
                ElseIf lbl_mplate(iCount).Tag = "ord5" Then
                    Call Plate_Seq_Create(iOrd5_Curr_Row, sSeq, "U")
                Else
                    Call Plate_Seq_Create(iOrd6_Curr_Row, sSeq, "U")
                End If
                    
            End If
            
            cmd_mplate_complete.Enabled = False
            sdb_plate_len.Value = 0
            
        End If
    
        If iCount = 1 Then
            lbl_mplate(iCount).Left = Shape1.Left
        Else
            If lbl_mplate(iCount - 1).Caption <> "删除" Then
                lbl_mplate(iCount).Left = lbl_mplate(iCount - 1).Left + lbl_mplate(iCount - 1).Width
            Else
                lbl_mplate(iCount).Left = lbl_mplate(iCount - 1).Left + lbl_mplate(iCount - 1).Width - 30
            End If
        End If
        
    Next iCount
    
    iVisible_Cnt = 0
    For iCount = 1 To iMplate_cnt
    
        If lbl_mplate(iCount).Visible Then
            iVisible_Cnt = iVisible_Cnt + 1
        End If
    
    Next iCount
    
    'EP_SLAB_EDT3_D DATA DELETE
    If iVisible_Cnt = 0 Then
    
        For iCount = 1 To iMplate_cnt
            Unload lbl_mplate(iCount)
        Next iCount
        
        iMplate_cnt = 0
        Call Plate_Seq_Create(lMain_row, "00", "D")
        sdb_asroll_thk.Value = 0
        sdb_asroll_wid.Value = 0
        sdb_slab_thk.Value = 0
        sdb_slab_wid.Value = 0
        cmd_mplate_del.Enabled = False
        cmd_mplate_design.Enabled = False
        
        If iSlab_cnt <= 0 Then
            sHTM_METH = ""
        End If
        
    End If
    
End Sub

Private Sub cmd_mplate_design_Click()

On Error GoTo Process_Exec_ERROR
    
    Dim sQuery As String
    Dim iCount As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim OutParam(2, 4) As Variant
    Dim iVisible_Cnt As Integer
    
    Dim adoCmd As ADODB.Command
    Dim AdoRs As ADODB.Recordset
    
    Set AdoRs = New ADODB.Recordset
    
    '----------------------------------------------------
    
    For iCount = 1 To iMplate_cnt
        If lbl_mplate(iCount).Visible Then
            iVisible_Cnt = iVisible_Cnt + 1
        End If
    Next iCount
    
    If iVisible_Cnt = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    'SLAB_NO, SLAB_CUT_SEQ, BLOCK_SEQ, EMP_NO
    sQuery = "{call CEG2062P ('9999999999','99','99','" + sUserID + "',?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'DESIGN LEN
'    sQuery = "SELECT NVL(LEN,0) FROM nisco.EP_SLAB_EDT3_D WHERE SLAB_NO = '9999999999' AND BLOCK_SEQ = '99' AND  SEQ = '00' "
'    sdb_plate_len.Value = Gf_FloatFind(M_CN1, sQuery)
    
    sQuery = "         SELECT  NVL(THK,0) ,NVL(WID,0) , NVL(LEN,0)"
    sQuery = sQuery + "  FROM  EP_SLAB_EDT3_D "
    sQuery = sQuery + " WHERE  SLAB_NO = '9999999999' AND SLAB_CUT_SEQ = '99' AND BLOCK_SEQ = '99' AND  SEQ = '00' "
    
    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    Do Until AdoRs.EOF
        sdb_asroll_thk.Value = Val(AdoRs.Fields(0) & "")
        sdb_asroll_wid.Value = Val(AdoRs.Fields(1) & "")
        'sdb_asroll_len.Value = Val(AdoRs.Fields(2) & "")
        sdb_plate_len.Value = Val(AdoRs.Fields(2) & "")
        
        AdoRs.MoveNext
    Loop
    AdoRs.Close
    
    'Process Error Check
    If adoCmd("arg_e_code") <> "Y" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
        cmd_mplate_complete.Enabled = True
    Else
        cmd_mplate_complete.Enabled = True
    End If
    
    iVisible_Cnt = 0
    
    'Caption Rewrite
    For iCount = 1 To iMplate_cnt
    
        If lbl_mplate(iCount).Visible Then
        
            iVisible_Cnt = iVisible_Cnt + 1
            
            lbl_mplate(iVisible_Cnt).Visible = True
            lbl_mplate(iVisible_Cnt).BackColor = &HC0C0C0
            lbl_mplate(iVisible_Cnt).ForeColor = &HFF0000
            
            If iVisible_Cnt < 10 Then
                lbl_mplate(iVisible_Cnt).Caption = "0" & iVisible_Cnt
            Else
                lbl_mplate(iVisible_Cnt).Caption = Trim(Str(iVisible_Cnt))
            End If
            
            lbl_mplate(iVisible_Cnt).Top = 150
            lbl_mplate(iVisible_Cnt).Height = 500
            
            If lbl_mplate(iCount).Tag = "ord1" Then
                lbl_mplate(iVisible_Cnt).Width = (Shape1.Width / lCool_max) * sdb_ord1_len.Value
            ElseIf lbl_mplate(iCount).Tag = "ord2" Then
                lbl_mplate(iVisible_Cnt).Width = (Shape1.Width / lCool_max) * sdb_ord2_len.Value
            ElseIf lbl_mplate(iCount).Tag = "ord3" Then
                lbl_mplate(iVisible_Cnt).Width = (Shape1.Width / lCool_max) * sdb_ord3_len.Value
            ElseIf lbl_mplate(iCount).Tag = "ord4" Then
                lbl_mplate(iVisible_Cnt).Width = (Shape1.Width / lCool_max) * sdb_ord4_len.Value
            ElseIf lbl_mplate(iCount).Tag = "ord5" Then
                lbl_mplate(iVisible_Cnt).Width = (Shape1.Width / lCool_max) * sdb_ord5_len.Value
            Else
                lbl_mplate(iVisible_Cnt).Width = (Shape1.Width / lCool_max) * sdb_ord6_len.Value
            End If
                
            lbl_mplate(iVisible_Cnt).Tag = lbl_mplate(iCount).Tag
            
            If iVisible_Cnt = 1 Then
                lbl_mplate(iVisible_Cnt).Left = Shape1.Left
            Else
                If lbl_mplate(iVisible_Cnt - 1).Caption <> "删除" Then
                    lbl_mplate(iVisible_Cnt).Left = lbl_mplate(iVisible_Cnt - 1).Left + lbl_mplate(iVisible_Cnt - 1).Width
                Else
                    lbl_mplate(iVisible_Cnt).Left = lbl_mplate(iVisible_Cnt - 1).Left + lbl_mplate(iVisible_Cnt - 1).Width - 30
                End If
                
            End If
                
        End If
    
    Next iCount
    
    'Remain Plate Delete
    For iCount = iVisible_Cnt + 1 To iMplate_cnt
        Unload lbl_mplate(iCount)
    Next iCount
    
    iMplate_cnt = iVisible_Cnt
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Err.Raise Err.Number, Err.Description & sQuery
    
End Sub

Private Sub cmd_mplate_init_Click()

    Dim iCnt As Long
    Dim iRow As Integer
    
    For iCnt = 1 To iMplate_cnt
        lbl_mplate(iCnt).Caption = "删除"
    Next iCnt

    Call cmd_mplate_del_Click
    
    txt_ord_no1.Text = ""
    txt_ord_no2.Text = ""
    txt_ord_no3.Text = ""
    txt_ord_no4.Text = ""
    txt_ord_no5.Text = ""
    txt_ord_no6.Text = ""
    sdb_ord11_cnt.Value = 0
    sdb_ord12_cnt.Value = 0
    sdb_ord21_cnt.Value = 0
    sdb_ord22_cnt.Value = 0
    sdb_ord31_cnt.Value = 0
    sdb_ord32_cnt.Value = 0
    sdb_ord41_cnt.Value = 0
    sdb_ord42_cnt.Value = 0
    sdb_ord51_cnt.Value = 0
    sdb_ord52_cnt.Value = 0
    sdb_ord61_cnt.Value = 0
    sdb_ord62_cnt.Value = 0
    sdb_ord1_len.Value = 0
    sdb_ord2_len.Value = 0
    sdb_ord3_len.Value = 0
    sdb_ord4_len.Value = 0
    sdb_ord5_len.Value = 0
    sdb_ord6_len.Value = 0
    sdb_asroll_thk.Value = 0
    sdb_asroll_wid.Value = 0
    sdb_slab_thk.Value = 0
    sdb_slab_wid.Value = 0
    sdb_slab_len.Value = 0
    sdb_asroll_len.Value = 0
    sdb_asroll_prod_len.Value = 0
    sdb_plate_len.Value = 0
    iOrd1_Curr_Row = 0
    iOrd2_Curr_Row = 0
    iOrd3_Curr_Row = 0
    iOrd4_Curr_Row = 0
    iOrd5_Curr_Row = 0
    iOrd6_Curr_Row = 0
    lMain_row = 0
    oRd_cnt = 0
    
    iMplate_cnt = 0
    cmd_mplate_del.Enabled = False
    cmd_mplate_complete.Enabled = False
    
    If iSlab_cnt <= 0 Then
        sHTM_METH = ""
        sdb_slab_thk1.Value = 0
        sdb_slab_wid1.Value = 0
        sdb_slab_len1.Value = 0
        sdb_slab_wgt1.Value = 0
    End If
    
    For iRow = 1 To ss2.MaxRows
        ss2.Row = iRow
        ss2.Col = 0
        ss2.Text = ""
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, iRow, iRow)
        ss2.Col = 15
        If ss2.Text = "单定尺" Then
            ss2.Col = 14
            ss2.Lock = False
            Call Gp_Sp_CellColor(ss2, 14, iRow, , &HC0FFFF)
        End If
    Next iRow
            
End Sub

Private Sub cmd_ord1_Click()

    Dim sSeq As String
    Dim sQuery As String
    
'    If sdb_ord12_cnt.Value = 0 Then Exit Sub
    If sdb_ord11_cnt.Value = 0 Then Exit Sub
    
    If sdb_asroll_len.Value + sdb_ord1_len.Value >= lCool_max Then
        Call Gp_MsgBoxDisplay("母板长度 > " & lCool_max)
        Exit Sub
    End If
    
    sdb_ord12_cnt.Value = sdb_ord12_cnt.Value - 1
    iMplate_cnt = iMplate_cnt + 1
    cmd_mplate_del.Enabled = True
    cmd_mplate_design.Enabled = True
    cmd_mplate_complete.Enabled = False
    
    If iMplate_cnt < 10 Then
        sSeq = "0" & iMplate_cnt
    Else
        sSeq = Trim(Str(iMplate_cnt))
    End If
    
    sdb_asroll_len.Value = sdb_asroll_len.Value + sdb_ord1_len.Value
    
    Load lbl_mplate(iMplate_cnt)
    lbl_mplate(iMplate_cnt).Tag = "ord1"
    lbl_mplate(iMplate_cnt).Caption = sSeq
    lbl_mplate(iMplate_cnt).Top = 150
    lbl_mplate(iMplate_cnt).Height = 500
    lbl_mplate(iMplate_cnt).Width = (Shape1.Width / lCool_max) * sdb_ord1_len.Value
        
    If iMplate_cnt = 1 Then
        lbl_mplate(iMplate_cnt).Left = Shape1.Left
        lbl_mplate(iMplate_cnt).Visible = True
        
        Call Asroll_Thk(txt_ord_no1.Text)
        Call Asroll_Wid(txt_ord_no1.Text)
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ='00'
        Call Plate_Block_Seq_Create(iOrd1_Curr_Row, "I")
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd1_Curr_Row, sSeq, "I")
        
    Else
    
        If lbl_mplate(iMplate_cnt - 1).Caption <> "删除" Then
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width
        Else
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width - 30
        End If
        
        lbl_mplate(iMplate_cnt).Visible = True
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd1_Curr_Row, sSeq, "I")
        
    End If
    
End Sub

Private Sub cmd_ord2_Click()

    Dim sSeq As String
    Dim sQuery As String
    
'    If sdb_ord22_cnt.Value = 0 Then Exit Sub
    If sdb_ord21_cnt.Value = 0 Then Exit Sub
    
    If sdb_asroll_len.Value + sdb_ord2_len.Value >= lCool_max Then
        Call Gp_MsgBoxDisplay("母板长度 > " & lCool_max)
        Exit Sub
    End If
    
    sdb_ord22_cnt.Value = sdb_ord22_cnt.Value - 1
    iMplate_cnt = iMplate_cnt + 1
    cmd_mplate_del.Enabled = True
    cmd_mplate_design.Enabled = True
    cmd_mplate_complete.Enabled = False
    
    If iMplate_cnt < 10 Then
       sSeq = "0" & iMplate_cnt
    Else
       sSeq = Trim(Str(iMplate_cnt))
    End If
    
    sdb_asroll_len.Value = sdb_asroll_len.Value + sdb_ord2_len.Value
    
    Load lbl_mplate(iMplate_cnt)
    lbl_mplate(iMplate_cnt).Tag = "ord2"
    lbl_mplate(iMplate_cnt).Caption = sSeq
    lbl_mplate(iMplate_cnt).Top = 150
    lbl_mplate(iMplate_cnt).Height = 500
    lbl_mplate(iMplate_cnt).Width = (Shape1.Width / lCool_max) * sdb_ord2_len.Value
        
    If iMplate_cnt = 1 Then
        lbl_mplate(iMplate_cnt).Left = Shape1.Left
        lbl_mplate(iMplate_cnt).Visible = True
        
        Call Asroll_Thk(txt_ord_no2.Text)
        Call Asroll_Wid(txt_ord_no2.Text)
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ='00'
        Call Plate_Block_Seq_Create(iOrd2_Curr_Row, "I")
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd2_Curr_Row, sSeq, "I")
        
    Else
    
        If lbl_mplate(iMplate_cnt - 1).Caption <> "删除" Then
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width
        Else
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width - 30
        End If
        
        lbl_mplate(iMplate_cnt).Visible = True
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd2_Curr_Row, sSeq, "I")
        
    End If
    
End Sub

Private Sub cmd_ord3_Click()

    Dim sSeq As String
    Dim sQuery As String
    
'    If sdb_ord32_cnt.Value = 0 Then Exit Sub
    If sdb_ord31_cnt.Value = 0 Then Exit Sub
    
    If sdb_asroll_len.Value + sdb_ord3_len.Value >= lCool_max Then
        Call Gp_MsgBoxDisplay("母板长度 > " & lCool_max)
        Exit Sub
    End If
    
    sdb_ord32_cnt.Value = sdb_ord32_cnt.Value - 1
    iMplate_cnt = iMplate_cnt + 1
    cmd_mplate_del.Enabled = True
    cmd_mplate_design.Enabled = True
    cmd_mplate_complete.Enabled = False
    
    If iMplate_cnt < 10 Then
       sSeq = "0" & iMplate_cnt
    Else
       sSeq = Trim(Str(iMplate_cnt))
    End If
    
    sdb_asroll_len.Value = sdb_asroll_len.Value + sdb_ord3_len.Value
    
    Load lbl_mplate(iMplate_cnt)
    lbl_mplate(iMplate_cnt).Tag = "ord3"
    lbl_mplate(iMplate_cnt).Caption = sSeq
    lbl_mplate(iMplate_cnt).Top = 150
    lbl_mplate(iMplate_cnt).Height = 500
    lbl_mplate(iMplate_cnt).Width = (Shape1.Width / lCool_max) * sdb_ord3_len.Value
        
    If iMplate_cnt = 1 Then
        lbl_mplate(iMplate_cnt).Left = Shape1.Left
        lbl_mplate(iMplate_cnt).Visible = True
        
        Call Asroll_Thk(txt_ord_no3.Text)
        Call Asroll_Wid(txt_ord_no3.Text)
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ='00'
        Call Plate_Block_Seq_Create(iOrd3_Curr_Row, "I")
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd3_Curr_Row, sSeq, "I")
        
    Else
    
        If lbl_mplate(iMplate_cnt - 1).Caption <> "删除" Then
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width
        Else
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width - 30
        End If
        
        lbl_mplate(iMplate_cnt).Visible = True
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd3_Curr_Row, sSeq, "I")
        
    End If
    
End Sub

Private Sub cmd_ord4_Click()

    Dim sSeq As String
    Dim sQuery As String
    
'    If sdb_ord42_cnt.Value = 0 Then Exit Sub
    If sdb_ord41_cnt.Value = 0 Then Exit Sub
    
    If sdb_asroll_len.Value + sdb_ord4_len.Value >= lCool_max Then
        Call Gp_MsgBoxDisplay("母板长度 > " & lCool_max)
        Exit Sub
    End If
    
    sdb_ord42_cnt.Value = sdb_ord42_cnt.Value - 1
    iMplate_cnt = iMplate_cnt + 1
    cmd_mplate_del.Enabled = True
    cmd_mplate_design.Enabled = True
    cmd_mplate_complete.Enabled = False
    
    If iMplate_cnt < 10 Then
       sSeq = "0" & iMplate_cnt
    Else
       sSeq = Trim(Str(iMplate_cnt))
    End If
    
    sdb_asroll_len.Value = sdb_asroll_len.Value + sdb_ord4_len.Value
    
    Load lbl_mplate(iMplate_cnt)
    lbl_mplate(iMplate_cnt).Tag = "ord4"
    lbl_mplate(iMplate_cnt).Caption = sSeq
    lbl_mplate(iMplate_cnt).Top = 150
    lbl_mplate(iMplate_cnt).Height = 500
    lbl_mplate(iMplate_cnt).Width = (Shape1.Width / lCool_max) * sdb_ord4_len.Value
        
    If iMplate_cnt = 1 Then
        lbl_mplate(iMplate_cnt).Left = Shape1.Left
        lbl_mplate(iMplate_cnt).Visible = True
        
        Call Asroll_Thk(txt_ord_no4.Text)
        Call Asroll_Wid(txt_ord_no4.Text)
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ='00'
        Call Plate_Block_Seq_Create(iOrd4_Curr_Row, "I")
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd4_Curr_Row, sSeq, "I")
        
    Else
    
        If lbl_mplate(iMplate_cnt - 1).Caption <> "删除" Then
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width
        Else
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width - 30
        End If
        
        lbl_mplate(iMplate_cnt).Visible = True
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd4_Curr_Row, sSeq, "I")
        
    End If
    
End Sub

Private Sub cmd_ord5_Click()

    Dim sSeq As String
    Dim sQuery As String
    
'    If sdb_ord52_cnt.Value = 0 Then Exit Sub
    If sdb_ord51_cnt.Value = 0 Then Exit Sub
    
    If sdb_asroll_len.Value + sdb_ord5_len.Value >= lCool_max Then
        Call Gp_MsgBoxDisplay("母板长度 > " & lCool_max)
        Exit Sub
    End If
    
    sdb_ord52_cnt.Value = sdb_ord52_cnt.Value - 1
    iMplate_cnt = iMplate_cnt + 1
    cmd_mplate_del.Enabled = True
    cmd_mplate_design.Enabled = True
    cmd_mplate_complete.Enabled = False
    
    If iMplate_cnt < 10 Then
       sSeq = "0" & iMplate_cnt
    Else
       sSeq = Trim(Str(iMplate_cnt))
    End If
    
    sdb_asroll_len.Value = sdb_asroll_len.Value + sdb_ord5_len.Value
    
    Load lbl_mplate(iMplate_cnt)
    lbl_mplate(iMplate_cnt).Tag = "ord5"
    lbl_mplate(iMplate_cnt).Caption = sSeq
    lbl_mplate(iMplate_cnt).Top = 150
    lbl_mplate(iMplate_cnt).Height = 500
    lbl_mplate(iMplate_cnt).Width = (Shape1.Width / lCool_max) * sdb_ord5_len.Value
        
    If iMplate_cnt = 1 Then
        lbl_mplate(iMplate_cnt).Left = Shape1.Left
        lbl_mplate(iMplate_cnt).Visible = True
        
        Call Asroll_Thk(txt_ord_no5.Text)
        Call Asroll_Wid(txt_ord_no5.Text)
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ='00'
        Call Plate_Block_Seq_Create(iOrd5_Curr_Row, "I")
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd5_Curr_Row, sSeq, "I")
        
    Else
    
        If lbl_mplate(iMplate_cnt - 1).Caption <> "删除" Then
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width
        Else
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width - 30
        End If
        
        lbl_mplate(iMplate_cnt).Visible = True
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd5_Curr_Row, sSeq, "I")
        
    End If
    
End Sub

Private Sub cmd_ord6_Click()

    Dim sSeq As String
    Dim sQuery As String
    
'    If sdb_ord62_cnt.Value = 0 Then Exit Sub
    If sdb_ord61_cnt.Value = 0 Then Exit Sub
    
    If sdb_asroll_len.Value + sdb_ord6_len.Value >= lCool_max Then
        Call Gp_MsgBoxDisplay("母板长度 > " & lCool_max)
        Exit Sub
    End If
    
    sdb_ord62_cnt.Value = sdb_ord62_cnt.Value - 1
    iMplate_cnt = iMplate_cnt + 1
    cmd_mplate_del.Enabled = True
    cmd_mplate_design.Enabled = True
    cmd_mplate_complete.Enabled = False
    
    If iMplate_cnt < 10 Then
       sSeq = "0" & iMplate_cnt
    Else
       sSeq = Trim(Str(iMplate_cnt))
    End If
    
    sdb_asroll_len.Value = sdb_asroll_len.Value + sdb_ord6_len.Value
    
    Load lbl_mplate(iMplate_cnt)
    lbl_mplate(iMplate_cnt).Tag = "ord6"
    lbl_mplate(iMplate_cnt).Caption = sSeq
    lbl_mplate(iMplate_cnt).Top = 150
    lbl_mplate(iMplate_cnt).Height = 500
    lbl_mplate(iMplate_cnt).Width = (Shape1.Width / lCool_max) * sdb_ord6_len.Value
        
    If iMplate_cnt = 1 Then
        lbl_mplate(iMplate_cnt).Left = Shape1.Left
        lbl_mplate(iMplate_cnt).Visible = True
        
        Call Asroll_Thk(txt_ord_no6.Text)
        Call Asroll_Wid(txt_ord_no6.Text)
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ='00'
        Call Plate_Block_Seq_Create(iOrd6_Curr_Row, "I")
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd6_Curr_Row, sSeq, "I")
        
    Else
    
        If lbl_mplate(iMplate_cnt - 1).Caption <> "删除" Then
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width
        Else
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width - 30
        End If
        
        lbl_mplate(iMplate_cnt).Visible = True
        
        'EP_SLAB_EDT3_D INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd6_Curr_Row, sSeq, "I")
        
    End If
    
End Sub

Private Sub cmd_slab_complete_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iRow As Integer

    Dim adoCmd As ADODB.Command
    
    If sdb_slab_thk1.Value = 0 Then
        Call Gp_MsgBoxDisplay("板坯厚度必须输入", "I")
        Exit Sub
    End If
    
    If sdb_slab_wid1.Value = 0 Then
        Call Gp_MsgBoxDisplay("板坯宽度必须输入", "I")
        Exit Sub
    End If

    If sdb_slab_len1.Value = 0 Then
        Call Gp_MsgBoxDisplay("板坯长度必须输入", "I")
        Exit Sub
    End If

    If sdb_slab_wgt1.Value = 0 Then
        Call Gp_MsgBoxDisplay("板坯重量必须输入", "I")
        Exit Sub
    End If

    If sdb_slab_wgt.Value > sdb_slab_wgt1.Value Then
       Call Gp_MsgBoxDisplay("轧件重量 > 板坯重量..!!", "I")
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass

    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256

    sQuery = "{call CEG2065P ('9999999999','99'," & sdb_slab_thk1.Value & "," & sdb_slab_wid1.Value & ","
    sQuery = sQuery & sdb_slab_len1.Value & "," & sdb_slab_wgt1.Value & ",?)}"

    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command

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
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    Else

        iSlab_Complete = iSlab_Complete + 1
        sdb_slab_cnt.Value = sdb_slab_cnt.Value + 1
        iLastSlab_cnt = iSlab_cnt              'Complete Slab Count

        cmd_slab_design.Enabled = False
        cmd_slab_del.Enabled = False

        'Spread Sheet Refresh
        'Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1)

        If iOrd1_Curr_Row <> 0 Then
            sQuery = Gf_Sp_MakeQuery(sc2.Item("Spread"), sc2.Item("P-O"), "O", sc2.Item("pColumn"), iOrd1_Curr_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, sc2.Item("Spread"), iOrd1_Curr_Row)
        End If

        If iOrd2_Curr_Row <> 0 Then
            sQuery = Gf_Sp_MakeQuery(sc2.Item("Spread"), sc2.Item("P-O"), "O", sc2.Item("pColumn"), iOrd2_Curr_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, sc2.Item("Spread"), iOrd2_Curr_Row)
        End If

        If iOrd3_Curr_Row <> 0 Then
            sQuery = Gf_Sp_MakeQuery(sc2.Item("Spread"), sc2.Item("P-O"), "O", sc2.Item("pColumn"), iOrd3_Curr_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, sc2.Item("Spread"), iOrd3_Curr_Row)
        End If

        If iOrd4_Curr_Row <> 0 Then
            sQuery = Gf_Sp_MakeQuery(sc2.Item("Spread"), sc2.Item("P-O"), "O", sc2.Item("pColumn"), iOrd4_Curr_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, sc2.Item("Spread"), iOrd4_Curr_Row)
        End If

        If iOrd5_Curr_Row <> 0 Then
            sQuery = Gf_Sp_MakeQuery(sc2.Item("Spread"), sc2.Item("P-O"), "O", sc2.Item("pColumn"), iOrd5_Curr_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, sc2.Item("Spread"), iOrd5_Curr_Row)
        End If

        If iOrd6_Curr_Row <> 0 Then
            sQuery = Gf_Sp_MakeQuery(sc2.Item("Spread"), sc2.Item("P-O"), "O", sc2.Item("pColumn"), iOrd6_Curr_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, sc2.Item("Spread"), iOrd6_Curr_Row)
        End If

'        For iRow = 1 To ss1.MaxRows
'
'            If iOrd1_Curr_Row = iRow Or iOrd2_Curr_Row = iRow Or iOrd3_Curr_Row = iRow Then
'                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFFFF80)
'                ss1.Col = 0
'                ss1.Text = "选择"
'            End If
'
'        Next iRow

    End If

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)

End Sub

Private Sub cmd_slab_del_Click()

    Dim sSeq As String
    
    Dim iCount As Integer
    Dim iVisible_Cnt As Integer
    
    If iSlab_cnt = 0 Then Exit Sub
    
    For iCount = 1 To iSlab_cnt
        
        If lbl_slab(iCount).Caption = "删除" Then  'Delete
            
            If lbl_slab(iCount).Visible Then
            
                lbl_slab(iCount).Width = 0
                lbl_slab(iCount).Visible = False
                
                sdb_slab_len.Value = sdb_slab_len.Value - Val(lbl_slab(iCount).Tag)
                sdb_slab_wgt1.Value = 0
                sdb_asroll_prod_len.Value = 0
                
                If iCount < 10 Then
                    sSeq = "0" & iCount
                Else
                    sSeq = Trim(Str(iCount))
                End If
                
                'EP_SLAB_EDT3_D UPDATE  BLOCK_SEQ='01', SEQ      --> LEN = 0
                Call Slab_Seq_Create(sSeq, "U")
                    
            End If
            
            cmd_slab_complete.Enabled = False
            
        End If
    
        If iCount = 1 Then
            lbl_slab(iCount).Left = Shape4.Left
        Else
            If lbl_slab(iCount - 1).Caption <> "删除" Then
                If iCount Mod 3 = 1 Then
                    lbl_slab(iCount).Left = lbl_slab(iCount - 1).Left + lbl_slab(iCount - 1).Width - 10
                Else
                    lbl_slab(iCount).Left = lbl_slab(iCount - 1).Left + lbl_slab(iCount - 1).Width
                End If
            Else
                lbl_slab(iCount).Left = lbl_slab(iCount - 1).Left + lbl_slab(iCount - 1).Width - 30
            End If
        End If
    
    Next iCount
    
    iVisible_Cnt = 0
    For iCount = 1 To iSlab_cnt
    
        If lbl_slab(iCount).Visible Then
            iVisible_Cnt = iVisible_Cnt + 1
        End If
    
    Next iCount
    
    'EP_SLAB_EDT3_D DATA DELETE
    If iVisible_Cnt = 0 Then
   
        For iCount = 1 To iSlab_cnt
            Unload lbl_slab(iCount)
        Next iCount
    
        iSlab_Complete = 0
        iSlab_cnt = 0
        
        sdb_slab_wgt.Value = 0
        sdb_slab_ratio.Value = 0
        
        sdb_slab_len1.Value = 0
        sdb_slab_wgt1.Value = 0

        sdb_slab_len.Value = 0
        'sdb_slab_lenq.Value = 0
        sdb_asroll_prod_len.Value = 0
        
        Call Slab_Seq_Create("00", "D")
        cmd_slab_del.Enabled = False
        cmd_slab_design.Enabled = False
        cmd_design_modify.Enabled = False
        
        vENDUSE_CD = ""
        vSTLGRD = ""
        vPROD_THK = ""
        vPROD_WID = ""
        vORD_HCR_FL = ""
        vCR_CD = ""
        vORD_TRIM_FL = ""
        vUST_FL = ""
        
        If iMplate_cnt <= 0 Then
            sHTM_METH = ""
        End If
        
        If txt_ord_no1.Text = "" Then
            sdb_slab_thk1.Value = 0
            sdb_slab_wid1.Value = 0
        End If
   
    End If
    
End Sub

Private Sub cmd_slab_design_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim sSeq As String
    Dim iCount As Integer
    Dim iVisible_Cnt As Integer
    Dim dWgt As Double
    Dim P_SLAB_NO As String
    
    Dim AdoRs As ADODB.Recordset
    Dim adoCmd As ADODB.Command
    Set AdoRs = New ADODB.Recordset
    
    If sdb_slab_thk1.Value = 0 Then
        Call Gp_MsgBoxDisplay("板坯厚度必须输入", "I")
        Exit Sub
    End If
    
    If sdb_slab_wid1.Value = 0 Then
        Call Gp_MsgBoxDisplay("板坯宽度必须输入", "I")
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    sQuery = "{call CEG2063P ('9999999999','99','" + sUserID + "',?,?)}"
    
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
    If adoCmd("arg_e_code") <> "Y" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
        cmd_slab_complete.Enabled = False
        cmd_design_modify.Enabled = True
    Else
        cmd_slab_complete.Enabled = True
        cmd_design_modify.Enabled = True
        
        'Plate Redisplay
        sQuery = "         SELECT  THK, WID, LEN, WGT "
        sQuery = sQuery + "  FROM  NISCO.EP_SLAB_EDT3_D "
        sQuery = sQuery + " WHERE  SLAB_NO         =  '9999999999' "
        sQuery = sQuery + "   AND  SLAB_CUT_SEQ    =  '99' "
        sQuery = sQuery + "   AND  BLOCK_SEQ       =  '00' "
        sQuery = sQuery + "   AND  SEQ             =  '00' "
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
        Do Until AdoRs.EOF
            sdb_slab_thk.Value = Val(AdoRs.Fields(0) & "")
            sdb_slab_wid.Value = Val(AdoRs.Fields(1) & "")
            sdb_slab_len.Value = Val(AdoRs.Fields(2) & "")
            sdb_slab_wgt.Value = Val(AdoRs.Fields(3) & "")
            sdb_slab_wgt1.Value = Val(AdoRs.Fields(3) & "")
            AdoRs.MoveNext
        Loop
        AdoRs.Close
        
        Call Slab_Size("LEN")
        
        sQuery = "         SELECT  SUM(LEN) "
        sQuery = sQuery + "  FROM  NISCO.EP_SLAB_EDT3_D "
        sQuery = sQuery + " WHERE  SLAB_NO         =   '9999999999' "
        sQuery = sQuery + "   AND  SLAB_CUT_SEQ    =   '99' "
        sQuery = sQuery + "   AND  BLOCK_SEQ       NOT IN ('00','99') "
        sQuery = sQuery + "   AND  SEQ             =   '00' "
        sQuery = sQuery + "   AND  LEN             <>  0  "
        
        sdb_asroll_prod_len.Value = Gf_FloatFind(M_CN1, sQuery)
        
        'Slab_ratio
        sQuery = "         SELECT  SUM(WGT) "
        sQuery = sQuery + "  FROM  NISCO.EP_SLAB_EDT3_D "
        sQuery = sQuery + " WHERE  SLAB_NO         =   '9999999999' "
        sQuery = sQuery + "   AND  SLAB_CUT_SEQ    =   '99' "
        sQuery = sQuery + "   AND  BLOCK_SEQ       NOT IN ('00','99') "
        sQuery = sQuery + "   AND  SEQ             >   '00' "
        sQuery = sQuery + "   AND  LEN             <>  0  "
        
        dWgt = Gf_FloatFind(M_CN1, sQuery)
        sdb_slab_ratio.Value = (dWgt / sdb_slab_wgt1.Value) * 100
        
        'Plate Delete
        For iCount = 1 To iSlab_cnt
            Unload lbl_slab(iCount)
        Next iCount
    
        'Plate Redisplay
        sQuery = "         SELECT  BLOCK_SEQ, NVL(LEN,0) FROM NISCO.EP_SLAB_EDT3_D "
        sQuery = sQuery + " WHERE  SLAB_NO          =  '9999999999' "
        sQuery = sQuery + "   AND  SLAB_CUT_SEQ     =  '99' "
        sQuery = sQuery + "   AND  BLOCK_SEQ        NOT IN ('00','99') "
        sQuery = sQuery + "   AND  SEQ              =  '00' "
    
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
        iVisible_Cnt = 0
        
        If Not AdoRs.BOF And Not AdoRs.EOF Then
    
            While Not AdoRs.EOF
    
                iVisible_Cnt = iVisible_Cnt + 1
    
                Load lbl_slab(iVisible_Cnt)
                lbl_slab(iVisible_Cnt).Visible = True
                lbl_slab(iVisible_Cnt).BackColor = &HFF8080
                lbl_slab(iVisible_Cnt).ForeColor = &HFF0000
    
                lbl_slab(iVisible_Cnt).Caption = AdoRs.Fields(0)
    
                lbl_slab(iVisible_Cnt).Top = 180
                lbl_slab(iVisible_Cnt).Height = 500
    
                lbl_slab(iVisible_Cnt).Width = (Shape4.Width / sdb_slab_lenq.Value) * AdoRs.Fields(1)
                lbl_slab(iVisible_Cnt).Tag = Str(AdoRs.Fields(1))
    
                If iVisible_Cnt = 1 Then
                    lbl_slab(iVisible_Cnt).Left = Shape4.Left
                Else
                    If lbl_slab(iVisible_Cnt - 1).Caption <> "删除" Then
                        lbl_slab(iVisible_Cnt).Left = lbl_slab(iVisible_Cnt - 1).Left + lbl_slab(iVisible_Cnt - 1).Width
                    Else
                        lbl_slab(iVisible_Cnt).Left = lbl_slab(iVisible_Cnt - 1).Left + lbl_slab(iVisible_Cnt - 1).Width - 30
                    End If
    
                End If
                AdoRs.MoveNext
    
            Wend
    
        End If
    
        iSlab_cnt = iVisible_Cnt
        AdoRs.Close
        
    End If
    
    Set AdoRs = Nothing
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    AdoRs.Close
    Set AdoRs = Nothing
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

Private Sub cmd_slab_init_Click()

    Dim iCnt As Long
    Dim iRow As Integer
    
    For iCnt = 1 To iSlab_cnt
        lbl_slab(iCnt).Caption = "删除"
    Next iCnt

    Call cmd_slab_del_Click
    
    iSlab_cnt = 0
    sdb_slab_thk.Value = 0
    sdb_slab_wid.Value = 0
    sdb_slab_len.Value = 0
    sdb_slab_wgt.Value = 0
    sdb_slab_len1.Value = 0
    sdb_slab_wgt1.Value = 0
    sdb_slab_ratio.Value = 0
    sdb_slab_cnt.Value = 0
    sdb_asroll_prod_len.Value = 0
    vENDUSE_CD = ""
    vSTLGRD = ""
    vPROD_THK = ""
    vPROD_WID = ""
    vORD_HCR_FL = ""
    vCR_CD = ""
    vORD_TRIM_FL = ""
    vUST_FL = ""
    cmd_slab_del.Enabled = False
    cmd_slab_complete.Enabled = False
    cmd_design_modify.Enabled = False
    
    If iMplate_cnt <= 0 Then
        sHTM_METH = ""
    End If
    
    If txt_ord_no1.Text = "" Then
        sdb_slab_thk1.Value = 0
        sdb_slab_wid1.Value = 0
    End If
    
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
    
    If Mid(sAuthority, 3, 1) <> "1" Then
        cmd_ord1.Enabled = False
        cmd_ord2.Enabled = False
        cmd_ord3.Enabled = False
        cmd_ord4.Enabled = False
        cmd_ord5.Enabled = False
        cmd_ord6.Enabled = False
    End If

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    
    lCool_max = Gf_FloatFind(M_CN1, "SELECT MAXI FROM EP_SLABDESIGN WHERE PLT = 'C3' AND APLY_ITEM = 'SLABDESIGN008' AND PRC_LINE = '1'")
    
    If lCool_max = 0 Then
        Label4.Caption = "0(M)"
    Else
        Label4.Caption = lCool_max / 1000 & "(M)"
    End If
    
    lAsroll_max = Gf_FloatFind(M_CN1, "SELECT MAXI FROM EP_SLABDESIGN WHERE PLT = 'C3' AND APLY_ITEM = 'SLABDESIGN003' AND PRC_LINE = '1'")
    
    If lAsroll_max = 0 Then
        Label5.Caption = "0(M)"
    Else
        Label5.Caption = lAsroll_max / 1000 & "(M)"
    End If
    
    sdb_slab_lenq.Value = lAsroll_max

    sdb_thk_to.Value = 9999.99
    sdb_wid_to.Value = 9999.99
    sdb_len_to.Value = 99999.9
    
    Call cmd_slab_del_Click
    Call cmd_mplate_del_Click
            
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim iCount As Integer
    
    If iMplate_cnt > 0 Then
        Call Gp_MsgBoxDisplay("Must plate data clear necessarily")
        Cancel = 1
        Exit Sub
    End If
    
    If iSlab_cnt > 0 Then
        Call Gp_MsgBoxDisplay("Must moplate data clear necessarily")
        Cancel = 1
        Exit Sub
    End If
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    
    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing
    
    Set pContro2 = Nothing
    Set nContro2 = Nothing
    Set iContro2 = Nothing
    Set rContro2 = Nothing
    Set cContro2 = Nothing
    Set aContro2 = Nothing
    Set lContro2 = Nothing
    Set mContro2 = Nothing
    
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

End Sub

Public Sub Form_Cls()
    
    Dim iCnt As Long
    
    If iMplate_cnt > 0 Then
        Call Gp_MsgBoxDisplay("Must plate data clear necessarily")
        Exit Sub
    End If
    
    If iSlab_cnt > 0 Then
        Call Gp_MsgBoxDisplay("Must moplate data clear necessarily")
        Exit Sub
    End If
    
    If Gf_Sp_Cls(sc2) Then
    
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        
        ss2.SetFocus
        
        sdb_thk_to.Value = 9999.99
        sdb_wid_to.Value = 9999.99
        sdb_len_to.Value = 99999.9
        
        oRd_cnt = 0
        iSlab_Row = 0
        iOrd1_Curr_Row = 0
        iOrd2_Curr_Row = 0
        iOrd3_Curr_Row = 0
        iOrd4_Curr_Row = 0
        iOrd5_Curr_Row = 0
        iOrd6_Curr_Row = 0
        iSLAB_EDT_SEQ = 0
        cmd_mplate_del.Enabled = False
        
        For iCnt = 1 To iMplate_cnt
            Unload lbl_mplate(iCnt)
        Next iCnt
        
        For iCnt = 1 To iSlab_cnt
            Unload lbl_slab(iCnt)
        Next iCnt
        
        iMplate_cnt = 0
        iSlab_cnt = 0
        sHTM_METH = ""
        
    End If

End Sub

Public Sub Form_Ref()

    Dim sQuery As String
    Dim dValue As String
    Dim iCnt As Long
    Dim iRow As Long
    
    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
    
    'EP_PLATE_EDT DATA DELETE
    If iMplate_cnt > 0 Then
        Call Gp_MsgBoxDisplay("Must plate data clear necessarily")
        Exit Sub
    End If
        
    'EP_PLATE_EDT DATA DELETE
    If iSlab_cnt > 0 Then
        Call Gp_MsgBoxDisplay("Must moslab data clear necessarily")
        Exit Sub
    End If
    
    
    If Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl")) Then
        'Call Gp_Sp_EvenRowBackcolor(Proc_Sc("Sc").Item("Spread"))
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss2.OperationMode = OperationModeNormal
        txt_ord_no1.Text = ""
        txt_ord_no2.Text = ""
        txt_ord_no3.Text = ""
        txt_ord_no4.Text = ""
        txt_ord_no5.Text = ""
        txt_ord_no6.Text = ""
        sdb_ord11_cnt.Value = 0
        sdb_ord12_cnt.Value = 0
        sdb_ord21_cnt.Value = 0
        sdb_ord22_cnt.Value = 0
        sdb_ord31_cnt.Value = 0
        sdb_ord32_cnt.Value = 0
        sdb_ord41_cnt.Value = 0
        sdb_ord42_cnt.Value = 0
        sdb_ord51_cnt.Value = 0
        sdb_ord52_cnt.Value = 0
        sdb_ord61_cnt.Value = 0
        sdb_ord62_cnt.Value = 0
        sdb_ord1_len.Value = 0
        sdb_ord2_len.Value = 0
        sdb_ord3_len.Value = 0
        sdb_ord4_len.Value = 0
        sdb_ord5_len.Value = 0
        sdb_ord6_len.Value = 0
        sdb_asroll_thk.Value = 0
        sdb_asroll_wid.Value = 0
        sdb_asroll_len.Value = 0
        sdb_asroll_prod_len.Value = 0
        sdb_slab_ratio.Value = 0
        sdb_slab_thk1.Value = 0
        sdb_slab_wid1.Value = 0
        sdb_slab_len1.Value = 0
        sdb_slab_wgt1.Value = 0
        
        iSlab_Row = 0
        iOrd1_Curr_Row = 0
        iOrd2_Curr_Row = 0
        iOrd3_Curr_Row = 0
        iOrd4_Curr_Row = 0
        iOrd5_Curr_Row = 0
        iOrd6_Curr_Row = 0
        lMain_row = 0
        oRd_cnt = 0
        
        For iCnt = 1 To iMplate_cnt
            Unload lbl_mplate(iCnt)
        Next iCnt
        
        For iCnt = 1 To iSlab_cnt
            Unload lbl_slab(iCnt)
        Next iCnt
        
        iMplate_cnt = 0
        iSlab_cnt = 0
        iSlab_Complete = 0
        
        iSLAB_EDT_SEQ = 0
        cmd_mplate_del.Enabled = False
        cmd_slab_del.Enabled = False
        
        For iRow = 1 To ss2.MaxRows
            ss2.Row = iRow
            ss2.Col = 15
            If ss2.Text = "单定尺" Then
                ss2.Col = 14
                ss2.Lock = False
                Call Gp_Sp_CellColor(ss2, 14, iRow, , &HC0FFFF)
            End If
        Next iRow
        
    End If
            
End Sub

Public Sub Form_Pro()

End Sub

Public Sub Form_Ins()
    
End Sub

Public Sub Spread_Cpy()

End Sub

Public Sub Spread_Pst()

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
    
End Sub

Private Sub lbl_mplate_DblClick(Index As Integer)

    Dim sSeq As String
    
    If Index < 10 Then
        sSeq = "0" & Index
    Else
        sSeq = Trim(Str(Index))
    End If
    
    If lbl_mplate(Index).BackColor = &HE0E0E0 Then
        lbl_mplate(Index).BackColor = &HC0C0C0
        lbl_mplate(Index).ForeColor = &HFF0000
        lbl_mplate(Index).Caption = sSeq
        
        If lbl_mplate(Index).Tag = "ord1" Then
            sdb_ord12_cnt = sdb_ord12_cnt - 1
        ElseIf lbl_mplate(Index).Tag = "ord2" Then
            sdb_ord22_cnt = sdb_ord22_cnt - 1
        ElseIf lbl_mplate(Index).Tag = "ord3" Then
            sdb_ord32_cnt = sdb_ord32_cnt - 1
        ElseIf lbl_mplate(Index).Tag = "ord4" Then
            sdb_ord42_cnt = sdb_ord42_cnt - 1
        ElseIf lbl_mplate(Index).Tag = "ord5" Then
            sdb_ord52_cnt = sdb_ord52_cnt - 1
        Else
            sdb_ord62_cnt = sdb_ord62_cnt - 1
        End If
    Else
        lbl_mplate(Index).BackColor = &HE0E0E0
        lbl_mplate(Index).ForeColor = &HFF0000
        lbl_mplate(Index).Caption = "删除"

        If lbl_mplate(Index).Tag = "ord1" Then
            sdb_ord12_cnt = sdb_ord12_cnt + 1
        ElseIf lbl_mplate(Index).Tag = "ord2" Then
            sdb_ord22_cnt = sdb_ord22_cnt + 1
        ElseIf lbl_mplate(Index).Tag = "ord3" Then
            sdb_ord32_cnt = sdb_ord32_cnt + 1
        ElseIf lbl_mplate(Index).Tag = "ord4" Then
            sdb_ord42_cnt = sdb_ord42_cnt + 1
        ElseIf lbl_mplate(Index).Tag = "ord5" Then
            sdb_ord52_cnt = sdb_ord52_cnt + 1
        Else
            sdb_ord62_cnt = sdb_ord62_cnt + 1
        End If
        
    End If
    
End Sub

Private Sub lbl_slab_DblClick(Index As Integer)

    Dim sSeq As String
    
    If Index < 10 Then
        sSeq = "0" & Index
    Else
        sSeq = Trim(Str(Index))
    End If
    
    If lbl_slab(Index).BackColor = &HFFC0C0 Then
        lbl_slab(Index).BackColor = &HFF8080
        lbl_slab(Index).ForeColor = &HFF0000
        lbl_slab(Index).Caption = sSeq
    Else
        lbl_slab(Index).BackColor = &HFFC0C0
        lbl_slab(Index).ForeColor = &HFF0000
        lbl_slab(Index).Caption = "删除"
    End If
    
End Sub

Private Sub sdb_slab_len1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String
    Dim dWgt As Double
    
    Call Slab_Size("WGT")
    
    If sdb_slab_wgt1.Value = 0 Then Exit Sub
    
    'Slab_ratio
    sQuery = "         SELECT  SUM(WGT) "
    sQuery = sQuery + "  FROM  NISCO.EP_SLAB_EDT3_D "
    sQuery = sQuery + " WHERE  SLAB_NO         =   '9999999999' "
    sQuery = sQuery + "   AND  SLAB_CUT_SEQ    =   '99' "
    sQuery = sQuery + "   AND  BLOCK_SEQ       NOT IN ('00','99') "
    sQuery = sQuery + "   AND  SEQ             >   '00' "
    sQuery = sQuery + "   AND  LEN             <>  0  "
    
    dWgt = Gf_FloatFind(M_CN1, sQuery)
    sdb_slab_ratio.Value = (dWgt / sdb_slab_wgt1.Value) * 100
    
End Sub

Private Sub sdb_slab_thk1_KeyUp(KeyCode As Integer, Shift As Integer)

    Call Slab_Size("LEN")
    
End Sub

Private Sub sdb_slab_wid1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String
    Dim dWgt As Double
    
    Call Slab_Size("LEN")
    
    If sdb_slab_wgt1.Value = 0 Then Exit Sub
    
    'Slab_ratio
    sQuery = "         SELECT  SUM(WGT) "
    sQuery = sQuery + "  FROM  NISCO.EP_SLAB_EDT3_D "
    sQuery = sQuery + " WHERE  SLAB_NO         =   '9999999999' "
    sQuery = sQuery + "   AND  SLAB_CUT_SEQ    =   '99' "
    sQuery = sQuery + "   AND  BLOCK_SEQ       NOT IN ('00','99') "
    sQuery = sQuery + "   AND  SEQ             >   '00' "
    sQuery = sQuery + "   AND  LEN             <>  0  "
    
    dWgt = Gf_FloatFind(M_CN1, sQuery)
    sdb_slab_ratio.Value = (dWgt / sdb_slab_wgt1.Value) * 100
    
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim sTemp_ord As String
    Dim sOrd_No As String
    Dim sOrd_item As String
    Dim iRow As Integer
    Dim iCnt As Long
    Dim dWgt As Double
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If ss2.MaxRows < 1 Or Row < 1 Then Exit Sub
    If Col <> 0 Then Exit Sub
    
    ss2.Row = Row
    
    ss2.Col = 0
    
    If ss2.Text <> "选择" Then
    
        If oRd_cnt = 6 Then Exit Sub
        
        If txt_ord_no1.Text = "" Then
        
            ss2.Row = Row
            
            If iSlab_cnt > 0 Then
                If First_Condition_Compare(Row) = False Then Exit Sub
            Else
                ss2.Col = 20
                sHTM_METH = ss2.Text
                ss2.Col = 1
                sOrd_No = ss2.Text
                ss2.Col = 2
                sOrd_item = ss2.Text
                sdb_slab_thk1.Value = Gf_FloatFind(M_CN1, "SELECT  SLAB_THK FROM  NISCO.CP_ORD_SL_D " & _
                                                           "WHERE  ORD_NO    = '" & sOrd_No & "' " & _
                                                           "  AND  ORD_ITEM  = '" & sOrd_item & "' " & _
                                                           "  AND  CNF_EMP   = 'SYSTEM' ")
                sdb_slab_wid1.Value = Gf_FloatFind(M_CN1, "SELECT  SLAB_WID FROM  NISCO.CP_ORD_SL_D " & _
                                                           "WHERE  ORD_NO    = '" & sOrd_No & "' " & _
                                                           "  AND  ORD_ITEM  = '" & sOrd_item & "' " & _
                                                           "  AND  CNF_EMP   = 'SYSTEM' ")
            End If
            
            ss2.Row = Row
            ss2.Col = 0
            ss2.Text = "选择"
            ss2.Col = 1
            txt_ord_no1.Text = ss2.Text
            ss2.Col = 2
            txt_ord_no1.Text = txt_ord_no1.Text & "-" & ss2.Text
            
            'PROD_THK
            'ss1.Col = 8
            'sdb_asroll_thk.Value = ss1.Value
            
            'PROD_WID
            'ss1.Col = 11
            'sdb_asroll_wid.Value = ss1.Value
            
            'PROD_LEN
            ss2.Col = 14
            sdb_ord1_len.Value = ss2.Value
            
            'PROD_WGT
            ss2.Col = 18
            dWgt = ss2.Value
            
            'DESIGN_REM_WGT / PROD_WGT
            ss2.Col = 28
            sdb_ord11_cnt.Value = Round((ss2.Value / dWgt) + 0.5)
            sdb_ord12_cnt.Value = Round((ss2.Value / dWgt) + 0.5)
            
            lMain_row = Row
            
            'Select Order1 Row
            iOrd1_Curr_Row = Row
            
        ElseIf txt_ord_no2.Text = "" Then
        
            If Condition_Compare(Row) = False Then Exit Sub
            
            ss2.Row = Row
            ss2.Col = 0
            ss2.Text = "选择"
            ss2.Col = 1
            txt_ord_no2.Text = ss2.Text
            ss2.Col = 2
            txt_ord_no2.Text = txt_ord_no2.Text & "-" & ss2.Text
            
            'PROD_LEN
            ss2.Col = 14
            sdb_ord2_len.Value = ss2.Value
            
            'PROD_WGT
            ss2.Col = 18
            dWgt = ss2.Value
            
            'DESIGN_REM_WGT / PROD_WGT
            ss2.Col = 28
            sdb_ord21_cnt.Value = Round((ss2.Value / dWgt) + 0.5)
            sdb_ord22_cnt.Value = Round((ss2.Value / dWgt) + 0.5)
            
            'Select Order2 Row
            iOrd2_Curr_Row = Row
        
        ElseIf txt_ord_no3.Text = "" Then
        
            If Condition_Compare(Row) = False Then Exit Sub
            
            ss2.Row = Row
            ss2.Col = 0
            ss2.Text = "选择"
            ss2.Col = 1
            txt_ord_no3.Text = ss2.Text
            ss2.Col = 2
            txt_ord_no3.Text = txt_ord_no3.Text & "-" & ss2.Text
            
            'PROD_LEN
            ss2.Col = 14
            sdb_ord3_len.Value = ss2.Value
            
            'PROD_WGT
            ss2.Col = 18
            dWgt = ss2.Value
            
            'DESIGN_REM_WGT / PROD_WGT
            ss2.Col = 28
            sdb_ord31_cnt.Value = Round((ss2.Value / dWgt) + 0.5)
            sdb_ord32_cnt.Value = Round((ss2.Value / dWgt) + 0.5)
            
            'Select Order3 Row
            iOrd3_Curr_Row = Row
            
        ElseIf txt_ord_no4.Text = "" Then
        
            If Condition_Compare(Row) = False Then Exit Sub
            
            ss2.Row = Row
            ss2.Col = 0
            ss2.Text = "选择"
            ss2.Col = 1
            txt_ord_no4.Text = ss2.Text
            ss2.Col = 2
            txt_ord_no4.Text = txt_ord_no4.Text & "-" & ss2.Text
            
            'PROD_LEN
            ss2.Col = 14
            sdb_ord4_len.Value = ss2.Value
            
            'PROD_WGT
            ss2.Col = 18
            dWgt = ss2.Value
            
            'DESIGN_REM_WGT / PROD_WGT
            ss2.Col = 28
            sdb_ord41_cnt.Value = Round((ss2.Value / dWgt) + 0.5)
            sdb_ord42_cnt.Value = Round((ss2.Value / dWgt) + 0.5)
            
            'Select Order4 Row
            iOrd4_Curr_Row = Row
            
        ElseIf txt_ord_no5.Text = "" Then
        
            If Condition_Compare(Row) = False Then Exit Sub
            
            ss2.Row = Row
            ss2.Col = 0
            ss2.Text = "选择"
            ss2.Col = 1
            txt_ord_no5.Text = ss2.Text
            ss2.Col = 2
            txt_ord_no5.Text = txt_ord_no5.Text & "-" & ss2.Text
            
            'PROD_LEN
            ss2.Col = 14
            sdb_ord5_len.Value = ss2.Value
            
            'PROD_WGT
            ss2.Col = 18
            dWgt = ss2.Value
            
            'DESIGN_REM_WGT / PROD_WGT
            ss2.Col = 28
            sdb_ord51_cnt.Value = Round((ss2.Value / dWgt) + 0.5)
            sdb_ord52_cnt.Value = Round((ss2.Value / dWgt) + 0.5)
            
            'Select Order5 Row
            iOrd5_Curr_Row = Row
            
        Else
        
            If Condition_Compare(Row) = False Then Exit Sub
            
            ss2.Row = Row
            ss2.Col = 0
            ss2.Text = "选择"
            ss2.Col = 1
            txt_ord_no6.Text = ss2.Text
            ss2.Col = 2
            txt_ord_no6.Text = txt_ord_no6.Text & "-" & ss2.Text
            
            'PROD_LEN
            ss2.Col = 14
            sdb_ord6_len.Value = ss2.Value
            
            'PROD_WGT
            ss2.Col = 18
            dWgt = ss2.Value
            
            'DESIGN_REM_WGT / PROD_WGT
            ss2.Col = 28
            sdb_ord61_cnt.Value = Round((ss2.Value / dWgt) + 0.5)
            sdb_ord62_cnt.Value = Round((ss2.Value / dWgt) + 0.5)
            
            'Select Order6 Row
            iOrd6_Curr_Row = Row
            
        End If
        
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, Row, Row, , &HFFFF80)
        oRd_cnt = oRd_cnt + 1
        ss2.Col = 14
        ss2.Lock = True
    
    Else
    
        If iMplate_cnt > 0 Then Exit Sub
    
        ss2.Text = ""
        
        ss2.Col = 1
        sTemp_ord = ss2.Text
        ss2.Col = 2
        sTemp_ord = sTemp_ord & "-" & ss2.Text
        
        If txt_ord_no1.Text = sTemp_ord Then
            txt_ord_no1.Text = ""
            txt_ord_no2.Text = ""
            txt_ord_no3.Text = ""
            txt_ord_no4.Text = ""
            txt_ord_no5.Text = ""
            txt_ord_no6.Text = ""
            sdb_ord11_cnt.Value = 0
            sdb_ord12_cnt.Value = 0
            sdb_ord21_cnt.Value = 0
            sdb_ord22_cnt.Value = 0
            sdb_ord31_cnt.Value = 0
            sdb_ord32_cnt.Value = 0
            sdb_ord41_cnt.Value = 0
            sdb_ord42_cnt.Value = 0
            sdb_ord51_cnt.Value = 0
            sdb_ord52_cnt.Value = 0
            sdb_ord61_cnt.Value = 0
            sdb_ord62_cnt.Value = 0
            sdb_ord1_len.Value = 0
            sdb_ord2_len.Value = 0
            sdb_ord3_len.Value = 0
            sdb_ord4_len.Value = 0
            sdb_ord5_len.Value = 0
            sdb_ord6_len.Value = 0
            oRd_cnt = 1
            lMain_row = 0
            iOrd1_Curr_Row = 0
            iOrd2_Curr_Row = 0
            iOrd3_Curr_Row = 0
            iOrd4_Curr_Row = 0
            iOrd5_Curr_Row = 0
            iOrd6_Curr_Row = 0
            
            For iRow = 1 To ss2.MaxRows
                ss2.Row = iRow
                ss2.Col = 0
                ss2.Text = ""
                Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, iRow, iRow)
                ss2.Row = iRow
                ss2.Col = 15
                If ss2.Text = "单定尺" Then
                    ss2.Col = 14
                    ss2.Lock = False
                    Call Gp_Sp_CellColor(ss2, 14, iRow, , &HC0FFFF)
                End If
            Next iRow
            
            For iCnt = 1 To iMplate_cnt
                Unload lbl_mplate(iCnt)
            Next iCnt
            iMplate_cnt = 0
            
        ElseIf txt_ord_no2.Text = sTemp_ord Then
            txt_ord_no2.Text = ""
            sdb_ord21_cnt.Value = 0
            sdb_ord22_cnt.Value = 0
            sdb_ord2_len.Value = 0
            iOrd2_Curr_Row = 0
        
        ElseIf txt_ord_no3.Text = sTemp_ord Then
            txt_ord_no3.Text = ""
            sdb_ord31_cnt.Value = 0
            sdb_ord32_cnt.Value = 0
            sdb_ord3_len.Value = 0
            iOrd3_Curr_Row = 0
            
        ElseIf txt_ord_no4.Text = sTemp_ord Then
            txt_ord_no4.Text = ""
            sdb_ord41_cnt.Value = 0
            sdb_ord42_cnt.Value = 0
            sdb_ord4_len.Value = 0
            iOrd4_Curr_Row = 0
            
        ElseIf txt_ord_no5.Text = sTemp_ord Then
            txt_ord_no5.Text = ""
            sdb_ord51_cnt.Value = 0
            sdb_ord52_cnt.Value = 0
            sdb_ord5_len.Value = 0
            iOrd5_Curr_Row = 0
            
        Else
            txt_ord_no6.Text = ""
            sdb_ord61_cnt.Value = 0
            sdb_ord62_cnt.Value = 0
            sdb_ord6_len.Value = 0
            iOrd6_Curr_Row = 0
            
        End If
            
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, Row, Row)
        oRd_cnt = oRd_cnt - 1
        
        ss2.Row = Row
        ss2.Col = 15
        If ss2.Text = "单定尺" Then
            ss2.Col = 14
            ss2.Lock = False
            Call Gp_Sp_CellColor(ss2, 14, Row, , &HC0FFFF)
        End If
    
    End If
        
End Sub

Private Function Condition_Compare(iRow As Long) As Boolean

    Dim sTemp   As String
    Dim dTemp   As Double
    Dim dWidMin As Double
    Dim dWidMax As Double
    Dim dThkMin As Double
    Dim dThkMax As Double
    
    Condition_Compare = True
    
    'STLGRD
    ss2.Row = lMain_row
    ss2.Col = 5
    sTemp = ss2.Text
    ss2.Row = iRow
    
    If sTemp <> ss2.Text Then
        Call Gp_MsgBoxDisplay("钢种不一致")
        Condition_Compare = False
        Exit Function
    End If
    
    'PROD_THK
    ss2.Row = lMain_row
    ss2.Col = 8
    dTemp = ss2.Value
    
    ss2.Row = iRow
    ss2.Col = 9
    dThkMin = ss2.Value
    ss2.Col = 10
    dThkMax = ss2.Value
    
    If dTemp < dThkMin Or dTemp > dThkMax Then
        Call Gp_MsgBoxDisplay("厚度不一致")
        Condition_Compare = False
        Exit Function
    End If
    
    'PROD_WID
    'Call Range_Wid(iRow, dWidMin, dWidMax)
    
    ss2.Row = lMain_row
    ss2.Col = 11
    dTemp = ss2.Value
    
    ss2.Row = iRow
    ss2.Col = 12
    dWidMin = ss2.Value
    ss2.Col = 13
    dWidMax = ss2.Value
    
'    ss2.Row = iRow
    If dTemp < dWidMin Or dTemp > dWidMax Then
        Call Gp_MsgBoxDisplay("宽度不一致")
        Condition_Compare = False
        Exit Function
    End If
    
    'ORD_TRIM_FL
    ss2.Row = lMain_row
    ss2.Col = 22
    sTemp = ss2.Text
    ss2.Row = iRow

    If sTemp <> ss2.Text Then
        Call Gp_MsgBoxDisplay("切边不一致")
        Condition_Compare = False
        Exit Function
    End If

    'HTM_METH
    ss2.Row = iRow
    ss2.Col = 20
    
    If sHTM_METH = "" Then
        If ss2.Text <> "" Then
            Call Gp_MsgBoxDisplay("热处理不一致")
            Condition_Compare = False
            Exit Function
        End If
    Else
        If ss2.Text = "" Then
            Call Gp_MsgBoxDisplay("热处理不一致")
            Condition_Compare = False
            Exit Function
        End If
    End If

'
'    'ENDUSE_CD
'    ss2.Row = lMain_row
'    ss2.Col = 7
'    sTemp = ss2.Text
'    ss2.Row = iRow
'
'    If sTemp <> ss2.Text Then
'        Call Gp_MsgBoxDisplay("用途不一致")
'        Condition_Compare = False
'        Exit Function
'    End If
'
'    'ORD_HCR_FL
'    ss2.Row = lMain_row
'    ss2.Col = 18
'    sTemp = ss2.Text
'    ss2.Row = iRow
'
'    If sTemp <> ss2.Text Then
'        Call Gp_MsgBoxDisplay("H/C 不一致")
'        Condition_Compare = False
'        Exit Function
'    End If
'
'    'CR_CD
'    ss2.Row = lMain_row
'    ss2.Col = 19
'    sTemp = ss2.Text
'    ss2.Row = iRow
'
'    If sTemp <> ss2.Text Then
'        Call Gp_MsgBoxDisplay("控轧不一致")
'        Condition_Compare = False
'        Exit Function
'    End If
'
'    'UST_FL
'    ss2.Row = lMain_row
'    ss2.Col = 21
'    sTemp = ss2.Text
'    ss2.Row = iRow
'
'    If sTemp <> ss2.Text Then
'        Call Gp_MsgBoxDisplay("UST 不一致")
'        Condition_Compare = False
'        Exit Function
'    End If

End Function

Private Function First_Condition_Compare(iRow As Long) As Boolean

    Dim sTemp   As String
    Dim dTemp   As Double
    Dim dWidMin As Double
    Dim dWidMax As Double
    Dim dThkMin As Double
    Dim dThkMax As Double
    
    First_Condition_Compare = True
    ss2.Row = iRow
    
    'STLGRD
    ss2.Col = 5
    If vSTLGRD <> ss2.Text Then
        Call Gp_MsgBoxDisplay("钢种不一致")
        First_Condition_Compare = False
        Exit Function
    End If
    
    'PROD_THK
    ss2.Col = 9
    dThkMin = ss2.Value
    ss2.Col = 10
    dThkMax = ss2.Value
    
    If vPROD_THK < dThkMin Or vPROD_THK > dThkMax Then
        Call Gp_MsgBoxDisplay("厚度不一致")
        First_Condition_Compare = False
        Exit Function
    End If
    
    'PROD_WID
    'Call Range_Wid(iRow, dWidMin, dWidMax)
'    ss2.Col = lMain_row
'    ss2.Col = 12
'    dWidMin = ss2.Value
'    ss2.Col = 13
'    dWidMax = ss2.Value
'
'    If vPROD_WID < dWidMin Or vPROD_WID > dWidMax Then
'        Call Gp_MsgBoxDisplay("宽度不一致")
'        First_Condition_Compare = False
'        Exit Function
'    End If
    
    'ORD_TRIM_FL
    ss2.Col = 22
    If vORD_TRIM_FL <> ss2.Text Then
        Call Gp_MsgBoxDisplay("切边不一致")
        First_Condition_Compare = False
        Exit Function
    End If
    
    'HTM_METH
    ss2.Col = 20
    If sHTM_METH = "" Then
        If ss2.Text <> "" Then
            Call Gp_MsgBoxDisplay("热处理不一致")
            First_Condition_Compare = False
            Exit Function
        End If
    Else
        If ss2.Text = "" Then
            Call Gp_MsgBoxDisplay("热处理不一致")
            First_Condition_Compare = False
            Exit Function
        End If
    End If

'    'ENDUSE_CD
'    ss2.Col = 7
'    If vENDUSE_CD <> ss2.Text Then
'        Call Gp_MsgBoxDisplay("用途不一致")
'        First_Condition_Compare = False
'        Exit Function
'    End If
'
'    'ORD_HCR_FL
'    ss2.Col = 18
'    If vORD_HCR_FL <> ss2.Text Then
'        Call Gp_MsgBoxDisplay("H/C 不一致")
'        First_Condition_Compare = False
'        Exit Function
'    End If
'
'    'CR_CD
'    ss2.Col = 19
'    If vCR_CD <> ss2.Text Then
'        Call Gp_MsgBoxDisplay("控轧不一致")
'        First_Condition_Compare = False
'        Exit Function
'    End If
'
'    'UST_FL
'    ss2.Col = 21
'    If vUST_FL <> ss2.Text Then
'        Call Gp_MsgBoxDisplay("UST 不一致")
'        First_Condition_Compare = False
'        Exit Function
'    End If

End Function

Private Sub Plate_Block_Seq_Create(Current_Row As Variant, iType As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    ss2.Row = Current_Row
    
    'SLAB_NO, BLOCK_SEQ, SEQ
    sQuery = "{call CEG2060C.P_MODIFY1 ('" + iType + "','9999999999','99','00',"
    
    'ORD_NO
    ss2.Col = 1
    sQuery = sQuery + "'" + ss2.Text + "',"
    
    'ORD_ITEM
    ss2.Col = 2
    sQuery = sQuery + "'" + ss2.Text + "',"
    
    'PROD_CD
    sQuery = sQuery + "'PP',"
        
    'STLGRD
    ss2.Col = 5
    sQuery = sQuery + "'" + ss2.Text + "',"
    
    'THK
    ss2.Col = 8
    sQuery = sQuery & ss2.Value & ","
    
    'WID
    ss2.Col = 11
    sQuery = sQuery & ss2.Value & ","
    
    'LEN
    ss2.Col = 14
    sQuery = sQuery & ss2.Value & ","
    
    'WGT
    ss2.Col = 18
    sQuery = sQuery & ss2.Value & ","
    
    'CR_CD
    ss2.Col = 21
    sQuery = sQuery + "'" + ss2.Text + "',"
    
    'UST_FL
    ss2.Col = 23
    sQuery = sQuery + "'" + ss2.Text + "',"
    
    'TRIM_FL
    ss2.Col = 22
    sQuery = sQuery + "'" + ss2.Text + "',?,?)}"
    
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
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)

End Sub

Private Sub Plate_Seq_Create(Current_Row As Variant, Seq As String, iType As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    ss2.Row = Current_Row
    
    'SLAB_NO, BLOCK_SEQ, SEQ
    sQuery = "{call CEG2060C.P_MODIFY1 ('" + iType + "','9999999999','99','" + Seq + "',"
    
    'ORD_NO
    ss2.Col = 1
    sQuery = sQuery + "'" + ss2.Text + "',"
    
    'ORD_ITEM
    ss2.Col = 2
    sQuery = sQuery + "'" + ss2.Text + "',"
    
    'PROD_CD
    sQuery = sQuery + "'PP',"
        
    'STLGRD
    ss2.Col = 5
    sQuery = sQuery + "'" + ss2.Text + "',"
    
    'THK
    ss2.Col = 8
    sQuery = sQuery & ss2.Value & ","
    
    'WID
    ss2.Col = 11
    sQuery = sQuery & ss2.Value & ","
    
    'LEN
    ss2.Col = 14
    sQuery = sQuery & ss2.Value & ","
    
    'WGT
    ss2.Col = 18
    sQuery = sQuery & ss2.Value & ","
    
    'CR_CD
    ss2.Col = 21
    sQuery = sQuery + "'" + ss2.Text + "',"
    
    'UST_FL
    ss2.Col = 23
    sQuery = sQuery + "'" + ss2.Text + "',"
    
    'TRIM_FL
    ss2.Col = 22
    sQuery = sQuery + "'" + ss2.Text + "',?,?)}"
    
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
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)

End Sub

Private Sub Slab_Block_Seq_Create(iType As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim P_SLAB_NO As String
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    'Max SLAB_EDT_SEQ READ
    'sQuery = "Select Max(slab_edt_seq) from EP_SLAB_EDT "
    'P_SLAB_EDT_SEQ = Gf_FloatFind(M_CN1, sQuery) + 1
    'iSLAB_EDT_SEQ = P_SLAB_EDT_SEQ
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    'SLAB_NO, BLOCK_SEQ, SEQ
    sQuery = "{call CEG2060C.P_MODIFY2 ('" + iType + "','9999999999','00','00',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
        
    sQuery = sQuery + "'',"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',?,?)}"
    
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
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

Private Sub Slab_Seq_Create(Seq As String, iType As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim P_SLAB_NO As String
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    'SLAB_NO, BLOCK_SEQ, SEQ
    sQuery = "{call CEG2060C.P_MODIFY2 ('" + iType + "','9999999999','" + Seq + "','00',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
        
    sQuery = sQuery + "'',"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',?,?)}"
    
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
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)

End Sub

Private Sub Slab_Cut_Block_Seq_Create(Seq As String, iType As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    'SLAB_NO, CUT_SEQ, BLOCK_SEQ, SEQ, SLAB_WGT
    sQuery = "{call CEG2060C.P_MODIFY3 ('" + iType + "','9999999999','" & Seq & "','00','00'," & sdb_slab_wgt.Value & ", "
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
        
    sQuery = sQuery + "'',"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',?,?)}"
    
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
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

Private Sub Slab_Cut_Seq_Create(Seq As String, iType As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim P_SLAB_NO As String
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    'SLAB_NO, CUT_SEQ, BLOCK_SEQ, SEQ
    sQuery = "{call CEG2060C.P_MODIFY3 ('" + iType + "','9999999999','" & Seq & "','99','00'," & sdb_slab_wgt.Value & ","
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
        
    sQuery = sQuery + "'',"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',?,?)}"
    
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
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)

End Sub

Private Sub Slab_Size(Size_Fl As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim sStlgrd As String
    
    Dim adoCmd As ADODB.Command
    
    ss2.Row = lMain_row
    ss2.Col = 5
    sStlgrd = ss2.Text
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "p_size"
    OutParam(1, 2) = adVariant
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    sQuery = "{call GP_JP_WGT ('" & Size_Fl & "','" & sStlgrd & "'," & sdb_slab_thk1.Value & "," & sdb_slab_wid1.Value
    sQuery = sQuery & "," & sdb_slab_len1.Value & "," & sdb_slab_wgt1.Value & ",?,?)}"
    
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
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        'Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        If Size_Fl = "THK" Then
            sdb_slab_thk1.Value = adoCmd("p_size")
        ElseIf Size_Fl = "WID" Then
            sdb_slab_wid1.Value = adoCmd("p_size")
        ElseIf Size_Fl = "LEN" Then
            sdb_slab_len1.Value = adoCmd("p_size")
        ElseIf Size_Fl = "WGT" Then
            sdb_slab_wgt1.Value = adoCmd("p_size")
        End If
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

Public Sub Asroll_Thk(sOrderNo As String)

    Dim sQuery As String
    
    'Asroll Thk
    sQuery = "         SELECT  NVL(MILL_TGT_THK,0) "
    sQuery = sQuery + "  FROM  NISCO.QP_QLTY_TECH "
    sQuery = sQuery + " WHERE  ORD_NO    = '" & Mid(sOrderNo, 1, 11) & "' "
    sQuery = sQuery + "   AND  ORD_ITEM  = '" & Mid(sOrderNo, 13, 2) & "' "
    sQuery = sQuery + "   AND  KND       = (SELECT  MAX(KND) "
    sQuery = sQuery + "                       FROM  NISCO.QP_QLTY_TECH "
    sQuery = sQuery + "                      WHERE  ORD_NO    = '" & Mid(sOrderNo, 1, 11) & "' "
    sQuery = sQuery + "                        AND  ORD_ITEM  = '" & Mid(sOrderNo, 13, 2) & "') "
    
    sdb_asroll_thk.Value = Gf_FloatFind(M_CN1, sQuery)
    sdb_slab_thk.Value = sdb_asroll_thk.Value
        
End Sub

Public Sub Asroll_Wid(sOrderNo As String)

    Dim sQuery As String
    Dim str_knd As String
    Dim squery_new As String
    str_knd = "1"
    
    'Asroll Wid
    sQuery = "         SELECT  NVL(MILL_TGT_WID,0) "
    sQuery = sQuery + "  FROM  NISCO.QP_QLTY_TECH "
    sQuery = sQuery + " WHERE  ORD_NO    = '" & Mid(sOrderNo, 1, 11) & "' "
    sQuery = sQuery + "   AND  ORD_ITEM  = '" & Mid(sOrderNo, 13, 2) & "' "
    sQuery = sQuery + "   AND  KND       = (SELECT  MAX(KND) "
    sQuery = sQuery + "                       FROM  NISCO.QP_QLTY_TECH "
    sQuery = sQuery + "                      WHERE  ORD_NO    = '" & Mid(sOrderNo, 1, 11) & "' "
    sQuery = sQuery + "                        AND  ORD_ITEM  = '" & Mid(sOrderNo, 13, 2) & "') "
    
'   sdb_asroll_wid.Value = Gf_FloatFind(M_CN1, squery)  (modified by mr.kim on 05-04-20)
    sdb_asroll_wid.Value = Gf_FloatFind(M_CN1, sQuery)
    'squery_new = "SELECT GF_CAL_WID('1',0," & sdb_asroll_wid.Value & ", 0 ," & sdb_slab_wid1.Value & ") FROM DUAL"
    'sdb_asroll_wid.Value = Gf_FloatFind(M_CN1, squery_new)
    'sdb_slab_wid.Value = sdb_asroll_wid.Value
        
End Sub

Public Sub Range_Wid(iRow As Long, dWidMin As Double, dWidMax As Double)

    Dim sQuery As String
    Dim dWid   As Double
    
    Set AdoRs = New ADODB.Recordset
    
    'Asroll Wid
    ss2.Row = iRow
    sQuery = "         SELECT  WID_TOL_MIN, WID_TOL_MAX "
    sQuery = sQuery + "  FROM  NISCO.QP_QLTY_DELV "
    ss2.Col = 1
    sQuery = sQuery + " WHERE  ORD_NO    = '" & Trim(ss2.Text) & "' "
    ss2.Col = 2
    sQuery = sQuery + "   AND  ORD_ITEM  = '" & Trim(ss2.Text) & "' "
    sQuery = sQuery + "   AND  KND       = '4'"
    
    AdoRs.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    ss2.Col = 9
    dWid = Val(ss2.Value & "")
    
    dWidMin = dWid + Val(AdoRs(0) & "")
    dWidMax = dWid + Val(AdoRs(1) & "")
    
End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim Prod_Thk As Double
    Dim Prod_Wid As Double
    Dim Prod_Len As Double
    Dim Prod_Wgt As Double
    
    If Col <> 14 Then Exit Sub
    ss2.Row = Row
    
    If ChangeMade Then
    
        ss2.Col = 8
        Prod_Thk = ss2.Value
        ss2.Col = 11
        Prod_Wid = ss2.Value
        ss2.Col = 14
        Prod_Len = ss2.Value
    
        Prod_Wgt = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('WGT',''," & Prod_Thk & "," & Prod_Wid & "," & Prod_Len & ",0) FROM DUAL ")
        
        ss2.Col = 18
        ss2.Value = Prod_Wgt
    
    End If
    
End Sub

Private Sub ss2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    Dim iCol As Integer
    Dim iRow As Integer
    Dim dMin As Double
    Dim dMax As Double
    Dim cValue As Double
    Dim sQuery As String
   
    If Row < 0 Or Row = 0 Then Exit Sub
    
    With ss2
            
        If .CellTag = "False" Then Exit Sub
        
        .Row = Row
              
        Select Case Col
        
            Case 14      'Design Product Len
            
                .Col = Col
                cValue = .Value
                
                .Col = Col + 2
                If .Value = "" Then
                    dMin = 0
                Else
                    dMin = .Value
                End If
                
                .Col = Col + 3
                If .Value = "" Then
                    dMax = 0
                Else
                    dMax = .Value
                End If
                                
                If cValue > dMax Or cValue < dMin Then
                
                    .Col = Col
                    .Row = Row
                    .CellTag = "False"
                 
                    Call Gp_MsgBoxDisplay("已超出最大/最小值...!!")
                  
                    .Col = Col
                    .Row = Row
                    .CellTag = ""
                    
                    .Value = 0
                    .TabStop = True
                    .SetFocus
                    .SetActiveCell Col, Row
                    .Action = SS_ACTION_ACTIVE_CELL
                    .EditMode = True
                    .TabStop = False
                    
                End If
           
        End Select
            
   End With

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss2
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If
    
End Sub

Private Sub TxT_stdgrd_DblClick()

    Call TxT_stdgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TxT_stdgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

         DD.sWitch = "MS"
         DD.rControl.Add Item:=txt_stdgrd
        
         DD.nameType = "1"
         Call Gf_Stlgrd_DD(M_CN1, KeyCode)

    End If

End Sub
