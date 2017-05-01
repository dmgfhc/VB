VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACE1030C 
   Caption         =   "可替代余材选定_ACE1030C"
   ClientHeight    =   9225
   ClientLeft      =   930
   ClientTop       =   4305
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9135
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   16113
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "ACE1030C.frx":0000
      Begin Threed.SSFrame SSFrame1 
         Height          =   1365
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   2408
         _Version        =   196609
         BackColor       =   14737632
         ShadowStyle     =   1
         Begin VB.TextBox TXT_MAT_NO 
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
            Left            =   10365
            MaxLength       =   14
            TabIndex        =   23
            Tag             =   "物料号"
            Top             =   530
            Width           =   1335
         End
         Begin VB.TextBox txt_STLGRD_Name 
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
            Left            =   6705
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   22
            Tag             =   "钢种(标准号)"
            Top             =   120
            Width           =   1890
         End
         Begin VB.TextBox Text_PROD_CD_mate 
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
            Left            =   1915
            MaxLength       =   2
            TabIndex        =   21
            Tag             =   "BIZ_AREA"
            Text            =   "板坯"
            Top             =   125
            Width           =   1230
         End
         Begin VB.TextBox text_prod_cd 
            BackColor       =   &H00C0FFFF&
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
            Left            =   1530
            MaxLength       =   2
            TabIndex        =   11
            Tag             =   "产品"
            Text            =   "SL"
            Top             =   125
            Width           =   375
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
            Height          =   310
            Left            =   10365
            MaxLength       =   11
            TabIndex        =   10
            Tag             =   "订单号"
            Top             =   125
            Width           =   1335
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
            Height          =   310
            Left            =   5400
            MaxLength       =   12
            TabIndex        =   9
            Tag             =   "钢种"
            Top             =   125
            Width           =   1305
         End
         Begin VB.ComboBox prod_combo 
            Height          =   300
            ItemData        =   "ACE1030C.frx":0052
            Left            =   2010
            List            =   "ACE1030C.frx":0062
            OLEDragMode     =   1  'Automatic
            TabIndex        =   8
            Top             =   3120
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.ComboBox Combo_ORD_ITEM 
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
            Left            =   11700
            TabIndex        =   6
            Top             =   120
            Width           =   630
         End
         Begin VB.TextBox txt_loc 
            Height          =   315
            Left            =   5400
            TabIndex        =   5
            Top             =   525
            Width           =   1305
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
            Left            =   1915
            TabIndex        =   4
            Top             =   525
            Width           =   1230
         End
         Begin VB.TextBox text_cur_inv_code 
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
            Left            =   1530
            MaxLength       =   2
            TabIndex        =   3
            Top             =   525
            Width           =   375
         End
         Begin CSTextLibCtl.sidbEdit sdb_prod_wgt_to 
            Height          =   315
            Left            =   12465
            TabIndex        =   7
            Top             =   2370
            Visible         =   0   'False
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
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
            RawData         =   "100.00"
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
            MaxValue        =   9999999
            MinValue        =   0
            Undo            =   0
            Data            =   99.999
         End
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Index           =   0
            Left            =   150
            Top             =   120
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "产品"
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
            Left            =   4005
            Top             =   120
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "钢种"
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
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   0
            Left            =   8970
            Top             =   120
            Width           =   1350
            _ExtentX        =   2381
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
         Begin CSTextLibCtl.sidbEdit sdb_prod_thk_fr 
            Height          =   315
            Left            =   1530
            TabIndex        =   12
            Tag             =   "产品厚度（MIN）"
            Top             =   930
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
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
            FmtControl      =   1
            NumDecDigits    =   2
            NumIntDigits    =   4
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Left            =   150
            Top             =   930
            Width           =   1350
            _ExtentX        =   2381
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
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   4005
            Top             =   930
            Width           =   1350
            _ExtentX        =   2381
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
            Index           =   0
            Left            =   8970
            Top             =   930
            Width           =   1350
            _ExtentX        =   2381
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
            Left            =   2565
            TabIndex        =   13
            Tag             =   "产品厚度（MAX）"
            Top             =   930
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
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
            RawData         =   "999.99"
            Text            =   " 999.99"
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
            NumIntDigits    =   3
            Undo            =   0
            Data            =   999.99
         End
         Begin CSTextLibCtl.sidbEdit sdb_prod_len_fr 
            Height          =   315
            Left            =   10365
            TabIndex        =   14
            Tag             =   "产品长度（MIN）"
            Top             =   930
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
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
            Left            =   11400
            TabIndex        =   15
            Tag             =   "产品长度（MIN）"
            Top             =   930
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
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
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "999999"
            Text            =   " 999,999"
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
            MaxValue        =   100000
            MinValue        =   0
            Undo            =   0
            Data            =   999999
         End
         Begin CSTextLibCtl.sidbEdit sdb_prod_wid_fr 
            Height          =   315
            Left            =   5400
            TabIndex        =   16
            Tag             =   "产品宽度（MIN）"
            Top             =   930
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
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
            Left            =   6435
            TabIndex        =   17
            Tag             =   "产品宽度（MAX）"
            Top             =   930
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
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
            RawData         =   "9999"
            Text            =   " 9,999"
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
            Data            =   9999
         End
         Begin Threed.SSCommand Command_ALLSELECT 
            Height          =   450
            Left            =   13590
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   120
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   794
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "板坯选定"
            BevelWidth      =   3
         End
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Index           =   0
            Left            =   4005
            Top             =   525
            Width           =   1350
            _ExtentX        =   2381
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Index           =   1
            Left            =   10395
            Top             =   2370
            Visible         =   0   'False
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            Caption         =   "产品重量"
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
         Begin Threed.SSCommand cmd_cancel 
            Height          =   450
            Left            =   13590
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   840
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   794
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   12583104
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "取消替代"
            BevelWidth      =   3
         End
         Begin CSTextLibCtl.sidbEdit sdb_prod_wgt_fr 
            Height          =   315
            Left            =   11445
            TabIndex        =   20
            Top             =   2370
            Visible         =   0   'False
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
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
            MaxValue        =   100000
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   150
            Top             =   525
            Width           =   1350
            _ExtentX        =   2381
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
         Begin InDate.ULabel ULabel20 
            Height          =   315
            Left            =   8970
            Top             =   525
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "物料号"
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
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   7740
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1395
         Width           =   15135
         _Version        =   393216
         _ExtentX        =   26696
         _ExtentY        =   13653
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
         MaxCols         =   27
         MaxRows         =   2
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACE1030C.frx":008E
         VisibleCols     =   1
      End
   End
End
Attribute VB_Name = "ACE1030C"
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
'-- Program ID        ACE1030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Yang Zhibin
'-- Date              2003.9.8
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
Public Active_CForm As String       'Form Active

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iCount As Integer

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(text_prod_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(Combo_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_prod_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_prod_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_prod_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_prod_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_prod_len_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_prod_len_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_loc, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_prod_wgt_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_prod_wgt_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
   ' Call Gp_Sp_ColHidden(ss1, 1, True)
    
    'Duplicate Count
    'iDupCnt = 1
    
    'Sum Column Count
    'iSumCnt = 1
    
    'Sum Column Setting
    'iSumCol.Add Item:=5
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub Cmd_Cancel_Click()

    Dim iRow As Integer
    Dim Mat_no As String
    Dim Prod_cd As String
    Dim sRef_yn As Boolean
    
    sRef_yn = False
    
    For iRow = 1 To ss1.MaxRows
        
        ss1.Row = iRow
        ss1.Col = 2
        Mat_no = ss1.Text
        ss1.Col = 1
        Prod_cd = ss1.Text
      '  Prod_cd = Text_PROD_CD.Text
        
        ss1.Col = 0
        If ss1.Text <> "" Then
            If Cancel_Pro(Prod_cd, Mat_no) Then
                ss1.Col = 0
                ss1.Text = ""
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow)
                sRef_yn = True
            Else
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , vbYellow)
                Exit For
            End If
        End If
        
    Next

    If sRef_yn Then
        Call Form_Ref
    End If
        
End Sub

Private Sub Combo_ORD_ITEM_Change()
  
'    If combo_ord_item.Text <> "" Then
'         If combo_ord_item.Text > combo_ord_item.ListCount Then
'           combo_ord_item.Text = ""
'         End If
'    End If

End Sub

Private Sub Combo_ORD_ITEM_KeyPress(KeyAscii As Integer)
    'KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub Combo_ORD_ITEM_LostFocus()

    Dim S As String
  
    If Len(Combo_ORD_ITEM.Text) = 1 Then
        S = Combo_ORD_ITEM.Text
        Combo_ORD_ITEM.Text = "0" + S
    End If
    
End Sub

Private Sub Command_ALLSELECT_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4)      As Variant
    Dim ret_Result_ErrMsg   As String
    Dim sQuery              As String
    Dim iCount              As Integer
    Dim sProd_kind          As String
    Dim stlgrd              As String
    Dim prod_loc            As String
    Dim sCur_Inv            As String
    Dim SMESG               As String
    Dim sMat_No             As String
    
    Dim minSIZEthk          As Single     '--thick
    Dim maxSIZEthk          As Single
    Dim minSIZEwid          As Single     '--wide
    Dim maxSIZEwid          As Single
    Dim minSIZElen          As Single     '--lenth
    Dim maxSIZElen          As Single
    Dim minSIZEwgt          As Single     '--wight
    Dim maxSIZEwgt          As Single
    
    Dim adoCmd As ADODB.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    sProd_kind = "4" 'Left(prod_combo.Text, 1)

  '  squery = "{call ACE1020P ('C1','" + sProd_kind + "','" + slab_confirm.kqnum + "','" + sUserID + "',?)}"

'-------------------------------------------------------------------------
    '----物料选定的条件：钢种，物料位置，产品重量，长，宽，厚
    minSIZEthk = sdb_prod_thk_fr.Value
    maxSIZEthk = sdb_prod_thk_to.Value
    minSIZEwid = sdb_prod_wid_fr.Value
    maxSIZEwid = sdb_prod_wid_to.Value
    minSIZElen = sdb_prod_len_fr.Value
    maxSIZElen = sdb_prod_len_to.Value
    minSIZEwgt = sdb_prod_wgt_fr.Value
    maxSIZEwgt = sdb_prod_wgt_to.Value
    stlgrd = txt_stlgrd.Text
    prod_loc = txt_loc.Text
    sCur_Inv = text_cur_inv_code.Text
     
     '如果位置的长度大于10，则提示位置信息输入错误
        
    If Len(prod_loc) > 10 Then
           Call MsgBox("物料位置输入错误！" & Chr(10) & "请重新输入。", vbExclamation + vbOKOnly, "警告")
           Screen.MousePointer = vbDefault
           Exit Sub
    End If
    If Len(prod_loc) > 0 Then
           If Left(prod_loc, 1) <> Left(text_prod_cd.Text, 1) Then
               Call MsgBox("产品代码与物料位置不符！" & Chr(10) & "请检查后输入。", vbExclamation + vbOKOnly, "警告")
               Screen.MousePointer = vbDefault
               Exit Sub
           End If
    End If
    
    If maxSIZEthk < minSIZEthk Then
        Call MsgBox("厚度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
    ElseIf maxSIZEwid < minSIZEwid Then
        Call MsgBox("宽度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
    ElseIf maxSIZElen < minSIZElen Then
        Call MsgBox("长度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
    ElseIf maxSIZEwgt < minSIZEwgt Then
        Call MsgBox("重量区间不符合规范！" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
    Else
        SMESG = Gf_Ms_NeceCheck(nControl)
        If SMESG <> "OK" Then
            SMESG = SMESG + " Must input necessarily"
            Call Gp_MsgBoxDisplay(SMESG)
        Else
            SMESG = Gf_Ms_NeceCheck2(mControl)
            If SMESG <> "OK" Then
                SMESG = SMESG + " Must input according to length of item"
                Call Gp_MsgBoxDisplay(SMESG)
            Else
        
                M_CN1.CursorLocation = adUseServer
                  '---------squery(CALL ACE1030P)----------------------
        
                   sQuery = "{CALL ACE1030P('C1',            '" & _
                                             sProd_kind & "','" & _
                                             sCur_Inv & "',  '" & _
                                             stlgrd & "',    '" & _
                                             TXT_MAT_NO & "','" & _
                                             prod_loc & "',   " & _
                                             minSIZEthk & ",  " & _
                                             maxSIZEthk & ",  " & _
                                             minSIZEwid & ",  " & _
                                             maxSIZEwid & ",  " & _
                                             minSIZElen & ",  " & _
                                             maxSIZElen & ",  " & _
                                             minSIZEwgt & ",  " & _
                                             maxSIZEwgt & ", '" & _
                                             sUserID & "',?)}"
        
                   '-------------------------------------------------------
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
                Else
                    Call Gp_MsgBoxDisplay("选定完了..!!", "I")
                    Call Form_Ref
                End If
            End If
        End If
    End If

'-------------------------------------------------------------------------
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

Private Sub Form_Activate()
    
    If Active_CForm <> "" Then
        Call Form_Ref
        Active_CForm = ""
    End If
    
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
        cmd_cancel.Enabled = False
        Command_ALLSELECT.Enabled = False
    End If
    
    Call Form_Define
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    'ERP I/F START CHECK -- ONLY SLAB SELECT
'    If Gf_ErpSystem_Chk Then
        text_prod_cd.Text = "SL"
        text_prod_cd.Enabled = False
'    End If

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        Combo_ORD_ITEM.Clear
        
        'ERP I/F START CHECK -- ONLY SLAB SELECT
        If Gf_ErpSystem_Chk Then
            text_prod_cd.Text = "SL"
            text_prod_cd.Enabled = False
        End If
    
        sdb_prod_thk_to.Value = 9999.99
        sdb_prod_wid_to.Value = 999999
        sdb_prod_len_to.Value = 9999999
        sdb_prod_wgt_to.Value = 99999.99
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    Dim sQuery          As String
    Dim SMESG           As String
    Dim sProduct        As String
    Dim S               As String
    Dim minSIZEthk      As Single     '--thick
    Dim maxSIZEthk      As Single
    Dim minSIZEwid      As Single     '--wide
    Dim maxSIZEwid      As Single
    Dim minSIZElen      As Single     '--lenth
    Dim maxSIZElen      As Single
    Dim minSIZEwgt      As Single     '--wight
    Dim maxSIZEwgt      As Single
    Dim stlgrd          As String
    Dim prod_loc        As String
    Dim str_ord_no      As String
    Dim str_ord_item    As String
    Dim sCur_Inv        As String
    
    '----物料查询的条件：钢种，物料位置，产品重量，长，宽，厚，订单号
    minSIZEthk = sdb_prod_thk_fr.Value
    maxSIZEthk = sdb_prod_thk_to.Value
    minSIZEwid = sdb_prod_wid_fr.Value
    maxSIZEwid = sdb_prod_wid_to.Value
    minSIZElen = sdb_prod_len_fr.Value
    maxSIZElen = sdb_prod_len_to.Value
    minSIZEwgt = sdb_prod_wgt_fr.Value
    maxSIZEwgt = sdb_prod_wgt_to.Value
    stlgrd = txt_stlgrd.Text
    prod_loc = txt_loc.Text
    str_ord_no = txt_ord_no.Text
    str_ord_item = Combo_ORD_ITEM.Text
    sCur_Inv = text_cur_inv_code.Text
        
    If Len(prod_loc) > 9 Then
        Call MsgBox("物料位置输入错误！" & Chr(10) & "请重新输入。", vbExclamation + vbOKOnly, "警告")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If Len(prod_loc) > 0 Then
        If Left(prod_loc, 1) <> Left(text_prod_cd.Text, 1) Then
            Call MsgBox("产品代码与物料位置不符！" & Chr(10) & "请检查后输入。", vbExclamation + vbOKOnly, "警告")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    If Combo_ORD_ITEM.Text <> "" Then
        If Len(Combo_ORD_ITEM.Text) = 1 Then
            S = Combo_ORD_ITEM.Text
            Combo_ORD_ITEM.Text = "0" + S
        End If
    End If
    
    '---------------------根据产品的类型-------------------------------
    If text_prod_cd.Text <> "" Then
    
       If text_prod_cd.Text = "SL" Then
       
           sQuery = " Select  'SL',  A.SLAB_NO, "
           sQuery = sQuery + " A.STLGRD, A.THK, A.WID, A.LEN, A.WGT, DECODE(A.SLAB_CUT_FL,'Y','是','否'),A.SLAB_CUT_CNT,A.SLAB_REM_WGT,A.SLAB_REM_LEN, "
           sQuery = sQuery + " A.ORD_NO||'-'||A.ORD_ITEM, B.PROD_THK, B.PROD_WID, B.PROD_LEN, B.PROD_WGT, "
           sQuery = sQuery + " A.TRIM_FL, A.UST_FL, TO_DATE(SUBSTR(A.PROD_DATE,1,8),'YYYY-MM-DD'), Gf_ComnNameFind('C0008',A.WOO_RSN), "
           sQuery = sQuery + " A.CR_CD, Gf_ComnNameFind('C0013',A.CUR_INV), A.LOC, A.ORG_ORD_NO||'-'||A.ORG_ORD_ITEM, "
           sQuery = sQuery + " DECODE(C.MAT_OWNER_FL,'B','委托加工',''), C.MAT_OWNER_CUST_CD, A.MLT_PROC_CD MLT_PROC_CD "
           sQuery = sQuery + " From CP_REP_SLAB A, CP_REP_ORD B, FP_SLAB C "
           sQuery = sQuery + "  Where NVL(A.STLGRD,' ')   Like '" + Trim(txt_stlgrd.Text) + "%' "
           sQuery = sQuery + "    AND NVL(A.ORD_NO,' ')   Like '" + Trim(txt_ord_no.Text) + "%' "
           sQuery = sQuery + "    AND NVL(A.ORD_ITEM,' ') Like '" + Trim(Combo_ORD_ITEM.Text) + "%' "
           sQuery = sQuery + "    AND A.THK  BETWEEN " & sdb_prod_thk_fr.Value & " AND " & sdb_prod_thk_to.Value
           sQuery = sQuery + "    AND A.WID  BETWEEN " & sdb_prod_wid_fr.Value & " AND " & sdb_prod_wid_to.Value
           sQuery = sQuery + "    AND A.LEN  BETWEEN " & sdb_prod_len_fr.Value & " AND " & sdb_prod_len_to.Value
           sQuery = sQuery + "    AND NVL(A.ORD_NO,' ')   =    NVL(B.ORD_NO(+),' ')"
           sQuery = sQuery + "    AND NVL(A.ORD_ITEM,' ') =    NVL(B.ORD_ITEM(+),' ') "
           sQuery = sQuery + "    AND A.WGT  BETWEEN " & minSIZEwgt & " AND " & maxSIZEwgt                         'GENG 重量范围
           sQuery = sQuery + "    AND NVL(A.LOC,' ')  LIKE '" + prod_loc + "%'"                              '根据位置查询
           sQuery = sQuery + "    AND A.SLAB_NO       LIKE '" + TXT_MAT_NO + "%'"                            '物料号
           sQuery = sQuery + "    AND A.SLAB_NO          =  C.SLAB_NO "
           If sCur_Inv <> "" Then
                sQuery = sQuery + "AND NVL(A.CUR_INV,'00') = '" + sCur_Inv + "'"
           End If
           sQuery = sQuery + "  ORDER BY A.LOC ASC "
           
       End If
       ''''''''DANGZERONG
     
       If txt_ord_no.Text <> "" Then
           
           ''''' SL
           sQuery = sQuery + " Select 'SL',   A.SLAB_NO, "
           sQuery = sQuery + " A.STLGRD, A.THK, A.WID, A.LEN, A.WGT, DECODE(A.SLAB_CUT_FL,'Y','是','否'),A.SLAB_CUT_CNT,A.SLAB_REM_WGT,A.SLAB_REM_LEN, "
           sQuery = sQuery + " A.ORD_NO||'-'||A.ORD_ITEM, B.PROD_THK, B.PROD_WID, B.PROD_LEN, B.PROD_WGT, "
           sQuery = sQuery + " A.TRIM_FL, A.UST_FL, SUBSTR(A.PROD_DATE,1,8),Gf_ComnNameFind('C0008',A.WOO_RSN), "
           sQuery = sQuery + " A.CR_CD, Gf_ComnNameFind('C0013',A.CUR_INV), A.LOC, A.ORG_ORD_NO||'-'||A.ORG_ORD_ITEM, "
           sQuery = sQuery + " DECODE(C.MAT_OWNER_FL,'B','委托加工',''), C.MAT_OWNER_CUST_CD, A.MLT_PROC_CD MLT_PROC_CD "
           sQuery = sQuery + " From CP_REP_SLAB A, CP_REP_ORD B, FP_SLAB C  "
           sQuery = sQuery + "  Where NVL(A.STLGRD,' ')   Like '" + Trim(txt_stlgrd.Text) + "%' "
           sQuery = sQuery + "    AND NVL(A.ORD_NO,' ')   Like '" + Trim(txt_ord_no.Text) + "%' "
           sQuery = sQuery + "    AND NVL(A.ORD_ITEM,' ') Like '" + Trim(Combo_ORD_ITEM.Text) + "%' "
           sQuery = sQuery + "    AND A.THK  BETWEEN " & sdb_prod_thk_fr.Value & " AND " & sdb_prod_thk_to.Value
           sQuery = sQuery + "    AND A.WID  BETWEEN " & sdb_prod_wid_fr.Value & " AND " & sdb_prod_wid_to.Value
           sQuery = sQuery + "    AND A.LEN  BETWEEN " & sdb_prod_len_fr.Value & " AND " & sdb_prod_len_to.Value
           sQuery = sQuery + "    AND NVL(A.ORD_NO,' ')   =    NVL(B.ORD_NO(+),' ')"
           sQuery = sQuery + "    AND NVL(A.ORD_ITEM,' ') =    NVL(B.ORD_ITEM(+),' ') "
           sQuery = sQuery + "    AND A.WGT  BETWEEN " & minSIZEwgt & " AND " & maxSIZEwgt                         'GENG 重量范围
           sQuery = sQuery + "    AND NVL(A.LOC,' ')  LIKE '" + prod_loc + "%'"                              '根据位置查询
           sQuery = sQuery + "    AND A.SLAB_NO       LIKE '" + TXT_MAT_NO + "%'"                            '物料号
           sQuery = sQuery + "    AND A.SLAB_NO      =  C.SLAB_NO "
           If sCur_Inv <> "" Then
                sQuery = sQuery + "AND NVL(A.CUR_INV,'00') = '" + sCur_Inv + "'"
           End If
           
           sQuery = sQuery + "  ORDER BY LOC ASC "
       End If
     
     
       'Debug.Print squery
       
       SMESG = Gf_Ms_NeceCheck(nControl)
       If SMESG = "OK" Then
       
           SMESG = Gf_Ms_NeceCheck2(mControl)
           If SMESG = "OK" Then
    
               If Gf_Only_Display(M_CN1, Proc_Sc("Sc"), sQuery, , , False) Then
                   ss1.OperationMode = OperationModeNormal
                   Call Gp_Ms_ControlLock(Mc1("lControl"), True)
                   Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
               End If
               
           Else
               SMESG = SMESG + " Must input according to length of item"
               Call Gp_MsgBoxDisplay(SMESG)
               Screen.MousePointer = vbDefault
           End If
       
       Else
           SMESG = SMESG + " Must input necessarily"
           Call Gp_MsgBoxDisplay(SMESG)
           Screen.MousePointer = vbDefault
       End If
       
    Else
    
       Call MsgBox("产品分类代码不能为空！" & Chr(10) & "请重试。", vbExclamation + vbOKOnly, "警告")
       Screen.MousePointer = vbDefault
       text_prod_cd.Text = ""
       text_prod_cd.SetFocus
       
    End If

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

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub Label2_Click()

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
    
    If Row < 1 Then Exit Sub
    If ss1.MaxRows < 1 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 9
    If Trim(ss1.Text) = "-" Or Trim(ss1.Text) = "" Then Exit Sub
    
    ss1.Col = 0
    
    If ss1.Text <> "选择" Then
        ss1.Text = "选择"
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
    Else
        ss1.Text = ""
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
    End If

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim P_SLAB_NO As String
    
    ss1.Col = 1
    ss1.Row = Row
    
    If ss1.Text = "SL" Then
        Load ACE1151C
        
        ss1.Row = Row
        ss1.Col = 2
        P_SLAB_NO = ss1.Text
        
        ACE1151C.txt_slab_no.Text = P_SLAB_NO
        ACE1151C.Show 1
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
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
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
        
        If text_prod_cd.Text = "SL" And text_cur_inv_code.Text = "ZB" Then
            text_cur_inv_code.Text = ""
            text_cur_inv.Text = ""
        End If
       
    Else
    
        If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
            text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
        Else
            text_cur_inv.Text = ""
        End If
        
        If text_prod_cd.Text = "SL" And text_cur_inv_code.Text = "ZB" Then
            text_cur_inv_code.Text = ""
            text_cur_inv.Text = ""
        End If
        
    End If
    
End Sub

Private Sub Text_PROD_CD_Change()

    Select Case text_prod_cd.Text
           Case "S", "s", "SL"
               text_prod_cd.Text = "SL"
           Case "P", "p", "PP"
               text_prod_cd.Text = "PP"
           Case "H", "h", "HC"
               text_prod_cd.Text = "HC"
           Case ""
               text_prod_cd.Text = ""
           Case Else
               text_prod_cd.Text = ""
               Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
               Text_PROD_CD_mate.Text = ""
    End Select

End Sub

Private Sub Text_PROD_CD_DblClick()

    Call Text_PROD_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=text_prod_cd
'        DD.rControl.Add Item:=Text_PROD_CD_mate
   
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() 用于客户代码

        Exit Sub
        
    End If

'    If Len(Trim(text_prod_cd.Text)) = text_prod_cd.MaxLength Then
'
'        Text_PROD_CD_mate.Text = Gf_ComnNameFind(M_CN1, "B0005", text_prod_cd.Text, 2)
'    Else
'        Text_PROD_CD_mate.Text = ""
'    End If


End Sub

Private Sub text_prod_cd_LostFocus()

    If text_prod_cd.Text <> "" Then
        If (Len(text_prod_cd.Text) < text_prod_cd.MaxLength) Then
            Call Gp_MsgBoxDisplay("产品分类不符合规范！")
            'Text_PROD_CD.Text = ""
            text_prod_cd.SetFocus
        End If
    End If
    
End Sub

Private Sub txt_mat_no_Change()

    TXT_MAT_NO.Text = Replace(TXT_MAT_NO, vbCrLf, "")
    
End Sub

Private Sub txt_stlgrd_Change()

    txt_stlgrd.Text = Replace(txt_stlgrd, vbCrLf, "")

    If Len(txt_stlgrd.Text) <> 11 Then txt_STLGRD_Name.Text = ""
   
End Sub

Private Sub txt_stlgrd_DblClick()

    Call txt_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stlgrd
        DD.rControl.Add Item:=txt_STLGRD_Name
        DD.nameType = "1"
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_stlgrd.Text)) >= 10 Then
            txt_STLGRD_Name.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stlgrd.Text))
        Else
            txt_STLGRD_Name.Text = ""
        End If
        
    End If
        
End Sub

Public Function Cancel_Pro(Prod_cd As String, Mat_no As String) As Boolean

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call ACE1200P ('" + Prod_cd + "','" + Mat_no + "',?)}"
    
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
        Cancel_Pro = False
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    Cancel_Pro = True
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Function

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Cancel_Pro = False
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)

End Function

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String
    
    If Len(Trim(txt_ord_no.Text)) = txt_ord_no.MaxLength Then
    
        If Combo_ORD_ITEM.Text <> "" Then Exit Sub
        
        txt_ord_no.Text = StrConv(txt_ord_no.Text, vbUpperCase)
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(txt_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, Combo_ORD_ITEM, sQuery)
        
       ' If combo_ord_item.ListCount <> 0 Then
       '       combo_ord_item.ListIndex = 0
       ' End If
    Else
        Combo_ORD_ITEM.Clear
    End If

End Sub
