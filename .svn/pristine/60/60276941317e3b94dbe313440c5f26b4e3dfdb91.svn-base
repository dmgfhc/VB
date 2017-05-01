VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACE1010C 
   Caption         =   "可替代订单查询_ACE1010C"
   ClientHeight    =   9225
   ClientLeft      =   345
   ClientTop       =   1905
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   Tag             =   "交货期"
   WindowState     =   2  'Maximized
   Begin VB.ComboBox prod_combo 
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
      ItemData        =   "ACE1010C.frx":0000
      Left            =   15720
      List            =   "ACE1010C.frx":000A
      OLEDragMode     =   1  'Automatic
      TabIndex        =   24
      Top             =   690
      Visible         =   0   'False
      Width           =   420
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9135
      Left            =   60
      TabIndex        =   4
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
      PaneTree        =   "ACE1010C.frx":001E
      Begin Threed.SSFrame SSFrame1 
         Height          =   1350
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   2381
         _Version        =   196609
         BackColor       =   14737632
         ShadowStyle     =   1
         Begin VB.TextBox txt_stlgrd_name 
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
            Left            =   10080
            TabIndex        =   26
            Top             =   530
            Width           =   2355
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
            Left            =   8670
            MaxLength       =   11
            TabIndex        =   25
            Top             =   530
            Width           =   1395
         End
         Begin VB.TextBox TXT_ORD_NO 
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
            Left            =   4920
            MaxLength       =   11
            TabIndex        =   23
            Tag             =   "CD_MANA_NO"
            Top             =   135
            Width           =   1380
         End
         Begin VB.ComboBox CBO_ORD_ITEM 
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
            Left            =   6315
            TabIndex        =   22
            Top             =   135
            Width           =   660
         End
         Begin VB.TextBox text_prod_cd 
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
            Left            =   1545
            MaxLength       =   2
            TabIndex        =   0
            Tag             =   "BIZ_AREA"
            Top             =   140
            Width           =   480
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
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   12
            Tag             =   "BIZ_AREA"
            Top             =   140
            Width           =   1125
         End
         Begin VB.TextBox txt_stdspec 
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
            TabIndex        =   11
            Top             =   135
            Width           =   2100
         End
         Begin VB.TextBox txt_cust_cd 
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
            Left            =   12450
            MaxLength       =   6
            TabIndex        =   10
            Top             =   135
            Width           =   810
         End
         Begin VB.TextBox TXT_CUST_DES 
            Height          =   315
            Left            =   13290
            TabIndex        =   9
            Top             =   135
            Width           =   1605
         End
         Begin VB.TextBox Text_size_knd_name 
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
            Height          =   310
            Left            =   2040
            TabIndex        =   8
            Tag             =   "钢种"
            Top             =   530
            Width           =   1125
         End
         Begin VB.TextBox Text_size_knd 
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
            Height          =   310
            Left            =   1545
            MaxLength       =   2
            TabIndex        =   7
            Tag             =   "钢种"
            Top             =   530
            Width           =   480
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   1
            Left            =   150
            Top             =   135
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "订单产品"
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
            Left            =   3525
            Top             =   530
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "订单交货期"
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
         Begin InDate.UDate UDate_DEL_TO_b 
            Height          =   315
            Left            =   4920
            TabIndex        =   13
            Tag             =   "INS_DATE"
            Top             =   530
            Width           =   1425
            _ExtentX        =   2514
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
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Index           =   1
            Left            =   7290
            Top             =   135
            Width           =   1350
            _ExtentX        =   2381
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
            Left            =   150
            Top             =   915
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
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Index           =   1
            Left            =   3525
            Top             =   900
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
            Left            =   7290
            Top             =   915
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
         Begin CSTextLibCtl.sidbEdit sidbEdit_size_Athk 
            Height          =   315
            Left            =   1545
            TabIndex        =   14
            Top             =   915
            Width           =   795
            _Version        =   262145
            _ExtentX        =   1402
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   -2147483640
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
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sidbEdit_size_Bthk 
            Height          =   315
            Left            =   2370
            TabIndex        =   15
            Top             =   915
            Width           =   795
            _Version        =   262145
            _ExtentX        =   1402
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   -2147483640
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
            MaxValue        =   999.99
            MinValue        =   0
            Undo            =   0
            Data            =   999.99
         End
         Begin CSTextLibCtl.sidbEdit sidbEdit_size_Awid 
            Height          =   315
            Left            =   4920
            TabIndex        =   16
            Top             =   915
            Width           =   1005
            _Version        =   262145
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   -2147483640
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
            NumIntDigits    =   4
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sidbEdit_size_Bwid 
            Height          =   315
            Left            =   5940
            TabIndex        =   17
            Top             =   915
            Width           =   1005
            _Version        =   262145
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   -2147483640
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
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   9999
         End
         Begin CSTextLibCtl.sidbEdit sidbEdit_size_Alen 
            Height          =   315
            Left            =   8670
            TabIndex        =   18
            Top             =   915
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   -2147483640
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
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sidbEdit_size_Blen 
            Height          =   315
            Left            =   9720
            TabIndex        =   19
            Top             =   915
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   -2147483640
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
            NumIntDigits    =   6
            MaxValue        =   999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   999999
         End
         Begin Threed.SSCommand cmd_confirm 
            Height          =   390
            Left            =   13380
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   900
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   688
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
            Caption         =   "替代确定处理"
            BevelWidth      =   3
         End
         Begin Threed.SSCommand Command_ALLSELECT 
            Height          =   390
            Left            =   11400
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   900
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   688
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
            Caption         =   "订单选定"
            BevelWidth      =   3
         End
         Begin InDate.ULabel ULabel10 
            Height          =   315
            Left            =   11070
            Top             =   135
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "客户"
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
            Left            =   150
            Top             =   525
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "定尺区分"
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
            Left            =   3525
            Top             =   135
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
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   7290
            Top             =   530
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
         Height          =   7755
         Left            =   0
         TabIndex        =   6
         Top             =   1380
         Width           =   15135
         _Version        =   393216
         _ExtentX        =   26696
         _ExtentY        =   13679
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
         MaxCols         =   34
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACE1010C.frx":0070
      End
   End
   Begin VB.OptionButton opt_Prod_Kind 
      BackColor       =   &H00E0E0E0&
      Caption         =   "成品"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   16440
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.OptionButton opt_Prod_Kind 
      BackColor       =   &H00E0E0E0&
      Caption         =   "半成品"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   15390
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   960
   End
   Begin Threed.SSCommand Command_REP 
      Height          =   390
      Left            =   17250
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   688
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "替代处理"
   End
End
Attribute VB_Name = "ACE1010C"
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
'-- Program ID        ACE1010C
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
Dim t_ord_no As String
Dim t_ord_item As String
Dim sdel As String

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
Dim nColumn1 As New Collection      'Spread necessary Column1 Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column1 Collection
Dim iColumn1 As New Collection      'Spread Insert Column1 Collection
Dim aColumn1 As New Collection      'Master -> Spread Column1 Collection
Dim lColumn1 As New Collection      'Spread Lock Column1 Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column1

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SPD_ORD_NO = 1
Const SPD_ORD_ITEM = 2
Const SPD_PROD_CD = 6
Const SPD_PROD_LEN = 13
Const SPD_SIZE_KND = 14
Const SPD_CNF_WGT = 28
Const SPD_END_WGT = 29

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
  'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(text_prod_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_cust_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sidbEdit_size_Awid, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sidbEdit_size_Bwid, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sidbEdit_size_Athk, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sidbEdit_size_Bthk, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sidbEdit_size_Alen, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sidbEdit_size_Blen, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(UDate_DEL_TO_b, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(Text_size_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(cbo_ord_item, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    'Call Spread_Collection("Column1_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    Call Gp_Sp_ColColor(ss1, 1)
    Call Gp_Sp_ColColor(ss1, 2)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACE1010C.P_REFER", Key:="P-R"
    sc1.Add Item:="ACE1010C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
'    'Duplicate Count
'    iDupCnt = 1
'
'    'Sum Column1 Count
'    iSumCnt = 1
'
'    'Sum Column1 Setting
'    iSumCol.Add Item:=5
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub cmd_confirm_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call ACE1210P ('C1','','" + sUserID + "',?)}"
    
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
    Else
        Call Gp_MsgBoxDisplay("确定处理完了..!!", "I")
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

Private Sub Command_ALLSELECT_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    Dim sProd_Fl As String
    Dim sProd_kind As String

    Dim SMESG As String
    Dim maxDATE As String
    Dim minSIZEthk As Single
    Dim maxSIZEthk As Single
    Dim minSIZEwid As Single
    Dim maxSIZEwid As Single
    Dim minSIZElen As Single
    Dim maxSIZElen As Single
    
    Dim adoCmd As ADODB.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    sProd_kind = Left(prod_combo.Text, 1)
    
    If text_prod_cd.Text = "" Then
        Call MsgBox("请选择订单产品!" & Chr(10) & "请选择。", vbExclamation + vbOKOnly, "警告")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
  '  squery = "{call ACE1020P ('C1','" + sProd_kind + "','" + slab_confirm.kqnum + "','" + sUserID + "',?)}"

'-------------------------------------------------------------------------
        
        minSIZEthk = sidbEdit_size_Athk.Value
        maxSIZEthk = sidbEdit_size_Bthk.Value
        minSIZEwid = sidbEdit_size_Awid.Value
        maxSIZEwid = sidbEdit_size_Bwid.Value
        minSIZElen = sidbEdit_size_Alen.Value
        maxSIZElen = sidbEdit_size_Blen.Value
        
'    If UDate_DEL_TO_b.RawData = "" Then
'        maxDATE = "99991231"
'    Else
        maxDATE = UDate_DEL_TO_b.RawData
'    End If
    
    If maxSIZEthk >= minSIZEthk Then
        If maxSIZEwid >= minSIZEwid Then
            If maxSIZElen >= minSIZElen Then
       
           SMESG = Gf_Ms_NeceCheck(nControl)
                If SMESG = "OK" Then
                
                    SMESG = Gf_Ms_NeceCheck2(mControl)
                    If SMESG = "OK" Then
            
                        M_CN1.CursorLocation = adUseServer
                          '---------squery(CALL ACE1020P)----------------------
                        
                        If opt_Prod_Kind(0) Then
                            sProd_Fl = "1"      'Semi Product
                        Else
                            sProd_Fl = "2"      'Product
                        End If
                        sQuery = "{call ACE1020P('C1','" & text_prod_cd.Text & "','" & txt_cust_cd.Text & "','','" & txt_stlgrd.Text & "','" + txt_stdspec.Text + "','" + txt_ORD_NO.Text + "','" + cbo_ord_item.Text + "','" + maxDATE + "'," & minSIZEthk & "," & maxSIZEthk & "," & minSIZEwid & "," & maxSIZEwid & "," & minSIZElen & "," & maxSIZElen & ",'" + sUserID + "',?)}"
                
             '   Debug.Print squery
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
                        
                    Else
                        SMESG = SMESG + " Must input according to length of item"
                        Call Gp_MsgBoxDisplay(SMESG)
                    End If
                
                Else
                    SMESG = SMESG + " Must input necessarily"
                    Call Gp_MsgBoxDisplay(SMESG)
                End If

            Else
                Call MsgBox("长度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
            End If
        Else
            Call MsgBox("宽度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
        End If
    Else
        Call MsgBox("厚度区间不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
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

Private Sub Command_REP_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call ACE1070P (?)}"
    
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
    Else
        Call Gp_MsgBoxDisplay("替代处理完了..!!", "I")
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call Form_Button_Edit
'    With MDIMain.MenuTool
'        .Buttons(4).Enabled = True                 'Save
'        .Buttons(9).Enabled = True                 'Delete
'    End With
  
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
    'sAuthority = "0001"
    
    Call Form_Define
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
   ' Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call Form_Button_Edit
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    If Mid(sAuthority, 3, 1) <> "1" Then
       Command_ALLSELECT.Enabled = False
       Command_REP.Enabled = False
       cmd_confirm.Enabled = False
    End If
    
    text_prod_cd.Text = "PP"
    Call Text_PROD_CD_KeyUp(0, 0)
    
    'ERP I/F START CHECK -- SCREEN HIDDEN
    If Gf_ErpSystem_Chk Then
        opt_Prod_Kind(1).Enabled = False
    End If

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Button_Edit()

    MDIMain.MenuTool.Buttons(7).Enabled = False              'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False              'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False              'Row Cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False             'Row Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False             'Row Paste
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
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
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Form_Button_Edit
    End If
    
    UDate_DEL_TO_b.RawData = ""
'    txt_stlgrd.Text = ""
    
    text_prod_cd.Text = "PP"
    Call Text_PROD_CD_KeyUp(0, 0)
   
    sidbEdit_size_Athk = 0
    sidbEdit_size_Bthk = 9999.99
    sidbEdit_size_Awid.Text = 0
    sidbEdit_size_Bwid.Text = 9999
    sidbEdit_size_Alen.Text = 0
    sidbEdit_size_Blen.Text = 9999999
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    Dim iRow As Integer
    
    If Gf_Sp_ProceExist(sc1.Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl")) Then
        ss1.OperationMode = OperationModeNormal
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Form_Button_Edit
        
    End If


    With ss1
        For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = SPD_SIZE_KND
            If Trim(.Text) <> "定尺" Then
                .Col = SPD_PROD_LEN:    .Lock = False
                Call Gp_Sp_CellColor(ss1, SPD_PROD_LEN, iRow, , &HC0FFFF)
            Else
                .Col = SPD_PROD_LEN:    .Lock = True
                Call Gp_Sp_CellColor(ss1, SPD_PROD_LEN, iRow)
            End If
        Next iRow
    End With
    
End Sub

Public Sub Form_Pro()
    
    If Gf_Sp_Process(M_CN1, sc1, Mc1, False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Form_Button_Edit
    End If
    
End Sub

Public Sub Spread_Column1sSort()

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


Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim SMESG As String

    If ss1.MaxRows < 1 Or Row < 1 Then Exit Sub
    
    If Col <> SPD_CNF_WGT And Col <> SPD_PROD_LEN Then
        Unload ACE1030C
        Load ACE1030C
        
        ss1.Row = Row
        
        ss1.Col = SPD_ORD_NO
        ACE1030C.txt_ORD_NO.Text = ss1.Text
        
        ss1.Col = SPD_ORD_ITEM
        ACE1030C.Combo_ORD_ITEM.Text = Trim(ss1.Value)
        
        ss1.Col = SPD_PROD_CD
        ACE1030C.text_prod_cd.Text = ss1.Text
        
        ACE1030C.Active_CForm = "ACE1030C"
        
        ACE1030C.Show
        ACE1030C.SetFocus
        
    Else
        ss1.Col = SPD_END_WGT
        ss1.Row = Row
        If ss1.Value <> 0 Then
            SMESG = "已经进行替代不能修改！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
        End If
       
    End If
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
   If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
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

Private Sub Text_PROD_CD_Change()
'    ULabel2.Caption = "钢种"
    Select Case text_prod_cd.Text
'           Case "S", "s", "SL"
'               Text_PROD_CD.Text = "SL"
           Case "P", "p", "PP"
               text_prod_cd.Text = "PP"
'                ULabel2.Caption = "标准号"
           Case "H", "h", "HC"
               text_prod_cd.Text = "HC"
'                ULabel2.Caption = "标准号"
           Case "", "**"
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
        DD.rControl.Add Item:=Text_PROD_CD_mate
        DD.nameType = "1"
        Call Gf_Common_DD(M_CN1, KeyCode)

    Else

        If Len(Trim(text_prod_cd.Text)) = text_prod_cd.MaxLength Then
            Text_PROD_CD_mate.Text = Gf_ComnNameFind(M_CN1, "B0005", text_prod_cd.Text, 1)
        Else
            Text_PROD_CD_mate.Text = ""
        End If
    
    End If
    
End Sub

Private Sub Text_size_knd_DblClick()

    Call Text_size_knd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_cust_cd_DblClick()

    Call txt_cust_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TXT_ORD_NO_Change()

    txt_ORD_NO.Text = Replace(txt_ORD_NO, vbCrLf, "")
    
End Sub

Private Sub txt_stdspec_Change()

    txt_stdspec.Text = Replace(txt_stdspec, vbCrLf, "")
    
End Sub

Private Sub txt_stdspec_DblClick()

    Call txt_stdspec_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stdspec_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then
            
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec
        
        Call Gf_StdSPEC_DD(M_CN1, KeyCode)
        
    End If
        
End Sub

Private Sub Text_size_knd_Change()

    If Len(Trim(Text_size_knd.Text)) = Text_size_knd.MaxLength Then
        Text_size_knd_name.Text = Gf_ComnNameFind(M_CN1, "B0043", Text_size_knd.Text, 2)
        Exit Sub
    Else
        Text_size_knd_name.Text = ""
    End If
    
End Sub

Private Sub Text_size_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"
        DD.rControl.Add Item:=Text_size_knd
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(Text_size_knd.Text)) = Text_size_knd.MaxLength Then
            Text_size_knd_name.Text = Gf_ComnNameFind(M_CN1, "B0043", Text_size_knd.Text, 2)
        Else
            Text_size_knd_name.Text = ""
        End If
    End If
    
End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"

        DD.rControl.Add Item:=txt_cust_cd
        DD.rControl.Add Item:=TXT_CUST_DES

        DD.nameType = "2"

        Call Gf_Customer_DD(M_CN1, KeyCode)

    Else
    
        If txt_cust_cd.Text <> "" Then
            TXT_CUST_DES.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
        Else
            TXT_CUST_DES.Text = ""
        End If
        
    End If

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    Dim i As Integer
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    Dim Row1 As Long
    Dim Row2 As Long
    Dim Col As Long
    
    Col = BlockCol
    Row1 = BlockRow
    Row2 = BlockRow2
    
    If Col = -1 Then
    
        With ss1
        
            For i = BlockRow To BlockRow2
                .Row = i
                .Col = 0
                If .Text = "Delete" Then
                   .Text = ""
                    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row1, Row2)
                    Call Gp_Sp_BlockColor(ss1, SPD_CNF_WGT, SPD_CNF_WGT, Row1, Row2, , &HC0FFFF)
                Else: .Text = "Delete"
                    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row1, Row2, , &HFFFF80)
                End If
               
            Next
         
        End With
        
   End If
     
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub
  
Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String

    If Len(Trim(txt_ORD_NO.Text)) = txt_ORD_NO.MaxLength Then
    
        If cbo_ord_item.Text <> "" Then Exit Sub
        
        txt_ORD_NO.Text = StrConv(txt_ORD_NO.Text, vbUpperCase)
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(txt_ORD_NO.Text) & "'"
        Call Gf_ComboAdd(M_CN1, cbo_ord_item, sQuery)

    Else
        cbo_ord_item.Clear
    End If

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
    
        If Len(Trim(txt_stlgrd.Text)) = txt_stlgrd.MaxLength Then
            txt_STLGRD_Name.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stlgrd.Text))
        Else
            txt_STLGRD_Name.Text = ""
        End If
        
    End If
    
End Sub
