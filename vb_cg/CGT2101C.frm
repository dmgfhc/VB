VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGT2101C 
   Caption         =   "物料全息查询_CGT2101C"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   14835
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9105
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   16060
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      PaneTree        =   "CGT2101C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   7755
         Left            =   0
         TabIndex        =   1
         Top             =   1350
         Width           =   15255
         _Version        =   393216
         _ExtentX        =   26908
         _ExtentY        =   13679
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
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
         MaxCols         =   139
         MaxRows         =   11
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGT2101C.frx":0052
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   1290
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   2275
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.ComboBox CBO_PRODGRD 
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
            ItemData        =   "CGT2101C.frx":304A
            Left            =   1275
            List            =   "CGT2101C.frx":3060
            TabIndex        =   29
            Tag             =   "等级"
            Top             =   495
            Width           =   1065
         End
         Begin VB.ComboBox CBO_SURFGRD 
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
            ItemData        =   "CGT2101C.frx":309A
            Left            =   4725
            List            =   "CGT2101C.frx":30B0
            TabIndex        =   28
            Tag             =   "等级"
            Top             =   495
            Width           =   1065
         End
         Begin VB.TextBox TXT_ORD_ITEM 
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
            MaxLength       =   2
            TabIndex        =   27
            Top             =   120
            Width           =   675
         End
         Begin VB.ComboBox CBO_GROUP 
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
            ItemData        =   "CGT2101C.frx":30EA
            Left            =   11640
            List            =   "CGT2101C.frx":30FA
            TabIndex        =   17
            Tag             =   "班别"
            Top             =   870
            Width           =   795
         End
         Begin VB.TextBox TXT_CD 
            Height          =   315
            Left            =   14580
            TabIndex        =   14
            Top             =   810
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox TXT_SP_CD 
            Alignment       =   2  'Center
            Height          =   270
            Left            =   15720
            TabIndex        =   11
            Top             =   960
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox TXT_TRNS_CMPY_CD 
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
            MaxLength       =   4
            TabIndex        =   9
            Top             =   495
            Width           =   1305
         End
         Begin VB.TextBox TXT_CUST_CD 
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
            MaxLength       =   6
            TabIndex        =   8
            Top             =   120
            Width           =   1305
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
            Left            =   4725
            MaxLength       =   11
            TabIndex        =   7
            Top             =   120
            Width           =   1305
         End
         Begin VB.TextBox TXT_SLAB_NO 
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
            Left            =   1275
            MaxLength       =   10
            TabIndex        =   6
            Top             =   120
            Width           =   1305
         End
         Begin VB.TextBox TXT_STDSPEC 
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
            Left            =   10620
            MaxLength       =   18
            TabIndex        =   3
            Tag             =   "标准号"
            Top             =   495
            Width           =   2505
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   1
            Left            =   9495
            Top             =   495
            Width           =   1125
            _ExtentX        =   1984
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
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   120
            Top             =   495
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            Caption         =   "产品等级"
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
            Left            =   3570
            Top             =   495
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            Caption         =   "表面等级"
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
            Top             =   120
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
            Left            =   3570
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
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   7020
            Top             =   120
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            Caption         =   "用户代码"
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
            Left            =   7020
            Top             =   495
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            Caption         =   "分段号"
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
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Left            =   9495
            Top             =   120
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            Caption         =   "轧制日期"
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
         Begin Threed.SSOption OPT_MILL 
            Height          =   330
            Left            =   13980
            TabIndex        =   12
            Top             =   930
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   2
            ForeColor       =   8421504
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
            Caption         =   "轧钢"
         End
         Begin Threed.SSOption OPT_SMSMILL 
            Height          =   330
            Left            =   12900
            TabIndex        =   13
            Top             =   930
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   2
            ForeColor       =   8421504
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
            Caption         =   "综合"
            Value           =   -1
         End
         Begin InDate.UDate TXT_MILL_DATE 
            Height          =   315
            Left            =   10620
            TabIndex        =   15
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
         Begin InDate.UDate TXT_MILL_DATE_TO 
            Height          =   315
            Left            =   12375
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
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   10485
            Top             =   870
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            Caption         =   "班别"
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
            Left            =   120
            Top             =   870
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
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   3570
            Top             =   870
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
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   7020
            Top             =   870
            Width           =   1110
            _ExtentX        =   1958
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
         Begin CSTextLibCtl.sidbEdit TXT_WID 
            Height          =   315
            Left            =   4725
            TabIndex        =   18
            Top             =   870
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
         Begin CSTextLibCtl.sidbEdit TXT_LEN 
            Height          =   315
            Left            =   8160
            TabIndex        =   19
            Top             =   870
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
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.0"
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
            NumDecDigits    =   1
            NumIntDigits    =   6
            ShowZero        =   0   'False
            MaxValue        =   999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit TXT_THK 
            Height          =   315
            Left            =   1275
            TabIndex        =   20
            Top             =   870
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
            NumIntDigits    =   3
            ShowZero        =   0   'False
            MaxValue        =   999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit TXT_THK_TO 
            Height          =   315
            Left            =   2370
            TabIndex        =   21
            Top             =   870
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
            NumIntDigits    =   3
            ShowZero        =   0   'False
            MaxValue        =   999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit TXT_WID_TO 
            Height          =   315
            Left            =   5820
            TabIndex        =   22
            Top             =   870
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
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.000"
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
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit TXT_LEN_TO 
            Height          =   315
            Left            =   9270
            TabIndex        =   23
            Top             =   870
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
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.0"
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
            NumDecDigits    =   1
            NumIntDigits    =   6
            ShowZero        =   0   'False
            MaxValue        =   999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSPanel SSP6 
            Height          =   315
            Left            =   13950
            TabIndex        =   30
            Top             =   480
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   8454143
            BackColor       =   16711935
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "出口订单"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSP5 
            Height          =   315
            Left            =   13950
            TabIndex        =   31
            Top             =   90
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   8454143
            BackColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "定制配送"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   9120
            TabIndex        =   26
            Top             =   1020
            Width           =   195
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   5700
            TabIndex        =   25
            Top             =   1020
            Width           =   195
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   2250
            TabIndex        =   24
            Top             =   1020
            Width           =   195
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   12180
            TabIndex        =   10
            Top             =   270
            Width           =   195
         End
      End
   End
   Begin CSTextLibCtl.sitxEdit SDT_PROD_DATE 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   2100
      Width           =   1200
      _Version        =   262145
      _ExtentX        =   2117
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__ __-__-__"
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
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   "____-__-__"
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
      Mask            =   "____-__-__"
      Justification   =   1
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   0
      ValidateMask    =   0   'False
   End
   Begin CSTextLibCtl.sitxEdit SDT_PROD_TO_DATE 
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   2130
      Width           =   1200
      _Version        =   262145
      _ExtentX        =   2117
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__ __-__-__"
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
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   "____-__-__"
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
      Mask            =   "____-__-__"
      Justification   =   1
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   0
      ValidateMask    =   0   'False
   End
End
Attribute VB_Name = "CGT2101C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Production DayReport Final Steel Grade
'-- Sub_System Name
'-- Program Name
'-- Program ID        CGT2101C
'-- Document No       Q-00-0010(Specification)
'-- Designer          LiQian
'-- Coder             LiQian
'-- Date              2011.05.05
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

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim ls_PChangeName                  'To Record P control Name

Const SS1_SLENP = 9                   '计划长度
Const SS1_RM_CR_STAGE3_TIME = 13      '订单数量
Const SS1_ORD_NO = 17                 '订单号
Const SS1_ORD_ITEM = 18               '订单序号
Const SS1_CUST_CD_CODE = 19           '用户名称
Const SS1_COOLING_TIME = 24           '堆冷时间
Const SS1_CHA_UNCHA_IND = 25          '再装炉指示
Const SS1_PRE_TOP_SLAB_TEMP = 30      '预热段温度上表
Const SS1_PRE_BOT_SLAB_TEMP = 31      '预热段温度下表
Const SS1_HT_TOP_SLAB_TEMP = 32       '加热段温度下限
Const SS1_HT_BOT_SLAB_TEMP_TEG = 33   '加热段温度上限
Const SS1_HT_BOT_SLAB_TEMP = 34       '加热段温度实绩
Const SS1_HT_ZONE_TIME = 35           '加热段温度驻段时间
Const SS1_EXT_TEMP_TEG = 45           '出炉温度目标
Const SS1_EXT_TEMP = 46               '出炉温度实绩
Const SS1_PDT_UNI_TEMP_TEG = 55       '温度均匀性目标
Const SS1_PDT_UNI_TEMP = 56           '温度均匀性实绩
Const SS1_DISCHARGE_DATE = 57         '出炉时间
Const SS1_GAS = 58                    '煤气热值
Const SS1_O2 = 59                     '炉内残氧
Const SS1_HT_TEMP1 = 60               '除鳞压力开始
Const SS1_TEMP2 = 61                  '除鳞压力终了
Const SS1_T1 = 62                     '除鳞出口温度
Const SS1_RM_MILL_END_AIM_TEMP = 65   '粗轧终轧温度目标
Const SS1_RM_MILL_END_AVE_TEMP = 66   '粗轧终轧温度实绩
Const SS1_CR_STAGE3_TIME = 68         '粗轧机轧制模式
Const SS1_RM_AVE_WID = 72             '宽度(平均测宽)
Const SS1_RM_SLAB_MILL_LEN = 73       '长度(计算)
Const SS1_T12 = 74                    '气雾冷却开冷温度目标
Const SS1_T13 = 75                    '气雾冷却开冷温度实绩
Const SS1_T14 = 76                    '气雾冷却终冷温度目标
Const SS1_T15 = 77                    '气雾冷却终冷温度实绩
Const SS1_T16 = 78                    '气雾冷却速度目标
Const SS1_RM_COOL_RATE = 79           '气雾冷却速度实绩
Const SS1_T20 = 82                    '精轧开轧厚度比目标
Const SS1_T21 = 83                    '精轧开轧厚度比实绩
Const SS1_ROLLING_METHOD = 88         '精轧机轧制模式
Const SS1_AIM_THK = 91                '目标厚度
Const SS1_T32 = 98                    'ACC冷却终冷温度目标
Const SS1_EXT_STK_TEMP = 99           'ACC冷却终冷温度实绩
Const SS1_ACC_UD_QT_RT = 100           '上下流量比
Const SS1_HT_T35 = 101                 'ACC冷却速度目标
Const SS1_COOL_RATE = 102              'ACC冷却速度实绩

Const SS1_T40 = 113                    '矩形度
Const SS1_T41 = 114                    '不平度
Const SS1_SIZE_KND = 116               '定尺
Const SS1_PROD_GRD = 117               '产品等级
Const SS1_SURF_GRD = 118              '表面等级
Const SS1_T42 = 119                   '缺陷
Const SS1_SLAB_NO1 = 120              '设计成材率
Const SS1_SLAB_NO2 = 121              '实设成材率
Const SS1_T43 = 122                   '实际成材率
Const SS1_YP_RST = 123                '力学性能屈服
Const SS1_TS_RST = 124                '力学性能抗拉
Const SS1_EL_RST = 125                '力学性能延伸率
Const SS1_IMPACT_RST_AVE = 126        '力学性能冲击均值
Const SS1_DWTT_YP_RST = 127           '力学性能DWTT
Const SS1_HTM_METHOD = 128            '热处理实绩热处理方式
Const SS1_HEAT_RATIO = 129            '热处理实绩加热速率
Const SS1_HT_TEMP = 130               '热处理实绩加热温度
Const SS1_UNIFORM_DT = 131            '热处理实绩保温时间
Const SS1_COL_OUT_TEMP = 132          '热处理实绩出炉温度

Const SS1_HT_BOT_SLAB_AIM_TEMP2 = 36
Const SS1_HT_BOT_SLAB_TEMP2 = 37
Const SS1_HT_ZONE_TIME2 = 38

Const SS1_FLAG = 138
Const SS1_EXPORT = 139

Const SS1_PLATE_NO = 1
Private Sub Form_Define()

   Dim i As Integer
   Dim iRow As Integer
   
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Hsheet"
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_slab_no, "p", " ", " ", " i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '板坯号
          Call Gp_Ms_Collection(TXT_ORD_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '订单号
        Call Gp_Ms_Collection(TXT_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '序列
         Call Gp_Ms_Collection(TXT_CUST_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '客户代码
       Call Gp_Ms_Collection(TXT_MILL_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '轧制日期
    Call Gp_Ms_Collection(TXT_MILL_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '轧制日期
         Call Gp_Ms_Collection(TXT_STDSPEC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '标准
         Call Gp_Ms_Collection(CBO_PRODGRD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '产品等级
         Call Gp_Ms_Collection(CBO_SURFGRD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '表面等级
    Call Gp_Ms_Collection(TXT_TRNS_CMPY_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '分段号
             Call Gp_Ms_Collection(txt_thk, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '厚
          Call Gp_Ms_Collection(TXT_THK_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '厚
             Call Gp_Ms_Collection(txt_wid, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '宽
          Call Gp_Ms_Collection(TXT_WID_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '宽
             Call Gp_Ms_Collection(txt_len, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '长
          Call Gp_Ms_Collection(TXT_LEN_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '长
           Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '班别
           
             ' Call Gp_Ms_Collection(TXT_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) ' 综合/轧钢, 不含炼钢信息
     
     'MASTER Collection
    Mc1.Add Item:="CGT2101C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="CGT2101C.P_MREFER", Key:="P-R"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGT2101C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").ROW = 0
    sc1.Item("Spread").Text = "◎"
       
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    'FormType = "Sheet"
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    Call Gp_Sp_ColHidden(ss1, SS1_HT_BOT_SLAB_AIM_TEMP2, True)  'modify by LiQian at 2012.09.17 加热二段无数据,不显示
    Call Gp_Sp_ColHidden(ss1, SS1_HT_BOT_SLAB_TEMP2, True)
    Call Gp_Sp_ColHidden(ss1, SS1_HT_ZONE_TIME2, True)
    Call Gp_Sp_ColHidden(ss1, SS1_SLENP, True)                       '计划长度
    Call Gp_Sp_ColHidden(ss1, SS1_RM_CR_STAGE3_TIME, True)           '订单数量
    Call Gp_Sp_ColHidden(ss1, SS1_ORD_NO, True)                      '订单号
    Call Gp_Sp_ColHidden(ss1, SS1_ORD_ITEM, True)                    '订单序号
    Call Gp_Sp_ColHidden(ss1, SS1_CUST_CD_CODE, True)                '用户名称
    Call Gp_Sp_ColHidden(ss1, SS1_COOLING_TIME, True)                '堆冷时间
    Call Gp_Sp_ColHidden(ss1, SS1_CHA_UNCHA_IND, True)               '再装炉指示
    Call Gp_Sp_ColHidden(ss1, SS1_PRE_TOP_SLAB_TEMP, True)           '预热段温度上表
    Call Gp_Sp_ColHidden(ss1, SS1_PRE_BOT_SLAB_TEMP, True)           '预热段温度下表
    Call Gp_Sp_ColHidden(ss1, SS1_EXT_TEMP_TEG, True)                '出炉温度目标
    Call Gp_Sp_ColHidden(ss1, SS1_EXT_TEMP, True)                    '出炉温度实绩
    Call Gp_Sp_ColHidden(ss1, SS1_PDT_UNI_TEMP_TEG, True)            '温度均匀性目标
    Call Gp_Sp_ColHidden(ss1, SS1_PDT_UNI_TEMP, True)                '温度均匀性目标
    Call Gp_Sp_ColHidden(ss1, SS1_DISCHARGE_DATE, True)              '出炉时间
    Call Gp_Sp_ColHidden(ss1, SS1_GAS, True)                         '煤气热值
    Call Gp_Sp_ColHidden(ss1, SS1_O2, True)                          '炉内残氧
    Call Gp_Sp_ColHidden(ss1, SS1_RM_MILL_END_AIM_TEMP, True)        '粗轧终轧温度目标
    Call Gp_Sp_ColHidden(ss1, SS1_RM_MILL_END_AVE_TEMP, True)        '粗轧终轧温度实绩
    Call Gp_Sp_ColHidden(ss1, SS1_CR_STAGE3_TIME, True)              '粗轧机轧制模式
    Call Gp_Sp_ColHidden(ss1, SS1_RM_AVE_WID, True)                  '宽度(平均测宽)
    Call Gp_Sp_ColHidden(ss1, SS1_RM_SLAB_MILL_LEN, True)            '长度(计算)
    Call Gp_Sp_ColHidden(ss1, SS1_ROLLING_METHOD, True)              '精轧机轧制模式
    Call Gp_Sp_ColHidden(ss1, SS1_AIM_THK, True)                     '目标厚度
    Call Gp_Sp_ColHidden(ss1, SS1_T32, True)                         'ACC冷却终冷温度目标
    Call Gp_Sp_ColHidden(ss1, SS1_EXT_STK_TEMP, True)                'ACC冷却终冷温度实绩
    Call Gp_Sp_ColHidden(ss1, SS1_ACC_UD_QT_RT, True)                '上下流量比
    Call Gp_Sp_ColHidden(ss1, SS1_HT_T35, True)                      'ACC冷却速度目标
    Call Gp_Sp_ColHidden(ss1, SS1_COOL_RATE, True)                   'ACC冷却速度实绩
    Call Gp_Sp_ColHidden(ss1, SS1_HT_TEMP1, True)                    '除鳞压力开始
    Call Gp_Sp_ColHidden(ss1, SS1_TEMP2, True)                       '除鳞压力终了
    Call Gp_Sp_ColHidden(ss1, SS1_T1, True)                          '除鳞出口温度
    Call Gp_Sp_ColHidden(ss1, SS1_T12, True)                         '气雾冷却开冷温度目标
    Call Gp_Sp_ColHidden(ss1, SS1_T13, True)                         '气雾冷却开冷温度实绩
    Call Gp_Sp_ColHidden(ss1, SS1_T14, True)                         '气雾冷却终冷温度目标
    Call Gp_Sp_ColHidden(ss1, SS1_T15, True)                         '气雾冷却终冷温度实绩
    Call Gp_Sp_ColHidden(ss1, SS1_T16, True)                         '气雾冷却速度目标
    Call Gp_Sp_ColHidden(ss1, SS1_RM_COOL_RATE, True)                '气雾冷却速度实绩
    Call Gp_Sp_ColHidden(ss1, SS1_T20, True)                         '精轧开轧厚度比目标
    Call Gp_Sp_ColHidden(ss1, SS1_T21, True)                         '精轧开轧厚度比实绩
    Call Gp_Sp_ColHidden(ss1, SS1_HT_T35, True)                      'ACC冷却速度目标
    Call Gp_Sp_ColHidden(ss1, SS1_COOL_RATE, True)                   'ACC冷却速度实绩
    Call Gp_Sp_ColHidden(ss1, SS1_T40, True)                         '矩形度
    Call Gp_Sp_ColHidden(ss1, SS1_T41, True)                         '不平度
    Call Gp_Sp_ColHidden(ss1, SS1_SIZE_KND, True)                    '定尺
    Call Gp_Sp_ColHidden(ss1, SS1_PROD_GRD, True)                    '产品等级
    Call Gp_Sp_ColHidden(ss1, SS1_SURF_GRD, True)                    '表面等级
    Call Gp_Sp_ColHidden(ss1, SS1_T42, True)                         '缺陷
    Call Gp_Sp_ColHidden(ss1, SS1_SLAB_NO1, True)                    '设计成材率
    Call Gp_Sp_ColHidden(ss1, SS1_SLAB_NO2, True)                    '实设成材率
    Call Gp_Sp_ColHidden(ss1, SS1_T43, True)                         '实际成材率
    Call Gp_Sp_ColHidden(ss1, SS1_YP_RST, True)                      '力学性能屈服
    Call Gp_Sp_ColHidden(ss1, SS1_TS_RST, True)                      '力学性能抗拉
    Call Gp_Sp_ColHidden(ss1, SS1_EL_RST, True)                      '力学性能延伸率
    Call Gp_Sp_ColHidden(ss1, SS1_IMPACT_RST_AVE, True)              '力学性能冲击均值
    Call Gp_Sp_ColHidden(ss1, SS1_DWTT_YP_RST, True)                 '力学性能DWTT
    Call Gp_Sp_ColHidden(ss1, SS1_HTM_METHOD, True)                  '热处理实绩热处理方式
    Call Gp_Sp_ColHidden(ss1, SS1_HEAT_RATIO, True)                  '热处理实绩加热速率
    Call Gp_Sp_ColHidden(ss1, SS1_HT_TEMP, True)                     '热处理实绩加热温度
    Call Gp_Sp_ColHidden(ss1, SS1_UNIFORM_DT, True)                  '热处理实绩保温时间
    Call Gp_Sp_ColHidden(ss1, SS1_COL_OUT_TEMP, True)                '热处理实绩出炉温度
    
    CBO_GROUP.Clear
    
    CBO_GROUP.AddItem "A"
    CBO_GROUP.AddItem "B"
    CBO_GROUP.AddItem "C"
    CBO_GROUP.AddItem "D"
    
    'OPT_MILL.Value = True '不含炼钢信息
    
    Screen.MousePointer = vbDefault
   
End Sub

Public Sub Form_Cls()
    
    Call Gf_Sp_Cls(sc1)
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_Cls(Mc1("rControl"))
        
End Sub


Public Sub Form_Pro()

If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
     Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
End If
    
End Sub

Public Sub Form_Ref()
    
    Dim sFlag As String
    Dim sexport As String
    Dim iCount   As Integer

    If Len(txt_slab_no) <= 8 Then
        If Trim(TXT_MILL_DATE.RawData) = "" Or Trim(TXT_MILL_DATE_TO.RawData) = "" Then
           MsgBox "查询日期未输入!", vbCritical, "系统提示信息"
           Exit Sub
        End If
        
        If Trim(TXT_MILL_DATE.RawData) <> Trim(TXT_MILL_DATE_TO.RawData) Then
           MsgBox "只能查询一天内信息!", vbCritical, "系统提示信息"
           Exit Sub
        End If
    End If
        
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Nothing) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    With ss1
        For iCount = 1 To .MaxRows
        
            .ROW = iCount:
            .Col = SS1_FLAG:       sFlag = Trim(.Text)
            .Col = SS1_EXPORT:     sexport = Trim(.Text)
            
            '是否定制配送
            If sFlag = "Y" Then
               Call Gp_Sp_BlockColor(ss1, SS1_PLATE_NO, SS1_PLATE_NO, iCount, iCount, SSP5.BackColor)
            End If
            '是否出口订单
            
            If sexport = "Y" Then
               Call Gp_Sp_BlockColor(ss1, SS1_PLATE_NO, SS1_PLATE_NO, iCount, iCount, SSP6.BackColor)
            End If
        Next iCount
    End With
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
End Sub

Public Sub Sp_Setting(ByVal sPname As Variant)

    Dim iRow As Integer

    With sPname

        .RowHeight(-1) = 13

        .BackColorStyle = BackColorStyleUnderGrid

        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040

        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040


        .OperationMode = OperationModeNormal
        .RetainSelBlock = True
        .UserResize = UserResizeColumns

        .ProcessTab = True
        .ScrollBarExtMode = True
        .ButtonDrawMode = 1
        .TabStop = False

        .Col = 0: .Col2 = -1
        .ROW = 0: .Row2 = -1


        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False

        .Col = -1
        .ROW = 0
        .FontBold = True
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
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

'Private Sub OPT_MILL_Click(Value As Integer)
'
'    Dim iRow As Integer
'    Dim sTemp As String
'
'    If OPT_MILL.Value = True Then
'        OPT_MILL.ForeColor = &HFF&
'        OPT_SMSMILL.ForeColor = &H808080
'        TXT_CD = "M"
'        Call Gf_Sp_Cls(sc1)
'        For iRow = 1 To SPD_SMSMILL
'           ''''''''''''''''''''''''''   Call Gp_Sp_ColHidden(ss1, iRow, True)
'        Next iRow
'        For iRow = SPD_SMSMILL + 1 To SPD_MAX
'                Call Gp_Sp_ColHidden(ss1, iRow, False)
'        Next iRow
'    Else
'        OPT_MILL.ForeColor = &H808080
'        TXT_CD = "A"
'    End If
'
'End Sub

'Private Sub OPT_SMSMILL_Click(Value As Integer)
'
'    Dim iRow As Integer
'    Dim sTemp As String
'
'    If OPT_SMSMILL.Value = True Then
'        OPT_SMSMILL.ForeColor = &HFF&
'        OPT_MILL.ForeColor = &H808080
'        TXT_CD = "A"
'        Call Gf_Sp_Cls(sc1)
'        For iRow = 1 To SPD_MAX 'SPD_SMSMILL
'                Call Gp_Sp_ColHidden(ss1, iRow, False)
'        Next iRow
'    Else
'        OPT_SMSMILL.ForeColor = &H808080
'        TXT_CD = "M"
'    End If
'
'End Sub

Private Sub TXT_MILL_DATE_GotFocus()
     If TXT_MILL_DATE.RawData = "" Then
        TXT_MILL_DATE.RawData = Gf_DTSet(M_CN1, "D")
     End If
     If TXT_MILL_DATE_TO.RawData = "" Then
        TXT_MILL_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub TXT_MILL_DATE_TO_GotFocus()
     If TXT_MILL_DATE_TO.RawData = "" Then
        TXT_MILL_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").ROW = 0
    sc1.Item("Spread").Text = "◎"

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub ss1_LostFocus()
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub txt_stdspec_DblClick()
    Call txt_STDSPEC_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_STDSPEC

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub
