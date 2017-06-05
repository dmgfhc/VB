VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form EGC1050C 
   Caption         =   "火切实绩查询_EGC1050C"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9075
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   16007
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      Locked          =   -1  'True
      PaneTree        =   "EGC1050C.frx":0000
      Begin Threed.SSFrame Single 
         Height          =   1695
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   2990
         _Version        =   196609
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cbo_ord_fl 
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
            ItemData        =   "EGC1050C.frx":0052
            Left            =   4365
            List            =   "EGC1050C.frx":0054
            TabIndex        =   32
            Tag             =   "班别"
            Top             =   1320
            Width           =   1530
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
            Left            =   1485
            MaxLength       =   15
            TabIndex        =   30
            Tag             =   "移拨码单号"
            Top             =   1320
            Width           =   1680
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
            Left            =   11730
            MaxLength       =   2
            TabIndex        =   29
            Top             =   540
            Width           =   780
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
            Left            =   12540
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   540
            Width           =   2610
         End
         Begin VB.TextBox TXT_LOT_NO 
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
            MaxLength       =   14
            TabIndex        =   14
            Top             =   540
            Width           =   1680
         End
         Begin VB.ComboBox CBO_SHIFT 
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
            ItemData        =   "EGC1050C.frx":0056
            Left            =   5655
            List            =   "EGC1050C.frx":0063
            TabIndex        =   13
            Top             =   135
            Width           =   750
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
            Left            =   8055
            MaxLength       =   18
            TabIndex        =   12
            Tag             =   "标准号"
            Top             =   540
            Width           =   2380
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
            ItemData        =   "EGC1050C.frx":0070
            Left            =   11730
            List            =   "EGC1050C.frx":0089
            TabIndex        =   11
            Tag             =   "等级"
            Top             =   135
            Width           =   1065
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
            ItemData        =   "EGC1050C.frx":00CC
            Left            =   7575
            List            =   "EGC1050C.frx":00DC
            TabIndex        =   10
            Tag             =   "班别"
            Top             =   135
            Width           =   750
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
            Left            =   4725
            TabIndex        =   9
            Top             =   540
            Width           =   1680
         End
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
            ItemData        =   "EGC1050C.frx":00EC
            Left            =   12780
            List            =   "EGC1050C.frx":0102
            TabIndex        =   8
            Tag             =   "等级"
            Top             =   135
            Width           =   1065
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
            Left            =   13740
            TabIndex        =   7
            Tag             =   "标准代码"
            Top             =   945
            Width           =   1395
         End
         Begin VB.ComboBox CBO_MILLGROUP 
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
            ItemData        =   "EGC1050C.frx":013C
            Left            =   9690
            List            =   "EGC1050C.frx":014C
            TabIndex        =   6
            Tag             =   "班别"
            Top             =   135
            Width           =   750
         End
         Begin VB.ComboBox cbo_cut_place 
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
            ItemData        =   "EGC1050C.frx":015C
            Left            =   11730
            List            =   "EGC1050C.frx":015E
            TabIndex        =   5
            Tag             =   "班别"
            Top             =   945
            Width           =   1065
         End
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   105
            Top             =   540
            Width           =   1350
            _ExtentX        =   2381
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
         End
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   105
            Top             =   135
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            Caption         =   "切割时间"
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
         Begin InDate.ULabel ULabel13 
            Height          =   315
            Left            =   4740
            Top             =   135
            Width           =   885
            _ExtentX        =   1561
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
         Begin InDate.ULabel ULabel22 
            Height          =   300
            Index           =   4
            Left            =   6660
            Top             =   540
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   529
            Caption         =   "标准号"
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
         Begin InDate.UDate SDT_PROD_TO_DATE 
            Height          =   315
            Left            =   3090
            TabIndex        =   15
            Tag             =   "起始日期"
            Top             =   135
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
         End
         Begin InDate.UDate SDT_PROD_DATE 
            Height          =   315
            Left            =   1485
            TabIndex        =   16
            Tag             =   "起始日期"
            Top             =   135
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
         End
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   6660
            Top             =   135
            Width           =   885
            _ExtentX        =   1561
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
            Left            =   3345
            Top             =   540
            Width           =   1350
            _ExtentX        =   2381
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
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   10680
            Top             =   135
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            Caption         =   "表面/综合"
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
         Begin Threed.SSCommand CMD_CARD 
            Height          =   345
            Left            =   13860
            TabIndex        =   17
            Top             =   120
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   609
            _Version        =   196609
            ForeColor       =   255
            Caption         =   "码堆报告"
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   2
            Left            =   12810
            Top             =   945
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   556
            Caption         =   "垛位"
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
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   3345
            Top             =   945
            Width           =   990
            _ExtentX        =   1746
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
            Left            =   4365
            TabIndex        =   18
            Top             =   945
            Width           =   900
            _Version        =   262145
            _ExtentX        =   1587
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
            Modified        =   -1  'True
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
            FmtThousands    =   0
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
            Left            =   6675
            Top             =   945
            Width           =   1350
            _ExtentX        =   2381
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
         Begin CSTextLibCtl.sidbEdit SDB_LEN 
            Height          =   315
            Left            =   8055
            TabIndex        =   19
            Top             =   945
            Width           =   1050
            _Version        =   262145
            _ExtentX        =   1852
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
            FmtThousands    =   0
            FmtControl      =   1
            NumDecDigits    =   1
            NumIntDigits    =   6
            ShowZero        =   0   'False
            MaxValue        =   999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_WID_TO 
            Height          =   315
            Left            =   5505
            TabIndex        =   20
            Top             =   945
            Width           =   900
            _Version        =   262145
            _ExtentX        =   1587
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
            FmtThousands    =   0
            FmtControl      =   1
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_LEN_TO 
            Height          =   315
            Left            =   9390
            TabIndex        =   21
            Top             =   945
            Width           =   1050
            _Version        =   262145
            _ExtentX        =   1852
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
            FmtThousands    =   0
            FmtControl      =   1
            NumDecDigits    =   1
            NumIntDigits    =   6
            ShowZero        =   0   'False
            MaxValue        =   999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_THK 
            Height          =   315
            Left            =   1365
            TabIndex        =   22
            Top             =   940
            Width           =   765
            _Version        =   262145
            _ExtentX        =   1349
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
            Modified        =   -1  'True
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
            MaxValue        =   999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   105
            Top             =   940
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
         Begin CSTextLibCtl.sidbEdit SDB_THK_TO 
            Height          =   315
            Left            =   2400
            TabIndex        =   23
            Top             =   940
            Width           =   765
            _Version        =   262145
            _ExtentX        =   1349
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
            MaxValue        =   999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   8565
            Top             =   135
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "轧制班别"
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
            Left            =   10680
            Top             =   945
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            Caption         =   "切割场地"
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
            Left            =   10680
            Top             =   540
            Width           =   1020
            _ExtentX        =   1799
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
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   105
            Top             =   1320
            Width           =   1350
            _ExtentX        =   2381
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
         Begin InDate.ULabel ULabel14 
            Height          =   315
            Left            =   3345
            Top             =   1320
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   556
            Caption         =   "余材替代"
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
         Begin VB.Label Label6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   2220
            TabIndex        =   27
            Top             =   1065
            Width           =   195
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   2940
            TabIndex        =   26
            Top             =   255
            Width           =   195
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   5340
            TabIndex        =   25
            Top             =   1065
            Width           =   195
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   9195
            TabIndex        =   24
            Top             =   1065
            Width           =   195
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   7320
         Left            =   0
         TabIndex        =   31
         Top             =   1755
         Width           =   15225
         _Version        =   393216
         _ExtentX        =   26855
         _ExtentY        =   12912
         _StockProps     =   64
         ColsFrozen      =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   49
         MaxRows         =   501
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "EGC1050C.frx":0160
      End
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   1920
      Top             =   1620
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "检查时间"
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
   Begin InDate.UDate CHK_TO_DATE 
      Height          =   315
      Left            =   18390
      TabIndex        =   0
      Tag             =   "起始日期"
      Top             =   1620
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
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
   Begin InDate.UDate CHK_FROM_DATE 
      Height          =   315
      Left            =   16650
      TabIndex        =   1
      Tag             =   "起始日期"
      Top             =   1620
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
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
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   120
      Left            =   18180
      TabIndex        =   2
      Top             =   1710
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "EGC1050C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   ZB HTM System
'-- Program Name      火切实绩查询界面
'-- Program ID        EGC1050C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2010.7.23
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

Dim pColumn  As New Collection      'Spread Primary Key Collection
Dim nColumn  As New Collection      'Spread necessary Column Collection
Dim mColumn  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn  As New Collection      'Spread Insert Column Collection
Dim aColumn  As New Collection      'Master -> Spread Column Collection
Dim lColumn  As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_PLATE_NO = 1
Const SS1_PROC_CD = 8
Const SS1_STDSPEC = 9
Const SS1_AFER_SIZE = 12
Const SS1_AFER_WGT = 13
Const SS1_EMP_NO = 28
Const SS1_LOC = 29
Const SS1_REMARK = 31
Const SS1_CUST_CD = 37
Const SS1_ORD = 36
Const SS1_SIZE = 38


Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(TXT_LOT_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDT_PROD_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(SDT_PROD_TO_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(CHK_FROM_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(CHK_TO_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(CBO_SURFGRD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(CBO_PRODGRD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_loc, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'ADDED BY GUOLI AT 20080604
           Call Gp_Ms_Collection(SDB_THK, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDB_THK_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_WID, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDB_WID_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_LEN, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDB_LEN_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(CBO_MILLGROUP, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(cbo_cut_place, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_mv_lst_no, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(cbo_ord_fl, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
        Mc1.Add Item:=pControl, Key:="pControl"
        Mc1.Add Item:=nControl, Key:="nControl"
        Mc1.Add Item:=mControl, Key:="mControl"
        Mc1.Add Item:=iControl, Key:="iControl"
        Mc1.Add Item:=rControl, Key:="rControl"
        Mc1.Add Item:=cControl, Key:="cControl"
        Mc1.Add Item:=aControl, Key:="aControl"
        Mc1.Add Item:=lControl, Key:="lControl"

     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '试样号状态
     Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '委托单号
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)      '热处理指示 by 2011-09-07
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)      '矫直指示   by 2011-09-07
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)      '探伤指示   by 2011-09-07
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
 
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="EGC1050C.P_SREFER", Key:="P-R"
    sc1.Add Item:="EGC1050C.P_SONEROW", Key:="P-O"
    sc1.Add Item:="EGC1050C.P_SMODIFY", Key:="P-M"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"

     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0
     
End Sub

Private Sub CBO_SHIFT_Click()
    ss1.MaxRows = 0
End Sub

Private Sub CMD_CARD_Click()
    If SDT_PROD_DATE.RawData = "" Or SDT_PROD_TO_DATE.RawData = "" Or Trim(CBO_SHIFT.Text) = "" Then
        Call Gp_MsgBoxDisplay("切割时间和班次必须输入！！", "", "错误提示")
        Exit Sub
    End If
    
    Call ExcelPrn_Pile
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False                  'Row Insert

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

    Call Gp_Ms_Cls(Mc1("rControl"))
    SDT_PROD_DATE.Text = ""
    SDT_PROD_TO_DATE.Text = ""
    CHK_FROM_DATE.Text = ""
    CHK_TO_DATE.Text = ""
    
    cbo_cut_place.AddItem ""
    cbo_cut_place.AddItem "中板厂"
    cbo_cut_place.AddItem "修建部"
    cbo_cut_place.AddItem "港南库"
    cbo_cut_place.AddItem "板卷线" '20170511
    cbo_cut_place.AddItem "GT库" '20170522
    cbo_cut_place.ListIndex = 0
    
    cbo_ord_fl.AddItem ""
    cbo_ord_fl.AddItem "订单材-余材"
    cbo_ord_fl.AddItem "余材-订单材"
    cbo_ord_fl.AddItem "订单材-订单材"
    cbo_ord_fl.ListIndex = 0

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "EG-System.INI", Me.Name)
        
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "EG-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set iColumn = Nothing
    Set pColumn = Nothing
    Set lColumn = Nothing
    Set nColumn = Nothing
    Set mColumn = Nothing
    Set aColumn = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub
Public Sub Form_Exc()
    lBlkcol1 = 1: lBlkcol2 = ss1.MaxCols: lBlkrow1 = 0: lBlkrow2 = ss1.MaxRows
    
    Call Gp_Sp_Excel(Me, ss1, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub
Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
       Call Gp_Ms_Cls(Mc1("rControl"))
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       
       MDIMain.MenuTool.Buttons(7).Enabled = False                  'Row Insert
       
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
       SDT_PROD_DATE.Text = ""
       SDT_PROD_TO_DATE.Text = ""
       CHK_FROM_DATE.Text = ""
       CHK_TO_DATE.Text = ""
    End If

End Sub

Public Sub Form_Ref()
    
    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    If SDT_PROD_DATE.RawData = "" Or SDT_PROD_TO_DATE.RawData = "" Then
       SDT_PROD_DATE.RawData = Gf_DTSet(M_CN1, "D")
       SDT_PROD_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
    End If
    
    If Val(SDT_PROD_DATE.RawData) - Val(SDT_PROD_TO_DATE.RawData) > 0 Then
         sMesg = " 时间范围输入错误，请重新输入时间信息 ！！！"
         Call Gp_MsgBoxDisplay(sMesg)
         Exit Sub
    End If
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        MDIMain.MenuTool.Buttons(7).Enabled = False                  'Row Insert
    End If

End Sub

Public Sub Form_Pro()
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call Form_Ref
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub
Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Private Sub SDT_PROD_DATE_GotFocus()
     SDT_PROD_DATE.RawData = Gf_DTSet(M_CN1, "D")
     SDT_PROD_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
End Sub

Private Sub SDT_PROD_to_DATE_GotFocus()
     SDT_PROD_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        ss1.ROW = ss1.ActiveRow
        ss1.Col = SS1_EMP_NO
        ss1.Text = sUserID
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

        Exit Sub

    End If
End Sub

Private Sub ExcelPrn_Pile()

    Dim i               As Integer
    Dim j               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    Dim sDate           As String
    Dim sShift          As String
    
    Dim sPage_Num       As Integer
    Dim sPage_X         As Integer
    Dim sPage           As Double
    Dim sLastPage       As Double
    Dim sRow1           As Integer
    Dim sRow2           As Integer
    
    Dim xl_A            As String
    Dim xl_B            As String
    Dim xl_C            As String
    Dim xl_D            As String
    Dim xl_E            As String
    Dim xl_F            As String
    Dim xl_G            As String
    Dim xl_H            As String
    Dim xl_I            As String
    
    Dim xl_clr_body     As String
    Dim xl_clr_sum      As String
    
    Dim Xl_Cnt          As String
    Dim Xl_Wgt          As String
    Dim Xl_Wgt_Val      As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\EGC1050C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    sDate = SDT_PROD_DATE.Text
    
    If SDT_PROD_DATE.Text <> SDT_PROD_TO_DATE.Text Then
        xlApp.Range("A2").Value = "日期: " & Left(sDate, 4) + "年" + Mid(sDate, 6, 2) + "月" + Mid(sDate, 9, 2) + "日 - " + Mid(SDT_PROD_TO_DATE.Text, 9, 2) + "日"
    Else
        xlApp.Range("A2").Value = "日期: " & Left(sDate, 4) + "年" + Mid(sDate, 6, 2) + "月" + Mid(sDate, 9, 2) + "日"
    End If
    
    If CBO_SHIFT.Text = "1" Then
       sShift = "大夜班"
    ElseIf CBO_SHIFT.Text = "2" Then
       sShift = "白班"
    ElseIf CBO_SHIFT.Text = "3" Then
       sShift = "小夜班"
    Else
       sShift = ""
    End If
    
    xlApp.Range("C2").Value = Mid(xlApp.Range("C2").Value, 1, 3) & sShift

    xlApp.Range("D2").Value = "码堆员：" & sUserName
        
    sPage_Num = 31
    sPage_X = 32
    
    sPage = Int(ss1.MaxRows / sPage_Num) + 1
    sLastPage = ss1.MaxRows - Int(ss1.MaxRows / sPage_Num) * sPage_Num
    
    For i = 0 To 10
        xl_clr_body = "A" + CStr(4 + i * sPage_X) + ":" + "I" + CStr(33 + i * sPage_X)
        xl_clr_sum = "G" + CStr(34 + i * sPage_X) + ":" + "G" + CStr(35 + i * sPage_X)
        xlApp.Range(xl_clr_body).Value = Null
        xlApp.Range(xl_clr_sum).Value = Null
    Next i
    
    For i = 0 To sPage - 1
       
        sRow1 = 1 + sPage_Num * i
        sRow2 = sPage_Num * (i + 1)

        If i = sPage - 1 Then
           sRow2 = sPage_Num * i + sLastPage
        End If

        xl_A = "A" + CStr(4 + i * sPage_X)
        xl_B = "B" + CStr(4 + i * sPage_X)
        xl_C = "C" + CStr(4 + i * sPage_X)
        xl_D = "D" + CStr(4 + i * sPage_X)
        xl_E = "E" + CStr(4 + i * sPage_X)
        xl_F = "F" + CStr(4 + i * sPage_X)
        xl_G = "G" + CStr(4 + i * sPage_X)
        xl_H = "H" + CStr(4 + i * sPage_X)
        xl_I = "I" + CStr(4 + i * sPage_X)
        
        Xl_Cnt = "C" + CStr(3 + (i + 1) * sPage_X)
        Xl_Wgt = "G" + CStr(3 + (i + 1) * sPage_X)
        
        Clipboard.Clear
        ss1.SetSelection SS1_LOC, sRow1, SS1_LOC, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_A).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_STDSPEC, sRow1, SS1_STDSPEC, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_B).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_PROC_CD, sRow1, SS1_PROC_CD, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_C).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_PLATE_NO, sRow1, SS1_PLATE_NO, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_D).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_AFER_SIZE, sRow1, SS1_AFER_SIZE, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_E).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_AFER_WGT, sRow1, SS1_AFER_WGT, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_F).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_CUST_CD, sRow1, SS1_CUST_CD, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_G).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_ORD, sRow1, SS1_ORD, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_H).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
                
        Clipboard.Clear
        ss1.SetSelection SS1_SIZE, sRow1, SS1_SIZE, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_I).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
                        
                
        If i = sPage - 1 Then
           xlApp.Range(Xl_Cnt).Value = sLastPage
        Else
           xlApp.Range(Xl_Cnt).Value = sPage_Num
        End If
        
        For j = 1 To sPage_Num
            Xl_Wgt_Val = "F" & CStr((Val(Mid(xl_F, 2)) + j - 1))
            xlApp.Range(Xl_Wgt).Value = xlApp.Range(Xl_Wgt).Value + xlApp.Range(Xl_Wgt_Val).Value
        Next j
                      
    Next i
    
    ss1.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Exit Sub

ErrHandle:
    MsgBox Error
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
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

