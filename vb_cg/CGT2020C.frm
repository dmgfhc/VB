VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGT2020C 
   Caption         =   "ĸ��ֶ�ʵ����ѯ_CGT2020C"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   15135
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9135
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   16113
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      Locked          =   -1  'True
      PaneTree        =   "CGT2020C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   7020
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2115
         Width           =   15120
         _Version        =   393216
         _ExtentX        =   26670
         _ExtentY        =   12383
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ButtonDrawMode  =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   68
         MaxRows         =   20
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CGT2020C.frx":0052
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   2055
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   15120
         _ExtentX        =   26670
         _ExtentY        =   3625
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.TextBox TXT_TRNS_CMPY_CD_TO 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9810
            MaxLength       =   3
            TabIndex        =   39
            Top             =   1260
            Width           =   975
         End
         Begin VB.TextBox txt_STLGRD_Name 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3045
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   34
            Tag             =   "����(��׼��)"
            Top             =   1650
            Width           =   2025
         End
         Begin VB.TextBox txt_stlgrd 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1500
            MaxLength       =   12
            TabIndex        =   33
            Tag             =   "����"
            Top             =   1650
            Width           =   1515
         End
         Begin VB.TextBox TXT_TRNS_CMPY_CD 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8580
            MaxLength       =   3
            TabIndex        =   32
            Top             =   1260
            Width           =   975
         End
         Begin VB.TextBox TXT_MILL_LOT_NO 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1500
            MaxLength       =   14
            TabIndex        =   10
            Top             =   480
            Width           =   1740
         End
         Begin VB.TextBox TXT_SLAB_NO 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1500
            MaxLength       =   12
            TabIndex        =   9
            Top             =   870
            Width           =   1740
         End
         Begin VB.ComboBox CBO_SHIFT 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "CGT2020C.frx":1E71
            Left            =   6390
            List            =   "CGT2020C.frx":1E73
            TabIndex        =   8
            Top             =   90
            Width           =   735
         End
         Begin VB.TextBox txt_stdspec 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1500
            MaxLength       =   18
            TabIndex        =   7
            Tag             =   "��׼��"
            Top             =   1260
            Width           =   2445
         End
         Begin VB.ComboBox cbo_gas_fl 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "CGT2020C.frx":1E75
            Left            =   6390
            List            =   "CGT2020C.frx":1E77
            TabIndex        =   6
            Top             =   480
            Width           =   735
         End
         Begin VB.ComboBox cbo_plate_yn 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "CGT2020C.frx":1E79
            Left            =   6390
            List            =   "CGT2020C.frx":1E7B
            TabIndex        =   5
            Top             =   870
            Width           =   735
         End
         Begin VB.ComboBox cbo_scrap 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "CGT2020C.frx":1E7D
            Left            =   6390
            List            =   "CGT2020C.frx":1E7F
            TabIndex        =   4
            Top             =   1260
            Width           =   735
         End
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   105
            Top             =   870
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "������"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   110
            Top             =   90
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "�и�ʱ��"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   4995
            Top             =   90
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "���"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.UDate SDT_PROD_DATE_FROM 
            Height          =   315
            Left            =   1500
            TabIndex        =   11
            Tag             =   "��ʼ����"
            Top             =   90
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   3225
            TabIndex        =   12
            Tag             =   "��ʼ����"
            Top             =   90
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   105
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "������"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   7410
            Top             =   480
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "ĸ���"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   8580
            TabIndex        =   13
            Top             =   480
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   7410
            Top             =   870
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "ĸ�峤"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   8580
            TabIndex        =   14
            Top             =   870
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin CSTextLibCtl.sidbEdit SDB_WID_TO 
            Height          =   315
            Left            =   9810
            TabIndex        =   15
            Top             =   480
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin CSTextLibCtl.sidbEdit SDB_LEN_TO 
            Height          =   315
            Left            =   9810
            TabIndex        =   16
            Top             =   870
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin CSTextLibCtl.sidbEdit SDB_THK 
            Height          =   315
            Left            =   8580
            TabIndex        =   17
            Top             =   90
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   1
            Left            =   105
            Top             =   1260
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "��׼��"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16711680
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   7410
            Top             =   90
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "ĸ���"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   9810
            TabIndex        =   18
            Top             =   90
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   4995
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "�и�ָʾ"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   11130
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "��������"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CSTextLibCtl.sidbEdit SDB_WID_PLAN 
            Height          =   315
            Left            =   12525
            TabIndex        =   19
            Top             =   480
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin InDate.ULabel ULabel10 
            Height          =   315
            Left            =   11130
            Top             =   870
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "��������"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CSTextLibCtl.sidbEdit SDB_LEN_PLAN 
            Height          =   315
            Left            =   12525
            TabIndex        =   20
            Top             =   870
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin CSTextLibCtl.sidbEdit SDB_WID_PLAN_TO 
            Height          =   315
            Left            =   13755
            TabIndex        =   21
            Top             =   480
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin CSTextLibCtl.sidbEdit SDB_LEN_PLAN_TO 
            Height          =   315
            Left            =   13755
            TabIndex        =   22
            Top             =   870
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin CSTextLibCtl.sidbEdit SDB_THK_PLAN 
            Height          =   315
            Left            =   12525
            TabIndex        =   23
            Top             =   90
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Left            =   11130
            Top             =   90
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "�������"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CSTextLibCtl.sidbEdit SDB_THK_PLAN_TO 
            Height          =   315
            Left            =   13755
            TabIndex        =   24
            Top             =   90
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   4995
            Top             =   870
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "���иְ�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   4995
            Top             =   1260
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "�Ƿ��з�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   7410
            Top             =   1260
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "�ֶκ�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.ULabel ULabel15 
            Height          =   315
            Left            =   105
            Top             =   1650
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "ԭʼ���ϸ���"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16711680
         End
         Begin Threed.SSPanel SSPpdt 
            Height          =   315
            Left            =   12570
            TabIndex        =   35
            Top             =   1650
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   255
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "������ǰ��������"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSP1 
            Height          =   315
            Left            =   10380
            TabIndex        =   36
            Top             =   1650
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   16711680
            BackColor       =   16761087
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "һ���ඩ��"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSCommand CMD_CARD 
            Height          =   315
            Left            =   6510
            TabIndex        =   37
            Top             =   1650
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   255
            Caption         =   "ȡ��������"
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   315
            Left            =   8190
            TabIndex        =   40
            Top             =   1650
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "�ȴ���ָʾ"
            FloodColor      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSP4 
            Height          =   315
            Left            =   5100
            TabIndex        =   41
            Top             =   1650
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   33023
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "�ص㶩��"
            FloodColor      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   315
            Left            =   11130
            TabIndex        =   42
            Top             =   1280
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   0
            BackColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "�����ʶ"
            FloodColor      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.Label Label8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   9630
            TabIndex        =   38
            Top             =   1380
            Width           =   195
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   13570
            TabIndex        =   31
            Top             =   990
            Width           =   195
         End
         Begin VB.Label Label6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   13570
            TabIndex        =   30
            Top             =   600
            Width           =   195
         End
         Begin VB.Label Label5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   13570
            TabIndex        =   29
            Top             =   210
            Width           =   195
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   9630
            TabIndex        =   28
            Top             =   990
            Width           =   195
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   9630
            TabIndex        =   27
            Top             =   600
            Width           =   195
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   9630
            TabIndex        =   26
            Top             =   210
            Width           =   195
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   3060
            TabIndex        =   25
            Top             =   210
            Width           =   195
         End
      End
   End
   Begin VB.ComboBox cbo_plate_next 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "CGT2020C.frx":1E81
      Left            =   4470
      List            =   "CGT2020C.frx":1E83
      TabIndex        =   0
      Top             =   5970
      Visible         =   0   'False
      Width           =   735
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   3330
      Top             =   5970
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "���¼���"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "CGT2020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   Mill System
'-- Program Name      ����ʵ����ѯ����
'-- Program ID        CGT2020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2008.5.8
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
Const SS1_TRNS_CMPY_CD = 4
Const SS1_MILL_LOT_NO = 6
Const SS1_APLY_STDSPEC = 16
Const SS1_PROD_THK = 17
Const SS1_ORD_THK = 22
Const SS1_ORD_WID = 23
Const SS1_ORD_LEN = 24
Const SS1_PROD_CAL_UNIT_WGT = 25
Const SS1_TRIM_FL = 26
Const SS1_SIZE_KND = 27
Const SS1_HTM_METH1 = 30  '�Ƿ��ȴ��������ɫ 2013-05-28  by  LiQian
Const SS1_PROD_DATE = 33
Const SS1_SHIFT = 37
Const SS1_GROUP_CD = 38
Const SS1_LEN_TOL = 48
Const SS1_DEL_TO_DATE = 51
Const SS1_ORD_CNT = 53
Const SS1_URGNT_FL = 54   '����������ɫ��� 2012-08-16  by  LiQian
Const SS1_IMP_CONT = 57
Const SS1_PILECOOL = 58

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_SLAB_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(TXT_MILL_LOT_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'ADDED BY GUOLI AT 20080604
       Call Gp_Ms_Collection(TXT_STDSPEC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_THK, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDB_THK_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_WID, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDB_WID_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_LEN, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDB_LEN_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(CBO_GAS_FL, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'ADDED BY GUOLI AT 20080918
'    Call Gp_Ms_Collection(cbo_plate_next, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_THK_PLAN, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_THK_PLAN_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_WID_PLAN, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_WID_PLAN_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_LEN_PLAN, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_LEN_PLAN_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(cbo_plate_yn, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(cbo_scrap, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 
 'ADDED BY ��ѧ�� for add �ֶκŲ�ѯ���� at 20100903
   Call Gp_Ms_Collection(TXT_TRNS_CMPY_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 'Added by lichao for add �ֶκŲ�ѯ���� at 20120611
Call Gp_Ms_Collection(TXT_TRNS_CMPY_CD_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  'ADDED BY zhangjinbo for add ԭʼ���ϸ��ֲ�ѯ���� at 20110121
        Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
        Mc1.Add Item:=pControl, Key:="pControl"
        Mc1.Add Item:=nControl, Key:="nControl"
        Mc1.Add Item:=mControl, Key:="mControl"
        Mc1.Add Item:=iControl, Key:="iControl"
        Mc1.Add Item:=rControl, Key:="rControl"
        Mc1.Add Item:=cControl, Key:="cControl"
        Mc1.Add Item:=aControl, Key:="aControl"
        Mc1.Add Item:=lControl, Key:="lControl"

     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 53, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)    '����������ɫ��� 2012-08-16  by  LiQian
    Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)    '�Ƿ񷭰�
    Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)    '�Ƿ�RH 2013-05-28  by  LiQian
    Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 58, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 59, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 60, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 61, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 62, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 63, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 64, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 65, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 66, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 67, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 68, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGT2020C.P_SREFER", Key:="P-R"
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


Private Sub CMD_CARD_Click()
    Call ExcelPrn_Pile(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
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

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
        
    Screen.MousePointer = vbDefault
    
    CBO_GAS_FL.AddItem "Y"
    CBO_GAS_FL.AddItem "N"
    
    cbo_plate_next.AddItem "Y"
    cbo_plate_next.AddItem "N"
    
    cbo_plate_yn.AddItem "Y"
    cbo_plate_yn.AddItem "N"
    
    cbo_scrap.AddItem "Y"
    cbo_scrap.AddItem "N"
    
    CBO_SHIFT.AddItem " "
    CBO_SHIFT.AddItem "1"
    CBO_SHIFT.AddItem "2"
    CBO_SHIFT.AddItem "3"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
'    Set pControl1 = Nothing
'    Set nControl1 = Nothing
'    Set iControl1 = Nothing
'    Set rControl1 = Nothing
'    Set cControl1 = Nothing
'    Set aControl1 = Nothing
'    Set lControl1 = Nothing
'    Set mControl1 = Nothing

    Set Mc1 = Nothing
'    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub
Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub
Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
       Call Gp_Ms_Cls(Mc1("rControl"))
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If
    
    TXT_TRNS_CMPY_CD.Text = ""
    TXT_TRNS_CMPY_CD_TO.Text = ""

End Sub

Public Sub Form_Ref()
    
    Dim iRow As Integer
    Dim iCol As Integer
    Dim sCurDate As String
    Dim sDel_To_Date As String
    Dim sord_cnt As Integer
    Dim sBackColor As Variant
    Dim sHtm_Meth As String
    Dim simpcont As String
    Dim PILECOOL As String
    
    sCurDate = Format(Now, "YYYYMM")
    
    If TXT_TRNS_CMPY_CD.Text <> "" Or TXT_TRNS_CMPY_CD_TO.Text <> "" Then
            If SDT_PROD_DATE_FROM.Text = SDT_PROD_DATE_TO.Text Then
                Else: Call Gp_MsgBoxDisplay("�и�ʱ�����Ϊͬһ��", "W")
                Exit Sub
            End If
    End If
    

    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        With ss1
              
              For iRow = 1 To .MaxRows
                .ROW = iRow:            .Col = SS1_ORD_CNT:        sord_cnt = Val(.Text)
                
                If sord_cnt > 1 Then
                     Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, , SSP1.BackColor)
                End If
                .ROW = iRow:            .Col = SS1_DEL_TO_DATE
                sDel_To_Date = Mid(.Value, 1, 6)
                If sDel_To_Date < sCurDate Then
                     If sord_cnt > 1 Then
                        sBackColor = SSP1.BackColor
                    Else
                        sBackColor = vbWhite
                    End If
                     Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, .ROW, .ROW, &HFF&, sBackColor)
                End If
                
                .ROW = iRow:             .Col = SS1_HTM_METH1:            sHtm_Meth = .Text
                If sHtm_Meth <> "" Then
                     Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, .ROW, .ROW, , SSPanel1.BackColor)
                End If
                
                 '����������ɫ��� 2012-08-16  by  LiQian
                 .Col = SS1_URGNT_FL
                If ss1.Text = "Y" Then
                     Call Gp_Sp_BlockColor(ss1, SS1_PLATE_NO, .MaxRows, .ROW, .ROW, &HC000&)
                End If
                
                .ROW = iRow:
                .Col = SS1_IMP_CONT:    simpcont = Trim(.Text)
                If simpcont = "Y" Then
                  Call Gp_Sp_BlockColor(ss1, SS1_PLATE_NO, SS1_PLATE_NO, iRow, iRow, SSP4.BackColor)
                  Call Gp_Sp_BlockColor(ss1, SS1_IMP_CONT, SS1_IMP_CONT, iRow, iRow, SSP4.BackColor)
                End If
                
                .ROW = iRow:
                .Col = SS1_PILECOOL: PILECOOL = Trim(.Text)
                .Col = SS1_IMP_CONT: simpcont = Trim(.Text)
                
                If PILECOOL = "Y" And simpcont <> "Y" Then
                    Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, vbBlack, vbYellow)
                End If
              
              
              
              Next iRow
        
        End With
        
    End If
    
        
'    Call ss1.SetActiveCell(1, ss1.MaxRows)
               
End Sub

Private Sub SDT_PROD_DATE_FROM_GotFocus()
     SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
     SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
End Sub


Private Sub SDT_PROD_DATE_TO_GotFocus()
     SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    If ROW = 0 Then
        Call Gp_Sp_Sort(ss1, Col, ROW)
    End If
End Sub

Private Sub txt_stdspec_DblClick()
    Call txt_STDSPEC_KeyUp(vbKeyF4, 1)
End Sub

Private Sub txt_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_STDSPEC

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)
    End If
End Sub


Private Sub txt_stlgrd_Change()
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
        
    End If
        
End Sub


Public Sub ExcelPrn_Pile(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

On Error GoTo Excel_Error

    Dim ret         As Boolean
    Dim xlApp       As Object
    Dim xlBpp       As Object
    Dim xlBook      As Object
    Dim xlSheet     As Object
    Dim ColIndex    As Integer
    Dim sExlRange   As String
    Dim sExlRange1  As String
    Dim iExlCol     As Integer
    
    
    With sPname
    
        If .MaxRows = 0 Then Exit Sub
        
        If bLkcol1 = 0 Then
           bLkcol1 = 1
        End If
        
        If bLkcol2 = 0 Then
            bLkcol2 = -1
        End If
        
        If bLkrow2 = 0 Then
            bLkrow2 = -1
        End If
        
        Clipboard.Clear
        
        .Col = bLkcol1: .Col2 = bLkcol2
        .ROW = bLkrow1: .Row2 = bLkrow2
        Clipboard.SetText .Clip
        
        'Call Excel
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
    
        xlApp.Visible = True
                        
        xlSheet.Cells.NumberFormatLocal = "G/ͨ�ø�ʽ"
        
        sExlRange1 = ""
        For ColIndex = 1 To .MaxCols
            .Col = ColIndex
            .ROW = 1

            iExlCol = ColIndex
'            If IsNumeric(.Text) And (Left(.Text, 1) = "0" Or Left(.Text, 1) = "1" Or Left(.Text, 1) = "7") And _
'               (Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14) Then
            If .CellType = SS_CELL_TYPE_EDIT Then
                If ColIndex > 104 Then
                    sExlRange1 = "D"
                    iExlCol = ColIndex - 104
                ElseIf ColIndex > 78 Then
                    sExlRange1 = "C"
                    iExlCol = ColIndex - 78
                ElseIf ColIndex > 52 Then
                    sExlRange1 = "B"
                    iExlCol = ColIndex - 52
                ElseIf ColIndex > 26 Then
                    sExlRange1 = "A"
                    iExlCol = ColIndex - 26
                End If

                sExlRange = sExlRange1 & Chr(iExlCol + 64) & "1:" & sExlRange1 & Chr(iExlCol + 64) & .MaxRows + 5
                If Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14 Then
                     xlSheet.Range(sExlRange).NumberFormat = "@"
                End If
            End If
        Next
        
       ss1.ROW = 0
       ss1.Col = SS1_PLATE_NO:             xlApp.Range("A1").Value = ss1.Text
       ss1.Col = SS1_TRNS_CMPY_CD:         xlApp.Range("B1").Value = ss1.Text
       ss1.Col = SS1_MILL_LOT_NO:          xlApp.Range("C1").Value = ss1.Text
       ss1.Col = SS1_APLY_STDSPEC:         xlApp.Range("D1").Value = ss1.Text
       ss1.Col = SS1_PROD_THK:             xlApp.Range("E1").Value = ss1.Text
       ss1.Col = SS1_ORD_THK:              xlApp.Range("F1").Value = ss1.Text
       ss1.Col = SS1_ORD_WID:              xlApp.Range("G1").Value = ss1.Text
       ss1.Col = SS1_ORD_LEN:              xlApp.Range("H1").Value = ss1.Text
       ss1.Col = SS1_PROD_CAL_UNIT_WGT:    xlApp.Range("I1").Value = ss1.Text
       ss1.Col = SS1_TRIM_FL:              xlApp.Range("J1").Value = ss1.Text
       ss1.Col = SS1_SIZE_KND:             xlApp.Range("K1").Value = ss1.Text
       ss1.Col = SS1_PROD_DATE:            xlApp.Range("L1").Value = ss1.Text
       ss1.Col = SS1_SHIFT:                xlApp.Range("M1").Value = ss1.Text
       ss1.Col = SS1_GROUP_CD:             xlApp.Range("N1").Value = ss1.Text
       ss1.Col = SS1_LEN_TOL:              xlApp.Range("O1").Value = ss1.Text
       
        xlApp.Range("E2").Value = "��ȹ���"
        xlApp.Range("F2").Value = "���"
        xlApp.Range("G2").Value = "����"
        xlApp.Range("H2").Value = "����"
        xlApp.Range("I2").Value = "����"
        xlApp.Range("J2").Value = "�б�"
        xlApp.Range("K2").Value = "����"
       
        Clipboard.Clear
        ss1.SetSelection SS1_PLATE_NO, 1, SS1_PLATE_NO, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("A3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_TRNS_CMPY_CD, 1, SS1_TRNS_CMPY_CD, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("B3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_MILL_LOT_NO, 1, SS1_MILL_LOT_NO, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("C3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_APLY_STDSPEC, 1, SS1_APLY_STDSPEC, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("D3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_PROD_THK, 1, SS1_SIZE_KND, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("E3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_PROD_THK, 1, SS1_SIZE_KND, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("E3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_PROD_DATE, 1, SS1_PROD_DATE, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("L3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_SHIFT, 1, SS1_GROUP_CD, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("M3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_LEN_TOL, 1, SS1_LEN_TOL, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("O3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        
    
'        xlSheet.Range("A1").Select
'        xlSheet.Paste
'        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
            
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
        
    End With
    
    Exit Sub
    
Excel_Error:
    Call Gp_MsgBoxDisplay("���Ļ�����δ��װExcel", "W")

End Sub


