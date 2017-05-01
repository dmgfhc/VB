VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGC2010C 
   Caption         =   "精轧作业实绩查询及修改界面_CGC2010C"
   ClientHeight    =   9240
   ClientLeft      =   585
   ClientTop       =   1365
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   15075
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   525
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   926
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox txt_RmFinTmp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   14550
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   " "
         Top             =   105
         Width           =   675
      End
      Begin VB.TextBox txt_RMRollingSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   " "
         Top             =   105
         Width           =   2000
      End
      Begin VB.TextBox txt_RollingSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10890
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   " "
         Top             =   105
         Width           =   2235
      End
      Begin VB.TextBox txt_SlabNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Top             =   105
         Width           =   1635
      End
      Begin VB.TextBox txt_SlabSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4155
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   " "
         Top             =   105
         Width           =   2085
      End
      Begin VB.TextBox TXT_CB 
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
         Left            =   14595
         TabIndex        =   1
         Text            =   "CC"
         Top             =   90
         Visible         =   0   'False
         Width           =   345
      End
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   120
         Top             =   105
         Width           =   1185
         _ExtentX        =   2090
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
      Begin InDate.ULabel ULabel43 
         Height          =   315
         Left            =   3060
         Top             =   105
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "板坯规格"
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
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   9780
         Top             =   105
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "轧制规格"
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
      Begin InDate.ULabel ULabel24 
         Height          =   315
         Left            =   6360
         Top             =   105
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         Caption         =   "粗轧后规格"
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
      Begin InDate.ULabel ULabel26 
         Height          =   315
         Left            =   13260
         Top             =   105
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "粗轧结束温度"
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
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   8655
      Left            =   30
      TabIndex        =   2
      Top             =   510
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   15266
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox txt_HTM 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10830
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   68
         Top             =   900
         Width           =   1095
      End
      Begin VB.TextBox txt_CrCd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11220
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   66
         Top             =   1270
         Width           =   1320
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "精轧实绩"
         ForeColor       =   &H00FF0000&
         Height          =   1035
         Left            =   120
         TabIndex        =   27
         Top             =   3090
         Width           =   15075
         Begin VB.TextBox TXT_EMP3 
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
            Left            =   13680
            MaxLength       =   8
            TabIndex        =   41
            Top             =   570
            Width           =   1215
         End
         Begin VB.TextBox TXT_EMP2 
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
            Left            =   12450
            MaxLength       =   8
            TabIndex        =   37
            Top             =   570
            Width           =   1215
         End
         Begin VB.TextBox TXT_EMP1 
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
            Left            =   11220
            MaxLength       =   8
            TabIndex        =   36
            Top             =   570
            Width           =   1215
         End
         Begin VB.TextBox TXT_GROUP 
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
            Left            =   10515
            MaxLength       =   1
            TabIndex        =   35
            Top             =   570
            Width           =   705
         End
         Begin VB.TextBox TXT_SHIFT 
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
            Left            =   9810
            MaxLength       =   1
            TabIndex        =   34
            Top             =   570
            Width           =   705
         End
         Begin CSTextLibCtl.sidbEdit txt_thk 
            Height          =   315
            Left            =   5010
            TabIndex        =   28
            Tag             =   "厚度"
            Top             =   230
            Width           =   675
            _Version        =   262145
            _ExtentX        =   1191
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
         Begin CSTextLibCtl.sidbEdit txt_wid 
            Height          =   315
            Left            =   6735
            TabIndex        =   29
            Tag             =   "宽度"
            Top             =   230
            Width           =   735
            _Version        =   262145
            _ExtentX        =   1296
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
            RawData         =   ""
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
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   90
            Top             =   230
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   556
            Caption         =   "开轧时间"
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
            Top             =   630
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   556
            Caption         =   "终轧时间"
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
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   4080
            Top             =   230
            Width           =   900
            _ExtentX        =   1588
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
         Begin InDate.ULabel ULabel23 
            Height          =   315
            Left            =   5820
            Top             =   230
            Width           =   900
            _ExtentX        =   1588
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
         Begin InDate.ULabel ULabel27 
            Height          =   315
            Left            =   7620
            Top             =   230
            Width           =   900
            _ExtentX        =   1588
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
         Begin CSTextLibCtl.sitxEdit TXT_MILL_STA_TIME 
            Height          =   315
            Left            =   1770
            TabIndex        =   30
            Tag             =   "开轧时间"
            Top             =   230
            Width           =   2145
            _Version        =   262145
            _ExtentX        =   3784
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   "____-__-__ __-__-__"
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   "____-__-__ __:__:__ "
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
            Mask            =   "____-__-__ __:__:__ "
            CharacterTable  =   ""
            BorderStyle     =   0
            MaxLength       =   0
            ValidateMask    =   0   'False
         End
         Begin CSTextLibCtl.sitxEdit TXT_MILL_END_TIME 
            Height          =   315
            Left            =   1770
            TabIndex        =   31
            Tag             =   "终轧时间"
            Top             =   630
            Width           =   2145
            _Version        =   262145
            _ExtentX        =   3784
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   "____-__-__ __-__-__"
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
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   "____-__-__ __:__:__"
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
            Mask            =   "____-__-__ __:__:__"
            CharacterTable  =   ""
            BorderStyle     =   0
            MaxLength       =   14
            ValidateMask    =   0   'False
         End
         Begin CSTextLibCtl.sidbEdit txt_len 
            Height          =   315
            Left            =   8550
            TabIndex        =   32
            Tag             =   "长度"
            Top             =   230
            Width           =   1095
            _Version        =   262145
            _ExtentX        =   1931
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
            RawData         =   ""
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
            NumDecDigits    =   0
            NumIntDigits    =   5
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit txt_LastTemp 
            Height          =   315
            Left            =   5670
            TabIndex        =   33
            Top             =   600
            Width           =   855
            _Version        =   262145
            _ExtentX        =   1508
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
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
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel20 
            Height          =   315
            Left            =   4080
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            Caption         =   "轧制结束温度"
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
         Begin InDate.ULabel ULabel34 
            Height          =   315
            Left            =   9810
            Top             =   230
            Width           =   705
            _ExtentX        =   1244
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
         Begin InDate.ULabel ULabel35 
            Height          =   315
            Left            =   10515
            Top             =   230
            Width           =   705
            _ExtentX        =   1244
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
         Begin InDate.ULabel ULabel36 
            Height          =   315
            Left            =   11235
            Top             =   230
            Width           =   3665
            _ExtentX        =   6456
            _ExtentY        =   556
            Caption         =   "作业人员"
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
         Begin Threed.SSCommand cmd_LPass 
            Height          =   360
            Left            =   8220
            TabIndex        =   51
            Top             =   600
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   635
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   128
            BackColor       =   14737632
            Caption         =   "空过"
         End
         Begin Threed.SSCommand cmd_Pass 
            Height          =   360
            Left            =   6720
            TabIndex        =   49
            Top             =   600
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   635
            _Version        =   196609
            ForeColor       =   16711680
            BackColor       =   14737632
            BackStyle       =   1
            ActiveColors    =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "废钢"
         End
      End
      Begin VB.TextBox txt_Roll_Stlgrd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10830
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   120
         Width           =   2190
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "冷却实绩"
         ForeColor       =   &H00FF0000&
         Height          =   1125
         Left            =   120
         TabIndex        =   43
         Top             =   2040
         Width           =   6015
         Begin CSTextLibCtl.sidbEdit SDB_COOL_AVE_TEMP 
            Height          =   315
            Left            =   1465
            TabIndex        =   44
            Tag             =   "冷却平均温度"
            Top             =   255
            Width           =   825
            _Version        =   262145
            _ExtentX        =   1455
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
            RawData         =   ""
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
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_COOL_EXT_TEMP 
            Height          =   315
            Left            =   5025
            TabIndex        =   45
            Top             =   255
            Width           =   825
            _Version        =   262145
            _ExtentX        =   1455
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
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
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_COOL_ENT_TEMP 
            Height          =   315
            Left            =   4170
            TabIndex        =   46
            Top             =   255
            Width           =   825
            _Version        =   262145
            _ExtentX        =   1455
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
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
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_COOL_WGT 
            Height          =   315
            Left            =   4170
            TabIndex        =   47
            Tag             =   "冷却水量"
            Top             =   675
            Width           =   825
            _Version        =   262145
            _ExtentX        =   1455
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
            RawData         =   ""
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
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   90
            Top             =   255
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "冷却平均温度"
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
            Left            =   2460
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            Caption         =   "冷却入\出口温度"
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
            Left            =   2460
            Top             =   675
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            Caption         =   "冷却水量"
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
         Begin CSTextLibCtl.sidbEdit sidbEdit2 
            Height          =   315
            Left            =   1465
            TabIndex        =   48
            Tag             =   "冷却水量"
            Top             =   675
            Width           =   825
            _Version        =   262145
            _ExtentX        =   1455
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
            RawData         =   ""
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
            NumDecDigits    =   0
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel25 
            Height          =   315
            Left            =   90
            Top             =   675
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "冷却速率"
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
      End
      Begin VB.TextBox txt_CoolMth 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12630
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   40
         Top             =   1650
         Width           =   825
      End
      Begin VB.TextBox txt_TrimFl 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   14310
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   39
         Top             =   135
         Width           =   810
      End
      Begin VB.TextBox txt_Stlgrd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10830
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   38
         Top             =   510
         Width           =   1710
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "控轧实绩"
         ForeColor       =   &H00FF0000&
         Height          =   1125
         Left            =   6090
         TabIndex        =   16
         Top             =   2040
         Width           =   9105
         Begin VB.TextBox TXT_EXCEPTION 
            BackColor       =   &H8000000B&
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
            Left            =   5310
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   75
            Tag             =   "控轧代码"
            Text            =   " "
            Top             =   150
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.CheckBox CHK_C 
            BackColor       =   &H00E0E0E0&
            Caption         =   "工艺超下限值"
            Height          =   285
            Left            =   3780
            TabIndex        =   74
            Top             =   720
            Width           =   1605
         End
         Begin VB.CheckBox CHK_A 
            BackColor       =   &H00E0E0E0&
            Caption         =   "工艺正常"
            Height          =   285
            Left            =   3780
            TabIndex        =   73
            Top             =   180
            Width           =   1455
         End
         Begin VB.CheckBox CHK_B 
            BackColor       =   &H00E0E0E0&
            Caption         =   "工艺超上限值"
            Height          =   285
            Left            =   3780
            TabIndex        =   72
            Top             =   450
            Width           =   1605
         End
         Begin VB.CheckBox CHK_ROLLING_OP 
            BackColor       =   &H00E0E0E0&
            Caption         =   "人工干预"
            Height          =   285
            Left            =   2550
            TabIndex        =   22
            Top             =   675
            Width           =   1095
         End
         Begin VB.CheckBox CHK_ROLLING_AUTO 
            BackColor       =   &H00E0E0E0&
            Caption         =   "自动"
            Height          =   285
            Left            =   1770
            TabIndex        =   21
            Top             =   675
            Width           =   735
         End
         Begin VB.TextBox TXT_ROLLING_METHOD 
            BackColor       =   &H8000000B&
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
            Left            =   1125
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   20
            Tag             =   "轧制方式"
            Text            =   " "
            Top             =   675
            Width           =   585
         End
         Begin VB.TextBox TXT_CR_CD 
            BackColor       =   &H8000000B&
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
            Left            =   1125
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   19
            Tag             =   "控轧代码"
            Text            =   " "
            Top             =   255
            Width           =   585
         End
         Begin VB.CheckBox CHK_CR_CD 
            BackColor       =   &H00E0E0E0&
            Caption         =   "控轧"
            Height          =   285
            Left            =   1770
            TabIndex        =   18
            Top             =   270
            Width           =   735
         End
         Begin VB.CheckBox CHK_NON_CR_CD 
            BackColor       =   &H00E0E0E0&
            Caption         =   "否"
            Height          =   285
            Left            =   2550
            TabIndex        =   17
            Top             =   270
            Width           =   1095
         End
         Begin InDate.ULabel ULabel80 
            Height          =   315
            Left            =   120
            Top             =   675
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            Caption         =   "轧制方式"
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
         Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE1_THK 
            Height          =   315
            Left            =   6690
            TabIndex        =   23
            Tag             =   "一阶段厚度"
            Top             =   255
            Width           =   735
            _Version        =   262145
            _ExtentX        =   1296
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
            RawData         =   ""
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
            NumDecDigits    =   0
            NumIntDigits    =   3
            ShowZero        =   0   'False
            MaxValue        =   999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE2_THK 
            Height          =   315
            Left            =   6690
            TabIndex        =   24
            Tag             =   "二阶段厚度"
            Top             =   675
            Width           =   735
            _Version        =   262145
            _ExtentX        =   1296
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
            RawData         =   ""
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
            NumDecDigits    =   0
            NumIntDigits    =   3
            ShowZero        =   0   'False
            MaxValue        =   999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE1_TEMP 
            Height          =   315
            Left            =   8355
            TabIndex        =   25
            Tag             =   "一阶段温度"
            Top             =   255
            Width           =   675
            _Version        =   262145
            _ExtentX        =   1191
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
            ShowZero        =   0   'False
            MaxValue        =   9999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE2_TEMP 
            Height          =   315
            Left            =   8355
            TabIndex        =   26
            Tag             =   "二阶段温度"
            Top             =   675
            Width           =   675
            _Version        =   262145
            _ExtentX        =   1191
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
            ShowZero        =   0   'False
            MaxValue        =   9999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel94 
            Height          =   315
            Left            =   7530
            Top             =   255
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            Caption         =   "温度"
            Alignment       =   1
            BackColor       =   14804174
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
         Begin InDate.ULabel ULabel95 
            Height          =   315
            Left            =   7530
            Top             =   675
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            Caption         =   "温度"
            Alignment       =   1
            BackColor       =   14804174
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
         Begin InDate.ULabel ULabel105 
            Height          =   315
            Left            =   120
            Top             =   255
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            Caption         =   "控轧代码"
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
         Begin InDate.ULabel ULabel91 
            Height          =   315
            Left            =   5520
            Top             =   255
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "一阶段 厚度"
            Alignment       =   1
            BackColor       =   14804174
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
         Begin InDate.ULabel ULabel92 
            Height          =   315
            Left            =   5520
            Top             =   675
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "二阶段 厚度"
            Alignment       =   1
            BackColor       =   14804174
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
      End
      Begin InDate.ULabel ULabel32 
         Height          =   315
         Left            =   9660
         Top             =   1270
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         Caption         =   "控轧代码"
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
      Begin FPSpread.vaSpread ss1 
         Height          =   1840
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   9420
         _Version        =   393216
         _ExtentX        =   16616
         _ExtentY        =   3246
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
         MaxCols         =   6
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGC2010C.frx":0000
      End
      Begin CSTextLibCtl.sidbEdit txt_CrMillRatet3 
         Height          =   315
         Left            =   11220
         TabIndex        =   7
         Top             =   1650
         Width           =   645
         _Version        =   262145
         _ExtentX        =   1138
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
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
         NumDecDigits    =   0
         NumIntDigits    =   1
         ShowZero        =   0   'False
         MaxValue        =   9.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   12630
         Top             =   1270
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         Caption         =   "冷却方式 速率 温度"
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
      Begin CSTextLibCtl.sidbEdit txt_CoolSpeed 
         Height          =   315
         Left            =   13455
         TabIndex        =   8
         Top             =   1650
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
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
         NumDecDigits    =   0
         NumIntDigits    =   1
         ShowZero        =   0   'False
         MaxValue        =   9.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_CoolTemp 
         Height          =   315
         Left            =   14295
         TabIndex        =   9
         Top             =   1650
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         MaxValue        =   9.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   13140
         Top             =   135
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "切边代码"
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
      Begin TabDlg.SSTab tab1 
         Height          =   4335
         Left            =   120
         TabIndex        =   10
         Top             =   4170
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   7646
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   4
         TabHeight       =   520
         TabMaxWidth     =   3528
         BackColor       =   14737632
         TabCaption(0)   =   "精轧等待"
         TabPicture(0)   =   "CGC2010C.frx":19BE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSP2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "SSP4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "SSP1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "SSSplitter1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "精轧实绩"
         TabPicture(1)   =   "CGC2010C.frx":19DA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txt_RstToDate"
         Tab(1).Control(1)=   "txt_RstFormDate"
         Tab(1).Control(2)=   "ss3"
         Tab(1).ControlCount=   3
         Begin SSSplitter.SSSplitter SSSplitter1 
            Height          =   3855
            Left            =   60
            TabIndex        =   61
            Top             =   390
            Width           =   14940
            _ExtentX        =   26353
            _ExtentY        =   6800
            _Version        =   196609
            SplitterBarWidth=   3
            BorderStyle     =   1
            PaneTree        =   "CGC2010C.frx":19F6
            Begin FPSpread.vaSpread ss2 
               Height          =   3825
               Left            =   15
               TabIndex        =   62
               Top             =   15
               Width           =   11490
               _Version        =   393216
               _ExtentX        =   20267
               _ExtentY        =   6747
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   37
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               SpreadDesigner  =   "CGC2010C.frx":1A48
            End
            Begin FPSpread.vaSpread ss4 
               Height          =   3825
               Left            =   11565
               TabIndex        =   63
               Top             =   15
               Width           =   3360
               _Version        =   393216
               _ExtentX        =   5927
               _ExtentY        =   6747
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   4
               MaxRows         =   9
               RetainSelBlock  =   0   'False
               SpreadDesigner  =   "CGC2010C.frx":2846
            End
         End
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3855
            Left            =   -74910
            TabIndex        =   11
            Top             =   390
            Width           =   14700
            _Version        =   393216
            _ExtentX        =   25929
            _ExtentY        =   6800
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
            MaxCols         =   10
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CGC2010C.frx":2C5E
         End
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   3855
            Left            =   -74910
            TabIndex        =   12
            Top             =   390
            Width           =   14700
            _Version        =   393216
            _ExtentX        =   25929
            _ExtentY        =   6800
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
            MaxCols         =   10
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CGC2010C.frx":474D
         End
         Begin FPSpread.vaSpread ss3 
            Height          =   3495
            Left            =   -74910
            TabIndex        =   13
            Top             =   750
            Width           =   14940
            _Version        =   393216
            _ExtentX        =   26352
            _ExtentY        =   6165
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   13
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CGC2010C.frx":623C
         End
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   -74910
            Top             =   420
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            Caption         =   "装时间"
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
         Begin CSTextLibCtl.sitxEdit txt_RstFormDate 
            Height          =   315
            Left            =   -74910
            TabIndex        =   14
            Tag             =   "装炉时间"
            Top             =   420
            Width           =   1830
            _Version        =   262145
            _ExtentX        =   3228
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
            Text            =   "____-__-__ __-__-__"
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
            Mask            =   "____-__-__ __:__"
            CharacterTable  =   ""
            BorderStyle     =   0
            MaxLength       =   0
            ValidateMask    =   0   'False
         End
         Begin CSTextLibCtl.sitxEdit txt_RstToDate 
            Height          =   315
            Left            =   -73080
            TabIndex        =   15
            Tag             =   "装炉时间"
            Top             =   420
            Width           =   1800
            _Version        =   262145
            _ExtentX        =   3175
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
            Text            =   "____-__-__ __-__-__"
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
            Mask            =   "____-__-__ __:__"
            CharacterTable  =   ""
            BorderStyle     =   0
            MaxLength       =   0
            ValidateMask    =   0   'False
         End
         Begin Threed.SSPanel SSP1 
            Height          =   285
            Left            =   6000
            TabIndex        =   69
            Top             =   0
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   196609
            ForeColor       =   0
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
         Begin Threed.SSPanel SSP4 
            Height          =   285
            Left            =   4590
            TabIndex        =   70
            Top             =   0
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   196609
            ForeColor       =   0
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
         Begin Threed.SSPanel SSP2 
            Height          =   285
            Left            =   7380
            TabIndex        =   71
            Top             =   0
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   196609
            ForeColor       =   16711680
            BackColor       =   8454143
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "一坯多订单"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   9660
         Top             =   120
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "轧制标准"
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
      Begin InDate.ULabel ULabel28 
         Height          =   315
         Left            =   9660
         Top             =   510
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "板坯钢种"
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
      Begin CSTextLibCtl.sidbEdit txt_millTemp 
         Height          =   315
         Left            =   12630
         TabIndex        =   53
         Top             =   900
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
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
         NumDecDigits    =   0
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel29 
         Height          =   315
         Left            =   12630
         Top             =   510
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         Caption         =   "终轧温度 最小 最大"
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
      Begin CSTextLibCtl.sidbEdit txt_millTemp_min 
         Height          =   315
         Left            =   13455
         TabIndex        =   54
         Top             =   900
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
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
         NumDecDigits    =   0
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_millTemp_max 
         Height          =   315
         Left            =   14295
         TabIndex        =   55
         Top             =   900
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
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
         ShowZero        =   0   'False
         MaxValue        =   9.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   9660
         Top             =   1650
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         Caption         =   "T1厚度比 温度"
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
         Height          =   315
         Left            =   9660
         Top             =   900
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "热处理"
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
      Begin CSTextLibCtl.sidbEdit txt_CrMillTmpt3 
         Height          =   315
         Left            =   11895
         TabIndex        =   67
         Top             =   1650
         Width           =   645
         _Version        =   262145
         _ExtentX        =   1138
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
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
         NumDecDigits    =   0
         NumIntDigits    =   1
         ShowZero        =   0   'False
         MaxValue        =   9.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   360
      Left            =   7680
      TabIndex        =   56
      Top             =   9660
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   635
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   14737632
      BackStyle       =   1
      ActiveColors    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "再板坯"
   End
   Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE3_TEMP 
      Height          =   315
      Left            =   6435
      TabIndex        =   57
      Tag             =   "三阶段温度"
      Top             =   9750
      Visible         =   0   'False
      Width           =   675
      _Version        =   262145
      _ExtentX        =   1191
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
      ShowZero        =   0   'False
      MaxValue        =   9999
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel93 
      Height          =   315
      Left            =   3945
      Top             =   9750
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      Caption         =   "厚度"
      Alignment       =   1
      BackColor       =   14804174
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
   Begin InDate.ULabel ULabel96 
      Height          =   315
      Left            =   5610
      Top             =   9750
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      Caption         =   "温度"
      Alignment       =   1
      BackColor       =   14804174
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
   Begin InDate.ULabel ULabel102 
      Height          =   315
      Left            =   2820
      Top             =   9750
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Caption         =   "三阶段"
      Alignment       =   1
      BackColor       =   14804174
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
   Begin CSTextLibCtl.sidbEdit SDB_CR_STAGE3_THK 
      Height          =   315
      Left            =   4770
      TabIndex        =   58
      Tag             =   "三阶段厚度"
      Top             =   9750
      Visible         =   0   'False
      Width           =   735
      _Version        =   262145
      _ExtentX        =   1296
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
      ShowZero        =   0   'False
      MaxValue        =   999
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_CrMillRatet5 
      Height          =   315
      Left            =   16725
      TabIndex        =   59
      Top             =   3120
      Visible         =   0   'False
      Width           =   690
      _Version        =   262145
      _ExtentX        =   1217
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   14737632
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
      NumIntDigits    =   1
      ShowZero        =   0   'False
      MaxValue        =   9.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   15750
      Top             =   3150
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   "T3厚度比"
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
   Begin CSTextLibCtl.sidbEdit txt_CrMillTmpt5 
      Height          =   315
      Left            =   16725
      TabIndex        =   60
      Top             =   3540
      Visible         =   0   'False
      Width           =   690
      _Version        =   262145
      _ExtentX        =   1217
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   14737632
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
      NumDecDigits    =   2
      NumIntDigits    =   1
      ShowZero        =   0   'False
      MaxValue        =   9.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   15750
      Top             =   3540
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   "T3温度"
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
   Begin CSTextLibCtl.sidbEdit txt_CrMillRatet4 
      Height          =   315
      Left            =   17280
      TabIndex        =   64
      Top             =   2760
      Visible         =   0   'False
      Width           =   645
      _Version        =   262145
      _ExtentX        =   1138
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   14737632
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
      NumIntDigits    =   1
      ShowZero        =   0   'False
      MaxValue        =   9.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_CrMillTmpt4 
      Height          =   315
      Left            =   17940
      TabIndex        =   65
      Top             =   2760
      Visible         =   0   'False
      Width           =   645
      _Version        =   262145
      _ExtentX        =   1138
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   14737632
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
      NumDecDigits    =   2
      NumIntDigits    =   1
      ShowZero        =   0   'False
      MaxValue        =   9.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   15720
      Top             =   2760
      Visible         =   0   'False
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "T2厚度比 温度"
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
End
Attribute VB_Name = "CGC2010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   OLD PLATE Mill System
'-- Program Name      (FM)轧制作业实绩查询及修改界面
'-- Program ID        CGC2010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          SHIN.C.S
'-- Coder             SHIN.C.S
'-- Date              2007.7.23
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
Public sQuery_Rt As String          'Active Form sQuery Setting
       
Dim pControl1 As New Collection      'Master Primary Key Collection
Dim nControl1 As New Collection      'Master Necessary Collection
Dim mControl1 As New Collection      'Master Maxlength check Collection
Dim iControl1 As New Collection      'Master Insert Collection
Dim rControl1 As New Collection      'Master Refer Collection
Dim cControl1 As New Collection      'Master Copy Collection
Dim aControl1 As New Collection      'Master -> Spread Collection
Dim lControl1 As New Collection      'Master Lock Collection

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

Dim pControl3 As New Collection      'Master Primary Key Collection
Dim nControl3 As New Collection      'Master Necessary Collection
Dim mControl3 As New Collection      'Master Maxlength check Collection
Dim iControl3 As New Collection      'Master Insert Collection
Dim rControl3 As New Collection      'Master Refer Collection
Dim cControl3 As New Collection      'Master Copy Collection
Dim aControl3 As New Collection      'Master -> Spread Collection
Dim lControl3 As New Collection      'Master Lock Collection

Dim pControl4 As New Collection      'Master Primary Key Collection
Dim nControl4 As New Collection      'Master Necessary Collection
Dim mControl4 As New Collection      'Master Maxlength check Collection
Dim iControl4 As New Collection      'Master Insert Collection
Dim rControl4 As New Collection      'Master Refer Collection
Dim cControl4 As New Collection      'Master Copy Collection
Dim aControl4 As New Collection      'Master -> Spread Collection
Dim lControl4 As New Collection      'Master Lock Collection

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

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim pColumn4 As New Collection      'Spread Primary Key Collection
Dim nColumn4 As New Collection      'Spread necessary Column Collection
Dim mColumn4 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn4 As New Collection      'Spread Insert Column Collection
Dim aColumn4 As New Collection      'Master -> Spread Column Collection
Dim lColumn4 As New Collection      'Spread Lock Column Collection


Dim Mc1 As New Collection           'Master Collectionn
Dim Mc2 As New Collection           'Master Collectionn
Dim Mc3 As New Collection           'Master Collectionn
Dim Mc4 As New Collection           'Master Collectionn

Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection
Dim sc4 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS2_SLAB_NO = 1
Const SS2_UST_FL = 25           '是否探伤 add by liqian 2013-04-08
Const SS2_DEL_TO_DATE = 28      '超交货期用红色显示 add by liqian 2012-06-11
Const SS2_URGNT_FL = 29         '紧急订单绿色标记显示 add by liqian 2012-08-15
Const SS2_FLAG_FL = 30   '定制配送
Const SS2_EXPORT_FL = 31 '出口订单
Const SS2_ORD_CNT = 32

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          'Rolling default order
            Call Gp_Ms_Collection(txt_SlabNo, "p", "n", " ", "i", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_SlabSize, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_RMRollingSize, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_RollingSize, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_RmFinTmp, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            Call Gp_Ms_Collection(txt_stlgrd, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
               '2010.09.09 015725 加热轧/热处理交货状态显示
               Call Gp_Ms_Collection(txt_HTM, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_CrCd, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            Call Gp_Ms_Collection(txt_TrimFl, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
     'Control1 rolling order
      Call Gp_Ms_Collection(txt_CrMillRatet3, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_CrMillRatet4, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_CrMillRatet5, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_CrMillTmpt3, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_CrMillTmpt4, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_CrMillTmpt5, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     
          'cooling order
           Call Gp_Ms_Collection(txt_CoolMth, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_CoolSpeed, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_CoolTemp, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     
    'Cooling result
     Call Gp_Ms_Collection(SDB_COOL_AVE_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(SDB_COOL_EXT_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(SDB_COOL_ENT_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(SDB_COOL_WGT, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
           
           'Control1ed rolling result
             Call Gp_Ms_Collection(TXT_CR_CD, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(TXT_ROLLING_METHOD, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(SDB_CR_STAGE1_THK, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(SDB_CR_STAGE1_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(SDB_CR_STAGE2_THK, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(SDB_CR_STAGE2_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(SDB_CR_STAGE3_THK, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(SDB_CR_STAGE3_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)


         'Rolling result
     Call Gp_Ms_Collection(TXT_MILL_STA_TIME, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MILL_END_TIME, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
               Call Gp_Ms_Collection(txt_thk, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
               Call Gp_Ms_Collection(txt_wid, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
               Call Gp_Ms_Collection(txt_len, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_LastTemp, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             Call Gp_Ms_Collection(txt_Shift, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             Call Gp_Ms_Collection(TXT_GROUP, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(TXT_EMP1, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(TXT_EMP2, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(TXT_EMP3, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_Roll_Stlgrd, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    'Added by guoli at 20080806232000
          Call Gp_Ms_Collection(txt_millTemp, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_millTemp_min, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_millTemp_max, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      
         Call Gp_Ms_Collection(TXT_EXCEPTION, " ", " ", " ", "i", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
    'MASTER Collection
     Mc1.Add Item:="CGC2010C.P_MODIFY1", Key:="P-M"
     Mc1.Add Item:="CGC2010C.P_SEFER1", Key:="P-R"
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"
     
   
    Call Gp_Ms_Collection(txt_SlabNo, "p", "n", " ", " ", " ", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
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
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

   'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGC2010C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
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
   Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '订单号add by liqian 2012-06-11
   Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '序列号add by liqian 2012-06-11
   Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '是否探伤 add by liqian 2013-04-08
   Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '超交货期用红色显示 add by liqian 2012-06-11
   Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '超交货期用红色显示 add by liqian 2012-06-11
   Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '紧急订单绿色显示 add by liqian 2012-08-15
   Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 31, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '订单数量 20150119
   Call Gp_Sp_Collection(ss2, 32, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '订单数量 20150119
   Call Gp_Sp_Collection(ss2, 33, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '订单数量 20150119
   Call Gp_Sp_Collection(ss2, 34, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '订单数量 20150119
   Call Gp_Sp_Collection(ss2, 35, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '订单数量 20150119
   Call Gp_Sp_Collection(ss2, 36, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 37, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   
   
   'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CGC2010C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Call Gp_Ms_Collection(txt_RstFormDate, "p", "n", " ", " ", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
      Call Gp_Ms_Collection(txt_RstToDate, "p", "n", " ", " ", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
    
    'MASTER Collection
     Mc3.Add Item:=pControl3, Key:="pControl"
     Mc3.Add Item:=nControl3, Key:="nControl"
     Mc3.Add Item:=mControl3, Key:="mControl"
     Mc3.Add Item:=iControl3, Key:="iControl"
     Mc3.Add Item:=rControl3, Key:="rControl"
     Mc3.Add Item:=cControl3, Key:="cControl"
     Mc3.Add Item:=aControl3, Key:="aControl"
     Mc3.Add Item:=lControl3, Key:="lControl"
    
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 13, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   
   'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="CGC2010C.P_REFER3", Key:="P-R"
    sc3.Add Item:=pColumn3, Key:="pColumn"
    sc3.Add Item:=nColumn3, Key:="nColumn"
    sc3.Add Item:=aColumn3, Key:="aColumn"
    sc3.Add Item:=mColumn3, Key:="mColumn"
    sc3.Add Item:=iColumn3, Key:="iColumn"
    sc3.Add Item:=lColumn3, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss4, 1, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 3, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 4, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   
   'Spread_Collection
    sc4.Add Item:=ss4, Key:="Spread"
    sc4.Add Item:="CGC2010C.P_REFER4", Key:="P-R"
    sc4.Add Item:=pColumn4, Key:="pColumn"
    sc4.Add Item:=nColumn4, Key:="nColumn"
    sc4.Add Item:=aColumn4, Key:="aColumn"
    sc4.Add Item:=mColumn4, Key:="mColumn"
    sc4.Add Item:=iColumn4, Key:="iColumn"
    sc4.Add Item:=lColumn4, Key:="lColumn"
    sc4.Add Item:=1, Key:="First"
    sc4.Add Item:=ss4.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    
          'Rolling default order
            Call Gp_Ms_Collection(txt_SlabNo, "p", "n", " ", "i", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
          Call Gp_Ms_Collection(txt_SlabSize, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
     Call Gp_Ms_Collection(txt_RMRollingSize, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
       Call Gp_Ms_Collection(txt_RollingSize, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
          Call Gp_Ms_Collection(txt_RmFinTmp, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
            Call Gp_Ms_Collection(txt_stlgrd, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
               '2010.09.09 015725 加热轧/热处理交货状态显示
               Call Gp_Ms_Collection(txt_HTM, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl1)
              Call Gp_Ms_Collection(txt_CrCd, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
            Call Gp_Ms_Collection(txt_TrimFl, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    
     'Control4 rolling order
      Call Gp_Ms_Collection(txt_CrMillRatet3, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
      Call Gp_Ms_Collection(txt_CrMillRatet4, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
      Call Gp_Ms_Collection(txt_CrMillRatet5, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
       Call Gp_Ms_Collection(txt_CrMillTmpt3, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
       Call Gp_Ms_Collection(txt_CrMillTmpt4, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
       Call Gp_Ms_Collection(txt_CrMillTmpt5, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
     
          'cooling order
           Call Gp_Ms_Collection(txt_CoolMth, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
         Call Gp_Ms_Collection(txt_CoolSpeed, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
          Call Gp_Ms_Collection(txt_CoolTemp, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
     
    'Cooling result
     Call Gp_Ms_Collection(SDB_COOL_AVE_TEMP, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
     Call Gp_Ms_Collection(SDB_COOL_EXT_TEMP, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
     Call Gp_Ms_Collection(SDB_COOL_ENT_TEMP, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
          Call Gp_Ms_Collection(SDB_COOL_WGT, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
           
           'Control4ed rolling result
             Call Gp_Ms_Collection(TXT_CR_CD, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    Call Gp_Ms_Collection(TXT_ROLLING_METHOD, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
     Call Gp_Ms_Collection(SDB_CR_STAGE1_THK, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    Call Gp_Ms_Collection(SDB_CR_STAGE1_TEMP, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
     Call Gp_Ms_Collection(SDB_CR_STAGE2_THK, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    Call Gp_Ms_Collection(SDB_CR_STAGE2_TEMP, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
     Call Gp_Ms_Collection(SDB_CR_STAGE3_THK, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    Call Gp_Ms_Collection(SDB_CR_STAGE3_TEMP, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)


         'Rolling result
     Call Gp_Ms_Collection(TXT_MILL_STA_TIME, " ", "n", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
     Call Gp_Ms_Collection(TXT_MILL_END_TIME, " ", "n", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
               Call Gp_Ms_Collection(txt_thk, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
               Call Gp_Ms_Collection(txt_wid, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
               Call Gp_Ms_Collection(txt_len, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
          Call Gp_Ms_Collection(txt_LastTemp, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
             Call Gp_Ms_Collection(txt_Shift, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
             Call Gp_Ms_Collection(TXT_GROUP, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
              Call Gp_Ms_Collection(TXT_EMP1, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
              Call Gp_Ms_Collection(TXT_EMP2, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
              Call Gp_Ms_Collection(TXT_EMP3, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
       Call Gp_Ms_Collection(txt_Roll_Stlgrd, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    
      'Added by guoli at 20080806232000
          Call Gp_Ms_Collection(txt_millTemp, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
      Call Gp_Ms_Collection(txt_millTemp_min, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
      Call Gp_Ms_Collection(txt_millTemp_max, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
      
         Call Gp_Ms_Collection(TXT_EXCEPTION, " ", " ", " ", "i", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    
    'MASTER Collection
     Mc4.Add Item:="CGC2010C.P_SEFER2", Key:="P-R"
     Mc4.Add Item:=pControl4, Key:="pControl"
     Mc4.Add Item:=nControl4, Key:="nControl"
     Mc4.Add Item:=mControl4, Key:="mControl"
     Mc4.Add Item:=iControl4, Key:="iControl"
     Mc4.Add Item:=rControl4, Key:="rControl"
     Mc4.Add Item:=cControl4, Key:="cControl"
     Mc4.Add Item:=aControl4, Key:="aControl"
     Mc4.Add Item:=lControl4, Key:="lControl"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub CHK_A_Click()
    If CHK_A.Value = ssCBChecked Then
       TXT_EXCEPTION.Text = "A"
       CHK_B.Value = ssCBUnchecked
       chk_c.Value = ssCBUnchecked
       CHK_A.ForeColor = &HFF&
       CHK_B.ForeColor = &H80000012
       chk_c.ForeColor = &H80000012
    End If
    If CHK_A.Value = ssCBUnchecked And CHK_B.Value = ssCBUnchecked And chk_c.Value = ssCBUnchecked Then
       CHK_A.ForeColor = &H80000012
       CHK_B.ForeColor = &H80000012
       chk_c.ForeColor = &H80000012
       TXT_EXCEPTION.Text = ""
    End If
End Sub

Private Sub CHK_B_Click()
    If CHK_B.Value = ssCBChecked Then
       TXT_EXCEPTION.Text = "B"
       CHK_A.Value = ssCBUnchecked
       chk_c.Value = ssCBUnchecked
       CHK_B.ForeColor = &HFF&
       CHK_A.ForeColor = &H80000012
       chk_c.ForeColor = &H80000012
    End If
    If CHK_A.Value = ssCBUnchecked And CHK_B.Value = ssCBUnchecked And chk_c.Value = ssCBUnchecked Then
       CHK_A.ForeColor = &H80000012
       CHK_B.ForeColor = &H80000012
       chk_c.ForeColor = &H80000012
       TXT_EXCEPTION.Text = ""
    End If
End Sub

Private Sub chk_c_Click()
    If chk_c.Value = ssCBChecked Then
       TXT_EXCEPTION.Text = "C"
       CHK_A.Value = ssCBUnchecked
       CHK_B.Value = ssCBUnchecked
       chk_c.ForeColor = &HFF&
       CHK_A.ForeColor = &H80000012
       CHK_B.ForeColor = &H80000012
    End If
    If CHK_A.Value = ssCBUnchecked And CHK_B.Value = ssCBUnchecked And chk_c.Value = ssCBUnchecked Then
       CHK_A.ForeColor = &H80000012
       CHK_B.ForeColor = &H80000012
       chk_c.ForeColor = &H80000012
       TXT_EXCEPTION.Text = ""
    End If
End Sub

Private Sub CHK_CR_CD_Click()
   If CHK_CR_CD.Value = ssCBUnchecked Then
       If CHK_NON_CR_CD.Value = ssCBUnchecked Then
'          CHK_CR_CD.Value = ssCBChecked
          TXT_CR_CD.Text = ""
          CHK_CR_CD.ForeColor = &H80000012
          CHK_NON_CR_CD.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_CR_CD.Text = "1"
   
   CHK_CR_CD.ForeColor = &HFF&
   CHK_CR_CD.Value = ssCBChecked

   CHK_NON_CR_CD.ForeColor = &H808080
   CHK_NON_CR_CD.Value = ssCBUnchecked
End Sub

Private Sub CHK_NON_CR_CD_Click()
   If CHK_NON_CR_CD.Value = ssCBUnchecked Then
       If CHK_CR_CD.Value = ssCBUnchecked Then
'          CHK_NON_CR_CD.Value = ssCBChecked
          TXT_CR_CD.Text = ""
          CHK_NON_CR_CD.ForeColor = &H80000012
          CHK_CR_CD.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_CR_CD.Text = "0"
   
   CHK_NON_CR_CD.ForeColor = &HFF&
   CHK_NON_CR_CD.Value = ssCBChecked

   CHK_CR_CD.ForeColor = &H808080
   CHK_CR_CD.Value = ssCBUnchecked
End Sub

Private Sub Chk_Rolling_Auto_Click()
   
  If CHK_ROLLING_AUTO.Value = ssCBUnchecked Then
       If CHK_ROLLING_OP.Value = ssCBUnchecked Then
          TXT_ROLLING_METHOD.Text = ""
          CHK_ROLLING_AUTO.ForeColor = &H80000012
          CHK_ROLLING_OP.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_ROLLING_METHOD.Text = "0"
   
   CHK_ROLLING_AUTO.ForeColor = &HFF&
   CHK_ROLLING_AUTO.Value = ssCBChecked

   CHK_ROLLING_OP.ForeColor = &H808080
   CHK_ROLLING_OP.Value = ssCBUnchecked
  
End Sub

Private Sub Chk_Rolling_Op_Click()
  
  If CHK_ROLLING_OP.Value = ssCBUnchecked Then
       If CHK_ROLLING_AUTO.Value = ssCBUnchecked Then
          TXT_ROLLING_METHOD.Text = ""
          CHK_ROLLING_OP.ForeColor = &H80000012
          CHK_ROLLING_AUTO.ForeColor = &H80000012
       End If
       Exit Sub
   End If
   
   TXT_ROLLING_METHOD.Text = "1"
   
   CHK_ROLLING_OP.ForeColor = &HFF&
   CHK_ROLLING_OP.Value = ssCBChecked

   CHK_ROLLING_AUTO.ForeColor = &H808080
   CHK_ROLLING_AUTO.Value = ssCBUnchecked
      
End Sub

Private Sub cmd_LPass_Click()

If Not Gf_MessConfirm("您确定板坯号 " & txt_SlabNo.Text & " 仅通过精轧机吗？", "W", "") Then
   Exit Sub
End If

    Dim OutParam(2, 4) As Variant
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    Dim SMESG As String
    
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
    
    If Mid(TXT_MILL_END_TIME, 1, 1) <> "2" Then
         SMESG = " 请输入终轧时间...！"
         Call Gp_MsgBoxDisplay(SMESG)
         Screen.MousePointer = DEFAULT
         Exit Sub
    End If
    
    sQuery = "{call CGC2010C.P_PASS('" & Trim(txt_SlabNo.Text) & "','" & TXT_CB & "',?,?)}"
    
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
    TXT_MILL_END_TIME = ""
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Pass_Click()

If Not Gf_MessConfirm("您确定要对板坯号 " & txt_SlabNo.Text & " 做轧废处理吗？", "W", "") Then
   Exit Sub
End If

    Dim OutParam(2, 4) As Variant
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    Dim SMESG As String
    
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
    
    If Len(Trim(TXT_MILL_END_TIME.RawData)) <> 14 Then
         SMESG = " 请输入终轧时间...！"
         Call Gp_MsgBoxDisplay(SMESG)
         Screen.MousePointer = DEFAULT
         Exit Sub
    End If
    
    sQuery = "{call CGC2010C.P_SCRAP('" & Trim(txt_SlabNo.Text) & "','" & txt_Shift & "','" & TXT_GROUP & "','" & TXT_EMP1 & "','" & TXT_CB & "',?,?)}"
    
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
    TXT_MILL_END_TIME = ""
    
    Screen.MousePointer = vbDefault

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
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_Cls(Mc3("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_ControlLock(Mc2("lControl"), True)
    Call Gp_Ms_ControlLock(Mc3("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    Call Gp_Ms_NeceColor(Mc3("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    Call Gp_Sp_Setting(sc3.Item("Spread"))
    Call Gp_Sp_Setting(sc4.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(sc3)
    Call Gf_Sp_Cls(sc4)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc3.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc4.Item("Spread"), "CG-System.INI", Me.Name)

    tab1.Tab = 0
    Call Form_Ref

    txt_Shift = Gf_ShiftSet3(M_CN1)
    TXT_GROUP = Gf_GroupSet(M_CN1, Trim(txt_Shift), Gf_DTSet(M_CN1, , "X"))
    TXT_EMP1 = sUserID
    
    If Mid(sAuthority, 1, 3) = "111" Then
       cmd_Pass.Enabled = True
       SSCommand1.Enabled = True
       cmd_LPass.Enabled = True
    Else
       cmd_Pass.Enabled = False
       SSCommand1.Enabled = False
       cmd_LPass.Enabled = False
    End If

    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc3.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc4.Item("Spread"), "CG-System.INI", Me.Name)

    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
    Set pControl3 = Nothing
    Set nControl3 = Nothing
    Set iControl3 = Nothing
    Set rControl3 = Nothing
    Set cControl3 = Nothing
    Set aControl3 = Nothing
    Set lControl3 = Nothing
    Set mControl3 = Nothing
    
    Set pControl4 = Nothing
    Set nControl4 = Nothing
    Set iControl4 = Nothing
    Set rControl4 = Nothing
    Set cControl4 = Nothing
    Set aControl4 = Nothing
    Set lControl4 = Nothing
    Set mControl4 = Nothing
    
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
    
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing
    
    Set iColumn4 = Nothing
    Set pColumn4 = Nothing
    Set lColumn4 = Nothing
    Set nColumn4 = Nothing
    Set mColumn4 = Nothing
    Set aColumn4 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set Mc4 = Nothing
    
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set sc3 = Nothing
    Set sc4 = Nothing
    Set Proc_Sc = Nothing

     Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()


    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_Cls(Mc3("rControl"))

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(sc3)
    Call Gf_Sp_Cls(sc4)
    
    CHK_A.Value = ssCBUnchecked
    CHK_B.Value = ssCBUnchecked
    chk_c.Value = ssCBUnchecked
    
    TXT_EXCEPTION.Text = ""

End Sub


Public Sub Form_Ref()
    Dim iRow As Integer
    Dim iCol As Integer
    Dim sCurDate As String
    Dim sDel_To_Date As String
    Dim sUrgnt_Fl As String
    Dim sUst_Fl As String
    Dim sFlag As String
    Dim sexport As String
    Dim sOrdcnt As String
    
    sCurDate = Format(Now, "YYYYMM")
     
    If tab1.Tab = 0 Then
        Call Gf_Sp_Refer(M_CN1, sc2, , , , False)
        Call Gf_Sp_Refer(M_CN1, sc4, , , , False)
        ss2.Col = 1
        ss2.ROW = 1
        If ss2.Text <> "" Then
            Call ss2_DblClick(1, 1)
        Else
            txt_SlabNo.Text = ""
        End If
        '超交货期用红色显示 add by liqian 2012-06-11
         With ss2
              For iRow = 1 To .MaxRows
                 .ROW = iRow:             .Col = SS2_DEL_TO_DATE
                  sDel_To_Date = Mid(.Value, 1, 6)
                  .Col = SS2_URGNT_FL:    sUrgnt_Fl = Trim(.Text)
                  If sDel_To_Date < sCurDate Then
                       Call Gp_Sp_BlockColor(ss2, 1, .MaxCols, iRow, iRow, &HFF&)
                  End If
                  '紧急订单绿色显示 add by liqian 2012-08-15
                  If sUrgnt_Fl = "Y" Then
                     Call Gp_Sp_BlockColor(ss2, 1, .MaxCols, iRow, iRow, &HC000&)
                  End If
                  '是否探伤 add by liqian 2013-04-08
                  .ROW = iRow:
                  .Col = SS2_UST_FL:   sUst_Fl = Trim(.Text)
                  If sUst_Fl = "是" Then
                     Call Gp_Sp_BlockColor(ss2, SS2_SLAB_NO, SS2_SLAB_NO, iRow, iRow, &HFF00FF)
                     Call Gp_Sp_BlockColor(ss2, SS2_UST_FL, SS2_UST_FL, iRow, iRow, &HFF00FF)
                  End If
                  '是否定制配送
                  .ROW = iRow:
                  .Col = SS2_FLAG_FL: sFlag = Trim(.Text)
                  If sFlag = "Y" Then
                     Call Gp_Sp_BlockColor(ss2, SS2_SLAB_NO, SS2_SLAB_NO, iRow, iRow, SSP4.BackColor)
                  End If
                  '是否出口订单
                  .ROW = iRow:
                  .Col = SS2_EXPORT_FL: sexport = Trim(.Text)
                  If sexport = "Y" Then
                     Call Gp_Sp_BlockColor(ss2, SS2_SLAB_NO, SS2_SLAB_NO, iRow, iRow, SSP1.BackColor)
                  End If
                  '是否一坯多订单
                  .ROW = iRow:
                  .Col = SS2_ORD_CNT: sOrdcnt = Val(.Text)
                  If sOrdcnt > 1 Then
                    Call Gp_Sp_BlockColor(ss2, 1, .MaxCols, iRow, iRow, , SSP2.BackColor)
                  End If
              Next iRow
        End With
    ElseIf tab1.Tab = 1 Then
        Call Gf_Sp_Refer(M_CN1, sc3, Mc3, Mc3("nControl"), Mc3("mControl"), False)
        ss3.Col = 1
        ss3.ROW = 1
        If ss3.Text <> "" Then
            Call ss3_DblClick(1, 1)
        End If
    End If

    txt_millTemp.Enabled = True
    txt_millTemp.ReadOnly = True
    txt_millTemp_min.Enabled = True
    txt_millTemp_min.ReadOnly = True
    txt_millTemp_max.Enabled = True
    txt_millTemp_max.ReadOnly = True
    
End Sub

Public Sub Form_Pro()
Dim SMESG As String
    
    If Not Gp_DateCheck(TXT_MILL_STA_TIME) Then
            SMESG = " 请正确输入开轧时间 ！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
    End If
    
    If Not Gp_DateCheck(TXT_MILL_END_TIME) Then
            SMESG = " 请正确输入终轧时间 ！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
    End If

    Call Gf_Ms_Process(M_CN1, Mc1, sAuthority)
    Call Form_Ref
    
    TXT_MILL_END_TIME = ""
    txt_Shift = Gf_ShiftSet3(M_CN1)
    TXT_GROUP = Gf_GroupSet(M_CN1, Trim(txt_Shift), Gf_DTSet(M_CN1, , "X"))
    TXT_EMP1 = sUserID
    
'    Dim sMesg As String
'    Dim Temp_no As String
'
'    Temp_no = CBO_SLAB_NO.Text
'
'    TXT_UPD_EMP = sUserID
'
'    Select Case SSTab1.Tab
'
'          Case 0
'
'                 If Not Gp_DateCheck(TXT_MILL_STA_TIME) Then
'                      sMesg = " 请正确输入开轧时间 ！"
'                      Call Gp_MsgBoxDisplay(sMesg)
'                      Exit Sub
'                 End If
'
'                 If TXT_MILL_STA_TIME.RawData = "" And TXT_MILL_END_TIME.RawData = "" Then
'                      sMesg = " 请输入开轧时间 ！"
'                      Call Gp_MsgBoxDisplay(sMesg)
'                      Exit Sub
'                 ElseIf TXT_MILL_STA_TIME.RawData = "" And TXT_MILL_END_TIME.RawData <> "" Then
'                      sMesg = " 请首先输入开轧时间 ！"
'                      Call Gp_MsgBoxDisplay(sMesg)
'                      Exit Sub
'                 ElseIf TXT_MILL_STA_TIME.RawData <> "" And TXT_MILL_END_TIME.RawData <> "" Then
'                        If Not Gp_DateCheck(TXT_MILL_END_TIME) Then
'                             sMesg = " 请正确输入终轧时间 ！"
'                             Call Gp_MsgBoxDisplay(sMesg)
'                             Exit Sub
'                        End If
'                        If Val(TXT_MILL_STA_TIME.RawData) - Val(TXT_MILL_END_TIME.RawData) > 0 Then
'                             sMesg = " 终轧时间应大于开轧时间，请正确输入时间信息 ！"
'                             Call Gp_MsgBoxDisplay(sMesg)
'                             Exit Sub
'                        End If
'                 End If
'
'                 If Trim(TXT_CR_CD) = "1" Then
'                    If Trim(SDB_CR_STAGE1_THK) = "" And Trim(SDB_CR_STAGE1_TEMP) = "" And Trim(SDB_CR_STAGE1_TIME) = "" Then
'                        sMesg = " 请输入控轧一阶段厚度，温度，待轧时间 ！"
'                        Call Gp_MsgBoxDisplay(sMesg)
'                        Exit Sub
'                    End If
'                 Else
'                    SDB_CR_STAGE1_THK = ""
'                    SDB_CR_STAGE1_TEMP = ""
'                    SDB_CR_STAGE1_TIME = ""
'                    SDB_CR_STAGE2_THK = ""
'                    SDB_CR_STAGE2_TEMP = ""
'                    SDB_CR_STAGE2_TIME = ""
'                    SDB_CR_STAGE3_THK = ""
'                    SDB_CR_STAGE3_TEMP = ""
'                    SDB_CR_STAGE3_TIME = ""
'                 End If
'                 If Gf_Mc_Authority(sAuthority, Mc1) Then
'                   ' txt_ins_emp.Text = sUserID
'                   If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
'                 End If
''                 Call Gf_Mill_ComboAdd(M_CN1, CBO_SLAB_NO, "CB")
'                ' Call Gf_Common_ComboSet(M_CN1, CBO_SLAB_NO, "CA")
'                 CBO_SLAB_NO.Text = Temp_no
'          Case 1
'                 If TXT_CUTEND_CD.Text = "Y" And TXT_COMFRM.Text = "2" Then
'                    sMesg = " （缺号母板确定） 与 （母板剪切结束确定）不能同时操作 ！"
'                    Call Gp_MsgBoxDisplay(sMesg)
'                    Exit Sub
'                 End If
'
'                 If TXT_CUTEND_CD.Text = "Y" Then
'                    sMesg = " 确定此轧件剪切母板结束 ？ "
'                 ElseIf TXT_COMFRM.Text = "2" Then
'                    sMesg = " 确定以下母板缺号 ？ "
'                 Else
'                    If Gf_Mc_Authority(sAuthority, Mc1) Then
'                       If Gf_Ms_Process(M_CN1, Mc2, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
'                    End If
'                    Exit Sub
'                 End If
'                 If Gp_MsgBox(sMesg, "C") = 6 Then
'                    If Gf_Mc_Authority(sAuthority, Mc1) Then
'                       If Gf_Ms_Process(M_CN1, Mc2, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
'                    End If
'                 End If
'
'   End Select
   
End Sub

Public Sub Form_Del()

'    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub


Private Sub Label2_Click()

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        ss2.ROW = ROW
        ss2.Col = 1
        txt_SlabNo.Text = ss2.Text
        
        Call Gf_Ms_Refer(M_CN1, Mc1, , , False)
        Call Gf_Sp_Refer(M_CN1, sc1, Mc2, Mc2("nControl"), Mc2("mControl"))

        TXT_MILL_STA_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        'TXT_MILL_END_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        
        txt_Shift = Gf_ShiftSet3(M_CN1)
        TXT_GROUP = Gf_GroupSet(M_CN1, Trim(txt_Shift), Gf_DTSet(M_CN1, , "X"))
        TXT_EMP1 = sUserID
    
    End If
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        ss2.ROW = ROW
        ss2.Col = 1
        txt_SlabNo.Text = ss2.Text
        
        Call Gf_Ms_Refer(M_CN1, Mc1, , , False)
        Call Gf_Sp_Refer(M_CN1, sc1, Mc2, Mc2("nControl"), Mc2("mControl"))

        TXT_MILL_STA_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        'TXT_MILL_END_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        
        txt_Shift = Gf_ShiftSet3(M_CN1)
        TXT_GROUP = Gf_GroupSet(M_CN1, Trim(txt_Shift), Gf_DTSet(M_CN1, , "X"))
        TXT_EMP1 = sUserID
    
    End If
End Sub

Private Sub ss3_DblClick(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        ss3.ROW = ROW
        ss3.Col = 1
        txt_SlabNo.Text = ss3.Text
        
        Call Gf_Ms_Refer(M_CN1, Mc4, , , False)
        Call Gf_Sp_Refer(M_CN1, sc1, Mc2, Mc2("nControl"), Mc2("mControl"))
        
        txt_millTemp.Enabled = True
        txt_millTemp.ReadOnly = True
        txt_millTemp_min.Enabled = True
        txt_millTemp_min.ReadOnly = True
        txt_millTemp_max.Enabled = True
        txt_millTemp_max.ReadOnly = True
    
    End If
End Sub

Private Sub SSCommand1_Click()
    Dim OutParam(2, 4) As Variant
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    Dim SMESG As String
    
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
    
'    If Len(Trim(TXT_SLABNO.Text)) < 10 Then
'         sMesg = " 请输入板坯号码 ！"
'         Call Gp_MsgBoxDisplay(sMesg)
'         Screen.MousePointer = DEFAULT
'         Exit Sub
'    End If
'
'    If txt_thk.RawData < 100 Then
'         sMesg = " 请输入厚度 ！"
'         Call Gp_MsgBoxDisplay(sMesg)
'         Screen.MousePointer = DEFAULT
'         Exit Sub
'    ElseIf txt_wid.RawData < 1000 Then
'         sMesg = " 请输入宽度 ！"
'         Call Gp_MsgBoxDisplay(sMesg)
'         Screen.MousePointer = DEFAULT
'         Exit Sub
'    ElseIf txt_len.RawData < 1000 Then
'         sMesg = " 请输入长度 ！"
'         Call Gp_MsgBoxDisplay(sMesg)
'         Screen.MousePointer = DEFAULT
'         Exit Sub
'    End If
    
    sQuery = "{call CGC2010C.P_REJECT('" & Trim(txt_SlabNo.Text) & "'," & txt_thk.RawData & "," & txt_wid.RawData & "," & txt_len.RawData & ",'" & txt_Shift & "','" & TXT_GROUP & "','" & TXT_EMP1 & "','" & TXT_CB & "',?,?)}"
    
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
    TXT_MILL_END_TIME = ""
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub tab1_Click(PreviousTab As Integer)
    If tab1.Tab = "1" Then
        txt_Shift = Gf_ShiftSet3(M_CN1)
        If txt_Shift = "1" Then
            txt_RstFormDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "000001"
            txt_RstToDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "081459"
        ElseIf txt_Shift = "2" Then
            txt_RstFormDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "081500"
            txt_RstToDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "155959"
        ElseIf txt_Shift = "3" Then
            txt_RstFormDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "160000"
            txt_RstToDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "235959"
        End If
    ElseIf tab1.Tab = "0" Then
        TXT_MILL_STA_TIME.RawData = Gf_DTSet(M_CN1, , "X") 'Format(Now, "YYYYMMDDHHMMSS")
    End If
End Sub

Private Sub TXT_EMP1_DblClick()
    Call TXT_EMP1_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_EMP1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        TXT_EMP1.Text = ""
        DD.rControl.Add Item:=TXT_EMP1

        Call Gf_EmpID_DD(M_CN1, vbKeyF4, "1ZB")

        Exit Sub
    End If
End Sub

Private Sub TXT_EMP2_DblClick()
    Call TXT_EMP2_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_EMP2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        TXT_EMP2.Text = ""
        DD.rControl.Add Item:=TXT_EMP2

        Call Gf_EmpID_DD(M_CN1, vbKeyF4, "1ZB")

        Exit Sub
    End If
End Sub

Private Sub TXT_EMP3_DblClick()
    Call TXT_EMP3_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_EMP3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        TXT_EMP3.Text = ""
        DD.rControl.Add Item:=TXT_EMP3

        Call Gf_EmpID_DD(M_CN1, vbKeyF4, "1ZB")

        Exit Sub
    End If
End Sub

'Private Sub TXT_MILL_END_TIME_LostFocus()
'    txt_Shift = Gf_ShiftSet3(M_CN1, Mid(TXT_MILL_END_TIME, 12, 2))
'    TXT_GROUP = Gf_GroupSet(M_CN1, Trim(txt_Shift), Gf_DTSet(M_CN1, , "X"))
'    TXT_EMP1 = sUserID
'End Sub

Private Sub TXT_MILL_STA_TIME_DblClick()

    TXT_MILL_STA_TIME.RawData = Gf_DTSet(M_CN1, , "X") 'Format(Now, "YYYYMMDDHHMMSS")

End Sub

Private Sub TXT_MILL_END_TIME_DblClick()

    TXT_MILL_END_TIME.RawData = Gf_DTSet(M_CN1, , "X") 'Format(Now, "YYYYMMDDHHMMSS")

End Sub

Private Sub txt_millTemp_max_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    txt_millTemp_max.ToolTipText = "终轧温度最大偏差"
End Sub

Private Sub txt_millTemp_min_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    txt_millTemp_min.ToolTipText = "终轧温度最小偏差"
End Sub

Private Sub txt_millTemp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    txt_millTemp.ToolTipText = "终轧目标温度"
End Sub

Private Sub txt_RstFormDate_DblClick()
    txt_RstFormDate.RawData = Gf_DTSet(M_CN1, , "X")
    txt_RstToDate.RawData = Gf_DTSet(M_CN1, , "X")
End Sub
