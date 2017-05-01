VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AEC0000C 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "炼钢紧急编制"
   ClientHeight    =   4170
   ClientLeft      =   4650
   ClientTop       =   4050
   ClientWidth     =   5640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5640
   Begin Threed.SSPanel pnl_first 
      Height          =   3345
      Left            =   105
      TabIndex        =   18
      Top             =   90
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5900
      _Version        =   196609
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
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cbo_mill_plt 
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
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "AEC0000C.frx":0000
         Left            =   1440
         List            =   "AEC0000C.frx":0002
         TabIndex        =   24
         Tag             =   "使用工厂"
         Top             =   180
         Width           =   675
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   30
         Left            =   30
         TabIndex        =   22
         Top             =   930
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   53
         _Version        =   196609
         Caption         =   "SSPanel1"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.ComboBox cbo_ccm_no 
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
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "AEC0000C.frx":0004
         Left            =   2130
         List            =   "AEC0000C.frx":0006
         TabIndex        =   21
         Tag             =   "连铸机号"
         Top             =   555
         Width           =   675
      End
      Begin VB.TextBox txt_stlgrd 
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
         Left            =   3930
         MaxLength       =   18
         TabIndex        =   1
         Tag             =   "钢种"
         Top             =   180
         Width           =   1365
      End
      Begin VB.ComboBox cbo_line_no 
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
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "AEC0000C.frx":0008
         Left            =   1440
         List            =   "AEC0000C.frx":000A
         TabIndex        =   0
         Tag             =   "炉座号"
         Top             =   555
         Width           =   675
      End
      Begin InDate.ULabel ULabel48 
         Height          =   315
         Left            =   120
         Tag             =   "使用工厂"
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "使用工厂"
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
         Left            =   2850
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "钢    种"
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
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   405
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "板坯类型 1"
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
      Begin CSTextLibCtl.sidbEdit sdb_wid_1 
         Height          =   315
         Left            =   1545
         TabIndex        =   2
         Tag             =   "板坯宽度"
         Top             =   1440
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         NumIntDigits    =   6
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   2850
         Top             =   555
         Width           =   1035
         _ExtentX        =   1826
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
      End
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   405
         Top             =   1785
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "板坯类型 2"
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
      Begin CSTextLibCtl.sidbEdit sdb_wid_2 
         Height          =   315
         Left            =   1545
         TabIndex        =   5
         Tag             =   "板坯宽度2"
         Top             =   1785
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         NumIntDigits    =   6
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   120
         Top             =   555
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "炉座/连铸机"
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
         Left            =   1545
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
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
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   2715
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
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
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   3870
         Top             =   1095
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "板坯数量"
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
      Begin CSTextLibCtl.sidbEdit sdb_cnt_1 
         Height          =   315
         Left            =   3855
         TabIndex        =   4
         Tag             =   "板坯数量1"
         Top             =   1440
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
      Begin CSTextLibCtl.sidbEdit sdb_len_1 
         Height          =   315
         Left            =   2700
         TabIndex        =   3
         Tag             =   "板坯宽度1"
         Top             =   1425
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
      Begin CSTextLibCtl.sidbEdit sdb_len_2 
         Height          =   315
         Left            =   2700
         TabIndex        =   6
         Tag             =   "板坯长度2"
         Top             =   1785
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
      Begin CSTextLibCtl.sidbEdit sdb_cnt_2 
         Height          =   315
         Left            =   3855
         TabIndex        =   7
         Tag             =   "板坯数量2"
         Top             =   1800
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   390
         Top             =   2145
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "板坯类型 3"
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
      Begin CSTextLibCtl.sidbEdit sdb_wid_3 
         Height          =   315
         Left            =   1545
         TabIndex        =   8
         Tag             =   "板坯宽度3"
         Top             =   2145
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         NumIntDigits    =   6
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   390
         Top             =   2490
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "板坯类型 4"
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
      Begin CSTextLibCtl.sidbEdit sdb_wid_4 
         Height          =   315
         Left            =   1545
         TabIndex        =   11
         Tag             =   "板坯宽度4"
         Top             =   2490
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         NumIntDigits    =   6
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_cnt_3 
         Height          =   315
         Left            =   3855
         TabIndex        =   10
         Tag             =   "板坯数量3"
         Top             =   2145
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
      Begin CSTextLibCtl.sidbEdit sdb_len_3 
         Height          =   315
         Left            =   2700
         TabIndex        =   9
         Tag             =   "板坯长度3"
         Top             =   2145
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
      Begin CSTextLibCtl.sidbEdit sdb_len_4 
         Height          =   315
         Left            =   2700
         TabIndex        =   12
         Tag             =   "板坯长度4"
         Top             =   2490
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
      Begin CSTextLibCtl.sidbEdit sdb_cnt_4 
         Height          =   315
         Left            =   3855
         TabIndex        =   13
         Tag             =   "板坯数量4"
         Top             =   2505
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   390
         Top             =   2850
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "板坯类型 5"
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
      Begin CSTextLibCtl.sidbEdit sdb_wid_5 
         Height          =   315
         Left            =   1545
         TabIndex        =   14
         Tag             =   "板坯宽度5"
         Top             =   2850
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         NumIntDigits    =   6
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_len_5 
         Height          =   315
         Left            =   2700
         TabIndex        =   15
         Tag             =   "板坯长度5"
         Top             =   2850
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
      Begin CSTextLibCtl.sidbEdit sdb_cnt_5 
         Height          =   315
         Left            =   3855
         TabIndex        =   16
         Tag             =   "板坯数量5"
         Top             =   2865
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
      Begin CSTextLibCtl.sidbEdit sdb_thk 
         Height          =   315
         Left            =   3930
         TabIndex        =   23
         Tag             =   "板坯厚度"
         Top             =   555
         Width           =   1365
         _Version        =   262145
         _ExtentX        =   2408
         _ExtentY        =   547
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   135
         Width           =   105
      End
   End
   Begin Threed.SSCommand cmd_OK 
      Height          =   465
      Left            =   1545
      TabIndex        =   17
      Top             =   3600
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
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
      Caption         =   "確認"
   End
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   465
      Left            =   3075
      TabIndex        =   20
      Top             =   3600
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
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
      Caption         =   "取消"
   End
End
Attribute VB_Name = "AEC0000C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name
'-- Sub_System Name
'-- Program Name      炼钢紧急编制
'-- Program ID        AEC0000C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim S.H
'-- Coder             Kim S.H
'-- Date              2005.12.29
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Private Sub cbo_ccm_no_Change()

    Dim strQuery As String
    
    strQuery = "SELECT A.THK_MAX FROM EP_CCM_CON A WHERE  A.PRC_LINE ='" + Trim(cbo_ccm_no.Text) + "' AND USE_FL = '1' "
    sdb_thk.Value = Gf_FloatFind(M_CN1, strQuery)
    
    'cbo_CCM_NO.Text = cbo_line_no.Text

End Sub

Private Sub cbo_ccm_no_Click()

    Dim strQuery As String
    
    strQuery = "SELECT A.THK_MAX FROM EP_CCM_CON A WHERE  A.PRC_LINE ='" + Trim(cbo_ccm_no.Text) + "' AND USE_FL = '1' "
    sdb_thk.Value = Gf_FloatFind(M_CN1, strQuery)
    
    'cbo_CCM_NO.Text = cbo_line_no.Text

End Sub

Private Sub Cmd_Cancel_Click()

    Unload Me
    
End Sub

Private Sub Cmd_Ok_Click()

    If Trim(cbo_mill_plt.Text) = "" Or (Trim(cbo_mill_plt.Text) <> "C1" And Trim(cbo_mill_plt.Text) <> "C2" And Trim(cbo_mill_plt.Text) <> "C3" And Trim(cbo_mill_plt.Text) <> "CZ") Then
        Call Gp_MsgBoxDisplay(cbo_mill_plt.Tag & "必须输入", "", "错误提示")
        Exit Sub
    End If
    
    If Trim(cbo_line_no.Text) = "" Then
        Call Gp_MsgBoxDisplay(cbo_line_no.Tag & "必须输入", "", "错误提示")
        Exit Sub
    End If
    
    If Trim(cbo_ccm_no.Text) = "" Then
        Call Gp_MsgBoxDisplay(cbo_ccm_no.Tag & "必须输入", "", "错误提示")
        Exit Sub
    End If

    If Trim(txt_stlgrd.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_stlgrd.Tag & "必须输入", "", "错误提示")
        Exit Sub
    End If
              
    If sdb_thk.Value = 0 Then
        Call Gp_MsgBoxDisplay(sdb_thk.Tag & "必须输入", "", "错误提示")
        Exit Sub
    End If
       
    If sdb_wid_1.Value = 0 And _
       sdb_wid_2.Value = 0 And _
       sdb_wid_3.Value = 0 And _
       sdb_wid_4.Value = 0 And _
       sdb_wid_5.Value = 0 Then
        Call Gp_MsgBoxDisplay(sdb_wid_1.Tag & "必须输入", "", "错误提示")
        Exit Sub
    End If
         
    If sdb_wid_1.Value > 0 Then
        If sdb_len_1.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_len_1.Tag & "必须输入", "", "错误提示")
        End If
        If sdb_cnt_1.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_cnt_1.Tag & "必须输入", "", "错误提示")
        End If
    End If
    
    If sdb_wid_2.Value > 0 Then
        If sdb_len_2.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_len_2.Tag & "必须输入", "", "错误提示")
        End If
        If sdb_cnt_2.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_cnt_2.Tag & "必须输入", "", "错误提示")
        End If
    End If
    
    If sdb_wid_3.Value > 0 Then
        If sdb_len_3.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_len_3.Tag & "必须输入", "", "错误提示")
        End If
        If sdb_cnt_3.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_cnt_3.Tag & "必须输入", "", "错误提示")
        End If
    End If
        If sdb_wid_4.Value > 0 Then
        If sdb_len_4.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_len_4.Tag & "必须输入", "", "错误提示")
        End If
        If sdb_cnt_4.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_cnt_4.Tag & "必须输入", "", "错误提示")
        End If
    End If
    
    If sdb_wid_5.Value > 0 Then
        If sdb_len_5.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_len_5.Tag & "必须输入", "", "错误提示")
        End If
        If sdb_cnt_5.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_cnt_5.Tag & "必须输入", "", "错误提示")
        End If
    End If
    
    Call Gp_Process_Exec
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    cbo_line_no.Clear
    
    cbo_line_no.AddItem "1"
    cbo_line_no.AddItem "2"
    cbo_line_no.AddItem "3"
    cbo_line_no.ListIndex = 0

    cbo_ccm_no.AddItem "1"
    cbo_ccm_no.AddItem "2"
    cbo_ccm_no.AddItem "3"
    cbo_ccm_no.ListIndex = 0
    
    cbo_mill_plt.AddItem "C1"
    cbo_mill_plt.AddItem "C2"
    cbo_mill_plt.AddItem "C3"
    cbo_mill_plt.AddItem "CZ"  ' 外售板坯
    cbo_mill_plt.ListIndex = 0
    
    Call Gp_FormCenter(Me)
    Me.BackColor = &HE0E0E0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set Active_Spread = Nothing

End Sub

Public Sub Gp_Process_Exec()

    Dim OutParam(1, 4)      As Variant
    Dim ret_Result_ErrMsg   As String
    Dim sQuery              As String
    
    Dim adoCmd As ADODB.Command

    On Error GoTo Process_Exec_ERROR

    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
                                 
    sQuery = "{call AEC0000P ('B1'," & _
                             "'" & Trim(cbo_mill_plt.Text) & "'," & _
                             "'" & Trim(cbo_line_no.Text) & "'," & _
                             "'" & Trim(cbo_ccm_no.Text) & "'," & _
                             "'" & Trim(txt_stlgrd.Text) & "'," & _
                                   sdb_thk.Value & "," & _
                                   sdb_wid_1.Value & "," & _
                                   sdb_len_1.Value & "," & _
                                   sdb_cnt_1.Value & "," & _
                                   sdb_wid_2.Value & "," & _
                                   sdb_len_2.Value & "," & _
                                   sdb_cnt_2.Value & "," & _
                                   sdb_wid_3.Value & "," & _
                                   sdb_len_3.Value & "," & _
                                   sdb_cnt_3.Value & "," & _
                                   sdb_wid_4.Value & "," & _
                                   sdb_len_4.Value & "," & _
                                   sdb_cnt_4.Value & "," & _
                                   sdb_wid_5.Value & "," & _
                                   sdb_len_5.Value & "," & _
                                   sdb_cnt_5.Value & "," & _
                             "'" & sUserID & "',?)}"

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
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay("Error Mesg : " & ret_Result_ErrMsg)
        Set adoCmd = Nothing
        Exit Sub
    Else
        Call Gp_MsgBoxDisplay("炼钢紧急编制完了..!!", "I")
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        Call AEC1070C.Form_Ref
        Unload Me
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

Private Sub txt_stlgrd_DblClick()

    Call txt_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stlgrd
        
        DD.nameType = "1"
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
    End If
    
End Sub
