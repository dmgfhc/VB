VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AKN2041C 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Ǽƻ��ı���"
   ClientHeight    =   4170
   ClientLeft      =   6105
   ClientTop       =   4425
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
      TabIndex        =   20
      Top             =   90
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5900
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Begin VB.ComboBox cbo_plan_name 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "AKN2041C.frx":0000
         Left            =   1395
         List            =   "AKN2041C.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "�ƻ���"
         Top             =   540
         Width           =   1485
      End
      Begin VB.ComboBox cbo_mill_plt 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "AKN2041C.frx":0004
         Left            =   1395
         List            =   "AKN2041C.frx":0006
         TabIndex        =   0
         Tag             =   "ʹ�ù���"
         Top             =   180
         Width           =   675
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   30
         Left            =   30
         TabIndex        =   23
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
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "AKN2041C.frx":0008
         Left            =   4290
         List            =   "AKN2041C.frx":000A
         TabIndex        =   1
         Tag             =   "������"
         Top             =   195
         Width           =   675
      End
      Begin InDate.ULabel ULabel48 
         Height          =   315
         Left            =   180
         Tag             =   "ʹ�ù���"
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "ʹ�ù���"
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   180
         Top             =   540
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "�ƻ���"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   180
         Top             =   1440
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "�������� 1"
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
      Begin CSTextLibCtl.sidbEdit sdb_wid_1 
         Height          =   315
         Left            =   1395
         TabIndex        =   4
         Tag             =   "��������1"
         Top             =   1440
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   16711680
         BackColor       =   12648447
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
         MaxValue        =   0
         MinValue        =   99999
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   3075
         Top             =   555
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "�������"
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
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   180
         Top             =   1785
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "�������� 2"
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
      Begin CSTextLibCtl.sidbEdit sdb_wid_2 
         Height          =   315
         Left            =   1395
         TabIndex        =   7
         Tag             =   "��������2"
         Top             =   1785
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   16711680
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         MaxValue        =   0
         MinValue        =   99999
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   3075
         Top             =   195
         Width           =   1185
         _ExtentX        =   2090
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   1395
         Top             =   1080
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "��������"
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
         Left            =   2625
         Top             =   1080
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "��������"
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
         Left            =   3870
         Top             =   1080
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "��������"
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
      Begin CSTextLibCtl.sidbEdit sdb_cnt_1 
         Height          =   315
         Left            =   3870
         TabIndex        =   6
         Tag             =   "��������1"
         Top             =   1440
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         Left            =   2625
         TabIndex        =   5
         Tag             =   "��������1"
         Top             =   1440
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         Left            =   2625
         TabIndex        =   8
         Tag             =   "��������2"
         Top             =   1785
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         Left            =   3870
         TabIndex        =   9
         Tag             =   "��������2"
         Top             =   1785
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         Left            =   180
         Top             =   2145
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "�������� 3"
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
      Begin CSTextLibCtl.sidbEdit sdb_wid_3 
         Height          =   315
         Left            =   1395
         TabIndex        =   10
         Tag             =   "��������3"
         Top             =   2145
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   16711680
         BackColor       =   12648447
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
         MaxValue        =   0
         MinValue        =   99999
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   180
         Top             =   2490
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "�������� 4"
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
      Begin CSTextLibCtl.sidbEdit sdb_wid_4 
         Height          =   315
         Left            =   1395
         TabIndex        =   13
         Tag             =   "��������4"
         Top             =   2490
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   16711680
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         MaxValue        =   0
         MinValue        =   99999
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_cnt_3 
         Height          =   315
         Left            =   3870
         TabIndex        =   12
         Tag             =   "��������3"
         Top             =   2145
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         Left            =   2625
         TabIndex        =   11
         Tag             =   "��������3"
         Top             =   2145
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         Left            =   2625
         TabIndex        =   14
         Tag             =   "��������4"
         Top             =   2490
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         Left            =   3870
         TabIndex        =   15
         Tag             =   "��������4"
         Top             =   2490
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         Left            =   180
         Top             =   2850
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "�������� 5"
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
      Begin CSTextLibCtl.sidbEdit sdb_wid_5 
         Height          =   315
         Left            =   1395
         TabIndex        =   16
         Tag             =   "��������5"
         Top             =   2850
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   16711680
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         MaxValue        =   0
         MinValue        =   99999
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_len_5 
         Height          =   315
         Left            =   2625
         TabIndex        =   17
         Tag             =   "��������5"
         Top             =   2850
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         Left            =   3870
         TabIndex        =   18
         Tag             =   "��������5"
         Top             =   2850
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
         Left            =   4290
         TabIndex        =   3
         Tag             =   "�������"
         Top             =   555
         Width           =   915
         _Version        =   262145
         _ExtentX        =   1614
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   16711680
         BackColor       =   12648447
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
         MaxValue        =   0
         MinValue        =   999
         Undo            =   0
         Data            =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   135
         Width           =   105
      End
   End
   Begin Threed.SSCommand cmd_OK 
      Height          =   465
      Left            =   1545
      TabIndex        =   19
      Top             =   3600
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�_�J"
   End
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   465
      Left            =   3075
      TabIndex        =   22
      Top             =   3600
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ȡ��"
   End
End
Attribute VB_Name = "AKN2041C"
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
'-- Program Name      ���ֽ�������
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

Private Sub cbo_ccm_no_Click()

    Dim sQuery As String
    Dim Dynamic_Slab As String
    
    If cbo_ccm_no.Text = "" Then
    
        cbo_plan_name.Clear
        
    ElseIf cbo_ccm_no.Text = "1" Then
        
        Dynamic_Slab = "SC1"
        sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
        
        If Gf_CodeFind(M_CN1, sQuery) <> "Y" Then
            cbo_plan_name.Clear
            Exit Sub
        End If
    
    ElseIf cbo_ccm_no.Text = "2" Then
    
        Dynamic_Slab = "SC2"
        sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
        
        If Gf_CodeFind(M_CN1, sQuery) <> "Y" Then
            cbo_plan_name.Clear
            Exit Sub
        End If
    
    ElseIf cbo_ccm_no.Text = "3" Then
    
        Dynamic_Slab = "SC3"
        sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
        
        If Gf_CodeFind(M_CN1, sQuery) <> "Y" Then
            cbo_plan_name.Clear
            Exit Sub
        End If
        
    End If
    
    sQuery = " SELECT DISTINCT PLAN_NAME FROM EP_SLAB_INS WHERE PRC_STS = 'B' AND BOF_RSLT = 'Y' AND CCM_PRC_LINE = '" & cbo_ccm_no.Text & "' ORDER BY PLAN_NAME "
    Call Gf_ComboAdd(M_CN1, cbo_plan_name, sQuery)

    sdb_thk.Value = 0
    sdb_wid_1.Value = 0
    sdb_wid_2.Value = 0
    sdb_wid_3.Value = 0
    sdb_wid_4.Value = 0
    sdb_wid_5.Value = 0
    sdb_len_1.Value = 0
    sdb_len_2.Value = 0
    sdb_len_3.Value = 0
    sdb_len_4.Value = 0
    sdb_len_5.Value = 0
    sdb_cnt_1.Value = 0
    sdb_cnt_2.Value = 0
    sdb_cnt_3.Value = 0
    sdb_cnt_4.Value = 0
    sdb_cnt_5.Value = 0
    
End Sub

Private Sub cbo_plan_name_Click()

    Dim sQuery As String
    
    sQuery = " SELECT SLAB_THK FROM EP_SLAB_INS WHERE PLAN_NAME     = '" & cbo_plan_name.Text & "' "
    sQuery = sQuery + "                           AND PRC_STS       = 'B' "
    sQuery = sQuery + "                           AND SLAB_IN_PLAN  = (SELECT MAX(SLAB_IN_PLAN) FROM EP_SLAB_INS WHERE PLAN_NAME  = '" & cbo_plan_name.Text & "' AND PRC_STS = 'B') "
    sdb_thk.Value = Gf_FloatFind(M_CN1, sQuery)
    
    sQuery = " SELECT SLAB_WID FROM EP_SLAB_INS WHERE PLAN_NAME     = '" & cbo_plan_name.Text & "' "
    sQuery = sQuery + "                           AND PRC_STS       = 'B' "
    sQuery = sQuery + "                           AND SLAB_IN_PLAN  = (SELECT MAX(SLAB_IN_PLAN) FROM EP_SLAB_INS WHERE PLAN_NAME  = '" & cbo_plan_name.Text & "' AND PRC_STS = 'B') "
    
    sdb_wid_1.Value = Gf_FloatFind(M_CN1, sQuery)
    sdb_wid_2.Value = sdb_wid_1.Value
    sdb_wid_3.Value = sdb_wid_1.Value
    sdb_wid_4.Value = sdb_wid_1.Value
    sdb_wid_5.Value = sdb_wid_1.Value

End Sub

Private Sub Cmd_Cancel_Click()

    Unload Me
    
End Sub

Private Sub Cmd_Ok_Click()

    If Trim(cbo_mill_plt.Text) = "" Or (Trim(cbo_mill_plt.Text) <> "C1" And Trim(cbo_mill_plt.Text) <> "C3") Then
        Call Gp_MsgBoxDisplay(cbo_mill_plt.Tag & "��������", "", "������ʾ")
        Exit Sub
    End If
    
    If Trim(cbo_ccm_no.Text) = "" Then
        Call Gp_MsgBoxDisplay(cbo_ccm_no.Tag & "��������", "", "������ʾ")
        Exit Sub
    End If

    If Trim(cbo_plan_name.Text) = "" Then
        Call Gp_MsgBoxDisplay(cbo_plan_name.Tag & "��������", "", "������ʾ")
        Exit Sub
    End If
    
    If sdb_thk.Value = 0 Then
        Call Gp_MsgBoxDisplay(sdb_thk.Tag & "��������", "", "������ʾ")
        Exit Sub
    End If
       
    If sdb_len_1.Value + sdb_cnt_1.Value <> 0 Then
        If sdb_wid_1.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_wid_1.Tag & "��������", "", "������ʾ")
            Exit Sub
        ElseIf sdb_len_1.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_len_1.Tag & "��������", "", "������ʾ")
            Exit Sub
        
        ElseIf sdb_cnt_1.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_cnt_1.Tag & "��������", "", "������ʾ")
            Exit Sub
            
        End If
    End If
    
    If sdb_len_2.Value + sdb_cnt_2.Value <> 0 Then
        If sdb_wid_2.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_wid_2.Tag & "��������", "", "������ʾ")
            Exit Sub
        ElseIf sdb_len_2.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_len_2.Tag & "��������", "", "������ʾ")
            Exit Sub
        
        ElseIf sdb_cnt_2.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_cnt_2.Tag & "��������", "", "������ʾ")
            Exit Sub
            
        End If
    End If
    
    If sdb_len_3.Value + sdb_cnt_3.Value <> 0 Then
        If sdb_wid_3.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_wid_3.Tag & "��������", "", "������ʾ")
            Exit Sub
        ElseIf sdb_len_3.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_len_3.Tag & "��������", "", "������ʾ")
            Exit Sub
        
        ElseIf sdb_cnt_3.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_cnt_3.Tag & "��������", "", "������ʾ")
            Exit Sub
            
        End If
    End If
    
    If sdb_len_4.Value + sdb_cnt_4.Value <> 0 Then
        If sdb_wid_4.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_wid_4.Tag & "��������", "", "������ʾ")
            Exit Sub
        ElseIf sdb_len_4.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_len_4.Tag & "��������", "", "������ʾ")
            Exit Sub
        
        ElseIf sdb_cnt_4.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_cnt_4.Tag & "��������", "", "������ʾ")
            Exit Sub
            
        End If
    End If
    
    If sdb_len_5.Value + sdb_cnt_5.Value <> 0 Then
        If sdb_wid_5.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_wid_5.Tag & "��������", "", "������ʾ")
            Exit Sub
        ElseIf sdb_len_5.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_len_5.Tag & "��������", "", "������ʾ")
            Exit Sub
        
        ElseIf sdb_cnt_5.Value = 0 Then
            Call Gp_MsgBoxDisplay(sdb_cnt_5.Tag & "��������", "", "������ʾ")
            Exit Sub
            
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

    cbo_ccm_no.AddItem "1"
    cbo_ccm_no.AddItem "2"
    cbo_ccm_no.AddItem "3"
    cbo_ccm_no.Text = ""
    
    cbo_mill_plt.AddItem "C1"
    cbo_mill_plt.AddItem "C3"
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
                                 
    sQuery = "{call AFZ4610P ('B1'," & _
                             "'" & Trim(cbo_mill_plt.Text) & "'," & _
                             "'" & Trim(cbo_ccm_no.Text) & "'," & _
                             "'" & Trim(cbo_plan_name.Text) & "'," & _
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
        Call Gp_MsgBoxDisplay("�Ǽƻ��ı�������..!!", "I")
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        Call AKN2040C.Form_Ref
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