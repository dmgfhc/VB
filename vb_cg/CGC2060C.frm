VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGC2060C 
   Caption         =   "ĸ��ֶμ�ʵ������_CGC2060C"
   ClientHeight    =   9270
   ClientLeft      =   615
   ClientTop       =   1455
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   645
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   1138
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox txt_RollingSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   " "
         Top             =   150
         Width           =   2325
      End
      Begin VB.TextBox txt_Stlgrd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   " "
         Top             =   150
         Width           =   2385
      End
      Begin VB.TextBox txt_HotLevTmp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   " "
         Top             =   150
         Width           =   825
      End
      Begin VB.TextBox txt_RollingNo 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1980
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   22
         Top             =   150
         Width           =   1605
      End
      Begin VB.TextBox TXT_CB 
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
         Left            =   14115
         TabIndex        =   1
         Text            =   "CF"
         Top             =   180
         Visible         =   0   'False
         Width           =   675
      End
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   450
         Top             =   150
         Width           =   1485
         _ExtentX        =   2619
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
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   3750
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "���ƺ���"
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
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   7560
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "����"
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
         Left            =   11400
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "�Ƚ�ֱ���¶�"
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "��"
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
         Left            =   13590
         TabIndex        =   31
         Top             =   225
         Width           =   255
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   3135
      Left            =   90
      TabIndex        =   2
      Top             =   780
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   5530
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox txt_pdt_inf 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7230
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   64
         Text            =   "CGC2060C.frx":0000
         Top             =   2460
         Width           =   5745
      End
      Begin VB.TextBox TXT_SEQ 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "����"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4560
         MaxLength       =   5
         TabIndex        =   63
         Top             =   2130
         Width           =   1635
      End
      Begin VB.TextBox txt_CutYN6 
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
         Left            =   7140
         TabIndex        =   62
         Top             =   1650
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt_CutYN5 
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
         Left            =   7140
         TabIndex        =   61
         Top             =   1260
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt_CutYN4 
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
         Left            =   7140
         TabIndex        =   60
         Top             =   900
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt_CutYN3 
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
         Left            =   150
         TabIndex        =   59
         Top             =   1650
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt_CutYN2 
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
         Left            =   150
         TabIndex        =   58
         Top             =   1260
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt_CutYN1 
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
         Left            =   150
         TabIndex        =   57
         Top             =   900
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   7950
         TabIndex        =   55
         Top             =   1740
         Width           =   210
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   7950
         TabIndex        =   54
         Top             =   1350
         Width           =   210
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   7950
         TabIndex        =   53
         Top             =   960
         Width           =   210
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1170
         TabIndex        =   52
         Top             =   1740
         Width           =   210
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1170
         TabIndex        =   51
         Top             =   1350
         Width           =   210
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1170
         TabIndex        =   50
         Top             =   960
         Width           =   210
      End
      Begin VB.TextBox txt_Emp 
         Alignment       =   2  'Center
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
         Left            =   1950
         MaxLength       =   8
         TabIndex        =   37
         Top             =   2490
         Width           =   1215
      End
      Begin VB.TextBox txt_Group 
         Alignment       =   2  'Center
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
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   36
         Top             =   2490
         Width           =   735
      End
      Begin VB.TextBox txt_Shift 
         Alignment       =   2  'Center
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
         Left            =   450
         MaxLength       =   1
         TabIndex        =   35
         Top             =   2490
         Width           =   735
      End
      Begin VB.TextBox TXT_MOTHER_PLATE4 
         Alignment       =   2  'Center
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
         Left            =   8235
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "04"
         Top             =   900
         Width           =   400
      End
      Begin VB.TextBox TXT_MOTHER_PLATE6 
         Alignment       =   2  'Center
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
         Left            =   8235
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "06"
         Top             =   1680
         Width           =   400
      End
      Begin VB.TextBox TXT_MOTHER_PLATE2 
         Alignment       =   2  'Center
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
         Left            =   1455
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "02"
         Top             =   1290
         Width           =   400
      End
      Begin VB.TextBox TXT_MOTHER_PLATE3 
         Alignment       =   2  'Center
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
         Left            =   1455
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "03"
         Top             =   1680
         Width           =   400
      End
      Begin VB.TextBox TXT_MOTHER_PLATE5 
         Alignment       =   2  'Center
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
         Left            =   8235
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "05"
         Top             =   1290
         Width           =   400
      End
      Begin VB.TextBox TXT_MOTHER_PLATE1 
         Alignment       =   2  'Center
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
         Left            =   1455
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "01"
         Top             =   900
         Width           =   400
      End
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN1 
         Height          =   315
         Left            =   4170
         TabIndex        =   9
         Top             =   900
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN2 
         Height          =   315
         Left            =   4170
         TabIndex        =   10
         Top             =   1290
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN3 
         Height          =   315
         Left            =   4170
         TabIndex        =   11
         Top             =   1680
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN4 
         Height          =   315
         Left            =   10950
         TabIndex        =   12
         Top             =   900
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN5 
         Height          =   315
         Left            =   10950
         TabIndex        =   13
         Top             =   1290
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_PLATE_LEN6 
         Height          =   315
         Left            =   10950
         TabIndex        =   14
         Top             =   1680
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel51 
         Height          =   315
         Left            =   450
         Top             =   900
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "ĸ��1"
         Alignment       =   0
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
      Begin InDate.ULabel ULabel52 
         Height          =   315
         Left            =   450
         Top             =   1290
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "ĸ��2"
         Alignment       =   0
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
      Begin InDate.ULabel ULabel53 
         Height          =   315
         Left            =   450
         Top             =   1680
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "ĸ��3"
         Alignment       =   0
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
      Begin InDate.ULabel ULabel54 
         Height          =   315
         Left            =   7230
         Top             =   900
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "ĸ��4"
         Alignment       =   0
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
      Begin InDate.ULabel ULabel55 
         Height          =   315
         Left            =   7230
         Top             =   1290
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "ĸ��5"
         Alignment       =   0
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
      Begin InDate.ULabel ULabel56 
         Height          =   315
         Left            =   7230
         Top             =   1680
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "ĸ��6"
         Alignment       =   0
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
      Begin InDate.ULabel ULabel57 
         Height          =   315
         Left            =   3675
         Top             =   900
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Caption         =   "����"
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
      Begin InDate.ULabel ULabel58 
         Height          =   315
         Left            =   3675
         Top             =   1290
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Caption         =   "����"
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
      Begin InDate.ULabel ULabel59 
         Height          =   315
         Left            =   3675
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Caption         =   "����"
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
      Begin InDate.ULabel ULabel60 
         Height          =   315
         Left            =   10455
         Top             =   900
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Caption         =   "����"
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
      Begin InDate.ULabel ULabel61 
         Height          =   315
         Left            =   10455
         Top             =   1290
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Caption         =   "����"
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
      Begin InDate.ULabel ULabel62 
         Height          =   315
         Left            =   10455
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Caption         =   "����"
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
      Begin Threed.SSCommand cmd_Pass 
         Height          =   720
         Left            =   435
         TabIndex        =   15
         Top             =   150
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1270
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   14737632
         BackStyle       =   1
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   20.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�չ�"
      End
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN1 
         Height          =   315
         Left            =   5175
         TabIndex        =   16
         Top             =   900
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN2 
         Height          =   315
         Left            =   5175
         TabIndex        =   17
         Top             =   1290
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN3 
         Height          =   315
         Left            =   5175
         TabIndex        =   18
         Top             =   1680
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN4 
         Height          =   315
         Left            =   11955
         TabIndex        =   19
         Top             =   900
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN5 
         Height          =   315
         Left            =   11955
         TabIndex        =   20
         Top             =   1290
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_MOTHER_SCH_LEN6 
         Height          =   315
         Left            =   11955
         TabIndex        =   21
         Top             =   1680
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1764
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   4170
         Top             =   570
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "ʵ��"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   5190
         Top             =   570
         Width           =   985
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "ָʾ"
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
         Left            =   10950
         Top             =   570
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "ʵ��"
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
         Left            =   11970
         Top             =   570
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "ָʾ"
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
      Begin InDate.ULabel ULabel25 
         Height          =   315
         Left            =   450
         Top             =   2160
         Width           =   735
         _ExtentX        =   1296
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
      Begin InDate.ULabel ULabel35 
         Height          =   315
         Left            =   1200
         Top             =   2160
         Width           =   735
         _ExtentX        =   1296
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
      Begin InDate.ULabel ULabel36 
         Height          =   315
         Left            =   1950
         Top             =   2160
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "��ҵ��Ա"
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
      Begin CSTextLibCtl.sidbEdit txt_Thk1 
         Height          =   315
         Left            =   1860
         TabIndex        =   38
         Top             =   900
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_Wid1 
         Height          =   315
         Left            =   2760
         TabIndex        =   39
         Top             =   900
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_Thk4 
         Height          =   315
         Left            =   8640
         TabIndex        =   40
         Top             =   900
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_Wid4 
         Height          =   315
         Left            =   9540
         TabIndex        =   41
         Top             =   900
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_Thk2 
         Height          =   315
         Left            =   1860
         TabIndex        =   42
         Top             =   1290
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_Wid2 
         Height          =   315
         Left            =   2760
         TabIndex        =   43
         Top             =   1290
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_Thk5 
         Height          =   315
         Left            =   8640
         TabIndex        =   44
         Top             =   1290
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_Wid5 
         Height          =   315
         Left            =   9540
         TabIndex        =   45
         Top             =   1290
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_Thk3 
         Height          =   315
         Left            =   1860
         TabIndex        =   46
         Top             =   1680
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_Wid3 
         Height          =   315
         Left            =   2760
         TabIndex        =   47
         Top             =   1680
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_Thk6 
         Height          =   315
         Left            =   8640
         TabIndex        =   48
         Top             =   1680
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_Wid6 
         Height          =   315
         Left            =   9540
         TabIndex        =   49
         Top             =   1680
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Enabled         =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         ReadOnly        =   -1  'True
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   1860
         Top             =   570
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         Caption         =   "ʵ����"
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
         Left            =   2745
         Top             =   570
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         Caption         =   "ʵ����"
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
         Left            =   8640
         Top             =   570
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         Caption         =   "ʵ����"
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
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   8265
         Top             =   900
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         Caption         =   "ʵ����"
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
      Begin CSTextLibCtl.sitxEdit txt_CutDate 
         Height          =   315
         Left            =   3660
         TabIndex        =   56
         Tag             =   "�и�ʱ��"
         Top             =   180
         Width           =   2160
         _Version        =   262145
         _ExtentX        =   3810
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
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
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   1860
         Tag             =   "�и�ʱ��"
         Top             =   180
         Width           =   1770
         _ExtentX        =   3122
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
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   9540
         Top             =   570
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         Caption         =   "ʵ����"
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
      Begin InDate.ULabel ULabel12 
         Height          =   345
         Left            =   7230
         Top             =   2100
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   609
         Caption         =   "�����ڲ�Ʒ��Ϣ"
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
      Begin Threed.SSCommand cmd_scrap 
         Height          =   720
         Left            =   13170
         TabIndex        =   68
         Top             =   2100
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1270
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   14737632
         BackStyle       =   1
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   20.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�ϸ�"
      End
      Begin InDate.ULabel ULabel13 
         Height          =   675
         Left            =   3660
         Tag             =   "�и�ʱ��"
         Top             =   2130
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1191
         Caption         =   "�ֶκ�"
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
      Begin Threed.SSPanel SSP4 
         Height          =   315
         Left            =   8490
         TabIndex        =   70
         Top             =   120
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   255
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
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   315
         Left            =   9630
         TabIndex        =   71
         Top             =   120
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   16711680
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
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSP6 
         Height          =   315
         Left            =   10860
         TabIndex        =   72
         Top             =   120
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "���ڶ���"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSP5 
         Height          =   315
         Left            =   12030
         TabIndex        =   73
         Top             =   120
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   255
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "��������"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   5055
      Left            =   90
      TabIndex        =   23
      Top             =   4050
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabMaxWidth     =   3528
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
      TabCaption(0)   =   "�ȴ�ĸ��ֶ�"
      TabPicture(0)   =   "CGC2060C.frx":0002
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSSplitter1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ĸ��ֶ�ʵ����ѯ"
      TabPicture(1)   =   "CGC2060C.frx":001E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ss2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txt_RstToDate"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txt_RstFormDate"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin SSSplitter.SSSplitter SSSplitter1 
         Height          =   4515
         Left            =   90
         TabIndex        =   65
         Top             =   420
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   7964
         _Version        =   196609
         SplitterBarWidth=   3
         BorderStyle     =   1
         PaneTree        =   "CGC2060C.frx":003A
         Begin FPSpread.vaSpread ss1 
            Height          =   4485
            Left            =   15
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   15
            Width           =   12450
            _Version        =   393216
            _ExtentX        =   21960
            _ExtentY        =   7911
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
            MaxCols         =   29
            MaxRows         =   5
            ProcessTab      =   -1  'True
            Protect         =   0   'False
            SpreadDesigner  =   "CGC2060C.frx":008C
         End
         Begin FPSpread.vaSpread ss3 
            Height          =   4485
            Left            =   12525
            TabIndex        =   67
            Top             =   15
            Width           =   2340
            _Version        =   393216
            _ExtentX        =   4128
            _ExtentY        =   7911
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   6
            MaxRows         =   9
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CGC2060C.frx":0F81
         End
      End
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   -74910
         Top             =   390
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "��ֱʱ��"
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
      Begin CSTextLibCtl.sitxEdit SDT_PROD_DATE 
         Height          =   315
         Left            =   -73515
         TabIndex        =   24
         Top             =   390
         Width           =   1260
         _Version        =   262145
         _ExtentX        =   2222
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
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
         Mask            =   "____-__-__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin CSTextLibCtl.sitxEdit SDT_PROD_TO_DATE 
         Height          =   315
         Left            =   -72090
         TabIndex        =   25
         Top             =   390
         Width           =   1260
         _Version        =   262145
         _ExtentX        =   2222
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
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
         Mask            =   "____-__-__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   5130
         Left            =   -74910
         TabIndex        =   26
         Top             =   750
         Width           =   14910
         _Version        =   393216
         _ExtentX        =   26300
         _ExtentY        =   9049
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
         MaxCols         =   17
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGC2060C.frx":1455
      End
      Begin CSTextLibCtl.sitxEdit txt_RstFormDate 
         Height          =   315
         Left            =   -74910
         TabIndex        =   32
         Tag             =   "װ¯ʱ��"
         Top             =   360
         Width           =   1830
         _Version        =   262145
         _ExtentX        =   3228
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __:__"
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
         TabIndex        =   33
         Tag             =   "װ¯ʱ��"
         Top             =   360
         Width           =   1800
         _Version        =   262145
         _ExtentX        =   3175
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __:__"
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
      Begin FPSpread.vaSpread ss2 
         Height          =   4215
         Left            =   -74910
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   720
         Width           =   14910
         _Version        =   393216
         _ExtentX        =   26300
         _ExtentY        =   7435
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
         MaxCols         =   13
         MaxRows         =   20
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CGC2060C.frx":3402
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   120
         Left            =   -72240
         TabIndex        =   27
         Top             =   510
         Width           =   195
      End
   End
   Begin Threed.SSCommand cmd_Seq 
      Height          =   390
      Left            =   15330
      TabIndex        =   69
      Top             =   1860
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   688
      _Version        =   196609
      Font3D          =   3
      ForeColor       =   255
      BackColor       =   14737632
      BackStyle       =   1
      Enabled         =   0   'False
      ActiveColors    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "��ǰ�ֶκ�"
      Alignment       =   6
   End
End
Attribute VB_Name = "CGC2060C"
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
'-- Program Name      HOT SHEAR ��ҵʵ����ѯ���޸Ľ���
'-- Program ID        CGC2060C
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

Dim Mc1 As New Collection           'Master Collectionn
Dim Mc2 As New Collection           'Master Collectionn
Dim Mc3 As New Collection           'Master Collectionn
Dim Mc4 As New Collection           'Master Collectionn

Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2
Dim CUT_SEQ As String

Const SS1_URGNT_FL = 20  '����������ɫ�����ʾ add by liqian 2012-08-16
Const SS1_IMP_CONT = 21
Const SS1_PILECOOL = 22
Const SS1_FLAG = 23
Const SS1_EXPORT = 24
Const SS1_SLAB_NO = 1

Const SS2_DUILENG = 13
Const SS2_MPLATE_NO = 1


Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        
         Call Gp_Ms_Collection(txt_RollingNo, "p", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_RollingSize, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            Call Gp_Ms_Collection(txt_Stlgrd, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_HotLevTmp, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     
           Call Gp_Ms_Collection(txt_CutDate, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     
            Call Gp_Ms_Collection(txt_CutYN1, " ", " ", " ", "i", " ", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_Thk1, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_Wid1, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE1, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN1, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   
     
            Call Gp_Ms_Collection(txt_CutYN2, " ", " ", " ", "i", " ", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_Thk2, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_Wid2, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE2, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN2, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN2, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   
     
            Call Gp_Ms_Collection(txt_CutYN3, " ", " ", " ", "i", " ", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_Thk3, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_Wid3, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE3, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN3, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN3, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   
     
            Call Gp_Ms_Collection(txt_CutYN4, " ", " ", " ", "i", " ", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_Thk4, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_Wid4, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE4, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN4, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN4, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   
     
            Call Gp_Ms_Collection(txt_CutYN5, " ", " ", " ", "i", " ", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_Thk5, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_Wid5, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE5, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN5, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN5, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   
     
            Call Gp_Ms_Collection(txt_CutYN6, " ", " ", " ", "i", " ", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_Thk6, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_Wid6, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE6, " ", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN6, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN6, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   
             
             Call Gp_Ms_Collection(txt_Shift, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             Call Gp_Ms_Collection(txt_Group, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
               Call Gp_Ms_Collection(txt_Emp, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
           Call Gp_Ms_Collection(txt_pdt_inf, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
               Call Gp_Ms_Collection(TXT_SEQ, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
                                

    'MASTER Collection
     Mc1.Add Item:="CGC2060C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="CGC2060C.P_SEFER1", Key:="P-R"
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"
     
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)   '����������ɫ�����ʾ add by liqian 2012-08-16
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '�������
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '�������
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '�������
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '�������
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '�������
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       
     'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGC2060C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Call Gp_Ms_Collection(txt_RstFormDate, "p", "n", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(txt_RstToDate, "p", "n", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
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
   
     'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CGC2060C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   
     'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="CGC2060C.P_REFER3", Key:="P-R"
    sc3.Add Item:=pColumn3, Key:="pColumn"
    sc3.Add Item:=nColumn3, Key:="nColumn"
    sc3.Add Item:=aColumn3, Key:="aColumn"
    sc3.Add Item:=mColumn3, Key:="mColumn"
    sc3.Add Item:=iColumn3, Key:="iColumn"
    sc3.Add Item:=lColumn3, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        
         Call Gp_Ms_Collection(txt_RollingNo, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
       Call Gp_Ms_Collection(txt_RollingSize, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
            Call Gp_Ms_Collection(txt_Stlgrd, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
         Call Gp_Ms_Collection(txt_HotLevTmp, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     
           Call Gp_Ms_Collection(txt_CutDate, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     
            Call Gp_Ms_Collection(txt_CutYN1, " ", " ", " ", "i", " ", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(txt_Thk1, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(txt_Wid1, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE1, " ", " ", " ", " ", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN1, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN1, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     
            Call Gp_Ms_Collection(txt_CutYN2, " ", " ", " ", "i", " ", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(txt_Thk2, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(txt_Wid2, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE2, " ", " ", " ", " ", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN2, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN2, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     
            Call Gp_Ms_Collection(txt_CutYN3, " ", " ", " ", "i", " ", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(txt_Thk3, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(txt_Wid3, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE3, " ", " ", " ", " ", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN3, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN3, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     
            Call Gp_Ms_Collection(txt_CutYN4, " ", " ", " ", "i", " ", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(txt_Thk4, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(txt_Wid4, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE4, " ", " ", " ", " ", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN4, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN4, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     
            Call Gp_Ms_Collection(txt_CutYN5, " ", " ", " ", "i", " ", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(txt_Thk5, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(txt_Wid5, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE5, " ", " ", " ", " ", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN5, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN5, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     
            Call Gp_Ms_Collection(txt_CutYN6, " ", " ", " ", "i", " ", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(txt_Thk6, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(txt_Wid6, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     Call Gp_Ms_Collection(TXT_MOTHER_PLATE6, " ", " ", " ", " ", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
 Call Gp_Ms_Collection(SDB_MOTHER_PLATE_LEN6, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
   Call Gp_Ms_Collection(SDB_MOTHER_SCH_LEN6, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
             
             Call Gp_Ms_Collection(txt_Shift, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
             Call Gp_Ms_Collection(txt_Group, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
               Call Gp_Ms_Collection(txt_Emp, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
                                

    'MASTER Collection
     Mc3.Add Item:="CGC2060C.P_MODIFY", Key:="P-M"
     Mc3.Add Item:="CGC2060C.P_SEFER2", Key:="P-R"
     Mc3.Add Item:=pControl3, Key:="pControl"
     Mc3.Add Item:=nControl3, Key:="nControl"
     Mc3.Add Item:=mControl3, Key:="mControl"
     Mc3.Add Item:=iControl3, Key:="iControl"
     Mc3.Add Item:=rControl3, Key:="rControl"
     Mc3.Add Item:=cControl3, Key:="cControl"
     Mc3.Add Item:=aControl3, Key:="aControl"
     Mc3.Add Item:=lControl3, Key:="lControl"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    CUT_SEQ = "SELECT NVL(SEQ_NO,0) FROM NISCO.GP_MP_IDX WHERE PLT='C3'"

End Sub

Private Sub Check1_Click()
    If SDB_MOTHER_PLATE_LEN1 = "" Then Check1.Value = 0
    If Check1.Value = 1 Then
        txt_CutYN1 = "Y"
    Else
        txt_CutYN1 = ""
    End If
End Sub

Private Sub Check2_Click()
    If SDB_MOTHER_PLATE_LEN2 = "" Then Check2.Value = 0
    If Check2.Value = 1 Then
        txt_CutYN2 = "Y"
    Else
        txt_CutYN2 = ""
    End If
End Sub

Private Sub Check3_Click()
    
    If SDB_MOTHER_PLATE_LEN3 = "" Then Check3.Value = 0
    If Check3.Value = 1 Then
        txt_CutYN3 = "Y"
    Else
        txt_CutYN3 = ""
    End If
End Sub

Private Sub Check4_Click()
    If SDB_MOTHER_PLATE_LEN4 = "" Then Check4.Value = 0
    If Check4.Value = 1 Then
        txt_CutYN4 = "Y"
    Else
        txt_CutYN4 = ""
    End If
End Sub

Private Sub Check5_Click()
    If SDB_MOTHER_PLATE_LEN5 = "" Then Check5.Value = 0
    If Check5.Value = 1 Then
        txt_CutYN5 = "Y"
    Else
        txt_CutYN5 = ""
    End If
End Sub

Private Sub Check6_Click()
    If SDB_MOTHER_PLATE_LEN6 = "" Then Check6.Value = 0
    If Check6.Value = 1 Then
        txt_CutYN6 = "Y"
    Else
        txt_CutYN6 = ""
    End If
End Sub

Private Sub cmd_Pass_Click()
Dim CNT As Long
Dim sMesg As String

    If Mid(txt_CutDate, 1, 1) <> "2" Then
         sMesg = " �������и�ʱ��...��"
         Call Gp_MsgBoxDisplay(sMesg)
         Screen.MousePointer = DEFAULT
         Exit Sub
    End If

    CNT = 0
    If Check1.Value = ssCBChecked Then
       CNT = CNT + 1
    End If
    
    If Check2.Value = ssCBChecked Then
       CNT = CNT + 1
    End If
    
    If Check3.Value = ssCBChecked Then
       CNT = CNT + 1
    End If
    
    If Check4.Value = ssCBChecked Then
       CNT = CNT + 1
    End If
    
    If Check5.Value = ssCBChecked Then
       CNT = CNT + 1
    End If
    
    If Check6.Value = ssCBChecked Then
       CNT = CNT + 1
    End If

    If CNT > 1 Then
        sMesg = "һ������ĸ�岻�ܿչ�.......��"
        Call Gp_MsgBoxDisplay(sMesg)
         Screen.MousePointer = DEFAULT
        Exit Sub
    Else
        Call Form_Pro
    End If
    
    txt_CutDate = ""
    
    
End Sub

Private Sub cmd_scrap_Click()

If Not Gf_MessConfirm("��ȷ��Ҫ��ĸ��� " & txt_RollingNo.Text & " ���ϸִ�����", "W", "") Then
   Exit Sub
End If

    Dim OutParam(2, 4) As Variant
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    Dim sMesg As String
    
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
    
    sQuery = "{call CGC2060C.P_SCRAP('" & Trim(txt_RollingNo.Text) & "','" & txt_Shift & "','" & txt_Group & "','" & sUserID & "','" & TXT_CB & "',?,?)}"
    
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
        Call Gp_MsgBoxDisplay("ʵ������ʧ�ܣ���ȷ��=> " & adoCmd("arg_loaction2"))
    Else
        Call Gp_MsgBoxDisplay("�ϸִ����ɹ�=> " & Trim(txt_RollingNo.Text), "I", "ϵͳ��ʾ��Ϣ")
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    txt_CutDate = ""
    
End Sub

Private Sub cmd_Seq_Click()

If Not Gf_MessConfirm("��ȷ��Ҫ��ʼ����ǰ�ֶκ���", "W", "") Then
   TXT_SEQ = Gf_FloatFind(M_CN1, CUT_SEQ)
   Exit Sub
End If
    
    Dim OutParam(2, 4) As Variant
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    Dim sMesg As String
        
    On Error Resume Next
    
    If Not Gf_Sc_Authority(sAuthority, "U") Then
         sMesg = " ��û��Ȩ�޲��������� ��"
         Call Gp_MsgBoxDisplay(sMesg)
         Exit Sub
    End If

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
    
    If CInt(TXT_SEQ) < 0 Then
         sMesg = " ������������� ��"
         Call Gp_MsgBoxDisplay(sMesg)
         Exit Sub
    End If
    
    sQuery = "{call CGC2060C.P_CreSeq(" & CInt(TXT_SEQ) & ",?,?)}"
    
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
        Call Gp_MsgBoxDisplay("ʵ������ʧ�ܣ���ȷ��=> " & adoCmd("arg_loaction2"))
    End If
    
    Set adoCmd = Nothing
    
    Call Form_Ref
    
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

    Call Gp_Mill_ControlLock(Mc1("lControl"), True)
    Call Gp_Mill_ControlLock(Mc2("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    Call Gp_Sp_Setting(sc3.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(sc3)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc3.Item("Spread"), "CG-System.INI", Me.Name)
    
    Tab1.Tab = 0
    Call Form_Ref
'    Call ss1_DblClick(1, 1)
    
    txt_Shift = Gf_ShiftSet3(M_CN1)
    txt_Group = Gf_GroupSet(M_CN1, Trim(txt_Shift), Gf_DTSet(M_CN1, , "X"))
    txt_Emp = sUserID
    If Mid(sAuthority, 1, 3) = "111" Then
       cmd_Pass.Enabled = True
       cmd_Seq.Enabled = True
       cmd_scrap.Enabled = True
    Else
       cmd_Pass.Enabled = False
       cmd_Seq.Enabled = False
       cmd_scrap.Enabled = False
    End If

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc3.Item("Spread"), "CG-System.INI", Me.Name)

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
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set Mc4 = Nothing
    
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set sc3 = Nothing
    
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(sc3)
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = False                'Excel
    End With
    
End Sub


Public Sub Form_Ref()
    Dim iRow As Integer
    Dim iCol As Integer
    Dim sUrgnt_Fl As String
    Dim simpcont  As String
    Dim PILECOOL  As String
    Dim sFlag  As String
    Dim sexport  As String
    
    Dim sDuileng  As String

    If Tab1.Tab = 0 Then
        Call Gf_Sp_Refer(M_CN1, sc1, , , , False)
        ss1.OperationMode = OperationModeNormal
'        TXT_SEQ = Gf_FloatFind(M_CN1, CUT_SEQ)
        Call ss1_DblClick(1, 1)
        
         '����������ɫ��ʾ add by liqian 2012-08-16
         With ss1
              For iRow = 1 To .MaxRows
                 .ROW = iRow:
                  .Col = SS1_URGNT_FL:    sUrgnt_Fl = Trim(.Text)
                  .Col = SS1_IMP_CONT:    simpcont = Trim(.Text)
                  .Col = SS1_PILECOOL:    PILECOOL = Trim(.Text)
                  .Col = SS1_FLAG:        sFlag = Trim(.Text)
                  .Col = SS1_EXPORT:      sexport = Trim(.Text)
                  
                  If sUrgnt_Fl = "Y" Then
                     Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, &HC000&)
                  End If
                  If simpcont = "Y" Then
                    Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, SSP4.BackColor)
                  End If
                  If PILECOOL = "Y" And simpcont <> "Y" Then
                    Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, vbBlack, vbYellow)
                  End If
                  
                  '�Ƿ�������
                   
                    If sFlag = "Y" Then
                       Call Gp_Sp_BlockColor(ss1, SS1_SLAB_NO, SS1_SLAB_NO, iRow, iRow, SSP5.BackColor)
                    End If
                  '�Ƿ���ڶ���
                    If sexport = "Y" Then
                       Call Gp_Sp_BlockColor(ss1, SS1_SLAB_NO, SS1_SLAB_NO, iRow, iRow, SSP6.BackColor)
                    End If
              Next iRow
        End With
        
        With ss2
              For iRow = 1 To .MaxRows
                 .ROW = iRow:
                 .Col = SS2_DUILENG:    sDuileng = Trim(.Text)
                  
                    If sDuileng = "Y" Then
                       Call Gp_Sp_BlockColor(ss2, SS2_MPLATE_NO, SS2_MPLATE_NO, iRow, iRow, SSPanel1.BackColor)
                    End If
              Next iRow
        End With
             
    ElseIf Tab1.Tab = 1 Then
        Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl"), False)
        ss2.OperationMode = OperationModeNormal
'        TXT_SEQ = Gf_FloatFind(M_CN1, CUT_SEQ)

    End If

    
End Sub

Public Sub Form_Pro()
    Dim sMesg As String

    If Not Gp_DateCheck(txt_CutDate) Then
            sMesg = " ����ȷ�����и�ʱ�� ��"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
    End If
    
    Call Gf_Ms_Process(M_CN1, Mc1, sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Form_Ref
    
    txt_Shift = Gf_ShiftSet3(M_CN1)
    txt_Group = Gf_GroupSet(M_CN1, Trim(txt_Shift), Gf_DTSet(M_CN1, , "X"))
    txt_Emp = sUserID
    
    txt_CutDate.RawData = ""
   
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub

Private Sub SDB_MOTHER_PLATE_LEN1_Change()
    If Len(SDB_MOTHER_PLATE_LEN1) > 0 Then
       Check1.Value = 1
       txt_CutYN1 = "Y"
    Else
       Check1.Value = 0
       txt_CutYN1 = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN1_DblClick()
    If SDB_MOTHER_PLATE_LEN1.Text = "" Then
        SDB_MOTHER_PLATE_LEN1.Text = SDB_MOTHER_SCH_LEN1.Text
    Else
        SDB_MOTHER_PLATE_LEN1.Text = ""
    End If
End Sub


Private Sub SDB_MOTHER_PLATE_LEN2_Change()
    If Len(SDB_MOTHER_PLATE_LEN2) > 0 Then
       Check2.Value = 1
       txt_CutYN2 = "Y"
    Else
       Check2.Value = 0
       txt_CutYN2 = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN2_DblClick()
    If SDB_MOTHER_PLATE_LEN2.Text = "" Then
        SDB_MOTHER_PLATE_LEN2.Text = SDB_MOTHER_SCH_LEN2.Text
    Else
        SDB_MOTHER_PLATE_LEN2.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN3_Change()
    If Len(SDB_MOTHER_PLATE_LEN3) > 0 Then
       Check3.Value = 1
       txt_CutYN3 = "Y"
    Else
       Check3.Value = 0
       txt_CutYN3 = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN3_DblClick()
    If SDB_MOTHER_PLATE_LEN3.Text = "" Then
        SDB_MOTHER_PLATE_LEN3.Text = SDB_MOTHER_SCH_LEN3.Text
    Else
        SDB_MOTHER_PLATE_LEN3.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN4_Change()
    If Len(SDB_MOTHER_PLATE_LEN4) > 0 Then
       Check4.Value = 1
       txt_CutYN4 = "Y"
    Else
       Check4.Value = 0
       txt_CutYN4 = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN4_DblClick()
    If SDB_MOTHER_PLATE_LEN4.Text = "" Then
        SDB_MOTHER_PLATE_LEN4.Text = SDB_MOTHER_SCH_LEN4.Text
    Else
        SDB_MOTHER_PLATE_LEN4.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN5_Change()
    If Len(SDB_MOTHER_PLATE_LEN5) > 0 Then
       Check5.Value = 1
       txt_CutYN5 = "Y"
    Else
       Check5.Value = 0
       txt_CutYN5 = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN5_DblClick()
    If SDB_MOTHER_PLATE_LEN5.Text = "" Then
        SDB_MOTHER_PLATE_LEN5.Text = SDB_MOTHER_SCH_LEN5.Text
    Else
        SDB_MOTHER_PLATE_LEN5.Text = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN6_Change()
    If Len(SDB_MOTHER_PLATE_LEN6) > 0 Then
       Check6.Value = 1
       txt_CutYN6 = "Y"
    Else
       Check6.Value = 0
       txt_CutYN6 = ""
    End If
End Sub

Private Sub SDB_MOTHER_PLATE_LEN6_DblClick()
    If SDB_MOTHER_PLATE_LEN6.Text = "" Then
        SDB_MOTHER_PLATE_LEN6.Text = SDB_MOTHER_SCH_LEN6.Text
    Else
        SDB_MOTHER_PLATE_LEN6.Text = ""
    End If
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        ss1.ROW = ROW
        ss1.Col = 1
        txt_RollingNo.Text = ss1.Text
        
        If Trim(txt_RollingNo.Text) <> "" Then
        
            Call Gf_Ms_Refer(M_CN1, Mc1, , , False)
            Call Gf_Sp_Refer(M_CN1, sc3, Mc1, , , False)
            ss3.OperationMode = OperationModeNormal
            
            If CDbl(SDB_MOTHER_PLATE_LEN1.Value) < 500 Then
               txt_Thk1 = ""
               txt_Wid1 = ""
               Check1.Value = ssCBUnchecked
            Else
               Check1.Value = ssCBChecked
            End If
            
            If CDbl(SDB_MOTHER_PLATE_LEN2.Value) < 500 Then
               txt_Thk2 = ""
               txt_Wid2 = ""
               Check2.Value = ssCBUnchecked
            Else
               Check2.Value = ssCBChecked
            End If
            
            If CDbl(SDB_MOTHER_PLATE_LEN3.Value) < 500 Then
               txt_Thk3 = ""
               txt_Wid3 = ""
               Check3.Value = ssCBUnchecked
            Else
               Check3.Value = ssCBChecked
            End If
            
            If CDbl(SDB_MOTHER_PLATE_LEN4.Value) < 500 Then
               txt_Thk4 = ""
               txt_Wid4 = ""
               Check4.Value = ssCBUnchecked
            Else
               Check4.Value = ssCBChecked
            End If
            
            If CDbl(SDB_MOTHER_PLATE_LEN5.Value) < 500 Then
               txt_Thk5 = ""
               txt_Wid5 = ""
               Check5.Value = ssCBUnchecked
            Else
               Check5.Value = ssCBChecked
            End If
            
            If CDbl(SDB_MOTHER_PLATE_LEN6.Value) < 500 Then
               txt_Thk6 = ""
               txt_Wid6 = ""
               Check6.Value = ssCBUnchecked
            Else
               Check6.Value = ssCBChecked
            End If
            
            txt_CutDate.RawData = ""
        End If
        
        txt_Shift = Gf_ShiftSet3(M_CN1)
        txt_Group = Gf_GroupSet(M_CN1, Trim(txt_Shift), Gf_DTSet(M_CN1, , "X"))
        txt_Emp = sUserID
        
        'txt_CutDate.RawData = Gf_DTSet(M_CN1, , "X")
    
    End If
End Sub


Private Sub ss2_DblClick(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        ss2.ROW = ROW
        ss2.Col = 1
        txt_RollingNo.Text = ss2.Text
        If Trim(txt_RollingNo.Text) <> "" Then
            Call Gf_Ms_Refer(M_CN1, Mc3, , , False)
        End If
    End If
End Sub

Private Sub tab1_Click(PreviousTab As Integer)
    If Tab1.Tab = "1" Then
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
    End If
End Sub

Private Sub txt_CutDate_Click()
    txt_CutDate.RawData = Gf_DTSet(M_CN1, , "X")
End Sub

Private Sub txt_RstFormDate_DblClick()
    txt_RstFormDate.RawData = Gf_DTSet(M_CN1, , "X")
    txt_RstToDate.RawData = Gf_DTSet(M_CN1, , "X")
End Sub