VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form EGA1080C 
   Caption         =   "������ʵ����ѯ���޸�_EGA1080C"
   ClientHeight    =   9405
   ClientLeft      =   585
   ClientTop       =   1680
   ClientWidth     =   15150
   BeginProperty Font 
      Name            =   "����"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame sf1 
      Height          =   4665
      Left            =   540
      TabIndex        =   21
      Top             =   9330
      Visible         =   0   'False
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   8229
      _Version        =   196609
      Font3D          =   2
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " ����ȱ��"
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "β��"
         Height          =   240
         Index           =   8
         Left            =   3330
         TabIndex        =   102
         Tag             =   "B"
         Top             =   1815
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�в�"
         Height          =   195
         Index           =   7
         Left            =   3330
         TabIndex        =   101
         Tag             =   "M"
         Top             =   1590
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ͷ��"
         Height          =   195
         Index           =   6
         Left            =   3330
         TabIndex        =   100
         Tag             =   "T"
         Top             =   1365
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   3330
         TabIndex        =   99
         Text            =   " "
         Top             =   1020
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   7
         Left            =   3330
         TabIndex        =   98
         Top             =   690
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Height          =   315
         Index           =   5
         Left            =   705
         TabIndex        =   64
         Top             =   555
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Height          =   315
         Index           =   4
         Left            =   390
         TabIndex        =   63
         Top             =   555
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Height          =   315
         Index           =   7
         Left            =   705
         TabIndex        =   62
         Top             =   225
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "β��"
         Height          =   240
         Index           =   17
         Left            =   3345
         TabIndex        =   51
         Tag             =   "B"
         Top             =   3825
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�в�"
         Height          =   195
         Index           =   16
         Left            =   3345
         TabIndex        =   50
         Tag             =   "M"
         Top             =   3600
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ͷ��"
         Height          =   195
         Index           =   15
         Left            =   3345
         TabIndex        =   49
         Tag             =   "T"
         Top             =   3375
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "β��"
         Height          =   240
         Index           =   14
         Left            =   2370
         TabIndex        =   48
         Tag             =   "B"
         Top             =   3825
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�в�"
         Height          =   195
         Index           =   13
         Left            =   2370
         TabIndex        =   47
         Tag             =   "M"
         Top             =   3600
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ͷ��"
         Height          =   195
         Index           =   12
         Left            =   2370
         TabIndex        =   46
         Tag             =   "T"
         Top             =   3375
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "β��"
         Height          =   240
         Index           =   11
         Left            =   1410
         TabIndex        =   45
         Tag             =   "B"
         Top             =   3825
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�в�"
         Height          =   195
         Index           =   10
         Left            =   1410
         TabIndex        =   44
         Tag             =   "M"
         Top             =   3600
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ͷ��"
         Height          =   195
         Index           =   9
         Left            =   1410
         TabIndex        =   43
         Tag             =   "T"
         Top             =   3375
         Width           =   810
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   1410
         TabIndex        =   42
         Text            =   " "
         Top             =   3030
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   2370
         TabIndex        =   41
         Text            =   " "
         Top             =   3030
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   3345
         TabIndex        =   40
         Text            =   " "
         Top             =   3030
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   4
         Left            =   2370
         TabIndex        =   39
         Top             =   2700
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   5
         Left            =   3345
         TabIndex        =   38
         Top             =   2700
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   1
         Left            =   2370
         TabIndex        =   7
         Top             =   690
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   2385
         TabIndex        =   30
         Text            =   " "
         Top             =   1020
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   1410
         TabIndex        =   29
         Text            =   " "
         Top             =   1020
         Width           =   960
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ͷ��"
         Height          =   195
         Index           =   0
         Left            =   1410
         TabIndex        =   28
         Tag             =   "T"
         Top             =   1365
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�в�"
         Height          =   195
         Index           =   1
         Left            =   1410
         TabIndex        =   27
         Tag             =   "M"
         Top             =   1590
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "β��"
         Height          =   240
         Index           =   2
         Left            =   1410
         TabIndex        =   26
         Tag             =   "B"
         Top             =   1815
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ͷ��"
         Height          =   195
         Index           =   3
         Left            =   2370
         TabIndex        =   25
         Tag             =   "T"
         Top             =   1365
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�в�"
         Height          =   195
         Index           =   4
         Left            =   2370
         TabIndex        =   24
         Tag             =   "M"
         Top             =   1590
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "β��"
         Height          =   240
         Index           =   5
         Left            =   2370
         TabIndex        =   23
         Tag             =   "B"
         Top             =   1815
         Width           =   810
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   0
         Left            =   1410
         TabIndex        =   8
         Top             =   2070
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         NumDecDigits    =   1
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   1
         Left            =   2370
         TabIndex        =   9
         Top             =   2070
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         NumDecDigits    =   1
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   225
         Top             =   1020
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "ȱ�ݲ�λ"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Left            =   225
         Top             =   2070
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "ȱ�ݳߴ�"
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   3
         Left            =   1410
         TabIndex        =   52
         Top             =   4095
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         NumDecDigits    =   1
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   4
         Left            =   2370
         TabIndex        =   53
         Top             =   4095
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         NumDecDigits    =   1
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   5
         Left            =   3345
         TabIndex        =   54
         Top             =   4095
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         NumDecDigits    =   1
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   225
         Top             =   2700
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "�ϱ���"
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
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   225
         Top             =   3030
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "ȱ�ݲ�λ"
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
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   225
         Top             =   4080
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "ȱ�ݳߴ�"
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
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   1410
         Top             =   360
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         Caption         =   "��Ҫȱ��"
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   2370
         Top             =   360
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         Caption         =   "Сȱ��1"
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   2
         Left            =   3330
         TabIndex        =   103
         Top             =   2070
         Visible         =   0   'False
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         NumDecDigits    =   1
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   3330
         Top             =   360
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         Caption         =   "Сȱ��2"
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
   Begin Threed.SSFrame SSFrame2 
      Height          =   2145
      Left            =   5070
      TabIndex        =   114
      Top             =   4035
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   3784
      _Version        =   196609
      Font3D          =   2
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " ��ĥ"
      Begin VB.CheckBox CHK_BOT_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�ϸ�"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   3765
         TabIndex        =   122
         Tag             =   "Y"
         Top             =   990
         Width           =   735
      End
      Begin VB.TextBox TXT_BOT_GRID_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1290
         MaxLength       =   1
         TabIndex        =   121
         Text            =   " "
         Top             =   980
         Width           =   690
      End
      Begin VB.TextBox TXT_TOP_GRID_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1290
         MaxLength       =   1
         TabIndex        =   120
         Text            =   " "
         Top             =   600
         Width           =   690
      End
      Begin VB.CheckBox CHK_GRID_FLAG 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�Ƿ���ĥ"
         Height          =   240
         Left            =   165
         TabIndex        =   119
         Tag             =   "G"
         Top             =   300
         Width           =   1110
      End
      Begin VB.TextBox TXT_GRID_EMP_CD 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1290
         MaxLength       =   7
         TabIndex        =   118
         Tag             =   "��ҵ��Ա"
         Top             =   1360
         Width           =   1035
      End
      Begin VB.CheckBox CHK_TOP_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "���ϸ�"
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   3765
         TabIndex        =   117
         Tag             =   "N"
         Top             =   675
         Width           =   900
      End
      Begin VB.CheckBox CHK_TOP_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�ϸ�"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   3765
         TabIndex        =   116
         Tag             =   "Y"
         Top             =   420
         Width           =   735
      End
      Begin VB.CheckBox CHK_BOT_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "���ϸ�"
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   3765
         TabIndex        =   115
         Tag             =   "N"
         Top             =   1260
         Width           =   900
      End
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   165
         Top             =   1380
         Width           =   1095
         _ExtentX        =   1931
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
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Index           =   0
         Left            =   165
         Top             =   990
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "�±���"
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
         Index           =   2
         Left            =   165
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "�ϱ���"
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
         Index           =   1
         Left            =   1290
         Top             =   240
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   556
         Caption         =   "�ж�/ �����%/ ���"
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
      Begin CSTextLibCtl.sidbEdit SDB_TOP_GRID_DEEP 
         Height          =   315
         Left            =   2880
         TabIndex        =   123
         Top             =   600
         Width           =   840
         _Version        =   262145
         _ExtentX        =   1482
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_TOP_GRID_YRD 
         Height          =   315
         Left            =   2010
         TabIndex        =   124
         Top             =   600
         Width           =   840
         _Version        =   262145
         _ExtentX        =   1482
         _ExtentY        =   556
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumDecDigits    =   2
         NumIntDigits    =   3
         ShowZero        =   0   'False
         MaxValue        =   999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_BOT_GRID_YRD 
         Height          =   315
         Left            =   2010
         TabIndex        =   125
         Top             =   975
         Width           =   840
         _Version        =   262145
         _ExtentX        =   1482
         _ExtentY        =   556
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         Left            =   165
         Top             =   1755
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "��ĥʱ��"
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
      Begin CSTextLibCtl.sitxEdit TXT_GRID_TIME 
         Height          =   315
         Left            =   1290
         TabIndex        =   126
         Top             =   1740
         Width           =   2085
         _Version        =   262145
         _ExtentX        =   3678
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
      Begin CSTextLibCtl.sidbEdit SDB_BOT_GRID_DEEP 
         Height          =   315
         Left            =   2880
         TabIndex        =   127
         Top             =   975
         Width           =   840
         _Version        =   262145
         _ExtentX        =   1482
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
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1935
      Left            =   5070
      TabIndex        =   110
      Top             =   7095
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   3413
      _Version        =   196609
      Font3D          =   2
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox TXT_INSP_MAN_TAIL 
         Height          =   330
         Left            =   3750
         MaxLength       =   7
         TabIndex        =   150
         Tag             =   "���Ա"
         Top             =   60
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_MAN 
         Height          =   330
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   113
         Tag             =   "���Ա"
         Top             =   60
         Width           =   960
      End
      Begin VB.TextBox TXT_EMP_CD1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1215
         MaxLength       =   7
         TabIndex        =   112
         Tag             =   "��ҵ��Ա"
         Top             =   810
         Visible         =   0   'False
         Width           =   1035
      End
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   90
         Top             =   435
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "���ʱ��"
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
      Begin CSTextLibCtl.sitxEdit TXT_INSP_OCCR_TIME 
         Height          =   315
         Left            =   1320
         TabIndex        =   111
         Tag             =   "���ʱ��"
         Top             =   435
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   90
         Top             =   60
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "ͷ�����鹤"
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
      Begin InDate.ULabel ULabel26 
         Height          =   315
         Left            =   90
         Top             =   825
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
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
      Begin InDate.ULabel ULabel47 
         Height          =   315
         Left            =   2520
         Top             =   60
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "β�����鹤"
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
   End
   Begin Threed.SSFrame Single 
      Height          =   945
      Left            =   90
      TabIndex        =   20
      Top             =   60
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   1667
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox TXT_PLATE_NO 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1425
         MaxLength       =   14
         TabIndex        =   0
         Top             =   105
         Width           =   2010
      End
      Begin VB.TextBox TXT_STLGRD 
         Height          =   285
         Left            =   4275
         TabIndex        =   90
         Top             =   120
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox TXT_APLY_ENDUSE_CD 
         Height          =   285
         Left            =   4065
         TabIndex        =   89
         Top             =   105
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox TXT_PROC_FLAG 
         Height          =   270
         Left            =   3855
         TabIndex        =   88
         Top             =   105
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox TXT_UST_FLAG 
         Height          =   270
         Left            =   3645
         TabIndex        =   87
         Top             =   105
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox txt_stdspec_chg_ref 
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
         Left            =   1425
         MaxLength       =   18
         TabIndex        =   1
         Tag             =   "��׼��"
         Top             =   525
         Width           =   2925
      End
      Begin VB.ComboBox CBO_SHIFT 
         Height          =   315
         ItemData        =   "EGA1080C.frx":0000
         Left            =   7680
         List            =   "EGA1080C.frx":000D
         TabIndex        =   4
         Top             =   525
         Width           =   1005
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   6465
         Top             =   105
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "����ʱ��"
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
      Begin CSTextLibCtl.sitxEdit SDT_PROD_DATE 
         Height          =   315
         Left            =   7680
         TabIndex        =   2
         Top             =   105
         Width           =   1200
         _Version        =   262145
         _ExtentX        =   2117
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
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   6465
         Top             =   525
         Width           =   1185
         _ExtentX        =   2090
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
      Begin InDate.ULabel ULabel22 
         Height          =   300
         Index           =   4
         Left            =   210
         Top             =   525
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         Caption         =   "��׼��"
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
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   11925
         Top             =   105
         Width           =   1185
         _ExtentX        =   2090
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
      Begin CSTextLibCtl.sidbEdit SDB_THK_REF 
         Height          =   315
         Left            =   13140
         TabIndex        =   5
         Top             =   105
         Width           =   1065
         _Version        =   262145
         _ExtentX        =   1879
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
      Begin InDate.ULabel ULabel24 
         Height          =   315
         Left            =   11925
         Top             =   525
         Width           =   1185
         _ExtentX        =   2090
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
      Begin CSTextLibCtl.sidbEdit SDB_WID_REF 
         Height          =   315
         Left            =   13140
         TabIndex        =   6
         Top             =   525
         Width           =   1065
         _Version        =   262145
         _ExtentX        =   1879
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
      Begin CSTextLibCtl.sitxEdit SDT_PROD_TO_DATE 
         Height          =   315
         Left            =   9060
         TabIndex        =   3
         Top             =   105
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
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   210
         Top             =   105
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "�ְ��"
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
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   120
         Left            =   8910
         TabIndex        =   75
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.TextBox txt_ResonCd 
      Height          =   285
      Left            =   16140
      TabIndex        =   105
      Text            =   " "
      Top             =   600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.ComboBox cbo_ResonDesc 
      Height          =   315
      ItemData        =   "EGA1080C.frx":001A
      Left            =   16380
      List            =   "EGA1080C.frx":001C
      TabIndex        =   104
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TXT_INSP_FLAW 
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   61
      Top             =   10005
      Visible         =   0   'False
      Width           =   285
   End
   Begin InDate.ULabel ULabel1 
      Height          =   330
      Left            =   10410
      Top             =   9690
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin Threed.SSFrame SSFrame5 
      Height          =   705
      Left            =   11610
      TabIndex        =   91
      Top             =   9690
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1244
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.CheckBox chkGrid 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ĥ"
         Height          =   210
         Left            =   750
         TabIndex        =   95
         Tag             =   "G"
         Top             =   90
         Width           =   720
      End
      Begin VB.CheckBox chkGas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "GAS"
         Height          =   210
         Left            =   60
         TabIndex        =   94
         Tag             =   "C"
         Top             =   90
         Width           =   645
      End
      Begin VB.TextBox txtGas 
         Height          =   285
         Left            =   870
         TabIndex        =   93
         Top             =   300
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox txtGrid 
         Height          =   285
         Left            =   1170
         TabIndex        =   92
         Top             =   330
         Visible         =   0   'False
         Width           =   210
      End
   End
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Index           =   1
      Left            =   15330
      Top             =   600
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "����ԭ��"
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
      ForeColor       =   16711680
   End
   Begin Threed.SSCommand cmd_Off 
      Height          =   375
      Left            =   16530
      TabIndex        =   106
      Top             =   150
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      _Version        =   196609
      Caption         =   "����"
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   3030
      Left            =   90
      TabIndex        =   72
      Top             =   1005
      Width           =   15210
      _Version        =   393216
      _ExtentX        =   26829
      _ExtentY        =   5345
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   23
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "EGA1080C.frx":001E
   End
   Begin Threed.SSFrame sf3 
      Height          =   4995
      Left            =   90
      TabIndex        =   22
      Top             =   4035
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   8811
      _Version        =   196609
      Font3D          =   2
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " �ߴ�"
      Begin VB.TextBox TXT_SIZE_KND_NAME 
         Height          =   315
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   152
         Tag             =   "����"
         Top             =   3990
         Width           =   1050
      End
      Begin VB.TextBox TXT_SIZE_KND 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   151
         Tag             =   "ԭ��"
         Top             =   3990
         Width           =   840
      End
      Begin VB.TextBox TXT_WAVE1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3630
         MaxLength       =   2
         TabIndex        =   142
         Top             =   3240
         Width           =   990
      End
      Begin VB.TextBox txtCl 
         Height          =   285
         Left            =   2130
         TabIndex        =   133
         Top             =   2970
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CheckBox chkCl 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ֱָʾ"
         Height          =   210
         Left            =   2100
         TabIndex        =   132
         Tag             =   "G"
         Top             =   3690
         UseMaskColor    =   -1  'True
         Width           =   1080
      End
      Begin VB.TextBox TXT_VERT_DEG 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   109
         Top             =   3630
         Width           =   840
      End
      Begin VB.TextBox TXT_RECT_DEG 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3990
         MaxLength       =   2
         TabIndex        =   108
         Top             =   3630
         Width           =   840
      End
      Begin VB.TextBox TXT_WAVE 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   107
         Top             =   3240
         Width           =   840
      End
      Begin VB.TextBox TXT_INSP_WGT_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2415
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_THK_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2415
         Width           =   840
      End
      Begin VB.TextBox TXT_INSP_LEN_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2415
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_WID_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1950
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2415
         Width           =   960
      End
      Begin InDate.ULabel ULabel28 
         Height          =   315
         Left            =   1950
         Top             =   285
         Width           =   960
         _ExtentX        =   1693
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
      Begin InDate.ULabel ULabel29 
         Height          =   315
         Left            =   1080
         Top             =   285
         Width           =   840
         _ExtentX        =   1482
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
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Left            =   2940
         Top             =   285
         Width           =   960
         _ExtentX        =   1693
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
      Begin InDate.ULabel ULabel33 
         Height          =   315
         Left            =   60
         Top             =   2415
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "�ж����"
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
      Begin CSTextLibCtl.sidbEdit SDB_WGT_ORD 
         Height          =   315
         Left            =   3930
         TabIndex        =   17
         Top             =   1335
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_WGT 
         Height          =   315
         Left            =   3930
         TabIndex        =   18
         Top             =   615
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_WID_MX 
         Height          =   315
         Left            =   1950
         TabIndex        =   12
         Top             =   1695
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   14737632
         BackColor       =   16777215
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LEN_MX 
         Height          =   315
         Left            =   2940
         TabIndex        =   15
         Top             =   1695
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_WID_MN 
         Height          =   315
         Left            =   1950
         TabIndex        =   13
         Top             =   2055
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_THK_MN 
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   2055
         Width           =   840
         _Version        =   262145
         _ExtentX        =   1482
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LEN_MN 
         Height          =   315
         Left            =   2940
         TabIndex        =   16
         Top             =   2055
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_PWGT_MN 
         Height          =   315
         Left            =   3930
         TabIndex        =   19
         Top             =   2055
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_WID 
         Height          =   315
         Left            =   1950
         TabIndex        =   10
         Top             =   615
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
      Begin CSTextLibCtl.sidbEdit SDB_THK 
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   615
         Width           =   840
         _Version        =   262145
         _ExtentX        =   1482
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
      Begin CSTextLibCtl.sidbEdit SDB_LEN 
         Height          =   315
         Left            =   2940
         TabIndex        =   35
         Top             =   615
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel38 
         Height          =   315
         Left            =   60
         Top             =   2055
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "�¹���"
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
      Begin InDate.ULabel ULabel43 
         Height          =   315
         Left            =   60
         Top             =   615
         Width           =   990
         _ExtentX        =   1746
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_THK_MX 
         Height          =   315
         Left            =   1080
         TabIndex        =   55
         Top             =   1695
         Width           =   840
         _Version        =   262145
         _ExtentX        =   1482
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_PWGT_MX 
         Height          =   315
         Left            =   3930
         TabIndex        =   56
         Top             =   1695
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   60
         Top             =   1695
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "�Ϲ���"
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
      Begin InDate.ULabel ULabel44 
         Height          =   315
         Left            =   3945
         Top             =   285
         Width           =   930
         _ExtentX        =   1640
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
      Begin CSTextLibCtl.sidbEdit SDB_ORD_WID 
         Height          =   315
         Left            =   1950
         TabIndex        =   57
         Top             =   1350
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
      Begin CSTextLibCtl.sidbEdit SDB_ORD_THK 
         Height          =   315
         Left            =   1080
         TabIndex        =   58
         Top             =   1335
         Width           =   840
         _Version        =   262145
         _ExtentX        =   1482
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
      Begin CSTextLibCtl.sidbEdit SDB_ORD_LEN 
         Height          =   315
         Left            =   2940
         TabIndex        =   59
         Top             =   1335
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel45 
         Height          =   315
         Left            =   60
         Top             =   1335
         Width           =   990
         _ExtentX        =   1746
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   5
         Left            =   60
         Top             =   3240
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "��ƽ��(/m)"
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   6
         Left            =   60
         Top             =   3630
         Width           =   990
         _ExtentX        =   1746
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   7
         Left            =   3210
         Top             =   3630
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         Caption         =   "��б"
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
      Begin CSTextLibCtl.sidbEdit SDB_WID_R 
         Height          =   315
         Left            =   1950
         TabIndex        =   135
         Top             =   980
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
      Begin CSTextLibCtl.sidbEdit SDB_THK_R 
         Height          =   315
         Left            =   1080
         TabIndex        =   136
         Top             =   980
         Width           =   840
         _Version        =   262145
         _ExtentX        =   1482
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
      Begin CSTextLibCtl.sidbEdit SDB_LEN_R 
         Height          =   315
         Left            =   2940
         TabIndex        =   137
         Top             =   980
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   60
         Top             =   980
         Width           =   990
         _ExtentX        =   1746
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
      Begin InDate.ULabel ULabel27 
         Height          =   315
         Left            =   60
         Top             =   2760
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "�Խ���1"
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_DIAGONAL1 
         Height          =   315
         Left            =   1080
         TabIndex        =   138
         Top             =   2760
         Width           =   1350
         _Version        =   262145
         _ExtentX        =   2381
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel31 
         Height          =   315
         Left            =   2520
         Top             =   2760
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "�Խ���2"
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_DIAGONAL2 
         Height          =   315
         Left            =   3540
         TabIndex        =   139
         Top             =   2760
         Width           =   1350
         _Version        =   262145
         _ExtentX        =   2381
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   8
         Left            =   2460
         Top             =   3240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         Caption         =   "��ƽ��(/2m)"
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
      Begin InDate.ULabel ULabel48 
         Height          =   315
         Left            =   60
         Top             =   3990
         Width           =   990
         _ExtentX        =   1746
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
         ForeColor       =   16711680
      End
   End
   Begin Threed.SSFrame SF4 
      Height          =   4995
      Left            =   10140
      TabIndex        =   36
      Top             =   4035
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   8811
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   16711680
      BackColor       =   14737632
      Caption         =   "�ж�"
      Begin VB.CheckBox CHK_FLAW_YN 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�±��Ƿ����"
         Height          =   240
         Left            =   3300
         TabIndex        =   149
         Tag             =   "G"
         Top             =   3750
         Width           =   1620
      End
      Begin VB.TextBox txt_Color_name 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   141
         Top             =   3720
         Width           =   1395
      End
      Begin VB.TextBox txt_Color_code 
         Height          =   300
         Left            =   1380
         MaxLength       =   2
         TabIndex        =   140
         Tag             =   "ԭ��"
         Top             =   3720
         Width           =   405
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   3990
         MaxLength       =   3
         TabIndex        =   97
         Top             =   3330
         Width           =   945
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   2
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   3330
         Width           =   2595
      End
      Begin VB.TextBox TXT_PROC_CD 
         Alignment       =   2  'Center
         BackColor       =   &H00E1E4CD&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   74
         Tag             =   "�����ж�"
         Text            =   " "
         Top             =   1875
         Width           =   840
      End
      Begin CSTextLibCtl.sidbEdit SDB_Mn 
         Height          =   225
         Left            =   1230
         TabIndex        =   73
         Top             =   1470
         Width           =   840
         _Version        =   262145
         _ExtentX        =   1482
         _ExtentY        =   397
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   255
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
         Text            =   ""
         StartText.x     =   2
         StartText.y     =   0
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
      Begin VB.TextBox txt_Scrap_name 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3555
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   1815
         Width           =   1395
      End
      Begin VB.TextBox txt_Scrap_code 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3135
         MaxLength       =   1
         TabIndex        =   70
         Tag             =   "ԭ��"
         Top             =   1815
         Width           =   405
      End
      Begin VB.TextBox txt_stdspec_yy 
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
         Height          =   330
         Left            =   3750
         MaxLength       =   40
         TabIndex        =   69
         Tag             =   "STDSPEC"
         Top             =   2190
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txt_stdspec_name_chg 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   68
         Tag             =   "STDSPEC"
         Top             =   2910
         Width           =   2840
      End
      Begin VB.TextBox txt_stdspec_chg 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         MaxLength       =   18
         TabIndex        =   67
         Tag             =   "��׼��"
         Top             =   2910
         Width           =   1965
      End
      Begin VB.TextBox txt_stdspec_name 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   66
         Tag             =   "STDSPEC"
         Top             =   2580
         Width           =   2840
      End
      Begin VB.TextBox txt_stdspec 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   65
         Tag             =   "��׼����"
         Top             =   2580
         Width           =   1965
      End
      Begin VB.TextBox TXT_SURF_GRD 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1610
         Locked          =   -1  'True
         TabIndex        =   60
         Tag             =   "�����ж�"
         Text            =   " "
         Top             =   300
         Width           =   840
      End
      Begin VB.TextBox TXT_INSP_MAIN_GRD 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1610
         Locked          =   -1  'True
         TabIndex        =   37
         Tag             =   "����ȼ��ж�"
         Top             =   750
         Width           =   840
      End
      Begin InDate.ULabel ULabel22 
         Height          =   330
         Index           =   0
         Left            =   135
         Top             =   750
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         Caption         =   "����ȼ��ж�"
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
      Begin InDate.ULabel ULabel36 
         Height          =   330
         Left            =   135
         Top             =   300
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         Caption         =   "�����ж�"
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
      Begin InDate.ULabel ULabel22 
         Height          =   300
         Index           =   1
         Left            =   135
         Top             =   2250
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   529
         Caption         =   "��׼��"
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
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Left            =   2490
         Top             =   1815
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   529
         Caption         =   "ԭ��"
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
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel22 
         Height          =   300
         Index           =   2
         Left            =   135
         Top             =   1410
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   529
         Caption         =   "Mn �ɷ� (         )"
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
         ForeColor       =   255
      End
      Begin InDate.ULabel ULabel22 
         Height          =   300
         Index           =   3
         Left            =   135
         Top             =   1815
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   529
         Caption         =   "��   �� (         )"
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
         ForeColor       =   255
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   315
         Left            =   2490
         TabIndex        =   76
         Top             =   300
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   3300
            TabIndex        =   77
            Text            =   " "
            Top             =   30
            Width           =   225
         End
         Begin Threed.SSOption opt_CHK_SUR_GRD 
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   78
            Top             =   30
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   196609
            Font3D          =   1
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "�ϸ�"
         End
         Begin Threed.SSOption opt_CHK_SUR_GRD 
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   79
            Top             =   30
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   450
            _Version        =   196609
            Font3D          =   1
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "���ϸ�"
         End
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   1005
         Left            =   2490
         TabIndex        =   80
         Top             =   750
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   1773
         _Version        =   196609
         BackColor       =   14737632
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   81
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            _Version        =   196609
            Font3D          =   1
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "��Ʒ"
         End
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   82
            Top             =   420
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   450
            _Version        =   196609
            Font3D          =   1
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "����"
         End
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   83
            Top             =   720
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   450
            _Version        =   196609
            Font3D          =   1
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Э��"
         End
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   255
            Index           =   3
            Left            =   1590
            TabIndex        =   84
            Top             =   90
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   196609
            Font3D          =   1
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "����"
         End
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   255
            Index           =   4
            Left            =   1590
            TabIndex        =   85
            Top             =   390
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   196609
            Font3D          =   1
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "��Ʒ"
         End
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   255
            Index           =   5
            Left            =   1590
            TabIndex        =   86
            Top             =   720
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   196609
            Font3D          =   1
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "�ϸ�"
         End
      End
      Begin InDate.ULabel ULabel25 
         Height          =   315
         Left            =   150
         Top             =   3330
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "����ȱ��"
         Alignment       =   1
         BackColor       =   8421631
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
      Begin InDate.ULabel ULabel32 
         Height          =   315
         Left            =   150
         Top             =   3720
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "������ɫ"
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
         Left            =   150
         Top             =   4110
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         Caption         =   "���1"
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
      Begin CSTextLibCtl.sidbEdit SDB_HD1 
         Height          =   315
         Left            =   810
         TabIndex        =   143
         Top             =   4110
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
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
      Begin InDate.ULabel ULabel39 
         Height          =   315
         Left            =   1650
         Top             =   4110
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         Caption         =   "���2"
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
      Begin CSTextLibCtl.sidbEdit SDB_HD2 
         Height          =   315
         Left            =   2280
         TabIndex        =   144
         Top             =   4110
         Width           =   780
         _Version        =   262145
         _ExtentX        =   1376
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
      Begin InDate.ULabel ULabel40 
         Height          =   315
         Left            =   3090
         Top             =   4110
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         Caption         =   "���3"
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
      Begin CSTextLibCtl.sidbEdit SDB_HD3 
         Height          =   315
         Left            =   3720
         TabIndex        =   145
         Top             =   4110
         Width           =   750
         _Version        =   262145
         _ExtentX        =   1323
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
      Begin InDate.ULabel ULabel41 
         Height          =   315
         Left            =   150
         Top             =   4500
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         Caption         =   "���4"
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
      Begin CSTextLibCtl.sidbEdit SDB_HD4 
         Height          =   315
         Left            =   840
         TabIndex        =   146
         Top             =   4500
         Width           =   780
         _Version        =   262145
         _ExtentX        =   1376
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
      Begin InDate.ULabel ULabel42 
         Height          =   315
         Left            =   1650
         Top             =   4500
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         Caption         =   "���5"
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
      Begin CSTextLibCtl.sidbEdit SDB_HD5 
         Height          =   315
         Left            =   2280
         TabIndex        =   147
         Top             =   4500
         Width           =   780
         _Version        =   262145
         _ExtentX        =   1376
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
      Begin InDate.ULabel ULabel46 
         Height          =   315
         Left            =   3090
         Top             =   4500
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         Caption         =   "���6"
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
      Begin CSTextLibCtl.sidbEdit SDB_HD6 
         Height          =   315
         Left            =   3720
         TabIndex        =   148
         Top             =   4500
         Width           =   750
         _Version        =   262145
         _ExtentX        =   1323
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
   End
   Begin VB.Frame sf5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ȱ��"
      Height          =   945
      Left            =   5040
      TabIndex        =   128
      Top             =   6180
      Width           =   5385
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   3
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   134
         Top             =   210
         Width           =   2220
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   3570
         TabIndex        =   131
         Top             =   210
         Width           =   855
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   3570
         TabIndex        =   130
         Top             =   540
         Width           =   855
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   129
         Top             =   540
         Width           =   2220
      End
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   180
         Top             =   540
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         Caption         =   "�±���"
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
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   180
         Top             =   210
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         Caption         =   "�ϱ���"
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
   End
End
Attribute VB_Name = "EGA1080C"
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
'-- Program Name      ������ʵ����ѯ���޸Ľ���
'-- Program ID        EGA1080C
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
Public sQuery_Rt As String          'Active Form sQuery Setting

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

Dim sControl  As New Collection      'Master Clear Key Collection
Dim MC        As New Collection      'Master Collection
Dim Mc1       As New Collection      'Master Collection

Dim sc1       As New Collection      'Spread Collection
Dim Proc_Sc   As New Collection      'Spread Struc Collection

Dim sCheck  As String
Dim sQuery  As String

Private Sub Form_Define()
    Dim iIndex As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDT_PROD_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDT_PROD_TO_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_stdspec_chg_ref, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_THK_REF, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_WID_REF, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_UST_FLAG, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_PROC_FLAG, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_APLY_ENDUSE_CD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_STLGRD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                                                                                                                                                
     Call Gp_Ms_Collection(TXT_INSP_FLAW(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_FLAW(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_FLAW(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_PART(3), " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_PART(4), " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_PART(5), " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LTH(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LTH(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LTH(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_FLAW(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_FLAW(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_FLAW(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_PART(0), " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_PART(1), " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_PART(2), " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LTH(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LTH(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LTH(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_THK, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_THK_MX, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_THK_MN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_WID, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_WID_MX, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_WID_MN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_LEN, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LEN_MX, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LEN_MN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_WGT_ORD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_WGT, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_PWGT_MX, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_PWGT_MN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
     Call Gp_Ms_Collection(TXT_INSP_THK_GRD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_WID_GRD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_LEN_GRD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_WGT_GRD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_SURF_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_INSP_MAIN_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
        'Call Gp_Ms_Collection(TXT_NEXT_PROC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
               Call Gp_Ms_Collection(txtGas, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txtGrid, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txtCl, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(TXT_INSP_MAN, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(TXT_INSP_OCCR_TIME, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_ORD_WID, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_ORD_THK, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_ORD_LEN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_GRID_EMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_GRID_TIME, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_TOP_GRID_GRD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_TOP_GRID_YRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_TOP_GRID_DEEP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_BOT_GRID_GRD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_BOT_GRID_YRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_BOT_GRID_DEEP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_stdspec, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_stdspec_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_stdspec_chg, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_stdspec_name_chg, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Scrap_code, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Scrap_name, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(SDB_Mn, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_PROC_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            'add by liqian at 20120322
          Call Gp_Ms_Collection(TXT_EMP_CD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_WAVE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_VERT_DEG, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_RECT_DEG, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           'ADD BY LIQIAN at 2013-05-29 ʵ�ʲ����ߴ�
            Call Gp_Ms_Collection(SDB_THK_R, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDB_WID_R, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDB_LEN_R, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_INSP_DIAGONAL1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '�Խ���1
   Call Gp_Ms_Collection(SDB_INSP_DIAGONAL2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '�Խ���2
       Call Gp_Ms_Collection(txt_Color_code, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(TXT_WAVE1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
    Call Gp_Ms_Collection(TXT_INSP_MAN_TAIL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
              Call Gp_Ms_Collection(SDB_HD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_HD2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_HD3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_HD4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_HD5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_HD6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                
          Call Gp_Ms_Collection(CHK_FLAW_YN, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_size_knd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

        
    For iIndex = 0 To 17
        Call Gp_Clear_Collection(CHK_PART(iIndex), "s", sControl)
    Next iIndex
    
     Call Gp_Clear_Collection(CHK_TOP_GRD(0), "s", sControl)
     Call Gp_Clear_Collection(CHK_TOP_GRD(1), "s", sControl)
     Call Gp_Clear_Collection(CHK_BOT_GRD(0), "s", sControl)
     Call Gp_Clear_Collection(CHK_BOT_GRD(1), "s", sControl)
     Call Gp_Clear_Collection(CHK_BOT_GRD(1), "s", sControl)
     
    
    MC.Add Item:=sControl, Key:="sControl"
    
    'MASTER Collection
    Mc1.Add Item:="EGA1080C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="EGA1080C.P_REFER", Key:="P-R"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
      
    'Spread_Collection
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
     Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="EGA1080C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
        
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub cbo_ResonDesc_Click()
    txt_ResonCd = Mid(cbo_ResonDesc.Text, 1, 1)
End Sub

Private Sub CHK_TOP_GRD_Click(Index As Integer)
    Dim iNext       As Integer
    
    If sCheck <> "" Then Exit Sub

    sCheck = "**"
    
    If Index = 0 Then
        iNext = 1
    Else
        iNext = 0
    End If
    
    If CHK_TOP_GRD(Index).Value = ssCBUnchecked Then
        If CHK_TOP_GRD(iNext).Value = ssCBUnchecked Then
            TXT_TOP_GRID_GRD.Text = ""
            CHK_TOP_GRD(Index).ForeColor = &H808080
            sCheck = ""
            Exit Sub
        End If
    End If
    
    CHK_TOP_GRD(Index).ForeColor = &HFF&
    CHK_TOP_GRD(Index).Value = ssCBChecked
                
    CHK_TOP_GRD(iNext).ForeColor = &H808080
    CHK_TOP_GRD(iNext).Value = ssCBUnchecked

    TXT_TOP_GRID_GRD.Text = CHK_TOP_GRD(Index).Tag
    sCheck = ""
    
End Sub

Private Sub chkCl_Click()
    If chkCl.Value = ssCBChecked Then
        txtCl.Text = "Y"
        chkCl.ForeColor = &HFF&       'red
    Else
        txtCl.Text = "N"
        chkCl.ForeColor = &H80000012       'red
    End If
End Sub

Private Sub chkGas_Click()
    If chkGas.Value Then
        txtGas = "Y"
        chkGas.ForeColor = &HFF&       'red
    Else
        txtGas = "N"
        chkGas.ForeColor = &H80000012       'red
    End If
End Sub

Private Sub chkGrid_Click()
    If chkGrid.Value Then
        txtGrid = "Y"
        chkGrid.ForeColor = &HFF&       'red
    Else
        txtGrid = "N"
        chkGrid.ForeColor = &H80000012       'red
    End If
End Sub

'Private Sub cmd_Off_Click()
'    Dim OutParam(2, 4) As Variant
'    Dim sQuery As String
'    Dim adoCmd As ADODB.Command
'
'
'    On Error Resume Next
'
'    Screen.MousePointer = vbHourglass
'
'
'    'Return loaction1 Parameter
'    OutParam(1, 1) = "arg_loaction1"
'    OutParam(1, 2) = adVarChar
'    OutParam(1, 3) = adParamOutput
'    OutParam(1, 4) = 10
'
'    'Return loaction2 Parameter
'    OutParam(2, 1) = "arg_loaction2"
'    OutParam(2, 2) = adVarChar
'    OutParam(2, 3) = adParamOutput
'    OutParam(2, 4) = 10
'
'    sQuery = "{call CGD2050C.P_LINEOFF('" & Trim(TXT_PLATE_NO.Text) & "','" & txt_PrcLine & "','" & txt_ResonCd & "','" & Gf_ShiftSet3(M_CN1) & "','" & sUserID & "',?,?)}"
'
'    'Ado Setting
'    M_CN1.CursorLocation = adUseServer
'    Set adoCmd = New ADODB.Command
'
'    adoCmd.CommandType = adCmdText
'    Set adoCmd.ActiveConnection = M_CN1
'
'    adoCmd.CommandText = sQuery
'
'    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
'    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
'
'    adoCmd.Execute , , adExecuteNoRecords
'
'    'Process Error Check
'    If Trim(adoCmd("arg_loaction2")) <> "" Then
'        Call Gp_MsgBoxDisplay("ʵ������ʧ�ܣ���ȷ��=> " & adoCmd("arg_loaction2"))
'    End If
'
'    Set adoCmd = Nothing
'
'    Call Form_Ref
'
'    Screen.MousePointer = vbDefault
'
'End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        If Len(TXT_PLATE_NO.Text) >= 8 Then
           Call Form_Ref
        End If
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
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))

    Call Gf_Sp_Cls(sc1)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "EG-System.INI", Me.Name)
    
   
    If TXT_PLATE_NO <> "" Then
       Call Form_Ref
    End If
        
    cbo_ResonDesc.AddItem "1:�豸�쳣"
    cbo_ResonDesc.AddItem "2:�߹�����"
    cbo_ResonDesc.AddItem "3:��Ʒ�쳣"
    
    Screen.MousePointer = vbDefault
    
    If Mid(sAuthority, 1, 3) = "111" Then
       cmd_Off.Enabled = True
    Else
       cmd_Off.Enabled = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "EG-System.INI", Me.Name)

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
    
    Set sControl = Nothing
    Set MC = Nothing

    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    Dim iCount As Integer
    
    If Gf_Sp_Cls(sc1) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        TXT_PLATE_NO = ""
        Call Gp_SSCheck_Cls(MC("sControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)

        TXT_INSP_MAN = ""
        TXT_EMP_CD1 = sUserID
        
        For iCount = 0 To 5
            TXT_INSP_FLAW_NAME(iCount).Text = ""
        Next iCount
        
        ss1.BlockMode = True
        ss1.ROW = -1
        ss1.Col = -1
        ss1.BackColor = &HFFFFFF
        ss1.BlockMode = False
        
        chkCl.Value = 0
        CHK_FLAW_YN.Value = 0
    End If
End Sub

Public Sub Form_Ref()
Dim i As Integer
    
    If SDT_PROD_DATE.RawData = "" Then
       SDT_PROD_DATE.RawData = Format(Now, "yyyymmdd")
    End If
    
    If SDT_PROD_TO_DATE.RawData = "" Then
       SDT_PROD_TO_DATE.RawData = Format(Now, "yyyymmdd")
    End If
    
    Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1)
    
    If ss1.MaxRows > 0 Then
       ss1.ROW = 1
       ss1.Col = 1
       TXT_PLATE_NO.Text = ss1.Text
       For i = 0 To 21  '17
           CHK_PART(i).Value = 0
       Next
    End If
    
    If Len(TXT_PLATE_NO.Text) = 14 Then
        If Gf_Ms_Refer(M_CN1, Mc1, , , False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            
            If txt_SURF_GRD = "Y" Then
               opt_CHK_SUR_GRD(0).Value = True
            Else
               opt_CHK_SUR_GRD(1).Value = True
            End If
            
            
            If Len(TXT_INSP_MAIN_GRD) = 1 Then
                If TXT_INSP_MAIN_GRD = "7" Then
                   opt_CHK_PRD_GRD(5).Value = True
                Else
                   opt_CHK_PRD_GRD(TXT_INSP_MAIN_GRD - 1).Value = True
                End If
            End If
            If TXT_INSP_OCCR_TIME.RawData = "" Then
               TXT_INSP_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
            End If
            'TXT_INSP_MAN = sUserID
            TXT_EMP_CD1.Text = sUserID
            'Call Display_Data_Edit
        End If
    End If
     
End Sub

Public Sub Form_Pro()

    Dim sMesg   As String
    Dim iCount  As Integer
    
'    For icount = 0 To 5
'        If TXT_INSP_FLAW_NAME(icount).Text <> "" And TXT_INSP_PART(icount).Text = "" Then
'            sMesg = " ������ȱ�ݲ�λ ��"
'            Call Gp_MsgBoxDisplay(sMesg)
'            Exit Sub
'        End If
'    Next icount
        
    If Trim(TXT_INSP_MAIN_GRD.Text) <> "4" Then
        If Trim(txt_SURF_GRD.Text) = "" Then
            sMesg = " ����������ж� ��"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        End If
        
    End If
    
    If Not Gp_DateCheck(TXT_INSP_OCCR_TIME) Then
        sMesg = " ����ȷ������ʱ�� ��"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    If CHK_GRID_FLAG.Value = ssCBChecked Then
        If Not Gp_DateCheck(TXT_GRID_TIME) Then
            sMesg = " ����ȷ������ĥʱ�� ��"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        End If
        If Trim(TXT_GRID_EMP_CD.Text) = "" Then
            TXT_GRID_EMP_CD.Text = sUserID
        End If
        If TXT_TOP_GRID_GRD.Text = "" Then
            sMesg = " ����ȷ�����ϱ�����ĥ���ж� ��"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        End If
        If TXT_BOT_GRID_GRD.Text = "" Then
            sMesg = " ����ȷ�����±�����ĥ���ж� ��"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        End If
    End If
    
    
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        'TXT_INSP_MAN.Text = sUserID
        TXT_EMP_CD1.Text = sUserID
       If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    End If

End Sub

Private Sub opt_CHK_PRD_GRD_Click(Index As Integer, Value As Integer)
    If Index = 0 Then
       TXT_INSP_MAIN_GRD = "1"
       opt_CHK_PRD_GRD(0).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
       txt_Scrap_code.Text = ""
       txt_Scrap_code.Enabled = False
    ElseIf Index = 1 Then
       TXT_INSP_MAIN_GRD = "2"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
       txt_Scrap_code.Text = ""
       txt_Scrap_code.Enabled = False
    ElseIf Index = 2 Then
        TXT_INSP_MAIN_GRD = "3"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
       txt_Scrap_code.Text = ""
       txt_Scrap_code.Enabled = False
    ElseIf Index = 3 Then
        TXT_INSP_MAIN_GRD = "4"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
       txt_Scrap_code.Text = ""
       txt_Scrap_code.Enabled = False
    ElseIf Index = 4 Then
        TXT_INSP_MAIN_GRD = "5"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
       txt_Scrap_code.Text = ""
       txt_Scrap_code.Enabled = False
    ElseIf Index = 5 Then
        TXT_INSP_MAIN_GRD = "7"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &HFF&       'red
       txt_Scrap_code.Enabled = True
    End If
End Sub

Private Sub opt_CHK_SUR_GRD_Click(Index As Integer, Value As Integer)
    If Index = 0 Then
       opt_CHK_SUR_GRD(0).ForeColor = &HFF&       'red
       opt_CHK_SUR_GRD(1).ForeColor = &H80000012  'black
        txt_SURF_GRD = "Y"
    Else
        txt_SURF_GRD = "N"
       opt_CHK_SUR_GRD(1).ForeColor = &HFF&       'red
       opt_CHK_SUR_GRD(0).ForeColor = &H80000012  'black
    End If
End Sub

'Private Sub opt_LineFlag_Click(Index As Integer, Value As Integer)
''    Call Form_Cls
''    TXT_PLATE_NO = ""
'    If opt_LineFlag(0).Value = True Then
'       txt_PrcLine = "1"
'       opt_LineFlag(0).ForeColor = &HFF&       'red
'       opt_LineFlag(1).ForeColor = &H80000012  'black
'       opt_LineFlag(2).ForeColor = &H80000012  'black
'    ElseIf opt_LineFlag(1).Value = True Then
'       txt_PrcLine = "2"
'       opt_LineFlag(0).ForeColor = &H80000012       'black
'       opt_LineFlag(1).ForeColor = &HFF&  'red
'       opt_LineFlag(2).ForeColor = &H80000012       'black
'    ElseIf opt_LineFlag(2).Value = True Then
'       txt_PrcLine = "3"
'       opt_LineFlag(0).ForeColor = &H80000012       'black
'       opt_LineFlag(1).ForeColor = &H80000012       'black
'       opt_LineFlag(2).ForeColor = &HFF&  'red
'    End If
'End Sub

Private Sub SDB_THK_Change()
    Call PRD_WEIGHT_CALC
End Sub
    
Private Sub SDB_WID_Change()
    Call PRD_WEIGHT_CALC
End Sub

Private Sub SDB_LEN_Change()
    Call PRD_WEIGHT_CALC
End Sub

Private Sub PRD_WEIGHT_CALC()

    Dim dThk        As Double
    Dim dWid        As Double
    Dim dLen        As Double
    
    dThk = Val(Format(SDB_THK.Text, "####0.##") & "")
    dWid = Val(Format(SDB_WID.Text, "###0") & "")
    dLen = Val(Format(SDB_LEN.Text, "###0.##") & "")
    If dThk > 0 And dWid > 0 And dLen > 0 Then
        SDB_WGT.Text = Cal_Plate_Wgt("WGT", dThk, dWid, dLen)
    End If
    
    Call Size_Grade_Edit
End Sub

Private Function Cal_Plate_Wgt(sMode As String, dThk As Double, dWid As Double, dLen As Double) As Double

    Dim RS  As New ADODB.Recordset
    
    Cal_Plate_Wgt = 0
    
    sQuery = "SELECT  Gf_Cal_Plate_Wgt('" & sMode & "'" & vbCrLf
    sQuery = sQuery & "             ,'" & Trim(txt_APLY_ENDUSE_CD.Text) & "'" & vbCrLf
    sQuery = sQuery & "             ,'" & Trim(txt_STLGRD.Text) & "'" & vbCrLf
    sQuery = sQuery & "             ," & dThk & vbCrLf
    sQuery = sQuery & "             ," & dWid & vbCrLf
    sQuery = sQuery & "             ," & dLen & vbCrLf
    sQuery = sQuery & "             ,0 )" & vbCrLf
    sQuery = sQuery & "       FROM  DUAL " & vbCrLf
    RS.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        Cal_Plate_Wgt = Val(RS(0).Value & "")
    End If
    
    RS.Close
    Set RS = Nothing
     
End Function

Private Sub SDT_PROD_DATE_DblClick()
     SDT_PROD_DATE.RawData = Gf_DTSet(M_CN1, "D")
     SDT_PROD_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
End Sub
Private Sub SDT_PROD_TO_DATE_DblClick()
     SDT_PROD_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
End Sub

Private Sub txt_Color_code_Change()
If Len(Trim(txt_Color_code)) = txt_Color_code.MaxLength Then
        txt_Color_name.Text = Gf_ComnNameFind(M_CN1, "CG002", Trim(txt_Color_code.Text), 1)
    Else
        txt_Color_name.Text = ""
    End If
End Sub

Private Sub txt_Color_code_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
            
        DD.sWitch = "MS"
        DD.sKey = "CG002"
        DD.rControl.Add Item:=txt_Color_code
        DD.rControl.Add Item:=txt_Color_name
        
        DD.nameType = "1"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    End If

End Sub

Private Sub txt_Color_code_DblClick()
    Call txt_Color_code_KeyUp(vbKeyF4, 0)
End Sub


Private Sub TXT_GRID_EMP_CD_DblClick()
    TXT_GRID_EMP_CD.Text = sUserID
End Sub

Private Sub TXT_GRID_TIME_DblClick()
    TXT_GRID_TIME.RawData = Gf_DTSet(M_CN1, , "X")
End Sub

Private Sub TXT_INSP_FLAW_Change(Index As Integer)
    TXT_INSP_FLAW_NAME(Index).Text = Gf_ComnNameFind(M_CN1, "G0002", TXT_INSP_FLAW(Index).Text, 1)
End Sub



Private Sub TXT_INSP_FLAW_NAME_DblClick(Index As Integer)
    DD.sWitch = "MS"
    DD.sKey = "G0002"
    DD.rControl.Add Item:=TXT_INSP_FLAW(Index)

    DD.nameType = "2"

    Call Gf_Common_DD(M_CN1, vbKeyF4)
    
    If Len(Trim(TXT_INSP_FLAW(Index).Text)) = 3 Then
        TXT_INSP_FLAW_NAME(Index).Text = Gf_ComnNameFind(M_CN1, "G0002", Trim(TXT_INSP_FLAW(Index).Text), 1)
    Else
        TXT_INSP_FLAW_NAME(Index).Text = ""
    End If
End Sub

Private Sub TXT_INSP_FLAW_NAME_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   Call TXT_INSP_FLAW_NAME_DblClick(Index)
End If
End Sub

Private Sub TXT_INSP_OCCR_TIME_DblClick()
    TXT_INSP_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
End Sub


Private Sub CHK_PART_Click(Index As Integer)
    Dim iCount      As Integer
    Dim iIndexTxt   As Integer
    Dim iIndexChk   As Integer
    Dim iIndexStr   As Integer
    
    If sCheck <> "" Then Exit Sub
    
    iIndexTxt = Index \ 3
    iIndexChk = iIndexTxt * 3
    iCount = 0
    sCheck = "**"
            
    If CHK_PART(Index).Value = ssCBUnchecked Then
        For iIndexStr = iIndexChk To iIndexChk + 2
            If CHK_PART(iIndexStr).Value = ssCBChecked Then
               iCount = iCount + 1
            End If
        Next iIndexStr
        If iCount = 0 Then
            TXT_INSP_PART(iIndexTxt).Text = ""
            TXT_INSP_FLAW(iIndexTxt).Text = ""
            TXT_INSP_FLAW_NAME(iIndexTxt).Text = ""
            CHK_PART(Index).ForeColor = &H808080
            sCheck = ""
            Exit Sub
        End If
    Else
        For iIndexStr = iIndexChk To iIndexChk + 2
            CHK_PART(iIndexStr).ForeColor = &H808080
            CHK_PART(iIndexStr).Value = ssCBUnchecked
        Next iIndexStr
    End If
    
    CHK_PART(Index).ForeColor = &HFF&
    CHK_PART(Index).Value = ssCBChecked

    TXT_INSP_PART(iIndexTxt).Text = CHK_PART(Index).Tag
    sCheck = ""
    
End Sub

Private Sub CHK_BOT_GRD_Click(Index As Integer)
    Dim iNext       As Integer
    
    If sCheck <> "" Then Exit Sub

    sCheck = "**"
    
    If Index = 0 Then
        iNext = 1
    Else
        iNext = 0
    End If
    
    If CHK_BOT_GRD(Index).Value = ssCBUnchecked Then
        If CHK_BOT_GRD(iNext).Value = ssCBUnchecked Then
            TXT_BOT_GRID_GRD.Text = ""
            CHK_BOT_GRD(Index).ForeColor = &H808080
            sCheck = ""
            Exit Sub
        End If
    End If
    
    CHK_BOT_GRD(Index).ForeColor = &HFF&
    CHK_BOT_GRD(Index).Value = ssCBChecked
                
    CHK_BOT_GRD(iNext).ForeColor = &H808080
    CHK_BOT_GRD(iNext).Value = ssCBUnchecked

    TXT_BOT_GRID_GRD.Text = CHK_BOT_GRD(Index).Tag
    sCheck = ""
    
End Sub

Private Sub CHK_GRID_FLAG_Click()
    If CHK_GRID_FLAG.Value = ssCBUnchecked Then
        CHK_TOP_GRD(0).Enabled = False:        CHK_TOP_GRD(0).Value = ssCBUnchecked
        CHK_TOP_GRD(1).Enabled = False:        CHK_TOP_GRD(1).Value = ssCBUnchecked
        CHK_BOT_GRD(0).Enabled = False:        CHK_BOT_GRD(0).Value = ssCBUnchecked
        CHK_BOT_GRD(1).Enabled = False:        CHK_BOT_GRD(1).Value = ssCBUnchecked
        SDB_TOP_GRID_YRD.Enabled = False:      SDB_TOP_GRID_YRD.Text = ""
        SDB_BOT_GRID_YRD.Enabled = False:      SDB_BOT_GRID_YRD.Text = ""
        SDB_TOP_GRID_DEEP.Enabled = False:     SDB_TOP_GRID_DEEP.Text = ""
        SDB_BOT_GRID_DEEP.Enabled = False:     SDB_BOT_GRID_DEEP.Text = ""
        TXT_GRID_EMP_CD.Enabled = False:       TXT_GRID_EMP_CD.Text = ""
        TXT_GRID_TIME.Enabled = False:         TXT_GRID_TIME.Text = ""
                
'        CHK_NEXT_PRC(1).Enabled = True
    Else
        CHK_TOP_GRD(0).Enabled = True
        CHK_TOP_GRD(1).Enabled = True
        CHK_BOT_GRD(0).Enabled = True
        CHK_BOT_GRD(1).Enabled = True
        SDB_TOP_GRID_YRD.Enabled = True
        SDB_BOT_GRID_YRD.Enabled = True
        SDB_TOP_GRID_DEEP.Enabled = True
        SDB_BOT_GRID_DEEP.Enabled = True
        TXT_GRID_EMP_CD.Enabled = True
        TXT_GRID_TIME.Enabled = True
        
        TXT_GRID_EMP_CD.Text = sUserID
        TXT_GRID_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        
        CHK_TOP_GRD(0).Value = ssCBChecked
        Call CHK_TOP_GRD_Click(0)
        CHK_BOT_GRD(0).Value = ssCBChecked
        Call CHK_BOT_GRD_Click(0)
        

'        CHK_NEXT_PRC(2).Value = ssCBChecked
'        Call CHK_NEXT_PRC_Click(2)

'        TXT_NEXT_PROC.Text = ""

    End If
End Sub

Private Sub Display_Data_Edit()
    Dim iIndexChk   As Integer
    Dim iIndexStr   As Integer
    
    sCheck = "**"
    
    For iIndexStr = 0 To 5
        For iIndexChk = iIndexStr * 3 To (iIndexStr * 3) + 2
            If TXT_INSP_PART(iIndexStr).Text = CHK_PART(iIndexChk).Tag Then
                CHK_PART(iIndexChk).ForeColor = &HFF&
                CHK_PART(iIndexChk).Value = ssCBChecked
            Else
                CHK_PART(iIndexChk).ForeColor = &H808080
                CHK_PART(iIndexChk).Value = ssCBUnchecked
            End If
        Next iIndexChk
    Next iIndexStr
        

    
    If Trim(TXT_TOP_GRID_GRD.Text) <> "" Then CHK_GRID_FLAG.Value = ssCBChecked
    
    If TXT_TOP_GRID_GRD.Text = "Y" Then
        CHK_TOP_GRD(0).Value = ssCBChecked
        CHK_TOP_GRD(1).Value = ssCBUnchecked
    ElseIf TXT_TOP_GRID_GRD.Text = "N" Then
        CHK_TOP_GRD(0).Value = ssCBUnchecked
        CHK_TOP_GRD(1).Value = ssCBChecked
    End If
    
    If TXT_BOT_GRID_GRD.Text = "Y" Then
        CHK_BOT_GRD(0).Value = ssCBChecked
        CHK_BOT_GRD(1).Value = ssCBUnchecked
    ElseIf TXT_BOT_GRID_GRD.Text = "N" Then
        CHK_BOT_GRD(0).Value = ssCBUnchecked
        CHK_BOT_GRD(1).Value = ssCBChecked
    End If
    
    If txtGas = "Y" Then
        chkGas.Value = 1
    End If
    If txtGrid = "Y" Then
        chkGrid.Value = 1
    End If
    If txtCl = "Y" Then
        chkCl.Value = 1
    End If

    If TXT_INSP_MAIN_GRD = "1" Then
        opt_CHK_PRD_GRD(0).Value = True
    ElseIf TXT_INSP_MAIN_GRD = "2" Then
        opt_CHK_PRD_GRD(1).Value = True
    ElseIf TXT_INSP_MAIN_GRD = "3" Then
        opt_CHK_PRD_GRD(2).Value = True
    ElseIf TXT_INSP_MAIN_GRD = "4" Then
        opt_CHK_PRD_GRD(3).Value = True
    ElseIf TXT_INSP_MAIN_GRD = "5" Then
        opt_CHK_PRD_GRD(4).Value = True
    ElseIf TXT_INSP_MAIN_GRD = "7" Then
        opt_CHK_PRD_GRD(5).Value = True
    End If
    
    If txt_SURF_GRD = "Y" Then
        opt_CHK_SUR_GRD(0).Value = True
    ElseIf txt_SURF_GRD = "N" Then
        opt_CHK_SUR_GRD(1).Value = True
    End If
    
    '''''''''ADD BY GUOLI AT 200712071330''''''''''
    If opt_CHK_SUR_GRD(0).Value = True Then
       txt_SURF_GRD = "Y"
    ElseIf opt_CHK_SUR_GRD(1).Value = True Then
       txt_SURF_GRD = "N"
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''

End Sub

Private Sub Size_Grade_Edit()
    Dim sGradeFlag As String
    
    sGradeFlag = ""
    
    If TXT_PROC_FLAG.Text <> "CGD" Then Exit Sub
    
    ' THICK GRAND CHECK
    If Val(SDB_THK & "") >= Val(SDB_ORD_THK & "") + Val(SDB_INSP_THK_MN & "") And _
       Val(SDB_THK & "") <= Val(SDB_ORD_THK & "") + Val(SDB_INSP_THK_MX & "") Then
        TXT_INSP_THK_GRD = "Y"
        SDB_THK.ForeColor = &H80000012
    Else
        TXT_INSP_THK_GRD = "N"
        SDB_THK.ForeColor = &HFF&
        sGradeFlag = "N"
    End If
    
    ' WIDTH GRAND CHECK
    If Val(SDB_WID & "") >= Val(SDB_ORD_WID & "") + Val(SDB_INSP_WID_MN & "") And _
       Val(SDB_WID & "") <= Val(SDB_ORD_WID & "") + Val(SDB_INSP_WID_MX & "") Then
        TXT_INSP_WID_GRD = "Y"
        SDB_WID.ForeColor = &H80000012
    Else
        TXT_INSP_WID_GRD = "N"
        SDB_WID.ForeColor = &HFF&
        sGradeFlag = "N"
    End If
        
    ' LENGTH GRAND CHECK
    If Val(SDB_LEN & "") >= Val(SDB_ORD_LEN & "") + Val(SDB_INSP_LEN_MN & "") And _
       Val(SDB_LEN & "") <= Val(SDB_ORD_LEN & "") + Val(SDB_INSP_LEN_MX & "") Then
        TXT_INSP_LEN_GRD = "Y"
        SDB_LEN.ForeColor = &H80000012
    Else
        TXT_INSP_LEN_GRD = "N"
        SDB_LEN.ForeColor = &HFF&
        sGradeFlag = "N"
    End If
    
    ' WEIGHT GRAND CHECK
    If Val(SDB_WGT & "") >= Val(SDB_WGT_ORD & "") + Val(SDB_PWGT_MN & "") And _
       Val(SDB_WGT & "") <= Val(SDB_WGT_ORD & "") + Val(SDB_PWGT_MX & "") Then
        TXT_INSP_WGT_GRD = "Y"
        SDB_WGT.ForeColor = &H80000012
    Else
        TXT_INSP_WGT_GRD = "N"
        SDB_WGT.ForeColor = &HFF&
        sGradeFlag = "N"
    End If
    


End Sub
  
Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       ' Call Gp_Ms_ControlLock(Mc1("pControl"), False)
     End If

End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
End Sub

'Private Sub ss1_EditChange(ByVal Col As Long, ByVal Row As Long)
'    Dim dThk        As Double
'    Dim dWid        As Double
'    Dim dLen        As Double
'    Dim dLenSum     As Double
'
'    Dim iIdr        As Integer
'
'    ss1.Row = Row
'    ss1.Col = 2:  dThk = Val(ss1.Text & "")
'    ss1.Col = 3:  dWid = Val(ss1.Text & "")
'    ss1.Col = 4:  dLen = Val(ss1.Text & "")
'
'    ss1.Col = 5
'    ss1.Text = Cal_Plate_Wgt("WGT", dThk, dWid, dLen)
'
'    For iIdr = 1 To ss1.MaxRows - 1
'        ss1.Row = iIdr
'        ss1.Col = 4
'        dLenSum = dLenSum + Val(ss1.Text & "")
'    Next iIdr
'
'    ss1.Row = ss1.MaxRows
'    ss1.Col = 4
'    dLen = ss1.Text 'SDB_LEN.Value - dLenSum
'    ss1.Text = dLen
'    ss1.Col = 5
'    ss1.Text = Cal_Plate_Wgt("WGT", dThk, dWid, dLen)
'End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)
    If ROW < 1 Then Exit Sub
    
    ss1.ROW = ROW
    ss1.Col = 1
    TXT_PLATE_NO.Text = ss1.Text
    CHK_GRID_FLAG.Value = ssCBUnchecked
    
    If Len(TXT_PLATE_NO.Text) = 14 Then
        Call Gp_SSCheck_Cls(MC("sControl"))
        If Gf_Ms_Refer(M_CN1, Mc1, , , True) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            ''''''''''''''''''ADD BY GUOLI AT 200712071330''''''''''
            If opt_CHK_SUR_GRD(0).Value = True Then
               txt_SURF_GRD = "Y"
            ElseIf opt_CHK_SUR_GRD(1).Value = True Then
               txt_SURF_GRD = "N"
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If Len(TXT_INSP_MAIN_GRD) = 1 Then
                If TXT_INSP_MAIN_GRD = "7" Then
                   opt_CHK_PRD_GRD(5).Value = True
                Else
                   opt_CHK_PRD_GRD(TXT_INSP_MAIN_GRD - 1).Value = True
                End If
            End If

            'Call Display_Data_Edit
        End If
        If TXT_INSP_OCCR_TIME.RawData = "" Then
           TXT_INSP_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        End If
        'TXT_INSP_MAN = sUserID
        TXT_EMP_CD1 = sUserID
        
    End If
    
End Sub

Private Sub TXT_INSP_PART_Change(Index As Integer)
Dim i As Integer
For i = 0 To 5
    If TXT_INSP_PART(i).Text = "T" Then
       CHK_PART(i * 3).Value = 1
    ElseIf TXT_INSP_PART(i).Text = "M" Then
       CHK_PART(i * 3 + 1).Value = 1
    ElseIf TXT_INSP_PART(i).Text = "B" Then
       CHK_PART(i * 3 + 2).Value = 1
    End If
Next
End Sub

Private Sub txt_Scrap_code_DblClick()
    Call txt_Scrap_code_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stdspec_chg_DblClick()
         DD.sWitch = "MS"
         DD.DataDicType = "C"
         DD.rControl.Add Item:=txt_stdspec_chg
         DD.rControl.Add Item:=txt_stdspec_name_chg
        
         Call Pf_Common_DD(M_CN1, vbKeyF4)
         
End Sub

Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        txt_stdspec_yy.Text = ""
        DD.rControl.Add Item:=txt_stdspec_chg
        DD.rControl.Add Item:=txt_stdspec_yy
        DD.rControl.Add Item:=txt_stdspec_name_chg

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

Private Function Pf_Common_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sSelect = False
        DD.sWhere = ""
        DD.sKey = ""
        
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sSelect = False
        DD.sWhere = ""
        DD.sKey = ""
        
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "HC"        'Common Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    DD.sQuery = "SELECT CD_SHORT_NAME ""��׼����"", CD_NAME ""��׼������"" FROM ZP_CD WHERE CD_MANA_NO = 'G0030'"
    
    Call Gf_DD_Display(Conn, DD.sQuery, False)
    
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function


Private Sub txt_Scrap_code_Change()
    
    If Len(Trim(txt_Scrap_code)) = txt_Scrap_code.MaxLength Then
        txt_Scrap_name.Text = Gf_ComnNameFind(M_CN1, "G0017", Trim(txt_Scrap_code.Text), 1)
    Else
        txt_Scrap_name.Text = ""
    End If
    
End Sub

Private Sub txt_Scrap_code_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
            
        DD.sWitch = "MS"
        DD.sKey = "G0017"
        DD.rControl.Add Item:=txt_Scrap_code
        DD.rControl.Add Item:=txt_Scrap_name
        
        DD.nameType = "1"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    End If

End Sub

Private Sub txt_stdspec_chg_ref_DblClick()
    Call txt_stdspec_chg_ref_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stdspec_chg_ref_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec_chg_ref

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

Private Sub TXT_INSP_MAN_DblClick()
    Call TXT_INSP_MAN_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_INSP_MAN_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0054"

        DD.rControl.Add Item:=TXT_INSP_MAN

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub TXT_SIZE_KND_Change()
    If Len(Trim(txt_size_knd.Text)) = txt_size_knd.MaxLength Then
        TXT_SIZE_KND_NAME.Text = Gf_ComnNameFind(M_CN1, "B0043", txt_size_knd.Text, 2)
    Else
        TXT_SIZE_KND_NAME.Text = ""
    End If
End Sub

Private Sub txt_size_knd_DblClick()
    Call txt_size_knd_KeyUp(vbKeyF4, 0)
End Sub
Private Sub txt_size_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sSize_knd As String
    sSize_knd = txt_size_knd.Text

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"

        DD.rControl.Add Item:=txt_size_knd

        DD.nameType = "2"
        txt_size_knd.Text = ""
        Call Gf_Common_DD(M_CN1, KeyCode)
        If txt_size_knd.Text = "" Then
            txt_size_knd.Text = sSize_knd
        End If
        
    End If
    
End Sub

Private Sub TXT_INSP_MAN_TAIL_DblClick()
    Call TXT_INSP_MAN_TAIL_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_INSP_MAN_TAIL_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0054"

        DD.rControl.Add Item:=TXT_INSP_MAN_TAIL

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub
