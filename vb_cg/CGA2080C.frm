VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGA2080C 
   Caption         =   "�����и���ҵ����_CGA2080C"
   ClientHeight    =   9345
   ClientLeft      =   -465
   ClientTop       =   2175
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7725
      Left            =   90
      TabIndex        =   24
      Top             =   1380
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   13626
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "�����и�"
      TabPicture(0)   =   "CGA2080C.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSSplitter1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "�и�ƻ�"
      TabPicture(1)   =   "CGA2080C.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ss3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "���ּƻ�"
      TabPicture(2)   =   "CGA2080C.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ss4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin SSSplitter.SSSplitter SSSplitter1 
         Height          =   7425
         Left            =   0
         TabIndex        =   25
         Top             =   300
         Width           =   15045
         _ExtentX        =   26538
         _ExtentY        =   13097
         _Version        =   196609
         SplitterBarWidth=   3
         BorderStyle     =   0
         PaneTree        =   "CGA2080C.frx":0054
         Begin FPSpread.vaSpread ss1 
            Height          =   4485
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   15045
            _Version        =   393216
            _ExtentX        =   26538
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
            MaxCols         =   25
            MaxRows         =   50
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CGA2080C.frx":00C6
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   570
            Left            =   0
            TabIndex        =   27
            Top             =   4545
            Width           =   15045
            _ExtentX        =   26538
            _ExtentY        =   1005
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
            Begin VB.TextBox TXT_SLABNO 
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
               Left            =   12330
               MaxLength       =   10
               TabIndex        =   29
               Top             =   90
               Visible         =   0   'False
               Width           =   2040
            End
            Begin VB.ComboBox cbo_cutcnt 
               Enabled         =   0   'False
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
               ItemData        =   "CGA2080C.frx":0F12
               Left            =   1575
               List            =   "CGA2080C.frx":0F14
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Tag             =   "��������"
               Top             =   120
               Width           =   705
            End
            Begin InDate.ULabel ULabel4 
               Height          =   315
               Left            =   240
               Top             =   120
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
               Caption         =   "�и����"
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
            Begin InDate.ULabel ULabel10 
               Height          =   315
               Left            =   2640
               Top             =   120
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
               Caption         =   "�ܳ���"
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
            Begin CSTextLibCtl.sidbEdit txt_total_len 
               Height          =   315
               Left            =   3960
               TabIndex        =   30
               Top             =   120
               Width           =   915
               _Version        =   262145
               _ExtentX        =   1614
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   255
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
               DataProperty    =   1
               ReadOnly        =   -1  'True
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
               MaxValue        =   20
               MinValue        =   10
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel11 
               Height          =   315
               Left            =   5130
               Top             =   120
               Width           =   1305
               _ExtentX        =   2302
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
               ForeColor       =   0
            End
            Begin CSTextLibCtl.sidbEdit txt_total_wgt 
               Height          =   315
               Left            =   6450
               TabIndex        =   31
               Top             =   120
               Width           =   915
               _Version        =   262145
               _ExtentX        =   1614
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   255
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
               DataProperty    =   1
               ReadOnly        =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   "0.000"
               Text            =   " 0.000"
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
               MaxValue        =   20
               MinValue        =   10
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel12 
               Height          =   315
               Left            =   7620
               Top             =   120
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   556
               Caption         =   "�ϸ�����"
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
            Begin CSTextLibCtl.sidbEdit txt_scrap_wgt 
               Height          =   315
               Left            =   8940
               TabIndex        =   32
               Top             =   120
               Width           =   915
               _Version        =   262145
               _ExtentX        =   1614
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0"
               ForeColor       =   255
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
               DataProperty    =   1
               ReadOnly        =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   "0.000"
               Text            =   " 0.000"
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
               MaxValue        =   20
               MinValue        =   10
               Undo            =   0
               Data            =   0
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   2250
            Left            =   0
            TabIndex        =   33
            Top             =   5175
            Width           =   15045
            _Version        =   393216
            _ExtentX        =   26538
            _ExtentY        =   3969
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
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
            MaxCols         =   16
            MaxRows         =   50
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CGA2080C.frx":0F16
         End
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   7425
         Left            =   -75000
         TabIndex        =   34
         Top             =   300
         Width           =   15045
         _Version        =   393216
         _ExtentX        =   26538
         _ExtentY        =   13097
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
         MaxCols         =   20
         MaxRows         =   50
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGA2080C.frx":1AB0
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   7425
         Left            =   -75000
         TabIndex        =   37
         Top             =   300
         Width           =   15045
         _Version        =   393216
         _ExtentX        =   26538
         _ExtentY        =   13097
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
         MaxCols         =   15
         MaxRows         =   50
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGA2080C.frx":276B
      End
   End
   Begin VB.TextBox txt_IST_DATE 
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
      Left            =   15630
      MaxLength       =   20
      TabIndex        =   19
      Top             =   810
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txt_tmpPLT 
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
      Left            =   15630
      MaxLength       =   20
      TabIndex        =   18
      Top             =   450
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txt_plt_dec 
      Enabled         =   0   'False
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
      Left            =   2100
      MaxLength       =   11
      TabIndex        =   17
      Top             =   570
      Width           =   1440
   End
   Begin VB.TextBox txt_plt 
      Alignment       =   2  'Center
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
      Left            =   1460
      MaxLength       =   2
      TabIndex        =   0
      Top             =   570
      Width           =   630
   End
   Begin VB.TextBox txt_Status 
      Alignment       =   2  'Center
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
      Left            =   2910
      MaxLength       =   11
      TabIndex        =   14
      Top             =   -30
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.OptionButton opt_prc_status2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "��ѯ���޸�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   3405
      TabIndex        =   13
      Top             =   180
      Width           =   1365
   End
   Begin VB.OptionButton opt_prc_status1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�����и�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1800
      TabIndex        =   12
      Top             =   180
      Value           =   -1  'True
      Width           =   1155
   End
   Begin VB.TextBox txt_act_stlgrd_dec 
      Enabled         =   0   'False
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
      Left            =   6510
      MaxLength       =   11
      TabIndex        =   11
      Top             =   570
      Width           =   1710
   End
   Begin VB.TextBox txt_MOSLAB 
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
      Left            =   6510
      MaxLength       =   10
      TabIndex        =   8
      Top             =   150
      Width           =   1710
   End
   Begin VB.TextBox txt_LOC 
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
      Left            =   9855
      MaxLength       =   11
      TabIndex        =   9
      Top             =   570
      Width           =   1440
   End
   Begin VB.TextBox txt_act_stlgrd 
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
      Left            =   5175
      MaxLength       =   11
      TabIndex        =   7
      Top             =   570
      Width           =   1305
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   3840
      Top             =   570
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   8520
      Top             =   570
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "��λ��"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   5175
      Top             =   150
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "ĸ������"
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
      Left            =   120
      Top             =   150
      Width           =   1305
      _ExtentX        =   2302
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   120
      Top             =   990
      Width           =   1305
      _ExtentX        =   2302
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   3840
      Top             =   990
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   7890
      Top             =   990
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   120
      Top             =   570
      Width           =   1305
      _ExtentX        =   2302
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
      ForeColor       =   16711680
   End
   Begin CSTextLibCtl.sidbEdit txt_wid 
      Height          =   315
      Left            =   5175
      TabIndex        =   3
      Top             =   990
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
      DataProperty    =   1
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
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_len 
      Height          =   315
      Left            =   9225
      TabIndex        =   5
      Top             =   990
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
      DataProperty    =   1
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
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_thk 
      Height          =   315
      Left            =   1460
      TabIndex        =   1
      Top             =   990
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
      DataProperty    =   1
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
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_thk_to 
      Height          =   315
      Left            =   2625
      TabIndex        =   2
      Top             =   990
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
      DataProperty    =   1
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
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_wid_to 
      Height          =   315
      Left            =   6360
      TabIndex        =   4
      Top             =   990
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
      DataProperty    =   1
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
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_len_to 
      Height          =   315
      Left            =   10380
      TabIndex        =   6
      Top             =   990
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
      DataProperty    =   1
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
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin Threed.SSCommand cmd_Cancel 
      Height          =   435
      Left            =   13170
      TabIndex        =   20
      Top             =   120
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   767
      _Version        =   196609
      ForeColor       =   16576
      BackColor       =   14737632
      BackStyle       =   1
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
      Caption         =   "ָʾȡ��"
   End
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   8520
      Top             =   150
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "ָʾ����"
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
   Begin InDate.UDate U_FROM_DATE 
      Height          =   315
      Left            =   9855
      TabIndex        =   21
      Tag             =   "��ʼ����"
      Top             =   150
      Width           =   1440
      _ExtentX        =   2540
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
   Begin InDate.UDate U_TO_DATE 
      Height          =   315
      Left            =   11505
      TabIndex        =   22
      Tag             =   "��ʼ����"
      Top             =   150
      Width           =   1440
      _ExtentX        =   2540
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
   Begin Threed.SSPanel SSP1 
      Height          =   375
      Left            =   11520
      TabIndex        =   35
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      Caption         =   "���и�"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSP3 
      Height          =   375
      Left            =   11520
      TabIndex        =   36
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�и�ƻ�"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSP6 
      Height          =   375
      Left            =   13200
      TabIndex        =   38
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      Caption         =   "���º�ͬ"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSP4 
      Height          =   375
      Left            =   13200
      TabIndex        =   39
      Top             =   570
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   196609
      ForeColor       =   65535
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
      Caption         =   "�ص㶩��"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   11340
      TabIndex        =   23
      Top             =   270
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   10170
      TabIndex        =   16
      Top             =   1110
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   6135
      TabIndex        =   15
      Top             =   1110
      Width           =   195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   2400
      TabIndex        =   10
      Top             =   1110
      Width           =   195
   End
End
Attribute VB_Name = "CGA2080C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name      �����и���ҵ����
'-- Program ID        CGA2080c
'-- Designer          SHIN.C.S
'-- Coder             SHIN.C.S
'-- Date              2007.7.25
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String            'Form Type
Public Toolbar_St As String          'Active Form ToolBar Setting
Public sAuthority As String          'Active Form Authority Setting
Public sDateTime As String           'Active Form Authority Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pControl2 As New Collection       'Master Primary Key Collection
Dim nControl2 As New Collection       'Master Necessary Collection
Dim mControl2 As New Collection       'Master Maxlength check Collection
Dim iControl2 As New Collection       'Master Insert Collection
Dim rControl2 As New Collection       'Master Refer Collection
Dim cControl2 As New Collection       'Master Copy Collection
Dim aControl2 As New Collection       'Master -> Spread Collection
Dim lControl2 As New Collection       'Master Lock Collection

Dim pColumn As New Collection        'Spread Primary Key Collection
Dim nColumn As New Collection        'Spread necessary Column Collection
Dim mColumn As New Collection        'Spread Maxlength check Column Collection
Dim iColumn As New Collection        'Spread Insert Column Collection
Dim aColumn As New Collection        'Master -> Spread Column Collection
Dim lColumn As New Collection        'Spread Lock Column Collection

Dim pColumn1 As New Collection       'Spread Primary Key Collection
Dim nColumn1 As New Collection       'Spread necessary Column Collection
Dim mColumn1 As New Collection       'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection       'Spread Insert Column Collection
Dim aColumn1 As New Collection       'Master -> Spread Column Collection
Dim lColumn1 As New Collection       'Spread Lock Column Collection

Dim pColumn2 As New Collection       'Spread Primary Key Collection
Dim nColumn2 As New Collection       'Spread necessary Column Collection
Dim mColumn2 As New Collection       'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection       'Spread Insert Column Collection
Dim aColumn2 As New Collection       'Master -> Spread Column Collection
Dim lColumn2 As New Collection       'Spread Lock Column Collection

Dim pColumn3 As New Collection       'Spread Primary Key Collection
Dim nColumn3 As New Collection       'Spread necessary Column Collection
Dim mColumn3 As New Collection       'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection       'Spread Insert Column Collection
Dim aColumn3 As New Collection       'Master -> Spread Column Collection
Dim lColumn3 As New Collection       'Spread Lock Column Collection

Dim pColumn4 As New Collection       'Spread Primary Key Collection
Dim nColumn4 As New Collection       'Spread necessary Column Collection
Dim mColumn4 As New Collection       'Spread Maxlength check Column Collection
Dim iColumn4 As New Collection       'Spread Insert Column Collection
Dim aColumn4 As New Collection       'Master -> Spread Column Collection
Dim lColumn4 As New Collection       'Spread Lock Column Collection

Dim Mc1 As New Collection            'Master Collection
Dim Mc2 As New Collection

Dim sc1 As New Collection            'Spread Collection
Dim sc2 As New Collection            'Spread Collection
Dim sc3 As New Collection            'Spread Collection
Dim sc4 As New Collection            'Spread Collection
Dim Proc_Sc As New Collection        'Spread Struc Collection

Dim lBlkcol1 As Long                 'To Excel Block Col1
Dim lBlkcol2 As Long                 'To Excel Block Col2
Dim lBlkrow1 As Long                 'To Excel Block Row1
Dim lBlkrow2 As Long                 'To Excel Block Row2

'DOTHER SLAB LENGTH,WGT CUALUCATE


Public cSlabno As String              'dother slab no
Public cSlabthk As Double
Public cSlabwid As Double
Public cSlabLen As Double              'Mother Slab Length
Public cSlabWgt As Double              'Mother Slab Wgt
Public cSlabCalWgt As Double           'Mother Slab Cal Wgt
Public cStlgrd As String
Public cOrdno As String
Public cProddate As String
Public cLoc As String
Public cRcvDate As String
Public tmWgt As Double
Public tmpSlabNo As String
Public NEWSLABNO As String
Public cfLen, cfWgt, cfCalWgt As Double
Public addSlabNo As String
Public lCurrRow As Long
Public SCRAP_NO As String

Dim sQuery As String

Const SS2_BLOCK_SEQ = 2
Const SS1_URGNT_FL = 25  'Add by LiQian at 2012-08-30 �Ƿ��������
Const SS3_IMP_CONT = 20
Const SS3_SLAB_NO = 1

Public Sub Form_Ins()
    If ss2.SelBlockRow2 = ss2.MaxRows Then
       ss2.ROW = ss2.MaxRows
       ss2.Col = 0
       If ss2.Text <> "Delete" Then
            Call Gp_Sp_Ins(Proc_Sc("Sc2"))
            
            With ss1
                .ROW = .ActiveRow
                .Col = 8
                .Text = sUserID
            End With
            
            Call INS_WGT_CAL
        End If
    End If

End Sub
Public Sub Spread_Del()
Dim i%
       For i = 1 To ss2.MaxRows
           ss2.ROW = i
           ss2.Col = 0
           If UCase(ss2.Text) = "" Then
              ss2.Text = "Delete"
           End If
       Next i

End Sub

Public Sub Spread_Can()

    If ss2.SelBlockRow2 = ss2.MaxRows Then
       ss2.ROW = ss2.MaxRows
       ss2.Col = 0
       If ss2.Text = "Input" Then
            ss2.MaxRows = ss2.MaxRows - 1
            addSlabNo = Mid(addSlabNo, 1, 8) & CStr(CInt(Mid(addSlabNo, 9, 2)) - 1)
            Call CANCEL_WGT_CAL
       End If
    End If
End Sub

Public Sub WGT_CAL()
Dim tmThk As Double
Dim tmWid As Double
Dim tmLen As Double
Dim tempWgt As Double
Dim tot_cal_total As Double
Dim cal_wgt As Double
Dim tmp_rat As Double
Dim tmTotalLen As Double
Dim tmpLen As Double
Dim sub_wgt As Double
Dim sub_len As Double
Dim tmCalCut As Double
Dim tmCalMo As Double
Dim tmCalCutOne As Double



Dim i As Integer

    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    tmCalMo = cSlabthk * cSlabwid * cSlabLen
    
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 0
        If ss2.Text <> "Delete" Then
            ss2.ROW = i
            ss2.Col = 2
            tmThk = ss2.Value
            ss2.Col = 3
            tmWid = ss2.Value
            ss2.Col = 4
            tmLen = ss2.Value
            tmTotalLen = tmTotalLen + ss2.Value
            
            tmCalCut = tmCalCut + (tmThk * tmWid * tmLen)
        End If
    Next i
        
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 0
        If ss2.Text <> "Delete" Then
            ss2.ROW = i
            ss2.Col = 2
            tmThk = ss2.Value
            ss2.Col = 3
            tmWid = ss2.Value
            ss2.Col = 4
            tmLen = ss2.Value
            
            tmCalCutOne = tmThk * tmWid * tmLen
            
            ss2.Col = 5
            If tmCalCut <= tmCalMo Then
                tempWgt = tempWgt + Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
                sub_wgt = sub_wgt - Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
                ss2.Value = Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
            Else
                tempWgt = tempWgt + Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
                sub_wgt = sub_wgt - Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
                ss2.Value = Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
            End If
            
            ss2.Col = 6
            ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
        End If
    Next i
    
    If tmCalCut = tmCalMo Then
        sub_len = cSlabLen
        sub_wgt = cSlabWgt
        For i = 1 To ss2.MaxRows
            ss2.ROW = i
            If i < ss2.MaxRows Then
               ss2.Col = 5
               sub_wgt = sub_wgt - ss2.Value
            End If
        Next i
        ss2.ROW = ss2.MaxRows

        ss2.Col = 5
        ss2.Text = sub_wgt
    End If
    
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 0
        If ss2.Text <> "Delete" Then
            ss2.ROW = i
            ss2.Col = 4
            tmTotalLen = tmTotalLen + ss2.Value
            
            ss2.Col = 5
            tempWgt = tempWgt + ss2.Value
        End If

    Next i
    
    For i = 1 To ss2.MaxRows
         ss2.ROW = i
         ss2.Col = 0
         If UCase(ss2.Text) = "" Then
            ss2.Text = "Update"
         End If
    Next i
    
    
    If tmTotalLen = cSlabLen Then
       txt_total_len.ForeColor = &H0&
    Else
       txt_total_len.ForeColor = &HFF&
    End If
    txt_total_len = tmTotalLen
    
    txt_total_wgt = tempWgt
    If CDbl(txt_total_wgt) - cSlabWgt = 0 Then
       txt_total_wgt.ForeColor = &H0&
    Else
       txt_total_wgt.ForeColor = &HFF&
    End If
    
    txt_scrap_wgt = Format(cSlabWgt - tempWgt, "###0.000")
    If cSlabWgt - tempWgt = 0 Then
       txt_scrap_wgt.ForeColor = &H0&
    Else
       txt_scrap_wgt.ForeColor = &HFF&
    End If

       
       
End Sub
Public Sub INS_WGT_CAL()
Dim tmThk As Double
Dim tmWid As Double
Dim tmLen As Double
Dim tempWgt As Double
Dim tot_cal_total As Double
Dim cal_wgt As Double
Dim tmp_rat As Double
Dim tmTotalLen As Double
Dim tmpLen As Double
Dim sub_wgt As Double
Dim sub_len As Double
Dim S1 As String
Dim S2 As Double
Dim S3 As Double
Dim S4 As Double
Dim S5 As Double
Dim S6 As Double
Dim S7 As String
Dim S8 As String
Dim S9 As String
Dim S10 As String
Dim S11 As String
Dim S12 As String

Dim i, delete_cnt As Integer
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&

    delete_cnt = 0
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 0
        If UCase(ss2.Text) <> "DELETE" Then
           delete_cnt = delete_cnt + 1
        End If
    Next i

    cfLen = Format(cSlabLen / ss2.MaxRows, "####0")
    cfWgt = Round(cSlabWgt / ss2.MaxRows, 3)
    cfCalWgt = Round(cSlabCalWgt / ss2.MaxRows, 3)
        
    ' DATA COPY
    ss2.ROW = ss2.MaxRows - 1
    
    ss2.Col = 1
    S1 = ss2.Value
    
    ss2.Col = 2
    S2 = ss2.Value
    
    ss2.Col = 3
    S3 = ss2.Value
    
    ss2.Col = 4
    S4 = ss2.Value
    
    ss2.Col = 5
    S5 = ss2.Value
    
    ss2.Col = 6
    S6 = ss2.Value
    
    ss2.Col = 7
    S7 = ss2.Text
    
    ss2.Col = 8
    S8 = ss2.Text
    
    ss2.Col = 9
    S9 = ss2.Text
    
    ss2.Col = 10
    S10 = ss2.Text
    
    ss2.Col = 11
    S11 = ss2.Text
    
    ss2.Col = 12
    S12 = ss2.Text

    ' DATA PAST
    ss2.ROW = ss2.MaxRows
    ss2.Col = 1
    ss2.Text = addSlabNo
    addSlabNo = Mid(addSlabNo, 1, 8) & CStr(CInt(Mid(addSlabNo, 9, 2)) + 1)
    ss2.Col = 2
    ss2.Text = S2
    tmThk = S2
    ss2.Col = 3
    ss2.Text = S3
    tmWid = S3
    ss2.Col = 4
    ss2.Text = S4
    tmLen = S4
    ss2.Col = 5
    ss2.Text = S5
    ss2.Col = 6
    ss2.Text = S6
    ss2.Col = 7
    ss2.Text = S7
    ss2.Col = 8
    ss2.Text = S8
    ss2.Col = 9
    ss2.Text = S9
    ss2.Col = 10
    ss2.Text = S10
    ss2.Col = 11
    ss2.Text = S11
    ss2.Col = 12
    ss2.Text = S12
    
    tmp_rat = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
         ss2.ROW = i
         ss2.Col = 4
         ss2.Text = cSlabLen / ss2.MaxRows
         ss2.Col = 5
         ss2.Text = Round(cSlabWgt * ((cSlabLen / ss2.MaxRows) / cSlabLen), 3)
         tempWgt = tempWgt + Round(cSlabWgt * ((cSlabLen / ss2.MaxRows) / cSlabLen), 3)
         sub_wgt = sub_wgt - Round(cSlabWgt * ((cSlabLen / ss2.MaxRows) / cSlabLen), 3)
         
         ss2.Col = 6
         ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
    Next i
    
    sub_len = cSlabLen
    sub_wgt = cSlabWgt
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        If i <> ss2.MaxRows Then
           ss2.Col = 4
           sub_len = sub_len - ss2.Value
           
           ss2.Col = 5
           sub_wgt = sub_wgt - ss2.Value
        End If
    Next i
    ss2.ROW = ss2.MaxRows
    ss2.Col = 4
    ss2.Text = sub_len
    
    ss2.Col = 5
    ss2.Text = sub_wgt
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 4
        tmTotalLen = tmTotalLen + ss2.Value
        
        ss2.Col = 5
        tempWgt = tempWgt + ss2.Value

    Next i
    
    If tmTotalLen = cSlabLen Then
       txt_total_len.ForeColor = &H0&
    Else
       txt_total_len.ForeColor = &HFF&
    End If
    txt_total_len = tmTotalLen
    
    txt_total_wgt = tempWgt
    If CDbl(txt_total_wgt) = cSlabWgt Then
       txt_total_wgt.ForeColor = &H0&
    Else
       txt_total_wgt.ForeColor = &HFF&
    End If

    
    txt_scrap_wgt = Format(cSlabWgt - tempWgt, "###0.000")
    If cSlabWgt - tempWgt = 0 Then
       txt_scrap_wgt.ForeColor = &H0&
    Else
       txt_scrap_wgt.ForeColor = &HFF&
    End If


    
    For i = 1 To ss2.MaxRows
         ss2.ROW = i
         ss2.Col = 0
         If UCase(ss2.Text) = "" Then
            ss2.Text = "Update"
         End If
    Next i
        
End Sub
Public Sub DEL_WGT_CAL()
Dim tmThk As Double
Dim tmWid As Double
Dim tmLen As Double
Dim tempWgt As Double
Dim tot_cal_total As Double
Dim cal_wgt As Double
Dim tmp_rat As Double
Dim tmTotalLen As Double
Dim tmpLen As Double
Dim sub_wgt As Double
Dim sub_len As Double


Dim i, delete_cnt As Integer
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&

    delete_cnt = 0
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 0
        If UCase(ss2.Text) <> "DELETE" Then
           delete_cnt = delete_cnt + 1
        End If
    Next i

    cfLen = Format(cSlabLen / delete_cnt, "####0")
    cfWgt = Round(cSlabWgt / delete_cnt, 3)
    cfCalWgt = Round(cSlabCalWgt / delete_cnt, 3)
    
        
    tempWgt = 0
    For i = 1 To delete_cnt
         ss2.ROW = i
         
         ss2.Col = 2
         tmThk = ss2.Value
         
         ss2.Col = 3
         tmWid = ss2.Value
         
         ss2.Col = 4
         ss2.Text = cfLen
         tmLen = cfLen
         
         ss2.Col = 4
         If ss2.ROW = ss2.MaxRows Then
            ss2.Text = cSlabLen - tmTotalLen
            tmTotalLen = tmTotalLen + ss2.Text
         Else
            ss2.Text = cfLen
            tmLen = cfLen
            tmTotalLen = tmTotalLen + cfLen
         End If
         
         ss2.Col = 5
         ss2.Text = cfWgt
         tmWgt = tmWgt + ss2.Value
         
         ss2.Col = 6
         ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
    Next i
    
    sub_len = cSlabLen
    sub_wgt = cSlabWgt
    For i = 1 To delete_cnt
        ss2.ROW = i
        If i <> delete_cnt Then
           ss2.Col = 4
           sub_len = sub_len - ss2.Value
           
           ss2.Col = 5
           sub_wgt = sub_wgt - ss2.Value
        End If
    Next i
    ss2.ROW = delete_cnt
    ss2.Col = 4
    ss2.Text = sub_len
    
    ss2.Col = 5
    ss2.Text = sub_wgt
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To delete_cnt
        ss2.ROW = i
        ss2.Col = 4
        tmTotalLen = tmTotalLen + ss2.Value
        
        ss2.Col = 5
        tempWgt = tempWgt + ss2.Value

    Next i
    
    For i = 1 To ss2.MaxRows
         ss2.ROW = i
         ss2.Col = 0
         If UCase(ss2.Text) = "" Then
            ss2.Text = "Update"
         End If
    Next i
    
    If tmTotalLen = cSlabLen Then
       txt_total_len.ForeColor = &H0&
    Else
       txt_total_len.ForeColor = &HFF&
    End If
    txt_total_len = tmTotalLen
    
    txt_total_wgt = tempWgt
    If CDbl(txt_total_wgt) = cSlabWgt Then
       txt_total_wgt.ForeColor = &H0&
    Else
       txt_total_wgt.ForeColor = &HFF&
    End If

    txt_scrap_wgt = Format(cSlabWgt - tempWgt, "###0.000")
    If cSlabWgt - tempWgt = 0 Then
       txt_scrap_wgt.ForeColor = &H0&
    Else
       txt_scrap_wgt.ForeColor = &HFF&
    End If

       

End Sub

Public Sub CANCEL_WGT_CAL()
Dim tmThk As Double
Dim tmWid As Double
Dim tmLen As Double
Dim tempWgt As Double
Dim tot_cal_total As Double
Dim cal_wgt As Double
Dim tmp_rat As Double
Dim tmTotalLen As Double
Dim tmpLen As Double
Dim sub_wgt As Double
Dim sub_len As Double
Dim i As Integer

    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    tmTotalLen = 0
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        
        ss2.Col = 4
        ss2.Text = cSlabLen / ss2.MaxRows
        tmLen = cSlabLen / ss2.MaxRows
        tmTotalLen = tmTotalLen + tmLen
        
    Next i
    
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 2
        tmThk = ss2.Value
        
        ss2.Col = 3
        tmWid = ss2.Value
        
        ss2.Col = 4
        tmLen = ss2.Value
        
        ss2.Col = 5
        ss2.Text = Round(cSlabWgt * (tmLen / tmTotalLen), 3)
        
        ss2.Col = 6
        ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
    Next i
    
    sub_len = cSlabLen
    sub_wgt = cSlabWgt
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        If i <> ss2.MaxRows Then
           ss2.Col = 4
           sub_len = sub_len - ss2.Value
           
           ss2.Col = 5
           sub_wgt = sub_wgt - ss2.Value
        End If
    Next i
    
    ss2.ROW = ss2.MaxRows
    ss2.Col = 4
    ss2.Text = sub_len
    
    ss2.Col = 5
    ss2.Text = sub_wgt
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 4
        tmTotalLen = tmTotalLen + ss2.Value
        
        ss2.Col = 5
        tempWgt = tempWgt + ss2.Value

    Next i
    If tmTotalLen = cSlabLen Then
       txt_total_len.ForeColor = &H0&
    Else
       txt_total_len.ForeColor = &HFF&
    End If
    txt_total_len = tmTotalLen
    
    txt_total_wgt = tempWgt
    If CDbl(txt_total_wgt) = cSlabWgt Then
       txt_total_wgt.ForeColor = &H0&
    Else
       txt_total_wgt.ForeColor = &HFF&
    End If

    txt_scrap_wgt = Format(cSlabWgt - tempWgt, "###0.000")
    If cSlabWgt - tempWgt = 0 Then
       txt_scrap_wgt.ForeColor = &H0&
    Else
       txt_scrap_wgt.ForeColor = &HFF&
    End If
    
End Sub

Public Sub LENMODIFY_WGT_CAL(ByVal Col As Long, ByVal ROW As Long)
Dim tmThk As Double
Dim tmWid As Double
Dim tmLen As Double
Dim tempWgt As Double
Dim tot_cal_total As Double
Dim cal_wgt As Double
Dim tmp_rat As Double
Dim tmTotalLen As Double
Dim tmpLen As Double
Dim sub_wgt As Double
Dim sub_len As Double
Dim i As Integer

Dim tmCalCut As Double
Dim tmCalMo As Double
Dim tmCalCutOne As Double

    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    tmCalMo = cSlabthk * cSlabwid * cSlabLen
    
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 0
        If UCase(ss2.Text) <> "DELETE" Then
            ss2.Col = 4
            tmTotalLen = tmTotalLen + ss2.Value
        End If
    Next i
    
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 0
        If UCase(ss2.Text) <> "DELETE" Then
            ss2.ROW = i
            ss2.Col = 2
            tmThk = ss2.Value
            ss2.Col = 3
            tmWid = ss2.Value
            ss2.Col = 4
            tmLen = ss2.Value
            
            tmCalCut = tmCalCut + (tmThk * tmWid * tmLen)
        End If
    Next i
        
    tmp_rat = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 2
        tmThk = ss2.Value
        ss2.Col = 3
        tmWid = ss2.Value
        ss2.Col = 4
        tmLen = ss2.Value
        
        tmCalCutOne = tmThk * tmWid * tmLen
        
        ss2.Col = 5
        If tmCalCut <= tmCalMo Then
            tempWgt = tempWgt + Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
            sub_wgt = sub_wgt - Round((cSlabWgt * (cfLen / cSlabLen)), 3)
            ss2.Value = Round((cSlabWgt * (tmCalCutOne / tmCalMo)), 3)
        Else
            tempWgt = tempWgt + Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
            sub_wgt = sub_wgt - Round((cSlabWgt * (cfLen / tmTotalLen)), 3)
            ss2.Value = Round((cSlabWgt * (tmCalCutOne / tmCalCut)), 3)
        End If
        
        ss2.Col = 6
        ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
    Next i
    
    
    If tmCalCut >= tmCalMo Then
        sub_len = cSlabLen
        sub_wgt = cSlabWgt
        For i = 1 To ss2.MaxRows
            ss2.ROW = i
            If i <> ss2.MaxRows Then
               ss2.Col = 5
               sub_wgt = sub_wgt - ss2.Value
            End If
        Next i
        ss2.ROW = ss2.MaxRows

        ss2.Col = 5
        ss2.Text = sub_wgt
    End If
    
    
    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 4
        tmTotalLen = tmTotalLen + ss2.Value
        
        ss2.Col = 5
        tempWgt = tempWgt + ss2.Value

    Next i
    
    If tmTotalLen = cSlabLen Then
       txt_total_len.ForeColor = &H0&
    Else
       txt_total_len.ForeColor = &HFF&
    End If
    txt_total_len = tmTotalLen
    
    txt_total_wgt = tempWgt
    If CDbl(txt_total_wgt) = cSlabWgt Then
       txt_total_wgt.ForeColor = &H0&
    Else
       txt_total_wgt.ForeColor = &HFF&
    End If

    txt_scrap_wgt = Format(cSlabWgt - tempWgt, "###0.000")
    If cSlabWgt - tempWgt = 0 Then
       txt_scrap_wgt.ForeColor = &H0&
    Else
       txt_scrap_wgt.ForeColor = &HFF&
    End If
    
       
       

End Sub

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(txt_Status, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_act_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_MOSLAB, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_loc, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_plt_dec, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_thk, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_THK_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_wid, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_WID_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_len, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_LEN_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(U_FROM_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(U_TO_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)   '�������
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) 'Add by LiQian at 2012-08-30 �Ƿ��������
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGA2080C.P_REFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    
    Call Gp_Ms_Collection(txt_tmpPLT, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  Call Gp_Ms_Collection(txt_IST_DATE, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(txt_SlabNo, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
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
    Call Gp_Sp_Collection(ss2, 1, "p", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CGA2080C.P_MODIFY1", Key:="P-M"
    sc2.Add Item:="CGA2080C.P_REFER1", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 13, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 14, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 15, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 16, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 17, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 18, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 19, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 20, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    
    'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="CGA2080C.P_REFER2", Key:="P-R"
    sc3.Add Item:=pColumn3, Key:="pColumn"
    sc3.Add Item:=nColumn3, Key:="nColumn"
    sc3.Add Item:=aColumn3, Key:="aColumn"
    sc3.Add Item:=mColumn3, Key:="mColumn"
    sc3.Add Item:=iColumn3, Key:="iColumn"
    sc3.Add Item:=lColumn3, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss4, 1, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 3, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 4, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 5, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 6, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 7, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 8, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 9, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 10, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 11, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 12, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 13, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 14, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 15, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    
    'Spread_Collection
    sc4.Add Item:=ss4, Key:="Spread"
    sc4.Add Item:="CGA2080C.P_REFER3", Key:="P-R"
    sc4.Add Item:=pColumn4, Key:="pColumn"
    sc4.Add Item:=nColumn4, Key:="nColumn"
    sc4.Add Item:=aColumn4, Key:="aColumn"
    sc4.Add Item:=mColumn4, Key:="mColumn"
    sc4.Add Item:=iColumn4, Key:="iColumn"
    sc4.Add Item:=lColumn4, Key:="lColumn"
    sc4.Add Item:=1, Key:="First"
    sc4.Add Item:=ss4.MaxCols, Key:="Last"
    
    Call Gp_Sp_ColHidden(ss2, 6, True)
    'Call Gp_Sp_ColHidden(ss2, 10, True)
    Call Gp_Sp_ColHidden(ss2, 12, True)
    Call Gp_Sp_ColHidden(ss2, 13, True)

    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub


Private Sub cbo_cutcnt_Click()
Dim i, j As Integer

Dim tmThk, tmWid, tmLen As Double
Dim tmTotalLen As Double
Dim tmpLen As Double

    
Dim tempWgt As Double
Dim tot_cal_total As Double
Dim cal_wgt As Double
Dim sub_wgt As Double
Dim tmp_rat As Double

    If txt_SlabNo.Text = "" Then Exit Sub
    
    If cbo_cutcnt.ListIndex = 0 Then Exit Sub
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    ss2.MaxRows = 0
    ss2.MaxRows = CInt(cbo_cutcnt)
    
    sQuery = "          SELECT MAX(SLAB_NO) "
    sQuery = sQuery & "   FROM NISCO.FP_SLAB "
    sQuery = sQuery & "  WHERE SLAB_NO LIKE '" & Mid(cSlabno, 1, 8) & "%'"
    
    tmpSlabNo = Gf_CodeFind(M_CN1, sQuery)
    If CInt(Mid(tmpSlabNo, 9, 2)) < 30 Then
       tmpSlabNo = Mid(tmpSlabNo, 1, 8) & "30"
    End If
    

    cfLen = Format(cSlabLen / cbo_cutcnt, "####0")
    cfWgt = Round(cSlabWgt / cbo_cutcnt, 3)
    cfCalWgt = Round(cSlabCalWgt / cbo_cutcnt, 3)
    
    For i = 1 To cbo_cutcnt
        ss2.ROW = i
        ss2.Col = 1
        
        NEWSLABNO = Mid(tmpSlabNo, 1, 4) & Mid(tmpSlabNo, 5, 6) + i
        If Len(Mid(NEWSLABNO, 5, 6)) = 5 Then
           NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "0" & Mid(NEWSLABNO, 5, 5)
        ElseIf Len(Mid(NEWSLABNO, 5, 6)) = 4 Then
           NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "00" & Mid(NEWSLABNO, 5, 5)
        ElseIf Len(Mid(NEWSLABNO, 5, 6)) = 3 Then
           NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "000" & Mid(NEWSLABNO, 5, 5)
        End If
        
        ss2.Text = NEWSLABNO
    
        ss2.Col = 2
        ss2.Text = cSlabthk
        tmThk = cSlabthk
    
        ss2.Col = 3
        ss2.Text = cSlabwid
        tmWid = cSlabwid
    
        ss2.Col = 4
        If ss2.ROW = ss2.MaxRows Then
            ss2.Text = cSlabLen - tmTotalLen
            tmTotalLen = tmTotalLen + ss2.Text
        Else
            ss2.Text = cfLen
            tmLen = cfLen
            tmTotalLen = tmTotalLen + cfLen
        End If
        
        ss2.Col = 5
        ss2.Text = cfWgt
        tmWgt = tmWgt + ss2.Value
    
        ss2.Col = 6
        ss2.Text = cfCalWgt
    
        ss2.Col = 7
        ss2.Text = Format(Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD') FROM DUAL"), "YYYY-MM-DD")
    
        ss2.Col = 8
        ss2.Text = Format(Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'HH24:MI') FROM DUAL"), "HH:MM")
    
        
        ss2.Col = 9
        ss2.Text = cLoc
        
        ss2.Col = 10
        ss2.Text = sUserID
        
        ss2.Col = 11
        ss2.Text = txt_SlabNo
        
        ss2.Col = 12
        If i = cbo_cutcnt Then
            ss2.Text = "Y"
        Else
            ss2.Text = ""
        End If
        
        ss2.Col = 0
        ss2.ROW = i
        ss2.Text = "Input"
    
    Next i
    SCRAP_NO = txt_SlabNo
    
    Call WGT_CAL
    
    MDIMain.MenuTool.Buttons(1).Enabled = True                 'Save
    MDIMain.MenuTool.Buttons(2).Enabled = True                 'Delete
    MDIMain.MenuTool.Buttons(4).Enabled = True                 'Separator
    MDIMain.MenuTool.Buttons(14).Enabled = True                  'Row Delete
    
    MDIMain.MenuTool.Buttons(5).Enabled = False                 'Save
    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Row Insert
    MDIMain.MenuTool.Buttons(9).Enabled = False                 'Separator
    MDIMain.MenuTool.Buttons(11).Enabled = False                 'Row Insert
    MDIMain.MenuTool.Buttons(12).Enabled = False                 'Row Delete
    MDIMain.MenuTool.Buttons(15).Enabled = False                 'Row Delete
    
End Sub

Private Sub Cmd_Cancel_Click()
    Dim OutParam(2, 4) As Variant
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
    
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
    
    sQuery = "{call CGA2080C.P_ORDCANCEL('" & Trim(txt_SlabNo.Text) & "','" & sUserID & "',?,?)}"
    
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
   


    MDIMain.MenuTool.Buttons(14).Enabled = True                 'Save
'    MDIMain.MenuTool.Buttons(2).Enabled = True                 'Delete
'    MDIMain.MenuTool.Buttons(4).Enabled = True                 'Separator
'    MDIMain.MenuTool.Buttons(14).Enabled = True                  'Row Delete
'
'    MDIMain.MenuTool.Buttons(5).Enabled = False                 'Save
'    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Delete
'    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Row Insert
'    MDIMain.MenuTool.Buttons(9).Enabled = False                 'Separator
'    MDIMain.MenuTool.Buttons(11).Enabled = False                 'Row Insert
'    MDIMain.MenuTool.Buttons(12).Enabled = False                 'Row Delete
'    MDIMain.MenuTool.Buttons(15).Enabled = False                 'Row Delete

   
End Sub

Private Sub Form_Load()
Dim sQuery As String
    
    Dim i, j As Integer
    
    Screen.MousePointer = vbHourglass
    
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Screen.MousePointer = vbDefault
    
    cbo_cutcnt.AddItem "0"
    cbo_cutcnt.AddItem "1"
    cbo_cutcnt.AddItem "2"
    cbo_cutcnt.AddItem "3"
    cbo_cutcnt.AddItem "4"
    cbo_cutcnt.AddItem "5"
    cbo_cutcnt.AddItem "6"
    cbo_cutcnt.AddItem "7"
    cbo_cutcnt.AddItem "8"
    cbo_cutcnt.AddItem "9"
    cbo_cutcnt.AddItem "10"
    
    Call opt_prc_status1_click
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    Call Gp_Sp_Setting(sc3.Item("Spread"))
    Call Gp_Sp_Setting(sc4.Item("Spread"))

    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(sc3)
    Call Gf_Sp_Cls(sc4)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc3.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc4.Item("Spread"), "CG-System.INI", Me.Name)
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    txt_plt = "B3"
    txt_plt_dec = "�����ֳ�"
    txt_thk = 150
    TXT_THK_TO = 320
    txt_wid = 1000
    TXT_WID_TO = 4000
    txt_len = 1000
    TXT_LEN_TO = 99999
    
    U_FROM_DATE.RawData = Format(Now, "YYYYMM") + "01"
    
    If Mid(sAuthority, 1, 3) = "111" Then
       cmd_cancel.Enabled = True
    Else
       cmd_cancel.Enabled = False
    End If
    
   
'    MDIMain.MenuTool.Buttons(1).Enabled = True                 'Save
'    MDIMain.MenuTool.Buttons(2).Enabled = True                 'Delete
'    MDIMain.MenuTool.Buttons(4).Enabled = True                 'Separator
'    MDIMain.MenuTool.Buttons(8).Enabled = True                 'Row Insert
'    MDIMain.MenuTool.Buttons(14).Enabled = True                  'Row Delete
'
'    MDIMain.MenuTool.Buttons(5).Enabled = False                 'Save
'    MDIMain.MenuTool.Buttons(7).Enabled = True                 'Delete
'    MDIMain.MenuTool.Buttons(9).Enabled = True                 'Separator
'    MDIMain.MenuTool.Buttons(11).Enabled = False                 'Row Insert
'    MDIMain.MenuTool.Buttons(12).Enabled = False                 'Row Delete
'    MDIMain.MenuTool.Buttons(15).Enabled = False                 'Row Delete

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc3.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc4.Item("Spread"), "CG-System.INI", Me.Name)
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
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

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(sc3)
    Call Gf_Sp_Cls(sc4)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    txt_act_stlgrd_dec = ""
    txt_SlabNo.Text = ""
    cbo_cutcnt.ListIndex = 0
    txt_total_len.Value = 0
    txt_total_wgt.Value = 0
    txt_scrap_wgt.Value = 0
    
    Call opt_prc_status1_click
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
        


End Sub

Public Sub Form_Pro()
    Dim i As Integer
    
    If SSTab1.Tab = 0 Then
    
    'ADD BY YIDUJUN AT 2010-12-21 ����Ӱ����ţ�Ϊ������ʾ����
        For i = 1 To ss2.MaxRows
            ss2.ROW = i
            ss2.Col = 5
            If ss2.Value < 0 Then
                ss2.Col = 1
                MsgBox "��ȷ���Ӱ����� " + ss2.Text + " ������", vbCritical, "������ʾ"
                Exit Sub
            End If

        Next i
       
        Call Gf_Sp_Process(M_CN1, Proc_Sc("Sc2"), Mc1, True)
    '    Call Scrap_Pro   '''COMMENT BY GUOLI AT 20081026
        
        If opt_prc_status1.Value = True Then
             Call Form_Ref
        End If
        If opt_prc_status2.Value = True Then
             If ss2.MaxRows < 1 Then
                Call Form_Ref
             End If
        End If
    Else
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    
    MDIMain.MenuTool.Buttons(1).Enabled = True                   'Save
    MDIMain.MenuTool.Buttons(2).Enabled = True                   'Delete
    MDIMain.MenuTool.Buttons(4).Enabled = True                   'Separator
    MDIMain.MenuTool.Buttons(8).Enabled = True                   'Row Insert
    MDIMain.MenuTool.Buttons(14).Enabled = True                  'Row Delete
    
    MDIMain.MenuTool.Buttons(5).Enabled = False                  'Save
    MDIMain.MenuTool.Buttons(7).Enabled = True                   'Delete
    MDIMain.MenuTool.Buttons(9).Enabled = True                   'Separator
    MDIMain.MenuTool.Buttons(11).Enabled = False                 'Row Insert
    MDIMain.MenuTool.Buttons(12).Enabled = False                 'Row Delete
    MDIMain.MenuTool.Buttons(15).Enabled = False                 'Row Delete


End Sub
Public Sub Scrap_Pro()
    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iType As String
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adInteger
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    If txt_scrap_wgt.Value = 0 Then
        iType = "D"
    Else
        iType = "I"
    End If
    
    sQuery = "{call CGA2080C.P_SCRAP ('" & iType & "','" & SCRAP_NO & "'," & txt_scrap_wgt.Value & ",'" & sUserID & "',?,?)}"
    
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
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Exit Sub
        
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault

End Sub

Public Sub Form_Ref()

    Dim ForCnt As Integer
    Dim tmWgt As Long
    Dim tmLen As Long
    
    Dim lRow As Long
    Dim sBlockSeq As String
    Dim iRow As Integer
    Dim i As Integer
    Dim TIME As String
    
    Dim sUrgnt_Fl As String
    Dim simpcont As String
    
    If Not Gf_Sp_Cls(sc2) Then Exit Sub
    
   
    
    If Len(Trim(txt_MOSLAB)) <> 0 Then
        If Len(Trim(txt_MOSLAB)) < 8 Then
           MsgBox "��ȷ��ĸ������"
           txt_MOSLAB.SetFocus
           Exit Sub
        End If
    End If
    
    If Len(Trim(txt_MOSLAB)) <> 8 Then
        If Len(Trim(txt_plt)) = 0 Then
            MsgBox "��ȷ�Ϲ���"
            txt_plt.SetFocus
            Exit Sub
        End If
    End If
    
    If txt_len.Value <= 0 Then
        MsgBox "��ȷ�ϳ��ȷ�Χ"
        txt_len.SetFocus
        Exit Sub
    End If
    
    txt_SlabNo.Text = ""
    cbo_cutcnt.ListIndex = 0
    txt_total_len.Value = 0
    txt_total_wgt.Value = 0
    txt_scrap_wgt.Value = 0
    
    If SSTab1.Tab = 0 Then
    
        Call Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"))
        ss1.OperationMode = OperationModeNormal
        Call Gf_Sp_Cls(sc3)
        Call Gf_Sp_Cls(sc4)
        
    ElseIf SSTab1.Tab = 1 Then
    
        If txt_Status.Text = "2" Then
             Call Gp_MsgBoxDisplay("����ȷѡ��������", "I")
        Exit Sub
        End If
        
        Call Gf_Sp_Refer(M_CN1, sc3, Mc1, Mc1("nControl"), Mc1("mControl"))
        ss3.OperationMode = OperationModeNormal
'        ss1.MaxRows = 0
'        ss2.MaxRows = 0
        Call Gf_Sp_Cls(sc1)
        Call Gf_Sp_Cls(sc2)
        
        If ss3.MaxRows > 0 Then
        
            For lRow = 1 To ss3.MaxRows
            
                ss3.ROW = lRow
                ss3.Col = SS2_BLOCK_SEQ: sBlockSeq = ss3.Text
                ss3.Col = SS3_IMP_CONT: simpcont = ss3.Text
                
                If sBlockSeq = "00" Then
                    Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, ss3.ROW, ss3.ROW, , SSP1.BackColor)
                Else
                    Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, ss3.ROW, ss3.ROW, , SSP3.BackColor)
                End If
                '�ص㶩��
                If simpcont = "Y" Then
                    Call Gp_Sp_BlockColor(ss3, SS3_SLAB_NO, SS3_SLAB_NO, lRow, lRow, SSP4.BackColor)
                    Call Gp_Sp_BlockColor(ss3, SS3_IMP_CONT, SS3_IMP_CONT, lRow, lRow, SSP4.BackColor)
                End If
            Next lRow
            
        End If
        
    ElseIf SSTab1.Tab = 2 Then
    
        If txt_Status.Text = "2" Then
             Call Gp_MsgBoxDisplay("����ȷѡ��������", "I")
        Exit Sub
        End If
        
        Call Gf_Sp_Refer(M_CN1, sc4, Mc1, Mc1("nControl"), Mc1("mControl"))
        ss4.OperationMode = OperationModeNormal
        Call Gf_Sp_Cls(sc1)
        Call Gf_Sp_Cls(sc2)
        
    End If
    
    
    
     TIME = Format(Now, "YYYY-MM")
 
     
     For iRow = 1 To ss1.MaxRows
    
      ss1.ROW = iRow
      ss1.Col = 23
        If Mid(ss1.Text, 1, 7) < TIME Then
          For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.BackColor = &HFF&
          Next
        End If

        If ss1.Text = "" Then
           Exit For
        End If
        
      ss1.Col = SS1_URGNT_FL:     sUrgnt_Fl = Trim(ss1.Text)
      
      '����������ɫ��ʾ add by liqian 2012-08-30
        If sUrgnt_Fl = "Y" Then
           Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, &HC000&)
        End If
        
    Next iRow
    
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    If opt_prc_status1.Value = True Then
        
      MDIMain.MenuTool.Buttons(7).Enabled = False                 'Save
      MDIMain.MenuTool.Buttons(8).Enabled = False                 'Delete
      MDIMain.MenuTool.Buttons(9).Enabled = False                'Separator
  
   End If
    
End Sub

Private Sub opt_prc_status1_click()

If opt_prc_status1.Value = True Then
   cmd_cancel.Visible = True
End If
    If opt_prc_status1.Tag <> "" Then
       opt_prc_status1.Tag = ""
       Exit Sub
    End If
    
    If Gf_Sp_Cls(sc2) = False Then
        opt_prc_status2.Tag = "A"
        opt_prc_status2.Value = True
        txt_Status.Text = "2"
        Exit Sub
    End If
    opt_prc_status1.Value = True
    
    opt_prc_status1.ForeColor = &HFF&
    opt_prc_status2.ForeColor = &H80000011
    
    Call Gf_Sp_Cls(sc1)
    
    txt_Status.Text = "1"
    
    txt_act_stlgrd_dec = ""
    txt_MOSLAB.Text = ""
    txt_SlabNo.Text = ""
    
    'cbo_cutcnt.Enabled = True
    
    cbo_cutcnt.Clear
    cbo_cutcnt.AddItem "0"
    cbo_cutcnt.AddItem "1"
    cbo_cutcnt.AddItem "2"
    cbo_cutcnt.AddItem "3"
    cbo_cutcnt.AddItem "4"
    cbo_cutcnt.AddItem "5"
    cbo_cutcnt.AddItem "6"
    cbo_cutcnt.AddItem "7"
    cbo_cutcnt.AddItem "8"
    cbo_cutcnt.AddItem "9"
    cbo_cutcnt.AddItem "10"
    
    cbo_cutcnt.ListIndex = 0
    txt_total_len.Value = 0
    txt_total_wgt.Value = 0
    txt_scrap_wgt.Value = 0
    
    txt_plt = "B3"
    txt_plt_dec = "�����ֳ�"
    txt_thk = 150
    TXT_THK_TO = 320
    txt_wid = 1000
    TXT_WID_TO = 4000
    txt_len = 1000
    TXT_LEN_TO = 99999
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False                 'Separator

End Sub

Private Sub opt_prc_status2_Click()
If opt_prc_status2.Value = True Then
    cmd_cancel.Visible = False
End If
    If opt_prc_status2.Tag <> "" Then
       opt_prc_status2.Tag = ""
       Exit Sub
    End If
    
    If Gf_Sp_Cls(sc2) = False Then
        opt_prc_status1.Tag = "A"
        opt_prc_status1.Value = True
        txt_Status.Text = "1"
        Exit Sub
    End If
    
    opt_prc_status2.Value = True
    
    opt_prc_status2.ForeColor = &HFF&
    opt_prc_status1.ForeColor = &H80000011
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    txt_Status.Text = "2"
    
    txt_act_stlgrd_dec = ""
    txt_MOSLAB.Text = ""
    txt_SlabNo.Text = ""
    
    cbo_cutcnt.Enabled = False
    cbo_cutcnt.ListIndex = 0
    txt_total_len.Value = 0
    txt_total_wgt.Value = 0
    txt_scrap_wgt.Value = 0
    
    txt_plt = "B3"
    txt_plt_dec = "�����ֳ�"
    txt_thk = 150
    TXT_THK_TO = 320
    txt_wid = 1000
    TXT_WID_TO = 4000
    txt_len = 1000
    TXT_LEN_TO = 99999
    
    U_FROM_DATE.RawData = Format(Now, "YYYYMM") + "01"
    
    txt_total_len.ForeColor = &H0&
    txt_total_wgt.ForeColor = &H0&
    txt_scrap_wgt.ForeColor = &H0&
    
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    MDIMain.MenuTool.Buttons(7).Enabled = True                 'Delete
    MDIMain.MenuTool.Buttons(8).Enabled = True                 'Delete
    MDIMain.MenuTool.Buttons(9).Enabled = True                 'Separator
    
    'Call Form_Cls
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
'Dim cSlabLen As Long

Dim iRow1, iRow2, iCol   As Integer
Dim sColor, sHeat, sTemp As String
Dim sChgPrcLine          As String
Dim sL2SendFL            As String
Dim i                    As Integer
Dim ForCnt               As Integer
Dim tmLen                As Double
Dim tmWgt                As Double
Dim TIME As String
    
Dim tmThk As Double
Dim tmWid As Double

Dim tempWgt As Double
Dim tot_cal_total As Double
Dim cal_wgt As Double
Dim tmp_rat As Double
Dim tmTotalLen As Double
Dim tmpLen As Double
Dim sub_wgt As Double
Dim sub_len As Double
Dim tmCalCut As Double
Dim tmCalMo As Double
Dim tmCalCutOne As Double
Dim iRow As Integer


    


    
    
  
    
    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW, "&H00000000", "&HFFFF80")
      
    For i = 1 To ss1.MaxRows
        If i <> ROW Then
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, i, i)
            
            
        End If
    Next
    
     TIME = Format(Now, "YYYY-MM")

     For iRow = 1 To ss1.MaxRows

      ss1.ROW = iRow
      ss1.Col = 24
        If Mid(ss1.Text, 1, 7) < TIME Then
          For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.BackColor = &HFF&
          Next
        End If

        If ss1.Text = "" Then
           Exit For
        End If

    Next iRow
    
    If ROW <> 0 Then

        ss1.Col = 1
        ss1.ROW = ROW
        txt_SlabNo = ss1.Text
        SCRAP_NO = ss1.Text
        cSlabno = ss1.Text
        ss1.Col = 8
        cSlabLen = ss1.Value
        ss1.Col = 9
        cSlabWgt = ss1.Value
        ss1.Col = 17
        txt_tmpPLT = ss1.Value
        ss1.Col = 18
        txt_IST_DATE = ss1.Value
    End If

    sQuery = "          SELECT MAX(SLAB_NO) "
    sQuery = sQuery & "   FROM NISCO.FP_SLAB "
    sQuery = sQuery & "  WHERE SLAB_NO LIKE '" & Mid(cSlabno, 1, 8) & "%'"
    
    tmpSlabNo = Gf_CodeFind(M_CN1, sQuery)
    If CInt(Mid(tmpSlabNo, 9, 2)) < 30 Or CInt(Mid(tmpSlabNo, 9, 2)) >= 97 Then  'modified by guoli at 20080418
       tmpSlabNo = Mid(tmpSlabNo, 1, 8) & "30"
    End If
    
    ss1.ROW = ROW
    ss1.Col = 1

    lBlkrow1 = ROW
    lBlkrow2 = ROW
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").ROW = 0
    sc1.Item("Spread").Text = "��"
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").ROW = 0
    sc2.Item("Spread").Text = ""

    If ROW = 0 Then Exit Sub

    If ROW = 0 Then Call Gp_Sp_Sort(sc1.Item("Spread"), Col, ROW)
    
    If opt_prc_status2 Then
        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2, Nothing, Mc2("mControl"), False)
        For i = 1 To ss2.MaxRows
             ss2.ROW = i
             ss2.Col = 15
             If Trim(ss2.Text) = "������" Then
                Call Gp_Sp_BlockLock(ss2, 3, 4, i, i)
             End If
        Next i
        Exit Sub
    End If
    

    Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2, Nothing, Mc2("mControl"), False)


    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 1
        
        NEWSLABNO = Mid(tmpSlabNo, 1, 8) & Mid(tmpSlabNo, 9, 2) + i
        If Len(Mid(NEWSLABNO, 5, 6)) = 5 Then
           NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "0" & Mid(NEWSLABNO, 5, 5)
        ElseIf Len(Mid(NEWSLABNO, 5, 6)) = 4 Then
           NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "00" & Mid(NEWSLABNO, 5, 5)
        ElseIf Len(Mid(NEWSLABNO, 5, 6)) = 3 Then
           NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "000" & Mid(NEWSLABNO, 5, 5)
        End If
        
        ss2.Text = NEWSLABNO
        
        ss2.Col = 2
        tmThk = ss2.Value
        
        ss2.Col = 3
        tmWid = ss2.Value
        
        ss2.Col = 4
        tmLen = ss2.Value
            
        ss2.Col = 6
        ss2.Text = ((tmThk * tmWid * tmLen) * 7.85) / 1000000000
        
        ss2.Col = 11
        ss2.Text = sUserID
        
        ss2.Col = 12
        ss2.Text = txt_SlabNo
        
        ss2.Col = 13
        If i = ss2.MaxRows Then
            ss2.Text = "Y"
        Else
            ss2.Text = ""
        End If
    Next

    tmTotalLen = 0
    tempWgt = 0
    For i = 1 To ss2.MaxRows
        ss2.ROW = i
        ss2.Col = 0
        If ss2.Text <> "Delete" Then
            ss2.ROW = i
            ss2.Col = 4
            tmTotalLen = tmTotalLen + ss2.Value
            
            ss2.Col = 5
            tempWgt = tempWgt + ss2.Value
        End If

    Next i
    
    If txt_Status = "1" Then
        For i = 1 To ss2.MaxRows
             ss2.ROW = i
             ss2.Col = 0
             If UCase(ss2.Text) = "" Then
                ss2.Text = "Input"
             End If
             ss2.Col = 15
             If Trim(ss2.Text) = "������" Then
                Call Gp_Sp_BlockLock(ss2, 3, 4, i, i)
             End If
             
        Next i
    End If
    
    If tmTotalLen = cSlabLen Then
       txt_total_len.ForeColor = &H0&
    Else
       txt_total_len.ForeColor = &HFF&
    End If
    txt_total_len = tmTotalLen
    
    txt_total_wgt = tempWgt
    If CDbl(txt_total_wgt) - cSlabWgt = 0 Then
       txt_total_wgt.ForeColor = &H0&
    Else
       txt_total_wgt.ForeColor = &HFF&
    End If
    
    txt_scrap_wgt = Format(cSlabWgt - tempWgt, "###0.000")
    If cSlabWgt - tempWgt = 0 Then
       txt_scrap_wgt.ForeColor = &H0&
    Else
       txt_scrap_wgt.ForeColor = &HFF&
    End If
End Sub


Public Sub Form_Exc()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If SSTab1.Tab = 0 Then
       Call Gp_Sp_Excel(Me, ss1, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf SSTab1.Tab = 1 Then
       Call Gp_Sp_Excel(Me, ss3, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf SSTab1.Tab = 2 Then
       Call Gp_Sp_Excel(Me, ss4, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If

End Sub



Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If
End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
Dim tmThk As Double
Dim tmWid As Double
Dim tmLen As Double
Dim tempWgt As Double
Dim i As Integer


    If Gf_Sc_Authority(sAuthority, "U") Then Call Gp_Sp_UpdateMake(Proc_Sc("SC2")("Spread"), Mode)
    
    
    If Col <> 2 And Col <> 3 And Col <> 4 Then Exit Sub
    
    If ChangeMade Then
        Call WGT_CAL
    End If
    
End Sub


Private Sub txt_act_stlgrd_Change()
    If Len(Trim(txt_act_stlgrd.Text)) = 0 Then txt_act_stlgrd_dec.Text = ""
End Sub

Private Sub txt_act_stlgrd_DblClick()
    Call txt_act_stlgrd_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_act_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        'txt_act_stlgrd.Text = ""
        DD.rControl.Add Item:=txt_act_stlgrd
        DD.rControl.Add Item:=txt_act_stlgrd_dec

        Call Gf_Stlgrd_DD(M_CN1, vbKeyF4)

        Exit Sub
    End If
End Sub

Private Sub txt_plt_DblClick()
    Call txt_plt_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=txt_plt_dec

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

    If Len(Trim(txt_plt)) = txt_plt.MaxLength Then
        txt_plt_dec.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        txt_plt_dec.Text = ""
    End If
End Sub