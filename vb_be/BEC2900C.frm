VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form BEC2900C 
   Caption         =   "������λ���ƽ���޸�_BEC2900C"
   ClientHeight    =   7650
   ClientLeft      =   315
   ClientTop       =   2100
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_color_len 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   10455
      MaxLength       =   50
      TabIndex        =   44
      Tag             =   "����"
      Top             =   60
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt_color 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   10005
      MaxLength       =   50
      TabIndex        =   43
      Tag             =   "����"
      Top             =   60
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txt_to_no 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   8790
      MaxLength       =   50
      TabIndex        =   42
      Tag             =   "����"
      Top             =   60
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox txt_target_no 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   9480
      MaxLength       =   50
      TabIndex        =   41
      Tag             =   "����"
      Top             =   60
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox txt_from_no 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   8220
      MaxLength       =   50
      TabIndex        =   40
      Tag             =   "����"
      Top             =   60
      Visible         =   0   'False
      Width           =   1200
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   4110
      Left            =   90
      TabIndex        =   9
      Top             =   900
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   7250
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin InDate.ULabel lbl_wid 
         Height          =   45
         Index           =   0
         Left            =   405
         Top             =   1530
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   79
         Caption         =   ""
         Alignment       =   1
         BackColor       =   255
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         LargeChange     =   50
         Left            =   45
         SmallChange     =   10
         TabIndex        =   10
         Top             =   3870
         Width           =   15045
      End
      Begin InDate.ULabel lbl_thk 
         Height          =   30
         Index           =   0
         Left            =   405
         Top             =   1590
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   53
         Caption         =   ""
         Alignment       =   1
         BackColor       =   16776960
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel lbl_len 
         Height          =   30
         Index           =   0
         Left            =   405
         Top             =   3735
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   53
         Caption         =   ""
         Alignment       =   1
         BackColor       =   16761087
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WCR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   14715
         TabIndex        =   22
         Top             =   60
         Width           =   285
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00000000&
         BorderColor     =   &H00000000&
         FillColor       =   &H0080C0FF&
         FillStyle       =   0  'Solid
         Height          =   90
         Left            =   14490
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   195
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "8,600"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   150
         Left            =   90
         TabIndex        =   21
         Top             =   3105
         Width           =   510
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFC0FF&
         FillStyle       =   0  'Solid
         Height          =   90
         Left            =   14175
         Shape           =   4  'Rounded Rectangle
         Top             =   2475
         Width           =   195
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   14400
         TabIndex        =   20
         Top             =   2445
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   14850
         TabIndex        =   19
         Top             =   2445
         Width           =   150
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFC0&
         FillStyle       =   0  'Solid
         Height          =   90
         Left            =   14625
         Shape           =   4  'Rounded Rectangle
         Top             =   2475
         Width           =   195
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "����(mm)"
         Height          =   915
         Left            =   45
         TabIndex        =   18
         Top             =   3420
         Width           =   375
      End
      Begin VB.Line Line5 
         X1              =   450
         X2              =   450
         Y1              =   2610
         Y2              =   3825
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   450
         X2              =   14985
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   225
         X2              =   14985
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BorderColor     =   &H00000000&
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   90
         Left            =   13950
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CCR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   14175
         TabIndex        =   15
         Top             =   60
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HCR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   13635
         TabIndex        =   14
         Top             =   60
         Width           =   270
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BorderColor     =   &H00000000&
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   90
         Left            =   13410
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   195
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   225
         X2              =   14985
         Y1              =   1575
         Y2              =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "���(mm)"
         Height          =   420
         Left            =   45
         TabIndex        =   12
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "����(mm)"
         Height          =   780
         Left            =   45
         TabIndex        =   11
         Top             =   90
         Width           =   375
      End
      Begin VB.Line Line2 
         X1              =   450
         X2              =   450
         Y1              =   90
         Y2              =   2520
      End
   End
   Begin FPSpread.vaSpread ss3 
      Height          =   1230
      Left            =   7845
      TabIndex        =   6
      Top             =   9255
      Visible         =   0   'False
      Width           =   4650
      _Version        =   393216
      _ExtentX        =   8202
      _ExtentY        =   2170
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
      MaxCols         =   7
      MaxRows         =   0
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "BEC2900C.frx":0000
   End
   Begin VB.TextBox txt_roll_mana_no 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8040
      MaxLength       =   5
      TabIndex        =   3
      Tag             =   "HEAT_MANA_NO"
      Top             =   90
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox txt_heat_mana_no 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   5640
      MaxLength       =   8
      TabIndex        =   2
      Tag             =   "¯�ι�����"
      Top             =   75
      Width           =   1140
   End
   Begin VB.TextBox txt_plt 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1350
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "����"
      Top             =   65
      Width           =   375
   End
   Begin VB.TextBox txt_plt_name 
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
      Left            =   1725
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "����"
      Top             =   60
      Width           =   2085
   End
   Begin SSSplitter.SSSplitter spl_splitter 
      Height          =   4215
      Left            =   135
      TabIndex        =   5
      Top             =   5025
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   7435
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "BEC2900C.frx":02E1
      Begin FPSpread.vaSpread ss1 
         Height          =   4215
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   8970
         _Version        =   393216
         _ExtentX        =   15822
         _ExtentY        =   7435
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
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "BEC2900C.frx":0333
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   4215
         Left            =   9030
         TabIndex        =   8
         Top             =   0
         Width           =   6060
         _Version        =   393216
         _ExtentX        =   10689
         _ExtentY        =   7435
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
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "BEC2900C.frx":0CD9
      End
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_edt_seq 
      Height          =   225
      Left            =   6390
      TabIndex        =   4
      Tag             =   "HEAT_EDT_SEQ"
      Top             =   150
      Visible         =   0   'False
      Width           =   150
      _Version        =   262145
      _ExtentX        =   265
      _ExtentY        =   397
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   16
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   90
      Top             =   60
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "����"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   4410
      Top             =   60
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "¯�ι�����"
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
   Begin Threed.SSCommand cmd_process 
      Height          =   405
      Left            =   13800
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   450
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
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
      Caption         =   "����"
   End
   Begin Threed.SSCommand cmd_roll1 
      Height          =   405
      Left            =   11040
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   15
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
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
      Caption         =   "����������λ"
   End
   Begin Threed.SSCommand cmd_confirm 
      Height          =   405
      Left            =   13800
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   15
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16576
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
      Caption         =   "ָʾȷ��"
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   420
      Left            =   2295
      TabIndex        =   23
      Top             =   435
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   741
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
      Begin Threed.SSOption opt_move 
         Height          =   285
         Left            =   360
         TabIndex        =   24
         Top             =   90
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   "�ƶ�"
         Value           =   -1
      End
      Begin Threed.SSOption opt_split 
         Height          =   285
         Left            =   2265
         TabIndex        =   25
         Top             =   90
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
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
         Caption         =   "�ֿ�"
      End
      Begin Threed.SSOption opt_unification 
         Height          =   285
         Left            =   1320
         TabIndex        =   26
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
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
         Caption         =   "ͳ��"
      End
      Begin Threed.SSOption opt_delete 
         Height          =   285
         Left            =   3210
         TabIndex        =   27
         Top             =   90
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
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
         Caption         =   "ɾ��"
      End
   End
   Begin Threed.SSPanel pnl_first 
      Height          =   420
      Left            =   6615
      TabIndex        =   28
      Top             =   435
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   741
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
      Begin VB.TextBox txt_from 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   825
         MaxLength       =   50
         TabIndex        =   31
         Tag             =   "����"
         Top             =   75
         Width           =   1200
      End
      Begin VB.TextBox txt_to 
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
         Height          =   310
         Left            =   2190
         MaxLength       =   50
         TabIndex        =   30
         Tag             =   "����"
         Top             =   75
         Width           =   1200
      End
      Begin VB.TextBox txt_target 
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
         Height          =   310
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   29
         Tag             =   "����"
         Top             =   75
         Width           =   1200
      End
      Begin Threed.SSOption opt_top 
         Height          =   285
         Left            =   5655
         TabIndex        =   32
         Top             =   90
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   "ǰ"
         Value           =   -1
      End
      Begin Threed.SSOption opt_bottom 
         Height          =   285
         Left            =   6315
         TabIndex        =   33
         Top             =   90
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
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
         Caption         =   "��"
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ŀ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   3720
         TabIndex        =   36
         Top             =   135
         Width           =   420
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   2055
         TabIndex        =   35
         Top             =   195
         Width           =   105
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   300
         TabIndex        =   34
         Top             =   120
         Width           =   420
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   420
      Left            =   75
      TabIndex        =   37
      Top             =   435
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   741
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
      Begin Threed.SSOption opt_roll 
         Height          =   285
         Left            =   270
         TabIndex        =   38
         Top             =   90
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   "����"
         Value           =   -1
      End
      Begin Threed.SSOption opt_slab 
         Height          =   285
         Left            =   1185
         TabIndex        =   39
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
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
         Caption         =   "����"
      End
   End
   Begin Threed.SSCommand cmd_sample 
      Height          =   420
      Left            =   12420
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   0
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   12583104
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
      Caption         =   "��ȡ����Ϣ"
      BevelWidth      =   3
   End
End
Attribute VB_Name = "BEC2900C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       DAILY SCHEDULE
'-- Sub_System Name
'-- Program Name
'-- Program ID        BEC2900C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.6.26
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'   2.01  07.10.24 KIM SUNG HO  EP_SLAB_EDT --> EP_SLAB_EDT2  PROGRAM ID CHANGE
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting

Public Complete As Boolean          'Move Status Setting

Dim OldRoll As String               'Old Roll_no Setting

Dim FstRollNo As String             'Fisrt Roll No Setting
Dim SecRollNo As String             'Second Roll No Setting

Dim iEcount As Integer              'Chart End Count
Dim iCurrent As Integer             'Chart Current count
Dim lbl_index As Integer            'ss1 Redraw index

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

Dim pContro2 As New Collection      'Master Primary Key Collection
Dim nContro2 As New Collection      'Master Necessary Collection
Dim mContro2 As New Collection      'Master Maxlength check Collection
Dim iContro2 As New Collection      'Master Insert Collection
Dim rContro2 As New Collection      'Master Refer Collection
Dim cContro2 As New Collection      'Master Copy Collection
Dim aContro2 As New Collection      'Master -> Spread Collection
Dim lContro2 As New Collection      'Master Lock Collection

Dim pContro3 As New Collection      'Master Primary Key Collection
Dim nContro3 As New Collection      'Master Necessary Collection
Dim mContro3 As New Collection      'Master Maxlength check Collection
Dim iContro3 As New Collection      'Master Insert Collection
Dim rContro3 As New Collection      'Master Refer Collection
Dim cContro3 As New Collection      'Master Copy Collection
Dim aContro3 As New Collection      'Master -> Spread Collection
Dim lContro3 As New Collection      'Master Lock Collection

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
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim sLoc        As String
Dim P_PLT       As String
Dim P_UNIT      As String
Dim P_STATUS    As String
Dim P_LINE      As String
Dim P_MODE      As String
Dim P_POSITION  As String

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
             Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(txt_heat_mana_no, "p", " ", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
            Call Gp_Ms_Collection(txt_from, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
              Call Gp_Ms_Collection(txt_to, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(txt_target, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_roll_mana_no, "p", "n", "m", " ", "r", " ", "l", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
    
    'MASTER Collection
    Mc2.Add Item:=pContro2, Key:="pControl"
    Mc2.Add Item:=nContro2, Key:="nControl"
    Mc2.Add Item:=mContro2, Key:="mControl"
    Mc2.Add Item:=iContro2, Key:="iControl"
    Mc2.Add Item:=rContro2, Key:="rControl"
    Mc2.Add Item:=cContro2, Key:="cControl"
    Mc2.Add Item:=aContro2, Key:="aControl"
    Mc2.Add Item:=lContro2, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
   Call Gp_Ms_Collection(sdb_slab_edt_seq, "p", " ", " ", " ", "r", " ", "l", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
    
    'MASTER Collection
    Mc3.Add Item:=pContro3, Key:="pControl"
    Mc3.Add Item:=nContro3, Key:="nControl"
    Mc3.Add Item:=mContro3, Key:="mControl"
    Mc3.Add Item:=iContro3, Key:="iControl"
    Mc3.Add Item:=rContro3, Key:="rControl"
    Mc3.Add Item:=cContro3, Key:="cControl"
    Mc3.Add Item:=aContro3, Key:="aControl"
    Mc3.Add Item:=lContro3, Key:="lControl"
    
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
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="BEC2900C.P_REFER3", Key:="P-R"
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
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="BEC2900C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc"
    
    Call Gp_Sp_ColHidden(sc1.Item("Spread"), 13, True)  'SLAB_EDT_NO
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    lbl_thk(0).Visible = False
    lbl_wid(0).Visible = False
    lbl_len(0).Visible = False
        
End Sub

Private Sub cmd_confirm_Click()

    Complete = False

    Set Active_Spread = Me.ss1
    
    Load Roll_Confirm
       
    Roll_Confirm.P_MODE = "R"            'ROLL
    Roll_Confirm.P_PLT = txt_plt.Text    'PLT
    Roll_Confirm.P_LINE = "1"            'LINE
    
    Roll_Confirm.P_CurrentCol = 1
    
    Roll_Confirm.Show 1
    
    'If Complete Then
    '    Call Form_Ref
    'End If

End Sub

Private Sub cmd_roll_Click()

'    Complete = False
'
'    If ss1.MaxRows = 0 Then Exit Sub
'    Set Active_Spread = Me.ss1
'
'    Load Process_Change
'
'    Process_Change.P_PLT = txt_plt.Text     'SMS
'    Process_Change.P_UNIT = "R"             'Roll
'    Process_Change.P_STATUS = "D"           'Daily
'    Process_Change.P_LINE = "1"             'Line
'
'    Process_Change.P_CurrentCol = 1
'    Process_Change.P_Tcurrent = 2           'SLAB
'
'    Process_Change.Show 1
'
'    If Complete Then
'        Call Form_Ref
'    End If

End Sub

Private Sub cmd_roll1_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "SELECT COUNT(*) FROM EP_SLAB_EDT2 WHERE SLAB_EDT_FL = '1' "
    iCount = Gf_FloatFind(M_CN1, sQuery)
    
    'If iCount > 0 Then  'HCR
        'sQuery = "{call AEC2030P ('" + txt_plt.Text + "','1',?)}"
    'Else                'CCR
        sQuery = "{call AED1040P ('" + txt_plt.Text + "','1',?)}"
    'End If
                
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
        Call Gp_MsgBoxDisplay("����������λ����..!!", "I")
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub cmd_slab_Click()

'    Complete = False
'
'    If ss1.MaxRows = 0 Then Exit Sub
'
'    Set Active_Spread = Me.ss1
'
'    Load Process_Change
'
'    Process_Change.P_PLT = txt_plt.Text     'SMS
'    Process_Change.P_UNIT = "S"             'Slab
'    Process_Change.P_STATUS = "D"           'Daily
'    Process_Change.P_LINE = "1"             'Line
'
'    Process_Change.P_CurrentCol = 2         'SLAB
'
'    Process_Change.Show 1
'
'    If Complete Then
'        Call Form_Ref
'        Call lbl_thk_Click(lbl_index)
'    End If

End Sub

Private Sub cmd_sample_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call AED2000P ('" + sUserID + "',?)}"
                
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
        Call Gp_MsgBoxDisplay("��ȡ����Ϣ����..!!", "I")
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)

End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    MDIMain.MenuTool.Buttons(4).Enabled = False
    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False

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
    
    'UPDATE AUTHORITY
    If Mid(sAuthority, 3, 1) <> "1" Then
        cmd_process.Enabled = False
        cmd_roll1.Enabled = False
        cmd_confirm.Enabled = False
    End If

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    MDIMain.MenuTool.Buttons(4).Enabled = False
    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Spl_SizeGet(spl_splitter, "E-System.INI", Me.Name, "W")

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "E-System.INI", Me.Name)
    
    txt_plt.Text = "C1"
    Call txt_plt_KeyUp(0, 0)
    
    P_MODE = "M"
    P_POSITION = "T"
    P_UNIT = "R"
    sLoc = "F"
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(spl_splitter, "E-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "E-System.INI", Me.Name)
    
    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing
    
    Set pContro2 = Nothing
    Set nContro2 = Nothing
    Set iContro2 = Nothing
    Set rContro2 = Nothing
    Set cContro2 = Nothing
    Set aContro2 = Nothing
    Set lContro2 = Nothing
    Set mContro2 = Nothing
    
    Set pContro3 = Nothing
    Set nContro3 = Nothing
    Set iContro3 = Nothing
    Set rContro3 = Nothing
    Set cContro3 = Nothing
    Set aContro3 = Nothing
    Set lContro3 = Nothing
    Set mContro3 = Nothing
        
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
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    iCurrent = 0
    iEcount = 0
    lbl_index = 0
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    Dim iCount As Integer
    
    If Gf_Sp_Cls(sc2) Then
    
        If Gf_Sp_Cls(sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            MDIMain.MenuTool.Buttons(4).Enabled = False
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(9).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
                        
            If iEcount <> 0 Then
                
                For iCount = 1 To 135  'iEcount
                    Unload lbl_wid(iCount)
                    Unload lbl_thk(iCount)
                    Unload lbl_len(iCount)
                Next iCount
            
            End If
            
            iEcount = 0
            iCurrent = 0
            lbl_index = 0
            
            ss3.MaxRows = 0
            
            txt_plt.Text = "C1"
            Call txt_plt_KeyUp(0, 0)
            
            sLoc = "F"
            txt_from.BackColor = &HC0FFFF
            txt_to.BackColor = &H80000005
            txt_target.BackColor = &H80000005
            
            rContro1(1).SetFocus
        End If
        
    End If
    
End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    cmd_confirm.Visible = True
    
    sMesg = Gf_Ms_NeceCheck(nContro1)
    If sMesg = "OK" Then
    
        sMesg = Gf_Ms_NeceCheck2(mContro1)
        If sMesg = "OK" Then

            lbl_index = 0
            
            If Chart_Refer Then
                Call Gf_Sp_Cls(sc1)
                Call Gf_Sp_Cls(sc2)
                Call Gp_Ms_ControlLock(Mc1("lControl"), True)
                
                ss3.Row = 1
                ss3.Col = 4
                
                If ss3.Text = "1" Then  'HCR
                    cmd_process.Enabled = False
                    'cmd_slab.Enabled = False
                    'cmd_roll.Enabled = False
                End If
                
                sQuery = "SELECT COUNT(*) FROM EP_SLAB_EDT2 WHERE SLAB_EDT_FL = '1' "
                iCount = Gf_FloatFind(M_CN1, sQuery)
                
                If iCount > 0 Then  'HCR
                    cmd_confirm.Visible = False
                Else                'CCR
                    cmd_confirm.Visible = True
                End If
                
                ss1.SetFocus
            End If
            
        Else
            sMesg = sMesg + " Must input according to length of item"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
   
    Else
        sMesg = sMesg + " Must input necessarily"
        Call Gp_MsgBoxDisplay(sMesg)
    End If
    
End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        MDIMain.MenuTool.Buttons(4).Enabled = False
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(9).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
    End If
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
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

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub HScroll1_Change()

    Call Chart_Draw(HScroll1.VALUE)

End Sub

Private Sub opt_roll_Click(VALUE As Integer)
    If opt_roll.VALUE Then
        P_UNIT = "R"
        Call Prod_Button_Edit
    End If
End Sub

Private Sub opt_slab_Click(VALUE As Integer)
    If opt_slab.VALUE Then
        P_UNIT = "S"
        Call Prod_Button_Edit
    End If
End Sub

Private Sub Prod_Button_Edit()
    sLoc = ""
    
    opt_slab.ForeColor = &H808080
    opt_roll.ForeColor = &H808080
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    opt_move.Enabled = True
    opt_unification.Enabled = True
    opt_split.Enabled = True
    opt_delete.Enabled = True
    
    Select Case P_UNIT
        Case "R"    'Roll
            opt_roll.ForeColor = &HFF&
        Case "S"    'Slab
            opt_slab.ForeColor = &HFF&
            opt_unification.Enabled = False
            opt_split.Enabled = False
    End Select
    opt_move.VALUE = True
    
End Sub

Private Sub opt_move_Click(VALUE As Integer)

    If opt_move.VALUE = True Then
        P_MODE = "M"
        Call Process_Button_Edit
    Else
        opt_move.ForeColor = &H808080
    End If

End Sub

Private Sub opt_unification_Click(VALUE As Integer)

    If opt_unification.VALUE = True Then
        P_MODE = "U"
        Call Process_Button_Edit
    Else
        opt_unification.ForeColor = &H808080
    End If

End Sub

Private Sub opt_split_Click(VALUE As Integer)

    If opt_split.VALUE = True Then
        P_MODE = "S"
        Call Process_Button_Edit
    Else
        opt_split.ForeColor = &H808080
    End If

End Sub

Private Sub opt_delete_Click(VALUE As Integer)

    If opt_delete.VALUE = True Then
        P_MODE = "D"
        Call Process_Button_Edit
    Else
        opt_move.ForeColor = &H808080
    End If
    
End Sub

Private Sub Process_Button_Edit()
    sLoc = ""
    
    opt_move.ForeColor = &H808080
    opt_unification.ForeColor = &H808080
    opt_split.ForeColor = &H808080
    opt_delete.ForeColor = &H808080
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    txt_from.Enabled = True
    txt_to.Enabled = False
    txt_target.Enabled = False
    
    opt_bottom.Enabled = True
    opt_top.Enabled = True
    opt_top.VALUE = True
    
    Select Case P_MODE
        Case "M"    'Move
            opt_move.ForeColor = &HFF&
            If opt_slab.VALUE = True Then
                txt_to.Enabled = True
            End If
            txt_target.Enabled = True
        Case "U"    'Unification
            opt_unification.ForeColor = &HFF&
            txt_target.Enabled = True
        Case "S"    'Split
            opt_split.ForeColor = &HFF&
            txt_target.Enabled = True
        Case "D"   'Delete
            opt_delete.ForeColor = &HFF&
            If opt_slab.VALUE = True Then
                txt_to.Enabled = True
            End If
            opt_top.Enabled = False
            opt_bottom.Enabled = False
    End Select
    
    Call txt_from_Click
End Sub

Private Sub opt_top_Click(VALUE As Integer)

    If opt_top.VALUE = True Then
        opt_top.ForeColor = &HFF&
        opt_bottom.ForeColor = &H808080
        P_POSITION = "T"
    Else
        opt_top.ForeColor = &H808080
        P_POSITION = "B"
    End If

End Sub

Private Sub opt_bottom_Click(VALUE As Integer)
    If opt_bottom.VALUE = True Then
        opt_bottom.ForeColor = &HFF&
        opt_top.ForeColor = &H808080
        P_POSITION = "B"
    Else
        opt_bottom.ForeColor = &H808080
        P_POSITION = "T"
    End If
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sMatFullNo  As String
    Dim iCount      As Integer
    
    If Row < 1 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 1:   sMatFullNo = Trim(ss1.Text)
    ss1.Col = 2:   sMatFullNo = sMatFullNo & Trim(ss1.Text)
    If sLoc = "A" And opt_split.VALUE = True Then
        ss1.Col = 2
    ElseIf opt_roll.VALUE = True Then
        ss1.Col = 1
    Else
        ss1.Col = 2
    End If
    
    For iCount = 1 To 135
        If Trim(lbl_wid(iCount).Tag) = Trim(sMatFullNo) Then
            If sLoc = "F" Then
                lbl_wid(iCount).BackColor = &H80FF80
                lbl_thk(iCount).BackColor = &H80FF80
                lbl_len(iCount).BackColor = &H80FF80
            ElseIf sLoc = "T" Then
                lbl_wid(iCount).BackColor = &H80CC00
                lbl_thk(iCount).BackColor = &H80CC00
                lbl_len(iCount).BackColor = &H80CC00
            ElseIf sLoc = "A" Then
                lbl_wid(iCount).BackColor = &HFFFF&
                lbl_thk(iCount).BackColor = &HFFFF&
                lbl_len(iCount).BackColor = &HFFFF&
            End If
            iCount = 999
        End If
    Next iCount
        
    Call Location_edit(ss1.Text, sMatFullNo)
    Call Chart_Color_Edit
    
End Sub

Private Sub txt_from_Click()
    sLoc = "F"
    txt_from.BackColor = &HC0FFFF
    txt_to.BackColor = &H80000005
    txt_target.BackColor = &H80000005
End Sub

Private Sub txt_plt_DblClick()

    Call txt_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_to_Click()
    sLoc = "T"
    txt_to.BackColor = &HC0FFFF
    txt_from.BackColor = &H80000005
    txt_target.BackColor = &H80000005
End Sub

Private Sub txt_target_Click()
    sLoc = "A"
    txt_target.BackColor = &HC0FFFF
    txt_from.BackColor = &H80000005
    txt_to.BackColor = &H80000005
End Sub

Private Sub txt_from_Change()
    If txt_from.Text = "" Then
        txt_from_no.Text = ""
        txt_to.Text = ""
    End If
    
    If Trim(txt_to.Text) = "" And txt_to.Enabled = False Then
        txt_to.Text = txt_from.Text
    End If
    Call Chart_Color_Edit
End Sub

Private Sub txt_to_Change()
    If txt_to.Text = "" Then
        txt_to_no.Text = ""
        txt_target.Text = ""
    End If
    Call Chart_Color_Edit
End Sub

Private Sub txt_target_Change()
    If txt_target.Text = "" Then txt_target_no.Text = ""
    Call Chart_Color_Edit
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    'Call Gp_Sp_Sort(Sc1.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If ss1.MaxRows < 1 Or Row < 1 Then
        sdb_slab_edt_seq.VALUE = 0
        Call Gf_Sp_Cls(sc2)
        Exit Sub
    End If
    
    ss1.Row = Row
    
    ss1.Col = 15
    sdb_slab_edt_seq.VALUE = ss1.Text
    
    Call Gf_Sp_Refer(M_CN1, sc2, Mc3, Mc3("nControl"), Mc3("mControl"), False)
    'Call Gp_Sp_EvenRowBackcolor(Sc2.Item("Spread"))

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    End If
    
End Sub

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub
Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=txt_plt_name
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
        txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        txt_plt_name.Text = ""
    End If

End Sub

Public Sub Chart_Draw(DrawCnt As Integer)

    Dim iCount As Integer
    Dim SlabNo As String

    If iCurrent <> 0 Then
    
        If iEcount = 135 Then
            
            For iCount = 1 To iEcount
                lbl_wid(iCount).Visible = False
                lbl_thk(iCount).Visible = False
                lbl_len(iCount).Visible = False
            Next iCount
        
        Else
            
            For iCount = 1 To iEcount - 1
                lbl_wid(iCount).Visible = False
                lbl_thk(iCount).Visible = False
                lbl_len(iCount).Visible = False
            Next iCount
            
        End If
        
    End If
    
    Screen.MousePointer = vbHourglass
    
    With ss3
    
        For iCount = 1 To 135
    
            iEcount = iCount
            
            If iCount + DrawCnt > .MaxRows Then Exit For
            
            .Row = iCount + DrawCnt
            
            .Col = 5
            SlabNo = .Text
            
            .Col = 1
            
            If iCount = 1 And .Text <> "" Then
            
                lbl_wid(iCount).Tag = .Text + SlabNo
                lbl_thk(iCount).Tag = .Text + SlabNo
                lbl_len(iCount).Tag = .Text + SlabNo
                
                .Col = 2
                lbl_wid(iCount).Height = (1485 / 4600) * .VALUE
                lbl_wid(iCount).Top = 90 + (1485 - lbl_wid(iCount).Height)
                lbl_wid(iCount).Left = lbl_wid(0).Left + lbl_wid(0).Width
                
                .Col = 3
                lbl_thk(iCount).Top = 1590
                lbl_thk(iCount).Left = lbl_thk(0).Left + lbl_thk(0).Width
                lbl_thk(iCount).Height = (930 / 50) * .VALUE
                
                .Col = 7
                lbl_len(iCount).Height = (1155 / 18000) * .VALUE
                lbl_len(iCount).Top = 2610 + (1155 - lbl_len(iCount).Height)
                lbl_len(iCount).Left = lbl_len(0).Left + lbl_len(0).Width
                
                .Col = 4
                If .Text = "H" Then
                    lbl_wid(iCount).BackColor = &H8080FF
                    lbl_thk(iCount).BackColor = &H8080FF
                ElseIf .Text = "C" Then
                    lbl_wid(iCount).BackColor = &HFF8080
                    lbl_thk(iCount).BackColor = &HFF8080
                Else
                    lbl_wid(iCount).BackColor = &H80C0FF
                    lbl_thk(iCount).BackColor = &H80C0FF
                End If
                
                .Col = 6
                If .Text = "HC" Then
                    lbl_len(iCount).BackColor = &HFFC0FF
                Else
                    lbl_len(iCount).BackColor = &HFFFFC0
                End If
                
'                lbl_wid(iCount).BorderStyle = 1
'                lbl_thk(iCount).BorderStyle = 1
'                lbl_len(iCount).BorderStyle = 1
                    
                lbl_wid(iCount).Visible = True
                lbl_thk(iCount).Visible = True
                lbl_len(iCount).Visible = True
        
            Else
            
                If .Text <> "" Then
                
                    lbl_wid(iCount).Tag = .Text + SlabNo
                    lbl_thk(iCount).Tag = .Text + SlabNo
                    lbl_len(iCount).Tag = .Text + SlabNo
                
                    .Col = 2
                    lbl_wid(iCount).Left = lbl_wid(iCount - 1).Left + lbl_wid(iCount - 1).Width
                    lbl_wid(iCount).Height = (1485 / 4600) * .VALUE
                    lbl_wid(iCount).Top = 90 + (1485 - lbl_wid(iCount).Height)
                        
                    .Col = 3
                    lbl_thk(iCount).Top = 1590
                    lbl_thk(iCount).Left = lbl_thk(iCount - 1).Left + lbl_thk(iCount - 1).Width
                    lbl_thk(iCount).Height = (930 / 50) * .VALUE
                    
                    .Col = 7
                    lbl_len(iCount).Left = lbl_len(iCount - 1).Left + lbl_len(iCount - 1).Width
                    lbl_len(iCount).Height = (1155 / 18000) * .VALUE
                    lbl_len(iCount).Top = 2610 + (1155 - lbl_len(iCount).Height)
                    
                    .Col = 4
                    If .Text = "H" Then
                        lbl_wid(iCount).BackColor = &H8080FF
                        lbl_thk(iCount).BackColor = &H8080FF
                    ElseIf .Text = "C" Then
                        lbl_wid(iCount).BackColor = &HFF8080
                        lbl_thk(iCount).BackColor = &HFF8080
                    Else
                        lbl_wid(iCount).BackColor = &H80C0FF
                        lbl_thk(iCount).BackColor = &H80C0FF
                    End If
                    
                    .Col = 6
                    If .Text = "HC" Then
                        lbl_len(iCount).BackColor = &HFFC0FF
                    Else
                        lbl_len(iCount).BackColor = &HFFFFC0
                    End If
                
'                    lbl_wid(iCount).BorderStyle = 1
'                    lbl_thk(iCount).BorderStyle = 1
'                    lbl_len(iCount).BorderStyle = 1
                    
                    lbl_wid(iCount).Visible = True
                    lbl_thk(iCount).Visible = True
                    lbl_len(iCount).Visible = True
                    
                ElseIf .Text = "" And iCount = 1 Then
                
                    lbl_wid(iCount).Tag = .Text + SlabNo
                    lbl_thk(iCount).Tag = .Text + SlabNo
                    lbl_len(iCount).Tag = .Text + SlabNo
                
                    lbl_wid(iCount).Top = 90
                    lbl_wid(iCount).Left = lbl_wid(iCount - 1).Left + lbl_wid(iCount - 1).Width
                    lbl_wid(iCount).Height = 1485
                    
                    lbl_thk(iCount).Top = 1590
                    lbl_thk(iCount).Left = lbl_thk(iCount - 1).Left + lbl_thk(iCount - 1).Width
                    lbl_thk(iCount).Height = 930
                    
                    lbl_len(iCount).Top = 2610
                    lbl_len(iCount).Left = lbl_len(iCount - 1).Left + lbl_len(iCount - 1).Width
                    lbl_len(iCount).Height = 1155
                    
                    lbl_wid(iCount).BackColor = &HE0E0E0
                    lbl_thk(iCount).BackColor = &HE0E0E0
                    lbl_len(iCount).BackColor = &HE0E0E0
                    
'                    lbl_wid(iCount).BorderStyle = 0
'                    lbl_thk(iCount).BorderStyle = 0
'                    lbl_len(iCount).BorderStyle = 0
'
'                    lbl_wid(iCount).Visible = True
'                    lbl_thk(iCount).Visible = True
'                    lbl_len(iCount).Visible = True
                    
                    lbl_wid(iCount).Visible = False
                    lbl_thk(iCount).Visible = False
                    lbl_len(iCount).Visible = False
                Else
                
                    lbl_wid(iCount).Tag = .Text + SlabNo
                    lbl_thk(iCount).Tag = .Text + SlabNo
                    lbl_len(iCount).Tag = .Text + SlabNo
                
                    lbl_wid(iCount).Top = 90
                    lbl_wid(iCount).Left = lbl_wid(iCount - 1).Left + lbl_wid(iCount - 1).Width
                    lbl_wid(iCount).Height = 1485
                    
                    lbl_thk(iCount).Top = 1590
                    lbl_thk(iCount).Left = lbl_thk(iCount - 1).Left + lbl_thk(iCount - 1).Width
                    lbl_thk(iCount).Height = 930
                    
                    lbl_len(iCount).Top = 2610
                    lbl_len(iCount).Left = lbl_len(iCount - 1).Left + lbl_len(iCount - 1).Width
                    lbl_len(iCount).Height = 1155
                    
                    lbl_wid(iCount).BackColor = &HE0E0E0
                    lbl_thk(iCount).BackColor = &HE0E0E0
                    lbl_len(iCount).BackColor = &HE0E0E0
                    
'                    lbl_wid(iCount).BorderStyle = 0
'                    lbl_thk(iCount).BorderStyle = 0
'                    lbl_len(iCount).BorderStyle = 0
'
'                    lbl_wid(iCount).Visible = True
'                    lbl_thk(iCount).Visible = True
'                    lbl_len(iCount).Visible = True
                    
                    lbl_wid(iCount).Visible = False
                    lbl_thk(iCount).Visible = False
                    lbl_len(iCount).Visible = False
                    
                End If
            
            End If
            
            txt_color.BackColor = lbl_wid(iCount).BackColor
            txt_color_len.BackColor = lbl_len(iCount).BackColor
            
            If txt_from.Text = "" And txt_to.Text = "" And txt_target.Text = "" Then
            
            Else
                If txt_from.Text <> "" Then
                    If lbl_wid(iCount).Tag = txt_from_no Then
                        lbl_wid(iCount).BackColor = &H80FF80
                        lbl_thk(iCount).BackColor = &H80FF80
                    End If
                End If
                If txt_to.Text <> "" Then
                    If lbl_wid(iCount).Tag = txt_to_no Then
                        lbl_wid(iCount).BackColor = &H80CC00
                        lbl_thk(iCount).BackColor = &H80CC00
                    End If
                End If
                If txt_target.Text <> "" Then
                    If lbl_wid(iCount).Tag = txt_target_no Then
                        lbl_wid(iCount).BackColor = &HFFFF&
                        lbl_thk(iCount).BackColor = &HFFFF&
                    End If
                End If
            End If
            
        Next iCount
        
        iCurrent = .Row
        
    End With
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Function Chart_Refer() As Boolean

On Error GoTo Chart_Refer_Error

    Dim sHcr As String
    Dim sTemp As String
    Dim sQuery As String
    
    Dim iRcnt As Integer
    Dim iCount As Integer
    Dim lRowCount As Integer
    Dim iSeparator As Integer
    
    Dim dYvalueMax As Double
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
    Set AdoRs = New ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    
    sQuery = "SELECT COUNT(*) FROM (SELECT COUNT(*) FROM EP_SLAB_EDT2 WHERE MILL_PLT = '" & txt_plt.Text & "' "
    sQuery = sQuery + " AND ROLL_MANA_NO IS NOT NULL GROUP BY ROLL_MANA_NO) "
    
    iSeparator = Gf_FloatFind(M_CN1, sQuery)
    
    If iSeparator = 0 Then
        Call Gp_MsgBoxDisplay("There is No Relevant Data", "I")
        Chart_Refer = False
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    iSeparator = iSeparator - 1

    sQuery = "SELECT ROLL_MANA_NO, ASROLL_WID, ASROLL_THK, HCR_FL, SLAB_NO, PROD_CD, SLAB_LEN  "
    sQuery = sQuery + " FROM EP_SLAB_EDT2 WHERE MILL_PLT = '" & txt_plt.Text & "' "
    sQuery = sQuery + " AND ROLL_MANA_NO IS NOT NULL ORDER BY ROLL_MANA_NO, ROLL_SLAB_SEQ  "
        
    With ss3

        Chart_Refer = True
        .ReDraw = False
        .MaxRows = 0
        
        If iCurrent <> 0 Then
    
            If lbl_wid(1).Visible Then
            
                For iCount = 1 To iEcount - 1
                    lbl_wid(iCount).Visible = False
                    lbl_thk(iCount).Visible = False
                    lbl_len(iCount).Visible = False
                Next iCount
                
            Else
            
                For iCount = 1 To 135
                    Load lbl_wid(iCount)
                    Load lbl_thk(iCount)
                    Load lbl_len(iCount)
                Next iCount
                
            End If
            
            iCurrent = 0
            
        Else
        
            For iCount = 1 To 135
                Load lbl_wid(iCount)
                Load lbl_thk(iCount)
                Load lbl_len(iCount)
            Next iCount
            
        End If
    
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Call Gp_MsgBoxDisplay("There is No Relevant Data", "I")
                
            Chart_Refer = False
            .ReDraw = True
            
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing
        
        If UBound(ArrayRecords, 1) <> 0 Then
    
            .MaxRows = UBound(ArrayRecords, 2) + iSeparator + 1
            
            HScroll1.max = .MaxRows
            HScroll1.VALUE = 0
            
            iRcnt = 1
            
            For lRowCount = 0 To UBound(ArrayRecords, 2)
            
                .Row = iRcnt
                
                If .Row = 1 Then
                    sTemp = Trim(ArrayRecords(0, lRowCount))
                    
                    For iCount = 0 To 6
                    
                        .Col = iCount + 1
                        
                        If VarType(ArrayRecords(iCount, lRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(iCount, lRowCount))
                            
                        End If
                    Next iCount
                    
                Else
                
                    If sTemp <> Trim(ArrayRecords(0, lRowCount)) Then
                        sTemp = Trim(ArrayRecords(0, lRowCount))
                        
                        iRcnt = iRcnt + 1
                        .Row = iRcnt
                    
                        For iCount = 0 To 6
                            
                            .Col = iCount + 1
                            
                            If VarType(ArrayRecords(iCount, lRowCount)) = vbNull Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(iCount, lRowCount))
                            End If
                            
                        Next iCount
                        
                    Else
                    
                        For iCount = 0 To 6
                            
                            .Col = iCount + 1
                            
                            If VarType(ArrayRecords(iCount, lRowCount)) = vbNull Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(iCount, lRowCount))
                            End If
                            
                        Next iCount
                    
                    End If
                        
                End If
                
                iRcnt = iRcnt + 1
                
            Next lRowCount
            
        End If
                                            
        Chart_Refer = True
        .ReDraw = True
        
    End With
    
    Call Chart_Draw(0)
    
    Screen.MousePointer = vbDefault
    Exit Function
    
Chart_Refer_Error:

    Chart_Refer = False
    Screen.MousePointer = vbDefault

End Function

Private Sub lbl_thk_Click(Index As Integer)
    If lbl_thk(Index).Tag = "" Then Exit Sub
        
    Call Chart_Click_Process(Index)
End Sub

Private Sub lbl_thk_DblClick(Index As Integer)
    If lbl_thk(Index).Tag = "" Then Exit Sub
    
    Call Chart_DblClick_Process(Index)
End Sub

Private Sub lbl_wid_click(Index As Integer)
    If lbl_wid(Index).Tag = "" Then Exit Sub
        
    Call Chart_Click_Process(Index)
End Sub

Private Sub lbl_wid_DblClick(Index As Integer)
    If lbl_wid(Index).Tag = "" Then Exit Sub
    
    Call Chart_DblClick_Process(Index)
    
End Sub

Private Sub lbl_len_click(Index As Integer)
    If lbl_len(Index).Tag = "" Then Exit Sub
    
    Call Chart_Click_Process(Index)
End Sub

Private Sub lbl_len_DblClick(Index As Integer)
    If lbl_len(Index).Tag = "" Then Exit Sub
    
    Call Chart_DblClick_Process(Index)
End Sub

Private Sub Chart_Click_Process(Index As Integer)

    Dim iCount As Integer
    Dim SlabNo As String
    Dim RollNo As String
    Dim sTemp As String
    
    txt_roll_mana_no.Text = Mid(lbl_wid(Index).Tag, 1, 5)

    If ss1.Row = 0 Then
    Else
        ss1.Row = ss1.ActiveRow
        ss1.Col = 1
        RollNo = ss1.Text
        ss1.Col = 2
        SlabNo = ss1.Text
    End If
    
    If lbl_wid(Index).Tag = RollNo + SlabNo Then Exit Sub
    If lbl_wid(Index).Tag = "" Then Exit Sub
    
    If txt_roll_mana_no.Text <> RollNo Then
        Call Gf_Sp_Refer(M_CN1, sc1, Mc2, Mc2("nControl"), Mc2("mControl"), False)
        'Call Gp_Sp_EvenRowBackcolor(Sc1.Item("Spread"))
        Call Gf_Sp_Cls(sc2)
        lbl_index = Index
    Else
        'Call Gp_Sp_EvenRowBackcolor(Sc1.Item("Spread"))
        Call Gf_Sp_Cls(sc2)
    End If
    
    With ss1
    
        For iCount = 1 To .MaxRows
        
            .Row = iCount
            .Col = 1
            RollNo = .Text
            .Col = 2
            SlabNo = .Text
            
            If RollNo = Mid(lbl_wid(Index).Tag, 1, 5) And SlabNo = Mid(lbl_wid(Index).Tag, 6, 10) Then
                .TabStop = True
                .Row = iCount + 13
                .Col = 1
                .Action = ActionActiveCell
                
                .Row = iCount
                .Col = 1
                .Action = ActionActiveCell
                
                .BlockMode = True
                .Row = iCount: .Row2 = iCount
                .Col = 1: .Col2 = -1
                '.BackColor = &HFFFF00      '  &HF2F2F2   'RGB(241, 236, 255)   '&HFFC0FF
                .BlockMode = False
                
                .TabStop = False
                
                ss1.Col = 15
                sdb_slab_edt_seq.VALUE = ss1.Text
                
                Call Gf_Sp_Refer(M_CN1, sc2, Mc3, Mc3("nControl"), Mc3("mControl"), False)
                'Call Gp_Sp_EvenRowBackcolor(Sc2.Item("Spread"))

            End If
            
        Next iCount
        
        For iCount = 1 To .MaxRows
                
            .Row = iCount
            .Col = 1
            
            If iCount = 1 Then sTemp = .Text
            
            If sTemp <> .Text Then
                sTemp = .Text
                Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iCount, iCount, , &HFFC0FF)
            End If
        
        Next iCount

    End With
    
End Sub

Private Sub Chart_DblClick_Process(Index As Integer)

    Dim iCount  As Integer
    Dim sRollNo As String
    Dim sSlabNo As String
    Dim sMatNo  As String
    
    If lbl_wid(Index).BackColor = &H80FF80 Or lbl_wid(Index).BackColor = &H80CC00 Or lbl_wid(Index).BackColor = &HFFFF& Then    'green, yellow
            
        If lbl_wid(Index).BackColor = &H80FF80 Then
            sLoc = "F"
        ElseIf lbl_wid(Index).BackColor = &H80CC00 Then
            sLoc = "T"
        ElseIf lbl_wid(Index).BackColor = &HFFFF& Then
            sLoc = "A"
        End If
        
        For iCount = 1 To ss3.MaxRows
        
            ss3.Row = iCount
            
            ss3.Col = 1
            sRollNo = ss3.Text
            
            ss3.Col = 5
            sSlabNo = ss3.Text
        
            If lbl_wid(Index).Tag = sRollNo + sSlabNo Then
            
                ss3.Col = 4
                
                If ss3.Text = "H" Then
                    lbl_wid(Index).BackColor = &H8080FF
                    lbl_thk(Index).BackColor = &H8080FF
                ElseIf ss3.Text = "C" Then
                    lbl_wid(Index).BackColor = &HFF8080
                    lbl_thk(Index).BackColor = &HFF8080
                Else
                    lbl_wid(Index).BackColor = &H80C0FF
                    lbl_thk(Index).BackColor = &H80C0FF
                End If
                
                ss3.Col = 6
                
                If ss3.Text = "HC" Then
                    lbl_len(iCount).BackColor = &HFFC0FF
                Else
                    lbl_len(iCount).BackColor = &HFFFFC0
                End If
                
            End If
            
        Next iCount

        Call Location_edit("", "")
    Else
    
        If sLoc = "F" Then
            lbl_wid(Index).BackColor = &H80FF80
            lbl_thk(Index).BackColor = &H80FF80
            lbl_len(Index).BackColor = &H80FF80
        ElseIf sLoc = "T" Then
            lbl_wid(Index).BackColor = &H80CC00
            lbl_thk(Index).BackColor = &H80CC00
            lbl_len(Index).BackColor = &H80CC00
        ElseIf sLoc = "A" Then
            lbl_wid(Index).BackColor = &HFFFF&
            lbl_thk(Index).BackColor = &HFFFF&
            lbl_len(Index).BackColor = &HFFFF&
        End If
        
        If sLoc = "A" And opt_split.VALUE = True Then
            sMatNo = Mid(lbl_wid(Index).Tag, 6, 10)
        ElseIf opt_roll.VALUE = True Then
            sMatNo = Mid(lbl_wid(Index).Tag, 1, 5)
        Else
            sMatNo = Mid(lbl_wid(Index).Tag, 6, 10)
        End If
        Call Location_edit(sMatNo, Left(lbl_wid(Index).Tag, 15))
    
    End If
    
End Sub

Private Sub Location_edit(ByVal sMatNo As String, ByVal sMatFullNo As String)

    Select Case sLoc
        Case "F"
            txt_from_no.Text = sMatFullNo
            txt_from.Text = sMatNo
            If txt_from.Text <> "" Then
                If txt_to.Enabled = True Then
                    Call txt_to_Click
                ElseIf txt_target.Enabled = True Then
                    Call txt_target_Click
                End If
            Else
                Call txt_from_Click
            End If
        Case "T"
            txt_to_no.Text = sMatFullNo
            txt_to.Text = sMatNo
            If txt_to.Text <> "" And txt_target.Enabled = True Then
                Call txt_target_Click
            Else
                Call txt_to_Click
            End If
        Case "A"
            txt_target_no.Text = sMatFullNo
            txt_target.Text = sMatNo
    End Select
    Call Chart_Color_Edit
End Sub
        
Private Sub Chart_Color_Edit()
    Dim iCount      As Integer
    Dim sMatNoFrom  As String
    Dim sMatNoTo    As String
    Dim sMatNoAim   As String
        
    sMatNoFrom = Trim(txt_from_no.Text)
    sMatNoTo = Trim(txt_to_no.Text)
    sMatNoAim = Trim(txt_target_no.Text)

    For iCount = 1 To 135
        If Trim(lbl_len(iCount).Tag) <> sMatNoFrom And _
           Trim(lbl_len(iCount).Tag) <> sMatNoTo And _
           Trim(lbl_len(iCount).Tag) <> sMatNoAim Then
            If lbl_wid(iCount).BackColor = &H80FF80 Or _
               lbl_wid(iCount).BackColor = &H80CC00 Or _
               lbl_wid(iCount).BackColor = &HFFFF& Or _
               lbl_thk(iCount).BackColor = &H80FF80 Or _
               lbl_thk(iCount).BackColor = &H80CC00 Or _
               lbl_thk(iCount).BackColor = &HFFFF& Or _
               lbl_len(iCount).BackColor = &H80FF80 Or _
               lbl_len(iCount).BackColor = &H80CC00 Or _
               lbl_len(iCount).BackColor = &HFFFF& Then
                    
                lbl_wid(iCount).BackColor = txt_color.BackColor
                lbl_thk(iCount).BackColor = txt_color.BackColor
                lbl_len(iCount).BackColor = txt_color_len.BackColor
            End If
        End If
        If Trim(lbl_len(iCount).Tag) = "" Then
            lbl_wid(iCount).BackColor = &HE0E0E0
            lbl_thk(iCount).BackColor = &HE0E0E0
            lbl_len(iCount).BackColor = &HE0E0E0
        End If
    Next iCount
    
End Sub

Private Sub cmd_process_Click()
    If txt_from.Text = "" Or txt_to.Text = "" Or txt_target.Text = "" Then
        If P_MODE = "D" Then
            If txt_from.Text = "" Or txt_to.Text = "" Then
                Call Gp_MsgBoxDisplay("Must input Value of From, To item")
            End If
        Else
            Call Gp_MsgBoxDisplay("Must input From, To, Value of Target item")
            Exit Sub
        End If
    End If
    
    If Trim(txt_from.Text) <= Trim(txt_target.Text) And _
       Trim(txt_target.Text) <= Trim(txt_to.Text) And _
       opt_split.VALUE = False And opt_unification.VALUE = False Then
        Call Gp_MsgBoxDisplay("Value of Target item is between from and to..")
        Exit Sub
    End If
    
    Call Lf_Process_Exec
End Sub

Public Sub Lf_Process_Exec()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    P_PLT = txt_plt.Text     'SMS
    P_STATUS = "D"           'Daily
    P_LINE = "1"             'Line
        
    sQuery = "{call BEZ5000P ('" + P_PLT + "','" + Trim(str(P_LINE)) + "','" + P_STATUS + "','" + P_MODE + "','" + P_UNIT + "','"
    sQuery = sQuery + Trim(txt_from.Text) + "','" + Trim(txt_to.Text) + "','" + Trim(txt_target.Text) + "','"
    sQuery = sQuery + P_POSITION + "','" + sUserID + "',?)}"
    
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
        Call Form_Ref
        
        If P_UNIT = "S" Then
            Call lbl_thk_Click(lbl_index)
        End If
        
        txt_from.BackColor = &HC0FFFF
        txt_to.BackColor = &H80000005
        txt_target.BackColor = &H80000005
        txt_from.Text = ""
        txt_to.Text = ""
        txt_target.Text = ""
        sLoc = "F"
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub
