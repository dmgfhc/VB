VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB4100C 
   Caption         =   "���϶�����Ϣ�������_ACB4100C"
   ClientHeight    =   9225
   ClientLeft      =   345
   ClientTop       =   2400
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15285
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7410
      Left            =   30
      TabIndex        =   0
      Top             =   1770
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   13070
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACB4100C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   3795
         Left            =   0
         TabIndex        =   38
         Top             =   3615
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   6694
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   14737632
         PaneTree        =   "ACB4100C.frx":0052
         Begin Threed.SSPanel SSPanel1 
            Height          =   570
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   1005
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_thk_cd 
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
               Left            =   1320
               MaxLength       =   1
               TabIndex        =   42
               Tag             =   "CD_MANA_NO"
               Text            =   "N"
               Top             =   120
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox txt_spec 
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
               Left            =   6330
               MaxLength       =   20
               TabIndex        =   41
               Tag             =   "����(��׼��)"
               Top             =   120
               Width           =   2490
            End
            Begin VB.TextBox prod_txt_prod_cd 
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
               Height          =   315
               Left            =   4015
               MaxLength       =   2
               TabIndex        =   40
               Tag             =   "�����Ʒ����"
               Top             =   120
               Width           =   435
            End
            Begin InDate.ULabel ULabel2 
               Height          =   315
               Index           =   1
               Left            =   2895
               Top             =   120
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               Caption         =   "��Ʒ"
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
            Begin CSTextLibCtl.sidbEdit sdb_thk 
               Height          =   315
               Left            =   11190
               TabIndex        =   43
               Top             =   120
               Width           =   810
               _Version        =   262145
               _ExtentX        =   1429
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
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
               BorderEffect    =   2
               DataProperty    =   2
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
               FmtControl      =   1
               NumDecDigits    =   2
               NumIntDigits    =   4
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel7 
               Height          =   315
               Index           =   1
               Left            =   9510
               Top             =   120
               Width           =   1650
               _ExtentX        =   2910
               _ExtentY        =   556
               Caption         =   "�� x �� x ��"
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
            End
            Begin CSTextLibCtl.sidbEdit sdb_wid 
               Height          =   315
               Left            =   12270
               TabIndex        =   44
               Top             =   120
               Width           =   960
               _Version        =   262145
               _ExtentX        =   1693
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
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
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit sdb_len 
               Height          =   315
               Left            =   13500
               TabIndex        =   45
               Top             =   120
               Width           =   1020
               _Version        =   262145
               _ExtentX        =   1799
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
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
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel11 
               Height          =   315
               Left            =   5100
               Top             =   120
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   556
               Caption         =   "��׼��"
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
            Begin Threed.SSCheck ssc_thk_cd 
               Height          =   285
               Left            =   270
               TabIndex        =   46
               Top             =   150
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   2
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "ͬ���"
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "x"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   13290
               TabIndex        =   48
               Top             =   90
               Width           =   150
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "x"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   12060
               TabIndex        =   47
               Top             =   90
               Width           =   150
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   3195
            Left            =   0
            TabIndex        =   49
            Top             =   600
            Width           =   15195
            _Version        =   393216
            _ExtentX        =   26802
            _ExtentY        =   5636
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColsFrozen      =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   13
            MaxRows         =   1
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "ACB4100C.frx":00A4
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   3555
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15195
         _Version        =   393216
         _ExtentX        =   26802
         _ExtentY        =   6271
         _StockProps     =   64
         ColsFrozen      =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   40
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACB4100C.frx":0856
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1740
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   3069
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.TextBox TXT_ORD_NO 
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
         Left            =   1230
         MaxLength       =   11
         TabIndex        =   20
         Tag             =   "CD_MANA_NO"
         Top             =   531
         Width           =   1380
      End
      Begin VB.TextBox Text_STLGRD 
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
         Left            =   5475
         MaxLength       =   20
         TabIndex        =   19
         Tag             =   "����(��׼��)"
         Top             =   531
         Width           =   3060
      End
      Begin VB.ComboBox CBO_ORD_ITEM 
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
         Left            =   2640
         TabIndex        =   18
         Top             =   531
         Width           =   750
      End
      Begin VB.TextBox text_cur_inv_code 
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
         Left            =   5475
         MaxLength       =   2
         TabIndex        =   17
         Top             =   1305
         Width           =   450
      End
      Begin VB.TextBox text_cur_inv 
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
         Left            =   5940
         TabIndex        =   16
         Top             =   1305
         Width           =   1080
      End
      Begin VB.TextBox Text_size_knd_name 
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
         Height          =   310
         Left            =   1785
         TabIndex        =   15
         Tag             =   "����"
         Top             =   1305
         Width           =   1620
      End
      Begin VB.TextBox Text_size_knd 
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
         Left            =   1230
         MaxLength       =   2
         TabIndex        =   14
         Tag             =   "����"
         Top             =   1305
         Width           =   525
      End
      Begin VB.TextBox txt_prod_grd 
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
         Left            =   10080
         MaxLength       =   1
         TabIndex        =   13
         Top             =   540
         Width           =   915
      End
      Begin VB.TextBox txt_TRIM_NAME 
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
         Height          =   310
         Left            =   1785
         TabIndex        =   12
         Tag             =   "����"
         Top             =   927
         Width           =   1620
      End
      Begin VB.TextBox txt_TRIM_FL 
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
         Left            =   1230
         MaxLength       =   1
         TabIndex        =   11
         Tag             =   "����"
         Top             =   927
         Width           =   525
      End
      Begin VB.TextBox TXT_HTM 
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
         Left            =   10080
         MaxLength       =   1
         TabIndex        =   10
         Top             =   1320
         Width           =   915
      End
      Begin VB.TextBox TXT_MAT_NO 
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
         Left            =   5475
         MaxLength       =   14
         TabIndex        =   9
         Tag             =   "���Ϻ�"
         Top             =   927
         Width           =   1965
      End
      Begin VB.ComboBox CBO_PROD_CD 
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
         Height          =   315
         ItemData        =   "ACB4100C.frx":1ACA
         Left            =   1230
         List            =   "ACB4100C.frx":1AD4
         TabIndex        =   8
         Text            =   "PP"
         Top             =   135
         Width           =   750
      End
      Begin VB.ComboBox CBO_PLT 
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "ACB4100C.frx":1AE0
         Left            =   3240
         List            =   "ACB4100C.frx":1AF0
         TabIndex        =   7
         Text            =   "C1"
         Top             =   135
         Width           =   750
      End
      Begin VB.TextBox TXT_ENDUSE_CD 
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
         Left            =   10080
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "CD_MANA_NO"
         Top             =   930
         Width           =   915
      End
      Begin VB.TextBox TXT_BED_PILE_DATE 
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
         Left            =   11040
         MaxLength       =   1
         TabIndex        =   5
         Tag             =   "CD_MANA_NO"
         Text            =   "N"
         Top             =   150
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TXT_UST_FL 
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
         Left            =   7080
         MaxLength       =   1
         TabIndex        =   4
         Tag             =   "CD_MANA_NO"
         Text            =   "N"
         Top             =   1305
         Visible         =   0   'False
         Width           =   375
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
         ItemData        =   "ACB4100C.frx":1B04
         Left            =   10080
         List            =   "ACB4100C.frx":1B11
         TabIndex        =   3
         Top             =   135
         Width           =   915
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   315
         Left            =   11250
         TabIndex        =   21
         Top             =   135
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.OptionButton Opt_rk_y 
            BackColor       =   &H00E0E0E0&
            Caption         =   "���"
            Height          =   195
            Left            =   1485
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   60
            Width           =   690
         End
         Begin VB.OptionButton Opt_all 
            BackColor       =   &H00E0E0E0&
            Caption         =   "ȫ��"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   390
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   60
            Value           =   -1  'True
            Width           =   750
         End
         Begin VB.OptionButton Opt_rk_n 
            BackColor       =   &H00E0E0E0&
            Caption         =   "δ���"
            Height          =   195
            Left            =   2550
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   60
            Width           =   915
         End
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   130
         Top             =   531
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "������"
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
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   0
         Left            =   4240
         Top             =   135
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "��������"
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
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   0
         Left            =   130
         Top             =   135
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "��Ʒ"
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   4240
         Top             =   531
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "��׼��"
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
      Begin CSTextLibCtl.sidbEdit sdb_thk_fr 
         Height          =   315
         Left            =   12510
         TabIndex        =   25
         Top             =   525
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1799
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
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
         BorderEffect    =   2
         DataProperty    =   2
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
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   4
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.UDate DTP_PROD_FR 
         Height          =   315
         Left            =   5475
         TabIndex        =   26
         Tag             =   "INS_DATE"
         Top             =   135
         Width           =   1410
         _ExtentX        =   2487
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
      Begin InDate.UDate DTP_PROD_TO 
         Height          =   315
         Left            =   7130
         TabIndex        =   27
         Tag             =   "INS_DATE"
         Top             =   135
         Width           =   1410
         _ExtentX        =   2487
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Index           =   0
         Left            =   11250
         Top             =   525
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "���"
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
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   11250
         Top             =   930
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "����"
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
      End
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   11250
         Top             =   1320
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "����"
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
      End
      Begin CSTextLibCtl.sidbEdit sdb_wid_fr 
         Height          =   315
         Left            =   12510
         TabIndex        =   28
         Top             =   930
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1799
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
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
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_thk_to 
         Height          =   315
         Left            =   13845
         TabIndex        =   29
         Top             =   525
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1799
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
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
         BorderEffect    =   2
         DataProperty    =   2
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
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   4
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_wid_to 
         Height          =   315
         Left            =   13845
         TabIndex        =   30
         Top             =   930
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1799
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
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
      Begin CSTextLibCtl.sidbEdit sdb_len_to 
         Height          =   315
         Left            =   13845
         TabIndex        =   31
         Top             =   1320
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1799
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Left            =   4240
         Top             =   1305
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "�ѷŲֿ�"
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
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   135
         Top             =   1305
         Width           =   1065
         _ExtentX        =   1879
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
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   8820
         Top             =   540
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "����ȼ�"
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
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   135
         Top             =   927
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "�б�"
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
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   8820
         Top             =   1320
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "�ȴ�������"
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   4240
         Top             =   927
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "���Ϻ�"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   2220
         Top             =   135
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
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin Threed.SSCheck SSC_UST_FL 
         Height          =   285
         Left            =   7530
         TabIndex        =   32
         Top             =   1320
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   2
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
         Caption         =   "�Ƿ�̽��"
      End
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   8820
         Top             =   930
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "������;"
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
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   8820
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
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
      Begin CSTextLibCtl.sidbEdit sdb_len_fr 
         Height          =   315
         Left            =   12510
         TabIndex        =   33
         Top             =   1320
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1799
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0.00"
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
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         Height          =   120
         Left            =   6960
         TabIndex        =   37
         Top             =   195
         Width           =   90
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         Height          =   120
         Left            =   13650
         TabIndex        =   36
         Top             =   630
         Width           =   90
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         Height          =   120
         Left            =   13650
         TabIndex        =   35
         Top             =   1050
         Width           =   90
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         Height          =   120
         Left            =   13650
         TabIndex        =   34
         Top             =   1410
         Width           =   90
      End
   End
End
Attribute VB_Name = "ACB4100C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name
'-- Program Name
'-- Program ID        ACB1020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Yang Zhibin
'-- Date              2003.9.8
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
Public STR1 As String
Public BASE As String
Public AIMNO As String
Dim sQuery As String

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

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column1 Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column1 Collection
Dim iColumn1 As New Collection      'Spread Insert Column1 Collection
Dim aColumn1 As New Collection      'Master -> Spread Column1 Collection
Dim lColumn1 As New Collection      'Spread Lock Column1 Collection

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column1 Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column1 Collection
Dim iColumn2 As New Collection      'Spread Insert Column1 Collection
Dim aColumn2 As New Collection      'Master -> Spread Column1 Collection
Dim lColumn2 As New Collection      'Spread Lock Column1 Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iCount As Integer

Dim iOrd_row As Integer
Const SPD_MAT_NO = 1
Const SPD_THK = 9
Const SPD_WID = 10
Const SPD_LEN = 11
Const SPD_APLY_STDSPEC = 8
Const SPD_ORD_NO = 6
Const SPD_ORD_ITEM = 7

Private Sub Form_Define()

    Dim iRow As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(CBO_PROD_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(CBO_PLT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(DTP_PROD_FR, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(DTP_PROD_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(cbo_ord_item, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_prod_grd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_enduse_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(text_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_len_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_len_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_mat_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(TXT_HTM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_BED_PILE_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_UST_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_TRIM_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(Text_size_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

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
    
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    For iRow = 2 To 5
        Call Gp_Sp_Collection(ss1, iRow, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Next iRow
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    For iRow = 8 To 39
        Call Gp_Sp_Collection(ss1, iRow, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Next iRow
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB4100C.P_SREFER1", Key:="P-R"
    sc1.Add Item:="ACB4100C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="ACB4100C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
 ' Call Gp_Ms_Collection(prod_txt_prod_cd, "p", "n ", " ", " ", "r", " ", "", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(txt_thk_cd, "p", "n", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   Call Gp_Ms_Collection(prod_txt_prod_cd, "p", "n", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(txt_spec, "p", "n", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(sdb_thk, "p", "n", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(sdb_wid, "p", "n", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(sdb_len, "p", "n", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)

       'MASTER Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
   
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)

    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ACB4100C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
'    Call Gp_Sp_ColHidden(ss1, ss1.MaxCols - 1, True)
'    Call Gp_Sp_ColHidden(ss1, 2, True)

End Sub



Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    
    Call Gp_Sp_Setting(ss1)
    Call Gp_Sp_Setting(ss2)
'    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "C-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(ss1, "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss2, "C-System.INI", Me.Name)
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    If App.Title = "CE" Then
        text_cur_inv_code.Text = "ZB"
        CBO_PROD_CD.Text = "SL"
        CBO_PLT.Text = "C3"
    ElseIf App.Title = "DE" Then
        text_cur_inv_code.Text = "00"
        CBO_PROD_CD.Text = "PP"
        CBO_PLT.Text = "C1"
    ElseIf App.Title = "CG" Then
        text_cur_inv_code.Text = "ZB"
        CBO_PROD_CD.Text = "PP"
        CBO_PLT.Text = "C3"
    ElseIf App.Title = "BG" Then
        text_cur_inv_code.Text = "00"
        CBO_PROD_CD.Text = "PP"
        CBO_PLT.Text = "C1"
    Else
        text_cur_inv_code.Text = "00"
        CBO_PROD_CD.Text = "PP"
        CBO_PLT.Text = "C1"
    End If
    
    Call text_cur_inv_code_KeyUp(0, 0)
    iOrd_row = 0
    ssc_thk_cd.Value = ssCBChecked
    
    Call Gp_Sp_ColHidden(ss1, 4, True)

    DTP_PROD_FR.Text = Format(DateAdd("d", 0, CDate(DTP_PROD_TO.Text)), "YYYY-MM-DD")

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Combo_ORD_ITEM_LostFocus()
    
    Dim S As String
  
    If Len(cbo_ord_item.Text) = 1 Then
        S = cbo_ord_item.Text
        cbo_ord_item.Text = "0" + S
    End If
    
End Sub

Private Sub Opt_Click()

End Sub

Private Sub opt_all_Click()

    If opt_all.Value = True Then
        opt_all.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = ""
    Else
        opt_all.ForeColor = &H80000012
    End If
    
    If Opt_rk_y.Value = True Then
        Opt_rk_y.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = "Y"
    Else
        Opt_rk_y.ForeColor = &H80000012
    End If
    
    If Opt_rk_n.Value = True Then
        Opt_rk_n.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = "N"
    Else
        Opt_rk_n.ForeColor = &H80000012
    End If

End Sub

Private Sub Opt_rk_n_Click()
    If opt_all.Value = True Then
        opt_all.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = ""
    Else
        opt_all.ForeColor = &H80000012
    End If
    
    If Opt_rk_y.Value = True Then
        Opt_rk_y.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = "Y"
    Else
        Opt_rk_y.ForeColor = &H80000012
    End If
    
    If Opt_rk_n.Value = True Then
        Opt_rk_n.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = "N"
    Else
        Opt_rk_n.ForeColor = &H80000012
    End If
End Sub

Private Sub Opt_rk_y_Click()
    If opt_all.Value = True Then
        opt_all.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = ""
    Else
        opt_all.ForeColor = &H80000012
    End If
    
    If Opt_rk_y.Value = True Then
        Opt_rk_y.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = "Y"
    Else
        Opt_rk_y.ForeColor = &H80000012
    End If
    
    If Opt_rk_n.Value = True Then
        Opt_rk_n.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = "N"
    Else
        Opt_rk_n.ForeColor = &H80000012
    End If
End Sub

Private Sub sdb_len_fr_Change()
    If sdb_len_fr.Value > 0 And sdb_len_to.Value < sdb_len_fr.Value Then
        sdb_len_to.Value = sdb_len_fr.Value
    End If
End Sub

Private Sub sdb_thk_fr_Change()
    If sdb_thk_fr.Value > 0 And sdb_thk_to.Value < sdb_thk_fr.Value Then
        sdb_thk_to.Value = sdb_thk_fr.Value
    End If
End Sub

Private Sub sdb_wid_fr_Change()
    If sdb_wid_fr.Value > 0 And sdb_wid_to.Value < sdb_wid_fr.Value Then
        sdb_wid_to.Value = sdb_wid_fr.Value
    End If
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    
End Sub



Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)

    Dim i As Integer
    
    Dim iOrd_no As String
    Dim iOrd_item As String

    If Col = 0 And Row > 0 Then

        For i = 1 To ss2.MaxRows
           ss2.Col = 0
           ss2.Text = ""
           Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, i, i)
        Next
       Call ss2_row_Click(Col, Row)
       ss2.Row = ss2.ActiveRow:    ss2.Col = 1:            iOrd_no = ss2.Text
                                   ss2.Col = 2:            iOrd_item = ss2.Text
       ss1.Row = iOrd_row:         ss1.Col = SPD_ORD_NO:   ss1.Text = iOrd_no
                                   ss1.Col = SPD_ORD_ITEM: ss1.Text = iOrd_item
       
    End If
End Sub


Private Sub SSC_BED_PILE_DATE_Click(Value As Integer)

End Sub

Private Sub ssc_thk_cd_Click(Value As Integer)
    If ssc_thk_cd.Value = -1 Then
       ssc_thk_cd.ForeColor = &HFF&
       txt_thk_cd = "Y"
    Else
       ssc_thk_cd.ForeColor = &H808080
       txt_thk_cd = "N"
    End If
End Sub

Private Sub SSC_UST_FL_Click(Value As Integer)
    If SSC_UST_FL.Value = -1 Then
       SSC_UST_FL.ForeColor = &HFF&
       TXT_UST_FL = "Y"
    Else
       SSC_UST_FL.ForeColor = &H808080
       TXT_UST_FL = "N"
    End If
End Sub

Private Sub text_cur_inv_code_DblClick()

    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_enduse_cd_DblClick()

    Call txt_enduse_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_enduse_cd_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Err_Track:

              
    If KeyCode = vbKeyF4 Then
                 
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_enduse_cd
        DD.nameType = "2"
            
        Call Gf_Usage_DD(M_CN1, KeyCode)
        
    End If
    
Err_Track:
    
End Sub

Private Sub TXT_HTM_DblClick()

    Call TXT_HTM_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TXT_HTM_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Q0073"
        
        DD.rControl.Add Item:=TXT_HTM
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If
    
End Sub

Private Sub txt_ord_no_LostFocus()

'   If Len(TXT_ORD_NO.Text) >= 2 Then
'      If Mid(TXT_ORD_NO.Text, 1, 2) <> "OD" Then
'         TXT_ORD_NO = ""
'         Call MsgBox("������Ӧ��Ϊ�ƻ���������ȷ��", vbExclamation + vbOKOnly, "����")
'      End If
'   End If
   
End Sub

Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuTool_ReSet

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Spl_SizeSet(SSSplitter1, "C-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(ss1, "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(ss2, "C-System.INI", Me.Name)
    
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
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) And Gf_Sp_Cls(sc2) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        cbo_ord_item.Clear
    End If
    
    If App.Title = "CE" Then
        text_cur_inv_code.Text = "ZB"
        CBO_PROD_CD.Text = "SL"
    End If
    
    Call text_cur_inv_code_KeyUp(0, 0)
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1, True) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        Call Gf_Sp_Cls(sc2)
        'Call Gp_Sp_EvenRowBackcolor(Proc_Sc("SC").Item("Spread"), 1)
    End If
       
End Sub

Public Sub Form_Ref()

     Dim SMESG As String
     Dim S As String
     
     If CBO_PLT.Text = "C1" Then
        Call Gp_Sp_ColHidden(ss1, 3, True)
     Else
        Call Gp_Sp_ColHidden(ss1, 3, False)
     End If
     
     If CBO_PROD_CD.Text = "" Then
        Call Gp_Sp_ColHidden(ss1, 2, False)
     Else
        Call Gp_Sp_ColHidden(ss1, 2, True)
     End If
     
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl")) Then
        Call Gf_Sp_Cls(sc2)
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If

End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    'Call Gp_Sp_EvenRowBackcolor(Proc_Sc("SC").Item("Spread"), 1)
    
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

Public Sub Form_Exit()

    Unload Me
    
End Sub

Private Sub text_cur_inv_code_Change()

    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
          text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
          Exit Sub
    Else
          text_cur_inv.Text = ""
    End If
    
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
           DD.sWitch = "MS"
           DD.sKey = "C0013"
    
           DD.rControl.Add Item:=text_cur_inv_code
           DD.rControl.Add Item:=text_cur_inv
           
    
           DD.nameType = "2"
           Call Gf_Common_DD(M_CN1, KeyCode)
    
    End If
    
    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
        text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
        Exit Sub
    Else
        text_cur_inv.Text = ""
    End If
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_row_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If Row < 1 Then Exit Sub
    If ss1.MaxRows < 1 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 0
    
    ss1.ReDraw = False
    
    If ss1.Text <> "Update" Then
                
        ss1.Text = "Update"
        iOrd_row = Row
        ss1.Col = 40
        ss1.Text = sUserID
        Call Gp_Sp_BlockColor(ss1, 1, -1, Row, Row, , &HFFFF80)
        
        If Row > 0 Then
        
            If ss1.MaxRows = 0 Then Exit Sub
            
            If prod_txt_prod_cd = "" Then
               prod_txt_prod_cd.Text = CBO_PROD_CD.Text
            End If
               ss1.Col = SPD_APLY_STDSPEC:     ss1.Row = ss1.ActiveRow
               txt_spec.Text = ss1.Text
               ss1.Col = SPD_THK:              ss1.Row = ss1.ActiveRow
               sdb_thk.Text = ss1.Text
               ss1.Col = SPD_WID:              ss1.Row = ss1.ActiveRow
               sdb_wid.Text = ss1.Text
               ss1.Col = SPD_LEN:              ss1.Row = ss1.ActiveRow
               sdb_len.Text = ss1.Text
            
            If Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl")) Then
                ss2.OperationMode = OperationModeNormal
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Exit Sub
            End If
            
        End If
        
    Else
       
        ss1.Col = 0
        ss1.Text = ""
        iOrd_row = 0
        ss1.Col = 40
        ss1.Text = ""
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
       
    End If
    ss1.ReDraw = True
    
End Sub
Private Sub ss2_row_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If Row < 1 Then Exit Sub
    If ss2.MaxRows < 1 Then Exit Sub
    
    ss2.Row = Row
    ss2.Col = 0
    
    ss2.ReDraw = False
    
    If ss2.Text <> "Update" Then
        
        ss2.Text = "Update"
        Call Gp_Sp_BlockColor(ss2, 1, -1, Row, Row, , &HFFFF80)
        
    Else
       
        ss2.Text = ""
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, Row, Row)
       
    End If
    ss2.ReDraw = True
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 Then
      Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    End If
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If Col = 0 Then
    
        Dim i As Integer
        For i = 1 To ss1.MaxRows
           ss1.Col = 0
           ss1.Text = ""
           Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, i, i)
        Next
    
       Call ss1_row_Click(Col, Row)
       
    End If
    
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub Text_size_knd_DblClick()

    Call Text_size_knd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_stlgrd_DblClick()

    Call text_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=text_stlgrd
           
        If CBO_PROD_CD.Text = "SL" Then
            DD.nameType = "1"
            Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        Else
            Call Gf_StdSPEC_DD2(M_CN1, KeyCode)
        End If
    End If
        
End Sub

Private Sub txt_prod_grd_DblClick()

    Call txt_prod_grd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_prod_grd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0034"

        DD.rControl.Add Item:=txt_prod_grd

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
    
End Sub

Private Sub Text_size_knd_Change()

    If Len(Trim(Text_size_knd.Text)) = Text_size_knd.MaxLength Then
        Text_size_knd_name.Text = Gf_ComnNameFind(M_CN1, "B0043", Text_size_knd.Text, 2)
        Exit Sub
    Else
        Text_size_knd_name.Text = ""
    End If
    
End Sub

Private Sub Text_size_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"

        DD.rControl.Add Item:=Text_size_knd

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
 
    If Col <> SPD_MAT_NO Then Exit Sub
    
    If Row > 0 And Col > 0 Then
    
        If ss1.MaxRows = Row Then Exit Sub
        
        If prod_txt_prod_cd = "" Then
           prod_txt_prod_cd.Text = CBO_PROD_CD.Text
        End If
           ss1.Col = SPD_APLY_STDSPEC:     ss1.Row = ss1.ActiveRow
           txt_spec.Text = ss1.Text
           ss1.Col = SPD_THK:              ss1.Row = ss1.ActiveRow
           sdb_thk.Text = ss1.Text
           ss1.Col = SPD_WID:              ss1.Row = ss1.ActiveRow
           sdb_wid.Text = ss1.Text
           ss1.Col = SPD_LEN:              ss1.Row = ss1.ActiveRow
           sdb_len.Text = ss1.Text
        
        If Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl")) Then
            ss2.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Exit Sub
        End If
        
    End If

End Sub

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String

    If Len(Trim(txt_ORD_NO.Text)) = txt_ORD_NO.MaxLength Then
    
        If cbo_ord_item.Text <> "" Then Exit Sub
        
        txt_ORD_NO.Text = StrConv(txt_ORD_NO.Text, vbUpperCase)
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(txt_ORD_NO.Text) & "'"
        Call Gf_ComboAdd(M_CN1, cbo_ord_item, sQuery)

    Else
        cbo_ord_item.Clear
    End If

End Sub

Private Sub txt_trim_fl_Change()

    If Len(Trim(txt_TRIM_FL.Text)) = txt_TRIM_FL.MaxLength Then
        txt_TRIM_NAME.Text = Gf_ComnNameFind(M_CN1, "B0021", txt_TRIM_FL.Text, 2)
        txt_TRIM_FL.Text = Trim(txt_TRIM_FL.Text)
        Exit Sub
    Else
        txt_TRIM_NAME.Text = ""
        txt_TRIM_FL.Text = ""
    End If

End Sub

Private Sub txt_trim_fl_DblClick()

    Call txt_trim_fl_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_trim_fl_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0021"

        DD.rControl.Add Item:=txt_TRIM_FL

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If

End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(11).Enabled = False                'Spread Copy
        .Buttons(12).Enabled = False                'Paste
    End With

End Sub