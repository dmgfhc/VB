VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACE6020C 
   Caption         =   "����Ʒ�������(������׼)_ACE6020C"
   ClientHeight    =   9225
   ClientLeft      =   330
   ClientTop       =   2235
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text_PROC_CD 
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
      Left            =   13020
      MaxLength       =   3
      TabIndex        =   32
      Tag             =   "CD_MANA_NO"
      Top             =   105
      Width           =   915
   End
   Begin VB.ComboBox Text_GROUP 
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
      ItemData        =   "ACE6020C.frx":0000
      Left            =   12915
      List            =   "ACE6020C.frx":0002
      TabIndex        =   31
      Top             =   870
      Width           =   780
   End
   Begin VB.TextBox txt_plt_nm 
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
      Left            =   1875
      Locked          =   -1  'True
      TabIndex        =   21
      Tag             =   "����Ͷ��"
      Top             =   105
      Width           =   1590
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
      Left            =   1410
      MaxLength       =   2
      TabIndex        =   20
      Tag             =   "����Ͷ��"
      Top             =   110
      Width           =   465
   End
   Begin VB.TextBox txt_enduse_cd 
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
      Left            =   5445
      MaxLength       =   4
      TabIndex        =   19
      Top             =   480
      Width           =   585
   End
   Begin VB.TextBox txt_cust_nm 
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
      Left            =   10215
      TabIndex        =   18
      Top             =   105
      Width           =   1275
   End
   Begin VB.TextBox txt_cust 
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
      Left            =   9375
      MaxLength       =   6
      TabIndex        =   17
      Top             =   105
      Width           =   840
   End
   Begin VB.ComboBox cbo_ord_item 
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
      Left            =   6825
      TabIndex        =   16
      Top             =   105
      Width           =   660
   End
   Begin VB.TextBox txt_ord_no 
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
      Left            =   5445
      MaxLength       =   11
      TabIndex        =   15
      Tag             =   "CD_MANA_NO"
      Top             =   105
      Width           =   1380
   End
   Begin VB.TextBox txt_stdspec 
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
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "��׼��"
      Top             =   480
      Width           =   2250
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7950
      Left            =   60
      TabIndex        =   0
      Top             =   1260
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   14023
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACE6020C.frx":0004
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   3525
         Left            =   0
         TabIndex        =   22
         Top             =   4425
         Width           =   15120
         _ExtentX        =   26670
         _ExtentY        =   6218
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   14737632
         PaneTree        =   "ACE6020C.frx":0056
         Begin FPSpread.vaSpread ss1 
            Height          =   2955
            Left            =   0
            TabIndex        =   23
            Top             =   570
            Width           =   15120
            _Version        =   393216
            _ExtentX        =   26670
            _ExtentY        =   5212
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColsFrozen      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   31
            MaxRows         =   1
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "ACE6020C.frx":00A8
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   540
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   15120
            _ExtentX        =   26670
            _ExtentY        =   953
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_ord_no1 
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
               Left            =   5490
               MaxLength       =   11
               TabIndex        =   28
               Tag             =   "CD_MANA_NO"
               Top             =   90
               Visible         =   0   'False
               Width           =   420
            End
            Begin VB.TextBox txt_ord_item1 
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
               Left            =   6000
               TabIndex        =   27
               Tag             =   "CD_MANA_NO"
               Top             =   90
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.TextBox txt_cur_inv 
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
               Height          =   310
               Left            =   1350
               MaxLength       =   2
               TabIndex        =   26
               Top             =   90
               Width           =   435
            End
            Begin VB.TextBox txt_cur_inv_nm 
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
               Left            =   1800
               TabIndex        =   25
               Top             =   90
               Width           =   3075
            End
            Begin CSTextLibCtl.sidbEdit sdb_mat_wgt 
               Height          =   315
               Left            =   13620
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   90
               Width           =   1380
               _Version        =   262145
               _ExtentX        =   2434
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
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
               DataProperty    =   2
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
               NumIntDigits    =   12
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel2 
               Height          =   315
               Index           =   0
               Left            =   12360
               Top             =   90
               Width           =   1230
               _ExtentX        =   2170
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
            Begin InDate.ULabel ULabel9 
               Height          =   315
               Left            =   9450
               Top             =   90
               Width           =   1230
               _ExtentX        =   2170
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
            Begin CSTextLibCtl.sidbEdit sdb_ord_wgt 
               Height          =   315
               Left            =   10710
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   90
               Width           =   1380
               _Version        =   262145
               _ExtentX        =   2434
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   16711680
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
               NumIntDigits    =   12
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel12 
               Height          =   315
               Left            =   90
               Top             =   90
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   556
               Caption         =   "�ѷŲֿ�"
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
         End
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   4365
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   15120
         _Version        =   393216
         _ExtentX        =   26670
         _ExtentY        =   7699
         _StockProps     =   64
         ColsFrozen      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   24
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACE6020C.frx":0FAC
      End
   End
   Begin InDate.ULabel ULabel3 
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   6315
      Top             =   495
      Width           =   1230
      _ExtentX        =   2170
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   150
      Top             =   480
      Width           =   1230
      _ExtentX        =   2170
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
      Left            =   1410
      TabIndex        =   2
      Top             =   870
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
   Begin InDate.UDate dpt_del_fr 
      Height          =   315
      Left            =   7575
      TabIndex        =   3
      Tag             =   "INS_DATE"
      Top             =   495
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
   Begin InDate.UDate dpt_del_to 
      Height          =   315
      Left            =   9105
      TabIndex        =   4
      Tag             =   "INS_DATE"
      Top             =   495
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
      Left            =   150
      Top             =   870
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "��Ʒ���"
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
      Left            =   4170
      Top             =   870
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "��Ʒ����"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Index           =   0
      Left            =   8115
      Top             =   870
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "��Ʒ����"
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
      Left            =   5445
      TabIndex        =   5
      Top             =   870
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
   Begin CSTextLibCtl.sidbEdit sdb_len_fr 
      Height          =   315
      Left            =   9375
      TabIndex        =   6
      Top             =   870
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
   Begin CSTextLibCtl.sidbEdit sdb_thk_to 
      Height          =   315
      Left            =   2610
      TabIndex        =   7
      Top             =   870
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
      Left            =   6630
      TabIndex        =   8
      Top             =   870
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
      Left            =   10560
      TabIndex        =   9
      Top             =   870
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   4170
      Top             =   105
      Width           =   1230
      _ExtentX        =   2170
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   8115
      Top             =   105
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "�ͻ�"
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
      Index           =   0
      Left            =   4170
      Top             =   480
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
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   150
      Top             =   105
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "����Ͷ��"
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   11760
      Top             =   105
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "����״̬"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   11640
      Top             =   870
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
         Size            =   9.75
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
      Left            =   10800
      Top             =   480
      Width           =   1260
      _ExtentX        =   2223
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
   Begin InDate.UDate Udate_BB_CONF_TO 
      Height          =   315
      Left            =   13740
      TabIndex        =   33
      Tag             =   "������"
      Top             =   480
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
      MaxLength       =   10
   End
   Begin InDate.UDate Udate_BB_CONF_FR 
      Height          =   315
      Left            =   12090
      TabIndex        =   34
      Tag             =   "������"
      Top             =   480
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
      MaxLength       =   10
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   120
      Left            =   13590
      TabIndex        =   35
      Top             =   585
      Width           =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   120
      Left            =   8985
      TabIndex        =   13
      Top             =   570
      Width           =   90
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   120
      Left            =   2475
      TabIndex        =   12
      Top             =   975
      Width           =   90
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   120
      Left            =   6525
      TabIndex        =   11
      Top             =   975
      Width           =   90
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   120
      Left            =   10440
      TabIndex        =   10
      Top             =   960
      Width           =   90
   End
End
Attribute VB_Name = "ACE6020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO PRODUCTION MANAGEMENT SYSTEM
'-- Sub_System Name   PROGRESS MANAGEMENT SYSTEM
'-- Program Name      SUBSTITUTE WORK IN PROCESS
'-- Program ID        ACE6020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2009.10.12
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
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iCurr_Row1 As Integer           'SS1 Current Row
Dim iCurr_Row2 As Integer           'SS2 Current Row

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_FORCE_CD = 26             ' �Ƿ�ǿ�����    24->26
Const SS1_REMARK = 27               ' �����ע        25->27
Const SS1_ORDER_NO = 28             ' �������        26->28
Const SS1_ORDER_ITEM = 29           ' �������        27->29
Const SS1_USER_ID = 30              ' �����          28->30
Const SS1_REASON = 31               ' ��עԭ��


Private Sub Form_Define()

    Dim i As Integer
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(txt_plt_nm, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(cbo_ord_item, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(txt_cust, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
      Call Gp_Ms_Collection(txt_cust_nm, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(dpt_del_fr, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(dpt_del_to, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
      Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(txt_enduse_cd, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(sdb_thk_fr, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(SDB_THK_TO, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(sdb_wid_fr, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(SDB_WID_TO, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(sdb_len_fr, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(SDB_LEN_TO, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
      Call Gp_Ms_Collection(sdb_ord_wgt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
      Call Gp_Ms_Collection(sdb_mat_wgt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       
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
       Call Gp_Ms_Collection(txt_ord_no1, "p", " ", " ", " ", "r", " ", "l", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
     Call Gp_Ms_Collection(txt_ord_item1, "p", " ", " ", " ", "r", " ", "l", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
       Call Gp_Ms_Collection(txt_cur_inv, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
    Call Gp_Ms_Collection(txt_cur_inv_nm, " ", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
  Call Gp_Ms_Collection(Udate_BB_CONF_FR, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
  Call Gp_Ms_Collection(Udate_BB_CONF_TO, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
      Call Gp_Ms_Collection(Text_PROC_CD, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
        Call Gp_Ms_Collection(Text_GROUP, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
        
    'MASTER Collection
    Mc2.Add Item:=pContro2, Key:="pControl"
    Mc2.Add Item:=nContro2, Key:="nControl"
    Mc2.Add Item:=mContro2, Key:="mControl"
    Mc2.Add Item:=iContro2, Key:="iControl"
    Mc2.Add Item:=rContro2, Key:="rControl"
    Mc2.Add Item:=cContro2, Key:="cControl"
    Mc2.Add Item:=aContro2, Key:="aControl"
    Mc2.Add Item:=lContro2, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    
    For i = 2 To ss1.MaxCols - 6
        Call Gp_Sp_Collection(ss1, i, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next i
    
    
    Call Gp_Sp_Collection(ss1, SS1_FORCE_CD, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
      Call Gp_Sp_Collection(ss1, SS1_REMARK, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, SS1_ORDER_NO, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, SS1_ORDER_ITEM, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, SS1_USER_ID, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
      Call Gp_Sp_Collection(ss1, SS1_REASON, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACE6020C.P_REFER1", Key:="P-R"
    sc1.Add Item:="ACE6020C.P_MODIFY1", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, "p", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    For i = 3 To ss2.MaxCols
        Call Gp_Sp_Collection(ss2, i, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next i
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ACE6020C.P_REFER2", Key:="P-R"
    sc2.Add Item:="ACE6020C.P_ONEROW2", Key:="P-O"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 4, True)
    Call Gp_Sp_ColHidden(ss1, SS1_ORDER_NO, True)
    Call Gp_Sp_ColHidden(ss1, SS1_ORDER_ITEM, True)
    Call Gp_Sp_ColHidden(ss1, SS1_USER_ID, True)
    
    Call Gp_Sp_ColHidden(ss2, 6, True)
    Call Gp_Sp_ColHidden(ss2, 7, True)
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").ROW = 0
    sc1.Item("Spread").Text = "��"
    
    iCurr_Row1 = 0
    iCurr_Row2 = 0
    
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

Private Sub Form_Load()

    Text_GROUP.AddItem "A"
    Text_GROUP.AddItem "B"
    Text_GROUP.AddItem "C"
    Text_GROUP.AddItem "D"

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    
'    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    If App.Title = "AC" Then
        txt_plt.Text = "C1"
    ElseIf App.Title = "BG" Then
        txt_plt.Text = "C1"
    ElseIf App.Title = "CG" Then
        txt_plt.Text = "C3"
    End If

    Call txt_plt_KeyUp(0, 0)

    dpt_del_fr.RawData = ""
    dpt_del_to.RawData = ""
    
    Call Gp_Spl_SizeGet(SSSplitter1, "C-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "C-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    
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
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
        If Gf_Sp_Cls(sc2) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gp_Ms_Cls(Mc2("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call MenuTool_ReSet
            rContro1(1).SetFocus
            iCurr_Row1 = 0
            iCurr_Row2 = 0
            
            If App.Title = "AC" Then
                txt_plt.Text = "C1"
            ElseIf App.Title = "BG" Then
                txt_plt.Text = "C1"
            ElseIf App.Title = "CG" Then
                txt_plt.Text = "C3"
            End If
        
            Call txt_plt_KeyUp(0, 0)
            
            dpt_del_fr.RawData = ""
            dpt_del_to.RawData = ""
        End If
    End If
    
End Sub

Public Sub Form_Ref()

    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc2, Mc1) Then
        ss2.OperationMode = OperationModeNormal
        Call Gp_Sp_EvenRowBackcolor(ss2)
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        Call Gf_Sp_Cls(sc1)
        Call Gp_Ms_Cls(Mc2("rControl"))
        iCurr_Row1 = 0
        iCurr_Row2 = 0
        sdb_ord_wgt.Value = 0
        sdb_mat_wgt.Value = 0
    End If
            
End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, sc1, Mc2, False) Then
        ss1.OperationMode = OperationModeNormal
        Call Gp_Sp_EvenRowBackcolor(ss1)
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        Call Gp_Sp_OneRowDisplay(M_CN1, Gf_Sp_MakeQuery(sc2.Item("Spread"), sc2.Item("P-O"), "O", sc2.Item("pColumn"), iCurr_Row2), sc2.Item("Spread"), iCurr_Row2)
        ss2.ROW = iCurr_Row2
        ss2.Col = ss2.MaxCols
        sdb_ord_wgt.Value = ss2.Value
        sdb_mat_wgt.Value = 0
    End If
    
End Sub

Public Sub Form_Ins()
    
End Sub

Public Sub Spread_Cpy()

End Sub

Public Sub Spread_Pst()

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
    
    Call Gp_Sp_Excel(Me, ss2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Spread_Del()
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If ss1.MaxRows < 1 Or ROW < 1 Then Exit Sub
    
    ss1.ROW = ROW
    ss1.Col = 0
    
    If iCurr_Row2 = 0 Then Exit Sub
    
    If ss1.Text = "Update" Then
    
        ss1.Col = 0
        ss1.Text = ""
        ss1.Col = SS1_ORDER_NO
        ss1.Text = ""
        ss1.Col = SS1_ORDER_ITEM
        ss1.Text = ""
        ss1.Col = SS1_USER_ID
        ss1.Text = ""
        ss1.Col = 10
        sdb_mat_wgt.Value = sdb_mat_wgt.Value - ss1.Value
        
        If ROW Mod 2 <> 0 Then
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW, , &HF2F2F2)
        Else
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW)
        End If
        
    Else
    
        ss1.Col = 10
        'If sdb_ord_wgt.Value < sdb_mat_wgt.Value + ss1.Value Then Exit Sub
    
        ss2.ROW = iCurr_Row2
        ss2.Col = 1
        txt_ord_no1.Text = ss2.Text
        ss2.Col = 2
        txt_ord_item1.Text = ss2.Text
        ss2.Col = ss2.MaxCols
    
        ss1.ROW = ROW
        ss1.Col = 0
        ss1.Text = "Update"
        ss1.Col = SS1_ORDER_NO
        ss1.Text = txt_ord_no1.Text
        ss1.Col = SS1_ORDER_ITEM
        ss1.Text = txt_ord_item1.Text
        ss1.Col = SS1_USER_ID
        ss1.Text = sUserID
        ss1.Col = 10
        sdb_mat_wgt.Value = sdb_mat_wgt.Value + ss1.Value
    
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW, , &HFFFF80)
        
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

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If ss2.MaxRows < 1 Or ROW < 1 Then Exit Sub
    
    If iCurr_Row2 <> 0 Then
        
        If iCurr_Row2 Mod 2 <> 0 Then
            Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, iCurr_Row2, iCurr_Row2, , &HF2F2F2)
        Else
            Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, iCurr_Row2, iCurr_Row2)
        End If
    
    End If
    
    ss2.ROW = ROW
    ss2.Col = 1
    txt_ord_no1.Text = ss2.Text
    ss2.Col = 2
    txt_ord_item1.Text = ss2.Text
    ss2.Col = ss2.MaxCols
    sdb_ord_wgt.Value = ss2.Value
    
    iCurr_Row2 = ROW
    Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, iCurr_Row2, iCurr_Row2, , &HFFFF80)
    
    Call Gf_Sp_Refer(M_CN1, sc1, Mc2, , , False)
    ss1.OperationMode = OperationModeNormal
    Call Gp_Sp_EvenRowBackcolor(ss1)
    iCurr_Row1 = 0
    sdb_mat_wgt.Value = 0
    
End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If ROW > 0 Then
        Set Active_Spread = Me.ss2
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If

End Sub

Private Sub Text_PROC_CD_DblClick()

    Call Text_PROC_CD_KeyUp(vbKeyF4, 0)
    
End Sub
Private Sub Text_PROC_CD_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "C0004"

        DD.rControl.Add Item:=Text_PROC_CD
'        DD.rControl.Add Item:=Text_PROC_CD_Name
   
        DD.nameType = "2"
        'DD.nameType="1" ���������Ʋ�ѯ
        'DD.nameType="2" ��Ӣ�����Ʋ�ѯ
       
        Call Gf_Common_DD(M_CN1, KeyCode)

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() ���ڿͻ�����
        Exit Sub
        
    End If

'    If Len(Trim(Text_PROC_CD.Text)) = Text_PROC_CD.MaxLength Then
'       '  Gf_ComnNAME_Find( �����ַ���, DD.sKEy���� ,DD.nameType)
'       ' Gf_CustNameFind( �����ַ���, �ͻ���������,DD.nameType)
'        Text_PROC_CD_Name.Text = Gf_ComnNameFind(M_CN1, "C0004", Text_PROC_CD.Text, 2)
'    Else
'        Text_PROC_CD_Name.Text = ""
'    End If
    
End Sub

Private Sub txt_cur_inv_DblClick()

    Call txt_cur_inv_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_cur_inv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
           DD.sWitch = "MS"
           DD.sKey = "C0013"
    
           DD.rControl.Add Item:=txt_cur_inv
           DD.rControl.Add Item:=txt_cur_inv_nm
    
           DD.nameType = "2"
           Call Gf_Common_DD(M_CN1, KeyCode)
    
    Else
    
        If Len(Trim(txt_cur_inv.Text)) = txt_cur_inv.MaxLength Then
            txt_cur_inv_nm.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv.Text, 2)
        Else
            txt_cur_inv_nm.Text = ""
        End If
        
    End If

End Sub

Private Sub txt_cust_DblClick()

    Call txt_cust_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_cust_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_cust
        DD.rControl.Add Item:=txt_cust_nm

        DD.nameType = "1"
        Call Gf_Customer_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_cust)) = txt_cust.MaxLength Then
            txt_cust_nm.Text = Gf_CodeFind(M_CN1, "SELECT  CUST_NM FROM NISCO.BP_CUST_CD WHERE CUST_CD = '" & Trim(txt_cust.Text) & "'")
        Else
            txt_cust_nm.Text = ""
        End If

    End If

End Sub

Private Sub txt_enduse_cd_DblClick()

    Call txt_enduse_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_enduse_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
                 
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_enduse_cd
        DD.nameType = "2"
            
        Call Gf_Usage_DD(M_CN1, KeyCode)
        
    End If

End Sub

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String

    If Len(Trim(txt_ord_no.Text)) = txt_ord_no.MaxLength Then
    
        If cbo_ord_item.Text <> "" Then Exit Sub
        
        txt_ord_no.Text = StrConv(txt_ord_no.Text, vbUpperCase)
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(txt_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, cbo_ord_item, sQuery)

    Else
        cbo_ord_item.Clear
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
        DD.rControl.Add Item:=txt_plt_nm

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

    Else

        If Len(Trim(txt_plt)) = txt_plt.MaxLength Then
            txt_plt_nm.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
        Else
            txt_plt_nm.Text = ""
        End If
    
    End If

End Sub

Private Sub txt_stdspec_DblClick()

    Call txt_stdspec_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stdspec_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec
        DD.nameType = "1"
        
        Call Gf_StdSPEC_DD(M_CN1, KeyCode)
        
    End If
    
End Sub

Private Sub MenuTool_ReSet()

    MDIMain.MenuTool.Buttons(7).Enabled = False              'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False              'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False              'Row Cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False             'Row Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False             'Row Paste
    
End Sub