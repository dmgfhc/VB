VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AEH2010C 
   Caption         =   "���Ϸֶ���ҵָʾ_AEH2010C"
   ClientHeight    =   8850
   ClientLeft      =   240
   ClientTop       =   2040
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
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
      Left            =   10365
      TabIndex        =   36
      Top             =   480
      Width           =   1800
   End
   Begin VB.TextBox txt_ord_fl 
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
      Left            =   13530
      TabIndex        =   25
      Top             =   1230
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   1275
      MaxLength       =   2
      TabIndex        =   14
      Text            =   "B1"
      Top             =   90
      Width           =   360
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
      Left            =   6990
      TabIndex        =   13
      Top             =   90
      Width           =   660
   End
   Begin VB.TextBox txt_ord_no 
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
      Left            =   5685
      MaxLength       =   11
      TabIndex        =   12
      Tag             =   "��Ʒ"
      Top             =   95
      Width           =   1305
   End
   Begin VB.TextBox txt_plt_name 
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
      Left            =   1635
      TabIndex        =   11
      Tag             =   "����"
      Top             =   95
      Width           =   1815
   End
   Begin VB.TextBox txt_slab_no 
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
      Left            =   13800
      TabIndex        =   10
      Top             =   1230
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txt_prod_no 
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
      Left            =   1275
      MaxLength       =   10
      TabIndex        =   0
      Top             =   480
      Width           =   1245
   End
   Begin VB.TextBox txt_loc 
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
      Left            =   10365
      MaxLength       =   10
      TabIndex        =   3
      Top             =   870
      Width           =   1170
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
      Left            =   5685
      MaxLength       =   11
      TabIndex        =   1
      Top             =   485
      Width           =   1275
   End
   Begin VB.TextBox txt_stlgrd_name 
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
      Left            =   6960
      MaxLength       =   11
      TabIndex        =   2
      Top             =   485
      Width           =   1755
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_thk_fr 
      Height          =   315
      Left            =   1275
      TabIndex        =   4
      Top             =   1260
      Width           =   1365
      _Version        =   262145
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   240
      Top             =   1260
      Width           =   990
      _ExtentX        =   1746
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   4650
      Top             =   1260
      Width           =   990
      _ExtentX        =   1746
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   1
      Left            =   9330
      Top             =   1260
      Width           =   990
      _ExtentX        =   1746
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_thk_to 
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   1260
      Width           =   1365
      _Version        =   262145
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_len_fr 
      Height          =   315
      Left            =   10365
      TabIndex        =   7
      Top             =   1260
      Width           =   1365
      _Version        =   262145
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_len_to 
      Height          =   315
      Left            =   11730
      TabIndex        =   8
      Top             =   1260
      Width           =   1365
      _Version        =   262145
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumIntDigits    =   7
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_wid_fr 
      Height          =   315
      Left            =   5685
      TabIndex        =   9
      Top             =   1260
      Width           =   1365
      _Version        =   262145
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_wid_to 
      Height          =   315
      Left            =   7050
      TabIndex        =   6
      Top             =   1260
      Width           =   1365
      _Version        =   262145
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   4650
      Top             =   480
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   9330
      Top             =   870
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "����λ��"
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
      Index           =   2
      Left            =   240
      Top             =   480
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "���Ϻ�"
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
      Left            =   4650
      Top             =   90
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   240
      Top             =   90
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "�� ��"
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
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7530
      Left            =   30
      TabIndex        =   15
      Top             =   1650
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   13282
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AEH2010C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   2940
         Left            =   0
         TabIndex        =   16
         Top             =   4590
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   5186
         _Version        =   196609
         SplitterBarWidth=   3
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   14737632
         PaneTree        =   "AEH2010C.frx":0052
         Begin Threed.SSPanel SSPanel1 
            Height          =   540
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   15240
            _ExtentX        =   26882
            _ExtentY        =   953
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.ComboBox cbo_slab_cut 
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
               ItemData        =   "AEH2010C.frx":00A4
               Left            =   4260
               List            =   "AEH2010C.frx":00A6
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   120
               Width           =   810
            End
            Begin InDate.ULabel ULabel6 
               Height          =   315
               Left            =   3120
               Top             =   120
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               Caption         =   "�и���"
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
               ForeColor       =   0
            End
            Begin InDate.ULabel ULabel11 
               Height          =   315
               Index           =   0
               Left            =   5220
               Top             =   450
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
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
            Begin CSTextLibCtl.sidbEdit sdb_slab_len 
               Height          =   315
               Left            =   6345
               TabIndex        =   18
               Top             =   450
               Visible         =   0   'False
               Width           =   1185
               _Version        =   262145
               _ExtentX        =   2090
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   16711680
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
               ReadOnly        =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   "0.0"
               Text            =   " 0.0"
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
               MaxValue        =   9999.99
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel7 
               Height          =   315
               Index           =   3
               Left            =   5340
               Top             =   120
               Width           =   1095
               _ExtentX        =   1931
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
            Begin CSTextLibCtl.sidbEdit sdb_slab_wgt 
               Height          =   315
               Left            =   6480
               TabIndex        =   19
               Top             =   120
               Width           =   1185
               _Version        =   262145
               _ExtentX        =   2090
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   16711680
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
               MaxValue        =   9999.99
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin InDate.UDate udt_ins_date 
               Height          =   315
               Left            =   1320
               TabIndex        =   26
               Tag             =   "ָʾ����"
               Top             =   120
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
               BackColor       =   12648447
               MaxLength       =   10
            End
            Begin InDate.ULabel ULabel12 
               Height          =   315
               Left            =   180
               Top             =   120
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               Caption         =   "ָʾ����"
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
               ForeColor       =   0
            End
            Begin InDate.ULabel ULabel11 
               Height          =   315
               Index           =   1
               Left            =   10440
               Top             =   120
               Visible         =   0   'False
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               Caption         =   "�и��"
               Alignment       =   1
               BackColor       =   16761087
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
            Begin CSTextLibCtl.sidbEdit sdb_cut_len 
               Height          =   315
               Left            =   11565
               TabIndex        =   30
               Top             =   120
               Visible         =   0   'False
               Width           =   1185
               _Version        =   262145
               _ExtentX        =   2090
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   255
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
               ReadOnly        =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   "0.0"
               Text            =   " 0.0"
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
               MaxValue        =   9999.99
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel7 
               Height          =   315
               Index           =   0
               Left            =   7950
               Top             =   120
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               Caption         =   "�и�����"
               Alignment       =   1
               BackColor       =   16761087
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
            Begin CSTextLibCtl.sidbEdit sdb_cut_wgt 
               Height          =   315
               Left            =   9090
               TabIndex        =   31
               Top             =   120
               Width           =   1185
               _Version        =   262145
               _ExtentX        =   2090
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   255
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
               MaxValue        =   9999.99
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   2355
            Left            =   0
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   585
            Width           =   15240
            _Version        =   393216
            _ExtentX        =   26882
            _ExtentY        =   4154
            _StockProps     =   64
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
            MaxCols         =   10
            MaxRows         =   2
            ProcessTab      =   -1  'True
            Protect         =   0   'False
            SpreadDesigner  =   "AEH2010C.frx":00A8
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   4530
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   15240
         _Version        =   393216
         _ExtentX        =   26882
         _ExtentY        =   7990
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
         MaxCols         =   22
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AEH2010C.frx":06F5
      End
   End
   Begin Threed.SSOption opt_ord1_fl 
      Height          =   330
      Left            =   9360
      TabIndex        =   23
      Top             =   75
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   582
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
      Caption         =   "������"
      Value           =   -1
   End
   Begin Threed.SSOption opt_ord2_fl 
      Height          =   330
      Left            =   10320
      TabIndex        =   24
      Top             =   75
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   582
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
      Caption         =   "���"
   End
   Begin CSTextLibCtl.sidbEdit sdb_cal_wgt 
      Height          =   315
      Left            =   120
      TabIndex        =   27
      Top             =   9240
      Visible         =   0   'False
      Width           =   1185
      _Version        =   262145
      _ExtentX        =   2090
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   16711680
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
      ReadOnly        =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.000"
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
      NumIntDigits    =   12
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_thk1 
      Height          =   315
      Left            =   1350
      TabIndex        =   28
      Top             =   9240
      Visible         =   0   'False
      Width           =   1365
      _Version        =   262145
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_wid1 
      Height          =   315
      Left            =   2730
      TabIndex        =   29
      Top             =   9240
      Visible         =   0   'False
      Width           =   1365
      _Version        =   262145
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin CSTextLibCtl.sidbEdit sdb_asroll_thk_fr 
      Height          =   315
      Left            =   1275
      TabIndex        =   32
      Top             =   870
      Width           =   1365
      _Version        =   262145
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   240
      Top             =   870
      Width           =   990
      _ExtentX        =   1746
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
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   4650
      Top             =   870
      Width           =   990
      _ExtentX        =   1746
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
   Begin CSTextLibCtl.sidbEdit sdb_asroll_thk_to 
      Height          =   315
      Left            =   2640
      TabIndex        =   33
      Top             =   870
      Width           =   1365
      _Version        =   262145
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin CSTextLibCtl.sidbEdit sdb_asroll_wid_fr 
      Height          =   315
      Left            =   5685
      TabIndex        =   34
      Top             =   870
      Width           =   1365
      _Version        =   262145
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin CSTextLibCtl.sidbEdit sdb_asroll_wid_to 
      Height          =   315
      Left            =   7050
      TabIndex        =   35
      Top             =   870
      Width           =   1365
      _Version        =   262145
      _ExtentX        =   2408
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   9330
      Top             =   480
      Width           =   990
      _ExtentX        =   1746
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
End
Attribute VB_Name = "AEH2010C"
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
'-- Program Name      SLAB CUT SEARCH
'-- Program ID        AEH2010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2008.4.8
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

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pControl2 As New Collection     'Master Primary Key Collection
Dim nControl2 As New Collection     'Master Necessary Collection
Dim mControl2 As New Collection     'Master Maxlength check Collection
Dim iControl2 As New Collection     'Master Insert Collection
Dim rControl2 As New Collection     'Master Refer Collection
Dim cControl2 As New Collection     'Master Copy Collection
Dim aControl2 As New Collection     'Master -> Spread Collection
Dim lControl2 As New Collection     'Master Lock Collection

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
Dim Sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim oRd_cnt As Integer
Dim Ord_Len As Double
Dim bSelect As Boolean
Dim lCurrRow As Long
Dim iCount As Integer

Const SS1_SLAB_NO = 1
Const SS1_URGNT_FL = 22          '����������ɫ���  2012-11-07 by CaoLei
Const SS1_ORD_NO = 12
Const SS1_ORD_ITEM = 13

Private Sub Form_Define()

    Dim iCol As Integer
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
              Call Gp_Ms_Collection(TXT_PLT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_PLT_NAME, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(cbo_ord_item, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_prod_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_STLGRD_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_loc, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ord_fl, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_slab_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_slab_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_slab_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_slab_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_slab_len_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_slab_len_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_asroll_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_asroll_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_asroll_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_asroll_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(txt_slab_no, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(cbo_slab_cut, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(sdb_slab_thk1, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(sdb_slab_wid1, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(sdb_slab_len, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(sdb_slab_wgt, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(sdb_cal_wgt, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(sdb_cut_len, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(sdb_cut_wgt, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(udt_ins_date, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
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
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    
    For iCol = 2 To ss1.MaxCols - 3     '2->3
        Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
   
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '����������ɫ���  2012-11-07 by CaoLei
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AEH2010C.P_REFER1", Key:="P-R"
    Sc1.Add Item:="AEH2010C.P_ONEROW1", Key:="P-O"
    Sc1.Add Item:="AEH2010C.P_MODIFY1", Key:="P-M"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
        
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AEH2010C.P_REFER2", Key:="P-R"
    sc2.Add Item:="AEH2010C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    cbo_slab_cut.AddItem "0"
    cbo_slab_cut.AddItem "1"
    cbo_slab_cut.AddItem "2"
    cbo_slab_cut.AddItem "3"
    cbo_slab_cut.AddItem "4"
    cbo_slab_cut.AddItem "5"
    cbo_slab_cut.AddItem "6"
    cbo_slab_cut.AddItem "7"
    cbo_slab_cut.AddItem "8"
    cbo_slab_cut.AddItem "9"
    cbo_slab_cut.AddItem "10"
    
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").Row = 0
    sc2.Item("Spread").Text = "��"
    
    Call Gp_Sp_ColHidden(ss1, 19, True)
    Call Gp_Sp_ColHidden(ss1, 20, True)
    Call Gp_Sp_ColHidden(ss1, 21, True)

    Call Gp_Sp_ColHidden(ss2, 1, True)
    Call Gp_Sp_ColHidden(ss2, 7, True)
    Call Gp_Sp_ColHidden(ss2, 8, True)
    Call Gp_Sp_ColHidden(ss2, 9, True)

    bSelect = False
        
End Sub

Private Sub cbo_slab_cut_Click()

    Dim iRow As Integer
    Dim uThk, uWid, uLen, uWgt, tdLen, lLen, tdWgt As Double
    Dim sTlgrd As String
        
    If cbo_slab_cut.ListIndex = 0 Then
        If opt_ord1_fl.VALUE Then
            cbo_slab_cut.ListIndex = oRd_cnt
        Else
            ss2.MaxRows = 0
            sdb_cut_len.VALUE = 0
            sdb_cut_wgt.VALUE = 0
            Exit Sub
        End If
    End If
    
    If cbo_slab_cut.ListIndex <= oRd_cnt Then
        cbo_slab_cut.ListIndex = oRd_cnt
    End If
    
    If cbo_slab_cut.Tag = "C" Then
        cbo_slab_cut.Tag = ""
        Exit Sub
    End If
    
    If txt_slab_no.Text = "" Then Exit Sub
    
    tdLen = 0
    tdWgt = 0
    
    sdb_cut_len.VALUE = 0
    sdb_cut_wgt.VALUE = 0
    
    'stlgrd
    ss1.Row = lCurrRow
    ss1.Col = 2
    sTlgrd = ss1.Text
    
    ss2.MaxRows = Val(cbo_slab_cut.Text)
    
    uLen = Format((sdb_slab_len.VALUE - Ord_Len) / IIf((Val(cbo_slab_cut.Text) - oRd_cnt) = 0, oRd_cnt, Val(cbo_slab_cut.Text) - oRd_cnt), "####0")
    
    For iRow = 1 To Val(cbo_slab_cut.Text)
    
        ss2.Row = iRow
            
        'Slab_No
        ss2.Col = 1
        ss2.Text = txt_slab_no.Text
        
        'Cut_Seq
        ss2.Col = 2
        ss2.Text = Right("0" & iRow, 2)
    
        'Slab_Thk
        ss2.Col = 3
        ss2.VALUE = sdb_slab_thk1.VALUE
        uThk = sdb_slab_thk1.VALUE
        
        'Slab_Wid
        ss2.Col = 4
        ss2.VALUE = sdb_slab_wid1.VALUE
        uWid = sdb_slab_wid1.VALUE
    
        'Slab_Len
        ss2.Col = 10
        
        If ss2.Text = "2" Or opt_ord2_fl.VALUE Then 'ORD_FL
            ss2.Col = 5
            If ss2.Row = ss2.MaxRows Then
                ss2.VALUE = sdb_slab_len.VALUE - tdLen
                sdb_cut_len.VALUE = sdb_cut_len.VALUE + ss2.VALUE
                ss2.Col = 10:   ss2.Text = "2"
            Else
                ss2.VALUE = uLen
                sdb_cut_len.VALUE = sdb_cut_len.VALUE + uLen
                tdLen = tdLen + uLen
                ss2.Col = 10:   ss2.Text = "2"
            End If
            ss2.Col = 4:   ss2.Lock = False
            ss2.Col = 5:   ss2.Lock = False
        ElseIf ss2.Text = "" Then
            ss2.Col = 5
            ss2.VALUE = uLen
            sdb_cut_len.VALUE = sdb_cut_len.VALUE + uLen
            tdLen = tdLen + uLen
            ss2.Col = 10:   ss2.Text = "2"
            ss2.Col = 4:    ss2.Lock = False
            ss2.Col = 5:    ss2.Lock = False
        Else 'ORD_FL = '1'
            ss2.Col = 5
            sdb_cut_len.VALUE = sdb_cut_len.VALUE + ss2.VALUE
            tdLen = tdLen + ss2.VALUE
            ss2.Lock = True
            ss2.Col = 4:    ss2.Lock = True
        End If
        
        'Slab_Wgt
        ss2.Col = 10
        
        If ss2.Text = "2" Then 'ORD_FL
        
            ss2.Col = 6
            If ss2.Row = ss2.MaxRows Then
                ss2.VALUE = sdb_slab_wgt.VALUE - tdWgt
                sdb_cut_wgt.VALUE = sdb_cut_wgt.VALUE + ss2.VALUE
            Else
                'ss2.Value = Round(sdb_slab_wgt.Value * ((uThk * uWid * uLen) / (uThk * uWid * sdb_slab_len.Value)), 3)
                uWgt = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('WGT','" & sTlgrd & "'," & uThk & "," & uWid & "," & uLen & ",0) FROM DUAL")
                ss2.VALUE = Val(uWgt)
                sdb_cut_wgt.VALUE = sdb_cut_wgt.VALUE + ss2.VALUE
                tdWgt = tdWgt + ss2.VALUE
            End If
        Else
            ss2.Col = 6
            sdb_cut_wgt.VALUE = sdb_cut_wgt.VALUE + ss2.VALUE
            tdWgt = tdWgt + ss2.VALUE
        End If
        
        ss2.Col = 7
        ss2.Text = sUserID
        
        ss2.Col = 0
        ss2.Row = iRow
        ss2.Text = "Input"
    
    Next iRow
    
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

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Sp_Setting(Sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    'Call Gp_Sp_ReadOnlySet(Sc1.Item("Spread"))
    
    Call Gf_Sp_Cls(Sc1)
    Call Gf_Sp_Cls(sc2)
    
    TXT_PLT.Text = "B1"
    Call txt_plt_KeyUp(0, 0)
    opt_ord1_fl.VALUE = True
    
    Call Gp_Spl_SizeGet(SSSplitter1, "E-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "E-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "E-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(Sc1.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "E-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
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
    Set Sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Sc1)
    
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc2) Then
        Call Gf_Sp_Cls(Sc1)
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        TXT_PLT.Text = "B1"
        Call txt_plt_KeyUp(0, 0)
        opt_ord1_fl.VALUE = True
        cbo_slab_cut.Enabled = False
        opt_ord1_fl.Enabled = True
        opt_ord2_fl.Enabled = True
        bSelect = False
    End If
    
End Sub

Public Sub Form_Ref()

    Dim lRow As Integer
    
    If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call Gf_Sp_Cls(sc2)
        cbo_slab_cut.Enabled = False
        opt_ord1_fl.Enabled = False
        opt_ord2_fl.Enabled = False
        bSelect = False
        
'        For lRow = 1 To ss1.MaxRows
'            ss1.Row = lRow
'            ss1.Col = 15
'            If ss1.Text <> "1" Then
'                ss1.Col = 4:    ss1.Lock = False
'                Call Gp_Sp_CellColor(ss1, 4, lRow, , &HC0FFFF)
'            Else
'                ss1.Col = 4:    ss1.Lock = True
'                Call Gp_Sp_CellColor(ss1, 4, lRow, , vbWhite)
'            End If
'        Next lRow
        
    End If
    
    '����������ɫ���  2012-11-07   by   CaoLei
    Call SS1_CHANGE_COLOR

End Sub


Private Sub SS1_CHANGE_COLOR()

    With ss1

        If .MaxRows <= 0 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .Row = iCount

             '����������ɫ��� 2012-11-07  by  CaoLei
            ss1.Row = .Row:       ss1.Col = SS1_URGNT_FL
            If ss1.Text = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, SS1_SLAB_NO, SS1_SLAB_NO, .Row, .Row, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_ORD_NO, SS1_ORD_NO, .Row, .Row, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_ORD_ITEM, SS1_ORD_ITEM, .Row, .Row, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_URGNT_FL, SS1_URGNT_FL, .Row, .Row, &HC000&)
            End If

        Next iCount

    End With

End Sub

Public Sub Form_Pro()

    Dim iRow As Integer
    Dim sDatetime As String
    
    If Val(str(sdb_slab_wgt.VALUE)) < Val(str(sdb_cut_wgt.VALUE)) Then
        Call Gp_MsgBoxDisplay("ĸ�������� < �и�����.....", "I")
        Exit Sub
    End If
    
    If Val(str(sdb_slab_wgt.VALUE)) <> Val(str(sdb_cut_wgt.VALUE)) Then
        If Not Gf_MessConfirm("ĸ�������� <> �и�����.....", "I") Then
            Exit Sub
        End If
    End If
    
    sDatetime = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDDHH24MISS') FROM DUAL")
    
    If udt_ins_date.RawData = "" Then
        Call Gp_MsgBoxDisplay(udt_ins_date.Tag + "��������", "I")
        Exit Sub
    End If

    If Len(udt_ins_date.RawData) <> 8 Then
        Call Gp_MsgBoxDisplay(udt_ins_date.Tag + "���Ȳ���ȷ", "I")
        Exit Sub
    End If
    
    If udt_ins_date.RawData < Mid(sDatetime, 1, 8) Then
        Call Gp_MsgBoxDisplay("ָʾ���� < ��������", "I")
        Exit Sub
    End If
    
'    If opt_ord1_fl Then
'
'        If ss1.MaxRows <= 0 Then Exit Sub
'
'        For irow = 1 To ss1.MaxRows
'            ss1.Row = irow
'            ss1.Col = 0
'            If ss1.Text = "Update" Then
'                ss1.Col = 21
'                ss1.Text = udt_ins_date.RawData
'            End If
'        Next irow
'
'        If Gf_Sp_Process(M_CN1, Sc1, Mc1) Then
'            ss1.OperationMode = OperationModeNormal
'            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'            Call MenuTool_ReSet
'            Call Gp_Ms_Cls(Mc2("rControl"))
'            Call Gf_Sp_Cls(Sc2)
'            cbo_slab_cut.Enabled = False
'        End If
'    Else
    
        If ss1.MaxRows <= 0 Then Exit Sub
        
        For iRow = 1 To ss2.MaxRows
            ss2.Row = iRow
            ss2.Col = 7
            ss2.Text = sUserID
            ss2.Col = 8
            ss2.Text = sDatetime
            ss2.Col = 9
            ss2.Text = udt_ins_date.RawData
        Next iRow
        
        If Gf_Sp_Process(M_CN1, sc2, Mc2) Then
            Call Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl"))
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call MenuTool_ReSet
            Call Gp_Ms_Cls(Mc2("rControl"))
            Call Gf_Sp_Cls(sc2)
            lCurrRow = 0
            cbo_slab_cut.Enabled = False
        End If
'    End If
    
    bSelect = False
    
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
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    
End Sub

Private Sub opt_ord1_fl_Click(VALUE As Integer)

    If opt_ord1_fl Then
        txt_ord_fl.Text = "1"
        opt_ord1_fl.ForeColor = &HFF&
        opt_ord2_fl.ForeColor = &H808080
    Else
        txt_ord_fl.Text = "2"
        opt_ord1_fl.ForeColor = &H808080
        opt_ord2_fl.ForeColor = &HFF&
    End If
    
End Sub

Private Sub opt_ord2_fl_Click(VALUE As Integer)

    If opt_ord2_fl Then
        txt_ord_fl.Text = "2"
        opt_ord2_fl.ForeColor = &HFF&
        opt_ord1_fl.ForeColor = &H808080
    Else
        txt_ord_fl.Text = "1"
        opt_ord2_fl.ForeColor = &H808080
        opt_ord1_fl.ForeColor = &HFF&
    End If

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim iRow As Integer
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If ss1.MaxRows < 1 Or Row < 1 Then Exit Sub
    
    'If Not Gf_Sp_Cls(Sc2) Then Exit Sub
    
    sdb_cut_len.VALUE = 0
    sdb_cut_wgt.VALUE = 0
    
    oRd_cnt = 0
    Ord_Len = 0
    
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    If lCurrRow <> 0 Then
        ss1.Col = 0
        ss1.Row = lCurrRow
        ss1.Text = ""
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lCurrRow, lCurrRow)
    End If

    lCurrRow = Row
    
    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
    ss1.Col = 0
    ss1.Row = Row
    ss1.Text = "ѡ��"
        
    ss1.Row = Row
    ss1.Col = 1:     txt_slab_no.Text = ss1.Text
    ss1.Col = 5:     sdb_slab_thk1.VALUE = ss1.VALUE
    ss1.Col = 6:     sdb_slab_wid1.VALUE = ss1.VALUE
    ss1.Col = 7:     sdb_slab_len.VALUE = ss1.VALUE
    ss1.Col = 8:     sdb_slab_wgt.VALUE = ss1.VALUE
    ss1.Col = 19:    sdb_cal_wgt.VALUE = ss1.VALUE
    cbo_slab_cut.Enabled = True
    
    ss1.Col = 15
    If ss1.Text = "1" Then
        Call Gf_Sp_Refer(M_CN1, sc2, Mc2, , , False)
        ss2.OperationMode = OperationModeNormal
    End If
    
    For iRow = 1 To ss2.MaxRows
        ss2.Row = iRow
        ss2.Col = 0
        ss2.Text = "Input"
        
        ss2.Col = 5
        sdb_cut_len.VALUE = sdb_cut_len.VALUE + IIf(ss2.VALUE = "", 0, ss2.VALUE)
        
        ss2.Col = 6
        sdb_cut_wgt.VALUE = sdb_cut_wgt.VALUE + IIf(ss2.VALUE = "", 0, ss2.VALUE)
        
        ss2.Col = 10
        If ss2.Text = "2" Then   'ORD_FL
            ss2.Col = 4:    ss2.Lock = False
            ss2.Col = 5:    ss2.Lock = False
            Call Gp_Sp_CellColor(ss2, 4, iRow, , &HC0FFFF)
            Call Gp_Sp_CellColor(ss2, 5, iRow, , &HC0FFFF)
        Else
            oRd_cnt = oRd_cnt + 1
            ss2.Col = 4:    ss2.Lock = True
            ss2.Col = 5:    ss2.Lock = True
            Ord_Len = Ord_Len + IIf(ss2.VALUE = "", 0, ss2.VALUE)
            Call Gp_Sp_CellColor(ss2, 4, iRow, , vbWhite)
            Call Gp_Sp_CellColor(ss2, 5, iRow, , vbWhite)
        End If
    
    Next iRow
    
    If ss2.MaxRows <> 0 Then
        cbo_slab_cut.Tag = "C"
        cbo_slab_cut.ListIndex = ss2.MaxRows
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

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    End If

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Col <> 5 And Col <> 4 Then Exit Sub
    
    ss2.Row = Row
    ss2.Col = Col
    
    If Mode = 1 Then
        ss2.Tag = ss2.VALUE
    Else
        If ss2.Tag <> ss2.VALUE Then
            Call Wgt_Cal(Row)
        End If
    End If
    
End Sub

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String

    If Len(Trim(txt_ord_no.Text)) = txt_ord_no.MaxLength Then

        If cbo_ord_item.Text <> "" Then Exit Sub
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(txt_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, cbo_ord_item, sQuery)
        
        'If Combo1.ListCount <> 0 Then
        '      Combo1.ListIndex = 0
        'End If
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
        DD.rControl.Add Item:=TXT_PLT
        DD.rControl.Add Item:=TXT_PLT_NAME

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

    If Len(Trim(TXT_PLT)) = TXT_PLT.MaxLength Then
        TXT_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(TXT_PLT.Text), 2)
    Else
        TXT_PLT_NAME.Text = ""
    End If

End Sub

Private Sub txt_stlgrd_DblClick()

    Call txt_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        
        DD.nameType = "1"
        DD.sWitch = "MS"
        
        DD.rControl.Add Item:=txt_stlgrd
        DD.rControl.Add Item:=txt_STLGRD_NAME
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_stlgrd.Text)) = txt_stlgrd.MaxLength Then
            txt_STLGRD_NAME.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stlgrd.Text))
        Else
            txt_STLGRD_NAME.Text = ""
        End If
        
    End If
    
End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

Private Sub Wgt_Cal(Row As Long)

    Dim iRow As Integer
    Dim uThk, uWid, uLen, uWgt, tdLen, lLen, tdWgt As Double
    Dim sTlgrd As String
        
    If cbo_slab_cut.ListIndex = 0 Then Exit Sub
    If txt_slab_no.Text = "" Then Exit Sub
    
    tdLen = 0
    tdWgt = 0
    
    'stlgrd
    ss1.Row = lCurrRow
    ss1.Col = 2
    sTlgrd = ss1.Text
    
    sdb_cut_len.VALUE = 0
    sdb_cut_wgt.VALUE = 0
    
    ss2.Row = Row
    ss2.Col = 3
    uThk = ss2.VALUE
    ss2.Col = 4
    uWid = ss2.VALUE
    ss2.Col = 5
    uLen = ss2.VALUE
    
    ss2.Col = 6
    'ss2.Value = Round(sdb_slab_wgt.Value * ((uThk * uWid * uLen) / (sdb_slab_thk1.Value * sdb_slab_wid1.Value * sdb_slab_len.Value)), 3)
    
    If Row <> ss2.MaxRows Then
        uWgt = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('WGT','" & sTlgrd & "'," & uThk & "," & uWid & "," & uLen & ",0) FROM DUAL")
        ss2.VALUE = Val(uWgt)
    End If
    
    For iRow = 1 To ss2.MaxRows

        ss2.Row = iRow

        'Slab_Len
        'ss2.Col = 5
        'uLen = ss2.Value
        'sdb_cut_len.Value = sdb_cut_len.Value + uLen

        'Slab_Wgt
        ss2.Col = 6
        'ss2.Value = Round(sdb_slab_wgt.Value * (uLen / sdb_slab_len.Value), 3)
        
        If iRow = ss2.MaxRows Then
            ss2.VALUE = sdb_slab_wgt.VALUE - tdWgt
        Else
            tdWgt = tdWgt + ss2.VALUE
        End If
        
        sdb_cut_wgt.VALUE = sdb_cut_wgt.VALUE + ss2.VALUE
        
    Next iRow

End Sub