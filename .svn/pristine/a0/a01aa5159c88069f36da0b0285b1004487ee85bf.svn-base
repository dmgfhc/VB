VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AEC3000C 
   Caption         =   "坯料申请信息查询/选定工序计划炼钢_AEC3000C"
   ClientHeight    =   9585
   ClientLeft      =   180
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9585
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_mill_plt2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   14640
      MaxLength       =   2
      TabIndex        =   32
      Tag             =   "工厂"
      Top             =   4920
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox txt_mill_plt 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1560
      MaxLength       =   2
      TabIndex        =   31
      Tag             =   "工厂"
      Top             =   120
      Width           =   465
   End
   Begin VB.TextBox txt_mill_plt_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   2025
      MaxLength       =   50
      TabIndex        =   30
      Tag             =   "工厂"
      Top             =   120
      Width           =   2250
   End
   Begin VB.ComboBox cbo_ord_item 
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
      Left            =   12585
      TabIndex        =   21
      Top             =   90
      Width           =   660
   End
   Begin VB.TextBox txt_ord_no 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   11190
      MaxLength       =   11
      TabIndex        =   20
      Tag             =   "产品"
      Top             =   95
      Width           =   1395
   End
   Begin InDate.ULabel ULabel1 
      Height          =   225
      Left            =   11400
      Top             =   570
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   397
      Caption         =   "对象编制重量/对象编制数/炼钢编制数"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
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
      ForeColor       =   16711680
   End
   Begin VB.TextBox txt_plt_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   13740
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   60
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txt_plt 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   13275
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   60
      Visible         =   0   'False
      Width           =   465
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   120
      Tag             =   "申请工厂"
      Top             =   90
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "申请工厂"
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
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8340
      Left            =   60
      TabIndex        =   2
      Top             =   870
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   14711
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AEC3000C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   4365
         Left            =   0
         TabIndex        =   3
         Top             =   3975
         Width           =   15180
         _ExtentX        =   26776
         _ExtentY        =   7699
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   14737632
         PaneTree        =   "AEC3000C.frx":0052
         Begin Threed.SSPanel SSPanel1 
            Height          =   570
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   15180
            _ExtentX        =   26776
            _ExtentY        =   1005
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin InDate.ULabel ULabel3 
               Height          =   315
               Left            =   3930
               Top             =   120
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   556
               Caption         =   "钢种"
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
               ForeColor       =   16711680
            End
            Begin VB.TextBox txt_prc_line 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               Left            =   2055
               MaxLength       =   1
               TabIndex        =   25
               Tag             =   "转炉"
               Top             =   120
               Width           =   390
            End
            Begin VB.TextBox txt_ccm_line 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               Left            =   2445
               MaxLength       =   1
               TabIndex        =   24
               Tag             =   "连浇机号"
               Top             =   120
               Width           =   390
            End
            Begin VB.TextBox txt_stlgrd_name 
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
               Left            =   5730
               TabIndex        =   17
               Top             =   120
               Width           =   1965
            End
            Begin VB.TextBox txt_sms_plt 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   14640
               MaxLength       =   2
               TabIndex        =   14
               Tag             =   "工厂"
               Text            =   "B1"
               Top             =   240
               Visible         =   0   'False
               Width           =   390
            End
            Begin VB.TextBox txt_stlgrd 
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
               Left            =   4380
               MaxLength       =   11
               TabIndex        =   6
               Top             =   125
               Width           =   1335
            End
            Begin Threed.SSCheck chk_sel 
               Height          =   345
               Left            =   120
               TabIndex        =   5
               Top             =   120
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   609
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BackStyle       =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "选择"
            End
            Begin CSTextLibCtl.sidbEdit sdb_slab_thk_fr 
               Height          =   315
               Left            =   8460
               TabIndex        =   7
               Top             =   120
               Width           =   645
               _Version        =   262145
               _ExtentX        =   1138
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
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
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel4 
               Height          =   315
               Left            =   7770
               Top             =   120
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   556
               Caption         =   "厚度"
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
            Begin InDate.ULabel ULabel12 
               Height          =   315
               Left            =   9810
               Top             =   120
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   556
               Caption         =   "宽度"
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
            Begin CSTextLibCtl.sidbEdit sdb_slab_thk_to 
               Height          =   315
               Left            =   9105
               TabIndex        =   8
               Top             =   120
               Width           =   645
               _Version        =   262145
               _ExtentX        =   1138
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
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
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit sdb_slab_wid_fr 
               Height          =   315
               Left            =   10500
               TabIndex        =   9
               Top             =   120
               Width           =   645
               _Version        =   262145
               _ExtentX        =   1138
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
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
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit sdb_slab_wid_to 
               Height          =   315
               Left            =   11145
               TabIndex        =   10
               Top             =   120
               Width           =   645
               _Version        =   262145
               _ExtentX        =   1138
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
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
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel13 
               Height          =   315
               Left            =   11850
               Top             =   120
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   556
               Caption         =   "长度"
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
            Begin CSTextLibCtl.sidbEdit sdb_slab_len_fr 
               Height          =   315
               Left            =   12540
               TabIndex        =   11
               Top             =   120
               Width           =   855
               _Version        =   262145
               _ExtentX        =   1508
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
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
               Left            =   13395
               TabIndex        =   12
               Top             =   120
               Width           =   855
               _Version        =   262145
               _ExtentX        =   1508
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
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
               FocusSelect     =   -1  'True
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
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
               NumDecDigits    =   0
               NumIntDigits    =   7
               Undo            =   0
               Data            =   0
            End
            Begin Threed.SSCommand cmd_refer 
               Height          =   375
               Left            =   14310
               TabIndex        =   15
               Top             =   90
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   661
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "查询"
               BevelWidth      =   3
            End
            Begin InDate.ULabel ULabel8 
               Height          =   315
               Left            =   990
               Top             =   120
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   556
               Caption         =   "转炉/连铸"
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
               ForeColor       =   255
            End
            Begin Threed.SSCheck chk_key 
               Height          =   285
               Left            =   3000
               TabIndex        =   33
               Top             =   120
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "重合"
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   3765
            Left            =   0
            TabIndex        =   13
            Top             =   600
            Width           =   15180
            _Version        =   393216
            _ExtentX        =   26776
            _ExtentY        =   6641
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
            MaxCols         =   23
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AEC3000C.frx":00A4
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   3915
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   15180
         _Version        =   393216
         _ExtentX        =   26776
         _ExtentY        =   6906
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
         MaxCols         =   0
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         SpreadDesigner  =   "AEC3000C.frx":0CCE
      End
   End
   Begin Threed.SSCheck chk_hcr 
      Height          =   285
      Left            =   9780
      TabIndex        =   18
      Top             =   510
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "HCR"
      Value           =   1
   End
   Begin Threed.SSCheck chk_ccr 
      Height          =   285
      Left            =   10500
      TabIndex        =   19
      Top             =   510
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CCR"
      Value           =   1
   End
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   9795
      Top             =   90
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "订单号"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   120
      Top             =   480
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "申请日期"
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
   Begin InDate.UDate udt_req_date_fr 
      Height          =   315
      Left            =   1515
      TabIndex        =   22
      Tag             =   "申请日期"
      Top             =   480
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.74
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MaxLength       =   10
   End
   Begin InDate.UDate udt_req_date_to 
      Height          =   315
      Left            =   2925
      TabIndex        =   23
      Tag             =   "申请日期"
      Top             =   480
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.74
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MaxLength       =   10
   End
   Begin InDate.UDate udt_plan_date_fr 
      Height          =   315
      Left            =   6030
      TabIndex        =   26
      Tag             =   "申请日期"
      Top             =   480
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.74
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MaxLength       =   10
   End
   Begin InDate.UDate udt_plan_date_to 
      Height          =   315
      Left            =   7440
      TabIndex        =   27
      Tag             =   "申请日期"
      Top             =   480
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.74
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      MaxLength       =   10
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   4630
      Top             =   480
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "计划使用"
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   4630
      Top             =   90
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "交货期"
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
   Begin InDate.UDate udt_del_date_fr 
      Height          =   315
      Left            =   6030
      TabIndex        =   28
      Tag             =   "INS_DATE"
      Top             =   90
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
   Begin InDate.UDate udt_del_date_to 
      Height          =   315
      Left            =   7440
      TabIndex        =   29
      Tag             =   "INS_DATE"
      Top             =   90
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
   Begin InDate.ULabel ULabel9 
      Height          =   225
      Left            =   13320
      Top             =   120
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   397
      Caption         =   "重点订单用红色字体显示"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
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
      ForeColor       =   255
   End
End
Attribute VB_Name = "AEC3000C"
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
'-- Program Name      SLAB REQUIRE CONFIRM
'-- Program ID        AEC3000C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2007.10.24
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

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

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
Dim iCount As Integer

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS2_IMP_CONT = 19             '重点订单红色标记  2013-11-14 by CaoLei

Private Sub Form_Define()

    Dim iCol As Integer
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
             Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(cbo_ord_item, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(udt_req_date_fr, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(udt_req_date_to, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udt_plan_date_fr, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udt_plan_date_to, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(chk_hcr, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(chk_ccr, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(udt_del_date_fr, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(udt_del_date_to, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_mill_plt, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
         Call Gp_Ms_Collection(txt_sms_plt, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(cbo_ord_item, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(udt_req_date_fr, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(udt_req_date_to, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(udt_plan_date_fr, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(udt_plan_date_to, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_STLGRD_NAME, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(sdb_slab_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(sdb_slab_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(sdb_slab_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(sdb_slab_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(sdb_slab_len_fr, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(sdb_slab_len_to, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
             Call Gp_Ms_Collection(chk_hcr, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
             Call Gp_Ms_Collection(chk_ccr, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(txt_PRC_LINE, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(txt_ccm_line, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(udt_del_date_fr, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(udt_del_date_to, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(txt_mill_plt, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
             Call Gp_Ms_Collection(chk_key, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            
    'MASTER Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
         
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To SS2.MaxCols - 4
        Call Gp_Sp_Collection(SS2, iCol, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next iCol
    
    Call Gp_Sp_Collection(SS2, SS2.MaxCols - 3, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, SS2.MaxCols - 2, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, SS2.MaxCols - 1, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
        Call Gp_Sp_Collection(SS2, SS2.MaxCols, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc2.Add Item:=SS2, Key:="Spread"
    sc2.Add Item:="AEC3000C.P_REFER2", Key:="P-R"
    sc2.Add Item:="AEC3000C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=SS2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").Row = 0
    sc2.Item("Spread").Text = "◎"
    
    Call Gp_Sp_ColHidden(SS2, SS2.MaxCols - 3, True)
    Call Gp_Sp_ColHidden(SS2, SS2.MaxCols - 2, True)
    Call Gp_Sp_ColHidden(SS2, SS2.MaxCols - 1, True)
    Call Gp_Sp_ColHidden(SS2, SS2.MaxCols, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, SpreadHeader + (ss1.RowHeaderCols - 1), True)
    ss1.Row = SpreadHeader + (ss1.ColHeaderRows - 1)
    ss1.RowHidden = True

End Sub

Public Sub Sp_Setting()

    ss1.ColWidth(SpreadHeader + (ss1.RowHeaderCols - 3)) = 16
    ss1.ColWidth(SpreadHeader + (ss1.RowHeaderCols - 2)) = 5
    ss1.MaxCols = 0

End Sub

Private Sub chk_sel_Click(Value As Integer)

    Dim iRow As Integer
    
    If chk_sel Then
        For iRow = 1 To SS2.MaxRows
            SS2.Row = iRow
            SS2.Col = 0
            SS2.Text = "Input"
            Call Gp_Sp_BlockColor(SS2, 1, SS2.MaxCols, iRow, iRow, , &HFFFF80)
        Next iRow
    Else
        For iRow = 1 To SS2.MaxRows
            SS2.Row = iRow
            SS2.Col = 0
            SS2.Text = ""
            Call Gp_Sp_BlockColor(SS2, 1, SS2.MaxCols, iRow, iRow)
        Next iRow

    End If
    
End Sub

Private Sub cmd_refer_Click()

    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, , , False)
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call MenuTool_ReSet
    SS2.OperationMode = OperationModeNormal
    chk_sel.Value = ssCBUnchecked
    
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
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(Sc1.Item("Spread"))
'    Call Gp_Sp_ReadOnlySet(Sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(Sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Sp_Setting
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet

    txt_plt.Text = "B1"
    Call txt_plt_KeyUp(0, 0)
    
    txt_mill_plt.Text = "C3"
    Call txt_mill_plt_KeyUp(0, 0)
    
    chk_hcr.Value = ssCBChecked
    chk_ccr.Value = ssCBChecked
    
    txt_PRC_LINE.Text = "2"
    txt_ccm_line.Text = "2"
    
'      udt_del_date_to.RawData = Format(Now, "YYYYMMDD")
     udt_del_date_to.RawData = Gf_DTSet(M_CN1, "D")

    udt_del_date_fr.Text = Mid(DateAdd("M", -1, udt_del_date_to.Text), 1, 8) & "20"
    
    ss1.RowHeight(SpreadHeader + (ss1.ColHeaderRows - 2)) = 24
    
    Call Gp_Spl_SizeGet(SSSplitter1, "E-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "E-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub udt_del_date_fr_DblClick()
    
    udt_del_date_fr.RawData = Gf_DTSet(M_CN1, "T", "X")

End Sub

Private Sub udt_del_date_to_DblClick()
    
    udt_del_date_to.RawData = Gf_DTSet(M_CN1, "T", "X")

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Spl_SizeSet(SSSplitter1, "E-System.INI", Me.Name)
    
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

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc2) Then
        Call Gf_Sp_Cls(Sc1)
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        txt_plt.Text = "B1"
        txt_sms_plt.Text = "B1"
        Call txt_plt_KeyUp(0, 0)
        txt_PRC_LINE.Text = "2"
        txt_ccm_line.Text = "2"
        ss1.MaxCols = 0
        chk_hcr.Value = ssCBChecked
        chk_ccr.Value = ssCBChecked
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, SS2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
End Sub

Public Sub Form_Ref()

    Dim sQuery1 As String   'Header Display
    Dim sQuery2 As String   'Data Display
    Dim sQuery3 As String   'STLGRD SUM Display
    Dim sQuery4 As String   'WID, THK SUM Display
    Dim sQuery5 As String   'TOTAL SUM Display
    Dim sMesg As String
    Dim sHcr_Fl As String
    Dim sCcr As String
    Dim sMill_Plt As String
    
    If Not Gf_Sp_Cls(sc2) Then Exit Sub
    
    If chk_hcr.Value And chk_ccr.Value Then
        sHcr_Fl = ""
    ElseIf chk_hcr.Value And Not chk_ccr.Value Then
        sHcr_Fl = "H"
    ElseIf Not chk_hcr.Value And chk_ccr.Value Then
        sHcr_Fl = "C"
    Else
        sHcr_Fl = "X"
    End If
    
    'EP_SLAB_EDT_CHECK
    If Ep_Slab_Edt_Chk = False Then Exit Sub
    
    If udt_req_date_fr.RawData = "" Then
       udt_req_date_fr.RawData = "20080101"
    End If
    If udt_req_date_to.RawData = "" Then
       udt_req_date_to.RawData = "20200101"
    End If
    If udt_plan_date_fr.RawData = "" Then
       udt_plan_date_fr.RawData = "20080101"
    End If
    If udt_plan_date_to.RawData = "" Then
       udt_plan_date_to.RawData = "20200101"
    End If
    
    If udt_del_date_fr.RawData = "" Then
       udt_del_date_fr.RawData = "20080101"
    End If
    If udt_del_date_to.RawData = "" Then
       udt_del_date_to.RawData = "20200101"
    End If
    
    If txt_mill_plt.Text = "**" Then
        sMill_Plt = ""
    Else
        sMill_Plt = txt_mill_plt.Text
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    'Header Display
    sQuery1 = "SELECT  DISTINCT  SLAB_THK "
    sQuery1 = sQuery1 + "  FROM  NISCO.EP_REQ_SLAB A ,NISCO.BP_ORDER_ITEM B "
    
    If txt_ord_no.Text = "" Then
        sQuery1 = sQuery1 + "     WHERE  A.INS_DATE       BETWEEN '" & udt_req_date_fr.RawData & "' AND '" & udt_req_date_to.RawData & "' "
        sQuery1 = sQuery1 + "       AND  A.REQ_LIMI_DATE  BETWEEN '" & udt_plan_date_fr.RawData & "000000' AND '" & udt_plan_date_to.RawData & "999999' "
        sQuery1 = sQuery1 + "       AND  B.DEL_TO_DATE  BETWEEN '" & udt_del_date_fr.RawData & "' AND '" & udt_del_date_to.RawData & "' "
        sQuery1 = sQuery1 + "   AND  A.REC_STS        IN ('1') "
    Else
        sQuery1 = sQuery1 + "     WHERE  A.INS_DATE       BETWEEN '00000000'       AND '99999999' "
        sQuery1 = sQuery1 + "       AND  A.REQ_LIMI_DATE  BETWEEN '00000000000000' AND '99999999999999' "
        sQuery1 = sQuery1 + "   AND  A.REC_STS        IN ('1','2') "
        sQuery1 = sQuery1 + "   AND  A.REQ_PLT        LIKE '" & sMill_Plt & "%' "
    End If
    
    sQuery1 = sQuery1 + "   AND  CNF_FL         =  'F' "
    sQuery1 = sQuery1 + "   AND  REQ_HCR_FL     LIKE  '" & sHcr_Fl & "%'"
    sQuery1 = sQuery1 + " ORDER  BY SLAB_THK ASC "
    
    '炉次编制量标准 B1:150 B3:30
    
    'Data Display
    sQuery2 = " {call AEC3000C.P_DATA ( '" & txt_ord_no.Text & "','" & cbo_ord_item.Text & "','" & _
                                             udt_req_date_fr.RawData & "','" & udt_req_date_to.RawData & "','" & _
                                             udt_plan_date_fr.RawData & "','" & udt_plan_date_to.RawData & "','" & sHcr_Fl & "','" & _
                                             udt_plan_date_fr.RawData & "','" & udt_plan_date_to.RawData & "','" & sMill_Plt & "')} "
     
    'STLGRD, WID SUM Display
    sQuery3 = " {call AEC3000C.P_STLGRD_WID ( '" & txt_ord_no.Text & "','" & cbo_ord_item.Text & "','" & _
                                                   udt_req_date_fr.RawData & "','" & udt_req_date_to.RawData & "','" & _
                                                   udt_plan_date_fr.RawData & "','" & udt_plan_date_to.RawData & "','" & sHcr_Fl & "','" & _
                                             udt_plan_date_fr.RawData & "','" & udt_plan_date_to.RawData & "','" & sMill_Plt & "')} "
    
    'THK SUM Display
    sQuery4 = " {call AEC3000C.P_THK ( '" & txt_ord_no.Text & "','" & cbo_ord_item.Text & "','" & _
                                            udt_req_date_fr.RawData & "','" & udt_req_date_to.RawData & "','" & _
                                            udt_plan_date_fr.RawData & "','" & udt_plan_date_to.RawData & "','" & sHcr_Fl & "','" & _
                                             udt_plan_date_fr.RawData & "','" & udt_plan_date_to.RawData & "','" & sMill_Plt & "')} "
    
    'SUM Display
    sQuery5 = " {call AEC3000C.P_TOTAL ( '" & txt_ord_no.Text & "','" & cbo_ord_item.Text & "','" & _
                                              udt_req_date_fr.RawData & "','" & udt_req_date_to.RawData & "','" & _
                                              udt_plan_date_fr.RawData & "','" & udt_plan_date_to.RawData & "','" & sHcr_Fl & "','" & _
                                             udt_plan_date_fr.RawData & "','" & udt_plan_date_to.RawData & "','" & sMill_Plt & "')} "
 
    sMesg = Gf_Ms_NeceCheck(Mc1("nControl"))
    If sMesg = "OK" Then
    
        sMesg = Gf_Ms_NeceCheck2(Mc1("mControl"))
        If sMesg = "OK" Then

            'Header Display
            If Sp_Header_Refer1(ss1, sQuery1) Then       'Header Display
        
                'Data Display
                If Sp_Data_Refer1(ss1, sQuery2) Then     'SLAB Data Display
                    Call Sp_Data_Refer2(ss1, sQuery3)    'STLGRD, WID SUM Display
                    Call Sp_Data_Refer3(ss1, sQuery4)    'THK SUM Display
                    Call Sp_Data_Refer4(ss1, sQuery5)    'TOTAL SUM Display
                    ss1.OperationMode = OperationModeNormal
                    Call Gp_Sp_ReadOnlySet(Sc1.Item("Spread"))
                End If
            
            End If
            
        Else
            Call Gp_MsgBoxDisplay(Trim(sMesg) + "长度不正确", "I")
        End If
    
    Else
        Call Gp_MsgBoxDisplay(Trim(sMesg) + "必须输入", "I")
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub Form_Pro()
    
    Dim iRow As Integer
    Dim sMesg As String
    
    If txt_PRC_LINE.Text = "" Then
        sMesg = txt_PRC_LINE.Tag + "必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    If txt_ccm_line.Text = "" Then
        sMesg = txt_ccm_line.Tag + "必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    For iRow = 1 To SS2.MaxRows
    
        SS2.Row = iRow
        SS2.Col = 0
        
        If SS2.Text = "Input" Then
            SS2.Col = SS2.MaxCols - 2
            SS2.Text = txt_PRC_LINE.Text
            SS2.Col = SS2.MaxCols - 1
            SS2.Text = txt_ccm_line.Text
            SS2.Col = SS2.MaxCols
            SS2.Text = sUserID
        End If
    
    Next iRow
    
    If Gf_Sp_Process(M_CN1, sc2, Mc2) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        Call Form_Ref
    End If
    
End Sub

Public Sub Spread_Can()

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

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If Row <= 0 Or ss1.MaxRows = Row Then Exit Sub
    If ss1.MaxCols - 1 = Col Or ss1.MaxCols - 2 = Col Or ss1.MaxCols = Col Then Exit Sub
    If Col Mod 3 = 0 Or Col Mod 3 = 2 Then Exit Sub
    
    ss1.Col = Col
    ss1.Row = SpreadHeader + (ss1.ColHeaderRows - 2)
    sdb_slab_thk_fr.Value = Val(ss1.Text)
    sdb_slab_thk_to.Value = Val(ss1.Text)
    
    ss1.Row = Row
    ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 2)
    sdb_slab_wid_fr.Value = Val(ss1.Text)
    sdb_slab_wid_to.Value = Val(ss1.Text)
    
    ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 1)
    txt_stlgrd.Text = ss1.Text
    
    ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 3)
    txt_STLGRD_NAME.Text = ss1.Text
    
    sdb_slab_len_fr.Value = 0
    sdb_slab_len_to.Value = 9999999
    
    If Col Mod 3 = 1 Then   'B1
        txt_sms_plt.Text = "B1"
    End If
    
    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, , , False)
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call MenuTool_ReSet
    SS2.OperationMode = OperationModeNormal
    chk_sel.Value = ssCBUnchecked
    
    '重点订单红色标记  2013-11-14   by   CaoLei
    Call SS2_CHANGE_COLOR
    
End Sub

Private Sub SS2_CHANGE_COLOR()

    With SS2

        If .MaxRows <= 0 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .Row = iCount

             '重点订单红色标记 2013-11-14  by  CaoLei
            SS2.Row = .Row:          SS2.Col = SS2_IMP_CONT
            If SS2.Text = "Y" Then
                 Call Gp_Sp_BlockColor(SS2, 1, SS2.MaxCols, .Row, .Row, &HFF&)
            End If

        Next iCount

    End With

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(sc2.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If Row <= 0 Then Exit Sub
    
    SS2.Col = 0
    SS2.Row = Row
    
    If SS2.Text = "" Then
        SS2.Col = 0:              SS2.Text = "Input"
        SS2.Col = SS2.MaxCols:    SS2.Text = sUserID
        Call Gp_Sp_BlockColor(SS2, 1, SS2.MaxCols, Row, Row, , &HFFFF80)
    Else
        SS2.Col = 0:              SS2.Text = ""
        SS2.Col = SS2.MaxCols:    SS2.Text = ""
        Call Gp_Sp_BlockColor(SS2, 1, SS2.MaxCols, Row, Row)
    End If

End Sub

Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.SS2
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
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

Private Sub txt_mill_plt_DblClick()

    Call txt_mill_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_mill_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_mill_plt
        DD.rControl.Add Item:=txt_mill_plt_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(txt_mill_plt.Text)) = txt_mill_plt.MaxLength Then
        txt_mill_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_mill_plt.Text), 2)
    Else
        txt_mill_plt_name.Text = ""
    End If

End Sub


Private Function Sp_Header_Refer1(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay1_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim iColCnt As Integer
    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New adodb.Recordset
    
    With sPname

        Sp_Header_Refer1 = True
        
        .ReDraw = False
        .MaxRows = 0:  .MaxCols = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Header_Refer1 = False
            '.ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = (UBound(ArrayRecords, 2) + 1) * 3
            For iCol = 0 To .MaxCols - 1 Step 3
            
                For iColCnt = 1 To 3
                    
                    .Row = SpreadHeader + (.ColHeaderRows - 2)
                    .Col = iCol + iColCnt
                    
                    If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(0, iCnt))
                    End If
                    
                    .ColWidth(iCol + iColCnt) = 10
    
                    .Col = iCol + iColCnt: .Col2 = iCol + iColCnt
                    .Row = 1: .ROW2 = -1
                    .BlockMode = True
                    .TypeHAlign = TypeHAlignCenter
                    .TypeVAlign = TypeVAlignCenter
                    .BlockMode = False
                    
                    .Row = SpreadHeader + (.ColHeaderRows - 1)
                    .Col = iCol + iColCnt
                    
                    Select Case iColCnt
                        Case 1
                            .Text = "板卷厂"
                        Case 2
                            .Text = "老炼厂"
                            .ColHidden = True
                        Case 3
                            .Text = "合计"
                            .ColHidden = True
                    End Select
                    
                    If iColCnt = 3 Then
                        Call Gp_Sp_ColHidden(ss1, .Col, True)
                    End If
                    
                Next iColCnt
                
                iCnt = iCnt + 1
                
            Next iCol
            
            '合计 Col
            For iColCnt = 1 To 3
                
                .MaxCols = .MaxCols + 1
                .Col = .MaxCols
                .Row = SpreadHeader + (.ColHeaderRows - 2)
                .Text = "合计(t)"
                .Row = SpreadHeader + (.ColHeaderRows - 1)
                
                Select Case iColCnt
                    Case 1
                        .Text = "板卷厂"
                    Case 2
                        .Text = "老炼厂"
                        .ColHidden = True
                    Case 3
                        .Text = "合计"
                        .ColHidden = True
                End Select
                    
                .ColWidth(.Col) = 12
                    
                .Col = .MaxCols: .Col2 = .MaxCols
                .Row = 1: .ROW2 = -1
                .BlockMode = True
                .TypeHAlign = TypeHAlignCenter
                .TypeVAlign = TypeVAlignCenter
                .BlockMode = False
                
            Next iColCnt
            
        End If
        
        .BlockMode = True
        .Col = .MaxCols:  .Col2 = .MaxCols
        .Row = 1: .ROW2 = -1
        .ForeColor = &HFF&  '&H00FF0000&
        .BlockMode = False
        
        For iColCnt = 3 To .MaxCols - 3 Step 3
            .BlockMode = True
            .Col = iColCnt:  .Col2 = iColCnt
            .Row = 1: .ROW2 = -1
            .ForeColor = &HFF0000
            .BlockMode = False
        Next iColCnt
        
        .BlockMode = True
        .Row = SpreadHeader + (.ColHeaderRows - 2)
        .Col = 1
        .ROW2 = SpreadHeader + (.ColHeaderRows - 2)
        .Col2 = .MaxCols - 3
        .RowMerge = MergeAlways
        '.ColMerge = MergeAlways
        .BlockMode = False
        
        .BlockMode = True
        .Row = SpreadHeader + (.ColHeaderRows - 2)
        .Col = .MaxCols - 2
        .ROW2 = SpreadHeader + (.ColHeaderRows - 1)
        .Col2 = .MaxCols - 2
        .RowMerge = MergeAlways
        ''.ColMerge = MergeAlways
        .BlockMode = False
        
        .ReDraw = True
        .Refresh
        
        Screen.MousePointer = vbDefault
        
    End With
        
    Exit Function

SpreadDisplay1_Error:
    
    Set AdoRs = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer1 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay1_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer1(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay1_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    
    Dim iBas As Integer
    Dim iCot As Integer
    
    Dim sCol_a As String
    Dim sCol_b As String
    Dim sStlgrd As String
    Dim sWid As String
    
    Dim ColSum As Double
    
    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New adodb.Recordset
    
    With sPname

        Sp_Data_Refer1 = True
        .ReDraw = False
        .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer1 = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            For iCnt = 0 To UBound(ArrayRecords, 2)

                If iCnt = 0 Or sStlgrd <> Trim(ArrayRecords(2, iCnt)) Or sWid <> Trim(ArrayRecords(1, iCnt)) Then
                    sStlgrd = Trim(ArrayRecords(2, iCnt))
                    sWid = Trim(ArrayRecords(1, iCnt))
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = SpreadHeader + (.RowHeaderCols - 3)
                    .Text = Trim(ArrayRecords(0, iCnt))
                    .Col = SpreadHeader + (.RowHeaderCols - 2)
                    .Text = Trim(ArrayRecords(1, iCnt))
                    .Col = SpreadHeader + (.RowHeaderCols - 1)
                    .Text = Trim(ArrayRecords(2, iCnt))
                End If
                
                For iCol = 1 To .MaxCols - 1 Step 3
                
                    .Col = iCol
                    .Row = SpreadHeader + (.ColHeaderRows - 2)
                    
                    If .Text = Trim(ArrayRecords(3, iCnt)) Then

                        .Row = .MaxRows
                        
                        If VarType(ArrayRecords(4, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(4, iCnt)) = "0/0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(4, iCnt))
                            End If
                        End If
                        
                        .Col = iCol + 1
                        If VarType(ArrayRecords(5, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(5, iCnt)) = "0/0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(5, iCnt))
                            End If
                        End If
                        
                        .Col = iCol + 2
                        If VarType(ArrayRecords(6, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(6, iCnt))
                        End If
                    End If
                        
                Next iCol
                
            Next iCnt
            
        End If
        
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 0
        .Text = "合计(t)"
        
        Call Gp_Sp_EvenRowBackcolor(sPname, 1)
        
        .BlockMode = True
        .Row = .MaxRows:  .ROW2 = .MaxRows
        .Col = 1: .Col2 = -1
        .ForeColor = &HFF&
        .BlockMode = False
        
        For iCol = 3 To .MaxCols - 3 Step 3
            .BlockMode = True
            .Col = iCol:  .Col2 = iCol
            .Row = .MaxRows: .ROW2 = .MaxRows
            .ForeColor = &HFF0000
            .BlockMode = False
        Next iCol
        
        .ReDraw = True
        Call Gp_Ms_ControlLock(Mc1("lControl"), True)
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay1_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer1 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay1_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer2(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay2_Error

    Dim iRow As Integer
    Dim iCnt As Integer
    
    Dim sStlgrd As String
    Dim sWid As String
    
    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New adodb.Recordset
    
    With sPname

        Sp_Data_Refer2 = True
        .ReDraw = False
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer2 = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            For iCnt = 0 To UBound(ArrayRecords, 2)
                
                For iRow = 1 To .MaxRows
                    
                    .Row = iRow
                    .Col = SpreadHeader + (.RowHeaderCols - 1)
                    sStlgrd = .Text
                    .Col = SpreadHeader + (.RowHeaderCols - 2)
                    sWid = .Text
                    
                    If sStlgrd = Trim(ArrayRecords(1, iCnt)) And sWid = Trim(ArrayRecords(2, iCnt)) Then
    
                        .Col = .MaxCols - 2
                        If VarType(ArrayRecords(3, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(3, iCnt)) = "0/0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(3, iCnt))
                            End If
                        End If
                        
                        .Col = .MaxCols - 1
                        If VarType(ArrayRecords(4, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(4, iCnt)) = "0/0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(4, iCnt))
                            End If
                        End If
                        
                        .Col = .MaxCols
                        If VarType(ArrayRecords(5, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(5, iCnt)) = "0/0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(5, iCnt))
                            End If
                        End If
                        
                        Exit For
                        
                    End If
                    
                Next iRow

            Next iCnt
                
        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay2_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer2 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay2_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer3(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay3_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    
    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New adodb.Recordset
    
    With sPname

        Sp_Data_Refer3 = True
        .ReDraw = False
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer3 = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            For iCnt = 0 To UBound(ArrayRecords, 2)

                For iCol = 1 To .MaxCols - 1 Step 3
                
                    .Col = iCol
                    .Row = SpreadHeader + (.ColHeaderRows - 2)
                    
                    If .Text = Trim(ArrayRecords(0, iCnt)) Then
                        .Row = .MaxRows
                        
                        .Col = iCol
                        If VarType(ArrayRecords(1, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(1, iCnt)) = "0/0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(1, iCnt))
                            End If
                        End If
                        
                        .Col = iCol + 1
                        If VarType(ArrayRecords(2, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(2, iCnt)) = "0/0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(2, iCnt))
                            End If
                        End If
                        
                        .Col = iCol + 2
                        If VarType(ArrayRecords(3, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(3, iCnt)) = "0/0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(3, iCnt))
                            End If
                        End If
                        
                        Exit For
                        
                    End If

                Next iCol
                
            Next iCnt
            
        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay3_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer3 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay3_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer4(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay4_Error

    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New adodb.Recordset
    
    With sPname

        Sp_Data_Refer4 = True
        .ReDraw = False
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer4 = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
                            
            .Row = .MaxRows
            
            .Col = .MaxCols - 2
            If VarType(ArrayRecords(0, 0)) = vbNull Then
                .Text = ""
            Else
                If Trim(ArrayRecords(0, 0)) = "0/0/0" Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, 0))
                End If
            End If
            
            .Col = .MaxCols - 1
            If VarType(ArrayRecords(1, 0)) = vbNull Then
                .Text = ""
            Else
                If Trim(ArrayRecords(1, 0)) = "0/0/0" Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, 0))
                End If
            End If
            
            .Col = .MaxCols
            If VarType(ArrayRecords(2, 0)) = vbNull Then
                .Text = ""
            Else
                If Trim(ArrayRecords(2, 0)) = "0/0/0" Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(2, 0))
                End If
            End If
            
        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay4_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer4 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay3_Error : " & Error)
    
End Function

Private Function Ep_Slab_Edt_Chk() As Boolean

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    
    Dim sQuery As String
    Dim adoCmd As adodb.Command
    
    Screen.MousePointer = vbHourglass
    
    Ep_Slab_Edt_Chk = True
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call AEC3000C.P_SLAB_EDT_CHK (?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
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
        Ep_Slab_Edt_Chk = False
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Function

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Ep_Slab_Edt_Chk : " & Error)
    Ep_Slab_Edt_Chk = False
    
End Function

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(9).Enabled = False                  'Row Cancel
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String

    If Len(Trim(txt_ord_no.Text)) = txt_ord_no.MaxLength Then
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(txt_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, cbo_ord_item, sQuery)
    Else
        cbo_ord_item.Clear
    End If
    
End Sub

Private Sub txt_prc_line_Change()

    txt_ccm_line.Text = txt_PRC_LINE.Text
    
End Sub

Private Sub txt_prc_line_KeyUp(KeyCode As Integer, Shift As Integer)

    txt_ccm_line.Text = txt_PRC_LINE.Text
    
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

