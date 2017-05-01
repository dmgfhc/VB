VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACE1280C 
   Caption         =   "余材替代录入_ACE1280C"
   ClientHeight    =   9225
   ClientLeft      =   330
   ClientTop       =   1950
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   14670
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9135
      Left            =   90
      TabIndex        =   7
      Top             =   45
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   16113
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACE1280C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter3 
         Height          =   3705
         Left            =   0
         TabIndex        =   9
         Top             =   5430
         Width           =   15180
         _ExtentX        =   26776
         _ExtentY        =   6535
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   14737632
         PaneTree        =   "ACE1280C.frx":0052
         Begin Threed.SSPanel SSPanel2 
            Height          =   555
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   15180
            _ExtentX        =   26776
            _ExtentY        =   979
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00E0E1FE&
               Caption         =   "模糊替代查询"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   90
               TabIndex        =   38
               Tag             =   "J"
               Top             =   120
               Width           =   1530
            End
            Begin VB.TextBox TXT_STDSPEC 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   10695
               MaxLength       =   18
               TabIndex        =   37
               Tag             =   "标准号"
               Top             =   90
               Width           =   3015
            End
            Begin InDate.ULabel ULabel15 
               Height          =   315
               Left            =   1680
               Top             =   90
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "宽度上/下限"
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
            Begin InDate.ULabel ULabel16 
               Height          =   315
               Left            =   5400
               Top             =   90
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "长度上/下限"
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
            Begin CSTextLibCtl.sidbEdit SDB_WID_FROM 
               Height          =   315
               Left            =   3090
               TabIndex        =   39
               Top             =   90
               Width           =   885
               _Version        =   262145
               _ExtentX        =   1561
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   -2147483640
               BackColor       =   16777215
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
               MaxValue        =   9999.99
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit SDB_WID_TO 
               Height          =   315
               Left            =   4155
               TabIndex        =   40
               Top             =   90
               Width           =   885
               _Version        =   262145
               _ExtentX        =   1561
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   -2147483640
               BackColor       =   16777215
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
               MaxValue        =   9999.99
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit SDB_LEN_FROM 
               Height          =   315
               Left            =   6810
               TabIndex        =   41
               Top             =   90
               Width           =   975
               _Version        =   262145
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   -2147483640
               BackColor       =   16777215
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
               NumIntDigits    =   5
               MaxValue        =   9999999.9
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit SDB_LEN_TO 
               Height          =   315
               Left            =   7950
               TabIndex        =   42
               Top             =   90
               Width           =   975
               _Version        =   262145
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   -2147483640
               BackColor       =   16777215
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
               NumIntDigits    =   5
               MaxValue        =   9999999.9
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel22 
               Height          =   315
               Index           =   1
               Left            =   9300
               Top             =   90
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "标准号"
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
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "~"
               Height          =   180
               Left            =   4020
               TabIndex        =   44
               Top             =   210
               Width           =   90
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "~"
               Height          =   180
               Left            =   7830
               TabIndex        =   43
               Top             =   210
               Width           =   90
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   3120
            Left            =   0
            TabIndex        =   13
            Top             =   585
            Width           =   15180
            _Version        =   393216
            _ExtentX        =   26776
            _ExtentY        =   5503
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
            MaxCols         =   30
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "ACE1280C.frx":00A4
         End
      End
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   5370
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   15180
         _ExtentX        =   26776
         _ExtentY        =   9472
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   14737632
         PaneTree        =   "ACE1280C.frx":0C41
         Begin Threed.SSPanel SSPanel1 
            Height          =   1245
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   15180
            _ExtentX        =   26776
            _ExtentY        =   2196
            _Version        =   196609
            BackColor       =   14737632
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
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
               Left            =   8760
               MaxLength       =   2
               TabIndex        =   45
               Tag             =   "轧钢投入工厂"
               Top             =   465
               Width           =   465
            End
            Begin VB.TextBox txt_prod_cd 
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
               Left            =   1170
               MaxLength       =   2
               TabIndex        =   28
               Tag             =   "产品"
               Top             =   90
               Width           =   465
            End
            Begin VB.TextBox txt_prod_cd_name 
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
               Left            =   1635
               MaxLength       =   40
               TabIndex        =   27
               Tag             =   "产品"
               Top             =   90
               Visible         =   0   'False
               Width           =   210
            End
            Begin VB.TextBox txt_enduse_cd 
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
               Left            =   13185
               MaxLength       =   4
               TabIndex        =   26
               Top             =   90
               Width           =   585
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
               Left            =   5910
               MaxLength       =   18
               TabIndex        =   25
               Tag             =   "钢种"
               Top             =   90
               Width           =   1635
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "替代区分"
               ForeColor       =   &H80000007&
               Height          =   240
               Left            =   9630
               TabIndex        =   22
               Top             =   495
               Visible         =   0   'False
               Width           =   2235
               Begin Threed.SSOption Opt_Cond 
                  Height          =   240
                  Left            =   240
                  TabIndex        =   23
                  Top             =   285
                  Width           =   1200
                  _ExtentX        =   2117
                  _ExtentY        =   423
                  _Version        =   196609
                  Font3D          =   1
                  ForeColor       =   255
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
                  Caption         =   "条件替代"
                  Value           =   -1
               End
               Begin Threed.SSOption Opt_No_Cond 
                  Height          =   240
                  Left            =   1575
                  TabIndex        =   24
                  Top             =   285
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   423
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
                  Caption         =   "强制替代"
               End
            End
            Begin VB.TextBox txt_sale_way_name 
               Height          =   315
               Left            =   3300
               TabIndex        =   21
               Top             =   90
               Width           =   1005
            End
            Begin VB.TextBox txt_sale_way 
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
               Left            =   2970
               MaxLength       =   2
               TabIndex        =   20
               Tag             =   "销售方式"
               Top             =   90
               Width           =   345
            End
            Begin VB.TextBox TXT_CUST_DES 
               Height          =   315
               Left            =   9690
               TabIndex        =   19
               Top             =   90
               Width           =   2220
            End
            Begin VB.TextBox txt_cust_cd 
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
               Left            =   8760
               MaxLength       =   6
               TabIndex        =   18
               Top             =   90
               Width           =   915
            End
            Begin VB.TextBox txt_TRIM_NAME 
               Enabled         =   0   'False
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
               Left            =   9180
               TabIndex        =   17
               Tag             =   "钢种"
               Top             =   810
               Width           =   1080
            End
            Begin VB.TextBox txt_TRIM_FL 
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
               Left            =   8760
               MaxLength       =   1
               TabIndex        =   16
               Tag             =   "钢种"
               Top             =   810
               Width           =   420
            End
            Begin VB.TextBox Text_size_knd 
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
               Left            =   5910
               MaxLength       =   2
               TabIndex        =   15
               Tag             =   "钢种"
               Top             =   810
               Width           =   345
            End
            Begin VB.TextBox Text_size_knd_name 
               Enabled         =   0   'False
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
               Left            =   6255
               TabIndex        =   14
               Tag             =   "钢种"
               Top             =   810
               Width           =   1170
            End
            Begin InDate.ULabel ULabel11 
               Height          =   315
               Left            =   90
               Top             =   450
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   556
               Caption         =   "产品厚度"
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
            Begin InDate.ULabel ULabel7 
               Height          =   315
               Left            =   2490
               Top             =   450
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   556
               Caption         =   "产品宽度"
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
            Begin InDate.ULabel ULabel6 
               Height          =   315
               Left            =   4830
               Top             =   450
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               Caption         =   "产品长度"
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
            Begin InDate.ULabel ULabel4 
               Height          =   315
               Left            =   90
               Top             =   90
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               Caption         =   "产品"
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
            Begin CSTextLibCtl.sidbEdit sdb_prod_thk_to 
               Height          =   315
               Left            =   1155
               TabIndex        =   29
               Tag             =   "产品厚度（MAX）"
               Top             =   450
               Width           =   1095
               _Version        =   262145
               _ExtentX        =   1931
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
               NumIntDigits    =   4
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit sdb_prod_len_to 
               Height          =   315
               Left            =   5910
               TabIndex        =   30
               Tag             =   "产品长度（MIN）"
               Top             =   450
               Width           =   1095
               _Version        =   262145
               _ExtentX        =   1931
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
            Begin CSTextLibCtl.sidbEdit sdb_prod_wid_to 
               Height          =   315
               Left            =   3540
               TabIndex        =   31
               Tag             =   "产品宽度（MAX）"
               Top             =   450
               Width           =   1095
               _Version        =   262145
               _ExtentX        =   1931
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
            Begin InDate.ULabel ULabel9 
               Height          =   315
               Left            =   90
               Top             =   810
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   556
               Caption         =   "板坯宽度"
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
            Begin CSTextLibCtl.sidbEdit sdb_slab_wid_fr 
               Height          =   315
               Left            =   1155
               TabIndex        =   32
               Tag             =   "板坯宽度（MIN）"
               Top             =   810
               Width           =   1095
               _Version        =   262145
               _ExtentX        =   1931
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
               Left            =   2490
               TabIndex        =   33
               Tag             =   "板坯宽度（MAX）"
               Top             =   810
               Width           =   1095
               _Version        =   262145
               _ExtentX        =   1931
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
            Begin CSTextLibCtl.sidbEdit sdb_mat_wgt 
               Height          =   315
               Left            =   13185
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   810
               Width           =   1860
               _Version        =   262145
               _ExtentX        =   3281
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   255
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
            Begin InDate.ULabel ULabel8 
               Height          =   315
               Left            =   12105
               Top             =   90
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               Caption         =   "订单用途"
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
            Begin InDate.ULabel ULabel3 
               Height          =   315
               Left            =   4830
               Top             =   90
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               Caption         =   "标准号"
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
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Left            =   12105
               Top             =   810
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   556
               Caption         =   "材料重量"
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
            Begin InDate.ULabel ULabel10 
               Height          =   315
               Left            =   12105
               Top             =   450
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   556
               Caption         =   "订单重量"
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
            Begin CSTextLibCtl.sidbEdit sdb_ord_wgt 
               Height          =   315
               Left            =   13185
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   450
               Width           =   1860
               _Version        =   262145
               _ExtentX        =   3281
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   16711680
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
               Left            =   1875
               Top             =   90
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   556
               Caption         =   "销售方式"
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
            Begin InDate.ULabel ULabel13 
               Height          =   315
               Left            =   7695
               Top             =   90
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               Caption         =   "客户"
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
            Begin InDate.ULabel ULabel14 
               Height          =   315
               Left            =   4830
               Top             =   810
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               Caption         =   "定尺区分"
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
            Begin InDate.ULabel ULabel23 
               Height          =   315
               Left            =   7695
               Top             =   810
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               Caption         =   "切边"
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
            Begin InDate.ULabel ULabel17 
               Height          =   315
               Left            =   7695
               Top             =   450
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               Caption         =   "订单投入"
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
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "~"
               Height          =   180
               Left            =   2310
               TabIndex        =   36
               Top             =   945
               Width           =   90
            End
         End
         Begin FPSpread.vaSpread ss1 
            Height          =   4095
            Left            =   0
            TabIndex        =   12
            Top             =   1275
            Width           =   15180
            _Version        =   393216
            _ExtentX        =   26776
            _ExtentY        =   7223
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
            MaxCols         =   22
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "ACE1280C.frx":0C93
         End
      End
   End
   Begin VB.TextBox txt_ord_no_1 
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
      Left            =   4470
      MaxLength       =   11
      TabIndex        =   6
      Top             =   8730
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox txt_ord_item 
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
      Left            =   5760
      MaxLength       =   11
      TabIndex        =   5
      Top             =   8715
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.ComboBox cbo_ord_item 
      Height          =   315
      Left            =   12795
      TabIndex        =   1
      Top             =   8580
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.TextBox txt_ord_no 
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
      Left            =   10185
      MaxLength       =   11
      TabIndex        =   0
      Top             =   8565
      Visible         =   0   'False
      Width           =   1290
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_thk_fr 
      Height          =   315
      Left            =   450
      TabIndex        =   2
      Tag             =   "产品厚度（MIN）"
      Top             =   8670
      Visible         =   0   'False
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
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
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_prod_len_fr 
      Height          =   315
      Left            =   3060
      TabIndex        =   4
      Tag             =   "产品长度（MIN）"
      Top             =   8700
      Visible         =   0   'False
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
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
   Begin CSTextLibCtl.sidbEdit sdb_prod_wid_fr 
      Height          =   315
      Left            =   1755
      TabIndex        =   3
      Tag             =   "产品宽度（MIN）"
      Top             =   8700
      Visible         =   0   'False
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   11715
      Top             =   8565
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "订单序列"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   9105
      Top             =   8565
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "订单号"
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
End
Attribute VB_Name = "ACE1280C"
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
'-- Program Name      HMI
'-- Program ID        ACE1280C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.10.4
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

Dim oRd_cnt As Integer              'Select Order Count
Dim iCurr_Row As Integer            'SS1 Current Row

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_prod_cd, "p", "n", "m", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(txt_prod_cd_name, " ", "n", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(txt_sale_way, "p", "n", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(cbo_ord_item, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(txt_enduse_cd, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(sdb_prod_thk_fr, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(sdb_prod_thk_to, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(sdb_prod_wid_fr, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(sdb_prod_wid_to, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(sdb_prod_len_fr, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(sdb_prod_len_to, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(sdb_slab_wid_fr, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(sdb_slab_wid_to, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(sdb_ord_wgt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(sdb_mat_wgt, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(txt_cust_cd, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(TXT_CUST_DES, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(Text_size_knd, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(txt_TRIM_FL, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
             Call Gp_Ms_Collection(txt_plt, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       
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
     Call Gp_Ms_Collection(txt_prod_cd, "p", "n", "m", " ", "r", " ", "l", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
    Call Gp_Ms_Collection(txt_ord_no_1, "p", "n", " ", " ", "r", " ", "l", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
    Call Gp_Ms_Collection(txt_ord_item, "p", "n", " ", " ", "r", " ", "l", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
        
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
      Call Gp_Ms_Collection(txt_prod_cd, "p", "n", "m", " ", "r", " ", "l", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
     Call Gp_Ms_Collection(txt_ord_no_1, "p", "n", " ", " ", "r", " ", "l", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
     Call Gp_Ms_Collection(txt_ord_item, "p", "n", " ", " ", "r", " ", "l", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
     Call Gp_Ms_Collection(SDB_WID_FROM, "p", " ", " ", " ", "r", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
       Call Gp_Ms_Collection(sdb_wid_to, "p", " ", " ", " ", "r", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
     Call Gp_Ms_Collection(SDB_WID_FROM, "p", " ", " ", " ", "r", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
       Call Gp_Ms_Collection(sdb_len_to, "p", " ", " ", " ", "r", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
      Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", "r", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
        
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
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACE1280C.P_REFER1", Key:="P-R"
    sc1.Add Item:="ACE1280C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
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
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ACE1280C.P_REFER3", Key:="P-R"
    sc2.Add Item:="ACE1280C.P_MODIFY", Key:="P-M"
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
    
    Call Gp_Sp_ColColor(ss1, 1)
    Call Gp_Sp_ColColor(ss1, 2)
    
    Call Gp_Sp_ColColor(ss2, 1)
    Call Gp_Sp_ColColor(ss2, 14)
        
    Call Gp_Sp_ColHidden(ss2, 19, True)
    Call Gp_Sp_ColHidden(ss2, 20, True)
    Call Gp_Sp_ColHidden(ss2, 21, True)
    Call Gp_Sp_ColHidden(ss2, 22, True)
    
    'Call Gp_Sp_ColHidden(ss1, 20, True)
    'Call Gp_Sp_ColHidden(ss2, 23, True)
    
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").Row = 0
    sc2.Item("Spread").Text = "◎"
    
    iCurr_Row = 0
    
End Sub

Private Sub chk_Cond_Click()

    If chk_Cond = False Then
        If Gf_Sp_Cls(sc2) Then
            Call Gp_Ms_Cls(Mc3("rControl"))
            txt_prod_cd = "PP"
        End If
    End If
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call Form_Button_Edit

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Button_Edit()

    MDIMain.MenuTool.Buttons(7).Enabled = False              'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False              'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False              'Row Cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False             'Row Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False             'Row Paste
    
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call Form_Button_Edit
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_Cls(Mc3("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
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
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc2) Then
        If Gf_Sp_Cls(sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gp_Ms_Cls(Mc2("rControl"))
            Call Gp_Ms_Cls(Mc3("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call Form_Button_Edit
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            cbo_ord_item.Clear
            txt_sale_way_name.Text = ""
            iCurr_Row = 0
            oRd_cnt = 0
        End If
        Call Gp_Ms_ControlLock(Mc2("lControl"), False)
        txt_prod_cd.SetFocus
    End If
    
End Sub

Public Sub Form_Ref()

    Dim sTemp As String
    Dim iRow As Integer
    
    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
'        Call Gf_Sp_Refer(M_CN1, Sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False)
        ss2.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Form_Button_Edit
        
        sdb_ord_wgt.Value = 0
        sdb_mat_wgt.Value = 0
        iCurr_Row = 0
        oRd_cnt = 0
    End If
            
End Sub

Public Sub Form_Pro()

    Dim sQuery As String
    
    If chk_Cond Then
       Exit Sub
    End If
    
    If Gf_Sp_Process(M_CN1, sc2, Mc2, False) Then
        Call OrderStausProcess
        
        sQuery = Gf_Sp_MakeQuery(sc1.Item("Spread"), sc1.Item("P-O"), "O", sc1.Item("pColumn"), iCurr_Row)
        Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, sc1.Item("Spread"), iCurr_Row)
                    
        'Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iCurr_Row, iCurr_Row)
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        sdb_mat_wgt.Value = 0
        Call Form_Button_Edit
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

Private Sub Opt_Cond_Click(Value As Integer)
    Opt_Cond.ForeColor = &HFF
    Opt_No_Cond.ForeColor = &H80000012
End Sub

Private Sub Opt_No_Cond_Click(Value As Integer)
    Opt_Cond.ForeColor = &H80000012
    Opt_No_Cond.ForeColor = &HFF
End Sub


Private Sub sdb_prod_thk_to_Change()
    sdb_prod_thk_fr.Value = sdb_prod_thk_to.Value
End Sub

Private Sub sdb_prod_wid_to_Change()
    sdb_prod_wid_fr.Value = sdb_prod_wid_to.Value
End Sub

Private Sub sdb_prod_len_to_Change()
    sdb_prod_len_fr.Value = sdb_prod_len_to.Value
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim SMESG       As String
    Dim sCustSpcNo  As String
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If ss1.MaxRows < 1 Or Row < 1 Then Exit Sub
    
    sdb_ord_wgt.Value = 0
    sdb_mat_wgt.Value = 0
    txt_ord_no_1.Text = ""
    txt_ord_item.Text = ""
            
    ss1.BlockMode = True
    ss1.Row = 1
    ss1.Row2 = ss1.MaxRows
    ss1.Col = -1
    ss1.BackColor = &HFFFFFF
    ss1.BlockMode = False
    
    ss1.Row = Row
    ss1.BackColor = &HFFFFC1
    
    ss1.Row = Row
    iCurr_Row = Row
    
    ss1.Col = 1
    txt_ord_no_1.Text = Trim(ss1.Text)
    ss1.Col = 2
    txt_ord_item.Text = Trim(ss1.Text)
    
    'REM_WGT
    ss1.Col = 16
    sdb_ord_wgt.Value = ss1.Value
      
    'CUST SPECIAL REQUEST
    ss1.Col = 20
    sCustSpcNo = Trim(ss1.Text)
    
    sc2.Remove "P-R"
    
    If Opt_Cond Then

        If chk_Cond Then
           sc2.Add Item:="ACE1280C.P_REFER2", Key:="P-R"
        Else
           sc2.Add Item:="ACE1280C.P_REFER3", Key:="P-R"
        End If
    Else
        sc2.Add Item:="ACE1280C.P_REFER2", Key:="P-R"
    End If
    
    If Opt_No_Cond And sCustSpcNo <> "" Then
        ss2.MaxRows = 0
    Else
        If chk_Cond Then
           Call Gf_Sp_Refer(M_CN1, sc2, Mc3, Mc3("nControl"), Mc3("mControl"), False)
        Else
           Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl"), False)
        End If
    End If
    
    ss2.OperationMode = OperationModeNormal
    
End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
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

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If Row < 1 Then Exit Sub
    If ss2.MaxRows < 1 Then Exit Sub
    If chk_Cond Then
       Exit Sub
    End If
    
    ss2.Row = Row
    ss2.Col = 0
    
    If ss2.Text <> "Update" Then
    
'        If oRd_cnt = 0 Then Exit Sub
        
        'REM_WGT
        ss2.Col = 14
        
        If sdb_mat_wgt.Value + ss2.Value > sdb_ord_wgt.Value Then
        
            If Gf_MessConfirm("订单重量 < 材料重量", "Q") Then
            Else
                Exit Sub
            End If
        
        End If
    
        sdb_mat_wgt.Value = sdb_mat_wgt.Value + ss2.Value
        
        ss2.Col = 0
        ss2.Text = "Update"
        
        'Order_No
        ss1.Col = 1
        ss1.Row = iCurr_Row
        ss2.Col = 19
        ss2.Text = ss1.Text
        
        'Order_Item
        ss1.Col = 2
        ss2.Col = 20
        ss2.Text = ss1.Text
        
        'Prod_CD
        ss2.Col = 21
        ss2.Text = txt_prod_cd.Text
        
        'INS_EMP
        ss2.Col = 22
        ss2.Text = sUserID
        
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, Row, Row, , &HFFFF80)
    Else
        ss2.Text = ""
        
        ss2.Col = 19
        ss2.Text = ""
        ss2.Col = 20
        ss2.Text = ""
        ss2.Col = 21
        ss2.Text = ""
        ss2.Col = 22
        ss2.Text = ""
        
        'REM_WGT
        ss2.Col = 14
        sdb_mat_wgt.Value = sdb_mat_wgt.Value - ss2.Value
        
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, Row, Row)
    End If

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
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
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

Private Sub Text_size_knd_DblClick()

    Call Text_size_knd_KeyUp(vbKeyF4, 0)
    
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

Private Sub txt_cust_cd_DblClick()

    Call txt_cust_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_enduse_cd_DblClick()

    Call txt_enduse_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_plt_DblClick()

    Call txt_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

    End If
    
End Sub

Private Sub txt_prod_cd_DblClick()

    Call txt_prod_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_sale_way_DblClick()

    Call txt_sale_way_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stlgrd_DblClick()

    Call txt_stlgrd_KeyUp(vbKeyF4, 0)
    
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

Private Sub txt_enduse_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
                
        If txt_prod_cd.Text = "" Then
            Call Gp_MsgBoxDisplay("产品 Must input necessarily")
            Exit Sub
        End If
        
        DD.sWitch = "MS"
        DD.sKey = Mid(txt_prod_cd, 1, 1)
        DD.rControl.Add Item:=txt_enduse_cd
        
        DD.nameType = "2"
        Call Gf_Usage_DD(M_CN1, KeyCode)
    End If
    
End Sub

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim sQuery As String
    
    If Len(Trim(txt_ord_no.Text)) = txt_ord_no.MaxLength Then
    
        If cbo_ord_item.Text <> "" Then Exit Sub
        
        sQuery = " SELECT DISTINCT(ORD_ITEM) FROM CP_REP_ORD WHERE ORD_NO = '" & Trim(txt_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, cbo_ord_item, sQuery)
       
       'If cbo_ord_item.ListCount <> 0 Then
       '   cbo_ord_item.ListIndex = 0
       'End If
    Else
        cbo_ord_item.Clear
    End If
    
End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case txt_prod_cd.Text
       Case "S", "s", "SL"
           txt_prod_cd.Text = "SL"
           ULabel3.Caption = "钢种"
       Case "P", "p", "PP"
           txt_prod_cd.Text = "PP"
           ULabel3.Caption = "标准号"
       Case "H", "h", "HC"
           txt_prod_cd.Text = "HC"
           ULabel3.Caption = "标准号"
    End Select
    
    If KeyCode = vbKeyF4 Then
        
        DD.sWitch = "MS"
        DD.sKey = "B0005"
        
        DD.rControl.Add Item:=txt_prod_cd
        DD.rControl.Add Item:=txt_prod_cd_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If
        
    If Len(Trim(txt_prod_cd.Text)) = txt_prod_cd.MaxLength Then
        txt_prod_cd_name.Text = Gf_ComnNameFind(M_CN1, "B0005", txt_prod_cd.Text, 2)
    Else
        txt_prod_cd_name.Text = ""
    End If
    
End Sub

Private Sub txt_sale_way_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0010"
        DD.rControl.Add Item:=txt_sale_way
        DD.rControl.Add Item:=txt_sale_way_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_sale_way)) = txt_sale_way.MaxLength Then
        txt_sale_way_name.Text = Gf_ComnNameFind(M_CN1, "B0010", Trim(txt_sale_way.Text), 2)
    Else
        txt_sale_way_name.Text = ""
    End If
 
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        If txt_prod_cd.Text = "SL" Then
            DD.sWitch = "MS"
            DD.rControl.Add Item:=txt_stlgrd
            
            DD.nameType = "1"
            Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        Else
            DD.sWitch = "MS"
            DD.rControl.Add Item:=txt_stlgrd
    
            Call Gf_StdSPEC_DD(M_CN1, KeyCode)
        End If
        
    End If
    
End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"

        DD.rControl.Add Item:=txt_cust_cd
        DD.rControl.Add Item:=TXT_CUST_DES

        DD.nameType = "2"

        Call Gf_Customer_DD(M_CN1, KeyCode)

    End If
    
    If txt_cust_cd.Text <> "" Then
        TXT_CUST_DES.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
    Else
        TXT_CUST_DES.Text = ""
    End If

End Sub

Private Sub txt_stdspec_DblClick()
    
    Call txt_stdspec_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_stdspec_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec

        Call Gf_StdSPEC_DD(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

