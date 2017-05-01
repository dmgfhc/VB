VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AFM2040C 
   Caption         =   "板坯修磨及废钢实绩修改及查询界面_AFM2040C"
   ClientHeight    =   9225
   ClientLeft      =   570
   ClientTop       =   2145
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame Frame5 
      Height          =   5565
      Left            =   90
      TabIndex        =   65
      Top             =   3600
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   9816
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin SSSplitter.SSSplitter SSSplitter1 
         Height          =   1725
         Left            =   60
         TabIndex        =   67
         Top             =   60
         Width           =   14925
         _ExtentX        =   26326
         _ExtentY        =   3043
         _Version        =   196609
         SplitterBarWidth=   3
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   16761087
         PaneTree        =   "AFM2040C.frx":0000
         Begin Threed.SSPanel SSPanel1 
            Height          =   645
            Left            =   0
            TabIndex        =   68
            Top             =   0
            Width           =   14925
            _ExtentX        =   26326
            _ExtentY        =   1138
            _Version        =   196609
            BackColor       =   14737632
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox TXT_SCRAP_NO_1 
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
               Left            =   10260
               MaxLength       =   10
               ScrollBars      =   1  'Horizontal
               TabIndex        =   98
               Tag             =   "废钢号"
               Top             =   135
               Width           =   1260
            End
            Begin VB.ComboBox CBO_PLT 
               BackColor       =   &H00C0FFFF&
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
               Height          =   315
               ItemData        =   "AFM2040C.frx":0052
               Left            =   14445
               List            =   "AFM2040C.frx":0059
               TabIndex        =   80
               Top             =   90
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.ComboBox CBO_LINE 
               BackColor       =   &H00C0FFFF&
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
               ItemData        =   "AFM2040C.frx":0060
               Left            =   14445
               List            =   "AFM2040C.frx":0067
               TabIndex        =   79
               Text            =   "1"
               Top             =   210
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.TextBox TXT_PRC_NAME 
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
               Height          =   315
               Left            =   5565
               Locked          =   -1  'True
               TabIndex        =   73
               Top             =   135
               Width           =   1365
            End
            Begin VB.TextBox TXT_PRC 
               Height          =   315
               Left            =   5130
               MaxLength       =   2
               TabIndex        =   72
               Top             =   135
               Width           =   435
            End
            Begin VB.TextBox TXT_SCRAP_CD 
               Height          =   315
               Left            =   9210
               TabIndex        =   71
               Top             =   135
               Visible         =   0   'False
               Width           =   135
            End
            Begin VB.ComboBox CBO_SCRAP_CD 
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
               ItemData        =   "AFM2040C.frx":006E
               Left            =   7800
               List            =   "AFM2040C.frx":0078
               TabIndex        =   70
               Top             =   135
               Width           =   1440
            End
            Begin InDate.ULabel ULabel10 
               Height          =   315
               Left            =   4365
               Tag             =   "工序"
               Top             =   135
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               Caption         =   "工序"
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
            Begin InDate.ULabel ULabel13 
               Height          =   315
               Left            =   7020
               Tag             =   "种类"
               Top             =   135
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   556
               Caption         =   "种类"
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
            Begin InDate.ULabel ULabel15 
               Height          =   315
               Left            =   120
               Tag             =   "发生日期"
               Top             =   135
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               Caption         =   "发生日期"
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
            Begin InDate.ULabel ULabel16 
               Height          =   315
               Left            =   11625
               Top             =   135
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               Caption         =   "废钢总量"
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
            Begin CSTextLibCtl.sidbEdit SDB_TOT_WGT 
               Height          =   315
               Left            =   12720
               TabIndex        =   74
               Tag             =   "废钢总量"
               Top             =   135
               Width           =   1470
               _Version        =   262145
               _ExtentX        =   2593
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0"
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
               Enabled         =   0   'False
               BorderEffect    =   2
               DataProperty    =   2
               FocusSelect     =   -1  'True
               Modified        =   -1  'True
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
               NumIntDigits    =   5
               ShowZero        =   0   'False
               MaxValue        =   99999.999
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin InDate.UDate TXT_From_Date 
               Height          =   315
               Left            =   1230
               TabIndex        =   77
               Tag             =   "发生日期"
               Top             =   135
               Width           =   1455
               _ExtentX        =   2566
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
            Begin InDate.UDate TXT_To_Date 
               Height          =   315
               Left            =   2820
               TabIndex        =   78
               Tag             =   "日期"
               Top             =   135
               Width           =   1455
               _ExtentX        =   2566
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
            Begin InDate.ULabel ULabel24 
               Height          =   315
               Left            =   9435
               Top             =   135
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   556
               Caption         =   "废钢号"
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
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "t"
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
               Left            =   14250
               TabIndex        =   76
               Top             =   150
               Width           =   180
            End
            Begin VB.Label Label3 
               BackColor       =   &H00E0E0E0&
               Caption         =   "~"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   150
               Left            =   2685
               TabIndex        =   75
               Top             =   255
               Width           =   195
            End
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   1035
            Left            =   0
            TabIndex        =   69
            Top             =   690
            Width           =   14925
            _ExtentX        =   26326
            _ExtentY        =   1826
            _Version        =   196609
            BackColor       =   14737632
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox SDB_SCRAP_REMARK 
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
               Height          =   315
               Left            =   5400
               Locked          =   -1  'True
               TabIndex        =   102
               Top             =   600
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.TextBox TXT_END_CD 
               Height          =   315
               Left            =   13980
               MaxLength       =   100
               TabIndex        =   100
               Text            =   "Text4"
               Top             =   210
               Visible         =   0   'False
               Width           =   450
            End
            Begin VB.TextBox TXT_END_TIME 
               Height          =   300
               Left            =   13590
               MaxLength       =   100
               TabIndex        =   99
               Text            =   "Text3"
               Top             =   45
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.ComboBox cbo_ths_d_mat_var 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "AFM2040C.frx":0092
               Left            =   12330
               List            =   "AFM2040C.frx":0094
               TabIndex        =   96
               Tag             =   "增减量(+,-)"
               Top             =   555
               Width           =   600
            End
            Begin VB.TextBox txt_Flag 
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
               Height          =   315
               Left            =   14475
               Locked          =   -1  'True
               TabIndex        =   91
               Top             =   90
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.TextBox txt_main_res_cd 
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
               Height          =   315
               Left            =   7665
               Locked          =   -1  'True
               TabIndex        =   90
               Top             =   150
               Width           =   3120
            End
            Begin VB.TextBox txt_code 
               Alignment       =   2  'Center
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
               Left            =   7275
               MaxLength       =   1
               TabIndex        =   89
               Tag             =   "原因"
               Top             =   150
               Width           =   390
            End
            Begin VB.TextBox TXT_SCRAP_NO 
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
               Left            =   12330
               MaxLength       =   10
               ScrollBars      =   1  'Horizontal
               TabIndex        =   88
               Tag             =   "废钢号"
               Top             =   150
               Width           =   1260
            End
            Begin VB.ComboBox cbo_group1 
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
               ItemData        =   "AFM2040C.frx":0096
               Left            =   7275
               List            =   "AFM2040C.frx":0098
               TabIndex        =   87
               Top             =   555
               Width           =   615
            End
            Begin VB.ComboBox cbo_shift1 
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
               ItemData        =   "AFM2040C.frx":009A
               Left            =   4575
               List            =   "AFM2040C.frx":009C
               TabIndex        =   86
               Tag             =   "班次"
               Top             =   555
               Width           =   675
            End
            Begin VB.TextBox TXT_PRC_INPUT_NAME 
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
               Height          =   315
               Left            =   1665
               Locked          =   -1  'True
               TabIndex        =   85
               Top             =   150
               Width           =   1590
            End
            Begin VB.TextBox TXT_PRC_INPUT 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1230
               MaxLength       =   2
               TabIndex        =   84
               Tag             =   "工序"
               Top             =   150
               Width           =   435
            End
            Begin VB.TextBox TXT_SCRAP_INPUT 
               Height          =   315
               Left            =   5790
               TabIndex        =   83
               Top             =   30
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.ComboBox CBO_SCRAP_INPUT 
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
               ItemData        =   "AFM2040C.frx":009E
               Left            =   4575
               List            =   "AFM2040C.frx":00A8
               TabIndex        =   82
               Tag             =   "种类"
               Top             =   150
               Width           =   1440
            End
            Begin VB.TextBox txt_UserId 
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
               Height          =   315
               Left            =   14460
               Locked          =   -1  'True
               TabIndex        =   81
               Top             =   420
               Visible         =   0   'False
               Width           =   255
            End
            Begin InDate.ULabel ULabel4 
               Height          =   315
               Left            =   120
               Tag             =   "发生日"
               Top             =   555
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               Caption         =   "发生日"
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
            Begin InDate.ULabel ULabel8 
               Height          =   315
               Left            =   11235
               Top             =   150
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               Caption         =   "废钢号"
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
            Begin InDate.ULabel ULabel14 
               Height          =   315
               Left            =   8490
               Top             =   555
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               Caption         =   "废钢重量"
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
            Begin CSTextLibCtl.sidbEdit SDB_SCRAP_WGT 
               Height          =   315
               Left            =   9585
               TabIndex        =   92
               Tag             =   "废钢重量"
               Top             =   555
               Width           =   1185
               _Version        =   262145
               _ExtentX        =   2090
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0"
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
               NumIntDigits    =   3
               ShowZero        =   0   'False
               MaxValue        =   9999.999
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel9 
               Height          =   315
               Left            =   6195
               Top             =   150
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               Caption         =   "原因"
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
            Begin InDate.ULabel ULabel6 
               Height          =   315
               Left            =   11235
               Top             =   555
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               Caption         =   "增减量"
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
            Begin InDate.ULabel ULabel5 
               Height          =   315
               Left            =   6195
               Tag             =   "班别"
               Top             =   555
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               Caption         =   "班别"
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
               Left            =   3465
               Top             =   555
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               Caption         =   "班次"
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
            Begin InDate.ULabel ULabel11 
               Height          =   315
               Left            =   120
               Top             =   150
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               Caption         =   "工序"
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
            Begin InDate.ULabel ULabel12 
               Height          =   315
               Left            =   3465
               Top             =   150
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   556
               Caption         =   "种类"
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
            Begin InDate.UDate TXT_OCCR_TIME 
               Height          =   315
               Left            =   1230
               TabIndex        =   95
               Tag             =   "发生日期"
               Top             =   555
               Width           =   1455
               _ExtentX        =   2566
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
            Begin CSTextLibCtl.sidbEdit sdb_ths_d_mat_var 
               Height          =   315
               Left            =   12915
               TabIndex        =   97
               Tag             =   "增减量"
               Top             =   555
               Width           =   1185
               _Version        =   262145
               _ExtentX        =   2090
               _ExtentY        =   556
               _StockProps     =   125
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
               NumIntDigits    =   3
               ShowZero        =   0   'False
               Undo            =   0
               Data            =   0
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "t"
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
               Left            =   14160
               TabIndex        =   94
               Top             =   615
               Width           =   150
            End
            Begin VB.Label Label36 
               BackStyle       =   0  'Transparent
               Caption         =   "t"
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
               Left            =   10830
               TabIndex        =   93
               Top             =   615
               Width           =   150
            End
         End
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   3720
         Left            =   75
         TabIndex        =   66
         Top             =   1785
         Width           =   14895
         _Version        =   393216
         _ExtentX        =   26273
         _ExtentY        =   6562
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   21
         MaxRows         =   2
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AFM2040C.frx":00C2
      End
   End
   Begin Threed.SSFrame Frame1 
      Height          =   2640
      Left            =   90
      TabIndex        =   48
      Top             =   510
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   4657
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.TextBox TXT_CN 
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
         Left            =   8430
         MaxLength       =   10
         ScrollBars      =   1  'Horizontal
         TabIndex        =   101
         Tag             =   "废钢号"
         Top             =   150
         Width           =   900
      End
      Begin Threed.SSCheck CHK_MAIN_GRD2 
         Height          =   240
         Left            =   11670
         TabIndex        =   60
         Top             =   1710
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   423
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
         BackColor       =   14737632
         Caption         =   "订单外一级"
      End
      Begin VB.TextBox txt_grd 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
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
         Left            =   11130
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   1395
         Width           =   390
      End
      Begin VB.TextBox txt_quality_id 
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
         Left            =   11130
         TabIndex        =   56
         Top             =   960
         Width           =   930
      End
      Begin VB.ComboBox cbo_group 
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
         Left            =   13845
         TabIndex        =   54
         Top             =   540
         Width           =   750
      End
      Begin VB.ComboBox cbo_shift 
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
         ItemData        =   "AFM2040C.frx":0C6B
         Left            =   11130
         List            =   "AFM2040C.frx":0C6D
         TabIndex        =   53
         Top             =   540
         Width           =   750
      End
      Begin VB.ComboBox cbo_emp_cd 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "1234567"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
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
         Left            =   1680
         TabIndex        =   52
         Top             =   540
         Width           =   1170
      End
      Begin VB.TextBox txt_det_code 
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
         Left            =   5240
         TabIndex        =   51
         Top             =   540
         Width           =   930
      End
      Begin VB.TextBox txt_det_name 
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
         Left            =   6165
         TabIndex        =   50
         Top             =   540
         Width           =   3180
      End
      Begin VB.TextBox cbo_slab_no 
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
         Left            =   1680
         MaxLength       =   10
         ScrollBars      =   1  'Horizontal
         TabIndex        =   49
         Tag             =   "废钢号"
         Top             =   150
         Width           =   1440
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   1635
         Left            =   60
         TabIndex        =   55
         Top             =   930
         Width           =   9285
         _Version        =   393216
         _ExtentX        =   16378
         _ExtentY        =   2884
         _StockProps     =   64
         ColsFrozen      =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         MaxRows         =   3
         RetainSelBlock  =   0   'False
         ScrollBars      =   0
         SpreadDesigner  =   "AFM2040C.frx":0C6F
         UserResize      =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_sf_yield 
         Height          =   315
         Left            =   13845
         TabIndex        =   58
         Top             =   960
         Width           =   930
         _Version        =   262145
         _ExtentX        =   1640
         _ExtentY        =   556
         _StockProps     =   125
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   9630
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "质量代码"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   12345
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "修磨面积比(%)"
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   9630
         Top             =   1395
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "判定等级"
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
      Begin InDate.ULabel ULabel48 
         Height          =   315
         Left            =   180
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "板坯号"
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
      Begin InDate.ULabel ULabel49 
         Height          =   315
         Left            =   3735
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "作业时间"
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
      Begin InDate.ULabel ULabel53 
         Height          =   315
         Left            =   9630
         Top             =   540
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "班次"
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
      Begin InDate.ULabel ULabel54 
         Height          =   315
         Left            =   12345
         Top             =   540
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "班别"
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
      Begin InDate.ULabel ULabel55 
         Height          =   315
         Left            =   180
         Top             =   540
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "作业人员"
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
      Begin CSTextLibCtl.sitxEdit txt_work_date 
         Height          =   315
         Left            =   5240
         TabIndex        =   59
         Top             =   150
         Width           =   1800
         _Version        =   262145
         _ExtentX        =   3175
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__"
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   3735
         Top             =   540
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Caption         =   "评审原因代码"
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
      Begin Threed.SSCheck CHK_MAIN_GRD1 
         Height          =   240
         Left            =   11670
         TabIndex        =   61
         Top             =   1440
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   423
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
         BackColor       =   14737632
         Caption         =   "正品"
      End
      Begin Threed.SSCheck CHK_MAIN_GRD3 
         Height          =   240
         Left            =   11670
         TabIndex        =   62
         Top             =   1980
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   423
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
         BackColor       =   14737632
         Caption         =   "订单外二级"
      End
      Begin Threed.SSCheck CHK_MAIN_GRD5 
         Height          =   240
         Left            =   12990
         TabIndex        =   63
         Top             =   1710
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   423
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
         BackColor       =   14737632
         Caption         =   "次品"
      End
      Begin Threed.SSCheck CHK_MAIN_GRD7 
         Height          =   240
         Left            =   12990
         TabIndex        =   64
         Top             =   1980
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   423
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
         BackColor       =   14737632
         Caption         =   "废钢"
      End
      Begin InDate.ULabel ULabel25 
         Height          =   315
         Left            =   7140
         Top             =   150
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         Caption         =   "板坯类型"
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
   End
   Begin VB.TextBox txt_rec_sts 
      Height          =   285
      Left            =   3690
      TabIndex        =   47
      Top             =   135
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txt_proc_cd 
      Height          =   285
      Left            =   2565
      TabIndex        =   46
      Top             =   135
      Visible         =   0   'False
      Width           =   885
   End
   Begin Threed.SSCheck Chk_ss3 
      Height          =   330
      Left            =   9945
      TabIndex        =   39
      Top             =   9555
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   2
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
      Caption         =   "2.板坯分板"
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   9945
      TabIndex        =   27
      Top             =   9840
      Visible         =   0   'False
      Width           =   5160
      Begin VB.TextBox Text2 
         BackColor       =   &H80000004&
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
         Left            =   8490
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1815
         Width           =   930
      End
      Begin VB.TextBox Text1 
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
         Left            =   8490
         TabIndex        =   35
         Top             =   1080
         Width           =   930
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "废钢"
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   8550
         TabIndex        =   34
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "次品"
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   8550
         TabIndex        =   33
         Top             =   2295
         Width           =   975
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "订单外二级"
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   7020
         TabIndex        =   32
         Top             =   2640
         Width           =   1500
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "订单外一级"
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   7020
         TabIndex        =   31
         Top             =   2400
         Width           =   1500
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "正品"
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   7020
         TabIndex        =   30
         Top             =   2175
         Width           =   1500
      End
      Begin VB.ComboBox cbo_slab_no1 
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
         Left            =   1155
         TabIndex        =   28
         Top             =   270
         Width           =   1470
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   1755
         Left            =   330
         TabIndex        =   29
         Top             =   1080
         Width           =   4530
         _Version        =   393216
         _ExtentX        =   7990
         _ExtentY        =   3096
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
         MaxCols         =   4
         MaxRows         =   3
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AFM2040C.frx":6E7F
      End
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   345
         Top             =   270
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   556
         Caption         =   "板坯号"
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
      Begin CSTextLibCtl.sidbEdit sidbEdit1 
         Height          =   315
         Left            =   8490
         TabIndex        =   37
         Top             =   1470
         Width           =   915
         _Version        =   262145
         _ExtentX        =   1614
         _ExtentY        =   556
         _StockProps     =   125
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   7005
         Top             =   1080
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         Caption         =   "质量代码"
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
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   7005
         Top             =   1455
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         Caption         =   "修磨面积比(%)"
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
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   7005
         Top             =   1815
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         Caption         =   "判定等级"
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
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   6525
         Top             =   645
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "作业时间"
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
      Begin CSTextLibCtl.sitxEdit sitxEdit1 
         Height          =   315
         Left            =   7590
         TabIndex        =   38
         Top             =   645
         Width           =   1845
         _Version        =   262145
         _ExtentX        =   3263
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__"
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
         Mask            =   "____-__-__ __:__"
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Left            =   330
         Top             =   660
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "板坯分板数"
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
      Begin Threed.SSCommand cmd_divide 
         Height          =   315
         Left            =   1965
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   660
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   196609
         Font3D          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "分板"
      End
      Begin CSTextLibCtl.sidbEdit SDB_DIVIDE_CNT 
         Height          =   315
         Left            =   1515
         TabIndex        =   41
         Top             =   660
         Width           =   420
         _Version        =   262145
         _ExtentX        =   741
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         MaxValue        =   50
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin Threed.SSCommand cmd_divide_ok 
         Height          =   315
         Left            =   2760
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   255
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "确认分板"
      End
      Begin Threed.SSCommand cmd_divide_delete 
         Height          =   315
         Left            =   3870
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   255
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   196609
         Font3D          =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "取消分板"
      End
      Begin CSTextLibCtl.sidbEdit SDB_LEN 
         Height          =   315
         Left            =   3060
         TabIndex        =   44
         Top             =   660
         Visible         =   0   'False
         Width           =   765
         _Version        =   262145
         _ExtentX        =   1349
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         MaxValue        =   9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_WID 
         Height          =   315
         Left            =   3870
         TabIndex        =   45
         Top             =   660
         Visible         =   0   'False
         Width           =   765
         _Version        =   262145
         _ExtentX        =   1349
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         MaxValue        =   9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   23
      Left            =   10470
      TabIndex        =   26
      Text            =   "Text24"
      Top             =   405
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   22
      Left            =   9765
      TabIndex        =   25
      Text            =   "Text23"
      Top             =   405
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   270
      Index           =   21
      Left            =   9060
      TabIndex        =   24
      Text            =   "Text22"
      Top             =   405
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   270
      Index           =   20
      Left            =   8355
      TabIndex        =   23
      Text            =   "Text21"
      Top             =   405
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   270
      Index           =   19
      Left            =   7665
      TabIndex        =   22
      Text            =   "Text20"
      Top             =   405
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   270
      Index           =   18
      Left            =   6990
      TabIndex        =   21
      Text            =   "Text19"
      Top             =   405
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   17
      Left            =   6300
      TabIndex        =   20
      Text            =   "Text18"
      Top             =   405
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   16
      Left            =   5595
      TabIndex        =   19
      Text            =   "Text17"
      Top             =   405
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   15
      Left            =   10575
      TabIndex        =   18
      Text            =   "Text16"
      Top             =   210
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   14
      Left            =   9855
      TabIndex        =   17
      Text            =   "Text15"
      Top             =   210
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   13
      Left            =   9135
      TabIndex        =   16
      Text            =   "Text14"
      Top             =   210
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   12
      Left            =   8430
      TabIndex        =   15
      Text            =   "Text13"
      Top             =   210
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   11
      Left            =   7725
      TabIndex        =   14
      Text            =   "Text12"
      Top             =   210
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   10
      Left            =   7005
      TabIndex        =   13
      Text            =   "Text11"
      Top             =   210
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   9
      Left            =   6300
      TabIndex        =   12
      Text            =   "Text10"
      Top             =   210
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   8
      Left            =   5595
      TabIndex        =   11
      Text            =   "Text9"
      Top             =   210
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   7
      Left            =   10620
      TabIndex        =   10
      Text            =   "Text8"
      Top             =   30
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   6
      Left            =   9885
      TabIndex        =   9
      Text            =   "Text7"
      Top             =   30
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   270
      Index           =   5
      Left            =   9150
      TabIndex        =   8
      Text            =   "Text6"
      Top             =   30
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   270
      Index           =   4
      Left            =   8415
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   30
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   270
      Index           =   3
      Left            =   7695
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   30
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   270
      Index           =   2
      Left            =   6990
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   30
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   1
      Left            =   6285
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   30
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox Text 
      Height          =   270
      Index           =   0
      Left            =   5580
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   30
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txt_oper 
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
      Left            =   1710
      TabIndex        =   0
      Top             =   135
      Visible         =   0   'False
      Width           =   300
   End
   Begin Threed.SSCheck Chk_ss1 
      Height          =   330
      Left            =   165
      TabIndex        =   2
      Top             =   150
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1.修磨"
      Value           =   1
   End
   Begin Threed.SSCheck Chk_ss2 
      Height          =   330
      Left            =   165
      TabIndex        =   1
      Top             =   3240
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "3.板坯库废钢"
   End
End
Attribute VB_Name = "AFM2040C"
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
'-- Program Name      CAST
'-- Program ID        AFH2010C
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2003.7.25
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
Public sDateTime As String              'Active Form Authority Setting

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

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection
Dim Mc3 As New Collection
Dim Mc4 As New Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim sQuery As String


Const SS1_SCR_ORNOT1 = 1            'SCR_ORNOT1
Const SS1_SCR_ORNOT8 = 8            'SCR_ORNOT8
Const SS2_SCRAP_DATE = 1            'SCRAP_DATE
Const SS2_SHIFT = 2                 'SHIFT
Const SS2_GROUP_CD = 3              'GROUP_CD
Const SS2_PRC = 4                   'PRC
Const SS2_PRC_NAME = 5              'PRC_NAME
Const SS2_MAT_KIND = 6              'MAT_KIND
Const SS2_MAT_NO = 7                'MAT_NO
Const SS2_SCRAP_WGT = 8             'SCRAP_WGT
Const SS2_SCRAP_RES = 9             'SCRAP_RES
Const SS2_SCRAP_RES1 = 10           'SCRAP_RES1
Const SS2_END_DATE = 11             'END_DATE
Const SS2_END_CD = 12               'END_CD
Const SS2_LOC = 13                  'LOC
Const SS2_PLT = 14                  'PLT
Const SS3_WID = 3                   'WID
Const SS3_LEN = 4                   'LEN



Private Sub Form_Define()
   'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      'Call Gp_Ms_Collection(txt_oper, "p", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(cbo_slab_no, "p", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(cbo_shift, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(cbo_group, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(cbo_emp_cd, " ", "n", " ", "i", " ", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
  Call Gp_Ms_Collection(txt_work_date, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      
        Call Gp_Ms_Collection(Text(0), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(Text(1), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(Text(2), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(Text(3), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(Text(4), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(Text(5), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(Text(6), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(Text(7), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(Text(8), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(Text(9), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(10), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(11), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(12), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(13), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(14), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(15), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(16), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(17), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(18), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(19), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(20), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(21), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(22), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(Text(23), " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        
 Call Gp_Ms_Collection(txt_quality_id, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(txt_sf_yield, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(txt_grd, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(txt_det_code, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(txt_det_name, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_proc_cd, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_rec_sts, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(TXT_CN, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 
    'MASTER Collection
     Mc1.Add Item:="AFM2040C.P_MODIFY1", Key:="P-M"
     Mc1.Add Item:="AFM2040C.P_REFER1", Key:="P-R"
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(txt_from_DATE, "p", "n", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(txt_to_DATE, "p", "n", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(cbo_shift1, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(txt_PRC, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(CBO_SCRAP_CD, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(TXT_SCRAP_CD, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(SDB_TOT_WGT, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(txt_Flag, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(TXT_SCRAP_NO_1, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    'MASTER Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
 
           Call Gp_Ms_Collection(cbo_plt, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     Call Gp_Ms_Collection(TXT_PRC_INPUT, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
Call Gp_Ms_Collection(TXT_PRC_INPUT_NAME, " ", " ", " ", " ", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
          Call Gp_Ms_Collection(CBO_LINE, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     Call Gp_Ms_Collection(TXT_OCCR_TIME, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
        Call Gp_Ms_Collection(cbo_shift1, "p", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
        Call Gp_Ms_Collection(cbo_group1, "p", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
   Call Gp_Ms_Collection(CBO_SCRAP_INPUT, " ", "n", " ", " ", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
   Call Gp_Ms_Collection(TXT_SCRAP_INPUT, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
      Call Gp_Ms_Collection(TXT_SCRAP_NO, "p", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
          Call Gp_Ms_Collection(txt_code, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
   Call Gp_Ms_Collection(txt_main_res_cd, " ", " ", " ", " ", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     Call Gp_Ms_Collection(SDB_SCRAP_WGT, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
 Call Gp_Ms_Collection(cbo_ths_d_mat_var, " ", " ", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
 Call Gp_Ms_Collection(sdb_ths_d_mat_var, " ", " ", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
        Call Gp_Ms_Collection(txt_UserId, " ", " ", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
      Call Gp_Ms_Collection(TXT_END_TIME, " ", " ", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
        Call Gp_Ms_Collection(TXT_END_CD, " ", " ", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
  Call Gp_Ms_Collection(SDB_SCRAP_REMARK, " ", " ", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
  
    'MASTER Collection
    Mc3.Add Item:="AGF2080C.P_MODIFY", Key:="P-M"
    Mc3.Add Item:=pControl3, Key:="pControl"
    Mc3.Add Item:=nControl3, Key:="nControl"
    Mc3.Add Item:=mControl3, Key:="mControl"
    Mc3.Add Item:=iControl3, Key:="iControl"
    Mc3.Add Item:=rControl3, Key:="rControl"
    Mc3.Add Item:=aControl3, Key:="aControl"
    Mc3.Add Item:=lControl3, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(cbo_slab_no1, "p", "n", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
     
    'MASTER Collection
    Mc4.Add Item:=pControl4, Key:="pControl"
    Mc4.Add Item:=nControl4, Key:="nControl"
    Mc4.Add Item:=mControl4, Key:="mControl"
    Mc4.Add Item:=iControl4, Key:="iControl"
    Mc4.Add Item:=rControl4, Key:="rControl"
    Mc4.Add Item:=aControl4, Key:="aControl"
    Mc4.Add Item:=lControl4, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss2, Key:="Spread"
    sc1.Add Item:="AFM2040C.P_SREFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
     
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc2.Add Item:=ss3, Key:="Spread"
    sc2.Add Item:="AFM2040C.P_SREFER", Key:="P-R"
    sc2.Add Item:="AFM2040C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss3.MaxCols, Key:="Last"
    
    Call Gp_Sp_ColHidden(ss2, SS2_END_DATE, True)     '11
    Call Gp_Sp_ColHidden(ss2, SS2_END_CD, True)       '12
    Call Gp_Sp_ColHidden(ss2, SS2_LOC, True)          '13
    Call Gp_Sp_ColHidden(ss2, SS2_PLT, True)          '14
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

    Call MenuTool_ReSet
End Sub

Private Sub CBO_SCRAP_CD_Change()
    TXT_SCRAP_CD.Text = Mid(Trim(CBO_SCRAP_CD.Text), 1, 2)
End Sub

Private Sub cbo_slab_no_Change()
    
    If Len(cbo_slab_no.Text) = 10 Then
       ''''''modified by guoli at 200709201355''''''
       sQuery = "SELECT * FROM FP_SLAB WHERE SLAB_NO = '" & Trim(cbo_slab_no.Text) & "'"
       '''''''''''''''''''''''''''''''''''''''''''''
       If Gf_FloatFind(M_CN1, sQuery) = 0 Then
            MsgBox "该板坯不存在，板坯号无效！请重新输入板坯号！", vbCritical, "系统提示信息"
       End If
    End If
End Sub

Private Sub cbo_slab_no1_Change()
    
    If Len(cbo_slab_no1.Text) = 10 Then
       sQuery = "SELECT * FROM FP_SLAB WHERE SLAB_NO = '" + cbo_slab_no1.Text + "' AND ORD_FL = '2' AND REC_STS = '2'"
       If Gf_FloatFind(M_CN1, sQuery) = 0 Then
            MsgBox "该板坯不存在，板坯号无效！请重新输入板坯号！", vbCritical, "系统提示信息"
            cbo_slab_no1.ListIndex = -1
       End If
    End If
End Sub

Private Sub CHK_MAIN_GRD1_Click(VALUE As Integer)
    If CHK_MAIN_GRD1.VALUE = ssCBUnchecked Then
        If CHK_MAIN_GRD2.VALUE = ssCBUnchecked And CHK_MAIN_GRD3 = ssCBUnchecked And CHK_MAIN_GRD5 = ssCBUnchecked And CHK_MAIN_GRD7 = ssCBUnchecked Then
          ' CHK_MAIN_GRD1.Value = ssCBChecked
           txt_grd.Text = ""
           CHK_MAIN_GRD1.ForeColor = &H808080
        End If
        Exit Sub
    End If
    
    txt_grd.Text = "1"
    
    CHK_MAIN_GRD1.ForeColor = &HFF&
    CHK_MAIN_GRD1.VALUE = ssCBChecked
    
    CHK_MAIN_GRD2.ForeColor = &H808080
    CHK_MAIN_GRD2.VALUE = ssCBUnchecked
    CHK_MAIN_GRD3.ForeColor = &H808080
    CHK_MAIN_GRD3.VALUE = ssCBUnchecked
    CHK_MAIN_GRD5.ForeColor = &H808080
    CHK_MAIN_GRD5.VALUE = ssCBUnchecked
    CHK_MAIN_GRD7.ForeColor = &H808080
    CHK_MAIN_GRD7.VALUE = ssCBUnchecked
End Sub

Private Sub CHK_MAIN_GRD2_Click(VALUE As Integer)

    If CHK_MAIN_GRD2.VALUE = ssCBUnchecked Then
        If CHK_MAIN_GRD1.VALUE = ssCBUnchecked And CHK_MAIN_GRD3 = ssCBUnchecked And CHK_MAIN_GRD5 = ssCBUnchecked And CHK_MAIN_GRD7 = ssCBUnchecked Then
          ' CHK_MAIN_GRD2.Value = ssCBChecked
           txt_grd.Text = ""
           CHK_MAIN_GRD2.ForeColor = &H808080
        End If
        Exit Sub
    End If
    
    txt_grd.Text = "2"
    
    CHK_MAIN_GRD2.ForeColor = &HFF&
    CHK_MAIN_GRD2.VALUE = ssCBChecked
    
    CHK_MAIN_GRD1.ForeColor = &H808080
    CHK_MAIN_GRD1.VALUE = ssCBUnchecked
    CHK_MAIN_GRD3.ForeColor = &H808080
    CHK_MAIN_GRD3.VALUE = ssCBUnchecked
    CHK_MAIN_GRD5.ForeColor = &H808080
    CHK_MAIN_GRD5.VALUE = ssCBUnchecked
    CHK_MAIN_GRD7.ForeColor = &H808080
    CHK_MAIN_GRD7.VALUE = ssCBUnchecked
   
End Sub

Private Sub CHK_MAIN_GRD3_Click(VALUE As Integer)

    If CHK_MAIN_GRD3.VALUE = ssCBUnchecked Then
        If CHK_MAIN_GRD1.VALUE = ssCBUnchecked And CHK_MAIN_GRD2 = ssCBUnchecked And CHK_MAIN_GRD5 = ssCBUnchecked And CHK_MAIN_GRD7 = ssCBUnchecked Then
          ' CHK_MAIN_GRD3.Value = ssCBChecked
           txt_grd.Text = ""
           CHK_MAIN_GRD3.ForeColor = &H808080
        End If
        Exit Sub
    End If
    
    txt_grd.Text = "3"
    
    CHK_MAIN_GRD3.ForeColor = &HFF&
    CHK_MAIN_GRD3.VALUE = ssCBChecked
    
    CHK_MAIN_GRD2.ForeColor = &H808080
    CHK_MAIN_GRD2.VALUE = ssCBUnchecked
    CHK_MAIN_GRD1.ForeColor = &H808080
    CHK_MAIN_GRD1.VALUE = ssCBUnchecked
    CHK_MAIN_GRD5.ForeColor = &H808080
    CHK_MAIN_GRD5.VALUE = ssCBUnchecked
    CHK_MAIN_GRD7.ForeColor = &H808080
    CHK_MAIN_GRD7.VALUE = ssCBUnchecked
   
End Sub

Private Sub CHK_MAIN_GRD5_Click(VALUE As Integer)

    If CHK_MAIN_GRD5.VALUE = ssCBUnchecked Then
        If CHK_MAIN_GRD1.VALUE = ssCBUnchecked And CHK_MAIN_GRD2 = ssCBUnchecked And CHK_MAIN_GRD3 = ssCBUnchecked And CHK_MAIN_GRD7 = ssCBUnchecked Then
          ' CHK_MAIN_GRD5.Value = ssCBChecked
           txt_grd.Text = ""
           CHK_MAIN_GRD3.ForeColor = &H808080
        End If
        Exit Sub
    End If
    
    txt_grd.Text = "5"
    
    CHK_MAIN_GRD5.ForeColor = &HFF&
    CHK_MAIN_GRD5.VALUE = ssCBChecked
    
    CHK_MAIN_GRD2.ForeColor = &H808080
    CHK_MAIN_GRD2.VALUE = ssCBUnchecked
    CHK_MAIN_GRD1.ForeColor = &H808080
    CHK_MAIN_GRD1.VALUE = ssCBUnchecked
    CHK_MAIN_GRD3.ForeColor = &H808080
    CHK_MAIN_GRD3.VALUE = ssCBUnchecked
    CHK_MAIN_GRD7.ForeColor = &H808080
    CHK_MAIN_GRD7.VALUE = ssCBUnchecked
    
End Sub

Private Sub CHK_MAIN_GRD7_Click(VALUE As Integer)

    If CHK_MAIN_GRD7.VALUE = ssCBUnchecked Then
        If CHK_MAIN_GRD1.VALUE = ssCBUnchecked And CHK_MAIN_GRD2 = ssCBUnchecked And CHK_MAIN_GRD3 = ssCBUnchecked And CHK_MAIN_GRD5 = ssCBUnchecked Then
          ' CHK_MAIN_GRD7.Value = ssCBChecked
           txt_grd.Text = ""
           CHK_MAIN_GRD7.ForeColor = &H808080
        End If
        Exit Sub
    End If
    
    txt_grd.Text = "7"
    
    CHK_MAIN_GRD7.ForeColor = &HFF&
    CHK_MAIN_GRD7.VALUE = ssCBChecked
    
    CHK_MAIN_GRD2.ForeColor = &H808080
    CHK_MAIN_GRD2.VALUE = ssCBUnchecked
    CHK_MAIN_GRD1.ForeColor = &H808080
    CHK_MAIN_GRD1.VALUE = ssCBUnchecked
    CHK_MAIN_GRD3.ForeColor = &H808080
    CHK_MAIN_GRD3.VALUE = ssCBUnchecked
    CHK_MAIN_GRD5.ForeColor = &H808080
    CHK_MAIN_GRD5.VALUE = ssCBUnchecked
   
End Sub

Private Sub CBO_SCRAP_CD_Click()
    TXT_SCRAP_CD.Text = Mid(Trim(CBO_SCRAP_CD.Text), 1, 2)
End Sub

Private Sub CBO_SCRAP_INPUT_Click()
    TXT_SCRAP_INPUT.Text = Mid(Trim(CBO_SCRAP_INPUT.Text), 1, 2)

    cbo_ths_d_mat_var.Text = ""
    sdb_ths_d_mat_var.Text = ""
End Sub

Private Sub Form_Activate()

   Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
   
End Sub

Private Sub Form_Load()
Dim sQuery As String
    
    Dim i, j As Integer
    
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_Cls(Mc3("rControl"))
    Call Gp_Ms_Cls(Mc4("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_ControlLock(Mc2("lControl"), True)
    Call Gp_Ms_ControlLock(Mc3("lControl"), True)
    Call Gp_Ms_ControlLock(Mc4("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    Call Gp_Ms_NeceColor(Mc3("nControl"))
    Call Gp_Ms_NeceColor(Mc4("nControl"))
    
    With ss1
        .RowHeight(-1) = 12.54
        .RowHeight(0) = 24
        
        .ColWidth(0) = 6
        
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        
        .SelBackColor = &H808040
     
        '.OperationMode = OperationModeRow
        .UserResize = UserResizeNone
        .ProcessTab = True
        .ScrollBarExtMode = True
        .ButtonDrawMode = 1
        .TabStop = False
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False
        
        .Col = -1
        .Row = 0
        .FontBold = True
                
        .MaxRows = 3
        .Col = 0
        .Row = 0
        .Text = "◎"
        .Row = 1
        .Text = "宽面"
        .FontBold = True
        .Row = 2
        .Text = "窄面"
        .FontBold = True
        .Row = 3
        .Text = "角面"
        .FontBold = True
        
        .RowHeight(1) = 16
        .RowHeight(2) = 16
        .RowHeight(3) = 16
        
        For i = 1 To 3
            For j = 1 To 8
                .Row = i
                .Col = j
                 If i = 1 Then
                    Text(j - 1).Text = .VALUE
                 ElseIf i = 2 Then
                    Text(j + 7).Text = .VALUE
                 ElseIf i = 3 Then
                    Text(j + 15).Text = .VALUE
                 End If
                 .ColWidth(j) = 8.81
'                 .ColWidth(j) = 8.9
            Next j
        Next i
        
    End With
    
    txt_oper.Text = "1"
    cbo_ths_d_mat_var.AddItem "+"
    cbo_ths_d_mat_var.AddItem "-"
    
    cbo_shift.AddItem "1"
    cbo_shift.AddItem "2"
    cbo_shift.AddItem "3"
    
    cbo_group.AddItem "A"
    cbo_group.AddItem "B"
    cbo_group.AddItem "C"
    cbo_group.AddItem "D"
    
    cbo_shift1.AddItem "1"
    cbo_shift1.AddItem "2"
    cbo_shift1.AddItem "3"
    
    cbo_group1.AddItem "A"
    cbo_group1.AddItem "B"
    cbo_group1.AddItem "C"
    cbo_group1.AddItem "D"
    
    cbo_emp_cd.Text = sUserID
    txt_UserId.Text = sUserID
    cbo_plt.Text = "B1"
    Frame5.Enabled = False
    
    txt_from_DATE.RawData = ""
    txt_to_DATE.RawData = ""
    TXT_OCCR_TIME.RawData = ""
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "F-System.INI", Me.Name)
    'Call Gp_Sp_ColGet(sc2.Item("Spread"), "F-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Chk_ss1_Click(VALUE As Integer)

    If Chk_ss1.VALUE = ssCBUnchecked Then
      If Chk_ss2.VALUE = ssCBUnchecked And Chk_ss3.VALUE = ssCBUnchecked Then
         Chk_ss1.VALUE = ssCBChecked
      End If
      Exit Sub
    End If
    
    If Chk_ss1.VALUE = -1 Then
            Chk_ss1.ForeColor = &HFF&
            Chk_ss2.ForeColor = &H0&
            Chk_ss2.VALUE = ssCBUnchecked
            Chk_ss3.ForeColor = &H0&
            Chk_ss3.VALUE = ssCBUnchecked

            Frame1.Enabled = True
            Frame5.Enabled = False
            Frame1.ShadowStyle = ssRaisedShadow
            Frame5.ShadowStyle = ssInsetShadow
            
            txt_oper = "1"
            cbo_slab_no.Enabled = True
            cbo_slab_no.SetFocus
            MDIMain.MenuTool.Buttons(4).Enabled = True
    End If

 
End Sub

Private Sub Chk_ss2_Click(VALUE As Integer)

    If Chk_ss2.VALUE = ssCBUnchecked Then
        If Chk_ss1.VALUE = ssCBUnchecked And Chk_ss3.VALUE = ssCBUnchecked Then
            Chk_ss2.VALUE = ssCBChecked
        End If
        Exit Sub
    End If
    
    If Chk_ss2.VALUE = -1 Then
            Chk_ss1.ForeColor = &H0&
            Chk_ss1.VALUE = ssCBUnchecked
            Chk_ss3.ForeColor = &H0&
            Chk_ss3.VALUE = ssCBUnchecked
            Chk_ss2.ForeColor = &HFF&

            Frame1.Enabled = False
            Frame5.Enabled = True
            Frame1.ShadowStyle = ssInsetShadow
            Frame5.ShadowStyle = ssRaisedShadow
            
            txt_from_DATE.RawData = Format(Date, "yyyymm" + "01")
            txt_to_DATE.RawData = Format(Date, "yyyymmdd")
            
            txt_oper = "2"
            CBO_SCRAP_CD.Enabled = True
            txt_from_DATE.SetFocus
            MDIMain.MenuTool.Buttons(4).Enabled = True
    End If
    
    If Mid(sAuthority, 4, 1) = "1" Then
       MDIMain.MenuTool.Buttons(5).Enabled = True
    End If

End Sub

Private Sub Chk_ss3_Click(VALUE As Integer)

    If Chk_ss3.VALUE = ssCBUnchecked Then
        If Chk_ss1.VALUE = ssCBUnchecked And Chk_ss2.VALUE = ssCBUnchecked Then
            Chk_ss3.VALUE = ssCBChecked
        End If
        Exit Sub
    End If
    
    If Chk_ss3.VALUE = -1 Then
            Chk_ss1.ForeColor = &H0&
            Chk_ss1.VALUE = ssCBUnchecked
            Chk_ss2.ForeColor = &H0&
            Chk_ss2.VALUE = ssCBUnchecked
            Chk_ss3.ForeColor = &HFF&

            Frame1.Enabled = False
            Frame2.Enabled = True
            Frame5.Enabled = False
                  
            cbo_slab_no1.Clear
            sQuery = "SELECT * FROM (select slab_no from FP_SLAB WHERE ORD_FL = '2' AND REC_STS = '2' order by SLAB_NO desc) WHERE ROWNUM <= 10"
            Call Gf_ComboAdd(M_CN1, cbo_slab_no1, sQuery)
   
            txt_oper = "3"
            MDIMain.MenuTool.Buttons(4).Enabled = False
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "F-System.INI", Me.Name)
    'Call Gp_Sp_ColSet(sc2.Item("Spread"), "F-System.INI", Me.Name)
    
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
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set Mc4 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Form_Cls()
    Dim iRow, iCol As Integer
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
    
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call Gp_Ms_Cls(Mc3("rControl"))
        Call Gp_Ms_Cls(Mc4("rControl"))
        
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        Call Gp_Ms_ControlLock(Mc2("pControl"), False)
        Call Gp_Ms_ControlLock(Mc3("pControl"), False)
        Call Gp_Ms_ControlLock(Mc4("pControl"), False)
        
        With ss1
            For iRow = 1 To 3
                .Row = iRow
                For iCol = SS1_SCR_ORNOT1 To SS1_SCR_ORNOT8   '1,8
                    .Col = iCol
                    .VALUE = 0
                 
                Next iCol
            Next iRow
        End With
        
        Chk_ss1.VALUE = ssCBChecked
        Chk_ss1.ForeColor = &HFF&
        
        CHK_MAIN_GRD1.VALUE = 0
        CHK_MAIN_GRD1.ForeColor = &H808080
        CHK_MAIN_GRD2.VALUE = 0
        CHK_MAIN_GRD2.ForeColor = &H808080
        CHK_MAIN_GRD3.VALUE = 0
        CHK_MAIN_GRD3.ForeColor = &H808080
        CHK_MAIN_GRD5.VALUE = 0
        CHK_MAIN_GRD5.ForeColor = &H808080
        CHK_MAIN_GRD7.VALUE = 0
        CHK_MAIN_GRD7.ForeColor = &H808080
        
        pControl1(1).SetFocus
        Frame1.Enabled = True
        Frame2.Enabled = False
        Frame5.Enabled = False
        
        cbo_emp_cd.Text = sUserID
        
        txt_main_res_cd.Text = ""
        cbo_ths_d_mat_var.Text = ""
        sdb_ths_d_mat_var.Text = ""
        CBO_SCRAP_INPUT.Text = ""
        SDB_DIVIDE_CNT.Text = ""
        txt_from_DATE.RawData = ""
        txt_to_DATE.RawData = ""
        TXT_OCCR_TIME.RawData = ""
    
    End If
 
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
End Sub

Public Sub Form_Ref()

    Dim i, j  As Integer
    Dim iRow  As Integer
    Dim sMesg As String
    
    If Chk_ss1.VALUE = -1 Then
        If cbo_slab_no.Text = "" Then
           MsgBox "板坯号必须输入！", vbCritical, "系统提示信息"
           Exit Sub
        End If
       
        If Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl")) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           Call Gp_Ms_ControlLock(Mc1("pControl"), True)
           Call MenuTool_ReSet
           
           With ss1
              For i = 1 To 3
                  For j = SS1_SCR_ORNOT1 To SS1_SCR_ORNOT8         '1,8
                     .Row = i
                     .Col = j
                      If i = 1 Then
                        .Text = Text(j - 1)
                      ElseIf i = 2 Then
                        .Text = Text(j + 7)
                      ElseIf i = 3 Then
                        .Text = Text(j + 15)
                      End If
                  Next
              Next
           End With
           
           If txt_grd.Text = "1" Then
              CHK_MAIN_GRD1.ForeColor = &HFF&
              CHK_MAIN_GRD1.VALUE = ssCBChecked
           ElseIf txt_grd.Text = "2" Then
              CHK_MAIN_GRD2.ForeColor = &HFF&
              CHK_MAIN_GRD2.VALUE = ssCBChecked
           ElseIf txt_grd.Text = "3" Then
              CHK_MAIN_GRD3.ForeColor = &HFF&
              CHK_MAIN_GRD3.VALUE = ssCBChecked
           ElseIf txt_grd.Text = "5" Then
              CHK_MAIN_GRD5.ForeColor = &HFF&
              CHK_MAIN_GRD5.VALUE = ssCBChecked
           ElseIf txt_grd.Text = "7" Then
              CHK_MAIN_GRD7.ForeColor = &HFF&
              CHK_MAIN_GRD7.VALUE = ssCBChecked
           End If
        End If

    End If

    If Chk_ss2.VALUE = -1 Then
    
        If Trim(txt_from_DATE.RawData) = "" Or Trim(txt_to_DATE.RawData) = "" Then
           MsgBox "发生时间必须输入！", vbCritical, "系统提示信息"
           Exit Sub
        End If
        
        If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
        SDB_TOT_WGT.VALUE = 0
        
        If Mid(txt_PRC.Text, 1, 1) = "B" Then
            txt_Flag.Text = "B1"
        ElseIf Mid(txt_PRC.Text, 1, 1) = "C" Then
            txt_Flag.Text = "C1"
        End If
                
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc2, Mc2("nControl"), Mc2("mControl")) Then
            If ss2.MaxRows > 0 Then
               For iRow = 1 To ss2.MaxRows
                   ss2.Row = iRow
                   ss2.Col = SS2_SCRAP_WGT       '8
                   SDB_TOT_WGT.VALUE = SDB_TOT_WGT.VALUE + Val(ss2.VALUE)
               Next iRow
            End If
            Call Gp_Sp_EvenRowBackcolor(ss2)
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            If Mid(sAuthority, 4, 1) = "1" Then
               MDIMain.MenuTool.Buttons(5).Enabled = True
            End If
            Call MenuTool_ReSet
        End If
        
        cbo_ths_d_mat_var.Text = ""
        sdb_ths_d_mat_var.Text = ""
    End If
    
    If Chk_ss3.VALUE = -1 Then
        If cbo_slab_no1.Text = "" Then
           MsgBox "板坯号必须输入！", vbCritical, "系统提示信息"
           Exit Sub
        End If
                      
        If Gf_Sp_Refer(M_CN1, sc2, Mc4, Mc4("nControl")) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call MenuTool_ReSet
        End If
        
        If ss3.MaxRows > 0 Then
            cmd_divide_ok.Enabled = False
            cmd_divide_delete.Enabled = True
        Else
            cmd_divide_ok.Enabled = True
            cmd_divide_delete.Enabled = False
        End If

    End If

End Sub

Public Sub Form_Pro()
    
    If txt_oper.Text = "1" Then
        If Len(cbo_slab_no.Text) <> 10 Then
            MsgBox "板坯号不正确！", vbCritical, "系统提示信息"
            Exit Sub
        End If
    
        If Trim(txt_work_date.Text) = "" Then
            MsgBox "作业日期必须输入！", vbCritical, "系统提示信息"
            Exit Sub
        End If
    
        If Trim(txt_det_code.Text) <> "" And Trim(txt_det_name.Text) = "" Then
            MsgBox "评审原因代码不正确！", vbCritical, "系统提示信息"
            Exit Sub
        End If
        
        If txt_proc_cd = "CAD" And Trim(txt_det_code.Text) <> "" Then
            If MsgBox("板坯号 " & cbo_slab_no.Text & " 已处于评审中，确定要再次录入评审原因吗？", vbExclamation + vbYesNo, "系统提示信息") = vbNo Then
               Exit Sub
            End If
        End If
        
        If Gf_Mc_Authority(sAuthority, Mc1) Then
            If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
        End If
        
    ElseIf txt_oper.Text = "2" Then
        If Not IsDate(TXT_OCCR_TIME) Then
           MsgBox "请正确输入发生时间！", vbCritical, "系统提示信息"
           Exit Sub
        End If
        
        If Mid(Trim(TXT_PRC_INPUT.Text), 1, 1) <> "B" And Mid(Trim(TXT_PRC_INPUT.Text), 1, 1) <> "C" Then
           MsgBox "工序代码错误！", vbCritical, "系统提示信息"
           Exit Sub
        End If
        
        If Gf_Mc_Authority(sAuthority, Mc2) Then
           txt_UserId.Text = sUserID
           If Mid(TXT_PRC_INPUT.Text, 1, 1) = "B" Then
            cbo_plt.Text = "B1"
           ElseIf Mid(TXT_PRC_INPUT.Text, 1, 1) = "C" Then
            cbo_plt.Text = "C1"
           End If
           CBO_LINE.Text = "1"
           txt_Flag.Text = "B1"
           
            If Gf_Ms_Process(M_CN1, Mc3, sAuthority) Then
               Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc2, Mc2("nControl"), Mc2("mControl"))
               Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
               Call MenuTool_ReSet
            End If
        End If
    End If
'    Call Form_Cls
'    Chk_ss2.Value = -1
    Call Form_Ref

End Sub

Public Sub Form_Del()

    Dim sQuery As String
    Dim sMessg As String
    Dim OutParam(2, 4) As Variant
    
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

    If Not Gf_MessConfirm("您确定要删除当前数据吗？", "Q") Then Exit Sub
    
    'Delete Make Query
    sQuery = Gf_Ms_MakeQuery(Mc3.Item("P-M"), "D", Mc3.Item("iControl"))
    
    If sQuery = "FAIL" Then
        Call Gp_MsgBoxDisplay("Delete Query Error : ")
        Exit Sub
    End If

    'sMessg = Gf_Ms_Display(Conn, sQuery, Mc.Item("rControl"), Mc.Item("lControl"))
    
    'Query Process
    If Gf_Ms_ExecQuery(OutParam, M_CN1, sQuery) Then
        Call Gp_Ms_ControlLock(Mc3!pControl, False)
        Call Form_Ref
        MDIMain.StatusBar1.Panels(1) = "提示信息：数据删除成功"
        TXT_PRC_INPUT.Text = ""
        CBO_SCRAP_INPUT.Text = ""
        txt_code.Text = ""
        TXT_SCRAP_NO.Text = ""
        TXT_OCCR_TIME.Text = ""
        cbo_shift1.Text = ""
        cbo_group1.Text = ""
        SDB_SCRAP_WGT.Text = ""
        cbo_ths_d_mat_var.Text = ""
        sdb_ths_d_mat_var.Text = ""
    Else
        Call Gp_MsgBoxDisplay("删除错误!")
    End If
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim i, j As Integer
    
    With ss1

        .Row = .ActiveRow
        .Col = .ActiveCol
         If .Row = 1 Then
             If Text(.Col - 1).Text = "Y" Then
                Text(.Col - 1).Text = "N"
             Else
                Text(.Col - 1).Text = "Y"
             End If
         ElseIf .Row = 2 Then
             If Text(.Col + 7).Text = "Y" Then
                Text(.Col + 7).Text = "N"
             Else
                Text(.Col + 7).Text = "Y"
             End If
         ElseIf .Row = 3 Then
             If Text(.Col + 15).Text = "Y" Then
                Text(.Col + 15).Text = "N"
             Else
                Text(.Col + 15).Text = "Y"
             End If
         End If

    End With
    
End Sub


Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
If Row > 0 Then
    With ss2
        .Row = Row
        .Col = SS2_SCRAP_DATE     '1
         TXT_OCCR_TIME.RawData = Mid(.Text, 1, 4) + Mid(.Text, 6, 2) + Mid(.Text, 9, 2)
        .Col = SS2_SHIFT          '2
         cbo_shift1.Text = .Text
        .Col = SS2_GROUP_CD       '3
         cbo_group1.Text = .Text
        .Col = SS2_PRC            '4
         TXT_PRC_INPUT.Text = .Text
        .Col = SS2_MAT_KIND       '6
         CBO_SCRAP_INPUT.Text = .Text
         TXT_SCRAP_INPUT.Text = Left(.Text, 2)
        .Col = SS2_MAT_NO         '7
         TXT_SCRAP_NO.Text = .Text
        .Col = SS2_SCRAP_WGT      '8
         SDB_SCRAP_WGT.Text = .Text
        .Col = SS2_SCRAP_RES      '9
         txt_code.Text = .Text
        .Col = SS2_SCRAP_RES1     '10
         txt_main_res_cd.Text = .Text
         .Col = SS2_PLT           '14
         cbo_plt.Text = .Text
    End With
End If
End Sub

Private Sub ss3_EditChange(ByVal Col As Long, ByVal Row As Long)
    Dim iIdr   As Integer
    Dim dLen   As Double
    Dim dWid   As Double
    
    If Row < 1 Then Exit Sub
    
    dLen = 0
    dWid = 0
    
    If Col = SS3_WID Then        '3
        ss3.Row = Row
        ss3.Col = Col
        If ss3.Text > SDB_WID.VALUE Then
            Call MsgBox("宽度错误!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
            Exit Sub
        End If
    ElseIf Col = SS3_LEN Then    '4
        For iIdr = 1 To ss3.MaxRows - 1
            ss3.Row = iIdr
            ss3.Col = SS3_LEN    '4
            dLen = dLen + Val(ss3.Text & "")
        Next iIdr
        
        ss3.Row = ss3.MaxRows
        ss3.Col = SS3_LEN        '4
        

    End If
End Sub

Private Sub SSCheck1_Click(VALUE As Integer)

End Sub

Private Sub txt_code_DblClick()

    Call txt_code_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_det_code_DblClick()

    Call txt_det_code_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_det_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
       Set DD.sPname = ss1
       DD.sWitch = "MS"
       DD.sKey = "C0017"
       DD.rControl.Add Item:=txt_det_code
       DD.rControl.Add Item:=txt_det_name
    
       DD.nameType = "1"
      
       Call Gf_Common_DD(M_CN1, KeyCode)
    ElseIf Len(Trim(txt_det_code.Text)) = 4 Then
       txt_det_name.Text = Gf_ComnNameFind(M_CN1, "C0017", Trim(txt_det_code), 1)
    Else
       txt_det_name.Text = ""
    End If

End Sub

Private Sub TXT_From_Date_DblClick()
    txt_from_DATE.RawData = Format(Date, "YYYYMMDD")
End Sub

Private Sub txt_PRC_DblClick()

    Call txt_PRC_KeyUp(vbKeyF4, 0)

End Sub

Private Sub TXT_PRC_INPUT_DblClick()

    Call TXT_PRC_INPUT_KeyUp(vbKeyF4, 0)

End Sub

Private Sub TXT_SCRAP_NO_Change()
Dim sQuery As String
Dim WGT As Double

    If Len(TXT_SCRAP_NO.Text) = 10 Then
       sQuery = "SELECT WGT FROM FP_SLAB WHERE SLAB_NO = '" + TXT_SCRAP_NO + "'"
       WGT = Gf_FloatFind(M_CN1, sQuery)
       If WGT = 0 Then
          MsgBox "该板坯不存在，板坯号无效！", vbCritical, "系统提示信息"
       Else
          SDB_SCRAP_WGT.Text = WGT
       End If
    End If
    
End Sub

Private Sub TXT_To_Date_DblClick()
    txt_to_DATE.RawData = Format(Date, "YYYYMMDD")
End Sub

Private Sub TXT_OCCR_TIME_DblClick()
  
    TXT_OCCR_TIME.RawData = Format(Date, "YYYYMMDD")
         
End Sub

Private Sub txt_code_Change()
    
    If Mid(TXT_PRC_INPUT.Text, 1, 1) = "C" Then
        
        If Len(Trim(txt_code)) = txt_code.MaxLength Then
            txt_main_res_cd.Text = Gf_ComnNameFind(M_CN1, "G0043", Trim(txt_code.Text), 1)
        Else
            txt_main_res_cd.Text = ""
        End If
 
        
    ElseIf Mid(TXT_PRC_INPUT.Text, 1, 1) = "B" Then
   
        If Len(Trim(txt_code)) = txt_code.MaxLength Then
            txt_main_res_cd.Text = Gf_ComnNameFind(M_CN1, "F0011", Trim(txt_code.Text), 1)
        Else
            txt_main_res_cd.Text = ""
        End If
        
    End If
    
End Sub

Private Sub txt_code_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sMesg As String
    
    If KeyCode = vbKeyF4 Then
            
        If Mid(TXT_PRC_INPUT.Text, 1, 1) = "C" Then
        
            DD.sWitch = "MS"
            DD.sKey = "G0043"
            DD.rControl.Add Item:=txt_code
            DD.rControl.Add Item:=txt_main_res_cd
    
            DD.nameType = "1"
    
            Call Gf_Common_DD(M_CN1, KeyCode)
            Exit Sub
            
        ElseIf Mid(TXT_PRC_INPUT.Text, 1, 1) = "B" Then
        
            DD.sWitch = "MS"
            DD.sKey = "F0011"
            DD.rControl.Add Item:=txt_code
            DD.rControl.Add Item:=txt_main_res_cd
    
            DD.nameType = "1"
    
            Call Gf_Common_DD(M_CN1, KeyCode)
            Exit Sub
            
        Else
            sMesg = " 工序代码错误 ！"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        
        End If
            
    End If

End Sub

Private Sub TXT_PRC_Change()
    
    If Len(Trim(txt_PRC)) = txt_PRC.MaxLength Then
        TXT_PRC_NAME.Text = Gf_ComnNameFind(M_CN1, "C0002", Trim(txt_PRC.Text), 2)
    Else
        TXT_PRC_NAME.Text = ""
    End If
    
End Sub

Private Sub txt_PRC_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
            
        DD.sWitch = "MS"
        DD.sKey = "C0002"
        DD.rControl.Add Item:=txt_PRC
        
        DD.nameType = "1"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    End If
    
End Sub

Private Sub TXT_PRC_INPUT_Change()
    
    If Len(Trim(TXT_PRC_INPUT)) = TXT_PRC_INPUT.MaxLength Then
        TXT_PRC_INPUT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0002", Trim(TXT_PRC_INPUT.Text), 2)
    Else
        TXT_PRC_INPUT_NAME.Text = ""
    End If
    
    txt_code.Text = ""
    
End Sub

Private Sub TXT_PRC_INPUT_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
            
        DD.sWitch = "MS"
        DD.sKey = "C0002"
        DD.rControl.Add Item:=TXT_PRC_INPUT
        
        DD.nameType = "1"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    End If
    
End Sub

Private Sub txt_work_date_DblClick()
         
    txt_work_date.RawData = Format(Now, "YYYYMMDDHH24MM")

End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Spread Copy
        .Buttons(12).Enabled = False                'Paste
    End With

End Sub


Private Sub cmd_divide_Click()

    Dim sTemp           As String
    Dim iDivCnt         As Integer
    Dim iIdr            As Integer
    Dim iIdc            As Integer
    Dim nSlab_no        As Integer
    
    
    If SDB_DIVIDE_CNT.VALUE < 2 Then
        Call MsgBox("板坯分板数错误！" & Chr(10) & "请重新输入。", vbExclamation + vbOKOnly, "警告")
        Exit Sub
    End If
    
    iDivCnt = SDB_DIVIDE_CNT.VALUE
    
    Set AdoRs = New ADODB.Recordset
    
    sQuery = "         SELECT  MAX(SLAB_NO)                  " & vbCrLf
    sQuery = sQuery & "  FROM  FP_SLAB                       " & vbCrLf
    sQuery = sQuery & " WHERE  SLAB_NO  LIKE '" & Mid(cbo_slab_no1.Text, 1, 8) & "%'" & vbCrLf
'    sQuery = sQuery & "   AND  ORD_FL   = '2'                " & vbCrLf
'    sQuery = sQuery & "   AND  REC_STS  = '2'                " & vbCrLf
    
    AdoRs.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    If CInt(Mid(AdoRs(0), 9, 2)) > 60 Then
       nSlab_no = CInt(Mid(AdoRs(0), 9, 2))
    Else
       nSlab_no = 60
    End If
    AdoRs.Close
     
    ss3.ReDraw = False
    
    ss3.MaxRows = iDivCnt
    
    sQuery = "         SELECT        SLAB_NO                      " & vbCrLf
    sQuery = sQuery & "             ,THK                          " & vbCrLf
    sQuery = sQuery & "             ,WID                          " & vbCrLf
    sQuery = sQuery & "             ,LEN                          " & vbCrLf
    sQuery = sQuery & "       FROM  FP_SLAB                       " & vbCrLf
    sQuery = sQuery & "      WHERE  SLAB_NO  = '" & cbo_slab_no1.Text & "'" & vbCrLf
    sQuery = sQuery & "        AND  ORD_FL   = '2'                " & vbCrLf
    sQuery = sQuery & "        AND  REC_STS  = '2'                " & vbCrLf

    With ss3
        If Gf_Only_Display(M_CN1, sc2, sQuery, , , False) Then
            .MaxRows = iDivCnt
            .BlockMode = True
            .Row = -1:  .Col = -1:     .Lock = True
            .BlockMode = False
                        
            For iIdr = 1 To iDivCnt
                nSlab_no = nSlab_no + 1
                .Row = iIdr
                .Col = 1
                .Text = Left(cbo_slab_no1.Text, 8) & CStr(nSlab_no)
                
                If iIdr < iDivCnt Then
                    For iIdc = 2 To .MaxCols
                        .Col = iIdc
                        .Row = iIdr
                        sTemp = .Text
                        
                        .Row = iIdr + 1
                        .Text = sTemp
                    Next iIdc
                End If
            Next iIdr
            .Row = 1:       .Col = SS3_WID:            SDB_WID.VALUE = .VALUE
            .Row = 1:       .Col = SS3_LEN:            SDB_LEN.VALUE = .VALUE
            
            .Row = -1
            .Col = SS3_WID:       .Lock = False:      .BackColor = &HC0FFFF
            .Col = SS3_LEN:       .Lock = False:      .BackColor = &HC0FFFF
        End If
    End With
               
    ss3.ReDraw = True
    
    cmd_divide_ok.Enabled = True
End Sub

Private Sub cmd_divide_ok_Click()
On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4)      As Variant
    Dim ret_Result_ErrMsg   As String
    Dim sSlabNo             As String
    Dim sEndCd              As String
    Dim icount              As Single
    Dim iIdr                As Single
    Dim iIdc                As Single
    Dim iThk                As Single
    Dim iWid                As Single
    Dim iLen                As Single
    Dim iNum                As Integer
    
    Dim adoCmd As ADODB.Command
    
    iLen = 0
    
    For iIdr = 1 To ss3.MaxRows
        ss3.Row = iIdr
        ss3.Col = SS3_WID       '3
        If Val(ss3.Text & "") <= 0 Then
            Call MsgBox("宽度错误!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
            Exit Sub
        End If
        ss3.Row = iIdr
        ss3.Col = SS3_LEN       '4
        If Val(ss3.Text & "") <= 0 Then
            Call MsgBox("长度错误!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
            Exit Sub
        End If
        iLen = iLen + Val(ss3.Text & "")
    Next iIdr
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    iNum = Val(SDB_DIVIDE_CNT.Text)
    For iIdr = 1 To ss3.MaxRows
        ss3.Row = iIdr
        ss3.Col = 1
        sSlabNo = ss3.Text
        ss3.Col = 2
        iThk = Val(ss3.Text & "")
        ss3.Col = SS3_WID      '3
        iWid = Val(ss3.Text & "")
        ss3.Col = SS3_LEN      '4
        iLen = Val(ss3.Text & "")
        sEndCd = ""
        If iIdr = ss3.MaxRows Then sEndCd = "Y"
        
        M_CN1.CursorLocation = adUseServer
        Set adoCmd = New ADODB.Command
        Set adoCmd.ActiveConnection = M_CN1
                
        '---------squery(CALL AFM2040P)----------------------
        sQuery = "{CALL AFM2040P('I','" & _
                                 sSlabNo & "','" & _
                                 cbo_slab_no1.Text & "'," & _
                                 iThk & ",    " & _
                                 iWid & ",    " & _
                                 iLen & ",   " & _
                                 iNum & ",    '" & _
                                 sEndCd & "','" & _
                                 sUserID & "',?)}"
        
        '-------------------------------------------------------
        
        adoCmd.CommandType = adCmdText
        adoCmd.CommandText = sQuery
        adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
        adoCmd.Execute , , adExecuteNoRecords
        
        'Process Error Check
        If adoCmd("arg_e_msg") <> "" Then
            ret_Result_ErrMsg = adoCmd("arg_e_msg")
            GoTo Process_Exec_ERROR
        End If
        Set adoCmd = Nothing
    Next iIdr
    
    Call Gp_MsgBoxDisplay("处理完了..!!", "I")
    SDB_DIVIDE_CNT.VALUE = 0
    cmd_divide_ok.Enabled = False
    cmd_divide_delete.Enabled = True
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error & "   " & ret_Result_ErrMsg)
End Sub

Private Sub cmd_divide_delete_Click()
On Error GoTo Process_Delete_ERROR

    Dim OutParam(1, 4)      As Variant
    Dim ret_Result_ErrMsg   As String
    
    Dim adoCmd As ADODB.Command
    
    If Trim(cbo_slab_no1.Text) = "" Or Len(cbo_slab_no1.Text) <> 10 Then
        Call MsgBox("板坯号错误!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
        Exit Sub
    End If
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    M_CN1.CursorLocation = adUseServer
    
    Set adoCmd = New ADODB.Command

    '---------squery(CALL AFM2040P)-Delete-------------------------------------------------------
    sQuery = "{CALL AFM2040P('D','" & cbo_slab_no1.Text & "','',0,0,0,'','','" & sUserID & "',?)}"
    '--------------------------------------------------------------------------------------------
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        GoTo Process_Delete_ERROR
    End If

    Set adoCmd = Nothing
    
    Call Gp_MsgBoxDisplay("处理完了..!!", "I")
    ss3.MaxRows = 0
    cmd_divide_ok.Enabled = True
    cmd_divide_delete.Enabled = False
    Exit Sub

Process_Delete_ERROR:

    Set adoCmd = Nothing
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error & "   " & ret_Result_ErrMsg)
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

