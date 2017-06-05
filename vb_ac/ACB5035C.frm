VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACB5035C 
   Caption         =   "半成品卸车实绩录入_ACB5035C"
   ClientHeight    =   9420
   ClientLeft      =   705
   ClientTop       =   2070
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter2 
      Height          =   9360
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   16510
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   1
      PaneTree        =   "ACB5035C.frx":0000
      Begin Threed.SSFrame SSFrame1 
         Height          =   1335
         Left            =   15
         TabIndex        =   1
         Top             =   15
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   2355
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.TextBox txt_plate_no 
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
            Left            =   9960
            MaxLength       =   15
            TabIndex        =   32
            Tag             =   "钢板号"
            Top             =   930
            Visible         =   0   'False
            Width           =   1830
         End
         Begin VB.CheckBox CHE_LOT 
            BackColor       =   &H00E0E0E0&
            Caption         =   "轧批"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   12990
            TabIndex        =   12
            Top             =   870
            Width           =   735
         End
         Begin VB.CheckBox chk_Excel_Fl 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Excel下载后打印"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   12990
            TabIndex        =   11
            Top             =   570
            Width           =   1815
         End
         Begin VB.TextBox text_cur_inv_name 
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
            Left            =   3840
            TabIndex        =   10
            Tag             =   "起始库"
            Top             =   135
            Width           =   1230
         End
         Begin VB.TextBox txt_to_inv_name 
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
            Left            =   7020
            TabIndex        =   9
            Tag             =   "目标库"
            Top             =   135
            Width           =   1185
         End
         Begin VB.TextBox txt_mv_lst_no 
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
            Left            =   1395
            MaxLength       =   15
            TabIndex        =   8
            Tag             =   "移拨码单号"
            Top             =   525
            Width           =   1830
         End
         Begin VB.TextBox text_cur_inv_code 
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
            Left            =   3390
            MaxLength       =   2
            TabIndex        =   7
            Tag             =   "起始库"
            Top             =   135
            Width           =   435
         End
         Begin VB.TextBox txt_to_inv 
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
            Left            =   6570
            MaxLength       =   2
            TabIndex        =   6
            Tag             =   "目标库"
            Top             =   135
            Width           =   435
         End
         Begin VB.TextBox text_prod_cd 
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
            Left            =   1395
            MaxLength       =   2
            TabIndex        =   5
            Tag             =   "产品"
            Top             =   135
            Width           =   465
         End
         Begin VB.TextBox txt_trans_way 
            Height          =   345
            Left            =   5070
            TabIndex        =   4
            Top             =   90
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.ComboBox CBO_GATE 
            BackColor       =   &H00FFFFFF&
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
            ItemData        =   "ACB5035C.frx":0052
            Left            =   11220
            List            =   "ACB5035C.frx":0080
            TabIndex        =   3
            Top             =   915
            Width           =   780
         End
         Begin VB.TextBox TXT_PASS_NO 
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
            Left            =   13110
            TabIndex        =   2
            Tag             =   "移拨码单号"
            Top             =   900
            Visible         =   0   'False
            Width           =   1830
         End
         Begin Threed.SSCommand cmd_Multi_Print 
            Height          =   345
            Left            =   12990
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   120
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   609
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "移拨码单打印"
         End
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   120
            Top             =   915
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "数量合计"
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
            Left            =   3405
            Top             =   915
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "重量合计"
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
         Begin CSTextLibCtl.sidbEdit text_tot_wgt 
            Height          =   315
            Left            =   4680
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   915
            Width           =   1230
            _Version        =   262145
            _ExtentX        =   2170
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   255
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
            ReadOnly        =   -1  'True
            Insert          =   0   'False
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
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit text_tot_sheets 
            Height          =   315
            Left            =   1395
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   915
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   16711680
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
            ReadOnly        =   -1  'True
            Insert          =   0   'False
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
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   8430
            Top             =   135
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "转库日期"
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
         Begin InDate.UDate udate_in_plt_date_a 
            Height          =   315
            Left            =   9720
            TabIndex        =   16
            Tag             =   "转库日期"
            Top             =   135
            Width           =   1440
            _ExtentX        =   2540
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
         End
         Begin InDate.UDate udate_in_plt_date_b 
            Height          =   315
            Left            =   11370
            TabIndex        =   17
            Tag             =   "转库日期"
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
         End
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   2100
            Top             =   135
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "起始库"
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
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   5280
            Tag             =   "目标库"
            Top             =   135
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "目标库"
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
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   120
            Top             =   135
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "产品"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.76
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16711680
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   120
            Top             =   525
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "移拨码单号"
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
         Begin InDate.ULabel ULabel34 
            Height          =   315
            Left            =   3405
            Top             =   525
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "到达日期"
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
         Begin CSTextLibCtl.sitxEdit txt_input_date 
            Height          =   315
            Left            =   4680
            TabIndex        =   18
            Tag             =   "到达日期"
            Top             =   525
            Width           =   2115
            _Version        =   262145
            _ExtentX        =   3731
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   "____-__-__ __-__-__"
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
         Begin Threed.SSCommand cmd_input 
            Height          =   345
            Left            =   6885
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   510
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   609
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "录入转库实绩"
         End
         Begin Threed.SSCommand cmd_Print 
            Height          =   345
            Left            =   5850
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   2460
            Visible         =   0   'False
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   609
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "码单打印"
         End
         Begin Threed.SSCommand cmd_Gate 
            Height          =   345
            Left            =   9720
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   510
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   609
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "出门岗确认"
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   9720
            Top             =   915
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            Caption         =   "门 岗"
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
         Begin Threed.SSCommand CMD_CARD 
            Height          =   315
            Left            =   6450
            TabIndex        =   28
            Top             =   930
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   255
            Caption         =   "作业指示单打印"
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   315
            Left            =   8040
            TabIndex        =   29
            Top             =   930
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   255
            Caption         =   "Excel导出"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   135
            Left            =   11145
            TabIndex        =   25
            Top             =   225
            Width           =   255
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "吨"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6030
            TabIndex        =   24
            Top             =   975
            Width           =   195
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "件"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2445
            TabIndex        =   23
            Top             =   975
            Width           =   195
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "号岗"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   12090
            TabIndex        =   22
            Top             =   960
            Width           =   390
         End
      End
      Begin SSSplitter.SSSplitter SSSplitter1 
         Height          =   7935
         Left            =   15
         TabIndex        =   26
         Top             =   1410
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   13996
         _Version        =   196609
         SplitterBarWidth=   3
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   16761087
         PaneTree        =   "ACB5035C.frx":00BC
         Begin FPSpread.vaSpread ss2 
            Height          =   2895
            Left            =   0
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   0
            Width           =   15975
            _Version        =   393216
            _ExtentX        =   28178
            _ExtentY        =   5106
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ButtonDrawMode  =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   16
            MaxRows         =   5
            Protect         =   0   'False
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "ACB5035C.frx":012E
         End
         Begin FPSpread.vaSpread ss1 
            Height          =   4905
            Left            =   0
            TabIndex        =   30
            Top             =   2940
            Width           =   15975
            _Version        =   393216
            _ExtentX        =   28178
            _ExtentY        =   8652
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
            MaxCols         =   70
            MaxRows         =   5
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "ACB5035C.frx":0911
         End
         Begin FPSpread.vaSpread ss3 
            Height          =   45
            Left            =   0
            TabIndex        =   31
            Top             =   7890
            Visible         =   0   'False
            Width           =   15975
            _Version        =   393216
            _ExtentX        =   28178
            _ExtentY        =   79
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
            MaxCols         =   70
            MaxRows         =   5
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "ACB5035C.frx":24CC
         End
      End
   End
End
Attribute VB_Name = "ACB5035C"
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
'-- Program ID        ACB5030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2007.8.12
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
Public sAuthority_Cmd As String     'Active Button Authority Setting

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

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection



Dim SumCnt   As Integer
Dim SumCol   As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2
Dim INF_KND  As String

Const SS2_MV_LST_NO = 1
Const SS2_PROD_CD = 2
Const SS2_FR_INV = 3
Const SS2_TO_INV = 5
Const SS2_CAR_NO = 7
Const SS2_MV_NUM = 8
Const SS2_MV_WGT = 9
Const SS2_MV_DATE = 10
Const SS2_MV_EMP = 12
Const SS2_TRANS_WAY = 16

Const SS1_PLATE_NO = 3
Const SS1_SPEC_NAME = 5
Const SS1_THK = 7
Const SS1_WID = 8
Const SS1_LEN = 9
Const SS1_WGT = 10
Const SS1_APLY_STDSPEC = 11
Const SS1_ORD_NO = 13
Const SS1_CUST_CD = 14
Const SS1_TRIM_FL = 15
Const SS1_SIZE_KND = 16
Const SS1_RCV_DATE = 19
Const SS1_RCV_EMP = 20
Const SS1_OUT_SHEET_NO = 22
Const SS1_UST_STATUS = 23
Const SS1_GAS_STATUS = 24
Const SS1_CL_STATUS = 25
Const SS1_HTM_METH = 26
Const SS1_QT = 27
Const SS1_MAC = 28
Const SS1_TRANS_WAY = 29
Const SS1_TRANS_COMP = 30
Const SS1_TRANS_TOOL = 31
Const SS1_PLT = 32
Const SS1_THK_AVG = 33
Const SS1_ORD_THK = 34
Const SS1_ORD_WID = 35
Const SS1_ORD_LEN = 36
Const SS1_LEN_AVG = 37
Const SS1_WID_AVG = 38
Const SS1_ORD_REMARK = 39
Const SS1_STDSPEC_ORG_KND = 40
Const SS1_STDSPEC_STLGRD = 41
Const SS1_MV_DATE = 42
Const SS1_Shift = 43
Const SS1_TRNS_CMPY_CD = 44
Const SS1_CUST_CD1 = 45


Const SS1_PLATE_CON = 53   '子板数
Const SS1_PLATE_SIZE = 54   '子板尺寸
Const SS1_CE_APPR_FL = 55
Const SS1_QS_MARK_FL = 56
Const SS1_VESSEL_NO = 57
Const SS1_PAINTNUM = 58
Const SS1_GANGYIN = 59
Const SS1_SIDEMARK = 60
Const SS1_RM_CR_STAGE3_TIME = 61 '订单数量
Const SS1_PUNCH = 62 '钢印加冲
Const SS1_CUST = 63 '客户交货期
Const SS1_SURFACE_REQUESTS = 67 '客户交货期
Const SS1_PROD_REMARK = 70 '产品备注









Private Sub Form_Define()
        
     'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
      FormType = "Msheet"
         
           Call Gp_Ms_Collection(Text_PROD_CD, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(text_cur_inv_code, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_to_inv, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udate_in_plt_date_a, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udate_in_plt_date_b, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_mv_lst_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                                                            
      'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
                                                      
    ' control part   Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 53, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
   Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) ' ADD BY 李超 20141011
   Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 58, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 59, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 60, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 61, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 62, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '加冲钢印
   Call Gp_Sp_Collection(ss1, 63, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '客户交货期
   Call Gp_Sp_Collection(ss1, 64, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 65, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 66, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 67, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 68, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 69, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 70, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '产品备注
  
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB5035C.P_REFER", Key:="P-R"
    sc1.Add Item:="ACB5035C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="ACB5035C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
                                                  
    ' control part   Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(SS2, 1, "p", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 2, "p", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(SS2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(SS2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(SS2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '12
   Call Gp_Sp_Collection(SS2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '13
   Call Gp_Sp_Collection(SS2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '13-> 14
   Call Gp_Sp_Collection(SS2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '15
   Call Gp_Sp_Collection(SS2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '16
   
   
    'Spread_Collection
    sc2.Add Item:=SS2, Key:="Spread"
    sc2.Add Item:="ACB5035C.P_SREFER", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=2, Key:="First"
    sc2.Add Item:=SS2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(SS2, SS2_TRANS_WAY, True)
    
    
   Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, "", " ", " ", "", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, "", " ", " ", "", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", "", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 13, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 14, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 15, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 16, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 17, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 18, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 19, " ", " ", " ", "", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 20, " ", " ", " ", "", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 21, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 22, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 23, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 24, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 25, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 26, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 27, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 28, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 29, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 30, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 31, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 32, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 33, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 34, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 35, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 36, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 37, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 38, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 39, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 40, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 41, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 42, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 43, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 44, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 45, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 46, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 47, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 48, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 49, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 50, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 51, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 52, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 53, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 54, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   
   Call Gp_Sp_Collection(ss3, 55, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3) ' ADD BY 李超 20141011
   Call Gp_Sp_Collection(ss3, 56, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 57, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 58, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 59, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 60, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 61, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss1, 62, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3) '加冲钢印
   Call Gp_Sp_Collection(ss1, 63, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3) '客户交货期
   Call Gp_Sp_Collection(ss1, 64, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss1, 65, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss1, 66, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss1, 67, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss1, 68, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss1, 69, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 70, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3) '产品备注
  
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="ACB5035C.P_SREFER1", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=3, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc3, Key:="Sc3"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
End Sub

Private Sub CMD_CARD_Click()
    Call ExcelPrn_Pile(Me, ss1, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End Sub

Private Sub cmd_Gate_Click()

    Dim iDR As Long
    Dim sMsg As String
    
    If Not Gf_Sc_Authority(sAuthority_Cmd, "U") Then
       Call Gp_MsgBoxDisplay("您没有权限操作此功能，工号：" & sUserID)
       Exit Sub
    End If

    If Trim(txt_mv_lst_no.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_mv_lst_no.Tag & "必须输入")
        Exit Sub
    End If
    
    If Trim(txt_to_inv_name.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_to_inv.Tag & "必须输入")
        Exit Sub
    End If
    
    If Mid(Trim(txt_mv_lst_no.Text), 2, 4) <> text_cur_inv_code.Text & txt_to_inv.Text Then
        Call Gp_MsgBoxDisplay("移拨码单号与起始库/目标库代码不一致,请确认")
        Exit Sub
    End If
        
    If Trim(CBO_GATE.Text) = "" Then
        Call Gp_MsgBoxDisplay("请选择门岗号")
        Exit Sub
    End If
    
    sMsg = Cp_Move_Pass_Exec
    
    Call Gp_MsgBoxDisplay(sMsg)

    
End Sub
Public Function Cp_Move_Pass_Exec() As String

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrCode As String
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim sMsg As String
    Dim mResult As String

    Dim adoCmd As ADODB.Command

    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256

    sQuery = "{call ACB5032P ('" + txt_mv_lst_no.Text + "','" + CBO_GATE.Text + "','" + sUserID + "',?,?)}"

    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command

    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1

    adoCmd.CommandText = sQuery

    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))

    M_CN1.BeginTrans

    adoCmd.Execute , , adExecuteNoRecords
    
    ret_Result_ErrCode = adoCmd("arg_e_code")
    ret_Result_ErrMsg = adoCmd("arg_e_msg")
    

    'Process Error Check
    If ret_Result_ErrCode = "Y" Then
    
        M_CN1.CommitTrans
    
    Else
    
        M_CN1.RollbackTrans
        
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    If ret_Result_ErrCode = "Y" Then
       Cp_Move_Pass_Exec = "出门岗确认成功，码单号：" & ret_Result_ErrMsg
    Else
       Cp_Move_Pass_Exec = "出门岗确认失败：" & ret_Result_ErrMsg
    End If
    Exit Function

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Cp_Move_Pass_Exec = ""
    ERR.Raise ERR.Number, ERR.Description & sQuery

End Function

Private Sub cmd_input_Click()

    Dim iDx     As Long
    Dim sMvNo   As String
    
    If ss1.MaxRows = 0 Then Exit Sub
    
    If Not Gf_Sc_Authority(sAuthority, "U") Then Exit Sub
    
    If Trim(txt_mv_lst_no.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_mv_lst_no.Tag & "必须输入")
        Exit Sub
    End If
    
    If Not IsDate(txt_input_date.Text) Then
        Call Gp_MsgBoxDisplay(txt_input_date.Tag & "必须输入")
        Exit Sub
    End If
    
    sMvNo = Trim(txt_mv_lst_no.Text)
    
    With ss1
    
        For iDx = 1 To .MaxRows
            .ROW = iDx
            
            .Col = 2
            If sMvNo <> Trim(.Text) Then
                Call Gp_MsgBoxDisplay("移拨码单号不一样! 查询后处理一下..")
                Exit Sub
            End If
            
            .Col = SS1_RCV_DATE:     .Text = Trim(txt_input_date.Text)
            .Col = SS1_RCV_EMP:      .Text = sUserID
            .Col = 0:                .Text = "Update"
        Next iDx
        
    End With
    
End Sub

Private Sub cmd_Print_Click()

    Call Form_Exc
    
End Sub






Private Sub Text_PROD_CD_Change()
   
    If Len(Text_PROD_CD) <> 2 Then Exit Sub

    Select Case Text_PROD_CD.Text

        Case "PP", "pp"
            Text_PROD_CD.Text = "PP"
        Case "HC", "hc"
            Text_PROD_CD.Text = "HC"
        Case "MP", "mp"
            Text_PROD_CD.Text = "MP"
        Case "", "**"
            Text_PROD_CD.Text = ""
        Case Else
            Text_PROD_CD.Text = ""
            Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
    End Select

End Sub



Private Sub Text_PROD_CD_DblClick()

    Call Text_PROD_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_prod_cd_LostFocus()

    If Text_PROD_CD.Text <> "" Then
        If (Len(Text_PROD_CD.Text) < Text_PROD_CD.MaxLength) Then
            Call Gp_MsgBoxDisplay("产品分类代码输入未完成！")
            Text_PROD_CD.SetFocus
        End If
    End If

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

    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    sAuthority_Cmd = Gf_Pgm_Authority("ACB5032P")
    
    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"), False)
    Call Gp_Sp_Setting(Proc_Sc("Sc3")("Spread"), False)
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "C-System.INI", Me.Name)

    udate_in_plt_date_a.Text = Mid(udate_in_plt_date_a.Text, 1, 8) & "01"

    udate_in_plt_date_b.RawData = Gf_GetLastDay(udate_in_plt_date_b.RawData)
    
    Screen.MousePointer = vbDefault
    
    Text_PROD_CD.Text = "PP"
    
    text_cur_inv_code = "00"
    txt_to_inv.Text = "WD"

    Call txt_to_inv_KeyUp(0, 0)
        
    If Gf_Sc_Authority(sAuthority, "U") Then
        cmd_input.Enabled = True
    Else
        cmd_input.Enabled = False
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '查询结束

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc3.Item("Spread"), "C-System.INI", Me.Name)
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
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
        
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set Proc_Sc = Nothing
    Set SumCol = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gf_Sp_Cls(Proc_Sc("Sc2"))
        Call Gf_Sp_Cls(Proc_Sc("Sc3"))
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
    End If
 
    udate_in_plt_date_a.Text = Format(Date, "YYYY-MM-01")

    udate_in_plt_date_b.RawData = Gf_GetLastDay(udate_in_plt_date_b.RawData)
    text_tot_sheets.Value = 0
    text_tot_wgt.Value = 0
    txt_input_date.Text = ""
    Text_PROD_CD.Text = "PP"

    text_cur_inv_code = "00"
    txt_to_inv.Text = "WD"
    
    Call txt_to_inv_KeyUp(0, 0)
    chk_Excel_Fl.Value = 0
    
End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, ss1, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

'    Dim iDR As Long
'
'    If Trim(txt_mv_lst_no.Text) = "" Then
'        Call Gp_MsgBoxDisplay(txt_mv_lst_no.Tag & "必须输入")
'        Exit Sub
'    End If
'
'    If Trim(txt_to_inv_name.Text) = "" Then
'        Call Gp_MsgBoxDisplay(txt_to_inv.Tag & "必须输入")
'        Exit Sub
'    End If
'
'    Call ExcelPrn

End Sub

Public Sub Form_Ref()

    Dim sTotnum As Double
    Dim sTotwgt As Double
    Dim iCount As Double
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    
    ss1.MaxRows = 0
    txt_input_date.Text = ""
                    
    If Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        SS2.OperationMode = OperationModeNormal
    End If
    
    With SS2
    
        If .MaxRows <= 1 Then
           Exit Sub
        End If
        
        For iCount = 1 To .MaxRows
        
            .ROW = iCount:            .Col = SS2_MV_NUM
            sTotnum = sTotnum + .Value
            
            .ROW = iCount:            .Col = SS2_MV_WGT
            sTotwgt = sTotwgt + .Value
            
        Next iCount
        
        .MaxRows = .MaxRows + 1
        .ROW = .MaxRows:
        .Col = 1:           .Text = "  合  计 "
        .Col = SS2_MV_NUM:  .Value = sTotnum
        .Col = SS2_MV_WGT:  .Value = sTotwgt
        
    End With
    
    

    With SS2
    
        If .MaxRows = 0 Then
            text_tot_sheets.Text = "0"
            text_tot_wgt.Value = 0
        Else
            .ReDraw = False
            .ROW = .MaxRows
            .Col = SS2_MV_NUM:  text_tot_sheets.Text = Val(.Value & "")
            .Col = SS2_MV_WGT:  text_tot_wgt.Text = Val(.Value & "")
            .MaxRows = .MaxRows + 1
            .ROW = 1
            .Action = SS_ACTION_INSERT_ROW
            .Col = 1:   .Text = "  合  计 "
            .Col = SS2_MV_NUM:   .Text = text_tot_sheets.Text
            .Col = SS2_MV_WGT:  .Text = text_tot_wgt.Text
            Call Gp_Sp_BlockColor(sc2.Item("Spread"), 1, .MaxCols, 1, 1, BLACK, &HE6E6FF)
            .ROW = .MaxRows
            .Action = SS_ACTION_DELETE_ROW
            .MaxRows = .MaxRows - 1
            .ReDraw = True
        End If
        
    End With
    
End Sub

Public Sub Form_Pro()

    Dim iRow  As Long
    Dim iCount, max_row As Long
    
    iCount = 0
    max_row = ss1.MaxRows
    
    For iRow = 1 To ss1.MaxRows
        ss1.ROW = iRow
        ss1.Col = 0
        If ss1.Text = "Update" Then
            ss1.Col = SS1_RCV_DATE
            If Not IsDate(ss1.Text) Then
                Call Gp_MsgBoxDisplay("到达日期必须输入")
                Exit Sub
            End If
        End If
        
        If ss1.Text = "Delete" Then
           iCount = iCount + 1
        End If

    Next iRow
        
    Screen.MousePointer = vbHourglass
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        txt_input_date.Text = ""
    End If


'抛送erp计算运费
    If iCount > 0 And txt_trans_way.Text = 0 Then
       If max_row = iCount Then
          INF_KND = "D"
       Else
          INF_KND = "U"      '传erp  D/N两条记录
       End If
       Call carprice
    End If
       
    Screen.MousePointer = vbDefault
    
 
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Del()

    Dim iRow  As Long
    
    Call Gp_Sp_Del(sc1)
    
    For iRow = 1 To ss1.MaxRows
        ss1.ROW = iRow
        ss1.Col = 0
        If UCase(ss1.Text) = "DELETE" Then
            ss1.Col = SS1_RCV_EMP
            ss1.Text = sUserID
        End If
    Next iRow
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
          
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub text_cur_inv_code_DblClick()

    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_cur_inv_code_Change()
    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
        text_cur_inv_name.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
    Else
      text_cur_inv_name.Text = ""
    End If
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
         DD.sWitch = "MS"
         DD.sKey = "C0013"
    
         DD.rControl.Add Item:=text_cur_inv_code
    
         DD.nameType = "2"
         Call Gf_Common_DD(M_CN1, KeyCode)
    
    End If

End Sub

Public Sub Spread_Forzens_Setting()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal ROW As Long)

    Dim iRow As Integer
    Dim iCol As Integer
    Dim i As Integer

    If ROW < 1 Then Exit Sub

    SS2.ROW = ROW
    SS2.Col = SS2_MV_LST_NO
    If Len(Trim(SS2.Text)) > 10 Then
        txt_mv_lst_no.Text = SS2.Text
        SS2.Col = SS2_FR_INV
        text_cur_inv_code.Text = SS2.Text
        SS2.Col = SS2_TO_INV
        txt_to_inv.Text = SS2.Text
        SS2.Col = SS2_TRANS_WAY
        txt_trans_way.Text = SS2.Text
    Else
        txt_mv_lst_no.Text = ""
    End If

    SS2.Col = SS2_MV_NUM:  text_tot_sheets.Text = Val(SS2.Value & "")
    SS2.Col = SS2_MV_WGT: text_tot_wgt.Text = Val(SS2.Value & "")

    Call Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False)
    
    ss1.OperationMode = OperationModeNormal
    
     For iRow = 1 To ss1.MaxRows
    
          ss1.ROW = iRow
          ss1.Col = 51
           If ss1.Text = "Y" Then
              For i = 1 To ss1.MaxCols
                   ss1.Col = i
                   ss1.ForeColor = &HC000&
              Next
           End If
      
      Next iRow
      ss1.ROW = 1
      ss1.Col = SS1_PLATE_NO: TXT_PLATE_NO = ss1.Value
      Call Gf_Sp_Refer(M_CN1, Sc3, Mc1, Mc1("nControl"), Mc1("mControl"), False)
      ss3.OperationMode = OperationModeNormal
      
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_LostFocus()

'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
'   Call ss1_row_Click(Col, Row)

 ss1.ROW = ROW
      ss1.Col = SS1_PLATE_NO: TXT_PLATE_NO = ss1.Value
      
      
      Call Gf_Sp_Refer(M_CN1, Sc3, Mc1, Mc1("nControl"), Mc1("mControl"), False)
      ss3.OperationMode = OperationModeNormal

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)


    If ROW <= 0 Then Exit Sub
    
    ss1.ROW = ROW
    ss1.Col = Col

    If Mode = 1 Then
        ss1.Tag = ss1.Text
    Else
        If Trim(ss1.Tag) <> Trim(ss1.Text) Then
            ss1.Col = 0
            Select Case Trim(ss1.Text)
                Case "Input", "Update", "Delete"
                Case Else
                    ss1.Text = "Update"
                    ss1.Col = SS1_RCV_EMP:   ss1.Text = sUserID
            End Select
        End If
    End If
    
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
'        .Buttons(8).Enabled = False                 'Row Delete
'        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Spread Copy
        .Buttons(12).Enabled = False                'Paste
    End With

End Sub


Private Function Gf_GetLastDay(Optional DTDay As String = "") As Variant

On Error GoTo DGet_Error

    Dim sQuery As String
    Dim strDay As String
    
    If DTDay = "" Then
        sQuery = "SELECT TO_CHAR(LAST_DAY(SYSDATE),'YYYYMMDD') FROM DUAL"
    Else
       strDay = DTDay
       sQuery = "SELECT TO_CHAR(LAST_DAY(TO_DATE('" + strDay + "','YYYYMMDD')),'YYYYMMDD') FROM DUAL"
    End If
       
    Dim AdoRs As ADODB.Recordset
    
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            If VarType(AdoRs.Fields(0)) = vbNull Then
                Gf_GetLastDay = ""
            Else
                Gf_GetLastDay = AdoRs.Fields(0)
            End If
        End If
        
    Else
        Gf_GetLastDay = "00000000"
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

DGet_Error:

    Set AdoRs = Nothing
    Gf_GetLastDay = "00000000"

End Function

Private Sub txt_input_date_Click()

    txt_input_date.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDDHH24MISS') FROM DUAL")
    
End Sub

Private Sub txt_to_inv_DblClick()

    Call txt_to_inv_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_to_inv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
        DD.rControl.Add Item:=txt_to_inv
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    Else
        If Len(Trim(txt_to_inv.Text)) = txt_to_inv.MaxLength Then
            txt_to_inv_name.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_to_inv.Text, 2)
        Else
            txt_to_inv_name.Text = ""
        End If
    
    End If
    
End Sub

Private Sub Text_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=Text_PROD_CD
        'DD.rControl.Add Item:=Text_PROD_CD_Name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If
    
End Sub

Private Sub udate_in_plt_date_a_LostFocus()
'    UDate_IN_PLT_DATE_b.RawData = Gf_GetLastDay(UDate_IN_PLT_DATE_a.RawData)
End Sub


Private Sub ExcelPrn()

    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    Dim Wb              As Object
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If ERR.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    ERR.Clear

    Set Wb = xlApp.Workbooks.Open(App.Path & "\ACB5035C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    xlApp.Rows("5:200").Select
    xlApp.Selection.delete Shift:=1
    
    xlApp.Sheets("Sheet2").Select
    xlApp.Range("A1:J1").Select
    xlApp.Selection.Copy
    xlApp.Sheets("Sheet1").Select
    sRow = "A" & 5 & ":" & "J" & ss1.MaxRows + 5
    xlApp.Range(sRow).Select
    xlApp.ActiveSheet.Paste
            
    Select Case Text_PROD_CD.Text
        Case "PP"
            xlApp.Range("B3").Value = "钢板"
        Case "SL"
            xlApp.Range("B3").Value = "板坯"
        Case "HC"
            xlApp.Range("B3").Value = "钢卷"
    End Select
    
    xlApp.Range("D3").Value = Format(Date, "YYYY-MM-DD")
    xlApp.Range("G3").Value = txt_mv_lst_no.Text
    xlApp.Range("J3").Value = txt_to_inv_name.Text
          
    ss1.ROW = 1: ss1.Col = ss1.MaxCols
    If CHE_LOT = 1 Then
        Clipboard.Clear
        ss1.SetSelection SS1_OUT_SHEET_NO, 1, SS1_OUT_SHEET_NO, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("A5").Select
        xlApp.ActiveSheet.Paste
    Else
        Clipboard.Clear
        ss1.SetSelection SS1_PLATE_NO, 1, SS1_PLATE_NO, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("A5").Select
        xlApp.ActiveSheet.Paste
    End If
     
    Clipboard.Clear
    ss1.SetSelection SS1_THK, 1, SS1_WGT, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("G5").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection SS1_APLY_STDSPEC, 1, SS1_APLY_STDSPEC, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("B5").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection SS1_CUST_CD, 1, SS1_CUST_CD, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("C5").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection SS1_TRIM_FL, 1, SS1_SIZE_KND, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("D5").Select
    xlApp.ActiveSheet.Paste
    
    
    Clipboard.Clear
    ss1.SetSelection SS1_ORD_NO, 1, SS1_ORD_NO, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("F5").Select
    xlApp.ActiveSheet.Paste
    
    
    Clipboard.Clear
    ss1.SetSelection SS1_CUST, 1, SS1_CUST, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("K5").Select
    xlApp.ActiveSheet.Paste


    xlApp.Sheets("Sheet2").Select
    xlApp.Range("A2:J2").Select
    xlApp.Selection.Copy
    xlApp.Sheets("Sheet1").Select
    sRow = "A" & ss1.MaxRows + 5 & ":" & "J" & ss1.MaxRows + 5
    xlApp.Range(sRow).Select
    xlApp.ActiveSheet.Paste
        

    Clipboard.Clear
    sRow = "E" & ss1.MaxRows + 5
    xlApp.Range(sRow).Value = "数量合计: " & text_tot_sheets.Text

    Clipboard.Clear
    sRow = "I" & ss1.MaxRows + 5
    xlApp.Range(sRow).Value = "总计:"

    sRow = "J" & ss1.MaxRows + 5
    xlApp.Range(sRow).Value = text_tot_wgt.Text

    sRow = "A" & ss1.MaxRows + 6
    SS2.ROW = SS2.ActiveRow: SS2.Col = 14
    
'     sRow = "A" & ss1.MaxRows + 6
'    ss2.Row = ss2.ActiveRow: ss2.Col = 13
'
    xlApp.Range(sRow).Value = "转库发货员姓名:" & SS2.Text

    sRow = "D" & ss1.MaxRows + 6
    xlApp.Range(sRow).Value = "仓库操作员工号:" & sUserName

    sRow = "H" & ss1.MaxRows + 6
    SS2.ROW = SS2.ActiveRow: SS2.Col = SS2_CAR_NO
    xlApp.Range(sRow).Value = "车辆号:" & SS2.Text

    Clipboard.Clear
    xlApp.Range("A2").Select
    xlApp.ActiveSheet.Paste
    
    If chk_Excel_Fl = 0 Then
        xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    End If
    
    ss1.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    If chk_Excel_Fl = 0 Then
        xlApp.Application.Visible = False
        Wb.Close False
        xlApp.QuitSet
        Set Wb = Nothing
        Set xlApp = Nothing
    Else
        xlApp.Application.Visible = True
    End If
    
'    Wb.Close
'    xlApp.Quit
    
'    Set Wb = Nothing
'    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
'    xlApp.Application.Visible = True
    
    Set Wb = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub
Private Sub ExcelGatePrn()

    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    Dim Wb              As Object
    
    Dim iFr_inv         As String
    Dim iTo_inv         As String
    
    If ss1.MaxRows < 1 Or SS2.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If ERR.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    ERR.Clear

    Set Wb = xlApp.Workbooks.Open(App.Path & "\ACB5031C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select

    Select Case Text_PROD_CD.Text
        Case "PP"
            xlApp.Range("B9").Value = "热轧钢板"
        Case "MP"
            xlApp.Range("B9").Value = "母板"
        Case "HC"
            xlApp.Range("B9").Value = "热轧钢卷"
    End Select
    
    xlApp.Range("C3").Value = txt_mv_lst_no.Text
    xlApp.Range("J3").Value = TXT_PASS_NO.Text
    xlApp.Range("G13").Value = CBO_GATE.Text

    With SS2
        .ROW = .ActiveRow
        .Col = SS2_MV_LST_NO
        If .Text = txt_mv_lst_no.Text Then
        
           .Col = SS2_MV_DATE
           xlApp.Range("F3").Value = .Text
           
           .Col = SS2_FR_INV:           iFr_inv = .Text
           .Col = SS2_TO_INV:           iTo_inv = .Text
           xlApp.Range("J4").Value = iFr_inv & " / " & iTo_inv
           
           .Col = SS2_CAR_NO
           xlApp.Range("B7").Value = .Text
           
           .Col = SS2_FR_INV
           xlApp.Range("G7").Value = .Text
           
           .Col = SS2_MV_NUM
           xlApp.Range("I9").Value = .Text
           xlApp.Range("I11").Value = .Text
           
           .Col = SS2_MV_WGT
           xlApp.Range("J9").Value = .Text
           xlApp.Range("J11").Value = .Text
           
           .Col = SS2_MV_EMP
           xlApp.Range("B13").Value = .Text
           
           xlApp.Range("C13").Value = sUserName
           
        End If
    End With
    
    With ss1
    
           .ROW = 1
           .Col = SS1_PLT
           xlApp.Range("C5").Value = .Text
        
           .Col = SS1_TRANS_WAY
           xlApp.Range("C7").Value = .Text
           
           .Col = SS1_TRANS_TOOL
           xlApp.Range("D7").Value = .Text
           
           .Col = SS1_TRANS_COMP
           xlApp.Range("C11").Value = .Text

    End With
    
    If chk_Excel_Fl = 0 Then
        xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    End If
    
    ss1.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    If chk_Excel_Fl = 0 Then
        xlApp.Application.Visible = False
        Wb.Close False
        xlApp.QuitSet
        Set Wb = Nothing
        Set xlApp = Nothing
    Else
        xlApp.Application.Visible = True
    End If
        
    Exit Sub

ErrHandle:
    MsgBox Error
'    xlApp.Application.Visible = True
    
    Set Wb = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmd_Multi_Print_Click()

    Dim iDR As Long

    If Trim(txt_mv_lst_no.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_mv_lst_no.Tag & "必须输入")
        Exit Sub
    End If
    
    If Trim(txt_to_inv_name.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_to_inv.Tag & "必须输入")
        Exit Sub
    End If
    
    Call ExcelPrn

'    Dim iDR     As Long
'    Dim sFromNo As String
'    Dim sToNo   As String
'
'    If Not Gf_Sc_Authority(sAuthority, "U") Then Exit Sub
'
'    If lBlkrow1 < 2 And lBlkrow2 < 2 Then Exit Sub
'
'    ss2.Col = 1
'    ss2.Row = lBlkrow1:     sFromNo = ss2.Text
'    ss2.Row = lBlkrow2:     sToNo = ss2.Text
'
'    If Not Gf_MessConfirm("您确定要码单多种打印(" & sFromNo & " ~ " & sToNo & ")吗？", "Q") Then Exit Sub
'
'    For iDR = lBlkrow1 To lBlkrow2
'        If iDR > 1 Then
'            Call ss2_DblClick(1, iDR)
'            Call ExcelPrn
'        End If
'    Next iDR
        
End Sub


Private Function carprice() As Boolean
    
On Error GoTo PRODEND_Error

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
   
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    If txt_mv_lst_no.Text = "" Then
        Call MsgBox("装车单为空，传运费系统失败", vbInformation, "系统提示信息")
        Screen.MousePointer = vbDefault
        Exit Function
    End If
          
    sQuery = "{call ARC0180P( '" + INF_KND + "','" + txt_mv_lst_no.Text + "',?)}"
  '  sQuery = "{call ACA1031P ('" + txt_ord_no + "', '" + Combo1.Text + "','" + TXT_REASON + "','" + sUserName + "',?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Execute , , adExecuteNoRecords
    
    'OS Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
'    Else
'
'        Call MsgBox(CMD_PRODEND.Caption + "完成！", vbInformation, "系统提示信息")
'        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Function

PRODEND_Error:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("运费抛送失败: " & Error)

        
End Function

Public Sub ExcelPrn_Pile(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

On Error GoTo Excel_Error

    Dim ret         As Boolean
    Dim xlApp       As Object
    Dim xlBpp       As Object
    Dim xlBook      As Object
    Dim xlSheet     As Object
    Dim ColIndex    As Integer
    Dim sExlRange   As String
    Dim sExlRange1  As String
    Dim iExlCol     As Integer
    
    
    With sPname
    
        If .MaxRows = 0 Then Exit Sub
        
        If bLkcol1 = 0 Then
           bLkcol1 = 1
        End If
        
        If bLkcol2 = 0 Then
            bLkcol2 = -1
        End If
        
        If bLkrow2 = 0 Then
            bLkrow2 = -1
        End If
        
        Clipboard.Clear
        
        .Col = bLkcol1: .Col2 = bLkcol2
        .ROW = bLkrow1: .Row2 = bLkrow2
        Clipboard.SetText .Clip
        
        'Call Excel
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
    
        xlApp.Visible = True
                        
        xlSheet.Cells.NumberFormatLocal = "G/通用格式"
        
        sExlRange1 = ""
        For ColIndex = 1 To .MaxCols
            .Col = ColIndex
            .ROW = 1

            iExlCol = ColIndex
'            If IsNumeric(.Text) And (Left(.Text, 1) = "0" Or Left(.Text, 1) = "1" Or Left(.Text, 1) = "7") And _
'               (Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14) Then
            If .CellType = SS_CELL_TYPE_EDIT Then
                If ColIndex > 104 Then
                    sExlRange1 = "D"
                    iExlCol = ColIndex - 104
                ElseIf ColIndex > 78 Then
                    sExlRange1 = "C"
                    iExlCol = ColIndex - 78
                ElseIf ColIndex > 52 Then
                    sExlRange1 = "B"
                    iExlCol = ColIndex - 52
                ElseIf ColIndex > 26 Then
                    sExlRange1 = "A"
                    iExlCol = ColIndex - 26
                End If

                sExlRange = sExlRange1 & Chr(iExlCol + 64) & "1:" & sExlRange1 & Chr(iExlCol + 64) & .MaxRows + 5
                If Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14 Then
                     xlSheet.Range(sExlRange).NumberFormat = "@"
                End If
            End If
        Next
        
       ss1.ROW = 0
       ss1.Col = SS1_PLATE_NO:             xlApp.Range("A1").Value = ss1.Text
       ss1.Col = SS1_LEN:                  xlApp.Range("B1").Value = ss1.Text
       ss1.Col = SS1_CUST_CD:              xlApp.Range("C1").Value = ss1.Text
       ss1.Col = SS1_TRIM_FL:              xlApp.Range("D1").Value = ss1.Text
       ss1.Col = SS1_SIZE_KND:             xlApp.Range("E1").Value = ss1.Text
       ss1.Col = SS1_OUT_SHEET_NO:         xlApp.Range("F1").Value = ss1.Text
       ss1.Col = SS1_UST_STATUS:           xlApp.Range("G1").Value = ss1.Text
       ss1.Col = SS1_GAS_STATUS:           xlApp.Range("H1").Value = ss1.Text
       ss1.Col = SS1_HTM_METH:             xlApp.Range("I1").Value = ss1.Text
       ss1.Col = SS1_THK_AVG:              xlApp.Range("J1").Value = ss1.Text
       ss1.Col = SS1_ORD_THK:              xlApp.Range("K1").Value = ss1.Text
       ss1.Col = SS1_ORD_WID:              xlApp.Range("L1").Value = ss1.Text
       ss1.Col = SS1_ORD_LEN:              xlApp.Range("M1").Value = ss1.Text
       ss1.Col = SS1_LEN_AVG:              xlApp.Range("N1").Value = ss1.Text
       ss1.Col = SS1_ORD_REMARK:           xlApp.Range("O1").Value = ss1.Text
       ss1.Col = SS1_STDSPEC_ORG_KND:      xlApp.Range("P1").Value = ss1.Text
       ss1.Col = SS1_STDSPEC_STLGRD:       xlApp.Range("Q1").Value = ss1.Text
       ss1.Col = SS1_TRNS_CMPY_CD:         xlApp.Range("R1").Value = ss1.Text
       ss1.Col = SS1_CUST_CD1:             xlApp.Range("S1").Value = ss1.Text
       

       
        xlApp.Range("B2").Value = "长度"
        xlApp.Range("G2").Value = "探伤"
        xlApp.Range("H2").Value = "切割"
        xlApp.Range("I2").Value = "热处理"
        xlApp.Range("J2").Value = "厚度公差"
        xlApp.Range("K2").Value = "厚度"
        xlApp.Range("L2").Value = "宽度"
        xlApp.Range("M2").Value = "长度"
        xlApp.Range("N2").Value = "长度公差"
        xlApp.Range("O2").Value = "订单备注"
        xlApp.Range("P2").Value = "标识标准"
        xlApp.Range("Q2").Value = "钢种"
'        xlApp.Range("T2").Value = "产品备注"
        
       
        Clipboard.Clear
        ss1.SetSelection SS1_PLATE_NO, 1, SS1_PLATE_NO, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("A3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_LEN, 1, SS1_LEN, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("B3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_CUST_CD, 1, SS1_SIZE_KND, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("C3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_OUT_SHEET_NO, 1, SS1_OUT_SHEET_NO, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("F3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_UST_STATUS, 1, SS1_GAS_STATUS, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("G3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_HTM_METH, 1, SS1_HTM_METH, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("I3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_THK_AVG, 1, SS1_LEN_AVG, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("J3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_ORD_REMARK, 1, SS1_STDSPEC_STLGRD, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("O3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_TRNS_CMPY_CD, 1, SS1_CUST_CD1, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("R3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
    
'        xlSheet.Range("A1").Select
'        xlSheet.Paste
'        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
            
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
        
    End With
    
    Exit Sub
    
Excel_Error:
    Call Gp_MsgBoxDisplay("您的机器尚未安装Excel", "W")

End Sub

Private Sub SSCommand2_Click()
   
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
   If ss3.MaxRows < 1 Then
   
   Call Gp_ACB5035C_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

   Else
   
   Call Gp_ACB5035C_Excel1(Me, Proc_Sc("Sc3")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
   
   End If
   
   
End Sub


Private Sub Gp_ACB5035C_Excel(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

On Error GoTo Excel_Error

    Dim ret         As Boolean
    Dim xlApp       As Object
    Dim xlBpp       As Object
    Dim xlBook      As Object
    Dim xlSheet     As Object
    Dim ColIndex    As Integer
    Dim sExlRange   As String
    Dim sExlRange1  As String
    Dim iExlCol     As Integer
    

    
    
    Const xlCenter = -4108

   Const xlNone = -4142

   Const xlAutomatic = -4105

   Const xlDiagonalDown = 5

   Const xlDiagonalUp = 6

   Const xlEdgeLeft = 7

   Const xlEdgeTop = 8

   Const xlEdgeBottom = 9

   Const xlEdgeRight = 10

   Const xlInsideVertical = 11

   Const xlInsideHorizontal = 12

   Const xlContinuous = 1

   Const xlMedium = -4138

   Const xlThick = 4

   Const xlthin = 2
    
    With sPname
    
        If .MaxRows = 0 Then Exit Sub
        
        If bLkcol1 = 0 Then
           bLkcol1 = 1
        End If
        
        If bLkcol2 = 0 Then
            bLkcol2 = -1
        End If
        
        If bLkrow2 = 0 Then
            bLkrow2 = -1
        End If
        

        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
    
        xlApp.Visible = True
        
        xlSheet.Cells.NumberFormatLocal = "@"
        
        ss1.ROW = ss1.ActiveRow
        
        xlSheet.Range("A1").Value = "生产厂"
        xlSheet.Range("A2").Value = "日期"
        xlSheet.Range("A3").Value = "产品号"
        xlSheet.Range("A4").Value = "客户"
        xlSheet.Range("A5").Value = "母板长"
        xlSheet.Range("A6").Value = "厚度"
        xlSheet.Range("A7").Value = "宽度"
        xlSheet.Range("A8").Value = "长度"
        xlSheet.Range("A9").Value = "切边"
        xlSheet.Range("A10").Value = "订单备注"
        xlSheet.Range("A11").Value = "产品备注"
        
        xlSheet.Range("A12").Value = "探伤"
        xlSheet.Range("A13").Value = "切割"
        xlSheet.Range("A14").Value = "矫直"
        xlSheet.Range("A15").Value = "热处理"
        xlSheet.Range("A16").Value = "标识标准"
        xlSheet.Range("A17").Value = "标识钢种"
        xlSheet.Range("A18").Value = "其它"
        xlSheet.Range("A19").Value = "子板数"
        xlSheet.Range("A20").Value = "子板尺寸"
        xlSheet.Range("A22").Value = "子板标准号"     '
'

        
        
        xlSheet.Range("C1").Value = "轧批号"
        xlSheet.Range("C2").Value = "班次"
        xlSheet.Range("C3").Value = "分断号"
        xlSheet.Range("C4").Value = "客户代码"
        
        xlSheet.Range("C6").Value = "厚度公差"
        xlSheet.Range("C7").Value = "宽度公差"
        xlSheet.Range("C8").Value = "长度公差"
        xlSheet.Range("C9").Value = "定尺"
        xlSheet.Range("C10").Value = "订单数量"
        xlSheet.Range("C11").Value = "客户表面要求"
        xlSheet.Range("C12").Value = "是否加喷CE"
        xlSheet.Range("C13").Value = "重量"
        xlSheet.Range("C14").Value = "加喷内容"
        xlSheet.Range("C15").Value = "侧喷加喷"
        xlSheet.Range("C16").Value = "表喷次数"
        xlSheet.Range("C17").Value = "是否钢印"
        xlSheet.Range("C18").Value = "钢印加冲"
        
        xlSheet.Range("C23").Value = "交货期"
     
        
  
        xlSheet.Range("A20", "A21").Merge
        xlSheet.Range("B20", "D21").Merge
        
'        xlSheet.Range("A1").Value = "日期"
'        xlSheet.Range("A2").Value = "产品号"
'        xlSheet.Range("A3").Value = "客户"
'        xlSheet.Range("A4").Value = "母板长"
'        xlSheet.Range("A5").Value = "厚度"
'        xlSheet.Range("A6").Value = "宽度"
'        xlSheet.Range("A7").Value = "长度"
'        xlSheet.Range("A8").Value = "切边"
'        xlSheet.Range("A9").Value = "订单备注"
'        xlSheet.Range("A10").Value = "探伤"
'        xlSheet.Range("A11").Value = "切割"
'        xlSheet.Range("A12").Value = "矫直"
'        xlSheet.Range("A13").Value = "热处理"
'        xlSheet.Range("A14").Value = "标识标准"
'        xlSheet.Range("A15").Value = "标识钢种"
'        xlSheet.Range("A16").Value = "其它"
'        xlSheet.Range("A17").Value = "子板数"
'        xlSheet.Range("A18").Value = "子板尺寸"
'
'        xlSheet.Range("A20").Value = "子板标准号"     '         新增 2015 1 27  周岩
'        xlSheet.Range("A21").Value = "交货期"     '         新增 2015 1 27  周岩
''        xlSheet.Range("C21").Value = "钢印加冲"     '         新增 2015 1 27  周岩
'
'
'
'
'        xlSheet.Range("C1").Value = "班次"
'        xlSheet.Range("C2").Value = "分断号"
'        xlSheet.Range("C3").Value = "轧批号"
'        xlSheet.Range("C4").Value = "备注"
'        xlSheet.Range("C5").Value = "厚度公差"
'        xlSheet.Range("C6").Value = "宽度公差"
'        xlSheet.Range("C7").Value = "长度公差"
'        xlSheet.Range("C8").Value = "定尺"
'        xlSheet.Range("C9").Value = "订单数量"
'        xlSheet.Range("C10").Value = "是否加喷CE"
''        xlSheet.Range("C11").Value = "是否加喷TS"
'        xlSheet.Range("C11").Value = "重量"
'        xlSheet.Range("C12").Value = "加喷内容"
'        xlSheet.Range("C13").Value = "侧喷加喷"
'        xlSheet.Range("C14").Value = "表喷次数"
'        xlSheet.Range("C15").Value = "是否钢印"
'        xlSheet.Range("C16").Value = "钢印加冲"
'        xlSheet.Range("C17").Value = "客户代码"
'
'
'        xlSheet.Range("A18", "A19").Merge
'        xlSheet.Range("B18", "D19").Merge


'        ss1.Col = SS1_MV_DATE:           xlSheet.Range("B1").Value = ss1.Text
'        ss1.Col = SS1_PLATE_NO:          xlSheet.Range("B2").Value = ss1.Text
'        ss1.Col = SS1_CUST_CD:           xlSheet.Range("B3").Value = ss1.Text
'        ss1.Col = SS1_LEN:               xlSheet.Range("B4").Value = ss1.Text
'        ss1.Col = SS1_ORD_THK:           xlSheet.Range("B5").Value = ss1.Text
'        ss1.Col = SS1_ORD_WID:           xlSheet.Range("B6").Value = ss1.Text
'        ss1.Col = SS1_ORD_LEN:           xlSheet.Range("B7").Value = ss1.Text
'        ss1.Col = SS1_TRIM_FL:           xlSheet.Range("B8").Value = ss1.Text
'        ss1.Col = SS1_ORD_REMARK:        xlSheet.Range("B9").Value = ss1.Text
'        ss1.Col = SS1_UST_STATUS:        xlSheet.Range("B10").Value = ss1.Text
'        ss1.Col = SS1_GAS_STATUS:        xlSheet.Range("B11").Value = ss1.Text
'        ss1.Col = SS1_CL_STATUS:         xlSheet.Range("B12").Value = ss1.Text
'        ss1.Col = SS1_HTM_METH:          xlSheet.Range("B13").Value = ss1.Text
'        ss1.Col = SS1_STDSPEC_ORG_KND:   xlSheet.Range("B14").Value = ss1.Text
'        ss1.Col = SS1_STDSPEC_STLGRD:    xlSheet.Range("B15").Value = ss1.Text
'        ss1.Col = SS1_QT:                xlSheet.Range("B16").Value = ss1.Text
'        ss1.Col = SS1_PLATE_CON:         xlSheet.Range("B17").Value = ss1.Text
'        ss1.Col = SS1_PLATE_SIZE:        xlSheet.Range("B18").Value = ss1.Text
'        ss1.Col = SS1_APLY_STDSPEC:      xlSheet.Range("B20").Value = ss1.Text   '新增 2015 1 27 周岩
'
'        ss1.Col = SS1_Shift:             xlSheet.Range("D1").Value = ss1.Text
'        ss1.Col = SS1_TRNS_CMPY_CD:      xlSheet.Range("D2").Value = ss1.Text
'        ss1.Col = SS1_OUT_SHEET_NO:      xlSheet.Range("D3").Value = ss1.Text
'
'        ss1.Col = SS1_TRNS_CMPY_CD:      xlSheet.Range("D4").Value = ss1.Text
'
'
'
'        ss1.Col = SS1_THK_AVG:           xlSheet.Range("D5").Value = ss1.Text
'        ss1.Col = SS1_WID_AVG:           xlSheet.Range("D6").Value = ss1.Text
'        ss1.Col = SS1_LEN_AVG:           xlSheet.Range("D7").Value = ss1.Text
'        ss1.Col = SS1_SIZE_KND:          xlSheet.Range("D8").Value = ss1.Text
'        ss1.Col = SS1_RM_CR_STAGE3_TIME: xlSheet.Range("D9").Value = ss1.Text
'        ss1.Col = SS1_CE_APPR_FL:        xlSheet.Range("D10").Value = ss1.Text
''        ss1.Col = SS1_QS_MARK_FL:        xlSheet.Range("D11").Value = ss1.Text
'        ss1.Col = SS1_WGT:               xlSheet.Range("D11").Value = ss1.Text
'        ss1.Col = SS1_VESSEL_NO:         xlSheet.Range("D12").Value = ss1.Text
'        ss1.Col = SS1_SIDEMARK:          xlSheet.Range("D13").Value = ss1.Text
'        ss1.Col = SS1_PAINTNUM:          xlSheet.Range("D14").Value = ss1.Text
'        ss1.Col = SS1_GANGYIN:           xlSheet.Range("D15").Value = ss1.Text
'        ss1.Col = SS1_PUNCH:             xlSheet.Range("D16").Value = ss1.Text
'        ss1.Col = SS1_CUST_CD1:          xlSheet.Range("D17").Value = ss1.Text


        ss3.Col = SS1_PLT:               xlSheet.Range("B1").Value = ss3.Text
        ss3.Col = SS1_MV_DATE:           xlSheet.Range("B2").Value = ss3.Text
        ss3.Col = SS1_PLATE_NO:          xlSheet.Range("B3").Value = ss3.Text
        ss3.Col = SS1_CUST_CD:           xlSheet.Range("B4").Value = ss3.Text
        ss1.Col = SS1_LEN:               xlSheet.Range("B5").Value = ss1.Text
        ss3.Col = SS1_ORD_THK:           xlSheet.Range("B6").Value = ss3.Text
        ss3.Col = SS1_ORD_WID:           xlSheet.Range("B7").Value = ss3.Text
        ss3.Col = SS1_ORD_LEN:           xlSheet.Range("B8").Value = ss3.Text
        ss3.Col = SS1_TRIM_FL:           xlSheet.Range("B9").Value = ss3.Text
        ss3.Col = SS1_ORD_REMARK:        xlSheet.Range("B10").Value = ss3.Text
        ss3.Col = SS1_PROD_REMARK:       xlSheet.Range("B11").Value = ss3.Text
        ss3.Col = SS1_UST_STATUS:        xlSheet.Range("B12").Value = ss3.Text
        ss3.Col = SS1_GAS_STATUS:        xlSheet.Range("B13").Value = ss3.Text
        ss3.Col = SS1_CL_STATUS:         xlSheet.Range("B14").Value = ss3.Text
        ss3.Col = SS1_HTM_METH:          xlSheet.Range("B15").Value = ss3.Text
        ss3.Col = SS1_STDSPEC_ORG_KND:   xlSheet.Range("B16").Value = ss3.Text
        ss3.Col = SS1_STDSPEC_STLGRD:    xlSheet.Range("B17").Value = ss3.Text
        ss3.Col = SS1_QT:                xlSheet.Range("B18").Value = ss3.Text
        ss3.Col = SS1_PLATE_CON:         xlSheet.Range("B19").Value = ss3.MaxRows     'ss3.Text
        ss3.Col = SS1_PLATE_SIZE:        xlSheet.Range("B20").Value = ss3.Text
        ss3.Col = SS1_APLY_STDSPEC:      xlSheet.Range("B22").Value = ss3.Text   '新增 2015 1 27 周岩



        ss3.Col = SS1_OUT_SHEET_NO:      xlSheet.Range("D1").Value = ss3.Text
        ss3.Col = SS1_Shift:             xlSheet.Range("D2").Value = ss3.Text
        ss3.Col = SS1_TRNS_CMPY_CD:      xlSheet.Range("D3").Value = ss3.Text
        ss3.Col = SS1_CUST_CD1:          xlSheet.Range("D4").Value = ss3.Text
        
        ss3.Col = SS1_THK_AVG:           xlSheet.Range("D6").Value = ss3.Text
        ss3.Col = SS1_WID_AVG:           xlSheet.Range("D7").Value = ss3.Text
        ss3.Col = SS1_LEN_AVG:           xlSheet.Range("D8").Value = ss3.Text
        ss3.Col = SS1_SIZE_KND:          xlSheet.Range("D9").Value = ss3.Text
        ss3.Col = SS1_RM_CR_STAGE3_TIME: xlSheet.Range("D10").Value = ss3.Text
        ss3.Col = SS1_SURFACE_REQUESTS:  xlSheet.Range("D11").Value = ss3.Text
        ss3.Col = SS1_CE_APPR_FL:        xlSheet.Range("D12").Value = ss3.Text
'
        ss3.Col = SS1_WGT:               xlSheet.Range("D13").Value = ss3.Text
        ss3.Col = SS1_VESSEL_NO:         xlSheet.Range("D14").Value = ss3.Text
        ss3.Col = SS1_SIDEMARK:          xlSheet.Range("D15").Value = ss3.Text
        ss3.Col = SS1_PAINTNUM:          xlSheet.Range("D16").Value = ss3.Text
        ss3.Col = SS1_GANGYIN:           xlSheet.Range("D17").Value = ss3.Text
        ss3.Col = SS1_PUNCH:             xlSheet.Range("D18").Value = ss3.Text
       
        ss3.Col = SS1_CUST:              xlSheet.Range("D23").Value = ss3.Text
       
'
        
'
'        xlSheet.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
'
'        xlSheet.Application.Visible = True
        
        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
        
        
        
        xlApp.Range("A1:D18").Select
        xlApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        xlApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With xlApp.Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
            
            
            ss1.ClearSelection
            Screen.MousePointer = vbDefault
        
            Set xlSheet = Nothing
            Set xlBook = Nothing
            Set xlApp = Nothing
            
        End With
        
        Exit Sub
    
Excel_Error:

    Call Gp_MsgBoxDisplay("您的机器尚未安装Excel" & Error, "W")

End Sub

Private Sub Gp_ACB5035C_Excel1(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

On Error GoTo Excel_Error

    Dim ret         As Boolean
    Dim xlApp       As Object
    Dim xlBpp       As Object
    Dim xlBook      As Object
    Dim xlSheet     As Object
    Dim ColIndex    As Integer
    Dim sExlRange   As String
    Dim sExlRange1  As String
    Dim iExlCol     As Integer
    Dim i           As Integer
    
    Const xlCenter = -4108

   Const xlNone = -4142

   Const xlAutomatic = -4105

   Const xlDiagonalDown = 5

   Const xlDiagonalUp = 6

   Const xlEdgeLeft = 7

   Const xlEdgeTop = 8

   Const xlEdgeBottom = 9

   Const xlEdgeRight = 10

   Const xlInsideVertical = 11

   Const xlInsideHorizontal = 12

   Const xlContinuous = 1

   Const xlMedium = -4138

   Const xlThick = 4

   Const xlthin = 2
   
    
    
    With sPname
    
        If .MaxRows = 0 Then Exit Sub
        
        If bLkcol1 = 0 Then
           bLkcol1 = 1
        End If
        
        If bLkcol2 = 0 Then
            bLkcol2 = -1
        End If
        
        If bLkrow2 = 0 Then
            bLkrow2 = -1
        End If
        

        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        
        For i = 1 To ss3.MaxRows
        Set xlSheet = xlBook.Worksheets(i)
    
        xlApp.Visible = True
        
        xlSheet.Cells.NumberFormatLocal = "@"
        
        ss3.ROW = i
        
        xlSheet.Range("A1").Value = "生产厂"
        xlSheet.Range("A2").Value = "日期"
        xlSheet.Range("A3").Value = "产品号"
        xlSheet.Range("A4").Value = "客户"
        xlSheet.Range("A5").Value = "母板长"
        xlSheet.Range("A6").Value = "厚度"
        xlSheet.Range("A7").Value = "宽度"
        xlSheet.Range("A8").Value = "长度"
        xlSheet.Range("A9").Value = "切边"
        xlSheet.Range("A10").Value = "订单备注"
        xlSheet.Range("A11").Value = "产品备注"
        
        xlSheet.Range("A12").Value = "探伤"
        xlSheet.Range("A13").Value = "切割"
        xlSheet.Range("A14").Value = "矫直"
        xlSheet.Range("A15").Value = "热处理"
        xlSheet.Range("A16").Value = "标识标准"
        xlSheet.Range("A17").Value = "标识钢种"
        xlSheet.Range("A18").Value = "其它"
        xlSheet.Range("A19").Value = "子板数"
        xlSheet.Range("A20").Value = "子板尺寸"
        xlSheet.Range("A22").Value = "子板标准号"     '
'

        
        
        xlSheet.Range("C1").Value = "轧批号"
        xlSheet.Range("C2").Value = "班次"
        xlSheet.Range("C3").Value = "分断号"
        xlSheet.Range("C4").Value = "客户代码"
        
        xlSheet.Range("C6").Value = "厚度公差"
        xlSheet.Range("C7").Value = "宽度公差"
        xlSheet.Range("C8").Value = "长度公差"
        xlSheet.Range("C9").Value = "定尺"
        xlSheet.Range("C10").Value = "订单数量"
        xlSheet.Range("C11").Value = "客户表面要求"
        xlSheet.Range("C12").Value = "是否加喷CE"
        xlSheet.Range("C13").Value = "重量"
        xlSheet.Range("C14").Value = "加喷内容"
        xlSheet.Range("C15").Value = "侧喷加喷"
        xlSheet.Range("C16").Value = "表喷次数"
        xlSheet.Range("C17").Value = "是否钢印"
        xlSheet.Range("C18").Value = "钢印加冲"
        
        xlSheet.Range("C23").Value = "交货期"
     
        
  
        xlSheet.Range("A20", "A21").Merge
        xlSheet.Range("B20", "D21").Merge
        
        
        ss3.Col = SS1_PLT:               xlSheet.Range("B1").Value = ss3.Text
        ss3.Col = SS1_MV_DATE:           xlSheet.Range("B2").Value = ss3.Text
        ss3.Col = SS1_PLATE_NO:          xlSheet.Range("B3").Value = ss3.Text
        ss3.Col = SS1_CUST_CD:           xlSheet.Range("B4").Value = ss3.Text
        ss1.Col = SS1_LEN:               xlSheet.Range("B5").Value = ss1.Text
        ss3.Col = SS1_ORD_THK:           xlSheet.Range("B6").Value = ss3.Text
        ss3.Col = SS1_ORD_WID:           xlSheet.Range("B7").Value = ss3.Text
        ss3.Col = SS1_ORD_LEN:           xlSheet.Range("B8").Value = ss3.Text
        ss3.Col = SS1_TRIM_FL:           xlSheet.Range("B9").Value = ss3.Text
        ss3.Col = SS1_ORD_REMARK:        xlSheet.Range("B10").Value = ss3.Text
        ss3.Col = SS1_PROD_REMARK:       xlSheet.Range("B11").Value = ss3.Text
        ss3.Col = SS1_UST_STATUS:        xlSheet.Range("B12").Value = ss3.Text
        ss3.Col = SS1_GAS_STATUS:        xlSheet.Range("B13").Value = ss3.Text
        ss3.Col = SS1_CL_STATUS:         xlSheet.Range("B14").Value = ss3.Text
        ss3.Col = SS1_HTM_METH:          xlSheet.Range("B15").Value = ss3.Text
        ss3.Col = SS1_STDSPEC_ORG_KND:   xlSheet.Range("B16").Value = ss3.Text
        ss3.Col = SS1_STDSPEC_STLGRD:    xlSheet.Range("B17").Value = ss3.Text
        ss3.Col = SS1_QT:                xlSheet.Range("B18").Value = ss3.Text
        ss3.Col = SS1_PLATE_CON:         xlSheet.Range("B19").Value = ss3.MaxRows     'ss3.Text
        ss3.Col = SS1_PLATE_SIZE:        xlSheet.Range("B20").Value = ss3.Text
        ss3.Col = SS1_APLY_STDSPEC:      xlSheet.Range("B22").Value = ss3.Text   '新增 2015 1 27 周岩



        ss3.Col = SS1_OUT_SHEET_NO:      xlSheet.Range("D1").Value = ss3.Text
        ss3.Col = SS1_Shift:             xlSheet.Range("D2").Value = ss3.Text
        ss3.Col = SS1_TRNS_CMPY_CD:      xlSheet.Range("D3").Value = ss3.Text
        ss3.Col = SS1_CUST_CD1:          xlSheet.Range("D4").Value = ss3.Text
        
        ss3.Col = SS1_THK_AVG:           xlSheet.Range("D6").Value = ss3.Text
        ss3.Col = SS1_WID_AVG:           xlSheet.Range("D7").Value = ss3.Text
        ss3.Col = SS1_LEN_AVG:           xlSheet.Range("D8").Value = ss3.Text
        ss3.Col = SS1_SIZE_KND:          xlSheet.Range("D9").Value = ss3.Text
        ss3.Col = SS1_RM_CR_STAGE3_TIME: xlSheet.Range("D10").Value = ss3.Text
        ss3.Col = SS1_SURFACE_REQUESTS:  xlSheet.Range("D11").Value = ss3.Text
        ss3.Col = SS1_CE_APPR_FL:        xlSheet.Range("D12").Value = ss3.Text
'
        ss3.Col = SS1_WGT:               xlSheet.Range("D13").Value = ss3.Text
        ss3.Col = SS1_VESSEL_NO:         xlSheet.Range("D14").Value = ss3.Text
        ss3.Col = SS1_SIDEMARK:          xlSheet.Range("D15").Value = ss3.Text
        ss3.Col = SS1_PAINTNUM:          xlSheet.Range("D16").Value = ss3.Text
        ss3.Col = SS1_GANGYIN:           xlSheet.Range("D17").Value = ss3.Text
        ss3.Col = SS1_PUNCH:             xlSheet.Range("D18").Value = ss3.Text
       
        ss3.Col = SS1_CUST:              xlSheet.Range("D23").Value = ss3.Text
       
'
        
'
'        xlSheet.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
'
'        xlSheet.Application.Visible = True
        
        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
        
        
        
        xlApp.Range("A1:D23").Select
        xlApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        xlApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With xlApp.Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
            
            
            ss3.ClearSelection
            Screen.MousePointer = vbDefault
            
        Next i
        
            Set xlSheet = Nothing
            Set xlBook = Nothing
            Set xlApp = Nothing
            
        End With
        
        Exit Sub
    
Excel_Error:

    Call Gp_MsgBoxDisplay("您的机器尚未安装Excel" & Error, "W")

End Sub


