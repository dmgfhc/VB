VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACA1021C 
   Caption         =   "合同兑现查询_ACA1021C"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   16219
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACA1021C.frx":0000
      Begin VB.TextBox txt_ord_no_find 
         Height          =   1710
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   15210
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1710
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   3016
         _Version        =   196609
         BackColor       =   14737632
         ShadowStyle     =   1
         Begin VB.TextBox txt_stlgrd_dec 
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
            Left            =   8610
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   13
            Top             =   885
            Width           =   1770
         End
         Begin VB.TextBox TXT_UST_FL 
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
            Height          =   310
            Left            =   3660
            MaxLength       =   4
            TabIndex        =   12
            Tag             =   "钢种"
            Top             =   885
            Width           =   570
         End
         Begin VB.TextBox TXT_ORD_KND 
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
            Left            =   1410
            MaxLength       =   1
            TabIndex        =   11
            Top             =   885
            Width           =   645
         End
         Begin VB.TextBox TXT_REC_STS 
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
            Left            =   1410
            MaxLength       =   1
            TabIndex        =   10
            Top             =   495
            Width           =   645
         End
         Begin VB.TextBox TXT_HTM_METH 
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
            Height          =   310
            Left            =   3660
            MaxLength       =   1
            TabIndex        =   9
            Top             =   1260
            Width           =   570
         End
         Begin VB.ComboBox CBO_ORD_ITEM 
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
            Left            =   5010
            TabIndex        =   8
            Top             =   495
            Width           =   660
         End
         Begin VB.TextBox TXT_ORD_STS 
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
            Left            =   1410
            MaxLength       =   1
            TabIndex        =   7
            Top             =   1260
            Width           =   645
         End
         Begin VB.TextBox TXT_STLGRD 
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
            Left            =   7290
            MaxLength       =   11
            TabIndex        =   6
            Tag             =   "钢种"
            Top             =   885
            Width           =   1320
         End
         Begin VB.TextBox TXT_ORD_NO 
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
            Left            =   3660
            MaxLength       =   11
            TabIndex        =   5
            Top             =   495
            Width           =   1350
         End
         Begin VB.TextBox TXT_PROD_CD 
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
            Left            =   1410
            MaxLength       =   2
            TabIndex        =   4
            Tag             =   "产品"
            Top             =   120
            Width           =   645
         End
         Begin VB.TextBox TXT_STDSPEC 
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
            Left            =   7290
            MaxLength       =   20
            TabIndex        =   3
            Tag             =   "钢种(标准号)"
            Top             =   1260
            Width           =   3090
         End
         Begin VB.TextBox TXT_CFM_MILL_PLT 
            Alignment       =   2  'Center
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
            Left            =   3660
            MaxLength       =   2
            TabIndex        =   2
            Top             =   120
            Width           =   540
         End
         Begin Threed.SSCommand cmd_ord 
            Height          =   345
            Left            =   10740
            TabIndex        =   14
            Top             =   1260
            Visible         =   0   'False
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   609
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   12583104
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "订单兑现统计"
         End
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   120
            Top             =   120
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
            Left            =   6000
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "用户交货期"
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
         Begin InDate.UDate TXT_DEL_TO_DATE 
            Height          =   315
            Left            =   8940
            TabIndex        =   15
            Tag             =   "交货期"
            Top             =   120
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
            BackColor       =   16777215
            MaxLength       =   10
         End
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   2370
            Top             =   495
            Width           =   1260
            _ExtentX        =   2223
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
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   120
            Top             =   1260
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "订单状态"
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
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Left            =   10740
            Top             =   120
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            Caption         =   "订单厚度"
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
         Begin InDate.ULabel ULabel13 
            Height          =   315
            Left            =   6000
            Top             =   885
            Width           =   1260
            _ExtentX        =   2223
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
         Begin InDate.ULabel ULabel01 
            Height          =   315
            Index           =   14
            Left            =   2370
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "投入工厂"
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
         Begin InDate.ULabel ULabel15 
            Height          =   315
            Left            =   2370
            Top             =   1260
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "热处理"
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
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   120
            Top             =   495
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "信息状态"
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
         Begin InDate.UDate TXT_DEL_FR_DATE 
            Height          =   315
            Left            =   7290
            TabIndex        =   16
            Tag             =   "交货期"
            Top             =   120
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
            BackColor       =   16777215
            MaxLength       =   10
         End
         Begin CSTextLibCtl.sidbEdit SDB_ORD_LEN_FR 
            Height          =   315
            Left            =   11880
            TabIndex        =   17
            Tag             =   "订单长度"
            Top             =   885
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
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
         Begin CSTextLibCtl.sidbEdit SDB_ORD_WID_FR 
            Height          =   315
            Left            =   11880
            TabIndex        =   18
            Tag             =   "订单宽度"
            Top             =   495
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
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
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   120
            Top             =   885
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "订单种类"
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
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   2370
            Top             =   885
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "探伤"
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
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   6000
            Top             =   495
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "确认日期"
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
         Begin InDate.UDate TXT_CONF_DATE_TO 
            Height          =   315
            Left            =   8940
            TabIndex        =   19
            Tag             =   "交货期"
            Top             =   495
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
            BackColor       =   16777215
            MaxLength       =   10
         End
         Begin InDate.UDate TXT_CONF_DATE_FR 
            Height          =   315
            Left            =   7290
            TabIndex        =   20
            Tag             =   "交货期"
            Top             =   495
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
            BackColor       =   16777215
            MaxLength       =   10
         End
         Begin CSTextLibCtl.sidbEdit SDB_ORD_THK_FR 
            Height          =   315
            Left            =   11880
            TabIndex        =   21
            Tag             =   "订单厚度"
            Top             =   120
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
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
         Begin CSTextLibCtl.sidbEdit SDB_ORD_THK_TO 
            Height          =   315
            Left            =   13110
            TabIndex        =   22
            Tag             =   "订单厚度"
            Top             =   120
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
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
         Begin InDate.ULabel ULabel17 
            Height          =   315
            Left            =   10740
            Top             =   495
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            Caption         =   "订单宽度"
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
         Begin CSTextLibCtl.sidbEdit SDB_ORD_WID_TO 
            Height          =   315
            Left            =   13110
            TabIndex        =   23
            Tag             =   "订单宽度"
            Top             =   495
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
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
         Begin InDate.ULabel ULabel18 
            Height          =   315
            Left            =   10740
            Top             =   885
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            Caption         =   "订单长度"
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
         Begin CSTextLibCtl.sidbEdit SDB_ORD_LEN_TO 
            Height          =   315
            Left            =   13110
            TabIndex        =   24
            Tag             =   "订单长度"
            Top             =   885
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
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
         Begin InDate.ULabel ULabel19 
            Height          =   315
            Left            =   6000
            Top             =   1260
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "标准号"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16711680
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            Height          =   120
            Left            =   8790
            TabIndex        =   29
            Top             =   615
            Width           =   90
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            Height          =   120
            Left            =   8790
            TabIndex        =   28
            Top             =   210
            Width           =   90
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            Height          =   75
            Left            =   12960
            TabIndex        =   27
            Top             =   240
            Width           =   90
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            Height          =   75
            Left            =   12960
            TabIndex        =   26
            Top             =   615
            Width           =   90
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            Height          =   75
            Left            =   12960
            TabIndex        =   25
            Top             =   1005
            Width           =   90
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   3660
         Left            =   0
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1770
         Width           =   15210
         _Version        =   393216
         _ExtentX        =   26829
         _ExtentY        =   6456
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
         MaxCols         =   55
         MaxRows         =   2
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "ACA1021C.frx":0072
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   3705
         Left            =   0
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   5520
         Width           =   15210
         _Version        =   393216
         _ExtentX        =   26829
         _ExtentY        =   6535
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
         MaxCols         =   61
         MaxRows         =   2
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "ACA1021C.frx":1F24
      End
   End
End
Attribute VB_Name = "ACA1021C"
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
'-- Program Name
'-- Program ID
'-- Designer
'-- Coder
'-- Date
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
Public QueryYN As Boolean

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

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

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
    Dim iRow As Integer
    Dim iRow_1 As Integer

'    Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_cfm_mill_plt, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_prod_cd, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_REC_STS, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ord_knd, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ord_sts, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_UST_FL, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_HTM_METH, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_STDSPEC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_del_fr_date, "p", "n", "m", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_del_to_date, "p", "n", "m", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_CONF_DATE_FR, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_CONF_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_ORD_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(CBO_ORD_ITEM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_ORD_THK_FR, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_ORD_THK_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_ORD_WID_FR, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_ORD_WID_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_ORD_LEN_FR, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_ORD_LEN_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

       
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  BELOW EDIT ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
   
  
   For iRow = 1 To 55
         Call Gp_Sp_Collection(ss1, iRow, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iRow
    
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACA1021C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    
    
    '    先注册1  查询明细使用
'    Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_ord_no_find, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
       
    'MASTER Collection
    Mc2.Add Item:=pContro2, Key:="pControl"
    Mc2.Add Item:=nContro2, Key:="nControl"
    Mc2.Add Item:=mContro2, Key:="mControl"
    Mc2.Add Item:=iContro2, Key:="iControl"
    Mc2.Add Item:=rContro2, Key:="rControl"
    Mc2.Add Item:=cContro2, Key:="cControl"
    Mc2.Add Item:=aContro2, Key:="aControl"
    Mc2.Add Item:=lContro2, Key:="lControl"
    
    
     ' control part   Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    
    For iRow_1 = 1 To 61
        Call Gp_Sp_Collection(ss2, iRow, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next iRow_1
   
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ACA1021C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=2, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub cmd_ord_Click()
On Error GoTo cmd_ord_Error

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
   
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
              
    sQuery = "{call ACA1021P('" + txt_cfm_mill_plt.Text + "','" + txt_prod_cd.Text + "','" + TXT_REC_STS.Text + "','" + txt_ord_knd.Text + "','" + _
                                  txt_ord_sts.Text + "','" + TXT_UST_FL.Text + "','" + TXT_HTM_METH.Text + "','" + txt_stlgrd.Text + "','" + TXT_STDSPEC.Text + "','" + _
                                  txt_del_fr_date.RawData + "','" + txt_del_to_date.RawData + "','" + TXT_CONF_DATE_FR.RawData + "','" + TXT_CONF_DATE_TO.RawData + "','" + TXT_ORD_NO.Text + "','" + _
                                  CBO_ORD_ITEM.Text + "','" + SDB_ORD_THK_FR.Text + "','" + SDB_ORD_THK_TO.Text + "','" + SDB_ORD_WID_FR.Text + "','" + SDB_ORD_WID_TO.Text + "','" + _
                                  SDB_ORD_LEN_FR.Text + "','" + SDB_ORD_LEN_TO.Text + "','" + sUserID + "',?)}"
      
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
    Else
   
        Call MsgBox(cmd_ord.Caption + "完成！", vbInformation, "系统提示信息")
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

cmd_ord_Error:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("订单兑现统计错误: " & Error)

End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("sc")("Spread"), False)
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("sc")("Spread"))
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc2")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("sc"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    
    Call Gp_Sp_ColGet(Proc_Sc("sc")("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "C-System.INI", Me.Name)
    
   
    If Gf_Sc_Authority(sAuthority, "U") Then
       cmd_ord.Visible = True
    Else
       cmd_ord.Visible = False
    End If
        
    Screen.MousePointer = vbDefault

    txt_prod_cd.Text = "PP"
    
    txt_del_fr_date.Text = DateAdd("m", -1, Date)
    txt_del_to_date.Text = DateAdd("m", 1, Date)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "AC-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "AC-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
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

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("pControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call Gf_Sp_Cls(Proc_Sc("SC2"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        
    End If

End Sub

Public Sub Form_Ref()
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss1.OperationMode = OperationModeNormal
        Call Gp_Sp_EvenRowBackcolor(ss1)
    End If
    Exit Sub
Refer_Err:

End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
        
End Sub

Public Sub Spread_Forzens_Setting()

    ss1.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

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

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

   Dim iRow As Long
   Dim iTxt As String
   Dim iOrd_no As String
   
   iRow = Row
   
   Call Gf_Sp_Cls(Proc_Sc("Sc2"))
   
   ss1.Row = iRow:  ss1.Col = 1:  iOrd_no = ss1.Text
   txt_ord_no_find.Text = iOrd_no
   
   If iOrd_no = "" Then
      Exit Sub
   End If
   
   If iRow > 0 Then
      
      
      If Gf_Sp_Refer(M_CN1, sc2, Mc2) Then
            ss2.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
      End If
   End If
   
End Sub


Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
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

Private Sub txt_prod_cd_Change()

   Select Case txt_prod_cd.Text
          Case "S", "s", "SL"
             txt_prod_cd.Text = "SL"
          Case "P", "p", "PP"
             txt_prod_cd.Text = "PP"
          Case "H", "h", "HC"
             txt_prod_cd.Text = "HC"
          Case "", "**"
             txt_prod_cd.Text = ""
          Case Else
             txt_prod_cd.Text = ""
             Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
    End Select

End Sub

Private Sub txt_prod_cd_DblClick()

    Call txt_prod_cd_KeyUp(vbKeyF4, 0)

End Sub
Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=txt_prod_cd
   
        DD.nameType = "2"
        'DD.nameType="1" 按中文名称查询
        'DD.nameType="2" 按英文名称查询
       
        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub
        
    End If
    
End Sub

Private Sub TXT_CFM_MILL_PLT_DblClick()

    Call txt_cfm_mill_plt_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_rec_sts_DblClick()

    Call txt_rec_sts_KeyUp(vbKeyF4, 0)
 
End Sub

Private Sub txt_rec_sts_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "Z0005"
        DD.rControl.Add Item:=TXT_REC_STS
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
    End If

End Sub

Private Sub txt_cfm_mill_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_cfm_mill_plt
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

End Sub

Private Sub TXT_HTM_METH_DblClick()

    Call TXT_HTM_METH_KeyUp(vbKeyF4, 0)

End Sub

Private Sub TXT_HTM_METH_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "Q0073"
        DD.rControl.Add Item:=TXT_HTM_METH
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If

End Sub

Private Sub txt_ord_knd_DblClick()

    Call txt_ord_knd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_ord_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "B0009"
        DD.rControl.Add Item:=txt_ord_knd
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    End If

End Sub

Private Sub txt_ord_sts_DblClick()

    Call txt_ord_sts_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_ord_sts_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "B0011"
        DD.rControl.Add Item:=txt_ord_sts
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If

End Sub

Private Sub txt_stdspec_DblClick()

   Call txt_STDSPEC_KeyUp(vbKeyF4, 0)
 
End Sub

Private Sub txt_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_STDSPEC
        If txt_prod_cd.Text = "SL" Then
            DD.nameType = "1"
            Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        Else
            Call Gf_StdSPEC_DD2(M_CN1, KeyCode)
        End If
    End If

End Sub

Private Sub txt_stlgrd_DblClick()

   Call txt_stlgrd_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stlgrd
        DD.rControl.Add Item:=txt_stlgrd_dec
        Call Gf_Stlgrd_DD(M_CN1, vbKeyF4)
        Exit Sub
    End If
    
    If Len(Trim(txt_stlgrd)) = txt_stlgrd.MaxLength Then
        txt_stlgrd_dec.Text = Gf_CustNameFind(M_CN1, Trim(txt_stlgrd.Text), 1)
    Else
        txt_stlgrd_dec.Text = ""
    End If
    
End Sub

Private Sub TXT_UST_FL_DblClick()

   Call TXT_UST_FL_KeyUp(vbKeyF4, 0)

End Sub

Private Sub TXT_UST_FL_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "Q0046"
        DD.rControl.Add Item:=TXT_UST_FL
        DD.nameType = "2"
        Call Gf_Mill_Common_DD(M_CN1, vbKeyF4)
    End If

End Sub
