VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AEB1090C 
   Caption         =   "板坯长度设计_AEB1090C"
   ClientHeight    =   9225
   ClientLeft      =   315
   ClientTop       =   2010
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxT_stdgrd 
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
      Left            =   5325
      MaxLength       =   11
      TabIndex        =   76
      Top             =   525
      Width           =   1335
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
      Left            =   8310
      MaxLength       =   2
      TabIndex        =   75
      Tag             =   "产品"
      Top             =   120
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
      Left            =   8775
      MaxLength       =   40
      TabIndex        =   74
      Tag             =   "产品"
      Top             =   120
      Width           =   1620
   End
   Begin VB.TextBox txt_prc_line 
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
      Left            =   4860
      MaxLength       =   1
      TabIndex        =   72
      Tag             =   "转炉"
      Top             =   120
      Width           =   465
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
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   71
      Tag             =   "工厂"
      Top             =   120
      Width           =   465
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
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   70
      Tag             =   "工厂"
      Top             =   120
      Width           =   1125
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
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   69
      Tag             =   "产品"
      Top             =   525
      Width           =   1545
   End
   Begin VB.TextBox txt_ord_item 
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
      Left            =   2865
      MaxLength       =   2
      TabIndex        =   68
      Tag             =   "产品"
      Top             =   525
      Width           =   435
   End
   Begin VB.TextBox txt_stlgrd_grp 
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
      Left            =   4860
      MaxLength       =   11
      TabIndex        =   67
      Tag             =   "钢种组"
      Top             =   525
      Width           =   465
   End
   Begin VB.TextBox txt_ccm_line 
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
      Left            =   5325
      MaxLength       =   1
      TabIndex        =   66
      Tag             =   "连浇机号"
      Top             =   120
      Width           =   465
   End
   Begin VB.TextBox txt_stdspec 
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
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   60
      Top             =   900
      Width           =   1980
   End
   Begin VB.TextBox txt_cust_cd 
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
      Left            =   4860
      MaxLength       =   11
      TabIndex        =   59
      Tag             =   "产品"
      Top             =   900
      Width           =   1035
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   870
      Left            =   75
      TabIndex        =   0
      Top             =   8400
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   1535
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   630
         Top             =   90
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "轧件厚度"
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   6465
         Top             =   90
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "轧件长度"
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
      Begin CSTextLibCtl.sidbEdit sdb_slab_len 
         Height          =   315
         Left            =   7860
         TabIndex        =   1
         Top             =   90
         Width           =   1410
         _Version        =   262145
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_slab_thk 
         Height          =   315
         Left            =   2010
         TabIndex        =   2
         Top             =   90
         Width           =   1140
         _Version        =   262145
         _ExtentX        =   2011
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Left            =   3540
         Top             =   90
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "轧件宽度"
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
      Begin CSTextLibCtl.sidbEdit sdb_slab_wid 
         Height          =   315
         Left            =   4920
         TabIndex        =   3
         Top             =   90
         Width           =   1140
         _Version        =   262145
         _ExtentX        =   2011
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   630
         Top             =   495
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "板坯厚度"
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
         Left            =   6465
         Top             =   495
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "板坯长度"
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
      Begin CSTextLibCtl.sidbEdit sdb_slab_len1 
         Height          =   315
         Left            =   7860
         TabIndex        =   4
         Top             =   495
         Width           =   1410
         _Version        =   262145
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.0"
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
         NumDecDigits    =   1
         NumIntDigits    =   7
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_slab_thk1 
         Height          =   315
         Left            =   2010
         TabIndex        =   5
         Top             =   495
         Width           =   1140
         _Version        =   262145
         _ExtentX        =   2011
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.00"
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
         NumDecDigits    =   2
         NumIntDigits    =   4
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   3540
         Top             =   495
         Width           =   1350
         _ExtentX        =   2381
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
      Begin CSTextLibCtl.sidbEdit sdb_slab_wid1 
         Height          =   315
         Left            =   4920
         TabIndex        =   6
         Top             =   495
         Width           =   1140
         _Version        =   262145
         _ExtentX        =   2011
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.00"
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
         NumDecDigits    =   2
         NumIntDigits    =   4
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   9675
         Top             =   495
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "板坯重量"
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
      Begin CSTextLibCtl.sidbEdit sdb_slab_wgt1 
         Height          =   315
         Left            =   11070
         TabIndex        =   7
         Top             =   495
         Width           =   1410
         _Version        =   262145
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         NumIntDigits    =   7
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel29 
         Height          =   315
         Left            =   9675
         Top             =   90
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "成材率"
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
      Begin CSTextLibCtl.sidbEdit sdb_slab_ratio 
         Height          =   315
         Left            =   11070
         TabIndex        =   8
         Top             =   90
         Width           =   1410
         _Version        =   262145
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
         Undo            =   0
         Data            =   0
      End
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   1275
      Left            =   75
      TabIndex        =   9
      Top             =   7200
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   2249
      _Version        =   196609
      BackColor       =   12640511
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   45
         Top             =   45
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Caption         =   "轧件"
         Alignment       =   1
         BackColor       =   12640511
         BackgroundStyle =   1
         BorderEffect    =   0
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
         ForeColor       =   255
      End
      Begin Threed.SSCommand cmd_slab_init 
         Height          =   375
         Left            =   13140
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "初始化"
         BevelWidth      =   3
      End
      Begin Threed.SSCommand cmd_slab_design 
         Height          =   375
         Left            =   13140
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   45
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "设计探讨"
         BevelWidth      =   3
      End
      Begin InDate.ULabel lbl_slab 
         Height          =   150
         Index           =   0
         Left            =   675
         Top             =   180
         Visible         =   0   'False
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   265
         Caption         =   ""
         Alignment       =   1
         BackColor       =   8421631
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin Threed.SSCommand cmd_slab_del 
         Height          =   375
         Left            =   14115
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   45
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   32896
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "删除"
         BevelWidth      =   3
      End
      Begin Threed.SSCommand cmd_slab_complete 
         Height          =   375
         Left            =   14115
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   450
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   12583104
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "设计确定"
         BevelWidth      =   3
      End
      Begin Threed.SSCommand cmd_design_modify 
         Height          =   375
         Left            =   14115
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   855
         Visible         =   0   'False
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "设计调整"
         BevelWidth      =   3
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "500(M)"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   13140
         TabIndex        =   15
         Top             =   990
         Width           =   510
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000000&
         Height          =   1095
         Left            =   630
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   12435
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   510
      Left            =   75
      TabIndex        =   16
      Top             =   6645
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   900
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   630
         Top             =   90
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "母板厚度"
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
         Left            =   6465
         Top             =   90
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "母板长度"
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
      Begin CSTextLibCtl.sidbEdit sdb_asroll_len 
         Height          =   315
         Left            =   7860
         TabIndex        =   17
         Top             =   90
         Width           =   1410
         _Version        =   262145
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_asroll_thk 
         Height          =   315
         Left            =   2010
         TabIndex        =   18
         Top             =   90
         Width           =   1140
         _Version        =   262145
         _ExtentX        =   2011
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   3540
         Top             =   90
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "母板宽度"
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
      Begin CSTextLibCtl.sidbEdit sdb_asroll_wid 
         Height          =   315
         Left            =   4920
         TabIndex        =   19
         Top             =   90
         Width           =   1140
         _Version        =   262145
         _ExtentX        =   2011
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   9675
         Top             =   90
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "探讨母板长度"
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
      Begin CSTextLibCtl.sidbEdit sdb_plate_len 
         Height          =   315
         Left            =   11070
         TabIndex        =   20
         Top             =   90
         Width           =   1410
         _Version        =   262145
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
         Undo            =   0
         Data            =   0
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   840
      Left            =   75
      TabIndex        =   21
      Top             =   4440
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   1482
      _Version        =   196609
      BackColor       =   14737918
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_ord_no6 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   310
         Left            =   10935
         MaxLength       =   14
         TabIndex        =   51
         Top             =   435
         Width           =   1650
      End
      Begin VB.TextBox txt_ord_no5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   310
         Left            =   5895
         MaxLength       =   14
         TabIndex        =   50
         Top             =   435
         Width           =   1650
      End
      Begin VB.TextBox txt_ord_no4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   310
         Left            =   900
         MaxLength       =   14
         TabIndex        =   49
         Top             =   435
         Width           =   1650
      End
      Begin VB.TextBox txt_ord_no1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   310
         Left            =   900
         MaxLength       =   14
         TabIndex        =   24
         Top             =   90
         Width           =   1650
      End
      Begin VB.TextBox txt_ord_no3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   310
         Left            =   10935
         MaxLength       =   14
         TabIndex        =   23
         Top             =   90
         Width           =   1650
      End
      Begin VB.TextBox txt_ord_no2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   310
         Left            =   5895
         MaxLength       =   14
         TabIndex        =   22
         Top             =   90
         Width           =   1650
      End
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   90
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
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
      End
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   5085
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
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
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   10125
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
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
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord11_cnt 
         Height          =   315
         Left            =   3420
         TabIndex        =   25
         Top             =   90
         Width           =   435
         _Version        =   262145
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Left            =   2610
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         Caption         =   "张数"
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
      Begin CSTextLibCtl.sidbEdit sdb_ord21_cnt 
         Height          =   315
         Left            =   8415
         TabIndex        =   26
         Top             =   90
         Width           =   435
         _Version        =   262145
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   7605
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         Caption         =   "张数"
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
      Begin CSTextLibCtl.sidbEdit sdb_ord31_cnt 
         Height          =   315
         Left            =   13455
         TabIndex        =   27
         Top             =   90
         Width           =   435
         _Version        =   262145
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   12645
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         Caption         =   "张数"
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
      Begin Threed.SSCommand cmd_ord1 
         Height          =   330
         Left            =   4365
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   90
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "适用"
         BevelWidth      =   3
      End
      Begin Threed.SSCommand cmd_ord2 
         Height          =   330
         Left            =   9360
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   90
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "适用"
         BevelWidth      =   3
      End
      Begin Threed.SSCommand cmd_ord3 
         Height          =   330
         Left            =   14400
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   90
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "适用"
         BevelWidth      =   3
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord32_cnt 
         Height          =   315
         Left            =   13905
         TabIndex        =   31
         Top             =   90
         Width           =   435
         _Version        =   262145
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord22_cnt 
         Height          =   315
         Left            =   8865
         TabIndex        =   32
         Top             =   90
         Width           =   435
         _Version        =   262145
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord12_cnt 
         Height          =   315
         Left            =   3870
         TabIndex        =   33
         Top             =   90
         Width           =   435
         _Version        =   262145
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord2_len 
         Height          =   315
         Left            =   5895
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   1410
         _Version        =   262145
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord3_len 
         Height          =   315
         Left            =   10935
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   1410
         _Version        =   262145
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord1_len 
         Height          =   315
         Left            =   900
         TabIndex        =   36
         Top             =   30
         Visible         =   0   'False
         Width           =   1410
         _Version        =   262145
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   5085
         Top             =   435
         Width           =   765
         _ExtentX        =   1349
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
      End
      Begin InDate.ULabel ULabel24 
         Height          =   315
         Left            =   10125
         Top             =   435
         Width           =   765
         _ExtentX        =   1349
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
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord41_cnt 
         Height          =   315
         Left            =   3420
         TabIndex        =   37
         Top             =   435
         Width           =   435
         _Version        =   262145
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel25 
         Height          =   315
         Left            =   2610
         Top             =   435
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         Caption         =   "张数"
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
      Begin CSTextLibCtl.sidbEdit sdb_ord51_cnt 
         Height          =   315
         Left            =   8415
         TabIndex        =   38
         Top             =   435
         Width           =   435
         _Version        =   262145
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel26 
         Height          =   315
         Left            =   7605
         Top             =   435
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         Caption         =   "张数"
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
      Begin CSTextLibCtl.sidbEdit sdb_ord61_cnt 
         Height          =   315
         Left            =   13455
         TabIndex        =   39
         Top             =   435
         Width           =   435
         _Version        =   262145
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   315
         Left            =   12645
         Top             =   435
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         Caption         =   "张数"
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
      Begin Threed.SSCommand cmd_ord4 
         Height          =   330
         Left            =   4365
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   435
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "适用"
         BevelWidth      =   3
      End
      Begin Threed.SSCommand cmd_ord5 
         Height          =   330
         Left            =   9360
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   435
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "适用"
         BevelWidth      =   3
      End
      Begin Threed.SSCommand cmd_ord6 
         Height          =   330
         Left            =   14400
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   435
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "适用"
         BevelWidth      =   3
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord62_cnt 
         Height          =   315
         Left            =   13905
         TabIndex        =   43
         Top             =   435
         Width           =   435
         _Version        =   262145
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord52_cnt 
         Height          =   315
         Left            =   8865
         TabIndex        =   44
         Top             =   435
         Width           =   435
         _Version        =   262145
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord42_cnt 
         Height          =   315
         Left            =   3870
         TabIndex        =   45
         Top             =   435
         Width           =   435
         _Version        =   262145
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         NumIntDigits    =   3
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel28 
         Height          =   315
         Left            =   90
         Top             =   435
         Width           =   765
         _ExtentX        =   1349
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
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord6_len 
         Height          =   315
         Left            =   10935
         TabIndex        =   46
         Top             =   345
         Visible         =   0   'False
         Width           =   1410
         _Version        =   262145
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord5_len 
         Height          =   315
         Left            =   5895
         TabIndex        =   47
         Top             =   345
         Visible         =   0   'False
         Width           =   1410
         _Version        =   262145
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_ord4_len 
         Height          =   315
         Left            =   900
         TabIndex        =   48
         Top             =   375
         Visible         =   0   'False
         Width           =   1410
         _Version        =   262145
         _ExtentX        =   2487
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Modified        =   -1  'True
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
         Undo            =   0
         Data            =   0
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   1275
      Left            =   75
      TabIndex        =   52
      Top             =   5340
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   2249
      _Version        =   196609
      BackColor       =   16761024
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin InDate.ULabel lbl_mplate 
         Height          =   150
         Index           =   0
         Left            =   705
         Top             =   195
         Visible         =   0   'False
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   265
         Caption         =   ""
         Alignment       =   1
         BackColor       =   12632256
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         Left            =   45
         Top             =   45
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Caption         =   "母板"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
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
         ForeColor       =   8421631
      End
      Begin Threed.SSCommand cmd_mplate_init 
         Height          =   375
         Left            =   13140
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   495
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "初始化"
         BevelWidth      =   3
      End
      Begin Threed.SSCommand cmd_mplate_design 
         Height          =   375
         Left            =   13140
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   90
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "设计探讨"
         BevelWidth      =   3
      End
      Begin Threed.SSCommand cmd_mplate_del 
         Height          =   375
         Left            =   14115
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   90
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   32896
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "删除"
         BevelWidth      =   3
      End
      Begin Threed.SSCommand cmd_mplate_complete 
         Height          =   375
         Left            =   14115
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   495
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   12583104
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "设计确定"
         BevelWidth      =   3
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         Height          =   1095
         Left            =   630
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   12435
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "50(M)"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   13140
         TabIndex        =   57
         Top             =   990
         Width           =   960
      End
   End
   Begin FPSpread.vaSpread SS1 
      Height          =   2760
      Left            =   75
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   1680
      Width           =   15105
      _Version        =   393216
      _ExtentX        =   26644
      _ExtentY        =   4868
      _StockProps     =   64
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
      MaxCols         =   35
      MaxRows         =   2
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "AEB1090C.frx":0000
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1035
      Left            =   11295
      TabIndex        =   61
      Top             =   120
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   1826
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   130
         Top             =   150
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "板坯宽度"
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
      Begin CSTextLibCtl.sidbEdit TXT_SlaB_WIDTH_FROM 
         Height          =   315
         Left            =   1455
         TabIndex        =   62
         Tag             =   "板坯宽度FROM"
         Top             =   150
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
      Begin CSTextLibCtl.sidbEdit TXT_SLAB_WIDTH_TO 
         Height          =   315
         Left            =   2475
         TabIndex        =   63
         Tag             =   "板坯宽度TO"
         Top             =   150
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   130
         Top             =   570
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "板坯变成宽度"
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
      Begin CSTextLibCtl.sidbEdit TXT_SLAB_WIDTH_TAG 
         Height          =   315
         Left            =   1455
         TabIndex        =   64
         Tag             =   "板坯变成宽度"
         Top             =   570
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
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
      Begin Threed.SSCommand cmd_Wid_Modify 
         Height          =   420
         Left            =   2640
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   525
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   12583104
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "宽度变更"
         BevelWidth      =   3
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Left            =   12135
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   741
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
      Caption         =   "订单编制"
      BevelWidth      =   3
   End
   Begin InDate.UDate txt_del_fr 
      Height          =   315
      Left            =   8310
      TabIndex        =   77
      Top             =   510
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
   End
   Begin CSTextLibCtl.sidbEdit txt_prod_thk_from 
      Height          =   315
      Left            =   1320
      TabIndex        =   78
      Top             =   1290
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   120
      Top             =   1290
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "产品厚度"
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
   Begin InDate.ULabel ULabel30 
      Height          =   315
      Left            =   3570
      Top             =   1290
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "产品宽度"
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
   Begin InDate.ULabel ULabel31 
      Height          =   315
      Left            =   7020
      Top             =   1290
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "产品长度"
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
   Begin InDate.ULabel ULabel32 
      Height          =   315
      Left            =   7020
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
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel33 
      Height          =   315
      Left            =   7020
      Top             =   510
      Width           =   1260
      _ExtentX        =   2223
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
   Begin InDate.ULabel ULabel34 
      Height          =   315
      Left            =   3570
      Top             =   525
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "钢种组/钢种"
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
   Begin CSTextLibCtl.sidbEdit txt_prod_thk_to 
      Height          =   315
      Left            =   2310
      TabIndex        =   79
      Top             =   1290
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
   Begin CSTextLibCtl.sidbEdit txt_prod_len_from 
      Height          =   315
      Left            =   8310
      TabIndex        =   80
      Top             =   1290
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
      FocusSelect     =   -1  'True
      Modified        =   -1  'True
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
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_prod_len_to 
      Height          =   315
      Left            =   9585
      TabIndex        =   81
      Top             =   1290
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
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_prod_wid_from 
      Height          =   315
      Left            =   4860
      TabIndex        =   82
      Top             =   1290
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
   Begin CSTextLibCtl.sidbEdit txt_prod_wid_to 
      Height          =   315
      Left            =   5835
      TabIndex        =   83
      Top             =   1290
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
   Begin InDate.ULabel ULabel35 
      Height          =   225
      Left            =   10515
      Top             =   1350
      Visible         =   0   'False
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   397
      Caption         =   "紧急订单"
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
   Begin InDate.ULabel ULabel36 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "工厂/机号"
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
   Begin InDate.ULabel ULabel37 
      Height          =   315
      Left            =   120
      Top             =   525
      Width           =   1170
      _ExtentX        =   2064
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
   Begin InDate.UDate txt_del_to 
      Height          =   315
      Left            =   9750
      TabIndex        =   84
      Top             =   510
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
   End
   Begin InDate.ULabel ULabel38 
      Height          =   315
      Left            =   3570
      Top             =   120
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "转炉/连浇"
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
   Begin InDate.ULabel ULabel39 
      Height          =   315
      Left            =   120
      Top             =   900
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "标准号"
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
   Begin InDate.ULabel ULabel40 
      Height          =   315
      Left            =   3570
      Top             =   900
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "客户代码"
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
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   30
      X2              =   15105
      Y1              =   7185
      Y2              =   7185
   End
End
Attribute VB_Name = "AEB1090C"
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
'-- Program ID        AEB1090C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Caolei
'-- Coder             Caolei
'-- Date              2013.11.21
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

Public Active_CForm As String       'Form Active

Public Complete As Boolean          'Plate Delete Setting

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

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim oRd_cnt As Integer              'Select Order Count
Dim iMplate_cnt As Integer          'Mplate Design Count
Dim iSlab_cnt As Integer            'Slab Design Count
Dim iLastSlab_cnt As Integer        'Last Slab Complte Count
Dim iSlab_Complete As Integer       'Slab Complete Count
Dim iOrd1_Curr_Row As Integer       'Select Order1 Row
Dim iOrd2_Curr_Row As Integer       'Select Order2 Row
Dim iOrd3_Curr_Row As Integer       'Select Order3 Row
Dim iOrd4_Curr_Row As Integer       'Select Order1 Row
Dim iOrd5_Curr_Row As Integer       'Select Order2 Row
Dim iOrd6_Curr_Row As Integer       'Select Order3 Row

Dim iSumCol As New Collection       'Sum Column

Dim vCR_CD As Variant               'First Slab CR_CD
Dim vSTLGRD As Variant              'First Slab STLGR
Dim vUST_FL As Variant              'First Slab UST_FL
Dim vSTDSPEC As Variant             'First Slab STDSPEC
Dim vISP_CMP As Variant             'First Slab ISP_CMP
Dim vPROD_WID As Variant            'First Slab PROD_WID
Dim vPROD_THK As Variant            'First Slab PROD_THK
Dim vENDUSE_CD As Variant           'First Slab ENDUSE_CD
Dim vORD_HCR_FL As Variant          'First Slab ORD_HCR_F
Dim vMLT_PROC_CD As Variant         'First Slab MLT_PROC_CD
Dim vORD_TRIM_FL As Variant         'First Slab ORD_TRIM_FL
Dim vCUST_SPEC_NO As Variant        'First Slab CUST_SPEC_NO
Dim vORD_NO As Variant              'First Slab ORD_NO
Dim vORD_ITEM As Variant            'First Slab ORD_ITEM

Dim mPlate_ORD_NO As Variant        'MPLATE First ORD_NO
Dim mPlate_ORD_ITEM As Variant      'MPLATE First ORD_ITEM

Dim sHTM_METH As String             'First Plate HTM_METH

Dim lMain_row As Long               'Main Row(Order no1)
Dim lSlab_left As Long              'Slab Left Position
Dim lMplate_left As Long            'Mplate Left Position
Dim iSLAB_EDT_SEQ As Long           'SLAB_EDT_SEQ Value

Dim lCool_max As Long               'COOLING BED LENGTH MAX SIZE

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
     
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
                Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_prc_line, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ccm_line, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_prod_cd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_prod_cd_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_del_fr, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_del_to, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_stlgrd_grp, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TxT_stdgrd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_prod_thk_from, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_prod_thk_to, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_prod_len_from, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_prod_len_to, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_prod_wid_from, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_prod_wid_to, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_SlaB_WIDTH_FROM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_SLAB_WIDTH_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_SLAB_WIDTH_TAG, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ord_item, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_cust_cd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
            
          
         Call Gp_Ms_Collection(txt_ord_no1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ord_no2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ord_no3, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ord_no4, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ord_no5, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_ord_no6, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ord11_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ord12_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ord21_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ord22_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ord31_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ord32_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ord41_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ord42_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ord51_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ord52_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ord61_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ord62_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_ord1_len, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_ord2_len, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_ord3_len, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_ord4_len, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_ord5_len, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_ord6_len, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      
      Call Gp_Ms_Collection(sdb_asroll_thk, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_asroll_wid, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_asroll_len, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_plate_len, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
        Call Gp_Ms_Collection(sdb_slab_thk, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_slab_wid, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_slab_len, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_slab_thk1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_slab_wid1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_slab_len1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_slab_wgt1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_slab_ratio, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
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
     Call Gp_Sp_Collection(SS1, 1, "p", "n", "m", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(SS1, 2, "p", "n", "m", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(SS1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(SS1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 15, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(SS1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 24, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 26, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(SS1, 27, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
     'Spread_Collection
    
    Sc1.Add Item:=SS1, Key:="Spread"
    Sc1.Add Item:="AEB1090C.P_SREFER", Key:="P-R"
    Sc1.Add Item:="AEB1090C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AEB1090C.P_SONEROW", Key:="P-O"
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  EDIT  End      ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=SS1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    'Sum Column Count
    iSumCnt = 4
    
    'Sum Column Setting
    iSumCol.Add Item:=26
    iSumCol.Add Item:=27
    iSumCol.Add Item:=28
    iSumCol.Add Item:=29
    
    Sc1.Item("Spread").Col = 0
    Sc1.Item("Spread").Row = 0
    Sc1.Item("Spread").Text = "◎"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub cmd_design_modify_Click()

    Complete = False
    Load AEB1091C
    AEB1091C.sdb_slab_edt_seq.Value = iSLAB_EDT_SEQ
    AEB1091C.sdb_slab_len.Value = sdb_slab_len1.Value
    AEB1091C.sdb_slab_wgt.Value = sdb_slab_wgt1.Value
    AEB1091C.Show 1

    If Complete Then
        Call cmd_slab_design_Click
        Call cmd_slab_del_Click
    End If

End Sub

Private Sub cmd_mplate_complete_Click()
    
    Dim sSeq As String
    Dim sQuery As String
    
    If sdb_plate_len.Value = 0 Then Exit Sub
    If iSlab_Complete > 0 Then Exit Sub
    
    If sdb_plate_len.Value + sdb_slab_len.Value >= 500000 Then
        Call Gp_MsgBoxDisplay("轧件长度 > 500,000")
        Exit Sub
    End If
    
    If iSlab_cnt = 0 Then
        
        SS1.Row = lMain_row
        
        SS1.Col = 1
        vORD_NO = SS1.Text
        SS1.Col = 2
        vORD_ITEM = SS1.Text
        SS1.Col = 5
        vENDUSE_CD = SS1.Text
        SS1.Col = 7
        vSTLGRD = SS1.Text
        SS1.Col = 11
        vPROD_THK = SS1.Value
        SS1.Col = 12
        vPROD_WID = SS1.Value
        SS1.Col = 18
        vMLT_PROC_CD = SS1.Text
        SS1.Col = 19
        vORD_HCR_FL = SS1.Text
        SS1.Col = 9
        vSTDSPEC = SS1.Text
        SS1.Col = 20
        vCR_CD = SS1.Text
        SS1.Col = 21
        vORD_TRIM_FL = SS1.Text
        SS1.Col = 22
        vUST_FL = SS1.Text
        SS1.Col = 25
        vCUST_SPEC_NO = SS1.Text
        
    End If
    
    iSlab_cnt = iSlab_cnt + 1
    cmd_slab_del.Enabled = True
    cmd_slab_design.Enabled = True
    cmd_slab_complete.Enabled = False
    
    If iSlab_cnt < 10 Then
        sSeq = "0" & iSlab_cnt
    Else
        sSeq = Trim(str(iSlab_cnt))
    End If
    
    sdb_slab_len.Value = sdb_slab_len.Value + sdb_plate_len.Value
    
    Load lbl_slab(iSlab_cnt)
    lbl_slab(iSlab_cnt).Tag = str(sdb_plate_len.Value)
    lbl_slab(iSlab_cnt).Caption = sSeq
    lbl_slab(iSlab_cnt).Top = 250
    lbl_slab(iSlab_cnt).Height = 780
    lbl_slab(iSlab_cnt).Width = (Shape4.Width / 500000) * sdb_plate_len.Value
        
    If iSlab_cnt = 1 Then
        lbl_slab(iSlab_cnt).Left = Shape4.Left
        lbl_slab(iSlab_cnt).Visible = True
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='00', SEQ='00'
        Call Slab_Block_Seq_Create("I")
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ=sSeq, SEQ= '00' ADD 1
        Call Slab_Seq_Create(sSeq, "I")
        
    Else
        If lbl_slab(iSlab_cnt - 1).Caption <> "删除" Then
            lbl_slab(iSlab_cnt).Left = lbl_slab(iSlab_cnt - 1).Left + lbl_slab(iSlab_cnt - 1).Width
        Else
            lbl_slab(iSlab_cnt).Left = lbl_slab(iSlab_cnt - 1).Left + lbl_slab(iSlab_cnt - 1).Width - 30
        End If
        
        lbl_slab(iSlab_cnt).Visible = True
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ=sSeq, SEQ= '00' ADD 1
        Call Slab_Seq_Create(sSeq, "I")
    End If
    
'    Call cmd_slab_design_Click
    
End Sub

Private Sub cmd_mplate_del_Click()

    Dim sSeq As String
    Dim iCount As Integer
    Dim iVisible_Cnt As Integer
    
    If iMplate_cnt = 0 Then Exit Sub
    
    For iCount = 1 To iMplate_cnt
        
        If lbl_mplate(iCount).Caption = "删除" Then
            
            If lbl_mplate(iCount).Visible Then
            
                lbl_mplate(iCount).Width = 0
                lbl_mplate(iCount).Visible = False
                
                If lbl_mplate(iCount).Tag = "ord1" Then
                    sdb_asroll_len.Value = sdb_asroll_len.Value - sdb_ord1_len.Value
                ElseIf lbl_mplate(iCount).Tag = "ord2" Then
                    sdb_asroll_len.Value = sdb_asroll_len.Value - sdb_ord2_len.Value
                ElseIf lbl_mplate(iCount).Tag = "ord3" Then
                    sdb_asroll_len.Value = sdb_asroll_len.Value - sdb_ord3_len.Value
                ElseIf lbl_mplate(iCount).Tag = "ord4" Then
                    sdb_asroll_len.Value = sdb_asroll_len.Value - sdb_ord4_len.Value
                ElseIf lbl_mplate(iCount).Tag = "ord5" Then
                    sdb_asroll_len.Value = sdb_asroll_len.Value - sdb_ord5_len.Value
                Else
                    sdb_asroll_len.Value = sdb_asroll_len.Value - sdb_ord6_len.Value
                End If
                
                If iCount < 10 Then
                    sSeq = "0" & iCount
                Else
                    sSeq = str(iCount)
                End If
                
                'EP_PLATE_EDT_CSL UPDATE  BLOCK_SEQ='01', SEQ      --> LEN = 0
                If lbl_mplate(iCount).Tag = "ord1" Then
                    Call Plate_Seq_Create(iOrd1_Curr_Row, sSeq, "U")
                ElseIf lbl_mplate(iCount).Tag = "ord2" Then
                    Call Plate_Seq_Create(iOrd2_Curr_Row, sSeq, "U")
                ElseIf lbl_mplate(iCount).Tag = "ord3" Then
                    Call Plate_Seq_Create(iOrd3_Curr_Row, sSeq, "U")
                ElseIf lbl_mplate(iCount).Tag = "ord4" Then
                    Call Plate_Seq_Create(iOrd4_Curr_Row, sSeq, "U")
                ElseIf lbl_mplate(iCount).Tag = "ord5" Then
                    Call Plate_Seq_Create(iOrd5_Curr_Row, sSeq, "U")
                Else
                    Call Plate_Seq_Create(iOrd6_Curr_Row, sSeq, "U")
                End If
                    
            End If
            
            cmd_mplate_complete.Enabled = False
            sdb_plate_len.Value = 0
                
        End If
    
        If iCount = 1 Then
            lbl_mplate(iCount).Left = Shape1.Left
        Else
            If lbl_mplate(iCount - 1).Caption <> "删除" Then
                lbl_mplate(iCount).Left = lbl_mplate(iCount - 1).Left + lbl_mplate(iCount - 1).Width
            Else
                lbl_mplate(iCount).Left = lbl_mplate(iCount - 1).Left + lbl_mplate(iCount - 1).Width - 30
            End If
        End If
        
    Next iCount
    
    iVisible_Cnt = 0
    For iCount = 1 To iMplate_cnt
    
        If lbl_mplate(iCount).Visible Then
            iVisible_Cnt = iVisible_Cnt + 1
        End If
    
    Next iCount
    
    'EP_PLATE_EDT_CSL DATA DELETE
    If iVisible_Cnt = 0 Then
    
        For iCount = 1 To iMplate_cnt
            Unload lbl_mplate(iCount)
        Next iCount
        
        iMplate_cnt = 0
        Call Plate_Seq_Create(lMain_row, "00", "D")
        cmd_mplate_del.Enabled = False
        cmd_mplate_design.Enabled = False
        
        mPlate_ORD_NO = ""
        mPlate_ORD_ITEM = ""
        
        If iSlab_cnt <= 0 Then
            sHTM_METH = ""
        End If
        
    End If
    
End Sub

Private Sub cmd_mplate_design_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    Dim iVisible_Cnt As Integer
    Dim lSlab_Edt_Seq As Double
    
    Dim adoCmd As adodb.Command
    Dim AdoRs As adodb.Recordset
    
    Set AdoRs = New adodb.Recordset
    
    For iCount = 1 To iMplate_cnt
        If lbl_mplate(iCount).Visible Then
            iVisible_Cnt = iVisible_Cnt + 1
        End If
    Next iCount
    
    If iVisible_Cnt = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'SLAB_EDT_SEQ Setting
    If txt_ccm_line.Text = "1" Then
        lSlab_Edt_Seq = 99999010
    ElseIf txt_ccm_line.Text = "2" Then
        lSlab_Edt_Seq = 99999020
    Else
        lSlab_Edt_Seq = 99999030
    End If
    
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
    
    sQuery = "{call AEB1092P (" & lSlab_Edt_Seq & ",'99','" + sUserID + "',?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'DESIGN LEN
'    sQuery = "SELECT NVL(LEN,0) FROM NISCO.EP_PLATE_EDT_CSL WHERE SLAB_EDT_SEQ = " & lSlab_Edt_Seq & " AND BLOCK_SEQ = '99' AND  SEQ = '00' "
'    sdb_plate_len.Value = Gf_FloatFind(M_CN1, sQuery)
    
    sQuery = "SELECT NVL(THK,0) ,NVL(WID,0) , NVL(LEN,0)"
    sQuery = sQuery + "  FROM EP_PLATE_EDT_CSL "
    sQuery = sQuery + " WHERE SLAB_EDT_SEQ = " & lSlab_Edt_Seq & " AND BLOCK_SEQ = '99' AND  SEQ = '00' "
    
    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    Do Until AdoRs.EOF
        sdb_asroll_thk.Value = Val(AdoRs.Fields(0) & "")
        sdb_asroll_wid.Value = Val(AdoRs.Fields(1) & "")
        sdb_asroll_len.Value = Val(AdoRs.Fields(2) & "")
        sdb_plate_len.Value = Val(AdoRs.Fields(2) & "")
        
        AdoRs.MoveNext
    Loop
    AdoRs.Close
    
    'Process Error Check
    If adoCmd("arg_e_code") <> "Y" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
        cmd_mplate_complete.Enabled = False
    Else
        cmd_mplate_complete.Enabled = True
    End If
    
    iVisible_Cnt = 0
    
    'Caption Rewrite
    For iCount = 1 To iMplate_cnt
    
        If lbl_mplate(iCount).Visible Then
        
            iVisible_Cnt = iVisible_Cnt + 1
            
            lbl_mplate(iVisible_Cnt).Visible = True
            lbl_mplate(iVisible_Cnt).BackColor = &HC0C0C0
            lbl_mplate(iVisible_Cnt).ForeColor = &HFF0000
            
            If iVisible_Cnt < 10 Then
                lbl_mplate(iVisible_Cnt).Caption = "0" & iVisible_Cnt
            Else
                lbl_mplate(iVisible_Cnt).Caption = Trim(str(iVisible_Cnt))
            End If
            
            lbl_mplate(iVisible_Cnt).Top = 250
            lbl_mplate(iVisible_Cnt).Height = 780
            
            If lbl_mplate(iCount).Tag = "ord1" Then
                lbl_mplate(iVisible_Cnt).Width = (Shape1.Width / lCool_max) * sdb_ord1_len.Value
            ElseIf lbl_mplate(iCount).Tag = "ord2" Then
                lbl_mplate(iVisible_Cnt).Width = (Shape1.Width / lCool_max) * sdb_ord2_len.Value
            ElseIf lbl_mplate(iCount).Tag = "ord3" Then
                lbl_mplate(iVisible_Cnt).Width = (Shape1.Width / lCool_max) * sdb_ord3_len.Value
            ElseIf lbl_mplate(iCount).Tag = "ord4" Then
                lbl_mplate(iVisible_Cnt).Width = (Shape1.Width / lCool_max) * sdb_ord4_len.Value
            ElseIf lbl_mplate(iCount).Tag = "ord5" Then
                lbl_mplate(iVisible_Cnt).Width = (Shape1.Width / lCool_max) * sdb_ord5_len.Value
            Else
                lbl_mplate(iVisible_Cnt).Width = (Shape1.Width / lCool_max) * sdb_ord6_len.Value
            End If
                
            lbl_mplate(iVisible_Cnt).Tag = lbl_mplate(iCount).Tag
            
            If iVisible_Cnt = 1 Then
                lbl_mplate(iVisible_Cnt).Left = Shape1.Left
            Else
                If lbl_mplate(iVisible_Cnt - 1).Caption <> "删除" Then
                    lbl_mplate(iVisible_Cnt).Left = lbl_mplate(iVisible_Cnt - 1).Left + lbl_mplate(iVisible_Cnt - 1).Width
                Else
                    lbl_mplate(iVisible_Cnt).Left = lbl_mplate(iVisible_Cnt - 1).Left + lbl_mplate(iVisible_Cnt - 1).Width - 30
                End If
            End If
                
        End If
    
    Next iCount
    
    'Remain Plate Delete
    For iCount = iVisible_Cnt + 1 To iMplate_cnt
        Unload lbl_mplate(iCount)
    Next iCount
    
    iMplate_cnt = iVisible_Cnt
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub cmd_mplate_init_Click()

    Dim iCnt As Long
    Dim iRow As Integer
    
    For iCnt = 1 To iMplate_cnt
        lbl_mplate(iCnt).Caption = "删除"
    Next iCnt

    Call cmd_mplate_del_Click
    
    txt_ord_no1.Text = ""
    txt_ord_no2.Text = ""
    txt_ord_no3.Text = ""
    txt_ord_no4.Text = ""
    txt_ord_no5.Text = ""
    txt_ord_no6.Text = ""
    sdb_ord11_cnt.Value = 0
    sdb_ord12_cnt.Value = 0
    sdb_ord21_cnt.Value = 0
    sdb_ord22_cnt.Value = 0
    sdb_ord31_cnt.Value = 0
    sdb_ord32_cnt.Value = 0
    sdb_ord41_cnt.Value = 0
    sdb_ord42_cnt.Value = 0
    sdb_ord51_cnt.Value = 0
    sdb_ord52_cnt.Value = 0
    sdb_ord61_cnt.Value = 0
    sdb_ord62_cnt.Value = 0
    sdb_ord1_len.Value = 0
    sdb_ord2_len.Value = 0
    sdb_ord3_len.Value = 0
    sdb_ord4_len.Value = 0
    sdb_ord5_len.Value = 0
    sdb_ord6_len.Value = 0
    sdb_asroll_thk.Value = 0
    sdb_asroll_wid.Value = 0
    sdb_asroll_len.Value = 0
    sdb_plate_len.Value = 0
    iOrd1_Curr_Row = 0
    iOrd2_Curr_Row = 0
    iOrd3_Curr_Row = 0
    iOrd4_Curr_Row = 0
    iOrd5_Curr_Row = 0
    iOrd6_Curr_Row = 0
    lMain_row = 0
    oRd_cnt = 0
    
    mPlate_ORD_NO = ""
    mPlate_ORD_ITEM = ""
    
    iMplate_cnt = 0
    cmd_mplate_del.Enabled = False
    cmd_mplate_complete.Enabled = False
    
    If iSlab_cnt <= 0 Then
        sHTM_METH = ""
    End If
    
    For iRow = 1 To SS1.MaxRows
        SS1.Row = iRow
        SS1.Col = 0
        SS1.Text = ""
        Call Gp_Sp_BlockColor(SS1, 1, SS1.MaxCols, iRow, iRow)
    Next iRow
            
End Sub

Private Sub cmd_ord1_Click()

    Dim sSeq As String
    Dim sQuery As String
    Dim lSlab_Edt_Seq As Double
    
'    If sdb_ord12_cnt.Value = 0 Then Exit Sub
    If sdb_ord11_cnt.Value = 0 Then Exit Sub
    
    'SLAB_EDT_SEQ Setting
    If txt_ccm_line.Text = "1" Then
        lSlab_Edt_Seq = 99999010
    ElseIf txt_ccm_line.Text = "2" Then
        lSlab_Edt_Seq = 99999020
    Else
        lSlab_Edt_Seq = 99999030
    End If
    
    If iMplate_cnt = 0 Then
        sQuery = " SELECT COUNT(*) FROM NISCO.EP_PLATE_EDT_CSL WHERE SLAB_EDT_SEQ = " & lSlab_Edt_Seq
        If Gf_FloatFind(M_CN1, sQuery) <> 0 Then
            Call Gp_MsgBoxDisplay("Another Job Processing..!!")
            Exit Sub
        End If
    Else
        If Plate_Setting_Check(Mid(txt_ord_no1.Text, 1, 11), Mid(txt_ord_no1.Text, 13, 2)) = False Then Exit Sub
    End If
    
    If sdb_asroll_len.Value + sdb_ord1_len.Value >= lCool_max Then
        Call Gp_MsgBoxDisplay("母板长度 >= " & lCool_max)
        Exit Sub
    End If
    
    sdb_ord12_cnt.Value = sdb_ord12_cnt.Value - 1
    iMplate_cnt = iMplate_cnt + 1
    cmd_mplate_del.Enabled = True
    cmd_mplate_design.Enabled = True
    
    If iMplate_cnt < 10 Then
        sSeq = "0" & iMplate_cnt
    Else
        sSeq = Trim(str(iMplate_cnt))
    End If
    
    sdb_asroll_len.Value = sdb_asroll_len.Value + sdb_ord1_len.Value
    
    Load lbl_mplate(iMplate_cnt)
    lbl_mplate(iMplate_cnt).Tag = "ord1"
    lbl_mplate(iMplate_cnt).Caption = sSeq
    lbl_mplate(iMplate_cnt).Top = 250
    lbl_mplate(iMplate_cnt).Height = 780
    lbl_mplate(iMplate_cnt).Width = (Shape1.Width / lCool_max) * sdb_ord1_len.Value
        
    If iMplate_cnt = 1 Then
        lbl_mplate(iMplate_cnt).Left = Shape1.Left
        lbl_mplate(iMplate_cnt).Visible = True
        
        Call Asroll_Thk(txt_ord_no1.Text)
        Call Asroll_Wid(txt_ord_no1.Text)
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ='00'
        Call Plate_Block_Seq_Create(iOrd1_Curr_Row, "I")
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd1_Curr_Row, sSeq, "I")
    Else
        If lbl_mplate(iMplate_cnt - 1).Caption <> "删除" Then
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width
        Else
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width - 30
        End If
        
        lbl_mplate(iMplate_cnt).Visible = True
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd1_Curr_Row, sSeq, "I")
    End If
    
End Sub

Private Sub cmd_ord2_Click()

    Dim sSeq As String
    Dim sQuery As String
    Dim lSlab_Edt_Seq As Double
    
'    If sdb_ord22_cnt.Value = 0 Then Exit Sub
    If sdb_ord21_cnt.Value = 0 Then Exit Sub
    
    'SLAB_EDT_SEQ Setting
    If txt_ccm_line.Text = "1" Then
        lSlab_Edt_Seq = 99999010
    ElseIf txt_ccm_line.Text = "2" Then
        lSlab_Edt_Seq = 99999020
    Else
        lSlab_Edt_Seq = 99999030
    End If
    
    If iMplate_cnt = 0 Then
        sQuery = " SELECT COUNT(*) FROM NISCO.EP_PLATE_EDT_CSL WHERE SLAB_EDT_SEQ = " & lSlab_Edt_Seq
        If Gf_FloatFind(M_CN1, sQuery) <> 0 Then
            Call Gp_MsgBoxDisplay("Another Job Processing..!!")
            Exit Sub
        End If
    Else
        If Plate_Setting_Check(Mid(txt_ord_no2.Text, 1, 11), Mid(txt_ord_no2.Text, 13, 2)) = False Then Exit Sub
    End If
    
    If sdb_asroll_len.Value + sdb_ord2_len.Value >= lCool_max Then
        Call Gp_MsgBoxDisplay("母板长度 >= " & lCool_max)
        Exit Sub
    End If
    
    sdb_ord22_cnt.Value = sdb_ord22_cnt.Value - 1
    iMplate_cnt = iMplate_cnt + 1
    cmd_mplate_del.Enabled = True
    cmd_mplate_design.Enabled = True
    
    If iMplate_cnt < 10 Then
       sSeq = "0" & iMplate_cnt
    Else
       sSeq = Trim(str(iMplate_cnt))
    End If
    
    sdb_asroll_len.Value = sdb_asroll_len.Value + sdb_ord2_len.Value
    
    Load lbl_mplate(iMplate_cnt)
    lbl_mplate(iMplate_cnt).Tag = "ord2"
    lbl_mplate(iMplate_cnt).Caption = sSeq
    lbl_mplate(iMplate_cnt).Top = 250
    lbl_mplate(iMplate_cnt).Height = 780
    lbl_mplate(iMplate_cnt).Width = (Shape1.Width / lCool_max) * sdb_ord2_len.Value
        
    If iMplate_cnt = 1 Then
        lbl_mplate(iMplate_cnt).Left = Shape1.Left
        lbl_mplate(iMplate_cnt).Visible = True
        
        Call Asroll_Thk(txt_ord_no2.Text)
        Call Asroll_Wid(txt_ord_no2.Text)
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ='00'
        Call Plate_Block_Seq_Create(iOrd2_Curr_Row, "I")
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd2_Curr_Row, sSeq, "I")
    Else
        If lbl_mplate(iMplate_cnt - 1).Caption <> "删除" Then
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width
        Else
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width - 30
        End If
        
        lbl_mplate(iMplate_cnt).Visible = True
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd2_Curr_Row, sSeq, "I")
    End If
    
End Sub

Private Sub cmd_ord3_Click()

    Dim sSeq As String
    Dim sQuery As String
    Dim lSlab_Edt_Seq As Double
    
'    If sdb_ord32_cnt.Value = 0 Then Exit Sub
    If sdb_ord31_cnt.Value = 0 Then Exit Sub
    
    'SLAB_EDT_SEQ Setting
    If txt_ccm_line.Text = "1" Then
        lSlab_Edt_Seq = 99999010
    ElseIf txt_ccm_line.Text = "2" Then
        lSlab_Edt_Seq = 99999020
    Else
        lSlab_Edt_Seq = 99999030
    End If
    
    If iMplate_cnt = 0 Then
        sQuery = " SELECT COUNT(*) FROM NISCO.EP_PLATE_EDT_CSL WHERE SLAB_EDT_SEQ = " & lSlab_Edt_Seq
        If Gf_FloatFind(M_CN1, sQuery) <> 0 Then
            Call Gp_MsgBoxDisplay("Another Job Processing..!!")
            Exit Sub
        End If
    Else
        If Plate_Setting_Check(Mid(txt_ord_no3.Text, 1, 11), Mid(txt_ord_no3.Text, 13, 2)) = False Then Exit Sub
    End If
    
    If sdb_asroll_len.Value + sdb_ord3_len.Value >= lCool_max Then
        Call Gp_MsgBoxDisplay("母板长度 >= " & lCool_max)
        Exit Sub
    End If
    
    sdb_ord32_cnt.Value = sdb_ord32_cnt.Value - 1
    iMplate_cnt = iMplate_cnt + 1
    cmd_mplate_del.Enabled = True
    cmd_mplate_design.Enabled = True
    
    If iMplate_cnt < 10 Then
       sSeq = "0" & iMplate_cnt
    Else
       sSeq = Trim(str(iMplate_cnt))
    End If
    
    sdb_asroll_len.Value = sdb_asroll_len.Value + sdb_ord3_len.Value
    
    Load lbl_mplate(iMplate_cnt)
    lbl_mplate(iMplate_cnt).Tag = "ord3"
    lbl_mplate(iMplate_cnt).Caption = sSeq
    lbl_mplate(iMplate_cnt).Top = 250
    lbl_mplate(iMplate_cnt).Height = 780
    lbl_mplate(iMplate_cnt).Width = (Shape1.Width / lCool_max) * sdb_ord3_len.Value
        
    If iMplate_cnt = 1 Then
        lbl_mplate(iMplate_cnt).Left = Shape1.Left
        lbl_mplate(iMplate_cnt).Visible = True
        
        Call Asroll_Thk(txt_ord_no3.Text)
        Call Asroll_Wid(txt_ord_no3.Text)
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ='00'
        Call Plate_Block_Seq_Create(iOrd3_Curr_Row, "I")
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd3_Curr_Row, sSeq, "I")
    Else
        If lbl_mplate(iMplate_cnt - 1).Caption <> "删除" Then
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width
        Else
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width - 30
        End If
        
        lbl_mplate(iMplate_cnt).Visible = True
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd3_Curr_Row, sSeq, "I")
    End If
    
End Sub

Private Sub cmd_ord4_Click()

    Dim sSeq As String
    Dim sQuery As String
    Dim lSlab_Edt_Seq As Double
    
'    If sdb_ord42_cnt.Value = 0 Then Exit Sub
    If sdb_ord41_cnt.Value = 0 Then Exit Sub
    
    'SLAB_EDT_SEQ Setting
    If txt_ccm_line.Text = "1" Then
        lSlab_Edt_Seq = 99999010
    ElseIf txt_ccm_line.Text = "2" Then
        lSlab_Edt_Seq = 99999020
    Else
        lSlab_Edt_Seq = 99999030
    End If
    
    If iMplate_cnt = 0 Then
        sQuery = " SELECT COUNT(*) FROM NISCO.EP_PLATE_EDT_CSL WHERE SLAB_EDT_SEQ = " & lSlab_Edt_Seq
        If Gf_FloatFind(M_CN1, sQuery) <> 0 Then
            Call Gp_MsgBoxDisplay("Another Job Processing..!!")
            Exit Sub
        End If
    Else
        If Plate_Setting_Check(Mid(txt_ord_no4.Text, 1, 11), Mid(txt_ord_no4.Text, 13, 2)) = False Then Exit Sub
    End If
    
    If sdb_asroll_len.Value + sdb_ord4_len.Value >= lCool_max Then
        Call Gp_MsgBoxDisplay("母板长度 >= " & lCool_max)
        Exit Sub
    End If
    
    sdb_ord42_cnt.Value = sdb_ord42_cnt.Value - 1
    iMplate_cnt = iMplate_cnt + 1
    cmd_mplate_del.Enabled = True
    cmd_mplate_design.Enabled = True
    
    If iMplate_cnt < 10 Then
       sSeq = "0" & iMplate_cnt
    Else
       sSeq = Trim(str(iMplate_cnt))
    End If
    
    sdb_asroll_len.Value = sdb_asroll_len.Value + sdb_ord4_len.Value
    
    Load lbl_mplate(iMplate_cnt)
    lbl_mplate(iMplate_cnt).Tag = "ord4"
    lbl_mplate(iMplate_cnt).Caption = sSeq
    lbl_mplate(iMplate_cnt).Top = 250
    lbl_mplate(iMplate_cnt).Height = 780
    lbl_mplate(iMplate_cnt).Width = (Shape1.Width / lCool_max) * sdb_ord4_len.Value
        
    If iMplate_cnt = 1 Then
        lbl_mplate(iMplate_cnt).Left = Shape1.Left
        lbl_mplate(iMplate_cnt).Visible = True
        
        Call Asroll_Thk(txt_ord_no4.Text)
        Call Asroll_Wid(txt_ord_no4.Text)
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ='00'
        Call Plate_Block_Seq_Create(iOrd4_Curr_Row, "I")
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd4_Curr_Row, sSeq, "I")
    Else
        If lbl_mplate(iMplate_cnt - 1).Caption <> "删除" Then
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width
        Else
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width - 30
        End If
        
        lbl_mplate(iMplate_cnt).Visible = True
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd4_Curr_Row, sSeq, "I")
    End If
    
End Sub

Private Sub cmd_ord5_Click()

    Dim sSeq As String
    Dim sQuery As String
    Dim lSlab_Edt_Seq As Double
    
'    If sdb_ord52_cnt.Value = 0 Then Exit Sub
    If sdb_ord51_cnt.Value = 0 Then Exit Sub
    
    'SLAB_EDT_SEQ Setting
    If txt_ccm_line.Text = "1" Then
        lSlab_Edt_Seq = 99999010
    ElseIf txt_ccm_line.Text = "2" Then
        lSlab_Edt_Seq = 99999020
    Else
        lSlab_Edt_Seq = 99999030
    End If
    
    If iMplate_cnt = 0 Then
        sQuery = " SELECT COUNT(*) FROM NISCO.EP_PLATE_EDT_CSL WHERE SLAB_EDT_SEQ = " & lSlab_Edt_Seq
        If Gf_FloatFind(M_CN1, sQuery) <> 0 Then
            Call Gp_MsgBoxDisplay("Another Job Processing..!!")
            Exit Sub
        End If
    Else
        If Plate_Setting_Check(Mid(txt_ord_no5.Text, 1, 11), Mid(txt_ord_no5.Text, 13, 2)) = False Then Exit Sub
    End If
    
    If sdb_asroll_len.Value + sdb_ord5_len.Value >= lCool_max Then
        Call Gp_MsgBoxDisplay("母板长度 >= " & lCool_max)
        Exit Sub
    End If
    
    sdb_ord52_cnt.Value = sdb_ord52_cnt.Value - 1
    iMplate_cnt = iMplate_cnt + 1
    cmd_mplate_del.Enabled = True
    cmd_mplate_design.Enabled = True
    
    If iMplate_cnt < 10 Then
       sSeq = "0" & iMplate_cnt
    Else
       sSeq = Trim(str(iMplate_cnt))
    End If
    
    sdb_asroll_len.Value = sdb_asroll_len.Value + sdb_ord5_len.Value
    
    Load lbl_mplate(iMplate_cnt)
    lbl_mplate(iMplate_cnt).Tag = "ord5"
    lbl_mplate(iMplate_cnt).Caption = sSeq
    lbl_mplate(iMplate_cnt).Top = 250
    lbl_mplate(iMplate_cnt).Height = 780
    lbl_mplate(iMplate_cnt).Width = (Shape1.Width / lCool_max) * sdb_ord5_len.Value
        
    If iMplate_cnt = 1 Then
        lbl_mplate(iMplate_cnt).Left = Shape1.Left
        lbl_mplate(iMplate_cnt).Visible = True
        
        Call Asroll_Thk(txt_ord_no5.Text)
        Call Asroll_Wid(txt_ord_no5.Text)
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ='00'
        Call Plate_Block_Seq_Create(iOrd5_Curr_Row, "I")
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd5_Curr_Row, sSeq, "I")
    Else
        If lbl_mplate(iMplate_cnt - 1).Caption <> "删除" Then
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width
        Else
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width - 30
        End If
        
        lbl_mplate(iMplate_cnt).Visible = True
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd5_Curr_Row, sSeq, "I")
    End If
    
End Sub

Private Sub cmd_ord6_Click()

    Dim sSeq As String
    Dim sQuery As String
    Dim lSlab_Edt_Seq As Double
    
'    If sdb_ord62_cnt.Value = 0 Then Exit Sub
    If sdb_ord61_cnt.Value = 0 Then Exit Sub
    
    'SLAB_EDT_SEQ Setting
    If txt_ccm_line.Text = "1" Then
        lSlab_Edt_Seq = 99999010
    ElseIf txt_ccm_line.Text = "2" Then
        lSlab_Edt_Seq = 99999020
    Else
        lSlab_Edt_Seq = 99999030
    End If
    
    If iMplate_cnt = 0 Then
        sQuery = " SELECT COUNT(*) FROM NISCO.EP_PLATE_EDT_CSL WHERE SLAB_EDT_SEQ = " & lSlab_Edt_Seq
        If Gf_FloatFind(M_CN1, sQuery) <> 0 Then
            Call Gp_MsgBoxDisplay("Another Job Processing..!!")
            Exit Sub
        End If
    Else
        If Plate_Setting_Check(Mid(txt_ord_no6.Text, 1, 11), Mid(txt_ord_no6.Text, 13, 2)) = False Then Exit Sub
    End If
    
    If sdb_asroll_len.Value + sdb_ord6_len.Value >= lCool_max Then
        Call Gp_MsgBoxDisplay("母板长度 >= " & lCool_max)
        Exit Sub
    End If
    
    sdb_ord62_cnt.Value = sdb_ord62_cnt.Value - 1
    iMplate_cnt = iMplate_cnt + 1
    cmd_mplate_del.Enabled = True
    cmd_mplate_design.Enabled = True
    
    If iMplate_cnt < 10 Then
       sSeq = "0" & iMplate_cnt
    Else
       sSeq = Trim(str(iMplate_cnt))
    End If
    
    sdb_asroll_len.Value = sdb_asroll_len.Value + sdb_ord6_len.Value
    
    Load lbl_mplate(iMplate_cnt)
    lbl_mplate(iMplate_cnt).Tag = "ord6"
    lbl_mplate(iMplate_cnt).Caption = sSeq
    lbl_mplate(iMplate_cnt).Top = 250
    lbl_mplate(iMplate_cnt).Height = 780
    lbl_mplate(iMplate_cnt).Width = (Shape1.Width / lCool_max) * sdb_ord6_len.Value
        
    If iMplate_cnt = 1 Then
        lbl_mplate(iMplate_cnt).Left = Shape1.Left
        lbl_mplate(iMplate_cnt).Visible = True
        
        Call Asroll_Thk(txt_ord_no6.Text)
        Call Asroll_Wid(txt_ord_no6.Text)
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ='00'
        Call Plate_Block_Seq_Create(iOrd6_Curr_Row, "I")
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd6_Curr_Row, sSeq, "I")
    Else
        If lbl_mplate(iMplate_cnt - 1).Caption <> "删除" Then
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width
        Else
            lbl_mplate(iMplate_cnt).Left = lbl_mplate(iMplate_cnt - 1).Left + lbl_mplate(iMplate_cnt - 1).Width - 30
        End If
        
        lbl_mplate(iMplate_cnt).Visible = True
        
        'EP_PLATE_EDT_CSL INSERT  BLOCK_SEQ='01', SEQ ADD 1
        Call Plate_Seq_Create(iOrd6_Curr_Row, sSeq, "I")
    End If
    
End Sub

Private Sub cmd_slab_complete_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iRow As Integer
    
    Dim adoCmd As adodb.Command
    
    If sdb_slab_wgt1.Value = 0 Then Exit Sub
    
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
    
    If iSlab_Complete = 0 Then
        sQuery = "{call AEB1095P (" & iSLAB_EDT_SEQ & ",'R',?,?)}"
    Else
        sQuery = "{call AEB1095P (" & iSLAB_EDT_SEQ & ",'C',?,?)}"
    End If
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_code") <> "Y" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        
        iSlab_Complete = iSlab_Complete + 1
        iLastSlab_cnt = iSlab_cnt              'Complete Slab Count
        
        cmd_slab_design.Enabled = False
        cmd_slab_del.Enabled = False
        
        'Spread Sheet Refresh
        'Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1)
        
        If iOrd1_Curr_Row <> 0 Then
            sQuery = Gf_Sp_MakeQuery(Sc1.Item("Spread"), Sc1.Item("P-O"), "O", Sc1.Item("pColumn"), iOrd1_Curr_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, Sc1.Item("Spread"), iOrd1_Curr_Row)
        End If
        
        If iOrd2_Curr_Row <> 0 Then
            sQuery = Gf_Sp_MakeQuery(Sc1.Item("Spread"), Sc1.Item("P-O"), "O", Sc1.Item("pColumn"), iOrd2_Curr_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, Sc1.Item("Spread"), iOrd2_Curr_Row)
        End If
        
        If iOrd3_Curr_Row <> 0 Then
            sQuery = Gf_Sp_MakeQuery(Sc1.Item("Spread"), Sc1.Item("P-O"), "O", Sc1.Item("pColumn"), iOrd3_Curr_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, Sc1.Item("Spread"), iOrd3_Curr_Row)
        End If
        
        If iOrd4_Curr_Row <> 0 Then
            sQuery = Gf_Sp_MakeQuery(Sc1.Item("Spread"), Sc1.Item("P-O"), "O", Sc1.Item("pColumn"), iOrd4_Curr_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, Sc1.Item("Spread"), iOrd4_Curr_Row)
        End If
        
        If iOrd5_Curr_Row <> 0 Then
            sQuery = Gf_Sp_MakeQuery(Sc1.Item("Spread"), Sc1.Item("P-O"), "O", Sc1.Item("pColumn"), iOrd5_Curr_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, Sc1.Item("Spread"), iOrd5_Curr_Row)
        End If
        
        If iOrd6_Curr_Row <> 0 Then
            sQuery = Gf_Sp_MakeQuery(Sc1.Item("Spread"), Sc1.Item("P-O"), "O", Sc1.Item("pColumn"), iOrd6_Curr_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, Sc1.Item("Spread"), iOrd6_Curr_Row)
        End If
        
        
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub cmd_slab_del_Click()

    Dim sSeq As String
    
    Dim iCount As Integer
    Dim iVisible_Cnt As Integer
    
    If iSlab_cnt = 0 Then Exit Sub
    
    For iCount = 1 To iSlab_cnt
        
        If lbl_slab(iCount).Caption = "删除" Then  'Delete
            
            If lbl_slab(iCount).Visible Then
            
                lbl_slab(iCount).Width = 0
                lbl_slab(iCount).Visible = False
                
                sdb_slab_len.Value = sdb_slab_len.Value - Val(lbl_slab(iCount).Tag)
                
                If iCount < 10 Then
                    sSeq = "0" & iCount
                Else
                    sSeq = Trim(str(iCount))
                End If
                
                'EP_PLATE_EDT UPDATE  BLOCK_SEQ='01', SEQ      --> LEN = 0
                Call Slab_Seq_Create(sSeq, "U")
                    
            End If
            
            sdb_slab_thk1.Value = 0
            sdb_slab_wid1.Value = 0
            sdb_slab_len1.Value = 0
            sdb_slab_ratio.Value = 0
            sdb_slab_wgt1.Value = 0
            cmd_slab_complete.Enabled = False
            
        End If
    
        If iCount = 1 Then
            lbl_slab(iCount).Left = Shape4.Left
        Else
            If lbl_slab(iCount - 1).Caption <> "删除" Then
                If iCount Mod 3 = 1 Then
                    lbl_slab(iCount).Left = lbl_slab(iCount - 1).Left + lbl_slab(iCount - 1).Width - 10
                Else
                    lbl_slab(iCount).Left = lbl_slab(iCount - 1).Left + lbl_slab(iCount - 1).Width
                End If
            Else
                lbl_slab(iCount).Left = lbl_slab(iCount - 1).Left + lbl_slab(iCount - 1).Width - 30
            End If
        End If
    
    Next iCount
    
    iVisible_Cnt = 0
    For iCount = 1 To iSlab_cnt
        If lbl_slab(iCount).Visible Then
            iVisible_Cnt = iVisible_Cnt + 1
        End If
    Next iCount
    
    'EP_PLATE_EDT_CSL DATA DELETE
    If iVisible_Cnt = 0 Then
   
        For iCount = 1 To iSlab_cnt
            Unload lbl_slab(iCount)
        Next iCount
    
        iSlab_Complete = 0
        iSlab_cnt = 0
        
        sdb_slab_thk.Value = 0
        sdb_slab_wid.Value = 0
        sdb_slab_len.Value = 0
        
        sdb_slab_thk1.Value = 0
        sdb_slab_wid1.Value = 0
        sdb_slab_len1.Value = 0
        sdb_slab_ratio.Value = 0
        sdb_slab_wgt1.Value = 0
        
        Call Slab_Seq_Create("00", "D")
        cmd_slab_del.Enabled = False
        cmd_slab_design.Enabled = False
        cmd_design_modify.Visible = False
        
        vORD_NO = ""
        vORD_ITEM = ""
        vENDUSE_CD = ""
        vSTLGRD = ""
        vPROD_THK = ""
        vPROD_WID = ""
        vMLT_PROC_CD = ""
        vORD_HCR_FL = ""
        vSTDSPEC = ""
        vISP_CMP = ""
        vCR_CD = ""
        vORD_TRIM_FL = ""
        vUST_FL = ""
        vCUST_SPEC_NO = ""
        
        If iMplate_cnt <= 0 Then
            sHTM_METH = ""
        End If
        
    End If
    
End Sub

Private Sub cmd_slab_design_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    Dim iVisible_Cnt As Integer
    
    Dim P_SLAB_EDT_SEQ As Long
    
    Dim AdoRs As adodb.Recordset
    Dim adoCmd As adodb.Command
    Set AdoRs = New adodb.Recordset
    
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
    
    sQuery = "{call AEB1093P (" & iSLAB_EDT_SEQ & ",'" + sUserID + "',?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'SLAB THK, WID, LEN, WGT Display
    Call Slab_Size
    
    'PLATE LEN, THK, WID
    Call Plate_Size
    
    'Process Error Check
    If adoCmd("arg_e_code") <> "Y" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
        cmd_slab_complete.Enabled = False
    Else
        cmd_slab_complete.Enabled = True
        cmd_design_modify.Visible = False
    End If
    
    cmd_design_modify.Visible = True
    iVisible_Cnt = 0
    
    'Plate Delete Setting
    For iCount = 1 To iSlab_cnt
        Unload lbl_slab(iCount)
    Next iCount
    
    'Plate Redisplay
    sQuery = "         SELECT BLOCK_SEQ, NVL(LEN,0) FROM NISCO.EP_PLATE_EDT_CSL "
    sQuery = sQuery + " WHERE SLAB_EDT_SEQ = " & iSLAB_EDT_SEQ
    sQuery = sQuery + "   AND BLOCK_SEQ    <> '00' "
    sQuery = sQuery + "   AND SEQ          =  '00' "
    sQuery = sQuery + " ORDER BY BLOCK_SEQ, SEQ "
    
    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        While Not AdoRs.EOF
        
            iVisible_Cnt = iVisible_Cnt + 1
            
            Load lbl_slab(iVisible_Cnt)
            lbl_slab(iVisible_Cnt).Visible = True
            lbl_slab(iVisible_Cnt).BackColor = &H8080FF
            lbl_slab(iVisible_Cnt).ForeColor = &HFF0000
            
            lbl_slab(iVisible_Cnt).Caption = AdoRs.Fields(0)
            
            lbl_slab(iVisible_Cnt).Top = 250
            lbl_slab(iVisible_Cnt).Height = 780
            
            lbl_slab(iVisible_Cnt).Tag = str(AdoRs.Fields(1))
            lbl_slab(iVisible_Cnt).Width = (Shape4.Width / 500000) * Val(lbl_slab(iVisible_Cnt).Tag)
            
            If iVisible_Cnt = 1 Then
                lbl_slab(iVisible_Cnt).Left = Shape4.Left
            Else
                If lbl_slab(iVisible_Cnt - 1).Caption <> "删除" Then
                    lbl_slab(iVisible_Cnt).Left = lbl_slab(iVisible_Cnt - 1).Left + lbl_slab(iVisible_Cnt - 1).Width
                Else
                    lbl_slab(iVisible_Cnt).Left = lbl_slab(iVisible_Cnt - 1).Left + lbl_slab(iVisible_Cnt - 1).Width - 30
                End If
            End If
            AdoRs.MoveNext
            
        Wend
        
    End If
    
    iSlab_cnt = iVisible_Cnt
    
    AdoRs.Close
    Set AdoRs = Nothing
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    AdoRs.Close
    Set AdoRs = Nothing
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub cmd_slab_init_Click()

    Dim iCnt As Long
    Dim iRow As Integer
    
    'Slab Complete Count
    If iSlab_Complete = 0 Then
        For iCnt = 1 To iSlab_cnt
            lbl_slab(iCnt).Caption = "删除"
        Next iCnt
    
        Call cmd_slab_del_Click
    Else
        For iCnt = 1 To iSlab_cnt
            Unload lbl_slab(iCnt)
        Next iCnt
        
        iSlab_Complete = 0
        sdb_slab_thk.Value = 0
        sdb_slab_wid.Value = 0
        sdb_slab_len.Value = 0
        
        sdb_slab_thk1.Value = 0
        sdb_slab_wid1.Value = 0
        sdb_slab_len1.Value = 0
        sdb_slab_ratio.Value = 0
        sdb_slab_wgt1.Value = 0
    End If
    
    iSlab_cnt = 0
    
    vORD_NO = ""
    vORD_ITEM = ""
    vENDUSE_CD = ""
    vSTLGRD = ""
    vPROD_THK = ""
    vPROD_WID = ""
    vMLT_PROC_CD = ""
    vORD_HCR_FL = ""
    vSTDSPEC = ""
    vISP_CMP = ""
    vCR_CD = ""
    vORD_TRIM_FL = ""
    vUST_FL = ""
    vCUST_SPEC_NO = ""
    
    cmd_slab_del.Enabled = False
    cmd_slab_complete.Enabled = False
    cmd_design_modify.Visible = False
    
    If iMplate_cnt <= 0 Then
        sHTM_METH = ""
    End If
    
End Sub

Private Sub cmd_Wid_Modify_Click()

    On Error GoTo Wid_Modify_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As adodb.Command
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
        
    If Trim(txt_plt.Text) = "" Then
       Call Gp_MsgBoxDisplay(txt_plt.Tag & "必须输入")
       Exit Sub
    End If
       
    If Trim(txt_prc_line.Text) = "" Then
       Call Gp_MsgBoxDisplay(txt_prc_line.Tag & "必须输入")
       Exit Sub
    End If
    
    If TXT_SlaB_WIDTH_FROM.Value = 0 Then
       Call Gp_MsgBoxDisplay(TXT_SlaB_WIDTH_FROM.Tag & "必须输入")
       Exit Sub
    End If
        
    If TXT_SLAB_WIDTH_TO.Value = 0 Then
       Call Gp_MsgBoxDisplay(TXT_SLAB_WIDTH_TO.Tag & "必须输入")
       Exit Sub
    End If
        
    If TXT_SLAB_WIDTH_TAG.Value = 0 Then
       Call Gp_MsgBoxDisplay(TXT_SLAB_WIDTH_TAG.Tag & "必须输入")
       Exit Sub
    End If
    
    If Not Gf_MessConfirm("您确定要板坯变成宽度吗？", "Q") Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    sQuery = "{call AEB1096P ('" & txt_plt.Text & "','" & txt_ccm_line.Text & "'," & _
                                   TXT_SlaB_WIDTH_FROM.Value & "," & _
                                   TXT_SLAB_WIDTH_TO.Value & "," & _
                                   TXT_SLAB_WIDTH_TAG.Value & ",?)}"
    
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
    Else
        TXT_SlaB_WIDTH_FROM = TXT_SLAB_WIDTH_TAG
        TXT_SLAB_WIDTH_TO = TXT_SLAB_WIDTH_TAG
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Wid_Modify_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub Form_Activate()
    
    If Active_CForm <> "" Then
        Call txt_prod_cd_KeyUp(0, 0)
        Call Form_Ref
        Active_CForm = ""
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    If Mid(sAuthority, 4, 1) <> "1" Then
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
    Else
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
       
    End If
    
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
        SSCommand1.Enabled = False
        
        cmd_ord1.Enabled = False
        cmd_ord2.Enabled = False
        cmd_ord3.Enabled = False
        cmd_ord4.Enabled = False
        cmd_ord5.Enabled = False
        cmd_ord6.Enabled = False
    End If
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    If Mid(sAuthority, 4, 1) <> "1" Then
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
    Else
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = True
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
       
    End If
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "E-System.INI", Me.Name)
    
    SS1.RetainSelBlock = False
    SS1.OperationMode = OperationModeNormal
    
    txt_plt.Text = "B1"
    Call txt_plt_KeyUp(0, 0)
    txt_prc_line.Text = "1"
    txt_ccm_line.Text = "1"
    txt_del_fr.Text = ""
    txt_del_to.Text = ""
    
    Active_CForm = ""
    
    lCool_max = Gf_FloatFind(M_CN1, "SELECT MAXI FROM EP_SLABDESIGN WHERE PLT = 'C1' AND APLY_ITEM = 'SLABDESIGN008' AND PRC_LINE = '1'")
    
    If lCool_max = 0 Then
        Label2.Caption = "0(M)"
    Else
        Label2.Caption = lCool_max / 1000 & "(M)"
    End If
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim iCount As Integer
    
    If iMplate_cnt > 0 Then
        Call Gp_MsgBoxDisplay("Must plate data clear necessarily")
        Cancel = 1
        Exit Sub
    End If
    
    If iSlab_cnt > 0 Then
        If iSlab_Complete < 1 Then
            Call Gp_MsgBoxDisplay("Must slab data clear necessarily")
            Cancel = 1
            Exit Sub
        End If
    End If
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "E-System.INI", Me.Name)
    
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
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Dim iRow As Integer
    
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    
    With SS1
        For iRow = 1 To .MaxRows - 1
            .Row = iRow
            .Col = 13
            If Trim(.Text) <> "定尺" Then
                .Col = 14:    .Lock = False
            Else
                .Col = 14:    .Lock = True
                Call Gp_Sp_BlockColor(SS1, 14, 14, iRow, iRow, BLACK, WHITE)
            End If
        Next iRow
    End With
      
End Sub

Public Sub Form_Cls()
    
    Dim iCnt As Long
    
    If iMplate_cnt > 0 Then
        Call Gp_MsgBoxDisplay("Must plate data clear necessarily")
        Exit Sub
    End If
    
    If iSlab_cnt > 0 Then
        Call Gp_MsgBoxDisplay("Must slab data clear necessarily")
        Exit Sub
    End If
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
    
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        If Mid(sAuthority, 4, 1) <> "1" Then
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        Else
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
           
        End If
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        
        rControl(1).SetFocus
        SS1.SetFocus
        
        txt_plt.Text = "B1"
        Call txt_plt_KeyUp(0, 0)
        txt_prc_line.Text = "1"
        txt_ccm_line.Text = "1"
        txt_prod_cd.Text = "PP"
        Call txt_prod_cd_KeyUp(0, 0)
        
        txt_del_fr.Text = ""
        txt_del_to.Text = ""
        txt_prod_cd_name.Text = ""
    
        oRd_cnt = 0
        iOrd1_Curr_Row = 0
        iOrd2_Curr_Row = 0
        iOrd3_Curr_Row = 0
        iOrd4_Curr_Row = 0
        iOrd5_Curr_Row = 0
        iOrd6_Curr_Row = 0
        iSLAB_EDT_SEQ = 0
'        opt_sort1.Value = True
        cmd_mplate_del.Enabled = False
        
        For iCnt = 1 To iMplate_cnt
            Unload lbl_mplate(iCnt)
        Next iCnt
        
        For iCnt = 1 To iSlab_cnt
            Unload lbl_slab(iCnt)
        Next iCnt
        
        iMplate_cnt = 0
        iSlab_cnt = 0
        
        mPlate_ORD_NO = ""
        mPlate_ORD_ITEM = ""
        
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sQuery As String
    Dim dValue As String
    
    Dim iCnt As Long
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    'EP_PLATE_EDT_CSL DATA DELETE
    If iMplate_cnt > 0 Then
        Call Gp_MsgBoxDisplay("Must plate data clear necessarily")
        Exit Sub
    End If
        
    'EP_PLATE_EDT_CSL DATA DELETE
    If iSlab_cnt > 0 Then
        Call Gp_MsgBoxDisplay("Must slab data clear necessarily")
        Exit Sub
    End If
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call Sp_Total
        'Call Gp_Sp_EvenRowBackcolor(Proc_Sc("Sc").Item("Spread"))
        SS1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        Call SS1_CHANGE_COLOR
        
        If Mid(sAuthority, 4, 1) <> "1" Then   'DELETE
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        Else
            MDIMain.MenuTool.Buttons(7).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        End If
        
        txt_ord_no1.Text = ""
        txt_ord_no2.Text = ""
        txt_ord_no3.Text = ""
        txt_ord_no4.Text = ""
        txt_ord_no5.Text = ""
        txt_ord_no6.Text = ""
        sdb_ord11_cnt.Value = 0
        sdb_ord12_cnt.Value = 0
        sdb_ord21_cnt.Value = 0
        sdb_ord22_cnt.Value = 0
        sdb_ord31_cnt.Value = 0
        sdb_ord32_cnt.Value = 0
        sdb_ord41_cnt.Value = 0
        sdb_ord42_cnt.Value = 0
        sdb_ord51_cnt.Value = 0
        sdb_ord52_cnt.Value = 0
        sdb_ord61_cnt.Value = 0
        sdb_ord62_cnt.Value = 0
        sdb_ord1_len.Value = 0
        sdb_ord2_len.Value = 0
        sdb_ord3_len.Value = 0
        sdb_ord4_len.Value = 0
        sdb_ord5_len.Value = 0
        sdb_ord6_len.Value = 0
        sdb_asroll_thk.Value = 0
        sdb_asroll_wid.Value = 0
        sdb_asroll_len.Value = 0
        sdb_slab_thk1.Value = 0
        sdb_slab_wid1.Value = 0
        sdb_slab_len1.Value = 0
        sdb_slab_ratio.Value = 0
        sdb_slab_wgt1.Value = 0
        
        iOrd1_Curr_Row = 0
        iOrd2_Curr_Row = 0
        iOrd3_Curr_Row = 0
        iOrd4_Curr_Row = 0
        iOrd5_Curr_Row = 0
        iOrd6_Curr_Row = 0
        lMain_row = 0
        oRd_cnt = 0
        
        mPlate_ORD_NO = ""
        mPlate_ORD_ITEM = ""
        
        For iCnt = 1 To iMplate_cnt
            Unload lbl_mplate(iCnt)
        Next iCnt
        
        For iCnt = 1 To iSlab_cnt
            Unload lbl_slab(iCnt)
        Next iCnt
        
        iMplate_cnt = 0
        iSlab_cnt = 0
        iSlab_Complete = 0
        
        iSLAB_EDT_SEQ = 0
        cmd_mplate_del.Enabled = False
        cmd_slab_del.Enabled = False
        
        Exit Sub
        
        
    End If
    
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()


End Sub

Public Sub Form_Ins()

   Call Gp_Sp_Ins(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Cpy()

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
    
End Sub

Private Sub SS1_CHANGE_COLOR()
Dim iCount As Integer


    With SS1

        If .MaxRows <= 0 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .Row = iCount

             '重点订单红色标记 2013-11-16  by  CaoLei
            SS1.Row = .Row:          SS1.Col = 33
            If SS1.Text = "Y" Then
                 Call Gp_Sp_BlockColor(SS1, 1, 35, .Row, .Row, &HFF&)
'                 Call Gp_Sp_BlockColor(SS1, 33, 33, .Row, .Row, &HFF&)
            End If

        Next iCount

    End With

End Sub

Public Sub Gp_Process_Exec(P_MODE As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As adodb.Command
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
        
    If txt_prod_thk_to.Value = 0 Then
       txt_prod_thk_to.Value = 9999.99
    End If
        
    If txt_prod_wid_to.Value = 0 Then
       txt_prod_wid_to.Value = 9999.99
    End If
        
    If txt_prod_len_to.Value = 0 Then
       txt_prod_len_to.Value = 9999999.9
    End If
    
    sQuery = "{call AEB1090P ('" + txt_plt.Text + "','" + txt_prc_line.Text + "','" + txt_ccm_line.Text + "','" & _
                                   Trim(txt_ord_no.Text) + "','" + Trim(txt_ord_item.Text) + "','" & _
                                   Trim(txt_prod_cd.Text) + "','" + Trim(txt_cust_cd.Text) + "','" + Trim(txt_stlgrd_grp.Text) + "','" & _
                                   Trim(TxT_stdgrd.Text) + "','" + Trim(txt_stdspec.Text) + "','" & _
                                   Trim(txt_del_fr.RawData) + "','" + Trim(txt_del_to.RawData) + "'," & _
                                   txt_prod_thk_from.Value & "," & txt_prod_thk_to.Value & "," & txt_prod_wid_from.Value & "," & _
                                   txt_prod_wid_to.Value & "," & txt_prod_len_from.Value & "," & txt_prod_len_to.Value & ",'" & _
                                   sUserID + "',?)}"
    
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
    Else
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

Public Sub Sp_Total()
    
    Dim j As Integer
    Dim iBas As Integer
    Dim iCot As Integer
    Dim iRow As Integer
    
    Dim sCol_a As String
    Dim sCol_b As String
    
    With SS1
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1
        
        Call Gp_Sp_BlockLock(SS1, 1, .MaxCols, .MaxRows, .MaxRows, True)
        Call Gp_Sp_BlockColor(SS1, 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HE6E6FF)
        
        For j = 1 To iSumCnt
            .Col = j
            If .ColHidden = False Then
                .Text = "合   计"
                j = iSumCnt
            End If
        Next j
        
        For j = 1 To iSumCnt
            .Col = iSumCol(j)
            
            If iSumCol(j) <= 26 Then
                sCol_a = Chr(iSumCol(j) + 64)
                .Formula = "sum(" + sCol_a + "1:" + sCol_a & .MaxRows - 1 & ")"
            Else
                iCot = Int(((iSumCol(j) - 1) / 26))
                iBas = 26 * iCot
                sCol_a = Chr((iSumCol(j) - iBas) + 64)
                sCol_b = Chr(iCot + 64)
                .Formula = "sum(" + sCol_b + sCol_a + "1:" + sCol_b + sCol_a & .MaxRows - 1 & ")"
            End If
        Next j
        
        For iRow = 1 To .MaxRows - 1
            .Row = iRow
            .Col = 13
            If Trim(.Text) <> "定尺" Then
                .Col = 14:    .Lock = False
            Else
                .Col = 14:    .Lock = True
                Call Gp_Sp_BlockColor(SS1, 14, 14, iRow, iRow, BLACK, WHITE)
            End If
        Next iRow
        
    End With
        
End Sub

Private Sub lbl_mplate_DblClick(Index As Integer)

    Dim sSeq As String
    
    If Index < 10 Then
        sSeq = "0" & Index
    Else
        sSeq = Trim(str(Index))
    End If
    
    If lbl_mplate(Index).BackColor = &HE0E0E0 Then
        lbl_mplate(Index).BackColor = &HC0C0C0
        lbl_mplate(Index).ForeColor = &HFF0000
        lbl_mplate(Index).Caption = sSeq
        
        If lbl_mplate(Index).Tag = "ord1" Then
            sdb_ord12_cnt = sdb_ord12_cnt - 1
        ElseIf lbl_mplate(Index).Tag = "ord2" Then
            sdb_ord22_cnt = sdb_ord22_cnt - 1
        ElseIf lbl_mplate(Index).Tag = "ord3" Then
            sdb_ord32_cnt = sdb_ord32_cnt - 1
        ElseIf lbl_mplate(Index).Tag = "ord4" Then
            sdb_ord42_cnt = sdb_ord42_cnt - 1
        ElseIf lbl_mplate(Index).Tag = "ord5" Then
            sdb_ord52_cnt = sdb_ord52_cnt - 1
        Else
            sdb_ord62_cnt = sdb_ord62_cnt - 1
        End If
    Else
        lbl_mplate(Index).BackColor = &HE0E0E0
        lbl_mplate(Index).ForeColor = &HFF0000
        lbl_mplate(Index).Caption = "删除"

        If lbl_mplate(Index).Tag = "ord1" Then
            sdb_ord12_cnt = sdb_ord12_cnt + 1
        ElseIf lbl_mplate(Index).Tag = "ord2" Then
            sdb_ord22_cnt = sdb_ord22_cnt + 1
        ElseIf lbl_mplate(Index).Tag = "ord3" Then
            sdb_ord32_cnt = sdb_ord32_cnt + 1
        ElseIf lbl_mplate(Index).Tag = "ord4" Then
            sdb_ord42_cnt = sdb_ord42_cnt + 1
        ElseIf lbl_mplate(Index).Tag = "ord5" Then
            sdb_ord52_cnt = sdb_ord52_cnt + 1
        Else
            sdb_ord62_cnt = sdb_ord62_cnt + 1
        End If
        
    End If
    
End Sub

Private Sub lbl_slab_DblClick(Index As Integer)

    Dim sSeq As String
    
    If Index < 10 Then
        sSeq = "0" & Index
    Else
        sSeq = Trim(str(Index))
    End If
    
    If lbl_slab(Index).BackColor = &HC0C0FF Then
        lbl_slab(Index).BackColor = &H8080FF
        lbl_slab(Index).ForeColor = &HFF0000
        lbl_slab(Index).Caption = sSeq
    Else
        lbl_slab(Index).BackColor = &HC0C0FF
        lbl_slab(Index).ForeColor = &HFF0000
        lbl_slab(Index).Caption = "删除"

    End If
    
End Sub

'Private Sub opt_sort1_Click(Value As Integer)
'
'    If opt_sort1.Value = True Then
'        txt_sort.Text = "1"
'        opt_sort1.ForeColor = &HFF&
'        opt_sort2.ForeColor = &H808080
'        opt_sort3.ForeColor = &H808080
'    End If
'
'End Sub
'
'Private Sub opt_sort2_Click(Value As Integer)
'
'    If opt_sort2.Value = True Then
'        txt_sort.Text = "2"
'        opt_sort2.ForeColor = &HFF&
'        opt_sort1.ForeColor = &H808080
'        opt_sort3.ForeColor = &H808080
'    End If
'
'End Sub
'
'Private Sub opt_sort3_Click(Value As Integer)
'
'    If opt_sort3.Value = True Then
'        txt_sort.Text = "3"
'        opt_sort3.ForeColor = &HFF&
'        opt_sort1.ForeColor = &H808080
'        opt_sort2.ForeColor = &H808080
'    End If
'
'End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim sTemp_ord As String
    Dim iRow As Integer
    Dim iCnt As Long
    Dim dWgt As Double
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If SS1.MaxRows < 1 Or Row < 1 Then Exit Sub
    
    SS1.Row = Row
    SS1.Col = 10
    If SS1.Text <> "PP" Then Exit Sub   'Only Plate Product
    
    SS1.Col = 0
    
    If SS1.Text = "" Then
    
        If oRd_cnt = 6 Then Exit Sub
        
        If txt_ord_no1.Text = "" Then
        
            If iSlab_cnt > 0 Then
                If First_Condition_Compare(Row) = False Then Exit Sub
            Else
                SS1.Col = 32
                sHTM_METH = SS1.Text
            End If
            
            SS1.Row = Row
            SS1.Col = 0
            SS1.Text = "选择"
            SS1.Col = 1
            txt_ord_no1.Text = SS1.Text
            SS1.Col = 2
            txt_ord_no1.Text = txt_ord_no1.Text & "-" & SS1.Text
            
           
            'PROD_LEN
            SS1.Col = 14
            sdb_ord1_len.Value = SS1.Value
            
            'PROD_WGT
            SS1.Col = 15
            dWgt = SS1.Value
            
            'DESIGN_REM_WGT / PROD_WGT
            SS1.Col = 29
            sdb_ord11_cnt.Value = (SS1.Value / dWgt)
            sdb_ord11_cnt.Value = Round(sdb_ord11_cnt.Value)
            sdb_ord12_cnt.Value = (SS1.Value / dWgt)
            sdb_ord12_cnt.Value = Round(sdb_ord12_cnt.Value)
            
            lMain_row = Row
            
            'Select Order1 Row
            iOrd1_Curr_Row = Row
            
        ElseIf txt_ord_no2.Text = "" Then
        
            If Condition_Compare(Row) = False Then Exit Sub
            
            SS1.Row = Row
            SS1.Col = 0
            SS1.Text = "选择"
            SS1.Col = 1
            txt_ord_no2.Text = SS1.Text
            SS1.Col = 2
            txt_ord_no2.Text = txt_ord_no2.Text & "-" & SS1.Text
            
            'PROD_LEN
            SS1.Col = 14
            sdb_ord2_len.Value = SS1.Value
            
            'PROD_WGT
            SS1.Col = 15
            dWgt = SS1.Value
            
            'DESIGN_REM_WGT / PROD_WGT
            SS1.Col = 29
            sdb_ord21_cnt.Value = (SS1.Value / dWgt)
            sdb_ord21_cnt.Value = Round(sdb_ord21_cnt.Value)
            sdb_ord22_cnt.Value = (SS1.Value / dWgt)
            sdb_ord22_cnt.Value = Round(sdb_ord22_cnt.Value)
            
            'Select Order2 Row
            iOrd2_Curr_Row = Row
        
        ElseIf txt_ord_no3.Text = "" Then
        
            If Condition_Compare(Row) = False Then Exit Sub
            
            SS1.Row = Row
            SS1.Col = 0
            SS1.Text = "选择"
            SS1.Col = 1
            txt_ord_no3.Text = SS1.Text
            SS1.Col = 2
            txt_ord_no3.Text = txt_ord_no3.Text & "-" & SS1.Text
            
            'PROD_LEN
            SS1.Col = 14
            sdb_ord3_len.Value = SS1.Value
            
            'PROD_WGT
            SS1.Col = 15
            dWgt = SS1.Value
            
            'DESIGN_REM_WGT / PROD_WGT
            SS1.Col = 29
            sdb_ord31_cnt.Value = (SS1.Value / dWgt)
            sdb_ord31_cnt.Value = Round(sdb_ord31_cnt.Value)
            sdb_ord32_cnt.Value = (SS1.Value / dWgt)
            sdb_ord32_cnt.Value = Round(sdb_ord32_cnt.Value)
            
            'Select Order3 Row
            iOrd3_Curr_Row = Row
            
        ElseIf txt_ord_no4.Text = "" Then
        
            If Condition_Compare(Row) = False Then Exit Sub
            
            SS1.Row = Row
            SS1.Col = 0
            SS1.Text = "选择"
            SS1.Col = 1
            txt_ord_no4.Text = SS1.Text
            SS1.Col = 2
            txt_ord_no4.Text = txt_ord_no4.Text & "-" & SS1.Text
            
            'PROD_LEN
            SS1.Col = 14
            sdb_ord4_len.Value = SS1.Value
            
            'PROD_WGT
            SS1.Col = 15
            dWgt = SS1.Value
            
            'DESIGN_REM_WGT / PROD_WGT
            SS1.Col = 29
            sdb_ord41_cnt.Value = (SS1.Value / dWgt)
            sdb_ord41_cnt.Value = Round(sdb_ord41_cnt.Value)
            sdb_ord42_cnt.Value = (SS1.Value / dWgt)
            sdb_ord42_cnt.Value = Round(sdb_ord42_cnt.Value)
            
            'Select Order4 Row
            iOrd4_Curr_Row = Row
            
        ElseIf txt_ord_no5.Text = "" Then
        
            If Condition_Compare(Row) = False Then Exit Sub
            
            SS1.Row = Row
            SS1.Col = 0
            SS1.Text = "选择"
            SS1.Col = 1
            txt_ord_no5.Text = SS1.Text
            SS1.Col = 2
            txt_ord_no5.Text = txt_ord_no5.Text & "-" & SS1.Text
            
            'PROD_LEN
            SS1.Col = 14
            sdb_ord5_len.Value = SS1.Value
            
            'PROD_WGT
            SS1.Col = 15
            dWgt = SS1.Value
            
            'DESIGN_REM_WGT / PROD_WGT
            SS1.Col = 29
            sdb_ord51_cnt.Value = (SS1.Value / dWgt)
            sdb_ord51_cnt.Value = Round(sdb_ord51_cnt.Value)
            sdb_ord52_cnt.Value = (SS1.Value / dWgt)
            sdb_ord52_cnt.Value = Round(sdb_ord52_cnt.Value)
            
            'Select Order5 Row
            iOrd5_Curr_Row = Row
            
        Else
        
            If Condition_Compare(Row) = False Then Exit Sub
            
            SS1.Row = Row
            SS1.Col = 0
            SS1.Text = "选择"
            SS1.Col = 1
            txt_ord_no6.Text = SS1.Text
            SS1.Col = 2
            txt_ord_no6.Text = txt_ord_no6.Text & "-" & SS1.Text
            
            'PROD_LEN
            SS1.Col = 14
            sdb_ord6_len.Value = SS1.Value
            
            'PROD_WGT
            SS1.Col = 15
            dWgt = SS1.Value
            
            'DESIGN_REM_WGT / PROD_WGT
            SS1.Col = 29
            sdb_ord61_cnt.Value = (SS1.Value / dWgt)
            sdb_ord61_cnt.Value = Round(sdb_ord61_cnt.Value)
            sdb_ord62_cnt.Value = (SS1.Value / dWgt)
            sdb_ord62_cnt.Value = Round(sdb_ord62_cnt.Value)
            
            'Select Order6 Row
            iOrd6_Curr_Row = Row
            
        End If
        
        Call Gp_Sp_BlockColor(SS1, 1, SS1.MaxCols, Row, Row, , &HFFFF80)
        oRd_cnt = oRd_cnt + 1
    
    Else
    
        If iMplate_cnt > 0 Then Exit Sub
        
        SS1.Text = ""
        
        SS1.Col = 1
        sTemp_ord = SS1.Text
        SS1.Col = 2
        sTemp_ord = sTemp_ord & "-" & SS1.Text
        
        If txt_ord_no1.Text = sTemp_ord Then
            txt_ord_no1.Text = ""
            txt_ord_no2.Text = ""
            txt_ord_no3.Text = ""
            txt_ord_no4.Text = ""
            txt_ord_no5.Text = ""
            txt_ord_no6.Text = ""
            sdb_ord11_cnt.Value = 0
            sdb_ord12_cnt.Value = 0
            sdb_ord21_cnt.Value = 0
            sdb_ord22_cnt.Value = 0
            sdb_ord31_cnt.Value = 0
            sdb_ord32_cnt.Value = 0
            sdb_ord41_cnt.Value = 0
            sdb_ord42_cnt.Value = 0
            sdb_ord51_cnt.Value = 0
            sdb_ord52_cnt.Value = 0
            sdb_ord61_cnt.Value = 0
            sdb_ord62_cnt.Value = 0
            sdb_ord1_len.Value = 0
            sdb_ord2_len.Value = 0
            sdb_ord3_len.Value = 0
            sdb_ord4_len.Value = 0
            sdb_ord5_len.Value = 0
            sdb_ord6_len.Value = 0
            sdb_asroll_thk.Value = 0
            sdb_asroll_wid.Value = 0
            oRd_cnt = 1
            lMain_row = 0
            iOrd1_Curr_Row = 0
            iOrd2_Curr_Row = 0
            iOrd3_Curr_Row = 0
            iOrd4_Curr_Row = 0
            iOrd5_Curr_Row = 0
            iOrd6_Curr_Row = 0
            
            For iRow = 1 To SS1.MaxRows
                SS1.Row = iRow
                SS1.Col = 0
                SS1.Text = ""
                Call Gp_Sp_BlockColor(SS1, 1, SS1.MaxCols, iRow, iRow)
            Next iRow
            
            For iCnt = 1 To iMplate_cnt
                Unload lbl_mplate(iCnt)
            Next iCnt
            iMplate_cnt = 0
        ElseIf txt_ord_no2.Text = sTemp_ord Then
            txt_ord_no2.Text = ""
            sdb_ord21_cnt.Value = 0
            sdb_ord22_cnt.Value = 0
            sdb_ord2_len.Value = 0
            iOrd2_Curr_Row = 0
        ElseIf txt_ord_no3.Text = sTemp_ord Then
            txt_ord_no3.Text = ""
            sdb_ord31_cnt.Value = 0
            sdb_ord32_cnt.Value = 0
            sdb_ord3_len.Value = 0
            iOrd3_Curr_Row = 0
        ElseIf txt_ord_no4.Text = sTemp_ord Then
            txt_ord_no4.Text = ""
            sdb_ord41_cnt.Value = 0
            sdb_ord42_cnt.Value = 0
            sdb_ord4_len.Value = 0
            iOrd4_Curr_Row = 0
        ElseIf txt_ord_no5.Text = sTemp_ord Then
            txt_ord_no5.Text = ""
            sdb_ord51_cnt.Value = 0
            sdb_ord52_cnt.Value = 0
            sdb_ord5_len.Value = 0
            iOrd5_Curr_Row = 0
        Else
            txt_ord_no6.Text = ""
            sdb_ord61_cnt.Value = 0
            sdb_ord62_cnt.Value = 0
            sdb_ord6_len.Value = 0
            iOrd6_Curr_Row = 0
        End If
            
        Call Gp_Sp_BlockColor(SS1, 1, SS1.MaxCols, Row, Row)
        Call SS1_CHANGE_COLOR
        oRd_cnt = oRd_cnt - 1
    
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
        Set Active_Spread = Me.SS1
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If

End Sub

Private Sub SSCommand1_Click()

 Call Gp_Process_Exec("1")
 
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
        
    Else

        If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
            txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
        Else
            txt_plt_name.Text = ""
        End If

    End If
    
End Sub

Public Sub Sp_Setting(ByVal sPname As Variant)

    With sPname
    
        .RowHeight(-1) = 12.54
        .RowHeight(0) = 24
        
        .ColWidth(0) = 9.5
        
        .ColWidth(1) = 13
        .ColWidth(2) = 13
        .ColWidth(3) = 13
        .ColWidth(4) = 13
        
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
        
    End With
    
End Sub

Private Sub txt_prod_cd_DblClick()

    Call txt_prod_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_prod_cd
        DD.rControl.Add Item:=txt_prod_cd_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else

        If Len(Trim(txt_prod_cd.Text)) = txt_prod_cd.MaxLength Then
            txt_prod_cd_name.Text = Gf_ComnNameFind(M_CN1, "B0005", Trim(txt_prod_cd.Text), 2)
        Else
            txt_prod_cd_name.Text = ""
        End If
    
    End If
    
End Sub

'Private Sub txt_sort_Change()
'
'    If txt_sort.Text = "" Then opt_sort1.Value = True
'
'End Sub

Private Function Condition_Compare(iRow As Long) As Boolean

    Dim sTemp   As String
    Dim sOrd1OrdNo   As String
    Dim sOrd1OrdItem   As String
    Dim sCurOrdNo   As String
    Dim sCurOrdItem   As String
    Dim sQuery   As String
    Dim sMessage As String
    Dim dTemp   As Double
    Dim dWidMin As Double
    Dim dWidMax As Double
    
    Condition_Compare = True
    
    'STLGRD,THK,WID,TRIM_FL以外 不要检察 2005.11.11
'---------------------------------------------------------------------------
    
    'STLGRD
    SS1.Row = iOrd1_Curr_Row
    SS1.Col = 7
    sTemp = SS1.Text
    SS1.Row = iRow

    If sTemp <> SS1.Text Then
        Call Gp_MsgBoxDisplay("钢种不一致")
        Condition_Compare = False
        Exit Function
    End If

    'PROD_THK
    SS1.Row = iOrd1_Curr_Row
    SS1.Col = 11
    dTemp = SS1.Value
    SS1.Row = iRow

    If dTemp <> SS1.Value Then
        Call Gp_MsgBoxDisplay("厚度不一致")
        Condition_Compare = False
        Exit Function
    End If
'
    
    '---------------------------------------------------------------------------
    'PROD_WID
'    Call Range_Wid(iRow, dWidMin, dWidMax)
'
'    ss1.Row = iOrd1_Curr_Row
'    ss1.Col = 9
'    dTemp = ss1.Value
    
'    ss1.Row = iRow
'    comment by yangmeng at 081208
'    If dTemp < dWidMin Or dTemp > dWidMax Then
'        Call Gp_MsgBoxDisplay("宽度不一致")
'        Condition_Compare = False
'        Exit Function
'    End If

        
    'ORD_TRIM_FL
    SS1.Row = iOrd1_Curr_Row
    SS1.Col = 21
    sTemp = SS1.Text
    SS1.Row = iRow

    If sTemp <> SS1.Text Then
        Call Gp_MsgBoxDisplay("切边不一致")
        Condition_Compare = False
        Exit Function
    End If
'
'    'HTM_METH
'    ss1.Row = iRow
'    ss1.Col = 29
'
'    If sHTM_METH = "" Then
'        If ss1.Text <> "" Then
'            Call Gp_MsgBoxDisplay("热处理不一致")
'            Condition_Compare = False
'            Exit Function
'        End If
'    Else
'        If ss1.Text = "" Then
'            Call Gp_MsgBoxDisplay("热处理不一致")
'            Condition_Compare = False
'            Exit Function
'        End If
'    End If
    
'---------------------------------------------------------------------------




'    'MLT_PROC_CD
'    ss1.Row = iOrd1_Curr_Row
'    ss1.Col = 14
'    sTemp = ss1.Text
'    ss1.Row = iRow
'
'    If sTemp <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("工艺流程不一致")
'        Condition_Compare = False
'        Exit Function
'    End If
'
'    'CUST_SPEC_NO
'    ss1.Row = iOrd1_Curr_Row
'    ss1.Col = 21
'    sTemp = ss1.Text
'    ss1.Row = iRow
'
'    If sTemp <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("客户要求特殊编号不一致")
'        Condition_Compare = False
'        Exit Function
'    End If
'
'    'ENDUSE_CD
'    ss1.Row = iOrd1_Curr_Row
'    ss1.Col = 5
'    sTemp = ss1.Text
'    ss1.Row = iRow
'
'    If sTemp <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("用途不一致")
'        Condition_Compare = False
'        Exit Function
'    End If
'
'    'STDSPEC
'    ss1.Row = iOrd1_Curr_Row
'    ss1.Col = 16
'    sTemp = ss1.Text
'    ss1.Row = iRow
'
'    If sTemp <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("标准代号不一致")
'        Condition_Compare = False
'        Exit Function
'    End If
'
'    'ISP_CMP
'    ss1.Row = iOrd1_Curr_Row
'    ss1.Col = 17
'    sTemp = ss1.Text
'    ss1.Row = iRow
'
'    If sTemp <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("检查机关不一致")
'        Condition_Compare = False
'        Exit Function
'    End If
'
'    'ORD_HCR_FL
'    ss1.Row = iOrd1_Curr_Row
'    ss1.Col = 15
'    sTemp = ss1.Text
'    ss1.Row = iRow
'
'    If sTemp <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("H/C 不一致")
'        Condition_Compare = False
'        Exit Function
'    End If
'
'    'CR_CD
'    ss1.Row = iOrd1_Curr_Row
'    ss1.Col = 18
'    sTemp = ss1.Text
'    ss1.Row = iRow
'
'    If sTemp <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("控轧不一致")
'        Condition_Compare = False
'        Exit Function
'    End If
'
'    'UST_FL
'    ss1.Row = iOrd1_Curr_Row
'    ss1.Col = 20
'    sTemp = ss1.Text
'    ss1.Row = iRow
'
'    If sTemp <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("UST 不一致")
'        Condition_Compare = False
'        Exit Function
'    End If
    
End Function

Private Function First_Condition_Compare(iRow As Long) As Boolean

    Dim sTemp   As String
    Dim sCurOrdNo As String
    Dim sCurOrdItem As String
    Dim sQuery As String
    Dim sMessage As String
    Dim dTemp   As Double
    Dim dWidMin As Double
    Dim dWidMax As Double
    
    First_Condition_Compare = True
    SS1.Row = iRow
    
    'STLGRD
    SS1.Col = 7
    If vSTLGRD <> SS1.Text Then
        Call Gp_MsgBoxDisplay("钢种不一致")
        First_Condition_Compare = False
        Exit Function
    End If
    
    'PROD_THK
    SS1.Col = 11
    If vPROD_THK <> SS1.Value Then
        Call Gp_MsgBoxDisplay("厚度不一致")
        First_Condition_Compare = False
        Exit Function
    End If

    'ORD_TRIM_FL
    SS1.Col = 21
    If vORD_TRIM_FL <> SS1.Text Then
        Call Gp_MsgBoxDisplay("切边不一致")
        First_Condition_Compare = False
        Exit Function
    End If
'
'    'HTM_METH
'    ss1.Col = 29
'    If sHTM_METH = "" Then
'        If ss1.Text <> "" Then
'            Call Gp_MsgBoxDisplay("热处理不一致")
'            First_Condition_Compare = False
'            Exit Function
'        End If
'    Else
'        If ss1.Text = "" Then
'            Call Gp_MsgBoxDisplay("热处理不一致")
'            First_Condition_Compare = False
'            Exit Function
'        End If
'    End If
    
'---------------------------------------------------------------------------

    
'    'MLT_PROC_CD
'    ss1.Col = 14
'    If vMLT_PROC_CD <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("工艺流程不一致")
'        First_Condition_Compare = False
'        Exit Function
'    End If
'
'    'CUST_SPEC_NO
'    ss1.Col = 21
'    If vCUST_SPEC_NO <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("客户要求特殊编号不一致")
'        First_Condition_Compare = False
'        Exit Function
'    End If
'
'    'ENDUSE_CD
'    ss1.Col = 5
'    If vENDUSE_CD <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("用途不一致")
'        First_Condition_Compare = False
'        Exit Function
'    End If
'
'    'STDSPEC
'    ss1.Col = 16
'    If vSTDSPEC <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("标准代号不一致")
'        First_Condition_Compare = False
'        Exit Function
'    End If
'
'    'ISP_CMP
'    ss1.Col = 17
'    If vISP_CMP <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("检查机关不一致")
'        First_Condition_Compare = False
'        Exit Function
'    End If
'
'    'ORD_HCR_FL
'    ss1.Col = 15
'    If vORD_HCR_FL <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("H/C 不一致")
'        First_Condition_Compare = False
'        Exit Function
'    End If
'
'    'CR_CD
'    ss1.Col = 18
'    If vCR_CD <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("控轧不一致")
'        First_Condition_Compare = False
'        Exit Function
'    End If
'
'    'UST_FL
'    ss1.Col = 20
'    If vUST_FL <> ss1.Text Then
'        Call Gp_MsgBoxDisplay("UST 不一致")
'        First_Condition_Compare = False
'        Exit Function
'    End If
'
End Function

Private Sub Plate_Block_Seq_Create(Current_Row As Variant, iType As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim P_SLAB_EDT_SEQ As Long
    Dim adoCmd As adodb.Command
    Dim lSlab_Edt_Seq As Double
    
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
    
    'SLAB_EDT_SEQ Setting
    If txt_ccm_line.Text = "1" Then
        lSlab_Edt_Seq = 99999010
    ElseIf txt_ccm_line.Text = "2" Then
        lSlab_Edt_Seq = 99999020
    Else
        lSlab_Edt_Seq = 99999030
    End If
    
    SS1.Row = Current_Row
    
    'SLAB_EDT_SEQ, BLOCK_SEQ, SEQ
    sQuery = "{call AEB1090C.P_MODIFY1 ('" + iType + "'," & lSlab_Edt_Seq & ",'99','00',"
    
    'ORD_NO
    SS1.Col = 1
    sQuery = sQuery + "'" + SS1.Text + "',"
    
    'ORD_ITEM
    SS1.Col = 2
    sQuery = sQuery + "'" + SS1.Text + "',"
    
    'PROD_CD
    SS1.Col = 10
    sQuery = sQuery + "'" + SS1.Text + "',"
        
    'STLGRD
    SS1.Col = 7
    sQuery = sQuery + "'" + SS1.Text + "',"
    
    'THK
    sQuery = sQuery & sdb_asroll_thk.Value & ","
    
    'WID
    sQuery = sQuery & sdb_asroll_wid.Value & ","
    
    'LEN
    sQuery = sQuery & sdb_asroll_len.Value & ","
    
    'WGT
    SS1.Col = 15
    sQuery = sQuery & SS1.Value & ","
    
    'CR_CD
    SS1.Col = 20
    sQuery = sQuery + "'" + SS1.Text + "',"
    
    'UST_FL
    SS1.Col = 22
    sQuery = sQuery + "'" + SS1.Text + "',"
    
    'TRIM_FL
    SS1.Col = 21
    sQuery = sQuery + "'" + SS1.Text + "',?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
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
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub Plate_Seq_Create(Current_Row As Variant, Seq As String, iType As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim lSlab_Edt_Seq As Double
    
    Dim adoCmd As adodb.Command
    
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
    
    'SLAB_EDT_SEQ Setting
    If txt_ccm_line.Text = "1" Then
        lSlab_Edt_Seq = 99999010
    ElseIf txt_ccm_line.Text = "2" Then
        lSlab_Edt_Seq = 99999020
    Else
        lSlab_Edt_Seq = 99999030
    End If
    
    SS1.Row = Current_Row
    
    'SLAB_EDT_SEQ, BLOCK_SEQ, SEQ
    sQuery = "{call AEB1090C.P_MODIFY1 ('" + iType + "'," & lSlab_Edt_Seq & ",'99','" + Seq + "',"
    
    'ORD_NO
    SS1.Col = 1
    sQuery = sQuery + "'" + SS1.Text + "',"
    
    'ORD_ITEM
    SS1.Col = 2
    sQuery = sQuery + "'" + SS1.Text + "',"
    
    'PROD_CD
    SS1.Col = 10
    sQuery = sQuery + "'" + SS1.Text + "',"
        
    'STLGRD
    SS1.Col = 7
    sQuery = sQuery + "'" + SS1.Text + "',"
    
    'THK
    SS1.Col = 11
    sQuery = sQuery & SS1.Value + ","
    
    'WID
    SS1.Col = 12
    sQuery = sQuery & SS1.Value + ","
    
    'LEN
    SS1.Col = 14
    sQuery = sQuery & SS1.Value + ","
    
    'WGT
    SS1.Col = 15
    sQuery = sQuery & SS1.Value & ","
    
    'CR_CD
    SS1.Col = 20
    sQuery = sQuery + "'" + SS1.Text + "',"
    
    'UST_FL
    SS1.Col = 22
    sQuery = sQuery + "'" + SS1.Text + "',"
    
    'TRIM_FL
    SS1.Col = 21
    sQuery = sQuery + "'" + SS1.Text + "',?,?)}"
    
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
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
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)

End Sub

Private Sub Slab_Block_Seq_Create(iType As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim P_SLAB_EDT_SEQ As Long
    Dim adoCmd As adodb.Command
    
    Screen.MousePointer = vbHourglass
    
    'Max SLAB_EDT_SEQ READ
    sQuery = "SELECT MAX(SLAB_EDT_SEQ) FROM EP_SLAB_EDT_CSL WHERE CCM_PRC_LINE = '" + txt_ccm_line.Text + "'"
    P_SLAB_EDT_SEQ = Gf_FloatFind(M_CN1, sQuery)
    
    If P_SLAB_EDT_SEQ = 0 Then
        If txt_ccm_line.Text = "1" Then
            P_SLAB_EDT_SEQ = 0
        ElseIf txt_ccm_line.Text = "2" Then
            P_SLAB_EDT_SEQ = 30000000
        Else
            P_SLAB_EDT_SEQ = 50000000
        End If
    
    End If
    
    P_SLAB_EDT_SEQ = P_SLAB_EDT_SEQ + 1
    iSLAB_EDT_SEQ = P_SLAB_EDT_SEQ
    
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
    
    'SLAB_EDT_SEQ, BLOCK_SEQ, SEQ
    sQuery = "{call AEB1090C.P_MODIFY2 ('" + iType + "','" + txt_ccm_line.Text + "'," & P_SLAB_EDT_SEQ & ",'00','00',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
        
    sQuery = sQuery + "'',"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
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
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub Slab_Seq_Create(Seq As String, iType As String)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As adodb.Command
    
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
    
    'SLAB_EDT_SEQ, BLOCK_SEQ, SEQ
    sQuery = "{call AEB1090C.P_MODIFY2 ('" + iType + "','" + txt_ccm_line.Text + "'," & iSLAB_EDT_SEQ & ",'" + Seq + "','00',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
        
    sQuery = sQuery + "'',"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery & "0,"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',"
    
    sQuery = sQuery + "'',?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
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
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)

End Sub

Public Sub Asroll_Thk(sOrderNo As String)

    Dim sQuery As String
    
    'Asroll Thk
    sQuery = "         SELECT  MILL_TGT_THK "
    sQuery = sQuery + "  FROM  NISCO.QP_QLTY_TECH "
    sQuery = sQuery + " WHERE  ORD_NO    = '" & Mid(sOrderNo, 1, 11) & "' "
    sQuery = sQuery + "   AND  ORD_ITEM  = '" & Mid(sOrderNo, 13, 2) & "' "
    sQuery = sQuery + "   AND  KND       = (SELECT  MAX(KND) "
    sQuery = sQuery + "                       FROM  NISCO.QP_QLTY_TECH "
    sQuery = sQuery + "                      WHERE  ORD_NO    = '" & Mid(sOrderNo, 1, 11) & "' "
    sQuery = sQuery + "                        AND  ORD_ITEM  = '" & Mid(sOrderNo, 13, 2) & "') "
    
    sdb_asroll_thk.Value = Gf_FloatFind(M_CN1, sQuery)
        
End Sub

Public Sub Asroll_Wid(sOrderNo As String)

    Dim sQuery As String
    
    'Asroll Wid
    sQuery = "         SELECT  MILL_TGT_WID "
    sQuery = sQuery + "  FROM  NISCO.QP_QLTY_TECH "
    sQuery = sQuery + " WHERE  ORD_NO    = '" & Mid(sOrderNo, 1, 11) & "' "
    sQuery = sQuery + "   AND  ORD_ITEM  = '" & Mid(sOrderNo, 13, 2) & "' "
    sQuery = sQuery + "   AND  KND       = (SELECT  MAX(KND) "
    sQuery = sQuery + "                       FROM  NISCO.QP_QLTY_TECH "
    sQuery = sQuery + "                      WHERE  ORD_NO    = '" & Mid(sOrderNo, 1, 11) & "' "
    sQuery = sQuery + "                        AND  ORD_ITEM  = '" & Mid(sOrderNo, 13, 2) & "') "
    
    sdb_asroll_wid.Value = Gf_FloatFind(M_CN1, sQuery)
        
End Sub

Public Sub Range_Wid(iRow As Long, dWidMin As Double, dWidMax As Double)

    Dim sQuery As String
    Dim dWid   As Double
    
    Set AdoRs = New adodb.Recordset
    
    'Asroll Wid
    SS1.Row = iRow
    sQuery = "         SELECT  WID_TOL_MIN, WID_TOL_MAX "
    sQuery = sQuery + "  FROM  NISCO.QP_QLTY_DELV "
    SS1.Col = 1
    sQuery = sQuery + " WHERE  ORD_NO    = '" & Trim(SS1.Text) & "' "
    SS1.Col = 2
    sQuery = sQuery + "   AND  ORD_ITEM  = '" & Trim(SS1.Text) & "' "
    sQuery = sQuery + "   AND  KND       = '4'"
    
    AdoRs.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    SS1.Col = 12
    dWid = Val(Format(SS1.Value, "###0.000") & "")
    
    dWidMin = dWid + Val(AdoRs(0) & "")
    dWidMax = dWid + Val(AdoRs(1) & "")
    
End Sub

Public Sub Slab_Size()

On Error GoTo Slab_Size_Error

    Dim sQuery As String
    Dim AdoRs As adodb.Recordset
    Set AdoRs = New adodb.Recordset
    
    'SLAB THK, WID, LEN, WGT
    sQuery = "         SELECT  NVL(SLAB_THK,0), NVL(SLAB_WID,0), NVL(SLAB_LEN,0), NVL(SLAB_WGT,0), NVL(DESIGN_RATIO,0) "
    sQuery = sQuery + "  FROM  NISCO.EP_SLAB_EDT_CSL "
    sQuery = sQuery + " WHERE  SLAB_EDT_SEQ  =  " & iSLAB_EDT_SEQ
    
    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            sdb_slab_thk1.Value = AdoRs.Fields(0)
            sdb_slab_wid1.Value = AdoRs.Fields(1)
            sdb_slab_len1.Value = AdoRs.Fields(2)
            sdb_slab_wgt1.Value = AdoRs.Fields(3)
            sdb_slab_ratio.Value = AdoRs.Fields(4)
        End If
        
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    Exit Sub

Slab_Size_Error:

    Call Gp_MsgBoxDisplay("Slab_Size Error : " & Error)
    Set AdoRs = Nothing

End Sub

Public Sub Plate_Size()

On Error GoTo Plate_Size_Error

    Dim sQuery As String
    Dim AdoRs As adodb.Recordset
    Set AdoRs = New adodb.Recordset
    
    'PLATE THK, WID, LEN, WGT
    sQuery = "         SELECT  NVL(LEN,0), NVL(THK,0), NVL(WID,0) "
    sQuery = sQuery + "  FROM  NISCO.EP_PLATE_EDT_CSL "
    sQuery = sQuery + " WHERE  SLAB_EDT_SEQ  =  " & iSLAB_EDT_SEQ
    sQuery = sQuery + "   AND  BLOCK_SEQ     =  '00' "
    sQuery = sQuery + "   AND  SEQ           =  '00' "
    
    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            sdb_slab_len.Value = AdoRs.Fields(0)
            sdb_slab_thk.Value = AdoRs.Fields(1)
            sdb_slab_wid.Value = AdoRs.Fields(2)
        End If
        
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    Exit Sub

Plate_Size_Error:

    Call Gp_MsgBoxDisplay("Plate_Size Error : " & Error)
    Set AdoRs = Nothing

End Sub

Private Function Plate_Setting_Check(sOrd_No As String, sOrd_item As String) As Boolean

    Dim sQuery As String
    Dim sMessage As String
    Dim fOrd_No As String
    Dim fOrd_Item As String
    Dim iPlate_Cnt As Integer
    Dim lSlab_Edt_Seq As Long
    
    Plate_Setting_Check = True
    
'    For iPlate_Cnt = iMplate_cnt To 1 Step -1
'
'        If lbl_mplate(iPlate_Cnt).Visible Then
'
'            If lbl_mplate(iPlate_Cnt).Tag = "ord1" Then
'                fOrd_No = Mid(txt_ord_no1.Text, 1, 11)
'                fOrd_Item = Mid(txt_ord_no1.Text, 13, 2)
'            ElseIf lbl_mplate(iPlate_Cnt).Tag = "ord2" Then
'                fOrd_No = Mid(txt_ord_no2.Text, 1, 11)
'                fOrd_Item = Mid(txt_ord_no2.Text, 13, 2)
'            ElseIf lbl_mplate(iPlate_Cnt).Tag = "ord3" Then
'                fOrd_No = Mid(txt_ord_no3.Text, 1, 11)
'                fOrd_Item = Mid(txt_ord_no3.Text, 13, 2)
'            ElseIf lbl_mplate(iPlate_Cnt).Tag = "ord4" Then
'                fOrd_No = Mid(txt_ord_no4.Text, 1, 11)
'                fOrd_Item = Mid(txt_ord_no4.Text, 13, 2)
'            ElseIf lbl_mplate(iPlate_Cnt).Tag = "ord5" Then
'                fOrd_No = Mid(txt_ord_no5.Text, 1, 11)
'                fOrd_Item = Mid(txt_ord_no5.Text, 13, 2)
'            ElseIf lbl_mplate(iPlate_Cnt).Tag = "ord6" Then
'                fOrd_No = Mid(txt_ord_no6.Text, 1, 11)
'                fOrd_Item = Mid(txt_ord_no6.Text, 13, 2)
'            End If
'
'            Exit For
'
'        End If
'
'    Next iPlate_Cnt

    If txt_ccm_line.Text = "1" Then
        lSlab_Edt_Seq = 99999010
    ElseIf txt_ccm_line.Text = "2" Then
        lSlab_Edt_Seq = 99999020
    Else
        lSlab_Edt_Seq = 99999030
    End If

    sQuery = " SELECT ORD_NO FROM EP_PLATE_EDT_CSL WHERE SLAB_EDT_SEQ = " & lSlab_Edt_Seq & " AND BLOCK_SEQ = '99' AND SEQ = '01' "
    fOrd_No = Gf_CodeFind(M_CN1, sQuery)

    sQuery = " SELECT ORD_ITEM FROM EP_PLATE_EDT_CSL WHERE SLAB_EDT_SEQ = " & lSlab_Edt_Seq & " AND BLOCK_SEQ = '99' AND SEQ = '01' "
    fOrd_Item = Gf_CodeFind(M_CN1, sQuery)

    If fOrd_No = "" Then
        Exit Function
    End If

    sQuery = " SELECT GF_HMI_DESIGN_ORDER_CHECK_CSL('C1','" + fOrd_No + "','" + fOrd_Item + "','" + sOrd_No + "','" + sOrd_item + "') FROM DUAL "
    sMessage = Gf_CodeFind(M_CN1, sQuery)

    If sMessage <> "" Then
        Call Gp_MsgBoxDisplay(sMessage)
        Plate_Setting_Check = False
        Exit Function
    End If

End Function


