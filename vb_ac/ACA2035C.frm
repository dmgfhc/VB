VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACA2035C 
   Caption         =   "特殊质量信息详细查询界面_ACA2035C"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   3030
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   5345
      _Version        =   196609
      AutoSize        =   1
      PaneTree        =   "ACA2035C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   675
         Left            =   30
         TabIndex        =   3
         Top             =   2325
         Width           =   4500
         _Version        =   393216
         _ExtentX        =   7937
         _ExtentY        =   1191
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
         MaxCols         =   57
         MaxRows         =   5
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACA2035C.frx":0052
      End
      Begin Threed.SSFrame FRAME1 
         Height          =   2205
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   3889
         _Version        =   196609
         BackColor       =   14737632
         ShadowStyle     =   1
         Begin VB.TextBox Text_BB_DOME_FL_mate 
            Height          =   270
            Left            =   9480
            TabIndex        =   37
            Top             =   1800
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox Text_BB_REC_STS_Name 
            Height          =   270
            Left            =   9240
            TabIndex        =   36
            Top             =   1800
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.TextBox Text_BB_PROD_CD_mate 
            Height          =   270
            Left            =   8400
            TabIndex        =   35
            Top             =   1800
            Visible         =   0   'False
            Width           =   735
         End
         Begin Threed.SSCheck CHK_SPEC 
            Height          =   375
            Left            =   10800
            TabIndex        =   28
            Top             =   0
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   196609
            BackColor       =   16777152
            Caption         =   "特殊尺寸外形要求"
         End
         Begin VB.TextBox Text_PONO 
            Height          =   315
            Left            =   3660
            TabIndex        =   24
            Top             =   495
            Width           =   1350
         End
         Begin VB.TextBox txt_cust_cd_name 
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
            Left            =   8145
            MaxLength       =   40
            TabIndex        =   23
            Tag             =   "客户"
            Top             =   1290
            Width           =   2235
         End
         Begin VB.TextBox txt_item_count 
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
            Left            =   5220
            MaxLength       =   3
            TabIndex        =   22
            Text            =   "99"
            Top             =   120
            Width           =   450
         End
         Begin VB.TextBox txt_cfm_mill_plt 
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
            TabIndex        =   21
            Top             =   120
            Width           =   540
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
            TabIndex        =   20
            Tag             =   "钢种(标准号)"
            Top             =   900
            Width           =   3090
         End
         Begin VB.TextBox txt_sale_way_name 
            Height          =   315
            Left            =   4230
            TabIndex        =   19
            Top             =   1260
            Width           =   1425
         End
         Begin VB.TextBox Text_BB_PROD_CD 
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
            TabIndex        =   18
            Tag             =   "产品"
            Top             =   120
            Width           =   645
         End
         Begin VB.TextBox Text_BB_DOME_FL 
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
            TabIndex        =   17
            Top             =   885
            Width           =   645
         End
         Begin VB.TextBox txt_sale_way 
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
            Left            =   3660
            MaxLength       =   2
            TabIndex        =   16
            Top             =   1260
            Width           =   570
         End
         Begin VB.TextBox Text_BB_ORD_NO 
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
            TabIndex        =   15
            Top             =   885
            Width           =   1350
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
            Left            =   7290
            MaxLength       =   6
            TabIndex        =   14
            Top             =   1290
            Width           =   835
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
            Left            =   7290
            MaxLength       =   11
            TabIndex        =   13
            Tag             =   "钢种"
            Top             =   525
            Width           =   1320
         End
         Begin VB.TextBox Text_size_knd 
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
            MaxLength       =   2
            TabIndex        =   12
            Tag             =   "钢种"
            Top             =   1680
            Width           =   570
         End
         Begin VB.TextBox Text_size_knd_name 
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
            Left            =   4230
            TabIndex        =   11
            Tag             =   "钢种"
            Top             =   1680
            Width           =   1425
         End
         Begin VB.TextBox txt_ord_sts 
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
            Top             =   1650
            Width           =   645
         End
         Begin VB.ComboBox Combo_ORD_ITEM 
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
            TabIndex        =   9
            Top             =   885
            Width           =   660
         End
         Begin VB.TextBox Text_BB_REC_STS 
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
            TabIndex        =   8
            Top             =   495
            Width           =   645
         End
         Begin VB.TextBox TXT_ENDUSE_CD 
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
            Left            =   7290
            MaxLength       =   4
            TabIndex        =   7
            Tag             =   "CD_MANA_NO"
            Top             =   1695
            Width           =   645
         End
         Begin VB.TextBox txt_ord_knd 
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
            TabIndex        =   6
            Top             =   1260
            Width           =   645
         End
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
            TabIndex        =   5
            Top             =   525
            Width           =   1770
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
            Top             =   885
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "订单分类"
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
         Begin InDate.UDate UDate_BB_DEL_TO 
            Height          =   315
            Left            =   8940
            TabIndex        =   25
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
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   2370
            Top             =   1260
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "销售方式"
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
            Left            =   2370
            Top             =   885
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
            Left            =   120
            Top             =   1650
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
            Left            =   6000
            Top             =   525
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
         Begin InDate.ULabel ULabel14 
            Height          =   315
            Left            =   2370
            Top             =   1680
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "定尺区分"
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
         Begin InDate.ULabel ULabel01 
            Height          =   315
            Index           =   14
            Left            =   2370
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "生产工厂"
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
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16711680
         End
         Begin InDate.UDate Udate_BB_DEL_FR 
            Height          =   315
            Left            =   7290
            TabIndex        =   26
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
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   6000
            Top             =   1695
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "订单用途"
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
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   120
            Top             =   1260
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
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16711680
         End
         Begin InDate.ULabel ULabel19 
            Height          =   315
            Left            =   6000
            Top             =   900
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
         Begin InDate.ULabel ULabel20 
            Height          =   315
            Left            =   4560
            Top             =   120
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            Caption         =   "项次数"
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
            Left            =   6000
            Top             =   1290
            Width           =   1260
            _ExtentX        =   2223
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
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Left            =   2370
            Top             =   495
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "合同编号"
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
         Begin Threed.SSCheck CHK_OPRA 
            Height          =   375
            Left            =   10800
            TabIndex        =   29
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   196609
            BackColor       =   16777152
            Caption         =   "特殊作业要求"
            MaskColor       =   16777152
         End
         Begin Threed.SSCheck CHK_SURFACE 
            Height          =   375
            Left            =   10800
            TabIndex        =   30
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   196609
            BackColor       =   16777152
            Caption         =   "表面质量特殊要求"
            MaskColor       =   16777152
         End
         Begin Threed.SSCheck CHK_MARK 
            Height          =   375
            Left            =   10800
            TabIndex        =   31
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   196609
            BackColor       =   16777152
            Caption         =   "特殊标识要求"
            MaskColor       =   16777152
         End
         Begin Threed.SSCheck CHK_JIT 
            Height          =   375
            Left            =   10800
            TabIndex        =   32
            Top             =   1440
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   196609
            BackColor       =   16777152
            Caption         =   "认证要求"
            MaskColor       =   16777152
         End
         Begin Threed.SSCheck CHK_MEMO 
            Height          =   375
            Left            =   10800
            TabIndex        =   33
            Top             =   1800
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   196609
            BackColor       =   16777152
            Caption         =   "各类订单备注"
            MaskColor       =   16777152
         End
         Begin Threed.SSCheck CHK_UST 
            Height          =   375
            Left            =   12600
            TabIndex        =   34
            Top             =   0
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   196609
            BackColor       =   16777152
            Caption         =   "特殊探伤要求"
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            Height          =   120
            Left            =   8790
            TabIndex        =   27
            Top             =   210
            Width           =   90
         End
      End
   End
   Begin CSTextLibCtl.sidbEdit sidbEdit1 
      Height          =   315
      Left            =   13020
      TabIndex        =   1
      Tag             =   "订单长度"
      Top             =   1200
      Width           =   780
      _Version        =   262145
      _ExtentX        =   1376
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
   Begin InDate.ULabel ULabel24 
      Height          =   315
      Left            =   11880
      Top             =   1200
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "特殊长度"
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
   Begin CSTextLibCtl.sidbEdit sidbEdit2 
      Height          =   315
      Left            =   14010
      TabIndex        =   2
      Tag             =   "订单长度"
      Top             =   1200
      Width           =   780
      _Version        =   262145
      _ExtentX        =   1376
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
End
Attribute VB_Name = "ACA2035C"
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
'-- Program ID        ACA2035C
'-- Document No       Q-00-0010(Specification)
'-- Designer          WANGLEI
'-- Coder             WANGLEI
'-- Date              2016.09.27
'-- Description nnnn
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'  -------------------------------------------------------------------------------

Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting
Public ORD_NO As String             'Transfer to ACA1030C
Public ORD_ITEM As String           'Transfer to ACA1030C


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
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim sCheck1 As String
Dim sCheck2 As String
Dim iCount As Integer

Const iSumColCnt = 12
Const iSumCol1 = 24
Const iSumCol2 = 25
Const iSumCol3 = 26
Const iSumCol4 = 27
Const iSumCol5 = 28
Const iSumCol6 = 29
Const iSumCol7 = 30
Const iSumCol8 = 31
Const iSumCol9 = 32
Const iSumCol10 = 33
Const iSumCol11 = 34
Const iSumCol12 = 35

Const SS1_URGNT_FL = 49          '紧急订单绿色标记
Const SS1_ORD_NO = 1
Const SS1_ORD_ITEM = 2


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
 '   FormType = "Msheet"
    FormType = "Refer"

   'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(Text_BB_PROD_CD, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Text_BB_DOME_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Text_BB_REC_STS, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Udate_BB_DEL_FR, "p", "n", "m", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(UDate_BB_DEL_TO, "p", "n", "m", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(Text_BB_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(Combo_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_sale_way, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_ord_sts, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_cust_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(Text_size_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_cfm_mill_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_ord_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_ENDUSE_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_STDSPEC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_item_count, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(Text_PONO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(CHK_SPEC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(CHK_SURFACE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(CHK_OPRA, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(CHK_MEMO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(CHK_MARK, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(CHK_JIT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(CHK_UST, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
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
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '报表编制日期
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
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '计划投入工厂
    
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '生产工厂
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
    Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     
     
   'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"

    sc1.Add Item:="ACA2035C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").ROW = 0
    sc1.Item("Spread").Text = "◎"
    
    
    'Sum Column Count
    iSumCnt = iSumColCnt
    
    'Sum Column Setting
    iSumCol.Add Item:=iSumCol1
    iSumCol.Add Item:=iSumCol2
    iSumCol.Add Item:=iSumCol3
    iSumCol.Add Item:=iSumCol4
    iSumCol.Add Item:=iSumCol5
    iSumCol.Add Item:=iSumCol6
    iSumCol.Add Item:=iSumCol7
    iSumCol.Add Item:=iSumCol8
    iSumCol.Add Item:=iSumCol9
    iSumCol.Add Item:=iSumCol10
    iSumCol.Add Item:=iSumCol11
    iSumCol.Add Item:=iSumCol12
    
        
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub



Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

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
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    
    Text_BB_PROD_CD.Text = "PP"
    
    Udate_BB_DEL_FR.Text = Mid(DateAdd("m", -1, Date), 1, 8) + "01"

    UDate_BB_DEL_TO.Text = Format(DateAdd("m", 1, Udate_BB_DEL_FR.Text), "YYYY-MM-DD")
    UDate_BB_DEL_TO.Text = DateAdd("d", -1, UDate_BB_DEL_TO.Text)

    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
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
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If
    
    
    
    UDate_BB_DEL_TO.Text = DateAdd("m", -1, Now)   '本月
    UDate_BB_DEL_TO.Text = DateAdd("d", -1, UDate_BB_DEL_TO.Text)     '上月最后一天
    
    Udate_BB_DEL_FR.Text = Mid(UDate_BB_DEL_TO.Text, 1, 8) + "01"     '上月第一天
    
   
   
End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
End Sub

Public Sub Form_Ref()
    
    Dim sMesg As String
    Dim sQuery As String
    
    If UDate_BB_DEL_TO.RawData >= Udate_BB_DEL_FR.RawData Or UDate_BB_DEL_TO.RawData = "" Then
    
        If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
        
        sMesg = Gf_Ms_NeceCheck(nControl)
        If sMesg = "OK" Then
        
            sMesg = Gf_Ms_NeceCheck2(mControl)
            If sMesg = "OK" Then
                
                sQuery = Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", pControl)
                If Gf_Total_Display(M_CN1, Proc_Sc("Sc"), sQuery, 0, iSumCnt, iSumCol) Then
                    ss1.OperationMode = OperationModeNormal
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                End If
        
            Else
                sMesg = sMesg + " Must input according to length of item"
                Call Gp_MsgBoxDisplay(sMesg)
            End If
                
        Else
           sMesg = sMesg + " Must input necessarily"
           Call Gp_MsgBoxDisplay(sMesg)
        End If
                 
    Else
       Call MsgBox("输入日期不符合规范!" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
    End If
    
    
   
    
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


Private Sub Text_BB_DOME_FL_Change()
               
    Text_BB_DOME_FL.Text = StrConv(Text_BB_DOME_FL.Text, vbUpperCase)
    
    Select Case Text_BB_DOME_FL.Text
    
        Case "E", "D", ""
        
        Case Else
           Text_BB_DOME_FL.Text = ""
           Call MsgBox("订单分类代码：  " & Chr(10) & "D：内销" & Chr(10) & "E：出口", vbExclamation + vbOKOnly, "警告")
           
    End Select
     
End Sub

Private Sub Text_BB_DOME_FL_DblClick()

    Call Text_BB_DOME_FL_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_BB_DOME_FL_KeyUp(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"

        DD.sKey = "B0002"
        
        DD.rControl.Add Item:=Text_BB_DOME_FL
        DD.rControl.Add Item:=Text_BB_DOME_FL_mate
   
        DD.nameType = "2"
        'DD.nameType="1" 按中文名称查询
        'DD.nameType="2" 按英文名称查询
        
        Call Gf_Common_DD(M_CN1, KeyCode)

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() 用于客户代码

        Exit Sub
        
    End If

    If Len(Trim(Text_BB_DOME_FL.Text)) = Text_BB_DOME_FL.MaxLength Then
       '  Gf_ComnNAME_Find( 连接字符串, DD.sKEy内容 ,DD.nameType)
       ' Gf_CustNameFind( 连接字符串, 客户代码内容,DD.nameType)
        Text_BB_DOME_FL_mate.Text = Gf_ComnNameFind(M_CN1, "B0002", Text_BB_DOME_FL.Text, 2)
    Else
        Text_BB_DOME_FL_mate.Text = ""
    End If
    
End Sub

Private Sub text_bb_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String
    
    If Len(Trim(Text_BB_ORD_NO.Text)) = Text_BB_ORD_NO.MaxLength Then
    
        If Combo_ORD_ITEM.Text <> "" Then Exit Sub
        
        Text_BB_ORD_NO.Text = StrConv(Text_BB_ORD_NO.Text, vbUpperCase)
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(Text_BB_ORD_NO.Text) & "'"
        Call Gf_ComboAdd(M_CN1, Combo_ORD_ITEM, sQuery)
       
       'If Combo_ORD_ITEM.ListCount <> 0 Then
       '   Combo_ORD_ITEM.ListIndex = 0
       'End If
    Else
    
        Combo_ORD_ITEM.Clear
        
    End If

End Sub

Private Sub Text_BB_PROD_CD_Change()

    Select Case Text_BB_PROD_CD.Text
         Case "S", "s", "SL"
             Text_BB_PROD_CD.Text = "SL"
         Case "P", "p", "PP"
             Text_BB_PROD_CD.Text = "PP"
         Case "H", "h", "HC"
             Text_BB_PROD_CD.Text = "HC"
         Case "", "**"
             Text_BB_PROD_CD.Text = ""
         Case Else
             Text_BB_PROD_CD.Text = ""
             Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
    End Select
           
End Sub

Private Sub Text_BB_PROD_CD_DblClick()

    Call Text_BB_PROD_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_BB_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)
   
   Text_BB_PROD_CD_mate = ""
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=Text_BB_PROD_CD
        DD.rControl.Add Item:=Text_BB_PROD_CD_mate
   
        DD.nameType = "2"
        'DD.nameType="1" 按中文名称查询
        'DD.nameType="2" 按英文名称查询
       
        Call Gf_Common_DD(M_CN1, KeyCode)

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        'Gf_Customer_DD() 用于客户代码

        Exit Sub
        
    End If

    If Len(Trim(Text_BB_PROD_CD.Text)) = Text_BB_PROD_CD.MaxLength Then
       '  Gf_ComnNAME_Find( 连接字符串, DD.sKEy内容 ,DD.nameType)
       ' Gf_CustNameFind( 连接字符串, 客户代码内容,DD.nameType)
        Text_BB_PROD_CD_mate.Text = Gf_ComnNameFind(M_CN1, "B0005", Text_BB_PROD_CD.Text, 2)
    Else
        Text_BB_PROD_CD_mate.Text = ""
    End If

End Sub

Private Sub Text_BB_REC_STS_Change()

    If Not Text_BB_REC_STS.Text = "" Then
        If Not Text_BB_REC_STS.Text = "1" Then
            If Not Text_BB_REC_STS.Text = "2" Then
                If Not Text_BB_REC_STS.Text = "3" Then
    '               Call MsgBox("状态代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
                    Text_BB_REC_STS.Text = ""
                End If
            End If
        End If
    End If
'
'    If Text_BB_REC_STS.Text = "2" Then
'        Check_ord_END.Value = 1
'    '   Check_ord_END.Enabled = False
'    Else
'        Check_ord_END.Value = 0
'    '   Check_ord_END.Enabled = True
'    End If

End Sub

Private Sub Text_BB_REC_STS_DblClick()

    Call Text_BB_REC_STS_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_BB_REC_STS_KeyUp(KeyCode As Integer, Shift As Integer)

    Text_BB_REC_STS_Name = ""
    
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "Z0005"
        DD.rControl.Add Item:=Text_BB_REC_STS
        DD.rControl.Add Item:=Text_BB_REC_STS_Name
        
        DD.nameType = "2"
        'DD.nameType="1" 按中文名称查询
        'DD.nameType="2" 按英文名称查询
        Call Gf_Common_DD(M_CN1, KeyCode)
        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() 用于客户代码
        
        Exit Sub
        
    End If

    If Len(Trim(Text_BB_REC_STS.Text)) = Text_BB_REC_STS.MaxLength Then
       '  Gf_ComnNAME_Find( 连接字符串, DD.sKEy内容 ,DD.nameType)
       ' Gf_CustNameFind( 连接字符串, 客户代码内容,DD.nameType)
        Text_BB_REC_STS_Name.Text = Gf_ComnNameFind(M_CN1, "Z0005", Text_BB_REC_STS.Text, 2)
    Else
        Text_BB_REC_STS_Name.Text = ""
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

Private Sub TXT_CFM_MILL_PLT_DblClick()

    Call txt_cfm_mill_plt_KeyUp(vbKeyF4, 0)
    
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

Private Sub txt_cust_cd_DblClick()

    Call txt_cust_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_cust_cd
        DD.rControl.Add Item:=txt_cust_cd_name

        DD.nameType = "1"

        Call Gf_Customer_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_cust_cd)) = txt_cust_cd.MaxLength Then
        txt_cust_cd_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
    Else
        txt_cust_cd_name.Text = ""
    End If

End Sub

Private Sub txt_enduse_cd_DblClick()

    Call txt_enduse_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_enduse_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
                 
        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_ENDUSE_CD
        DD.nameType = "2"
            
        Call Gf_Usage_DD(M_CN1, KeyCode)
        
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

Private Sub txt_sale_way_DblClick()

    Call txt_sale_way_KeyUp(vbKeyF4, 0)
    
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
Private Sub txt_stdspec_DblClick()
    Call txt_stdspec_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stdspec_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_STDSPEC
           
        If Text_BB_PROD_CD.Text = "SL" Then
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
        'txt_act_stlgrd.Text = ""
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

