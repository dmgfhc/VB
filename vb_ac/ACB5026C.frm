VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB5026C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "库内不同区域装车实绩录入_ACB5026C"
   ClientHeight    =   9225
   ClientLeft      =   315
   ClientTop       =   1800
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_trans_way_s 
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
      Left            =   11520
      TabIndex        =   40
      Top             =   9240
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt_area 
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
      TabIndex        =   36
      Top             =   60
      Width           =   1500
   End
   Begin VB.TextBox txt_loc 
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
      Left            =   11640
      MaxLength       =   10
      TabIndex        =   35
      Top             =   60
      Width           =   1080
   End
   Begin VB.TextBox txt_sizeKnd_s 
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
      Height          =   315
      Left            =   11640
      MaxLength       =   2
      TabIndex        =   34
      Top             =   420
      Width           =   435
   End
   Begin VB.TextBox txt_stlgrd_s 
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
      Left            =   15780
      TabIndex        =   26
      Top             =   1410
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.TextBox txt_stdspec_s 
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
      Left            =   1275
      TabIndex        =   25
      Top             =   435
      Width           =   1965
   End
   Begin VB.TextBox text_prod_cd 
      Alignment       =   2  'Center
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
      Left            =   1275
      MaxLength       =   2
      TabIndex        =   24
      Tag             =   "产品"
      Top             =   60
      Width           =   465
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
      Left            =   4785
      MaxLength       =   2
      TabIndex        =   23
      Tag             =   "仓库"
      Top             =   60
      Width           =   435
   End
   Begin VB.TextBox text_cur_inv 
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
      Left            =   5235
      TabIndex        =   22
      Top             =   60
      Width           =   1500
   End
   Begin VB.TextBox txt_enduse_s 
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
      Left            =   16695
      TabIndex        =   21
      Top             =   1875
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt_trim_fl 
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
      Height          =   315
      Left            =   11640
      MaxLength       =   1
      TabIndex        =   20
      Top             =   795
      Width           =   435
   End
   Begin VB.TextBox txt_trim_name 
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
      Left            =   12105
      TabIndex        =   19
      Tag             =   "钢种"
      Top             =   795
      Width           =   1200
   End
   Begin VB.ComboBox cbo_prod_grd 
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
      ItemData        =   "ACB5026C.frx":0000
      Left            =   8250
      List            =   "ACB5026C.frx":0002
      TabIndex        =   18
      Top             =   435
      Width           =   1440
   End
   Begin VB.TextBox text_size_knd_name_in 
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
      Left            =   12105
      TabIndex        =   17
      Tag             =   "钢种"
      Top             =   420
      Width           =   1200
   End
   Begin VB.TextBox txt_area_code 
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
      Left            =   8250
      MaxLength       =   2
      TabIndex        =   16
      Tag             =   "起始区"
      Top             =   60
      Width           =   480
   End
   Begin VB.TextBox txt_htm_ord_cd 
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
      Left            =   14640
      MaxLength       =   7
      TabIndex        =   15
      Top             =   60
      Width           =   480
   End
   Begin VB.TextBox txt_lot_no 
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
      Left            =   4785
      TabIndex        =   14
      Tag             =   "轧批号"
      Top             =   435
      Width           =   1965
   End
   Begin VB.TextBox txt_prod_grd_s 
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
      Left            =   13605
      TabIndex        =   3
      Top             =   9300
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt_cust_cd_s 
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
      Left            =   12810
      TabIndex        =   2
      Top             =   9300
      Visible         =   0   'False
      Width           =   795
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7995
      Left            =   75
      TabIndex        =   0
      Top             =   1185
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   14102
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "ACB5026C.frx":0004
      Begin Threed.SSPanel SSPanel1 
         Height          =   1005
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1773
         _Version        =   196609
         BackColor       =   14737918
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txt_trans_comp 
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
            Left            =   5145
            MaxLength       =   6
            TabIndex        =   39
            Tag             =   "运输公司"
            Top             =   480
            Width           =   1050
         End
         Begin VB.TextBox txt_trans_comp_name 
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
            Left            =   6270
            MaxLength       =   40
            TabIndex        =   38
            Top             =   480
            Width           =   2000
         End
         Begin VB.ComboBox Cbo_trans_way 
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
            ItemData        =   "ACB5026C.frx":0056
            Left            =   1365
            List            =   "ACB5026C.frx":0066
            TabIndex        =   37
            Top             =   480
            Width           =   1080
         End
         Begin CSTextLibCtl.sitxEdit stx_move_time 
            Height          =   315
            Left            =   10080
            TabIndex        =   11
            Top             =   90
            Width           =   960
            _Version        =   262145
            _ExtentX        =   1693
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   "__:__:__"
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
            DataProperty    =   1
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   ""
            Text            =   "__:__:__"
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
            Mask            =   "%%:%%:%%"
            CharacterTable  =   ""
            BorderStyle     =   0
            MaxLength       =   6
         End
         Begin InDate.UDate udt_move_date 
            Height          =   330
            Left            =   8595
            TabIndex        =   10
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            BackColor       =   16777215
         End
         Begin VB.TextBox txt_car_no 
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
            Left            =   5145
            MaxLength       =   15
            TabIndex        =   9
            Tag             =   "车辆号"
            Top             =   90
            Width           =   1635
         End
         Begin VB.TextBox txt_plt_name 
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
            Left            =   1875
            TabIndex        =   8
            Tag             =   "工 厂"
            Top             =   90
            Width           =   1800
         End
         Begin VB.TextBox txt_plt 
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
            Left            =   1365
            MaxLength       =   2
            TabIndex        =   7
            Tag             =   "目标库"
            Top             =   90
            Width           =   465
         End
         Begin VB.TextBox txt_sale_dept_name 
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
            Left            =   12945
            TabIndex        =   6
            Tag             =   "工 厂"
            Top             =   630
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.TextBox txt_sale_dept 
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
            Left            =   12480
            MaxLength       =   3
            TabIndex        =   5
            Tag             =   "部门代码"
            Top             =   630
            Visible         =   0   'False
            Width           =   450
         End
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   165
            Tag             =   "移 送 工 厂"
            Top             =   90
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "目标区"
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
         Begin InDate.ULabel ULabel10 
            Height          =   315
            Left            =   3945
            Top             =   90
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "车辆号"
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
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   11280
            Tag             =   "移 送 工 厂"
            Top             =   630
            Visible         =   0   'False
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "部门代码"
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
            Left            =   7380
            Top             =   90
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "装车日"
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
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   11250
            Top             =   90
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "已选择量"
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
         End
         Begin CSTextLibCtl.sidbEdit sdb_slab_num 
            Height          =   315
            Left            =   12465
            TabIndex        =   12
            Top             =   90
            Width           =   675
            _Version        =   262145
            _ExtentX        =   1191
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
         Begin CSTextLibCtl.sidbEdit sdb_slab_wgt 
            Height          =   315
            Left            =   13155
            TabIndex        =   13
            Top             =   90
            Width           =   1410
            _Version        =   262145
            _ExtentX        =   2487
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
            NumIntDigits    =   7
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel20 
            Height          =   315
            Left            =   3945
            Top             =   480
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "运输公司"
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
         Begin InDate.ULabel ULabel21 
            Height          =   315
            Left            =   165
            Top             =   480
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "运输别"
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
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   6960
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1035
         Width           =   15195
         _Version        =   393216
         _ExtentX        =   26802
         _ExtentY        =   12277
         _StockProps     =   64
         ButtonDrawMode  =   4
         ColsFrozen      =   3
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
         MaxRows         =   20
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACB5026C.frx":0086
      End
   End
   Begin CSTextLibCtl.sidbEdit txt_len_min_s 
      Height          =   330
      Left            =   8250
      TabIndex        =   27
      Top             =   780
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      NumIntDigits    =   0
      ShowZero        =   0   'False
      MaxValue        =   99999999
      MinValue        =   -99999999
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_wid_min_s 
      Height          =   330
      Left            =   4785
      TabIndex        =   28
      Top             =   780
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      NumIntDigits    =   9
      ShowZero        =   0   'False
      MaxValue        =   99999999
      MinValue        =   -99999999
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_thk_min_s 
      Height          =   330
      Left            =   1275
      TabIndex        =   29
      Top             =   780
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      NumIntDigits    =   4
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   120
      Top             =   60
      Width           =   1110
      _ExtentX        =   1958
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   3635
      Top             =   60
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "仓库"
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
   Begin CSTextLibCtl.sidbEdit txt_len_max_s 
      Height          =   330
      Left            =   9225
      TabIndex        =   30
      Top             =   780
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      NumIntDigits    =   0
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_wid_max_s 
      Height          =   330
      Left            =   5760
      TabIndex        =   31
      Top             =   780
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      NumIntDigits    =   9
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_thk_max_s 
      Height          =   330
      Left            =   2265
      TabIndex        =   32
      Top             =   780
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      NumIntDigits    =   4
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel15 
      Height          =   315
      Left            =   10545
      Top             =   420
      Width           =   1050
      _ExtentX        =   1852
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
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   7095
      Top             =   435
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "等级"
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   120
      Top             =   435
      Width           =   1110
      _ExtentX        =   1958
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
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Left            =   15615
      Top             =   1875
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Caption         =   "用途"
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
      Left            =   10545
      Top             =   795
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   556
      Caption         =   "切边"
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   120
      Top             =   795
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "厚度"
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
      Left            =   3635
      Top             =   795
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "宽度"
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Left            =   7095
      Top             =   795
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "长度"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   7095
      Top             =   60
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "起始区"
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   13425
      Top             =   60
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "热处理对象"
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
   Begin Threed.SSCheck chk_htm_shot_blast 
      Height          =   285
      Left            =   13560
      TabIndex        =   33
      Top             =   480
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "抛丸作业对象"
   End
   Begin InDate.ULabel ULabel19 
      Height          =   315
      Left            =   3635
      Top             =   435
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "轧批号"
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
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   10545
      Top             =   60
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "垛位"
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
   Begin Threed.SSCheck chk_sel 
      Height          =   345
      Left            =   13410
      TabIndex        =   41
      Top             =   810
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   609
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
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
      Caption         =   "批次取消/选择"
   End
End
Attribute VB_Name = "ACB5026C"
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
'-- Program Name      Insert Moving result/半产品
'-- Program ID        ACB5025C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim.Sung.Ho
'-- Coder             Kim.Sung.Ho
'-- Date              2008.1.18
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

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_MV_NO = 1
Const SS1_MV_LST_NO = 2
Const SS1_MOVE_DATE = 9
Const SS1_MOVE_TIME = 10
Const SS1_WGT = 14
Const SS1_CAR_NO = 32
Const SS1_INS_EMP = 33
Const SS1_PROD_CD = 34

Const SS1_AREA_TO = 38 '  目标区域
Const SS1_CAR_CD = 39  ' 运输别
Const SS1_CAR_NAME = 40  ' 运输公司


'Dim crxApplication As New CRAXDRT.Application
'
'Public Report As CRAXDRT.Report
'
'Dim crxDatabaseTable As CRAXDRT.DatabaseTable
'Dim crxSubreport As CRAXDRT.Report
'Dim CPProperties As CRAXDRT.ConnectionProperties

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
       
             Call Gp_Ms_Collection(text_prod_cd, "p", "n", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_stdspec_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_stlgrd_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(text_cur_inv_code, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(text_cur_inv, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_prod_grd_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(cbo_prod_grd, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_enduse_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_cust_cd_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SizeKnd_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(Text_size_knd_name_IN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_thk_min_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_thk_max_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_wid_min_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_wid_max_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_len_min_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_len_max_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_Trim_fl, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_Trim_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_area_code, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)    ' 区的选择，
                  Call Gp_Ms_Collection(txt_loc, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_plt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_plt_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_car_no, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_sale_dept, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_sale_dept_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(udt_move_date, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(stx_move_time, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_slab_num, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_slab_wgt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(chk_htm_shot_blast, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_htm_ord_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_lot_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
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
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
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
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
  
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB5026C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="ACB5026C.P_SREFER1", Key:="P-R"
    sc1.Add Item:="ACB5026C.P_MODIFY1", Key:="P-L"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, 1, True)
    Call Gp_Sp_ColHidden(ss1, 2, True)
    Call Gp_Sp_ColHidden(ss1, 9, True)
    Call Gp_Sp_ColHidden(ss1, 10, True)
    Call Gp_Sp_ColHidden(ss1, 32, True)
    Call Gp_Sp_ColHidden(ss1, 33, True)
    Call Gp_Sp_ColHidden(ss1, 34, True)
    
    Call Gp_Sp_ColHidden(ss1, SS1_AREA_TO, True)
    Call Gp_Sp_ColHidden(ss1, SS1_CAR_CD, True)
    Call Gp_Sp_ColHidden(ss1, SS1_CAR_NAME, True)
    
    

End Sub

Private Sub Cbo_trans_way_Change()
 
    If Trim(Cbo_trans_way.Text) <> "" Then
        txt_trans_way_s.Text = Left(Cbo_trans_way.Text, 1)
    Else
        txt_trans_way_s.Text = ""
    End If
End Sub

Private Sub Cbo_trans_way_CLICK()
 
    If Trim(Cbo_trans_way.Text) <> "" Then
        txt_trans_way_s.Text = Left(Cbo_trans_way.Text, 1)
    Else
        txt_trans_way_s.Text = ""
    End If
End Sub


Private Sub chk_sel_Click(Value As Integer)
    Dim iRow As Integer
    
    If chk_sel Then
        For iRow = 1 To ss1.MaxRows
            ss1.Row = iRow
            ss1.Col = 0
            ss1.Text = "Update"
            ss1.Col = SS1_INS_EMP:   ss1.Text = sUserID
            ss1.Col = SS1_WGT
            sdb_slab_num.Value = sdb_slab_num.Value + 1
            sdb_slab_wgt.Value = sdb_slab_wgt.Value + ss1.Value
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFFFF80)
        Next iRow
    Else
        For iRow = 1 To ss1.MaxRows
            ss1.Row = iRow
            ss1.Col = 0
            ss1.Text = ""
            ss1.Col = SS1_INS_EMP:    ss1.Text = ""
            ss1.Col = SS1_WGT
            sdb_slab_num.Value = 0
            sdb_slab_wgt.Value = 0#
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow)
        Next iRow

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

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call AC_ComboAdd(M_CN1, cbo_prod_grd, "Q0034")
    
    Call AC_ComboAdd(M_CN1, Cbo_trans_way, "C0024")
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gf_Sp_Cls(sc1)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    text_prod_cd.Text = "PP"
    
    If App.Title = "CE" Then
        text_cur_inv_code.Text = "ZB"
    Else
        text_cur_inv_code.Text = "00"
    End If
    
    Call text_cur_inv_code_KeyUp(0, 0)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
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
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        text_prod_cd.Text = "PP"
        If App.Title = "CE" Then
            text_cur_inv_code.Text = "ZB"
        Else
            text_cur_inv_code.Text = "00"
        End If
        Call text_cur_inv_code_KeyUp(0, 0)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    Dim iRow As Integer
    Dim iCol As Integer
    Dim i As Integer

    If Gf_Sp_ProceExist(sc1.Item("Spread")) Then Exit Sub
            
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        
      For iRow = 1 To ss1.MaxRows
    
          ss1.Row = iRow
          ss1.Col = 37
           If ss1.Text = "Y" Then
              For i = 1 To ss1.MaxCols
                   ss1.Col = i
                   ss1.ForeColor = &HC000&
              Next
           End If
      
      Next iRow
        
        
        txt_car_no.Text = ""
        stx_move_time.RawData = ""
        sdb_slab_num.Value = 0
        sdb_slab_wgt.Value = 0
    End If

End Sub

Public Sub Form_Pro()

    Dim sQuery      As String
    Dim sErrMessg   As String
    Dim MvNo        As String
    Dim TransNo     As String
    Dim iRow        As Integer
    Dim iCol        As Integer
    Dim iCount      As Integer
    Dim sTemp       As String
    Dim dTempInt    As Double
    Dim intLastRow  As Integer
    On Error GoTo Process_Exec_ERROR
    
    If Trim(txt_plt.Text) = "" Or Trim(txt_plt_name.Text) = "" Then
        Call Gp_MsgBoxDisplay("输入目标区...")
        Exit Sub
    End If
    
    If Trim(txt_car_no.Text) = "" Then
        Call Gp_MsgBoxDisplay("输入车辆号...")
        Exit Sub
    End If
    
    If Len(udt_move_date.RawData) <> 8 Or Len(stx_move_time.RawData) <> 6 Then
        Call Gp_MsgBoxDisplay("输入转区日...")
        Exit Sub
    End If
    
    If Trim(text_cur_inv_code.Text) = Trim(txt_plt.Text) Then
        Call Gp_MsgBoxDisplay("起始区 = 目标区")
        Exit Sub
    End If
    
    If UCase(Trim(txt_plt.Text)) = "ZZ" Then
        Call Gp_MsgBoxDisplay("错误目标区")
        Exit Sub
    End If
    
'    If Cbo_trans_way.ListIndex = 0 Then
'
'        If Trim(txt_trans_comp.Text) = "" Then
'            Call Gp_MsgBoxDisplay("输入运输公司...")
'            Exit Sub
'        End If
'
'    End If
    
    If Trim(Cbo_trans_way.Text) = "" Then
        Call Gp_MsgBoxDisplay("输入运输别...")
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    M_CN1.BeginTrans
    
    MvNo = "C" & UCase(txt_area_code.Text) & UCase(txt_plt.Text)
    
    Call MoveTransNoEdit(MvNo, TransNo)
            
    iCount = 0
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        ss1.Col = 0
        If ss1.Text = "Update" Then
            iCount = iCount + 1
            intLastRow = iRow
            
            ss1.Col = SS1_MV_NO
            ss1.Text = UCase(MvNo)
            
            ss1.Col = SS1_MV_LST_NO
            ss1.Text = UCase(TransNo)
            
            ss1.Col = SS1_MOVE_DATE
            ss1.Text = udt_move_date.RawData
            
            ss1.Col = SS1_MOVE_TIME
            ss1.Text = stx_move_time.RawData
            
            ss1.Col = SS1_CAR_NO
            ss1.Text = txt_car_no.Text
            
            ss1.Col = SS1_PROD_CD
            ss1.Text = text_prod_cd.Text
            
            ss1.Col = SS1_CAR_CD
            ss1.Text = Cbo_trans_way.ListIndex
            
            ss1.Col = SS1_CAR_NAME
            ss1.Text = txt_trans_comp.Text
            
            
            sErrMessg = ""
            Call Sp_Process(iRow, sErrMessg)
                        
            'Error Check
            If Trim(sErrMessg) <> "" Then
                Call Gp_Sp_RowColor(ss1, iRow, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                M_CN1.RollbackTrans
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
             
        End If
    Next iRow

    'HP_LOAD_WGT  USE
    If iCount > 0 Then
        Call Sp_Process(intLastRow, sErrMessg, True)
    End If
    
    M_CN1.CommitTrans
    
    If iCount > 0 Then
        
        If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call MenuTool_ReSet
        End If
        txt_car_no.Text = ""
        sdb_slab_num.Value = 0
        sdb_slab_wgt.Value = 0
    End If
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Process_Exec_ERROR:

    M_CN1.RollbackTrans
    Call Gp_MsgBoxDisplay(Error & sErrMessg)
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub MoveTransNoEdit(MoveIspNo As String, TransNo As String)
    
    Dim SQL    As String
    Dim sDate  As String
    Dim AdoRs  As New ADODB.Recordset
    
    SQL = "SELECT  TO_CHAR(SYSDATE,'YYMMDD') FROM  DUAL " & vbCrLf
    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    If AdoRs.EOF = False Then
        sDate = AdoRs(0).Value & ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    MoveIspNo = MoveIspNo & Left(sDate, 4) & "0001"
    
    SQL = " SELECT    SUBSTR(MAX(MV_LST_NO),1,11) || LPAD(TO_NUMBER(SUBSTR(MAX(MV_LST_NO),12,4)) + 1,4,'0')"
    SQL = SQL & "     FROM  CP_MOVE_AREA                                     " & vbCrLf
    SQL = SQL & "    WHERE  MV_NO = '" & MoveIspNo & "'" & vbCrLf
    SQL = SQL & "      AND  MV_LST_NO  LIKE '" & Left(MoveIspNo, 5) & sDate & "%'" & vbCrLf
    
    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly

    If AdoRs.EOF Or AdoRs.BOF Then
        TransNo = Left(MoveIspNo, 5) & sDate & "0001"
    Else
        TransNo = AdoRs.Fields(0) & ""
    End If
    
    If TransNo = "" Then
        TransNo = Left(MoveIspNo, 5) & sDate & "0001"
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
End Sub

Private Sub Sp_Process(iRow As Integer, sErrMessg As String, Optional bLast As Boolean = False)

    Dim iCount      As Integer
    Dim iCol        As Integer
    Dim sTemp       As String
    Dim dTempInt    As Double

    Dim adoCmd As ADODB.Command

    On Error GoTo Process_Exec_ERROR

    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command

    Set adoCmd.ActiveConnection = M_CN1
    adoCmd.CommandType = adCmdStoredProc
    If bLast Then
       adoCmd.CommandText = sc1.Item("P-L")
    Else
        adoCmd.CommandText = sc1.Item("P-M")
    End If
    'Create Parameter (Input) iType + iColumn
    For iCount = 0 To sc1.Item("iColumn").Count
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount

    adoCmd.Parameters(0).Value = "U"
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)

    sc1.Item("Spread").Row = iRow

    'Parameters Setting
    For iCol = 1 To sc1.Item("iColumn").Count

        sc1.Item("Spread").Col = sc1.Item("iColumn").Item(iCol)
        Select Case sc1.Item("Spread").CellType

            Case SS_CELL_TYPE_NUMBER
                If Trim(sc1.Item("Spread").Text) = "" Then
                    adoCmd.Parameters(iCol).Value = 0
                Else
                    dTempInt = sc1.Item("Spread").Text
                    adoCmd.Parameters(iCol).Value = Trim(Str(dTempInt))
                End If
                
            Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If Trim(sc1.Item("Spread").Value) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(sc1.Item("Spread").Value)
                        End If
                        
            Case Else
                sTemp = Replace(sc1.Item("Spread").Text, "'", "''")
                adoCmd.Parameters(iCol).Value = Trim(sTemp)

        End Select

    Next iCol

    adoCmd.Execute

    If adoCmd("Error") <> "0" Then
        sErrMessg = adoCmd("Messg")
    End If

    Set adoCmd = Nothing
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    sErrMessg = Error
End Sub

Public Sub Spread_ColumnsSort()
    Spread_ColSort.Show 1
End Sub

Public Sub Spread_Forzens_Setting()
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
End Sub

Public Sub Spread_Forzens_Cancel()
    Me.ActiveControl.ColsFrozen = 0
End Sub

Public Sub Spread_Del()

    Call Gp_Sp_Del(sc1)

End Sub

Public Sub Spread_Can()

    On Error Resume Next

    Dim sQuery As String
    Dim i As Integer
    Dim iRow, BR1, BR2 As Long

    With sc1
        
        .Item("Spread").ReDraw = False
        
        If .Item("Spread").MaxRows < 1 Or .Item("Spread").SelBlockRow < 1 Then
            Exit Sub
        End If
    
        BR1 = .Item("Spread").SelBlockRow
        BR2 = .Item("Spread").SelBlockRow2
        
        For iRow = .Item("Spread").SelBlockRow To BR2
            
            Select Case Trim(Gf_Sp_RcvData(.Item("Spread"), 0, iRow))
                
                Case "Delete"
                    Call Gp_Sp_SendData(.Item("Spread"), "", 0, iRow)
                    Call Gp_Sp_RowColor(.Item("Spread"), iRow)
                    
                    For i% = 1 To sc1!iColumn.Count
                        Call Gp_Sp_CellColor(.Item("Spread"), sc1!iColumn(i%), iRow, , &HC0FFFF)
                    Next i%
                Case Else
                    'sQuery = Gf_Sp_MakeQuery(.Item("Spread"), .Item("P-O"), "O", .Item("icolumn"), iRow)
                    'Call Gp_Sp_OneRowDisplay(Conn, sQuery, .Item("Spread"), iRow)
            End Select
            
            If iRow = BR2 Then
                Exit For
            End If

        Next iRow
        
        .Item("Spread").ReDraw = True
        
    End With
          
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

Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ss1.MaxRows > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row <= 0 Then Exit Sub
    
    Call Gp_Sp_Sort(sc1.Item("Spread"), Col, Row)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    ss1.Row = Row
    ss1.Col = 0
    
    If ss1.Text <> "Update" Then
        ss1.Col = 0:    ss1.Text = "Update"
        ss1.Col = SS1_INS_EMP:   ss1.Text = sUserID
        ss1.Col = SS1_WGT
        sdb_slab_num.Value = sdb_slab_num.Value + 1
        sdb_slab_wgt.Value = sdb_slab_wgt.Value + ss1.Value
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
   Else
        ss1.Col = 0:     ss1.Text = ""
        ss1.Col = SS1_INS_EMP:    ss1.Text = ""
        ss1.Col = SS1_WGT
        sdb_slab_num.Value = sdb_slab_num.Value - 1
        sdb_slab_wgt.Value = sdb_slab_wgt.Value - ss1.Value
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
   End If

End Sub

Private Sub stx_move_time_DblClick()

    stx_move_time.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'HH24MISS') FROM DUAL")
    
End Sub

Private Sub Text_PROD_CD_Change()

    If Len(text_prod_cd) <> 2 Then Exit Sub

    Select Case text_prod_cd.Text

        Case "PP", "pp"
            text_prod_cd.Text = "PP"
        Case "MP", "mp"
            text_prod_cd.Text = "MP"
        Case "", "**"
            text_prod_cd.Text = ""
        Case Else
            text_prod_cd.Text = ""
            Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
    End Select
        
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    If text_prod_cd.Text = "MP" Then
       text_cur_inv_code.Text = "ZB"
    End If
    
    cbo_prod_grd.Clear
    
    Select Case text_prod_cd.Text
    
        Case "S", "s", "SL"
        
            cbo_prod_grd.AddItem "0:合格"
            cbo_prod_grd.AddItem "1:表面不合格"
            cbo_prod_grd.AddItem "2:内部缺陷"
            cbo_prod_grd.AddItem "3:内外缺陷"
            cbo_prod_grd.AddItem "4:操作员变更"
            cbo_prod_grd.AddItem "5:长度不合格"
            
            txt_stlgrd_s.Visible = True
            txt_stdspec_s.Visible = False
        Case Else
            
            Call AC_ComboAdd(M_CN1, cbo_prod_grd, "Q0034")
            
            txt_stlgrd_s.Visible = False
            txt_stdspec_s.Visible = True
            
    End Select
    
End Sub

Private Sub Text_PROD_CD_DblClick()

    Call Text_PROD_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=text_prod_cd
    
        DD.nameType = "2"
    
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If

End Sub

Private Sub text_cur_inv_code_DblClick()

    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)


     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        
        DD.sKey = "C0013"

        DD.rControl.Add Item:=text_cur_inv_code
        DD.rControl.Add Item:=text_cur_inv
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
     
        If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
            text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
        Else
          text_cur_inv.Text = ""
        End If
        
    End If
    
End Sub


Private Sub txt_area_code_DblClick()

    Call txt_area_code_KeyUp(vbKeyF4, 0)
    
End Sub

' 区域
Private Sub txt_area_code_KeyUp(KeyCode As Integer, Shift As Integer)

Dim S_CODE As String  ' 公用表格
 
    If text_cur_inv_code.Text = "00" Then
            S_CODE = "C0023"
    ElseIf text_cur_inv_code.Text = "WG" Then
        S_CODE = "C0026"
    ElseIf text_cur_inv_code.Text = "HB" Then
        S_CODE = "C0028"
    End If
         If KeyCode = vbKeyF4 Then
        
            DD.sWitch = "MS"
            DD.sKey = S_CODE
    
            DD.rControl.Add Item:=txt_area_code
            DD.rControl.Add Item:=txt_area
            
            DD.nameType = "2"
            Call Gf_Common_DD(M_CN1, KeyCode)
            
        Else
         
            If Len(Trim(txt_area_code.Text)) = txt_area_code.MaxLength Then
                txt_area.Text = Gf_ComnNameFind(M_CN1, S_CODE, text_cur_inv_code.Text, 2)
            Else
              txt_area.Text = ""
            End If
            
        End If
 
  
End Sub



Private Sub txt_EndUse_s_DblClick()

    Call txt_EndUse_s_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_EndUse_s_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        If text_prod_cd.Text = "SL" Then
            DD.sKey = "S"
        Else
            DD.sKey = "P"
        End If
        
        DD.rControl.Add Item:=txt_enduse_s
        
        Call Gf_Usage_DD(M_CN1, KeyCode)
    End If
    
End Sub

Private Sub txt_htm_ord_cd_DblClick()

    Call txt_htm_ord_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_htm_ord_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Q0073"
        
        DD.rControl.Add Item:=txt_htm_ord_cd
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If
    
End Sub

Private Sub txt_plt_DblClick()

    Call txt_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

 Dim S_CODE As String
 If text_cur_inv_code.Text = "00" Then
         S_CODE = "C0023"
    ElseIf text_cur_inv_code.Text = "WG" Then
        S_CODE = "C0026"
    ElseIf text_cur_inv_code.Text = "HB" Then
        S_CODE = "C0028"
    End If
        If KeyCode = vbKeyF4 Then
        
            DD.sWitch = "MS"
            DD.sKey = S_CODE
            DD.rControl.Add Item:=txt_plt
            DD.rControl.Add Item:=txt_plt_name
            
            DD.nameType = "2"
            Call Gf_Common_DD(M_CN1, KeyCode)
            
        Else
        
            If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
                txt_plt_name.Text = Gf_ComnNameFind(M_CN1, S_CODE, txt_plt.Text, 2)
            Else
                  txt_plt_name.Text = ""
            End If
        
        End If
 
End Sub

Private Sub txt_Sale_dept_DblClick()

    Call txt_Sale_dept_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_Sale_dept_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Z0002"

        DD.rControl.Add Item:=txt_sale_dept
        DD.rControl.Add Item:=txt_sale_dept_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_sale_dept.Text)) = txt_sale_dept.MaxLength Then
            txt_sale_dept_name.Text = Gf_ComnNameFind(M_CN1, "Z0002", txt_sale_dept.Text, 2)
        Else
          txt_sale_dept_name.Text = ""
        End If
        
    End If
    
End Sub

Private Sub txt_car_no_DblClick()

    Call txt_car_no_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_car_no_KeyUp(KeyCode As Integer, Shift As Integer)

    If ULabel10.Caption <> "车辆号" Then Exit Sub
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_car_no
  '      DD.rControl.Add Item:=txt_fac_name

        DD.nameType = "2"

        Call Gf_CAR_NO_DD(M_CN1, KeyCode)

    End If

End Sub

Public Function Gf_CAR_NO_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String

    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "A"        'Apply Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "SELECT CAR_NO, CAR_KND,CAR_WGT_MAX,CAR_WGT_AVE,CAR_CMP_CD,Gf_Comnnamefind('H0002',CAR_CMP_CD) AS CAR_CMP_NAME FROM  HP_CAR_IMF "
    '    DD.sQuery = DD.sQuery + " WHERE "
        DD.sWhere = " WHERE CAR_NO like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere & " AND CAR_KND <> 'H' "

    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
    End If
    
    DD.sWitch = ""
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing
    
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

Private Sub cbo_prod_grd_Click()

    If Trim(cbo_prod_grd.Text) <> "" Then
        txt_prod_grd_s.Text = Left(cbo_prod_grd.Text, 1)
    Else
        txt_prod_grd_s.Text = ""
    End If
    
End Sub

Private Sub cbo_prod_grd_Change()

    If Trim(cbo_prod_grd.Text) <> "" Then
        txt_prod_grd_s.Text = Left(cbo_prod_grd.Text, 1)
    Else
        txt_prod_grd_s.Text = ""
    End If
    
End Sub

Private Sub txt_SizeKnd_s_DblClick()

    Call txt_SizeKnd_s_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_SizeKnd_s_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"

        DD.rControl.Add Item:=txt_SizeKnd_s
        DD.rControl.Add Item:=Text_size_knd_name_IN

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_SizeKnd_s.Text)) = txt_SizeKnd_s.MaxLength Then
            Text_size_knd_name_IN.Text = Gf_ComnNameFind(M_CN1, "B0043", txt_SizeKnd_s.Text, 2)
            Exit Sub
        Else
            Text_size_knd_name_IN.Text = ""
        End If
    End If
    
End Sub

Private Sub txt_stdspec_s_DblClick()

    Call txt_stdspec_s_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_trim_fl_DblClick()

    Call txt_trim_fl_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_trim_fl_KeyUp(KeyCode As Integer, Shift As Integer)

        If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0021"

        DD.rControl.Add Item:=txt_Trim_fl
        DD.rControl.Add Item:=txt_Trim_NAME

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_Trim_fl.Text)) = txt_Trim_fl.MaxLength Then
            txt_Trim_NAME.Text = Gf_ComnNameFind(M_CN1, "B0021", txt_Trim_fl.Text, 2)
            txt_Trim_fl.Text = Trim(txt_Trim_fl.Text)
            Exit Sub
        Else
            txt_Trim_NAME.Text = ""
            txt_Trim_fl.Text = ""
        End If
    
    End If

End Sub

Private Sub txt_stdspec_s_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec_s

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)
    End If
    
End Sub

Private Sub txt_stlgrd_s_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stlgrd_s
        
        DD.nameType = "1"
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
    End If
    
End Sub

Private Function AC_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, sPrc As String, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim sQuery As String
    Dim intCount As Integer
    Dim AdoRs As ADODB.Recordset
    
    If Trim(sPrc) = "" Then
        AC_ComboAdd = False: Exit Function
    End If
    
    intCount = 1
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then AC_ComboAdd = False: Exit Function
    End If
    
    sQuery = "SELECT CD_NAME FROM ZP_CD Where CD_MANA_NO = '" + Trim(sPrc) + "'"

    If ClsChk Then
        Cbo.Clear
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
                If intCount = 6 Then intCount = 7
                Cbo.AddItem Trim(Str(intCount)) + ":" + AdoRs.Fields(0)
                intCount = intCount + 1
            End If
            AdoRs.MoveNext
            
        Wend
        AC_ComboAdd = True
    Else
        AC_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    AC_ComboAdd = False

End Function
