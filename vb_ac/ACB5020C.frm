VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB5020C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "产品装车实绩录入_ACB5020C"
   ClientHeight    =   9225
   ClientLeft      =   660
   ClientTop       =   1755
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox text_send_plan_wgt 
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
      Left            =   14160
      TabIndex        =   42
      Top             =   840
      Width           =   780
   End
   Begin VB.TextBox text_send_no 
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
      Left            =   14160
      TabIndex        =   41
      Top             =   450
      Width           =   1620
   End
   Begin VB.TextBox txt_trns_no 
      Height          =   345
      Left            =   15090
      TabIndex        =   40
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   11415
      TabIndex        =   36
      Tag             =   "轧批号"
      Top             =   450
      Width           =   1515
   End
   Begin VB.TextBox txt_htm_ord_cd 
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
      Left            =   12990
      MaxLength       =   7
      TabIndex        =   30
      Top             =   75
      Width           =   570
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
      Left            =   4665
      MaxLength       =   10
      TabIndex        =   21
      Top             =   450
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
      Left            =   8700
      TabIndex        =   20
      Tag             =   "钢种"
      Top             =   75
      Width           =   1440
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
      ItemData        =   "ACB5020C.frx":0000
      Left            =   1950
      List            =   "ACB5020C.frx":0002
      TabIndex        =   19
      Top             =   450
      Width           =   1170
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
      Left            =   8700
      TabIndex        =   18
      Tag             =   "钢种"
      Top             =   450
      Width           =   1440
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
      Left            =   8190
      MaxLength       =   1
      TabIndex        =   17
      Top             =   450
      Width           =   495
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
      Left            =   8190
      MaxLength       =   2
      TabIndex        =   16
      Top             =   75
      Width           =   495
   End
   Begin VB.TextBox txt_enduse_s 
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
      Left            =   1410
      TabIndex        =   15
      Top             =   450
      Width           =   495
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
      TabIndex        =   14
      Top             =   9300
      Visible         =   0   'False
      Width           =   795
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
      Left            =   5175
      TabIndex        =   10
      Top             =   75
      Width           =   1440
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
      Left            =   4665
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "起始库"
      Top             =   75
      Width           =   495
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
      Left            =   1410
      MaxLength       =   2
      TabIndex        =   9
      Tag             =   "产品"
      Top             =   75
      Width           =   495
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
      TabIndex        =   8
      Top             =   9300
      Visible         =   0   'False
      Width           =   795
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
      Left            =   11415
      TabIndex        =   4
      Top             =   820
      Width           =   2100
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
      Left            =   11415
      TabIndex        =   3
      Top             =   820
      Width           =   2580
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7995
      Left            =   45
      TabIndex        =   1
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
      PaneTree        =   "ACB5020C.frx":0004
      Begin Threed.SSPanel SSPanel1 
         Height          =   885
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1561
         _Version        =   196609
         BackColor       =   14737918
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
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
            ItemData        =   "ACB5020C.frx":0056
            Left            =   1365
            List            =   "ACB5020C.frx":0060
            TabIndex        =   29
            Top             =   465
            Width           =   1080
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
            Left            =   9670
            MaxLength       =   40
            TabIndex        =   37
            Top             =   465
            Width           =   2000
         End
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
            Left            =   8610
            MaxLength       =   10
            TabIndex        =   32
            Tag             =   "运输公司"
            Top             =   465
            Width           =   1050
         End
         Begin CSTextLibCtl.sitxEdit stx_move_time 
            Height          =   315
            Left            =   10095
            TabIndex        =   28
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
            Height          =   315
            Left            =   8610
            TabIndex        =   27
            Top             =   90
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
            Left            =   5025
            MaxLength       =   15
            TabIndex        =   26
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
            TabIndex        =   25
            Tag             =   "工 厂"
            Top             =   90
            Width           =   1560
         End
         Begin VB.TextBox txt_plt 
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
            Left            =   1365
            MaxLength       =   2
            TabIndex        =   24
            Tag             =   "目标库"
            Top             =   90
            Width           =   495
         End
         Begin VB.TextBox txt_trans_tool_name 
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
            Left            =   5540
            TabIndex        =   23
            Top             =   465
            Width           =   1560
         End
         Begin VB.TextBox txt_trans_tool 
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
            Left            =   5025
            MaxLength       =   3
            TabIndex        =   31
            Tag             =   "车型"
            Top             =   465
            Width           =   495
         End
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   165
            Tag             =   "移 送 工 厂"
            Top             =   90
            Width           =   1170
            _ExtentX        =   2064
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
         Begin InDate.ULabel ULabel10 
            Height          =   315
            Left            =   3825
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
            Left            =   3825
            Top             =   465
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            Caption         =   "车型"
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
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "转库日"
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
            Left            =   11940
            Top             =   465
            Width           =   1260
            _ExtentX        =   2223
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
            Left            =   13230
            TabIndex        =   33
            Top             =   465
            Width           =   645
            _Version        =   262145
            _ExtentX        =   1138
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   0
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
            Enabled         =   0   'False
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
            Justification   =   1
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
            Left            =   13890
            TabIndex        =   35
            Top             =   465
            Width           =   990
            _Version        =   262145
            _ExtentX        =   1746
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   0
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
            Enabled         =   0   'False
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
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel13 
            Height          =   315
            Left            =   7380
            Top             =   465
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
         Begin InDate.ULabel ULabel14 
            Height          =   315
            Left            =   165
            Top             =   465
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
         Begin CSTextLibCtl.sidbEdit sdb_car_max_wgt 
            Height          =   315
            Left            =   13890
            TabIndex        =   39
            Top             =   90
            Width           =   990
            _Version        =   262145
            _ExtentX        =   1746
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
            Enabled         =   0   'False
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
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel20 
            Height          =   315
            Left            =   11940
            Top             =   90
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            Caption         =   "车辆最大载荷(吨)"
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
         Begin VB.Label Label1 
            BackColor       =   &H00E0E1FE&
            Caption         =   "*汽运抛erp系统计算运费"
            ForeColor       =   &H000000C0&
            Height          =   360
            Left            =   2550
            TabIndex        =   38
            Top             =   450
            Width           =   1125
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   7080
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   915
         Width           =   15195
         _Version        =   393216
         _ExtentX        =   26802
         _ExtentY        =   12488
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
         MaxCols         =   35
         MaxRows         =   2
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACB5020C.frx":0070
      End
   End
   Begin CSTextLibCtl.sidbEdit txt_len_min_s 
      Height          =   330
      Left            =   8190
      TabIndex        =   5
      Top             =   820
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
      Left            =   4665
      TabIndex        =   6
      Top             =   825
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
      Left            =   1410
      TabIndex        =   7
      Top             =   825
      Width           =   825
      _Version        =   262145
      _ExtentX        =   1455
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
      Left            =   210
      Top             =   75
      Width           =   1170
      _ExtentX        =   2064
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
      Left            =   3465
      Top             =   75
      Width           =   1170
      _ExtentX        =   2064
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
   Begin CSTextLibCtl.sidbEdit txt_len_max_s 
      Height          =   330
      Left            =   9180
      TabIndex        =   11
      Top             =   820
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
      Left            =   5640
      TabIndex        =   12
      Top             =   825
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
      TabIndex        =   13
      Top             =   825
      Width           =   825
      _Version        =   262145
      _ExtentX        =   1455
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
      Left            =   6990
      Top             =   75
      Width           =   1170
      _ExtentX        =   2064
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
      Left            =   1950
      Top             =   75
      Width           =   1170
      _ExtentX        =   2064
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
      Left            =   10215
      Top             =   825
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
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Left            =   210
      Top             =   450
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "用途"
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
   Begin InDate.ULabel ULabel23 
      Height          =   315
      Left            =   6990
      Top             =   450
      Width           =   1170
      _ExtentX        =   2064
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
      Left            =   210
      Top             =   820
      Width           =   1170
      _ExtentX        =   2064
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
      Left            =   3465
      Top             =   820
      Width           =   1170
      _ExtentX        =   2064
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
      Left            =   6990
      Top             =   820
      Width           =   1170
      _ExtentX        =   2064
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
      Left            =   3465
      Top             =   450
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "堆放位置"
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   11775
      Top             =   75
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
      Left            =   10170
      TabIndex        =   34
      Top             =   90
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
      Left            =   10215
      Top             =   450
      Width           =   1170
      _ExtentX        =   2064
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
   Begin InDate.ULabel ULabel21 
      Height          =   315
      Left            =   12960
      Top             =   450
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "派车申请单号"
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
   Begin Threed.SSCheck chk_sel 
      Height          =   345
      Left            =   13620
      TabIndex        =   43
      Top             =   60
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
Attribute VB_Name = "ACB5020C"
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
'-- Program Name      Insert Moving result
'-- Program ID        ACB5020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim.Sung.Ho
'-- Coder             Kim.Sung.Ho
'-- Date              2007.8.1
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
'
'Dim crxApplication As New CRAXDRT.Application
'
'Public Report As CRAXDRT.Report
'
'Dim crxDatabaseTable As CRAXDRT.DatabaseTable
'Dim crxSubreport As CRAXDRT.Report
'Dim CPProperties As CRAXDRT.ConnectionProperties

Const SPD_MV_NO = 1
Const SPD_MV_LST_NO = 2
Const SPD_MAT_NO = 3
Const SPD_MOPLATE_NO = 4
Const SPD_PROD_DATE = 5
Const SPD_LOC = 6
Const SPD_MOVE_DATE = 7
Const SPD_MOVE_TIME = 8
Const SPD_THK = 9
Const SPD_WID = 10
Const SPD_LEN = 11
Const SPD_WGT = 12
Const SPD_UST_FL = 24
Const SPD_UST_RLT_CD = 25
Const SPD_CL_FL = 26
Const SPD_HTM_SHOT_BLAST = 27
Const SPD_HTM_METH = 28
Const SPD_CAR_NO = 29
Const SPD_EMP_CD = 30
Const SPD_PROD_CD = 31
Const SPD_TRANS_WAY = 33
Const SPD_TRANS_COMP = 34
Const SPD_TRANS_TOOL = 35

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
                  Call Gp_Ms_Collection(txt_loc, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_plt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_plt_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_car_no, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_trans_tool, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_trans_tool_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(udt_move_date, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(stx_move_time, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_slab_num, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_slab_wgt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(chk_htm_shot_blast, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_htm_ord_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_lot_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_trns_no, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(Cbo_trans_way, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_trans_tool, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_trans_tool_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_trans_comp, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_trans_comp_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
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
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB5020C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="ACB5020C.P_SREFER1", Key:="P-R"
'    sc1.Add Item:="ACB5020C.P_MODIFY1", Key:="P-L"
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
    
    Call Gp_Sp_ColHidden(ss1, SPD_MV_NO, True)
    Call Gp_Sp_ColHidden(ss1, SPD_MV_LST_NO, True)
    Call Gp_Sp_ColHidden(ss1, SPD_MOVE_DATE, True)
    Call Gp_Sp_ColHidden(ss1, SPD_MOVE_TIME, True)
    Call Gp_Sp_ColHidden(ss1, SPD_UST_FL, True)
    Call Gp_Sp_ColHidden(ss1, SPD_UST_RLT_CD, True)
    Call Gp_Sp_ColHidden(ss1, SPD_CL_FL, True)
    Call Gp_Sp_ColHidden(ss1, SPD_CAR_NO, True)
    Call Gp_Sp_ColHidden(ss1, SPD_EMP_CD, True)
    Call Gp_Sp_ColHidden(ss1, SPD_PROD_CD, True)
    Call Gp_Sp_ColHidden(ss1, SPD_TRANS_WAY, True)
    Call Gp_Sp_ColHidden(ss1, SPD_TRANS_COMP, True)
    Call Gp_Sp_ColHidden(ss1, SPD_TRANS_TOOL, True)

End Sub


Private Sub Cbo_trans_way_CLICK()
   If Cbo_trans_way.Text <> "汽运" Then
      txt_trans_tool.Enabled = False
      txt_trans_tool_name.Enabled = False
      txt_trans_comp.Enabled = False
      txt_trans_comp_name.Enabled = False
      txt_trans_tool.Text = "R"
      txt_trans_tool_name.Text = "R"
      txt_trans_comp.Text = "R"
      txt_trans_comp_name.Text = "R"
   Else
      txt_trans_tool.Enabled = True
      txt_trans_tool_name.Enabled = True
      txt_trans_comp.Enabled = True
      txt_trans_comp_name.Enabled = True
      txt_trans_tool.Text = ""
      txt_trans_tool_name.Text = ""
      txt_trans_comp.Text = ""
      txt_trans_comp_name.Text = ""
   End If
End Sub

Private Sub chk_sel_Click(Value As Integer)
    Dim iRow As Integer
    
    If chk_sel Then
        For iRow = 1 To ss1.MaxRows
            ss1.Row = iRow
            ss1.Col = 0
            ss1.Text = "Update"
            ss1.Col = SPD_EMP_CD:   ss1.Text = sUserID
            ss1.Col = 1:            ss1.Text = text_send_no.Text
            ss1.Col = SPD_WGT
            sdb_slab_num.Value = sdb_slab_num.Value + 1
            sdb_slab_wgt.Value = sdb_slab_wgt.Value + ss1.Value
            text_send_plan_wgt.Text = Str(Val(text_send_plan_wgt.Text) - ss1.Value)
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFFFF80)
        Next iRow
    Else
        For iRow = 1 To ss1.MaxRows
            ss1.Row = iRow
            ss1.Col = 0
            ss1.Text = ""
            ss1.Col = SPD_EMP_CD:    ss1.Text = ""
            ss1.Col = 1:            ss1.Text = ""
            ss1.Col = SPD_WGT
            sdb_slab_num.Value = 0
            sdb_slab_wgt.Value = 0#
            text_send_plan_wgt.Text = 0
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
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gf_Sp_Cls(sc1)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    text_prod_cd.Text = "PP"
    text_cur_inv_code.Text = "00"
    Call text_cur_inv_code_KeyUp(0, 0)
    Cbo_trans_way.ListIndex = 0
    
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
        text_cur_inv_code.Text = "00"
        Call text_cur_inv_code_KeyUp(0, 0)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    If Gf_Sp_ProceExist(sc1.Item("Spread")) Then Exit Sub
            
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        txt_car_no.Text = ""
        stx_move_time.RawData = ""
        sdb_slab_num.Value = 0
        sdb_slab_wgt.Value = 0
        If text_prod_cd.Text = "PP" Then
            Call Gp_Sp_ColHidden(ss1, SPD_HTM_SHOT_BLAST, False)
            Call Gp_Sp_ColHidden(ss1, SPD_HTM_METH, False)
        Else
            Call Gp_Sp_ColHidden(ss1, SPD_HTM_SHOT_BLAST, True)
            Call Gp_Sp_ColHidden(ss1, SPD_HTM_METH, True)
        End If
    End If

End Sub

Public Sub Form_Pro()

    Dim sQuery      As String
    Dim sErrMessg   As String
    Dim MvNo        As String
    Dim TransNo     As String
    Dim MOVENO     As String
    Dim iRow        As Integer
    Dim iCol        As Integer
    Dim iCount      As Integer
    Dim sTemp       As String
    Dim dTempInt    As Double
    Dim intLastRow  As Integer
    
    Dim dDate_Limit As Date
    Dim dDate_Now   As Date
    Dim iDate_Num   As Double
    On Error GoTo Process_Exec_ERROR
    
    dDate_Now = Now
    dDate_Limit = CDate("2011-06-03 00:00:00")
    iDate_Num = dDate_Limit - dDate_Now
    iDate_Num = Round(iDate_Num)
    
    If sdb_car_max_wgt.Value = 0 Or sdb_car_max_wgt.Text = "" Then
       sdb_car_max_wgt.Value = 85
    End If
    
    If Trim(txt_plt.Text) = "" Or Trim(txt_plt_name.Text) = "" Then
        Call Gp_MsgBoxDisplay("输入目标库...")
        Exit Sub
    End If
    
    If Trim(txt_car_no.Text) = "" Then
        Call Gp_MsgBoxDisplay("输入车辆号...")
        Exit Sub
    End If
    
    If Cbo_trans_way.ListIndex = 0 Then

        If Trim(txt_trans_tool.Text) = "" Then
            Call Gp_MsgBoxDisplay("输入车型...")
            Exit Sub
        End If
        If Trim(txt_trans_comp.Text) = "" Then
            Call Gp_MsgBoxDisplay("输入运输公司...")
            Exit Sub
        End If
        
    End If
    
    If Trim(Cbo_trans_way.Text) = "" Then
        Call Gp_MsgBoxDisplay("输入运输别...")
        Exit Sub
    End If
    
    If Len(udt_move_date.RawData) <> 8 Or Len(stx_move_time.RawData) <> 6 Then
        Call Gp_MsgBoxDisplay("输入转库日...")
        Exit Sub
    End If
    
    If Trim(text_cur_inv_code.Text) = Trim(txt_plt.Text) Then
        Call Gp_MsgBoxDisplay("起始库 = 目标库")
        Exit Sub
    End If
    
    If UCase(Trim(txt_plt.Text)) = "ZZ" Then
        Call Gp_MsgBoxDisplay("错误目标库")
        Exit Sub
    End If
    
    If sdb_slab_wgt.Value > sdb_car_max_wgt.Value Then
        If dDate_Now > dDate_Limit Then
           Call Gp_MsgBoxDisplay("已超最大载荷上限 " & sdb_car_max_wgt.Value & " 吨 ，请驳载！")
           Exit Sub
        Else
'           Call Gp_MsgBoxDisplay("已超最大载荷上限 " & sdb_car_max_wgt.Value & " 吨 ，" & iDate_Num & " 天后将不允许超载！", "I")
           Call Gp_MsgBoxDisplay("已超最大载荷上限 " & sdb_car_max_wgt.Value & " 吨 ，N 天后将不允许超载！", "I")
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    M_CN1.BeginTrans
    
    MvNo = "C" & UCase(text_cur_inv_code.Text) & UCase(txt_plt.Text)
    
    Call MoveTransNoEdit(MvNo, TransNo)
            
    txt_trns_no.Text = TransNo
    
    iCount = 0
    
    If Trim(text_send_no.Text) = "" Then
        MOVENO = UCase(MvNo)
    Else
        MOVENO = UCase(text_send_no.Text)
    End If
        
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        ss1.Col = 0
        
        If sdb_slab_num.Value = iCount Then Exit For
        
        If ss1.Text = "Update" Then
            iCount = iCount + 1
            intLastRow = iRow
            
            ss1.Col = SPD_MV_NO:            ss1.Text = UCase(MOVENO)
            ss1.Col = SPD_MV_LST_NO:        ss1.Text = UCase(TransNo)
            ss1.Col = SPD_MOVE_DATE:        ss1.Text = udt_move_date.RawData
            ss1.Col = SPD_MOVE_TIME:        ss1.Text = stx_move_time.RawData
            ss1.Col = SPD_CAR_NO:           ss1.Text = txt_car_no.Text
            ss1.Col = SPD_PROD_CD:          ss1.Text = text_prod_cd.Text
            ss1.Col = SPD_TRANS_WAY:        ss1.Text = Cbo_trans_way.ListIndex
            ss1.Col = SPD_TRANS_COMP:       ss1.Text = txt_trans_comp.Text
            ss1.Col = SPD_TRANS_TOOL:       ss1.Text = txt_trans_tool.Text
            
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

    
    M_CN1.CommitTrans
    
    If iCount > 0 And Cbo_trans_way.ListIndex = 0 Then
        Call carprice       '传erp运费计算
    End If
    
    If iCount > 0 Then

        If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call MenuTool_ReSet
        End If
        txt_car_no.Text = ""
        txt_trans_comp.Text = ""
        txt_trans_tool.Text = ""
        txt_trans_comp_name.Text = ""
        txt_trans_tool_name.Text = ""
        text_send_no.Text = ""

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
    SQL = SQL & "     FROM  CP_MOVE_SLT                                     " & vbCrLf
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
                            adoCmd.Parameters(iCol).Value = Trim(Str(sc1.Item("Spread").Value))
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
    
'   If Val(text_send_plan_wgt.Text) > 0 Then
    If ss1.Text <> "Update" Then
        ss1.Col = 0:            ss1.Text = "Update"
        ss1.Col = SPD_EMP_CD:   ss1.Text = sUserID
        ss1.Col = 1:            ss1.Text = text_send_no.Text
        ss1.Col = SPD_WGT
        sdb_slab_num.Value = sdb_slab_num.Value + 1
        sdb_slab_wgt.Value = sdb_slab_wgt.Value + ss1.Value
        text_send_plan_wgt.Text = Str(Val(text_send_plan_wgt.Text) - ss1.Value)
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
    Else
        ss1.Col = 0:             ss1.Text = ""
        ss1.Col = SPD_EMP_CD:    ss1.Text = ""
        ss1.Col = 1:            ss1.Text = ""
        ss1.Col = SPD_WGT
        sdb_slab_num.Value = sdb_slab_num.Value - 1
        sdb_slab_wgt.Value = sdb_slab_wgt.Value - ss1.Value
        text_send_plan_wgt.Text = Str(Val(text_send_plan_wgt.Text) + ss1.Value)
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
   
    End If
'  Else
'    Call MsgBox("请输入派车申请单号或者装车超重")
'  End If

End Sub

Private Sub stx_move_time_DblClick()

    stx_move_time.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'HH24MISS') FROM DUAL")
    
End Sub

Private Sub Text_PROD_CD_Change()

    Select Case text_prod_cd.Text
'        Case "S", "s", "SL"
'            text_prod_cd.Text = "SL"
        Case "P", "p", "PP"
            text_prod_cd.Text = "PP"
        Case "H", "h", "HC"
            text_prod_cd.Text = "HC"
        Case "", "**"
            text_prod_cd.Text = ""
        Case Else
            text_prod_cd.Text = ""
            Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
        End Select
        
    Call Gp_Ms_Cls(Mc1("rControl"))
    
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

Private Sub text_send_no_DblClick()

    Call text_send_no_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_EndUse_s_DblClick()

    Call txt_EndUse_s_KeyUp(vbKeyF4, 0)
    
End Sub
Private Sub text_send_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    Dim iRow        As Integer
    Dim iCount        As Integer
    
    DD.DataDicType = "U"        'Order Usage Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=text_send_no
        DD.rControl.Add Item:=text_send_plan_wgt

        DD.sQuery = "            SELECT SEND_NO ""派车申请单号"" ,PLAN_WGT-FIN_WGT ""派车剩余重量""FROM NISCO.CP_MOVE_PLAN "
        DD.sQuery = DD.sQuery + " WHERE REC_STS             IN   ('2','3') "
        DD.sWhere = DD.sWhere + "   AND SEND_NO            like   '" & Trim(DD.rControl.Item(1).Text) & "%' "
        
        
        DD.sWhere = DD.sWhere + " ORDER  BY  SEND_NO  ASC "
        
     If Gf_DD_Display(M_CN1, DD.sQuery + DD.sWhere, False) Then
    
'        If DD.sWitch = "SP" Then
'
'            DD.sPname.Col = DD.rControl.Item(1)
'            sNew_Code = DD.sPname.Text
'
'            If DD.rControl.Count > 1 Then
'                DD.sPname.Col = DD.rControl.Item(2)
'                sNew_Name = DD.sPname.Text
'            End If
'
'            DD.sPname.TabStop = True
'            DD.sPname.SetFocus
'            DD.sPname.SetActiveCell DD.rControl.Item(1), DD.sPname.ActiveRow
'            DD.sPname.Action = SS_ACTION_ACTIVE_CELL
'            DD.sPname.EditMode = True
'            DD.sPname.TabStop = False
'
'            If DD.sSelect Then
'                If sOld_Code <> sNew_Code Then Call Gp_Sp_UpdateMake(DD.sPname, False)
'            End If
'        End If
    
    End If
    
    DD.sWitch = ""
    DD.sSelect = False
    
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        ss1.Col = 0
        
        If sdb_slab_num.Value = iCount Then Exit For
        
        If ss1.Text = "Update" Then
           ss1.Col = 1
           ss1.Text = text_send_no.Text
           ss1.Col = 12
           text_send_plan_wgt.Text = Val(text_send_plan_wgt.Text) - Val(ss1.Text)
           
        End If
    Next iRow
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing
         
    End If
    
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

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=txt_plt_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
            txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_plt.Text, 2)
        Else
              txt_plt_name.Text = ""
        End If
    
    End If
    
End Sub

Private Sub txt_trans_tool_DblClick()

    Call txt_trans_tool_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_trans_tool_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0015"

        DD.rControl.Add Item:=txt_trans_tool
        DD.rControl.Add Item:=txt_trans_tool_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_trans_tool.Text)) = txt_trans_tool.MaxLength Then
            txt_trans_tool_name.Text = Gf_ComnNameFind(M_CN1, "C0015", txt_trans_tool.Text, 2)
        Else
          txt_trans_tool_name.Text = ""
        End If
        
    End If
    
End Sub
Private Sub txt_trans_comp_DblClick()

    Call txt_trans_comp_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_trans_comp_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0016"

        DD.rControl.Add Item:=txt_trans_comp
        DD.rControl.Add Item:=txt_trans_comp_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_trans_comp.Text)) = txt_trans_comp.MaxLength Then
            txt_trans_comp_name.Text = Gf_ComnNameFind(M_CN1, "C0016", txt_trans_comp.Text, 2)
        Else
          txt_trans_comp_name.Text = ""
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
        DD.rControl.Add Item:=sdb_car_max_wgt

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
    
        DD.sQuery = "SELECT CAR_NO, CAR_WGT_MAX, CAR_WGT_AVE, CAR_KND, CAR_CMP_CD, Gf_Comnnamefind('H0002',CAR_CMP_CD) AS CAR_CMP_NAME FROM  HP_CAR_IMF "
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

        Call Gf_StdSPEC_DD(M_CN1, KeyCode)
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
    
    If txt_trns_no.Text = "" Then
        Call MsgBox("装车单为空，传运费系统失败", vbInformation, "系统提示信息")
        Screen.MousePointer = vbDefault
        Exit Function
    End If
          
    sQuery = "{call ARC0180P( 'N','" + txt_trns_no.Text + "',?)}"
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

