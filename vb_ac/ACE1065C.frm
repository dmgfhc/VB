VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACE1065C 
   Caption         =   "物料替代_ACE1065C"
   ClientHeight    =   9225
   ClientLeft      =   270
   ClientTop       =   1395
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter3 
      Height          =   9105
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   16060
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "ACE1065C.frx":0000
      Begin Threed.SSFrame SSFrame1 
         Height          =   1305
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   2302
         _Version        =   196609
         BackColor       =   14737632
         ShadowStyle     =   1
         Begin VB.TextBox txt_STLGRD_Name 
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
            Left            =   7095
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   26
            Tag             =   "钢种(标准号)"
            Top             =   120
            Width           =   1890
         End
         Begin VB.ComboBox ord_ord_item 
            Height          =   315
            Left            =   3015
            TabIndex        =   25
            Top             =   510
            Width           =   735
         End
         Begin VB.TextBox ord_ord_no 
            Height          =   315
            Left            =   1545
            MaxLength       =   11
            TabIndex        =   24
            Top             =   510
            Width           =   1470
         End
         Begin VB.TextBox ord_txt_prod_cd 
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
            Height          =   310
            Left            =   1545
            MaxLength       =   2
            TabIndex        =   23
            Tag             =   "订单产品类型"
            Top             =   120
            Width           =   420
         End
         Begin VB.TextBox ord_TxT_STLGRD 
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
            Left            =   5595
            MaxLength       =   11
            TabIndex        =   22
            Top             =   120
            Width           =   1500
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   0
            Left            =   180
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
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
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Index           =   0
            Left            =   4215
            Top             =   510
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "订单输入日期"
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
         Begin InDate.UDate UDate_DEL_TO_b 
            Height          =   315
            Left            =   5595
            TabIndex        =   27
            Tag             =   "INS_DATE"
            Top             =   510
            Width           =   1485
            _ExtentX        =   2619
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
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Index           =   0
            Left            =   4215
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "钢种"
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
            Index           =   0
            Left            =   180
            Top             =   915
            Width           =   1335
            _ExtentX        =   2355
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
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Index           =   0
            Left            =   4215
            Top             =   915
            Width           =   1335
            _ExtentX        =   2355
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Index           =   1
            Left            =   8415
            Top             =   915
            Width           =   1335
            _ExtentX        =   2355
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
         Begin CSTextLibCtl.sidbEdit ord_prod_thk_fr 
            Height          =   315
            Left            =   1545
            TabIndex        =   28
            Top             =   915
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
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
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit ord_prod_thk_to 
            Height          =   315
            Left            =   2580
            TabIndex        =   29
            Top             =   915
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
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
            MaxValue        =   999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit ord_prod_wid_fr 
            Height          =   315
            Left            =   5595
            TabIndex        =   30
            Top             =   915
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
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
         Begin CSTextLibCtl.sidbEdit ord_prod_wid_to 
            Height          =   315
            Left            =   6630
            TabIndex        =   31
            Top             =   915
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
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
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit ord_prod_len_fr 
            Height          =   315
            Left            =   9795
            TabIndex        =   32
            Top             =   915
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
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
            NumIntDigits    =   7
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit ord_prod_len_to 
            Height          =   315
            Left            =   10830
            TabIndex        =   33
            Top             =   915
            Width           =   1035
            _Version        =   262145
            _ExtentX        =   1826
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
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSCommand cmd_confirm 
            Height          =   450
            Left            =   13485
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   780
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   794
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
            Caption         =   "替代确定处理"
            BevelWidth      =   3
         End
         Begin Threed.SSCommand Command_REP 
            Height          =   450
            Left            =   13485
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   120
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   794
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
            Caption         =   "替代处理"
            BevelWidth      =   3
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   2
            Left            =   180
            Top             =   510
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "订单号"
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
      End
      Begin SSSplitter.SSSplitter SSSplitter1 
         Height          =   7770
         Left            =   0
         TabIndex        =   9
         Top             =   1335
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   13705
         _Version        =   196609
         SplitterBarWidth=   4
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   16761087
         PaneTree        =   "ACE1065C.frx":0052
         Begin SSSplitter.SSSplitter SSSplitter2 
            Height          =   3750
            Left            =   0
            TabIndex        =   10
            Top             =   4020
            Width           =   15165
            _ExtentX        =   26749
            _ExtentY        =   6615
            _Version        =   196609
            SplitterBarWidth=   2
            SplitterBarJoinStyle=   0
            SplitterBarAppearance=   0
            BorderStyle     =   0
            BackColor       =   14737632
            PaneTree        =   "ACE1065C.frx":00A4
            Begin Threed.SSPanel SSPanel1 
               Height          =   660
               Left            =   0
               TabIndex        =   11
               Top             =   0
               Width           =   15165
               _ExtentX        =   26749
               _ExtentY        =   1164
               _Version        =   196609
               BackColor       =   14737918
               BevelOuter      =   1
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
               Begin VB.TextBox txt_cur_inv 
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
                  Left            =   11970
                  MaxLength       =   2
                  TabIndex        =   36
                  Tag             =   "钢种"
                  Top             =   90
                  Width           =   480
               End
               Begin VB.TextBox prod_no 
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
                  Left            =   3060
                  MaxLength       =   14
                  TabIndex        =   17
                  Top             =   90
                  Width           =   1695
               End
               Begin VB.TextBox prod_txt_prod_cd 
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
                  Left            =   1275
                  MaxLength       =   2
                  TabIndex        =   16
                  Tag             =   "替代产品类型"
                  Text            =   "SL"
                  Top             =   90
                  Width           =   435
               End
               Begin VB.TextBox prod_ord_no 
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
                  Left            =   8690
                  MaxLength       =   11
                  TabIndex        =   15
                  Tag             =   "订单号"
                  Top             =   90
                  Width           =   1320
               End
               Begin VB.TextBox prod_txt_stlgrd 
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
                  Left            =   6030
                  MaxLength       =   11
                  TabIndex        =   14
                  Tag             =   "钢种"
                  Top             =   90
                  Width           =   1350
               End
               Begin VB.TextBox prod_loc 
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
                  Left            =   13770
                  TabIndex        =   13
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.ComboBox prod_ord_itm 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   10005
                  TabIndex        =   12
                  Top             =   90
                  Width           =   645
               End
               Begin InDate.ULabel ULabel2 
                  Height          =   315
                  Index           =   1
                  Left            =   165
                  Top             =   90
                  Width           =   1095
                  _ExtentX        =   1931
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
                  Index           =   1
                  Left            =   4920
                  Top             =   90
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
                  Caption         =   "钢种"
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
                  Index           =   1
                  Left            =   7575
                  Top             =   90
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
                  Caption         =   "订单号"
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
               Begin InDate.ULabel ULabel7 
                  Height          =   315
                  Index           =   1
                  Left            =   12720
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
                  Caption         =   "物料位置"
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
               Begin InDate.ULabel ULabel7 
                  Height          =   315
                  Index           =   2
                  Left            =   1950
                  Top             =   90
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
                  Caption         =   "物料号"
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
                  Index           =   3
                  Left            =   10860
                  Top             =   90
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
                  Caption         =   "堆放仓库"
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
               Begin Threed.SSCheck chk_use 
                  Height          =   270
                  Left            =   14100
                  TabIndex        =   37
                  Top             =   105
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   476
                  _Version        =   196609
                  Font3D          =   1
                  BackStyle       =   1
                  ActiveColors    =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "宋体"
                     Size            =   9.75
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin InDate.ULabel ULabel5 
                  Height          =   315
                  Left            =   12720
                  Top             =   90
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   556
                  Caption         =   "强制替代"
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
               Begin VB.Label Label4 
                  BackStyle       =   0  'Transparent
                  Caption         =   "——"
                  Height          =   375
                  Left            =   10425
                  TabIndex        =   18
                  Top             =   150
                  Width           =   255
               End
            End
            Begin FPSpread.vaSpread prod_ss 
               Height          =   3060
               Left            =   0
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   690
               Width           =   15165
               _Version        =   393216
               _ExtentX        =   26749
               _ExtentY        =   5398
               _StockProps     =   64
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
               MaxCols         =   27
               MaxRows         =   2
               ProcessTab      =   -1  'True
               RetainSelBlock  =   0   'False
               SpreadDesigner  =   "ACE1065C.frx":00F6
               VisibleCols     =   1
            End
         End
         Begin FPSpread.vaSpread ord_ss 
            Height          =   3960
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   15165
            _Version        =   393216
            _ExtentX        =   26749
            _ExtentY        =   6985
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
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "ACE1065C.frx":0E4C
         End
      End
   End
   Begin CSTextLibCtl.sidbEdit prod_prod_wgt_to 
      Height          =   315
      Left            =   13740
      TabIndex        =   0
      Top             =   9765
      Visible         =   0   'False
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
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
      RawData         =   "99999.99"
      Text            =   " 0.00"
      StartText.x     =   3
      StartText.y     =   4
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   13
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      Undo            =   0
      Data            =   99999.99
   End
   Begin CSTextLibCtl.sidbEdit prod_prod_wgt_fr 
      Height          =   315
      Left            =   12540
      TabIndex        =   1
      Top             =   9765
      Visible         =   0   'False
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
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
      StartText.y     =   4
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   13
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit prod_prod_thk_fr 
      Height          =   315
      Left            =   945
      TabIndex        =   2
      Tag             =   "产品厚度（MIN）"
      Top             =   9765
      Visible         =   0   'False
      Width           =   1395
      _Version        =   262145
      _ExtentX        =   2461
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
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Index           =   1
      Left            =   -45
      Top             =   9765
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   3855
      Top             =   9765
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   0
      Left            =   7635
      Top             =   9765
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
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
   Begin CSTextLibCtl.sidbEdit prod_prod_thk_to 
      Height          =   315
      Left            =   2340
      TabIndex        =   3
      Tag             =   "产品厚度（MAX）"
      Top             =   9765
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
      RawData         =   "9999.99"
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
      Data            =   9999.99
   End
   Begin CSTextLibCtl.sidbEdit prod_prod_len_fr 
      Height          =   315
      Left            =   8595
      TabIndex        =   4
      Tag             =   "产品长度（MIN）"
      Top             =   9765
      Visible         =   0   'False
      Width           =   1395
      _Version        =   262145
      _ExtentX        =   2461
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
   Begin CSTextLibCtl.sidbEdit prod_prod_len_to 
      Height          =   315
      Left            =   9990
      TabIndex        =   5
      Tag             =   "产品长度（MIN）"
      Top             =   9765
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
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "9999999"
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
      Data            =   9999999
   End
   Begin CSTextLibCtl.sidbEdit prod_prod_wid_fr 
      Height          =   315
      Left            =   4830
      TabIndex        =   6
      Tag             =   "产品宽度（MIN）"
      Top             =   9765
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
      NumIntDigits    =   4
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit prod_prod_wid_to 
      Height          =   315
      Left            =   6105
      TabIndex        =   7
      Tag             =   "产品宽度（MAX）"
      Top             =   9765
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
      RawData         =   "999999"
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
      NumIntDigits    =   4
      Undo            =   0
      Data            =   999999
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   2
      Left            =   11580
      Top             =   9765
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "产品重量"
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
Attribute VB_Name = "ACE1065C"
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
'-- Program ID        ACE1065C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Yang Zhibin
'-- Date              2003.9.8
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
Dim t_ord_no As String
Dim t_ord_item As String
Dim sdel As String

Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting
Public Active_CForm As String       'Form Active

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
Dim nColumn1 As New Collection      'Spread necessary Column1 Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column1 Collection
Dim iColumn1 As New Collection      'Spread Insert Column1 Collection
Dim aColumn1 As New Collection      'Master -> Spread Column1 Collection
Dim lColumn1 As New Collection      'Spread Lock Column1 Collection

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column1 Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column1 Collection
Dim iColumn2 As New Collection      'Spread Insert Column1 Collection
Dim aColumn2 As New Collection      'Master -> Spread Column1 Collection
Dim lColumn2 As New Collection      'Spread Lock Column1 Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim ord_sc As New Collection           'order Spread Collection
Dim prod_sc As New Collection          'product spread collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column1

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iCount As Integer
Const PROD_SS_MLT_PROC_CD = 25            ' 炼钢工艺流程
Const PROD_SS_CUR_INV = 26               ' 仓库


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
  'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(ord_txt_prod_cd, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(ord_TxT_STLGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(ord_ord_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(ord_ord_item, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'       Call Gp_Ms_Collection(txt_cust_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(ord_prod_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(ord_prod_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(ord_prod_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(ord_prod_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(ord_prod_len_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(ord_prod_len_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(UDate_DEL_TO_b, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    'Call Spread_Collection("Column1_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ord_ss, 1, "p", "n", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ord_ss, 2, "p", "n", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ord_ss, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ord_ss, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ord_ss, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ord_ss, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ord_ss, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ord_ss, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ord_ss, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ord_ss, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    ord_sc.Add Item:=ord_ss, Key:="Spread"
    ord_sc.Add Item:="ACE1010C.P_MODIFY", Key:="P-M"
    ord_sc.Add Item:="ACE1065C.P_SREFER1", Key:="P-R"
    ord_sc.Add Item:=pColumn1, Key:="pColumn"
    ord_sc.Add Item:=nColumn1, Key:="nColumn"
    ord_sc.Add Item:=aColumn1, Key:="aColumn"
    ord_sc.Add Item:=mColumn1, Key:="mColumn"
    ord_sc.Add Item:=iColumn1, Key:="iColumn"
    ord_sc.Add Item:=lColumn1, Key:="lColumn"
    ord_sc.Add Item:=1, Key:="First"
    ord_sc.Add Item:=ord_ss.MaxCols, Key:="ord_Last"
   
   
     '  Call Gp_Ms_Collection(prod_txt_prod_cd, "p", "n ", " ", " ", "r", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(prod_txt_prod_cd, "p", "n", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(ord_txt_prod_cd, "p", "n", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(prod_no, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(prod_txt_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(prod_ord_no, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
'     Call Gp_Ms_Collection(Text_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(prod_ord_itm, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(prod_loc, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)

   'Call Gp_Ms_Collection(prod_prod_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   'Call Gp_Ms_Collection(prod_prod_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   'Call Gp_Ms_Collection(prod_prod_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   'Call Gp_Ms_Collection(prod_prod_wid_to, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   'Call Gp_Ms_Collection(prod_prod_len_fr, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   'Call Gp_Ms_Collection(prod_prod_len_to, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                 
'    Call Gp_Ms_Collection(prod_prod_wgt_fr, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
'    Call Gp_Ms_Collection(prod_prod_wgt_to, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
'    Call Gp_Ms_Collection(prod_combo, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)

       'MASTER Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
   
     Call Gp_Sp_Collection(prod_ss, 1, " ", "n", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(prod_ss, 2, "p", "n", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(prod_ss, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(prod_ss, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(prod_ss, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(prod_ss, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(prod_ss, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(prod_ss, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(prod_ss, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 19, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 20, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 21, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 23, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 24, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 25, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 26, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(prod_ss, 27, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    

    'Spread_Collection
    prod_sc.Add Item:=prod_ss, Key:="Spread"
    prod_sc.Add Item:="ACE1065C.P_SREFER2", Key:="P-R"
    prod_sc.Add Item:=pColumn2, Key:="pColumn"
    prod_sc.Add Item:=nColumn2, Key:="nColumn"
    prod_sc.Add Item:=aColumn2, Key:="aColumn"
    prod_sc.Add Item:=mColumn2, Key:="mColumn"
    prod_sc.Add Item:=iColumn2, Key:="iColumn"
    prod_sc.Add Item:=lColumn2, Key:="lColumn"
    prod_sc.Add Item:=1, Key:="First"
    prod_sc.Add Item:=prod_ss.MaxCols, Key:="prod_Last"

   ' Proc_Sc.Add Item:=sc1, Key:="Sc"
    Proc_Sc.Add Item:=ord_sc, Key:="oSc"
    Proc_Sc.Add Item:=prod_sc, Key:="pSc"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(prod_ss, PROD_SS_CUR_INV, True)

End Sub

Private Sub cmd_confirm_Click()     '确定替代处理

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call ACE1210P ('C1','','" + sUserID + "',?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
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
        Call Gp_MsgBoxDisplay("确定处理完了..!!", "I")
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

Private Sub Command_REP_Click()           '替代处理

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    Dim lRow As Integer
    
    Dim adoCmd As ADODB.Command

    Dim str_ord_prod_cd As String
    Dim str_prod_prod_cd As String
    Dim str_ord_no As String
    Dim str_ord_item As String
    Dim str_prod_loc As String
    Dim str_prod_stlgrd As String
    Dim str_prod_no As String
    Dim str_prod_force_cd As String

    Screen.MousePointer = vbHourglass

    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256

    str_ord_prod_cd = ord_txt_prod_cd.Text
    str_prod_prod_cd = prod_txt_prod_cd.Text
    str_prod_stlgrd = prod_txt_stlgrd.Text
    str_ord_no = prod_ord_no.Text
    str_ord_item = prod_ord_itm.Text
    str_prod_loc = prod_loc.Text
    str_prod_no = prod_no.Text
    
    If chk_use.Value = "-1" Then
        str_prod_force_cd = "1"
    End If
   
    'ElseIf Len(str_prod_no) < 7 Then

    If Len(str_prod_loc) > 0 Then
           If Left(str_prod_loc, 1) <> Left(prod_txt_prod_cd.Text, 1) Then
               Call MsgBox("产品代码与物料位置不符！" & Chr(10) & "请检查后输入。", vbExclamation + vbOKOnly, "警告")
               Screen.MousePointer = vbDefault
               Exit Sub
           End If
    End If
            
'     For lRow = 1 To prod_ss.MaxRows
'
'        prod_ss.Row = lRow
'        prod_ss.Col = 2
'
'        If prod_ss.Text = str_prod_no Then
'            prod_ss.Col = PROD_SS_FORCE_CD
'            str_prod_force_cd = prod_ss.Text
'
'        End If
'
'    Next lRow
    
    If str_ord_prod_cd <> "" Then
    
        If str_prod_prod_cd <> "" Then
        
            If str_ord_prod_cd = "PP" And str_prod_prod_cd = "PP" Then
                sQuery = "{call ACE1080P('" + str_ord_no + "','" + str_ord_item + "','" + str_prod_no + "','" + str_prod_loc + "','" + str_prod_stlgrd + "',?)}"
            
            ElseIf str_ord_prod_cd = "HC" And str_prod_prod_cd = "HC" Then
                sQuery = "{call ACE1070P('" + str_ord_no + "','" + str_ord_item + "','" + str_prod_no + "','" + str_prod_loc + "','" + str_prod_stlgrd + "',?)}"
            
            ElseIf str_ord_prod_cd = "SL" And str_prod_prod_cd = "SL" Then
                sQuery = "{call ACE1090P('" + str_ord_no + "','" + str_ord_item + "','" + str_prod_no + "','" + str_prod_loc + "','" + str_prod_stlgrd + "',?)}"

            ElseIf (str_ord_prod_cd = "HC" Or str_ord_prod_cd = "PP") And str_prod_prod_cd = "SL" Then
                sQuery = "{call ACE1100P ('" + str_ord_no + "','" + str_ord_item + "','" + str_prod_no + "','" + txt_cur_inv.Text + "','" + str_prod_loc + "','" + str_prod_stlgrd + "','" + str_prod_force_cd + "',?)}"
            
            Else
                Call MsgBox("订单与物料产品代码输入错误" & Chr(10) & "请重试。", vbExclamation + vbOKOnly, "警告")
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
      
                            '    'Ado Setting
            M_CN1.CursorLocation = adUseServer
            Set adoCmd = New ADODB.Command
            
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
               
                Call Gp_MsgBoxDisplay("替代处理完了..!!", "I")
                'Call Form_Cls
                prod_no.Text = ""
                Call Form_Ref
                Call Gp_Ms_Cls(Mc2("rControl"))
                Call Gf_Sp_Cls(Proc_Sc("pSc"))
                'Call Prod_ss_Ref
            End If
       
            Set adoCmd = Nothing
            Screen.MousePointer = vbDefault
            Exit Sub
                
        Else
            Call MsgBox("产品分类代码不能为空！" & Chr(10) & "请重试。", vbExclamation + vbOKOnly, "警告")
            prod_txt_prod_cd.Text = ""
            prod_txt_prod_cd.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        Call MsgBox("产品分类代码不能为空！" & Chr(10) & "请重试。", vbExclamation + vbOKOnly, "警告")
        ord_txt_prod_cd.Text = ""
        ord_txt_prod_cd.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
        
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    With MDIMain.MenuTool
        .Buttons(4).Enabled = True                 'Save
        .Buttons(9).Enabled = True                 'Delete
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
    End With

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
    'sAuthority = "0001"
    
    Call Form_Define
    
    Call Gp_Sp_Setting(Proc_Sc("oSc")("Spread"), False)
    Call Gp_Sp_Setting(Proc_Sc("pSc")("Spread"), False)
   ' Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gf_Sp_Cls(Proc_Sc("oSc"))
    Call Gf_Sp_Cls(Proc_Sc("pSc"))
    
    Call Gp_Spl_SizeGet(SSSplitter1, "C-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(Proc_Sc("oSc")("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("pSc")("Spread"), "C-System.INI", Me.Name)
    
    ord_txt_prod_cd.Text = "PP"
    
     If Mid(sAuthority, 3, 1) <> "1" Then
''             Command_ALLSELECT.Enabled = False
             Command_REP.Enabled = False
             cmd_confirm.Enabled = False
     End If

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        
    Call Gp_Spl_SizeSet(SSSplitter1, "C-System.INI", Me.Name)

    Call Gp_Sp_ColSet(Proc_Sc("OSc")("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("PSc")("Spread"), "C-System.INI", Me.Name)
    
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
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set ord_sc = Nothing
    Set prod_sc = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("oSc")) And Gf_Sp_Cls(Proc_Sc("pSc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        With MDIMain.MenuTool
                .Buttons(4).Enabled = True                 'Save
                .Buttons(9).Enabled = True                 'Delete
                .Buttons(11).Enabled = False                'Copy
                .Buttons(12).Enabled = False                'Paste
        End With
    End If
    
    UDate_DEL_TO_b.RawData = ""
    ord_TxT_STLGRD.Text = ""
   
    ord_prod_thk_fr = 0
    ord_prod_thk_to = 0
    ord_prod_wid_fr.Text = 0
    ord_prod_wid_to.Text = 0
    ord_prod_len_fr.Text = 0
    ord_prod_len_to.Text = 0
    ord_txt_prod_cd.Text = "PP"
    ord_TxT_STLGRD.Text = ""
    'txt_cust_cd.Text = ""
    UDate_DEL_TO_b.Text = ""
    
    'prod_txt_prod_cd.Text = ""
    prod_txt_stlgrd.Text = ""
    prod_ord_no.Text = ""
    prod_ord_itm.Text = ""
    prod_loc.Text = ""
    prod_no.Text = ""
    ord_ord_no.Text = ""
    ord_ord_item.Text = ""
    
    prod_prod_thk_fr.Text = 0
    prod_prod_thk_to.Text = 0
    prod_prod_wid_fr.Text = 0
    prod_prod_wid_to.Text = 0
    prod_prod_len_fr.Text = 0
    prod_prod_len_to.Text = 0
    prod_prod_wgt_fr.Text = 0
    prod_prod_wgt_to.Text = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("ord_Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()
   
   On Error GoTo Refer_Err

    Dim SMESG As String
    
    If Gf_Sp_ProceExist(Proc_Sc("oSc").Item("Spread")) Then Exit Sub
    
        If Gf_Sp_Refer(M_CN1, Proc_Sc("oSc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            ord_ss.OperationMode = OperationModeNormal
             With MDIMain.MenuTool
                .Buttons(4).Enabled = True                 'Save
                .Buttons(9).Enabled = True                 'Delete
                .Buttons(11).Enabled = False                'Copy
                .Buttons(12).Enabled = False                'Paste
            End With

            Exit Sub
        End If
            
    Exit Sub

Refer_Err:

'    Call ord_sel
'    Call prod_sel

End Sub
Public Sub Form_Pro()
  
End Sub

Public Sub Spread_Column1sSort()

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

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc")("ord_Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub ord_ss_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("osc")("Spread"), Col, Row)
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0


End Sub

Private Sub ord_ss_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim SMESG As String
    If ord_ss.MaxRows < 1 Or Row < 1 Then Exit Sub
  
    ord_ss.BlockMode = True
    ord_ss.Row = 1
    ord_ss.Row2 = ord_ss.MaxRows
    ord_ss.Col = -1
    ord_ss.BackColor = &HFFFFFF
    ord_ss.BlockMode = False
    
    ord_ss.Row = Row
    ord_ss.BackColor = &HFFFFC1
    
    ord_ss.Row = Row
    
    ord_ss.Col = 1
    prod_ord_no.Text = ord_ss.Text
    
    ord_ss.Col = 2
    prod_ord_itm.Text = Trim(ord_ss.Value)
    
    prod_no.Text = ""
    prod_loc.Text = ""
    
    If prod_txt_prod_cd.Text = "" Then prod_txt_prod_cd.Text = "SL"
    'Call prod_sel
    Call Gf_Sp_Cls(Proc_Sc("pSc"))
    
    Call Prod_ss_Ref
    
End Sub

Private Sub ord_ss_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("oSC")("spread"), Mode)
        'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
    End If

End Sub

Private Sub ord_ss_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ord_ss_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ord_ss
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub ord_txt_prod_cd_DblClick()

    Call ord_txt_prod_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub ord_TxT_STLGRD_DblClick()

    Call ord_TxT_STLGRD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub prod_ss_Click(ByVal Col As Long, ByVal Row As Long)

 
    Call Gp_Sp_Sort(Proc_Sc("PSc")("Spread"), Col, Row)
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
End Sub

Private Sub prod_ss_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Col <> 2 And Col <> 3 And Col <> 4 Then Exit Sub
    If prod_ss.MaxRows < 1 Or Row < 1 Then Exit Sub
    
    prod_ss.Row = Row
    prod_ss.Col = Col
    
    If Col = 2 Then
        prod_no.Text = prod_ss.Text
    ElseIf Col = 3 Then
        prod_ss.Col = PROD_SS_CUR_INV
        txt_cur_inv.Text = prod_ss.Text
    ElseIf Col = 4 Then
        prod_loc.Text = prod_ss.Text
    End If
        
End Sub

Private Sub prod_ss_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("pSC")("spread"), Mode)
        'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
    End If
    
End Sub

Private Sub prod_ss_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub prod_ss_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.prod_ss
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub

Private Sub ord_txt_prod_cd_Change()
       
    If Len(ord_txt_prod_cd) <> 2 Then Exit Sub

    Select Case ord_txt_prod_cd.Text

        Case "PP", "pp"
            ord_txt_prod_cd.Text = "PP"
        Case "SL", "sl"
            ord_txt_prod_cd.Text = "SL"
        Case "HC", "hc"
            ord_txt_prod_cd.Text = "HC"
        Case Else
            ord_txt_prod_cd.Text = ""
            Call MsgBox("产品分类代码应该为 ( PP , HC , SL )，请更正", vbExclamation + vbOKOnly, "警告")
    End Select
     
End Sub

Private Sub ord_txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then  '          '          '                 '                 ' ''''''''''
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=ord_txt_prod_cd
'        DD.rControl.Add Item:=Text_PROD_CD_mate
   
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() 用于客户代码

        Exit Sub
        
    End If

'    If Len(Trim(ord_txt_prod_cd.Text)) = ord_txt_prod_cd.MaxLength Then
'
'        Text_PROD_CD_mate.Text = Gf_ComnNameFind(M_CN1, "B0005", ord_txt_prod_cd.Text, 2)
'    Else
'        Text_PROD_CD_mate.Text = ""
'    End If
    
End Sub

Private Sub ord_TxT_STLGRD_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then
            
        DD.nameType = "1"
        DD.sWitch = "MS"
        DD.rControl.Add Item:=ord_TxT_STLGRD
        DD.rControl.Add Item:=txt_STLGRD_Name
        
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(ord_TxT_STLGRD.Text)) >= 10 Then
            txt_STLGRD_Name.Text = Gf_StlgrdNameFind(M_CN1, Trim(ord_TxT_STLGRD.Text))
        Else
            txt_STLGRD_Name.Text = ""
        End If
        
    End If
        
End Sub

Private Sub prod_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String
    
    If Len(Trim(prod_ord_no.Text)) = prod_ord_no.MaxLength Then
    
        If prod_ord_itm.Text <> "" Then Exit Sub
        
        prod_ord_no.Text = StrConv(prod_ord_no.Text, vbUpperCase)
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(prod_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, prod_ord_itm, sQuery)
        
       ' If combo_ord_item.ListCount <> 0 Then
       '       combo_ord_item.ListIndex = 0
       ' End If
    Else
        prod_ord_itm.Clear
    End If

End Sub

Private Sub ord_ss_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    Dim i As Integer
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

'    Dim ROW1 As Long
'    Dim row2 As Long
'    Dim Col As Long
'
'    Col = BlockCol
'    ROW1 = BlockRow
'    row2 = BlockRow2
'
'    If Col = -1 Then
'        For i = BlockRow To BlockRow2
'           ord_ss.Row = i
'           ord_ss.Col = 0
'           If ord_ss.Text = "Delete" Then
'              ord_ss.Text = ""
'               Call Gp_Sp_BlockColor(ord_ss, 1, ord_ss.MaxCols, ROW1, row2)
'           Else: ord_ss.Text = "Delete"
'            Call Gp_Sp_BlockColor(ord_ss, 1, ord_ss.MaxCols, ROW1, row2, , &HFFFF80)
'           End If
'
'        Next
'    End If

End Sub

Private Sub prod_ss_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    Dim i As Integer
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0


     

'    Dim ROW1 As Long
'    Dim row2 As Long
'    Dim Col As Long
'    Col = BlockCol
'    ROW1 = BlockRow
'    row2 = BlockRow2
'    If Col = -1 Then
'     For i = BlockRow To BlockRow2
'        prod_ss.ROW = i
'        prod_ss.Col = 0
'        If prod_ss.Text = "Delete" Then
'           prod_ss.Text = ""
'            Call Gp_Sp_BlockColor(prod_ss, 1, prod_ss.MaxCols, ROW1, row2)
'        Else: prod_ss.Text = "Delete"
'         Call Gp_Sp_BlockColor(prod_ss, 1, prod_ss.MaxCols, ROW1, row2, , &HFFFF80)
'        End If
'
'     Next
'   End If

End Sub


'-------------------------------------------------------------------------------------
'--------------------------------------物料-------------------------------------------
'-------------------------------------------------------------------------------------

Private Sub prod_ord_itm_KeyPress(KeyAscii As Integer)
    'KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub prod_ord_itm_LostFocus()

    Dim S As String
  
    If Len(prod_ord_itm.Text) = 1 Then
        S = prod_ord_itm.Text
        prod_ord_itm.Text = "0" + S
    End If
    
End Sub

Private Sub prod_txt_prod_cd_Change()

    Dim sProdChk As String
    
    sProdChk = ""
    
    If Len(prod_txt_prod_cd) <> 2 Then Exit Sub

    Select Case prod_txt_prod_cd.Text
        Case "SL", "sl"
            prod_txt_prod_cd.Text = "SL"
        Case Else
            prod_txt_prod_cd.Text = ""
            Call MsgBox("产品分类代码应该为 ( SL )，请更正", vbExclamation + vbOKOnly, "警告")
    End Select
    
    If prod_txt_prod_cd.Text = "" Then Exit Sub
    
    Select Case ord_txt_prod_cd.Text
           Case "SL"
               If prod_txt_prod_cd.Text <> "SL" Then sProdChk = "Err"
           Case "PP"
               If prod_txt_prod_cd.Text <> "SL" And prod_txt_prod_cd.Text <> "PP" Then sProdChk = "Err"
           Case "HC"
               If prod_txt_prod_cd.Text <> "SL" And prod_txt_prod_cd.Text <> "HC" Then sProdChk = "Err"
    End Select
    
    If sProdChk = "Err" Then
        Call MsgBox("产品分类代码应该为 ( SL )，请更正", vbExclamation + vbOKOnly, "警告")
        sProdChk = ""
        prod_txt_prod_cd.Text = ""
    End If

End Sub

Private Sub prod_txt_prod_cd_DblClick()

    Call prod_txt_prod_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub prod_txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

   If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=prod_txt_prod_cd
'        DD.rControl.Add Item:=Text_PROD_CD_mate

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() 用于客户代码

        Exit Sub

    End If
'
'    If Len(Trim(prod_txt_prod_cd.Text)) = prod_txt_prod_cd.MaxLength Then
'
'        Text_PROD_CD_mate.Text = Gf_ComnNameFind(M_CN1, "B0005", prod_txt_prod_cd.Text, 2)
'    Else
'        Text_PROD_CD_mate.Text = ""
'    End If


End Sub

Private Sub prod_Txt_PROD_CD_LostFocus()

    If prod_txt_prod_cd.Text <> "" Then
        If (Len(prod_txt_prod_cd.Text) < prod_txt_prod_cd.MaxLength) Then
            Call Gp_MsgBoxDisplay("产品分类不符合规范！")
            'Text_PROD_CD.Text = ""
            prod_txt_prod_cd.SetFocus
        End If
    End If

End Sub

Private Sub prod_txt_stlgrd_DblClick()

    Call prod_txt_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub prod_txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then
            
        DD.nameType = "1"
        DD.sWitch = "MS"
        DD.rControl.Add Item:=prod_txt_stlgrd
        
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
    End If
        
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub


Private Sub ord_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String
    
    If Len(Trim(ord_ord_no.Text)) = ord_ord_no.MaxLength Then
    
        If ord_ord_item.Text <> "" Then Exit Sub
        
        ord_ord_no.Text = StrConv(ord_ord_no.Text, vbUpperCase)
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(ord_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, ord_ord_item, sQuery)
        
       ' If combo_ord_item.ListCount <> 0 Then
       '       combo_ord_item.ListIndex = 0
       ' End If
    Else
        ord_ord_item.Clear
    End If

End Sub

Public Sub Prod_ss_Ref()
   
   On Error GoTo Refer_Err
   
    If Trim(prod_txt_prod_cd.Text) = "" Then Exit Sub

    If Len(Trim(prod_no.Text)) > 10 Then
        Call MsgBox("物料号输入错误" & Chr(10) & "请重新输入", vbExclamation + vbOKOnly, "警告")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
            
    If Len(Trim(prod_loc.Text)) > 10 Then
        Call MsgBox("物料位置输入错误！" & Chr(10) & "请重新输入。", vbExclamation + vbOKOnly, "警告")
        Exit Sub
    End If
    
    If Len(Trim(prod_loc.Text)) > 0 Then
        If Left(Trim(prod_loc.Text), 1) <> Left(prod_txt_prod_cd.Text, 1) Then
            Call MsgBox("产品代码与物料位置不符！" & Chr(10) & "请检查后输入。", vbExclamation + vbOKOnly, "警告")
            Exit Sub
        End If
    End If
    
    If prod_ord_itm.Text <> "" Then
        If Len(prod_ord_itm.Text) = 1 Then
            'S = prod_ord_itm.Text
            prod_ord_itm.Text = "0" + prod_ord_itm.Text
        End If
    End If
    
    If Gf_Sp_ProceExist(Proc_Sc("pSc").Item("Spread")) Then Exit Sub

    If Gf_Sp_Refer(M_CN1, Proc_Sc("pSc"), Mc2, Mc2("nControl"), Mc2("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gp_Ms_ControlLock(Mc2("lControl"), True)
        With MDIMain.MenuTool
            .Buttons(4).Enabled = True                 'Save
            .Buttons(9).Enabled = True                 'Delete
            .Buttons(11).Enabled = False                'Copy
            .Buttons(12).Enabled = False                'Paste
        End With

        Exit Sub
    End If
            
    Exit Sub

Refer_Err:

'    Call ord_sel
'    Call prod_sel

End Sub
