VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGB3020C 
   Caption         =   "母板分段实绩查询与修改界面_AGB3020C"
   ClientHeight    =   9225
   ClientLeft      =   735
   ClientTop       =   2130
   ClientWidth     =   15465
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_org_mplate_no 
      Height          =   315
      Left            =   0
      MaxLength       =   12
      TabIndex        =   82
      Top             =   0
      Visible         =   0   'False
      Width           =   1620
   End
   Begin SSSplitter.SSSplitter SSSplitter2 
      Height          =   8430
      Left            =   60
      TabIndex        =   0
      Top             =   750
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   14870
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "AGB3020C.frx":0000
      Begin Threed.SSPanel SSPanel3 
         Height          =   615
         Left            =   0
         TabIndex        =   35
         Top             =   5400
         Width           =   15345
         _ExtentX        =   27067
         _ExtentY        =   1085
         _Version        =   196609
         BackColor       =   14737918
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txt_res_mplate_no 
            BackColor       =   &H00C0E0FF&
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   8595
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   39
            Top             =   150
            Width           =   1410
         End
         Begin VB.TextBox txt_chg_mplate_no 
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1785
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   36
            Top             =   150
            Width           =   1410
         End
         Begin CSTextLibCtl.sidbEdit sdb_org_mplate_len 
            Height          =   315
            Left            =   1860
            TabIndex        =   37
            Top             =   450
            Visible         =   0   'False
            Width           =   1005
            _Version        =   262145
            _ExtentX        =   1773
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
            Modified        =   -1  'True
            HideSelection   =   -1  'True
            RawData         =   "0.0"
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
            NumDecDigits    =   1
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_chg_mplate_len 
            Height          =   315
            Left            =   4530
            TabIndex        =   38
            Top             =   150
            Width           =   1005
            _Version        =   262145
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   255
            BackColor       =   12648384
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   600
            Top             =   150
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            Caption         =   "母板1"
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
         Begin CSTextLibCtl.sidbEdit sdb_res_mplate_len 
            Height          =   315
            Left            =   11340
            TabIndex        =   40
            Top             =   150
            Width           =   1005
            _Version        =   262145
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   255
            BackColor       =   12640511
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   7410
            Top             =   150
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            Caption         =   "母板2"
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
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   10155
            Top             =   150
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            Caption         =   "母板2长度"
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
         Begin CSTextLibCtl.sidbEdit sdb_res_len 
            Height          =   315
            Left            =   13965
            TabIndex        =   52
            Top             =   150
            Width           =   1215
            _Version        =   262145
            _ExtentX        =   2143
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel26 
            Height          =   315
            Left            =   12780
            Top             =   150
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            Caption         =   "剩余长度"
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
            ForeColor       =   255
         End
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   3345
            Top             =   150
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            Caption         =   "母板1长度"
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
         Begin Threed.SSCheck chk_end 
            Height          =   375
            Left            =   5700
            TabIndex        =   53
            Top             =   120
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   661
            _Version        =   196609
            BackColor       =   14737918
            Caption         =   "变更长度"
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2385
         Left            =   0
         TabIndex        =   4
         Top             =   6045
         Width           =   15345
         _ExtentX        =   27067
         _ExtentY        =   4207
         _Version        =   196609
         BackColor       =   14737632
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin InDate.ULabel ULabel15 
            Height          =   315
            Left            =   7740
            Top             =   480
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "06"
            Alignment       =   1
            BackColor       =   14737632
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
         Begin VB.TextBox txt_org_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   1
            Left            =   570
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   24
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txt_chg_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   1
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   23
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txt_org_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   6
            Left            =   8130
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   22
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txt_chg_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   6
            Left            =   9720
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   21
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txt_chg_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   7
            Left            =   9720
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   20
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txt_org_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   7
            Left            =   8130
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   19
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txt_chg_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   2
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   18
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txt_org_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   2
            Left            =   570
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   17
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txt_chg_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   8
            Left            =   9720
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   16
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txt_org_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   8
            Left            =   8130
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   15
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txt_chg_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   3
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   14
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txt_org_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   3
            Left            =   570
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   13
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txt_chg_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   9
            Left            =   9720
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   12
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox txt_org_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   9
            Left            =   8130
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   11
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox txt_chg_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   4
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   10
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox txt_org_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   4
            Left            =   570
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   9
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox txt_chg_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   10
            Left            =   9720
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   8
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txt_org_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   10
            Left            =   8130
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   7
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txt_chg_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   5
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   6
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txt_org_plate_no 
            Alignment       =   2  'Center
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   5
            Left            =   570
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   5
            Top             =   1920
            Width           =   1575
         End
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   570
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            Caption         =   "原始钢板号"
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
            Left            =   2160
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            Caption         =   "变更钢板号"
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
            ForeColor       =   255
         End
         Begin CSTextLibCtl.sidbEdit sdb_chg_len 
            Height          =   315
            Index           =   1
            Left            =   4700
            TabIndex        =   25
            Top             =   480
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel10 
            Height          =   315
            Left            =   4700
            Top             =   120
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            Caption         =   "变更长度"
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
            ForeColor       =   255
         End
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Left            =   180
            Top             =   480
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "01"
            Alignment       =   1
            BackColor       =   14737632
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
            Left            =   8130
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            Caption         =   "原始钢板号"
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
            Left            =   9720
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            Caption         =   "变更钢板号"
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
            ForeColor       =   255
         End
         Begin CSTextLibCtl.sidbEdit sdb_chg_len 
            Height          =   315
            Index           =   6
            Left            =   12270
            TabIndex        =   26
            Top             =   480
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel14 
            Height          =   315
            Left            =   12270
            Top             =   120
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            Caption         =   "变更长度"
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
            ForeColor       =   255
         End
         Begin CSTextLibCtl.sidbEdit sdb_chg_len 
            Height          =   315
            Index           =   2
            Left            =   4700
            TabIndex        =   27
            Top             =   840
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   180
            Top             =   840
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "02"
            Alignment       =   1
            BackColor       =   14737632
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
         Begin CSTextLibCtl.sidbEdit sdb_chg_len 
            Height          =   315
            Index           =   7
            Left            =   12270
            TabIndex        =   28
            Top             =   840
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel17 
            Height          =   315
            Left            =   7740
            Top             =   840
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "07"
            Alignment       =   1
            BackColor       =   14737632
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
         Begin CSTextLibCtl.sidbEdit sdb_chg_len 
            Height          =   315
            Index           =   3
            Left            =   4700
            TabIndex        =   29
            Top             =   1200
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel18 
            Height          =   315
            Left            =   180
            Top             =   1200
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "03"
            Alignment       =   1
            BackColor       =   14737632
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
         Begin CSTextLibCtl.sidbEdit sdb_chg_len 
            Height          =   315
            Index           =   8
            Left            =   12270
            TabIndex        =   30
            Top             =   1200
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel19 
            Height          =   315
            Left            =   7740
            Top             =   1200
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "08"
            Alignment       =   1
            BackColor       =   14737632
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
         Begin CSTextLibCtl.sidbEdit sdb_chg_len 
            Height          =   315
            Index           =   4
            Left            =   4700
            TabIndex        =   31
            Top             =   1560
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel20 
            Height          =   315
            Left            =   180
            Top             =   1560
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "04"
            Alignment       =   1
            BackColor       =   14737632
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
         Begin CSTextLibCtl.sidbEdit sdb_chg_len 
            Height          =   315
            Index           =   9
            Left            =   12270
            TabIndex        =   32
            Top             =   1560
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel21 
            Height          =   315
            Left            =   7740
            Top             =   1560
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "09"
            Alignment       =   1
            BackColor       =   14737632
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
         Begin CSTextLibCtl.sidbEdit sdb_chg_len 
            Height          =   315
            Index           =   5
            Left            =   4700
            TabIndex        =   33
            Top             =   1920
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Left            =   180
            Top             =   1920
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "05"
            Alignment       =   1
            BackColor       =   14737632
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
         Begin CSTextLibCtl.sidbEdit sdb_chg_len 
            Height          =   315
            Index           =   10
            Left            =   12270
            TabIndex        =   34
            Top             =   1920
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel23 
            Height          =   315
            Left            =   7740
            Top             =   1920
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            Caption         =   "10"
            Alignment       =   1
            BackColor       =   14737632
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
         Begin CSTextLibCtl.sidbEdit sdb_org_len 
            Height          =   315
            Index           =   1
            Left            =   3750
            TabIndex        =   42
            Top             =   480
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel24 
            Height          =   315
            Left            =   3750
            Top             =   120
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            Caption         =   "原始长度"
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
         Begin CSTextLibCtl.sidbEdit sdb_org_len 
            Height          =   315
            Index           =   2
            Left            =   3750
            TabIndex        =   43
            Top             =   840
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_org_len 
            Height          =   315
            Index           =   3
            Left            =   3750
            TabIndex        =   44
            Top             =   1200
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_org_len 
            Height          =   315
            Index           =   4
            Left            =   3750
            TabIndex        =   45
            Top             =   1560
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_org_len 
            Height          =   315
            Index           =   5
            Left            =   3750
            TabIndex        =   46
            Top             =   1920
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_org_len 
            Height          =   315
            Index           =   6
            Left            =   11330
            TabIndex        =   47
            Top             =   480
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel25 
            Height          =   315
            Left            =   11330
            Top             =   120
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            Caption         =   "原始长度"
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
         Begin CSTextLibCtl.sidbEdit sdb_org_len 
            Height          =   315
            Index           =   7
            Left            =   11330
            TabIndex        =   48
            Top             =   840
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_org_len 
            Height          =   315
            Index           =   8
            Left            =   11330
            TabIndex        =   49
            Top             =   1200
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_org_len 
            Height          =   315
            Index           =   9
            Left            =   11330
            TabIndex        =   50
            Top             =   1560
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_org_len 
            Height          =   315
            Index           =   10
            Left            =   11330
            TabIndex        =   51
            Top             =   1920
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_smp_len 
            Height          =   315
            Index           =   1
            Left            =   5640
            TabIndex        =   58
            Top             =   480
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   5640
            Top             =   120
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            Caption         =   "取样长度"
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
         Begin CSTextLibCtl.sidbEdit sdb_smp_len 
            Height          =   315
            Index           =   2
            Left            =   5640
            TabIndex        =   59
            Top             =   840
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_smp_len 
            Height          =   315
            Index           =   3
            Left            =   5640
            TabIndex        =   60
            Top             =   1200
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_smp_len 
            Height          =   315
            Index           =   4
            Left            =   5640
            TabIndex        =   61
            Top             =   1560
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_smp_len 
            Height          =   315
            Index           =   5
            Left            =   5640
            TabIndex        =   62
            Top             =   1920
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_tot_len 
            Height          =   315
            Index           =   1
            Left            =   6600
            TabIndex        =   63
            Top             =   480
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel27 
            Height          =   315
            Left            =   6600
            Top             =   120
            Width           =   915
            _ExtentX        =   1614
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
            ForeColor       =   16711680
         End
         Begin CSTextLibCtl.sidbEdit sdb_tot_len 
            Height          =   315
            Index           =   2
            Left            =   6600
            TabIndex        =   64
            Top             =   840
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_tot_len 
            Height          =   315
            Index           =   3
            Left            =   6600
            TabIndex        =   65
            Top             =   1200
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_tot_len 
            Height          =   315
            Index           =   4
            Left            =   6600
            TabIndex        =   66
            Top             =   1560
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_tot_len 
            Height          =   315
            Index           =   5
            Left            =   6600
            TabIndex        =   67
            Top             =   1920
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_smp_len 
            Height          =   315
            Index           =   6
            Left            =   13230
            TabIndex        =   68
            Top             =   480
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel28 
            Height          =   315
            Left            =   13230
            Top             =   120
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            Caption         =   "取样长度"
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
         Begin CSTextLibCtl.sidbEdit sdb_smp_len 
            Height          =   315
            Index           =   7
            Left            =   13230
            TabIndex        =   69
            Top             =   840
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_smp_len 
            Height          =   315
            Index           =   8
            Left            =   13230
            TabIndex        =   70
            Top             =   1200
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_smp_len 
            Height          =   315
            Index           =   9
            Left            =   13230
            TabIndex        =   71
            Top             =   1560
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_smp_len 
            Height          =   315
            Index           =   10
            Left            =   13230
            TabIndex        =   72
            Top             =   1920
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_tot_len 
            Height          =   315
            Index           =   6
            Left            =   14190
            TabIndex        =   73
            Top             =   480
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel29 
            Height          =   315
            Left            =   14190
            Top             =   120
            Width           =   915
            _ExtentX        =   1614
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
            ForeColor       =   16711680
         End
         Begin CSTextLibCtl.sidbEdit sdb_tot_len 
            Height          =   315
            Index           =   7
            Left            =   14190
            TabIndex        =   74
            Top             =   840
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_tot_len 
            Height          =   315
            Index           =   8
            Left            =   14190
            TabIndex        =   75
            Top             =   1200
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_tot_len 
            Height          =   315
            Index           =   9
            Left            =   14190
            TabIndex        =   76
            Top             =   1560
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_tot_len 
            Height          =   315
            Index           =   10
            Left            =   14190
            TabIndex        =   77
            Top             =   1920
            Width           =   915
            _Version        =   262145
            _ExtentX        =   1614
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
            FocusSelect     =   -1  'True
            Modified        =   -1  'True
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
            NumIntDigits    =   7
            ShowZero        =   0   'False
            MaxValue        =   9999999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   5370
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   15345
         _Version        =   393216
         _ExtentX        =   27067
         _ExtentY        =   9472
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
         MaxCols         =   17
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AGB3020C.frx":0072
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   660
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   1164
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cbo_shift 
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
         ItemData        =   "AGB3020C.frx":0A89
         Left            =   13950
         List            =   "AGB3020C.frx":0A96
         TabIndex        =   81
         Tag             =   "班次"
         Top             =   150
         Width           =   735
      End
      Begin VB.TextBox txt_onoff 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   120
         MaxLength       =   1
         TabIndex        =   78
         Text            =   " "
         Top             =   120
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt_emp_cd 
         Height          =   315
         Left            =   13650
         TabIndex        =   57
         Top             =   510
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txt_prc_line 
         Height          =   315
         Left            =   14850
         MaxLength       =   12
         TabIndex        =   41
         Text            =   "2"
         Top             =   510
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txt_mat_no 
         Height          =   315
         Left            =   5820
         MaxLength       =   12
         TabIndex        =   3
         Tag             =   "母板号"
         Top             =   150
         Width           =   1635
      End
      Begin VB.ComboBox cbo_plt 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "AGB3020C.frx":0AA3
         Left            =   14085
         List            =   "AGB3020C.frx":0AAD
         TabIndex        =   2
         Text            =   "C1"
         Top             =   510
         Visible         =   0   'False
         Width           =   750
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   4170
         Top             =   150
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         Caption         =   "母板号"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   7770
         Top             =   150
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         Caption         =   "生产日期"
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
      Begin InDate.UDate udt_date_fr 
         Height          =   315
         Left            =   9420
         TabIndex        =   55
         Tag             =   "INS_DATE"
         Top             =   150
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
      Begin InDate.UDate udt_date_to 
         Height          =   315
         Left            =   10860
         TabIndex        =   56
         Tag             =   "INS_DATE"
         Top             =   150
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
      Begin Threed.SSOption opt_on 
         Height          =   285
         Left            =   2280
         TabIndex        =   79
         Top             =   180
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
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
         Caption         =   "在线"
         Value           =   -1
      End
      Begin Threed.SSOption opt_off 
         Height          =   285
         Left            =   3150
         TabIndex        =   80
         Top             =   180
         Width           =   705
         _ExtentX        =   1244
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
         Caption         =   "离线"
      End
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Left            =   540
         Top             =   150
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         Caption         =   "在/离线"
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
      Begin InDate.ULabel ULabel31 
         Height          =   315
         Left            =   12600
         Top             =   150
         Width           =   1320
         _ExtentX        =   2328
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
   End
End
Attribute VB_Name = "AGB3020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   Mill System
'-- Program Name      母板分板实绩界面
'-- Program ID        AGB3020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM.SUNG.HO
'-- Coder             KIM.SUNG.HO
'-- Date              2010.7.20
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

Dim Mc1      As New Collection      'Master Collection
Dim Mc2      As New Collection      'Master Collection
Dim sc1      As New Collection      'Spread Collection
Dim Proc_Sc  As New Collection      'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim lMain_Row As Long
Dim Re_chk As Boolean
Dim chg_mplate_len As Long
Dim lDs_Head_Crop As Long
Dim lDs_Tail_Crop As Long
Dim lMplate_Spare_Len As Long
Dim lPlate_Spare_Len As Long


Private Sub Form_Define()

    Dim iCol As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
               Call Gp_Ms_Collection(txt_onoff, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(CBO_PLT, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_PRC_LINE, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(udt_date_fr, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(udt_date_to, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_emp_cd, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_org_mplate_no, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_org_mplate_len, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_chg_mplate_no, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_chg_mplate_len, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_res_mplate_no, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_res_mplate_len, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(chk_end, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_org_plate_no(1), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_chg_plate_no(1), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_chg_len(1), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_org_plate_no(2), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_chg_plate_no(2), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_chg_len(2), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_org_plate_no(3), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_chg_plate_no(3), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_chg_len(3), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_org_plate_no(4), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_chg_plate_no(4), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_chg_len(4), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_org_plate_no(5), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_chg_plate_no(5), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_chg_len(5), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_org_plate_no(6), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_chg_plate_no(6), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_chg_len(6), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_org_plate_no(7), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_chg_plate_no(7), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_chg_len(7), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_org_plate_no(8), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_chg_plate_no(8), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_chg_len(8), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_org_plate_no(9), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_chg_plate_no(9), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_chg_len(9), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_org_plate_no(10), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_chg_plate_no(10), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_chg_len(10), " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_res_len, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      
    Mc1.Add Item:="AGB3020C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="AGB3020C.P_REFER1", Key:="P-R"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_org_mplate_no, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_org_plate_no(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_org_len(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_smp_len(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_tot_len(1), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_org_plate_no(2), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_org_len(2), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_smp_len(2), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_tot_len(2), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_org_plate_no(3), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_org_len(3), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_smp_len(3), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_tot_len(3), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_org_plate_no(4), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_org_len(4), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_smp_len(4), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_tot_len(4), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_org_plate_no(5), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_org_len(5), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_smp_len(5), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_tot_len(5), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_org_plate_no(6), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_org_len(6), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_smp_len(6), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_tot_len(6), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_org_plate_no(7), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_org_len(7), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_smp_len(7), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_tot_len(7), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_org_plate_no(8), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_org_len(8), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_smp_len(8), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_tot_len(8), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_org_plate_no(9), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_org_len(9), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_smp_len(9), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(sdb_tot_len(9), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(txt_org_plate_no(10), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(sdb_org_len(10), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(sdb_smp_len(10), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(sdb_tot_len(10), " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    Mc2.Add Item:="AGB3020C.P_REFER2", Key:="P-R"
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    
    For iCol = 1 To ss1.MaxCols - 1
        Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
    
    Call Gp_Sp_Collection(ss1, ss1.MaxCols, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGB3020C.P_SREFER", Key:="P-R"
    sc1.Add Item:="AGB3020C.P_SMODIFY", Key:="P-M"
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
    
    Call Gp_Sp_ColHidden(ss1, ss1.MaxCols, True)
     
End Sub

Private Sub chk_end_Click(Value As Integer)

    Dim cIndex As Integer

    chg_mplate_len = 0
    
    If chk_end Then
    
        txt_chg_mplate_no.Text = txt_org_mplate_no.Text
        txt_res_mplate_no.Text = ""
        sdb_res_mplate_len.Value = 0
        sdb_res_len.Value = sdb_org_mplate_len.Value - sdb_chg_mplate_len.Value
        Call sdb_chg_mplate_len_LostFocus
             
        For cIndex = 1 To 10
        
            If txt_chg_plate_no(cIndex).Text = "" Then
                
                txt_chg_plate_no(cIndex).BackColor = &H80000005
                sdb_chg_len(cIndex).BackColor = &H80000005
                
                If txt_org_plate_no(cIndex).Text = "" Then
                    sdb_tot_len(cIndex).Value = 0
                End If
                
            End If
            
        Next cIndex
        
    Else
    
        If sdb_chg_mplate_len.Value = 0 Then Exit Sub
        
        If txt_chg_mplate_no.Text <> "" Then
            
            If Mid(txt_org_mplate_no.Text, 11, 1) = "0" Or Mid(txt_org_mplate_no.Text, 11, 1) = "1" Then
            
                If Mid(txt_org_mplate_no.Text, 11, 1) = "0" Then
                
                    txt_chg_mplate_no.Text = Mid(txt_org_mplate_no.Text, 1, 10) & "2" & Mid(txt_org_mplate_no.Text, 12, 1)
                    txt_res_mplate_no.Text = Mid(txt_org_mplate_no.Text, 1, 10) & "3" & Mid(txt_org_mplate_no.Text, 12, 1)
                    
                Else
                
                    txt_chg_mplate_no.Text = Mid(txt_org_mplate_no.Text, 1, 10) & "6" & Mid(txt_org_mplate_no.Text, 12, 1)
                    txt_res_mplate_no.Text = Mid(txt_org_mplate_no.Text, 1, 10) & "7" & Mid(txt_org_mplate_no.Text, 12, 1)
                    
                End If
            
            Else
            
                txt_chg_mplate_no.Text = txt_org_mplate_no.Text
                txt_res_mplate_no.Text = Mid(txt_chg_mplate_no.Text, 1, 10) & Val(Mid(txt_chg_mplate_no.Text, 11, 1)) + 1 & Mid(txt_chg_mplate_no.Text, 12, 1)
            
            End If
                
            Call sdb_chg_mplate_len_KeyUp(0, 0)
            Call sdb_chg_mplate_len_LostFocus
        
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

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet

    Call Gp_Ms_Cls(Mc1("iControl"))
    TXT_MAT_NO.Text = ""
    sdb_org_mplate_len.Value = 0
    sdb_res_len.Value = 0
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_ReadOnlySet(ss1)
    Call Gf_Sp_Cls(sc1)
    
    txt_onoff.Text = "I"
    CBO_PLT.ListIndex = 0
    txt_PRC_LINE.Text = "2"
    txt_emp_cd.Text = sUserID
    udt_date_fr.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
    udt_date_to.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
    lMain_Row = 0
    Re_chk = False
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    
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
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Cls()

    Dim iCnt As Integer
    
    If Gf_Sp_Cls(sc1) Then
        Call Gp_Ms_Cls(Mc1("iControl"))
        TXT_MAT_NO.Text = ""
        sdb_org_mplate_len.Value = 0
        sdb_res_len.Value = 0
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        Call MenuTool_ReSet
        txt_onoff.Text = "I"
        CBO_PLT.ListIndex = 0
        txt_PRC_LINE.Text = "2"
        CBO_SHIFT.Text = ""
        txt_emp_cd.Text = sUserID
        udt_date_fr.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
        udt_date_to.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
        lMain_Row = 0
        Re_chk = False
        
        For iCnt = 1 To 10
            txt_chg_plate_no(iCnt).BackColor = &H80000005
            sdb_chg_len(iCnt).BackColor = &H80000005
        Next
        
    End If

End Sub

Public Sub Form_Ref()
    
    Dim iCnt As Integer
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss1.OperationMode = OperationModeNormal
        Call Gp_Sp_EvenRowBackcolor(ss1)
        lMain_Row = 0
        Re_chk = False
        sdb_org_mplate_len.Value = 0
        sdb_org_mplate_len.Value = 0
        sdb_res_len.Value = 0
        Call Gp_Ms_Cls(Mc1("iControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        CBO_PLT.ListIndex = 0
        txt_PRC_LINE.Text = "2"
        txt_emp_cd.Text = sUserID

        For iCnt = 1 To 10
            txt_chg_plate_no(iCnt).BackColor = &H80000005
            sdb_chg_len(iCnt).BackColor = &H80000005
        Next
    End If
       
End Sub

Public Sub Form_Pro()

    Dim iIdx As Integer
    Dim iCnt As Integer
    
    txt_emp_cd.Text = sUserID
    
    If txt_chg_mplate_no.Text = "" Then
        
        For iCnt = 1 To ss1.MaxRows
            ss1.Row = iCnt
            ss1.Col = 0
            
            If ss1.Text <> "" Then
                ss1.Col = ss1.MaxCols
                ss1.Text = sUserID
            Else
                ss1.Col = ss1.MaxCols
                ss1.Text = ""
            End If
            
        Next
        
        If Gf_Sp_Process(M_CN1, sc1, Mc1) Then
            ss1.OperationMode = OperationModeNormal
            Call Gp_Sp_EvenRowBackcolor(ss1)
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            Call MenuTool_ReSet
            lMain_Row = 0
            Re_chk = False
        End If
        
    Else
    
        If chk_end Then
            
            If sdb_chg_mplate_len.Value > sdb_org_mplate_len.Value Then
                Call Gp_MsgBoxDisplay("第一块母板长度错误..!!", "I")
                Exit Sub
            End If
            
        Else
            
            If sdb_chg_mplate_len.Value = 0 Or sdb_res_mplate_len.Value = 0 Or sdb_res_len.Value < 0 Then
                Call Gp_MsgBoxDisplay("母板长度错误..!!", "I")
                Exit Sub
            End If
            
        End If
        
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
            Call Gf_Sp_Refer(M_CN1, sc1, Mc1)
            ss1.OperationMode = OperationModeNormal
            Call Gp_Sp_EvenRowBackcolor(ss1)
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            Call MenuTool_ReSet
            Call Gp_Ms_Cls(Mc1("iControl"))
            Call Gp_Ms_Cls(Mc2("rControl"))
            TXT_MAT_NO.Text = ""
            txt_emp_cd.Text = sUserID
            sdb_org_mplate_len.Value = 0
            sdb_res_len.Value = 0
            CBO_PLT.ListIndex = 0
            txt_PRC_LINE.Text = "2"
            lMain_Row = 0
            Re_chk = False
            
            For iCnt = 1 To 10
                txt_chg_plate_no(iCnt).BackColor = &H80000005
                sdb_chg_len(iCnt).BackColor = &H80000005
            Next
            
        End If
    
    End If
    
End Sub

Public Sub Form_Ins()

End Sub

Public Sub Spread_Can()
    
    If lMain_Row <> 0 Then Exit Sub
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    Call Gp_Sp_EvenRowBackcolor(ss1)
    
End Sub

Public Sub Spread_Del()

    If lMain_Row <> 0 Then Exit Sub
    Call Gp_Sp_Del(Proc_Sc("SC"))

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

Private Sub opt_off_Click(Value As Integer)

    If opt_off.Value Then
        opt_on.ForeColor = &H80000012
        opt_off.ForeColor = &HFF&
        txt_onoff.Text = "O"
    End If

End Sub

Private Sub opt_on_Click(Value As Integer)

    If opt_on.Value Then
        opt_off.ForeColor = &H80000012
        opt_on.ForeColor = &HFF&
        txt_onoff.Text = "I"
    End If

End Sub

Private Sub sdb_chg_mplate_len_GotFocus()

    chg_mplate_len = sdb_chg_mplate_len.Value
    
End Sub

Private Sub sdb_chg_mplate_len_KeyUp(KeyCode As Integer, Shift As Integer)

    If chk_end Then
        sdb_res_mplate_len.Value = 0
        sdb_res_len.Value = sdb_org_mplate_len.Value - sdb_chg_mplate_len.Value
        Exit Sub
    End If
    
    If sdb_org_mplate_len.Value <> 0 And sdb_chg_mplate_len.Value <> 0 Then
        sdb_res_mplate_len.Value = sdb_org_mplate_len.Value - sdb_chg_mplate_len.Value
        sdb_res_len.Value = sdb_org_mplate_len.Value - (sdb_chg_mplate_len.Value + sdb_res_mplate_len.Value)
    End If
    
End Sub

Private Sub sdb_chg_mplate_len_LostFocus()

    Dim iIdx As Integer
    Dim lLenSum As Long
    Dim lMo1 As Integer
    Dim lMo2 As Integer
    Dim sChg_No As Integer
    Dim Sw As Boolean
    
    Sw = True
    
    If lMain_Row = 0 Then Exit Sub
    If chg_mplate_len = sdb_chg_mplate_len.Value Then Exit Sub
    
    '1st MOPLATE NO Setting
    For iIdx = 1 To 10
    
        If txt_org_plate_no(iIdx).Text <> "" Then
        
            sChg_No = sChg_No + 1
            txt_chg_plate_no(iIdx).Text = txt_chg_mplate_no.Text & Right("0" & Trim(Str(sChg_No)), 2)
            sdb_chg_len(iIdx).Value = sdb_org_len(iIdx).Value
        
        Else
            
            txt_chg_plate_no(iIdx).Text = ""
            sdb_chg_len(iIdx).Value = 0
        
        End If
        
    Next iIdx
    
    sChg_No = 0
        
    For iIdx = 1 To 10
    
        If txt_org_plate_no(iIdx).Text = "" Then Exit For
    
        If Sw Then
        
            If sdb_chg_mplate_len.Value >= lLenSum + sdb_chg_len(iIdx).Value + sdb_smp_len(iIdx).Value + IIf(iIdx = 1, lDs_Head_Crop, lPlate_Spare_Len) Then
            
                sChg_No = sChg_No + 1
                txt_chg_plate_no(iIdx).Text = txt_chg_mplate_no.Text & Right("0" & Trim(Str(sChg_No)), 2)
                lLenSum = lLenSum + sdb_chg_len(iIdx).Value + sdb_smp_len(iIdx).Value + IIf(iIdx = 1, lDs_Head_Crop, lPlate_Spare_Len)
            
            Else
            
                If sdb_chg_mplate_len.Value - lLenSum - sdb_smp_len(iIdx).Value - lPlate_Spare_Len >= 2000 Then
                
                    sChg_No = sChg_No + 1
                    txt_chg_plate_no(iIdx).Text = txt_chg_mplate_no.Text & Right("0" & Trim(Str(sChg_No)), 2)
                    sdb_chg_len(iIdx).Value = sdb_chg_mplate_len.Value - lLenSum - sdb_smp_len(iIdx).Value - lPlate_Spare_Len
                    
                    sChg_No = 0
                    Sw = False
                    
                Else
                
                    sChg_No = 0
                    Sw = False
                
                    If chk_end Then
                        txt_chg_plate_no(iIdx).Text = ""
                        sdb_chg_len(iIdx).Value = 0
                        sdb_tot_len(iIdx).Value = sdb_org_len(iIdx).Value + sdb_tot_len(iIdx - 1).Value + lPlate_Spare_Len
                    Else
                        sChg_No = sChg_No + 1
                        txt_chg_plate_no(iIdx).Text = txt_res_mplate_no.Text & Right("0" & Trim(Str(sChg_No)), 2)
                    End If
                    
                End If
                    
            End If
        
        Else
        
            If chk_end Then
                txt_chg_plate_no(iIdx).Text = ""
                sdb_chg_len(iIdx).Value = 0
                sdb_tot_len(iIdx).Value = sdb_org_len(iIdx).Value + sdb_tot_len(iIdx - 1).Value + lPlate_Spare_Len
            Else
                sChg_No = sChg_No + 1
                txt_chg_plate_no(iIdx).Text = txt_res_mplate_no.Text & Right("0" & Trim(Str(sChg_No)), 2)
            End If
            
        End If
        
    Next iIdx
    
    Sw = True
    lLenSum = 0
    sChg_No = 0
    
    For iIdx = 1 To 10

        If txt_chg_plate_no(iIdx).Text <> "" And Mid(txt_chg_plate_no(iIdx).Text, 1, 12) = txt_res_mplate_no.Text Then

            If Sw Then
                
                If sdb_res_mplate_len.Value >= lLenSum + sdb_chg_len(iIdx).Value + sdb_smp_len(iIdx).Value + IIf(iIdx = 1, lDs_Head_Crop, lPlate_Spare_Len) Then
    
                    sChg_No = sChg_No + 1
                    txt_chg_plate_no(iIdx).Text = txt_res_mplate_no.Text & Right("0" & Trim(Str(sChg_No)), 2)
                    lLenSum = lLenSum + sdb_chg_len(iIdx).Value + sdb_smp_len(iIdx).Value + IIf(iIdx = 1, lDs_Head_Crop, lPlate_Spare_Len)
                    
                Else
    
                    If sdb_res_mplate_len.Value - lLenSum - sdb_smp_len(iIdx).Value - lPlate_Spare_Len >= 2000 Then
                    
                        sChg_No = sChg_No + 1
                        txt_chg_plate_no(iIdx).Text = txt_res_mplate_no.Text & Right("0" & Trim(Str(sChg_No)), 2)
                        sdb_chg_len(iIdx).Value = sdb_res_mplate_len.Value - lLenSum - sdb_smp_len(iIdx).Value - lPlate_Spare_Len
                        
                        sChg_No = 0
                        Sw = False
                        
                    Else
                    
                        sChg_No = 0
                        Sw = False
                    
                        sChg_No = sChg_No + 1
                        txt_chg_plate_no(iIdx).Text = ""
                        sdb_chg_len(iIdx).Value = 0
                    End If
                    
                End If
            
            Else
            
                txt_chg_plate_no(iIdx).Text = ""
                sdb_chg_len(iIdx).Value = 0
            
            End If
            
        End If

    Next iIdx
    
    'Backcolor Setting
    lMo1 = 0
    lMo2 = 0
    For iIdx = 1 To 10
    
        If txt_chg_mplate_no.Text = Mid(txt_chg_plate_no(iIdx).Text, 1, 12) Then
            
            txt_chg_plate_no(iIdx).BackColor = &HC0FFC0
            sdb_chg_len(iIdx).BackColor = &HC0FFC0
            
            If lMo1 = 0 Then
                sdb_tot_len(iIdx).Value = sdb_org_len(iIdx).Value + sdb_smp_len(iIdx).Value + lDs_Head_Crop
                lMo1 = 1
            Else
                sdb_tot_len(iIdx).Value = sdb_tot_len(iIdx - 1).Value + sdb_org_len(iIdx).Value + sdb_smp_len(iIdx).Value + lPlate_Spare_Len
            End If
            
        ElseIf txt_res_mplate_no.Text <> "" And txt_res_mplate_no.Text = Mid(txt_chg_plate_no(iIdx).Text, 1, 12) Then
        
            txt_chg_plate_no(iIdx).BackColor = &HC0E0FF
            sdb_chg_len(iIdx).BackColor = &HC0E0FF
        
            If lMo2 = 0 Then
                sdb_tot_len(iIdx).Value = sdb_org_len(iIdx).Value + sdb_smp_len(iIdx).Value
                lMo2 = 1
            Else
                sdb_tot_len(iIdx).Value = sdb_tot_len(iIdx - 1).Value + sdb_org_len(iIdx).Value + sdb_smp_len(iIdx).Value + lPlate_Spare_Len
            End If
        
        Else
        
            txt_chg_plate_no(iIdx).BackColor = &H80000005
            sdb_chg_len(iIdx).BackColor = &H80000005
        
        End If
    
    Next iIdx
    
    chg_mplate_len = 0
    
End Sub

Private Sub sdb_res_mplate_len_KeyUp(KeyCode As Integer, Shift As Integer)

    sdb_res_len.Value = sdb_org_mplate_len.Value - (sdb_chg_mplate_len.Value + sdb_res_mplate_len.Value)
    
End Sub

Private Sub sdb_res_mplate_len_LostFocus()

    Dim iIdx As Integer
    Dim lLenSum As Long
    Dim sChg_No As Integer
    Dim Sw As Boolean
    
    Sw = True
    
    If lMain_Row = 0 Then Exit Sub
    
    For iIdx = 1 To 10
    
        If Mid(txt_chg_plate_no(iIdx).Text, 1, 12) = txt_res_mplate_no.Text Then
        
            If Sw Then
            
                If sdb_res_mplate_len.Value >= lLenSum + sdb_chg_len(iIdx).Value + sdb_smp_len(iIdx).Value Then
                
                    sChg_No = sChg_No + 1
                    txt_chg_plate_no(iIdx).Text = txt_res_mplate_no.Text & Right("0" & Trim(Str(sChg_No)), 2)
                    lLenSum = lLenSum + sdb_chg_len(iIdx).Value + sdb_smp_len(iIdx).Value + IIf(sChg_No = 1, 0, lPlate_Spare_Len)
                    
                Else
                
                    If sdb_res_mplate_len.Value - lLenSum - sdb_smp_len(iIdx).Value - lPlate_Spare_Len >= 2000 Then
                    
                        sChg_No = sChg_No + 1
                        txt_chg_plate_no(iIdx).Text = txt_res_mplate_no.Text & Right("0" & Trim(Str(sChg_No)), 2)
                        sdb_chg_len(iIdx).Value = sdb_res_mplate_len.Value - lLenSum - sdb_smp_len(iIdx).Value - lPlate_Spare_Len
                        
                        sChg_No = 0
                        Sw = False
                        
                    Else
                    
                        sChg_No = 0
                        Sw = False
                    
                        txt_chg_plate_no(iIdx).Text = ""
                        sdb_chg_len(iIdx).Value = 0
                        txt_chg_plate_no(iIdx).BackColor = &H80000005
                        sdb_chg_len(iIdx).BackColor = &H80000005
                    
                    End If
                    
                End If
                    
            Else
                            
                txt_chg_plate_no(iIdx).Text = ""
                sdb_chg_len(iIdx).Value = 0
                txt_chg_plate_no(iIdx).BackColor = &H80000005
                sdb_chg_len(iIdx).BackColor = &H80000005
                
            End If
            
        Else
        
            If txt_org_plate_no(iIdx).Text <> "" And txt_chg_plate_no(iIdx).Text = "" Then
            
                If sdb_res_mplate_len.Value >= lLenSum + sdb_org_len(iIdx).Value + sdb_smp_len(iIdx).Value Then
                
                    sChg_No = sChg_No + 1
                    txt_chg_plate_no(iIdx).Text = txt_res_mplate_no.Text & Right("0" & Trim(Str(sChg_No)), 2)
                    lLenSum = lLenSum + sdb_chg_len(iIdx).Value + sdb_smp_len(iIdx).Value
                    sdb_chg_len(iIdx).Value = sdb_org_len(iIdx).Value
                    
                    txt_chg_plate_no(iIdx).BackColor = &HC0E0FF
                    sdb_chg_len(iIdx).BackColor = &HC0E0FF
                    
                Else
                
                    If sdb_res_mplate_len.Value - lLenSum - sdb_smp_len(iIdx).Value >= 2000 Then
                    
                        sChg_No = sChg_No + 1
                        txt_chg_plate_no(iIdx).Text = txt_res_mplate_no.Text & Right("0" & Trim(Str(sChg_No)), 2)
                        sdb_chg_len(iIdx).Value = sdb_res_mplate_len.Value - lLenSum - sdb_smp_len(iIdx).Value
                        
                        txt_chg_plate_no(iIdx).BackColor = &HC0E0FF
                        sdb_chg_len(iIdx).BackColor = &HC0E0FF
                        
                    Else
                    
                        txt_chg_plate_no(iIdx).Text = ""
                        sdb_chg_len(iIdx).Value = 0
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
    Next iIdx

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
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
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim iCnt As Integer
    Dim sQuery As String
    Dim sPlate As String
    Dim lThk As Long
    
    If ss1.MaxRows < 1 Or Row <= 0 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 0
    
    If ss1.Text = "" Then
    
        If lMain_Row <> 0 Then
    
            ss1.Row = lMain_Row
            ss1.Col = 0
            ss1.Text = ""
            
            If lMain_Row Mod 2 <> 0 Then
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lMain_Row, lMain_Row, , &HF2F2F2)
            Else
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lMain_Row, lMain_Row, , &HFFFFFF)
            End If
            
        End If
    
        For iCnt = 1 To ss1.MaxRows
        
            ss1.Row = iCnt
            ss1.Col = 0
            If ss1.Text <> "" Then
            
                ss1.Text = ""
                If iCnt Mod 2 <> 0 Then
                    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iCnt, iCnt, , &HF2F2F2)
                Else
                    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iCnt, iCnt, , &HFFFFFF)
                End If
            End If
            
        Next iCnt
        
        ss1.Row = Row
        ss1.Col = 0
        sPlate = TXT_MAT_NO.Text
        
        Call Gp_Ms_Cls(Mc1("iControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        TXT_MAT_NO.Text = sPlate
        sdb_org_mplate_len.Value = 0
        sdb_res_len.Value = 0
        CBO_PLT.ListIndex = 0
        txt_PRC_LINE.Text = "2"
        
        ss1.Text = "选择"
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.Row, ss1.Row, , CYAN)
        lMain_Row = Row
        
        ss1.Col = 1
        txt_org_mplate_no.Text = ss1.Text
        ss1.Col = 2
        lThk = ss1.Value
        ss1.Col = 4
        sdb_org_mplate_len.Value = ss1.Value
        
        If Mid(txt_org_mplate_no.Text, 11, 1) = "0" Or Mid(txt_org_mplate_no.Text, 11, 1) = "1" Then
        
            If Mid(txt_org_mplate_no.Text, 11, 1) = "0" Then
            
                txt_chg_mplate_no.Text = Mid(txt_org_mplate_no.Text, 1, 10) & "2" & Mid(txt_org_mplate_no.Text, 12, 1)
                txt_res_mplate_no.Text = Mid(txt_org_mplate_no.Text, 1, 10) & "3" & Mid(txt_org_mplate_no.Text, 12, 1)
                
            Else
            
                txt_chg_mplate_no.Text = Mid(txt_org_mplate_no.Text, 1, 10) & "6" & Mid(txt_org_mplate_no.Text, 12, 1)
                txt_res_mplate_no.Text = Mid(txt_org_mplate_no.Text, 1, 10) & "7" & Mid(txt_org_mplate_no.Text, 12, 1)
                
            End If
        
        Else
        
            txt_chg_mplate_no.Text = txt_org_mplate_no.Text
            txt_res_mplate_no.Text = Mid(txt_chg_mplate_no.Text, 1, 10) & Val(Mid(txt_chg_mplate_no.Text, 11, 1)) + 1 & Mid(txt_chg_mplate_no.Text, 12, 1)
        
        End If
        
        If Gf_CodeFind(M_CN1, "SELECT PLATE_NO FROM NISCO.GP_PLATE WHERE PLATE_NO = '" & txt_res_mplate_no.Text & "'") <> "" Then
        
            Call Gp_MsgBoxDisplay("该母板不能切割..!!")
            txt_res_mplate_no.Text = ""
            
            For iCnt = 1 To 10
                txt_chg_plate_no(iCnt).BackColor = &H80000005
                sdb_chg_len(iCnt).BackColor = &H80000005
            Next
            
            Exit Sub
            
        End If
        
        'Spare Len
        If Mid(txt_org_mplate_no.Text, 11, 1) = "0" Or Mid(txt_org_mplate_no.Text, 11, 1) = "1" Then
            lDs_Head_Crop = Gf_FloatFind(M_CN1, "SELECT NVL(MAXI,0) FROM NISCO.EP_PLATELEN_M WHERE PLT = 'C1' AND  PRC_LINE = '1' AND PROD_THK_MIN <= " & lThk & " AND PROD_THK_MAX > " & lThk & " AND APLY_ITEM = 'PLATELEN_M007'")
            lDs_Tail_Crop = Gf_FloatFind(M_CN1, "SELECT NVL(MAXI,0) FROM NISCO.EP_PLATELEN_M WHERE PLT = 'C1' AND  PRC_LINE = '1' AND PROD_THK_MIN <= " & lThk & " AND PROD_THK_MAX > " & lThk & " AND APLY_ITEM = 'PLATELEN_M008'")
        Else
            lDs_Head_Crop = 0
            lDs_Tail_Crop = 0
        End If
        
        lMplate_Spare_Len = Gf_FloatFind(M_CN1, "SELECT NVL(MAXI,0) FROM NISCO.EP_PLATELEN_M WHERE PLT = 'C1' AND  PRC_LINE = '1' AND PROD_THK_MIN <= " & lThk & " AND PROD_THK_MAX > " & lThk & " AND APLY_ITEM = 'PLATELEN_M010'")
        lPlate_Spare_Len = Gf_FloatFind(M_CN1, "SELECT NVL(MAXI,0) FROM NISCO.EP_PLATELEN_M WHERE PLT = 'C1' AND  PRC_LINE = '1' AND PROD_THK_MIN <= " & lThk & " AND PROD_THK_MAX > " & lThk & " AND APLY_ITEM = 'PLATELEN_M011'")
        
        sdb_chg_mplate_len.Value = sdb_org_mplate_len.Value
        sdb_res_len.Value = 0
        
        Call Gf_Ms_Outpara(M_CN1, Mc2, False)
        
        For iCnt = 1 To 10
        
            txt_chg_plate_no(iCnt).Text = ""
            sdb_chg_len(iCnt).Value = 0
            
            If txt_org_plate_no(iCnt).Text <> "" Then
            
                txt_chg_plate_no(iCnt).Text = txt_chg_mplate_no.Text & Right("0" & Trim(Str(iCnt)), 2)
                sdb_chg_len(iCnt).Value = sdb_org_len(iCnt).Value
                            
                txt_chg_plate_no(iCnt).BackColor = &HC0FFC0
                sdb_chg_len(iCnt).BackColor = &HC0FFC0
                
                If iCnt = 1 Then
                    sdb_tot_len(iCnt).Value = sdb_org_len(iCnt).Value + sdb_smp_len(iCnt).Value + lDs_Head_Crop
                Else
                    sdb_tot_len(iCnt).Value = sdb_tot_len(iCnt - 1).Value + sdb_org_len(iCnt).Value + sdb_smp_len(iCnt).Value + lPlate_Spare_Len
                End If
                
            Else
            
                txt_chg_plate_no(iCnt).BackColor = &H80000005
                sdb_chg_len(iCnt).BackColor = &H80000005
            
            End If
            
        Next
        
    Else
    
        ss1.Text = ""
        sPlate = TXT_MAT_NO.Text
        Call Gp_Ms_Cls(Mc1("iControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        TXT_MAT_NO.Text = sPlate
        sdb_org_mplate_len.Value = 0
        sdb_res_len.Value = 0
        CBO_PLT.ListIndex = 0
        txt_PRC_LINE.Text = "2"
        
        If Row Mod 2 <> 0 Then
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HF2F2F2)
        Else
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFFFF)
        End If
        
        For iCnt = 1 To 10
            txt_chg_plate_no(iCnt).BackColor = &H80000005
            sdb_chg_len(iCnt).BackColor = &H80000005
        Next
        
        lMain_Row = 0
        
    End If
    
    chk_end.Value = ssCBUnchecked
        
End Sub

Private Sub ss1_LostFocus()

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

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
    
        .Buttons(5).Enabled = False                 'Delete
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = True                  'Row Delete
        .Buttons(9).Enabled = True                  'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = True                 'Excel
            
    End With

End Sub
