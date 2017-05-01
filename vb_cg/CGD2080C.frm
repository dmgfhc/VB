VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form CGD2080C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "标识（标印、标签）打印信息发送界面_CGD2080C"
   ClientHeight    =   7245
   ClientLeft      =   780
   ClientTop       =   2565
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   13830
   WindowState     =   2  'Maximized
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
      ItemData        =   "CGD2080C.frx":0000
      Left            =   14370
      List            =   "CGD2080C.frx":0010
      TabIndex        =   51
      Top             =   90
      Width           =   645
   End
   Begin VB.TextBox txt_rec_sts 
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
      Height          =   285
      Left            =   4170
      MaxLength       =   1
      TabIndex        =   48
      Tag             =   "CD_MANA_NO"
      Text            =   "1"
      Top             =   1380
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox txt_plate_no 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      MaxLength       =   14
      TabIndex        =   7
      Tag             =   "物料号"
      Top             =   480
      Width           =   1965
   End
   Begin VB.TextBox txt_lot_no 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5025
      TabIndex        =   6
      Tag             =   "轧批号"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txt_line 
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
      MaxLength       =   1
      TabIndex        =   5
      Tag             =   "CD_MANA_NO"
      Text            =   "1"
      Top             =   90
      Width           =   480
   End
   Begin VB.TextBox txt_plt 
      CausesValidation=   0   'False
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
      Left            =   1350
      MaxLength       =   2
      TabIndex        =   4
      Tag             =   "生产工厂"
      Top             =   90
      Width           =   420
   End
   Begin VB.TextBox txt_plt_name 
      CausesValidation=   0   'False
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
      Left            =   1770
      TabIndex        =   3
      Tag             =   "机号"
      Top             =   90
      Width           =   1530
   End
   Begin VB.ComboBox CBO_SHIFT 
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
      ItemData        =   "CGD2080C.frx":0024
      Left            =   13725
      List            =   "CGD2080C.frx":0031
      TabIndex        =   2
      Top             =   90
      Width           =   645
   End
   Begin VB.TextBox txt_stdspec_chg 
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
      Left            =   12330
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txt_stdspec 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9420
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   6360
      Top             =   30
   End
   Begin InDate.UDate udt_date_fr 
      Height          =   315
      Left            =   9420
      TabIndex        =   8
      Tag             =   "INS_DATE"
      Top             =   90
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
      TabIndex        =   9
      Tag             =   "INS_DATE"
      Top             =   90
      Width           =   1500
      _ExtentX        =   2646
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
   Begin InDate.ULabel ULabel19 
      Height          =   315
      Left            =   3825
      Top             =   480
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel20 
      Height          =   315
      Left            =   135
      Top             =   480
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "钢板号"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   3825
      Top             =   90
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "精整线"
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
   Begin SSSplitter.SSSplitter SSSp1 
      Height          =   8325
      Left            =   90
      TabIndex        =   10
      Top             =   840
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   14684
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "CGD2080C.frx":0041
      Begin Threed.SSPanel SSPanel1 
         Height          =   3030
         Left            =   0
         TabIndex        =   11
         Tag             =   "172.18.151.145"
         Top             =   0
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   5345
         _Version        =   196609
         BackColor       =   12632319
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox TXT_TO_CUR_INV 
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
            Left            =   13380
            MaxLength       =   2
            TabIndex        =   65
            Top             =   3090
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox TXT_CUST_CD 
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
            Left            =   10575
            MaxLength       =   6
            TabIndex        =   64
            Top             =   3090
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TXT_WGT 
            Alignment       =   1  'Right Justify
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
            Left            =   4200
            MaxLength       =   14
            TabIndex        =   62
            Tag             =   "物料号"
            Top             =   1260
            Width           =   735
         End
         Begin VB.TextBox TXT_VESSEL_NO 
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
            Left            =   10320
            TabIndex        =   61
            Tag             =   "物料号"
            Top             =   2460
            Width           =   4515
         End
         Begin Threed.SSFrame SSFrame5 
            Height          =   315
            Left            =   4740
            TabIndex        =   54
            Top             =   480
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   556
            _Version        =   196609
            BackColor       =   12632319
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00C0C0FF&
               Caption         =   " 喷印"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   330
               TabIndex        =   57
               Top             =   30
               Value           =   1  'Checked
               Width           =   900
            End
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00C0C0FF&
               Caption         =   " 冲印"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   3
               Left            =   1695
               TabIndex        =   56
               Top             =   30
               Value           =   1  'Checked
               Width           =   900
            End
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00C0C0FF&
               Caption         =   " 侧喷"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   3030
               TabIndex        =   55
               Top             =   30
               Value           =   1  'Checked
               Width           =   900
            End
         End
         Begin VB.TextBox TXT_ORD_REMARK 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1950
            MultiLine       =   -1  'True
            TabIndex        =   53
            Tag             =   "物料号"
            Top             =   2370
            Width           =   7005
         End
         Begin VB.TextBox TXT_SPEC_DATE 
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
            Left            =   3720
            MaxLength       =   30
            TabIndex        =   52
            Top             =   1620
            Width           =   1215
         End
         Begin VB.TextBox Winsock 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   150
            TabIndex        =   50
            Tag             =   "轧批号"
            Top             =   3660
            Width           =   14835
         End
         Begin VB.CheckBox chk_Cond 
            BackColor       =   &H00C0C0FF&
            Caption         =   " 标签"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   13920
            TabIndex        =   33
            Top             =   150
            Width           =   900
         End
         Begin VB.TextBox TXT_Paint4 
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
            TabIndex        =   26
            Top             =   1980
            Width           =   2385
         End
         Begin VB.TextBox TXT_Paint3 
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
            TabIndex        =   25
            Top             =   1620
            Width           =   2385
         End
         Begin VB.TextBox TXT_Paint2 
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
            TabIndex        =   24
            Top             =   1260
            Width           =   2385
         End
         Begin VB.TextBox TXT_Paint1 
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
            TabIndex        =   23
            Top             =   900
            Width           =   2385
         End
         Begin VB.TextBox TXT_Bar 
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
            Left            =   10320
            TabIndex        =   22
            Top             =   1950
            Width           =   2175
         End
         Begin VB.TextBox TXT_Edge 
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
            Left            =   10320
            TabIndex        =   21
            Top             =   1590
            Width           =   4515
         End
         Begin VB.TextBox TXT_Punch2 
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
            Left            =   10320
            TabIndex        =   20
            Top             =   1230
            Width           =   4515
         End
         Begin VB.TextBox TXT_Punch1 
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
            Left            =   10320
            TabIndex        =   19
            Top             =   870
            Width           =   4515
         End
         Begin VB.TextBox TXT_SPEC 
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
            Left            =   1950
            TabIndex        =   18
            Top             =   1620
            Width           =   1755
         End
         Begin VB.TextBox TXT_MAT_NO 
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
            Left            =   1950
            MaxLength       =   14
            TabIndex        =   17
            Tag             =   "物料号"
            Top             =   900
            Width           =   1965
         End
         Begin VB.TextBox TXT_WID 
            Alignment       =   1  'Right Justify
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
            Left            =   2580
            MaxLength       =   14
            TabIndex        =   16
            Tag             =   "物料号"
            Top             =   1260
            Width           =   675
         End
         Begin VB.TextBox TXT_LEN 
            Alignment       =   1  'Right Justify
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
            Left            =   3270
            MaxLength       =   14
            TabIndex        =   15
            Tag             =   "物料号"
            Top             =   1260
            Width           =   915
         End
         Begin VB.TextBox TXT_P 
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
            Left            =   1950
            MaxLength       =   14
            TabIndex        =   14
            Tag             =   "物料号"
            Text            =   "2"
            Top             =   1980
            Width           =   585
         End
         Begin VB.TextBox TXT_H 
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
            Left            =   4350
            MaxLength       =   14
            TabIndex        =   13
            Tag             =   "物料号"
            Text            =   "2"
            Top             =   1980
            Width           =   585
         End
         Begin VB.TextBox TXT_THK 
            Alignment       =   1  'Right Justify
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
            Left            =   1950
            MaxLength       =   14
            TabIndex        =   12
            Tag             =   "物料号"
            Top             =   1260
            Width           =   615
         End
         Begin Threed.SSFrame SSFrame3 
            Height          =   315
            Left            =   1740
            TabIndex        =   27
            Top             =   120
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line1 
               Height          =   255
               Left            =   330
               TabIndex        =   28
               Top             =   30
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "# 1"
               Value           =   -1
            End
            Begin Threed.SSOption opt_line2 
               Height          =   255
               Left            =   1320
               TabIndex        =   29
               Top             =   30
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "# 2"
            End
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   315
            Left            =   1740
            TabIndex        =   30
            Top             =   510
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line3 
               Height          =   255
               Left            =   330
               TabIndex        =   31
               Top             =   30
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "计划"
               Value           =   -1
            End
            Begin Threed.SSOption opt_line4 
               Height          =   255
               Left            =   1290
               TabIndex        =   32
               Top             =   30
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "实绩"
            End
         End
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   180
            Top             =   120
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            Caption         =   "剪切线"
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
            ForeColor       =   255
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   180
            Top             =   510
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            Caption         =   "钢板状态"
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
            ForeColor       =   255
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   675
            Left            =   9660
            TabIndex        =   34
            Top             =   120
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   1191
            _Version        =   196609
            BackColor       =   12632319
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00C0C0FF&
               Caption         =   " 侧喷"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   8
               Left            =   150
               TabIndex        =   59
               Top             =   360
               Width           =   900
            End
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00C0C0FF&
               Caption         =   " 标印"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   150
               TabIndex        =   58
               Top             =   60
               Width           =   900
            End
            Begin VB.Label tcpMsg2 
               Height          =   225
               Left            =   1350
               TabIndex        =   60
               Top             =   360
               Width           =   2055
            End
            Begin VB.Shape tcpStatus2 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00000000&
               FillColor       =   &H0000FF00&
               Height          =   225
               Left            =   990
               Shape           =   3  'Circle
               Top             =   360
               Width           =   435
            End
            Begin VB.Label tcpMsg 
               Height          =   225
               Left            =   1350
               TabIndex        =   35
               Top             =   60
               Width           =   2055
            End
            Begin VB.Shape tcpStatus 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00000000&
               FillColor       =   &H0000FF00&
               Height          =   225
               Left            =   990
               Shape           =   3  'Circle
               Top             =   60
               Width           =   435
            End
         End
         Begin Threed.SSFrame SSFrame4 
            Height          =   315
            Left            =   6300
            TabIndex        =   36
            Top             =   120
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line5 
               Height          =   255
               Left            =   330
               TabIndex        =   37
               Top             =   30
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "钢板号"
               Value           =   -1
            End
            Begin Threed.SSOption opt_line6 
               Height          =   255
               Left            =   1470
               TabIndex        =   38
               Top             =   30
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "轧批号"
            End
         End
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   4740
            Top             =   120
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            Caption         =   "标印内容"
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
            ForeColor       =   255
         End
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   180
            Top             =   1620
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            Caption         =   "标准/牌号"
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
         Begin InDate.ULabel ULabel10 
            Height          =   315
            Left            =   180
            Top             =   900
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            Caption         =   "物料号"
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
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Left            =   180
            Top             =   1260
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            Caption         =   "厚*宽*长 / 重"
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
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   180
            Top             =   1980
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            Caption         =   "冲印深度(1,2,4)"
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
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel13 
            Height          =   315
            Left            =   2580
            Top             =   1980
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            Caption         =   "侧喷高度(1,2,4)"
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
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   180
            Top             =   2520
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            Caption         =   "订单备注"
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
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel14 
            Height          =   315
            Left            =   9150
            Top             =   2460
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            Caption         =   "加 喷"
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
            ForeColor       =   0
         End
         Begin Threed.SSPanel SSPpdt 
            Height          =   315
            Left            =   12690
            TabIndex        =   63
            Top             =   1950
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   255
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "当月以前交货订单"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin InDate.ULabel ULabel15 
            Height          =   315
            Left            =   11940
            Top             =   3090
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            Caption         =   "目的库"
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
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   9150
            Top             =   3090
            Visible         =   0   'False
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            Caption         =   "客户"
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
            ForeColor       =   0
         End
         Begin Threed.SSPanel SSP4 
            Height          =   315
            Left            =   13500
            TabIndex        =   66
            Top             =   480
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   255
            BackColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "重点订单"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "喷印 line 4:"
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
            Index           =   5
            Left            =   5280
            TabIndex        =   46
            Top             =   2010
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "喷印 line 3:"
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
            Index           =   4
            Left            =   5280
            TabIndex        =   45
            Top             =   1650
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "喷印 line 2:"
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
            Index           =   3
            Left            =   5280
            TabIndex        =   44
            Top             =   1290
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "喷印 line 1:"
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
            Index           =   2
            Left            =   5280
            TabIndex        =   43
            Top             =   930
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "      侧喷:"
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
            Index           =   11
            Left            =   9150
            TabIndex        =   42
            Top             =   1650
            Width           =   1125
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "    条形码:"
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
            Index           =   10
            Left            =   9150
            TabIndex        =   41
            Top             =   2010
            Width           =   1125
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "冲印 line1:"
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
            Index           =   7
            Left            =   9150
            TabIndex        =   40
            Top             =   900
            Width           =   1125
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "冲印 line2:"
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
            Index           =   8
            Left            =   9150
            TabIndex        =   39
            Top             =   1260
            Width           =   1125
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   5265
         Left            =   0
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   3060
         Width           =   15165
         _Version        =   393216
         _ExtentX        =   26749
         _ExtentY        =   9287
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ButtonDrawMode  =   4
         ColsFrozen      =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   59
         MaxRows         =   10
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CGD2080C.frx":0093
      End
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   0
      Left            =   135
      Top             =   90
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "生产工厂"
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
      Left            =   7470
      Top             =   90
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   556
      Caption         =   "生产日期"
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
   End
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   7470
      Top             =   480
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   556
      Caption         =   "标准号 / 改判"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   12660
      Top             =   90
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "班次/别"
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
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5820
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "172.18.128.101"
      RemotePort      =   2121
   End
   Begin CSTextLibCtl.sitxEdit TXT_CUT_TIME 
      Height          =   315
      Left            =   9630
      TabIndex        =   49
      Tag             =   "出炉时间"
      Top             =   9810
      Visible         =   0   'False
      Width           =   2130
      _Version        =   262145
      _ExtentX        =   3757
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
      Text            =   "____-__-__ __-__-__"
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
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   6810
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   25298
   End
End
Attribute VB_Name = "CGD2080C"
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
'-- Program Name      LABEL PRINTER SEND DATA
'-- Program ID        CGC2080C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM SUNG HO
'-- Coder             KIM SUNG HO
'-- Date              2008.3.24
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

Const SPD_LINE1 = 1
Const SPD_LINE2 = 2
Const SPD_PLATE_NO = 3
Const SPD_LOT_NO = 5
Const SPD_CUT_NO = 6
Const SPD_THK = 7
Const SPD_WID = 8
Const SPD_LEN = 9
Const SPD_WGT = 10
Const SPD_LAST_YN = 11
Const SPD_SIZE_KND = 12
Const SPD_TRIM_FL = 13
Const SPD_APLY_STDSPEC = 14
Const SPD_APLY_STDSPEC_NEW = 15
Const SPD_SURF_GRD = 16
Const SPD_MARK_YN = 17
Const SPD_STAMP_YN = 18
Const SPD_BAR_YN = 19
Const SPD_PROD_DATE = 20
Const SPD_EMP_CD = 21
Const SPD_PAINT = 22
Const SPD_LABEL = 23
Const SPD_LOTCD = 24
Const SPD_STDSPEC_YY = 25
Const SPD_STLGRD = 26
Const SPD_ORD_REMARK = 27
Const SPD_UST = 28
Const SPD_CUR_UST = 29
Const SPD_VESSEL_NO = 30
Const SPD_DEL_TO_DATE = 38
Const SPD_CUST_CD = 39
Const SPD_TO_CUR_INV = 40
Const SPD_CUST_CD_SHORT = 41
Const SPD_URGNT_FL = 42
Const SPD_IMP_CONT = 44
Const SPD_SIDE_MARK = 45
Const SPD_JIT_FLAG = 46

Const SS2_PRODSPECNOA_STD = 51 '多船级社标准一
Const SS2_PRODSPECNOB_STD = 52 '多船级社标准二
Const SS2_PRODSPECNOC_STD = 53 '多船级社标准三
Const SS2_PRODSPECNOA = 54 '多船级社牌号一
Const SS2_PRODSPECNOB = 55 '多船级社牌号二
Const SS2_PRODSPECNOC = 56 '多船级社牌号三
Const SS2_PRODSPECNOA1 = 57 '多船级社牌号一
Const SS2_PRODSPECNOB1 = 58 '多船级社牌号二
Const SS2_PRODSPECNOC1 = 59 '多船级社牌号三

Dim PRODSPECNOA As Integer '牌号一
Dim PRODSPECNOB As Integer '牌号二
Dim PRODSPECNOC As Integer '牌号三
Dim PRODSPECNOA1 As Integer '牌号一
Dim PRODSPECNOB1 As Integer '牌号二
Dim PRODSPECNOC1 As Integer '牌号三

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpvDest As Any, ByRef lpvSrc As Any, ByVal cbLength As Long)
Public Property Get LoByte(ByRef Word As Integer) As Byte
CopyMemory LoByte, ByVal VarPtr(Word), 1
End Property

Public Property Let LoByte(ByRef Word As Integer, ByVal LowByte As Byte)
CopyMemory Word, LowByte, 1
End Property

Public Property Get HiByte(ByRef Word As Integer) As Byte
CopyMemory HiByte, ByVal VarPtr(Word) + 1, 1
End Property

Public Property Let HiByte(ByRef Word As Integer, ByVal HighByte As Byte)
CopyMemory ByVal VarPtr(Word) + 1, HighByte, 1
End Property

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
       
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_plt_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_plate_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(udt_date_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(udt_date_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_line, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_lot_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_rec_sts, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
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
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    '2010.09.13 015725 钢板厚宽长公差要求
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 53, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 58, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 59, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGD2080C.P_REFER", Key:="P-R"
    sc1.Add Item:="CGD2080C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="CGD2080C.P_MODIFY", Key:="P-M"
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
    
'    Call Gp_Sp_ColHidden(ss1, 18, True)
    
End Sub

Private Sub CmdSEND_Click()

End Sub

Private Sub chk_Cond_Click(Index As Integer)

    Dim strState As String
    Dim strState2 As String
    
    If Index = 0 Then
       If chk_Cond(Index) = 1 Then
          Winsock1.Connect
       Else
          Winsock1.Close
          strState = "连接断线"
          tcpStatus.BackColor = &HFF&
          chk_Cond(0).ForeColor = &HFF&
          tcpMsg.Caption = "标印机状态 : " & strState
       End If
    End If
    
    If Index = 8 Then
       If chk_Cond(Index) = 1 Then
          Winsock2.Connect
       Else
          Winsock2.Close
          strState2 = "连接断线"
          tcpStatus2.BackColor = &HFF&
          chk_Cond(Index).ForeColor = &HFF&
          tcpMsg2.Caption = "侧喷机状态 : " & strState2
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
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    Call Gf_Sp_Cls(sc1)
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    
  '  Call Gp_Sp_ColHidden(ss1, SPD_MARK_YN, True)
  '  Call Gp_Sp_ColHidden(ss1, SPD_STAMP_YN, True)
  '  Call Gp_Sp_ColHidden(ss1, SPD_BAR_YN, True)
    Call Gp_Sp_ColHidden(ss1, SPD_PAINT, True)
    Call Gp_Sp_ColHidden(ss1, SPD_LABEL, True)
    Call Gp_Sp_ColHidden(ss1, SPD_LOTCD, True)
    
    If App.Title = "CG" Then
        txt_plt.Text = "C3"
    Else
        txt_plt.Text = "C1"
    End If
    
    Call txt_plt_KeyUp(0, 0)
    
    txt_line.Text = "1"
    txt_rec_sts.Text = "1"
    Call Gp_Sp_ColHidden(ss1, SPD_LINE2, True)
    opt_line1 = True
    opt_line3 = True
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Winsock1.State = 1 Or Winsock1.State = 7 Or Winsock1.State = 9 Then
       Winsock1.Close
    End If
    
    If Winsock2.State = 1 Or Winsock2.State = 7 Or Winsock2.State = 9 Then
       Winsock2.Close
    End If

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
        
        If App.Title = "CG" Then
            txt_plt.Text = "C3"
        Else
            txt_plt.Text = "C1"
        End If
        
        Call txt_plt_KeyUp(0, 0)
        txt_line.Text = "1"
        txt_rec_sts.Text = "1"
        opt_line3 = True
        txt_stdspec_chg = ""
        
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()
    
    Dim iCount       As Integer
    Dim iCol         As Integer
    Dim sCurDate     As String
    Dim sDel_To_Date As String
    Dim sPlateNo     As String
    Dim sUrgnt_Fl    As String
    Dim simpcont     As String
    
'    If Gf_Sp_ProceExist(sc1.Item("Spread")) Then Exit Sub

    sCurDate = Format(Now, "YYYYMM")
            
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss1.OperationMode = OperationModeNormal
    End If
    
    With ss1
        For iCount = 1 To .MaxRows
            .ROW = iCount:            .Col = SPD_PLATE_NO
             sPlateNo = .Text
            If Left(.Text, 12) = Left(sPlateNo, 12) Then
            Else
               .ROW = iCount - 1:           .Col = SPD_LAST_YN
               .Value = 1
            End If
            .ROW = iCount:            .Col = SPD_DEL_TO_DATE
            sDel_To_Date = Mid(.Value, 1, 6)
            If sDel_To_Date < sCurDate Then
                 Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iCount, iCount, &HFF&)
            End If
            '紧急订单绿色显示 add by liqian 2012-08-16
            .ROW = iCount:            .Col = SPD_URGNT_FL
            sUrgnt_Fl = Trim(.Text)
            If sUrgnt_Fl = "Y" Then
                     Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iCount, iCount, &HC000&)
            End If
            
            .ROW = iCount:
            .Col = SPD_IMP_CONT:   simpcont = Trim(.Text)
            If simpcont = "Y" Then
                Call Gp_Sp_BlockColor(ss1, SPD_PLATE_NO, SPD_PLATE_NO, iCount, iCount, SSP4.BackColor)
                Call Gp_Sp_BlockColor(ss1, SPD_IMP_CONT, SPD_IMP_CONT, iCount, iCount, SSP4.BackColor)
            End If
        Next iCount
    End With

End Sub

Public Sub Form_Pro()

    Dim iRow As Integer
    
    Dim sMark_no As String
    Dim sPlate_no As String
    Dim sThk As String
    Dim sWid As String
    Dim sLen As String
    Dim sWgt As String
    Dim sSpec As String
    Dim sStdspec_YY As String
    
    If txt_rec_sts = "1" Then
        If Gf_Sp_Pro(M_CN1, Proc_Sc("SC"), Mc1) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call MenuTool_ReSet
        End If
    End If
    
    For iRow = 1 To ss1.MaxRows
         ss1.Col = 0
         ss1.ROW = iRow
         If ss1.Text = "Update" Or ss1.Text = "Insert" Or ss1.Text = "Delete" Then
            ss1.Col = SPD_PLATE_NO:             sPlate_no = ss1.Text
            If opt_line5 Then
                ss1.Col = SPD_PLATE_NO:         sMark_no = ss1.Text
            Else
                ss1.Col = SPD_LOT_NO:           sMark_no = ss1.Text
            End If
            ss1.Col = SPD_THK:              sThk = Trim(Str(ss1.Text))
            ss1.Col = SPD_WID:              sWid = Trim(Str(ss1.Text))
            ss1.Col = SPD_LEN:              sLen = Trim(Str(ss1.Text))
            ss1.Col = SPD_WGT:              sWgt = Trim(Str(ss1.Text))
            If Mid(sWgt, 1, 1) = "." Then
               sWgt = "0" & sWgt
            End If
            ss1.Col = SPD_STDSPEC_YY:       sStdspec_YY = ss1.Text
            ss1.Col = SPD_STLGRD:           sSpec = ss1.Text
            ss1.Col = 0
            If (chk_Cond(0) Or chk_Cond(8)) And ss1.Text <> "Delete" Then
               Call Cmd_SEND(sMark_no, sThk, sWid, sLen, sWgt, sSpec, sStdspec_YY, sPlate_no)
            End If
            Exit For
         End If
    Next iRow
    
    Call Form_Ref

    iRow = iRow + 10
    If iRow > ss1.MaxRows Then
       iRow = ss1.MaxRows
    End If
    
    Call ss1.SetActiveCell(SPD_LEN, iRow)
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_Pro
'   2.Name         : Spread Data Process
'   3.Input  Value : Conn Connection, Sc Collection, Mc Collection, {RefChek Boolean}
'   4.Return Value : Boolean
'   5.Writer       : 杨猛
'   6.Create Date  : 2010. 12 .09
'   7.Modify Date  :
'   8.Comment      : Spread Data Process
'---------------------------------------------------------------------------------------
Public Function Gf_Sp_Pro(Conn As ADODB.Connection, Sc As Collection, Optional MC As Collection, _
                              Optional RefChek As Boolean = False) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iCount, iProcessCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim dTempFloat As Double
    
    Dim SMESG As String
    Dim sTemp As String
    Dim ProcessChk As String
    Dim DelYN As Boolean
    Dim Msg_Count As Integer
    Dim Msg_Yes As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command

    Gf_Sp_Pro = True
    iProcessCount = 0
    
    'MaxRow = 0 is Exit Function Or iCount = 0
    If Sc.Item("Spread").MaxRows < 1 Or Sc.Item("iColumn").Count = 0 Then
        Gf_Sp_Pro = False
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    Sc.Item("Spread").ReDraw = False
    
    'NeceCheck
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
            
            Case "Input", "Update"
            
                If Not MC Is Nothing Then
                    Call Gp_Sp_Move(iCount, Sc, MC)
                End If
                
                'Maxlength Check
                SMESG = Gf_Sp_NeceCheck2(Sc.Item("Spread"), Sc.Item("mColumn"), iCount, Sc.Item("nColumn"))
                        
                If Trim(SMESG) = "OK" Then
                    
                ElseIf Mid(SMESG, 1, 5) = "FALSE" Then
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    SMESG = Mid(SMESG, 6, Len(SMESG))
                    SMESG = SMESG + "长度不正确"
                    Call Gp_MsgBoxDisplay(SMESG)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Pro = False
                    Exit Function
                Else
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    SMESG = SMESG + "必须输入"
                    Call Gp_MsgBoxDisplay(SMESG)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Pro = False
                    Exit Function
                End If
        
        End Select
    
    Next iCount
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_Sp_Pro = False: Exit Function
    End If
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = Conn
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = Sc.Item("P-M")
    
    Conn.BeginTrans
    
    'Create Parameter (Input) iType + iColumn
    For iCount = 0 To Sc.Item("iColumn").Count
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    Msg_Count = 1
    For iCount = 1 To Sc.Item("Spread").MaxRows
        
        ProcessChk = "NO"
        DelYN = False
        
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
        
            Case "Input"
                adoCmd.Parameters(0).Value = "I"
                ProcessChk = "YES"
                
            Case "Update"
                adoCmd.Parameters(0).Value = "U"
                ProcessChk = "YES"
                
            Case "Delete"
                adoCmd.Parameters(0).Value = "D"
                If Msg_Count = 1 Then
                   DelYN = Gf_MessConfirm("您确定要删除状态为[Delete]的数据吗？", "Q")
                   If DelYN Then Msg_Yes = "yes"
                   Msg_Count = Msg_Count + 1
                End If
                If Msg_Yes = "yes" Then DelYN = True
        End Select
          
        If ProcessChk = "YES" Or DelYN Then
            
            'Parameters Setting
            For iCol = 1 To Sc.Item("iColumn").Count
            
                Sc.Item("Spread").Col = Sc.Item("iColumn").Item(iCol)
                
                Select Case Sc.Item("Spread").CellType
                
                    Case SS_CELL_TYPE_CURRENCY
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempFloat = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Trim(Str(dTempFloat))
                        End If
                        
                    Case SS_CELL_TYPE_NUMBER
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempInt = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Trim(Str(dTempInt))
                        End If
                        
                    Case SS_CELL_TYPE_CHECKBOX
                        If Sc.Item("Spread").Value = "1" Then
                            adoCmd.Parameters(iCol).Value = "1"
                        Else
                            adoCmd.Parameters(iCol).Value = "0"
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = "0"
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(Str(Sc.Item("Spread").Value))
                        End If
                        
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If Trim(Sc.Item("Spread").Value) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(Sc.Item("Spread").Value)
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Mid(Trim(Sc.Item("Spread").Text), 1, 4) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 6, 2) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 9, 2)
                        End If
                       
                    Case Else
                        sTemp = Replace(Sc.Item("Spread").Text, "'", "''")
                        adoCmd.Parameters(iCol).Value = Trim(sTemp)
                        
                End Select
           
            Next iCol
                           
            iProcessCount = iProcessCount + 1
            adoCmd.Execute
            
            'Error Check
            If adoCmd("Error") <> "0" Then
            
                ret_Result_ErrCode = adoCmd("Error")
                ret_Result_ErrMsg = adoCmd("Messg")
        
                sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                
                Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                Screen.MousePointer = vbDefault
                Set adoCmd = Nothing
                
                Conn.RollbackTrans
                Gf_Sp_Pro = False
                Exit Function
        
             End If
        
        End If
        
    Next iCount
    
    Conn.CommitTrans
    
    ' 0 Column Space
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
        
            Case "Input", "Update"
            
                sQuery = Gf_Sp_MakeQuery(Sc.Item("Spread"), Sc.Item("P-O"), "O", Sc.Item("pColumn"), iCount)
                Call Gp_Sp_OneRowDisplay(Conn, sQuery, Sc.Item("Spread"), iCount)
                
            Case "Delete"
                If DelYN Then
                   Call Gp_Sp_SendData(Sc.Item("Spread"), "", 0, iCount)
                   Call Gp_Sp_DeleteRow(Sc.Item("Spread"), iCount)
                   iCount = iCount - 1
                End If
        End Select
        
    Next iCount
    
    Sc.Item("Spread").ReDraw = True
            
    If iProcessCount > 0 Then
        If Not MC Is Nothing Then
            Call Gp_Ms_ControlLock(MC.Item("lControl"), True)
        End If
    Else
        Gf_Sp_Pro = False
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function

SpreadPro_Error:
    
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Gf_Sp_Pro = False
    Call Gp_MsgBoxDisplay("Gf_Sp_Pro Error : " & Error)
    Screen.MousePointer = vbDefault

End Function

Public Sub Form_Ins()
    Dim dThk        As Double
    Dim dWid        As Double
    Dim dLen        As Double
    Dim dWgt        As Double
    Dim lRow        As Long
    Dim sPlateNo    As String
    Dim sLotNo      As String
    Dim sCutNo      As String
    Dim sClipText   As String
    
    Dim sSize_knd   As Integer
    Dim sTrim_fl    As Integer
    Dim sAply_stdspec  As String
    Dim sEmp_cd     As String
    Dim sStdspec_YY As String
    Dim sStdspec As String
    Dim iCount As Integer
    
    sPlateNo = ""
    
    With ss1
        If .MaxRows = 0 Then
           If Len(txt_plate_no.Text) = 12 Then
               Call Gp_Sp_Ins(Proc_Sc("Sc"))
              .ROW = 1
              .Col = SPD_PLATE_NO
              .Text = txt_plate_no.Text & "01"
              .Col = SPD_THK:           .Value = 0
              .Col = SPD_WID:           .Value = 0
              .Col = SPD_LEN:           .Value = 0
              .Col = SPD_APLY_STDSPEC:  .Text = "GB-XXX"
           Else
               Call Gp_MsgBoxDisplay("请正确输入母板号 ！")
           End If
           Exit Sub
        End If
        For iCount = .ActiveRow To .MaxRows
            .ROW = iCount
            .Col = SPD_PLATE_NO
            If Left(.Text, 12) = Left(sPlateNo, 12) Or sPlateNo = "" Then
               sPlateNo = .Text
               lRow = iCount
            Else
               Exit For
            End If
        Next iCount
    End With
    
    sPlateNo = ""
    
    Call ss1.SetActiveCell(1, lRow)
    Call Gp_Sp_Ins(Proc_Sc("Sc"))

    With ss1
        .ReDraw = False
        If lRow > 0 Then
            .ROW = lRow
            .Col = SPD_PLATE_NO:      sPlateNo = .Text
            .Col = SPD_LOT_NO:        sLotNo = .Text
            .Col = SPD_CUT_NO:        sCutNo = .Text
            .Col = SPD_THK:           dThk = Val(.Value) 'Val(.Text & "")
            .Col = SPD_WID:           dWid = Val(.Value) 'Val(.Text & "")
            .Col = SPD_LEN:           dLen = Val(.Value) 'Val(.Text & "")
            .Col = SPD_WGT:           dWgt = Val(.Value) 'Val(.Text & "")
            .Col = SPD_SIZE_KND:      sSize_knd = .Value
            .Col = SPD_TRIM_FL:       sTrim_fl = .Value
            .Col = SPD_APLY_STDSPEC:  sAply_stdspec = .Text
            .Col = SPD_STDSPEC_YY:    sStdspec_YY = .Text
            .Col = SPD_EMP_CD:        sEmp_cd = .Text
            .Col = SPD_STLGRD:        sStdspec = .Text
        Else
            sPlateNo = txt_plate_no.Text & "00"
        End If

        .ROW = lRow + 1
        .Col = SPD_PLATE_NO:      .Text = sPlateNo
        .Col = SPD_LOT_NO:        .Text = sLotNo
        .Col = SPD_CUT_NO:        .Text = sCutNo
        .Col = SPD_THK:           .Value = dThk
        .Col = SPD_WID:           .Value = dWid
        .Col = SPD_LEN:           .Value = dLen
        .Col = SPD_WGT:           .Value = dWgt
        .Col = SPD_SIZE_KND:      .Value = sSize_knd
        .Col = SPD_TRIM_FL:       .Value = sTrim_fl
        .Col = SPD_APLY_STDSPEC:  .Text = sAply_stdspec
        .Col = SPD_EMP_CD:        .Text = sEmp_cd
        .Col = SPD_STDSPEC_YY:    .Text = sStdspec_YY
        .Col = SPD_STLGRD:        .Text = sStdspec
        .Col = 0: .Text = "Input"
        .Col = SPD_PLATE_NO: .Text = Mid(.Text, 1, 12) & Format(Val(Mid(.Text, 13, 2) & "") + 1, "00")
        .Col = SPD_SURF_GRD:      .Value = 1
        .Col = SPD_MARK_YN:       .Value = 1
        .Col = SPD_STAMP_YN:      .Value = 1
        .Col = SPD_BAR_YN:        .Value = 1
'        .Col = SPD_LINE1:         .Value = 1
        .Col = 0:                 .Text = "Input"
        
         Call .SetActiveCell(1, .ROW)
        .ReDraw = True
    End With

End Sub



'Public Sub Spread_ColumnsSort()
'    Spread_ColSort.Show 1
'End Sub

Public Sub Spread_Forzens_Setting()
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
End Sub

Public Sub Spread_Forzens_Cancel()
    Me.ActiveControl.ColsFrozen = 0
End Sub

Public Sub Spread_Del()

End Sub

Public Sub Spread_Can()
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub opt_line1_Click(Value As Integer)
    
    If opt_line1 Then
        opt_line1.ForeColor = &HFF&
        opt_line2.ForeColor = &H80000012
        txt_line = "1"
        If ss1.MaxRows > 0 Then Call Form_Ref
        Call Gp_Sp_ColHidden(ss1, SPD_LINE1, False)
        Call Gp_Sp_ColHidden(ss1, SPD_LINE2, True)
'        Winsock1.RemoteHost = "172.18.43.98" 'Gf_ComnNameFind(M_CN1, "G0034", "01", 1)
'        Winsock1.RemotePort = "2121" 'Gf_ComnNameFind(M_CN1, "G0034", "01", 2)
'        Winsock2.RemoteHost = "172.18.43.98" 'Gf_ComnNameFind(M_CN1, "G0034", "01", 1)
'        Winsock2.RemotePort = "25298" 'Gf_ComnNameFind(M_CN1, "G0034", "01", 2)
        Winsock1.RemoteHost = Gf_ComnNameFind(M_CN1, "G0034", "01", 1)
        Winsock1.RemotePort = Gf_ComnNameFind(M_CN1, "G0034", "01", 2)
        Winsock2.RemoteHost = Gf_ComnNameFind(M_CN1, "G0040", "01", 1)
        Winsock2.RemotePort = Gf_ComnNameFind(M_CN1, "G0040", "01", 2)
    End If
    
End Sub

Private Sub opt_line2_Click(Value As Integer)

    If opt_line2 Then
        opt_line2.ForeColor = &HFF&
        opt_line1.ForeColor = &H80000012
        txt_line = "2"
        If ss1.MaxRows > 0 Then Call Form_Ref
        Call Gp_Sp_ColHidden(ss1, SPD_LINE2, False)
        Call Gp_Sp_ColHidden(ss1, SPD_LINE1, True)
        Winsock1.RemoteHost = Gf_ComnNameFind(M_CN1, "G0034", "02", 1)
        Winsock1.RemotePort = Gf_ComnNameFind(M_CN1, "G0034", "02", 2)
        Winsock2.RemoteHost = Gf_ComnNameFind(M_CN1, "G0040", "02", 1)
        Winsock2.RemotePort = Gf_ComnNameFind(M_CN1, "G0040", "02", 2)
    End If
    
End Sub

Private Sub opt_line3_Click(Value As Integer)
    If opt_line3 Then
        opt_line3.ForeColor = &HFF&
        opt_line4.ForeColor = &H80000012
        txt_rec_sts = "1"
    End If
End Sub

Private Sub opt_line4_Click(Value As Integer)
    If opt_line4 Then
        opt_line4.ForeColor = &HFF&
        opt_line3.ForeColor = &H80000012
        txt_rec_sts = "2"
    End If
End Sub
Private Sub opt_line5_Click(Value As Integer)

    If opt_line5 Then
        opt_line5.ForeColor = &HFF&
        opt_line6.ForeColor = &H80000012
    End If
    
End Sub

Private Sub opt_line6_Click(Value As Integer)

    If opt_line6 Then
        opt_line6.ForeColor = &HFF&
        opt_line5.ForeColor = &H80000012
    End If
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    Dim lRow As Integer
    Dim sCheck1 As String
    Dim sCheck2 As String
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal ROW As Long, ByVal ButtonDown As Integer)

    Dim sCheck1 As Integer
    Dim sCheck2 As Integer
    
    Dim iCol As Long
    Dim iRow As Long
    Dim iMode As Integer
    
    Dim iRowNum As Long
    Dim iRowfr As Long
    Dim iRowto As Long
    
    iCol = Col
    iRow = ROW

    If ROW <= 0 Then Exit Sub
    If Col <> SPD_LINE1 And Col <> SPD_LINE2 Then Exit Sub
    If Not Gf_Sc_Authority(sAuthority, "U") Then Exit Sub
    
    iRowto = iRow - 1
    iRowfr = iRow + 1
    
    If iRowto > 0 Then
        For iRowNum = 1 To iRowto
             
             ss1.Col = 0
             ss1.ROW = iRowNum
             If ss1.Text <> "" Then
                ss1.Text = ""
                ss1.Col = SPD_LINE1
                ss1.Value = 0
                ss1.Col = SPD_LINE2
                ss1.Value = 0
                Exit For
             End If
        Next iRowNum
    End If
    
    If iRowfr <= ss1.MaxRows Then
        For iRowNum = iRowfr To ss1.MaxRows
             
             ss1.Col = 0
             ss1.ROW = iRowNum
             If ss1.Text <> "" Then
                ss1.Text = ""
                ss1.Col = SPD_LINE1
                ss1.Value = 0
                ss1.Col = SPD_LINE2
                ss1.Value = 0
                Exit For
             End If
        Next iRowNum
    End If

    ss1.ROW = iRow

    If Col = SPD_LINE1 And ButtonDown = 1 Then
        ss1.Col = SPD_LINE2
        ss1.Text = 0
    ElseIf Col = SPD_LINE2 And ButtonDown = 1 Then
        ss1.Col = SPD_LINE1
        ss1.Text = 0
    End If

    ss1.Col = 0
    ss1.Text = "Update"

    ss1.Col = SPD_LINE1
    sCheck1 = ss1.Value
    ss1.Col = SPD_LINE2
    sCheck2 = ss1.Value

    If sCheck1 = 0 And sCheck2 = 0 Then
        ss1.Col = 0
        ss1.Text = ""
    End If
    
        ss1.Col = SPD_EMP_CD
        ss1.Text = sUserID
        
        ss1.Col = SPD_LABEL
        If chk_Cond(1) Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        
        ss1.Col = SPD_PAINT
        If chk_Cond(0) Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        
        ss1.Col = SPD_LOTCD
        If opt_line6 Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        
'        ss1.Col = SPD_MARK_YN
'        If ss1.Value Then                'chk_Cond(2) hanchao 20140325
'           ss1.Value = 1
'        Else
'           ss1.Value = 0
'        End If
'        ss1.Col = SPD_STAMP_YN
'        If ss1.Value Then               'chk_Cond(3) hanchao 20140325
'           ss1.Value = 1
'        Else
'           ss1.Value = 0
'        End If
'        ss1.Col = SPD_BAR_YN
'        If ss1.Value Then             'chk_Cond(4) hanchao 20140325
'           ss1.Value = 1
'        Else
'           ss1.Value = 0
'        End If
        
        Call Cmd_SEND_SET(ROW)
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

'  Dim sStdspec As String
'  Dim sStdspec_YY As String

  If ROW <= 0 Then Exit Sub
  
  ss1.ROW = ROW
     
  If Col = SPD_APLY_STDSPEC_NEW Then
     ss1.Col = Col
     If ss1.Text = "" Then
        ss1.Text = txt_stdspec_chg
        If txt_stdspec_chg <> "" Then
'            sStdspec = txt_stdspec_chg
'            sStdspec_YY = "%"
            ss1.Col = SPD_SURF_GRD
            ss1.Value = 0
'            ss1.Col = SPD_STDSPEC_YY:          ss1.Text = Gf_qp_std_headFind(M_CN1, sStdspec, sStdspec_YY, 1)
'            ss1.Col = SPD_STLGRD:              ss1.Text = Gf_qp_std_headFind(M_CN1, sStdspec, sStdspec_YY, 2)
            ss1.Col = SPD_CUR_UST:             ss1.Text = "X"
        End If
     Else
            ss1.Col = SPD_APLY_STDSPEC
'            sStdspec = ss1.Text
'            sStdspec_YY = "%"
            ss1.Col = SPD_APLY_STDSPEC_NEW
            ss1.Text = ""
            ss1.Col = SPD_SURF_GRD
            ss1.Value = 1
'            ss1.Col = SPD_STDSPEC_YY:          ss1.Text = Gf_qp_std_headFind(M_CN1, sStdspec, sStdspec_YY, 1)
'            ss1.Col = SPD_STLGRD:              ss1.Text = Gf_qp_std_headFind(M_CN1, sStdspec, sStdspec_YY, 2)
            ss1.Col = SPD_CUR_UST:             ss1.Text = ""
     End If
  End If

  If Col = SPD_PROD_DATE Then
     TXT_CUT_TIME.RawData = Gf_DTSet(M_CN1, , "X")
     ss1.Col = SPD_PROD_DATE
     ss1.Text = TXT_CUT_TIME.Text
  End If
  
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_ComnNameFind
'   2.Name         : Common Code Name Return
'   3.Input  Value : Conn Connection, Cd_Mana_No String, Code String, nameType String
'   4.Return Value : Variant
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Common Code Name Return
'---------------------------------------------------------------------------------------
Public Function Gf_qp_std_headFind(Conn As ADODB.Connection, sStdspec As String, sStdspec_YY As String, nameType As String) As Variant

On Error GoTo qp_std_headFind_Error

    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_qp_std_headFind = "FAIL": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    Select Case nameType
    
        Case "1"        'Short Name
            sQuery = "SELECT MAX(STDSPEC_ORG_KND) FROM qp_std_head WHERE STDSPEC = '" & sStdspec & "' AND STDSPEC_YY LIKE '" & sStdspec_YY & "' AND NVL(STDSPEC_CHR_CD,'Y') <>'N' "
        Case "2"        'Full Name
            sQuery = "SELECT MAX(STDSPEC_STLGRD)  FROM qp_std_head WHERE STDSPEC = '" & sStdspec & "' AND STDSPEC_YY LIKE '" & sStdspec_YY & "' AND NVL(STDSPEC_CHR_CD,'Y') <>'N'"
        Case Else       'Full Name
            sQuery = "SELECT MAX(STDSPEC_STLGRD)  FROM qp_std_head WHERE STDSPEC = '" & sStdspec & "' AND STDSPEC_YY LIKE '" & sStdspec_YY & "' AND NVL(STDSPEC_CHR_CD,'Y') <>'N'"
            
    End Select
    
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            Gf_qp_std_headFind = IIf(VarType(AdoRs.Fields(0)) = vbNull, "", AdoRs.Fields(0))
        End If
        
    Else
        Gf_qp_std_headFind = ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

qp_std_headFind_Error:

    Set AdoRs = Nothing
    Gf_qp_std_headFind = "FAIL"

End Function

Private Sub Cmd_SEND_SET(ByVal ROW As Long)
    
    Dim Header As String * 2
    Dim Nisco As String
    Dim sFlag As String
    Dim sNull As String
    
    Dim sPlate_no As String
    Dim sThk As String
    Dim sWid As String
    Dim sLen As String
    Dim sSpec As String
    Dim sSpec_ALL As String
    Dim sSpec1 As String
    Dim sSpec2 As String
    Dim sSpec_Str As String
    Dim sStdspec_YY As String
    Dim sNum As Integer
    Dim sNumFL As String
    Dim sUST As String
    Dim sCUST_CD_SHORT As String
        
    ss1.ROW = ROW
    If opt_line5 Then
        ss1.Col = SPD_PLATE_NO:     TXT_MAT_NO = ss1.Text
    Else
        ss1.Col = SPD_LOT_NO:       TXT_MAT_NO = ss1.Text
    End If
    ss1.Col = SPD_THK:              TXT_THK = Trim(Str(ss1.Text))
    ss1.Col = SPD_WID:              TXT_WID = Trim(Str(ss1.Text))
    ss1.Col = SPD_LEN:              TXT_LEN = Trim(Str(ss1.Text))
    ss1.Col = SPD_WGT:              TXT_WGT = Trim(Str(ss1.Text))
    If Mid(TXT_WGT, 1, 1) = "." Then
       TXT_WGT = "0" & TXT_WGT
    End If
    ss1.Col = SPD_STDSPEC_YY:       TXT_SPEC = ss1.Text
    ss1.Col = SPD_STLGRD:           TXT_SPEC_DATE = ss1.Text
    ss1.Col = SPD_ORD_REMARK:       TXT_ORD_REMARK = ss1.Text
    ss1.Col = SPD_VESSEL_NO:        TXT_VESSEL_NO = ss1.Text
    ss1.Col = SPD_APLY_STDSPEC_NEW: sSpec_ALL = ss1.Text
    ss1.Col = SPD_UST:              sUST = ss1.Text
    ss1.Col = SPD_CUST_CD:          TXT_CUST_CD = ss1.Text
    ss1.Col = SPD_TO_CUR_INV:       TXT_TO_CUR_INV = ss1.Text
    ss1.Col = SPD_CUST_CD_SHORT:    sCUST_CD_SHORT = ss1.Text
    If sSpec_ALL = "" Then
       ss1.Col = SPD_APLY_STDSPEC:  sSpec_ALL = ss1.Text
    End If
    
    If sUST = "" Or sUST = "X" Then
       sUST = ""
    End If
    ss1.Col = 0
    
    Nisco = "NG"
    sFlag = "X"
    sNull = " "
    
    sPlate_no = TXT_MAT_NO
    sThk = TXT_THK
    sWid = TXT_WID
    sLen = TXT_LEN
    sSpec = TXT_SPEC_DATE
    sSpec1 = sSpec
    sSpec2 = sSpec
    sStdspec_YY = TXT_SPEC
        
    sNum = InStr(sSpec_ALL, "-")
    If sNum = 0 Then
        sNumFL = "Y"
        sNum = Len(sSpec_ALL)
    End If
    
    sSpec_Str = Mid(sSpec_ALL, 1, (sNum - 1))
    
    sSpec1 = sStdspec_YY
    sSpec2 = sStdspec_YY
    
    Select Case sSpec_Str
    
           Case "BV"

           Case "CCS"

           Case "DNV"
           
           Case "VL"

           Case "GL"

           Case "KR"

           Case "LR"

           Case "NK"
 
           Case "RINA"

           Case "ABS"
           
           Case "RS"
           
           Case Else
                sSpec1 = sSpec & " " & sStdspec_YY
                sSpec2 = sSpec
    End Select

    'TXT_Paint1 = Nisco & sNull & sPlate_no
    TXT_Paint1 = sPlate_no
    TXT_Paint2 = sSpec1
    TXT_Paint3 = sThk & sFlag & sWid & sFlag & sLen
    TXT_Paint4 = sUST & sNull & TXT_VESSEL_NO.Text
    TXT_Paint4 = Trim(TXT_Paint4.Text)
    TXT_Paint4 = sCUST_CD_SHORT & "  " & TXT_Paint4.Text
    TXT_Paint4 = Trim(TXT_Paint4.Text)

    TXT_Punch1 = sSpec2 & sNull & sPlate_no
    TXT_Punch2 = sPlate_no
    
    'TXT_Edge = sPlate_no & sNull & sThk & sFlag & sWid & sFlag & sLen & sNull & sSpec2 & sNull & TXT_VESSEL_NO & sNull & TXT_CUST_CD & TXT_TO_CUR_INV
    TXT_Edge = sPlate_no & sNull & sSpec2 & sNull & sThk & sFlag & sWid & sFlag & sLen & sNull & TXT_VESSEL_NO & sNull & TXT_CUST_CD & TXT_TO_CUR_INV
    TXT_Bar = TXT_MAT_NO
    
End Sub
Private Sub Cmd_SEND(iMark_no As String, iThk As String, iWid As String, iLen As String, iWGT As String, iSpec As String, iStdspec_yy As String, iPlate_no As String)

    Dim SMESG As String

    Dim i As Integer
        
    Dim sMark_no As String * 16
    Dim sPlate_no As String * 16
    Dim sThk As String
    Dim sWid As String
    Dim sLen As String
    Dim sWgt As String
    Dim sSpec As String
    Dim sSpec_ALL As String
    Dim sSpec1 As String
    Dim sSpec2 As String
    Dim sStdspec_YY As String
    Dim sUser As String * 10
    
    Dim Header As String * 2
    Dim Nisco_Logo As String
    Dim Nisco As String
    Dim sFlag As String
    Dim sPaint As Integer
    Dim sPunch As Integer
    Dim sEdge As Integer
    Dim sNull As String
    Dim sNullstr As String
        
    Dim sSpec_Str As String
    Dim sSpec_Logo As String
    Dim sSpec_IRS_Logo As String
    Dim sProd_Date As String
    Dim sGroup As String
    Dim sNum As Integer
    Dim sNumFL As String
        
    Dim PaintStr As String
    Dim PaintStr_CD As Integer
    Dim Paint(3) As String * 48
    
    Dim PunchStr As String
    Dim Punch(1) As String * 32
    
    Dim EdgeStr As String
    Dim EdgeStr_CD As Integer
    Dim Edge As String * 48
    Dim Bar As String * 18
    
    Dim sNisco As String
    
    Dim StrSend(10) As String
    
    Dim sUST As String
      
    Dim sEdgeString As String
    '2012-03-01  modify by liqian 侧标位数扩展50->65
    '2012-07-16  modify by liqian 侧标位数扩展65->90
    Dim sEdgeStr As String * 90
    Dim sVESSEL_NO As String
    Dim sideMark As String
    Dim sCUST_CD As String
    Dim sCUST_CD_SHORT As String
    Dim sTO_CUR_INV As String
    
    Dim sJIT_FLAG As String
    
    sUST = "T"
    
    sMark_no = iMark_no
    sPlate_no = iPlate_no
    sThk = iThk
    sWid = iWid
    sLen = iLen
    sWgt = iWGT
    sSpec = iSpec
    sSpec1 = iSpec
    sSpec2 = iSpec
    sStdspec_YY = iStdspec_yy
    sUser = sUserID
    
    Header = "MD"
    Nisco_Logo = Chr(127)
    Nisco = "NG"
    sFlag = "X"
    sNumFL = "N"
    
'    sPaint = 1
'    sPunch = 1
'    sEdge = 1
    
    sProd_Date = udt_date_fr.RawData
    sGroup = Trim(cbo_group.Text)
    
    If sGroup <> "A" And sGroup <> "B" And sGroup <> "C" And sGroup <> "D" Then
        SMESG = " 班别错误，请确认是否正确输入班别"
        Call Gp_MsgBoxDisplay(SMESG)
        Exit Sub
    End If
    
    
    ss1.Col = SPD_MARK_YN
    sPaint = ss1.Value
    ss1.Col = SPD_STAMP_YN
    sPunch = ss1.Value
    ss1.Col = SPD_BAR_YN
    sEdge = ss1.Value
    
    ss1.Col = SPD_APLY_STDSPEC_NEW: sSpec_ALL = ss1.Text
    If sSpec_ALL = "" Then
       ss1.Col = SPD_APLY_STDSPEC:  sSpec_ALL = ss1.Text
    End If
    
    ss1.Col = SS2_PRODSPECNOA:      PRODSPECNOA = ss1.Value
    ss1.Col = SS2_PRODSPECNOB:      PRODSPECNOB = ss1.Value
    ss1.Col = SS2_PRODSPECNOC:      PRODSPECNOC = ss1.Value
    ss1.Col = SS2_PRODSPECNOA1:     PRODSPECNOA1 = ss1.Value
    ss1.Col = SS2_PRODSPECNOB1:     PRODSPECNOB1 = ss1.Value
    ss1.Col = SS2_PRODSPECNOC1:     PRODSPECNOC1 = ss1.Value
    
    sNum = InStr(sSpec_ALL, "-")
    If sNum = 0 Then
        sNumFL = "Y"
        sNum = Len(sSpec_ALL)
    End If
    sSpec_Str = Mid(sSpec_ALL, 1, (sNum - 1))
    
    sSpec1 = sStdspec_YY & " " & PRODSPECNOA1 & " " & PRODSPECNOB1 & " " & PRODSPECNOC1
    sSpec2 = sStdspec_YY & " " & PRODSPECNOA1 & " " & PRODSPECNOB1 & " " & PRODSPECNOC1
        
    Select Case sSpec_Str
           Case "BV"
                 sSpec_Logo = Chr(44) 'ChrW(174)
           Case "CCS"
                 sSpec_Logo = Chr(33) 'ChrW(151)
           Case "DNV"
                 sSpec_Logo = Chr(39) 'ChrW(155)
           Case "VL"
                 sSpec_Logo = Chr(39) 'ChrW(155)
           Case "GL"
                 sSpec_Logo = Chr(39) 'ChrW(171) 36->39 20160127 LICHAO
           Case "KR"
                 sSpec_Logo = Chr(94) 'ChrW(176)
           Case "LR"
                 sSpec_Logo = Chr(96) 'ChrW(172)
           Case "NK"
                 sSpec_Logo = Chr(95) 'ChrW(224)
           Case "RINA"
                 sSpec_Logo = Chr(63) 'ChrW(166)
           Case "ABS"
                 sSpec_Logo = Chr(34) 'ChrW(225)    'CE 126
           Case "RS"  '俄罗斯船级社
                 sSpec_Logo = Chr(36) '" "-> 36 20161118 LICHAO
           Case "IRS"
                 sSpec_Logo = " "
                 sSpec_IRS_Logo = " "
           Case Else
                sSpec_Logo = " "
                sSpec_IRS_Logo = " "
                sSpec1 = sSpec & " " & sStdspec_YY & " " & PRODSPECNOA1 & " " & PRODSPECNOB1 & " " & PRODSPECNOC1
                sSpec2 = sSpec & " " & PRODSPECNOA1 & " " & PRODSPECNOB1 & " " & PRODSPECNOC1
    End Select
    
    sSpec1 = Trim(sSpec1)
    sSpec2 = Trim(sSpec2)
    
    StrSend(0) = Chr(46)
    StrSend(1) = Chr(46)
    sNull = StrSend(0) & StrSend(1)
    sNullstr = " "
    
    '2012-10-22  modify by liqian 标印第一行加 NISCO
     sNisco = sNullstr & "NISCO"
    
    '有重量标识要求的编辑重量信息
    If iStdspec_yy = "GB 713-2008" Or iStdspec_yy = "GB 3531-2008" Or iStdspec_yy = "GB 19189-2011" Or iStdspec_yy = "GB 713-2014" Or iStdspec_yy = "GB 3531-2014" Then
        sWgt = "  T.W. " & sWgt & " t"
    Else
        sWgt = ""
    End If
    
    ss1.ROW = ss1.ActiveRow
    '编辑探伤信息
    '如果钢板要求探伤，喷印第一行末尾加探伤内容
    ss1.Col = SPD_CUR_UST:    sUST = ss1.Text
    If sUST = "" Then
       ss1.Col = SPD_UST:    sUST = ss1.Text
    End If

'    Paint(0) = sNull & Nisco_Logo & sNullstr & sMark_no & sNisco & sWgt
    Paint(0) = sNull & Nisco_Logo & sNullstr & sMark_no & sWgt & " " & sUST
    Paint(1) = sNull & Nisco_Logo & sSpec_Logo & PRODSPECNOA & PRODSPECNOB & PRODSPECNOC & sSpec1
    Paint(2) = sNull & sNisco & "  " & sThk & sFlag & sWid & sFlag & sLen & sNullstr & sProd_Date & sNullstr & sGroup
    
'    ss1.ROW = ss1.ActiveRow
'    '编辑探伤信息
'    '如果钢板要求探伤，喷印第四行加喷 T
'    ss1.Col = SPD_CUR_UST:    sUST = ss1.Text
'    If sUST = "" Then
'       ss1.Col = SPD_UST:    sUST = ss1.Text
'    End If
'
'    If sUST = "" Or sUST = "X" Then
'       sUST = ""
'    Else
'       sUST = "T"
'    End If
    
    ss1.Col = SPD_JIT_FLAG
    If ss1.Text = "Y" Then
         sJIT_FLAG = "DZ" ' 17-DZ
    Else
         sJIT_FLAG = ""
    End If
    
    ss1.Col = SPD_VESSEL_NO:        sVESSEL_NO = ss1.Text
    ss1.Col = SPD_SIDE_MARK:        sideMark = ss1.Text
    ss1.Col = SPD_CUST_CD:          sCUST_CD = ss1.Text
    ss1.Col = SPD_TO_CUR_INV:       sTO_CUR_INV = ss1.Text
    ss1.Col = SPD_CUST_CD_SHORT:    sCUST_CD_SHORT = ss1.Text
    
    '编辑喷印第四行
    '如果钢板为子公司产品，喷印第四行首位喷子公司简码+（探伤标识）+（用户加喷信息）
    If opt_line5 Then
'            Paint(3) = sCUST_CD_SHORT & "  " & sUST & " " & sVESSEL_NO
             Paint(3) = sCUST_CD_SHORT & " " & sVESSEL_NO
    Else
'            Paint(3) = sPlate_no & " " & sCUST_CD_SHORT & "  " & sUST & " " & sVESSEL_NO
             Paint(3) = sPlate_no & " " & sCUST_CD_SHORT & " " & sVESSEL_NO
    End If
    
    If sJIT_FLAG = "" Then
       Paint(3) = sNull & "        " & Paint(3)
    Else
       Paint(3) = sNull & "        " & sJIT_FLAG & "  " & Paint(3)
    End If

    PaintStr = Paint(0) & Paint(1) & Paint(2) & Paint(3)

    StrSend(2) = Chr(30)
    StrSend(3) = Chr(30)
    sNull = StrSend(2) & StrSend(3)
    
    PaintStr_CD = Val(TXT_P)
    Punch(0) = sNull & sSpec_Logo & sSpec_IRS_Logo & sSpec2 & sNullstr & sMark_no
    Punch(1) = sNull '& sMark_no
    PunchStr = Punch(0) & Punch(1)
    
    StrSend(4) = Chr(46)
    StrSend(5) = Chr(46)
    sNull = StrSend(4) & StrSend(5)
    EdgeStr_CD = Val(TXT_H)
    Edge = sNull & sSpec2 & sNullstr & sMark_no & sThk & sFlag & sWid & sFlag & sLen
    
    StrSend(6) = Chr(16)
    StrSend(7) = Chr(16)
    sNull = StrSend(6) & StrSend(7)
    Bar = sNull & sMark_no
    EdgeStr = Bar & Edge
    
    '编辑侧喷信息
    '钢板号 + 尺寸 + 钢种 + （用户加喷信息） + （客户代码）+（目的库）
    If chk_Cond(8) = 1 And sEdge = 1 Then
    
'            sEdgeStr = sMark_no & " " & sThk & "X" & sWid & "X" & sLen & " " & sSpec2 & "  " & TXT_Paint4.Text & " " & TXT_CUST_CD.Text & " " & TXT_TO_CUR_INV.Text
            sEdgeString = sMark_no
            'sEdgeString = Trim(sEdgeString) & " " & Trim(sThk) & "X" & Trim(sWid) & "X" & Trim(sLen) & " " & Trim(sSpec2)
            sEdgeString = Trim(sEdgeString) & " " & Trim(sSpec2) & " " & Trim(sThk) & "X" & Trim(sWid) & "X" & Trim(sLen)
            sEdgeString = sEdgeString & " " & sideMark
            sEdgeString = Trim(sEdgeString) & " " & sCUST_CD & " " & sTO_CUR_INV
            sEdgeStr = Trim(sEdgeString)
      
            Winsock2.SendData sEdgeStr
        
    End If
    
    If chk_Cond(0) = 1 Then
        
            Winsock1.SendData Header & "  " & Chr(16) & Chr(14) & sMark_no & Chr(10) & Chr(10) & sUser
            Winsock1.SendData HiByte(Val(sWid))
            Winsock1.SendData LoByte(Val(sWid))
            Winsock1.SendData HiByte(sPaint)
            Winsock1.SendData LoByte(sPaint)
            Winsock1.SendData HiByte(sPunch)
            Winsock1.SendData LoByte(sPunch)
            Winsock1.SendData HiByte(sEdge)
            Winsock1.SendData LoByte(sEdge)
        
            Winsock1.SendData PaintStr
            
            Winsock1.SendData HiByte(PaintStr_CD)
            Winsock1.SendData LoByte(PaintStr_CD)
            Winsock1.SendData PunchStr
            
            Winsock1.SendData HiByte(EdgeStr_CD)
            Winsock1.SendData LoByte(EdgeStr_CD)
            Winsock1.SendData EdgeStr
    
    End If
    
'    Winsock = Header & "  " & Header & sMark_no & Header & sUser & HiByte(Val(sWid)) & LoByte(Val(sWid)) & HiByte(sPaint) & LoByte(sPaint) & HiByte(sPunch) & LoByte(sPunch) & HiByte(sEdge) & LoByte(sEdge) & PaintStr & HiByte(PaintStr_CD) & LoByte(PaintStr_CD) & PunchStr & HiByte(EdgeStr_CD) & LoByte(EdgeStr_CD) & EdgeStr
'    Winsock = Nisco_Logo & sSpec_Logo 'Header & sNull & Header & sMark_no & Header & sUser & HiByte(Val(sWid)) & LoByte(Val(sWid)) & HiByte(sPaint) & LoByte(sPaint) & HiByte(sPunch) & LoByte(sPunch) & HiByte(sEdge) & LoByte(sEdge) & PaintStr & HiByte(PaintStr_CD) & LoByte(PaintStr_CD) & PunchStr & HiByte(EdgeStr_CD) & LoByte(EdgeStr_CD) & EdgeStr
    
    
    
End Sub

Private Sub Timer1_Timer()

    'sckClosed            0 缺省的。--关闭 没有的
    'sckOpen              1 打开 --打开的
    'sckListening         2 侦听 --察看有没有请求进入的
    'sckConnectionPending 3 连接挂起
    'sckResolvingHost     4 识别主机
    'sckHostResolved      5 已识别主机
    'sckConnecting        6 正在连接
    'sckConnected         7 已连接
    'sckClosing           8 同级人员正在关闭连接 -说明对方关闭了你连接
    'sckError             9 错误
    
    Dim strState As String
    Dim strState2 As String
    
    If chk_Cond(0) <> 1 And chk_Cond(8) <> 1 Then
       Exit Sub
    Else
    
        If chk_Cond(0) = 1 Then
        
            Select Case Winsock1.State
                Case 0
                    strState = "连接关闭"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
'                    Winsock1.Connect
                Case 1
                    strState = "连接打开"
                Case 2
                    strState = "连接保留"
                Case 3
                    strState = "Close"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
                Case 4
                    strState = "Find Host...."
                Case 5
                    strState = "Finded Host"
                Case 6
                    strState = "正在连接"
                Case 7
                    strState = "连接正常"
                    tcpStatus.BackColor = &HC000&
                    chk_Cond(0).ForeColor = &HC000&
                Case 8
                    strState = "连接断线"
'                    Winsock1.Close
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
                Case 9
                    strState = "连接错误"
'                    Winsock1.Close
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
'        '            Winsock1.Connect
            Case Else
                strState = "StateNum:" & Winsock1.State
                tcpStatus.BackColor = &HFF&
                chk_Cond(0).ForeColor = &HFF&
            End Select

            tcpMsg.Caption = "标印机状态 : " & strState
            
        End If
        
        If chk_Cond(8) = 1 Then

            Select Case Winsock2.State
                Case 0
                    strState2 = "连接关闭"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(8).ForeColor = &HFF&
        '            Winsock2.Close
'                    Winsock2.Connect
                Case 1
                    strState2 = "连接打开"
                Case 2
                    strState2 = "连接保留"
                Case 3
                    strState2 = "Close"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(8).ForeColor = &HFF&
                Case 4
                    strState2 = "Find Host...."
                Case 5
                    strState2 = "找到主机"
                Case 6
                    strState2 = "正在连接"
                Case 7
                    strState2 = "连接正常"
                    tcpStatus2.BackColor = &HC000&
                    chk_Cond(8).ForeColor = &HC000&
                Case 8
                    strState2 = "连接断线"
'                    Winsock2.Close
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(8).ForeColor = &HFF&
                Case 9
                    strState2 = "连接错误"
'                    Winsock2.Close
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(8).ForeColor = &HFF&
'        '            Winsock2.Connect
            Case Else
                strState2 = "StateNum:" & Winsock2.State
                tcpStatus2.BackColor = &HFF&
                chk_Cond(8).ForeColor = &HFF&
            End Select

            tcpMsg2.Caption = "侧喷机状态 : " & strState2

        End If
        
    End If
    
End Sub


Private Sub txt_stdspec_chg_DblClick()
    Call txt_stdspec_chg_KeyUp(vbKeyF4, 0)
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim iCol As Long
    Dim iRow As Long
    Dim iMode As Integer
    
    Dim iRowNum As Long
    Dim iRowfr As Long
    Dim iRowto As Long
    
    iCol = Col
    iRow = ROW
    iMode = Mode

    If ROW <= 0 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") And Col > SPD_LINE2 Then
    
         iRowto = iRow - 1
         iRowfr = iRow + 1
        
        If iRowto > 0 Then
            For iRowNum = 1 To iRowto
                 
                 ss1.Col = 0
                 ss1.ROW = iRowNum
                 If ss1.Text <> "" Then
                    ss1.Text = ""
                    ss1.Col = SPD_LINE1
                    ss1.Value = 0
                    ss1.Col = SPD_LINE2
                    ss1.Value = 0
                    Exit For
                 End If
            Next iRowNum
        End If
        
        If iRowfr < ss1.MaxRows Then
            For iRowNum = iRowfr To ss1.MaxRows
                 
                 ss1.Col = 0
                 ss1.ROW = iRowNum
                 If ss1.Text <> "" Then
                    ss1.Text = ""
                    ss1.Col = SPD_LINE1
                    ss1.Value = 0
                    ss1.Col = SPD_LINE2
                    ss1.Value = 0
                    Exit For
                 End If
            Next iRowNum
        End If
    
        If Col = SPD_THK Or Col = SPD_WID Or Col = SPD_LEN Then
            If Mode = 1 Then
               ss1.Col = iCol
               ss1.ROW = iRow
               ss1.Text = 0
            End If
        End If
    
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), iMode)
        
        ss1.ROW = iRow  'ss1.ActiveRow
        ss1.Col = SPD_EMP_CD
        ss1.Text = sUserID
        
        ss1.Col = SPD_LABEL
        If chk_Cond(1) Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        
        ss1.Col = SPD_PAINT
        If chk_Cond(0) Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        
        ss1.Col = SPD_LOTCD
        If opt_line6 Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If

        ss1.Col = SPD_LINE1
        If opt_line1 Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        ss1.Col = SPD_LINE2
        If opt_line2 Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        
'        ss1.Col = SPD_MARK_YN
'        If chk_Cond(2) Then
'           ss1.Value = 1
'        Else
'           ss1.Value = 0
'        End If
'        ss1.Col = SPD_STAMP_YN
'        If chk_Cond(3) Then
'           ss1.Value = 1
'        Else
'           ss1.Value = 0
'        End If
'        ss1.Col = SPD_BAR_YN
'        If chk_Cond(4) Then
'           ss1.Value = 1
'        Else
'           ss1.Value = 0
'        End If
        
        Call Cmd_SEND_SET(iRow)
        
    End If

End Sub

Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ss1.MaxRows > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
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

        If Len(Trim(txt_plt)) = txt_plt.MaxLength Then
            txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
        Else
            txt_plt_name.Text = ""
        End If
    
    End If

End Sub

Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF4 Then
  
         DD.sWitch = "MS"
         DD.DataDicType = "C"
         DD.rControl.Add Item:=txt_stdspec_chg
        
         Call Pf_Common_DD(M_CN1, KeyCode)
         
         Exit Sub
  End If
End Sub

Private Sub txt_stdspec_DblClick()

    Call txt_STDSPEC_KeyUp(vbKeyF4, 0)
    
End Sub
Private Function Pf_Common_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sSelect = False
        DD.sWhere = ""
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
        DD.sSelect = False
        DD.sWhere = ""
        DD.sKey = ""
        
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "HC"        'Common Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    DD.sQuery = "SELECT CD_SHORT_NAME ""标准代号"", CD_NAME ""标准中文名"" FROM ZP_CD WHERE CD_MANA_NO = 'G0030'"
    
    Call Gf_DD_Display(Conn, DD.sQuery, False)
    
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function


Private Sub txt_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)
        
    End If
    
End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
'        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub


