VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form CGD2082C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "钢板标识信息发送界面_CGD2082C"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   6360
      Top             =   30
   End
   Begin VB.TextBox txt_stdspec 
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
      Left            =   9120
      TabIndex        =   8
      Top             =   480
      Width           =   2945
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
      ItemData        =   "CGD2082C.frx":0000
      Left            =   13635
      List            =   "CGD2082C.frx":000D
      TabIndex        =   7
      Top             =   90
      Width           =   765
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
      TabIndex        =   6
      Tag             =   "机号"
      Top             =   90
      Width           =   1530
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
      TabIndex        =   5
      Tag             =   "生产工厂"
      Top             =   90
      Width           =   420
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
      TabIndex        =   4
      Tag             =   "CD_MANA_NO"
      Text            =   "1"
      Top             =   90
      Width           =   480
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
      TabIndex        =   3
      Tag             =   "轧批号"
      Top             =   480
      Width           =   1935
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
      TabIndex        =   2
      Tag             =   "物料号"
      Top             =   480
      Width           =   1965
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
      TabIndex        =   1
      Tag             =   "CD_MANA_NO"
      Text            =   "1"
      Top             =   1380
      Visible         =   0   'False
      Width           =   390
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
      ItemData        =   "CGD2082C.frx":001D
      Left            =   14400
      List            =   "CGD2082C.frx":002D
      TabIndex        =   0
      Top             =   90
      Width           =   765
   End
   Begin InDate.UDate udt_date_fr 
      Height          =   315
      Left            =   9120
      TabIndex        =   9
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
      Left            =   10560
      TabIndex        =   10
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
         Size            =   9.76
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
         Size            =   9.76
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
         Size            =   9.76
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
      TabIndex        =   11
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
      PaneTree        =   "CGD2082C.frx":0041
      Begin Threed.SSPanel SSPanel1 
         Height          =   3000
         Left            =   0
         TabIndex        =   12
         Tag             =   "172.18.151.145"
         Top             =   0
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   5292
         _Version        =   196609
         BackColor       =   12632319
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox TXT_CE 
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
            Left            =   12360
            TabIndex        =   67
            Tag             =   "物料号"
            Top             =   2640
            Width           =   675
         End
         Begin VB.TextBox TXT_PAINTNUM 
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
            TabIndex        =   66
            Tag             =   "物料号"
            Top             =   2640
            Width           =   675
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
            TabIndex        =   65
            Top             =   3570
            Width           =   1095
         End
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
            TabIndex        =   64
            Top             =   3570
            Width           =   615
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
            TabIndex        =   37
            Tag             =   "物料号"
            Top             =   1260
            Width           =   615
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
            TabIndex        =   36
            Tag             =   "物料号"
            Text            =   "2"
            Top             =   1980
            Width           =   585
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
            TabIndex        =   35
            Tag             =   "物料号"
            Text            =   "2"
            Top             =   1980
            Width           =   585
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
            TabIndex        =   34
            Tag             =   "物料号"
            Top             =   1260
            Width           =   915
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
            TabIndex        =   33
            Tag             =   "物料号"
            Top             =   1260
            Width           =   675
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
            TabIndex        =   32
            Tag             =   "物料号"
            Top             =   900
            Width           =   1965
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
            TabIndex        =   31
            Top             =   1620
            Width           =   1755
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
            TabIndex        =   30
            Top             =   870
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
            TabIndex        =   29
            Top             =   1230
            Width           =   4515
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
            TabIndex        =   28
            Top             =   1590
            Width           =   4515
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
            TabIndex        =   27
            Top             =   1950
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
            TabIndex        =   26
            Top             =   900
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
            TabIndex        =   25
            Top             =   1260
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
            TabIndex        =   24
            Top             =   1620
            Width           =   2385
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
            TabIndex        =   23
            Top             =   1980
            Width           =   2385
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
            TabIndex        =   22
            Top             =   150
            Width           =   900
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
            TabIndex        =   21
            Tag             =   "轧批号"
            Top             =   3180
            Visible         =   0   'False
            Width           =   14835
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
            TabIndex        =   20
            Top             =   1620
            Width           =   1215
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
            Height          =   585
            Left            =   1950
            MultiLine       =   -1  'True
            TabIndex        =   19
            Tag             =   "物料号"
            Top             =   2370
            Width           =   7035
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
            TabIndex        =   14
            Tag             =   "物料号"
            Top             =   2280
            Width           =   4515
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
            TabIndex        =   13
            Tag             =   "物料号"
            Top             =   1260
            Width           =   735
         End
         Begin Threed.SSFrame SSFrame5 
            Height          =   315
            Left            =   4740
            TabIndex        =   15
            Top             =   480
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   556
            _Version        =   196609
            BackColor       =   12632319
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
               TabIndex        =   18
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
               TabIndex        =   17
               Top             =   30
               Value           =   1  'Checked
               Width           =   900
            End
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
               TabIndex        =   16
               Top             =   30
               Value           =   1  'Checked
               Width           =   900
            End
         End
         Begin Threed.SSFrame SSFrame3 
            Height          =   315
            Left            =   1740
            TabIndex        =   38
            Top             =   120
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line1 
               Height          =   255
               Left            =   330
               TabIndex        =   39
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
               TabIndex        =   40
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
            TabIndex        =   41
            Top             =   510
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line3 
               Height          =   255
               Left            =   1320
               TabIndex        =   42
               Top             =   30
               Visible         =   0   'False
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
            End
            Begin Threed.SSOption opt_line4 
               Height          =   255
               Left            =   330
               TabIndex        =   43
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
               Value           =   -1
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
               Size            =   9.76
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
               Size            =   9.76
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
            TabIndex        =   44
            Top             =   120
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   1191
            _Version        =   196609
            BackColor       =   12632319
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
               TabIndex        =   46
               Top             =   60
               Width           =   900
            End
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
               TabIndex        =   45
               Top             =   360
               Width           =   900
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
            Begin VB.Label tcpMsg 
               Height          =   225
               Left            =   1350
               TabIndex        =   48
               Top             =   60
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
            Begin VB.Label tcpMsg2 
               Height          =   225
               Left            =   1350
               TabIndex        =   47
               Top             =   360
               Width           =   2055
            End
         End
         Begin Threed.SSFrame SSFrame4 
            Height          =   315
            Left            =   6300
            TabIndex        =   49
            Top             =   120
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line5 
               Height          =   255
               Left            =   330
               TabIndex        =   50
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
               TabIndex        =   51
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
               Size            =   9.76
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
               Size            =   9.76
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
               Size            =   9.76
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
               Size            =   9.76
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
               Size            =   9.76
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
            Top             =   2490
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
               Size            =   9.76
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
            Top             =   2280
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
               Size            =   9.76
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel15 
            Height          =   315
            Left            =   11940
            Top             =   3570
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
            Top             =   3570
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
         Begin InDate.ULabel ULabel18 
            Height          =   315
            Left            =   9150
            Top             =   2640
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            Caption         =   "标识次数"
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
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel21 
            Height          =   315
            Left            =   11190
            Top             =   2640
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            Caption         =   "是否CE标识"
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
            ForeColor       =   0
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
            TabIndex        =   59
            Top             =   1260
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
            TabIndex        =   58
            Top             =   900
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
            TabIndex        =   57
            Top             =   2010
            Width           =   1125
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
            TabIndex        =   56
            Top             =   1650
            Width           =   1125
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
            TabIndex        =   55
            Top             =   930
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
            TabIndex        =   54
            Top             =   1290
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
            TabIndex        =   53
            Top             =   1650
            Width           =   1275
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
            TabIndex        =   52
            Top             =   2010
            Width           =   1275
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   5295
         Left            =   0
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   3030
         Width           =   15165
         _Version        =   393216
         _ExtentX        =   26749
         _ExtentY        =   9340
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
         MaxCols         =   57
         MaxRows         =   10
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CGD2082C.frx":0093
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
      Left            =   7470
      Top             =   90
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   7470
      Top             =   480
      Width           =   1620
      _ExtentX        =   2858
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   12570
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
         Size            =   9.76
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
      Left            =   16290
      TabIndex        =   61
      Tag             =   "出炉时间"
      Top             =   1110
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
   Begin Threed.SSPanel SSP1 
      Height          =   315
      Left            =   12570
      TabIndex        =   62
      Top             =   480
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   16777088
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "已喷印"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSP2 
      Height          =   315
      Left            =   13890
      TabIndex        =   63
      Top             =   480
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   16761087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "已选择"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
End
Attribute VB_Name = "CGD2082C"
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
'-- Program Name      钢板标识信息发送
'-- Program ID        CGC2082C
'-- Document No       Q-00-0010(Specification)
'-- Designer          杨猛
'-- Coder             杨猛
'-- Date              2011.02.12
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-- 1.01  20110212 杨猛         钢板标识信息发送
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
Const SPD_PLATE_NO = 4
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
Const SPD_SURF_GRD = 15
Const SPD_MARK_YN = 16
Const SPD_STAMP_YN = 17
Const SPD_BAR_YN = 18
Const SPD_PROD_REMARK = 19
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
Const SPD_CUST_CD = 37
Const SPD_TO_CUR_INV = 38
Const SPD_CUST_CD_SHORT = 39
Const SPD_ORD_CNT = 40        '一坯多订单  2011-08-23  by  LiQian
Const SPD_ORD_NO = 41
Const SPD_DEL_TO_DATE = 42
Const SPD_HTM_METH = 43
Const SPD_URGNT_FL = 44       '紧急订单绿色标记  2012-08-16  by  LiQian
Const SPD_SIDE_MARK = 45
Const SPD_SEALMEMO = 46   '加冲钢印 20150119
Const SPD_JIT_FLAG = 47
Const SPD_PAINTNUM = 48
Const SPD_CE = 49

Const SPD_PRODSPECNOA_STD = 50
Const SPD_PRODSPECNOB_STD = 51
Const SPD_PRODSPECNOC_STD = 52
Const SPD_PRODSPECNOA = 53
Const SPD_PRODSPECNOB = 54
Const SPD_PRODSPECNOC = 55

Const SPD_CLASS_CD = 56
Const SPD_CLASS_LVL = 57

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

Public Property Get HLByte(ByRef Word As Long, HL As Long) As Byte
CopyMemory HLByte, ByVal VarPtr(Word) + HL, 1
End Property

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
       
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_plt_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(udt_date_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(udt_date_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_line, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_STDSPEC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_LOT_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    '2012.04.19 015725 钢板厚宽长公差要求
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
   Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '加冲钢印 20150119
   Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '标识次数 add by lichao 20140928
   Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 53, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGD2082C.P_REFER", Key:="P-R"
    sc1.Add Item:="CGD2082C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="CGD2082C.P_MODIFY", Key:="P-M"
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
    
    Call Gp_Sp_ColHidden(ss1, SPD_LAST_YN, True)
    Call Gp_Sp_ColHidden(ss1, SPD_TRIM_FL, True)
    Call Gp_Sp_ColHidden(ss1, SPD_PAINT, True)
    Call Gp_Sp_ColHidden(ss1, SPD_LABEL, True)
    Call Gp_Sp_ColHidden(ss1, SPD_LOTCD, True)
    Call Gp_Sp_ColHidden(ss1, SPD_CE, True)
    
    
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
    opt_line4 = True
    
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
        opt_line4 = True
        
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()
    
    Dim iCount      As Double
    Dim sPlateNo    As String
    Dim sCnt_color  As Variant
    Dim sUrgnt_Fl   As String
    
    Dim inum As Integer
    Dim lRow As Integer
    
    
            
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss1.OperationMode = OperationModeNormal
    End If
    
    With ss1
        For iCount = 1 To .MaxRows
        
             ' 一坯多订单,字体显示红色  2011-08-23  by  LiQian
            ss1.ROW = lRow:       ss1.Col = SPD_ORD_CNT
            If ss1.Text <> "" Then
                If ss1.Text = "2" Then
                   sCnt_color = &HFF&
                Else
                   sCnt_color = vbBlack
                End If
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, sCnt_color)
            End If
        
            .ROW = iCount:       .Col = SPD_MARK_YN
            If .Value = 1 Then
                Call Gp_Sp_BlockColor(ss1, 1, -1, iCount, iCount, , SSP1.BackColor)
            End If
            
            '紧急订单绿色显示 add by liqian 2012-08-16
            .ROW = iCount:            .Col = SPD_URGNT_FL
            sUrgnt_Fl = Trim(.Text)
            If sUrgnt_Fl = "Y" Then
                     Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iCount, iCount, &HC000&)
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
    Dim iCount As Double
    
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
'            If (chk_Cond(0) Or chk_Cond(8)) And ss1.Text <> "Delete" Then
               Call Cmd_SEND(sMark_no, sThk, sWid, sLen, sWgt, sSpec, sStdspec_YY, sPlate_no)
'            End If
            Exit For
         End If
    Next iRow
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
    End If
    
    With ss1
        For iCount = 1 To .MaxRows
            .ROW = iCount:       .Col = SPD_MARK_YN
            If .Value = 1 Then
                Call Gp_Sp_BlockColor(ss1, 1, -1, iCount, iCount, , &HFFFF80)
            End If
        Next iCount
    End With

    iRow = iRow + 10
    If iRow > ss1.MaxRows Then
       iRow = ss1.MaxRows
    End If
    
    Call ss1.SetActiveCell(SPD_LEN, iRow)
    
End Sub

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
           If Len(TXT_PLATE_NO.Text) = 12 Then
               Call Gp_Sp_Ins(Proc_Sc("Sc"))
              .ROW = 1
              .Col = SPD_PLATE_NO
              .Text = TXT_PLATE_NO.Text & "01"
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
            sPlateNo = TXT_PLATE_NO.Text & "00"
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
                Call Gp_Sp_BlockColor(ss1, 1, -1, iRowNum, iRowNum)
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
                Call Gp_Sp_BlockColor(ss1, 1, -1, iRowNum, iRowNum)
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
    Call Gp_Sp_BlockColor(ss1, 1, -1, iRow, iRow, , SSP2.BackColor)

    ss1.Col = SPD_LINE1
    sCheck1 = ss1.Value
    ss1.Col = SPD_LINE2
    sCheck2 = ss1.Value

    If sCheck1 = 0 And sCheck2 = 0 Then
        ss1.Col = 0
        ss1.Text = ""
        Call Gp_Sp_BlockColor(ss1, 1, -1, iRow, iRow)
    End If
        
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
'        If chk_Cond(4) Or chk_Cond(8) Then
'           ss1.Value = 1
'        Else
'           ss1.Value = 0
'        End If
        
        Call Cmd_SEND_SET(ROW)
    
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
    Dim sCurDate     As String
    Dim sOrderNo     As String
    Dim sDel_To_Date As String
    
    Dim sAdd_W       As String
    Dim sAdd_S       As String
    Dim sAdd_T       As String
    Dim sAdd_H       As String
    Dim iPaint_Add   As String
    
    Dim sCUST_CD_SHORT As String
    Dim sideMark As String
    
    sCurDate = Format(Now, "YYYYMM")
        
    ss1.ROW = ROW
    If opt_line5 Then
        ss1.Col = SPD_PLATE_NO:     TXT_MAT_NO = ss1.Text
    Else
        ss1.Col = SPD_LOT_NO:       TXT_MAT_NO = ss1.Text
    End If
    ss1.Col = SPD_THK:              txt_thk = Trim(Str(ss1.Text))
    ss1.Col = SPD_WID:              txt_wid = Trim(Str(ss1.Text))
    ss1.Col = SPD_LEN:              txt_len = Trim(Str(ss1.Text))
    ss1.Col = SPD_WGT:              TXT_WGT = Trim(Str(ss1.Text))
    If Mid(TXT_WGT, 1, 1) = "." Then
       TXT_WGT = "0" & TXT_WGT
    End If
    ss1.Col = SPD_STDSPEC_YY:       txt_spec = ss1.Text
    ss1.Col = SPD_STLGRD:           TXT_SPEC_DATE = ss1.Text
    ss1.Col = SPD_ORD_REMARK:       TXT_ORD_REMARK = ss1.Text
    ss1.Col = SPD_VESSEL_NO:        TXT_VESSEL_NO = ss1.Text
    ss1.Col = SPD_UST:              sAdd_T = ss1.Text
    ss1.Col = SPD_CUST_CD:          TXT_CUST_CD = ss1.Text
    ss1.Col = SPD_TO_CUR_INV:       TXT_TO_CUR_INV = ss1.Text
    ss1.Col = SPD_CUST_CD_SHORT:    sCUST_CD_SHORT = ss1.Text
    ss1.Col = SPD_PAINTNUM:         TXT_PAINTNUM = ss1.Text
    ss1.Col = SPD_SIDE_MARK:        sideMark = ss1.Text
    ss1.Col = SPD_CE:               TXT_CE = ss1.Text
    
    If sSpec_ALL = "" Then
       ss1.Col = SPD_APLY_STDSPEC:  sSpec_ALL = ss1.Text
    End If
    
    ss1.Col = SPD_ORD_NO:   sOrderNo = Mid(ss1.Text, 1, 3)
    If sOrderNo = "OB5" Then
        sAdd_W = "W"
    End If
                                    
    ss1.Col = SPD_DEL_TO_DATE:   sDel_To_Date = Mid(ss1.Value, 1, 6)
    If sDel_To_Date < sCurDate Then
        sAdd_S = "S"
    End If
    
    '编辑探伤信息
    '如果钢板要求探伤，喷印第四行加喷 UT + 探伤标准
    ss1.Col = SPD_CUR_UST:    sAdd_T = ss1.Text
    If sAdd_T = "" Then
       ss1.Col = SPD_UST:    sAdd_T = ss1.Text
    End If
    
    If sAdd_T = "" Or sAdd_T = "X" Then
       sAdd_T = ""
    End If
    
    ss1.Col = SPD_HTM_METH:      sAdd_H = ss1.Text
    
    iPaint_Add = sAdd_W & sAdd_S & sAdd_T & sAdd_H
    
    
    ss1.Col = 0
    
    Nisco = "NG"
    sFlag = "X"
    sNull = " "
    
    sPlate_no = TXT_MAT_NO
    sThk = txt_thk
    sWid = txt_wid
    sLen = txt_len
    sSpec = TXT_SPEC_DATE
    sSpec1 = sSpec
    sSpec2 = sSpec
    sStdspec_YY = txt_spec
        
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
    TXT_Paint4 = iPaint_Add & sNull & TXT_VESSEL_NO.Text
    TXT_Paint4 = Trim(TXT_Paint4.Text)
    TXT_Paint4 = sCUST_CD_SHORT & "  " & TXT_Paint4.Text
    TXT_Paint4 = Trim(TXT_Paint4.Text)

    TXT_Punch1 = sSpec2 & sNull & sPlate_no
    TXT_Punch2 = sPlate_no
    
    'TXT_Edge = sPlate_no & sNull & sThk & sFlag & sWid & sFlag & sLen & sNull & sSpec2 & sNull & TXT_VESSEL_NO & sNull & TXT_CUST_CD & TXT_TO_CUR_INV
    TXT_Edge = sPlate_no & sNull & sSpec2 & sNull & sThk & sFlag & sWid & sFlag & sLen & sNull & sideMark & sNull & TXT_CUST_CD & TXT_TO_CUR_INV
    TXT_Bar = TXT_MAT_NO
    
End Sub

Private Sub Cmd_SEND(iMark_no As String, iThk As String, iWid As String, iLen As String, iWGT As String, iSpec As String, iStdspec_yy As String, iPlate_no As String)

    Dim SMESG As String

    Dim i As Integer
        
    Dim sMark_no As String * 16  '18---16
    Dim sPlate_no As String * 16 '18---16
    Dim sThk As String
    Dim sWid As String
    Dim sLen As String
    Dim sWgt As String
    Dim sSpec As String
    Dim sSpec_ALL As String
    Dim sSpec1 As String
    Dim sSpec2 As String
    Dim sStdspec_YY As String
    Dim sUser As String * 6  '10---6
    
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
    Dim sPaint_Logo1 As String
    Dim sPaint_Logo2 As String
    Dim sPaint_Logo3 As String
    Dim sPaint_Logo4 As String
    Dim sPunch_Logo1 As String
    Dim sPunch_Logo2 As String
    Dim sPunch_Logo3 As String
    Dim sPunch_Logo4 As String
    Dim sSpec_IRS_Logo As String
    Dim sProd_Date As String
    Dim sGroup As String
    Dim sNum As Integer
    Dim sNumFL As String
        
    Dim PaintStr As String
    Dim PaintStr_CD As Integer
    Dim Paint(3) As String * 78 '48---78
    
    Dim PunchStr As String
    Dim Punch(1) As String * 32
    
    Dim EdgeStr As String
    Dim EdgeStr_CD As Integer
    Dim Edge As String * 48
    Dim Bar As String * 18
    
    Dim StrSend(10) As String
    
    Dim sCurDate     As String
    Dim sOrderNo     As String
    Dim sDel_To_Date As String
    
    Dim sAdd_W       As String
    Dim sAdd_S       As String
    Dim sAdd_T       As String
    Dim sAdd_H       As String
    Dim iPaint_Add   As String
    
    Dim sNisco As String
    
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
    Dim sEALMEMO As String
    
    Dim sPAINTNUM As String
    
    Dim sSce_Str As String
    Dim sSce_Logo As String
    Dim sStr_Len As Integer
    
    Dim sSinspunita As String
    Dim sSinspunitb As String
    Dim sSinspunitc As String
    
    
    Dim PRODSPECNOA As Integer
    Dim PRODSPECNOB As Integer
    Dim PRODSPECNOC As Integer
    
    Dim sSpec_Punch_Logo As String
    
    Dim sClasscd As String
    Dim sClasslvl As String
    
    Dim sClass As String
    
    sCurDate = Format(Now, "YYYYMM")
    sAdd_T = "T"
    
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
    Nisco_Logo = Chr(1)
    Nisco = "NG"
    sFlag = "X"
    sNumFL = "N"
    
'    sPaint = 1
'    sPunch = 1
'    sEdge = 1
    
    sProd_Date = udt_date_fr.RawData
    sGroup = Trim(CBO_GROUP.Text)
    
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
    
    ss1.Col = SPD_PAINTNUM
    sPAINTNUM = ss1.Text
    If sPAINTNUM = "" Or sPAINTNUM = "0" Then
       sPAINTNUM = "1"
    End If
    
    ss1.Col = SPD_APLY_STDSPEC:  sSpec_ALL = ss1.Text
    
    ss1.Col = SPD_PRODSPECNOA_STD:      PRODSPECNOA = ss1.Value
    ss1.Col = SPD_PRODSPECNOB_STD:      PRODSPECNOB = ss1.Value
    ss1.Col = SPD_PRODSPECNOC_STD:      PRODSPECNOC = ss1.Value
    
    ss1.Col = SPD_PRODSPECNOA:          sSinspunita = ss1.Text
    ss1.Col = SPD_PRODSPECNOB:          sSinspunitb = ss1.Text
    ss1.Col = SPD_PRODSPECNOC:          sSinspunitc = ss1.Text
    
    ss1.Col = SPD_CLASS_CD:             sClasscd = ss1.Text
    ss1.Col = SPD_CLASS_LVL:            sClasslvl = ss1.Text
    
    If sClasscd = "Y" Then
       sClass = sClasslvl
    Else
       sClass = ""
    End If
    
    sNum = InStr(sSpec_ALL, "-")
    If sNum = 0 Then
        sNumFL = "Y"
        sNum = Len(sSpec_ALL)
    End If
    sSpec_Str = Mid(sSpec_ALL, 1, (sNum - 1))
    
    sSpec1 = sStdspec_YY
    sSpec2 = sStdspec_YY
    
    sSce_Str = TXT_CE.Text
    
    '喷印logo、冲印logo初始化。喷印logo1为双锤给1，logo2为船徽，logo3、4给0；冲印logo1为船徽，logo2、3、4给0
    sPaint_Logo1 = Chr(1)
    sPaint_Logo2 = Chr(0)
    sPaint_Logo3 = Chr(0)
    sPaint_Logo4 = Chr(0)
    sPunch_Logo1 = Chr(0)
    sPunch_Logo2 = Chr(0)
    sPunch_Logo3 = Chr(0)
    sPunch_Logo4 = Chr(0)
        
    Select Case sSpec_Str
           Case "BV"
                 sPaint_Logo2 = Chr(134 - 126) '44--134
                 sPunch_Logo1 = Chr(134 - 126) '44--134
                 sSpec_IRS_Logo = "  "
           Case "CCS"
                 sPaint_Logo2 = Chr(140 - 126) '33--140
                 sPunch_Logo1 = Chr(140 - 126) '33--140
                 sSpec_IRS_Logo = "  "
           Case "DNV"
                 sPaint_Logo2 = Chr(159 - 126) '39--159
                 sPunch_Logo1 = Chr(159 - 126) '39--159
                 sSpec_IRS_Logo = "  "
           Case "VL"
                 sPaint_Logo2 = Chr(159 - 126) '39--159
                 sPunch_Logo1 = Chr(159 - 126) '39--159
                 sSpec_IRS_Logo = "  "
           Case "GL"
                 sPaint_Logo2 = Chr(159 - 126) '39--159
                 sPunch_Logo1 = Chr(159 - 126) '39--159
                 sSpec_IRS_Logo = "  "
           Case "KR"
                 sPaint_Logo2 = Chr(146 - 126) '94--146
                 sPunch_Logo1 = Chr(146 - 126) '94--146
                 sSpec_IRS_Logo = "  "
           Case "LR"
                 sPaint_Logo2 = Chr(145 - 126) '96--145
                 sPunch_Logo1 = Chr(145 - 126) '96--145
                 sSpec_IRS_Logo = "  "
           Case "NK"
                 sPaint_Logo2 = Chr(133 - 126) '95--133
                 sPunch_Logo1 = Chr(133 - 126) '95--133
                 sSpec_IRS_Logo = "  "
           Case "RINA"
                 sPaint_Logo2 = Chr(161 - 126) '63--161
                 sPunch_Logo1 = Chr(161 - 126) '63--161
                 sSpec_IRS_Logo = "  "
           Case "ABS"
                 sPaint_Logo2 = Chr(158 - 126) '34--158
                 sPunch_Logo1 = Chr(158 - 126) '34--158
                 sSpec_IRS_Logo = "  "
           Case "RS"  '俄罗斯船级社
                 sPaint_Logo2 = Chr(142 - 126) '36--142
                 sPunch_Logo1 = Chr(142 - 126) '36--142
                 sSpec_IRS_Logo = "  "
           Case "IRS"
                 sPaint_Logo2 = Chr(148 - 126) '""--148
                 sPunch_Logo1 = Chr(148 - 126) '""--148
                 sSpec_IRS_Logo = "  "
           Case Else
                sSpec_IRS_Logo = ""
                sSpec1 = sSpec & " " & sStdspec_YY
                sSpec2 = sSpec
    End Select
    

    sPaint_Logo3 = Chr(PRODSPECNOA)
    sPaint_Logo4 = Chr(PRODSPECNOB)
    
    sPunch_Logo2 = Chr(PRODSPECNOA)
    sPunch_Logo3 = Chr(PRODSPECNOB)
    
    If sSinspunita <> "" Then
       sSpec_Punch_Logo = "  "
    Else
       sSpec_Punch_Logo = ""
    End If
    
    
    If sSce_Str = "是" Then
       sPaint_Logo2 = Chr(137 - 126)
    Else
       sSce_Logo = ""
    End If
    
    StrSend(0) = Chr(75)  '46-->75
    StrSend(1) = Chr(75)  '46-->75
    sNull = StrSend(0) & StrSend(1)
    sNullstr = " "
    
    '2012-10-22  modify by liqian 标印第一行加 NISCO
     sNisco = sNullstr & "NISCO "
'      sNisco = Chr(128)
    
    '有重量标识要求的编辑重量信息
    If iStdspec_yy = "GB 713-2008" Or iStdspec_yy = "GB 3531-2008" Or iStdspec_yy = "GB 19189-2011" Or iStdspec_yy = "GB 713-2014" Or iStdspec_yy = "GB 3531-2014" Then
        sWgt = "  T.W." & sWgt & "t"
    Else
        sWgt = ""
    End If
    
    '编辑探伤信息
    '如果钢板要求探伤，喷印第四行加喷 T
    ss1.Col = SPD_CUR_UST:    sAdd_T = ss1.Text
    If sAdd_T = "" Then
       ss1.Col = SPD_UST:    sAdd_T = ss1.Text
    End If
    
    sStr_Len = Len("   " & sMark_no & sWgt & " " & sAdd_T & " " & sClass)
    Paint(0) = Chr(75) & Chr(sStr_Len) & "   " & sMark_no & sWgt & " " & sAdd_T & " " & sClass
    
    sStr_Len = Len("   " & sSpec1 & " " & sSinspunita)
    Paint(1) = Chr(75) & Chr(sStr_Len) & "   " & sSpec1 & " " & sSinspunita
        
'    sStr_Len = Len("   " & sSpec1)
'    Paint(1) = Chr(75) & Chr(sStr_Len) & "   " & sSpec1
'    sStr_Len = Len("   " & sSpec1 & " " & sSinspunit)
'    Paint(1) = Chr(75) & Chr(sStr_Len) & "   " & sSpec1 & " " & sSinspunit
    sStr_Len = Len("   " & sThk & sFlag & sWid & sFlag & sLen & sNullstr & sProd_Date & sNullstr & sGroup)
    Paint(2) = Chr(75) & Chr(sStr_Len) & "   " & sThk & sFlag & sWid & sFlag & sLen & sNullstr & sProd_Date & sNullstr & sGroup
    
    ss1.ROW = ss1.ActiveRow
    
    ss1.Col = SPD_ORD_NO:   sOrderNo = Mid(ss1.Text, 1, 3)
    If sOrderNo = "OB5" Then
        sAdd_W = "W"
    End If
                                    
    ss1.Col = SPD_DEL_TO_DATE:   sDel_To_Date = Mid(ss1.Value, 1, 6)
    If sDel_To_Date < sCurDate Then
        sAdd_S = "S"
    End If
    
    ss1.Col = SPD_JIT_FLAG
    If ss1.Text = "Y" Then
         sJIT_FLAG = "DZ"  '17-DZ
    Else
         sJIT_FLAG = ""
    End If
    
    ss1.Col = SPD_HTM_METH:      sAdd_H = ss1.Text
    
    iPaint_Add = sAdd_W & sAdd_S & sAdd_H
    
    ss1.Col = SPD_VESSEL_NO:        sVESSEL_NO = ss1.Text
    ss1.Col = SPD_SIDE_MARK:        sideMark = ss1.Text
    ss1.Col = SPD_CUST_CD:          sCUST_CD = ss1.Text
    ss1.Col = SPD_TO_CUR_INV:       sTO_CUR_INV = ss1.Text
    ss1.Col = SPD_CUST_CD_SHORT:    sCUST_CD_SHORT = ss1.Text
    ss1.Col = SPD_SEALMEMO:         sEALMEMO = ss1.Text
    
    '编辑喷印第四行
    '如果钢板为子公司产品，喷印第四行首位喷子公司简码+（探伤标识）+（用户加喷信息）
    If opt_line5 Then
            Paint(3) = sCUST_CD_SHORT & "  " & iPaint_Add & " " & sVESSEL_NO
            sStr_Len = Len(sCUST_CD_SHORT & "  " & iPaint_Add & " " & sVESSEL_NO)
    Else
            Paint(3) = sPlate_no & " " & sCUST_CD_SHORT & "  " & iPaint_Add & " " & sVESSEL_NO
            sStr_Len = Len(sPlate_no & " " & sCUST_CD_SHORT & "  " & iPaint_Add & " " & sVESSEL_NO)
    End If
    
    If sJIT_FLAG = "" Then
       Paint(3) = Chr(75) & Chr(sStr_Len) & Paint(3)
    Else
       Paint(3) = Chr(75) & Chr(sStr_Len + 4) & sJIT_FLAG & "  " & Paint(3)
    End If

    PaintStr = Paint(0) & Paint(1) & Paint(2) & Paint(3)

    StrSend(2) = Chr(30)
    StrSend(3) = Chr(30)
    sNull = StrSend(2) & StrSend(3)
    
    PaintStr_CD = Val(TXT_P)
    
    If Len(sSpec_IRS_Logo & sSpec2 & sNullstr & sMark_no & sNullstr) > 30 Then
       sStr_Len = Len(sSpec_IRS_Logo & sSpec2 & sNullstr)
       Punch(0) = Chr(30) & Chr(sStr_Len) & sSpec_IRS_Logo & sSpec2 & sNullstr
       
       sStr_Len = Len(sSpec_Punch_Logo & sSinspunita & sMark_no & sNullstr & sEALMEMO)
       Punch(1) = Chr(30) & Chr(sStr_Len) & sSpec_Punch_Logo & sSinspunita & sMark_no & sNullstr & sEALMEMO
       
'       sStr_Len = Len(sMark_no & sNullstr & sEALMEMO)
'       Punch(1) = Chr(30) & Chr(sStr_Len) & sMark_no & sNullstr & sEALMEMO
    Else
       sStr_Len = Len(sSpec_IRS_Logo & sSpec2 & sNullstr & sMark_no & sNullstr)
       Punch(0) = Chr(30) & Chr(sStr_Len) & sSpec_IRS_Logo & sSpec2 & sNullstr & sMark_no & sNullstr
       
       sStr_Len = Len(sSpec_Punch_Logo & sSinspunita & sEALMEMO)
       Punch(1) = Chr(30) & Chr(sStr_Len) & sSpec_Punch_Logo & sSinspunita & sEALMEMO
       
'       sStr_Len = Len(sEALMEMO)
'       Punch(1) = Chr(30) & Chr(sStr_Len) & sEALMEMO
    End If
    
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
    
    If chk_Cond(8) = 1 And sEdge = 1 Then
    
            sEdgeString = sClass & " " & sMark_no
            sEdgeString = Trim(sEdgeString) & " " & Trim(sSpec2) & " " & sSinspunita & " " & Trim(sThk) & "X" & Trim(sWid) & "X" & Trim(sLen)
'            sEdgeString = Trim(sEdgeString) & " " & Trim(sSpec2) & " " & Trim(sThk) & "X" & Trim(sWid) & "X" & Trim(sLen)
'            sEdgeString = Trim(sEdgeString) & " " & Trim(sSpec2) & " " & Trim(sThk) & "X" & Trim(sWid) & "X" & Trim(sLen)
            sEdgeString = sEdgeString & " " & sideMark
            sEdgeString = Trim(sEdgeString) & " " & sCUST_CD & " " & sTO_CUR_INV
            sEdgeStr = Trim(sEdgeString)
      
            Winsock2.SendData sEdgeStr
        
    End If
    
    If chk_Cond(0) = 1 Then
        
            Winsock1.SendData Header & "  " & Chr(16) & Chr(14) & sMark_no
            
            Winsock1.SendData sPaint_Logo1 & sPaint_Logo2 & sPaint_Logo3 & sPaint_Logo4 & sPunch_Logo1 & sPunch_Logo2 & sPunch_Logo3 & sPunch_Logo4
            
            Winsock1.SendData HiByte(Val(sPAINTNUM)) '标识次数 新增
            Winsock1.SendData LoByte(Val(sPAINTNUM)) '标识次数 新增
            
            Winsock1.SendData HiByte(Val(0)) '标识位置，默认0 新增
            Winsock1.SendData LoByte(Val(0)) '标识位置，默认0 新增
            
            Winsock1.SendData HiByte(Val(sWid))
            Winsock1.SendData LoByte(Val(sWid))
            
            Winsock1.SendData HiByte(Val(sLen)) '长度 新增
            Winsock1.SendData LoByte(Val(sLen)) '长度 新增
            
            Winsock1.SendData HiByte(sPaint)
            Winsock1.SendData LoByte(sPaint)
            
            Winsock1.SendData HiByte(sPunch)
            Winsock1.SendData LoByte(sPunch)
            
'            Winsock1.SendData HiByte(sEdge)
'            Winsock1.SendData LoByte(sEdge)
        
            Winsock1.SendData PaintStr
            
            Winsock1.SendData HiByte(PaintStr_CD)
            Winsock1.SendData LoByte(PaintStr_CD)
            Winsock1.SendData PunchStr
            
'            Winsock1.SendData HiByte(EdgeStr_CD)   侧喷内容不要
'            Winsock1.SendData LoByte(EdgeStr_CD)
'            Winsock1.SendData EdgeStr
    
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
        DD.rControl.Add Item:=TXT_STDSPEC

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


