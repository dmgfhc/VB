VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form AGC2036C 
   Caption         =   "标印信息发送界面_AGC2036C"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13605
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   13605
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_line 
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
      Left            =   4530
      MaxLength       =   1
      TabIndex        =   73
      Tag             =   "CD_MANA_NO"
      Text            =   "2"
      Top             =   60
      Visible         =   0   'False
      Width           =   480
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
      Height          =   315
      Left            =   0
      MaxLength       =   1
      TabIndex        =   72
      Tag             =   "CD_MANA_NO"
      Text            =   "1"
      Top             =   1620
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.ComboBox cbo_group 
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
      ItemData        =   "AGC2036C.frx":0000
      Left            =   9600
      List            =   "AGC2036C.frx":0010
      TabIndex        =   7
      Top             =   60
      Width           =   645
   End
   Begin VB.TextBox txt_plate_no 
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
      Left            =   1260
      MaxLength       =   14
      TabIndex        =   6
      Tag             =   "物料号"
      Top             =   450
      Width           =   1755
   End
   Begin VB.TextBox txt_plt 
      CausesValidation=   0   'False
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
      Height          =   310
      Left            =   1260
      MaxLength       =   2
      TabIndex        =   5
      Tag             =   "生产工厂"
      Top             =   60
      Width           =   420
   End
   Begin VB.TextBox txt_plt_name 
      CausesValidation=   0   'False
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
      Height          =   310
      Left            =   1680
      TabIndex        =   4
      Tag             =   "机号"
      Top             =   60
      Width           =   1320
   End
   Begin VB.ComboBox CBO_SHIFT 
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
      ItemData        =   "AGC2036C.frx":0024
      Left            =   8955
      List            =   "AGC2036C.frx":0031
      TabIndex        =   3
      Top             =   60
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
      Left            =   7800
      TabIndex        =   2
      Top             =   450
      Width           =   2445
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
      Left            =   4860
      TabIndex        =   1
      Top             =   450
      Width           =   2925
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   16230
      Top             =   0
   End
   Begin VB.ComboBox CBO_sUserID 
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
      ItemData        =   "AGC2036C.frx":0041
      Left            =   10440
      List            =   "AGC2036C.frx":0051
      TabIndex        =   0
      Top             =   450
      Width           =   1215
   End
   Begin InDate.UDate udt_date_fr 
      Height          =   315
      Left            =   4860
      TabIndex        =   8
      Tag             =   "INS_DATE"
      Top             =   60
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
      Left            =   6300
      TabIndex        =   9
      Tag             =   "INS_DATE"
      Top             =   60
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
   Begin InDate.ULabel ULabel20 
      Height          =   315
      Left            =   45
      Top             =   450
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
   Begin SSSplitter.SSSplitter SSSp1 
      Height          =   8325
      Left            =   0
      TabIndex        =   10
      Top             =   810
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   14684
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AGC2036C.frx":0079
      Begin Threed.SSPanel SSPanel1 
         Height          =   2430
         Left            =   0
         TabIndex        =   11
         Tag             =   "172.18.151.145"
         Top             =   0
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   4286
         _Version        =   196609
         BackColor       =   12632319
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox TXT_STLGRD 
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
            Left            =   3060
            TabIndex        =   41
            Top             =   2040
            Width           =   2085
         End
         Begin VB.TextBox SOCK 
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
            TabIndex        =   40
            Tag             =   "轧批号"
            Top             =   3180
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
            Left            =   8970
            TabIndex        =   39
            Top             =   5790
            Visible         =   0   'False
            Width           =   870
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
            Left            =   5340
            TabIndex        =   38
            Top             =   5190
            Visible         =   0   'False
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
            Left            =   5340
            TabIndex        =   37
            Top             =   4830
            Visible         =   0   'False
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
            Left            =   5340
            TabIndex        =   36
            Top             =   4470
            Visible         =   0   'False
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
            Left            =   5340
            TabIndex        =   35
            Top             =   4110
            Visible         =   0   'False
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
            Left            =   9060
            TabIndex        =   34
            Top             =   5160
            Visible         =   0   'False
            Width           =   2385
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
            Left            =   9060
            TabIndex        =   33
            Top             =   4800
            Visible         =   0   'False
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
            Left            =   9060
            TabIndex        =   32
            Top             =   4440
            Visible         =   0   'False
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
            Left            =   9060
            TabIndex        =   31
            Top             =   4080
            Visible         =   0   'False
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
            Left            =   2160
            TabIndex        =   30
            Top             =   1680
            Width           =   2985
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
            Left            =   2160
            MaxLength       =   14
            TabIndex        =   29
            Tag             =   "物料号"
            Top             =   960
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
            Left            =   2790
            MaxLength       =   14
            TabIndex        =   28
            Tag             =   "物料号"
            Top             =   1320
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
            Left            =   3480
            MaxLength       =   14
            TabIndex        =   27
            Tag             =   "物料号"
            Top             =   1320
            Width           =   915
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
            Left            =   2160
            MaxLength       =   14
            TabIndex        =   26
            Tag             =   "物料号"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox TXT_Class_comp 
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
            Left            =   7755
            MaxLength       =   4
            TabIndex        =   25
            Tag             =   "物料号"
            Text            =   "140"
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox TXT_Paint 
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
            Left            =   11325
            MaxLength       =   50
            TabIndex        =   24
            Top             =   1320
            Width           =   3415
         End
         Begin VB.TextBox TXT_Punch 
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
            Left            =   11325
            MaxLength       =   30
            TabIndex        =   23
            Top             =   1680
            Width           =   3415
         End
         Begin VB.TextBox TXT_Special 
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
            Left            =   11325
            MaxLength       =   40
            TabIndex        =   22
            Top             =   2040
            Width           =   3415
         End
         Begin VB.TextBox TXT_Distance 
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
            Left            =   2160
            MaxLength       =   5
            TabIndex        =   21
            Tag             =   "物料号"
            Text            =   "6000"
            Top             =   2040
            Width           =   885
         End
         Begin VB.TextBox TXT_Producer 
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
            Left            =   11325
            MaxLength       =   50
            TabIndex        =   20
            Top             =   960
            Width           =   3415
         End
         Begin VB.TextBox TXT_X3 
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
            Left            =   6795
            MaxLength       =   17
            TabIndex        =   19
            Top             =   1680
            Width           =   2685
         End
         Begin VB.TextBox TXT_X2 
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
            Left            =   6795
            MaxLength       =   34
            TabIndex        =   18
            Top             =   1320
            Width           =   2685
         End
         Begin VB.TextBox TXT_X1 
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
            Left            =   6795
            MaxLength       =   50
            TabIndex        =   17
            Top             =   960
            Width           =   2685
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
            Left            =   6510
            TabIndex        =   16
            Top             =   6480
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.CheckBox chk_Cond 
            BackColor       =   &H00C0C0FF&
            Caption         =   "锁定"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   7
            Left            =   8640
            TabIndex        =   15
            Top             =   2070
            Width           =   810
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
            Left            =   4410
            MaxLength       =   14
            TabIndex        =   14
            Tag             =   "物料号"
            Top             =   1320
            Width           =   735
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
            Left            =   14130
            MaxLength       =   2
            TabIndex        =   13
            Top             =   2460
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
            Left            =   11325
            MaxLength       =   6
            TabIndex        =   12
            Top             =   2460
            Width           =   1095
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   345
            Left            =   270
            TabIndex        =   42
            Top             =   510
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line3 
               Height          =   255
               Left            =   1050
               TabIndex        =   43
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
               Left            =   3090
               TabIndex        =   44
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
         Begin Threed.SSCommand SSCmd_cnn 
            Height          =   315
            Left            =   630
            TabIndex        =   45
            Top             =   5070
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   255
            BackColor       =   14804173
            Caption         =   "连接状态"
         End
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   255
            Top             =   1680
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            Caption         =   "标准/ 牌号"
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
            Left            =   255
            Top             =   960
            Width           =   1875
            _ExtentX        =   3307
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
            Left            =   255
            Top             =   1320
            Width           =   1875
            _ExtentX        =   3307
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
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   5580
            Top             =   2040
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            Caption         =   "Class_comp(1-10)"
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
         Begin Threed.SSFrame SSFrame4 
            Height          =   345
            Left            =   5160
            TabIndex        =   46
            Top             =   105
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line5 
               Height          =   255
               Left            =   450
               TabIndex        =   47
               Top             =   30
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
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
               Caption         =   "Low"
            End
            Begin Threed.SSOption opt_line6 
               Height          =   255
               Left            =   1710
               TabIndex        =   48
               Top             =   30
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
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
               Caption         =   "Medium"
               Value           =   -1
            End
            Begin Threed.SSOption opt_line7 
               Height          =   255
               Left            =   3300
               TabIndex        =   49
               Top             =   30
               Width           =   825
               _ExtentX        =   1455
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
               Caption         =   "High"
            End
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   9900
            Top             =   1320
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            Caption         =   "Cust paint"
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
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   9900
            Top             =   1680
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            Caption         =   "Cust punch"
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
            Left            =   9900
            Top             =   2040
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            Caption         =   "Special"
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
         Begin InDate.ULabel ULabel15 
            Height          =   315
            Left            =   255
            Top             =   2040
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            Caption         =   "Repeat Distance"
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
         Begin Threed.SSFrame SSFrame5 
            Height          =   345
            Left            =   5160
            TabIndex        =   50
            Top             =   510
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line8 
               Height          =   255
               Left            =   450
               TabIndex        =   51
               Top             =   30
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
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
               Caption         =   "3mm"
            End
            Begin Threed.SSOption opt_line9 
               Height          =   255
               Left            =   1710
               TabIndex        =   52
               Top             =   30
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
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
               Caption         =   "8mm"
               Value           =   -1
            End
            Begin Threed.SSOption opt_line10 
               Height          =   255
               Left            =   3300
               TabIndex        =   53
               Top             =   30
               Width           =   825
               _ExtentX        =   1455
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
               Caption         =   "12mm"
            End
         End
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   9900
            Top             =   960
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            Caption         =   "Producer"
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
            Left            =   5580
            Top             =   1320
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "X2"
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
         Begin InDate.ULabel ULabel19 
            Height          =   315
            Left            =   5580
            Top             =   1680
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "X3"
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
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Left            =   5580
            Top             =   960
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "X1"
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
         Begin Threed.SSFrame SSFrame6 
            Height          =   345
            Left            =   9720
            TabIndex        =   54
            Top             =   105
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line11 
               Height          =   255
               Left            =   510
               TabIndex        =   55
               Top             =   30
               Width           =   1815
               _ExtentX        =   3201
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
               Caption         =   "Normal reading"
               Value           =   -1
            End
            Begin Threed.SSOption opt_line12 
               Height          =   255
               Left            =   2640
               TabIndex        =   56
               Top             =   30
               Width           =   1635
               _ExtentX        =   2884
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
               Caption         =   "Up-side down"
            End
         End
         Begin Threed.SSFrame SSFrame8 
            Height          =   345
            Left            =   270
            TabIndex        =   57
            Top             =   105
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   12632319
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
               Left            =   540
               TabIndex        =   60
               Top             =   60
               Value           =   1  'Checked
               Width           =   870
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
               Left            =   1860
               TabIndex        =   59
               Top             =   60
               Value           =   1  'Checked
               Width           =   870
            End
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00C0C0FF&
               Caption         =   " 侧喷加喷"
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
               Index           =   9
               Left            =   3180
               TabIndex        =   58
               Top             =   60
               Value           =   1  'Checked
               Width           =   1200
            End
         End
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   12690
            Top             =   2460
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
         Begin InDate.ULabel ULabel13 
            Height          =   315
            Left            =   9900
            Top             =   2460
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
            Left            =   4050
            TabIndex        =   68
            Top             =   5220
            Visible         =   0   'False
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
            Left            =   4050
            TabIndex        =   67
            Top             =   4860
            Visible         =   0   'False
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
            Left            =   4050
            TabIndex        =   66
            Top             =   4500
            Visible         =   0   'False
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
            Left            =   4050
            TabIndex        =   65
            Top             =   4140
            Visible         =   0   'False
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
            Left            =   7890
            TabIndex        =   64
            Top             =   4800
            Visible         =   0   'False
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
            Left            =   7890
            TabIndex        =   63
            Top             =   5160
            Visible         =   0   'False
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
            Left            =   7890
            TabIndex        =   62
            Top             =   4110
            Visible         =   0   'False
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
            Left            =   7890
            TabIndex        =   61
            Top             =   4470
            Visible         =   0   'False
            Width           =   1125
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   5865
         Left            =   0
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2460
         Width           =   7380
         _Version        =   393216
         _ExtentX        =   13018
         _ExtentY        =   10345
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ButtonDrawMode  =   4
         ColsFrozen      =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   25
         MaxRows         =   10
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "AGC2036C.frx":00EB
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   5865
         Left            =   7410
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2460
         Width           =   7755
         _Version        =   393216
         _ExtentX        =   13679
         _ExtentY        =   10345
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   47
         MaxRows         =   10
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "AGC2036C.frx":4392
      End
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   0
      Left            =   45
      Top             =   60
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
      Left            =   3210
      Top             =   60
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
      Left            =   3210
      Top             =   450
      Width           =   1620
      _ExtentX        =   2858
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
      Left            =   7950
      Top             =   60
      Width           =   975
      _ExtentX        =   1720
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
      Left            =   15330
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "191.168.1.100"
      RemotePort      =   5080
   End
   Begin CSTextLibCtl.sitxEdit TXT_CUT_TIME 
      Height          =   315
      Left            =   9540
      TabIndex        =   71
      Tag             =   "出炉时间"
      Top             =   9120
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
         Name            =   "Arial"
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
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   16
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   3735
      Top             =   60
      Visible         =   0   'False
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   10440
      Top             =   60
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "作业人员"
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
   Begin InDate.ULabel UL_conn 
      Height          =   315
      Left            =   16800
      Top             =   60
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "连接状态"
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
      Height          =   705
      Left            =   11670
      TabIndex        =   74
      Top             =   60
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1244
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "侧喷"
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
         Left            =   90
         TabIndex        =   77
         Top             =   390
         Width           =   750
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "标印"
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
         Left            =   90
         TabIndex        =   75
         Top             =   60
         Width           =   750
      End
      Begin VB.Label tcpMsg2 
         BackColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   1290
         TabIndex        =   78
         Top             =   420
         Width           =   2055
      End
      Begin VB.Shape tcpStatus2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   225
         Left            =   900
         Shape           =   3  'Circle
         Top             =   390
         Width           =   285
      End
      Begin VB.Shape tcpStatus 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   225
         Left            =   900
         Shape           =   3  'Circle
         Top             =   75
         Width           =   285
      End
      Begin VB.Label tcpMsg 
         BackColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   1290
         TabIndex        =   76
         Top             =   105
         Width           =   2055
      End
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   15330
      Top             =   870
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   25298
   End
End
Attribute VB_Name = "AGC2036C"
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
'-- Program ID        AGC2036C
'-- Document No       Q-00-0010(Specification)
'-- Designer          LiQian
'-- Coder             LiQian
'-- Date              2011.6.13
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

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

' 2012-02-20  liqian 加订单材/余材代码显示,需要改变左边列,以前左边列和右边列共用列变量,故重新定义左边变量,所有SS1变量地方做更改.
Const SS1_LINE1 = 1                        '1#线
Const SS1_LINE2 = 2                        '2#线
Const SS1_PLATE_NO = 3                     '钢板号
Const SS1_THK = 4                          '厚
Const SS1_WID = 5                          '宽
Const SS1_LEN = 6                          '长
Const SS1_WGT = 7                          '重
Const SS1_LAST_YN = 8                      '尾板
Const SS1_SIZE_KND = 9                     '定尺
Const SS1_TRIM_FL = 10                     '切边
Const SS1_APLY_STDSPEC = 11                '标准号
Const SS1_APLY_STDSPEC_NEW = 12            '改判标准号
Const SS1_SURF_GRD = 13                    '合格
Const SS1_MARK_YN = 14                     '喷印
Const SS1_STAMP_YN = 15                    '冲印
Const SS1_BAR_YN = 16                      '侧喷
Const SS1_PROD_DATE = 17                   '生产日期
Const SS1_EMP_CD = 18                      '作业人员
Const SS1_PAINT = 19                       '打印
Const SS1_LABEL = 20                       '标签
Const SS1_STANDARD = 21                    '打印标准
Const SS1_STLGRD = 22                      '钢种
Const SS1_PLAN_LEN = 24                    '计划长度
Const SS1_URGNT_FL = 25                    '紧急订单绿色标记 2012-08-16  by  LiQian
Const SS2_sPaint = 15   '喷印
Const SS2_sPunch = 16   '冲印
Const SS2_sEdge = 17    '侧喷

Const SS2_PRODSPECNOA_STD = 38 '多船级社标准一
Const SS2_PRODSPECNOB_STD = 39 '多船级社标准二
Const SS2_PRODSPECNOC_STD = 40 '多船级社标准三
Const SS2_PRODSPECNOA = 41 '多船级社牌号一
Const SS2_PRODSPECNOB = 42 '多船级社牌号二
Const SS2_PRODSPECNOC = 43 '多船级社牌号三
Const SS2_INSP_CD = 44     '见证机关
Const SS2_SIDEMARK = 37     '侧喷加喷
Const SS2_SURFACE_REQUESTS = 45
Const SS2_CLASS_CD = 46
Const SS2_CLASS_LVL = 47

Const SPD_LINE1 = 1                        '1#线
Const SPD_LINE2 = 2                        '2#线
Const SPD_ORD_FL = 3                              ' 订单材/余材 2012-02-20 LiQian
Const SPD_PLATE_NO = 4                     '钢板号
Const SPD_THK = 5                          '厚
Const SPD_WID = 6                          '宽
Const SPD_LEN = 7                          '长
Const SPD_WGT = 8                          '重
Const SPD_LAST_YN = 9                      '尾板
Const SPD_SIZE_KND = 10                    '定尺
Const SPD_TRIM_FL = 11                     '切边
Const SPD_APLY_STDSPEC = 12                '标准号
Const SPD_APLY_STDSPEC_NEW = 13            '改判标准号
Const SPD_SURF_GRD = 14                    '合格
Const SPD_MARK_YN = 15                     '喷印
Const SPD_STAMP_YN = 16                    '冲印
Const SPD_BAR_YN = 17                      '侧喷
Const SPD_PROD_DATE = 18                   '生产日期
Const SPD_EMP_CD = 19                      '作业人员
Const SPD_PAINT = 20                       '打印
Const SPD_LABEL = 21                       '标签
Const SPD_STANDARD = 22                    '打印标准
Const SPD_STLGRD = 23                      '钢种
Const SPD_VESSEL_NO = 24                   '加喷内容
Const SPD_COLOR_STROKE = 25                '色标及标准
Const SPD_CE_YN = 26                       'CE
Const SPD_TS_YN = 27                       'TS
Const SPD_JIS_YN = 28                      'JIS ADD HANCHAO
Const SPD_UST_FL = 29                      'UST MD HANCHAO
Const SPD_CUST_CD = 30                     '客户代码 MD HANCHAO
Const SPD_TO_CUR_INV = 31                  '目的库 MD HANCHAO
Const SPD_HTM_METH = 32                    '热处理指示 MD HANCHAO
Const SPD_DEL_TO_DATE = 33                 '交货期 MD HANCHAO
Const SPD_PROC_CD = 34                     '进程代码 MD HANCHAO
Const SPD_CERT_TYPE = 35                   '质保书类型 MD HANCHAO
Const SPD_CUST_PUNCH = 36                  'MD HANCHAO

Dim iSS As String
Dim iF_mm As Double
Dim iT_mm As Double
Dim PRODSPECNOA As Integer '牌号一
Dim PRODSPECNOB As Integer '牌号二
Dim PRODSPECNOC As Integer '牌号三
'更新多船级社相关字段 2013-12-31
Dim Ship_Emblem As Integer
Dim First_Number As Integer
Dim Second_Number As Integer
Dim Third_Number As Integer
Dim Firth_Number As Integer
Dim Fifth_Number As Integer
Dim Sixth_Number As Integer
Dim sPaint As Integer           '7         '是否标印
Dim sPunch As Integer           '8         '是否冲印
Dim sEdge As Integer            '9         '是否侧喷
'更新多船级社相关字段 2013-12-31
'Dim sComp As Integer            '17        '船徽

Dim sInsp_cd As String     '见证机关
Dim sideMark As String     '侧喷加喷


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpvDest As Any, ByRef lpvSrc As Any, ByVal cbLength As Long)
Public Property Get HLByte(ByRef Word As Long, HL As Long) As Byte
CopyMemory HLByte, ByVal VarPtr(Word) + HL, 1
End Property

Public Property Get HiByte(ByRef Word As Integer) As Byte
CopyMemory HiByte, ByVal VarPtr(Word) + 1, 1
End Property

Public Property Get LoByte(ByRef Word As Integer) As Byte
CopyMemory LoByte, ByVal VarPtr(Word), 1
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
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2036C.P_REFER", Key:="P-R"
    sc1.Add Item:="AGC2036C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="AGC2036C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 31, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 32, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 33, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 34, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 35, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '质保书类型
   Call Gp_Sp_Collection(ss2, 36, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 37, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 38, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '多船级社标准一 20140101
   Call Gp_Sp_Collection(ss2, 39, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '多船级社标准二 20140101
   Call Gp_Sp_Collection(ss2, 40, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '多船级社标准三 20140101
   Call Gp_Sp_Collection(ss2, 41, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '多船级社牌号一 20140101
   Call Gp_Sp_Collection(ss2, 42, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '多船级社牌号二 20140101
   Call Gp_Sp_Collection(ss2, 43, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '多船级社牌号三 20140101
   Call Gp_Sp_Collection(ss2, 44, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '见证机关 20140227
   Call Gp_Sp_Collection(ss2, 45, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 46, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 47, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AGC2036C.P_SREFER", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
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
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    
    Call Gp_Sp_ColHidden(ss1, SS1_LINE1, True)            '1#隐藏,2#显示
    Call Gp_Sp_ColHidden(ss1, SS1_PAINT, True)
    Call Gp_Sp_ColHidden(ss1, SS1_LABEL, True)
    Call Gp_Sp_ColHidden(ss1, SS1_MARK_YN, True)
    Call Gp_Sp_ColHidden(ss1, SS1_STAMP_YN, True)
    Call Gp_Sp_ColHidden(ss1, SS1_BAR_YN, True)

    Call Gp_Sp_ColHidden(ss2, SPD_LINE1, True)            '1#隐藏,2#显示
    Call Gp_Sp_ColHidden(ss2, SPD_SURF_GRD, True)
    Call Gp_Sp_ColHidden(ss2, SPD_LAST_YN, True)
    Call Gp_Sp_ColHidden(ss2, SPD_SIZE_KND, True)
    Call Gp_Sp_ColHidden(ss2, SPD_TRIM_FL, True)
    Call Gp_Sp_ColHidden(ss2, SPD_APLY_STDSPEC, True)
    Call Gp_Sp_ColHidden(ss2, SPD_APLY_STDSPEC_NEW, True)
    Call Gp_Sp_ColHidden(ss2, SPD_PROD_DATE, True)
    Call Gp_Sp_ColHidden(ss2, SPD_EMP_CD, True)
    Call Gp_Sp_ColHidden(ss2, SPD_PAINT, True)
    Call Gp_Sp_ColHidden(ss2, SPD_LABEL, True)
'    Call Gp_Sp_ColHidden(ss2, SPD_MARK_YN, True)
'    Call Gp_Sp_ColHidden(ss2, SPD_STAMP_YN, True)
'    Call Gp_Sp_ColHidden(ss2, SPD_BAR_YN, True)
    Call Gp_Sp_ColHidden(ss2, SPD_PROC_CD, True)
    Call Gp_Sp_ColHidden(ss2, SPD_CERT_TYPE, True)
    Call Gp_Sp_ColHidden(ss2, SS2_PRODSPECNOA, True) '多船级社牌号一隐藏 20140101
    Call Gp_Sp_ColHidden(ss2, SS2_PRODSPECNOB, True) '多船级社牌号二隐藏 20140101
    Call Gp_Sp_ColHidden(ss2, SS2_PRODSPECNOC, True) '多船级社牌号三隐藏 20140101
    
    txt_plt.Text = "C1"
    Call txt_plt_KeyUp(0, 0)
    Call Gf_USER_ComboAdd(M_CN1, CBO_sUserID, "AGC2036C")
    
    txt_line.Text = "2"                                  '2#线''
    txt_rec_sts.Text = "1"

    iSS = ""

    opt_line3 = True
    
'    Winsock1.RemoteHost = "172.18.56.194" 'Gf_ComnNameFind(M_CN1, "G0034", "01", 1)
'    Winsock1.RemotePort = "2222" 'Gf_ComnNameFind(M_CN1, "G0034", "01", 2)
'    Winsock2.RemoteHost = "172.18.43.113" 'Gf_ComnNameFind(M_CN1, "G0034", "01", 1)
'    Winsock2.RemotePort = "25298" 'Gf_ComnNameFind(M_CN1, "G0034", "01", 2)

    Winsock1.RemoteHost = Gf_ComnNameFind(M_CN1, "G0036", "02", 1)
    Winsock1.RemotePort = Gf_ComnNameFind(M_CN1, "G0036", "02", 2)
    Winsock2.RemoteHost = Gf_ComnNameFind(M_CN1, "G0038", "02", 1)
    Winsock2.RemotePort = Gf_ComnNameFind(M_CN1, "G0038", "02", 2)

    iF_mm = Val(Gf_ComnNameFind(M_CN1, "G0037", "01", 1))
    iT_mm = Val(Gf_ComnNameFind(M_CN1, "G0037", "02", 1))
    
    
    If Mid(sAuthority, 3, 1) = "1" Then
       chk_Cond(0).Enabled = True
       chk_Cond(8).Enabled = True
    Else
       chk_Cond(0).Enabled = False
       chk_Cond(8).Enabled = False
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Winsock1.State = 1 Or Winsock1.State = 7 Or Winsock1.State = 9 Then
       Winsock1.Close
    End If
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    
'    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
'        Cancel = 1
'        Exit Sub
'    End If
    
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
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) And Gf_Sp_Cls(sc2) Then
    
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        

        txt_plt.Text = "C1"
        Call txt_plt_KeyUp(0, 0)
        txt_line.Text = "2"                           '改2#线''
        txt_rec_sts.Text = "1"
        opt_line3 = True
        txt_stdspec_chg = ""
        
    End If
    
End Sub

Public Sub Form_Exc()
    
    If iSS = "ss2" Then
       Call Gp_Sp_Excel(Me, sc2.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf iSS = "ss1" Then
       Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Else
       Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If

End Sub

Public Sub Form_Ref()
    
    Dim iCount      As Integer
    Dim sPlateNo    As String
    
    Dim inum As Integer
    Dim lRow As Integer
    
    Dim sCurDate As String
    Dim sDel_To_Date As String
    Dim sSurf_Grd As Integer
    Dim sproc_cd As String
    Dim sUrgnt_Fl As String
    
    If Gf_Sp_ProceExist(sc1.Item("Spread")) Then Exit Sub
    
    sCurDate = Format(Now, "YYYYMM")
            
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
    
        '如果选中锁定框，查询时右侧表单不刷新
        If chk_Cond(7) Then
        Else
           Call Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False)
        End If
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss1.OperationMode = OperationModeNormal
        ss2.OperationMode = OperationModeNormal
        
    End If
    
    With ss1
        For iCount = 1 To .MaxRows
            .Row = iCount:            .Col = SS1_PLATE_NO
             sPlateNo = .Text
            If Left(.Text, 12) = Left(sPlateNo, 12) Then
            Else
               .Row = iCount - 1
               .Col = SS1_LAST_YN
               .Value = 1
            End If
            
             '紧急订单绿色标记 2012-11-08  by  LiQian
            .Row = iCount:            .Col = SS1_URGNT_FL
            sUrgnt_Fl = .Text
             If ss1.Text = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, SS1_PLATE_NO, SS1_PLATE_NO, .Row, .Row, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_URGNT_FL, SS1_URGNT_FL, .Row, .Row, &HC000&)
            End If
            
        Next iCount
    End With
    
    With ss2
        For iCount = 1 To .MaxRows
            '超交货期警示
            .Row = iCount:            .Col = SPD_DEL_TO_DATE
            sDel_To_Date = Mid(.Value, 1, 6)
            If sDel_To_Date < sCurDate Then
              .Row = iCount:           .Col = SPD_SURF_GRD:              sSurf_Grd = .Value
                                       .Col = SPD_PROC_CD:               sproc_cd = Mid(.Text, 1, 1)
              If sSurf_Grd = 1 And sproc_cd <> "X" Then
                 Call Gp_Sp_BlockColor(ss2, 1, .MaxCols, iCount, iCount, &HFF&)
              End If
            End If
        Next iCount
    End With

End Sub

Public Sub Form_Pro()

On Error GoTo Process_Exec_ERROR

    Dim SMESG As String

    Dim iRow As Integer
    
    Dim sPlate_No As String
    Dim sThk As String
    Dim sWid As String
    Dim sLen As String
    Dim sWgt As String
    Dim sEmp_Cd As String
    Dim sInv As String
    Dim sSpec As String
    Dim sStdspec_yy As String
    
    Dim sMesg_Fl As String
    Dim sRst_Len As Double
    Dim sPlan_Len As Double
    
    Dim sGroup As String
'    Dim sPaint As String
'    Dim sPunch As String
'    Dim sEdge As String

    sMesg_Fl = "N"
    
    sGroup = Trim(cbo_group.Text)
    sEmp_Cd = Trim(CBO_sUserID.Text)

'    If chk_Cond(0) Then
'           sPaint = 1
'    Else
'           sPaint = 0
'    End If
'    If chk_Cond(3) Then
'           sPunch = 1
'    Else
'           sPunch = 0
'    End If
'    If chk_Cond(4) Then
'           sEdge = 1
'    Else
'           sEdge = 0
'    End If

    If txt_rec_sts = "1" Or txt_rec_sts = "2" Then
        
        For iRow = 1 To ss1.MaxRows
        
             ss1.Col = 0
             ss1.Row = iRow
             If ss1.Text = "Update" Then
                ss1.Col = SS1_LEN:         sRst_Len = Val(ss1.Text)
                ss1.Col = SS1_PLAN_LEN:    sPlan_Len = Val(ss1.Text)
                If sPlan_Len <> sRst_Len Then
                   Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFFC0FF)
                   sMesg_Fl = "Y"
                End If
             End If

        Next iRow
        
       If sMesg_Fl = "Y" Then
            SMESG = " 钢板长度与计划长度不一致，您确认保存么？"
            If Gf_MessConfirm(SMESG, "Q") Then
                If Gf_Sp_Process(M_CN1, Proc_Sc("SC1"), Mc1) Then
                    If chk_Cond(7) Then
                    Else
                       Call Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False)
                    End If
                    ss1.OperationMode = OperationModeNormal
                    ss2.OperationMode = OperationModeNormal
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                    Call MenuTool_ReSet
                End If
            Else
                Exit Sub
            End If
       Else
            If Gf_Sp_Process(M_CN1, Proc_Sc("SC1"), Mc1) Then
                If chk_Cond(7) Then
                Else
                   Call Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False)
                End If
                ss1.OperationMode = OperationModeNormal
                ss2.OperationMode = OperationModeNormal
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call MenuTool_ReSet
            End If
       End If
        
    End If
    
    With ss2
        For iRow = 1 To .MaxRows
             .Col = 0
             .Row = iRow
             If ss2.Text = "Update" Or ss2.Text = "Insert" Or ss2.Text = "Delete" Then
                   .Col = SPD_PLATE_NO:         sPlate_No = .Text
                If (chk_Cond(0) Or chk_Cond(8)) And ss2.Text <> "Delete" Then
                   'Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, iRow, iRow, , &HFFC0FF)
                   If sGroup <> "A" And sGroup <> "B" And sGroup <> "C" And sGroup <> "D" Then
                        SMESG = " 班别错误，请确认是否正确输入班别"
                        Call Gp_MsgBoxDisplay(SMESG)
                        Exit Sub
                   End If
                   Call Cmd_SEND
                   If Gp_Send_Size_Exec(sPlate_No) <> "" Then
                      MsgBox (" 向几何尺寸测量设备发送指示失败 --〉 " + Gp_Send_Size_Exec(sPlate_No))
                   End If
                    '  2011-8-24  modify  by   LiQian
                '  更新标印、冲印、侧喷栏位,依照画面选择是否标印、冲印、或侧喷
                   Call Gp_Send_Paint(sPlate_No, sPaint, sPunch, sEdge, sEmp_Cd)
                '  重新查询画面,如果标印过,则此块从画面消失,无法再查询到,防止标印工重复点击造成的重号问题
                   Call Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False)
                   ss2.OperationMode = OperationModeNormal
                End If
                Exit For
             End If
        Next iRow
    End With

    iRow = iRow + 10
    If iRow > ss1.MaxRows Then
       iRow = ss1.MaxRows
    End If
    
    Call ss1.SetActiveCell(SS1_LEN, iRow)
    
Process_Exec_ERROR:

    Screen.MousePointer = vbDefault
    
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
    Dim sEmp_Cd     As String
    Dim sInv        As String
    Dim sStdspec_yy As String
    Dim sStdspec    As String
    
    Dim iCount As Integer
    
    sPlateNo = ""
    
    With ss1
        If .MaxRows = 0 Then
           If Len(txt_plate_no.Text) = 12 Then
               Call Gp_Sp_Ins(Proc_Sc("Sc1"))
              .Row = 1
              .Col = SS1_PLATE_NO
              .Text = txt_plate_no.Text & "01"
              .Col = SS1_THK:           .Value = 0
              .Col = SS1_WID:           .Value = 0
              .Col = SS1_LEN:           .Value = 0
              .Col = SS1_WGT:           .Value = 0
              .Col = SS1_APLY_STDSPEC:  .Text = "GB-XXX"
           Else
               Call Gp_MsgBoxDisplay("请正确输入母板号 ！")
           End If
           Exit Sub
        End If
        For iCount = .ActiveRow To .MaxRows
            .Row = iCount
            .Col = SS1_PLATE_NO
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
    Call Gp_Sp_Ins(Proc_Sc("Sc1"))

    With ss1
        .ReDraw = False
        If lRow > 0 Then
            .Row = lRow
            .Col = SS1_PLATE_NO:      sPlateNo = .Text
            .Col = SS1_THK:           dThk = Val(.Value) 'Val(.Text & "")
            .Col = SS1_WID:           dWid = Val(.Value) 'Val(.Text & "")
            .Col = SS1_LEN:           dLen = Val(.Value) 'Val(.Text & "")
            .Col = SS1_WGT:           dWgt = Val(.Value) 'Val(.Text & "")
            .Col = SS1_SIZE_KND:      sSize_knd = .Value
            .Col = SS1_TRIM_FL:       sTrim_fl = .Value
            .Col = SS1_APLY_STDSPEC:  sAply_stdspec = .Text
            .Col = SS1_STANDARD:      sStdspec_yy = .Text
            .Col = SS1_EMP_CD:        sEmp_Cd = .Text
            .Col = SS1_STLGRD:        sStdspec = .Text
        Else
            sPlateNo = txt_plate_no.Text & "00"
        End If

        .Row = lRow + 1
        .Col = SS1_PLATE_NO:      .Text = sPlateNo
        .Col = SS1_THK:           .Value = dThk
        .Col = SS1_WID:           .Value = dWid
        .Col = SS1_LEN:           .Value = dLen
        .Col = SS1_WGT:           .Value = dWgt
        .Col = SS1_SIZE_KND:      .Value = sSize_knd
        .Col = SS1_TRIM_FL:       .Value = sTrim_fl
        .Col = SS1_APLY_STDSPEC:  .Text = sAply_stdspec
        .Col = SS1_EMP_CD:        .Text = sEmp_Cd
        .Col = SS1_STANDARD:      .Text = sStdspec_yy
        .Col = SS1_STLGRD:        .Text = sStdspec
        .Col = 0:                 .Text = "Input"
        .Col = SS1_PLATE_NO: .Text = Mid(.Text, 1, 12) & Format(Val(Mid(.Text, 13, 2) & "") + 1, "00")
        .Col = SS1_SURF_GRD:      .Value = 1
        .Col = SS1_MARK_YN:       .Value = 1
        .Col = SS1_STAMP_YN:      .Value = 1
        .Col = SS1_BAR_YN:        .Value = 1
        .Col = SS1_LINE1:         .Value = 1
        .Col = 0:                 .Text = "Input"
        
         Call .SetActiveCell(1, .Row)
        .ReDraw = True
    End With

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
    ss1.Row = ss1.ActiveRow:        ss1.Col = SS1_EMP_CD:        ss1.Text = sUserID
    Call Gp_Sp_Del(Proc_Sc("Sc1"))
End Sub

Public Sub Spread_Can()
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc1"))
End Sub

Public Sub Form_Exit()
    Unload Me
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

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    Dim lRow As Integer
    Dim sCheck1 As String
    Dim sCheck2 As String
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_GotFocus()
    iSS = "ss1"
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    Dim lRow As Integer
    Dim sCheck1 As String
    Dim sCheck2 As String
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    Dim sCheck1 As Integer
    Dim sCheck2 As Integer
    
    Dim iCol As Long
    Dim iRow As Long
    Dim iMode As Integer
    
    Dim iRowNum As Long
    Dim iRowfr As Long
    Dim iRowto As Long
    
    iCol = Col
    iRow = Row

    If Row <= 0 Then Exit Sub
    If Col <> SS1_LINE1 And Col <> SS1_LINE2 Then Exit Sub
    If Not Gf_Sc_Authority(sAuthority, "U") Then Exit Sub
    
    iRowto = iRow - 1
    iRowfr = iRow + 1

    ss1.Row = iRow

    If Col = SS1_LINE1 And ButtonDown = 1 Then
        ss1.Col = SS1_LINE2
        ss1.Text = 0
    ElseIf Col = SS1_LINE2 And ButtonDown = 1 Then
        ss1.Col = SS1_LINE1
        ss1.Text = 0
    End If

    ss1.Col = 0
    ss1.Text = "Update"

    ss1.Col = SS1_LINE1
    sCheck1 = ss1.Value
    ss1.Col = SS1_LINE2
    sCheck2 = ss1.Value

    If sCheck1 = 0 And sCheck2 = 0 Then
        ss1.Col = 0
        ss1.Text = ""
    End If
    
        ss1.Col = SS1_EMP_CD
        ss1.Text = CBO_sUserID.Text
        
        ss1.Col = SS1_LABEL
        If chk_Cond(1) Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        
        ss1.Col = SS1_PAINT
        If chk_Cond(0) Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        
        ss1.Col = SS1_STAMP_YN
        If chk_Cond(3) Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        ss1.Col = SS1_BAR_YN
        If chk_Cond(4) Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        
        Call Cmd_SEND_SET(Row)
    
End Sub
Private Sub ss2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    Dim sCheck1 As Integer
    Dim sCheck2 As Integer
    
    Dim iCol As Long
    Dim iRow As Long
    Dim iMode As Integer
    
    Dim iRowNum As Long
    Dim iRowfr As Long
    Dim iRowto As Long
    
    iCol = Col
    iRow = Row

    If Row <= 0 Then Exit Sub
    If Col <> SPD_LINE1 And Col <> SPD_LINE2 Then Exit Sub
    If Not Gf_Sc_Authority(sAuthority, "U") Then Exit Sub
    
    iRowto = iRow - 1
    iRowfr = iRow + 1
    
    If iRowto > 0 Then
        For iRowNum = 1 To iRowto
             
             ss2.Col = 0
             ss2.Row = iRowNum
             If ss2.Text <> "" Then
                ss2.Text = ""
                ss2.Col = SPD_LINE1
                ss2.Value = 0
                ss2.Col = SPD_LINE2
                ss2.Value = 0
                Exit For
             End If
        Next iRowNum
    End If
    
    If iRowfr <= ss2.MaxRows Then
        For iRowNum = iRowfr To ss2.MaxRows
             
             ss2.Col = 0
             ss2.Row = iRowNum
             If ss2.Text <> "" Then
                ss2.Text = ""
                ss2.Col = SPD_LINE1
                ss2.Value = 0
                ss2.Col = SPD_LINE2
                ss2.Value = 0
                Exit For
             End If
        Next iRowNum
    End If

    ss2.Row = iRow

    If Col = SPD_LINE1 And ButtonDown = 1 Then
        ss2.Col = SPD_LINE2
        ss2.Text = 0
    ElseIf Col = SPD_LINE2 And ButtonDown = 1 Then
        ss2.Col = SPD_LINE1
        ss2.Text = 0
    End If

    ss2.Col = 0
    ss2.Text = "Update"
    Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, iRow, iRow, , &HFFC0FF)
    ss2.Col = SPD_LINE1
    sCheck1 = ss2.Value
    ss2.Col = SPD_LINE2
    sCheck2 = ss2.Value

    If sCheck1 = 0 And sCheck2 = 0 Then
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, iRow, iRow, , &H8000000E)
        ss2.Col = 0
        ss2.Text = ""
    End If
    
        ss2.Col = SPD_EMP_CD
        ss2.Text = CBO_sUserID.Text
        
        ss2.Col = SPD_LABEL
        If chk_Cond(1) Then
           ss2.Value = 1
        Else
           ss2.Value = 0
        End If
        
'将原有的根据控件修改标印的功能删除 2013-12-31
'        ss2.Col = SPD_PAINT
'        If chk_Cond(0) Then
'           ss2.Value = 1
'        Else
'           ss2.Value = 0
'        End If
'
'        ss2.Col = SPD_STAMP_YN
'        If chk_Cond(3) Then
'           ss2.Value = 1
'        Else
'           ss2.Value = 0
'        End If
'        ss2.Col = SPD_BAR_YN
'        If chk_Cond(4) Then
'           ss2.Value = 1
'        Else
'           ss2.Value = 0
'        End If
        
        Call Cmd_SEND_SET(Row)
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

  If Row <= 0 Then Exit Sub
  

  ss1.Row = Row
     
  If Col = SS1_APLY_STDSPEC_NEW Then
     ss1.Col = Col
     If ss1.Text = "" Then
        ss1.Text = txt_stdspec_chg
        If txt_stdspec_chg <> "" Then
            ss1.Col = SS1_SURF_GRD
            ss1.Value = 0
        End If
     Else
        ss1.Text = ""
        ss1.Col = SS1_SURF_GRD
        ss1.Value = 1
     End If
  End If

  If Col = SS1_PROD_DATE Then
     TXT_CUT_TIME.RawData = Gf_DTSet(M_CN1, , "X")
     ss1.Col = SS1_PROD_DATE
     ss1.Text = TXT_CUT_TIME.Text
  End If
  
End Sub

Private Sub ss2_GotFocus()
    iSS = "ss2"
End Sub

Private Sub Cmd_SEND_SET(ByVal Row As Long)
    
    Dim Header As String * 2
    Dim Nisco As String
    Dim sFlag As String
    Dim sNull As String
    
    Dim sPlate_No As String
    Dim sThk As String
    Dim sWid As String
    Dim sLen As String
    Dim sWgt As String
    Dim sSpec As String
    Dim sSpec_Str As String
    Dim sNum As Integer
    Dim sCE_YN As String
    Dim sTS_YN As String
    Dim jIS_YN As String 'md hanchao
    Dim sCEdate As String
    
    Dim sSURFACE As String
    Dim sClasscd As String
    Dim sClasslvl As String
    
    sCEdate = Format(Now, "YY")

    If iSS = "ss1" Or iSS = "" Then
        With ss1
            .Row = Row
            .Col = SS1_PLATE_NO:         TXT_MAT_NO = .Text
            .Col = SS1_THK:              TXT_THK = Trim(Str(.Text))
            .Col = SS1_WID:              TXT_WID = Trim(Str(.Text))
            .Col = SS1_LEN:              TXT_LEN = Trim(Str(.Text))
            .Col = SS1_WGT:              TXT_WGT = Trim(Str(.Text))
            .Col = SS1_STANDARD:         TXT_SPEC = .Text
            .Col = SS1_STLGRD:           TXT_STLGRD = .Text
            .Col = SS1_APLY_STDSPEC_NEW: sSpec = .Text
            If sSpec = "" Then
               .Col = SS1_APLY_STDSPEC:  sSpec = .Text
            End If
        End With
    ElseIf iSS = "ss2" Then
        With ss2
            .Row = Row
            .Col = SPD_PLATE_NO:         TXT_MAT_NO = .Text
            .Col = SPD_THK:              TXT_THK = Trim(Str(.Text))
            .Col = SPD_WID:              TXT_WID = Trim(Str(.Text))
            .Col = SPD_LEN:              TXT_LEN = Trim(Str(.Text))
            .Col = SPD_WGT:              TXT_WGT = Trim(Str(.Text))
            .Col = SPD_STANDARD:         TXT_SPEC = .Text
            .Col = SPD_CE_YN:            sCE_YN = .Text
            .Col = SPD_TS_YN:            sTS_YN = .Text
            .Col = SPD_JIS_YN:           jIS_YN = .Text
            .Col = SPD_UST_FL
             If .Text = "X" Then
                 TXT_Producer.Text = ""
             Else
                 TXT_Producer.Text = "UT " & .Text
             End If
             
            .Col = SS2_PRODSPECNOA:      PRODSPECNOA = .Value   '将牌号一的值保存到字段中
            .Col = SS2_PRODSPECNOB:      PRODSPECNOB = .Value   '将牌号二的值保存到字段中
            .Col = SS2_PRODSPECNOC:      PRODSPECNOC = .Value   '将牌号三的值保存到字段中
            .Col = SS2_INSP_CD:          sInsp_cd = .Text
            
            '在保存的时候，将喷印 冲印 侧喷的数据保存到相应变量中 2013-12-24
            .Col = SS2_sPaint:           sPaint = .Value '喷印
            .Col = SS2_sPunch:           sPunch = .Value '冲印
            .Col = SS2_sEdge:            sEdge = .Value '侧喷
            
            .Col = SPD_HTM_METH
             If .Text <> "" Then
                 TXT_Producer.Text = TXT_Producer.Text & " " & .Text
             End If
             
'             .Col = SS2_SURFACE_REQUESTS
            .Col = SS2_SURFACE_REQUESTS:     sSURFACE = .Text
            .Col = SS2_CLASS_CD:             sClasscd = .Text
            .Col = SS2_CLASS_LVL:            sClasslvl = .Text
            
            If sSURFACE <> "" And sClasscd <> "" Then
               TXT_Producer.Text = TXT_Producer.Text & "    " & sClasslvl
            ElseIf sSURFACE = "" And sClasscd <> "" Then
               TXT_Producer.Text = TXT_Producer.Text & "    " & sClasslvl
            ElseIf sSURFACE <> "" And sClasscd = "" Then
               TXT_Producer.Text = TXT_Producer.Text & "    " & sSURFACE
            End If
            
'             If .Text = "G" Then
'                 TXT_Producer.Text = TXT_Producer.Text & "    " & .Text
'             End If
            
            .Col = SPD_STLGRD:           TXT_STLGRD = .Text
            .Col = SPD_APLY_STDSPEC:     sSpec = .Text
            .Col = SPD_VESSEL_NO:        TXT_X1 = .Text
            .Col = SS2_SIDEMARK:         sideMark = .Text
            .Col = SPD_CUST_PUNCH:       TXT_Punch = .Text
             If jIS_YN = "D" Or jIS_YN = "E" Or jIS_YN = "F" Then
                TXT_Punch.Text = "NISCO" & " " & TXT_Punch.Text
             End If
            .Col = SPD_CUST_CD:          TXT_CUST_CD = .Text
            .Col = SPD_TO_CUR_INV:       TXT_TO_CUR_INV = .Text
        End With
        
        Ship_Emblem = 1
        
        '判断船级社数量
        If PRODSPECNOA <> 140 Then
            Ship_Emblem = Ship_Emblem + 1
        End If
        If PRODSPECNOB <> 140 Then
            Ship_Emblem = Ship_Emblem + 1
        End If
        If PRODSPECNOC <> 140 Then
            Ship_Emblem = Ship_Emblem + 1
        End If
        '判断船级社数量
        
        If Val(TXT_THK.Text) <= iF_mm Then
            opt_line8 = True
        ElseIf Val(TXT_THK.Text) <= iT_mm Then
            opt_line9 = True
        Else
            opt_line10 = True
        End If
        
    End If
    
    TXT_Special = TXT_SPEC
    
    Nisco = "NISCO"
    sFlag = "X"
    sNull = " "
    
    sPlate_No = TXT_MAT_NO
    sThk = TXT_THK
    sWid = TXT_WID
    sLen = TXT_LEN
    
    sNum = InStr(sSpec, "-")
    If sNum = 0 Then
        sNum = Len(sSpec)
    End If
    sSpec_Str = Mid(sSpec, 1, (sNum - 1))
    
    If sSpec = "ZH-ABS-A36" Then
       sSpec_Str = "ABS"
    End If

    Select Case sSpec_Str

           Case "ABS"                                 '美国船级社
                TXT_Class_comp = "127"
           Case "CCS"                                 '中国
                TXT_Class_comp = "128"
           Case "GL"                                  '德国
                TXT_Class_comp = "129"
           Case "BV"                                  '法国
                TXT_Class_comp = "130"
           Case "DNV"                                 '挪威
                TXT_Class_comp = "131"
           Case "VL"
                TXT_Class_comp = "131"
           Case "KR"                                  '韩国
                TXT_Class_comp = "132"
           Case "LR"                                  '英国
                TXT_Class_comp = "133"
           Case "RINA"                                '意大利
                TXT_Class_comp = "134"
           Case "NK"                                  '日本
                TXT_Class_comp = "135"
           Case "IRS"                                 '印度
                TXT_Class_comp = "139"
           Case "RS"                                 '俄罗斯
                TXT_Class_comp = "152"
           Case Else
           
                '2012-02-17 by liqian 暂时去掉,    杜萱后续还要恢复CE图标 12 的去喷
                '2012-08-02 by liqian 2012年图标可以使用
                                
                If sCE_YN = "Y" Then
                    TXT_Class_comp = "136"             'CE认证
                ElseIf sTS_YN = "Y" Then
                    TXT_Class_comp = "137"             'TS认证
                ElseIf sTS_YN = "A" Then '管线钢判断 add hanchao
                    TXT_Class_comp = "142"
                ElseIf jIS_YN = "A" Then 'JIS标识判断 add hanchao
                    TXT_Class_comp = "141"
                ElseIf jIS_YN = "D" Then 'API
                    TXT_Class_comp = "144"
                ElseIf jIS_YN = "E" Then 'API
                    TXT_Class_comp = "145"
                ElseIf jIS_YN = "F" Then 'API
                    TXT_Class_comp = "146"
                Else
                    TXT_Class_comp = "140"
'                    TXT_Special = ""
                    '加喷  2#线和1#线接口定义格式不同  by Liqian 2012-1-13
'                    TXT_Paint.Text = TXT_X1.Text
'                    TXT_X1.Text = ""
                End If

'                    TXT_Class_comp = "140"
                    TXT_Special = ""
    
    End Select
    
    If TXT_Class_comp.Text = "140" Then
       TXT_Class_comp.Text = sInsp_cd
    End If
    
End Sub

Private Sub Cmd_SEND()
    
    Dim SMESG As String
    
    '标印机接口文档编辑开始         '编号      '内容
    
    Dim Header As String * 2        '1         'MD
                                    '2         '空两格
    Dim sPlate_No As String * 16    '3         '钢板号
    Dim sLen As Long                '4         '长
    Dim sWid As Long                '5         '宽
    Dim sThk As Long                '6         '厚
    Dim sWgt As String                         '重
    
    Dim sEdge_Thk As Double                    '侧喷厚
    Dim sProducer As String * 50    '10        '画面输入(重量标识) '20150427 15
    
    Dim sStlgrd As String * 40      '12        '钢种
    Dim sStandard As String * 34    '11        '标准号
    Dim sInv As String * 3          '13        '仓库
    Dim sCust_Cd As String * 6      '14        '客户代码
    Dim sCust_Add_Cd As String                 '客户代码 + 流向
    Dim sProd_Date As String * 10   '15        '交货期
    Dim sGroup As String            '16        '班次
    Dim sComp As Integer            '17        '船徽
    Dim sX1 As String * 60          '18        '画面行1
    Dim sX2 As String * 60          '19        '画面行2
    Dim sX3 As String * 60          '20        '画面行3
    Dim sMark_disk As Long          '21        '冲印开始距离
    Dim sSpecial As String * 60                '加喷
    Dim Paint_line4 As String * 60  '22        '喷印行4
    Dim sMark_disk_ap As Integer    '23        '喷印到冲印距离
    Dim sCompression As Integer     '24        '压缩比
    Dim sPaint_font As Integer      '25        '是否加重字体
                                    '26        '重复喷印次数,默认为1,不重复
    Dim Repeat_Distance As Integer  '27        '重复喷印间隔距离
    Dim PaintStr_CD As Integer      '28        '喷印是否旋转
    Dim PunchStr_CD As Integer      '29        '冲印深度
    Dim Punch_line1 As String * 30  '30        '冲印行1
    Dim Punch_line2 As String * 30  '31        '冲印行2
    Dim Punch_line3 As String * 30  '32        '冲印行3
    Dim Punch_line4 As String * 30  '33        '冲印行4
    
    Dim UST_fl As String                       '是否探伤
    Dim sCert_type As String                   '质保书类型
    
    Dim sTopPaint As Integer
     
    Dim sSpec_Str As String
    
    Dim EdgeStr_CD As Integer
    Dim EU_Down_CD As Integer
   
    
    Dim sPunch_ori As Integer       '34        '冲印是否旋转
    Dim sEdge_hgt As Integer        '35        '侧喷高度
    
     
   
    '***  接口内容编辑结束  ***
   
    Dim sNum As Integer
    Dim sNumFL As String
    
    Dim sCEdate As String
    sCEdate = Format(Now, "YY")
    
    Dim sSpec_Logo As String
    
    Dim sEdgeStr As String * 80
    
    Dim sXm As String
    Dim sClasscd As String
    Dim sClasslvl As String
    
    Paint_line4 = ""                        '依照AGB3010P,默认空
    sMark_disk = 300                        '依照AGB3010P,默认是300
    sMark_disk_ap = 400                     '依照AGB3010P,默认是400
    Repeat_Distance = 6000                  '依照AGB3010P,默认是6000
    sPaint_font = 2                         '依照AGB3010P,默认2,加重字体
    PaintStr_CD = 2                         '依照AGB3010P,默认2,与长度方向平行

    '冲印深度
    If opt_line5 Then
       PunchStr_CD = 0
    End If
    If opt_line6 Then
       PunchStr_CD = 1
    End If
    If opt_line7 Then
       PunchStr_CD = 2
    End If
    
    sPunch_ori = 2                          '依照AGB3010P,默认2,与长度方向平行
    sCompression = 5                        '依照AGB3010P,默认5
    
    sSpecial = TXT_Special
    sX1 = TXT_X1                            '取画面值
    sX2 = TXT_X2                            '取画面值
    sX3 = TXT_X3                            '取画面值
    
    Header = "MD"
    
    sProd_Date = udt_date_fr.RawData        '交货期,画面取得
    sGroup = Trim(cbo_group.Text)           '班次,画面取得
    sComp = Val(TXT_Class_comp.Text)        '船徽,画面取得
    Repeat_Distance = Val(TXT_Distance)     '重复喷印间隔距离,画面取得
    
    '检查班次是否合法
    If sGroup <> "A" And sGroup <> "B" And sGroup <> "C" And sGroup <> "D" Then
        SMESG = " 班别错误，请确认是否正确输入班别"
        Call Gp_MsgBoxDisplay(SMESG)
        Exit Sub
    End If
    
    '检查船徽是否合法
    If sComp < 127 Or sComp > 256 Then
        SMESG = "Class_comp(127 ~ 256) 数据错误，请确认"
        Call Gp_MsgBoxDisplay(SMESG)
        Exit Sub
    End If
    
    '0 = Normal reading , 1 = Up-side down
    If opt_line11 Then
       EU_Down_CD = 0
    End If
    If opt_line12 Then
       EU_Down_CD = 1
    End If
    
    '
'    If chk_Cond(2) Then
'           sPaint = 1
'    Else
'           sPaint = 0
'    End If
'    If chk_Cond(3) Then
'           sPunch = 1
'    Else
'           sPunch = 0
'    End If
'    If chk_Cond(4) Then
'           sEdge = 1
'    Else
'           sEdge = 0
'    End If
    
    ss2.Row = ss2.ActiveRow
    
    ss2.Col = SPD_PLATE_NO:            sPlate_No = ss2.Text
    ss2.Col = SPD_THK:                 sThk = ss2.Value * 100 '按接口要求，厚度 * 100
    ss2.Col = SPD_WID:                 sWid = ss2.Value
    ss2.Col = SPD_LEN:                 sLen = ss2.Value
    ss2.Col = SPD_WGT:                 sWgt = ss2.Text
    ss2.Col = SPD_CUST_CD:             sCust_Cd = ss2.Text:     sCust_Add_Cd = ss2.Text
    ss2.Col = SPD_STLGRD:              sStlgrd = ss2.Text          '钢种
    ss2.Col = SPD_STANDARD:            sStandard = ss2.Text        '打印标准
    ss2.Col = SPD_TO_CUR_INV:          sInv = ss2.Text             '目的库
    
    ss2.Col = SPD_UST_FL:              UST_fl = ss2.Text           '探伤代码
    ss2.Col = SPD_CERT_TYPE:           sCert_type = ss2.Text       '质保书类型
    
    ss2.Col = SS2_SURFACE_REQUESTS:    sXm = ss2.Text
    ss2.Col = SS2_CLASS_CD:            sClasscd = ss2.Text
    ss2.Col = SS2_CLASS_LVL:           sClasslvl = ss2.Text
    
    '重量以小数点开始的 , 先在前面补0
    If sComp <> 136 And sComp <> 137 And sComp <> 141 And sComp <> 142 And (sComp <> 140 And sInsp_cd = 140) Then
       sStlgrd = sStandard
       sStandard = ""
    End If
    
    '侧喷高度
    If opt_line8 Then
       sEdge_hgt = 0
    End If
    If opt_line9 Then
       sEdge_hgt = 1
    End If
    If opt_line10 Then
       sEdge_hgt = 2
    End If
'    If sThk >= 4 * 100 And sThk <= 7.99 * 100 Then
'        sEdge_hgt = 1                            '3mm
'    ElseIf sThk >= 8 * 100 And sThk <= 15.99 * 100 Then
'        sEdge_hgt = 2                            '8mm
'    Else
'        sEdge_hgt = 4                            '12mm
'    End If
    
    '重量以小数点开始的 , 先在前面补0
    If Mid(sWgt, 1, 1) = "." Then
       sWgt = "0" & sWgt
    End If
    
     '有重量标识要求的编辑重量信息
     '标准年号为以下三种,标印重量加  T.W. 重量 t
     '画面取得,TXT_Producer
    If Trim(sStandard) = "GB 713-2008" Or Trim(sStandard) = "GB 3531-2008" Or Trim(sStandard) = "GB 19189-2011" Or Trim(sStandard) = "GB 713-2014" Or Trim(sStandard) = "GB 3531-2014" Or Trim(sStandard) = "GB/T 713-2014" Or Trim(sStandard) = "GB/T 3531-2014" Or Trim(sStandard) = "GB/T 19189-2011" Then
        sProducer = "T.W. " & sWgt & "t" & " " & TXT_Producer
    Else
        sProducer = TXT_Producer
    End If

    '***  喷印行1,2,3,接口文档暂未说明  ****

    
    '冲印行1,依照AGB3010P,默认空
    Punch_line1 = TXT_Punch.Text
    
    '冲印行2,依照AGB3010P,默认空
    Punch_line2 = ""
    
    '冲印行3,依照AGB3010P,默认空
    Punch_line3 = ""
    
    '冲印行4,依照AGB3010P,默认空
    Punch_line4 = ""
            
    ss2.Col = SPD_CE_YN
'    If ss2.Text = "Y" Then
'        sStandard = sStandard & " " & "CE 0038/" & sCEdate
'    End If
''   2012-02-13 liqian CE认证有图标显示,不需加喷
    '2012-02-17 by liqian 暂时去掉,    杜萱后续还要恢复CE图标 12 的去喷---  暂时为第2行加喷
 
    ' 2012-02-10 modified by liqian  CE认证显示
    
     '2012-08-02 by liqian 2012年图标可以使用
'     If ss2.Text = "Y" Then
'        sX2 = "CE 0038/" & sCEdate & sX2
'     End If
    
    
    If chk_Cond(8) = 1 Then
    
        sEdge_Thk = sThk / 100
            
'        If chk_Cond(9) = 1 Then
           'sEdgeStr = Trim(sPlate_No) & " " & sEdge_Thk & "X" & sWid & "X" & sLen & " " & Trim(sStlgrd) & "  " & sideMark
           sEdgeStr = Trim(sPlate_No) & " " & Trim(sStlgrd) & "  " & sEdge_Thk & "X" & sWid & "X" & sLen & " " & sideMark
'        Else
'           sEdgeStr = Trim(sPlate_No) & " " & sEdge_Thk & "X" & sWid & "X" & sLen & " " & sStlgrd
'        End If
        
        sEdgeStr = Trim(sEdgeStr) & " " & sCust_Add_Cd & " " & sInv
        
'        If sXm = "G" Then
'           sEdgeStr = sXm & "    " & Trim(sEdgeStr)
'        Else
'           sEdgeStr = Trim(sEdgeStr)
'        End If
        
        If sXm <> "" And sClasscd <> "" Then
           sEdgeStr = Trim(sEdgeStr) & "   " & sClasslvl
        ElseIf sXm = "" And sClasscd <> "" Then
           sEdgeStr = Trim(sEdgeStr) & "   " & sClasslvl
        ElseIf sXm <> "" And sClasscd = "" Then
           sEdgeStr = Trim(sEdgeStr) & "   " & sXm
        Else
           sEdgeStr = Trim(sEdgeStr)
        End If
        
'        sEdgeStr = Trim(sEdgeStr)
        
        Winsock2.SendData sEdgeStr
        
    End If
    
      '多船级社条件判断
    If Ship_Emblem = 1 Then
        Ship_Emblem = 0
        First_Number = 0
        Second_Number = 0
        Third_Number = 0
        Firth_Number = 0
        Fifth_Number = 0
        Sixth_Number = 0
    ElseIf Ship_Emblem = 2 Then
        First_Number = sComp
        Second_Number = PRODSPECNOA
        Third_Number = 0
        Firth_Number = 0
        Fifth_Number = 0
        Sixth_Number = 0
        sComp = 0
    ElseIf Ship_Emblem = 3 Then
        First_Number = sComp
        Second_Number = PRODSPECNOA
        Third_Number = PRODSPECNOB
        Firth_Number = 0
        Fifth_Number = 0
        Sixth_Number = 0
        sComp = 0
     ElseIf Ship_Emblem = 4 Then
        First_Number = sComp
        Second_Number = PRODSPECNOA
        Third_Number = PRODSPECNOB
        Firth_Number = PRODSPECNOC
        Fifth_Number = 0
        Sixth_Number = 0
        sComp = 0
    End If
    
    If chk_Cond(0) = 1 Then
   
        Winsock1.SendData Header & "  "
        Winsock1.SendData Chr(16) & Chr(14) & sPlate_No
        Winsock1.SendData HLByte(sLen, 3)
        Winsock1.SendData HLByte(sLen, 2)
        Winsock1.SendData HLByte(sLen, 1)
        Winsock1.SendData HLByte(sLen, 0)
        Winsock1.SendData HLByte(sWid, 3)
        Winsock1.SendData HLByte(sWid, 2)
        Winsock1.SendData HLByte(sWid, 1)
        Winsock1.SendData HLByte(sWid, 0)
        Winsock1.SendData HLByte(sThk, 3)
        Winsock1.SendData HLByte(sThk, 2)
        Winsock1.SendData HLByte(sThk, 1)
        Winsock1.SendData HLByte(sThk, 0)
        Winsock1.SendData HiByte(sPaint)
        Winsock1.SendData LoByte(sPaint)
        Winsock1.SendData HiByte(sPunch)
        Winsock1.SendData LoByte(sPunch)
        Winsock1.SendData HiByte(sEdge)
        Winsock1.SendData LoByte(sEdge)
        
        Winsock1.SendData Chr(50) & Chr(Len(Trim(sProducer))) & sProducer _
                        & Chr(20) & Chr(Len(Trim(sStlgrd))) & sStlgrd _
                        & Chr(34) & Chr(Len(Trim(sStandard))) & sStandard _
                        & Chr(3) & Chr(Len(Trim(sInv))) & sInv _
                        & Chr(6) & Chr(Len(Trim(sCust_Cd))) & sCust_Cd _
                        & Chr(10) & Chr(Len(Trim(sProd_Date))) & sProd_Date _
                        & sGroup
        Winsock1.SendData HiByte(sComp)
        Winsock1.SendData LoByte(sComp)
        Winsock1.SendData Chr(60) & Chr(Len(Trim(sX1))) & sX1 _
                        & Chr(60) & Chr(Len(Trim(sX2))) & sX2 _
                        & Chr(60) & Chr(Len(Trim(sX3))) & sX3
        Winsock1.SendData HLByte(sMark_disk, 3)
        Winsock1.SendData HLByte(sMark_disk, 2)
        Winsock1.SendData HLByte(sMark_disk, 1)
        Winsock1.SendData HLByte(sMark_disk, 0)
        Winsock1.SendData Chr(60) & Chr(Len(Trim(Paint_line4))) & Paint_line4
        Winsock1.SendData HiByte(sMark_disk_ap)
        Winsock1.SendData LoByte(sMark_disk_ap)
        Winsock1.SendData HiByte(sCompression)
        Winsock1.SendData LoByte(sCompression)
        Winsock1.SendData HiByte(sPaint_font)
        Winsock1.SendData LoByte(sPaint_font)
        Winsock1.SendData HiByte(1)
        Winsock1.SendData LoByte(1)
        Winsock1.SendData HiByte(Repeat_Distance)
        Winsock1.SendData LoByte(Repeat_Distance)
        Winsock1.SendData HiByte(PaintStr_CD)
        Winsock1.SendData LoByte(PaintStr_CD)
        Winsock1.SendData HiByte(PunchStr_CD)
        Winsock1.SendData LoByte(PunchStr_CD)
        Winsock1.SendData Chr(30) & Chr(Len(Trim(Punch_line1))) & Punch_line1 _
                        & Chr(30) & Chr(Len(Trim(Punch_line2))) & Punch_line2 _
                        & Chr(30) & Chr(Len(Trim(Punch_line3))) & Punch_line3 _
                        & Chr(30) & Chr(Len(Trim(Punch_line4))) & Punch_line4
        Winsock1.SendData HiByte(sPunch_ori)
        Winsock1.SendData LoByte(sPunch_ori)
        Winsock1.SendData HiByte(sEdge_hgt)
        Winsock1.SendData LoByte(sEdge_hgt)
         Winsock1.SendData HiByte(Ship_Emblem) '多少个船徽
        Winsock1.SendData LoByte(Ship_Emblem) '多少个船徽
        Winsock1.SendData HiByte(First_Number) '第一个代号
        Winsock1.SendData LoByte(First_Number) '第一个代号
        Winsock1.SendData HiByte(Second_Number) '第二个
        Winsock1.SendData LoByte(Second_Number) '第二个
        Winsock1.SendData HiByte(Third_Number) '第三个
        Winsock1.SendData LoByte(Third_Number) '第三个
        Winsock1.SendData HiByte(Firth_Number) '第四个
        Winsock1.SendData LoByte(Firth_Number) '第四个
        Winsock1.SendData HiByte(Fifth_Number) '第五个
        Winsock1.SendData LoByte(Fifth_Number) '第五个
        Winsock1.SendData HiByte(Sixth_Number) '第六个
        Winsock1.SendData LoByte(Sixth_Number) '第六个
    
    End If
      
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
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
                Case 9
                    strState = "连接错误"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
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
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(8).ForeColor = &HFF&
                Case 9
                    strState2 = "连接错误"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(8).ForeColor = &HFF&
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

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim iCol As Long
    Dim iRow As Long
    Dim iMode As Integer
    
    Dim iRowNum As Long
    Dim iRowfr As Long
    Dim iRowto As Long
    
    iCol = Col
    iRow = Row
    iMode = Mode

    If Row <= 0 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") And Col > SS1_LINE2 Then
    
         iRowto = iRow - 1
         iRowfr = iRow + 1
    
        If Col = SS1_THK Or Col = SS1_WID Or Col = SS1_LEN Then
            If Mode = 1 Then
               ss1.Col = iCol
               ss1.Row = iRow
               ss1.Text = 0
            End If
        End If
    
        Call Gp_Sp_UpdateMake(Proc_Sc("SC1")("Spread"), iMode)
        
        ss1.Row = iRow  'ss1.ActiveRow
        ss1.Col = SS1_EMP_CD
        ss1.Text = CBO_sUserID.Text
        
        ss1.Col = SS1_LABEL
        If chk_Cond(1) Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        
        ss1.Col = SS1_PAINT
        If chk_Cond(0) Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If

        ss1.Col = SS1_STAMP_YN
        If chk_Cond(3) Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        ss1.Col = SS1_BAR_YN
        If chk_Cond(4) Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        
        Call Cmd_SEND_SET(iRow)
        
    End If

End Sub
Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim iCol As Long
    Dim iRow As Long
    Dim iMode As Integer
    
    Dim iRowNum As Long
    Dim iRowfr As Long
    Dim iRowto As Long
    
    iCol = Col
    iRow = Row
    iMode = Mode

    If Row <= 0 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") And Col > SPD_LINE2 Then
    
         iRowto = iRow - 1
         iRowfr = iRow + 1
        
        If iRowto > 0 Then
            For iRowNum = 1 To iRowto
                 
                 ss2.Col = 0
                 ss2.Row = iRowNum
                 If ss2.Text <> "" Then
                    ss2.Text = ""
                    ss2.Col = SPD_LINE1
                    ss2.Value = 0
                    ss2.Col = SPD_LINE2
                    ss2.Value = 0
                    Exit For
                 End If
            Next iRowNum
        End If
        
        If iRowfr < ss2.MaxRows Then
            For iRowNum = iRowfr To ss2.MaxRows
                 
                 ss2.Col = 0
                 ss2.Row = iRowNum
                 If ss2.Text <> "" Then
                    ss2.Text = ""
                    ss2.Col = SPD_LINE1
                    ss2.Value = 0
                    ss2.Col = SPD_LINE2
                    ss2.Value = 0
                    Exit For
                 End If
            Next iRowNum
        End If
    
        If Col = SPD_THK Or Col = SPD_WID Or Col = SPD_LEN Then
            If Mode = 1 Then
               ss2.Col = iCol
               ss2.Row = iRow
               ss2.Text = 0
            End If
        End If
    
        Call Gp_Sp_UpdateMake(Proc_Sc("SC2")("Spread"), iMode)
        
        ss2.Row = iRow  'ss2.ActiveRow
        ss2.Col = SPD_EMP_CD
        ss2.Text = CBO_sUserID.Text
        
        ss2.Col = SPD_LABEL
        If chk_Cond(1) Then
           ss2.Value = 1
        Else
           ss2.Value = 0
        End If
        
'将原有的根据控件修改标印的功能删除 2013-12-31
'        ss2.Col = SPD_PAINT
'        If chk_Cond(0) Then
'           ss2.Value = 1
'        Else
'           ss2.Value = 0
'        End If
'
'        ss2.Col = SPD_STAMP_YN
'        If chk_Cond(3) Then
'           ss2.Value = 1
'        Else
'           ss2.Value = 0
'        End If
'        ss2.Col = SPD_BAR_YN
'        If chk_Cond(4) Then
'           ss2.Value = 1
'        Else
'           ss2.Value = 0
'        End If
        
        Call Cmd_SEND_SET(iRow)
        
    End If

End Sub


Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
'    iSS = ""

End Sub
Private Sub ss2_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
'    iSS = ""

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ss1.MaxRows > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub
Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ss2.MaxRows > 0 Then
        Set Active_Spread = Me.ss2
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
    
    DD.sQuery = "SELECT CD_SHORT_NAME ""标准代号"", CD_NAME ""标准中文名"" FROM ZP_CD WHERE CD_MANA_NO = 'G0035'"
    
    Call Gf_DD_Display(Conn, DD.sQuery, False)
    
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function


Private Sub txt_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        txt_stdspec.Text = ""
        DD.rControl.Add Item:=txt_stdspec

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

        Exit Sub

    End If
    
End Sub
Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
'        .Buttons(7).Enabled = False                  'Row Insert
'        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_USER_ComboAdd
'   2.Name         :
'   3.Input  Value : Conn Connection, Cbo Variant,sPRC String,
'                    {sFACT_CD,sPRC_LINE String, sADDNUM As Integer, ClsChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .25
'   7.Modify Date  :
'   8.Comment      : combo Add
'---------------------------------------------------------------------------------------
Public Function Gf_USER_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, sPgmId As String, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim sQuery As String

    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_USER_ComboAdd = False: Exit Function
    End If
    
    sQuery = "SELECT EMP_ID FROM ZP_AUTHORITY  "
    sQuery = sQuery + "    WHERE PGMID = '" & sPgmId & "'"
    sQuery = sQuery + "      AND UPD   = '1' AND EMP_ID <> '1JS6001'"
    sQuery = sQuery + "    ORDER BY EMP_ID"

    If ClsChk Then
        Cbo.Clear
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
                Cbo.AddItem AdoRs.Fields(0)
            End If
            AdoRs.MoveNext
            
        Wend
        Gf_USER_ComboAdd = True
    Else
        Gf_USER_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    Gf_USER_ComboAdd = False

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Send_Paint
'   2.Name         :
'   3.Input  Value :
'   4.Return Value : Boolean
'   5.Writer       : Li Qian
'   6.Create Date  : 2011. 08 .24
'   7.Modify Date  :
'   8.Comment      : 是否标识记录发送
'---------------------------------------------------------------------------------------

Public Function Gp_Send_Paint(sPlate_No As String, sMark As Integer, sStamp As Integer, sBar As Integer, sEmpCD As String) As String

'On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iRow As Integer
    
    Dim sPaint_in As String
    Dim sPunch_in As String
    Dim sEdge_in As String
    
    sPaint_in = Trim(Str(sMark))
    sPunch_in = Trim(Str(sStamp))
    sEdge_in = Trim(Str(sBar))

    Dim adoCmd As ADODB.Command

    Screen.MousePointer = vbHourglass

    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256

   sQuery = "{call AGC2039P ('" + sPlate_No + "','" + sPaint_in + "','" + sPunch_in + "','" + sEdge_in + "','" + sEmpCD + "',?)}"

    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command

    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1

    adoCmd.CommandText = sQuery

    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))

    adoCmd.Execute , , adExecuteNoRecords

'    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")

        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg

        Screen.MousePointer = vbDefault
'        Gp_Send_Size_Exec = sErrMessg
        Set adoCmd = Nothing
        Exit Function

    End If

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
'    Gp_Send_Size_Exec = ""
    Exit Function

'Process_Exec_ERROR:
'
'    Set adoCmd = Nothing
'    Screen.MousePointer = vbDefault
'    Gp_Send_Size_Exec = "Process_Exec_ERROR"
'    Err.Raise Err.Number, Err.Description & sQuery

End Function

Public Function Gp_Send_Size_Exec(sPlate_No As String) As String

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iRow As Integer

    Dim adoCmd As ADODB.Command

    Screen.MousePointer = vbHourglass

    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256

'    sQuery = "{call AKG2050P ('" + "C1" + "','" + "1" + "','" + Trim(txt_target.Text) + "',?)}"
    sQuery = "{call AGC2033P ('" + sPlate_No + "',?)}"

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

        Screen.MousePointer = vbDefault
        Gp_Send_Size_Exec = sErrMessg
        Set adoCmd = Nothing
        Exit Function

    End If

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Gp_Send_Size_Exec = ""
    Exit Function

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Gp_Send_Size_Exec = "Process_Exec_ERROR"
    Err.Raise Err.Number, Err.Description & sQuery

End Function


