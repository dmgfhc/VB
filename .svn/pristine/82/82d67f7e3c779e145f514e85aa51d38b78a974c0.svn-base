VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGD2081C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "钢板剪切、表面检查实绩查询及修改_CGD2081C"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_Color_name 
      Enabled         =   0   'False
      Height          =   300
      Left            =   13580
      Locked          =   -1  'True
      TabIndex        =   91
      Top             =   3510
      Width           =   810
   End
   Begin VB.TextBox txt_Color_code 
      Height          =   300
      Left            =   13150
      MaxLength       =   2
      TabIndex        =   90
      Tag             =   "原因"
      Top             =   3510
      Width           =   405
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
      Height          =   315
      Left            =   9525
      MaxLength       =   14
      TabIndex        =   68
      Tag             =   "物料号"
      Top             =   870
      Width           =   1725
   End
   Begin VB.TextBox TXT_sUserID_Tail 
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
      Left            =   5220
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   67
      Top             =   870
      Width           =   1095
   End
   Begin VB.TextBox TXT_sUserID 
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
      Left            =   1530
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   66
      Top             =   870
      Width           =   1095
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
      Left            =   5220
      TabIndex        =   5
      Top             =   480
      Width           =   2865
   End
   Begin InDate.UDate udt_date_fr 
      Height          =   315
      Left            =   5220
      TabIndex        =   6
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
      Left            =   6660
      TabIndex        =   7
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   3900
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
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
      Left            =   3900
      Top             =   480
      Width           =   1305
      _ExtentX        =   2302
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
   Begin VB.TextBox TXT_INSP_FLAW 
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
      Index           =   2
      Left            =   19410
      MaxLength       =   3
      TabIndex        =   21
      Top             =   2220
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TXT_INSP_FLAW_NAME 
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
      Left            =   16800
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2220
      Visible         =   0   'False
      Width           =   2595
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
      Left            =   3570
      MaxLength       =   1
      TabIndex        =   16
      Tag             =   "CD_MANA_NO"
      Text            =   "1"
      Top             =   480
      Visible         =   0   'False
      Width           =   330
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
      Left            =   11340
      MaxLength       =   1
      TabIndex        =   11
      Tag             =   "CD_MANA_NO"
      Text            =   "1"
      Top             =   90
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   15390
      Top             =   120
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
      ItemData        =   "CGD2081C.frx":0000
      Left            =   9525
      List            =   "CGD2081C.frx":000D
      TabIndex        =   4
      Top             =   90
      Width           =   705
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
      Left            =   17310
      TabIndex        =   3
      Tag             =   "机号"
      Top             =   1140
      Visible         =   0   'False
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
      Left            =   16890
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "生产工厂"
      Top             =   1140
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt_plate_no 
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
      Left            =   9525
      MaxLength       =   14
      TabIndex        =   1
      Tag             =   "物料号"
      Top             =   480
      Width           =   1725
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
      ItemData        =   "CGD2081C.frx":001D
      Left            =   10230
      List            =   "CGD2081C.frx":002D
      TabIndex        =   0
      Top             =   90
      Width           =   705
   End
   Begin InDate.ULabel ULabel19 
      Height          =   315
      Left            =   8190
      Top             =   870
      Width           =   1305
      _ExtentX        =   2302
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
      Left            =   8190
      Top             =   480
      Width           =   1305
      _ExtentX        =   2302
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
   Begin SSSplitter.SSSplitter SSSp1 
      Height          =   7905
      Left            =   90
      TabIndex        =   8
      Top             =   1260
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   13944
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      Locked          =   -1  'True
      PaneTree        =   "CGD2081C.frx":0041
      Begin Threed.SSPanel Winsock 
         Height          =   3450
         Left            =   0
         TabIndex        =   9
         Tag             =   "172.18.151.145"
         Top             =   0
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   6085
         _Version        =   196609
         BackColor       =   12632319
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox TXT_FLAW 
            Height          =   270
            Left            =   11340
            TabIndex        =   103
            Top             =   2370
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CheckBox CHK_FLAW_YN 
            BackColor       =   &H00E0E0E0&
            Caption         =   "下表是否检验"
            Height          =   240
            Left            =   9840
            TabIndex        =   102
            Tag             =   "G"
            Top             =   3090
            Width           =   1620
         End
         Begin VB.TextBox TXT_WAVE1 
            Alignment       =   1  'Right Justify
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
            Left            =   8550
            MaxLength       =   2
            TabIndex        =   87
            Top             =   120
            Width           =   885
         End
         Begin VB.TextBox TXT_LOC 
            Alignment       =   1  'Right Justify
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
            Left            =   13050
            MaxLength       =   7
            TabIndex        =   85
            Top             =   1890
            Width           =   1275
         End
         Begin VB.TextBox TXT_CL 
            Height          =   270
            Left            =   11340
            TabIndex        =   84
            Text            =   "Text1"
            Top             =   1800
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CheckBox CHK_CL_FL 
            BackColor       =   &H00E0E0E0&
            Caption         =   "矫直指示"
            Height          =   315
            Left            =   10230
            TabIndex        =   83
            Top             =   1750
            Width           =   1095
         End
         Begin VB.CheckBox CHK_BOT_GRD 
            BackColor       =   &H00C0C0C0&
            Caption         =   "合格"
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   8790
            TabIndex        =   75
            Tag             =   "Y"
            Top             =   1980
            Width           =   735
         End
         Begin VB.TextBox TXT_BOT_GRID_GRD 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6315
            MaxLength       =   1
            TabIndex        =   74
            Text            =   " "
            Top             =   1965
            Width           =   690
         End
         Begin VB.TextBox TXT_TOP_GRID_GRD 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6315
            MaxLength       =   1
            TabIndex        =   73
            Text            =   " "
            Top             =   1590
            Width           =   690
         End
         Begin VB.CheckBox SSCHK_GRID_YN 
            BackColor       =   &H00C0C0C0&
            Caption         =   "是否修磨"
            Height          =   240
            Left            =   5190
            TabIndex        =   72
            Tag             =   "G"
            Top             =   1290
            Width           =   1110
         End
         Begin VB.CheckBox CHK_TOP_GRD 
            BackColor       =   &H00C0C0C0&
            Caption         =   "不合格"
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   8790
            TabIndex        =   71
            Tag             =   "N"
            Top             =   1665
            Width           =   900
         End
         Begin VB.CheckBox CHK_TOP_GRD 
            BackColor       =   &H00C0C0C0&
            Caption         =   "合格"
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   8790
            TabIndex        =   70
            Tag             =   "Y"
            Top             =   1410
            Width           =   735
         End
         Begin VB.CheckBox CHK_BOT_GRD 
            BackColor       =   &H00C0C0C0&
            Caption         =   "不合格"
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   8790
            TabIndex        =   69
            Tag             =   "N"
            Top             =   2250
            Width           =   900
         End
         Begin VB.TextBox TXT_SIZE_KND 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   65
            Tag             =   "表面等级判定"
            Top             =   840
            Width           =   885
         End
         Begin VB.TextBox TXT_VERT_DEG 
            Alignment       =   1  'Right Justify
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
            Left            =   6360
            MaxLength       =   2
            TabIndex        =   64
            Top             =   480
            Width           =   885
         End
         Begin VB.TextBox TXT_TRIM_FL 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8550
            Locked          =   -1  'True
            TabIndex        =   63
            Tag             =   "表面等级判定"
            Top             =   840
            Width           =   885
         End
         Begin VB.TextBox TXT_RECT_DEG 
            Alignment       =   1  'Right Justify
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
            Left            =   8550
            MaxLength       =   2
            TabIndex        =   62
            Top             =   480
            Width           =   885
         End
         Begin VB.TextBox TXT_WAVE 
            Alignment       =   1  'Right Justify
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
            Left            =   6360
            MaxLength       =   2
            TabIndex        =   58
            Top             =   120
            Width           =   885
         End
         Begin VB.TextBox TXT_REMARK 
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
            Left            =   1440
            MaxLength       =   400
            MultiLine       =   -1  'True
            TabIndex        =   56
            Tag             =   "后道工序"
            Top             =   3030
            Width           =   8235
         End
         Begin VB.TextBox TXT_INSP_FLAW 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   5
            Left            =   13455
            MaxLength       =   3
            TabIndex        =   47
            Top             =   450
            Width           =   885
         End
         Begin VB.TextBox TXT_INSP_FLAW_NAME 
            Height          =   315
            Index           =   5
            Left            =   11385
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   450
            Width           =   2055
         End
         Begin VB.TextBox TXT_INSP_MAIN_GRD 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11025
            Locked          =   -1  'True
            TabIndex        =   42
            Tag             =   "表面等级判定"
            Top             =   825
            Width           =   885
         End
         Begin VB.TextBox TXT_INSP_FLAW 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   1
            Left            =   3870
            MaxLength       =   3
            TabIndex        =   30
            Top             =   5250
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox TXT_INSP_FLAW 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   4
            Left            =   3870
            MaxLength       =   3
            TabIndex        =   29
            Top             =   6315
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox TXT_INSP_FLAW 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   3
            Left            =   14250
            MaxLength       =   3
            TabIndex        =   28
            Top             =   3030
            Width           =   705
         End
         Begin VB.TextBox TXT_INSP_FLAW 
            Alignment       =   2  'Center
            Height          =   315
            Index           =   0
            Left            =   14250
            MaxLength       =   3
            TabIndex        =   27
            Top             =   2670
            Width           =   705
         End
         Begin VB.TextBox TXT_INSP_FLAW_NAME 
            Height          =   315
            Index           =   3
            Left            =   12780
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   3030
            Width           =   1455
         End
         Begin VB.TextBox TXT_INSP_FLAW_NAME 
            Height          =   315
            Index           =   4
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   6315
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.TextBox TXT_INSP_FLAW_NAME 
            Height          =   315
            Index           =   1
            Left            =   1470
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   5250
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.TextBox TXT_INSP_FLAW_NAME 
            Height          =   315
            Index           =   0
            Left            =   12780
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   2670
            Width           =   1455
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
            Left            =   11385
            TabIndex        =   15
            Top             =   120
            Width           =   2955
         End
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   10050
            Top             =   120
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "改判标准号"
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
            Left            =   270
            Top             =   4890
            Visible         =   0   'False
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "主要缺陷"
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
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   270
            Top             =   5250
            Visible         =   0   'False
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "小缺陷1"
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
         Begin InDate.ULabel ULabel26 
            Height          =   315
            Index           =   0
            Left            =   1470
            Top             =   4560
            Visible         =   0   'False
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            Caption         =   "下表面缺陷名称      /  代码"
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
         Begin InDate.ULabel ULabel26 
            Height          =   315
            Index           =   1
            Left            =   1470
            Top             =   5625
            Visible         =   0   'False
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            Caption         =   "上表面缺陷名称      /  代码"
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
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   270
            Top             =   5955
            Visible         =   0   'False
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "主要缺陷"
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
            Left            =   270
            Top             =   6315
            Visible         =   0   'False
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "小缺陷1"
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
         Begin InDate.ULabel ULabel28 
            Height          =   315
            Left            =   2340
            Top             =   120
            Width           =   1050
            _ExtentX        =   1852
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
         Begin InDate.ULabel ULabel29 
            Height          =   315
            Left            =   1440
            Top             =   120
            Width           =   870
            _ExtentX        =   1535
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
         Begin InDate.ULabel ULabel30 
            Height          =   315
            Left            =   3420
            Top             =   120
            Width           =   1125
            _ExtentX        =   1984
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
         Begin CSTextLibCtl.sidbEdit SDB_INSP_LEN_MX 
            Height          =   315
            Left            =   3420
            TabIndex        =   31
            Top             =   1545
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
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
            NumIntDigits    =   8
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_INSP_WID_MN 
            Height          =   315
            Left            =   2340
            TabIndex        =   32
            Top             =   1875
            Width           =   1050
            _Version        =   262145
            _ExtentX        =   1852
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_INSP_THK_MN 
            Height          =   315
            Left            =   1440
            TabIndex        =   33
            Top             =   1875
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_INSP_LEN_MN 
            Height          =   315
            Left            =   3420
            TabIndex        =   34
            Top             =   1875
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
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
            NumIntDigits    =   8
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_WID 
            Height          =   315
            Left            =   2340
            TabIndex        =   35
            Top             =   450
            Width           =   1050
            _Version        =   262145
            _ExtentX        =   1852
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
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_THK 
            Height          =   315
            Left            =   1440
            TabIndex        =   36
            Top             =   450
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
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
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_LEN 
            Height          =   315
            Left            =   3420
            TabIndex        =   37
            Top             =   450
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
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
            Modified        =   0   'False
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
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel38 
            Height          =   315
            Left            =   270
            Top             =   1875
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "下公差"
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
         Begin InDate.ULabel ULabel43 
            Height          =   315
            Left            =   270
            Top             =   450
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "公称"
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
         Begin CSTextLibCtl.sidbEdit SDB_INSP_THK_MX 
            Height          =   315
            Left            =   1440
            TabIndex        =   38
            Top             =   1545
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel37 
            Height          =   315
            Left            =   270
            Top             =   1545
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "上公差"
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
         Begin CSTextLibCtl.sidbEdit SDB_ORD_WID 
            Height          =   315
            Left            =   6930
            TabIndex        =   39
            Top             =   5775
            Width           =   1050
            _Version        =   262145
            _ExtentX        =   1852
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
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
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_ORD_THK 
            Height          =   315
            Left            =   6030
            TabIndex        =   40
            Top             =   5775
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
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
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_ORD_LEN 
            Height          =   315
            Left            =   8010
            TabIndex        =   41
            Top             =   5775
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
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
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel45 
            Height          =   315
            Left            =   4860
            Top             =   5775
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "订单"
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
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   0
            Left            =   10050
            Top             =   825
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   556
            Caption         =   "表面等级"
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
         Begin CSTextLibCtl.sidbEdit SDB_ACT_WID 
            Height          =   315
            Left            =   2340
            TabIndex        =   43
            Top             =   810
            Width           =   1050
            _Version        =   262145
            _ExtentX        =   1852
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
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_ACT_THK 
            Height          =   315
            Left            =   1440
            TabIndex        =   44
            Top             =   840
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
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
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Left            =   270
            Top             =   810
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "实测"
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
         Begin CSTextLibCtl.sidbEdit SDB_INSP_WID_MX 
            Height          =   315
            Left            =   2340
            TabIndex        =   45
            Top             =   1545
            Width           =   1050
            _Version        =   262145
            _ExtentX        =   1852
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel21 
            Height          =   315
            Left            =   10050
            Top             =   450
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "改判原因"
            Alignment       =   1
            BackColor       =   16777088
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
         Begin Threed.SSFrame SSFrame4 
            Height          =   1035
            Left            =   11970
            TabIndex        =   48
            Top             =   810
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   1826
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_CHK_PRD_GRD 
               Height          =   255
               Index           =   0
               Left            =   210
               TabIndex        =   49
               Top             =   60
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "正品"
            End
            Begin Threed.SSOption opt_CHK_PRD_GRD 
               Height          =   255
               Index           =   1
               Left            =   210
               TabIndex        =   50
               Top             =   390
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "改判"
            End
            Begin Threed.SSOption opt_CHK_PRD_GRD 
               Height          =   255
               Index           =   2
               Left            =   210
               TabIndex        =   51
               Top             =   720
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "协议"
            End
            Begin Threed.SSOption opt_CHK_PRD_GRD 
               Height          =   255
               Index           =   3
               Left            =   1380
               TabIndex        =   52
               Top             =   60
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "待判"
            End
            Begin Threed.SSOption opt_CHK_PRD_GRD 
               Height          =   255
               Index           =   4
               Left            =   1380
               TabIndex        =   53
               Top             =   720
               Width           =   975
               _ExtentX        =   1720
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
               Caption         =   "次品"
            End
         End
         Begin Threed.SSCheck SSCHK_SIZE_KND 
            Height          =   315
            Left            =   5190
            TabIndex        =   54
            Top             =   840
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   196609
            Font3D          =   2
            BackColor       =   14804173
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   " 定 尺"
         End
         Begin Threed.SSCheck SSCHK_TRIM_FL 
            Height          =   315
            Left            =   7380
            TabIndex        =   55
            Top             =   840
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   196609
            Font3D          =   2
            BackColor       =   14804173
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   " 切 边"
         End
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   270
            Top             =   3030
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "备注"
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
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   1
            Left            =   5190
            Top             =   120
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "不平度(/m)"
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
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   2
            Left            =   5190
            Top             =   480
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "镰刀弯"
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
         Begin Threed.SSCheck SSCHK_LAST_YN 
            Height          =   315
            Left            =   10230
            TabIndex        =   59
            Top             =   1350
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   196609
            Font3D          =   2
            BackColor       =   14804173
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   " 尾 板"
         End
         Begin Threed.SSCheck SSCHK_LY_YN 
            Height          =   315
            Left            =   10230
            TabIndex        =   60
            Top             =   2160
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   196609
            Font3D          =   2
            BackColor       =   14804173
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "留 样"
         End
         Begin CSTextLibCtl.sidbEdit SDB_ACT_LEN 
            Height          =   315
            Left            =   3420
            TabIndex        =   61
            Top             =   810
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
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
            Modified        =   0   'False
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
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   3
            Left            =   7380
            Top             =   480
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "切斜"
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
         Begin InDate.ULabel ULabel18 
            Height          =   315
            Index           =   0
            Left            =   5190
            Top             =   1980
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "下表面"
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
         Begin InDate.ULabel ULabel14 
            Height          =   315
            Index           =   2
            Left            =   5190
            Top             =   1590
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "上表面"
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
         Begin InDate.ULabel ULabel14 
            Height          =   315
            Index           =   1
            Left            =   6315
            Top             =   1230
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            Caption         =   "判定/ 面积比%/ 深度"
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
         Begin CSTextLibCtl.sidbEdit SDB_TOP_GRID_DEEP 
            Height          =   315
            Left            =   7905
            TabIndex        =   76
            Top             =   1590
            Width           =   840
            _Version        =   262145
            _ExtentX        =   1482
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
            Enabled         =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   3
            ShowZero        =   0   'False
            MaxValue        =   9999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_TOP_GRID_YRD 
            Height          =   315
            Left            =   7035
            TabIndex        =   77
            Top             =   1590
            Width           =   840
            _Version        =   262145
            _ExtentX        =   1482
            _ExtentY        =   556
            _StockProps     =   125
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
            Enabled         =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   3
            ShowZero        =   0   'False
            MaxValue        =   999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_BOT_GRID_YRD 
            Height          =   315
            Left            =   7035
            TabIndex        =   78
            Top             =   1965
            Width           =   840
            _Version        =   262145
            _ExtentX        =   1482
            _ExtentY        =   556
            _StockProps     =   125
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
            Enabled         =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   3
            ShowZero        =   0   'False
            MaxValue        =   999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel15 
            Height          =   315
            Left            =   5190
            Top             =   2325
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Caption         =   "修磨时间"
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
         Begin CSTextLibCtl.sitxEdit TXT_GRID_TIME 
            Height          =   315
            Left            =   6315
            TabIndex        =   79
            Top             =   2310
            Width           =   2085
            _Version        =   262145
            _ExtentX        =   3678
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
            Text            =   "____-__-__ __:__:__"
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
         Begin CSTextLibCtl.sidbEdit SDB_BOT_GRID_DEEP 
            Height          =   315
            Left            =   7905
            TabIndex        =   80
            Top             =   1965
            Width           =   840
            _Version        =   262145
            _ExtentX        =   1482
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
            Enabled         =   0   'False
            BorderEffect    =   2
            DataProperty    =   2
            FocusSelect     =   -1  'True
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   3
            ShowZero        =   0   'False
            MaxValue        =   9999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   270
            Top             =   2250
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "对角线1"
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
         Begin CSTextLibCtl.sidbEdit SDB_INSP_DIAGONAL1 
            Height          =   315
            Left            =   1440
            TabIndex        =   81
            Top             =   2250
            Width           =   1440
            _Version        =   262145
            _ExtentX        =   2540
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
            Modified        =   0   'False
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
            NumIntDigits    =   8
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel26 
            Height          =   315
            Index           =   2
            Left            =   270
            Top             =   2640
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "对角线2"
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
         Begin CSTextLibCtl.sidbEdit SDB_INSP_DIAGONAL2 
            Height          =   315
            Left            =   1440
            TabIndex        =   82
            Top             =   2640
            Width           =   1440
            _Version        =   262145
            _ExtentX        =   2540
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
            Modified        =   0   'False
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
            NumIntDigits    =   8
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   4
            Left            =   11970
            Top             =   1890
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            Caption         =   "跺位号"
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
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   5
            Left            =   7380
            Top             =   120
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "不平度(/2m)"
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
            Left            =   11970
            Top             =   2250
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            Caption         =   "表面颜色"
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
         Begin CSTextLibCtl.sidbEdit SDB_MS_WID 
            Height          =   315
            Left            =   2340
            TabIndex        =   92
            Top             =   1170
            Width           =   1050
            _Version        =   262145
            _ExtentX        =   1852
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
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_MS_THK 
            Height          =   315
            Left            =   1440
            TabIndex        =   93
            Top             =   1170
            Width           =   870
            _Version        =   262145
            _ExtentX        =   1535
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
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel23 
            Height          =   315
            Left            =   270
            Top             =   1170
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            Caption         =   "设备"
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
         Begin CSTextLibCtl.sidbEdit SDB_MS_LEN 
            Height          =   315
            Left            =   3420
            TabIndex        =   94
            Top             =   1170
            Width           =   1125
            _Version        =   262145
            _ExtentX        =   1984
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
            Modified        =   0   'False
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
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel24 
            Height          =   315
            Left            =   2940
            Top             =   2670
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            Caption         =   "厚度1"
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
         Begin CSTextLibCtl.sidbEdit SDB_HD1 
            Height          =   315
            Left            =   3600
            TabIndex        =   95
            Top             =   2670
            Width           =   810
            _Version        =   262145
            _ExtentX        =   1429
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
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel27 
            Height          =   315
            Left            =   4440
            Top             =   2670
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   556
            Caption         =   "厚度2"
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
         Begin CSTextLibCtl.sidbEdit SDB_HD2 
            Height          =   315
            Left            =   5070
            TabIndex        =   96
            Top             =   2670
            Width           =   780
            _Version        =   262145
            _ExtentX        =   1376
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
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel31 
            Height          =   315
            Left            =   5880
            Top             =   2670
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   556
            Caption         =   "厚度3"
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
         Begin CSTextLibCtl.sidbEdit SDB_HD3 
            Height          =   315
            Left            =   6510
            TabIndex        =   97
            Top             =   2670
            Width           =   750
            _Version        =   262145
            _ExtentX        =   1323
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
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel33 
            Height          =   315
            Left            =   7290
            Top             =   2670
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   556
            Caption         =   "厚度4"
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
         Begin CSTextLibCtl.sidbEdit SDB_HD4 
            Height          =   315
            Left            =   7980
            TabIndex        =   98
            Top             =   2670
            Width           =   780
            _Version        =   262145
            _ExtentX        =   1376
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
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel34 
            Height          =   315
            Left            =   8790
            Top             =   2670
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   556
            Caption         =   "厚度5"
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
         Begin CSTextLibCtl.sidbEdit SDB_HD5 
            Height          =   315
            Left            =   9480
            TabIndex        =   99
            Top             =   2670
            Width           =   780
            _Version        =   262145
            _ExtentX        =   1376
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
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel35 
            Height          =   315
            Left            =   10290
            Top             =   2670
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   556
            Caption         =   "厚度6"
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
         Begin CSTextLibCtl.sidbEdit SDB_HD6 
            Height          =   315
            Left            =   10920
            TabIndex        =   100
            Top             =   2670
            Width           =   750
            _Version        =   262145
            _ExtentX        =   1323
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
            Modified        =   0   'False
            HideSelection   =   -1  'True
            RawData         =   "0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel36 
            Height          =   315
            Left            =   11970
            Top             =   2670
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            Caption         =   "上表缺陷"
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
         Begin InDate.ULabel ULabel39 
            Height          =   315
            Left            =   11970
            Top             =   3030
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            Caption         =   "下表缺陷"
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
         Height          =   4425
         Left            =   0
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   3480
         Width           =   15165
         _Version        =   393216
         _ExtentX        =   26749
         _ExtentY        =   7805
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
         MaxCols         =   86
         MaxRows         =   10
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CGD2081C.frx":0093
      End
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   0
      Left            =   15465
      Top             =   1140
      Visible         =   0   'False
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   8190
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
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
   Begin CSTextLibCtl.sitxEdit TXT_CUT_TIME 
      Height          =   315
      Left            =   15510
      TabIndex        =   10
      Tag             =   "出炉时间"
      Top             =   1680
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
   Begin Threed.SSFrame SSFrame3 
      Height          =   315
      Left            =   1545
      TabIndex        =   12
      Top             =   90
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   556
      _Version        =   196609
      BackColor       =   14737632
      Begin Threed.SSOption opt_line1 
         Height          =   255
         Left            =   60
         TabIndex        =   13
         Top             =   30
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
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
         Caption         =   "# 1"
         Value           =   -1
      End
      Begin Threed.SSOption opt_line2 
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   30
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
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
      Begin Threed.SSOption opt_line5 
         Height          =   255
         Left            =   1470
         TabIndex        =   86
         Top             =   30
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "全部"
      End
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   120
      Top             =   90
      Width           =   1395
      _ExtentX        =   2461
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
      ForeColor       =   0
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   315
      Left            =   1545
      TabIndex        =   17
      Top             =   480
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   196609
      BackColor       =   14737632
      Begin Threed.SSOption opt_line3 
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   30
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
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
         Caption         =   "计划"
         Value           =   -1
      End
      Begin Threed.SSOption opt_line4 
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   30
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         Enabled         =   0   'False
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   120
      Top             =   480
      Width           =   1395
      _ExtentX        =   2461
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel25 
      Height          =   315
      Left            =   15480
      Top             =   2220
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "改判缺陷"
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
   Begin Threed.SSPanel SSP2 
      Height          =   315
      Left            =   13275
      TabIndex        =   22
      Top             =   870
      Width           =   1980
      _ExtentX        =   3493
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
   Begin Threed.SSPanel SSP1 
      Height          =   315
      Left            =   11280
      TabIndex        =   57
      Top             =   870
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "一坯多订单"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   120
      Top             =   870
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      Caption         =   "检验工(头部)"
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
      Index           =   0
      Left            =   3900
      Top             =   870
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "检验工(尾部)"
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
   Begin Threed.SSPanel SSP5 
      Height          =   285
      Left            =   13290
      TabIndex        =   88
      Top             =   510
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   503
      _Version        =   196609
      ForeColor       =   0
      BackColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "出口订单"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSP4 
      Height          =   285
      Left            =   11280
      TabIndex        =   89
      Top             =   510
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   503
      _Version        =   196609
      ForeColor       =   0
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
      Caption         =   "定制配送"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
End
Attribute VB_Name = "CGD2081C"
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
'-- Program Name      钢板剪切、表面检查实绩查询及修改
'-- Program ID        CGD2081C
'-- Document No       Q-00-0010(Specification)
'-- Designer          杨猛
'-- Coder             杨猛
'-- Date              2011.02.13
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-- 1.01  20110213 杨猛         钢板剪切、表面检查实绩查询及修改
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

Dim sCheck  As String

Const SPD_LINE1 = 1
Const SPD_LINE2 = 2
Const SPD_PLATE_NO = 3
Const SPD_LOT_NO = 4
Const SPD_CUT_NO = 5
Const SPD_THK = 6
Const SPD_WID = 7
Const SPD_LEN = 8
Const SPD_WGT = 9
Const SPD_ACT_THK = 10
Const SPD_ACT_WID = 11
Const SPD_ACT_LEN = 12
Const SPD_LAST_YN = 13
Const SPD_SIZE_KND = 14
Const SPD_TRIM_FL = 15
Const SPD_UST_FL = 16
Const SPD_APLY_STDSPEC = 17
Const SPD_APLY_STDSPEC_NEW = 18
Const SPD_INSP_CD = 19
Const SPD_GRID_YN = 20
Const SPD_INSP_CD1 = 21
Const SPD_INSP_CD2 = 22
Const SPD_INSP_CD3 = 23
Const SPD_INSP_CD4 = 24
Const SPD_SURF_YN = 25
Const SPD_SURF_GRD = 27
Const SPD_SURF_GRD_DET = 28
Const SPD_PROD_DATE = 29
Const SPD_EMP_CD = 30
Const SPD_THK_MIN = 31
Const SPD_THK_MAX = 32
Const SPD_WID_MIN = 34
Const SPD_WID_MAX = 35
Const SPD_LEN_MIN = 36
Const SPD_LEN_MAX = 37
Const SPD_DEL_DATE_FR = 50
Const SPD_DEL_DATE_TO = 51
Const SPD_ORD_CNT = 52
Const SPD_ORD_REMARK = 53
Const SPD_PROD_REMARK = 54
Const SPD_INS_MAN = 55
Const SPD_INSP_WAVE = 56
Const SPD_INSP_VERT_DEG = 57
Const SPD_INSP_RECT_DEG = 58
Const SPD_INS_MAN_TAIL = 59
Const SPD_TOP_GRID_GRD = 60
Const SPD_TOP_GRID_YRD = 61
Const SPD_TOP_GRID_DEEP = 62
Const SPD_BOT_GRID_GRD = 63
Const SPD_BOT_GRID_YRD = 64
Const SPD_BOT_GRID_DEEP = 65
Const SPD_GRID_TIME = 66
Const SPD_INSP_DIAGONAL1 = 67
Const SPD_INSP_DIAGONAL2 = 68
Const SPD_CL_FL = 69
Const SPD_LOC = 70
Const SPD_INSP_WAVE1 = 71
Const SPD_FLAG_FL = 72
Const SPD_EXPORT = 73
Const SPD_PLATE_COLOR = 74
Const SPD_THK1 = 80
Const SPD_THK2 = 81
Const SPD_THK3 = 82
Const SPD_THK4 = 83
Const SPD_THK5 = 84
Const SPD_THK6 = 85
Const SPD_FLAW_YN = 86



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
    Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '同板差
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
   Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 58, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 59, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  ' Add by LiQian at 2012-08-31 尾部检验工
   Call Gp_Sp_Collection(ss1, 60, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 61, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 62, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 63, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 64, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 65, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 66, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 67, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 68, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 69, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 70, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 71, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 72, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 73, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 74, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 75, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '坯料类别
   Call Gp_Sp_Collection(ss1, 76, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '坯料类别
   Call Gp_Sp_Collection(ss1, 77, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '坯料类别
   Call Gp_Sp_Collection(ss1, 78, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '坯料类别
   Call Gp_Sp_Collection(ss1, 79, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '坯料类别
   
   Call Gp_Sp_Collection(ss1, 80, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 81, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 82, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 83, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 84, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 85, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
   Call Gp_Sp_Collection(ss1, 86, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGD2081C.P_REFER", Key:="P-R"
    sc1.Add Item:="CGD2081C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="CGD2081C.P_MODIFY", Key:="P-M"
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
    
    Call Gp_Sp_ColHidden(ss1, SPD_LOT_NO, True)
    Call Gp_Sp_ColHidden(ss1, SPD_ACT_THK, True)
    Call Gp_Sp_ColHidden(ss1, SPD_ACT_WID, True)
    Call Gp_Sp_ColHidden(ss1, SPD_ACT_LEN, True)
    Call Gp_Sp_ColHidden(ss1, SPD_UST_FL, True)
    Call Gp_Sp_ColHidden(ss1, SPD_LAST_YN, True)
    Call Gp_Sp_ColHidden(ss1, SPD_GRID_YN, True)
    Call Gp_Sp_ColHidden(ss1, SPD_APLY_STDSPEC_NEW, True)
    Call Gp_Sp_ColHidden(ss1, SPD_INSP_CD, True)
    Call Gp_Sp_ColHidden(ss1, SPD_INSP_CD1, True)
    Call Gp_Sp_ColHidden(ss1, SPD_INSP_CD2, True)
    Call Gp_Sp_ColHidden(ss1, SPD_INSP_CD3, True)
    Call Gp_Sp_ColHidden(ss1, SPD_INSP_CD4, True)
    Call Gp_Sp_ColHidden(ss1, SPD_SURF_YN, True)
    'Call Gp_Sp_ColHidden(ss1, SPD_SURF_GRD, True)
    Call Gp_Sp_ColHidden(ss1, SPD_SURF_GRD_DET, True)
    Call Gp_Sp_ColHidden(ss1, SPD_PROD_DATE, True)
    Call Gp_Sp_ColHidden(ss1, SPD_EMP_CD, True)
    Call Gp_Sp_ColHidden(ss1, SPD_INSP_WAVE, True)
    Call Gp_Sp_ColHidden(ss1, SPD_INSP_VERT_DEG, True)
    Call Gp_Sp_ColHidden(ss1, SPD_INSP_RECT_DEG, True)
    Call Gp_Sp_ColHidden(ss1, SPD_TOP_GRID_GRD, True)
    Call Gp_Sp_ColHidden(ss1, SPD_TOP_GRID_YRD, True)
    Call Gp_Sp_ColHidden(ss1, SPD_TOP_GRID_DEEP, True)
    Call Gp_Sp_ColHidden(ss1, SPD_BOT_GRID_GRD, True)
    Call Gp_Sp_ColHidden(ss1, SPD_BOT_GRID_YRD, True)
    Call Gp_Sp_ColHidden(ss1, SPD_BOT_GRID_DEEP, True)
    Call Gp_Sp_ColHidden(ss1, SPD_GRID_TIME, True)
    Call Gp_Sp_ColHidden(ss1, SPD_INSP_DIAGONAL1, True)
    Call Gp_Sp_ColHidden(ss1, SPD_INSP_DIAGONAL2, True)
    Call Gp_Sp_ColHidden(ss1, SPD_LOC, True)
    Call Gp_Sp_ColHidden(ss1, SPD_INSP_WAVE1, True)
    
    Call Gp_Sp_ColHidden(ss1, SPD_THK1, True)
    Call Gp_Sp_ColHidden(ss1, SPD_THK2, True)
    Call Gp_Sp_ColHidden(ss1, SPD_THK3, True)
    Call Gp_Sp_ColHidden(ss1, SPD_THK4, True)
    Call Gp_Sp_ColHidden(ss1, SPD_THK5, True)
    Call Gp_Sp_ColHidden(ss1, SPD_THK6, True)

    txt_plt.Text = "C3"
'    Call txt_plt_KeyUp(0, 0)
    
    txt_line.Text = "1"
    txt_rec_sts.Text = "1"
    Call Gp_Sp_ColHidden(ss1, SPD_LINE2, True)
    opt_line1 = True
    opt_line3 = True
    opt_CHK_PRD_GRD(0).Value = True
    
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
        
        txt_plt.Text = "C3"
'        Call txt_plt_KeyUp(0, 0)
        txt_line.Text = "1"
        txt_rec_sts.Text = "1"
        opt_line3 = True
        txt_stdspec_chg = ""
        TXT_REMARK = ""  '2012-3-14 Modify by LiChao
        TXT_sUserID = ""
        TXT_sUserID_Tail = ""
        txt_Color_code = ""
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()
    
    Dim iCount      As Integer
    Dim sPlateNo    As String
    Dim iRow        As Integer
    Dim sord_cnt    As Integer
    Dim sFlag       As String
    Dim sexport     As String

    Dim inum As Integer
    Dim lRow As Integer
    
    Dim sCurDate As String
    Dim sDel_To_Date As String
    
    sCurDate = Format(Now, "YYYYMM")
    
'    If Gf_Sp_ProceExist(sc1.Item("Spread")) Then Exit Sub
            
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss1.OperationMode = OperationModeNormal
    End If
    
    With ss1
        For iCount = 1 To .MaxRows
            .ROW = iCount
            .Col = SPD_PLATE_NO
             sPlateNo = .Text
            If Left(.Text, 12) = Left(sPlateNo, 12) Then
            Else
               .ROW = iCount - 1
               .Col = SPD_LAST_YN
               .Value = 1
            End If
        Next iCount
        For iRow = 1 To .MaxRows
            .ROW = iRow:            .Col = SPD_ORD_CNT:        sord_cnt = Val(.Text)
                
             If sord_cnt > 1 Then
                Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, , SSP1.BackColor)
             End If
        Next iRow
        
        '超交货期用红色显示 add by liqian 2012-07-23
        For iRow = 1 To .MaxRows
            .ROW = iRow:             .Col = SPD_DEL_DATE_TO
             sDel_To_Date = Mid(.Value, 1, 6)
             If sDel_To_Date < sCurDate Then
                  Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, &HFF&)
             End If
        Next iRow
        
        For iRow = 1 To .MaxRows
            '是否定制配送
            .ROW = iRow:
            .Col = SPD_FLAG_FL: sFlag = Trim(.Text)
            If sFlag = "Y" Then
               Call Gp_Sp_BlockColor(ss1, SPD_PLATE_NO, SPD_PLATE_NO, iRow, iRow, SSP4.BackColor)
            End If
            '是否出口订单
            .ROW = iRow:
            .Col = SPD_EXPORT: sexport = Trim(.Text)
            If sexport = "Y" Then
               Call Gp_Sp_BlockColor(ss1, SPD_PLATE_NO, SPD_PLATE_NO, iRow, iRow, SSP5.BackColor)
            End If
        Next iRow
        
    End With

End Sub

Private Sub SSCHK_GRID_YN_Click()
    If SSCHK_GRID_YN.Value = ssCBUnchecked Then
        CHK_TOP_GRD(0).Enabled = False:        CHK_TOP_GRD(0).Value = ssCBUnchecked
        CHK_TOP_GRD(1).Enabled = False:        CHK_TOP_GRD(1).Value = ssCBUnchecked
        CHK_BOT_GRD(0).Enabled = False:        CHK_BOT_GRD(0).Value = ssCBUnchecked
        CHK_BOT_GRD(1).Enabled = False:        CHK_BOT_GRD(1).Value = ssCBUnchecked
        SDB_TOP_GRID_YRD.Enabled = False:      SDB_TOP_GRID_YRD.Text = ""
        SDB_BOT_GRID_YRD.Enabled = False:      SDB_BOT_GRID_YRD.Text = ""
        SDB_TOP_GRID_DEEP.Enabled = False:     SDB_TOP_GRID_DEEP.Text = ""
        SDB_BOT_GRID_DEEP.Enabled = False:     SDB_BOT_GRID_DEEP.Text = ""
        TXT_GRID_TIME.Enabled = False:         TXT_GRID_TIME.Text = ""
                
'        CHK_NEXT_PRC(1).Enabled = True
    Else
        CHK_TOP_GRD(0).Enabled = True
        CHK_TOP_GRD(1).Enabled = True
        CHK_BOT_GRD(0).Enabled = True
        CHK_BOT_GRD(1).Enabled = True
        SDB_TOP_GRID_YRD.Enabled = True
        SDB_BOT_GRID_YRD.Enabled = True
        SDB_TOP_GRID_DEEP.Enabled = True
        SDB_BOT_GRID_DEEP.Enabled = True
        TXT_GRID_TIME.Enabled = True
        
        TXT_GRID_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        
        CHK_TOP_GRD(0).Value = ssCBChecked
        Call CHK_TOP_GRD_Click(0)
        CHK_BOT_GRD(0).Value = ssCBChecked
        Call CHK_BOT_GRD_Click(0)

    End If
End Sub

Private Sub CHK_TOP_GRD_Click(Index As Integer)
    Dim iNext       As Integer
    
    If sCheck <> "" Then Exit Sub

    sCheck = "**"
    
    If Index = 0 Then
        iNext = 1
    Else
        iNext = 0
    End If
    
    If CHK_TOP_GRD(Index).Value = ssCBUnchecked Then
        If CHK_TOP_GRD(iNext).Value = ssCBUnchecked Then
            TXT_TOP_GRID_GRD.Text = ""
            CHK_TOP_GRD(Index).ForeColor = &H808080
            sCheck = ""
            Exit Sub
        End If
    End If
    
    CHK_TOP_GRD(Index).ForeColor = &HFF&
    CHK_TOP_GRD(Index).Value = ssCBChecked
                
    CHK_TOP_GRD(iNext).ForeColor = &H808080
    CHK_TOP_GRD(iNext).Value = ssCBUnchecked

    TXT_TOP_GRID_GRD.Text = CHK_TOP_GRD(Index).Tag
    sCheck = ""
    
End Sub

Private Sub CHK_BOT_GRD_Click(Index As Integer)
    Dim iNext       As Integer
    
    If sCheck <> "" Then Exit Sub

    sCheck = "**"
    
    If Index = 0 Then
        iNext = 1
    Else
        iNext = 0
    End If
    
    If CHK_BOT_GRD(Index).Value = ssCBUnchecked Then
        If CHK_BOT_GRD(iNext).Value = ssCBUnchecked Then
            TXT_BOT_GRID_GRD.Text = ""
            CHK_BOT_GRD(Index).ForeColor = &H808080
            sCheck = ""
            Exit Sub
        End If
    End If
    
    CHK_BOT_GRD(Index).ForeColor = &HFF&
    CHK_BOT_GRD(Index).Value = ssCBChecked
                
    CHK_BOT_GRD(iNext).ForeColor = &H808080
    CHK_BOT_GRD(iNext).Value = ssCBUnchecked

    TXT_BOT_GRID_GRD.Text = CHK_BOT_GRD(Index).Tag
    sCheck = ""
    
End Sub

Private Sub txt_Color_code_Change()
If Len(Trim(txt_Color_code)) = txt_Color_code.MaxLength Then
        txt_Color_name.Text = Gf_ComnNameFind(M_CN1, "CG002", Trim(txt_Color_code.Text), 1)
    Else
        txt_Color_name.Text = ""
    End If
End Sub

Private Sub txt_Color_code_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
            
        DD.sWitch = "MS"
        DD.sKey = "CG002"
        DD.rControl.Add Item:=txt_Color_code
        DD.rControl.Add Item:=txt_Color_name
        
        DD.nameType = "1"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    End If

End Sub

Private Sub txt_Color_code_DblClick()
    Call txt_Color_code_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_GRID_TIME_DblClick()
    TXT_GRID_TIME.RawData = Gf_DTSet(M_CN1, , "X")
End Sub

Public Sub Form_Pro()

    Dim iRow As Integer
    Dim sMesg   As String
    
    Dim sMark_no As String
    Dim sPlate_no As String
    Dim sThk As String
    Dim sWid As String
    Dim sLen As String
    Dim sWgt As String
    Dim sSpec As String
    Dim sStdspec_YY As String
    
    If SSCHK_GRID_YN.Value = ssCBChecked Then
        If Not Gp_DateCheck(TXT_GRID_TIME) Then
            sMesg = " 请正确输入修磨时间 ！"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        End If
        If TXT_TOP_GRID_GRD.Text = "" Then
            sMesg = " 请正确输入上表面修磨后判定 ！"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        End If
        If TXT_BOT_GRID_GRD.Text = "" Then
            sMesg = " 请正确输入下表面修磨后判定 ！"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        End If
    End If
    
    With ss1
         .ROW = .ActiveRow
         .Col = 0
         If .Text <> "Update" Then
            Call Gp_MsgBoxDisplay("请选择钢板号")
            Exit Sub
         End If

         .Col = SPD_THK:                 .Value = SDB_THK.Value
         .Col = SPD_WID:                 .Value = SDB_WID.Value
         .Col = SPD_LEN:                 .Value = SDB_LEN.Value
         .Col = SPD_ACT_THK:             .Value = SDB_ACT_THK.Value
         .Col = SPD_ACT_WID:             .Value = SDB_ACT_WID.Value
         .Col = SPD_ACT_LEN:             .Value = SDB_ACT_LEN.Value
         .Col = SPD_THK_MIN:             .Value = SDB_INSP_THK_MN.Value
         .Col = SPD_THK_MAX:             .Value = SDB_INSP_THK_MX.Value
         .Col = SPD_WID_MIN:             .Value = SDB_INSP_WID_MN.Value
         .Col = SPD_WID_MAX:             .Value = SDB_INSP_WID_MX.Value
         .Col = SPD_LEN_MIN:             .Value = SDB_INSP_LEN_MN.Value
         .Col = SPD_LEN_MAX:             .Value = SDB_INSP_LEN_MX.Value
         .Col = SPD_INSP_CD:             .Text = TXT_INSP_FLAW(5).Text
         .Col = SPD_INSP_CD1:            .Text = TXT_INSP_FLAW(0).Text
         .Col = SPD_INSP_CD2:            .Text = TXT_INSP_FLAW(1).Text
         .Col = SPD_INSP_CD3:            .Text = TXT_INSP_FLAW(3).Text
         .Col = SPD_INSP_CD4:            .Text = TXT_INSP_FLAW(4).Text
         .Col = SPD_APLY_STDSPEC_NEW:    .Text = txt_stdspec_chg.Text
         .Col = SPD_LAST_YN:             If SSCHK_LAST_YN.Value = -1 Then .Text = "Y" Else .Text = "N"
         .Col = SPD_SIZE_KND:            .Text = TXT_SIZE_KND.Text
         .Col = SPD_TRIM_FL:             .Text = TXT_TRIM_FL.Text
         .Col = SPD_GRID_YN:             If SSCHK_GRID_YN.Value = 1 Then .Text = "Y" Else .Text = "N"
         .Col = SPD_PROD_REMARK:         .Text = TXT_REMARK.Text
         .Col = SPD_SURF_GRD:            .Text = TXT_INSP_MAIN_GRD.Text
          If .Text <> "1" And .Text <> "2" And .Text <> "3" And .Text <> "4" And .Text <> "5" And .Text <> "7" Then
             Call Gp_MsgBoxDisplay("表面等级输入错误，请确认")
             Exit Sub
          End If
         .Col = SPD_TOP_GRID_GRD:        .Text = TXT_TOP_GRID_GRD.Text
         .Col = SPD_TOP_GRID_YRD:        .Text = SDB_TOP_GRID_YRD.Text
         .Col = SPD_TOP_GRID_DEEP:       .Text = SDB_TOP_GRID_DEEP.Text
         .Col = SPD_BOT_GRID_GRD:        .Text = TXT_BOT_GRID_GRD.Text
         .Col = SPD_BOT_GRID_YRD:        .Text = SDB_BOT_GRID_YRD.Text
         .Col = SPD_BOT_GRID_DEEP:       .Text = SDB_BOT_GRID_DEEP.Text
         .Col = SPD_GRID_TIME:           .Text = TXT_GRID_TIME.Text
         .Col = SPD_INSP_DIAGONAL1:      .Text = SDB_INSP_DIAGONAL1.Value
         .Col = SPD_INSP_DIAGONAL2:      .Text = SDB_INSP_DIAGONAL2.Value
         .Col = SPD_LOC:                 .Text = TXT_LOC.Text
         .Col = SPD_INSP_WAVE1:          .Value = TXT_WAVE1.Text
         .Col = SPD_PLATE_COLOR:         .Value = txt_Color_code.Text
         
         .Col = SPD_THK1:                .Value = SDB_HD1.Value
         .Col = SPD_THK2:                .Value = SDB_HD2.Value
         .Col = SPD_THK3:                .Value = SDB_HD3.Value
         .Col = SPD_THK4:                .Value = SDB_HD4.Value
         .Col = SPD_THK5:                .Value = SDB_HD5.Value
         .Col = SPD_THK6:                .Value = SDB_HD6.Value
        ' .Col = SPD_FLAW_YN:             If CHK_FLAW_YN.Value = -1 Then .Text = "Y" Else .Text = "N"
         
    End With
    
    If txt_rec_sts = "1" Then
        If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
'            TXT_REMARK = ""    '2012-3-14 Modify by LiChao
'            TXT_WAVE = ""
'            TXT_VERT_DEG = ""
'            TXT_RECT_DEG = ""
'            SSCHK_LY_YN.Value = 0
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call MenuTool_ReSet
        End If
    End If

    iRow = iRow + 10
    If iRow > ss1.MaxRows Then
       iRow = ss1.MaxRows
    End If
    
    'add by liqian at 2012-03-29 保存完成刷新整个查询画面
    Call Form_Ref
    
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
    
    Dim sSize_knd   As String
    Dim sTrim_fl    As String
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
            .Col = SPD_SIZE_KND:      sSize_knd = .Text
            .Col = SPD_TRIM_FL:       sTrim_fl = .Text
            .Col = SPD_APLY_STDSPEC:  sAply_stdspec = .Text
            .Col = SPD_EMP_CD:        sEmp_cd = .Text
    
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
        .Col = 0: .Text = "Input"
        .Col = SPD_PLATE_NO: .Text = Mid(.Text, 1, 12) & Format(Val(Mid(.Text, 13, 2) & "") + 1, "00")
        .Col = SPD_SURF_GRD:      .Value = 1
'        .Col = SPD_LINE1:         .Value = 1
        .Col = 0:                 .Text = "Input"
         Call Gp_Sp_BlockColor(ss1, 1, -1, .ActiveRow, .ActiveRow, , SSP2.BackColor)
        
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

Private Sub opt_CHK_PRD_GRD_Click(Index As Integer, Value As Integer)
    If Index = 0 Then
       TXT_INSP_MAIN_GRD = "1"
       opt_CHK_PRD_GRD(0).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
    ElseIf Index = 1 Then
       TXT_INSP_MAIN_GRD = "2"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       ULabel21.Caption = "改判原因"   '2012-03-14 Modify by LiChao
    ElseIf Index = 2 Then
        TXT_INSP_MAIN_GRD = "3"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       ULabel21.Caption = "协议原因"  '2012-03-14 Modify by LiChao
    ElseIf Index = 3 Then
        TXT_INSP_MAIN_GRD = "4"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       ULabel21.Caption = "待判原因"  '2012-03-14 Modify by LiChao
    ElseIf Index = 4 Then
        TXT_INSP_MAIN_GRD = "5"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &HFF&       'red
       ULabel21.Caption = "次品原因"  '2012-03-14 Modify by LiChao
    End If
    
End Sub

Private Sub opt_line1_Click(Value As Integer)
    
    If opt_line1 Then
        opt_line1.ForeColor = &HFF&
        opt_line2.ForeColor = &H80000012
        opt_line5.ForeColor = &H80000012
        txt_line = "1"
        If ss1.MaxRows > 0 Then Call Form_Ref
        Call Gp_Sp_ColHidden(ss1, SPD_LINE1, False)
        Call Gp_Sp_ColHidden(ss1, SPD_LINE2, True)
    End If
    
End Sub

Private Sub opt_line2_Click(Value As Integer)

    If opt_line2 Then
        opt_line2.ForeColor = &HFF&
        opt_line1.ForeColor = &H80000012
        opt_line5.ForeColor = &H80000012
        txt_line = "2"
        If ss1.MaxRows > 0 Then Call Form_Ref
        Call Gp_Sp_ColHidden(ss1, SPD_LINE2, False)
        Call Gp_Sp_ColHidden(ss1, SPD_LINE1, True)
    End If
    
End Sub

Private Sub opt_line5_Click(Value As Integer)

    If opt_line5 Then
        opt_line5.ForeColor = &HFF&
        opt_line1.ForeColor = &H80000012
        opt_line2.ForeColor = &H80000012
        txt_line = ""
        If ss1.MaxRows > 0 Then Call Form_Ref
        Call Gp_Sp_ColHidden(ss1, SPD_LINE2, False)
        Call Gp_Sp_ColHidden(ss1, SPD_LINE1, False)
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

'add by liqian at 2012-03-14 根据实测长度值计算公称长度
Private Sub SDB_ACT_LEN_Change()
 Dim iLen As Integer
     If TXT_SIZE_KND <> "01" Then
        If SDB_ACT_LEN.Value > 0 Then
           iLen = Int(SDB_ACT_LEN.Value / 50) * 50
           SDB_LEN.Value = iLen
        End If
     End If
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
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
    
    Dim sMesg As String
    
    iCol = Col
    iRow = ROW

    If ROW <= 0 Then Exit Sub
    If Col <> SPD_LINE1 And Col <> SPD_LINE2 And Col <> SPD_SURF_YN Then Exit Sub
    If Not Gf_Sc_Authority(sAuthority, "U") Then Exit Sub
    
    SSCHK_GRID_YN.Value = ssCBUnchecked
    
    iRowto = iRow - 1
    iRowfr = iRow + 1
    
    If iRowto > 0 Then
        For iRowNum = 1 To iRowto
             
             ss1.Col = 0
             ss1.ROW = iRowNum
             
             If ss1.Text <> "" Then
                ss1.Text = ""
                ss1.Col = SPD_LINE1:                ss1.Value = 0
                ss1.Col = SPD_LINE2:                ss1.Value = 0
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
                ss1.Col = SPD_LINE1:                ss1.Value = 0
                ss1.Col = SPD_LINE2:                ss1.Value = 0
                Call Gp_Sp_BlockColor(ss1, 1, -1, iRowNum, iRowNum)
                Exit For
             End If
        Next iRowNum
    End If

    ss1.ROW = iRow

    If Col = SPD_LINE1 And ButtonDown = 1 Then
        ss1.Col = SPD_LINE2:        ss1.Text = 0
    ElseIf Col = SPD_LINE2 And ButtonDown = 1 Then
        ss1.Col = SPD_LINE1:        ss1.Text = 0
    End If
    
    If Col = SPD_SURF_YN Then
       If ButtonDown = 1 Then
          ss1.Col = SPD_SURF_GRD:        ss1.Text = "1"
       Else
          ss1.Col = SPD_SURF_GRD:        ss1.Text = "4"
       End If
    End If
    

    ss1.Col = 0
    ss1.Text = "Update"
    
    TXT_INSP_FLAW(0).Text = ""
    TXT_INSP_FLAW(1).Text = ""
    TXT_INSP_FLAW(3).Text = ""
    TXT_INSP_FLAW(4).Text = ""
    TXT_INSP_FLAW(5).Text = ""
    
    ss1.Col = SPD_LINE1:    sCheck1 = ss1.Value
    ss1.Col = SPD_LINE2:    sCheck2 = ss1.Value

    If sCheck1 = 0 And sCheck2 = 0 Then
        ss1.Col = 0:        ss1.Text = ""
        Call Gp_Sp_BlockColor(ss1, 1, -1, iRow, iRow)
    Else
        Call Gp_Sp_BlockColor(ss1, 1, -1, iRow, iRow, , SSP2.BackColor)
'        TXT_INSP_FLAW(0).Text = ""
'        TXT_INSP_FLAW(1).Text = ""
'        TXT_INSP_FLAW(3).Text = ""
'        TXT_INSP_FLAW(4).Text = ""
'        TXT_INSP_FLAW(5).Text = ""
        opt_CHK_PRD_GRD(0).Value = True
        
        ' add by liqian at 2012-03-29  改判标准到下一块时自动清空
        txt_stdspec_chg.Text = ""
        TXT_INSP_FLAW_NAME(5).Text = ""
        TXT_INSP_FLAW(5).Text = ""
        ' add by liqian at 2012-04-18 留样下一块自动清空
        SSCHK_LY_YN.Value = 0
        SSCHK_LAST_YN.Value = 0
        CHK_FLAW_YN.Value = 0

        ss1.ROW = iRow:
        ss1.Col = SPD_THK:       SDB_THK.Value = ss1.Value
        ss1.Col = SPD_WID:       SDB_WID.Value = ss1.Value
        ss1.Col = SPD_LEN:       SDB_LEN.Value = ss1.Value
        ss1.Col = SPD_THK_MIN:   If ss1.Value = "" Then SDB_INSP_THK_MN.Value = 0 Else SDB_INSP_THK_MN.Value = ss1.Value
        ss1.Col = SPD_THK_MAX:   If ss1.Value = "" Then SDB_INSP_THK_MX.Value = 0 Else SDB_INSP_THK_MX.Value = ss1.Value
        ss1.Col = SPD_WID_MIN:   If ss1.Value = "" Then SDB_INSP_WID_MN.Value = 0 Else SDB_INSP_WID_MN.Value = ss1.Value
        ss1.Col = SPD_WID_MAX:   If ss1.Value = "" Then SDB_INSP_WID_MX.Value = 0 Else SDB_INSP_WID_MX.Value = ss1.Value
        ss1.Col = SPD_LEN_MIN:   If ss1.Value = "" Then SDB_INSP_LEN_MN.Value = 0 Else SDB_INSP_LEN_MN.Value = ss1.Value
        ss1.Col = SPD_LEN_MAX:   If ss1.Value = "" Then SDB_INSP_LEN_MX.Value = 0 Else SDB_INSP_LEN_MX.Value = ss1.Value
        ss1.Col = SPD_ACT_THK:   If ss1.Value = "" Then SDB_ACT_THK.Value = 0 Else SDB_ACT_THK.Value = ss1.Value
        ss1.Col = SPD_ACT_WID:   If ss1.Value = "" Then SDB_ACT_WID.Value = 0 Else SDB_ACT_WID.Value = ss1.Value
        ss1.Col = SPD_ACT_LEN:   If ss1.Value = "" Then SDB_ACT_LEN.Value = 0 Else SDB_ACT_LEN.Value = ss1.Value '2012-3-20 Modify by LiChao
        ss1.Col = SPD_INSP_WAVE:      If ss1.Value = "" Then TXT_WAVE.Text = "" Else TXT_WAVE.Text = ss1.Value
        ss1.Col = SPD_INSP_VERT_DEG:  If ss1.Value = "" Then TXT_VERT_DEG.Text = "" Else TXT_VERT_DEG.Text = ss1.Value
        ss1.Col = SPD_INSP_RECT_DEG:  If ss1.Value = "" Then TXT_RECT_DEG.Text = "" Else TXT_RECT_DEG.Text = ss1.Value
        ' add by liqian at 2012-04-18 留样下一块自动清空,备注重查
        ss1.Col = SPD_PROD_REMARK:  If ss1.Value = "" Then TXT_REMARK.Text = "" Else TXT_REMARK.Text = ss1.Value
        ss1.Col = SPD_TOP_GRID_GRD:   If ss1.Value = "" Then TXT_TOP_GRID_GRD.Text = "" Else TXT_TOP_GRID_GRD.Text = ss1.Value
        ss1.Col = SPD_TOP_GRID_YRD:   If ss1.Value = "" Then SDB_TOP_GRID_YRD.Value = 0 Else SDB_TOP_GRID_YRD.Value = ss1.Value
        ss1.Col = SPD_TOP_GRID_DEEP:  If ss1.Value = "" Then SDB_TOP_GRID_DEEP.Value = 0 Else SDB_TOP_GRID_DEEP.Value = ss1.Value
        ss1.Col = SPD_BOT_GRID_GRD:   If ss1.Value = "" Then TXT_BOT_GRID_GRD.Text = "" Else TXT_BOT_GRID_GRD.Text = ss1.Value
        ss1.Col = SPD_BOT_GRID_YRD:   If ss1.Value = "" Then SDB_BOT_GRID_YRD.Value = 0 Else SDB_BOT_GRID_YRD.Value = ss1.Value
        ss1.Col = SPD_BOT_GRID_DEEP:  If ss1.Value = "" Then SDB_BOT_GRID_DEEP.Value = 0 Else SDB_BOT_GRID_DEEP.Value = ss1.Value
        ss1.Col = SPD_GRID_TIME:      If ss1.Value = "" Then TXT_GRID_TIME.Text = "" Else TXT_GRID_TIME.Text = ss1.Value
        ss1.Col = SPD_INSP_DIAGONAL1: If ss1.Value = "" Then SDB_INSP_DIAGONAL1.Value = 0 Else SDB_INSP_DIAGONAL1.Value = ss1.Value
        ss1.Col = SPD_INSP_DIAGONAL2: If ss1.Value = "" Then SDB_INSP_DIAGONAL2.Value = 0 Else SDB_INSP_DIAGONAL2.Value = ss1.Value
        ss1.Col = SPD_LOC:            If ss1.Value = "" Then TXT_LOC.Text = "" Else TXT_LOC.Text = ss1.Value
        ss1.Col = SPD_INSP_WAVE1:     If ss1.Value = "" Then TXT_WAVE1.Text = "" Else TXT_WAVE1.Text = ss1.Value
        
        ss1.Col = SPD_THK1:           If ss1.Value = "" Then SDB_HD1.Value = 0 Else SDB_HD1.Value = ss1.Value
        ss1.Col = SPD_THK2:           If ss1.Value = "" Then SDB_HD2.Value = 0 Else SDB_HD2.Value = ss1.Value
        ss1.Col = SPD_THK3:           If ss1.Value = "" Then SDB_HD3.Value = 0 Else SDB_HD3.Value = ss1.Value
        ss1.Col = SPD_THK4:           If ss1.Value = "" Then SDB_HD4.Value = 0 Else SDB_HD4.Value = ss1.Value
        ss1.Col = SPD_THK5:           If ss1.Value = "" Then SDB_HD5.Value = 0 Else SDB_HD5.Value = ss1.Value
        ss1.Col = SPD_THK6:           If ss1.Value = "" Then SDB_HD6.Value = 0 Else SDB_HD6.Value = ss1.Value
        
        ss1.Col = SPD_INSP_CD1:       If ss1.Text = "" Then TXT_INSP_FLAW(0).Text = "" Else TXT_INSP_FLAW(0).Text = ss1.Text
        ss1.Col = SPD_INSP_CD3:       If ss1.Text = "" Then TXT_INSP_FLAW(3).Text = "" Else TXT_INSP_FLAW(3).Text = ss1.Text
        

        '2011-08-29   modified by liqian for 画面显示定尺改为汉字,双击时对应栏位还原回代码表示,保证定尺保存为01,02,06,08,..类型
        ss1.Col = SPD_SIZE_KND:  Select Case ss1.Text
                                 Case "定尺"
                                      TXT_SIZE_KND.Text = "01"
                                 Case "单定尺"
                                      TXT_SIZE_KND.Text = "02"
                                 Case "非尺"
                                      TXT_SIZE_KND.Text = "06"
                                 Case "小尺板"
                                      TXT_SIZE_KND.Text = "08"
                                 Case Else
                                      TXT_SIZE_KND.Text = ""
                                 End Select
        ss1.Col = SPD_TRIM_FL:   TXT_TRIM_FL.Text = ss1.Text
        
    End If
    
    ss1.Col = SPD_EMP_CD:      ss1.Text = sUserID
    
    TXT_CUT_TIME.RawData = Gf_DTSet(M_CN1, , "X")
    ss1.Col = SPD_PROD_DATE:   ss1.Text = TXT_CUT_TIME.Text
    
    ss1.Col = SPD_INS_MAN:        ss1.Text = TXT_sUserID.Text
    ss1.Col = SPD_INS_MAN_TAIL:   ss1.Text = TXT_sUserID_Tail.Text
     
    ss1.Col = SPD_PLATE_COLOR:    ss1.Text = txt_Color_code.Text
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
  Dim sMesg As String

  If ROW <= 0 Then Exit Sub

' Add by liqian at 2012-03-16
  If TXT_sUserID.Text = "" Then
     sMesg = "请选择头部检验人员工号！"
     Call Gp_MsgBoxDisplay(sMesg)
     Exit Sub
  End If

  If TXT_sUserID_Tail.Text = "" Then
     sMesg = "请选择尾部检验人员工号！"
     Call Gp_MsgBoxDisplay(sMesg)
     Exit Sub
  End If
  
  ss1.ROW = ROW

  If Col = SPD_PROD_DATE Then
     TXT_CUT_TIME.RawData = Gf_DTSet(M_CN1, , "X")
     ss1.Col = SPD_PROD_DATE:     ss1.Text = TXT_CUT_TIME.Text
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

Private Sub SSCHK_LY_YN_Click(Value As Integer)
    If Value = 0 Then
        TXT_REMARK.Text = TXT_REMARK.Text
    Else
        If TXT_REMARK.Text <> "" Then
            TXT_REMARK.Text = TXT_REMARK.Text + ";留样"
        Else
            TXT_REMARK.Text = "留样"
        End If
        ss1.Col = SPD_PROD_REMARK
        If ss1.Text <> "" Then
            ss1.Text = ss1.Text + ";留样"
        Else
            ss1.Text = "留样"
        End If
    End If
           
End Sub

'矫直指示
Private Sub CHK_CL_FL_Click()
ss1.Col = SPD_CL_FL
If CHK_CL_FL.Value = ssCBChecked Then
       TXT_CL.Text = "Y"
       ss1.Text = TXT_CL.Text
    Else
       TXT_CL.Text = ""
       ss1.Text = TXT_CL.Text
    End If

End Sub

Private Sub CHK_FLAW_YN_Click()
ss1.Col = SPD_FLAW_YN
If CHK_FLAW_YN.Value = ssCBChecked Then
       TXT_FLAW.Text = "Y"
       ss1.Text = TXT_FLAW.Text
    Else
       TXT_FLAW.Text = ""
       ss1.Text = TXT_FLAW.Text
    End If

End Sub

Private Sub SSCHK_SIZE_KND_Click(Value As Integer)

   If Value = 0 Then
      TXT_SIZE_KND.Text = "02"
   Else
      TXT_SIZE_KND.Text = "01"
   End If
    
End Sub

Private Sub SSCHK_TRIM_FL_Click(Value As Integer)
   If Value = 0 Then
      TXT_TRIM_FL.Text = "N"
   Else
      TXT_TRIM_FL.Text = "Y"
   End If
End Sub


Private Sub TXT_INSP_FLAW_Change(Index As Integer)
    
    TXT_INSP_FLAW_NAME(Index).Text = Gf_ComnNameFind(M_CN1, "G0002", TXT_INSP_FLAW(Index).Text, 1)

    If Len(Trim(TXT_INSP_FLAW(0).Text)) = 3 Then
       ss1.Col = SPD_INSP_CD1
       ss1.Text = TXT_INSP_FLAW(0).Text
    End If
    
    If Len(Trim(TXT_INSP_FLAW(0).Text)) = 0 Then
       ss1.Col = SPD_INSP_CD1
       ss1.Text = ""
    End If
    
    If Len(Trim(TXT_INSP_FLAW(3).Text)) = 3 Then
       ss1.Col = SPD_INSP_CD3
       ss1.Text = TXT_INSP_FLAW(3).Text
    End If
    
    If Len(Trim(TXT_INSP_FLAW(3).Text)) = 0 Then
       ss1.Col = SPD_INSP_CD3
       ss1.Text = ""
    End If
        
End Sub

Private Sub TXT_INSP_FLAW_NAME_DblClick(Index As Integer)
    DD.sWitch = "MS"
    DD.sKey = "G0002"
    DD.rControl.Add Item:=TXT_INSP_FLAW(Index)

    DD.nameType = "2"

    Call Gf_Common_DD(M_CN1, vbKeyF4)
    
    If Len(Trim(TXT_INSP_FLAW(Index).Text)) = 3 Then
        TXT_INSP_FLAW_NAME(Index).Text = Gf_ComnNameFind(M_CN1, "G0002", Trim(TXT_INSP_FLAW(Index).Text), 1)
    Else
        TXT_INSP_FLAW_NAME(Index).Text = ""
    End If
End Sub

Private Sub TXT_RECT_DEG_Change()
    ss1.Col = SPD_INSP_RECT_DEG
    If TXT_RECT_DEG.Text <> "" Then
        ss1.Text = TXT_RECT_DEG.Text
    End If
End Sub

Private Sub txt_size_knd_DblClick()
    Call txt_size_knd_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_size_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sSize_knd As String
    sSize_knd = TXT_SIZE_KND.Text

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"

        DD.rControl.Add Item:=TXT_SIZE_KND

        DD.nameType = "2"
        TXT_SIZE_KND.Text = ""
        Call Gf_Common_DD(M_CN1, KeyCode)
        If TXT_SIZE_KND.Text = "" Then
            TXT_SIZE_KND.Text = sSize_knd
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
    
        If Col = SPD_THK Or Col = SPD_WID Or Col = SPD_LEN Then
            If Mode = 1 Then
               ss1.Col = iCol:      ss1.ROW = iRow
               ss1.Text = 0
            End If
        End If
    
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), iMode)
        
        ss1.ROW = iRow  'ss1.ActiveRow
        ss1.Col = SPD_EMP_CD
        ss1.Text = sUserID

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
        
    End If

End Sub

Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

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

Private Sub TXT_sUserID_DblClick()
    Call TXT_sUserID_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_sUserID_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0054"

        DD.rControl.Add Item:=TXT_sUserID

        DD.nameType = "2"
'        TXT_sUserID.Text = ""
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub TXT_sUserID_Tail_DblClick()
    Call TXT_sUserID_KeyUp_Tail(vbKeyF4, 0)
End Sub

Private Sub TXT_sUserID_KeyUp_Tail(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0054"

        DD.rControl.Add Item:=TXT_sUserID_Tail

        DD.nameType = "2"
'        TXT_sUserID.Text = ""
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub TXT_TRIM_FL_DblClick()
    Call TXT_TRIM_FL_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_TRIM_FL_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sTrim_fl As String
    sTrim_fl = TXT_TRIM_FL.Text

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0021"

        DD.rControl.Add Item:=TXT_TRIM_FL

        DD.nameType = "2"
        TXT_TRIM_FL.Text = ""
        Call Gf_Common_DD(M_CN1, KeyCode)
        If TXT_TRIM_FL.Text = "" Then
            TXT_TRIM_FL.Text = sTrim_fl
        End If
        
    End If
End Sub

Private Sub TXT_VERT_DEG_Change()
    ss1.Col = SPD_INSP_VERT_DEG
    If TXT_VERT_DEG.Text <> "" Then
        ss1.Text = TXT_VERT_DEG.Text
    End If
End Sub

Private Sub TXT_WAVE_Change()
    ss1.Col = SPD_INSP_WAVE
    If TXT_WAVE.Text <> "" Then
        ss1.Text = TXT_WAVE.Text
    End If
End Sub

Private Sub TXT_LOC_Change()
    ss1.Col = SPD_LOC
    If TXT_LOC.Text <> "" Then
        ss1.Text = TXT_LOC.Text
    End If
End Sub

Private Sub TXT_WAVE1_Change()
    ss1.Col = SPD_INSP_WAVE1
    If TXT_WAVE1.Text <> "" Then
        ss1.Text = TXT_WAVE1.Text
    End If
End Sub

Private Sub SDB_HD1_Change()
    ss1.Col = SPD_THK1
    If SDB_HD1.Text <> "" Then
        ss1.Text = SDB_HD1.Text
    End If
End Sub

Private Sub SDB_HD2_Change()
    ss1.Col = SPD_THK2
    If SDB_HD2.Text <> "" Then
        ss1.Text = SDB_HD2.Text
    End If
End Sub

Private Sub SDB_HD3_Change()
    ss1.Col = SPD_THK3
    If SDB_HD3.Text <> "" Then
        ss1.Text = SDB_HD3.Text
    End If
End Sub

Private Sub SDB_HD4_Change()
    ss1.Col = SPD_THK4
    If SDB_HD4.Text <> "" Then
        ss1.Text = SDB_HD4.Text
    End If
End Sub

Private Sub SDB_HD5_Change()
    ss1.Col = SPD_THK5
    If SDB_HD5.Text <> "" Then
        ss1.Text = SDB_HD5.Text
    End If
End Sub

Private Sub SDB_HD6_Change()
    ss1.Col = SPD_THK6
    If SDB_HD6.Text <> "" Then
        ss1.Text = SDB_HD6.Text
    End If
End Sub
