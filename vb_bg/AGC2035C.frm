VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form AGC2035C 
   Caption         =   "��ӡ��Ϣ���ͽ���_AGC2035C"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   14160
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CBO_sUserID 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "AGC2035C.frx":0000
      Left            =   10500
      List            =   "AGC2035C.frx":0010
      TabIndex        =   74
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   16290
      Top             =   30
   End
   Begin VB.TextBox txt_stdspec 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   8
      Top             =   480
      Width           =   2925
   End
   Begin VB.TextBox txt_stdspec_chg 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7860
      TabIndex        =   7
      Top             =   480
      Width           =   2445
   End
   Begin VB.ComboBox CBO_SHIFT 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "AGC2035C.frx":0038
      Left            =   9015
      List            =   "AGC2035C.frx":0045
      TabIndex        =   6
      Top             =   90
      Width           =   645
   End
   Begin VB.TextBox txt_plt_name 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1740
      TabIndex        =   5
      Tag             =   "����"
      Top             =   90
      Width           =   1320
   End
   Begin VB.TextBox txt_plt 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   4
      Tag             =   "��������"
      Top             =   90
      Width           =   420
   End
   Begin VB.TextBox txt_plate_no 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   14
      TabIndex        =   2
      Tag             =   "���Ϻ�"
      Top             =   480
      Width           =   1755
   End
   Begin VB.ComboBox cbo_group 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "AGC2035C.frx":0055
      Left            =   9660
      List            =   "AGC2035C.frx":0065
      TabIndex        =   0
      Top             =   90
      Width           =   645
   End
   Begin InDate.UDate udt_date_fr 
      Height          =   315
      Left            =   4920
      TabIndex        =   9
      Tag             =   "INS_DATE"
      Top             =   90
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   6360
      TabIndex        =   10
      Tag             =   "INS_DATE"
      Top             =   90
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   105
      Top             =   480
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "�ְ��"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Height          =   8325
      Left            =   60
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
      PaneTree        =   "AGC2035C.frx":0079
      Begin Threed.SSPanel SSPanel1 
         Height          =   2445
         Left            =   0
         TabIndex        =   12
         Tag             =   "172.18.151.145"
         Top             =   0
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   4313
         _Version        =   196609
         BackColor       =   12632319
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox TXT_CUST_CD 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11325
            MaxLength       =   20
            TabIndex        =   83
            Top             =   2460
            Width           =   1095
         End
         Begin VB.TextBox TXT_TO_CUR_INV 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   82
            Top             =   2460
            Width           =   615
         End
         Begin VB.TextBox TXT_WGT 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   81
            Tag             =   "���Ϻ�"
            Top             =   1320
            Width           =   735
         End
         Begin VB.CheckBox chk_Cond 
            BackColor       =   &H00C0C0FF&
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   73
            Top             =   2070
            Width           =   810
         End
         Begin VB.CheckBox chk_Cond 
            BackColor       =   &H00C0C0FF&
            Caption         =   " ��ӡ"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   70
            Top             =   6480
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TXT_X1 
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   60
            Top             =   960
            Width           =   2685
         End
         Begin VB.TextBox TXT_X2 
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   59
            Top             =   1320
            Width           =   2685
         End
         Begin VB.TextBox TXT_X3 
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   58
            Top             =   1680
            Width           =   2685
         End
         Begin VB.TextBox TXT_Producer 
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   57
            Top             =   960
            Width           =   3415
         End
         Begin VB.TextBox TXT_Distance 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   51
            Tag             =   "���Ϻ�"
            Text            =   "6000"
            Top             =   2040
            Width           =   885
         End
         Begin VB.TextBox TXT_Special 
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   50
            Top             =   2040
            Width           =   3415
         End
         Begin VB.TextBox TXT_Punch 
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   49
            Top             =   1680
            Width           =   3415
         End
         Begin VB.TextBox TXT_Paint 
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   48
            Top             =   1320
            Width           =   3415
         End
         Begin VB.TextBox TXT_Class_comp 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7755
            MaxLength       =   14
            TabIndex        =   44
            Tag             =   "���Ϻ�"
            Text            =   "1"
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox TXT_THK 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   27
            Tag             =   "���Ϻ�"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox TXT_LEN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   26
            Tag             =   "���Ϻ�"
            Top             =   1320
            Width           =   915
         End
         Begin VB.TextBox TXT_WID 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   25
            Tag             =   "���Ϻ�"
            Top             =   1320
            Width           =   675
         End
         Begin VB.TextBox TXT_MAT_NO 
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   24
            Tag             =   "���Ϻ�"
            Top             =   960
            Width           =   1965
         End
         Begin VB.TextBox TXT_SPEC 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2160
            TabIndex        =   23
            Top             =   1680
            Width           =   2985
         End
         Begin VB.TextBox TXT_Punch1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9060
            TabIndex        =   22
            Top             =   4080
            Visible         =   0   'False
            Width           =   4515
         End
         Begin VB.TextBox TXT_Punch2 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9060
            TabIndex        =   21
            Top             =   4440
            Visible         =   0   'False
            Width           =   4515
         End
         Begin VB.TextBox TXT_Edge 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9060
            TabIndex        =   20
            Top             =   4800
            Visible         =   0   'False
            Width           =   4515
         End
         Begin VB.TextBox TXT_Bar 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9060
            TabIndex        =   19
            Top             =   5160
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.TextBox TXT_Paint1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5340
            TabIndex        =   18
            Top             =   4110
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.TextBox TXT_Paint2 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5340
            TabIndex        =   17
            Top             =   4470
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.TextBox TXT_Paint3 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5340
            TabIndex        =   16
            Top             =   4830
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.TextBox TXT_Paint4 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5340
            TabIndex        =   15
            Top             =   5190
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.TextBox SOCK 
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   150
            TabIndex        =   14
            Tag             =   "������"
            Top             =   3180
            Width           =   14835
         End
         Begin VB.TextBox TXT_SPEC_DATE 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3060
            TabIndex        =   13
            Top             =   2040
            Width           =   2085
         End
         Begin Threed.SSFrame SSFrame3 
            Height          =   345
            Left            =   2550
            TabIndex        =   28
            Top             =   510
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line1 
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   30
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Layout 1"
               Value           =   -1
            End
            Begin Threed.SSOption opt_line2 
               Height          =   255
               Left            =   1320
               TabIndex        =   30
               Top             =   30
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Layout 2"
            End
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   345
            Left            =   270
            TabIndex        =   31
            Top             =   510
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line3 
               Height          =   255
               Left            =   330
               TabIndex        =   32
               Top             =   30
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "�ƻ�"
               Value           =   -1
            End
            Begin Threed.SSOption opt_line4 
               Height          =   255
               Left            =   1290
               TabIndex        =   33
               Top             =   30
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "ʵ��"
            End
         End
         Begin Threed.SSCommand SSCmd_cnn 
            Height          =   315
            Left            =   630
            TabIndex        =   34
            Top             =   5070
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   255
            BackColor       =   14804173
            Caption         =   "����״̬"
         End
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   255
            Top             =   1680
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            Caption         =   "��׼/ �ƺ�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   255
            Top             =   960
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            Caption         =   "���Ϻ�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Left            =   255
            Top             =   1320
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            Caption         =   "��*��*�� / ��"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.76
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
               Name            =   "����"
               Size            =   9.76
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
            TabIndex        =   45
            Top             =   105
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line5 
               Height          =   255
               Left            =   450
               TabIndex        =   46
               Top             =   30
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               TabIndex        =   47
               Top             =   30
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               TabIndex        =   52
               Top             =   30
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               Name            =   "����"
               Size            =   9.76
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
               Size            =   9.76
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
            TabIndex        =   53
            Top             =   510
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line8 
               Height          =   255
               Left            =   450
               TabIndex        =   54
               Top             =   30
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "3.3mm"
            End
            Begin Threed.SSOption opt_line9 
               Height          =   255
               Left            =   1710
               TabIndex        =   55
               Top             =   30
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "6.2mm"
               Value           =   -1
            End
            Begin Threed.SSOption opt_line10 
               Height          =   255
               Left            =   3300
               TabIndex        =   56
               Top             =   30
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "9.5mm"
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
               Name            =   "����"
               Size            =   9.76
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
               Name            =   "����"
               Size            =   9.76
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
               Name            =   "����"
               Size            =   9.76
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
               Name            =   "����"
               Size            =   9.76
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
            TabIndex        =   61
            Top             =   510
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   12632319
            Begin Threed.SSOption opt_line11 
               Height          =   255
               Left            =   510
               TabIndex        =   62
               Top             =   30
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               TabIndex        =   63
               Top             =   30
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   1
               BackColor       =   12632319
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
         Begin Threed.SSFrame SSFrame7 
            Height          =   345
            Left            =   9720
            TabIndex        =   64
            Top             =   105
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   12632319
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Top Selected"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   540
               TabIndex        =   66
               Top             =   60
               Value           =   1  'Checked
               Width           =   1560
            End
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Bottom Selected"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   6
               Left            =   2670
               TabIndex        =   65
               Top             =   60
               Value           =   1  'Checked
               Width           =   1890
            End
         End
         Begin Threed.SSFrame SSFrame8 
            Height          =   345
            Left            =   270
            TabIndex        =   67
            Top             =   105
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   12632319
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00C0C0FF&
               Caption         =   " �������"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   80
               Top             =   60
               Value           =   1  'Checked
               Width           =   1200
            End
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00C0C0FF&
               Caption         =   " ����"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   69
               Top             =   60
               Value           =   1  'Checked
               Width           =   870
            End
            Begin VB.CheckBox chk_Cond 
               BackColor       =   &H00C0C0FF&
               Caption         =   " ��ӡ"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   68
               Top             =   60
               Value           =   1  'Checked
               Width           =   870
            End
         End
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   12690
            Top             =   2460
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            Caption         =   "Ŀ�Ŀ�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Caption         =   "�ͻ�"
            Alignment       =   1
            BackColor       =   14804173
            BackgroundStyle =   1
            ChiselText      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            Caption         =   "��ӡ line2:"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   42
            Top             =   4470
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ӡ line1:"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   41
            Top             =   4110
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "    ������:"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   40
            Top             =   5160
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "      ����:"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   39
            Top             =   4800
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ӡ line 1:"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   38
            Top             =   4140
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ӡ line 2:"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   37
            Top             =   4500
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ӡ line 3:"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   36
            Top             =   4860
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ӡ line 4:"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   35
            Top             =   5220
            Visible         =   0   'False
            Width           =   1275
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   5850
         Left            =   0
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2475
         Width           =   7620
         _Version        =   393216
         _ExtentX        =   13441
         _ExtentY        =   10319
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
         SpreadDesigner  =   "AGC2035C.frx":00EB
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   5850
         Left            =   7650
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2475
         Width           =   7515
         _Version        =   393216
         _ExtentX        =   13256
         _ExtentY        =   10319
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
         MaxCols         =   44
         MaxRows         =   10
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "AGC2035C.frx":36D4
      End
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   0
      Left            =   105
      Top             =   90
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "��������"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   3270
      Top             =   90
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      Caption         =   "��������"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   3270
      Top             =   480
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      Caption         =   "��׼�� / ����"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   8010
      Top             =   90
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "���/��"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   15390
      Top             =   330
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "191.168.1.34"
      RemotePort      =   2222
   End
   Begin CSTextLibCtl.sitxEdit TXT_CUT_TIME 
      Height          =   315
      Left            =   9600
      TabIndex        =   43
      Tag             =   "��¯ʱ��"
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
   Begin VB.TextBox txt_rec_sts 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      MaxLength       =   1
      TabIndex        =   1
      Tag             =   "CD_MANA_NO"
      Text            =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   390
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   3795
      Top             =   90
      Visible         =   0   'False
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "������"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin VB.TextBox txt_line 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4605
      MaxLength       =   1
      TabIndex        =   3
      Tag             =   "CD_MANA_NO"
      Text            =   "1"
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   10500
      Top             =   90
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "��ҵ��Ա"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel UL_conn 
      Height          =   315
      Left            =   16860
      Top             =   90
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "����״̬"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Left            =   11730
      TabIndex        =   75
      Top             =   90
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1244
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ɫ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   84
         Top             =   60
         Width           =   750
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   79
         Top             =   390
         Width           =   750
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ӡ"
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   78
         Top             =   60
         Width           =   750
      End
      Begin VB.Shape tcpStatus3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   225
         Left            =   3060
         Shape           =   3  'Circle
         Top             =   75
         Width           =   285
      End
      Begin VB.Shape tcpStatus2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   225
         Left            =   870
         Shape           =   3  'Circle
         Top             =   390
         Width           =   285
      End
      Begin VB.Label tcpMsg2 
         BackColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   1170
         TabIndex        =   77
         Top             =   420
         Width           =   945
      End
      Begin VB.Label tcpMsg 
         BackColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   1170
         TabIndex        =   76
         Top             =   90
         Width           =   945
      End
      Begin VB.Shape tcpStatus 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   225
         Left            =   870
         Shape           =   3  'Circle
         Top             =   75
         Width           =   285
      End
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   15390
      Top             =   870
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   25298
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   15390
      Top             =   1380
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   25298
   End
End
Attribute VB_Name = "AGC2035C"
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
'-- Program ID        AGC2035C
'-- Document No       Q-00-0010(Specification)
'-- Designer          ����
'-- Coder             ����
'-- Date              2008.3.24
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE        EDITOR       DESCRIPTION
'-- 1.3   2012.04.25  ����         ����ɫ��ӿ�
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

' 2012-02-13  liqian �Ӷ�����/��Ĵ�����ʾ,��Ҫ�ı������,��ǰ����к��ұ��й����б���,�����¶�����߱���,����SS1�����ط�������.
Const SS1_LINE1 = 1
Const SS1_LINE2 = 2
Const SS1_PLATE_NO = 3
Const SS1_THK = 4
Const SS1_WID = 5
Const SS1_LEN = 6
Const SS1_WGT = 7
Const SS1_LAST_YN = 8
Const SS1_SIZE_KND = 9
Const SS1_TRIM_FL = 10
Const SS1_APLY_STDSPEC = 11
Const SS1_APLY_STDSPEC_NEW = 12
Const SS1_SURF_GRD = 13
Const SS1_MARK_YN = 14
Const SS1_STAMP_YN = 15
Const SS1_BAR_YN = 16
Const SS1_PROD_DATE = 17
Const SS1_EMP_CD = 18
Const SS1_PAINT = 19
Const SS1_LABEL = 20
Const SS1_STDSPEC_YY = 21
Const SS1_STLGRD = 22
Const SS1_PLAN_LEN = 24
Const SS1_URGNT_FL = 25      'Add by LiQian at 2012-11-08 �����������
Const SS2_sPaint = 15   '��ӡ
Const SS2_sPunch = 16   '��ӡ
Const SS2_sEdge = 17    '����
Const SS2_PRODSPECNOA_STD = 37 '�ബ�����׼һ
Const SS2_PRODSPECNOB_STD = 38 '�ബ�����׼��
Const SS2_PRODSPECNOC_STD = 39 '�ബ�����׼��
Const SS2_PRODSPECNOA = 40 '�ബ�����ƺ�һ
Const SS2_PRODSPECNOB = 41 '�ബ�����ƺŶ�
Const SS2_PRODSPECNOC = 42 '�ബ�����ƺ���
Const SS2_INSP_CD = 43     '��֤����
Const SS2_SIDEMARK = 36 ' �������
Const SS2_SURFACE_REQUESTS = 44 ' ����ͻ�Ҫ��

Const SPD_LINE1 = 1
Const SPD_LINE2 = 2
Const SPD_ORD_FL = 3
Const SPD_PLATE_NO = 4
Const SPD_THK = 5
Const SPD_WID = 6
Const SPD_LEN = 7
Const SPD_WGT = 8
Const SPD_LAST_YN = 9
Const SPD_SIZE_KND = 10
Const SPD_TRIM_FL = 11
Const SPD_APLY_STDSPEC = 12
Const SPD_APLY_STDSPEC_NEW = 13
Const SPD_SURF_GRD = 14
Const SPD_MARK_YN = 15
Const SPD_STAMP_YN = 16
Const SPD_BAR_YN = 17
Const SPD_PROD_DATE = 18
Const SPD_EMP_CD = 19
Const SPD_PAINT = 20
Const SPD_LABEL = 21
Const SPD_STDSPEC_YY = 22
Const SPD_STLGRD = 23
Const SPD_VESSEL_NO = 24
Const SPD_COLOR_STROKE = 25
Const SPD_CE_YN = 26
Const SPD_TS_YN = 27
Const SPD_JIS_YN = 28 'ADD HANCHAO
Const SPD_UST_FL = 29 'MD HANCHAO
Const SPD_CUST_CD = 30 'MD HANCHAO
Const SPD_TO_CUR_INV = 31 ' MD HANCHAO
Const SPD_HTM_METH = 32 'MD HANCHAO
Const SPD_DEL_TO_DATE = 33 'MD HANCHAO
Const SPD_PROC_CD = 34 ' MD HANCHAO
Const SPD_CUST_PUNCH = 35 'MD HANCHAO

Dim iSS As String
Dim iF_mm As Double
Dim iT_mm As Double

 Dim sPaint As Integer
 Dim sPunch As Integer
 Dim sEdge As Integer

Dim Ship_Emblem As Integer '����������
Dim PRODSPECNOA As Integer '�ƺ�һ
Dim PRODSPECNOB As Integer '�ƺŶ�
Dim PRODSPECNOC As Integer '�ƺ���
  
'���¶ബ��������ֶ� 2013-12-31
Dim First_Number As Integer
Dim Second_Number As Integer
Dim Third_Number As Integer
Dim Firth_Number As Integer
Dim Fifth_Number As Integer
Dim Sixth_Number As Integer
'���¶ബ��������ֶ� 2013-12-31

Dim sInsp_cd As String     '��֤����
Dim sideMark As String     '�������

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
    Call Gp_Ms_Collection(TXT_PLT_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(udt_date_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(udt_date_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_line, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_STDSPEC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_rec_sts, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(cbo_shift, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
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
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '��ӡ
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '��ӡ
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '����
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2035C.P_REFER", Key:="P-R"
    sc1.Add Item:="AGC2035C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="AGC2035C.P_MODIFY", Key:="P-M"
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
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
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
    Call Gp_Sp_Collection(ss2, 35, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 36, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 37, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '�ബ�����׼һ 20140101
    Call Gp_Sp_Collection(ss2, 38, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '�ബ�����׼�� 20140101
    Call Gp_Sp_Collection(ss2, 39, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '�ബ�����׼�� 20140101
    Call Gp_Sp_Collection(ss2, 40, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '�ബ�����ƺ�һ 20140101
    Call Gp_Sp_Collection(ss2, 41, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '�ബ�����ƺŶ� 20140101
    Call Gp_Sp_Collection(ss2, 42, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '�ബ�����ƺ��� 20140101
    Call Gp_Sp_Collection(ss2, 43, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '��֤���� 20140227
    
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AGC2035C.P_SREFER", Key:="P-R"
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
          strState = "���Ӷ���"
          tcpStatus.BackColor = &HFF&
          chk_Cond(Index).ForeColor = &HFF&
          tcpMsg.Caption = "��ӡ��״̬ : " & strState
       End If
    End If
    If Index = 1 Then
       If chk_Cond(Index) = 1 Then
          Winsock3.Connect
       Else
          Winsock3.Close
          tcpStatus3.BackColor = &HFF&
          chk_Cond(Index).ForeColor = &HFF&
       End If
    End If
    If Index = 8 Then
       If chk_Cond(Index) = 1 Then
          Winsock2.Connect
       Else
          Winsock2.Close
          strState2 = "���Ӷ���"
          tcpStatus2.BackColor = &HFF&
          chk_Cond(Index).ForeColor = &HFF&
          tcpMsg2.Caption = "�����״̬ : " & strState2
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
    
    Call Gp_Sp_ColHidden(ss1, SS1_LINE2, True)
    Call Gp_Sp_ColHidden(ss1, SS1_PAINT, True)
'    Call Gp_Sp_ColHidden(ss1, SS1_LABEL, True)
    Call Gp_Sp_ColHidden(ss1, SS1_MARK_YN, True)
    Call Gp_Sp_ColHidden(ss1, SS1_STAMP_YN, True)
    Call Gp_Sp_ColHidden(ss1, SS1_BAR_YN, True)

    Call Gp_Sp_ColHidden(ss2, SPD_LINE2, True)
'    Call Gp_Sp_ColHidden(ss2, SPD_WGT, True)
    Call Gp_Sp_ColHidden(ss2, SPD_SURF_GRD, True)
    Call Gp_Sp_ColHidden(ss2, SPD_LAST_YN, True)
    Call Gp_Sp_ColHidden(ss2, SPD_SIZE_KND, True)
    Call Gp_Sp_ColHidden(ss2, SPD_TRIM_FL, True)
    Call Gp_Sp_ColHidden(ss2, SPD_APLY_STDSPEC, True)
    Call Gp_Sp_ColHidden(ss2, SPD_APLY_STDSPEC_NEW, True)
    Call Gp_Sp_ColHidden(ss2, SPD_PROD_DATE, True)
    Call Gp_Sp_ColHidden(ss2, SPD_EMP_CD, True)
    Call Gp_Sp_ColHidden(ss2, SPD_PAINT, True)
'    Call Gp_Sp_ColHidden(ss2, SPD_LABEL, True)
'    Call Gp_Sp_ColHidden(ss2, SPD_MARK_YN, True)
'    Call Gp_Sp_ColHidden(ss2, SPD_STAMP_YN, True)
'    Call Gp_Sp_ColHidden(ss2, SPD_BAR_YN, True)
    Call Gp_Sp_ColHidden(ss2, SPD_PROC_CD, True)
    Call Gp_Sp_ColHidden(ss2, SS2_PRODSPECNOA, True) '�ബ�����ƺ�һ���� 20140101
    Call Gp_Sp_ColHidden(ss2, SS2_PRODSPECNOB, True) '�ബ�����ƺŶ����� 20140101
    Call Gp_Sp_ColHidden(ss2, SS2_PRODSPECNOC, True) '�ബ�����ƺ������� 20140101
    
    txt_plt.Text = "C1"
    Call txt_plt_KeyUp(0, 0)
    Call Gf_USER_ComboAdd(M_CN1, CBO_sUserID, "AGC2035C")

    txt_line.Text = "1"
    txt_rec_sts.Text = "1"
'    TXT_Producer.Text = "NISCO"
    iSS = ""
    opt_line1 = True
    opt_line3 = True
'
'    Winsock1.RemoteHost = "172.18.56.194" 'Gf_ComnNameFind(M_CN1, "G0034", "01", 1)
'    Winsock1.RemotePort = "2222" 'Gf_ComnNameFind(M_CN1, "G0034", "01", 2)
'    Winsock2.RemoteHost = "172.18.43.113" 'Gf_ComnNameFind(M_CN1, "G0034", "01", 1)
'    Winsock2.RemotePort = "25298" 'Gf_ComnNameFind(M_CN1, "G0034", "01", 2)
'
    Winsock1.RemoteHost = Gf_ComnNameFind(M_CN1, "G0036", "01", 1)
    Winsock1.RemotePort = Gf_ComnNameFind(M_CN1, "G0036", "01", 2)
    Winsock2.RemoteHost = Gf_ComnNameFind(M_CN1, "G0038", "01", 1)
    Winsock2.RemotePort = Gf_ComnNameFind(M_CN1, "G0038", "01", 2)
    Winsock3.RemoteHost = Gf_ComnNameFind(M_CN1, "BG002", "01", 1)
    Winsock3.RemotePort = Gf_ComnNameFind(M_CN1, "BG002", "01", 2)

    iF_mm = Val(Gf_ComnNameFind(M_CN1, "G0037", "01", 1))
    iT_mm = Val(Gf_ComnNameFind(M_CN1, "G0037", "02", 1))
    
    If Mid(sAuthority, 3, 1) = "1" Then
       chk_Cond(0).Enabled = True
       chk_Cond(1).Enabled = True
       chk_Cond(8).Enabled = True
    Else
       chk_Cond(0).Enabled = False
       chk_Cond(1).Enabled = False
       chk_Cond(8).Enabled = False
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Winsock1.State = 1 Or Winsock1.State = 7 Or Winsock1.State = 9 Then
       Winsock1.Close
    End If
    
    If Winsock2.State = 1 Or Winsock2.State = 7 Or Winsock2.State = 9 Then
       Winsock2.Close
    End If
    
    If Winsock3.State = 1 Or Winsock3.State = 7 Or Winsock3.State = 9 Then
       Winsock3.Close
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
    
    Call Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False)
    
    If chk_Cond(7) Then
    Else
       Call Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False)
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call MenuTool_ReSet
    ss1.OperationMode = OperationModeNormal
    ss2.OperationMode = OperationModeNormal

    
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
            
             '����������ɫ��� 2012-11-08  by  LiQian
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
            '�������ھ�ʾ
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
    Dim sSpec As String
    Dim sStdspec_yy As String
    Dim sColorStr As String
    
    Dim sMesg_Fl As String
    Dim sRst_Len As Double
    Dim sPlan_Len As Double
    
    Dim sGroup As String
    Dim sEmp_Cd As String
    
    sMesg_Fl = "N"
    
    sGroup = Trim(CBO_GROUP.Text)
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
            SMESG = " �ְ峤����ƻ����Ȳ�һ�£���ȷ�ϱ���ô��"
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
             If .Text = "Update" Or .Text = "Insert" Or .Text = "Delete" Then
             
                .Col = SPD_PLATE_NO:         sPlate_No = .Text
                .Col = SPD_THK:              sThk = Trim(Str(.Text))
                .Col = SPD_WID:              sWid = Trim(Str(.Text))
                .Col = SPD_LEN:              sLen = Trim(Str(.Text))
                .Col = SPD_WGT:              sWgt = Trim(Str(.Text))
                If Mid(sWgt, 1, 1) = "." Then
                   sWgt = "0" & sWgt
                End If
                .Col = SPD_STLGRD:           sSpec = .Text
                .Col = SPD_STDSPEC_YY:       sStdspec_yy = .Text
                .Col = SPD_LABEL:            sColorStr = .Text
                .Col = 0
               
                If (chk_Cond(0) Or chk_Cond(1) Or chk_Cond(8)) And .Text <> "Delete" Then
                
                   'Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, iRow, iRow, , &HFFC0FF)
                   If sGroup <> "A" And sGroup <> "B" And sGroup <> "C" And sGroup <> "D" Then
                        SMESG = " ��������ȷ���Ƿ���ȷ������"
                        Call Gp_MsgBoxDisplay(SMESG)
                        Exit Sub
                   End If
                   If sEmp_Cd = "" Then
                        SMESG = " ��������ҵ��Ա���ţ�"
                        Call Gp_MsgBoxDisplay(SMESG)
                        Exit Sub
                   End If
                   Call Cmd_SEND(sPlate_No, sThk, sWid, sLen, sWgt, sSpec, sStdspec_yy, sColorStr)
                   If Gp_Send_Size_Exec(sPlate_No) <> "" Then
                      MsgBox (" �򼸺γߴ�����豸����ָʾʧ�� --�� " + Gp_Send_Size_Exec(sPlate_No))
                   End If
                '  2011-8-24  modify  by   LiQian
                '  ���±�ӡ����ӡ��������λ,���ջ���ѡ���Ƿ��ӡ����ӡ�������
                   Call Gp_Send_Paint(sPlate_No, sPaint, sPunch, sEdge, sEmp_Cd)
                '  ���²�ѯ����,�����ӡ��,��˿�ӻ�����ʧ,�޷��ٲ�ѯ��,��ֹ��ӡ���ظ������ɵ��غ�����
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
    Dim sStdspec_yy As String
    Dim sStdspec    As String
    
    Dim iCount As Integer
    
    sPlateNo = ""
    
    With ss1
        If .MaxRows = 0 Then
           If Len(TXT_PLATE_NO.Text) = 12 Then
               Call Gp_Sp_Ins(Proc_Sc("Sc1"))
              .Row = 1
              .Col = SS1_PLATE_NO
              .Text = TXT_PLATE_NO.Text & "01"
              .Col = SS1_THK:           .Value = 0
              .Col = SS1_WID:           .Value = 0
              .Col = SS1_LEN:           .Value = 0
              .Col = SS1_WGT:           .Value = 0
              .Col = SS1_APLY_STDSPEC:  .Text = "GB-XXX"
           Else
               Call Gp_MsgBoxDisplay("����ȷ����ĸ��� ��")
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
            .Col = SS1_STDSPEC_YY:    sStdspec_yy = .Text
            .Col = SS1_EMP_CD:        sEmp_Cd = .Text
            .Col = SS1_STLGRD:        sStdspec = .Text
        Else
            sPlateNo = TXT_PLATE_NO.Text & "00"
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
        .Col = SS1_STDSPEC_YY:    .Text = sStdspec_yy
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
        
        
'��ԭ�еĸ��ݿؼ��޸ı�ӡ�Ĺ���ɾ�� 2013-12-24
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
  
'  Call Cmd_SEND_SET(Row)
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
    Dim sSpec1 As String
    Dim sSpec2 As String
    Dim sSpec_Str As String
    Dim sStdspec_yy As String
    Dim sNum As Integer
    Dim sNumFL As String
    Dim sCE_YN As String
    Dim sTS_YN As String
    Dim jIS_YN As String
    Dim sCEdate As String
    
    Dim sXm As String
    
    sCEdate = Format(Now, "YY")

    If iSS = "ss1" Or iSS = "" Then
        With ss1
            .Row = Row
            .Col = SS1_PLATE_NO:         TXT_MAT_NO = .Text
            .Col = SS1_THK:              TXT_THK = Trim(Str(.Text))
            .Col = SS1_WID:              TXT_WID = Trim(Str(.Text))
            .Col = SS1_LEN:              TXT_LEN = Trim(Str(.Text))
            .Col = SS1_WGT:              TXT_WGT = Trim(Str(.Text))
            .Col = SS1_STDSPEC_YY:       TXT_SPEC = .Text
            .Col = SS1_STLGRD:           TXT_SPEC_DATE = .Text
            .Col = SS1_APLY_STDSPEC_NEW: sSpec1 = .Text
            If sSpec1 = "" Then
               .Col = SS1_APLY_STDSPEC:  sSpec1 = .Text
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
            .Col = SPD_STDSPEC_YY:       TXT_SPEC = .Text
            .Col = SPD_CE_YN:            sCE_YN = .Text
            .Col = SPD_TS_YN:            sTS_YN = .Text
            .Col = SPD_JIS_YN:           jIS_YN = .Text
            .Col = SPD_UST_FL
             If .Text = "X" Then
                 TXT_Producer.Text = ""
             Else
                 TXT_Producer.Text = " UT " & .Text
             End If
               
            .Col = SS2_PRODSPECNOA:      PRODSPECNOA = .Value   '���ƺ�һ��ֵ���浽�ֶ���
            .Col = SS2_PRODSPECNOB:      PRODSPECNOB = .Value   '���ƺŶ���ֵ���浽�ֶ���
            .Col = SS2_PRODSPECNOC:      PRODSPECNOC = .Value   '���ƺ�����ֵ���浽�ֶ���
            .Col = SS2_INSP_CD:          sInsp_cd = .Text
            
             '�ڱ����ʱ�򣬽���ӡ ��ӡ ��������ݱ��浽��Ӧ������ 2013-12-24
            .Col = SS2_sPaint:           sPaint = .Value '��ӡ
            .Col = SS2_sPunch:           sPunch = .Value '��ӡ
            .Col = SS2_sEdge:            sEdge = .Value '����
                

            .Col = SPD_HTM_METH
             If .Text <> "" Then
                 TXT_Producer.Text = TXT_Producer.Text & " " & .Text
             End If
             
            .Col = SS2_SURFACE_REQUESTS
            If .Text = "G" Then
               TXT_Producer.Text = TXT_Producer.Text & "    " & .Text
            End If
            .Col = SPD_STLGRD:           TXT_SPEC_DATE = .Text
            .Col = SPD_APLY_STDSPEC_NEW: sSpec1 = .Text
            .Col = SPD_VESSEL_NO:        TXT_X1 = .Text
            .Col = SS2_SIDEMARK:         sideMark = .Text
            .Col = SPD_CUST_PUNCH:       TXT_Punch.Text = .Text
             If jIS_YN = "D" Or jIS_YN = "E" Or jIS_YN = "F" Then
                TXT_Punch.Text = "NISCO" & " " & TXT_Punch.Text
             End If
            .Col = SPD_CUST_CD:          TXT_CUST_CD = .Text
            .Col = SPD_TO_CUR_INV:       TXT_TO_CUR_INV = .Text
            If sSpec1 = "" Then
               .Col = SPD_APLY_STDSPEC:  sSpec1 = .Text
            End If
        End With
        
        Ship_Emblem = 1
        
        '�жϴ���������
        If PRODSPECNOA <> 13 Then
            Ship_Emblem = Ship_Emblem + 1
        End If
        If PRODSPECNOB <> 13 Then
            Ship_Emblem = Ship_Emblem + 1
        End If
        If PRODSPECNOC <> 13 Then
            Ship_Emblem = Ship_Emblem + 1
        End If
        '�жϴ���������
                        
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
    sSpec = TXT_SPEC_DATE
    sStdspec_yy = TXT_SPEC
    
    sNum = InStr(sSpec1, "-")
    If sNum = 0 Then
        sNum = Len(sSpec1)
    End If
    sSpec_Str = Mid(sSpec1, 1, (sNum - 1))
    
    If sSpec1 = "ZH-ABS-A36" Then
       sSpec_Str = "ABS"
    End If
    
    Select Case sSpec_Str

           Case "BV"
                opt_line2 = True
                TXT_Class_comp = "5"
           Case "CCS"
                opt_line2 = True
                TXT_Class_comp = "2"
           Case "DNV"
                opt_line2 = True
                TXT_Class_comp = "6"
           Case "VL"
                opt_line2 = True
                TXT_Class_comp = "6"
           Case "GL"
                opt_line2 = True
                TXT_Class_comp = "4"
           Case "KR"
                opt_line2 = True
                TXT_Class_comp = "7"
           Case "LR"
                opt_line2 = True
                TXT_Class_comp = "8"
           Case "NK"
                opt_line2 = True
                TXT_Class_comp = "10"
           Case "RINA"
                opt_line2 = True
                TXT_Class_comp = "9"
           Case "ABS"
                opt_line2 = True
                TXT_Class_comp = "1"
           Case "IRS"
                opt_line2 = True
                TXT_Class_comp = "14"
           Case "RS"
                opt_line2 = True
                TXT_Class_comp = "25"
           Case Else
                opt_line1 = True
                TXT_Class_comp = "13"
                If sCE_YN = "Y" Then
                    TXT_Class_comp = "11"
                End If
                If sTS_YN = "Y" Then
                    TXT_Class_comp = "12"
                End If
                If sTS_YN = "A" Then '���߸��ж� add hanchao
                    TXT_Class_comp = "16"
                End If
                If jIS_YN = "A" Then 'JIS��ʶ�ж� add hanchao
                    TXT_Class_comp = "15"
                End If
                If jIS_YN = "D" Then 'API
                    TXT_Class_comp = "17"
                End If
                If jIS_YN = "E" Then 'API
                    TXT_Class_comp = "18"
                End If
                If jIS_YN = "F" Then 'API
                    TXT_Class_comp = "19"
                End If
                TXT_Special = ""
'                TXT_Paint.Text = TXT_X1.Text
'                TXT_X1.Text = ""
                
    End Select
    
    
    TXT_Paint.Text = TXT_X1.Text
    TXT_X1.Text = ""
    
    If TXT_Class_comp.Text = "13" Then
       TXT_Class_comp.Text = sInsp_cd
    End If
    
End Sub

Private Sub Cmd_SEND(iPlate_no As String, iThk As String, iWid As String, iLen As String, iWGT As String, iSpec As String, iStdspec_yy As String, iColorStr As String)
    
    Dim SMESG As String

    Dim Header As String * 2
    Dim sPlate_No As String * 14
    Dim sLen As Long
    Dim sWid As Long
    Dim sThk As Long
    Dim sWgt As String
    Dim sProducer As String * 50 '20150427
    Dim sSpec As String           ' ��׼��
    Dim sStdspec_yy As String     ' ��׼���
    Dim sSpec1 As String * 34     ' ��׼�ţ�����׼��ţ�
    Dim sSpec2 As String * 40     ' ����
    Dim sPaint_Cid As String * 50  '20150427
    Dim sPunch_Cid As String * 30   '20150427
    Dim sProd_Date As String * 13
    Dim sGroup As String * 2
    
    Dim sSpecial As String * 40
    Dim sX1 As String * 50  '20150427
    Dim sX2 As String * 34
    Dim sX3 As String * 17
    Dim sLayout As Integer
    Dim sTopPaint As Integer
    Dim sBottomPaint As Integer
   
        
    Dim sSpec_Str As String
    Dim PaintStr_CD As Integer
    Dim PunchStr_CD As Integer
    Dim EdgeStr_CD As Integer
    Dim EU_Down_CD As Integer
    
    Dim Repeat_Distance As Long
    Dim sNum As Integer
    Dim sNumFL As String
    
    Dim sCEdate As String
   
    sCEdate = Format(Now, "YY")
    
    Dim sSpec_Logo As String
    
    Dim sEdgeStr As String * 80
    Dim sColorStr As String
    Dim sComp As Integer
    
    Dim sXm As String
    
    sPlate_No = iPlate_no
    sThk = Val(iThk) * 100
    sWid = Val(iWid)
    sLen = Val(iLen)
    sWgt = iWGT
    sSpec = iSpec
    sSpec1 = iSpec
    sSpec2 = iSpec
    sStdspec_yy = iStdspec_yy
    sColorStr = iColorStr

    '��������ʶҪ��ı༭������Ϣ
    If iStdspec_yy = "GB 713-2008" Or iStdspec_yy = "GB 3531-2008" Or iStdspec_yy = "GB 19189-2011" Or iStdspec_yy = "GB 713-2014" Or iStdspec_yy = "GB 3531-2014" Then
        sProducer = "  T.W. " & sWgt & " t" & " " & TXT_Producer
    Else
        sProducer = TXT_Producer
    End If
    
    sPaint_Cid = TXT_Paint
    sPunch_Cid = TXT_Punch
    sSpecial = TXT_Special
    sX1 = TXT_X1
    sX2 = TXT_X2
    sX3 = TXT_X3
    
    Header = "MD"
    
    sProd_Date = udt_date_fr.RawData  '1111111111111
    sGroup = Trim(CBO_GROUP.Text)
    sComp = Val(TXT_Class_comp.Text)
    Repeat_Distance = Val(TXT_Distance)
    
    If sGroup <> "A " And sGroup <> "B " And sGroup <> "C " And sGroup <> "D " Then
        SMESG = " ��������ȷ���Ƿ���ȷ������"
        Call Gp_MsgBoxDisplay(SMESG)
        Exit Sub
    End If
    
    If TXT_Class_comp <> "" Then
        If Val(TXT_Class_comp) < 1 Or Val(TXT_Class_comp) > 50 Then
            SMESG = "Class_comp(1-50) ���ݴ�����ȷ��"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
        End If
    End If
    
    '���沼��
    If opt_line1 Then
       sLayout = 0
    End If
    If opt_line2 Then
       sLayout = 1
    End If
    
    '��ӡ���
    If opt_line5 Then
       PunchStr_CD = 0
    End If
    If opt_line6 Then
       PunchStr_CD = 1
    End If
    If opt_line7 Then
       PunchStr_CD = 2
    End If
    
    '����߶�
    If opt_line8 Then
       EdgeStr_CD = 0
    End If
    If opt_line9 Then
       EdgeStr_CD = 1
    End If
    If opt_line10 Then
       EdgeStr_CD = 2
    End If
    
    '0 = Normal reading , 1 = Up-side down
    If opt_line11 Then
       EU_Down_CD = 0
    End If
    If opt_line12 Then
       EU_Down_CD = 1
    End If
    
    
    
    
                
'                        If sPaint Then
'                            sPaint = 1
'                        Else
'                            sPaint = 0
'                        End If
'                        If sPunch Then
'                            sPunch = 1
'                        Else
'                            sPunch = 0
'                        End If
'                        If sEdge Then
'                            sEdge = 1
'                        Else
'                            sEdge = 0
'                        End If
                
                
'  '�ڱ����ʱ�򣬽���ӡ ��ӡ ��������ݱ��浽��Ӧ������ 2013-12-24


'    If chk_Cond(5) Then
'           sTopPaint = 1
'    Else
'           sTopPaint = 0
'    End If
'    If chk_Cond(6) Then
'           sBottomPaint = 1
'    Else
'           sBottomPaint = 0
'    End If
    
    sSpec1 = sStdspec_yy
    sSpec2 = sSpec
    If sComp <> 11 And sComp <> 12 And sComp <> 15 And sComp <> 16 And (sComp <> 13 And sInsp_cd = 13) Then
       sSpec2 = sStdspec_yy
    End If
    
'    ss2.Row = ss2.ActiveRow
'    ss2.Col = SPD_CE_YN
'    If ss2.Text = "Y" Then
'        sSpec1 = sStdspec_yy
'    End If

    ss2.Row = ss2.ActiveRow
    ss2.Col = SS2_SURFACE_REQUESTS: sXm = ss2.Text
    
    If chk_Cond(8) = 1 Then
            
'        If chk_Cond(9) = 1 Then
           'sEdgeStr = sPlate_No & " " & iTHK & "X" & iWID & "X" & iLEN & " " & Trim(sSpec2) & "  " & sideMark
           sEdgeStr = sPlate_No & " " & Trim(sSpec2) & "  " & iThk & "X" & iWid & "X" & iLen & " " & sideMark
'        Else
'           sEdgeStr = sPlate_No & " " & iThk & "X" & iWid & "X" & iLen & " " & sSpec2
'        End If
        
        sEdgeStr = Trim(sEdgeStr) & " " & TXT_CUST_CD & " " & TXT_TO_CUR_INV
        If sXm = "G" Then
           sEdgeStr = sXm & "    " & Trim(sEdgeStr)
        Else
           sEdgeStr = Trim(sEdgeStr)
        End If
        
        Winsock2.SendData sEdgeStr
        
    End If
        
    If chk_Cond(1) = 1 And sColorStr <> "" Then
        
        Winsock3.SendData sColorStr
        
    End If
    
    '�ബ���������ж�
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
'    ElseIf Ship_Emblem.Text = 5 Then
'    Ship_Emblem = 5
'    sComp = 0
'    Sixth_Number = 0
'    ElseIf Ship_Emblem.Text = 6 Then
'    Ship_Emblem = 6
'    sComp = 0
    End If
    sBottomPaint = 0
    
    If chk_Cond(0) = 1 Then
    
        Winsock1.SendData Header & "  "
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
        Winsock1.SendData Chr(14) & Chr(14) & sPlate_No _
                        & Chr(50) & Chr(Len(sProducer)) & sProducer _
                        & Chr(20) & Chr(Len(Trim(sSpec2))) & sSpec2 _
                        & Chr(34) & Chr(Len(Trim(sSpec1))) & sSpec1 _
                        & Chr(50) & Chr(Len(Trim(sPaint_Cid))) & sPaint_Cid _
                        & Chr(30) & Chr(Len(Trim(sPunch_Cid))) & sPunch_Cid _
                        & Chr(13) & Chr(Len(Trim(sProd_Date))) & sProd_Date _
                        & sGroup
        Winsock1.SendData HiByte(sComp)
        Winsock1.SendData LoByte(sComp)
        Winsock1.SendData Chr(12) & Chr(Len(Trim(sSpecial))) & sSpecial _
                        & Chr(50) & Chr(Len(Trim(sX1))) & sX1 _
                        & Chr(34) & Chr(Len(Trim(sX2))) & sX2 _
                        & Chr(17) & Chr(Len(Trim(sX3))) & sX3
        Winsock1.SendData HiByte(sLayout)
        Winsock1.SendData LoByte(sLayout)
        Winsock1.SendData HiByte(sPaint) '(sTopPaint) ʹ��sPaint�������� 20140101
        Winsock1.SendData LoByte(sPaint) '(sTopPaint) ʹ��sPaint�������� 20140101
        Winsock1.SendData HiByte(sBottomPaint) '(sBottomPaint) ʹ��sPaint�������� 20140101
        Winsock1.SendData LoByte(sBottomPaint) '(sBottomPaint) ʹ��sPaint�������� 20140101
        Winsock1.SendData HiByte(0)  'No of Repetitions
        Winsock1.SendData LoByte(0)  'No of Repetitions
        Winsock1.SendData HLByte(Repeat_Distance, 3) 'Repeat Distance
        Winsock1.SendData HLByte(Repeat_Distance, 2) 'Repeat Distance
        Winsock1.SendData HLByte(Repeat_Distance, 1) 'Repeat Distance
        Winsock1.SendData HLByte(Repeat_Distance, 0) 'Repeat Distance
        Winsock1.SendData HiByte(sPunch)
        Winsock1.SendData LoByte(sPunch)
        Winsock1.SendData HiByte(PunchStr_CD)
        Winsock1.SendData LoByte(PunchStr_CD)
        Winsock1.SendData HiByte(sEdge)
        Winsock1.SendData LoByte(sEdge)
        Winsock1.SendData HiByte(EdgeStr_CD)
        Winsock1.SendData LoByte(EdgeStr_CD)
        Winsock1.SendData HiByte(EU_Down_CD)
        Winsock1.SendData LoByte(EU_Down_CD)
        Winsock1.SendData HiByte(Ship_Emblem) '���ٸ�����
        Winsock1.SendData LoByte(Ship_Emblem) '���ٸ�����
        Winsock1.SendData HiByte(First_Number) '��һ������
        Winsock1.SendData LoByte(First_Number) '��һ������
        Winsock1.SendData HiByte(Second_Number) '�ڶ���
        Winsock1.SendData LoByte(Second_Number) '�ڶ���
        Winsock1.SendData HiByte(Third_Number) '������
        Winsock1.SendData LoByte(Third_Number) '������
        Winsock1.SendData HiByte(Firth_Number) '���ĸ�
        Winsock1.SendData LoByte(Firth_Number) '���ĸ�
        Winsock1.SendData HiByte(Fifth_Number) '�����
        Winsock1.SendData LoByte(Fifth_Number) '�����
        Winsock1.SendData HiByte(Sixth_Number) '������
        Winsock1.SendData LoByte(Sixth_Number) '������
        
    
    End If
    
'    SOCK.Text = Header & "  " & HLByte(Val(sLen), 3) & HLByte(Val(sLen), 2) & HLByte(Val(sLen), 1) & HLByte(Val(sLen), 0) _
'                              & HLByte(Val(sWid), 3) & HLByte(Val(sWid), 2) & HLByte(Val(sWid), 1) & HLByte(Val(sWid), 0) _
'                              & HLByte(Val(sThk), 3) & HLByte(Val(sThk), 2) & HLByte(Val(sThk), 1) & HLByte(Val(sThk), 0) _
'                              & Chr(14) & Chr(14) & sPlate_no & Chr(31) & Chr(31) & sProducer & Chr(14) & Chr(14) & sSpec2 _
'                              & Chr(34) & Chr(34) & sSpec1 & Chr(17) & Chr(17) & sPaint_Cid & Chr(18) & Chr(18) & sPunch_Cid _
'                              & Chr(13) & Chr(13) & sProd_Date & sGroup & HiByte(Val(sComp)) & LoByte(Val(sComp)) _
'                              & Chr(12) & Chr(12) & sSpecial & Chr(29) & Chr(29) & sX1 & Chr(34) & Chr(34) & sX2 & Chr(17) & Chr(17) & sX3 _
'                              & HiByte(Val(sLayout)) & LoByte(Val(sLayout)) & HiByte(Val(sTopPaint)) & LoByte(Val(sTopPaint)) _
'                              & HiByte(Val(sBottomPaint)) & LoByte(Val(sBottomPaint)) & HiByte(0) & LoByte(0) _
'                              & HLByte(10, 3) & HLByte(10, 2) & HLByte(10, 1) & HLByte(10, 0) & HiByte(sPaint) & LoByte(sPaint) _
'                              & HiByte(sPunch) & LoByte(sPunch) & HiByte(sEdge) & LoByte(sEdge) & HiByte(EdgeStr_CD) & LoByte(EdgeStr_CD) & HiByte(0) & LoByte(0)
                              
      
End Sub

Private Sub Timer1_Timer()

    'sckClosed            0 ȱʡ�ġ�--�ر� û�е�
    'sckOpen              1 �� --�򿪵�
    'sckListening         2 ���� --�쿴��û����������
    'sckConnectionPending 3 ���ӹ���
    'sckResolvingHost     4 ʶ������
    'sckHostResolved      5 ��ʶ������
    'sckConnecting        6 ��������
    'sckConnected         7 ������
    'sckClosing           8 ͬ����Ա���ڹر����� -˵���Է��ر���������
    'sckError             9 ����
    
    Dim strState As String
    Dim strState2 As String
    
    If chk_Cond(0) <> 1 And chk_Cond(1) <> 1 And chk_Cond(8) <> 1 Then
       Exit Sub
    Else
    
        If chk_Cond(0) = 1 Then
        
            Select Case Winsock1.State
                Case 0
                    strState = "���ӹر�"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
                Case 1
                    strState = "���Ӵ�"
                Case 2
                    strState = "���ӱ���"
                Case 3
                    strState = "Close"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
                Case 4
                    strState = "Find Host...."
                Case 5
                    strState = "Finded Host"
                Case 6
                    strState = "��������"
                Case 7
                    strState = "��������"
                    tcpStatus.BackColor = &HC000&
                    chk_Cond(0).ForeColor = &HC000&
                Case 8
                    strState = "���Ӷ���"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
                Case 9
                    strState = "���Ӵ���"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
            Case Else
                strState = "StateNum:" & Winsock1.State
                tcpStatus.BackColor = &HFF&
                chk_Cond(0).ForeColor = &HFF&
            End Select

            tcpMsg.Caption = "��ӡ��״̬ : " & strState
            
        End If
        
        If chk_Cond(1) = 1 Then
        
            Select Case Winsock3.State
                Case 0
                    tcpStatus3.BackColor = &HFF&
                    chk_Cond(1).ForeColor = &HFF&
                Case 1
                Case 2
                Case 3
                    tcpStatus3.BackColor = &HFF&
                    chk_Cond(1).ForeColor = &HFF&
                Case 4
                Case 5
                Case 6
                Case 7
                    tcpStatus3.BackColor = &HC000&
                    chk_Cond(1).ForeColor = &HC000&
                Case 8
                    tcpStatus3.BackColor = &HFF&
                    chk_Cond(1).ForeColor = &HFF&
                Case 9
                    strState = "���Ӵ���"
                    tcpStatus3.BackColor = &HFF&
                    chk_Cond(1).ForeColor = &HFF&
            Case Else
                tcpStatus3.BackColor = &HFF&
                chk_Cond(1).ForeColor = &HFF&
            End Select
            
        End If
        
        If chk_Cond(8) = 1 Then

            Select Case Winsock2.State
                Case 0
                    strState2 = "���ӹر�"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(8).ForeColor = &HFF&
                Case 1
                    strState2 = "���Ӵ�"
                Case 2
                    strState2 = "���ӱ���"
                Case 3
                    strState2 = "Close"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(8).ForeColor = &HFF&
                Case 4
                    strState2 = "Find Host...."
                Case 5
                    strState2 = "�ҵ�����"
                Case 6
                    strState2 = "��������"
                Case 7
                    strState2 = "��������"
                    tcpStatus2.BackColor = &HC000&
                    chk_Cond(8).ForeColor = &HC000&
                Case 8
                    strState2 = "���Ӷ���"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(8).ForeColor = &HFF&
                Case 9
                    strState2 = "���Ӵ���"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(8).ForeColor = &HFF&
            Case Else
                strState2 = "StateNum:" & Winsock2.State
                tcpStatus2.BackColor = &HFF&
                chk_Cond(8).ForeColor = &HFF&
            End Select

            tcpMsg2.Caption = "�����״̬ : " & strState2

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
    
        If Col = SS1_THK Or Col = SPD_WID Or Col = SS1_LEN Then
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
        
'��ԭ�еĸ��ݿؼ��޸ı�ӡ�Ĺ���ɾ�� 2013-12-24
        
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
        DD.rControl.Add Item:=TXT_PLT_NAME

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

    Else

        If Len(Trim(txt_plt)) = txt_plt.MaxLength Then
            TXT_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
        Else
            TXT_PLT_NAME.Text = ""
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
    
    DD.sQuery = "SELECT CD_SHORT_NAME ""��׼����"", CD_NAME ""��׼������"" FROM ZP_CD WHERE CD_MANA_NO = 'G0035'"
    
    Call Gf_DD_Display(Conn, DD.sQuery, False)
    
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function


Private Sub txt_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        TXT_STDSPEC.Text = ""
        DD.rControl.Add Item:=TXT_STDSPEC

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
    sQuery = sQuery + "      AND UPD   = '1' AND EMP_ID <> '1JS1005'"
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
    sQuery = "{call AGC2035P ('" + sPlate_No + "',?)}"

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

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Send_Paint
'   2.Name         :
'   3.Input  Value :
'   4.Return Value : Boolean
'   5.Writer       : Li Qian
'   6.Create Date  : 2011. 08 .24
'   7.Modify Date  :
'   8.Comment      : �Ƿ��ʶ��¼����
'---------------------------------------------------------------------------------------

Public Function Gp_Send_Paint(sPlate_No As String, sMark As Integer, sStamp As Integer, sBar As Integer, sEmpCD As String) As String

On Error GoTo Process_Exec_ERROR

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

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Err.Raise Err.Number, Err.Description
    
End Function

