VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2200C 
   Caption         =   "�ְ�ʵ����ѯ����_AGC2200C"
   ClientHeight    =   10680
   ClientLeft      =   15
   ClientTop       =   1740
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10680
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9165
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   16166
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      Locked          =   -1  'True
      PaneTree        =   "AGC2200C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   6255
         Left            =   0
         TabIndex        =   1
         Top             =   2910
         Width           =   15255
         _Version        =   393216
         _ExtentX        =   26908
         _ExtentY        =   11033
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   74
         MaxRows         =   5
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AGC2200C.frx":0072
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   765
         Left            =   0
         TabIndex        =   2
         Top             =   2085
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   1349
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.TextBox TXT_4_WGT 
            Alignment       =   1  'Right Justify
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
            Left            =   9630
            TabIndex        =   40
            Top             =   360
            Width           =   1125
         End
         Begin VB.TextBox TXT_4_NUM 
            Alignment       =   1  'Right Justify
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
            Left            =   8835
            TabIndex        =   39
            Top             =   360
            Width           =   780
         End
         Begin VB.TextBox TXT_2_WGT 
            Alignment       =   1  'Right Justify
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
            Left            =   5355
            TabIndex        =   38
            Top             =   360
            Width           =   1140
         End
         Begin VB.TextBox TXT_2_NUM 
            Alignment       =   1  'Right Justify
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
            Left            =   4575
            TabIndex        =   37
            Top             =   360
            Width           =   780
         End
         Begin VB.TextBox TXT_1_WGT 
            Alignment       =   1  'Right Justify
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
            Left            =   3240
            TabIndex        =   36
            Top             =   360
            Width           =   1140
         End
         Begin VB.TextBox TXT_1_NUM 
            Alignment       =   1  'Right Justify
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
            Left            =   2460
            TabIndex        =   35
            Top             =   360
            Width           =   780
         End
         Begin VB.TextBox TXT_TOT_WGT 
            Alignment       =   1  'Right Justify
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
            Left            =   1005
            TabIndex        =   34
            Top             =   360
            Width           =   1245
         End
         Begin VB.TextBox TXT_TOT_NUM 
            Alignment       =   1  'Right Justify
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
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   870
         End
         Begin VB.TextBox TXT_3_NUM 
            Alignment       =   1  'Right Justify
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
            Left            =   6705
            TabIndex        =   32
            Top             =   360
            Width           =   780
         End
         Begin VB.TextBox TXT_3_WGT 
            Alignment       =   1  'Right Justify
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
            Left            =   7485
            TabIndex        =   31
            Top             =   360
            Width           =   1140
         End
         Begin VB.TextBox TXT_5_NUM 
            Alignment       =   1  'Right Justify
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
            Left            =   10965
            TabIndex        =   30
            Top             =   360
            Width           =   780
         End
         Begin VB.TextBox TXT_5_WGT 
            Alignment       =   1  'Right Justify
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
            Left            =   11745
            TabIndex        =   29
            Top             =   360
            Width           =   1125
         End
         Begin VB.TextBox TXT_7_WGT 
            Alignment       =   1  'Right Justify
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
            Left            =   13875
            TabIndex        =   28
            Top             =   360
            Width           =   1125
         End
         Begin VB.TextBox TXT_7_NUM 
            Alignment       =   1  'Right Justify
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
            Left            =   13095
            TabIndex        =   27
            Top             =   360
            Width           =   780
         End
         Begin InDate.ULabel ULabel13 
            Height          =   315
            Left            =   130
            Top             =   30
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            Caption         =   "�ְ���/������(Ton)"
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
         Begin InDate.ULabel ULabel14 
            Height          =   315
            Left            =   2470
            Top             =   30
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            Caption         =   "��Ʒ��/����(Ton)"
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
         Begin InDate.ULabel ULabel18 
            Height          =   315
            Left            =   4590
            Top             =   30
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            Caption         =   "������/����(Ton)"
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
         Begin InDate.ULabel ULabel19 
            Height          =   315
            Left            =   8845
            Top             =   30
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            Caption         =   "������/����(Ton)"
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
         Begin InDate.ULabel ULabel21 
            Height          =   315
            Left            =   6720
            Top             =   30
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            Caption         =   "Э����/����(Ton)"
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
         Begin InDate.ULabel ULabel23 
            Height          =   315
            Left            =   10965
            Top             =   30
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            Caption         =   "��Ʒ��/����(Ton)"
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
         Begin InDate.ULabel ULabel24 
            Height          =   315
            Left            =   13095
            Top             =   30
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            Caption         =   "�з���/����(Ton)"
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
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   2025
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   3572
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.ComboBox CBO_GROUP 
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
            ItemData        =   "AGC2200C.frx":46F2
            Left            =   7080
            List            =   "AGC2200C.frx":4702
            TabIndex        =   52
            Top             =   90
            Width           =   735
         End
         Begin VB.TextBox txt_mill_stdspec 
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
            Left            =   9555
            MaxLength       =   18
            TabIndex        =   51
            Tag             =   "��׼��"
            Top             =   1650
            Width           =   2175
         End
         Begin VB.TextBox TXT_ENDUSE_CD 
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
            Left            =   6180
            MaxLength       =   3
            TabIndex        =   50
            Top             =   1650
            Width           =   885
         End
         Begin VB.TextBox txt_org_ord_item 
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
            MaxLength       =   2
            TabIndex        =   49
            Top             =   1650
            Width           =   645
         End
         Begin VB.TextBox txt_org_ord_no 
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
            Left            =   1365
            MaxLength       =   11
            TabIndex        =   48
            Top             =   1650
            Width           =   1680
         End
         Begin VB.TextBox txt_proc_cd 
            Alignment       =   1  'Right Justify
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
            Left            =   13545
            TabIndex        =   47
            Top             =   840
            Width           =   1215
         End
         Begin VB.ComboBox CBO_PRC_LINE 
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
            ItemData        =   "AGC2200C.frx":4712
            Left            =   6180
            List            =   "AGC2200C.frx":471C
            TabIndex        =   45
            Top             =   1260
            Width           =   885
         End
         Begin VB.TextBox TXT_SMP_NO 
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
            Left            =   6180
            TabIndex        =   44
            Top             =   870
            Width           =   1785
         End
         Begin VB.TextBox txt_stlgrd 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1365
            MaxLength       =   12
            TabIndex        =   43
            Tag             =   "����"
            Top             =   870
            Width           =   1335
         End
         Begin VB.TextBox txt_STLGRD_Name 
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
            Left            =   2730
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   42
            Tag             =   "����(��׼��)"
            Top             =   870
            Width           =   1845
         End
         Begin VB.ComboBox SLAB_THK 
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
            ItemData        =   "AGC2200C.frx":4728
            Left            =   9555
            List            =   "AGC2200C.frx":473B
            TabIndex        =   41
            Top             =   1260
            Width           =   975
         End
         Begin VB.ComboBox CBO_PRODGRD 
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
            ItemData        =   "AGC2200C.frx":4758
            Left            =   2760
            List            =   "AGC2200C.frx":476E
            TabIndex        =   9
            Tag             =   "�ȼ�"
            Top             =   1260
            Width           =   1365
         End
         Begin VB.TextBox TXT_PLATE_NO 
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
            Left            =   6180
            TabIndex        =   8
            Top             =   480
            Width           =   1785
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
            Left            =   1365
            MaxLength       =   18
            TabIndex        =   7
            Tag             =   "��׼��"
            Top             =   480
            Width           =   3195
         End
         Begin VB.ComboBox CBO_SURFGRD 
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
            ItemData        =   "AGC2200C.frx":47A8
            Left            =   1365
            List            =   "AGC2200C.frx":47C1
            TabIndex        =   6
            Tag             =   "�ȼ�"
            Top             =   1260
            Width           =   1365
         End
         Begin VB.TextBox TXT_SP_CD 
            Height          =   270
            Left            =   14970
            TabIndex        =   5
            Top             =   510
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox CBO_SHIFT 
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
            ItemData        =   "AGC2200C.frx":4804
            Left            =   6180
            List            =   "AGC2200C.frx":4811
            TabIndex        =   4
            Top             =   90
            Width           =   885
         End
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   120
            Top             =   90
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "����ʱ��"
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
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   4935
            Top             =   90
            Width           =   1215
            _ExtentX        =   2143
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
         Begin Threed.SSOption OPT_SLAB 
            Height          =   330
            Left            =   13800
            TabIndex        =   10
            Top             =   450
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   2
            ForeColor       =   8421504
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "����"
         End
         Begin Threed.SSOption OPT_PLATE 
            Height          =   330
            Left            =   12570
            TabIndex        =   11
            Top             =   450
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   2
            ForeColor       =   8421504
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "����"
         End
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Index           =   1
            Left            =   120
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "��׼��"
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
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   4935
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "��ѯ��"
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
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   120
            Top             =   1260
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "����/�ۺ�"
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
         Begin Threed.SSCommand ssc_cmd 
            Height          =   315
            Left            =   12300
            TabIndex        =   12
            Top             =   90
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   556
            _Version        =   196609
            Caption         =   "������ʵ��¼��"
         End
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   8310
            Top             =   90
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "���"
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
         Begin CSTextLibCtl.sidbEdit SDB_THK 
            Height          =   315
            Left            =   9555
            TabIndex        =   13
            Top             =   90
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            FmtThousands    =   0
            FmtControl      =   1
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   8310
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "����"
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
         Begin CSTextLibCtl.sidbEdit SDB_WID 
            Height          =   315
            Left            =   9555
            TabIndex        =   14
            Top             =   480
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            FmtThousands    =   0
            FmtControl      =   1
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_THK_TO 
            Height          =   315
            Left            =   10800
            TabIndex        =   15
            Top             =   90
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            FmtThousands    =   0
            FmtControl      =   1
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_WID_TO 
            Height          =   315
            Left            =   10800
            TabIndex        =   16
            Top             =   480
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            FmtThousands    =   0
            FmtControl      =   1
            NumDecDigits    =   2
            NumIntDigits    =   4
            ShowZero        =   0   'False
            MaxValue        =   9999.99
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.UDate SDT_PROD_DATE_FROM 
            Height          =   315
            Left            =   1365
            TabIndex        =   17
            Tag             =   "��ʼ����"
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
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
         End
         Begin InDate.UDate SDT_PROD_DATE_TO 
            Height          =   315
            Left            =   3120
            TabIndex        =   18
            Tag             =   "��ʼ����"
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
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
         End
         Begin InDate.ULabel ULabel10 
            Height          =   315
            Left            =   8310
            Top             =   870
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "����"
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
         Begin CSTextLibCtl.sidbEdit SDB_LEN 
            Height          =   315
            Left            =   9555
            TabIndex        =   19
            Top             =   870
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            FmtThousands    =   0
            FmtControl      =   1
            NumDecDigits    =   1
            NumIntDigits    =   8
            ShowZero        =   0   'False
            MaxValue        =   99999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit SDB_LEN_TO 
            Height          =   315
            Left            =   10800
            TabIndex        =   20
            Top             =   870
            Width           =   975
            _Version        =   262145
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            FmtThousands    =   0
            FmtControl      =   1
            NumDecDigits    =   1
            NumIntDigits    =   8
            ShowZero        =   0   'False
            MaxValue        =   99999.9
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   8310
            Top             =   1260
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "�������"
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
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   120
            Top             =   870
            Width           =   1215
            _ExtentX        =   2143
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
            ForeColor       =   16711680
         End
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   4935
            Top             =   870
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "������"
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
         Begin Threed.SSPanel SSPpdt 
            Height          =   315
            Left            =   10800
            TabIndex        =   46
            Top             =   1260
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   255
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "��������"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   12300
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
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
         End
         Begin InDate.ULabel ULabel29 
            Height          =   315
            Left            =   120
            Top             =   1650
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "ԭʼ����"
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
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Left            =   4935
            Top             =   1260
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "������"
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
         Begin InDate.ULabel ULabel15 
            Height          =   315
            Left            =   4935
            Top             =   1650
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "������;"
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   8310
            Top             =   1650
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "���Ʊ�׼"
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
         Begin Threed.SSPanel SSPANL2 
            Height          =   315
            Left            =   12210
            TabIndex        =   53
            Top             =   1260
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   16711680
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "���轻��"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   315
            Left            =   13620
            TabIndex        =   54
            Top             =   1260
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   32768
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "��������"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   10620
            TabIndex        =   24
            Top             =   600
            Width           =   195
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   10620
            TabIndex        =   23
            Top             =   210
            Width           =   195
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   2910
            TabIndex        =   22
            Top             =   210
            Width           =   195
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   10620
            TabIndex        =   21
            Top             =   990
            Width           =   195
         End
      End
   End
   Begin CSTextLibCtl.sitxEdit SDT_PROD_DATE 
      Height          =   315
      Left            =   1110
      TabIndex        =   25
      Top             =   2130
      Width           =   1200
      _Version        =   262145
      _ExtentX        =   2117
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__ __-__-__"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Text            =   "____-__-__"
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
      Mask            =   "____-__-__"
      Justification   =   1
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   0
      ValidateMask    =   0   'False
   End
   Begin CSTextLibCtl.sitxEdit SDT_PROD_TO_DATE 
      Height          =   315
      Left            =   3150
      TabIndex        =   26
      Top             =   2160
      Width           =   1200
      _Version        =   262145
      _ExtentX        =   2117
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__ __-__-__"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Text            =   "____-__-__"
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
      Mask            =   "____-__-__"
      Justification   =   1
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   0
      ValidateMask    =   0   'False
   End
End
Attribute VB_Name = "AGC2200C"
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
'-- Program Name      �ְ�ʵ����ѯ����
'-- Program ID        AGC2200C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
'-- Date              2003.7.23
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
Public sDateTime As String          'Active Form Time Setting
Public sQuery_load As String        'Active Form sQuery Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

'Dim pControl1 As New Collection      'Master Primary Key Collection
'Dim nControl1 As New Collection      'Master Necessary Collection
'Dim mControl1 As New Collection      'Master Maxlength check Collection
'Dim iControl1 As New Collection      'Master Insert Collection
'Dim rControl1 As New Collection      'Master Refer Collection
'Dim cControl1 As New Collection      'Master Copy Collection
'Dim aControl1 As New Collection      'Master -> Spread Collection
'Dim lControl1 As New Collection      'Master Lock Collection

Dim pColumn  As New Collection      'Spread Primary Key Collection
Dim nColumn  As New Collection      'Spread necessary Column Collection
Dim mColumn  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn  As New Collection      'Spread Insert Column Collection
Dim aColumn  As New Collection      'Master -> Spread Column Collection
Dim lColumn  As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
'Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2


Const SS1_FIRST_REMARK = 3
Const SPD_PROC_CD = 4
Const SPD_DEL_TO_DATE = 5

Const SPD_WGT = 22
Const SPD_PROD_GRD = 23
Const SPD_SURF_GRD = 24
Const SPD_EDT1 = 27
Const SPD_EDT2 = 28
Const SPD_EDT3 = 29
Const SPD_EDT4 = 30
Const SPD_EDT5 = 31
Const SPD_USERID = 47
Const SPD_REN_NAME = 70
Const SPD_DATE_YN = 71    '54 -> 55
Const SPD_LAST_YN = 74    '57 -> 58

Const SPD_SHOT_BLAST = 52
Const SPD_HTM = 53
Const SPD_URGNT_FL = 40

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(cbo_shift, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(CBO_SURFGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(CBO_PRODGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_stdspec_chg, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_SP_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_THK, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_WID, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_THK_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_WID_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_LEN, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_LEN_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SLAB_THK, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TXT_STLGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(CBO_PRC_LINE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_PROC_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_ORG_ORD_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) ' Add by Li Qian at 2012-08-02 ԭʼ�����Ų�ѯ
 Call Gp_Ms_Collection(txt_ORG_ORD_ITEM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_ENDUSE_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) ' Add by Li Qian at 2012-09-14 ԭʼ������;
 Call Gp_Ms_Collection(txt_mill_stdspec, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
        Mc1.Add Item:=pControl, Key:="pControl"
        Mc1.Add Item:=nControl, Key:="nControl"
        Mc1.Add Item:=mControl, Key:="mControl"
        Mc1.Add Item:=iControl, Key:="iControl"
        Mc1.Add Item:=rControl, Key:="rControl"
        Mc1.Add Item:=cControl, Key:="cControl"
        Mc1.Add Item:=aControl, Key:="aControl"
        Mc1.Add Item:=lControl, Key:="lControl"
        
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '�׼���ʶ
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)  '��ʶ����
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)  '��������
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)  '�������
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)  '��ӡ����
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 53, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 58, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 59, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 60, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 61, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 62, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 63, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 64, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 65, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 66, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) 'add 2014 03
    Call Gp_Sp_Collection(ss1, 67, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 68, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 69, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 70, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 71, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 72, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 73, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 74, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    
   
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2200C.P_REFER", Key:="P-R"
    sc1.Add Item:="AGC2200C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AGC2200C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"

     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0
     
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



Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
End Sub

Private Sub ssc_cmd_Click()
    AGC2033C.Show 1
'    AGC2033C.SetFocus
End Sub

'���Ʊ�׼˫���ؼ�
Private Sub txt_mill_stdspec_DblClick()
Call txt_stdspec_chg_KeyUp(vbKeyF4, 0)
End Sub

'���Ʊ�׼F4����
Private Sub txt_mill_stdspec_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec_chg

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

Private Sub TXT_PLATE_NO_Change()
   Dim SMESG As String
      If Len(TXT_PLATE_NO.Text) > 14 Then
      SMESG = "�����ų��Ȳ��ܳ���10λ����ȷ�ϰ����� ������"
      Call Gp_MsgBoxDisplay(SMESG)
   End If
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        If Len(TXT_PLATE_NO.Text) >= 8 Then
           Call Form_Ref
        End If
'        KeyAscii = 0
'        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    OPT_PLATE.Value = True
    
    SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
    SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set iColumn = Nothing
    Set pColumn = Nothing
    Set lColumn = Nothing
    Set nColumn = Nothing
    Set mColumn = Nothing
    Set aColumn = Nothing

    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub
Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub
Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
       Call Gp_Ms_Cls(Mc1("rControl"))
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If

End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
     End If

End Sub

Public Sub Form_Ref()
    
    Dim SMESG As String
    Dim iCount As Long
    
    Dim sTotnum As Double
    Dim sTotwgt As Double
    Dim s1num As Double
    Dim s1wgt As Double
    Dim s2num As Double
    Dim s2wgt As Double
    Dim s3num As Double
    Dim s3wgt As Double
    Dim s4num As Double
    Dim s4wgt As Double
    Dim s5num As Double
    Dim s5wgt As Double
    Dim s7num As Double
    Dim s7wgt As Double
    Dim sWgt As Double
    
    Dim sCurDate As String
    Dim sDel_To_Date As String
    Dim sproc_cd As String
    
    Dim sShot As String
    Dim sHtm As String
    Dim sURGNT As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Val(SDT_PROD_DATE_FROM.RawData) - Val(SDT_PROD_DATE_TO.RawData) > 0 Then
         SMESG = " ʱ�䷶Χ�����������������ʱ����Ϣ ������"
         Call Gp_MsgBoxDisplay(SMESG)
         Exit Sub
    End If
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If

    With ss1
    
        If .MaxRows < 1 Then
           Exit Sub
        End If
        
        sCurDate = Format(Now, "YYYYMM")
        
        For iCount = 1 To .MaxRows
            .Row = iCount:            .Col = SPD_WGT:        sWgt = Val(.Text):   .Col = SPD_SURF_GRD
            sTotnum = sTotnum + 1:    sTotwgt = sTotwgt + sWgt
            If .Text = "��Ʒ" Then
                s1num = s1num + 1:    s1wgt = s1wgt + sWgt
            End If
            If .Text = "����" Then
                s2num = s2num + 1:    s2wgt = s2wgt + sWgt
            End If
            If .Text = "Э��" Then
                s3num = s3num + 1:    s3wgt = s3wgt + sWgt
            End If
            If .Text = "����" Then
                s4num = s4num + 1:    s4wgt = s4wgt + sWgt
            End If
            If .Text = "��Ʒ" Then
                s5num = s5num + 1:    s5wgt = s5wgt + sWgt
            End If
            If .Text = "��Ʒ" Then
                s7num = s7num + 1:    s7wgt = s7wgt + sWgt
            End If
            '�������ھ�ʾ
            .Row = iCount:            .Col = SPD_DEL_TO_DATE
            sDel_To_Date = Mid(.Value, 1, 6)
            If sDel_To_Date < sCurDate Then
              .Row = iCount:           .Col = SPD_PROC_CD
              sproc_cd = Mid(.Text, 1, 1)
              If sproc_cd <> "X" Then
                 Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iCount, iCount, &HFF&)
                 .Row = iCount:       .Col = SPD_DATE_YN:
                 .Text = "Y"
              End If
            End If
            
            '���轻����ʾ
            .Row = iCount:            .Col = SPD_SHOT_BLAST
             sShot = .Text
            .Row = iCount:            .Col = SPD_HTM
             sHtm = .Text
              If sShot <> "N/N" And sHtm = "/ / /" Then
                 Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iCount, iCount, SSPANL2.ForeColor)
              End If
              
              '�Ƿ����������ʾ
            .Row = iCount:            .Col = SPD_URGNT_FL
             sURGNT = .Text
              If sURGNT = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iCount, iCount, SSPanel1.ForeColor)
              End If
            
        Next iCount
        
    End With
    
    TXT_TOT_NUM = sTotnum:      TXT_TOT_WGT = Round(sTotwgt, 3)
    TXT_1_NUM = s1num:          TXT_1_WGT = Round(s1wgt, 3)
    TXT_2_NUM = s2num:          TXT_2_WGT = Round(s2wgt, 3)
    TXT_3_NUM = s3num:          TXT_3_WGT = Round(s3wgt, 3)
    TXT_4_NUM = s4num:          TXT_4_WGT = Round(s4wgt, 3)
    TXT_5_NUM = s5num:          TXT_5_WGT = Round(s5wgt, 3)
    TXT_7_NUM = s7num:          TXT_7_WGT = Round(s7wgt, 3)

End Sub
Public Sub Form_Pro()

    Dim iCount      As Integer
    Dim sPlateNo    As String
    
    Dim inum As Integer
    Dim lRow As Integer
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))

End Sub

Private Sub SDT_PROD_DATE_FROM_GotFocus()
     If SDT_PROD_DATE_FROM.RawData = "" Then
        SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
     End If
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub SDT_PROD_DATE_TO_GotFocus()
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub
Private Sub OPT_SLAB_Click(Value As Integer)
    Dim iRow As Integer
    Dim sTemp As String

    If OPT_SLAB.Value = True Then
        OPT_SLAB.ForeColor = &HFF&
        OPT_PLATE.ForeColor = &H808080
        TXT_SP_CD = "S"
    Else
        OPT_SLAB.ForeColor = &H808080
        TXT_SP_CD = "P"
    End If

End Sub

Private Sub OPT_PLATE_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String

    If OPT_PLATE.Value = True Then
        OPT_PLATE.ForeColor = &HFF&
        OPT_SLAB.ForeColor = &H808080
        TXT_SP_CD = "P"
    Else
        OPT_PLATE.ForeColor = &H808080
        TXT_SP_CD = "S"
    End If

End Sub
Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    If ss1.MaxRows < 1 Then Exit Sub
    
    If Row <= 0 Then
       
        Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
        lBlkcol1 = 0
        lBlkcol2 = 0
        lBlkrow1 = 0
        lBlkrow2 = 0
        
    ElseIf (Col = SPD_EDT1 Or Col = SPD_EDT2 Or Col = SPD_EDT3 Or Col = SPD_EDT4 Or Col = SPD_EDT5 Or Col = SPD_LAST_YN Or Col = SPD_REN_NAME Or Col = SS1_FIRST_REMARK) Then
    
        ss1.Row = ss1.ActiveRow
        ss1.Col = 0
        ss1.Text = "Update"
        ss1.Col = SPD_USERID
        ss1.Text = sUserID
        
    End If

End Sub
Private Sub txt_stdspec_chg_DblClick()
    Call txt_stdspec_chg_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec_chg

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

        Exit Sub

    End If
    
End Sub

Private Sub txt_stlgrd_Change()
   If Len(TXT_STLGRD.Text) <> 11 Then txt_STLGRD_Name.Text = ""
End Sub

Private Sub txt_stlgrd_DblClick()

    Call txt_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_STLGRD
        DD.rControl.Add Item:=txt_STLGRD_Name
        DD.nameType = "1"
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
    End If
        
End Sub