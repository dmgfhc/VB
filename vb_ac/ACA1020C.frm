VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACA1020C 
   Caption         =   "����������״��ѯ_ACA1020C"
   ClientHeight    =   8010
   ClientLeft      =   375
   ClientTop       =   2460
   ClientWidth     =   15225
   FillStyle       =   2  'Horizontal Line
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_shape 
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
      Left            =   15780
      MaxLength       =   3
      TabIndex        =   50
      Text            =   "ss1"
      Top             =   1140
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txt_dzb_cd 
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
      Left            =   15630
      MaxLength       =   11
      TabIndex        =   16
      Top             =   4320
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox txt_ord_item_d 
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
      Left            =   15630
      MaxLength       =   2
      TabIndex        =   15
      Top             =   3870
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox txt_ord_no_d 
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
      Left            =   15630
      MaxLength       =   11
      TabIndex        =   14
      Top             =   3450
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox txt_ord_item 
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
      Left            =   15570
      MaxLength       =   11
      TabIndex        =   10
      Top             =   570
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox txt_ord_no 
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
      Left            =   15570
      MaxLength       =   11
      TabIndex        =   9
      Top             =   150
      Visible         =   0   'False
      Width           =   1410
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9195
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   16219
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACA1020C.frx":0000
      Begin Threed.SSFrame SSFrame1 
         Height          =   2745
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   4842
         _Version        =   196609
         BackColor       =   14737632
         ShadowStyle     =   1
         Begin VB.TextBox Text_CUST_LEVEL 
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
            Left            =   1410
            MaxLength       =   3
            TabIndex        =   69
            Tag             =   "CD_MANA_NO"
            Top             =   2400
            Width           =   915
         End
         Begin VB.TextBox Text_CUST_CLASS 
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
            Left            =   3660
            MaxLength       =   3
            TabIndex        =   68
            Tag             =   "CD_MANA_NO"
            Top             =   2400
            Width           =   915
         End
         Begin VB.TextBox Text_PONO 
            Height          =   315
            Left            =   3660
            TabIndex        =   67
            Top             =   495
            Width           =   1350
         End
         Begin VB.TextBox txt_ord_fl 
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
            Left            =   12120
            MaxLength       =   1
            TabIndex        =   62
            Top             =   1650
            Width           =   765
         End
         Begin VB.TextBox txt_cust_cd_name 
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
            Left            =   8145
            MaxLength       =   40
            TabIndex        =   61
            Tag             =   "�ͻ�"
            Top             =   1650
            Width           =   2235
         End
         Begin VB.TextBox txt_item_count 
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
            Left            =   5220
            MaxLength       =   3
            TabIndex        =   59
            Text            =   "99"
            Top             =   120
            Width           =   450
         End
         Begin VB.TextBox txt_cfm_mill_plt 
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
            Height          =   310
            Left            =   3660
            MaxLength       =   2
            TabIndex        =   25
            Top             =   120
            Width           =   540
         End
         Begin VB.TextBox TXT_STDSPEC 
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
            Left            =   7290
            MaxLength       =   20
            TabIndex        =   58
            Tag             =   "����(��׼��)"
            Top             =   1260
            Width           =   3090
         End
         Begin VB.TextBox txt_sale_way_name 
            Height          =   315
            Left            =   4230
            TabIndex        =   38
            Top             =   1260
            Width           =   1425
         End
         Begin VB.CheckBox Check_ord_END 
            BackColor       =   &H00E0E0E0&
            Caption         =   "��ɶ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   270
            Left            =   10710
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   135
            Width           =   1170
         End
         Begin VB.TextBox Text_BB_PROD_CD 
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
            Left            =   1410
            MaxLength       =   2
            TabIndex        =   36
            Tag             =   "��Ʒ"
            Top             =   120
            Width           =   645
         End
         Begin VB.TextBox Text_BB_DOME_FL 
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
            Left            =   1410
            MaxLength       =   1
            TabIndex        =   35
            Top             =   885
            Width           =   645
         End
         Begin VB.TextBox txt_sale_way 
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
            Left            =   3660
            MaxLength       =   2
            TabIndex        =   34
            Top             =   1260
            Width           =   570
         End
         Begin VB.CheckBox Check_CP_ORD_REM_WGT 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ƿ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   12390
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   135
            Width           =   1170
         End
         Begin VB.TextBox Text_BB_ORD_NO 
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
            Left            =   3660
            MaxLength       =   11
            TabIndex        =   32
            Top             =   885
            Width           =   1350
         End
         Begin VB.TextBox txt_cust_cd 
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
            Left            =   7290
            MaxLength       =   6
            TabIndex        =   31
            Top             =   1650
            Width           =   835
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
            Left            =   7290
            MaxLength       =   11
            TabIndex        =   30
            Tag             =   "����"
            Top             =   885
            Width           =   1320
         End
         Begin VB.TextBox Text_size_knd 
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
            Height          =   310
            Left            =   3660
            MaxLength       =   2
            TabIndex        =   29
            Tag             =   "����"
            Top             =   1680
            Width           =   570
         End
         Begin VB.TextBox Text_size_knd_name 
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
            Left            =   4230
            TabIndex        =   28
            Tag             =   "����"
            Top             =   1680
            Width           =   1425
         End
         Begin VB.TextBox txt_ord_sts 
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
            Left            =   1410
            MaxLength       =   1
            TabIndex        =   27
            Top             =   1650
            Width           =   645
         End
         Begin VB.ComboBox Combo_ORD_ITEM 
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
            Left            =   5010
            TabIndex        =   26
            Top             =   885
            Width           =   660
         End
         Begin VB.TextBox txt_next_plan_htm 
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
            Height          =   310
            Left            =   5700
            MaxLength       =   1
            TabIndex        =   24
            Top             =   2055
            Width           =   570
         End
         Begin VB.TextBox Text_BB_REC_STS 
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
            Left            =   1410
            MaxLength       =   1
            TabIndex        =   22
            Top             =   495
            Width           =   645
         End
         Begin VB.TextBox TXT_ENDUSE_CD 
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
            Left            =   1410
            MaxLength       =   4
            TabIndex        =   21
            Tag             =   "CD_MANA_NO"
            Top             =   2055
            Width           =   645
         End
         Begin VB.TextBox txt_ord_knd 
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
            Left            =   1410
            MaxLength       =   1
            TabIndex        =   20
            Top             =   1260
            Width           =   645
         End
         Begin VB.TextBox txt_ust 
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
            Height          =   310
            Left            =   3660
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "����"
            Top             =   2055
            Width           =   570
         End
         Begin VB.TextBox txt_stlgrd_dec 
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
            Left            =   8610
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   18
            Top             =   885
            Width           =   1770
         End
         Begin Threed.SSCommand cmd_ord 
            Height          =   435
            Left            =   13050
            TabIndex        =   23
            Top             =   1140
            Visible         =   0   'False
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   767
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   12583104
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   11.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "��������"
         End
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   120
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "��Ʒ"
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
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   120
            Top             =   885
            Width           =   1260
            _ExtentX        =   2223
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   6000
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "�û�������"
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
         Begin InDate.UDate UDate_BB_DEL_TO 
            Height          =   315
            Left            =   8940
            TabIndex        =   39
            Tag             =   "������"
            Top             =   120
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
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   2370
            Top             =   1260
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "���۷�ʽ"
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
            Left            =   2370
            Top             =   885
            Width           =   1260
            _ExtentX        =   2223
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
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   120
            Top             =   1650
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "����״̬"
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
         Begin InDate.ULabel ULabel13 
            Height          =   315
            Left            =   6000
            Top             =   885
            Width           =   1260
            _ExtentX        =   2223
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
            ForeColor       =   16711680
         End
         Begin InDate.ULabel ULabel14 
            Height          =   315
            Left            =   2370
            Top             =   1680
            Width           =   1260
            _ExtentX        =   2223
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
         Begin InDate.ULabel ULabel01 
            Height          =   315
            Index           =   14
            Left            =   2370
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
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
         Begin InDate.ULabel ULabel15 
            Height          =   315
            Left            =   4410
            Top             =   2055
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "�ȴ�������"
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
         Begin Threed.SSCheck chk_htm_shot_blast 
            Height          =   285
            Left            =   13650
            TabIndex        =   40
            Top             =   135
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   196609
            Font3D          =   1
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
            Caption         =   "�������"
         End
         Begin Threed.SSCommand cmd_ord_exc 
            Height          =   435
            Left            =   13050
            TabIndex        =   41
            Top             =   540
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   767
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   16576
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   11.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "���̵���"
         End
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   120
            Top             =   495
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "��Ϣ״̬"
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
         Begin InDate.UDate Udate_BB_DEL_FR 
            Height          =   315
            Left            =   7290
            TabIndex        =   42
            Tag             =   "������"
            Top             =   120
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
         Begin CSTextLibCtl.sidbEdit sdb_ord_len 
            Height          =   315
            Left            =   13020
            TabIndex        =   43
            Tag             =   "��������"
            Top             =   2055
            Width           =   780
            _Version        =   262145
            _ExtentX        =   1376
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
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
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord_wid 
            Height          =   315
            Left            =   10260
            TabIndex        =   44
            Tag             =   "��������"
            Top             =   2055
            Width           =   600
            _Version        =   262145
            _ExtentX        =   1058
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
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
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   120
            Top             =   2055
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "������;"
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
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   120
            Top             =   1260
            Width           =   1260
            _ExtentX        =   2223
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
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   2370
            Top             =   2055
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "̽��"
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
         Begin Threed.SSCheck SSC_UST_FL 
            Height          =   285
            Left            =   10680
            TabIndex        =   45
            Top             =   960
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   196609
            Font3D          =   2
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
            Caption         =   "�Ƿ�̽��"
         End
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   6000
            Top             =   495
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "ȷ������"
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
         Begin InDate.UDate Udate_BB_CONF_TO 
            Height          =   315
            Left            =   8940
            TabIndex        =   46
            Tag             =   "������"
            Top             =   495
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
         Begin InDate.UDate Udate_BB_CONF_FR 
            Height          =   315
            Left            =   7290
            TabIndex        =   47
            Tag             =   "������"
            Top             =   495
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
         Begin InDate.ULabel ULabel17 
            Height          =   315
            Left            =   9120
            Top             =   2055
            Width           =   1110
            _ExtentX        =   1958
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
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord_wid_to 
            Height          =   315
            Left            =   11070
            TabIndex        =   54
            Tag             =   "��������"
            Top             =   2055
            Width           =   600
            _Version        =   262145
            _ExtentX        =   1058
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
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
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel18 
            Height          =   315
            Left            =   11880
            Top             =   2055
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   556
            Caption         =   "��������"
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
         Begin CSTextLibCtl.sidbEdit sdb_ord_len_to 
            Height          =   315
            Left            =   14010
            TabIndex        =   56
            Tag             =   "��������"
            Top             =   2055
            Width           =   780
            _Version        =   262145
            _ExtentX        =   1376
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
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
         Begin InDate.ULabel ULabel19 
            Height          =   315
            Left            =   6000
            Top             =   1260
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "��׼��"
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
         Begin InDate.ULabel ULabel20 
            Height          =   315
            Left            =   4560
            Top             =   120
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   556
            Caption         =   "�����"
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
         Begin Threed.SSCheck chk_DZB 
            Height          =   285
            Left            =   13680
            TabIndex        =   60
            Top             =   1650
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   196609
            Font3D          =   1
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
            Caption         =   "�����ȴ�"
         End
         Begin InDate.ULabel ULabel21 
            Height          =   315
            Left            =   10680
            Top             =   1650
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            Caption         =   "����/�ص�Ʒ��"
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
            ForeColor       =   16711680
         End
         Begin Threed.SSCheck SS_KEY_ORD_FL 
            Height          =   285
            Left            =   10680
            TabIndex        =   63
            Top             =   1200
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   196609
            Font3D          =   2
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
            Caption         =   "�ص��ͬ"
            MaskColor       =   255
         End
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Left            =   6390
            Top             =   2055
            Width           =   1110
            _ExtentX        =   1958
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
         Begin CSTextLibCtl.sidbEdit sdb_ord_thk_from 
            Height          =   315
            Left            =   7530
            TabIndex        =   64
            Tag             =   "�������"
            Top             =   2055
            Width           =   570
            _Version        =   262145
            _ExtentX        =   1005
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_ord_thk_to 
            Height          =   315
            Left            =   8340
            TabIndex        =   65
            Tag             =   "�������"
            Top             =   2055
            Width           =   570
            _Version        =   262145
            _ExtentX        =   1005
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
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
            NumDecDigits    =   2
            NumIntDigits    =   4
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel10 
            Height          =   315
            Left            =   6000
            Top             =   1680
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "�ͻ�"
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
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Left            =   2370
            Top             =   495
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "��ͬ���"
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
         Begin InDate.ULabel ULabel26 
            Height          =   315
            Left            =   120
            Top             =   2400
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "�ͻ��ּ�"
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
         Begin InDate.ULabel ULabel23 
            Height          =   315
            Left            =   2370
            Top             =   2400
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            Caption         =   "�ͻ�����"
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
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            Height          =   75
            Left            =   8190
            TabIndex        =   66
            Top             =   2190
            Width           =   90
         End
         Begin VB.Label Lab3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   525
            Index           =   2
            Left            =   11850
            TabIndex        =   53
            Top             =   1050
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            Height          =   75
            Left            =   13860
            TabIndex        =   57
            Top             =   2160
            Width           =   90
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            Height          =   75
            Left            =   10920
            TabIndex        =   55
            Top             =   2190
            Width           =   90
         End
         Begin VB.Label Lab3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   525
            Index           =   1
            Left            =   11850
            TabIndex        =   52
            Top             =   540
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Lab3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "EXCEL ����"
            Height          =   435
            Index           =   0
            Left            =   10800
            TabIndex        =   51
            Top             =   540
            Width           =   1035
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            Height          =   120
            Left            =   8790
            TabIndex        =   49
            Top             =   210
            Width           =   90
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            Height          =   120
            Left            =   8790
            TabIndex        =   48
            Top             =   615
            Width           =   90
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   6390
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2805
         Width           =   7545
         _Version        =   393216
         _ExtentX        =   13309
         _ExtentY        =   11271
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   74
         MaxRows         =   2
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "ACA1020C.frx":0092
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   2730
         Left            =   7605
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2805
         Width           =   7605
         _Version        =   393216
         _ExtentX        =   13414
         _ExtentY        =   4815
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   15
         MaxRows         =   2
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACA1020C.frx":23B2
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   3600
         Left            =   7605
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   5595
         Width           =   7605
         _Version        =   393216
         _ExtentX        =   13414
         _ExtentY        =   6350
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
         MaxCols         =   8
         MaxRows         =   2
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACA1020C.frx":2C47
      End
   End
   Begin VB.TextBox text_ORD_REM_WGT 
      Height          =   345
      Left            =   13305
      TabIndex        =   7
      Top             =   1455
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox Text_BB_DOME_FL_mate 
      Height          =   300
      Left            =   13515
      TabIndex        =   6
      Top             =   1485
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   310
      Left            =   14850
      Max             =   1
      Min             =   99
      TabIndex        =   5
      Top             =   1470
      Value           =   1
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox Text_ORD_ITEM 
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
      Left            =   14550
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1470
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox text_VBCHK 
      Height          =   345
      Left            =   13035
      TabIndex        =   2
      Top             =   1455
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox Text_BB_PROD_CD_mate 
      Height          =   315
      Left            =   13980
      TabIndex        =   4
      Top             =   1455
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.TextBox Text_BB_REC_STS_Name 
      Height          =   315
      Left            =   13755
      TabIndex        =   3
      Top             =   1470
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.CheckBox Check_CP_DEL_DELAY 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�������ӳ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   15990
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2220
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1275
   End
   Begin InDate.ULabel ULabel01 
      Height          =   315
      Index           =   0
      Left            =   10800
      Top             =   1800
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "Ͷ�빤��"
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
      ForeColor       =   16711680
   End
End
Attribute VB_Name = "ACA1020C"
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
'-- Program ID        ACA1020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Yang Zhibin
'-- Date              2003.9.8
'-- Description nnnn
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'  -------------------------------------------------------------------------------

Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting
Public ORD_NO As String             'Transfer to ACA1030C
Public ORD_ITEM As String           'Transfer to ACA1030C

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pContro2 As New Collection      'Master Primary Key Collection
Dim nContro2 As New Collection      'Master Necessary Collection
Dim mContro2 As New Collection      'Master Maxlength check Collection
Dim iContro2 As New Collection      'Master Insert Collection
Dim rContro2 As New Collection      'Master Refer Collection
Dim cContro2 As New Collection      'Master Copy Collection
Dim aContro2 As New Collection      'Master -> Spread Collection
Dim lContro2 As New Collection      'Master Lock Collection

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

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim sCheck1 As String
Dim sCheck2 As String
Dim iCount As Integer

Const iSumColCnt = 13
Const iSumCol1 = 23
Const iSumCol2 = 24
Const iSumCol3 = 25
Const iSumCol4 = 26
Const iSumCol5 = 27
Const iSumCol6 = 28
Const iSumCol7 = 29
Const iSumCol8 = 30
Const iSumCol9 = 31
Const iSumCol10 = 32
Const iSumCol11 = 33
Const iSumCol12 = 34
Const iSumCol13 = 35


Const SS1_URGNT_FL = 56          '����������ɫ���  2012-11-07 by CaoLei      45->47->48
Const SS1_ORD_NO = 1
Const SS1_ORD_ITEM = 2
Const ss1_ORD_WGT = 23          '������
Const SS1_OVER_WGT = 57         '�����������Ų���    20130911 by CaoLei
Const SS1_OVER_FL = 58          '�����Ų����   ��N��- ������ > ������   ��Y��- ������ <= ������  20130911 by CaoLei
Const SS1_KEY_ORD_FL = 59


Const SS2_ORD_NO = 1
Const SS2_ORD_ITEM = 2
Const SS2_UST = 3
Const SS2_COOL = 4
Const SS2_GAS_FL = 5
Const SS2_HTM_N = 6
Const SS2_HTM_T = 7
Const SS2_HTM_Q = 8
Const SS2_GRID_FL = 9
Const SS2_NOPLAN_GAS = 10
Const SS2_CL_FL = 11
Const SS2_QAB = 12
Const SS2_ORD_FL = 13
Const SS2_JC = 14



Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
 '   FormType = "Msheet"
    FormType = "Refer"

   'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(Text_BB_PROD_CD, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Text_BB_DOME_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Text_BB_REC_STS, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Udate_BB_DEL_FR, "p", "n", "m", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(UDate_BB_DEL_TO, "p", "n", "m", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(Text_BB_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(Combo_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_sale_way, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(text_VBCHK, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_ord_sts, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_cust_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_ord_thk_from, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_ORD_THK_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_ord_wid, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_ord_len, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(Text_size_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_cfm_mill_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(chk_htm_shot_blast, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_next_plan_htm, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_ord_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_UST, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_ENDUSE_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(Udate_BB_CONF_FR, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(Udate_BB_CONF_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_ORD_WID_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_ORD_LEN_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_STDSPEC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_item_count, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_ORD_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SS_KEY_ORD_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(Text_PONO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Text_CUST_LEVEL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Text_CUST_CLASS, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_ord_no_d, "p", " ", " ", " ", "r", " ", "l", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
     Call Gp_Ms_Collection(txt_ord_item_d, "p", " ", " ", " ", "r", " ", "l", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
         Call Gp_Ms_Collection(txt_dzb_cd, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
        
    'MASTER Collection
    Mc2.Add Item:=pContro2, Key:="pControl"
    Mc2.Add Item:=nContro2, Key:="nControl"
    Mc2.Add Item:=mContro2, Key:="mControl"
    Mc2.Add Item:=iContro2, Key:="iControl"
    Mc2.Add Item:=rContro2, Key:="rControl"
    Mc2.Add Item:=cContro2, Key:="cControl"
    Mc2.Add Item:=aContro2, Key:="aControl"
    Mc2.Add Item:=lContro2, Key:="lControl"
       
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)     '������
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)   '���ȴ�   2013-06-28   by caolei
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '�������
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '��ӡ����
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 53, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '����������ɫ���  2012-11-07  by  CaoLei
    Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '������
    Call Gp_Sp_Collection(ss1, 58, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 59, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 60, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 61, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 62, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 63, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '��������
    Call Gp_Sp_Collection(ss1, 64, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 65, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    'ĸ���������
    Call Gp_Sp_Collection(ss1, 66, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    'ȡ����ʽ
    Call Gp_Sp_Collection(ss1, 67, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    'ȡ��������
    Call Gp_Sp_Collection(ss1, 68, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '��ͬ��
    Call Gp_Sp_Collection(ss1, 69, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '�Ƿ�ָ������
    Call Gp_Sp_Collection(ss1, 70, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '��������ԭ��
    Call Gp_Sp_Collection(ss1, 71, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '����������
    Call Gp_Sp_Collection(ss1, 72, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '�ͻ��ּ�
    Call Gp_Sp_Collection(ss1, 73, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '�ͻ�����
    Call Gp_Sp_Collection(ss1, 74, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '������ҵҪ��
     
   'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"

    sc1.Add Item:="ACA1020C.P_SREFER", Key:="P-R"
    sc1.Add Item:="ACA1020C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="ACA1020C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    ' control part   Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, "p", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
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
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
'    sc2.Add Item:="ACA1020C.P_SREFER1", Key:="P-R"
    sc2.Add Item:="ACA1020C.P_ONEROW1", Key:="P-O"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=2, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    ' control part   Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss3, 1, "p", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="ACA1020C.P_SREFER2", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=3, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc3, Key:="Sc3"
    
    'Sum Column Count
    iSumCnt = iSumColCnt
    
    'Sum Column Setting
    iSumCol.Add Item:=iSumCol1
    iSumCol.Add Item:=iSumCol2
    iSumCol.Add Item:=iSumCol3
    iSumCol.Add Item:=iSumCol4
    iSumCol.Add Item:=iSumCol5
    iSumCol.Add Item:=iSumCol6
    iSumCol.Add Item:=iSumCol7
    iSumCol.Add Item:=iSumCol8
    iSumCol.Add Item:=iSumCol9
    iSumCol.Add Item:=iSumCol10
    iSumCol.Add Item:=iSumCol11
    iSumCol.Add Item:=iSumCol12
    iSumCol.Add Item:=iSumCol13
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub Check_CP_DEL_DELAY_Click()

    If Check_CP_DEL_DELAY.Value = 1 Then
       Check_CP_DEL_DELAY.ForeColor = &HFF&
    Else
       Check_CP_DEL_DELAY.ForeColor = &H80000012
    End If
    
End Sub
'
'Private Sub Check_CP_ORD_REM_WGT_Click()
'
'    If Check_CP_ORD_REM_WGT.Value = 1 Then
'       Check_CP_ORD_REM_WGT.ForeColor = &HFF&
'    Else
'       Check_CP_ORD_REM_WGT.ForeColor = &H80000012
'    End If
'
'End Sub

Private Sub Check_ord_END_Click()

    If Check_ord_END.Value = 1 Then
        Text_BB_REC_STS.Text = 2
        Check_ord_END.ForeColor = &HFF&
        Check_CP_DEL_DELAY.ForeColor = &H80000012
'        Check_CP_ORD_REM_WGT.ForeColor = &H80000012
    Else
'        Text_BB_REC_STS.Text = ""
        Check_ord_END.ForeColor = &H80000012
    End If
    
End Sub

Private Sub chk_DZB_Click(Value As Integer)
    If chk_DZB.Value = ssCBUnchecked Then
       SSSplitter1.Panes(2).Width = 0
       SSSplitter1.Panes(3).Width = 0
       Call Gf_Sp_Cls(Proc_Sc("Sc2"))
       Call Gf_Sp_Cls(Proc_Sc("Sc3"))
       Lab3(1).Visible = False
       Lab3(2).Visible = False
    Else
       SSSplitter1.Panes(2).Width = SSSplitter1.Panes(0).Width / 3
       Lab3(1).Visible = True
       Lab3(2).Visible = True
    End If
End Sub

Private Sub cmd_ord_Click()

     Call Form_Pro

'    Dim iRow As Long
'
'    With ss1
'         For iRow = 1 To .MaxRows
'             .Row = iRow: .Col = 0
'             If .Text = "Update" Then
'                 .Col = 1: txt_ord_no.Text = .Text
'                 .Col = 2: txt_ord_item.Text = .Text
'                 Call Gp_CallACB3050P
'                 Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
'             End If
'         Next iRow
'    End With
'
'    txt_ord_no.Text = ""
'    txt_ord_item.Text = ""
    
End Sub

Private Sub cmd_ord_exc_Click()
    Call ExcelOrdPrn
End Sub

Private Sub Combo_ORD_ITEM_Change()

'    If combo_ord_item.Text <> "" Then
'         If combo_ord_item.Text > combo_ord_item.ListCount Then
'           combo_ord_item.Text = ""
'        End If
'    End If
    
End Sub

Private Sub Combo_ORD_ITEM_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub Combo_ORD_ITEM_LostFocus()

    Dim S As String
  
    If Len(Combo_ORD_ITEM.Text) = 1 Then
        S = Combo_ORD_ITEM.Text
        Combo_ORD_ITEM.Text = "0" + S
    End If
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

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
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"), False)
    Call Gp_Sp_Setting(Proc_Sc("Sc3")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc2")("Spread"))
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc3")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "C-System.INI", Me.Name)
    
    SSSplitter1.Panes(2).Width = 0
    SSSplitter1.Panes(0).LockHeight = True
    
    If Gf_Sc_Authority(sAuthority, "U") Then
       cmd_ord.Visible = True
       chk_DZB.Visible = True
    Else
       cmd_ord.Visible = False
       chk_DZB.Visible = False
    End If
    
    If Gf_Sc_Authority(sAuthority, "I") Then
       cmd_ord_exc.Visible = True
    Else
       cmd_ord_exc.Visible = False
    End If
    
    Text_BB_PROD_CD.Text = "PP"
    
    Udate_BB_DEL_FR.Text = Mid(Udate_BB_DEL_FR.Text, 1, 8) + "01"

    UDate_BB_DEL_TO.Text = Format(DateAdd("m", 1, Udate_BB_DEL_FR.Text), "YYYY MM DD")
    UDate_BB_DEL_TO.Text = Format(DateAdd("d", -1, UDate_BB_DEL_TO.Text), "YYYY MM DD")
    
    Udate_BB_CONF_FR.Text = ""
    Udate_BB_CONF_TO.Text = ""
    
    Text_BB_REC_STS.Text = "2"

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc3")("Spread"), "C-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set pContro2 = Nothing
    Set nContro2 = Nothing
    Set iContro2 = Nothing
    Set rContro2 = Nothing
    Set cContro2 = Nothing
    Set aContro2 = Nothing
    Set lContro2 = Nothing
    Set mContro2 = Nothing
            
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
    
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call Gf_Sp_Cls(Proc_Sc("SC2"))
        Call Gf_Sp_Cls(Proc_Sc("SC3"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        Combo_ORD_ITEM.Clear
        rControl(1).SetFocus
    End If

    Udate_BB_DEL_FR.Text = Mid(Udate_BB_DEL_FR.Text, 1, 8) + "01"
    
    UDate_BB_DEL_TO.Text = DateAdd("m", 1, Udate_BB_DEL_FR.Text)
    UDate_BB_DEL_TO.Text = DateAdd("d", -1, UDate_BB_DEL_TO.Text)

    Check_ord_END.Value = 0
    Check_CP_DEL_DELAY.Value = 0
'    Check_CP_ORD_REM_WGT.Value = 0
    Text_BB_PROD_CD_mate.Text = ""
    Text_BB_REC_STS_Name.Text = ""
    txt_sale_way_name.Text = ""
    Text_BB_DOME_FL_mate.Text = ""
    text_VBCHK.Text = ""
    iCount = 0
    
End Sub

Public Sub Form_Ref()

    Dim S As String
    Dim sMesg As String
    Dim sQuery As String
    
    Dim iCount As Long
    Dim iRow As Long
    Dim iCol As Long
    Dim iOrd_no As String
    Dim ORD_WGT As Double
    Dim OVER_WGT As Double
    Dim OVER_FL As String
    
    If Combo_ORD_ITEM.Text <> "" Then
        If Len(Combo_ORD_ITEM.Text) = 1 Then
            S = Combo_ORD_ITEM.Text
            Combo_ORD_ITEM.Text = "0" + S
        End If
    End If
    
    If Text_BB_ORD_NO.Text = "" Then
'        Text_ORD_ITEM.Text = ""
         Combo_ORD_ITEM.Text = ""
    End If
        
    If Check_CP_DEL_DELAY.Value = 1 And Check_CP_ORD_REM_WGT.Value = 1 Then
        sCheck1 = "11"
    ElseIf Check_CP_DEL_DELAY.Value = 1 And Check_CP_ORD_REM_WGT.Value = 0 Then
        sCheck1 = "10"
    ElseIf Check_CP_DEL_DELAY.Value = 0 And Check_CP_ORD_REM_WGT.Value = 1 Then
        sCheck1 = "01"
    Else
        sCheck1 = "00"
    End If

    text_VBCHK.Text = sCheck1
    
    SSSplitter1.Panes(2).Width = 0
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
           
    If UDate_BB_DEL_TO.RawData >= Udate_BB_DEL_FR.RawData Or UDate_BB_DEL_TO.RawData = "" Then
    
        If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
        
        sMesg = Gf_Ms_NeceCheck(nControl)
        If sMesg = "OK" Then
        
            sMesg = Gf_Ms_NeceCheck2(mControl)
            If sMesg = "OK" Then
                
                sQuery = Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", pControl)
                If Gf_Total_Display(M_CN1, Proc_Sc("Sc"), sQuery, 0, iSumCnt, iSumCol) Then
                    ss1.OperationMode = OperationModeNormal
'                If Gf_Multi_Stotal_Display(M_CN1, Proc_Sc("Sc"), sQuery, 4, 6, iSumCnt, iSumCol) Then
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                End If
        
            Else
                sMesg = sMesg + " Must input according to length of item"
                Call Gp_MsgBoxDisplay(sMesg)
            End If
                
        Else
           sMesg = sMesg + " Must input necessarily"
           Call Gp_MsgBoxDisplay(sMesg)
        End If
                 
    Else
       Call MsgBox("�������ڲ����Ϲ淶!" & Chr(10) & "�������", vbExclamation + vbOKOnly, "����")
    End If
    
    If chk_DZB.Value = -1 Then
    
       SSSplitter1.Panes(2).Width = SSSplitter1.Panes(0).Width / 3
       
       If ss1.MaxRows > 0 Then
          ss2.MaxRows = ss1.MaxRows - 1
       End If
           
       For iRow = 1 To ss1.MaxRows - 1
            ss1.ROW = iRow
            For iCol = 1 To 2
                ss1.Col = iCol
                iOrd_no = ss1.Text
                ss2.ROW = iRow
                ss2.Col = iCol
                ss2.Text = iOrd_no
                             
            Next iCol
                
       Next iRow
       
        

       Call Gp_Sp_DZB(M_CN1, Proc_Sc("Sc2"))
              
    End If
    
    
    '����������ɫ���  2012-11-07   by   CaoLei
    Call SS1_CHANGE_COLOR
    
    
    '�����Ų����
    With ss1
        If .MaxRows < 1 Then
          Exit Sub
        End If
        OVER_FL = "N"
     For iCount = 1 To .MaxRows - 1

         '������
         .ROW = iCount:            .Col = ss1_ORD_WGT
          ORD_WGT = Val(.Text)
          '������
          .ROW = iCount:           .Col = SS1_OVER_WGT
          OVER_WGT = Val(.Text)
          .ROW = iCount:           .Col = SS1_OVER_FL
          '������ < ������ �á�Y�� ���� �á�1��
          If ORD_WGT < OVER_WGT Then
                OVER_FL = "Y"
          Else
                OVER_FL = "N"
          End If
          .Text = OVER_FL
     Next iCount
           
   End With
       
   
 
End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    
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

Public Sub Form_Exc()
    
    If txt_shape.Text = "ss1" Then
       Call Gp_Sp_Excel(Me, ss1, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf txt_shape.Text = "ss2" Then
       Call Gp_ACA1020C_Excel(Me, ss2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf txt_shape.Text = "ss3" Then
       Call Gp_Sp_Excel(Me, ss3, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If
    

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub Lab3_Click(Index As Integer)

    If Index = 0 Then
       txt_shape.Text = "ss1"
       Lab3(0).Caption = "EXCEL ����"
       Lab3(1).Caption = ""
       Lab3(2).Caption = ""
       Lab3(0).BackColor = &HFFFFC0
       Lab3(1).BackColor = &HE0E0E0
       Lab3(2).BackColor = &HE0E0E0
    ElseIf Index = 1 Then
       txt_shape.Text = "ss2"
       Lab3(1).Caption = "EXCEL ����"
       Lab3(0).Caption = ""
       Lab3(2).Caption = ""
       Lab3(1).BackColor = &HFFFFC0
       Lab3(0).BackColor = &HE0E0E0
       Lab3(2).BackColor = &HE0E0E0
    ElseIf Index = 2 Then
       txt_shape.Text = "ss3"
       Lab3(2).Caption = "EXCEL ����"
       Lab3(0).Caption = ""
       Lab3(1).Caption = ""
       Lab3(2).BackColor = &HFFFFC0
       Lab3(0).BackColor = &HE0E0E0
       Lab3(1).BackColor = &HE0E0E0
    End If

End Sub



Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    Dim i As Integer
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

    Dim Row1 As Long
    Dim Row2 As Long
    Dim Col As Long
    
    Dim str_ord_fl As String
    Dim str_rec_sts As String
    
    If Not Gf_Sc_Authority(sAuthority, "U") Then
       Exit Sub
    End If
    
    Col = BlockCol
    Row1 = BlockRow
    Row2 = BlockRow2
  
    If Col = -1 Then

     For i = BlockRow To BlockRow2
        Call ss1_row_Click(1, i)
     Next
     
   End If

   Call ss1.SetActiveCell(1, Row2)

End Sub
Private Sub ss1_row_Click(ByVal Col As Long, ByVal ROW As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If ROW < 1 Or ROW = ss1.MaxRows Then Exit Sub
    If ss1.MaxRows < 1 Then Exit Sub
    
    ss1.ROW = ROW
    ss1.Col = 0
    
    ss1.ReDraw = False
    
    If ss1.Text <> "Update" Then
        
        ss1.Text = "Update"
        
        Call Gp_Sp_BlockColor(ss1, 1, -1, ROW, ROW, , &HFFFF80)
    Else
       
        ss1.Text = ""
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW)
        Call SS1_CHANGE_COLOR
       
    End If
    
    ss1.ReDraw = True
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)

End Sub
Public Sub Gp_Sp_Sort(sPname As Variant, Col As Variant, ROW As Variant, Optional CL As Boolean = False, Optional Key_Col As Long = 0)

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim sKey_Value() As String

    With sPname

        If .MaxRows < 1 Then Exit Sub
        
        If ROW <= 0 And Col > 0 Then
        
            If CL And Key_Col <> 0 Then
            
                ReDim sKey_Value(1 To .MaxRows)
                        
                For i = 1 To .MaxRows
                    .ROW = i
                    .Col = 0
                    
                    If .Text <> "" Then
                        j = j + 1
                        .Col = Key_Col
                        sKey_Value(j) = .Text
                        .Col = 0
                        .Text = ""
                        Call Gp_Sp_BlockColor(sPname, 1, .MaxCols, i, i, BLACK, WHITE)
                    End If
                Next i
                
            Else
            
                For i = 1 To .MaxRows
                    .ROW = i
                    .Col = 0
                    If .Text <> "" Then
                        Exit Sub
                    End If
                Next i
                
            End If
        
            .SortBy = SS_SORT_BY_ROW
            
            If .SortKey(1) = Col Then
                If .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING Then
                    .SortKeyOrder(1) = SS_SORT_ORDER_DESCENDING
                Else
                    .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
                End If
            Else
                If .SortKey(1) = -1 Then
                    .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
                End If
                .SortKey(1) = Col
                
            End If
            
            .Col = 1: .Col2 = .MaxCols
            .ROW = 0: .Row2 = .MaxRows
            
            .Action = SS_ACTION_SORT
            
            'CLEAR
            If CL And Key_Col <> 0 Then
                For i = 1 To j
                    For k = 1 To .MaxRows
                        .ROW = k
                        .Col = Key_Col
                        If .Text = sKey_Value(i) Then
                            Call Gp_Sp_BlockColor(sPname, 1, .MaxCols, k, k, WHITE, BLUE)
                            .Col = 0
                            .Text = "Select"
                        End If
                    Next k
                Next i
            ElseIf CL And Key_Col = 0 Then
                .Col = 0: .Col2 = 0
                .ROW = 1: .Row2 = .MaxRows
                .BlockMode = True
                .Text = ""
                .BlockMode = False
                Call Gp_Sp_BlockColor(sPname, 1, .MaxCols, 1, .MaxRows, BLACK, WHITE)
            End If
            
        End If
        
    End With
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal ROW As Long)

   Dim iRow As Long
   Dim iCol As Long
   Dim iTxt As String
   Dim iOrd_no As String
   Dim iOrd_item As String
   
   iRow = ROW
   iCol = Col
   
   Call Gf_Sp_Cls(Proc_Sc("Sc3"))
   
   ss2.ROW = iRow:  ss2.Col = iCol:  iTxt = ss2.Text
   If iTxt = "" Or Val(iTxt) = 0 Then
      Exit Sub
   End If
   
   If iRow > 0 And iCol > 1 Then
      ss2.ROW = iRow:                                             txt_dzb_cd = iCol
      ss2.Col = SS2_ORD_NO:        iOrd_no = ss2.Text:            txt_ord_no_d = iOrd_no
      ss2.Col = SS2_ORD_ITEM:      iOrd_item = ss2.Text:          txt_ord_item_d = iOrd_item
      If Gf_Sp_Refer(M_CN1, Sc3, Mc2) Then
            ss3.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
      End If
   End If
   
End Sub

Private Sub SSC_UST_FL_Click(Value As Integer)

    If SSC_UST_FL.Value = -1 Then
       SSC_UST_FL.ForeColor = &HFF&
       txt_UST.Text = "Y"
    Else
       SSC_UST_FL.ForeColor = &H808080
       txt_UST.Text = ""
    End If
    
End Sub


Private Sub SS1_CHANGE_COLOR()

    With ss1
      
        If .MaxRows <= 0 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .ROW = iCount
            
            '�ص��ͬ
            ss1.ROW = .ROW:       ss1.Col = SS1_KEY_ORD_FL
            If ss1.Text = "Y" Then
'                 Call Gp_Sp_BlockColor(ss1, 0, .MaxCols, .Row, .Row, &HFF&)
                 Call Gp_Sp_RowColor(ss1, .ROW, , &HC0C0FF)
            End If
            
              '����������ɫ��� 2012-11-07  by  GengXueyu
            ss1.ROW = .ROW:       ss1.Col = SS1_URGNT_FL
            If ss1.Text = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, SS1_ORD_NO, SS1_ORD_NO, .ROW, .ROW, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_ORD_ITEM, SS1_ORD_ITEM, .ROW, .ROW, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_URGNT_FL, SS1_URGNT_FL, .ROW, .ROW, &HC000&)
            End If
            
        Next iCount

    End With
    
End Sub



Private Sub Text_BB_DOME_FL_Change()
               
    Text_BB_DOME_FL.Text = StrConv(Text_BB_DOME_FL.Text, vbUpperCase)
    
    Select Case Text_BB_DOME_FL.Text
    
        Case "E", "D", ""
        
        Case Else
           Text_BB_DOME_FL.Text = ""
           Call MsgBox("����������룺  " & Chr(10) & "D������" & Chr(10) & "E������", vbExclamation + vbOKOnly, "����")
           
    End Select
     
End Sub

Private Sub Text_BB_DOME_FL_DblClick()

    Call Text_BB_DOME_FL_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_BB_DOME_FL_KeyUp(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"

        DD.sKey = "B0002"
        
        DD.rControl.Add Item:=Text_BB_DOME_FL
        DD.rControl.Add Item:=Text_BB_DOME_FL_mate
   
        DD.nameType = "2"
        'DD.nameType="1" ���������Ʋ�ѯ
        'DD.nameType="2" ��Ӣ�����Ʋ�ѯ
        
        Call Gf_Common_DD(M_CN1, KeyCode)

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() ���ڿͻ�����

        Exit Sub
        
    End If

    If Len(Trim(Text_BB_DOME_FL.Text)) = Text_BB_DOME_FL.MaxLength Then
       '  Gf_ComnNAME_Find( �����ַ���, DD.sKEy���� ,DD.nameType)
       ' Gf_CustNameFind( �����ַ���, �ͻ���������,DD.nameType)
        Text_BB_DOME_FL_mate.Text = Gf_ComnNameFind(M_CN1, "B0002", Text_BB_DOME_FL.Text, 2)
    Else
        Text_BB_DOME_FL_mate.Text = ""
    End If
    
End Sub

Private Sub text_bb_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String
    
    If Len(Trim(Text_BB_ORD_NO.Text)) = Text_BB_ORD_NO.MaxLength Then
    
        If Combo_ORD_ITEM.Text <> "" Then Exit Sub
        
        Text_BB_ORD_NO.Text = StrConv(Text_BB_ORD_NO.Text, vbUpperCase)
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(Text_BB_ORD_NO.Text) & "'"
        Call Gf_ComboAdd(M_CN1, Combo_ORD_ITEM, sQuery)
       
       'If Combo_ORD_ITEM.ListCount <> 0 Then
       '   Combo_ORD_ITEM.ListIndex = 0
       'End If
    Else
    
        Combo_ORD_ITEM.Clear
        
    End If

End Sub

Private Sub Text_BB_PROD_CD_Change()

    Select Case Text_BB_PROD_CD.Text
         Case "S", "s", "SL"
             Text_BB_PROD_CD.Text = "SL"
         Case "P", "p", "PP"
             Text_BB_PROD_CD.Text = "PP"
         Case "H", "h", "HC"
             Text_BB_PROD_CD.Text = "HC"
         Case "", "**"
             Text_BB_PROD_CD.Text = ""
         Case Else
             Text_BB_PROD_CD.Text = ""
             Call MsgBox("��Ʒ�������" & Chr(10) & "�����Ϲ淶! �������", vbExclamation + vbOKOnly, "����")
    End Select
           
End Sub

Private Sub Text_BB_PROD_CD_DblClick()

    Call Text_BB_PROD_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_BB_REC_STS_Change()

    If Not Text_BB_REC_STS.Text = "" Then
        If Not Text_BB_REC_STS.Text = "1" Then
            If Not Text_BB_REC_STS.Text = "2" Then
                If Not Text_BB_REC_STS.Text = "3" Then
    '               Call MsgBox("״̬����" & Chr(10) & "�����Ϲ淶! �������", vbExclamation + vbOKOnly, "����")
                    Text_BB_REC_STS.Text = ""
                End If
            End If
        End If
    End If
'
'    If Text_BB_REC_STS.Text = "2" Then
'        Check_ord_END.Value = 1
'    '   Check_ord_END.Enabled = False
'    Else
'        Check_ord_END.Value = 0
'    '   Check_ord_END.Enabled = True
'    End If

End Sub

Private Sub Text_BB_REC_STS_DblClick()

    Call Text_BB_REC_STS_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_BB_REC_STS_KeyUp(KeyCode As Integer, Shift As Integer)

    Text_BB_REC_STS_Name = ""
    
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "Z0005"
        DD.rControl.Add Item:=Text_BB_REC_STS
        DD.rControl.Add Item:=Text_BB_REC_STS_Name
        
        DD.nameType = "2"
        'DD.nameType="1" ���������Ʋ�ѯ
        'DD.nameType="2" ��Ӣ�����Ʋ�ѯ
        Call Gf_Common_DD(M_CN1, KeyCode)
        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() ���ڿͻ�����
        
        Exit Sub
        
    End If

    If Len(Trim(Text_BB_REC_STS.Text)) = Text_BB_REC_STS.MaxLength Then
       '  Gf_ComnNAME_Find( �����ַ���, DD.sKEy���� ,DD.nameType)
       ' Gf_CustNameFind( �����ַ���, �ͻ���������,DD.nameType)
        Text_BB_REC_STS_Name.Text = Gf_ComnNameFind(M_CN1, "Z0005", Text_BB_REC_STS.Text, 2)
    Else
        Text_BB_REC_STS_Name.Text = ""
    End If
    
End Sub

Private Sub Text_BB_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)
   
   Text_BB_PROD_CD_mate = ""
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=Text_BB_PROD_CD
        DD.rControl.Add Item:=Text_BB_PROD_CD_mate
   
        DD.nameType = "2"
        'DD.nameType="1" ���������Ʋ�ѯ
        'DD.nameType="2" ��Ӣ�����Ʋ�ѯ
       
        Call Gf_Common_DD(M_CN1, KeyCode)

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        'Gf_Customer_DD() ���ڿͻ�����

        Exit Sub
        
    End If

    If Len(Trim(Text_BB_PROD_CD.Text)) = Text_BB_PROD_CD.MaxLength Then
       '  Gf_ComnNAME_Find( �����ַ���, DD.sKEy���� ,DD.nameType)
       ' Gf_CustNameFind( �����ַ���, �ͻ���������,DD.nameType)
        Text_BB_PROD_CD_mate.Text = Gf_ComnNameFind(M_CN1, "B0005", Text_BB_PROD_CD.Text, 2)
    Else
        Text_BB_PROD_CD_mate.Text = ""
    End If
    
End Sub

Private Sub Text_CUST_LEVEL_DblClick()

    Call Text_CUST_LEVEL_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_CUST_LEVEL_KeyUp(KeyCode As Integer, Shift As Integer)

      If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "A0007"
        
        DD.rControl.Add Item:=Text_CUST_LEVEL
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        
    End If
    
End Sub
Private Sub Text_CUST_CLASS_DblClick()

    Call Text_CUST_CLASS_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_CUST_CLASS_KeyUp(KeyCode As Integer, Shift As Integer)

      If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "A0006"
        
        DD.rControl.Add Item:=Text_CUST_CLASS
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        
    End If
    
End Sub

Private Sub Text_size_knd_DblClick()

    Call Text_size_knd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TXT_CFM_MILL_PLT_DblClick()

    Call txt_cfm_mill_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_cfm_mill_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_cfm_mill_plt
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

End Sub

Private Sub txt_cust_cd_DblClick()

    Call txt_cust_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_next_plan_htm_DblClick()

    Call txt_next_plan_htm_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_next_plan_htm_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Q0073"
        
        DD.rControl.Add Item:=txt_next_plan_htm
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If
    
End Sub

Private Sub txt_ord_fl_Change()
   Call txt_ord_fl_KeyUp(vbKeyF4, 0)
End Sub


Private Sub txt_ord_fl_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0022"
        DD.rControl.Add Item:=TXT_ORD_FL

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
 End Sub
    

Private Sub txt_ord_knd_DblClick()

    Call txt_ord_knd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_ord_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0009"
        DD.rControl.Add Item:=txt_ord_knd

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
    
End Sub
Private Sub txt_enduse_cd_DblClick()

    Call txt_enduse_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_enduse_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
                 
        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_ENDUSE_CD
        DD.nameType = "2"
            
        Call Gf_Usage_DD(M_CN1, KeyCode)
        
    End If
    
End Sub

Private Sub txt_ord_sts_DblClick()

    Call txt_ord_sts_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_sale_way_DblClick()

    Call txt_sale_way_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_sale_way_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0010"
        DD.rControl.Add Item:=txt_sale_way
        DD.rControl.Add Item:=txt_sale_way_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_sale_way)) = txt_sale_way.MaxLength Then
        txt_sale_way_name.Text = Gf_ComnNameFind(M_CN1, "B0010", Trim(txt_sale_way.Text), 2)
    Else
        txt_sale_way_name.Text = ""
    End If
 
End Sub

Private Sub Text_ORD_ITEM_Change()

    If Text_ORD_ITEM.Text <> "" Then
        If Val(Text_ORD_ITEM.Text) > iCount Or Val(Text_ORD_ITEM.Text) < 0 Or Text_ORD_ITEM.Text = "00" Then
            Call MsgBox("����������벻��ȷ!" & Chr(10) & "�����ԡ�", vbExclamation + vbOKOnly, "����")
            Text_ORD_ITEM.Text = ""
        End If
    End If

End Sub

Private Sub Text_ORD_ITEM_KeyPress(KeyAscii As Integer)
    KeyAscii = txt_KeyPress(KeyAscii)
End Sub

Private Sub Text_ORD_ITEM_LostFocus()

    Dim S As String
  
    If Len(Text_ORD_ITEM.Text) = 1 Then
        S = Text_ORD_ITEM.Text
        Text_ORD_ITEM.Text = "0" + S
    End If
  
End Sub

Private Sub Text_size_knd_Change()
    If Len(Trim(Text_size_knd.Text)) = Text_size_knd.MaxLength Then
        Text_size_knd_name.Text = Gf_ComnNameFind(M_CN1, "B0043", Text_size_knd.Text, 2)
        Exit Sub
    Else
        Text_size_knd_name.Text = ""
    End If
End Sub

Private Sub Text_size_knd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"

        DD.rControl.Add Item:=Text_size_knd

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_cust_cd
        DD.rControl.Add Item:=txt_cust_cd_name

        DD.nameType = "1"

        Call Gf_Customer_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_cust_cd)) = txt_cust_cd.MaxLength Then
        txt_cust_cd_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
    Else
        txt_cust_cd_name.Text = ""
    End If

End Sub

Private Sub txt_ord_sts_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0011"
        DD.rControl.Add Item:=txt_ord_sts
   
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If
    
End Sub



Private Sub txt_stdspec_DblClick()
    Call txt_stdspec_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stdspec_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_STDSPEC
           
        If Text_BB_PROD_CD.Text = "SL" Then
            DD.nameType = "1"
            Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        Else
            Call Gf_StdSPEC_DD2(M_CN1, KeyCode)
        End If
    End If
    
End Sub

Private Sub txt_stlgrd_DblClick()

    Call txt_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        'txt_act_stlgrd.Text = ""
        DD.rControl.Add Item:=txt_stlgrd
        DD.rControl.Add Item:=txt_stlgrd_dec

        Call Gf_Stlgrd_DD(M_CN1, vbKeyF4)

        Exit Sub
    End If
    
    If Len(Trim(txt_stlgrd)) = txt_stlgrd.MaxLength Then
        txt_stlgrd_dec.Text = Gf_CustNameFind(M_CN1, Trim(txt_stlgrd.Text), 1)
    Else
        txt_stlgrd_dec.Text = ""
    End If
    
End Sub

Private Sub txt_UST_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.sKey = "Q0046"
        DD.rControl.Add Item:=txt_UST
    
        DD.nameType = "2"
        
        Call Gf_Mill_Common_DD(M_CN1, vbKeyF4)
        
    End If
End Sub

Private Sub VScroll1_Change()

    VScroll1.Min = iCount
    
    Select Case VScroll1.Value
        Case 1 To 9
            Text_ORD_ITEM.Text = "0" & VScroll1.Value
        Case 10 To 99
            Text_ORD_ITEM.Text = VScroll1.Value
    End Select
    
End Sub

Private Function txt_KeyPress(KeyAscii As Integer) As Integer

    Select Case KeyAscii
               
        'Case Is <= 32
        '   txt_KeyPress = KeyAscii
        Case 48 To 57
            txt_KeyPress = KeyAscii
        'Case 46
        '   txt_KeyPress = KeyAscii
        Case Else
             txt_KeyPress = 0
    End Select
                    
End Function

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)

    If ss1.MaxRows < 1 Or ROW = 0 Or ROW = -999 Or ss1.MaxRows = ROW Then Exit Sub
    
      If Col <> 55 Then
        Unload ACA1030C
        Load ACA1030C
        
        ACA1030C.txt_prod_cd = ACA1020C.Text_BB_PROD_CD
        
        ss1.ROW = ROW
        ss1.Col = 1
'<<<<<<< .mine
        ACA1030C.txt_ord_no.Text = Trim(ss1.Value)
'=======
        ACA1030C.txt_ord_no.Text = Trim(ss1.Value)
'>>>>>>> .r4464
        
        ss1.ROW = ROW
        ss1.Col = 2
        ACA1030C.cbo_ord_item.Text = Trim(ss1.Value)
        
        ACA1030C.Active_CForm = "ACA1030C"
        ACA1030C.Show
        ACA1030C.SetFocus
     Else
        Unload CGT2040C
        Load CGT2040C
        
    
        
        ss1.ROW = ROW
        ss1.Col = 1
'<<<<<<< .mine
        CGT2040C.txt_ord_item.Text = Trim(ss1.Value)
'=======
        CGT2040C.txt_ord_item.Text = Trim(ss1.Value)
'>>>>>>> .r4464
        
        ss1.ROW = ROW
        ss1.Col = 2
        CGT2040C.txt_ord_no.Text = Trim(ss1.Value)
'
        CGT2040C.SDT_PROD_DATE_FROM = DateSerial(Year(Now()), Month(Now()), 1)
        CGT2040C.SDT_PROD_DATE_TO = DateSerial(Year(Now()), Month(Now()) + 1, 0)
'
        
        CGT2040C.Show
        CGT2040C.SetFocus
        Call CGT2040C.Form_Ref
        
    End If
    
End Sub

Public Sub Gp_CallACB3050P()

On Error GoTo Gp_CallACB3050P_Error

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    

    sQuery = "{call ACB3050P ('" + txt_ord_no.Text + "','" + txt_ord_item.Text + "',?)}"

    sQuery = "{call ACB3050P ('" + txt_ord_no.Text + "','" + txt_ord_item.Text + "',?)}"

    
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
'        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Gp_CallACB3050P_Error:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Gp_CallACB3050P_Error : " & Error)
    
End Sub
Private Sub ExcelOrdPrn()
'---------------------------------------------------------------------------------------
'   1.ID           : ExcelOrdPrn
'   2.Name         : Spread --> Excel
'   3.Input  Value :
'   4.Return Value :
'   5.Writer       : YANG MENG
'   6.Create Date  : 2009. 11 .26
'   7.Modify Date  :
'   8.Comment      : Spread --> Excel
'---------------------------------------------------------------------------------------

    Dim i               As Long
    Dim II              As Long
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    Dim sDate           As String
    
    Dim sExlRange       As String
    Dim sRow_F          As Long
    Dim sRow_T          As Long
    Dim sOrd_Sts        As String
    Dim sOrd_Thk        As Double
    Dim sOrd_Wid        As Double
    Dim sOrd_Len        As Double
    Dim sOrd_Wgt        As Double
    Dim sOrd_RemWgt     As Double
    Dim sOrd_CulWgt     As Double
    Dim sOrd_ProcWgt(1 To 8)   As Double
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    sRow_F = 2
    sRow_T = ss1.MaxRows
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If ERR.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    ERR.Clear

    xlApp.Workbooks.Open (App.Path & "\ACA1020C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    Clipboard.Clear
    ss1.SetSelection 1, 1, 3, ss1.MaxRows - 1 '���������С�״̬
    ss1.ClipboardCopy
    xlApp.Range("A2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 5, 1, 5, ss1.MaxRows - 1 '�ͻ�
    ss1.ClipboardCopy
    xlApp.Range("D2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 6, 1, 6, ss1.MaxRows - 1 '��������
    ss1.ClipboardCopy
    xlApp.Range("F2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 14, 1, 17, ss1.MaxRows - 1 '��׼�����
    ss1.ClipboardCopy
    xlApp.Range("G2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 18, 1, 18, ss1.MaxRows - 1 '����
    ss1.ClipboardCopy
    xlApp.Range("K2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 22, 1, 23, ss1.MaxRows - 1 '�����ڡ�������
    ss1.ClipboardCopy
    xlApp.Range("L2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 25, 1, 25, ss1.MaxRows - 1 'Ƿ�������ޣ�
    ss1.ClipboardCopy
    xlApp.Range("N2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 26, 1, 28, ss1.MaxRows - 1 '����..���ֵȴ�
    ss1.ClipboardCopy
    xlApp.Range("O2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 31, 1, 35, ss1.MaxRows - 1 '�����ȴ�..�������     33->���ȴ�
    ss1.ClipboardCopy
    xlApp.Range("R2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 38, 1, 38, ss1.MaxRows - 1 'Ͷ�빤��
    ss1.ClipboardCopy
    xlApp.Range("W2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 39, 1, 39, ss1.MaxRows - 1 '��ǰ��������
    ss1.ClipboardCopy
    xlApp.Range("X2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 53, 1, 53, ss1.MaxRows - 1 '�ͻ�����    '42->44->47
    ss1.ClipboardCopy
    xlApp.Range("E2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 8, 1, 8, ss1.MaxRows - 1 '������;
    ss1.ClipboardCopy
    xlApp.Range("Y2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 54, 1, 54, ss1.MaxRows - 1 '����¼����    '43->45->46-48
    ss1.ClipboardCopy
    xlApp.Range("Z2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 67, 1, 67, ss1.MaxRows - 1 '��ͬ��
    ss1.ClipboardCopy
    xlApp.Range("AA2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 70, 1, 70, ss1.MaxRows - 1 '��������ԭ��
    ss1.ClipboardCopy
    xlApp.Range("AB2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection 71, 1, 71, ss1.MaxRows - 1 '����������
    ss1.ClipboardCopy
    xlApp.Range("AC2").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    ss1.ClearSelection
    
    For i = sRow_F To sRow_T
    
        sExlRange = "C" & i:        sOrd_Sts = xlApp.Range(sExlRange).Value          '����״̬
        sExlRange = "H" & i:        sOrd_Thk = xlApp.Range(sExlRange).Value          '���
        sExlRange = "I" & i:        sOrd_Wid = xlApp.Range(sExlRange).Value          '����
        sExlRange = "J" & i:        sOrd_Len = xlApp.Range(sExlRange).Value          '����
        sOrd_Wgt = Round((sOrd_Thk * sOrd_Wid * sOrd_Len * 7.85) / 1000000000, 3)    '��������
        sExlRange = "N" & i:        sOrd_RemWgt = xlApp.Range(sExlRange).Value       'Ƿ��(����)
        '            chr(79)
        sExlRange = "O" & i:        sOrd_ProcWgt(1) = xlApp.Range(sExlRange).Value   '����
        sExlRange = "P" & i:        sOrd_ProcWgt(2) = xlApp.Range(sExlRange).Value   '����
        sExlRange = "Q" & i:        sOrd_ProcWgt(3) = xlApp.Range(sExlRange).Value   '���ֵȴ�
        sExlRange = "R" & i:        sOrd_ProcWgt(4) = xlApp.Range(sExlRange).Value   '�����ȴ�
        sExlRange = "S" & i:        sOrd_ProcWgt(5) = xlApp.Range(sExlRange).Value   '���еȴ�
        sExlRange = "T" & i:        sOrd_ProcWgt(6) = xlApp.Range(sExlRange).Value   '�����ȴ�
        sExlRange = "U" & i:        sOrd_ProcWgt(7) = xlApp.Range(sExlRange).Value   '���ȴ�
        sExlRange = "V" & i:        sOrd_ProcWgt(8) = xlApp.Range(sExlRange).Value   '�������
        
        sExlRange = "O" & i
        '������������״̬���ǡ�������ϡ��͡�������ϡ��ġ�Ƿ��(����)����ֵ����
        '���㵥�أ��Ծ���ֵС�ڵ��صġ�Ƿ��(����)��������ֵ����
        If sOrd_Sts = "�������" Or sOrd_Sts = "�������" Or Abs(sOrd_RemWgt) < sOrd_Wgt Then xlApp.Range(sExlRange).Value = 0
        
        '��"���֡����������ֵȴ��������ȴ������еȴ��������ȴ�����ֵ��Ϊ�㣬�ѽ��롰������ϡ��ġ�Ƿ��(����)������
        If sOrd_ProcWgt(1) = 0 And sOrd_ProcWgt(2) = 0 And sOrd_ProcWgt(3) = 0 And sOrd_ProcWgt(4) = 0 And sOrd_ProcWgt(5) = 0 _
        And sOrd_ProcWgt(6) = 0 And sOrd_ProcWgt(7) = 0 And sOrd_ProcWgt(8) = 0 And sOrd_Sts = "�������" Then xlApp.Range(sExlRange).Value = 0
        
        '�����µġ�Ƿ��(����)����ֵ��"���֡����������ֵȴ��������ȴ������еȴ��������ȴ������ȴ���˳����"���֡����д��������㣬
        '��ֵ��Ϊ���������롰���������д��������㣬�Դ����ƣ�ֱ����ֵΪ�������Ϊֹ������ֵ�̶�����Ӧλ���ϡ�
        '�������������Ϊ�����㣻ȫ������Ƿ��(����)���㣨091201��
        If sOrd_RemWgt < 0 Then
            sOrd_CulWgt = sOrd_RemWgt
            For II = 1 To 8
                sOrd_CulWgt = sOrd_CulWgt + sOrd_ProcWgt(II)
                If sOrd_CulWgt >= 0 Then
                   sExlRange = Chr(78 + II) & i:                   xlApp.Range(sExlRange).Value = sOrd_CulWgt
                   Exit For
                Else
                   sExlRange = Chr(78 + II) & i:                   xlApp.Range(sExlRange).Value = 0
                End If
            Next II
                   sExlRange = "O" & i:                            xlApp.Range(sExlRange).Value = 0
        End If
                
    Next i
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_DZB
'   2.Name         : Spread Row Cancel (Insert, Update, Delete)
'   3.Input  Value : Conn Connection, Sc Collection
'   4.Return Value :
'   5.Writer       : YANGMENG
'   6.Create Date  : 2010. 01 .14
'   7.Modify Date  :
'   8.Comment      : Spread Row Cancel (Insert, Update, Delete)
'---------------------------------------------------------------------------------------
Private Sub Gp_Sp_DZB(Conn As ADODB.Connection, Sc As Collection)

On Error GoTo SpreadCancel_Error

    Dim sQuery As String
    Dim i As Integer
    Dim iRow, BR1, BR2 As Long

    With Sc
        
        Screen.MousePointer = vbHourglass
        .Item("Spread").ReDraw = False
        
        If .Item("Spread").MaxRows < 1 Or .Item("Spread").SelBlockRow < 1 Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        BR1 = 1
        BR2 = .Item("Spread").MaxRows
        
        For iRow = BR1 To BR2
            
                    sQuery = Gf_Sp_MakeQuery(.Item("Spread"), .Item("P-O"), "O", .Item("pColumn"), iRow)
                    Call Gp_Sp_OneRowDisplay(Conn, sQuery, .Item("Spread"), iRow)
                    Call Gp_Sp_SendData(.Item("Spread"), "", 0, iRow)
            
            If iRow = BR2 Then
                Exit For
            End If

        Next iRow
        
        .Item("Spread").ReDraw = True
        
    End With
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
SpreadCancel_Error:

    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Gp_Sp_Cancel Error : " & Error)
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_ACA1020C_Excel
'   2.Name         : Spread --> Excel
'   3.Input  Value : Fm Form, sPname Variant, bLkcol1 Long, bLkcol2 Long, bLkrow1 Long, bLkrow2 Long
'   4.Return Value :
'   5.Writer       : ����
'   6.Create Date  : 2010. 01 .18
'   7.Modify Date  :
'   8.Comment      : Spread --> Excel
'---------------------------------------------------------------------------------------
Private Sub Gp_ACA1020C_Excel(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

On Error GoTo Excel_Error

    Dim ret         As Boolean
    Dim xlApp       As Object
    Dim xlBpp       As Object
    Dim xlBook      As Object
    Dim xlSheet     As Object
    Dim ColIndex    As Integer
    Dim sExlRange   As String
    Dim sExlRange1  As String
    Dim iExlCol     As Integer
    
    'Call Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    With ss1
    
        If .MaxRows = 0 Then Exit Sub
        
        If bLkcol1 = 0 Then
           bLkcol1 = 1
        End If
        
        If bLkcol2 = 0 Then
            bLkcol2 = -1
        End If
        
        If bLkrow2 = 0 Then
            bLkrow2 = -1
        End If
        
        Clipboard.Clear
        
        .Col = bLkcol1: .Col2 = bLkcol2
        .ROW = bLkrow1: .Row2 = bLkrow2
        
        Clipboard.SetText .Clip
    
        xlApp.Visible = True
        
        xlSheet.Cells.NumberFormatLocal = "G/ͨ�ø�ʽ"
        xlSheet.Range("A1").Select
        xlSheet.Paste
        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
        
        sExlRange1 = ""
        For ColIndex = 1 To .MaxCols
            .Col = ColIndex
            .ROW = 1
            
            iExlCol = ColIndex
            If IsNumeric(.Text) And (Left(.Text, 1) = "0" Or Left(.Text, 1) = "1") And _
               (Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14) Then
                If ColIndex > 104 Then
                    sExlRange1 = "D" & sExlRange1
                    iExlCol = ColIndex - 104
                ElseIf ColIndex > 78 Then
                    sExlRange1 = "C" & sExlRange1
                    iExlCol = ColIndex - 78
                ElseIf ColIndex > 52 Then
                    sExlRange1 = "B" & sExlRange1
                    iExlCol = ColIndex - 52
                ElseIf ColIndex > 26 Then
                    sExlRange1 = "A"
                    iExlCol = ColIndex - 26
                End If
                
                sExlRange = sExlRange1 & Chr(iExlCol + 64) & "1:" & sExlRange1 & Chr(iExlCol + 64) & .MaxRows + 5
                If Len(.Text) = 8 Then
                    xlSheet.Range(sExlRange).NumberFormat = "00000000"
                ElseIf Len(.Text) = 10 Then
                    xlSheet.Range(sExlRange).NumberFormat = "0000000000"
                ElseIf Len(.Text) = 12 Then
                    xlSheet.Range(sExlRange).NumberFormat = "000000000000"
                ElseIf Len(.Text) = 14 Then
                    xlSheet.Range(sExlRange).NumberFormat = "00000000000000"
                End If
            End If
        Next
        
    End With
    
    With ss2
    
        If .MaxRows = 0 Then Exit Sub
        
        If bLkcol1 = 0 Then
           bLkcol1 = 1
        End If
        
        If bLkcol2 = 0 Then
            bLkcol2 = -1
        End If
        
        If bLkrow2 = 0 Then
            bLkrow2 = -1
        End If
        
        Clipboard.Clear
        
        .Col = bLkcol1: .Col2 = bLkcol2
        .ROW = bLkrow1: .Row2 = bLkrow2
        
        Clipboard.SetText .Clip
    
        xlApp.Visible = True
        
        xlSheet.Cells.NumberFormatLocal = "G/ͨ�ø�ʽ"
        xlSheet.Range("AQ1").Select
        xlSheet.Paste
        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
        
        sExlRange1 = ""
        For ColIndex = 1 To .MaxCols
            .Col = ColIndex
            .ROW = 1
            
            iExlCol = ColIndex
            If IsNumeric(.Text) And (Left(.Text, 1) = "0" Or Left(.Text, 1) = "1") And _
               (Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14) Then
                If ColIndex > 104 Then
                    sExlRange1 = "D" & sExlRange1
                    iExlCol = ColIndex - 104
                ElseIf ColIndex > 78 Then
                    sExlRange1 = "C" & sExlRange1
                    iExlCol = ColIndex - 78
                ElseIf ColIndex > 52 Then
                    sExlRange1 = "B" & sExlRange1
                    iExlCol = ColIndex - 52
                ElseIf ColIndex > 26 Then
                    sExlRange1 = "A"
                    iExlCol = ColIndex - 26
                End If
                
                sExlRange = sExlRange1 & Chr(iExlCol + 64) & "1:" & sExlRange1 & Chr(iExlCol + 64) & .MaxRows + 5
                If Len(.Text) = 8 Then
                    xlSheet.Range(sExlRange).NumberFormat = "00000000"
                ElseIf Len(.Text) = 10 Then
                    xlSheet.Range(sExlRange).NumberFormat = "0000000000"
                ElseIf Len(.Text) = 12 Then
                    xlSheet.Range(sExlRange).NumberFormat = "000000000000"
                ElseIf Len(.Text) = 14 Then
                    xlSheet.Range(sExlRange).NumberFormat = "00000000000000"
                End If
            End If
        Next
        
    End With
    
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    Exit Sub
    
Excel_Error:
    Call Gp_MsgBoxDisplay("���Ļ�����δ��װExcel", "W")

End Sub








