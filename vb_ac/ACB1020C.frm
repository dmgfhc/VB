VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB1020C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "���Ͽ����״��ѯ_ACB1020C"
   ClientHeight    =   10950
   ClientLeft      =   420
   ClientTop       =   1845
   ClientWidth     =   18105
   BeginProperty Font 
      Name            =   "����"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   18105
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin InDate.ULabel ULabel27 
      Height          =   315
      Left            =   6600
      Top             =   1695
      Width           =   1200
      _ExtentX        =   2117
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
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9150
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   16140
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "ACB1020C.frx":0000
      Begin Threed.SSFrame SSFrame1 
         Height          =   3285
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   5794
         _Version        =   196609
         BackColor       =   14737632
         ShadowStyle     =   1
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
            Left            =   7800
            MaxLength       =   3
            TabIndex        =   71
            Tag             =   "CD_MANA_NO"
            Top             =   1680
            Width           =   915
         End
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
            Left            =   5475
            MaxLength       =   3
            TabIndex        =   70
            Tag             =   "CD_MANA_NO"
            Top             =   1680
            Width           =   915
         End
         Begin InDate.ULabel ULabel26 
            Height          =   315
            Left            =   4245
            Top             =   1680
            Width           =   1200
            _ExtentX        =   2117
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
         Begin VB.TextBox txt_rep_remark 
            BackColor       =   &H00C0FFFF&
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
            Left            =   12510
            MaxLength       =   50
            TabIndex        =   67
            Tag             =   "������ע"
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox TXT_ORD_KND 
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
            Left            =   10080
            MaxLength       =   3
            TabIndex        =   64
            Tag             =   "CD_MANA_NO"
            Top             =   2160
            Width           =   915
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   4230
            TabIndex        =   60
            Top             =   30
            Width           =   4425
            Begin VB.OptionButton Opt_SHP_DATE 
               BackColor       =   &H00E0E0E0&
               Caption         =   "��������"
               Height          =   195
               Left            =   1500
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   480
               Width           =   1200
            End
            Begin VB.OptionButton Opt_IN_PLT_DATE 
               BackColor       =   &H00E0E0E0&
               Caption         =   "�������"
               Height          =   195
               Left            =   120
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   480
               Width           =   1200
            End
            Begin VB.OptionButton opt_del_date 
               BackColor       =   &H00E0E0E0&
               Caption         =   "��������"
               Height          =   195
               Left            =   3000
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   170
               Width           =   1200
            End
            Begin VB.OptionButton opt_mill_date 
               BackColor       =   &H00E0E0E0&
               Caption         =   "��������"
               Height          =   195
               Left            =   1500
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   170
               Width           =   1200
            End
            Begin VB.OptionButton opt_cut_date 
               BackColor       =   &H00E0E0E0&
               Caption         =   "��������"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   105
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   170
               Value           =   -1  'True
               Width           =   1200
            End
         End
         Begin VB.TextBox txt_date_fl 
            Height          =   285
            Left            =   8400
            TabIndex        =   59
            Top             =   270
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.TextBox txt_ORG_ORD_NO 
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
            Left            =   1380
            MaxLength       =   11
            TabIndex        =   58
            Tag             =   "CD_MANA_NO"
            Top             =   930
            Width           =   1400
         End
         Begin VB.ComboBox CBO_ORG_ORD_ITEM 
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
            Left            =   2805
            TabIndex        =   57
            Top             =   930
            Width           =   750
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
            ItemData        =   "ACB1020C.frx":0052
            Left            =   10080
            List            =   "ACB1020C.frx":005F
            TabIndex        =   56
            Top             =   135
            Width           =   915
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   315
            Left            =   11250
            TabIndex        =   52
            Top             =   525
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
            _Version        =   196609
            BackColor       =   14737632
            Begin VB.OptionButton Opt_rk_n 
               BackColor       =   &H00E0E0E0&
               Caption         =   "δ���"
               Height          =   195
               Left            =   2580
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   60
               Width           =   915
            End
            Begin VB.OptionButton Opt_all 
               BackColor       =   &H00E0E0E0&
               Caption         =   "ȫ��"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   390
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   60
               Value           =   -1  'True
               Width           =   750
            End
            Begin VB.OptionButton Opt_rk_y 
               BackColor       =   &H00E0E0E0&
               Caption         =   "���"
               Height          =   195
               Left            =   1515
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   60
               Width           =   690
            End
         End
         Begin VB.TextBox TXT_ORD_FL 
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
            Left            =   2850
            MaxLength       =   1
            TabIndex        =   49
            Tag             =   "CD_MANA_NO"
            Text            =   "A"
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox TXT_UST_FL 
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
            MaxLength       =   1
            TabIndex        =   47
            Tag             =   "CD_MANA_NO"
            Text            =   "N"
            Top             =   1800
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox TXT_BED_PILE_DATE 
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
            Left            =   3600
            MaxLength       =   1
            TabIndex        =   46
            Tag             =   "CD_MANA_NO"
            Text            =   "N"
            Top             =   1350
            Visible         =   0   'False
            Width           =   495
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
            Left            =   10080
            MaxLength       =   3
            TabIndex        =   45
            Tag             =   "CD_MANA_NO"
            Top             =   1320
            Width           =   915
         End
         Begin VB.ComboBox CBO_PLT 
            BackColor       =   &H00FFFFFF&
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
            ItemData        =   "ACB1020C.frx":006F
            Left            =   3390
            List            =   "ACB1020C.frx":0082
            TabIndex        =   42
            Text            =   "C1"
            Top             =   150
            Width           =   690
         End
         Begin VB.ComboBox CBO_PROD_CD 
            BackColor       =   &H00C0FFFF&
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
            ItemData        =   "ACB1020C.frx":009A
            Left            =   1380
            List            =   "ACB1020C.frx":00A7
            TabIndex        =   41
            Text            =   "PP"
            Top             =   150
            Width           =   750
         End
         Begin VB.TextBox TXT_MAT_NO 
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
            Left            =   5475
            MaxLength       =   14
            TabIndex        =   29
            Tag             =   "���Ϻ�"
            Top             =   2040
            Width           =   1485
         End
         Begin VB.TextBox TXT_LOT_NO 
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
            Left            =   1380
            TabIndex        =   28
            Tag             =   "������"
            Top             =   2865
            Width           =   1605
         End
         Begin VB.TextBox TXT_HTM 
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
            Left            =   10080
            MaxLength       =   1
            TabIndex        =   27
            Top             =   1710
            Width           =   915
         End
         Begin VB.TextBox txt_TRIM_FL 
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
            Left            =   1380
            MaxLength       =   1
            TabIndex        =   26
            Tag             =   "����"
            Top             =   2100
            Width           =   525
         End
         Begin VB.TextBox txt_TRIM_NAME 
            Enabled         =   0   'False
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
            Left            =   1935
            TabIndex        =   25
            Tag             =   "����"
            Top             =   2100
            Width           =   1620
         End
         Begin VB.TextBox txt_prod_grd_name 
            Enabled         =   0   'False
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
            Left            =   1935
            TabIndex        =   24
            Tag             =   "����"
            Top             =   1710
            Width           =   1620
         End
         Begin VB.TextBox txt_prod_grd 
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
            Left            =   1380
            MaxLength       =   1
            TabIndex        =   23
            Top             =   1710
            Width           =   525
         End
         Begin VB.TextBox Text_size_knd 
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
            Left            =   1380
            MaxLength       =   2
            TabIndex        =   22
            Tag             =   "����"
            Top             =   2490
            Width           =   525
         End
         Begin VB.TextBox Text_size_knd_name 
            Enabled         =   0   'False
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
            Left            =   1935
            TabIndex        =   21
            Tag             =   "����"
            Top             =   2490
            Width           =   1620
         End
         Begin VB.TextBox text_cur_inv 
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
            Left            =   5970
            TabIndex        =   20
            Top             =   2460
            Width           =   1110
         End
         Begin VB.TextBox text_cur_inv_code 
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
            Left            =   5475
            MaxLength       =   2
            TabIndex        =   19
            Top             =   2460
            Width           =   450
         End
         Begin VB.TextBox txt_woo_rsn 
            BackColor       =   &H00C0FFFF&
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
            Left            =   12510
            MaxLength       =   2
            TabIndex        =   18
            Tag             =   "���ԭ��"
            Top             =   930
            Width           =   735
         End
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
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   17
            Top             =   1320
            Width           =   1380
         End
         Begin VB.ComboBox CBO_ORD_ITEM 
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
            Left            =   2790
            TabIndex        =   16
            Top             =   540
            Width           =   750
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   11250
            TabIndex        =   12
            Top             =   30
            Width           =   3675
            Begin VB.OptionButton Option_ORD_FL_Y 
               BackColor       =   &H00E0E0E0&
               Caption         =   "����"
               Height          =   195
               Left            =   1515
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   170
               Width           =   690
            End
            Begin VB.OptionButton Option_ORD_FL_N 
               BackColor       =   &H00E0E0E0&
               Caption         =   "���"
               Height          =   195
               Left            =   2580
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   170
               Width           =   735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "ȫ��"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   390
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   170
               Value           =   -1  'True
               Width           =   750
            End
         End
         Begin VB.TextBox TXT_REC_STS 
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
            Left            =   10080
            MaxLength       =   1
            TabIndex        =   11
            Tag             =   "CD_MANA_NO"
            Text            =   "2"
            Top             =   525
            Width           =   915
         End
         Begin VB.TextBox Text_PROC_CD 
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
            Left            =   10080
            MaxLength       =   3
            TabIndex        =   10
            Tag             =   "CD_MANA_NO"
            Top             =   930
            Width           =   915
         End
         Begin VB.TextBox Text_STLGRD 
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
            Left            =   5475
            MaxLength       =   20
            TabIndex        =   9
            Tag             =   "����(��׼��)"
            Top             =   1290
            Width           =   3060
         End
         Begin VB.TextBox Text_LOC 
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
            Left            =   5475
            MaxLength       =   7
            TabIndex        =   8
            Tag             =   "CD_MANA_NO"
            Top             =   2850
            Width           =   1605
         End
         Begin VB.TextBox TXT_ORD_NO 
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
            Left            =   1380
            MaxLength       =   11
            TabIndex        =   7
            Tag             =   "CD_MANA_NO"
            Top             =   540
            Width           =   1380
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   4245
            Top             =   900
            Width           =   1200
            _ExtentX        =   2117
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
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   8820
            Top             =   525
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "��Ϣ״̬"
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
            Left            =   150
            Top             =   150
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "��Ʒ"
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   4245
            Top             =   1290
            Width           =   1200
            _ExtentX        =   2117
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
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Left            =   8820
            Top             =   930
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "����״̬"
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
            Left            =   4245
            Top             =   2850
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "��λ��"
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
         Begin CSTextLibCtl.sidbEdit sdb_thk_fr 
            Height          =   315
            Left            =   12510
            TabIndex        =   30
            Top             =   1680
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   -2147483640
            BackColor       =   16777215
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
            FmtControl      =   1
            NumDecDigits    =   2
            NumIntDigits    =   4
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.UDate DTP_PROD_FR 
            Height          =   315
            Left            =   5475
            TabIndex        =   31
            Tag             =   "INS_DATE"
            Top             =   900
            Width           =   1410
            _ExtentX        =   2487
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
         Begin InDate.UDate DTP_PROD_TO 
            Height          =   315
            Left            =   7140
            TabIndex        =   32
            Tag             =   "INS_DATE"
            Top             =   900
            Width           =   1410
            _ExtentX        =   2487
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
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   11250
            Top             =   1680
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "���"
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
         End
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   11250
            Top             =   2070
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "����"
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
         End
         Begin InDate.ULabel ULabel10 
            Height          =   315
            Left            =   11250
            Top             =   2460
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "����"
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
         End
         Begin CSTextLibCtl.sidbEdit sdb_wid_fr 
            Height          =   315
            Left            =   12510
            TabIndex        =   33
            Top             =   2070
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   -2147483640
            BackColor       =   16777215
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
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_len_fr 
            Height          =   315
            Left            =   12510
            TabIndex        =   34
            Top             =   2460
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   -2147483640
            BackColor       =   16777215
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
            NumIntDigits    =   7
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin InDate.ULabel ULabel11 
            Height          =   315
            Left            =   150
            Top             =   1320
            Width           =   1200
            _ExtentX        =   2117
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
         Begin CSTextLibCtl.sidbEdit sdb_thk_to 
            Height          =   315
            Left            =   13845
            TabIndex        =   35
            Top             =   1680
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   -2147483640
            BackColor       =   16777215
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
            FmtControl      =   1
            NumDecDigits    =   2
            NumIntDigits    =   4
            Undo            =   0
            Data            =   0
         End
         Begin CSTextLibCtl.sidbEdit sdb_wid_to 
            Height          =   315
            Left            =   13845
            TabIndex        =   36
            Top             =   2070
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   -2147483640
            BackColor       =   16777215
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
         Begin CSTextLibCtl.sidbEdit sdb_len_to 
            Height          =   315
            Left            =   13845
            TabIndex        =   37
            Top             =   2460
            Width           =   1020
            _Version        =   262145
            _ExtentX        =   1799
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   -2147483640
            BackColor       =   16777215
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
         Begin Threed.SSCommand cmd_fl_down 
            Height          =   375
            Left            =   13320
            TabIndex        =   38
            Top             =   900
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   661
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "��Ľ���"
            BevelWidth      =   3
         End
         Begin InDate.ULabel ULabel77 
            Height          =   315
            Left            =   11250
            Top             =   930
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "���ԭ��"
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
         Begin InDate.ULabel ULabel12 
            Height          =   315
            Left            =   4245
            Top             =   2460
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "�ѷŲֿ�"
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
            Left            =   150
            Top             =   2490
            Width           =   1200
            _ExtentX        =   2117
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
         Begin InDate.ULabel ULabel13 
            Height          =   315
            Left            =   150
            Top             =   1710
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "�ȼ�"
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
            Left            =   150
            Top             =   2100
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "�б�"
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
         Begin InDate.ULabel ULabel18 
            Height          =   315
            Left            =   8820
            Top             =   1710
            Width           =   1230
            _ExtentX        =   2170
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
            Left            =   13890
            TabIndex        =   39
            Top             =   2700
            Visible         =   0   'False
            Width           =   1485
            _ExtentX        =   2619
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
            Caption         =   "������ҵ����"
         End
         Begin InDate.ULabel ULabel19 
            Height          =   315
            Left            =   150
            Top             =   2865
            Width           =   1200
            _ExtentX        =   2117
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
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel20 
            Height          =   315
            Left            =   4245
            Top             =   2040
            Width           =   1200
            _ExtentX        =   2117
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
         Begin InDate.ULabel ULabel17 
            Height          =   315
            Left            =   2370
            Top             =   150
            Width           =   990
            _ExtentX        =   1746
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
         Begin Threed.SSCheck SSC_UST_FL 
            Height          =   285
            Left            =   7440
            TabIndex        =   43
            Top             =   2610
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
         Begin InDate.ULabel ULabel15 
            Height          =   315
            Left            =   8820
            Top             =   1320
            Width           =   1230
            _ExtentX        =   2170
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
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   8820
            Top             =   135
            Width           =   1230
            _ExtentX        =   2170
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
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   150
            Top             =   540
            Width           =   1200
            _ExtentX        =   2117
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
         End
         Begin InDate.ULabel ULabel21 
            Height          =   315
            Left            =   150
            Top             =   930
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            Caption         =   "ԭʼ������"
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
         Begin InDate.ULabel ULabel22 
            Height          =   315
            Left            =   8820
            Top             =   2160
            Width           =   1230
            _ExtentX        =   2170
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
         Begin InDate.ULabel ULabel24 
            Height          =   315
            Left            =   8820
            Top             =   2520
            Visible         =   0   'False
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "δ��������"
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
         Begin CSTextLibCtl.sidbEdit TXT_DZB_DATE 
            Height          =   315
            Left            =   10080
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   2520
            Visible         =   0   'False
            Width           =   495
            _Version        =   262145
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   125
            Text            =   " 0.00"
            ForeColor       =   16711680
            BackColor       =   16777215
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
            ReadOnly        =   -1  'True
            Insert          =   0   'False
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
            MaxValue        =   9999
            MinValue        =   0
            Undo            =   0
            Data            =   0
         End
         Begin Threed.SSCheck SSC_KEY_ORD_FL 
            Height          =   285
            Left            =   7440
            TabIndex        =   66
            Top             =   2880
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
            Caption         =   "�ص��ͬ"
         End
         Begin InDate.ULabel ULabel25 
            Height          =   315
            Left            =   11250
            Top             =   1320
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            Caption         =   "������ע"
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
         Begin Threed.SSCheck SSC_CROSS_MM 
            Height          =   285
            Left            =   7440
            TabIndex        =   68
            Top             =   2280
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
            Caption         =   "���º�ͬ"
         End
         Begin Threed.SSCheck SSC_URGNT_FL 
            Height          =   285
            Left            =   7440
            TabIndex        =   69
            Top             =   2040
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
            Caption         =   "��������"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Left            =   13650
            TabIndex        =   51
            Top             =   2550
            Width           =   90
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Left            =   13650
            TabIndex        =   50
            Top             =   2190
            Width           =   90
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Left            =   13650
            TabIndex        =   44
            Top             =   1770
            Width           =   90
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Left            =   6960
            TabIndex        =   40
            Top             =   1005
            Width           =   90
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   5835
         Left            =   0
         TabIndex        =   48
         Top             =   3315
         Width           =   15195
         _Version        =   393216
         _ExtentX        =   26802
         _ExtentY        =   10292
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   103
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACB1020C.frx":00B7
      End
   End
   Begin VB.TextBox Text_PROD_CD_Name 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   14190
      TabIndex        =   3
      Top             =   2565
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox Text_PROC_CD_Name 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   14550
      TabIndex        =   2
      Top             =   2265
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2535
      Left            =   13830
      Max             =   1
      Min             =   99
      TabIndex        =   0
      Top             =   2400
      Value           =   1
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox Text_ORD_ITEM 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "09"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   13650
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1830
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox Text_REC_STS_Name 
      Height          =   2535
      Left            =   13440
      MaxLength       =   2
      TabIndex        =   4
      Tag             =   "CD_MANA_NO"
      Top             =   2370
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "ACB1020C"
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
'-- Program ID        ACB1020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Yang Zhibin
'-- Date              2003.9.8
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE        EDITOR        DESCRIPTION
'-- 1.01  2003.09.08  Yang Zhibin   ���Ͽ����״��ѯ
'-- 1.02  2011.05.24  LiQian        ���Ͽ����״��ѯ�����Ӱ��������ڲ�ѯ����
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
Public STR1 As String
Public BASE As String
Public AIMNO As String
Public Refer_Fl As String
Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting

Dim sQuery As String
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

Dim iCount As Integer


Const SPD_WGT = 14   '11->12->14
Const SPD_SMP_FL = 18   '15->16->18
Const SPD_PROC_CD = 5
Const SPD_DEL_TO_DATE = 6
Const SPD_DATE_CD = 25  '��������/��������    / ��������   23->25
Const SPD_SPEC = 7  '��׼


Const SS1_PLATE_NO = 1
Const SS1_ORD_NO = 32   '20150331 ��25������������ʱ��  29--30->32
Const SS1_ORD_ITEM = 33 '20150331 ��25������������ʱ��  30--31->33
Const SS1_URGNT_FL = 82          '����������ɫ���  2012-11-07 by CaoLei    65->67    '20150331 ��25������������ʱ��  76--77->79
Const SS1_RH_FL = 31          '�Ƿ������      '20150331 ��25������������ʱ��  28--29->31
Const SS1_KEY_ORD_FL = 84         '�ص��ͬ   67->69     '20150331 ��25������������ʱ��  78--79 zhouyan   79->81
 

Private Sub Form_Define()

    Dim iRow As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(CBO_PROD_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(CBO_PLT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_REC_STS, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(DTP_PROD_FR, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(DTP_PROD_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(cbo_ord_item, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_cust_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_prod_grd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_ENDUSE_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(text_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(Text_PROC_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(Text_LOC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_thk_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_thk_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_wid_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_WID_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_len_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_LEN_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_mat_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_lot_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(TXT_HTM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_BED_PILE_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_UST_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_ORD_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_Trim_fl, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(Text_size_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_org_ord_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(CBO_ORG_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_date_fl, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_ord_knd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_DZB_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SSC_KEY_ORD_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SSC_URGNT_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SSC_CROSS_MM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
    
         
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    
    For iRow = 5 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, iRow, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Next iRow
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB1020C.P_SREFER", Key:="P-R"
    sc1.Add Item:="ACB1020C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="ACB1020C.P_MODIFY", Key:="P-M"
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
    
'    Call Gp_Sp_ColHidden(ss1, ss1.MaxCols - 1, True)
    Call Gp_Sp_ColHidden(ss1, 2, True)

End Sub






Private Sub CBO_PROD_CD_Click()
    If CBO_PROD_CD.Text = "SL" Then
       ss1.ROW = 0
       ss1.Col = SPD_SPEC
       ss1.Text = "��������"
       
       ss1.ROW = 0
       ss1.Col = 6  '  ############################################################################################################
       ss1.Text = "��������������"
       
       opt_cut_date.Value = True
       opt_mill_date.Value = False
       opt_del_date.Value = False
       Opt_IN_PLT_DATE.Value = False
       Opt_SHP_DATE.Value = False
       opt_cut_date.ForeColor = &HFF&
       opt_mill_date.ForeColor = &H80000012
       opt_del_date.ForeColor = &H80000012
       Opt_SHP_DATE.ForeColor = &H80000012
       Opt_IN_PLT_DATE.ForeColor = &H80000012
       
       txt_date_fl.Text = "C"
       ULabel1.Caption = "��������"
       ss1.ROW = 0
       ss1.Col = SPD_DATE_CD
       ss1.Text = "��������"
       
       
       ss1.ROW = 0
       ss1.Col = SS1_RH_FL
       ss1.Text = "�Ƿ������"
      
      
    Else
       ss1.ROW = 0
       ss1.Col = SPD_SPEC
       ss1.Text = "���⹤��"
       
       
       ss1.ROW = 0
       ss1.Col = 6  '  ############################################################################################################
       ss1.Text = "�ͻ�������"
       
        
       ss1.ROW = 0
       ss1.Col = SS1_RH_FL
       ss1.Text = "���۷�ʽ"
       
    End If
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
'    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    If App.Title = "CE" Then
        text_cur_inv_code.Text = "ZB"
        CBO_PROD_CD.Text = "SL"
        CBO_PLT.Text = "B1"
    ElseIf App.Title = "DE" Then
        text_cur_inv_code.Text = "00"
        CBO_PROD_CD.Text = "PP"
        CBO_PLT.Text = "C1"
    ElseIf App.Title = "CG" Then
        text_cur_inv_code.Text = "ZB"
        CBO_PROD_CD.Text = "PP"
        CBO_PLT.Text = "C3"
    ElseIf App.Title = "BG" Then
        text_cur_inv_code.Text = "00"
        CBO_PROD_CD.Text = "PP"
        CBO_PLT.Text = "C1"
    Else
        text_cur_inv_code.Text = "00"
        CBO_PROD_CD.Text = "PP"
        CBO_PLT.Text = "C1"
    End If
    
    TXT_REC_STS = "2"
    Option1.Value = True
    Call text_cur_inv_code_KeyUp(0, 0)
'    Call Gp_Sp_ColHidden(ss1, 42, True)         '40->41->42
'    Call Gp_Sp_ColHidden(ss1, 43, True)         '41->42->43

    Call Gp_Sp_ColHidden(ss1, 94, True)
    
    txt_date_fl.Text = "C"

    DTP_PROD_FR.Text = Format(DateAdd("d", -3, CDate(DTP_PROD_TO.Text)), "YYYY-MM-DD")

    Screen.MousePointer = vbDefault
    
End Sub



Private Sub chk_htm_shot_blast_Click(Value As Integer)

    If chk_htm_shot_blast Then
        Text_PROC_CD.Text = "DZB"
    End If
    
End Sub

Private Sub Combo_ORD_ITEM_LostFocus()
    
    Dim S As String
  
    If Len(cbo_ord_item.Text) = 1 Then
        S = cbo_ord_item.Text
        cbo_ord_item.Text = "0" + S
    End If
    
End Sub

Private Sub Opt_Click()

End Sub

Private Sub opt_all_Click()

    If opt_all.Value = True Then
        opt_all.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = ""
    Else
        opt_all.ForeColor = &H80000012
    End If
    
    If Opt_rk_y.Value = True Then
        Opt_rk_y.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = "Y"
    Else
        Opt_rk_y.ForeColor = &H80000012
    End If
    
    If Opt_rk_n.Value = True Then
        Opt_rk_n.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = "N"
    Else
        Opt_rk_n.ForeColor = &H80000012
    End If

End Sub

Private Sub opt_cut_date_Click()
    If opt_cut_date.Value = True Then
       opt_cut_date.ForeColor = &HFF&
       opt_mill_date.ForeColor = &H80000012
       opt_del_date.ForeColor = &H80000012
       Opt_IN_PLT_DATE.ForeColor = &H80000012
       Opt_SHP_DATE.ForeColor = &H80000012
       txt_date_fl.Text = "C"
       ULabel1.Caption = "��������"
'       ss1.Row = 0
'       ss1.Col = SPD_DATE_CD
'       ss1.Text = "��������"
    Else
       opt_mill_date.ForeColor = &HFF&
       opt_cut_date.ForeColor = &H80000012
       opt_mill_date.Value = True
       txt_date_fl.Text = "M"
       ULabel1.Caption = "��������"
'       ss1.Row = 0
'       ss1.Col = SPD_DATE_CD
'       ss1.Text = "��������"
    End If

End Sub

Private Sub opt_del_date_Click()
    If opt_del_date.Value = True Then
       opt_del_date.ForeColor = &HFF&
       opt_cut_date.ForeColor = &H80000012
       opt_mill_date.ForeColor = &H80000012
       Opt_IN_PLT_DATE.ForeColor = &H80000012
       Opt_SHP_DATE.ForeColor = &H80000012
       txt_date_fl.Text = "D"
       ULabel1.Caption = "��������"
'       ss1.Row = 0
'       ss1.Col = SPD_DATE_CD
'       ss1.Text = "��������"
    End If
End Sub

Private Sub Opt_IN_PLT_DATE_Click()
   If Opt_IN_PLT_DATE.Value = True Then
       Opt_IN_PLT_DATE.ForeColor = &HFF&
       opt_cut_date.ForeColor = &H80000012
       opt_mill_date.ForeColor = &H80000012
       opt_del_date.ForeColor = &H80000012
       Opt_SHP_DATE.ForeColor = &H80000012
       txt_date_fl.Text = "I"
       ULabel1.Caption = "�������"
    End If
End Sub
Private Sub Opt_SHP_DATE_Click()
   If Opt_SHP_DATE.Value = True Then
       Opt_SHP_DATE.ForeColor = &HFF&
       opt_cut_date.ForeColor = &H80000012
       opt_mill_date.ForeColor = &H80000012
       opt_del_date.ForeColor = &H80000012
       Opt_IN_PLT_DATE.ForeColor = &H80000012
       txt_date_fl.Text = "S"
       ULabel1.Caption = "��������"
    End If
End Sub

Private Sub opt_mill_date_Click()
    If opt_mill_date.Value = True Then
       opt_mill_date.ForeColor = &HFF&
       opt_cut_date.ForeColor = &H80000012
       opt_del_date.ForeColor = &H80000012
       Opt_IN_PLT_DATE.ForeColor = &H80000012
       Opt_SHP_DATE.ForeColor = &H80000012
       txt_date_fl.Text = "M"
       ULabel1.Caption = "��������"
'       ss1.Row = 0
'       ss1.Col = SPD_DATE_CD
'       ss1.Text = "��������"
    Else
       opt_cut_date.ForeColor = &HFF&
       opt_mill_date.ForeColor = &H80000012
       opt_cut_date.Value = True
       txt_date_fl.Text = "C"
       ss1.ROW = 0
'       ULabel1.Caption = "��������"
'       ss1.Col = SPD_DATE_CD
'       ss1.Text = "��������"
    End If
End Sub

Private Sub Opt_rk_n_Click()
    If opt_all.Value = True Then
        opt_all.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = ""
    Else
        opt_all.ForeColor = &H80000012
    End If
    
    If Opt_rk_y.Value = True Then
        Opt_rk_y.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = "Y"
    Else
        Opt_rk_y.ForeColor = &H80000012
    End If
    
    If Opt_rk_n.Value = True Then
        Opt_rk_n.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = "N"
    Else
        Opt_rk_n.ForeColor = &H80000012
    End If
End Sub

Private Sub Opt_rk_y_Click()
    If opt_all.Value = True Then
        opt_all.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = ""
    Else
        opt_all.ForeColor = &H80000012
    End If
    
    If Opt_rk_y.Value = True Then
        Opt_rk_y.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = "Y"
    Else
        Opt_rk_y.ForeColor = &H80000012
    End If
    
    If Opt_rk_n.Value = True Then
        Opt_rk_n.ForeColor = &HFF&
        TXT_BED_PILE_DATE.Text = "N"
    Else
        Opt_rk_n.ForeColor = &H80000012
    End If
End Sub



Private Sub Option_ORD_FL_N_Click()

    If Option_ORD_FL_Y.Value = True Then
        Option_ORD_FL_Y.ForeColor = &HFF&
        TXT_ORD_FL = "1"
    Else
        Option_ORD_FL_Y.ForeColor = &H80000012
    End If
    
    If Option_ORD_FL_N.Value = True Then
        Option_ORD_FL_N.ForeColor = &HFF&
        TXT_ORD_FL = "2"
    Else
        Option_ORD_FL_N.ForeColor = &H80000012
    End If
    
    If Option1.Value = True Then
        Option1.ForeColor = &HFF&
        TXT_ORD_FL = ""
    Else
        Option1.ForeColor = &H80000012
    End If

End Sub

Private Sub Option_ORD_FL_Y_Click()

    If Option_ORD_FL_Y.Value = True Then
        Option_ORD_FL_Y.ForeColor = &HFF&
        TXT_ORD_FL = "1"
    Else
        Option_ORD_FL_Y.ForeColor = &H80000012
    End If
    
    If Option_ORD_FL_N.Value = True Then
        Option_ORD_FL_N.ForeColor = &HFF&
        TXT_ORD_FL = "2"
    Else
        Option_ORD_FL_N.ForeColor = &H80000012
    End If
    
    If Option1.Value = True Then
        Option1.ForeColor = &HFF&
        TXT_ORD_FL = ""
    Else
        Option1.ForeColor = &H80000012
    End If

End Sub

Private Sub Option1_Click()

    If Option_ORD_FL_Y.Value = True Then
        Option_ORD_FL_Y.ForeColor = &HFF&
        TXT_ORD_FL = "1"
    Else
        Option_ORD_FL_Y.ForeColor = &H80000012
    End If
    
    If Option_ORD_FL_N.Value = True Then
        Option_ORD_FL_N.ForeColor = &HFF&
        TXT_ORD_FL = "2"
    Else
        Option_ORD_FL_N.ForeColor = &H80000012
    End If
    
    If Option1.Value = True Then
        Option1.ForeColor = &HFF&
        TXT_ORD_FL = ""
    Else
        Option1.ForeColor = &H80000012
    End If

End Sub

Private Sub Option2_Click()

End Sub

Private Sub sdb_len_fr_Change()
    If sdb_len_fr.Value > 0 And SDB_LEN_TO.Value < sdb_len_fr.Value Then
        SDB_LEN_TO.Value = sdb_len_fr.Value
    End If
End Sub

Private Sub sdb_thk_fr_Change()
    If sdb_thk_fr.Value > 0 And sdb_thk_to.Value < sdb_thk_fr.Value Then
        sdb_thk_to.Value = sdb_thk_fr.Value
    End If
End Sub

Private Sub sdb_wid_fr_Change()
    If sdb_wid_fr.Value > 0 And SDB_WID_TO.Value < sdb_wid_fr.Value Then
        SDB_WID_TO.Value = sdb_wid_fr.Value
    End If
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    
End Sub

Private Sub SSC_UST_FL_Click(Value As Integer)
    If SSC_UST_FL.Value = -1 Then
       SSC_UST_FL.ForeColor = &HFF&
       TXT_UST_FL = "Y"
    Else
       SSC_UST_FL.ForeColor = &H808080
       TXT_UST_FL = "N"
    End If
End Sub

Private Sub text_cur_inv_code_DblClick()

    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_PROC_CD_Change()

    If Not Text_PROC_CD.Text = "" Then
        If Len(Text_PROC_CD.Text) = Text_PROC_CD.MaxLength Then
            Text_PROC_CD.Text = StrConv(Text_PROC_CD.Text, vbUpperCase)
        End If
        If Text_PROC_CD.Text <> "DZB" Then
            chk_htm_shot_blast.Value = False
        End If
    End If

End Sub

Private Sub Text_PROC_CD_DblClick()

    Call Text_PROC_CD_KeyUp(vbKeyF4, 0)
    
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

Private Sub TXT_HTM_DblClick()

    Call TXT_HTM_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TXT_HTM_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Q0073"
        
        DD.rControl.Add Item:=TXT_HTM
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        If TXT_HTM.Text <> "" Then
            Text_PROC_CD.Text = "DZB"
        End If
        
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

Private Sub TXT_REC_STS_Change()
 
    If Not TXT_REC_STS.Text = "" Then
        If Not TXT_REC_STS.Text = "1" Then
            If Not TXT_REC_STS.Text = "2" Then
                If Not TXT_REC_STS.Text = "3" Then
    '               Call MsgBox("״̬����" & Chr(10) & "�����Ϲ淶! �������", vbExclamation + vbOKOnly, "����")
                    TXT_REC_STS.Text = ""
                End If
            End If
        End If
    End If
    
End Sub

Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuTool_ReSet
    
    Select Case Mid(sAuthority, 2, 3) 'Insert, Update, Delete

           Case "000"      'No Authority
             cmd_fl_down.Enabled = False
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
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

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        cbo_ord_item.Clear
    End If
    
    Option_ORD_FL_Y.Value = False
    Option_ORD_FL_N.Value = False
    Option1.Value = True
    txt_woo_rsn.Text = ""
    txt_rep_remark.Text = ""
    Text_PROD_CD_Name.Text = ""
    Text_PROC_CD_Name.Text = ""
    Text_REC_STS_Name.Text = ""
    
    If App.Title = "CE" Then
        text_cur_inv_code.Text = "ZB"
        CBO_PROD_CD.Text = "SL"
    End If
    
    Call text_cur_inv_code_KeyUp(0, 0)
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1, True) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        'Call Gp_Sp_EvenRowBackcolor(Proc_Sc("SC").Item("Spread"), 1)
    End If
       
End Sub

Public Sub Form_Ref()                   '###################################################################################################

    Dim sMesg As String
    Dim S As String
    
    Dim iRow As Long
    Dim iSumWgt As Double
    
    Dim iCount As Long
    Dim sCurDate As String
    Dim sDel_To_Date As String
    Dim sproc_cd As String

    
    If CBO_PROD_CD.Text = "PP" Then
        Call Gp_Sp_ColHidden(ss1, 3, False)
        Call Gp_Sp_ColHidden(ss1, 4, False)
    Else
        Call Gp_Sp_ColHidden(ss1, 3, True)
        Call Gp_Sp_ColHidden(ss1, 4, True)
    End If
    
    ss1.ROW = 0
    ss1.Col = SPD_SMP_FL
    
    If CBO_PROD_CD.Text = "SL" Then
        ss1.Text = "�ɷֵȼ�"
    Else
        ss1.Text = "��������"
    End If
     
    If Refer_Fl = "Y" Then
        If Len(txt_mat_no.Text) < 8 And Len(txt_lot_no.Text) < 12 And txt_ord_no.Text = "" And cbo_ord_item.Text = "" And (DTP_PROD_FR.RawData = "" Or DTP_PROD_TO.RawData = "") Then
            Call Gp_MsgBoxDisplay("�������ڻ򶩵��Ż����ϺŻ������ű�������", "I", "������ʾ")
            Exit Sub
        End If
    End If
    
    Refer_Fl = "Y"
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
   txt_woo_rsn.Text = ""
   txt_rep_remark.Text = ""
   
   With ss1
   
       If .MaxRows < 1 Then
           Exit Sub
       End If
       If CBO_PROD_CD.Text <> "SL" Then
       sCurDate = Format(Now, "YYYYMM")
       For iCount = 1 To .MaxRows

            '�������ھ�ʾ
            .ROW = iCount:            .Col = SPD_DEL_TO_DATE
            sDel_To_Date = Mid(.Value, 1, 6)
            If sDel_To_Date < sCurDate Then
              .ROW = iCount:           .Col = SPD_PROC_CD
              sproc_cd = Mid(.Value, 1, 1)
              If sproc_cd <> "X" Then
                 Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iCount, iCount, &HFF&)
              End If
            End If
            
         
            
             '�ص��ͬ
            ss1.ROW = .ROW:       ss1.Col = SS1_KEY_ORD_FL
            If ss1.Text = "Y" Then
'                 Call Gp_Sp_BlockColor(ss1, 0, .MaxCols, .Row, .Row, &HFF&)
                 Call Gp_Sp_RowColor(ss1, .ROW, , &HFF&)
            End If
            
               '����������ɫ��� 2012-11-07  by  CaoLei
            ss1.ROW = .ROW:       ss1.Col = SS1_URGNT_FL
            If ss1.Text = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, SS1_PLATE_NO, SS1_PLATE_NO, .ROW, .ROW, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_ORD_NO, SS1_ORD_NO, .ROW, .ROW, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_ORD_ITEM, SS1_ORD_ITEM, .ROW, .ROW, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_URGNT_FL, SS1_URGNT_FL, .ROW, .ROW, &HC000&)
            End If
            
        Next iCount
       End If
       
        .MaxRows = ss1.MaxRows + 1
          .ROW = .MaxRows:         .Col = 1
          .Text = "�ϼ�"
            For iRow = 1 To .MaxRows - 1
                .ROW = iRow
                .Col = SPD_WGT
                iSumWgt = iSumWgt + Val(.Text)
            Next iRow
          .ROW = ss1.MaxRows:          .Col = SPD_WGT
           iSumWgt = Round(iSumWgt, 3)
          .Text = iSumWgt
       
   End With

   

End Sub


Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    'Call Gp_Sp_EvenRowBackcolor(Proc_Sc("SC").Item("Spread"), 1)
    
End Sub

Private Sub SS1_CHANGE_COLOR()

   Dim Num As Integer
   Dim ordno As String
   Dim flag As Integer
   ordno = ""
   flag = 2
   Num = 0

    With ss1

 If .MaxRows <= 0 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
        .ROW = iCount
        ss1.ROW = .ROW:   ss1.Col = 3
            If ordno = "" Or ordno <> ss1.Text Then
               ordno = ss1.Text
               flag = flag + 1
               Num = flag Mod 2
            End If

            If Num = 1 Then     '&H000000��ʾ��ɫ
            Call Gp_Sp_BlockColor(ss1, 3, 3, .ROW, .ROW, vbBlue)
'
            Else
             Call Gp_Sp_BlockColor(ss1, 3, 3, .ROW, .ROW, &H0)
'
            End If

        Next iCount

    End With

End Sub



'Public Sub Spread_ColumnsSort()
'
''    Spread_ColSort.Show 1
'''    If ss1.ActiveCol = 3 Then
'''    Call SS1_CHANGE_COLOR
''    End If
''
'
'End Sub

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

Private Sub text_cur_inv_code_Change()

    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
          text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
          Exit Sub
    Else
          text_cur_inv.Text = ""
    End If
    
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
           DD.sWitch = "MS"
           DD.sKey = "C0013"
    
           DD.rControl.Add Item:=text_cur_inv_code
           DD.rControl.Add Item:=text_cur_inv
           
    
           DD.nameType = "2"
           Call Gf_Common_DD(M_CN1, KeyCode)
    
    End If
    
    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
        text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
        Exit Sub
    Else
        text_cur_inv.Text = ""
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
    
    If ss1.Text <> "����" Then
        
        ss1.Text = "����"
        
        Call Gp_Sp_BlockColor(ss1, 1, -1, ROW, ROW, , &HFFFF80)
    Else
       
        ss1.Text = ""
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW)
       
    End If
    ss1.ReDraw = True
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)
    
    If ss1.ActiveCol = 3 And ss1.ActiveRow = 1 Then
          Call SS1_CHANGE_COLOR
    End If
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
'   Call ss1_row_Click(Col, Row)

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

Private Sub Text_ORD_ITEM_Change()

    If Text_ORD_ITEM.Text <> "" Then
        If Val(Text_ORD_ITEM.Text) > iCount Or Val(Text_ORD_ITEM.Text) < 0 Or Text_ORD_ITEM.Text = "00" Then
            Call MsgBox("����������벻��ȷ!" & Chr(10) & "�����ԡ�", vbExclamation + vbOKOnly, "����")
            Text_ORD_ITEM.Text = ""
        End If
    End If

End Sub

Private Sub Text_PROC_CD_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "C0004"

        DD.rControl.Add Item:=Text_PROC_CD
        DD.rControl.Add Item:=Text_PROC_CD_Name
   
        DD.nameType = "2"
        'DD.nameType="1" ���������Ʋ�ѯ
        'DD.nameType="2" ��Ӣ�����Ʋ�ѯ
       
        Call Gf_Common_DD(M_CN1, KeyCode)

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        ' Gf_Customer_DD() ���ڿͻ�����
        Exit Sub
        
    End If

    If Len(Trim(Text_PROC_CD.Text)) = Text_PROC_CD.MaxLength Then
       '  Gf_ComnNAME_Find( �����ַ���, DD.sKEy���� ,DD.nameType)
       ' Gf_CustNameFind( �����ַ���, �ͻ���������,DD.nameType)
        Text_PROC_CD_Name.Text = Gf_ComnNameFind(M_CN1, "C0004", Text_PROC_CD.Text, 2)
    Else
        Text_PROC_CD_Name.Text = ""
    End If
    
End Sub

Private Sub Text_size_knd_DblClick()

    Call Text_size_knd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_stlgrd_DblClick()

    Call text_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=text_stlgrd
           
        If CBO_PROD_CD.Text = "SL" Then
            DD.nameType = "1"
            Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        Else
            Call Gf_StdSPEC_DD2(M_CN1, KeyCode)
        End If
    End If
        
End Sub

Private Sub txt_cust_cd_DblClick()

    Call txt_cust_cd_KeyUp(vbKeyF4, 0)
    
End Sub


Private Sub txt_prod_grd_Change()

    If Len(Trim(txt_prod_grd.Text)) = txt_prod_grd.MaxLength Then
        txt_prod_grd_name.Text = Gf_ComnNameFind(M_CN1, "Q0034", txt_prod_grd.Text, 1)
        Exit Sub
    Else
        txt_prod_grd_name.Text = ""
    End If
    
End Sub

Private Sub txt_prod_grd_DblClick()

    Call txt_prod_grd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_prod_grd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0034"

        DD.rControl.Add Item:=txt_prod_grd

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
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





Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)
 
    If Col = 3 Or Col = 4 Then Exit Sub
    
    If ROW > 0 And Col > 0 Then
    
        If ss1.MaxRows = ROW Then Exit Sub
    
        Unload ACB1030C
    
        ss1.Col = 1
        ss1.ROW = ROW
        AIMNO = Trim(ss1.Text)
        BASE = Trim(CBO_PROD_CD.Text)
        STR1 = Trim(sQuery)
       ' ACB1030C.TXT_FORM_NAME.Text = "ACB1020C"
        ACB1030C.FORM_A = "ACB1020C"
        ACB1030C.Show
        
    End If

End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_cust_cd

        DD.nameType = "1"
        Call Gf_Customer_DD(M_CN1, KeyCode)

    End If

End Sub

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String

    If Len(Trim(txt_ord_no.Text)) = txt_ord_no.MaxLength Then
    
        If cbo_ord_item.Text <> "" Then Exit Sub
        
        txt_ord_no.Text = StrConv(txt_ord_no.Text, vbUpperCase)
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(txt_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, cbo_ord_item, sQuery)

    Else
        cbo_ord_item.Clear
    End If

End Sub

'Private Sub txt_ord_no_LostFocus()

'    If TXT_ORD_NO.Text <> "" Then
'       If (Len(TXT_ORD_NO.Text) < TXT_ORD_NO.MaxLength) Then
'          Call Gp_MsgBoxDisplay("����������δ��ɣ�")
'          Text_ORD_ITEM.Text = ""
'          TXT_ORD_NO.SetFocus
'       End If
'    End If

'End Sub

Private Sub txt_trim_fl_Change()

    If Len(Trim(txt_Trim_fl.Text)) = txt_Trim_fl.MaxLength Then
        txt_Trim_NAME.Text = Gf_ComnNameFind(M_CN1, "B0021", txt_Trim_fl.Text, 2)
        txt_Trim_fl.Text = Trim(txt_Trim_fl.Text)
        Exit Sub
    Else
        txt_Trim_NAME.Text = ""
        txt_Trim_fl.Text = ""
    End If

End Sub

Private Sub txt_trim_fl_DblClick()

    Call txt_trim_fl_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_trim_fl_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0021"

        DD.rControl.Add Item:=txt_Trim_fl

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If

End Sub

Private Sub txt_woo_rsn_DblClick()

    Call txt_woo_rsn_KeyUp(vbKeyF4, 0)
    
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

Private Sub txt_woo_rsn_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0008"
        DD.rControl.Add Item:=txt_woo_rsn
        DD.rControl.Add Item:=txt_rep_remark
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
     Else
    
        If Len(Trim(txt_woo_rsn.Text)) = txt_woo_rsn.MaxLength Then
            txt_rep_remark.Text = Gf_ComnNameFind(M_CN1, "C0008", txt_woo_rsn.Text, 2)
        Else
            txt_rep_remark.Text = ""
        End If
        
    End If

End Sub

Private Sub cmd_fl_down_Click()

'On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    Dim idel As Integer
    Dim str_prod_cd As String
    Dim inum As Integer
    Dim str_slab_no As String
    Dim isel As Integer
    
    inum = 0
    
    If ss1.MaxRows <= 0 Then Exit Sub
    
     If Trim(txt_woo_rsn.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_woo_rsn.Tag + "��������")
        Exit Sub
    End If
    
    If Len(Trim(txt_woo_rsn.Text)) <> txt_woo_rsn.MaxLength Then
         Call Gp_MsgBoxDisplay(txt_woo_rsn.Tag + " �������")
        Exit Sub
    End If
    
    If Trim(txt_rep_remark.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_rep_remark.Tag + "��������")
        Exit Sub
    End If
    
'    If Len(Trim(txt_rep_remark.Text)) > txt_rep_remark.MaxLength Then
'         Call Gp_MsgBoxDisplay(txt_rep_remark.Tag + " 50 ���ַ���")
'        Exit Sub
'    End If
    
    Dim adoCmd As ADODB.Command
    
     Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    
    With ss1
      For idel = 1 To .MaxRows
        .ROW = idel
        .Col = 0
        If .Text = "����" Then
              inum = inum + 1
              .Col = 1
              str_prod_cd = .Text
              
              
               'COIL
              If Trim(CBO_PROD_CD.Text) = "HC" Then
                  sQuery = "{call ACE2010P ('" + str_prod_cd + "','" + txt_woo_rsn.Text + "','" + txt_rep_remark.Text + "','Y','" + sUserID + "',?)}"
              'PLATE
              ElseIf Trim(CBO_PROD_CD.Text) = "PP" Then
                  sQuery = "{call ACE2020P ('" + str_prod_cd + "','" + txt_woo_rsn.Text + "','" + txt_rep_remark.Text + "','Y','" + sUserID + "',?)}"
              Else
              'SLAB
                  sQuery = "{call ACE2030P ('" + str_prod_cd + "','" + txt_woo_rsn.Text + "','" + txt_rep_remark.Text + "','Y','" + sUserID + "',?)}"
              End If
                        
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
                Screen.MousePointer = vbDefault

                Call Form_Ref
                
                For isel = 1 To .MaxRows
                   .ROW = isel
                   .Col = 1
                   If .Text = str_prod_cd Then
                     Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, isel, isel, , &HFFFF80)
                    Exit Sub
                    End If
                Next
                
                Exit Sub
            End If
            
'            .Text = ""
    End If
    Next
    End With
    
    If inum = 0 Then
        Call Gp_MsgBoxDisplay("û��ѡ��������!!", "I")
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Call Gp_MsgBoxDisplay("��Ľ�������..!!", "I")
    Call Form_Ref
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Form_Ref
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(11).Enabled = False                'Spread Copy
        .Buttons(12).Enabled = False                'Paste
    End With

End Sub