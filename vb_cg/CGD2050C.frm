VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGD2050C 
   Caption         =   "表面检查实绩查询及修改_CGD2050C"
   ClientHeight    =   9405
   ClientLeft      =   585
   ClientTop       =   1680
   ClientWidth     =   15525
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   15525
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_EMP_CD1 
      Enabled         =   0   'False
      Height          =   330
      Left            =   8325
      MaxLength       =   7
      TabIndex        =   145
      Tag             =   "作业人员"
      Top             =   9510
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox txt_ResonCd 
      Height          =   285
      Left            =   16140
      TabIndex        =   110
      Text            =   " "
      Top             =   600
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.ComboBox cbo_ResonDesc 
      Height          =   315
      ItemData        =   "CGD2050C.frx":0000
      Left            =   16380
      List            =   "CGD2050C.frx":0002
      TabIndex        =   109
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TXT_INSP_FLAW 
      Height          =   315
      Index           =   1
      Left            =   600
      TabIndex        =   55
      Top             =   9945
      Visible         =   0   'False
      Width           =   285
   End
   Begin Threed.SSFrame SF4 
      Height          =   4755
      Left            =   10020
      TabIndex        =   29
      Top             =   3810
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   8387
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   16711680
      BackColor       =   14737632
      Caption         =   "判定"
      Begin VB.ComboBox COM_PF 
         Height          =   315
         ItemData        =   "CGD2050C.frx":0004
         Left            =   1290
         List            =   "CGD2050C.frx":0014
         TabIndex        =   163
         Top             =   990
         Width           =   1215
      End
      Begin VB.CheckBox CHK_FLAW_YN 
         BackColor       =   &H00E0E0E0&
         Caption         =   "下表是否检验"
         Height          =   240
         Left            =   3300
         TabIndex        =   161
         Tag             =   "G"
         Top             =   3570
         Width           =   1620
      End
      Begin VB.TextBox txt_Color_code 
         Height          =   300
         Left            =   1350
         MaxLength       =   2
         TabIndex        =   154
         Tag             =   "原因"
         Top             =   3540
         Width           =   405
      End
      Begin VB.TextBox txt_Color_name 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   153
         Top             =   3540
         Width           =   1395
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   103
         Top             =   3150
         Width           =   1035
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   2
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   3150
         Width           =   2595
      End
      Begin VB.TextBox TXT_PROC_CD 
         Alignment       =   2  'Center
         BackColor       =   &H00E1E4CD&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   69
         Tag             =   "表面判定"
         Text            =   " "
         Top             =   1785
         Width           =   840
      End
      Begin CSTextLibCtl.sidbEdit SDB_Mn 
         Height          =   225
         Left            =   1320
         TabIndex        =   68
         Top             =   1380
         Width           =   840
         _Version        =   262145
         _ExtentX        =   1482
         _ExtentY        =   397
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   255
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
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
         Text            =   ""
         StartText.x     =   2
         StartText.y     =   0
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin VB.TextBox txt_Scrap_name 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3615
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   1725
         Width           =   1395
      End
      Begin VB.TextBox txt_Scrap_code 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3195
         MaxLength       =   1
         TabIndex        =   63
         Tag             =   "原因"
         Top             =   1725
         Width           =   405
      End
      Begin VB.TextBox txt_stdspec_yy 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3690
         MaxLength       =   40
         TabIndex        =   62
         Tag             =   "STDSPEC"
         Top             =   2010
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txt_stdspec_name_chg 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2130
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   61
         Tag             =   "STDSPEC"
         Top             =   2790
         Width           =   2865
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
         Height          =   330
         Left            =   165
         MaxLength       =   18
         TabIndex        =   60
         Tag             =   "标准号"
         Top             =   2790
         Width           =   1965
      End
      Begin VB.TextBox txt_stdspec_name 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2130
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   59
         Tag             =   "STDSPEC"
         Top             =   2430
         Width           =   2865
      End
      Begin VB.TextBox txt_stdspec 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   58
         Tag             =   "标准代码"
         Top             =   2430
         Width           =   1965
      End
      Begin VB.TextBox TXT_SURF_GRD 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   53
         Tag             =   "表面判定"
         Text            =   " "
         Top             =   270
         Width           =   840
      End
      Begin VB.TextBox TXT_INSP_MAIN_GRD 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   30
         Tag             =   "表面等级判定"
         Top             =   630
         Width           =   840
      End
      Begin InDate.ULabel ULabel22 
         Height          =   330
         Index           =   0
         Left            =   165
         Top             =   630
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         Caption         =   "表面等级判定"
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
      Begin InDate.ULabel ULabel36 
         Height          =   330
         Left            =   165
         Top             =   270
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         Caption         =   "表面判定"
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
         Height          =   300
         Index           =   1
         Left            =   165
         Top             =   2070
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   529
         Caption         =   "标准号"
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Left            =   2550
         Top             =   1725
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   529
         Caption         =   "原因"
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
      Begin InDate.ULabel ULabel22 
         Height          =   300
         Index           =   2
         Left            =   165
         Top             =   1350
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   529
         Caption         =   "Mn 成分 (         )"
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
      Begin InDate.ULabel ULabel22 
         Height          =   300
         Index           =   3
         Left            =   165
         Top             =   1725
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   529
         Caption         =   "进   程 (         )"
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
      Begin Threed.SSFrame SSFrame3 
         Height          =   315
         Left            =   2520
         TabIndex        =   80
         Top             =   270
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   3300
            TabIndex        =   81
            Text            =   " "
            Top             =   30
            Width           =   225
         End
         Begin Threed.SSOption opt_CHK_SUR_GRD 
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   82
            Top             =   30
            Width           =   735
            _ExtentX        =   1296
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
            Caption         =   "合格"
         End
         Begin Threed.SSOption opt_CHK_SUR_GRD 
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   83
            Top             =   30
            Width           =   885
            _ExtentX        =   1561
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
            Caption         =   "不合格"
         End
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   1065
         Left            =   2520
         TabIndex        =   84
         Top             =   630
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1879
         _Version        =   196609
         BackColor       =   14737632
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   255
            Index           =   4
            Left            =   1530
            TabIndex        =   89
            Top             =   390
            Width           =   735
            _ExtentX        =   1296
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
            Caption         =   "次品"
         End
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   85
            Top             =   60
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
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
            Caption         =   "正品"
         End
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   86
            Top             =   390
            Width           =   885
            _ExtentX        =   1561
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
            Caption         =   "改判"
         End
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   87
            Top             =   690
            Width           =   885
            _ExtentX        =   1561
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
            Caption         =   "协议板"
         End
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   255
            Index           =   3
            Left            =   1530
            TabIndex        =   88
            Top             =   60
            Width           =   735
            _ExtentX        =   1296
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
            Caption         =   "待判"
         End
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   255
            Index           =   5
            Left            =   1530
            TabIndex        =   90
            Top             =   690
            Width           =   735
            _ExtentX        =   1296
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
            Caption         =   "废钢"
         End
      End
      Begin InDate.ULabel ULabel25 
         Height          =   315
         Left            =   165
         Top             =   3150
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "改判缺陷"
         Alignment       =   1
         BackColor       =   8421631
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
      Begin InDate.ULabel ULabel32 
         Height          =   315
         Left            =   165
         Top             =   3540
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "表面颜色"
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
      Begin InDate.ULabel ULabel35 
         Height          =   315
         Left            =   180
         Top             =   3930
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_HD1 
         Height          =   315
         Left            =   840
         TabIndex        =   155
         Top             =   3930
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
      Begin InDate.ULabel ULabel39 
         Height          =   315
         Left            =   1680
         Top             =   3930
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_HD2 
         Height          =   315
         Left            =   2310
         TabIndex        =   156
         Top             =   3930
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
      Begin InDate.ULabel ULabel40 
         Height          =   315
         Left            =   3120
         Top             =   3930
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_HD3 
         Height          =   315
         Left            =   3750
         TabIndex        =   157
         Top             =   3930
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
      Begin InDate.ULabel ULabel41 
         Height          =   315
         Left            =   180
         Top             =   4320
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_HD4 
         Height          =   315
         Left            =   870
         TabIndex        =   158
         Top             =   4320
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
      Begin InDate.ULabel ULabel42 
         Height          =   315
         Left            =   1680
         Top             =   4320
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         Caption         =   "厚度5"
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
      Begin CSTextLibCtl.sidbEdit SDB_HD5 
         Height          =   315
         Left            =   2310
         TabIndex        =   159
         Top             =   4320
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
      Begin InDate.ULabel ULabel46 
         Height          =   315
         Left            =   3120
         Top             =   4320
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_HD6 
         Height          =   315
         Left            =   3750
         TabIndex        =   160
         Top             =   4320
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
      Begin InDate.ULabel ULabel22 
         Height          =   330
         Index           =   9
         Left            =   165
         Top             =   990
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   582
         Caption         =   "判废库"
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
   Begin Threed.SSFrame sf3 
      Height          =   4755
      Left            =   60
      TabIndex        =   15
      Top             =   3810
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   8387
      _Version        =   196609
      Font3D          =   2
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
      Caption         =   " 尺寸"
      Begin VB.TextBox TXT_WAVE1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3990
         MaxLength       =   2
         TabIndex        =   150
         Top             =   3180
         Width           =   930
      End
      Begin VB.TextBox TXT_SIZE_KND 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1170
         MaxLength       =   2
         TabIndex        =   149
         Tag             =   "原因"
         Top             =   3960
         Width           =   840
      End
      Begin VB.TextBox TXT_SIZE_KND_NAME 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   148
         Tag             =   "钢种"
         Top             =   3960
         Width           =   1050
      End
      Begin VB.TextBox TXT_WAVE 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1170
         MaxLength       =   2
         TabIndex        =   133
         Top             =   3180
         Width           =   840
      End
      Begin VB.TextBox TXT_RECT_DEG 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   132
         Top             =   3540
         Width           =   840
      End
      Begin VB.TextBox TXT_VERT_DEG 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1170
         MaxLength       =   2
         TabIndex        =   131
         Top             =   3540
         Width           =   840
      End
      Begin VB.CheckBox CHK_CL_FL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "矫直指示"
         Height          =   315
         Left            =   3330
         TabIndex        =   130
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox TXT_INSP_WGT_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4095
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2415
         Width           =   855
      End
      Begin VB.TextBox TXT_INSP_THK_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2415
         Width           =   810
      End
      Begin VB.TextBox TXT_INSP_LEN_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2415
         Width           =   1035
      End
      Begin VB.TextBox TXT_INSP_WID_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2415
         Width           =   990
      End
      Begin InDate.ULabel ULabel28 
         Height          =   315
         Left            =   2010
         Top             =   285
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "宽度"
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
      Begin InDate.ULabel ULabel29 
         Height          =   315
         Left            =   1170
         Top             =   285
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         Caption         =   "厚度"
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
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Left            =   3030
         Top             =   285
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "长度"
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
      Begin InDate.ULabel ULabel33 
         Height          =   315
         Left            =   150
         Top             =   2430
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "判定结果"
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
      Begin CSTextLibCtl.sidbEdit SDB_WGT_ORD 
         Height          =   315
         Left            =   4095
         TabIndex        =   10
         Top             =   1335
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_WGT 
         Height          =   315
         Left            =   4095
         TabIndex        =   11
         Top             =   645
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
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
         RawData         =   "0.000"
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_WID_MX 
         Height          =   315
         Left            =   2010
         TabIndex        =   5
         Top             =   1710
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   14737632
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LEN_MX 
         Height          =   315
         Left            =   3030
         TabIndex        =   8
         Top             =   1710
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Left            =   2010
         TabIndex        =   6
         Top             =   2055
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Left            =   1170
         TabIndex        =   7
         Top             =   2055
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Left            =   3030
         TabIndex        =   9
         Top             =   2055
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
      Begin CSTextLibCtl.sidbEdit SDB_PWGT_MN 
         Height          =   315
         Left            =   4095
         TabIndex        =   12
         Top             =   2055
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
         Left            =   2010
         TabIndex        =   3
         Top             =   645
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
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
         Left            =   1170
         TabIndex        =   4
         Top             =   645
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
      Begin CSTextLibCtl.sidbEdit SDB_LEN 
         Height          =   315
         Left            =   3030
         TabIndex        =   28
         Top             =   645
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
         Left            =   150
         Top             =   2055
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "下公差"
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
      Begin InDate.ULabel ULabel43 
         Height          =   315
         Left            =   150
         Top             =   645
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "实际"
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_THK_MX 
         Height          =   315
         Left            =   1170
         TabIndex        =   48
         Top             =   1710
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
      Begin CSTextLibCtl.sidbEdit SDB_PWGT_MX 
         Height          =   315
         Left            =   4095
         TabIndex        =   49
         Top             =   1710
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   150
         Top             =   1725
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "上公差"
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
      Begin InDate.ULabel ULabel44 
         Height          =   315
         Left            =   4095
         Top             =   285
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "重量"
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
      Begin CSTextLibCtl.sidbEdit SDB_ORD_WID 
         Height          =   315
         Left            =   2010
         TabIndex        =   50
         Top             =   1335
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
      Begin CSTextLibCtl.sidbEdit SDB_ORD_THK 
         Height          =   315
         Left            =   1170
         TabIndex        =   51
         Top             =   1335
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
      Begin CSTextLibCtl.sidbEdit SDB_ORD_LEN 
         Height          =   315
         Left            =   3030
         TabIndex        =   52
         Top             =   1335
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
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
      Begin InDate.ULabel ULabel45 
         Height          =   315
         Left            =   150
         Top             =   1335
         Width           =   990
         _ExtentX        =   1746
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
      Begin CSTextLibCtl.sidbEdit SDB_WID_R 
         Height          =   315
         Left            =   2010
         TabIndex        =   127
         Top             =   975
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
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
      Begin CSTextLibCtl.sidbEdit SDB_THK_R 
         Height          =   315
         Left            =   1170
         TabIndex        =   128
         Top             =   975
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
      Begin CSTextLibCtl.sidbEdit SDB_LEN_R 
         Height          =   315
         Left            =   3030
         TabIndex        =   129
         Top             =   975
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   150
         Top             =   990
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "实测"
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
         Index           =   5
         Left            =   150
         Top             =   3180
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "不平度(/m)"
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
         Index           =   6
         Left            =   150
         Top             =   3540
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "镰刀弯"
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
         Index           =   7
         Left            =   3060
         Top             =   3540
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "切斜"
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
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   150
         Top             =   2760
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "对角线1"
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_DIAGONAL1 
         Height          =   315
         Left            =   1170
         TabIndex        =   143
         Top             =   2760
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
         Left            =   2700
         Top             =   2760
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "对角线2"
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_DIAGONAL2 
         Height          =   315
         Left            =   3720
         TabIndex        =   144
         Top             =   2760
         Width           =   1410
         _Version        =   262145
         _ExtentX        =   2487
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
      Begin InDate.ULabel ULabel31 
         Height          =   315
         Left            =   150
         Top             =   3960
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "定尺"
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   8
         Left            =   2760
         Top             =   3180
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "不平度(/2m)"
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
   Begin Threed.SSFrame sf1 
      Height          =   5025
      Left            =   930
      TabIndex        =   14
      Top             =   9540
      Visible         =   0   'False
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   8864
      _Version        =   196609
      Font3D          =   2
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
      Caption         =   " 表面缺陷"
      Begin VB.TextBox TXT_CL 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   112
         Top             =   1020
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "尾部"
         Height          =   240
         Index           =   8
         Left            =   3360
         TabIndex        =   107
         Tag             =   "B"
         Top             =   1815
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中部"
         Height          =   195
         Index           =   7
         Left            =   3360
         TabIndex        =   106
         Tag             =   "M"
         Top             =   1590
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头部"
         Height          =   195
         Index           =   6
         Left            =   3360
         TabIndex        =   105
         Tag             =   "T"
         Top             =   1365
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   270
         TabIndex        =   104
         Text            =   " "
         Top             =   1530
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Height          =   315
         Index           =   5
         Left            =   705
         TabIndex        =   57
         Top             =   555
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Height          =   315
         Index           =   4
         Left            =   390
         TabIndex        =   56
         Top             =   555
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "尾部"
         Height          =   240
         Index           =   17
         Left            =   3375
         TabIndex        =   44
         Tag             =   "B"
         Top             =   3825
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中部"
         Height          =   195
         Index           =   16
         Left            =   3375
         TabIndex        =   43
         Tag             =   "M"
         Top             =   3600
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头部"
         Height          =   195
         Index           =   15
         Left            =   3375
         TabIndex        =   42
         Tag             =   "T"
         Top             =   3375
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "尾部"
         Height          =   240
         Index           =   14
         Left            =   2400
         TabIndex        =   41
         Tag             =   "B"
         Top             =   3825
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中部"
         Height          =   195
         Index           =   13
         Left            =   2400
         TabIndex        =   40
         Tag             =   "M"
         Top             =   3600
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头部"
         Height          =   195
         Index           =   12
         Left            =   2400
         TabIndex        =   39
         Tag             =   "T"
         Top             =   3375
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "尾部"
         Height          =   240
         Index           =   11
         Left            =   1440
         TabIndex        =   38
         Tag             =   "B"
         Top             =   3825
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中部"
         Height          =   195
         Index           =   10
         Left            =   1440
         TabIndex        =   37
         Tag             =   "M"
         Top             =   3600
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头部"
         Height          =   195
         Index           =   9
         Left            =   1440
         TabIndex        =   36
         Tag             =   "T"
         Top             =   3375
         Width           =   810
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   1410
         TabIndex        =   35
         Text            =   " "
         Top             =   3030
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   2370
         TabIndex        =   34
         Text            =   " "
         Top             =   3030
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   3345
         TabIndex        =   33
         Text            =   " "
         Top             =   3030
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   4
         Left            =   2370
         TabIndex        =   32
         Top             =   2700
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   5
         Left            =   3345
         TabIndex        =   31
         Top             =   2700
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   1
         Left            =   2370
         TabIndex        =   0
         Top             =   690
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   2385
         TabIndex        =   23
         Text            =   " "
         Top             =   1020
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   1410
         TabIndex        =   22
         Text            =   " "
         Top             =   1020
         Width           =   960
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头部"
         Height          =   195
         Index           =   0
         Left            =   1440
         TabIndex        =   21
         Tag             =   "T"
         Top             =   1365
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中部"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   20
         Tag             =   "M"
         Top             =   1590
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "尾部"
         Height          =   240
         Index           =   2
         Left            =   1440
         TabIndex        =   19
         Tag             =   "B"
         Top             =   1815
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头部"
         Height          =   195
         Index           =   3
         Left            =   2400
         TabIndex        =   18
         Tag             =   "T"
         Top             =   1365
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中部"
         Height          =   195
         Index           =   4
         Left            =   2400
         TabIndex        =   17
         Tag             =   "M"
         Top             =   1590
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "尾部"
         Height          =   240
         Index           =   5
         Left            =   2400
         TabIndex        =   16
         Tag             =   "B"
         Top             =   1815
         Width           =   810
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   0
         Left            =   1410
         TabIndex        =   1
         Top             =   2070
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   1
         Left            =   2370
         TabIndex        =   2
         Top             =   2070
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   225
         Top             =   1020
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "缺陷部位"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Left            =   225
         Top             =   2070
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "缺陷尺寸"
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   3
         Left            =   1410
         TabIndex        =   45
         Top             =   4095
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   4
         Left            =   2370
         TabIndex        =   46
         Top             =   4095
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   5
         Left            =   3345
         TabIndex        =   47
         Top             =   4095
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   225
         Top             =   3030
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "缺陷部位"
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
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   225
         Top             =   4080
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "缺陷尺寸"
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
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   1410
         Top             =   360
         Width           =   960
         _ExtentX        =   1693
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   2370
         Top             =   360
         Width           =   960
         _ExtentX        =   1693
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   2
         Left            =   3330
         TabIndex        =   108
         Top             =   2070
         Visible         =   0   'False
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
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
         MaxValue        =   9999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
   End
   Begin Threed.SSFrame Single 
      Height          =   1200
      Left            =   60
      TabIndex        =   13
      Top             =   75
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   2117
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox TXT_STLGRD 
         Height          =   285
         Left            =   2940
         TabIndex        =   94
         Top             =   345
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox TXT_APLY_ENDUSE_CD 
         Height          =   285
         Left            =   2730
         TabIndex        =   93
         Top             =   330
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox TXT_PROC_FLAG 
         Height          =   270
         Left            =   2520
         TabIndex        =   92
         Top             =   330
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox TXT_UST_FLAG 
         Height          =   270
         Left            =   2310
         TabIndex        =   91
         Top             =   330
         Visible         =   0   'False
         Width           =   210
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1050
         Left            =   0
         TabIndex        =   75
         Top             =   0
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   1852
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.TextBox txt_PrcLine 
            Height          =   285
            Left            =   2310
            TabIndex        =   79
            Text            =   " "
            Top             =   30
            Visible         =   0   'False
            Width           =   225
         End
         Begin Threed.SSOption opt_LineFlag 
            Height          =   255
            Index           =   1
            Left            =   330
            TabIndex        =   76
            Top             =   615
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   0
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "# 2"
         End
         Begin Threed.SSOption opt_LineFlag 
            Height          =   255
            Index           =   0
            Left            =   330
            TabIndex        =   77
            Top             =   180
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   255
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "# 1"
         End
         Begin Threed.SSOption opt_LineFlag 
            Height          =   255
            Index           =   2
            Left            =   1530
            TabIndex        =   78
            Top             =   180
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   0
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "# 3"
         End
         Begin Threed.SSOption opt_LineFlag 
            Height          =   255
            Index           =   3
            Left            =   1530
            TabIndex        =   142
            Top             =   615
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   0
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "# 4"
         End
      End
      Begin VB.TextBox txt_stdspec_chg_ref 
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
         MaxLength       =   18
         TabIndex        =   71
         Tag             =   "标准号"
         Top             =   600
         Width           =   2925
      End
      Begin VB.ComboBox CBO_SHIFT 
         Height          =   315
         ItemData        =   "CGD2050C.frx":003C
         Left            =   8685
         List            =   "CGD2050C.frx":0049
         TabIndex        =   67
         Top             =   600
         Width           =   1005
      End
      Begin VB.TextBox TXT_PLATE_NO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4350
         MaxLength       =   14
         TabIndex        =   54
         Top             =   150
         Width           =   2010
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   3135
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "钢板号"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   6450
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "生产时间"
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
      Begin CSTextLibCtl.sitxEdit SDT_PROD_DATE 
         Height          =   315
         Left            =   7665
         TabIndex        =   66
         Top             =   180
         Width           =   1200
         _Version        =   262145
         _ExtentX        =   2117
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
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   7470
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "班次"
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
         Height          =   300
         Index           =   4
         Left            =   3135
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         Caption         =   "标准号"
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
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   12510
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "厚度"
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
      Begin CSTextLibCtl.sidbEdit SDB_THK_REF 
         Height          =   315
         Left            =   13545
         TabIndex        =   70
         Top             =   180
         Width           =   1065
         _Version        =   262145
         _ExtentX        =   1879
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
      Begin InDate.ULabel ULabel24 
         Height          =   315
         Left            =   12510
         Top             =   600
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "宽度"
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
      Begin CSTextLibCtl.sidbEdit SDB_WID_REF 
         Height          =   315
         Left            =   13545
         TabIndex        =   72
         Top             =   600
         Width           =   1065
         _Version        =   262145
         _ExtentX        =   1879
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
      Begin CSTextLibCtl.sitxEdit SDT_PROD_TO_DATE 
         Height          =   315
         Left            =   8880
         TabIndex        =   73
         Top             =   180
         Width           =   1260
         _Version        =   262145
         _ExtentX        =   2222
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
      Begin Threed.SSPanel SSP4 
         Height          =   315
         Left            =   10680
         TabIndex        =   147
         Top             =   720
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   33023
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
      Begin Threed.SSPanel SSP6 
         Height          =   315
         Left            =   10680
         TabIndex        =   151
         Top             =   390
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   8454143
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
      Begin Threed.SSPanel SSP5 
         Height          =   315
         Left            =   10680
         TabIndex        =   152
         Top             =   60
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   8454143
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
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   120
         Left            =   10080
         TabIndex        =   74
         Top             =   300
         Width           =   195
      End
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   2520
      Left            =   120
      TabIndex        =   65
      Top             =   1290
      Width           =   15165
      _Version        =   393216
      _ExtentX        =   26749
      _ExtentY        =   4445
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   30
      MaxRows         =   5
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "CGD2050C.frx":0056
   End
   Begin InDate.ULabel ULabel1 
      Height          =   330
      Left            =   10410
      Top             =   9690
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      Caption         =   "后道工序"
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
   Begin Threed.SSFrame SSFrame5 
      Height          =   855
      Left            =   11610
      TabIndex        =   95
      Top             =   9690
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1508
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.CheckBox chkCl 
         BackColor       =   &H00E0E0E0&
         Caption         =   "冷矫直"
         Height          =   210
         Left            =   1560
         TabIndex        =   101
         Tag             =   "G"
         Top             =   90
         Width           =   900
      End
      Begin VB.CheckBox chkGrid 
         BackColor       =   &H00E0E0E0&
         Caption         =   "修磨"
         Height          =   210
         Left            =   750
         TabIndex        =   100
         Tag             =   "G"
         Top             =   90
         Width           =   720
      End
      Begin VB.CheckBox chkGas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "GAS"
         Height          =   210
         Left            =   60
         TabIndex        =   99
         Tag             =   "C"
         Top             =   90
         Width           =   645
      End
      Begin VB.TextBox txtGas 
         Height          =   285
         Left            =   450
         TabIndex        =   98
         Top             =   360
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox txtGrid 
         Height          =   285
         Left            =   1050
         TabIndex        =   97
         Top             =   390
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox txtCl 
         Height          =   285
         Left            =   1770
         TabIndex        =   96
         Top             =   390
         Visible         =   0   'False
         Width           =   210
      End
   End
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Index           =   1
      Left            =   15330
      Top             =   600
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "下线原因"
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
   Begin Threed.SSCommand cmd_Off 
      Height          =   375
      Left            =   16530
      TabIndex        =   111
      Top             =   150
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      _Version        =   196609
      Caption         =   "下线"
   End
   Begin Threed.SSFrame SSFrame6 
      Height          =   1365
      Left            =   5190
      TabIndex        =   134
      Top             =   7200
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   2408
      _Version        =   196609
      Font3D          =   2
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
      Begin VB.TextBox TXT_EMP_CD5 
         Height          =   330
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   162
         Top             =   960
         Width           =   2160
      End
      Begin VB.TextBox TXT_INSP_MAN_TAIL 
         Height          =   330
         Left            =   3750
         MaxLength       =   7
         TabIndex        =   146
         Tag             =   "检查员"
         Top             =   120
         Width           =   1050
      End
      Begin VB.TextBox TXT_INSP_MAN 
         Height          =   330
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   135
         Tag             =   "检查员"
         Top             =   120
         Width           =   1050
      End
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   210
         Top             =   525
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "检查时间"
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
      Begin CSTextLibCtl.sitxEdit TXT_INSP_OCCR_TIME 
         Height          =   315
         Left            =   1440
         TabIndex        =   136
         Tag             =   "检查时间"
         Top             =   525
         Width           =   2160
         _Version        =   262145
         _ExtentX        =   3810
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
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   210
         Top             =   135
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "头部检验工"
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
      Begin InDate.ULabel ULabel27 
         Height          =   315
         Left            =   2520
         Top             =   120
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "尾部检验工"
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
      Begin InDate.ULabel ULabel47 
         Height          =   315
         Left            =   210
         Top             =   960
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "缺陷责任人"
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
   Begin VB.Frame sf5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "缺陷"
      Height          =   1125
      Left            =   5280
      TabIndex        =   137
      Top             =   6060
      Width           =   6195
      Begin VB.TextBox TXT_INSP_FLAW 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   3510
         TabIndex        =   141
         Top             =   300
         Width           =   765
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   3510
         TabIndex        =   140
         Top             =   630
         Width           =   765
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   0
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   139
         Top             =   630
         Width           =   2250
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   3
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   138
         Top             =   300
         Width           =   2250
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   120
         Top             =   300
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
         Left            =   120
         Top             =   630
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
   Begin Threed.SSFrame SSFrame2 
      Height          =   2235
      Left            =   5190
      TabIndex        =   113
      Top             =   3810
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   3942
      _Version        =   196609
      Font3D          =   2
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
      Caption         =   " 修磨"
      Begin VB.CheckBox CHK_BOT_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "不合格"
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   3795
         TabIndex        =   121
         Tag             =   "N"
         Top             =   1260
         Width           =   900
      End
      Begin VB.CheckBox CHK_TOP_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "合格"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   3795
         TabIndex        =   120
         Tag             =   "Y"
         Top             =   420
         Width           =   735
      End
      Begin VB.CheckBox CHK_TOP_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "不合格"
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   3795
         TabIndex        =   119
         Tag             =   "N"
         Top             =   675
         Width           =   900
      End
      Begin VB.TextBox TXT_GRID_EMP_CD 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   118
         Tag             =   "作业人员"
         Top             =   1360
         Width           =   1035
      End
      Begin VB.CheckBox CHK_GRID_FLAG 
         BackColor       =   &H00E0E0E0&
         Caption         =   "是否修磨"
         Height          =   240
         Left            =   195
         TabIndex        =   117
         Tag             =   "G"
         Top             =   300
         Width           =   1110
      End
      Begin VB.TextBox TXT_TOP_GRID_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   116
         Text            =   " "
         Top             =   600
         Width           =   690
      End
      Begin VB.TextBox TXT_BOT_GRID_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   115
         Text            =   " "
         Top             =   980
         Width           =   690
      End
      Begin VB.CheckBox CHK_BOT_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "合格"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   3795
         TabIndex        =   114
         Tag             =   "Y"
         Top             =   990
         Width           =   735
      End
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   195
         Top             =   1380
         Width           =   1095
         _ExtentX        =   1931
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
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Index           =   0
         Left            =   195
         Top             =   990
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
            Size            =   9.75
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
         Left            =   195
         Top             =   600
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
            Size            =   9.75
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
         Left            =   1320
         Top             =   240
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_TOP_GRID_DEEP 
         Height          =   315
         Left            =   2910
         TabIndex        =   122
         Top             =   600
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
         Left            =   2040
         TabIndex        =   123
         Top             =   600
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
         Left            =   2040
         TabIndex        =   124
         Top             =   975
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   195
         Top             =   1755
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sitxEdit TXT_GRID_TIME 
         Height          =   315
         Left            =   1320
         TabIndex        =   125
         Top             =   1740
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
         Left            =   2910
         TabIndex        =   126
         Top             =   975
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
   End
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   7380
      Top             =   9525
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      Caption         =   "检查人员"
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
Attribute VB_Name = "CGD2050C"
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
'-- Program Name      表面检查实绩查询及修改界面
'-- Program ID        AGC2020C
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
Public sQuery_Rt As String          'Active Form sQuery Setting

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

Dim sControl  As New Collection      'Master Clear Key Collection
Dim MC        As New Collection      'Master Collection
Dim Mc1       As New Collection      'Master Collection

Dim sc1       As New Collection      'Spread Collection
Dim Proc_Sc   As New Collection      'Spread Struc Collection

Dim sCheck  As String
Dim sQuery  As String

Const SS1_PLATE_NO = 1
Const SS1_IMP_CONT = 24
Const SS1_FLAG = 25
Const SS1_EXPORT = 26

Private Sub Form_Define()
    Dim iIndex As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_PrcLine, "p", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDT_PROD_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDT_PROD_TO_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_stdspec_chg_ref, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_THK_REF, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_WID_REF, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_UST_FLAG, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_PROC_FLAG, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(TXT_APLY_ENDUSE_CD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_stlgrd, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                                                                                                                                                
     Call Gp_Ms_Collection(TXT_INSP_FLAW(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_FLAW(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_FLAW(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_PART(3), " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_PART(4), " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_PART(5), " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LTH(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LTH(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LTH(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_FLAW(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_FLAW(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_FLAW(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_PART(0), " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_PART(1), " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_PART(2), " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LTH(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LTH(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LTH(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_THK, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_THK_MX, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_THK_MN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_WID, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_WID_MX, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_WID_MN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_LEN, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LEN_MX, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LEN_MN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_WGT_ORD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_WGT, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_PWGT_MX, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_PWGT_MN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
     Call Gp_Ms_Collection(TXT_INSP_THK_GRD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_WID_GRD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_LEN_GRD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_WGT_GRD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_SURF_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_INSP_MAIN_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
        'Call Gp_Ms_Collection(TXT_NEXT_PROC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
               Call Gp_Ms_Collection(txtGas, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txtGrid, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(TXT_CL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(TXT_INSP_MAN, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) '检验工（原来）现为头部检验工，两者一样
    Call Gp_Ms_Collection(TXT_INSP_MAN_TAIL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) 'ADD BY LICHAO 尾部检验工
   Call Gp_Ms_Collection(TXT_INSP_OCCR_TIME, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_ORD_WID, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_ORD_THK, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_ORD_LEN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_GRID_EMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_GRID_TIME, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_TOP_GRID_GRD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_TOP_GRID_YRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_TOP_GRID_DEEP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_BOT_GRID_GRD, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_BOT_GRID_YRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_BOT_GRID_DEEP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_STDSPEC, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_stdspec_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_stdspec_chg, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_stdspec_name_chg, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Scrap_code, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Scrap_name, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(SDB_Mn, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_PROC_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDB_THK_R, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDB_WID_R, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDB_LEN_R, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            'add by liqian at 20120322
          Call Gp_Ms_Collection(TXT_EMP_CD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_WAVE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_VERT_DEG, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_RECT_DEG, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
   Call Gp_Ms_Collection(SDB_INSP_DIAGONAL1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_INSP_DIAGONAL2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_SIZE_KND, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(TXT_WAVE1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Color_code, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
              Call Gp_Ms_Collection(SDB_HD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_HD2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_HD3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_HD4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_HD5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_HD6, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              
          Call Gp_Ms_Collection(CHK_FLAW_YN, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
          Call Gp_Ms_Collection(TXT_EMP_CD5, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(COM_PF, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
    For iIndex = 0 To 17
        Call Gp_Clear_Collection(CHK_PART(iIndex), "s", sControl)
    Next iIndex
    
     Call Gp_Clear_Collection(CHK_TOP_GRD(0), "s", sControl)
     Call Gp_Clear_Collection(CHK_TOP_GRD(1), "s", sControl)
     Call Gp_Clear_Collection(CHK_BOT_GRD(0), "s", sControl)
     Call Gp_Clear_Collection(CHK_BOT_GRD(1), "s", sControl)
     Call Gp_Clear_Collection(CHK_BOT_GRD(1), "s", sControl)
     
    
    MC.Add Item:=sControl, Key:="sControl"
    
    'MASTER Collection
    Mc1.Add Item:="CGD2050C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="CGD2050C.P_REFER", Key:="P-R"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
      
    'Spread_Collection
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGD2050C.P_SREFER", Key:="P-R"
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

Private Sub cbo_ResonDesc_Click()
    txt_ResonCd = Mid(cbo_ResonDesc.Text, 1, 1)
End Sub

Private Sub CHK_CL_FL_Click()
    If CHK_CL_FL.Value = ssCBChecked Then
       TXT_CL.Text = "Y"
    Else
       TXT_CL.Text = "N"
    End If
End Sub

Private Sub chkCl_Click()
    If chkCl.Value Then
        txtCl = "Y"
        chkCl.ForeColor = &HFF&       'red
    Else
        txtCl = "N"
        chkCl.ForeColor = &H80000012       'red
    End If
End Sub

Private Sub chkGas_Click()
    If chkGas.Value Then
        txtGas = "Y"
        chkGas.ForeColor = &HFF&       'red
    Else
        txtGas = "N"
        chkGas.ForeColor = &H80000012       'red
    End If
End Sub

Private Sub chkGrid_Click()
    If chkGrid.Value Then
        txtGrid = "Y"
        chkGrid.ForeColor = &HFF&       'red
    Else
        txtGrid = "N"
        chkGrid.ForeColor = &H80000012       'red
    End If
End Sub

Private Sub cmd_Off_Click()
    Dim OutParam(2, 4) As Variant
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
    
    On Error Resume Next

    Screen.MousePointer = vbHourglass

        
    'Return loaction1 Parameter
    OutParam(1, 1) = "arg_loaction1"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 10

    'Return loaction2 Parameter
    OutParam(2, 1) = "arg_loaction2"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 10
    
    sQuery = "{call CGD2050C.P_LINEOFF('" & Trim(TXT_PLATE_NO.Text) & "','" & txt_PrcLine & "','" & txt_ResonCd & "','" & Gf_ShiftSet3(M_CN1) & "','" & sUserID & "',?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If Trim(adoCmd("arg_loaction2")) <> "" Then
        Call Gp_MsgBoxDisplay("实绩处理失败，请确认=> " & adoCmd("arg_loaction2"))
    End If
    
    Set adoCmd = Nothing
    
    Call Form_Ref
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        If Len(TXT_PLATE_NO.Text) >= 8 Then
           Call Form_Ref
        End If
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
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))

    Call Gf_Sp_Cls(sc1)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    
   
    If TXT_PLATE_NO <> "" Then
       Call Form_Ref
    End If
    
    opt_LineFlag(0).Value = True
    
    cbo_ResonDesc.AddItem "1:设备异常"
    cbo_ResonDesc.AddItem "2:线过负荷"
    cbo_ResonDesc.AddItem "3:产品异常"
    
    Screen.MousePointer = vbDefault
    
    If Mid(sAuthority, 1, 3) = "111" Then
       cmd_Off.Enabled = True
    Else
       cmd_Off.Enabled = False
    End If
    
    CHK_CL_FL.Value = 0
    
    COM_PF.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "CG-System.INI", Me.Name)

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
    
    Set sControl = Nothing
    Set MC = Nothing

    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()

    Dim iCount As Integer
    
    If Gf_Sp_Cls(sc1) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        TXT_PLATE_NO = ""
        Call Gp_SSCheck_Cls(MC("sControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)

       ' TXT_INSP_MAN = sUserID
       'add by liqian at 20120322
        TXT_INSP_MAN = ""
        TXT_EMP_CD1 = sUserID
        
        For iCount = 0 To 5
            TXT_INSP_FLAW_NAME(iCount).Text = ""
        Next iCount
        
        ss1.BlockMode = True
        ss1.ROW = -1
        ss1.Col = -1
        ss1.BackColor = &HFFFFFF
        ss1.BlockMode = False
        CHK_CL_FL.Value = 0
        
        CHK_FLAW_YN.Value = 0
        
    End If
End Sub

Public Sub Form_Ref()
Dim i As Integer
'
'    Call Form_Cls
Dim simpcont As String
Dim iCount   As Integer
Dim sFlag As String
Dim sexport As String

    Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1)
    
    If ss1.MaxRows > 0 Then
       ss1.ROW = 1
       ss1.Col = 1
       TXT_PLATE_NO.Text = ss1.Text
       For i = 0 To 17
           CHK_PART(i).Value = 0
       Next
    End If
    
    If Len(TXT_PLATE_NO.Text) = 14 Then
        If Gf_Ms_Refer(M_CN1, Mc1, , , False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            
            If TXT_SURF_GRD = "Y" Then
               opt_CHK_SUR_GRD(0).Value = True
            Else
               opt_CHK_SUR_GRD(1).Value = True
            End If
            
            
            If Len(TXT_INSP_MAIN_GRD) = 1 Then
                If TXT_INSP_MAIN_GRD = "7" Then
                   opt_CHK_PRD_GRD(5).Value = True
                Else
                   opt_CHK_PRD_GRD(TXT_INSP_MAIN_GRD - 1).Value = True
                End If
            End If
            If TXT_INSP_OCCR_TIME.RawData = "" Then
               TXT_INSP_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
            End If
            'TXT_INSP_MAN = sUserID
            'ADD BY LIQIAN AT 20120322
            TXT_EMP_CD1 = sUserID
            
            'Call Display_Data_Edit
        End If
    End If
    
    With ss1
        For iCount = 1 To .MaxRows
        
            .ROW = iCount:
            .Col = SS1_IMP_CONT:   simpcont = Trim(.Text)
            .Col = SS1_FLAG:       sFlag = Trim(.Text)
            .Col = SS1_EXPORT:     sexport = Trim(.Text)
            If simpcont = "Y" Then
                Call Gp_Sp_BlockColor(ss1, SS1_PLATE_NO, SS1_PLATE_NO, iCount, iCount, SSP4.BackColor)
                Call Gp_Sp_BlockColor(ss1, SS1_IMP_CONT, SS1_IMP_CONT, iCount, iCount, SSP4.BackColor)
            End If
            
            '是否定制配送
                  If sFlag = "Y" Then
                     Call Gp_Sp_BlockColor(ss1, SS1_PLATE_NO, SS1_PLATE_NO, iCount, iCount, SSP5.BackColor)
                  End If
                  '是否出口订单
                  
                  If sexport = "Y" Then
                     Call Gp_Sp_BlockColor(ss1, SS1_PLATE_NO, SS1_PLATE_NO, iCount, iCount, SSP6.BackColor)
                  End If
        Next iCount
    End With
    
    CHK_CL_FL.Value = 0
     
End Sub

Public Sub Form_Pro()

    Dim SMESG   As String
    Dim iCount  As Integer
        
    If Trim(TXT_INSP_MAIN_GRD.Text) <> "4" Then
        If Trim(TXT_SURF_GRD.Text) = "" Then
            SMESG = " 请输入表面判定 ！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
        End If
        
    End If
    
    If Not Gp_DateCheck(TXT_INSP_OCCR_TIME) Then
        SMESG = " 请正确输入检查时间 ！"
        Call Gp_MsgBoxDisplay(SMESG)
        Exit Sub
    End If
    
    If CHK_GRID_FLAG.Value = ssCBChecked Then
        If Not Gp_DateCheck(TXT_GRID_TIME) Then
            SMESG = " 请正确输入修磨时间 ！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
        End If
        If Trim(TXT_GRID_EMP_CD.Text) = "" Then
            TXT_GRID_EMP_CD.Text = sUserID
        End If
        If TXT_TOP_GRID_GRD.Text = "" Then
            SMESG = " 请正确输入上表面修磨后判定 ！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
        End If
        If TXT_BOT_GRID_GRD.Text = "" Then
            SMESG = " 请正确输入下表面修磨后判定 ！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
        End If
    End If
        
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        'TXT_INSP_MAN.Text = sUserID
        'ADD BY LIQIAN AT 20120322
        TXT_EMP_CD1.Text = sUserID
        If TXT_INSP_MAN.Text = "" Then
            SMESG = " 请选择检验人员！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
        End If
       If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    End If
    
    CHK_CL_FL.Value = 0

End Sub

Private Sub opt_CHK_PRD_GRD_Click(Index As Integer, Value As Integer)
    If Index = 0 Then
       TXT_INSP_MAIN_GRD = "1"
       opt_CHK_PRD_GRD(0).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
       COM_PF.Enabled = False
       COM_PF.Text = ""
    ElseIf Index = 1 Then
       TXT_INSP_MAIN_GRD = "2"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
       COM_PF.Enabled = False
       COM_PF.Text = ""
    ElseIf Index = 2 Then
        TXT_INSP_MAIN_GRD = "3"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
       COM_PF.Enabled = False
       COM_PF.Text = ""
    ElseIf Index = 3 Then
        TXT_INSP_MAIN_GRD = "4"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
       COM_PF.Enabled = False
       COM_PF.Text = ""
    ElseIf Index = 4 Then
        TXT_INSP_MAIN_GRD = "5"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
       COM_PF.Enabled = False
       COM_PF.Text = ""
    ElseIf Index = 5 Then
        TXT_INSP_MAIN_GRD = "7"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &HFF&       'red
       txt_Scrap_code.Enabled = True
       COM_PF.Enabled = True
    End If
End Sub

Private Sub opt_CHK_SUR_GRD_Click(Index As Integer, Value As Integer)
    If Index = 0 Then
       opt_CHK_SUR_GRD(0).ForeColor = &HFF&       'red
       opt_CHK_SUR_GRD(1).ForeColor = &H80000012  'black
        TXT_SURF_GRD = "Y"
    Else
        TXT_SURF_GRD = "N"
       opt_CHK_SUR_GRD(1).ForeColor = &HFF&       'red
       opt_CHK_SUR_GRD(0).ForeColor = &H80000012  'black
    End If
End Sub

Private Sub opt_LineFlag_Click(Index As Integer, Value As Integer)
'    Call Form_Cls
'    TXT_PLATE_NO = ""
    If opt_LineFlag(0).Value = True Then
       txt_PrcLine = "1"
       opt_LineFlag(0).ForeColor = &HFF&       'red
       opt_LineFlag(1).ForeColor = &H80000012  'black
       opt_LineFlag(2).ForeColor = &H80000012  'black
       opt_LineFlag(3).ForeColor = &H80000012  'black
    ElseIf opt_LineFlag(1).Value = True Then
       txt_PrcLine = "2"
       opt_LineFlag(0).ForeColor = &H80000012  'black
       opt_LineFlag(1).ForeColor = &HFF&       'red
       opt_LineFlag(2).ForeColor = &H80000012  'black
       opt_LineFlag(3).ForeColor = &H80000012  'black
    ElseIf opt_LineFlag(2).Value = True Then
       txt_PrcLine = "3"
       opt_LineFlag(0).ForeColor = &H80000012  'black
       opt_LineFlag(1).ForeColor = &H80000012  'black
       opt_LineFlag(2).ForeColor = &HFF&       'red
       opt_LineFlag(3).ForeColor = &H80000012  'black
    ElseIf opt_LineFlag(3).Value = True Then
       txt_PrcLine = "4"
       opt_LineFlag(0).ForeColor = &H80000012  'black
       opt_LineFlag(1).ForeColor = &H80000012  'black
       opt_LineFlag(2).ForeColor = &H80000012  'black
       opt_LineFlag(3).ForeColor = &HFF&       'red
    End If
End Sub

''add by liqian at 2012-03-14 根据实测长度值计算公称长度
'Private Sub SDB_LEN_R_Change()
' Dim iLen As Integer
'     If TXT_SIZE_KND <> "01" Then
'        If SDB_LEN_R.Value > 0 Then
'           iLen = Int(SDB_LEN_R.Value / 50) * 50
'           SDB_LEN.Value = iLen
'        End If
'     End If
'End Sub

Private Sub SDB_THK_Change()
    Call PRD_WEIGHT_CALC
End Sub
    
Private Sub SDB_WID_Change()
    Call PRD_WEIGHT_CALC
End Sub

Private Sub SDB_LEN_Change()
    Call PRD_WEIGHT_CALC
End Sub

Private Sub PRD_WEIGHT_CALC()

    Dim dThk        As Double
    Dim dWid        As Double
    Dim dLen        As Double
    
    dThk = Val(Format(SDB_THK.Text, "####0.##") & "")
    dWid = Val(Format(SDB_WID.Text, "###0") & "")
    dLen = Val(Format(SDB_LEN.Text, "###0.##") & "")
    If dThk > 0 And dWid > 0 And dLen > 0 Then
        SDB_WGT.Text = Cal_Plate_Wgt("WGT", dThk, dWid, dLen)
    End If
    
    Call Size_Grade_Edit
End Sub

Private Function Cal_Plate_Wgt(sMode As String, dThk As Double, dWid As Double, dLen As Double) As Double

    Dim RS  As New ADODB.Recordset
    
    Cal_Plate_Wgt = 0
    
    sQuery = "SELECT  Gf_Cal_Plate_Wgt('" & sMode & "'" & vbCrLf
    sQuery = sQuery & "             ,'" & Trim(TXT_APLY_ENDUSE_CD.Text) & "'" & vbCrLf
    sQuery = sQuery & "             ,'" & Trim(txt_stlgrd.Text) & "'" & vbCrLf
    sQuery = sQuery & "             ," & dThk & vbCrLf
    sQuery = sQuery & "             ," & dWid & vbCrLf
    sQuery = sQuery & "             ," & dLen & vbCrLf
    sQuery = sQuery & "             ,0 )" & vbCrLf
    sQuery = sQuery & "       FROM  DUAL " & vbCrLf
    RS.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        Cal_Plate_Wgt = Val(RS(0).Value & "")
    End If
    
    RS.Close
    Set RS = Nothing
     
End Function

Private Sub SDT_PROD_DATE_DblClick()
     SDT_PROD_DATE.RawData = Gf_DTSet(M_CN1, "D")
     SDT_PROD_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
End Sub
Private Sub SDT_PROD_TO_DATE_DblClick()
     SDT_PROD_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
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

Private Sub TXT_GRID_EMP_CD_DblClick()
    TXT_GRID_EMP_CD.Text = sUserID
End Sub

Private Sub TXT_GRID_TIME_DblClick()
    TXT_GRID_TIME.RawData = Gf_DTSet(M_CN1, , "X")
End Sub

Private Sub TXT_INSP_FLAW_Change(Index As Integer)
    TXT_INSP_FLAW_NAME(Index).Text = Gf_ComnNameFind(M_CN1, "G0002", TXT_INSP_FLAW(Index).Text, 1)
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

'Private Sub TXT_INSP_MAN_DblClick()
'    TXT_INSP_MAN.Text = sUserID
'End Sub

Private Sub TXT_INSP_OCCR_TIME_DblClick()
    TXT_INSP_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
End Sub


Private Sub CHK_PART_Click(Index As Integer)
    Dim iCount      As Integer
    Dim iIndexTxt   As Integer
    Dim iIndexChk   As Integer
    Dim iIndexStr   As Integer
    
    If sCheck <> "" Then Exit Sub
    
    iIndexTxt = Index \ 3
    iIndexChk = iIndexTxt * 3
    iCount = 0
    sCheck = "**"
            
    If CHK_PART(Index).Value = ssCBUnchecked Then
        For iIndexStr = iIndexChk To iIndexChk + 2
            If CHK_PART(iIndexStr).Value = ssCBChecked Then
               iCount = iCount + 1
            End If
        Next iIndexStr
        If iCount = 0 Then
            TXT_INSP_PART(iIndexTxt).Text = ""
            TXT_INSP_FLAW(iIndexTxt).Text = ""
            TXT_INSP_FLAW_NAME(iIndexTxt).Text = ""
            CHK_PART(Index).ForeColor = &H808080
            sCheck = ""
            Exit Sub
        End If
    Else
        For iIndexStr = iIndexChk To iIndexChk + 2
            CHK_PART(iIndexStr).ForeColor = &H808080
            CHK_PART(iIndexStr).Value = ssCBUnchecked
        Next iIndexStr
    End If
    
    CHK_PART(Index).ForeColor = &HFF&
    CHK_PART(Index).Value = ssCBChecked

    TXT_INSP_PART(iIndexTxt).Text = CHK_PART(Index).Tag
    sCheck = ""
    
End Sub

Private Sub CHK_SUR_GRD_Click(Index As Integer)
    Dim iNext       As Integer
    
    If sCheck <> "" Then Exit Sub

    sCheck = "**"
    
    If Index = 0 Then
        iNext = 1
    Else
        iNext = 0
    End If

    sCheck = ""
    
End Sub

Private Sub CHK_PRD_GRD_Click(Index As Integer)
    Dim iCount      As Integer
    Dim iIndexStr   As Integer
    
    If sCheck <> "" Then Exit Sub

    iCount = 0
    sCheck = "**"
                 
    txt_stdspec_chg.Text = ""
    txt_stdspec_name_chg.Text = ""

        
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

Private Sub CHK_GRID_FLAG_Click()
    If CHK_GRID_FLAG.Value = ssCBUnchecked Then
        CHK_TOP_GRD(0).Enabled = False:        CHK_TOP_GRD(0).Value = ssCBUnchecked
        CHK_TOP_GRD(1).Enabled = False:        CHK_TOP_GRD(1).Value = ssCBUnchecked
        CHK_BOT_GRD(0).Enabled = False:        CHK_BOT_GRD(0).Value = ssCBUnchecked
        CHK_BOT_GRD(1).Enabled = False:        CHK_BOT_GRD(1).Value = ssCBUnchecked
        SDB_TOP_GRID_YRD.Enabled = False:      SDB_TOP_GRID_YRD.Text = ""
        SDB_BOT_GRID_YRD.Enabled = False:      SDB_BOT_GRID_YRD.Text = ""
        SDB_TOP_GRID_DEEP.Enabled = False:     SDB_TOP_GRID_DEEP.Text = ""
        SDB_BOT_GRID_DEEP.Enabled = False:     SDB_BOT_GRID_DEEP.Text = ""
        TXT_GRID_EMP_CD.Enabled = False:       TXT_GRID_EMP_CD.Text = ""
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
        TXT_GRID_EMP_CD.Enabled = True
        TXT_GRID_TIME.Enabled = True
        
        TXT_GRID_EMP_CD.Text = sUserID
        TXT_GRID_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        
        CHK_TOP_GRD(0).Value = ssCBChecked
        Call CHK_TOP_GRD_Click(0)
        CHK_BOT_GRD(0).Value = ssCBChecked
        Call CHK_BOT_GRD_Click(0)

    End If
End Sub

Private Sub Display_Data_Edit()
    Dim iIndexChk   As Integer
    Dim iIndexStr   As Integer
    
    sCheck = "**"
    
    For iIndexStr = 0 To 5
        For iIndexChk = iIndexStr * 3 To (iIndexStr * 3) + 2
            If TXT_INSP_PART(iIndexStr).Text = CHK_PART(iIndexChk).Tag Then
                CHK_PART(iIndexChk).ForeColor = &HFF&
                CHK_PART(iIndexChk).Value = ssCBChecked
            Else
                CHK_PART(iIndexChk).ForeColor = &H808080
                CHK_PART(iIndexChk).Value = ssCBUnchecked
            End If
        Next iIndexChk
    Next iIndexStr
        

    
    If Trim(TXT_TOP_GRID_GRD.Text) <> "" Then CHK_GRID_FLAG.Value = ssCBChecked
    
    If TXT_TOP_GRID_GRD.Text = "Y" Then
        CHK_TOP_GRD(0).Value = ssCBChecked
        CHK_TOP_GRD(1).Value = ssCBUnchecked
    ElseIf TXT_TOP_GRID_GRD.Text = "N" Then
        CHK_TOP_GRD(0).Value = ssCBUnchecked
        CHK_TOP_GRD(1).Value = ssCBChecked
    End If
    
    If TXT_BOT_GRID_GRD.Text = "Y" Then
        CHK_BOT_GRD(0).Value = ssCBChecked
        CHK_BOT_GRD(1).Value = ssCBUnchecked
    ElseIf TXT_BOT_GRID_GRD.Text = "N" Then
        CHK_BOT_GRD(0).Value = ssCBUnchecked
        CHK_BOT_GRD(1).Value = ssCBChecked
    End If
    
    If txtGas = "Y" Then
        chkGas.Value = 1
    End If
    If txtGrid = "Y" Then
        chkGrid.Value = 1
    End If
    If txtCl = "Y" Then
        chkCl.Value = 1
    End If

    If TXT_INSP_MAIN_GRD = "1" Then
        opt_CHK_PRD_GRD(0).Value = True
    ElseIf TXT_INSP_MAIN_GRD = "2" Then
        opt_CHK_PRD_GRD(1).Value = True
    ElseIf TXT_INSP_MAIN_GRD = "3" Then
        opt_CHK_PRD_GRD(2).Value = True
    ElseIf TXT_INSP_MAIN_GRD = "4" Then
        opt_CHK_PRD_GRD(3).Value = True
    ElseIf TXT_INSP_MAIN_GRD = "5" Then
        opt_CHK_PRD_GRD(4).Value = True
    ElseIf TXT_INSP_MAIN_GRD = "7" Then
        opt_CHK_PRD_GRD(5).Value = True
    End If
    
    If TXT_SURF_GRD = "Y" Then
        opt_CHK_SUR_GRD(0).Value = True
    ElseIf TXT_SURF_GRD = "N" Then
        opt_CHK_SUR_GRD(1).Value = True
    End If
    
    '''''''''ADD BY GUOLI AT 200712071330''''''''''
    If opt_CHK_SUR_GRD(0).Value = True Then
       TXT_SURF_GRD = "Y"
    ElseIf opt_CHK_SUR_GRD(1).Value = True Then
       TXT_SURF_GRD = "N"
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''

End Sub

Private Sub Size_Grade_Edit()
    Dim sGradeFlag As String
    
    sGradeFlag = ""
    
    If TXT_PROC_FLAG.Text <> "CGD" Then Exit Sub
    
    ' THICK GRAND CHECK
    If Val(SDB_THK & "") >= Val(SDB_ORD_THK & "") + Val(SDB_INSP_THK_MN & "") And _
       Val(SDB_THK & "") <= Val(SDB_ORD_THK & "") + Val(SDB_INSP_THK_MX & "") Then
        TXT_INSP_THK_GRD = "Y"
        SDB_THK.ForeColor = &H80000012
    Else
        TXT_INSP_THK_GRD = "N"
        SDB_THK.ForeColor = &HFF&
        sGradeFlag = "N"
    End If
    
    ' WIDTH GRAND CHECK
    If Val(SDB_WID & "") >= Val(SDB_ORD_WID & "") + Val(SDB_INSP_WID_MN & "") And _
       Val(SDB_WID & "") <= Val(SDB_ORD_WID & "") + Val(SDB_INSP_WID_MX & "") Then
        TXT_INSP_WID_GRD = "Y"
        SDB_WID.ForeColor = &H80000012
    Else
        TXT_INSP_WID_GRD = "N"
        SDB_WID.ForeColor = &HFF&
        sGradeFlag = "N"
    End If
        
    ' LENGTH GRAND CHECK
    If Val(SDB_LEN & "") >= Val(SDB_ORD_LEN & "") + Val(SDB_INSP_LEN_MN & "") And _
       Val(SDB_LEN & "") <= Val(SDB_ORD_LEN & "") + Val(SDB_INSP_LEN_MX & "") Then
        TXT_INSP_LEN_GRD = "Y"
        SDB_LEN.ForeColor = &H80000012
    Else
        TXT_INSP_LEN_GRD = "N"
        SDB_LEN.ForeColor = &HFF&
        sGradeFlag = "N"
    End If
    
    ' WEIGHT GRAND CHECK
    If Val(SDB_WGT & "") >= Val(SDB_WGT_ORD & "") + Val(SDB_PWGT_MN & "") And _
       Val(SDB_WGT & "") <= Val(SDB_WGT_ORD & "") + Val(SDB_PWGT_MX & "") Then
        TXT_INSP_WGT_GRD = "Y"
        SDB_WGT.ForeColor = &H80000012
    Else
        TXT_INSP_WGT_GRD = "N"
        SDB_WGT.ForeColor = &HFF&
        sGradeFlag = "N"
    End If
    


End Sub
  
Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       ' Call Gp_Ms_ControlLock(Mc1("pControl"), False)
     End If

End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)
    'If Row < 1 Or SDB_DIVIDE_CNT.Value > 0 Then Exit Sub
    
    ss1.ROW = ROW
    ss1.Col = 1
    TXT_PLATE_NO.Text = ss1.Text
    CHK_GRID_FLAG.Value = ssCBUnchecked
    
    If Len(TXT_PLATE_NO.Text) = 14 Then
        Call Gp_SSCheck_Cls(MC("sControl"))
        If Gf_Ms_Refer(M_CN1, Mc1, , , True) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            ''''''''''''''''''ADD BY GUOLI AT 200712071330''''''''''
            If opt_CHK_SUR_GRD(0).Value = True Then
               TXT_SURF_GRD = "Y"
            ElseIf opt_CHK_SUR_GRD(1).Value = True Then
               TXT_SURF_GRD = "N"
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If Len(TXT_INSP_MAIN_GRD) = 1 Then
                If TXT_INSP_MAIN_GRD = "7" Then
                   opt_CHK_PRD_GRD(5).Value = True
                Else
                   opt_CHK_PRD_GRD(TXT_INSP_MAIN_GRD - 1).Value = True
                End If
            End If

            'Call Display_Data_Edit
        End If
        If TXT_INSP_OCCR_TIME.RawData = "" Then
           TXT_INSP_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        End If
        
        ' add by liqian at 2012-04-12  下一块时改判标准到自动清空
        txt_stdspec_chg.Text = ""
        txt_stdspec_name_chg.Text = ""
        
        'TXT_INSP_MAN = sUserID
        'MODIFY BY LIQIAN AT 20120322
        TXT_EMP_CD1.Text = sUserID
        
    End If
    
End Sub

Private Sub TXT_INSP_PART_Change(Index As Integer)
Dim i As Integer
For i = 0 To 5
    If TXT_INSP_PART(i).Text = "T" Then
       CHK_PART(i * 3).Value = 1
    ElseIf TXT_INSP_PART(i).Text = "M" Then
       CHK_PART(i * 3 + 1).Value = 1
    ElseIf TXT_INSP_PART(i).Text = "B" Then
       CHK_PART(i * 3 + 2).Value = 1
    End If
Next
End Sub

Private Sub txt_Scrap_code_DblClick()
    Call txt_Scrap_code_KeyUp(vbKeyF4, 0)
End Sub




Private Sub TXT_SIZE_KND_Change()
    If Len(Trim(TXT_SIZE_KND.Text)) = TXT_SIZE_KND.MaxLength Then
        TXT_SIZE_KND_NAME.Text = Gf_ComnNameFind(M_CN1, "B0043", TXT_SIZE_KND.Text, 2)
    Else
        TXT_SIZE_KND_NAME.Text = ""
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
         DD.sWitch = "MS"
         DD.DataDicType = "C"
         DD.rControl.Add Item:=txt_stdspec_chg
         DD.rControl.Add Item:=txt_stdspec_name_chg
        
         Call Pf_Common_DD(M_CN1, vbKeyF4)
         
End Sub

Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        txt_stdspec_yy.Text = ""
        DD.rControl.Add Item:=txt_stdspec_chg
        DD.rControl.Add Item:=txt_stdspec_yy
        DD.rControl.Add Item:=txt_stdspec_name_chg

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

        Exit Sub

    End If
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


Private Sub txt_Scrap_code_Change()
    
    If Len(Trim(txt_Scrap_code)) = txt_Scrap_code.MaxLength Then
        txt_Scrap_name.Text = Gf_ComnNameFind(M_CN1, "G0017", Trim(txt_Scrap_code.Text), 1)
    Else
        txt_Scrap_name.Text = ""
    End If
    
End Sub

Private Sub txt_Scrap_code_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
            
        DD.sWitch = "MS"
        DD.sKey = "G0017"
        DD.rControl.Add Item:=txt_Scrap_code
        DD.rControl.Add Item:=txt_Scrap_name
        
        DD.nameType = "1"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    End If

End Sub

Private Sub txt_stdspec_chg_ref_DblClick()
    Call txt_stdspec_chg_ref_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stdspec_chg_ref_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec_chg_ref

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

Private Sub TXT_INSP_MAN_DblClick()
    Call TXT_INSP_MAN_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_INSP_MAN_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0054"

        DD.rControl.Add Item:=TXT_INSP_MAN

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub TXT_INSP_MAN_TAIL_DblClick()
    Call TXT_INSP_MAN_TAIL_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_INSP_MAN_TAIL_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0054"

        DD.rControl.Add Item:=TXT_INSP_MAN_TAIL

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub
