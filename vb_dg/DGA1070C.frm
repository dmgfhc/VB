VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form DGA1070C 
   Caption         =   "表面检查实绩查询及修改_DGA1070C"
   ClientHeight    =   9405
   ClientLeft      =   5925
   ClientTop       =   2085
   ClientWidth     =   15150
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   15150
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame2 
      Height          =   2325
      Left            =   4575
      TabIndex        =   65
      Top             =   6960
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   4101
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
      Begin VB.CheckBox CHK_GRID_FLAG 
         BackColor       =   &H00E0E0E0&
         Caption         =   "是否修磨"
         Height          =   240
         Left            =   225
         TabIndex        =   78
         Tag             =   "G"
         Top             =   300
         Width           =   1200
      End
      Begin VB.TextBox TXT_GRID_EMP_CD 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1245
         MaxLength       =   7
         TabIndex        =   77
         Tag             =   "作业人员"
         Top             =   1860
         Width           =   915
      End
      Begin VB.CheckBox CHK_TOP_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "不合格"
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   2430
         TabIndex        =   71
         Tag             =   "N"
         Top             =   900
         Width           =   900
      End
      Begin VB.CheckBox CHK_TOP_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "合格"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   1605
         TabIndex        =   70
         Tag             =   "Y"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox TXT_TOP_GRID_GRD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1590
         MaxLength       =   1
         TabIndex        =   69
         Text            =   " "
         Top             =   570
         Width           =   1815
      End
      Begin VB.TextBox TXT_BOT_GRID_GRD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3420
         MaxLength       =   1
         TabIndex        =   68
         Text            =   " "
         Top             =   570
         Width           =   1890
      End
      Begin VB.CheckBox CHK_BOT_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "不合格"
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   4350
         TabIndex        =   67
         Tag             =   "N"
         Top             =   900
         Width           =   900
      End
      Begin VB.CheckBox CHK_BOT_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "合格"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   3525
         TabIndex        =   66
         Tag             =   "Y"
         Top             =   900
         Width           =   735
      End
      Begin CSTextLibCtl.sidbEdit SDB_TOP_GRID_DEEP 
         Height          =   330
         Left            =   1590
         TabIndex        =   72
         Top             =   1515
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   2190
         Top             =   1860
         Width           =   990
         _ExtentX        =   1746
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
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Index           =   0
         Left            =   225
         Top             =   570
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "修磨后判定"
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
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   225
         Top             =   1170
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "修磨面积比%"
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
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   225
         Top             =   1515
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "修磨深度"
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
         Left            =   3420
         Top             =   240
         Width           =   1890
         _ExtentX        =   3334
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
      Begin CSTextLibCtl.sitxEdit TXT_GRID_TIME 
         Height          =   315
         Left            =   3225
         TabIndex        =   73
         Top             =   1860
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
      Begin CSTextLibCtl.sidbEdit SDB_TOP_GRID_YRD 
         Height          =   330
         Left            =   1590
         TabIndex        =   74
         Top             =   1170
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
      Begin CSTextLibCtl.sidbEdit SDB_BOT_GRID_DEEP 
         Height          =   330
         Left            =   3420
         TabIndex        =   75
         Top             =   1515
         Width           =   1890
         _Version        =   262145
         _ExtentX        =   3334
         _ExtentY        =   582
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_BOT_GRID_YRD 
         Height          =   330
         Left            =   3420
         TabIndex        =   76
         Top             =   1170
         Width           =   1890
         _Version        =   262145
         _ExtentX        =   3334
         _ExtentY        =   582
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Index           =   1
         Left            =   1590
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   225
         Top             =   1860
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   556
         Caption         =   "作业人员"
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
   Begin VB.TextBox TXT_INSP_FLAW 
      Height          =   315
      Index           =   1
      Left            =   450
      TabIndex        =   80
      Top             =   4845
      Visible         =   0   'False
      Width           =   285
   End
   Begin Threed.SSFrame SF4 
      Height          =   4665
      Left            =   10050
      TabIndex        =   36
      Top             =   4620
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8229
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   16711680
      BackColor       =   14737632
      Caption         =   "判定"
      Begin VB.TextBox TXT_PROC_CD 
         Alignment       =   2  'Center
         BackColor       =   &H00E1E4CD&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   95
         Tag             =   "表面判定"
         Text            =   " "
         Top             =   1875
         Width           =   840
      End
      Begin CSTextLibCtl.sidbEdit SDB_Mn 
         Height          =   225
         Left            =   1230
         TabIndex        =   94
         Top             =   1470
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
         Left            =   3555
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   1815
         Width           =   1335
      End
      Begin VB.TextBox txt_Scrap_code 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3135
         MaxLength       =   1
         TabIndex        =   90
         Tag             =   "原因"
         Top             =   1815
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
         Left            =   3660
         MaxLength       =   40
         TabIndex        =   89
         Tag             =   "STDSPEC"
         Top             =   2910
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
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   88
         Tag             =   "STDSPEC"
         Top             =   2910
         Width           =   2840
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
         Left            =   135
         MaxLength       =   18
         TabIndex        =   87
         Tag             =   "标准号"
         Top             =   2910
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
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   86
         Tag             =   "STDSPEC"
         Top             =   2580
         Width           =   2840
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
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   85
         Tag             =   "标准代码"
         Top             =   2580
         Width           =   1965
      End
      Begin VB.TextBox TXT_SURF_GRD 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1610
         Locked          =   -1  'True
         TabIndex        =   63
         Tag             =   "表面判定"
         Text            =   " "
         Top             =   300
         Width           =   840
      End
      Begin VB.TextBox TXT_INSP_MAN 
         Height          =   330
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   38
         Tag             =   "检查员"
         Top             =   3825
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_MAIN_GRD 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   1610
         Locked          =   -1  'True
         TabIndex        =   37
         Tag             =   "表面等级判定"
         Top             =   750
         Width           =   840
      End
      Begin InDate.ULabel ULabel22 
         Height          =   330
         Index           =   0
         Left            =   135
         Top             =   750
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
      Begin InDate.ULabel ULabel31 
         Height          =   330
         Left            =   150
         Top             =   3825
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         Caption         =   "检查员"
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
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   150
         Top             =   4230
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sitxEdit TXT_INSP_OCCR_TIME 
         Height          =   315
         Left            =   1380
         TabIndex        =   39
         Tag             =   "检查时间"
         Top             =   4230
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
      Begin InDate.ULabel ULabel1 
         Height          =   330
         Left            =   150
         Top             =   3420
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
      Begin InDate.ULabel ULabel36 
         Height          =   330
         Left            =   135
         Top             =   300
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
         Left            =   135
         Top             =   2250
         Width           =   4800
         _ExtentX        =   8467
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
         Left            =   2490
         Top             =   1815
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
         Left            =   135
         Top             =   1410
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
         Left            =   135
         Top             =   1815
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
         Left            =   2490
         TabIndex        =   98
         Top             =   300
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   3300
            TabIndex        =   99
            Text            =   " "
            Top             =   30
            Width           =   225
         End
         Begin Threed.SSOption opt_CHK_SUR_GRD 
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   100
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
            TabIndex        =   101
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
         Height          =   1005
         Left            =   2490
         TabIndex        =   102
         Top             =   750
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1773
         _Version        =   196609
         BackColor       =   14737632
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   103
            Top             =   30
            Width           =   765
            _ExtentX        =   1349
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
            Left            =   60
            TabIndex        =   104
            Top             =   270
            Width           =   1845
            _ExtentX        =   3254
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
            Caption         =   "订单外一级(改判)"
         End
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   105
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "订单外二级(协议板)"
         End
         Begin Threed.SSOption opt_CHK_PRD_GRD 
            Height          =   255
            Index           =   3
            Left            =   1710
            TabIndex        =   106
            Top             =   90
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
            Index           =   4
            Left            =   60
            TabIndex        =   107
            Top             =   720
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
            Height          =   255
            Index           =   5
            Left            =   1710
            TabIndex        =   108
            Top             =   720
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
      Begin Threed.SSFrame SSFrame5 
         Height          =   345
         Left            =   1380
         TabIndex        =   109
         Top             =   3420
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   609
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.TextBox txtCl 
            Height          =   285
            Left            =   1320
            TabIndex        =   115
            Top             =   180
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.TextBox txtGrid 
            Height          =   285
            Left            =   1110
            TabIndex        =   114
            Top             =   180
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.TextBox txtGas 
            Height          =   285
            Left            =   900
            TabIndex        =   113
            Top             =   180
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.CheckBox chkGas 
            BackColor       =   &H00E0E0E0&
            Caption         =   "GAS"
            Height          =   210
            Left            =   60
            TabIndex        =   112
            Tag             =   "C"
            Top             =   90
            Width           =   645
         End
         Begin VB.CheckBox chkGrid 
            BackColor       =   &H00E0E0E0&
            Caption         =   "修磨"
            Height          =   210
            Left            =   750
            TabIndex        =   111
            Tag             =   "G"
            Top             =   90
            Width           =   720
         End
         Begin VB.CheckBox chkCl 
            BackColor       =   &H00E0E0E0&
            Caption         =   "冷矫直"
            Height          =   210
            Left            =   1560
            TabIndex        =   110
            Tag             =   "G"
            Top             =   90
            Width           =   900
         End
      End
   End
   Begin Threed.SSFrame sf3 
      Height          =   2445
      Left            =   4575
      TabIndex        =   18
      Top             =   4620
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   4313
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
      Begin VB.TextBox TXT_INSP_WGT_GRD 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1935
         Width           =   1065
      End
      Begin VB.TextBox TXT_INSP_THK_GRD 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1935
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_LEN_GRD 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1935
         Width           =   1095
      End
      Begin VB.TextBox TXT_INSP_WID_GRD 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1935
         Width           =   960
      End
      Begin InDate.ULabel ULabel28 
         Height          =   315
         Left            =   2190
         Top             =   285
         Width           =   930
         _ExtentX        =   1640
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
         Left            =   1230
         Top             =   285
         Width           =   930
         _ExtentX        =   1640
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
         Left            =   3150
         Top             =   285
         Width           =   1065
         _ExtentX        =   1879
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
      Begin InDate.ULabel ULabel33 
         Height          =   315
         Left            =   210
         Top             =   1935
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_WGT_ORD 
         Height          =   315
         Left            =   4230
         TabIndex        =   13
         Top             =   945
         Width           =   1065
         _Version        =   262145
         _ExtentX        =   1879
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Left            =   4230
         TabIndex        =   14
         Top             =   615
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
         Left            =   2190
         TabIndex        =   8
         Top             =   1275
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   14737632
         BackColor       =   14737632
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
         Left            =   3150
         TabIndex        =   11
         Top             =   1275
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Left            =   2190
         TabIndex        =   9
         Top             =   1605
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Left            =   1230
         TabIndex        =   10
         Top             =   1605
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Left            =   3150
         TabIndex        =   12
         Top             =   1605
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Left            =   4245
         TabIndex        =   15
         Top             =   1605
         Width           =   1050
         _Version        =   262145
         _ExtentX        =   1852
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Left            =   2190
         TabIndex        =   6
         Top             =   615
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
         Left            =   1230
         TabIndex        =   7
         Top             =   615
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
         Left            =   3150
         TabIndex        =   35
         Top             =   615
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
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
         Left            =   210
         Top             =   1605
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
         Left            =   210
         Top             =   615
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
         Left            =   1230
         TabIndex        =   58
         Top             =   1275
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Left            =   4245
         TabIndex        =   59
         Top             =   1275
         Width           =   1050
         _Version        =   262145
         _ExtentX        =   1852
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Left            =   210
         Top             =   1275
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel44 
         Height          =   315
         Left            =   4245
         Top             =   285
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "重量"
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
         Left            =   2190
         TabIndex        =   60
         Top             =   945
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Left            =   1230
         TabIndex        =   61
         Top             =   945
         Width           =   960
         _Version        =   262145
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Left            =   3150
         TabIndex        =   62
         Top             =   945
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
         Left            =   210
         Top             =   945
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame sf1 
      Height          =   4665
      Left            =   60
      TabIndex        =   17
      Top             =   4620
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   8229
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
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   225
         Top             =   690
         Width           =   1185
         _ExtentX        =   2090
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
         ForeColor       =   16711680
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Height          =   315
         Index           =   5
         Left            =   705
         TabIndex        =   84
         Top             =   555
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Height          =   315
         Index           =   4
         Left            =   390
         TabIndex        =   83
         Top             =   555
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Height          =   315
         Index           =   3
         Left            =   90
         TabIndex        =   82
         Top             =   540
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Height          =   315
         Index           =   2
         Left            =   705
         TabIndex        =   81
         Top             =   225
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   79
         Top             =   225
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "尾部"
         Height          =   240
         Index           =   17
         Left            =   3345
         TabIndex        =   54
         Tag             =   "B"
         Top             =   3825
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中部"
         Height          =   195
         Index           =   16
         Left            =   3345
         TabIndex        =   53
         Tag             =   "M"
         Top             =   3600
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头部"
         Height          =   195
         Index           =   15
         Left            =   3345
         TabIndex        =   52
         Tag             =   "T"
         Top             =   3375
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "尾部"
         Height          =   240
         Index           =   14
         Left            =   2370
         TabIndex        =   51
         Tag             =   "B"
         Top             =   3825
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中部"
         Height          =   195
         Index           =   13
         Left            =   2370
         TabIndex        =   50
         Tag             =   "M"
         Top             =   3600
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头部"
         Height          =   195
         Index           =   12
         Left            =   2370
         TabIndex        =   49
         Tag             =   "T"
         Top             =   3375
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "尾部"
         Height          =   240
         Index           =   11
         Left            =   1410
         TabIndex        =   48
         Tag             =   "B"
         Top             =   3825
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中部"
         Height          =   195
         Index           =   10
         Left            =   1410
         TabIndex        =   47
         Tag             =   "M"
         Top             =   3600
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头部"
         Height          =   195
         Index           =   9
         Left            =   1410
         TabIndex        =   46
         Tag             =   "T"
         Top             =   3375
         Width           =   810
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   3
         Left            =   1410
         TabIndex        =   45
         Text            =   " "
         Top             =   3045
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   4
         Left            =   2370
         TabIndex        =   44
         Text            =   " "
         Top             =   3030
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   5
         Left            =   3345
         TabIndex        =   43
         Text            =   " "
         Top             =   3030
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   3
         Left            =   1425
         TabIndex        =   42
         Top             =   2685
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   4
         Left            =   2370
         TabIndex        =   41
         Top             =   2700
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   5
         Left            =   3345
         TabIndex        =   40
         Top             =   2700
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   2
         Left            =   3345
         TabIndex        =   2
         Top             =   690
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   1
         Left            =   2370
         TabIndex        =   1
         Top             =   690
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   0
         Left            =   1410
         TabIndex        =   0
         Top             =   675
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   2
         Left            =   3345
         TabIndex        =   30
         Text            =   " "
         Top             =   1020
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   1
         Left            =   2385
         TabIndex        =   29
         Text            =   " "
         Top             =   1020
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_PART 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   0
         Left            =   1410
         TabIndex        =   28
         Text            =   " "
         Top             =   1005
         Width           =   960
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头部"
         Height          =   195
         Index           =   0
         Left            =   1410
         TabIndex        =   27
         Tag             =   "T"
         Top             =   1365
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中部"
         Height          =   195
         Index           =   1
         Left            =   1410
         TabIndex        =   26
         Tag             =   "M"
         Top             =   1590
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "尾部"
         Height          =   240
         Index           =   2
         Left            =   1410
         TabIndex        =   25
         Tag             =   "B"
         Top             =   1815
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头部"
         Height          =   195
         Index           =   3
         Left            =   2370
         TabIndex        =   24
         Tag             =   "T"
         Top             =   1365
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中部"
         Height          =   195
         Index           =   4
         Left            =   2370
         TabIndex        =   23
         Tag             =   "M"
         Top             =   1590
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "尾部"
         Height          =   240
         Index           =   5
         Left            =   2370
         TabIndex        =   22
         Tag             =   "B"
         Top             =   1815
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头部"
         Height          =   195
         Index           =   6
         Left            =   3345
         TabIndex        =   21
         Tag             =   "T"
         Top             =   1365
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中部"
         Height          =   195
         Index           =   7
         Left            =   3345
         TabIndex        =   20
         Tag             =   "M"
         Top             =   1590
         Width           =   810
      End
      Begin VB.CheckBox CHK_PART 
         BackColor       =   &H00E0E0E0&
         Caption         =   "尾部"
         Height          =   240
         Index           =   8
         Left            =   3345
         TabIndex        =   19
         Tag             =   "B"
         Top             =   1815
         Width           =   810
      End
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   0
         Left            =   1410
         TabIndex        =   3
         Top             =   2085
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   1
         Left            =   2370
         TabIndex        =   4
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   2
         Left            =   3345
         TabIndex        =   5
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
            Size            =   9.76
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
         TabIndex        =   55
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   4
         Left            =   2370
         TabIndex        =   56
         Top             =   4080
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LTH 
         Height          =   315
         Index           =   5
         Left            =   3345
         TabIndex        =   57
         Top             =   4080
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   225
         Top             =   2700
         Width           =   1185
         _ExtentX        =   2090
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
         ForeColor       =   16711680
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
            Size            =   9.76
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
            Size            =   9.76
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
         Left            =   3345
         Top             =   360
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         Caption         =   "小缺陷2"
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
   Begin Threed.SSFrame Single 
      Height          =   1050
      Left            =   60
      TabIndex        =   16
      Top             =   75
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   1852
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
      Begin VB.TextBox text_cur_inv_code 
         Height          =   315
         Left            =   6810
         MaxLength       =   2
         TabIndex        =   121
         Tag             =   "起始库"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox text_cur_inv 
         Height          =   315
         Left            =   7200
         TabIndex        =   120
         Top             =   600
         Width           =   1560
      End
      Begin VB.TextBox TXT_STLGRD 
         Height          =   285
         Left            =   5040
         TabIndex        =   119
         Top             =   615
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox TXT_APLY_ENDUSE_CD 
         Height          =   285
         Left            =   4830
         TabIndex        =   118
         Top             =   600
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox TXT_PROC_FLAG 
         Height          =   270
         Left            =   4620
         TabIndex        =   117
         Top             =   630
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox TXT_UST_FLAG 
         Height          =   270
         Left            =   4410
         TabIndex        =   116
         Top             =   630
         Visible         =   0   'False
         Width           =   210
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
         Left            =   1435
         MaxLength       =   18
         TabIndex        =   96
         Tag             =   "标准号"
         Top             =   600
         Width           =   2925
      End
      Begin VB.TextBox TXT_PLATE_NO 
         Height          =   330
         Left            =   1435
         MaxLength       =   14
         TabIndex        =   64
         Top             =   180
         Width           =   1830
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   225
         Top             =   180
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
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   8940
         Top             =   600
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "生产时间"
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
      Begin CSTextLibCtl.sitxEdit SDT_PROD_DATE 
         Height          =   315
         Left            =   10125
         TabIndex        =   93
         Top             =   600
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
      Begin InDate.ULabel ULabel22 
         Height          =   300
         Index           =   4
         Left            =   225
         Top             =   600
         Width           =   1170
         _ExtentX        =   2064
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
      Begin CSTextLibCtl.sitxEdit SDT_PROD_TO_DATE 
         Height          =   315
         Left            =   11340
         TabIndex        =   97
         Top             =   600
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   5610
         Top             =   600
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "当库"
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
      Height          =   3435
      Left            =   60
      TabIndex        =   92
      Top             =   1140
      Width           =   15030
      _Version        =   393216
      _ExtentX        =   26511
      _ExtentY        =   6059
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
      MaxCols         =   17
      MaxRows         =   20
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "DGA1070C.frx":0000
   End
End
Attribute VB_Name = "DGA1070C"
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

Private Sub Form_Define()
    Dim iIndex As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDT_PROD_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDT_PROD_TO_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_stdspec_chg_ref, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_UST_FLAG, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_PROC_FLAG, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(TXT_APLY_ENDUSE_CD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_STLGRD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                                                                                                                                                
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
                Call Gp_Ms_Collection(txtCl, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(TXT_INSP_MAN, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
          Call Gp_Ms_Collection(txt_stdspec, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_stdspec_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_stdspec_chg, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_stdspec_name_chg, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Scrap_code, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Scrap_name, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(SDB_Mn, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_PROC_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
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
    Mc1.Add Item:="AGC2020C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="DGA1070C.P_REFER", Key:="P-R"
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
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="DGA1070C.P_SREFER", Key:="P-R"
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
    
   
    text_cur_inv_code = "00"
    text_cur_inv = "中厚板卷厂"
    
    If TXT_PLATE_NO <> "" Then
       Call Form_Ref
    End If
       
    Screen.MousePointer = vbDefault

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
        Call Gp_SSCheck_Cls(MC("sControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)

        TXT_INSP_MAN = sUserID
        
        For iCount = 0 To 5
            TXT_INSP_FLAW_NAME(iCount).Text = ""
        Next iCount
        
        ss1.BlockMode = True
        ss1.ROW = -1
        ss1.Col = -1
        ss1.BackColor = &HFFFFFF
        ss1.BlockMode = False
    End If
End Sub

Public Sub Form_Ref()

    Call Form_Cls
    Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1)
    
    If ss1.MaxRows > 0 Then
       ss1.ROW = 1
       ss1.Col = 1
       TXT_PLATE_NO.Text = ss1.Text
    End If
    
    If Len(TXT_PLATE_NO.Text) = 14 Then
        If Gf_Ms_Refer(M_CN1, Mc1, , , False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            ''''''''''''''''ADD BY GUOLI AT 200712071330''''''''''''''''''
            If opt_CHK_SUR_GRD(0).Value = True Then
               TXT_SURF_GRD = "Y"
            ElseIf opt_CHK_SUR_GRD(1).Value = True Then
               TXT_SURF_GRD = "N"
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            TXT_INSP_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
            TXT_INSP_MAN = sUserID
            'Call Display_Data_Edit
        End If
    End If
     
End Sub

Public Sub Form_Pro()

    Dim sMesg   As String
    Dim iCount  As Integer
    
    For iCount = 0 To 5
        If TXT_INSP_FLAW_NAME(iCount).Text <> "" And TXT_INSP_PART(iCount).Text = "" Then
            sMesg = " 请输入缺陷部位 ！"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        End If
    Next iCount
    
    
    If Trim(TXT_INSP_MAIN_GRD.Text) <> "4" Then
        If Trim(TXT_SURF_GRD.Text) = "" Then
            sMesg = " 请输入表面判定 ！"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        End If
        
    End If
    
    If Not Gp_DateCheck(TXT_INSP_OCCR_TIME) Then
        sMesg = " 请正确输入检查时间 ！"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    If CHK_GRID_FLAG.Value = ssCBChecked Then
        If Not Gp_DateCheck(TXT_GRID_TIME) Then
            sMesg = " 请正确输入修磨时间 ！"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
        End If
        If Trim(TXT_GRID_EMP_CD.Text) = "" Then
            TXT_GRID_EMP_CD.Text = sUserID
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
    
    
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        TXT_INSP_MAN.Text = sUserID
       If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    End If
    
    Call Form_Ref
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
    ElseIf Index = 1 Then
       TXT_INSP_MAIN_GRD = "2"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
    ElseIf Index = 2 Then
        TXT_INSP_MAIN_GRD = "3"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
    ElseIf Index = 3 Then
        TXT_INSP_MAIN_GRD = "4"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
    ElseIf Index = 4 Then
        TXT_INSP_MAIN_GRD = "5"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &HFF&       'red
       opt_CHK_PRD_GRD(5).ForeColor = &H80000012  'black
    ElseIf Index = 5 Then
        TXT_INSP_MAIN_GRD = "7"
       opt_CHK_PRD_GRD(0).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(1).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(2).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(3).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(4).ForeColor = &H80000012  'black
       opt_CHK_PRD_GRD(5).ForeColor = &HFF&       'red
       txt_Scrap_code.Enabled = True
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
    sQuery = sQuery & "             ,'" & Trim(TXT_STLGRD.Text) & "'" & vbCrLf
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

Private Sub text_cur_inv_code_DblClick()
    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=text_cur_inv_code
        DD.rControl.Add Item:=text_cur_inv
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
     
        If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
            text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
        Else
          text_cur_inv.Text = ""
        End If
        
    End If
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
    
    If Len(Trim(TXT_INSP_FLAW(Index).Text)) = 2 Then
        TXT_INSP_FLAW_NAME(Index).Text = Gf_ComnNameFind(M_CN1, "G0002", Trim(TXT_INSP_FLAW(Index).Text), 1)
    Else
        TXT_INSP_FLAW_NAME(Index).Text = ""
    End If
End Sub

Private Sub TXT_INSP_MAN_DblClick()
    TXT_INSP_MAN.Text = sUserID
End Sub

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
    
'    If CHK_NEXT_PRC(1).Value = ssCBChecked Then
'        CHK_SUR_GRD(0).ForeColor = &H808080
'        CHK_SUR_GRD(0).Value = ssCBUnchecked
'        CHK_SUR_GRD(1).ForeColor = &HFF&
'        CHK_SUR_GRD(1).Value = ssCBChecked
'        TXT_SURF_GRD.Text = CHK_SUR_GRD(1).Tag
'        sCheck = ""
'        Exit Sub
'    End If
    
    If Index = 0 Then
        iNext = 1
    Else
        iNext = 0
    End If
    
'    If CHK_SUR_GRD(Index).Value = ssCBUnchecked Then
'        If CHK_SUR_GRD(iNext).Value = ssCBUnchecked Then
'            TXT_SURF_GRD.Text = ""
'            CHK_SUR_GRD(Index).ForeColor = &H808080
'            sCheck = ""
'            Exit Sub
'        End If
'    Else
'        CHK_SUR_GRD(iNext).Value = ssCBUnchecked
'    End If
    
'    CHK_SUR_GRD(Index).ForeColor = &HFF&
'    CHK_SUR_GRD(Index).Value = ssCBChecked
'
'    CHK_SUR_GRD(iNext).ForeColor = &H808080
'    CHK_SUR_GRD(iNext).Value = ssCBUnchecked
'
'    TXT_SURF_GRD.Text = CHK_SUR_GRD(Index).Tag
    sCheck = ""
    
End Sub

Private Sub CHK_PRD_GRD_Click(Index As Integer)
    Dim iCount      As Integer
    Dim iIndexStr   As Integer
    
    If sCheck <> "" Then Exit Sub

    iCount = 0
    sCheck = "**"
    
'    If CHK_PRD_GRD(Index).Value = ssCBUnchecked Then
'        For iIndexStr = 0 To 5
'            If CHK_PRD_GRD(iIndexStr).Value = ssCBChecked Then
'               iCount = iCount + 1
'            End If
'        Next iIndexStr
'        If iCount = 0 Then
'            TXT_INSP_MAIN_GRD.Text = ""
'            CHK_PRD_GRD(Index).ForeColor = &H808080
'            sCheck = ""
'            Exit Sub
'        End If
'    Else
'        For iIndexStr = 0 To 5
'            CHK_PRD_GRD(iIndexStr).ForeColor = &H808080
'            CHK_PRD_GRD(iIndexStr).Value = ssCBUnchecked
'        Next iIndexStr
'    End If
'
'    CHK_PRD_GRD(Index).ForeColor = &HFF&
'    CHK_PRD_GRD(Index).Value = ssCBChecked
    
    'TXT_INSP_MAIN_GRD.Text = CHK_PRD_GRD(Index).Tag
                 
    txt_stdspec_chg.Text = ""
    txt_stdspec_name_chg.Text = ""
'    If CHK_PRD_GRD(0).Value = ssCBChecked Or CHK_PRD_GRD(1).Value = ssCBChecked Or CHK_PRD_GRD(2).Value = ssCBChecked Or CHK_PRD_GRD(3).Value = ssCBChecked Or CHK_PRD_GRD(5).Value = ssCBChecked Then
'        txt_stdspec_chg.Enabled = True
'    Else
'        txt_stdspec_chg.Enabled = False
'    End If
'
'    If CHK_PRD_GRD(4).Value = ssCBChecked Then
'        txt_Scrap_code.Enabled = True
'    Else
'        txt_Scrap_code.Enabled = False
'    End If
    
'   MODEFIED BY YANGMENG AT 07.01.30
'    待判时处理
'    sCheck = "**"
'    For iIndexStr = 0 To 2
'        If CHK_PRD_GRD(5).Value = ssCBChecked Then
'            TXT_NEXT_PROC.Text = ""
'        Else
'            CHK_NEXT_PRC(iIndexStr).Enabled = True
'        End If
'    Next iIndexStr
'
'    For iIndexStr = 0 To 1
'        If CHK_PRD_GRD(5).Value = ssCBChecked Then
'            CHK_SUR_GRD(iIndexStr).Enabled = False
'            TXT_SURF_GRD.Text = ""
'            CHK_SUR_GRD(iIndexStr).Value = ssCBUnchecked
'            CHK_SUR_GRD(iIndexStr).ForeColor = &H808080
'        Else
'            CHK_SUR_GRD(iIndexStr).Enabled = True
'        End If
'    Next iIndexStr
'
'    sCheck = ""
'
'    If Index = 0 Then
'        CHK_NEXT_PRC(2).Value = ssCBChecked
'        Call CHK_NEXT_PRC_Click(2)
'    End If
        
End Sub
'
'Private Sub CHK_NEXT_PRC_Click(Index As Integer)
'    Dim iCount      As Integer
'    Dim iIndexStr   As Integer
'
'    If sCheck <> "" Then Exit Sub
'
'    iCount = 0
'    sCheck = "**"
'
'    If CHK_NEXT_PRC(Index).Value = ssCBUnchecked Then
'        For iIndexStr = 0 To 2
'            If CHK_NEXT_PRC(iIndexStr).Value = ssCBChecked Then
'               iCount = iCount + 1
'            End If
'        Next iIndexStr
'        If iCount = 0 Then
'            TXT_NEXT_PROC.Text = ""
'            CHK_NEXT_PRC(Index).ForeColor = &H808080
'            sCheck = ""
'            Exit Sub
'        End If
'    Else
'        For iIndexStr = 0 To 2
'            CHK_NEXT_PRC(iIndexStr).ForeColor = &H808080
'            CHK_NEXT_PRC(iIndexStr).Value = ssCBUnchecked
'        Next iIndexStr
'    End If
'
'    CHK_NEXT_PRC(Index).ForeColor = &HFF&
'    CHK_NEXT_PRC(Index).Value = ssCBChecked
'
'    TXT_NEXT_PROC.Text = CHK_NEXT_PRC(Index).Tag
'
'    sCheck = ""
'
'    If CHK_NEXT_PRC(0).Value = ssCBChecked Or CHK_NEXT_PRC(1).Value = ssCBChecked Then
'        For iIndexStr = 0 To 4
''            CHK_PRD_GRD(iIndexStr).ForeColor = &H808080
''            CHK_PRD_GRD(iIndexStr).Value = ssCBUnchecked
''            TXT_INSP_MAIN_GRD.Text = ""
'        Next iIndexStr
'        If CHK_NEXT_PRC(1).Value = ssCBChecked Then
'            CHK_SUR_GRD(1).Value = ssCBChecked
'            Call CHK_SUR_GRD_Click(1)
''            CHK_GRID_FLAG.Value = ssCBUnchecked
''            Call CHK_GRID_FLAG_Click
'        End If
'    End If
'
'End Sub

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
        

'        CHK_NEXT_PRC(2).Value = ssCBChecked
'        Call CHK_NEXT_PRC_Click(2)

'        TXT_NEXT_PROC.Text = ""

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

Private Sub ss1_EditChange(ByVal Col As Long, ByVal ROW As Long)
    Dim dThk        As Double
    Dim dWid        As Double
    Dim dLen        As Double
    Dim dLenSum     As Double
    
    Dim iIdr        As Integer
                       
    ss1.ROW = ROW
    ss1.Col = 2:  dThk = Val(ss1.Text & "")
    ss1.Col = 3:  dWid = Val(ss1.Text & "")
    ss1.Col = 4:  dLen = Val(ss1.Text & "")

    ss1.Col = 5
    ss1.Text = Cal_Plate_Wgt("WGT", dThk, dWid, dLen)
     
    For iIdr = 1 To ss1.MaxRows - 1
        ss1.ROW = iIdr
        ss1.Col = 4
        dLenSum = dLenSum + Val(ss1.Text & "")
    Next iIdr
    
    ss1.ROW = ss1.MaxRows
    ss1.Col = 4
    dLen = ss1.Text 'SDB_LEN.Value - dLenSum
    ss1.Text = dLen
    ss1.Col = 5
    ss1.Text = Cal_Plate_Wgt("WGT", dThk, dWid, dLen)
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

            'Call Display_Data_Edit
        End If
        
     TXT_INSP_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
     TXT_INSP_MAN = sUserID
        
    End If
    
End Sub

Private Sub txt_Plt_Change()

End Sub



Private Sub txt_Scrap_code_DblClick()
    Call txt_Scrap_code_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stdspec_chg_DblClick()
    Call txt_stdspec_chg_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        txt_stdspec_yy.Text = ""
        DD.rControl.Add Item:=txt_stdspec_chg
        DD.rControl.Add Item:=txt_stdspec_yy
        DD.rControl.Add Item:=txt_stdspec_name_chg

        Call Gf_StdSPEC_DD(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

Private Sub txt_Scrap_code_Change()
    
    If Len(Trim(txt_Scrap_code)) = txt_Scrap_code.MaxLength Then
        txt_Scrap_name.Text = Gf_ComnNameFind(M_CN1, "F0011", Trim(txt_Scrap_code.Text), 1)
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

        Call Gf_StdSPEC_DD(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

