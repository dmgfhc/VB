VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2020C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "表面检查实绩查询及修改_AGC2020C"
   ClientHeight    =   9465
   ClientLeft      =   870
   ClientTop       =   1755
   ClientWidth     =   15240
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
   ScaleHeight     =   9465
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_NEXT_PROC 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   330
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   122
      Tag             =   "后道工序"
      Top             =   9450
      Visible         =   0   'False
      Width           =   870
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   2190
      Left            =   4920
      TabIndex        =   120
      Top             =   7035
      Width           =   5675
      _ExtentX        =   10001
      _ExtentY        =   3863
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_ORD_NO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         TabIndex        =   136
         Top             =   480
         Width           =   1485
      End
      Begin VB.TextBox TXT_SPEC_PROC_NAME 
         Height          =   315
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   124
         Tag             =   "后道工序"
         Top             =   810
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.TextBox TXT_SPEC_PROC 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1110
         MaxLength       =   1
         TabIndex        =   123
         Tag             =   "后道工序"
         Top             =   810
         Visible         =   0   'False
         Width           =   870
      End
      Begin InDate.ULabel ULabel27 
         Height          =   315
         Left            =   270
         Top             =   810
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "特殊工序"
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
      Begin VB.TextBox txt_woo_rsn 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4560
         MaxLength       =   2
         TabIndex        =   130
         Tag             =   "余材原因"
         Top             =   120
         Width           =   735
      End
      Begin VB.CheckBox CHK_CL_FL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "矫直指示"
         Height          =   315
         Left            =   2280
         TabIndex        =   129
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TXT_CL 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   128
         Top             =   480
         Width           =   870
      End
      Begin VB.TextBox TXT_GAS 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   127
         Top             =   120
         Width           =   870
      End
      Begin VB.CheckBox CHK_GAS_FL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "切割指示"
         Height          =   315
         Left            =   2280
         TabIndex        =   126
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox TXT_REMARK 
         Height          =   1095
         Left            =   1380
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   125
         Tag             =   "后道工序"
         Top             =   900
         Width           =   4020
      End
      Begin InDate.ULabel ULabel1 
         Height          =   1095
         Left            =   210
         Top             =   900
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1931
         Caption         =   "备注"
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
      Begin InDate.ULabel ULabel32 
         Height          =   315
         Left            =   210
         Top             =   480
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "矫直"
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
      Begin InDate.ULabel ULabel35 
         Height          =   315
         Left            =   210
         Top             =   120
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "切割"
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
      Begin InDate.ULabel ULabel77 
         Height          =   315
         Left            =   3390
         Top             =   120
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "余材原因"
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
      Begin InDate.ULabel ULabel40 
         Height          =   315
         Left            =   3390
         Top             =   480
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         Caption         =   "订单号"
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
      Height          =   315
      Left            =   15540
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   117
      Tag             =   "STDSPEC"
      Top             =   6690
      Visible         =   0   'False
      Width           =   1845
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
      Height          =   315
      Left            =   15540
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   116
      Tag             =   "STDSPEC"
      Top             =   7020
      Visible         =   0   'False
      Width           =   1845
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2295
      Left            =   90
      TabIndex        =   12
      Top             =   6930
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   4048
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
         Left            =   165
         TabIndex        =   26
         Tag             =   "G"
         Top             =   330
         Width           =   1110
      End
      Begin VB.TextBox TXT_GRID_EMP_CD 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   7
         TabIndex        =   25
         Tag             =   "作业人员"
         Top             =   1410
         Width           =   1065
      End
      Begin VB.CheckBox CHK_TOP_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "不合格"
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   3795
         TabIndex        =   18
         Tag             =   "N"
         Top             =   720
         Width           =   900
      End
      Begin VB.CheckBox CHK_TOP_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "合格"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   3795
         TabIndex        =   17
         Tag             =   "Y"
         Top             =   480
         Width           =   900
      End
      Begin VB.TextBox TXT_TOP_GRID_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1350
         MaxLength       =   1
         TabIndex        =   16
         Text            =   " "
         Top             =   630
         Width           =   870
      End
      Begin VB.TextBox TXT_BOT_GRID_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1350
         MaxLength       =   1
         TabIndex        =   15
         Text            =   " "
         Top             =   1020
         Width           =   870
      End
      Begin VB.CheckBox CHK_BOT_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "不合格"
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   3795
         TabIndex        =   14
         Tag             =   "N"
         Top             =   1260
         Width           =   900
      End
      Begin VB.CheckBox CHK_BOT_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "合格"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   3795
         TabIndex        =   13
         Tag             =   "Y"
         Top             =   1020
         Width           =   900
      End
      Begin CSTextLibCtl.sidbEdit SDB_TOP_GRID_DEEP 
         Height          =   315
         Left            =   2985
         TabIndex        =   19
         Top             =   630
         Width           =   735
         _Version        =   262145
         _ExtentX        =   1296
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   165
         Top             =   1800
         Width           =   1155
         _ExtentX        =   2037
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
         Left            =   1350
         Top             =   270
         Width           =   2370
         _ExtentX        =   4180
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
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   165
         Top             =   1020
         Width           =   1155
         _ExtentX        =   2037
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
         Left            =   1350
         TabIndex        =   20
         Top             =   1800
         Width           =   2145
         _Version        =   262145
         _ExtentX        =   3784
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
         Height          =   315
         Left            =   2230
         TabIndex        =   21
         Top             =   630
         Width           =   735
         _Version        =   262145
         _ExtentX        =   1296
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
      Begin CSTextLibCtl.sidbEdit SDB_BOT_GRID_DEEP 
         Height          =   315
         Left            =   2985
         TabIndex        =   22
         Top             =   1020
         Width           =   735
         _Version        =   262145
         _ExtentX        =   1296
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
      Begin CSTextLibCtl.sidbEdit SDB_BOT_GRID_YRD 
         Height          =   315
         Left            =   2230
         TabIndex        =   23
         Top             =   1020
         Width           =   735
         _Version        =   262145
         _ExtentX        =   1296
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
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Index           =   1
         Left            =   165
         Top             =   630
         Width           =   1155
         _ExtentX        =   2037
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
         Left            =   165
         Top             =   1410
         Width           =   1155
         _ExtentX        =   2037
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
   Begin Threed.SSFrame SF4 
      Height          =   5055
      Left            =   10560
      TabIndex        =   0
      Top             =   4170
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   8916
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   16711680
      BackColor       =   14737632
      Caption         =   " 判定"
      Begin VB.TextBox TXT_PROC_CD 
         Alignment       =   2  'Center
         BackColor       =   &H00E1E4CD&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2250
         Locked          =   -1  'True
         TabIndex        =   121
         Tag             =   "表面判定"
         Text            =   " "
         Top             =   330
         Width           =   1440
      End
      Begin VB.TextBox TXT_APLY_ENDUSE_CD 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   960
         TabIndex        =   119
         Top             =   1680
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   3690
         Width           =   2205
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   3660
         MaxLength       =   3
         TabIndex        =   52
         Top             =   3690
         Width           =   765
      End
      Begin VB.CheckBox CHK_PRD_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "待判"
         Height          =   240
         Index           =   5
         Left            =   3660
         TabIndex        =   38
         Tag             =   "4"
         Top             =   1260
         Width           =   1020
      End
      Begin VB.TextBox txt_Scrap_name 
         Height          =   315
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2250
         Width           =   2055
      End
      Begin VB.TextBox txt_Scrap_code 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   1665
         MaxLength       =   1
         TabIndex        =   35
         Tag             =   "原因"
         Top             =   2250
         Width           =   750
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
         Left            =   3930
         MaxLength       =   40
         TabIndex        =   29
         Tag             =   "STDSPEC"
         Top             =   4140
         Visible         =   0   'False
         Width           =   330
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
         Left            =   195
         MaxLength       =   18
         TabIndex        =   28
         Tag             =   "标准号"
         Top             =   3330
         Width           =   2805
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
         Height          =   315
         Left            =   195
         Locked          =   -1  'True
         TabIndex        =   27
         Tag             =   "标准代码"
         Top             =   3000
         Width           =   2805
      End
      Begin VB.CheckBox CHK_SUR_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "不合格"
         Height          =   195
         Index           =   1
         Left            =   3660
         TabIndex        =   11
         Tag             =   "N"
         Top             =   780
         Width           =   1020
      End
      Begin VB.CheckBox CHK_SUR_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "合格"
         Height          =   240
         Index           =   0
         Left            =   2610
         TabIndex        =   10
         Tag             =   "Y"
         Top             =   780
         Width           =   1050
      End
      Begin VB.TextBox TXT_SURF_GRD 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   9
         Tag             =   "表面判定"
         Text            =   " "
         Top             =   750
         Width           =   840
      End
      Begin VB.TextBox TXT_INSP_MAN 
         Height          =   315
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   7
         Tag             =   "检查员"
         Top             =   4155
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_MAIN_GRD 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   6
         Tag             =   "表面等级判定"
         Top             =   1290
         Width           =   840
      End
      Begin VB.CheckBox CHK_PRD_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "废钢"
         Height          =   240
         Index           =   4
         Left            =   195
         TabIndex        =   5
         Tag             =   "7"
         Top             =   1950
         Width           =   1050
      End
      Begin VB.CheckBox CHK_PRD_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "次品"
         Height          =   240
         Index           =   3
         Left            =   3660
         TabIndex        =   4
         Tag             =   "5"
         Top             =   1950
         Width           =   1020
      End
      Begin VB.CheckBox CHK_PRD_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "协议"
         Height          =   240
         Index           =   2
         Left            =   2610
         TabIndex        =   3
         Tag             =   "3"
         Top             =   1950
         Width           =   1050
      End
      Begin VB.CheckBox CHK_PRD_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "改判"
         Height          =   240
         Index           =   1
         Left            =   2610
         TabIndex        =   2
         Tag             =   "2"
         Top             =   1605
         Width           =   1050
      End
      Begin VB.CheckBox CHK_PRD_GRD 
         BackColor       =   &H00E0E0E0&
         Caption         =   "正品"
         Height          =   240
         Index           =   0
         Left            =   2610
         TabIndex        =   1
         Tag             =   "1"
         Top             =   1260
         Width           =   1050
      End
      Begin InDate.ULabel ULabel22 
         Height          =   330
         Index           =   0
         Left            =   195
         Top             =   1290
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
         Height          =   315
         Left            =   195
         Top             =   4155
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "检查人员"
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
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   195
         Top             =   4560
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
         Left            =   1440
         TabIndex        =   8
         Tag             =   "检查时间"
         Top             =   4560
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
      Begin InDate.ULabel ULabel36 
         Height          =   330
         Left            =   195
         Top             =   750
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
         Left            =   195
         Top             =   2670
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   529
         Caption         =   "     标准号            |    Mn成分"
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
         Height          =   315
         Left            =   195
         Top             =   2250
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         Caption         =   "判废原因"
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
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   195
         Top             =   3690
         Width           =   1200
         _ExtentX        =   2117
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
      Begin VB.TextBox TXT_PROC_FLAG 
         Height          =   285
         Left            =   1935
         TabIndex        =   24
         Top             =   1695
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox TXT_UST_FLAG 
         Height          =   270
         Left            =   1725
         TabIndex        =   37
         Top             =   1695
         Visible         =   0   'False
         Width           =   210
      End
      Begin CSTextLibCtl.sidbEdit SDB_Mn 
         Height          =   645
         Left            =   3060
         TabIndex        =   118
         Top             =   3000
         Width           =   1440
         _Version        =   262145
         _ExtentX        =   2540
         _ExtentY        =   1138
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   255
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   18
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
         StartText.y     =   8
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   27
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   1
         BorderStyle     =   0
         FmtControl      =   1
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel22 
         Height          =   360
         Index           =   3
         Left            =   195
         Top             =   285
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   635
         Caption         =   "钢板状态 (          )"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
   End
   Begin VB.TextBox TXT_INSP_PART 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   5
      Left            =   8280
      TabIndex        =   77
      Text            =   " "
      Top             =   10020
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TXT_INSP_PART 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   4
      Left            =   7305
      TabIndex        =   76
      Text            =   " "
      Top             =   10020
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TXT_INSP_PART 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   3
      Left            =   6345
      TabIndex        =   75
      Text            =   " "
      Top             =   10020
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "头部"
      Height          =   195
      Index           =   9
      Left            =   6345
      TabIndex        =   74
      Tag             =   "T"
      Top             =   10365
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "中部"
      Height          =   195
      Index           =   10
      Left            =   6345
      TabIndex        =   73
      Tag             =   "M"
      Top             =   10590
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "尾部"
      Height          =   240
      Index           =   11
      Left            =   6345
      TabIndex        =   72
      Tag             =   "B"
      Top             =   10815
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "头部"
      Height          =   195
      Index           =   12
      Left            =   7305
      TabIndex        =   71
      Tag             =   "T"
      Top             =   10365
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "中部"
      Height          =   195
      Index           =   13
      Left            =   7305
      TabIndex        =   70
      Tag             =   "M"
      Top             =   10590
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "尾部"
      Height          =   240
      Index           =   14
      Left            =   7305
      TabIndex        =   69
      Tag             =   "B"
      Top             =   10815
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "头部"
      Height          =   195
      Index           =   15
      Left            =   8280
      TabIndex        =   68
      Tag             =   "T"
      Top             =   10365
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "中部"
      Height          =   195
      Index           =   16
      Left            =   8280
      TabIndex        =   67
      Tag             =   "M"
      Top             =   10590
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "尾部"
      Height          =   240
      Index           =   17
      Left            =   8280
      TabIndex        =   66
      Tag             =   "B"
      Top             =   10815
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "尾部"
      Height          =   240
      Index           =   8
      Left            =   3420
      TabIndex        =   65
      Tag             =   "B"
      Top             =   10845
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "中部"
      Height          =   195
      Index           =   7
      Left            =   3420
      TabIndex        =   64
      Tag             =   "M"
      Top             =   10620
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "头部"
      Height          =   195
      Index           =   6
      Left            =   3420
      TabIndex        =   63
      Tag             =   "T"
      Top             =   10395
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "尾部"
      Height          =   240
      Index           =   5
      Left            =   2415
      TabIndex        =   62
      Tag             =   "B"
      Top             =   10845
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "中部"
      Height          =   195
      Index           =   4
      Left            =   2415
      TabIndex        =   61
      Tag             =   "M"
      Top             =   10620
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "头部"
      Height          =   195
      Index           =   3
      Left            =   2415
      TabIndex        =   60
      Tag             =   "T"
      Top             =   10395
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "尾部"
      Height          =   240
      Index           =   2
      Left            =   1455
      TabIndex        =   59
      Tag             =   "B"
      Top             =   10845
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "中部"
      Height          =   195
      Index           =   1
      Left            =   1455
      TabIndex        =   58
      Tag             =   "M"
      Top             =   10620
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CheckBox CHK_PART 
      BackColor       =   &H00E0E0E0&
      Caption         =   "头部"
      Height          =   195
      Index           =   0
      Left            =   1455
      TabIndex        =   57
      Tag             =   "T"
      Top             =   10395
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox TXT_INSP_PART 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   0
      Left            =   1515
      TabIndex        =   56
      Text            =   " "
      Top             =   9990
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TXT_INSP_PART 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   1
      Left            =   2490
      TabIndex        =   55
      Text            =   " "
      Top             =   9990
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TXT_INSP_PART 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   2
      Left            =   3450
      TabIndex        =   54
      Text            =   " "
      Top             =   9990
      Visible         =   0   'False
      Width           =   945
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   4065
      Left            =   90
      TabIndex        =   39
      Top             =   45
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   7170
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      Locked          =   -1  'True
      PaneTree        =   "AGC2020C.frx":0000
      Begin Threed.SSFrame Single 
         Height          =   915
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   1614
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
         Begin VB.ComboBox CBO_PLT 
            Height          =   315
            ItemData        =   "AGC2020C.frx":0052
            Left            =   14340
            List            =   "AGC2020C.frx":005C
            TabIndex        =   46
            Text            =   "C1"
            Top             =   150
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox TXT_PLATE_NO 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1435
            MaxLength       =   14
            TabIndex        =   45
            Top             =   90
            Width           =   2070
         End
         Begin VB.ComboBox CBO_SHIFT 
            Height          =   315
            ItemData        =   "AGC2020C.frx":0068
            Left            =   10035
            List            =   "AGC2020C.frx":0075
            TabIndex        =   44
            Top             =   90
            Width           =   1005
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
            Left            =   5265
            MaxLength       =   18
            TabIndex        =   43
            Tag             =   "标准号"
            Top             =   510
            Width           =   3225
         End
         Begin VB.TextBox text_cur_inv 
            Height          =   315
            Left            =   1950
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   510
            Width           =   1560
         End
         Begin VB.TextBox text_cur_inv_code 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1435
            MaxLength       =   2
            TabIndex        =   41
            Tag             =   "起始库"
            Top             =   510
            Width           =   495
         End
         Begin InDate.ULabel ULabel16 
            Height          =   345
            Left            =   225
            Top             =   90
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   609
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
            Left            =   3960
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
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
         Begin InDate.ULabel ULabel13 
            Height          =   315
            Left            =   8940
            Top             =   90
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            Caption         =   "班次"
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
            Index           =   4
            Left            =   3960
            Top             =   510
            Width           =   1275
            _ExtentX        =   2249
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
         Begin InDate.ULabel ULabel23 
            Height          =   315
            Left            =   8940
            Top             =   510
            Width           =   1065
            _ExtentX        =   1879
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
         Begin CSTextLibCtl.sidbEdit SDB_THK_REF 
            Height          =   315
            Left            =   10035
            TabIndex        =   47
            Top             =   510
            Width           =   1005
            _Version        =   262145
            _ExtentX        =   1773
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
            Left            =   11340
            Top             =   510
            Width           =   1065
            _ExtentX        =   1879
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
         Begin CSTextLibCtl.sidbEdit SDB_WID_REF 
            Height          =   315
            Left            =   12435
            TabIndex        =   48
            Top             =   510
            Width           =   1185
            _Version        =   262145
            _ExtentX        =   2090
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
         Begin InDate.ULabel ULabel25 
            Height          =   315
            Left            =   225
            Top             =   510
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "当前库"
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
         Begin InDate.UDate SDT_PROD_DATE 
            Height          =   315
            Left            =   5265
            TabIndex        =   50
            Tag             =   "起始日期"
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
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
         End
         Begin InDate.UDate SDT_PROD_TO_DATE 
            Height          =   315
            Left            =   7035
            TabIndex        =   51
            Tag             =   "起始日期"
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
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
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "~"
            Height          =   120
            Left            =   6810
            TabIndex        =   49
            Top             =   240
            Width           =   195
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   3090
         Left            =   0
         TabIndex        =   114
         Top             =   975
         Width           =   15285
         _Version        =   393216
         _ExtentX        =   26961
         _ExtentY        =   5450
         _StockProps     =   64
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   28
         MaxRows         =   5
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AGC2020C.frx":0082
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1050
      Left            =   5760
      TabIndex        =   30
      Top             =   1380
      Width           =   3285
      _ExtentX        =   5794
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   645
         Top             =   150
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "钢板分板数"
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
      Begin Threed.SSCommand cmd_divide 
         Height          =   360
         Left            =   660
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   540
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   196609
         Font3D          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "分板"
      End
      Begin CSTextLibCtl.sidbEdit SDB_DIVIDE_CNT 
         Height          =   315
         Left            =   1995
         TabIndex        =   32
         Top             =   150
         Width           =   480
         _Version        =   262145
         _ExtentX        =   847
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
         RawData         =   ""
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
         NumDecDigits    =   0
         NumIntDigits    =   1
         ShowZero        =   0   'False
         MaxValue        =   9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin Threed.SSCommand cmd_divide_ok 
         Height          =   360
         Left            =   2910
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   150
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   635
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "确认分板"
      End
      Begin Threed.SSCommand cmd_divide_delete 
         Height          =   360
         Left            =   2910
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   540
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   635
         _Version        =   196609
         Font3D          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "取消分板"
      End
   End
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   300
      Top             =   9990
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Caption         =   "下缺陷部位"
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   5130
      Top             =   10020
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Caption         =   "上缺陷部位"
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
      Left            =   11505
      TabIndex        =   78
      Top             =   10110
      Visible         =   0   'False
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
      Left            =   12465
      TabIndex        =   79
      Top             =   10110
      Visible         =   0   'False
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
      Left            =   13440
      TabIndex        =   80
      Top             =   10110
      Visible         =   0   'False
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   10320
      Top             =   10110
      Visible         =   0   'False
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
      Index           =   0
      Left            =   11565
      TabIndex        =   81
      Top             =   10560
      Visible         =   0   'False
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
      Left            =   12525
      TabIndex        =   82
      Top             =   10560
      Visible         =   0   'False
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
      Left            =   13500
      TabIndex        =   83
      Top             =   10560
      Visible         =   0   'False
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Left            =   10380
      Top             =   10560
      Visible         =   0   'False
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
   Begin Threed.SSFrame sf1 
      Height          =   2865
      Left            =   90
      TabIndex        =   84
      Top             =   4170
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   5054
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
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   0
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   615
         Width           =   2385
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   1
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   975
         Width           =   2385
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   5
         Left            =   1350
         TabIndex        =   93
         Top             =   2400
         Width           =   2385
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   4
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   2040
         Width           =   2385
      End
      Begin VB.TextBox TXT_INSP_FLAW_NAME 
         Height          =   315
         Index           =   3
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   1680
         Width           =   2385
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   3750
         MaxLength       =   3
         TabIndex        =   90
         Top             =   615
         Width           =   885
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   3750
         MaxLength       =   3
         TabIndex        =   89
         Top             =   1680
         Width           =   885
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   4
         Left            =   3750
         MaxLength       =   3
         TabIndex        =   88
         Top             =   2040
         Width           =   885
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   5
         Left            =   3750
         MaxLength       =   3
         TabIndex        =   87
         Top             =   2400
         Width           =   885
      End
      Begin VB.TextBox TXT_STLGRD 
         Height          =   285
         Left            =   1560
         TabIndex        =   86
         Top             =   4110
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TXT_INSP_FLAW 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   3750
         MaxLength       =   3
         TabIndex        =   85
         Top             =   975
         Width           =   885
      End
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   150
         Top             =   615
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
         Left            =   150
         Top             =   975
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
            Size            =   9.76
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
         Left            =   1350
         Top             =   285
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
         Index           =   1
         Left            =   1350
         Top             =   1350
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   150
         Top             =   1680
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   150
         Top             =   2040
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
         Left            =   150
         Top             =   2400
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "探伤缺陷"
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
   Begin Threed.SSFrame sf3 
      Height          =   2895
      Left            =   4920
      TabIndex        =   96
      Top             =   4170
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   5106
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
      Begin VB.TextBox TXT_UNIT 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   4530
         MaxLength       =   2
         TabIndex        =   135
         Tag             =   "原因"
         Top             =   2400
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox TXT_INSP_WID_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   134
         Top             =   2055
         Width           =   1050
      End
      Begin VB.TextBox TXT_SIZE_KND_NAME 
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   133
         Tag             =   "钢种"
         Top             =   2420
         Width           =   1050
      End
      Begin VB.TextBox TXT_SIZE_KND 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1380
         MaxLength       =   2
         TabIndex        =   115
         Tag             =   "原因"
         Top             =   2420
         Width           =   870
      End
      Begin VB.TextBox TXT_INSP_LEN_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   2055
         Width           =   1125
      End
      Begin VB.TextBox TXT_INSP_THK_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   2055
         Width           =   870
      End
      Begin VB.TextBox TXT_INSP_WGT_GRD 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4515
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   2055
         Width           =   975
      End
      Begin InDate.ULabel ULabel28 
         Height          =   315
         Left            =   2280
         Top             =   285
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
         Left            =   1380
         Top             =   285
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
         Left            =   3360
         Top             =   285
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
      Begin InDate.ULabel ULabel33 
         Height          =   315
         Left            =   210
         Top             =   2055
         Width           =   1140
         _ExtentX        =   2011
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
         Left            =   4515
         TabIndex        =   100
         Top             =   975
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
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
         Left            =   4515
         TabIndex        =   101
         Top             =   615
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_LEN_MX 
         Height          =   315
         Left            =   3360
         TabIndex        =   102
         Top             =   1335
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
         Left            =   2280
         TabIndex        =   103
         Top             =   1695
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
         Left            =   1380
         TabIndex        =   104
         Top             =   1695
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
         Left            =   3360
         TabIndex        =   105
         Top             =   1695
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
      Begin CSTextLibCtl.sidbEdit SDB_PWGT_MN 
         Height          =   315
         Left            =   4515
         TabIndex        =   106
         Top             =   1695
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
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
         Left            =   2280
         TabIndex        =   107
         Top             =   615
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
         Left            =   1380
         TabIndex        =   108
         Top             =   615
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
         Left            =   3360
         TabIndex        =   109
         Top             =   615
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
         Left            =   210
         Top             =   1695
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
         Left            =   210
         Top             =   615
         Width           =   1140
         _ExtentX        =   2011
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
      Begin CSTextLibCtl.sidbEdit SDB_PWGT_MX 
         Height          =   315
         Left            =   4515
         TabIndex        =   110
         Top             =   1335
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
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
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   210
         Top             =   1335
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
      Begin InDate.ULabel ULabel44 
         Height          =   315
         Left            =   4515
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
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
         Left            =   2280
         TabIndex        =   111
         Top             =   975
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
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ORD_THK 
         Height          =   315
         Left            =   1380
         TabIndex        =   112
         Top             =   975
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
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ORD_LEN 
         Height          =   315
         Left            =   3360
         TabIndex        =   113
         Top             =   975
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
         Top             =   975
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
         Left            =   210
         Top             =   2420
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         Caption         =   "定尺"
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_THK_MX 
         Height          =   315
         Left            =   1380
         TabIndex        =   131
         Top             =   1335
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
      Begin CSTextLibCtl.sidbEdit SDB_INSP_WID_MX 
         Height          =   315
         Left            =   2280
         TabIndex        =   132
         Top             =   1335
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
      Begin InDate.ULabel ULabel39 
         Height          =   345
         Left            =   3360
         Top             =   2400
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   609
         Caption         =   "不平度"
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
End
Attribute VB_Name = "AGC2020C"
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
Const SS1_URGNT_FL = 27


Private Sub Form_Define()
    Dim iIndex As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(CBO_PLT, "p", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(SDT_PROD_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDT_PROD_TO_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_stdspec_chg_ref, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_THK_REF, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_WID_REF, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_UST_FLAG, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_PROC_FLAG, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(TXT_APLY_ENDUSE_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
      Call Gp_Ms_Collection(SDB_INSP_THK_MX, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_THK_MN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_WID, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_WID_MX, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_WID_MN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_LEN, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LEN_MX, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_INSP_LEN_MN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_WGT_ORD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_WGT, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_PWGT_MX, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_PWGT_MN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
     Call Gp_Ms_Collection(TXT_INSP_THK_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_WID_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_LEN_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_INSP_WGT_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_SURF_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_INSP_MAIN_GRD, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_NEXT_PROC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_INSP_MAN, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(TXT_INSP_OCCR_TIME, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_ORD_WID, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_ORD_THK, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_ORD_LEN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_GRID_EMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_GRID_TIME, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_TOP_GRID_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_TOP_GRID_YRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_TOP_GRID_DEEP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_BOT_GRID_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDB_BOT_GRID_YRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_BOT_GRID_DEEP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_stdspec, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_stdspec_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_stdspec_chg, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_stdspec_name_chg, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Scrap_code, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Scrap_name, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(SDB_Mn, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_PROC_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_SIZE_KND, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_SPEC_PROC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_REMARK, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)  '增加备注录入，只显示录入生产备注  20110425  ADD BY GUHF
              Call Gp_Ms_Collection(TXT_GAS, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)  '火切指示 20110425  ADD BY GUHF
               Call Gp_Ms_Collection(TXT_CL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)  '矫直指示 20110425  ADD BY GUHF
          Call Gp_Ms_Collection(txt_woo_rsn, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)  '余材代码 20110608  ADD BY LiQian 如有发生降余材，提示必须输入余材代码，将余材代码做记录
             Call Gp_Ms_Collection(TXT_UNIT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_ORD_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
    For iIndex = 0 To 17
        Call Gp_Clear_Collection(CHK_PART(iIndex), "s", sControl)
    Next iIndex
    
     Call Gp_Clear_Collection(CHK_SUR_GRD(0), "s", sControl)
     Call Gp_Clear_Collection(CHK_SUR_GRD(1), "s", sControl)
     Call Gp_Clear_Collection(CHK_PRD_GRD(0), "s", sControl)
     Call Gp_Clear_Collection(CHK_PRD_GRD(1), "s", sControl)
     Call Gp_Clear_Collection(CHK_PRD_GRD(2), "s", sControl)
     Call Gp_Clear_Collection(CHK_PRD_GRD(3), "s", sControl)
     Call Gp_Clear_Collection(CHK_PRD_GRD(4), "s", sControl)
     Call Gp_Clear_Collection(CHK_TOP_GRD(0), "s", sControl)
     Call Gp_Clear_Collection(CHK_TOP_GRD(1), "s", sControl)
     Call Gp_Clear_Collection(CHK_BOT_GRD(0), "s", sControl)
     Call Gp_Clear_Collection(CHK_BOT_GRD(1), "s", sControl)
     Call Gp_Clear_Collection(CHK_GAS_FL, "s", sControl)  '火切指示  20110425  ADD BY GUHF
     Call Gp_Clear_Collection(CHK_CL_FL, "s", sControl)   '矫直指示  20110425  ADD BY GUHF
    
    MC.Add Item:=sControl, Key:="sControl"
    
    'MASTER Collection
    Mc1.Add Item:="AGC2020C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="AGC2020C.P_REFER", Key:="P-R"
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
      
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2020C.P_SREFER", Key:="P-R"
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

Private Sub CHK_CL_FL_Click()
'矫直指示   20110425  ADD BY GUHF
Dim V_CL As String


V_CL = Right(Trim(TXT_CL.Text), 2)
If CHK_CL_FL.Value = ssCBChecked Then
   TXT_CL.Text = "Y" + V_CL
Else
   TXT_CL.Text = "N" + V_CL
End If

V_CL = ""

End Sub

Private Sub CHK_GAS_FL_Click()
 '火切指示  20110425  ADD BY GUHF
 Dim V_GAS As String
 
V_GAS = Right(TXT_GAS.Text, 2)
If CHK_GAS_FL.Value = ssCBChecked Then
   TXT_GAS.Text = "Y" + V_GAS
ElseIf CHK_GAS_FL.Value = ssCBUnchecked Then
   TXT_GAS.Text = "N" + V_GAS
End If

V_GAS = ""
  
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

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    
    CBO_PLT.ListIndex = 0

    If App.Title = "BG" Then
        text_cur_inv_code = "00"
    ElseIf App.Title = "DG" Then
        text_cur_inv_code = "SG"
    End If

    Call text_cur_inv_code_KeyUp(0, 0)
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)

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
        
        CHK_GAS_FL.Enabled = False
        CHK_CL_FL.Enabled = False
        TXT_GAS.Text = ""      ' 清空火切 20110425  ADD BY GUHF
        TXT_CL.Text = ""       ' 清空矫直
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)

        TXT_INSP_MAN = sUserID
        CBO_PLT.ListIndex = 0
        SDB_DIVIDE_CNT.Value = 0
        
        
        For iCount = 0 To 5
            TXT_INSP_FLAW_NAME(iCount).Text = ""
        Next iCount
        
        ss1.BlockMode = True
        ss1.Row = -1
        ss1.Col = -1
        ss1.BackColor = &HFFFFFF
        ss1.BlockMode = False
    End If
    
    If App.Title = "BG" Then
        text_cur_inv_code = "00"
'    ElseIf App.Title = "DG" Then
'        text_cur_inv_code = "WD"
    End If

    Call text_cur_inv_code_KeyUp(0, 0)
End Sub

Public Sub Form_Ref()

    Dim iCount As Integer
    
    Call Form_Cls
    Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1)
    ss1.OperationMode = OperationModeNormal
    
    If Len(TXT_PLATE_NO.Text) = 14 Then
        If Gf_Ms_Refer(M_CN1, Mc1, , , True) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Display_Data_Edit
        End If
    End If
    
    With ss1
    '紧急订单绿色标记 2012-11-08  BY  LICHAO
    
        For iCount = 1 To .MaxRows
            .Row = iCount
           ss1.Row = .Row:       ss1.Col = SS1_URGNT_FL
           If ss1.Text = "Y" Then
                Call Gp_Sp_BlockColor(ss1, SS1_PLATE_NO, SS1_PLATE_NO, .Row, .Row, &HC000&)
                Call Gp_Sp_BlockColor(ss1, SS1_URGNT_FL, SS1_URGNT_FL, .Row, .Row, &HC000&)
           End If
        Next iCount
    End With
     
End Sub

Public Sub Form_Pro()

    Dim SMESG   As String
    Dim iCount  As Integer
    
'    For iCount = 0 To 5
'        If TXT_INSP_FLAW_NAME(iCount).Text <> "" And TXT_INSP_PART(iCount).Text = "" Then
'            SMESG = " 请输入缺陷部位 ！"
'            Call Gp_MsgBoxDisplay(SMESG)
'            Exit Sub
'        End If
'    Next iCount
    
    If TXT_NEXT_PROC.Text = "Y" Then
        SMESG = " 请输入最终等级判定 ！"
        If Not Gf_MessConfirm(SMESG, "Q") Then
           Exit Sub
        End If
    End If
    
    If TXT_NEXT_PROC.Text = "P" And Trim(TXT_INSP_MAIN_GRD.Text) = "" Then
        SMESG = " 请输入最终等级判定 ！"
        Call Gp_MsgBoxDisplay(SMESG)
        Exit Sub
    End If
    
    If Trim(TXT_INSP_MAIN_GRD.Text) <> "4" Then
        If Trim(TXT_SURF_GRD.Text) = "" Then
            SMESG = " 请输入表面判定 ！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
        End If
        
'        If Trim(txt_NEXT_PROC.Text) = "" Then
'            sMesg = " 请输入后道工序 ！"
'            Call Gp_MsgBoxDisplay(sMesg)
'            Exit Sub
'        End If
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
    
    If CHK_PRD_GRD(4).Value = ssCBChecked Then
        If txt_Scrap_code.Text = "" Then
            SMESG = " 请正确输入废钢原因 ！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
        End If
    End If
            
    If TXT_INSP_FLAW(2).Text <> "" And TXT_INSP_FLAW_NAME(2).Text = "" Then
        SMESG = " 请输入正确的改判原因 ！"
            Call Gp_MsgBoxDisplay(SMESG)
            Exit Sub
    End If

'    TXT_GAS.Text = Mid(TXT_GAS.Text, 1, 1)
'    TXT_CL.Text = Mid(TXT_CL.Text, 1, 1)
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        TXT_INSP_MAN.Text = sUserID
       If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
       TXT_PLATE_NO.Enabled = True
'       TXT_PLATE_NO.Text = ""
'       Call Form_Ref
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

Private Sub SDT_PROD_DATE_GotFocus()
     If SDT_PROD_DATE.RawData = "" Then
        SDT_PROD_DATE.RawData = Gf_DTSet(M_CN1, "D")
     End If
     If SDT_PROD_TO_DATE.RawData = "" Then
        SDT_PROD_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub SDT_PROD_TO_DATE_GotFocus()
     If SDT_PROD_TO_DATE.RawData = "" Then
        SDT_PROD_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub text_cur_inv_code_DblClick()
    'Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
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
    TXT_GRID_TIME.RawData = Gf_DTSet(M_CN1)
End Sub

Private Sub TXT_INSP_FLAW_Change(Index As Integer)
    If Len(Trim(TXT_INSP_FLAW(Index).Text)) = 3 Then
        TXT_INSP_FLAW_NAME(Index).Text = Gf_ComnNameFind(M_CN1, "G0002", Trim(TXT_INSP_FLAW(Index).Text), 1)
        TXT_INSP_FLAW(Index).Text = UCase(TXT_INSP_FLAW(Index).Text)
    Else
        TXT_INSP_FLAW_NAME(Index).Text = ""
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

Private Sub TXT_INSP_MAN_DblClick()
    TXT_INSP_MAN.Text = sUserID
End Sub

Private Sub TXT_INSP_OCCR_TIME_DblClick()
    TXT_INSP_OCCR_TIME.RawData = Gf_DTSet(M_CN1)
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
    
    If CHK_SUR_GRD(Index).Value = ssCBUnchecked Then
        If CHK_SUR_GRD(iNext).Value = ssCBUnchecked Then
            TXT_SURF_GRD.Text = ""
            CHK_SUR_GRD(Index).ForeColor = &H808080
            sCheck = ""
            Exit Sub
        End If
    Else
        CHK_SUR_GRD(iNext).Value = ssCBUnchecked
    End If
    
    CHK_SUR_GRD(Index).ForeColor = &HFF&
    CHK_SUR_GRD(Index).Value = ssCBChecked
                
    CHK_SUR_GRD(iNext).ForeColor = &H808080
    CHK_SUR_GRD(iNext).Value = ssCBUnchecked

    TXT_SURF_GRD.Text = CHK_SUR_GRD(Index).Tag
    sCheck = ""
    
End Sub

Private Sub CHK_PRD_GRD_Click(Index As Integer)
    Dim iCount      As Integer
    Dim iIndexStr   As Integer
    
    If sCheck <> "" Then Exit Sub

    iCount = 0
    sCheck = "**"
            
    If CHK_PRD_GRD(Index).Value = ssCBUnchecked Then
        For iIndexStr = 0 To 5
            If CHK_PRD_GRD(iIndexStr).Value = ssCBChecked Then
               iCount = iCount + 1
            End If
        Next iIndexStr
        If iCount = 0 Then
            TXT_INSP_MAIN_GRD.Text = ""
            CHK_PRD_GRD(Index).ForeColor = &H808080
            sCheck = ""
            Exit Sub
        End If
    Else
        For iIndexStr = 0 To 5
            CHK_PRD_GRD(iIndexStr).ForeColor = &H808080
            CHK_PRD_GRD(iIndexStr).Value = ssCBUnchecked
        Next iIndexStr
    End If
    
    CHK_PRD_GRD(Index).ForeColor = &HFF&
    CHK_PRD_GRD(Index).Value = ssCBChecked
    
    TXT_INSP_MAIN_GRD.Text = CHK_PRD_GRD(Index).Tag
                 
    txt_stdspec_chg.Text = ""
    txt_stdspec_name_chg.Text = ""
    If CHK_PRD_GRD(0).Value = ssCBChecked Or CHK_PRD_GRD(1).Value = ssCBChecked Or CHK_PRD_GRD(2).Value = ssCBChecked Or CHK_PRD_GRD(3).Value = ssCBChecked Or CHK_PRD_GRD(5).Value = ssCBChecked Then
        txt_stdspec_chg.Enabled = True
    Else
        txt_stdspec_chg.Enabled = False
    End If
         
    If CHK_PRD_GRD(4).Value = ssCBChecked Then
        txt_Scrap_code.Enabled = True
    Else
        txt_Scrap_code.Enabled = False
    End If
    
'   MODEFIED BY YANGMENG AT 07.01.30
'    待判时处理
    sCheck = "**"
    
    For iIndexStr = 0 To 1
        If CHK_PRD_GRD(5).Value = ssCBChecked Then
            CHK_SUR_GRD(iIndexStr).Enabled = False
            TXT_SURF_GRD.Text = ""
            CHK_SUR_GRD(iIndexStr).Value = ssCBUnchecked
            CHK_SUR_GRD(iIndexStr).ForeColor = &H808080
        Else
            CHK_SUR_GRD(iIndexStr).Enabled = True
        End If
    Next iIndexStr
        
    sCheck = ""
        
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
        TXT_GRID_TIME.RawData = Gf_DTSet(M_CN1)
        
        CHK_TOP_GRD(0).Value = ssCBChecked
        Call CHK_TOP_GRD_Click(0)
        CHK_BOT_GRD(0).Value = ssCBChecked
        Call CHK_BOT_GRD_Click(0)

    End If
End Sub

Private Sub Display_Data_Edit()
    Dim iIndexChk   As Integer
    Dim iIndexStr   As Integer
    Dim V_GAS_FL    As String
    Dim V_CL_FL     As String
        
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
        
    For iIndexChk = 0 To 1
        If TXT_SURF_GRD.Text = CHK_SUR_GRD(iIndexChk).Tag Then
            CHK_SUR_GRD(iIndexChk).ForeColor = &HFF&
            CHK_SUR_GRD(iIndexChk).Value = ssCBChecked
        Else
            CHK_SUR_GRD(iIndexChk).ForeColor = &H808080
            CHK_SUR_GRD(iIndexChk).Value = ssCBUnchecked
        End If
    Next iIndexChk

    For iIndexChk = 0 To 5
        If TXT_INSP_MAIN_GRD.Text = CHK_PRD_GRD(iIndexChk).Tag Then
            CHK_PRD_GRD(iIndexChk).ForeColor = &HFF&
            CHK_PRD_GRD(iIndexChk).Value = ssCBChecked
        Else
            CHK_PRD_GRD(iIndexChk).ForeColor = &H808080
            CHK_PRD_GRD(iIndexChk).Value = ssCBUnchecked
        End If
    Next iIndexChk
    
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
    
  ' 切割矫直指示 20110425  ADD BY GUHF
    V_GAS_FL = Mid(TXT_GAS, 1, 1)
    V_CL_FL = Mid(TXT_CL, 1, 1)
    
    If Right(TXT_GAS.Text, 1) = "Y" Then  '有切割实绩则不允许做切割指示变更
    CHK_GAS_FL.Enabled = False
    Else
    CHK_GAS_FL.Enabled = True
    End If
    If Right(TXT_CL.Text, 1) = "Y" Then   '有矫直实绩则不允许做矫直指示变更
    CHK_CL_FL.Enabled = False
    Else
    CHK_CL_FL.Enabled = True
    End If

    
    If V_GAS_FL = "Y" Then
    CHK_GAS_FL.Value = ssCBChecked
   
    End If
    
    If V_CL_FL = "Y" Then
    CHK_CL_FL.Value = ssCBChecked
    End If
    
    V_GAS_FL = ""
    V_CL_FL = ""
    
    '20110425  ADD BY GUHF
    
    sCheck = ""
    
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
    
    If TXT_INSP_MAIN_GRD.Text = "4" Then
        CHK_PRD_GRD(5).Value = ssCBChecked
        Call CHK_PRD_GRD_Click(5)
    Else
        CHK_SUR_GRD(0).Value = ssCBChecked
        Call CHK_SUR_GRD_Click(0)
        
        If sGradeFlag = "N" Then
            CHK_PRD_GRD(1).Value = ssCBChecked
            Call CHK_PRD_GRD_Click(1)
    '        CHK_PRD_GRD(0).Enabled = False
        Else
            CHK_PRD_GRD(0).Value = ssCBChecked
            Call CHK_PRD_GRD_Click(0)
    '        CHK_PRD_GRD(0).Enabled = True
        End If
        
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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Or SDB_DIVIDE_CNT.Value > 0 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 1
    TXT_PLATE_NO.Text = ss1.Text
    
    CHK_GRID_FLAG.Value = ssCBUnchecked
  
    
    If Len(TXT_PLATE_NO.Text) = 14 Then
        Call Gp_SSCheck_Cls(MC("sControl"))
        If Gf_Ms_Refer(M_CN1, Mc1, , , True) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Display_Data_Edit
        End If
    End If
    
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

Private Sub TXT_SPEC_PROC_Change()
    If Len(Trim(TXT_SPEC_PROC.Text)) = TXT_SPEC_PROC.MaxLength Then
        TXT_SPEC_PROC_NAME.Text = Gf_ComnNameFind(M_CN1, "G0046", TXT_SPEC_PROC.Text, 2)
    Else
        TXT_SPEC_PROC_NAME.Text = ""
    End If
End Sub

Private Sub TXT_SPEC_PROC_DblClick()
    Call TXT_SPEC_PROC_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_SPEC_PROC_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sSpec_proc As String
    sSpec_proc = TXT_SPEC_PROC.Text

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0046"

        DD.rControl.Add Item:=TXT_SPEC_PROC

        DD.nameType = "2"
        TXT_SPEC_PROC.Text = ""
        Call Gf_Common_DD(M_CN1, KeyCode)
        If TXT_SPEC_PROC.Text = "" Then
            TXT_SPEC_PROC.Text = sSpec_proc
        End If
        
    End If
End Sub

Private Sub txt_stdspec_chg_DblClick()
'    Call txt_stdspec_chg_KeyUp(vbKeyF4, 0)

         DD.sWitch = "MS"
         DD.DataDicType = "C"
         DD.rControl.Add Item:=txt_stdspec_chg
        
         Call Pf_Common_DD(M_CN1, vbKeyF4)
         
         Exit Sub
         
End Sub

Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        txt_stdspec_yy.Text = ""
        DD.rControl.Add Item:=txt_stdspec_chg
        DD.rControl.Add Item:=txt_stdspec_yy

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

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


Private Sub txt_woo_rsn_DblClick()

    Call txt_woo_rsn_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_woo_rsn_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0008"
        DD.rControl.Add Item:=txt_woo_rsn
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

End Sub
