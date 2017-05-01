VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB1026C 
   Caption         =   "板坯转库实绩录入_ACB1026C"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   900
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   14940
   WindowState     =   2  'Maximized
   Begin VB.TextBox text_to_cur_inv_code 
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
      Left            =   6495
      MaxLength       =   2
      TabIndex        =   19
      Top             =   615
      Width           =   465
   End
   Begin VB.TextBox text_to_cur_inv 
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
      Height          =   315
      Left            =   6975
      TabIndex        =   18
      Top             =   615
      Width           =   1455
   End
   Begin VB.TextBox text_fr_cur_inv_code 
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
      Left            =   6495
      MaxLength       =   2
      TabIndex        =   17
      Top             =   180
      Width           =   465
   End
   Begin VB.TextBox text_fr_cur_inv 
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
      Height          =   315
      Left            =   6975
      TabIndex        =   16
      Top             =   180
      Width           =   1455
   End
   Begin VB.TextBox Text_PROD_CD 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1230
      MaxLength       =   2
      TabIndex        =   15
      Tag             =   "产品"
      Text            =   "SL"
      Top             =   180
      Width           =   375
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   225
      Top             =   180
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "产品"
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
   Begin VB.TextBox txt_emp 
      Height          =   255
      Left            =   13755
      TabIndex        =   13
      Top             =   1065
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt_MOVE_CAR_NO 
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
      Left            =   10875
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "机号"
      Top             =   615
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txt_MOVE_SHEET_NO 
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
      Left            =   13710
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "机号"
      Top             =   615
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox TXT_PRC_LINTTO 
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
      Left            =   13710
      MaxLength       =   1
      TabIndex        =   8
      Tag             =   "机号"
      Text            =   "1"
      Top             =   180
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt_plt_to 
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
      Left            =   10875
      MaxLength       =   2
      TabIndex        =   7
      Tag             =   "工 厂"
      Top             =   180
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox TXT_PRC_STS 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4875
      MaxLength       =   11
      TabIndex        =   3
      Top             =   615
      Width           =   465
   End
   Begin VB.TextBox txt_PLT_NAME 
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
      Left            =   1980
      TabIndex        =   1
      Tag             =   "工 厂"
      Top             =   180
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.TextBox txt_PLT 
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
      Left            =   1500
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "工 厂"
      Text            =   "C1"
      Top             =   180
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt_PRC_line 
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
      Left            =   4890
      MaxLength       =   1
      TabIndex        =   2
      Tag             =   "机号"
      Text            =   "1"
      Top             =   180
      Width           =   465
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   3960
      Top             =   1035
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "移送指示日期"
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
   Begin InDate.UDate dtp_ins_date_PROD_DATE1 
      Height          =   315
      Left            =   5235
      TabIndex        =   4
      Tag             =   "INS_DATE"
      Top             =   1035
      Width           =   1515
      _ExtentX        =   2672
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
   Begin InDate.UDate dtp_ins_date_PROD_DATE2 
      Height          =   315
      Left            =   6960
      TabIndex        =   5
      Tag             =   "INS_DATE"
      Top             =   1035
      Width           =   1470
      _ExtentX        =   2593
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   225
      Tag             =   "移 送 工 厂"
      Top             =   180
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "来源工厂"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   3960
      Top             =   180
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      Caption         =   "来源机号"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   3960
      Top             =   615
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      Caption         =   "状态"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   7770
      Left            =   225
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1440
      Width           =   14850
      _Version        =   393216
      _ExtentX        =   26194
      _ExtentY        =   13705
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
      MaxCols         =   24
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACB1026C.frx":0000
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   9600
      Tag             =   "移 送 工 厂"
      Top             =   180
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "目的库"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   12435
      Top             =   180
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "目的机号"
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
   Begin InDate.ULabel ULabel55 
      Height          =   315
      Left            =   9600
      Top             =   615
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "转库车辆号"
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
   Begin InDate.ULabel ULabel56 
      Height          =   315
      Left            =   12435
      Top             =   615
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "转库提货单号"
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
      Left            =   225
      Top             =   1035
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "已选择数量"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_num 
      Height          =   315
      Left            =   1410
      TabIndex        =   11
      Top             =   1050
      Width           =   1350
      _Version        =   262145
      _ExtentX        =   2381
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_wgt 
      Height          =   315
      Left            =   1410
      TabIndex        =   12
      Top             =   615
      Width           =   1350
      _Version        =   262145
      _ExtentX        =   2381
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
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
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.000"
      Text            =   " 0.000"
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
      NumIntDigits    =   7
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   225
      Top             =   615
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "已选择总重量"
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   9600
      Top             =   1035
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "执行时间"
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
   Begin CSTextLibCtl.sitxEdit txt_MOVE_ts 
      Height          =   315
      Left            =   10875
      TabIndex        =   14
      Top             =   1035
      Visible         =   0   'False
      Width           =   2085
      _Version        =   262145
      _ExtentX        =   3678
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__ __:__:__"
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
   End
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Left            =   5415
      Top             =   180
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "来源仓库"
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
      Left            =   5415
      Top             =   615
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "目标仓库"
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
   Begin VB.Line Line1 
      X1              =   6795
      X2              =   6915
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line2 
      X1              =   8070
      X2              =   8190
      Y1              =   915
      Y2              =   915
   End
End
Attribute VB_Name = "ACB1026C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       MOVE ANOTHER FACTORY
'-- Sub_System Name
'-- Program Name
'-- Program ID        ACB1026C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             ZHENG WEN
'-- Date              2004.8.26
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
Dim sQuery As String

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
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iCount As Integer
Dim iCol As Integer
Dim iRow As Integer


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       
                'Call Gp_Ms_Collection(txt_PLT, "p", "n", " ", "i", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(text_prod_cd, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(txt_PRC_LINE, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   Call Gp_Ms_Collection(TXT_PRC_STS, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(dtp_ins_date_prod_date1, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(dtp_ins_date_prod_date2, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(text_fr_cur_inv_code, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(text_to_cur_inv_code, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

       
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
       
 
    'Spread_Collection
        'Call Spread_Collection("Column1_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")

     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB1026C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
        
    Call Gp_Sp_ColColor(ss1, 1)
    Call Gp_Sp_ColColor(ss1, 2)
    Call Gp_Sp_ColColor(ss1, 3)
    Call Gp_Sp_ColColor(ss1, 5)
    Call Gp_Sp_ColColor(ss1, 12)
    Call Gp_Sp_ColColor(ss1, 14)
    Call Gp_Sp_ColColor(ss1, 19)
    Call Gp_Sp_ColColor(ss1, 20)
    Call Gp_Sp_ColColor(ss1, 21)
    Call Gp_Sp_ColColor(ss1, 22)
    Call Gp_Sp_ColColor(ss1, 23)
    
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
'    Call Gp_Sp_ColHidden(ss1, 20, True)
'    Call Gp_Sp_ColHidden(ss1, 21, True)
'    Call Gp_Sp_ColHidden(ss1, 22, True)
'
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    text_fr_cur_inv_code.Text = "00"
End Sub


Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
      '  Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
      '  Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 4)
    End If
    
End Sub


Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub
Private Sub Form_Activate()
    
    Call FormMenuSetting1(Me, FormType, Toolbar_St, sAuthority)
  
   
End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)

'If ss1.ActiveCol = 15 Then
'
'    If KeyCode = vbKeyF4 Then
'
''        txt_f_addr.Text = "S"
'        DD.sWitch = "MS"
'        DD.sKey = "F0009"
'        DD.rControl.Add Item:=txt_f_addr
'
'
'        DD.nameType = "2"
'
'        Call Gf_Common_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'Else
'    Exit Sub
'End If



'--------------------------------------------------------------------------------

 Dim sTemp_Code As String

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    If ss1.ActiveCol = 15 Then

        If KeyCode = vbKeyF4 Then
            ss1.Row = ss1.ActiveRow
            Set DD.sPname = Me.ss1
            
            DD.sWitch = "SP"
            DD.sKey = "F0009"
            DD.rControl.Add Item:=15
            
            DD.nameType = "2"
            Call Gf_Common_DD(M_CN1, KeyCode)
            
        End If
                
    End If

End Sub

Private Sub text_fr_cur_inv_code_Change()
    If Len(Trim(text_fr_cur_inv_code.Text)) = text_fr_cur_inv_code.MaxLength Then
          text_fr_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_fr_cur_inv_code.Text, 2)
          Exit Sub
    Else
          text_fr_cur_inv.Text = ""
    End If
End Sub

Private Sub text_fr_cur_inv_code_DblClick()
    Call text_fr_cur_inv_code_KeyUp(vbKeyF4, 0)
End Sub

Private Sub text_fr_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=text_fr_cur_inv_code
        DD.rControl.Add Item:=text_fr_cur_inv
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
       
        If Len(Trim(text_fr_cur_inv_code.Text)) = text_fr_cur_inv_code.MaxLength Then
            text_fr_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_fr_cur_inv_code.Text, 2)
            Exit Sub
        Else
            text_fr_cur_inv.Text = ""
        End If
    End If
End Sub

Private Sub Text_PROD_CD_DblClick()

    Call Text_PROD_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_to_cur_inv_code_Change()
    If Len(Trim(text_to_cur_inv_code.Text)) = text_to_cur_inv_code.MaxLength Then
        text_to_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_to_cur_inv_code.Text, 2)
        Exit Sub
    Else
        text_to_cur_inv.Text = ""
    End If
End Sub

Private Sub text_to_cur_inv_code_DblClick()
    Call text_to_cur_inv_code_KeyUp(vbKeyF4, 0)
End Sub

Private Sub text_to_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=text_to_cur_inv_code
        DD.rControl.Add Item:=text_to_cur_inv
        

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
       
        If Len(Trim(text_to_cur_inv_code.Text)) = text_to_cur_inv_code.MaxLength Then
            text_to_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_to_cur_inv_code.Text, 2)
            Exit Sub
        Else
            text_to_cur_inv.Text = ""
        End If
    End If
End Sub

Private Sub Text_PROD_CD_Change()
   
    Select Case text_prod_cd.Text
        Case "S", "s", "SL"
            text_prod_cd.Text = "SL"
'        Case "P", "p", "PP"
'            Text_PROD_CD.Text = "PP"
'        Case "H", "h", "HC"
'            Text_PROD_CD.Text = "HC"
        Case Else
            text_prod_cd.Text = ""
            Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
    End Select

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
    Call FormMenuSetting1(Me, FormType, "FS", sAuthority)
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
'     Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    txt_emp = sUserID
    txt_plt = "C1"
    Call txt_plt_KeyUp(0, 0)
    Call Gp_Sp_HdColColor(ss1, 15)
    TXT_PRC_STS.Text = "A"
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

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
    Set iSumCol = Nothing
    
    Call FormMenuSetting1(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call FormMenuSetting1(Me, FormType, "CLS", sAuthority)
  
    End If
    
    text_prod_cd.Text = "SL"
    dtp_ins_date_prod_date1.RawData = ""
    dtp_ins_date_prod_date2.RawData = ""
    txt_plt = "C1"
    txt_plt_name = ""
    text_prod_cd.Text = "SL"
    txt_PRC_LINE = "1"
    sdb_slab_num.Value = 0
    sdb_slab_wgt.Value = 0
    txt_plt_to = ""
    TXT_PRC_LINTTO = "1"
    txt_MOVE_CAR_NO = ""
    txt_MOVE_SHEET_NO = ""
    txt_MOVE_ts = ""
    ULabel2.Visible = False
    ULabel3.Visible = False
    ULabel55.Visible = False
    ULabel56.Visible = False
    ULabel11.Visible = False
    txt_plt_to.Visible = False
    TXT_PRC_LINTTO.Visible = False
    txt_MOVE_CAR_NO.Visible = False
    txt_MOVE_SHEET_NO.Visible = False
    txt_MOVE_ts.Visible = False
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

     Dim sMesg As String
     Dim S As String
     Dim frtable As String
     txt_plt_to = ""
     TXT_PRC_LINTTO = "1"
     sdb_slab_num.Value = 0
     sdb_slab_wgt.Value = 0
     txt_MOVE_CAR_NO = ""
     txt_MOVE_SHEET_NO = ""
     txt_MOVE_ts = ""
     ULabel2.Visible = False
     ULabel3.Visible = False
     ULabel55.Visible = False
     ULabel56.Visible = False
     ULabel11.Visible = False
     txt_plt_to.Visible = False
     TXT_PRC_LINTTO.Visible = False
     txt_MOVE_CAR_NO.Visible = False
     txt_MOVE_SHEET_NO.Visible = False
     txt_MOVE_ts.Visible = False
    
    If dtp_ins_date_prod_date1.RawData = "" Then
       dtp_ins_date_prod_date1.RawData = Format(Date, "YYYYMM") + "01"
    End If
    If dtp_ins_date_prod_date2.RawData = "" Then
       dtp_ins_date_prod_date2.RawData = Format(Date, "YYYYMMDD")
    End If
     
 
   If text_prod_cd.Text <> " " Then
        If text_prod_cd.Text = "SL" Then
           frtable = "fp_slab"
        ElseIf text_prod_cd.Text = "PP" Then
           frtable = "gp_plate"
        ElseIf text_prod_cd.Text = "HC" Then
           frtable = "gp_coil"
        Else
           Call MsgBox("产品名称错误！" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
           Exit Sub
        End If
        
        sQuery = "SELECT a.mat_no,a.ins_date,a.ins_time,GF_EMPNAMEFIND(a.INS_EMP) AS INS_NAME,  " & vbCr
        sQuery = sQuery + "     a.prod_cd,a.aply_stdspec,a.stlgrd,a.thk,a.wid,a.len,a.wgt,      " & vbCr
        sQuery = sQuery + "     Gf_ComnNameFind('C0013',a.to_plt),to_prc,to_prc_line,nvl(b.loc,' ')," & vbCr
        sQuery = sQuery + "     Gf_ComnNameFind('C0013',a.from_plt),a.from_prc, " & vbCr
        sQuery = sQuery + "     a.from_prc_line,a.move_car_no,a.move_sheet_no,a.move_date,      " & vbCr
        sQuery = sQuery + "     a.move_time,GF_EMPNAMEFIND(a.MOVE_EMP) AS MOVE_NAME ,a.prc_sts  " & vbCr
        sQuery = sQuery + "   FROM CP_MOVE_INS a," + frtable + " b                              " & vbCr
        sQuery = sQuery + "  WHERE a.PROD_CD = '" + Trim(text_prod_cd) + "'                     " & vbCr
        sQuery = sQuery + "    AND NVL(a.PRC_STS,' ') like '" + Trim(TXT_PRC_STS) + "%'         " & vbCr
        sQuery = sQuery + "    AND NVL(a.INS_DATE,' ') BETWEEN '" + Trim(dtp_ins_date_prod_date1.RawData) + "' AND "
        sQuery = sQuery + "                                    '" + Trim(dtp_ins_date_prod_date2.RawData) + "' " & vbCr
        sQuery = sQuery + "    AND a.from_plt like '" + Trim(text_fr_cur_inv_code.Text) + "%'   " & vbCr
        sQuery = sQuery + "    AND a.to_plt like '" + Trim(text_to_cur_inv_code.Text) + "%'     "
  
        If text_prod_cd.Text = "SL" Then
           sQuery = sQuery + " AND a.mat_no = b.slab_no  " & vbCr
        ElseIf text_prod_cd.Text = "PP" Then
           sQuery = sQuery + " AND a.mat_no = b.plate_no " & vbCr
        ElseIf text_prod_cd.Text = "HC" Then
           sQuery = sQuery + " AND a.mat_no = b.coil_no  " & vbCr
        End If
        sQuery = sQuery + "  ORDER BY a.ins_date,a.ins_time " & vbCr
    Else
       Call MsgBox("产品名不能为空！" & Chr(10) & "请更正。", vbExclamation + vbOKOnly, "警告")
       Exit Sub
    End If
         
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
    
        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then
'            If Gf_Sp_Display(M_CN1, sc1.Item("Spread"), squery, "R", sc1.Item("pColumn"), True) Then
            If Gf_Sp_Display(M_CN1, sc1.Item("Spread"), sQuery, sc1.Item("lColumn"), True) Then
                
                Call FormMenuSetting1(Me, FormType, "RE", sAuthority)
            End If
        Else
            sMesg = sMesg + " Must input according to length of item"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
    Else
        sMesg = sMesg + " Must input necessarily"
        Call Gp_MsgBoxDisplay(sMesg)
    End If

End Sub
Public Sub Form_Pro()
Dim i As Long
Dim iCount As Integer

 With ss1

    For i = 1 To .MaxRows
        .Col = 0
        .Row = i
        If .Text = "Update" Then
           If txt_plt_to <> "" Then
               .Col = 12
               .Text = Trim(txt_plt_to)
           End If
           If TXT_PRC_LINTTO <> "" Then
               .Col = 14
               .Text = Trim(TXT_PRC_LINTTO)
           End If
           If txt_MOVE_CAR_NO <> "" Then
               .Col = 19
               .Text = Trim(txt_MOVE_CAR_NO)
           End If
           If txt_MOVE_SHEET_NO <> "" Then
               .Col = 20
               .Text = Trim(txt_MOVE_SHEET_NO)
           End If
           If txt_MOVE_ts <> "____-__-__ __:__:__" Then
               .Col = 21
               .Text = Left(txt_MOVE_ts.RawData, 4) + "-" + Mid(txt_MOVE_ts.RawData, 5, 2) + "-" + Mid(txt_MOVE_ts.RawData, 7, 2)
               .Col = 22
               .Text = Mid(txt_MOVE_ts.RawData, 9, 2) + ":" + Mid(txt_MOVE_ts.RawData, 11, 2) + ":" + Mid(txt_MOVE_ts.RawData, 13, 2)

            End If
            .Col = 23
            .Text = Trim(txt_emp)
        End If
        
        If .Text = "Delete" Then
           .Col = 23
           .Text = Trim(txt_emp)
        End If
    Next i
 End With
 
  With ss1
  
  For iCount = 1 To ss1.MaxRows
       .Col = 0
       .Row = iCount
      
       If .Text = "Update" Then
          If txt_plt_to.Text = "" Then
              Gp_MsgBoxDisplay ("目的库必须选择")
                 Exit Sub
            ElseIf txt_MOVE_CAR_NO.Text = "" Then
                Gp_MsgBoxDisplay ("移送车辆号必须输入")
                  Exit Sub
            ElseIf txt_MOVE_SHEET_NO.Text = "" Then
                  Gp_MsgBoxDisplay ("移送提货单号必须输入")
                    Exit Sub
            ElseIf txt_MOVE_ts.RawData = "" Then
                   Gp_MsgBoxDisplay ("执行时间必须输入")
                     Exit Sub
           End If
          
      End If
        
       
  Next iCount
  If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call FormMenuSetting1(Me, FormType, "RE", sAuthority)

  End With
  
  
  
'   If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

'    Call Form_Ref
    

End Sub
'Public Function Gf_Sp_Process(Conn As ADODB.Connection, Sc As Collection, Optional Mc As Collection, _
'                              Optional RefChek As Boolean = False) As Boolean
'
''On Error GoTo SpreadPro_Error
'
'    Dim iCol, iCount, iProcessCount As Integer
'    Dim ret_Result_ErrCode As Integer
'    Dim ret_Result_ErrMsg As String
'
'    Dim dTempInt As Double
'    Dim dTempFloat As Double
'
'    Dim sMesg As String
'    Dim sTemp As String
'    Dim ProcessChk As String
'    Dim DelYN As Boolean
'    Dim Msg_Count As Integer
'    Dim Msg_Yes As String
'
'    Dim adoCmd As ADODB.Command
'
'    Gf_Sp_Process = True
'    iProcessCount = 0
'
'    'MaxRow = 0 is Exit Function Or iCount = 0
'    If Sc.Item("Spread").MaxRows < 1 Or Sc.Item("iColumn").Count = 0 Then
'        Gf_Sp_Process = False
'        Exit Function
'    End If
'
'    Screen.MousePointer = vbHourglass
'    Sc.Item("Spread").ReDraw = False
'
'    'NeceCheck
'    For iCount = 1 To Sc.Item("Spread").MaxRows
'
'        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
'
'            Case "Input", "Update"
'
'                If Not Mc Is Nothing Then
'                    Call Gp_Sp_Move(iCount, Sc, Mc)
'                End If
'
'                'Maxlength Check
'                sMesg = Gf_Sp_NeceCheck2(Sc.Item("Spread"), Sc.Item("mColumn"), iCount, Sc.Item("nColumn"))
'
'                If Trim(sMesg) = "OK" Then
'
'                ElseIf Mid(sMesg, 1, 5) = "FALSE" Then
'                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
'                    sMesg = Mid(sMesg, 6, Len(sMesg))
'                    sMesg = sMesg + "长度不正确"
'                    Call Gp_MsgBoxDisplay(sMesg)
'                    Screen.MousePointer = vbDefault
'                    Set adoCmd = Nothing
'                    Gf_Sp_Process = False
'                    Exit Function
'                Else
'                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
'                    sMesg = sMesg + "必须输入"
'                    Call Gp_MsgBoxDisplay(sMesg)
'                    Screen.MousePointer = vbDefault
'                    Set adoCmd = Nothing
'                    Gf_Sp_Process = False
'                    Exit Function
'                End If
'
'        End Select
'
'    Next iCount
'
'    'Db Connection Check
'    If Conn Is Nothing Then
'        If GF_DbConnect = False Then Gf_Sp_Process = False: Exit Function
'    End If
'
'    'Ado Setting
'    Conn.CursorLocation = adUseServer
'    Set adoCmd = New ADODB.Command
'
'    Set adoCmd.ActiveConnection = Conn
'    adoCmd.CommandType = adCmdStoredProc
'    adoCmd.CommandText = Sc.Item("P-M")
'
'    Conn.BeginTrans
'
'    'Create Parameter (Input) iType + iColumn
'    For iCount = 0 To Sc.Item("iColumn").Count
'        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
'    Next iCount
'
'    'Create Parameter (Output)
'    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
'    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
'
'    Msg_Count = 1
'    For iCount = 1 To Sc.Item("Spread").MaxRows
'
'        ProcessChk = "NO"
'        DelYN = False
'
'        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
'
'            Case "Input"
'                adoCmd.Parameters(0).Value = "I"
'                ProcessChk = "YES"
'
'            Case "Update"
'                adoCmd.Parameters(0).Value = "U"
'                ProcessChk = "YES"
'
'            Case "Delete"
'                adoCmd.Parameters(0).Value = "D"
'                If Msg_Count = 1 Then
'                   DelYN = Gf_MessConfirm("您确定要删除状态为[Delete]的数据吗？", "Q")
'                   If DelYN Then Msg_Yes = "yes"
'                   Msg_Count = Msg_Count + 1
'                End If
'                If Msg_Yes = "yes" Then DelYN = True
'        End Select
'
'        If ProcessChk = "YES" Or DelYN Then
'
'            'Parameters Setting
'            For iCol = 1 To Sc.Item("iColumn").Count
'
'                Sc.Item("Spread").Col = Sc.Item("iColumn").Item(iCol)
'
'                Select Case Sc.Item("Spread").CellType
'
'                    Case SS_CELL_TYPE_CURRENCY
'                        If Trim(Sc.Item("Spread").Text) = "" Then
'                            adoCmd.Parameters(iCol).Value = 0
'                        Else
'                            dTempFloat = Sc.Item("Spread").Text
'                            adoCmd.Parameters(iCol).Value = Str(dTempFloat)
'                        End If
'
'                    Case SS_CELL_TYPE_NUMBER
'                        If Trim(Sc.Item("Spread").Text) = "" Then
'                            adoCmd.Parameters(iCol).Value = 0
'                        Else
'                            dTempInt = Sc.Item("Spread").Text
'                            adoCmd.Parameters(iCol).Value = Str(dTempInt)
'                        End If
'
'                    Case SS_CELL_TYPE_CHECKBOX
'                        If Sc.Item("Spread").Text = "1" Then
'                            adoCmd.Parameters(iCol).Value = "1"
'                        Else
'                            adoCmd.Parameters(iCol).Value = "0"
'                        End If
'
'                    Case SS_CELL_TYPE_COMBOBOX
'                        If Trim(Sc.Item("Spread").Text) = "" Then
'                            adoCmd.Parameters(iCol).Value = "0"
'                        Else
'                            adoCmd.Parameters(iCol).Value = Trim(Str(Sc.Item("Spread").Value))
'                        End If
'
'                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
'                        If Trim(Sc.Item("Spread").Value) = "" Then
'                            adoCmd.Parameters(iCol).Value = ""
'                        Else
'                            adoCmd.Parameters(iCol).Value = Trim(Str(Sc.Item("Spread").Value))
'                        End If
'
'                    Case SS_CELL_TYPE_DATE
'                        If Trim(Sc.Item("Spread").Text) = "" Then
'                            adoCmd.Parameters(iCol).Value = ""
'                        Else
'                            adoCmd.Parameters(iCol).Value = Mid(Trim(Sc.Item("Spread").Text), 1, 4) & _
'                                                            Mid(Trim(Sc.Item("Spread").Text), 6, 2) & _
'                                                            Mid(Trim(Sc.Item("Spread").Text), 9, 2)
'                        End If
'
'                    Case Else
'                        sTemp = Replace(Sc.Item("Spread").Text, "'", "''")
'                        adoCmd.Parameters(iCol).Value = Trim(sTemp)
'
'                End Select
'
'            Next iCol
'
'            iProcessCount = iProcessCount + 1
'            adoCmd.Execute
'
'            'Error Check
'            If adoCmd("Error") <> "0" Then
'
'                ret_Result_ErrCode = adoCmd("Error")
'                ret_Result_ErrMsg = adoCmd("Messg")
'
'                sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
'
'                Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
'                Call Gp_MsgBoxDisplay(sErrMessg)
'
'                Screen.MousePointer = vbDefault
'                Set adoCmd = Nothing
'
'                Conn.RollbackTrans
'                Gf_Sp_Process = False
'                Exit Function
'
'             End If
'
'        End If
'
'    Next iCount
'
'    Conn.CommitTrans
'
'    ' 0 Column Space
'    For iCount = 1 To Sc.Item("Spread").MaxRows
'
'        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
'
'            Case "Input", "Update"
'                Call Gp_Sp_SendData(Sc.Item("Spread"), "", 0, iCount)
'
'            Case "Delete"
'                If DelYN Then
'                   Call Gp_Sp_SendData(Sc.Item("Spread"), "", 0, iCount)
'                   Call Gp_Sp_DeleteRow(Sc.Item("Spread"), iCount)
'                   iCount = iCount - 1
'                End If
'        End Select
'
'    Next iCount
'
'    Sc.Item("Spread").ReDraw = True
'
'    If iProcessCount > 0 Then
'        If Not Mc Is Nothing Then
'            If RefChek = False Then Call Form_Ref
'
'        Else
'            If RefChek = False Then Exit Function
'        End If
'
'        MDIMain.StatusBar1.Panels(1) = "提示信息：成功处理了" & iProcessCount & "条记录"
'        'Call Gp_MsgBoxDisplay("Data that handle is " & iProcessCount & " items", "I")
'
'    End If
'
'    If iProcessCount > 0 Then
'        If Not Mc Is Nothing Then
'            Call Gp_Ms_ControlLock(Mc.Item("lControl"), True)
'        End If
'    Else
'        Gf_Sp_Process = False
'    End If
'
'    Screen.MousePointer = vbDefault
'    Exit Function
'
'SpreadPro_Error:
'
'    Set adoCmd = Nothing
'    Conn.RollbackTrans
'    Gf_Sp_Process = False
'    Call Gp_MsgBoxDisplay("Gf_Sp_Process Error : " & Error)
'    Screen.MousePointer = vbDefault
'
'End Function


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

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub
Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
Dim sMesg As String
Dim PRE As Long

     Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)

    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

 If Row < 1 Then Exit Sub
 If ss1.MaxRows < 1 Then Exit Sub
 ss1.Row = Row
 ss1.Col = 24
 If ss1.Text <> "A" Then
    sMesg = "只能选择状态为“A”的物料"
    Call Gp_MsgBoxDisplay(sMesg)
    Exit Sub
 End If
    ss1.Row = Row
    ss1.Col = 0
    
    If ss1.Text <> "Update" And ss1.Text <> "Delete" Then
       ss1.Col = 0
       ss1.Text = "Update"
       ss1.Col = 11
       sdb_slab_num.Value = sdb_slab_num.Value + 1
       sdb_slab_wgt.Value = sdb_slab_wgt.Value + ss1.Value
       Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
    
   Else
       
       ss1.Col = 11
       sdb_slab_num.Value = sdb_slab_num.Value - 1
       sdb_slab_wgt.Value = sdb_slab_wgt.Value - ss1.Value
       Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
       PRE = Row
       ss1.Row = PRE - 1
       ss1.Col = 0
       If PRE <> 0 Then
          ss1.Row = Row
          ss1.Text = Trim(Str(Row))
       Else
          ss1.Row = Row
          ss1.Text = "1"
       End If
   End If


    ULabel2.Visible = True
    'ULabel3.Visible = True
    ULabel55.Visible = True
    ULabel56.Visible = True
    ULabel11.Visible = True
    txt_plt_to.Visible = True
    'TXT_PRC_LINTTO.Visible = True
    txt_MOVE_CAR_NO.Visible = True
    txt_MOVE_SHEET_NO.Visible = True
    txt_MOVE_ts.Visible = True

End Sub

Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub Text_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=text_prod_cd

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

End Sub



Private Sub txt_MOVE_CAR_NO_DblClick()
    Call txt_MOVE_CAR_NO_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_MOVE_CAR_NO_KeyUp(KeyCode As Integer, Shift As Integer)
    If ULabel55.Caption <> "转库车辆号" Then Exit Sub
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_MOVE_CAR_NO
  '      DD.rControl.Add Item:=txt_fac_name

        DD.nameType = "2"

        Call ACB4020C.Gf_CAR_NO_DD(M_CN1, KeyCode)

    End If

End Sub

Private Sub txt_MOVE_ts_DblClick()

    txt_MOVE_ts.RawData = Format(Now, "YYYYMMDDHHMMSS")

End Sub


Private Sub txt_PLT_DblClick()
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
        Exit Sub
        
    End If

    If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
        txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        txt_plt_name.Text = ""
    End If

End Sub

Private Sub txt_plt_to_DblClick()
    Call txt_plt_to_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_plt_to_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
        DD.rControl.Add Item:=txt_plt_to
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

End Sub

Private Sub TXT_PRC_STS_DblClick()
    Call TXT_PRC_STS_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_PRC_STS_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Z0004"
        DD.rControl.Add Item:=TXT_PRC_STS
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

End Sub

Public Sub Spread_Del()
    
    Dim i As Long
    
    With ss1
        
        If .MaxRows < 1 Then Exit Sub
        If .SelBlockRow < 1 Then Exit Sub
        
        For i = .SelBlockRow To .SelBlockRow2
            .Row = i
            .Col = 0
            
            If Trim(.Text) = "Update" Then
                .Text = "Delete"
            End If
        Next i
        
    End With
End Sub
Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

'Public Sub Gp_Sp_Collection1(sPname As Variant, Num As Integer, pcol As String, ncol As String, mcol As String, _
'                                                               iCol As String, acol As String, lCol As String, _
'                            pColumn As Collection, nColumn As Collection, mColumn As Collection, iColumn As Collection, _
'                            aColumn As Collection, lColumn As Collection)
'
'    If LCase(Trim(pcol)) = "p" Then       'PK Column
'        pColumn.Add Item:=Num
'    End If
'
'    If LCase(Trim(ncol)) = "n" Then       'Necessary Column
'        nColumn.Add Item:=Num
'        'Call Gp_Sp_ColColor(SpName, Num, , &H80FF80)
'    End If
'
'    If LCase(Trim(mcol)) = "m" Then       'Spread Maxlength check Column
'        mColumn.Add Item:=Num
'    End If
'
'    If LCase(Trim(iCol)) = "i" Then       'Spread Insert Column
'        iColumn.Add Item:=Num
'
'    End If
'
'    If LCase(Trim(acol)) = "a" Then       'Master -> Spread Column
'        aColumn.Add Item:=Num
'        Call Gp_Sp_ColHidden(sPname, Num, True)
'    End If
'
'    If LCase(Trim(lCol)) = "l" Then       'Spread Lock Column
'        lColumn.Add Item:=Num
'        Call Gp_Sp_ColLock(sPname, Num, True)
'    End If
'
'
'End Sub
'
Public Sub FormMenuSetting1(Fm As Variant, FormType As String, ButtonType As String, sAuthority As String)



On Error Resume Next

    With MDIMain.MenuTool

        Select Case FormType

               Case "Start"
                    .Buttons(1).Enabled = False                 'Screen Clear
                    .Buttons(2).Enabled = False                 'Refer
                    .Buttons(3).Enabled = False                 'Separator
                    .Buttons(4).Enabled = False                 'Save
                    .Buttons(5).Enabled = False                 'Delete
                    .Buttons(6).Enabled = False                 'Separator
                    .Buttons(7).Enabled = False                 'Row Insert
                    .Buttons(8).Enabled = False                 'Row Delete
                    .Buttons(9).Enabled = False                 'Row Cancel
                    .Buttons(10).Enabled = False                'Separator
                    .Buttons(11).Enabled = False                'Copy
                    .Buttons(12).Enabled = False                'Paste
                    .Buttons(13).Enabled = False                'Separator
                    .Buttons(14).Enabled = False                'Excel
                    .Buttons(15).Enabled = False                'Print
                    .Buttons(16).Enabled = False                'Separator
                    .Buttons(17).Visible = True                 'Exit

                  Case "Msheet"
                    .Buttons(1).Enabled = True                  'Screen Clear
                    .Buttons(2).Enabled = True                  'Refer
                    .Buttons(3).Enabled = True                  'Separator
                    .Buttons(4).Enabled = True                  'Save
                    .Buttons(5).Enabled = False                 'Delete
                    .Buttons(6).Enabled = True                  'Separator
                    .Buttons(7).Enabled = False                 'Row Insert
                    .Buttons(8).Enabled = True                 'Row Delete
                    .Buttons(9).Enabled = False                 'Row Cancel
                    .Buttons(10).Enabled = True                 'Separator

                    .Buttons(11).Enabled = False                'Copy
                    .Buttons(11).ButtonMenus(1).Enabled = False 'All Copy
                    .Buttons(11).ButtonMenus(2).Enabled = False 'Master Copy
                    .Buttons(11).ButtonMenus(3).Enabled = True  'Spread Copy

                    .Buttons(12).Enabled = False                 'Paste
                    .Buttons(12).ButtonMenus(1).Enabled = False 'All Paste
                    .Buttons(12).ButtonMenus(2).Enabled = False 'Master Paste
                    .Buttons(12).ButtonMenus(3).Enabled = False 'Spread Paste

                    .Buttons(13).Enabled = True                 'Separator
                    .Buttons(14).Enabled = True                 'Excel
                    .Buttons(15).Enabled = False                'Print
                    .Buttons(16).Enabled = True                 'Separator
                    .Buttons(17).Enabled = True                 'Exit

        End Select

        Fm.Toolbar_St = ButtonType

        Select Case ButtonType
                 'Save, Refer
            Case "SE", "RE"

                Select Case FormType

                    Case "Msheet"
                        .Buttons(7).Enabled = False             'Row Insert
                        .Buttons(8).Enabled = True              'Row Delete
                        .Buttons(9).Enabled = False             'Row Cancel
                        .Buttons(14).Enabled = True             'Excel
                     End Select

                 'Form Start, Screen Clear
            Case "FS", "CLS"

                Select Case FormType

                    Case "Msheet"
                        .Buttons(7).Enabled = False             'Row Insert
                        .Buttons(8).Enabled = False             'Row Delete
                        .Buttons(9).Enabled = False             'Row Cancel
                        .Buttons(14).Enabled = False            'Excel

                End Select

            Case "Acopy"

                .Buttons(12).ButtonMenus(1).Enabled = True      'All Paste
                .Buttons(12).ButtonMenus(2).Enabled = False     'Master Paste
                .Buttons(12).ButtonMenus(3).Enabled = False     'Spread Paste

            Case "Mcopy"

                .Buttons(12).ButtonMenus(1).Enabled = False     'All Paste
                .Buttons(12).ButtonMenus(2).Enabled = True      'Master Paste
                .Buttons(12).ButtonMenus(3).Enabled = False     'Spread Paste

            Case "Scopy"

                .Buttons(12).ButtonMenus(1).Enabled = False     'All Paste
                .Buttons(12).ButtonMenus(2).Enabled = False     'Master Paste
                .Buttons(12).ButtonMenus(3).Enabled = True      'Spread Paste

        End Select

        'Autority Inquiry Check
        If Mid(sAuthority, 1, 1) = "0" Then
            .Buttons(2).Enabled = False                         'Refer
        End If

        Select Case Mid(sAuthority, 2, 3) 'Insert, Update, Delete

            Case "000"      'No Authority
                .Buttons(4).Enabled = False                     'Save
                .Buttons(5).Enabled = False                     'Delete
                .Buttons(7).Enabled = False                     'Row Insert
                .Buttons(8).Enabled = False                     'Row Delete
                .Buttons(9).Enabled = False                     'Row Cancel
                .Buttons(11).Enabled = False                    'Copy
                .Buttons(12).Enabled = False                    'Paste

            Case "001"      'Delete Authority
                .Buttons(7).Enabled = False                     'Row Insert
                .Buttons(11).Enabled = False                    'Copy
                .Buttons(12).Enabled = False                    'Paste

            Case "010"      'Update Authority
                .Buttons(5).Enabled = False                     'Delete
                .Buttons(7).Enabled = False                     'Row Insert
                .Buttons(8).Enabled = False                     'Row Delete
                .Buttons(11).Enabled = False                    'Copy
                .Buttons(12).Enabled = False                    'Paste

            Case "011"      'Update, Delete Authority
                .Buttons(7).Enabled = False                     'Row Insert
                .Buttons(11).Enabled = False                    'Copy
                .Buttons(12).Enabled = False                    'Paste

            Case "100"      'Insert Authority
                .Buttons(5).Enabled = False                     'Delete
                .Buttons(8).Enabled = False                     'Row Delete

            Case "101"      'Insert, Delete Authority

            Case "110"      'Insert, Update Authority
                .Buttons(5).Enabled = False                     'Delete
                .Buttons(8).Enabled = False                     'Row Delete

            Case "111"      'Insert, Update, Delete Authority

        End Select

        .Wrappable = True

    End With

End Sub

  
Private Function Gf_Sp_Process(Conn As ADODB.Connection, Sc As Collection, Optional MC As Collection, _
                              Optional RefChek As Boolean = False) As Boolean

'On Error GoTo SpreadPro_Error

    Dim iCol, iCount, iProcessCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim dTempFloat As Double
    
    Dim sMesg As String
    Dim sTemp As String
    Dim ProcessChk As String
    Dim DelYN As Boolean
    Dim Msg_Count As Integer
    Dim Msg_Yes As String
    
    Dim adoCmd As ADODB.Command

    Gf_Sp_Process = True
    iProcessCount = 0
    
    'MaxRow = 0 is Exit Function Or iCount = 0
    If Sc.Item("Spread").MaxRows < 1 Or Sc.Item("iColumn").Count = 0 Then
        Gf_Sp_Process = False
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    Sc.Item("Spread").ReDraw = False
    
    'NeceCheck
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
            
            Case "Input", "Update"
            
                If Not MC Is Nothing Then
                    Call Gp_Sp_Move(iCount, Sc, MC)
                End If
                
                'Maxlength Check
                sMesg = Gf_Sp_NeceCheck2(Sc.Item("Spread"), Sc.Item("mColumn"), iCount, Sc.Item("nColumn"))
                        
                If Trim(sMesg) = "OK" Then
                    
                ElseIf Mid(sMesg, 1, 5) = "FALSE" Then
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    sMesg = Mid(sMesg, 6, Len(sMesg))
                    sMesg = sMesg + "长度不正确"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Process = False
                    Exit Function
                Else
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    sMesg = sMesg + "必须输入"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Process = False
                    Exit Function
                End If
        
        End Select
    
    Next iCount
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_Sp_Process = False: Screen.MousePointer = vbDefault: Exit Function
    End If
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = Conn
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = Sc.Item("P-M")
    
    Conn.BeginTrans
    
    'Create Parameter (Input) iType + iColumn
    For iCount = 0 To Sc.Item("iColumn").Count
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    Msg_Count = 1
    For iCount = 1 To Sc.Item("Spread").MaxRows
        
        ProcessChk = "NO"
        DelYN = False
        
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
        
            Case "Input"
                adoCmd.Parameters(0).Value = "I"
                ProcessChk = "YES"
                
            Case "Update"
                adoCmd.Parameters(0).Value = "U"
                ProcessChk = "YES"
                
            Case "Delete"
                adoCmd.Parameters(0).Value = "D"
                If Msg_Count = 1 Then
                   DelYN = Gf_MessConfirm("您确定要删除状态为[Delete]的数据吗？", "Q")
                   If DelYN Then Msg_Yes = "yes"
                   Msg_Count = Msg_Count + 1
                End If
                If Msg_Yes = "yes" Then DelYN = True
        End Select
          
        If ProcessChk = "YES" Or DelYN Then
            
            'Parameters Setting
            For iCol = 1 To Sc.Item("iColumn").Count
            
                Sc.Item("Spread").Col = Sc.Item("iColumn").Item(iCol)
                
                Select Case Sc.Item("Spread").CellType
                
                    Case SS_CELL_TYPE_CURRENCY
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempFloat = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Str(dTempFloat)
                        End If
                        
                    Case SS_CELL_TYPE_NUMBER
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempInt = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Str(dTempInt)
                        End If
                        
                    Case SS_CELL_TYPE_CHECKBOX
                        If Sc.Item("Spread").Text = "1" Then
                            adoCmd.Parameters(iCol).Value = "1"
                        Else
                            adoCmd.Parameters(iCol).Value = "0"
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = "0"
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(Str(Sc.Item("Spread").Value))
                        End If
                        
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If Trim(Sc.Item("Spread").Value) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(Str(Sc.Item("Spread").Value))
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Mid(Trim(Sc.Item("Spread").Text), 1, 4) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 6, 2) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 9, 2)
                        End If
                       
                    Case Else
                        sTemp = Replace(Sc.Item("Spread").Text, "'", "''")
                        adoCmd.Parameters(iCol).Value = Trim(sTemp)
                        
                End Select
           
            Next iCol
                           
            iProcessCount = iProcessCount + 1
            adoCmd.Execute
            
            'Error Check
            If adoCmd("Error") <> "0" Then
            
                ret_Result_ErrCode = adoCmd("Error")
                ret_Result_ErrMsg = adoCmd("Messg")
        
                sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                
                Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                Screen.MousePointer = vbDefault
                Set adoCmd = Nothing
                
                Conn.RollbackTrans
                Gf_Sp_Process = False
                Exit Function
        
             End If
        
        End If
        
    Next iCount
    
    Conn.CommitTrans
    
    ' 0 Column Space
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
        
            Case "Input", "Update"
                Call Gp_Sp_SendData(Sc.Item("Spread"), "", 0, iCount)
                
            Case "Delete"
                If DelYN Then
                   Call Gp_Sp_SendData(Sc.Item("Spread"), "", 0, iCount)
                   Call Gp_Sp_DeleteRow(Sc.Item("Spread"), iCount)
                   iCount = iCount - 1
                End If
        End Select
        
    Next iCount
    
    Sc.Item("Spread").ReDraw = True
    
    If iProcessCount > 0 Then
        If Not MC Is Nothing Then
            If RefChek = False Then Call Form_Ref
                                                    
        Else
            If RefChek = False Then Screen.MousePointer = vbDefault: Exit Function
        End If
        
        MDIMain.StatusBar1.Panels(1) = "提示信息：成功处理了" & iProcessCount & "条记录"
        'Call Gp_MsgBoxDisplay("Data that handle is " & iProcessCount & " items", "I")
        
    End If
            
    If iProcessCount > 0 Then
        If Not MC Is Nothing Then
            Call Gp_Ms_ControlLock(MC.Item("lControl"), True)
        End If
    Else
        Gf_Sp_Process = False
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function

SpreadPro_Error:
    
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Gf_Sp_Process = False
    Call Gp_MsgBoxDisplay("Gf_Sp_Process Error : " & Error)
    Screen.MousePointer = vbDefault

End Function
