VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0030C 
   Caption         =   "��׼������Ϣ��ѯ - AQA0030C"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15630
   BeginProperty Font 
      Name            =   "����"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   15630
   WindowState     =   2  'Maximized
   Begin Threed.SSRibbon srbt_TEST_FL 
      Height          =   300
      Left            =   13170
      TabIndex        =   13
      Top             =   465
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   529
      _Version        =   196609
      PictureFrames   =   1
      Picture         =   "AQA0030C.frx":0000
      PictureDnFrames =   1
      Outline         =   0   'False
      PictureDn       =   "AQA0030C.frx":0354
      Value           =   -1  'True
   End
   Begin VB.TextBox txt_TEST_FL 
      Height          =   300
      Left            =   13770
      MaxLength       =   1
      TabIndex        =   12
      Text            =   "Y"
      Top             =   465
      Visible         =   0   'False
      Width           =   375
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   315
      Left            =   8235
      TabIndex        =   8
      Top             =   60
      Width           =   1965
      _Version        =   393216
      _ExtentX        =   3466
      _ExtentY        =   556
      _StockProps     =   64
      ColHeaderDisplay=   0
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   2
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "AQA0030C.frx":0468
   End
   Begin VB.TextBox txt_THK_MIN 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   13260
      TabIndex        =   10
      Tag             =   "�����-��С"
      Top             =   0
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt_THK_MAX 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   14070
      TabIndex        =   9
      Tag             =   "�����-���"
      Top             =   0
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmd_ListView 
      Caption         =   " <<"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txt_STDSPEC 
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
      Left            =   1425
      MaxLength       =   18
      TabIndex        =   0
      Top             =   60
      Width           =   2595
   End
   Begin VB.TextBox txt_STDSPEC_YY 
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
      Left            =   5475
      MaxLength       =   4
      TabIndex        =   1
      Top             =   60
      Width           =   735
   End
   Begin VB.TextBox txt_SMP_LOC 
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
      Left            =   1425
      MaxLength       =   1
      TabIndex        =   2
      Top             =   450
      Width           =   555
   End
   Begin VB.TextBox txt_SMP_LOC_NAME 
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
      Height          =   315
      Left            =   1995
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   6
      Top             =   450
      Width           =   2025
   End
   Begin VB.TextBox txt_INS_EMP 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12840
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   405
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   120
      Top             =   60
      Width           =   1275
      _ExtentX        =   2249
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   1
      Left            =   4140
      Top             =   60
      Width           =   1275
      _ExtentX        =   2249
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   0
      Left            =   6870
      Top             =   60
      Width           =   1275
      _ExtentX        =   2249
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
   Begin CSTextLibCtl.sidbEdit txt_SMP_LEN 
      Height          =   315
      Left            =   5475
      TabIndex        =   3
      Top             =   450
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
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
      AutoScroll      =   0   'False
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
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   4
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   90
      Top             =   450
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "ȡ����λ��"
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
      Index           =   2
      Left            =   4140
      Top             =   450
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "ȡ���ϳ���"
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
   Begin CSTextLibCtl.sidbEdit txt_PRE_SMP_QTY 
      Height          =   315
      Left            =   8235
      TabIndex        =   4
      Top             =   450
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
      AutoScroll      =   0   'False
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
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   0
      Left            =   6870
      Top             =   450
      Width           =   1275
      _ExtentX        =   2249
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
   Begin InDate.ULabel ULabel2 
      Height          =   300
      Index           =   2
      Left            =   9330
      Top             =   450
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "ȡ��������"
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
   Begin CSTextLibCtl.sidbEdit sdb_SMP_STD_WGT 
      Height          =   300
      Left            =   10710
      TabIndex        =   11
      Top             =   450
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
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
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   5
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel2 
      Height          =   300
      Index           =   3
      Left            =   11820
      Top             =   465
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "Ҫ�����"
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
   Begin TabDlg.SSTab ssT 
      Height          =   10095
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   17806
      _Version        =   393216
      Tabs            =   10
      Tab             =   9
      TabsPerRow      =   10
      TabHeight       =   520
      WordWrap        =   0   'False
      BackColor       =   14737632
      TabCaption(0)   =   "��������"
      TabPicture(0)   =   "AQA0030C.frx":07E0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txt_DRAW_DSC_CD(11)"
      Tab(0).Control(1)=   "txt_RA_DSC_CD(2)"
      Tab(0).Control(2)=   "txt_SMP_WID(0)"
      Tab(0).Control(3)=   "txt_SP_EL_CD(0)"
      Tab(0).Control(4)=   "txt_EL_CD(0)"
      Tab(0).Control(5)=   "txt_EL_CD(1)"
      Tab(0).Control(6)=   "txt_SP_EL_CD(1)"
      Tab(0).Control(7)=   "txt_SG_EL_CD(0)"
      Tab(0).Control(8)=   "txt_SG_EL_CD(1)"
      Tab(0).Control(9)=   "txt_SNPP_EL_CD(0)"
      Tab(0).Control(10)=   "txt_SNPP_EL_CD(1)"
      Tab(0).Control(11)=   "txt_SP_EL_SMP_CD(0)"
      Tab(0).Control(12)=   "txt_TENCIL_SMP_CD(0)"
      Tab(0).Control(13)=   "txt_YP_DSC_CD(0)"
      Tab(0).Control(14)=   "txt_SP_EL_DSC_CD(0)"
      Tab(0).Control(15)=   "txt_SG_EL_DSC_CD(0)"
      Tab(0).Control(16)=   "txt_SNPP_EL_DSC_CD(0)"
      Tab(0).Control(17)=   "txt_EL_DSC_CD(0)"
      Tab(0).Control(18)=   "txt_RA_DSC_CD(0)"
      Tab(0).Control(19)=   "txt_TS_DSC_CD(0)"
      Tab(0).Control(20)=   "txt_YR_DSC_CD(0)"
      Tab(0).Control(21)=   "txt_RA_DIR_CD(0)"
      Tab(0).Control(22)=   "txt_RA_DIR_NAME(0)"
      Tab(0).Control(23)=   "txt_HTM_CD(0)"
      Tab(0).Control(24)=   "ULabel4(1)"
      Tab(0).Control(25)=   "ULabel4(0)"
      Tab(0).Control(26)=   "ULabel4(21)"
      Tab(0).Control(27)=   "ULabel4(20)"
      Tab(0).Control(28)=   "ULabel4(19)"
      Tab(0).Control(29)=   "ULabel4(18)"
      Tab(0).Control(30)=   "ULabel4(17)"
      Tab(0).Control(31)=   "ULabel4(16)"
      Tab(0).Control(32)=   "ULabel4(2)"
      Tab(0).Control(33)=   "ULabel4(3)"
      Tab(0).Control(34)=   "ULabel4(4)"
      Tab(0).Control(35)=   "ULabel4(5)"
      Tab(0).Control(36)=   "ULabel4(7)"
      Tab(0).Control(37)=   "ULabel4(6)"
      Tab(0).Control(38)=   "ULabel4(8)"
      Tab(0).Control(39)=   "ULabel4(22)"
      Tab(0).Control(40)=   "sdb_DRAW_MIN(1)"
      Tab(0).Control(41)=   "sdb_YP_MAX(0)"
      Tab(0).Control(42)=   "ULabel4(9)"
      Tab(0).Control(43)=   "sdb_DRAW_MIN(2)"
      Tab(0).Control(44)=   "sdb_TS_MAX(0)"
      Tab(0).Control(45)=   "ULabel4(10)"
      Tab(0).Control(46)=   "sdb_DRAW_MIN(3)"
      Tab(0).Control(47)=   "ULabel4(11)"
      Tab(0).Control(48)=   "sdb_DRAW_MIN(6)"
      Tab(0).Control(49)=   "sdb_EL_MAX(0)"
      Tab(0).Control(50)=   "ULabel4(12)"
      Tab(0).Control(51)=   "sdb_DRAW_MIN(8)"
      Tab(0).Control(52)=   "sdb_SNPP_EL_MAX(0)"
      Tab(0).Control(53)=   "ULabel4(13)"
      Tab(0).Control(54)=   "sdb_DRAW_MIN(9)"
      Tab(0).Control(55)=   "sdb_SG_EL_MAX(0)"
      Tab(0).Control(56)=   "ULabel4(14)"
      Tab(0).Control(57)=   "sdb_DRAW_MIN(10)"
      Tab(0).Control(58)=   "sdb_SP_EL_MAX(0)"
      Tab(0).Control(59)=   "ULabel4(15)"
      Tab(0).Control(60)=   "ULabel4(59)"
      Tab(0).Control(61)=   "sdb_DRAW_MIN(7)"
      Tab(0).Control(62)=   "sdb_YR_MAX(0)"
      Tab(0).Control(63)=   "ULabel4(60)"
      Tab(0).Control(64)=   "ULabel4(95)"
      Tab(0).Control(65)=   "ULabel4(111)"
      Tab(0).Control(66)=   "ULabel4(138)"
      Tab(0).Control(67)=   "sdb_DRAW_MIN(4)"
      Tab(0).Control(68)=   "ULabel4(139)"
      Tab(0).Control(69)=   "ULabel4(142)"
      Tab(0).Control(70)=   "sdb_RA_MAX(0)"
      Tab(0).Control(71)=   "sdb_DRAW_MIN(5)"
      Tab(0).Control(72)=   "ULabel4(143)"
      Tab(0).Control(73)=   "sdb_DRAW_MIN(11)"
      Tab(0).Control(74)=   "sdb_DRAW_MAX(11)"
      Tab(0).Control(75)=   "ULabel4(149)"
      Tab(0).Control(76)=   "ULabel4(150)"
      Tab(0).Control(77)=   "Line49(2)"
      Tab(0).Control(78)=   "Line3(0)"
      Tab(0).Control(79)=   "Line3(1)"
      Tab(0).Control(80)=   "Line3(2)"
      Tab(0).Control(81)=   "Line3(3)"
      Tab(0).Control(82)=   "Line3(4)"
      Tab(0).ControlCount=   83
      TabCaption(1)   =   "׷����������"
      TabPicture(1)   =   "AQA0030C.frx":07FC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_A_DRAW_DSC_CD(11)"
      Tab(1).Control(1)=   "txt_RA_DSC_CD(3)"
      Tab(1).Control(2)=   "txt_SMP_WID(1)"
      Tab(1).Control(3)=   "txt_RA_DIR_NAME(1)"
      Tab(1).Control(4)=   "txt_RA_DIR_CD(1)"
      Tab(1).Control(5)=   "txt_YR_DSC_CD(1)"
      Tab(1).Control(6)=   "txt_TS_DSC_CD(1)"
      Tab(1).Control(7)=   "txt_RA_DSC_CD(1)"
      Tab(1).Control(8)=   "txt_EL_DSC_CD(1)"
      Tab(1).Control(9)=   "txt_SNPP_EL_DSC_CD(1)"
      Tab(1).Control(10)=   "txt_SG_EL_DSC_CD(1)"
      Tab(1).Control(11)=   "txt_SP_EL_DSC_CD(1)"
      Tab(1).Control(12)=   "txt_YP_DSC_CD(1)"
      Tab(1).Control(13)=   "txt_TENCIL_SMP_CD(1)"
      Tab(1).Control(14)=   "txt_SP_EL_SMP_CD(1)"
      Tab(1).Control(15)=   "txt_SNPP_EL_CD(3)"
      Tab(1).Control(16)=   "txt_SNPP_EL_CD(2)"
      Tab(1).Control(17)=   "txt_SG_EL_CD(3)"
      Tab(1).Control(18)=   "txt_SG_EL_CD(2)"
      Tab(1).Control(19)=   "txt_SP_EL_CD(3)"
      Tab(1).Control(20)=   "txt_EL_CD(3)"
      Tab(1).Control(21)=   "txt_EL_CD(2)"
      Tab(1).Control(22)=   "txt_SP_EL_CD(2)"
      Tab(1).Control(23)=   "txt_HTM_CD(1)"
      Tab(1).Control(24)=   "ULabel4(70)"
      Tab(1).Control(25)=   "ULabel4(71)"
      Tab(1).Control(26)=   "ULabel4(72)"
      Tab(1).Control(27)=   "ULabel4(73)"
      Tab(1).Control(28)=   "ULabel4(74)"
      Tab(1).Control(29)=   "ULabel4(75)"
      Tab(1).Control(30)=   "ULabel4(76)"
      Tab(1).Control(31)=   "ULabel4(77)"
      Tab(1).Control(32)=   "ULabel4(78)"
      Tab(1).Control(33)=   "ULabel4(79)"
      Tab(1).Control(34)=   "ULabel4(80)"
      Tab(1).Control(35)=   "ULabel4(81)"
      Tab(1).Control(36)=   "ULabel4(82)"
      Tab(1).Control(37)=   "ULabel4(83)"
      Tab(1).Control(38)=   "ULabel4(84)"
      Tab(1).Control(39)=   "ULabel4(85)"
      Tab(1).Control(40)=   "sdb_A_DRAW_MIN(1)"
      Tab(1).Control(41)=   "sdb_YP_MAX(1)"
      Tab(1).Control(42)=   "ULabel4(86)"
      Tab(1).Control(43)=   "sdb_A_DRAW_MIN(2)"
      Tab(1).Control(44)=   "sdb_TS_MAX(1)"
      Tab(1).Control(45)=   "ULabel4(87)"
      Tab(1).Control(46)=   "sdb_A_DRAW_MIN(3)"
      Tab(1).Control(47)=   "sdb_RA_MAX(1)"
      Tab(1).Control(48)=   "ULabel4(88)"
      Tab(1).Control(49)=   "sdb_A_DRAW_MIN(6)"
      Tab(1).Control(50)=   "sdb_EL_MAX(1)"
      Tab(1).Control(51)=   "ULabel4(89)"
      Tab(1).Control(52)=   "sdb_A_DRAW_MIN(8)"
      Tab(1).Control(53)=   "sdb_SNPP_EL_MAX(1)"
      Tab(1).Control(54)=   "ULabel4(90)"
      Tab(1).Control(55)=   "sdb_A_DRAW_MIN(9)"
      Tab(1).Control(56)=   "sdb_SG_EL_MAX(1)"
      Tab(1).Control(57)=   "ULabel4(91)"
      Tab(1).Control(58)=   "sdb_A_DRAW_MIN(10)"
      Tab(1).Control(59)=   "sdb_SP_EL_MAX(1)"
      Tab(1).Control(60)=   "ULabel4(92)"
      Tab(1).Control(61)=   "ULabel4(93)"
      Tab(1).Control(62)=   "sdb_A_DRAW_MIN(7)"
      Tab(1).Control(63)=   "sdb_YR_MAX(1)"
      Tab(1).Control(64)=   "ULabel4(94)"
      Tab(1).Control(65)=   "ULabel4(96)"
      Tab(1).Control(66)=   "ULabel4(135)"
      Tab(1).Control(67)=   "ULabel4(140)"
      Tab(1).Control(68)=   "sdb_A_DRAW_MIN(4)"
      Tab(1).Control(69)=   "ULabel4(141)"
      Tab(1).Control(70)=   "ULabel4(144)"
      Tab(1).Control(71)=   "sdb_A_DRAW_MIN(5)"
      Tab(1).Control(72)=   "ULabel4(145)"
      Tab(1).Control(73)=   "sdb_A_DRAW_MAX(11)"
      Tab(1).Control(74)=   "sdb_A_DRAW_MIN(11)"
      Tab(1).Control(75)=   "ULabel4(147)"
      Tab(1).Control(76)=   "ULabel4(146)"
      Tab(1).Control(77)=   "Line49(3)"
      Tab(1).Control(78)=   "Line3(31)"
      Tab(1).Control(79)=   "Line3(32)"
      Tab(1).Control(80)=   "Line3(33)"
      Tab(1).Control(81)=   "Line3(34)"
      Tab(1).Control(82)=   "Line3(35)"
      Tab(1).ControlCount=   83
      TabCaption(2)   =   "������������"
      TabPicture(2)   =   "AQA0030C.frx":0818
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt_HGT_RA_DSC_CD(2)"
      Tab(2).Control(1)=   "txt_SMP_WID(2)"
      Tab(2).Control(2)=   "txt_HGT_RA_DIR_NAME(0)"
      Tab(2).Control(3)=   "txt_HGT_RA_DIR_CD(0)"
      Tab(2).Control(4)=   "txt_HGT_TENCIL_TMP_UNIT(0)"
      Tab(2).Control(5)=   "txt_HGT_TENCIL_TMP(0)"
      Tab(2).Control(6)=   "txt_HGT_TENCIL_SMP_CD(0)"
      Tab(2).Control(7)=   "txt_HGT_SP_EL_SMP_CD(0)"
      Tab(2).Control(8)=   "txt_HGT_EL_CD(1)"
      Tab(2).Control(9)=   "txt_HGT_EL_CD(0)"
      Tab(2).Control(10)=   "txt_HGT_SNPP_EL_CD(1)"
      Tab(2).Control(11)=   "txt_HGT_SNPP_EL_CD(0)"
      Tab(2).Control(12)=   "txt_HGT_SP_EL_CD(1)"
      Tab(2).Control(13)=   "txt_HGT_SP_EL_CD(0)"
      Tab(2).Control(14)=   "txt_HGT_YP_DSC_CD(0)"
      Tab(2).Control(15)=   "txt_HGT_SP_EL_DSC_CD(0)"
      Tab(2).Control(16)=   "txt_HGT_RA_DSC_CD(0)"
      Tab(2).Control(17)=   "txt_HGT_TS_DSC_CD(0)"
      Tab(2).Control(18)=   "txt_HGT_SNPP_EL_DSC_CD(0)"
      Tab(2).Control(19)=   "txt_HGT_EL_DSC_CD(0)"
      Tab(2).Control(20)=   "txt_HTM_CD(2)"
      Tab(2).Control(21)=   "ULabel27(3)"
      Tab(2).Control(22)=   "ULabel27(4)"
      Tab(2).Control(23)=   "ULabel27(5)"
      Tab(2).Control(24)=   "ULabel27(6)"
      Tab(2).Control(25)=   "ULabel27(7)"
      Tab(2).Control(26)=   "ULabel27(2)"
      Tab(2).Control(27)=   "sdb_HGT_YP_MIN(0)"
      Tab(2).Control(28)=   "sdb_HGT_YP_MAX(0)"
      Tab(2).Control(29)=   "ULabel27(15)"
      Tab(2).Control(30)=   "sdb_HGT_TS_MIN(0)"
      Tab(2).Control(31)=   "sdb_HGT_TS_MAX(0)"
      Tab(2).Control(32)=   "ULabel27(16)"
      Tab(2).Control(33)=   "sdb_HGT_RA_MIN(0)"
      Tab(2).Control(34)=   "sdb_HGT_RA_MAX(0)"
      Tab(2).Control(35)=   "ULabel27(17)"
      Tab(2).Control(36)=   "sdb_HGT_SP_EL_MIN(0)"
      Tab(2).Control(37)=   "sdb_HGT_SP_EL_MAX(0)"
      Tab(2).Control(38)=   "ULabel27(20)"
      Tab(2).Control(39)=   "sdb_HGT_SNPP_EL_MIN(0)"
      Tab(2).Control(40)=   "sdb_HGT_SNPP_EL_MAX(0)"
      Tab(2).Control(41)=   "ULabel27(19)"
      Tab(2).Control(42)=   "sdb_HGT_EL_MIN(0)"
      Tab(2).Control(43)=   "sdb_HGT_EL_MAX(0)"
      Tab(2).Control(44)=   "ULabel27(18)"
      Tab(2).Control(45)=   "ULabel4(24)"
      Tab(2).Control(46)=   "ULabel4(25)"
      Tab(2).Control(47)=   "ULabel4(26)"
      Tab(2).Control(48)=   "ULabel4(27)"
      Tab(2).Control(49)=   "ULabel4(28)"
      Tab(2).Control(50)=   "ULabel4(29)"
      Tab(2).Control(51)=   "ULabel4(30)"
      Tab(2).Control(52)=   "ULabel4(31)"
      Tab(2).Control(53)=   "ULabel4(23)"
      Tab(2).Control(54)=   "ULabel4(97)"
      Tab(2).Control(55)=   "ULabel4(136)"
      Tab(2).Control(56)=   "ULabel27(32)"
      Tab(2).Control(57)=   "sdb_HGT_RA_MIN(2)"
      Tab(2).Control(58)=   "ULabel27(33)"
      Tab(2).Control(59)=   "ULabel27(36)"
      Tab(2).Control(60)=   "sdb_HGT_RA_MIN(4)"
      Tab(2).Control(61)=   "ULabel27(37)"
      Tab(2).Control(62)=   "Line49(4)"
      Tab(2).Control(63)=   "Line3(9)"
      Tab(2).Control(64)=   "Line3(8)"
      Tab(2).Control(65)=   "Line3(7)"
      Tab(2).Control(66)=   "Line3(6)"
      Tab(2).Control(67)=   "Line3(5)"
      Tab(2).ControlCount=   68
      TabCaption(3)   =   "׷�Ӹ�������"
      TabPicture(3)   =   "AQA0030C.frx":0834
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Line3(30)"
      Tab(3).Control(1)=   "Line3(29)"
      Tab(3).Control(2)=   "Line3(28)"
      Tab(3).Control(3)=   "Line3(27)"
      Tab(3).Control(4)=   "Line3(26)"
      Tab(3).Control(5)=   "Line49(7)"
      Tab(3).Control(6)=   "ULabel27(39)"
      Tab(3).Control(7)=   "sdb_HGT_RA_MIN(5)"
      Tab(3).Control(8)=   "ULabel27(38)"
      Tab(3).Control(9)=   "ULabel27(35)"
      Tab(3).Control(10)=   "sdb_HGT_RA_MIN(3)"
      Tab(3).Control(11)=   "ULabel27(34)"
      Tab(3).Control(12)=   "ULabel4(137)"
      Tab(3).Control(13)=   "ULabel4(98)"
      Tab(3).Control(14)=   "ULabel4(69)"
      Tab(3).Control(15)=   "ULabel4(68)"
      Tab(3).Control(16)=   "ULabel4(67)"
      Tab(3).Control(17)=   "ULabel4(66)"
      Tab(3).Control(18)=   "ULabel4(65)"
      Tab(3).Control(19)=   "ULabel4(64)"
      Tab(3).Control(20)=   "ULabel4(63)"
      Tab(3).Control(21)=   "ULabel4(62)"
      Tab(3).Control(22)=   "ULabel4(61)"
      Tab(3).Control(23)=   "ULabel27(30)"
      Tab(3).Control(24)=   "sdb_HGT_EL_MAX(1)"
      Tab(3).Control(25)=   "sdb_HGT_EL_MIN(1)"
      Tab(3).Control(26)=   "ULabel27(29)"
      Tab(3).Control(27)=   "sdb_HGT_SNPP_EL_MAX(1)"
      Tab(3).Control(28)=   "sdb_HGT_SNPP_EL_MIN(1)"
      Tab(3).Control(29)=   "ULabel27(28)"
      Tab(3).Control(30)=   "sdb_HGT_SP_EL_MAX(1)"
      Tab(3).Control(31)=   "sdb_HGT_SP_EL_MIN(1)"
      Tab(3).Control(32)=   "ULabel27(27)"
      Tab(3).Control(33)=   "sdb_HGT_RA_MAX(1)"
      Tab(3).Control(34)=   "sdb_HGT_RA_MIN(1)"
      Tab(3).Control(35)=   "ULabel27(14)"
      Tab(3).Control(36)=   "sdb_HGT_TS_MAX(1)"
      Tab(3).Control(37)=   "sdb_HGT_TS_MIN(1)"
      Tab(3).Control(38)=   "ULabel27(13)"
      Tab(3).Control(39)=   "sdb_HGT_YP_MAX(1)"
      Tab(3).Control(40)=   "sdb_HGT_YP_MIN(1)"
      Tab(3).Control(41)=   "ULabel27(12)"
      Tab(3).Control(42)=   "ULabel27(11)"
      Tab(3).Control(43)=   "ULabel27(10)"
      Tab(3).Control(44)=   "ULabel27(9)"
      Tab(3).Control(45)=   "ULabel27(8)"
      Tab(3).Control(46)=   "ULabel27(1)"
      Tab(3).Control(47)=   "txt_HTM_CD(3)"
      Tab(3).Control(48)=   "txt_HGT_EL_DSC_CD(1)"
      Tab(3).Control(49)=   "txt_HGT_SNPP_EL_DSC_CD(1)"
      Tab(3).Control(50)=   "txt_HGT_TS_DSC_CD(1)"
      Tab(3).Control(51)=   "txt_HGT_RA_DSC_CD(1)"
      Tab(3).Control(52)=   "txt_HGT_SP_EL_DSC_CD(1)"
      Tab(3).Control(53)=   "txt_HGT_YP_DSC_CD(1)"
      Tab(3).Control(54)=   "txt_HGT_SP_EL_CD(2)"
      Tab(3).Control(55)=   "txt_HGT_SP_EL_CD(3)"
      Tab(3).Control(56)=   "txt_HGT_SNPP_EL_CD(2)"
      Tab(3).Control(57)=   "txt_HGT_SNPP_EL_CD(3)"
      Tab(3).Control(58)=   "txt_HGT_EL_CD(2)"
      Tab(3).Control(59)=   "txt_HGT_EL_CD(3)"
      Tab(3).Control(60)=   "txt_HGT_SP_EL_SMP_CD(1)"
      Tab(3).Control(61)=   "txt_HGT_TENCIL_SMP_CD(1)"
      Tab(3).Control(62)=   "txt_HGT_TENCIL_TMP(1)"
      Tab(3).Control(63)=   "txt_HGT_TENCIL_TMP_UNIT(1)"
      Tab(3).Control(64)=   "txt_HGT_RA_DIR_CD(1)"
      Tab(3).Control(65)=   "txt_HGT_RA_DIR_NAME(1)"
      Tab(3).Control(66)=   "txt_SMP_WID(3)"
      Tab(3).Control(67)=   "txt_HGT_RA_DSC_CD(3)"
      Tab(3).ControlCount=   68
      TabCaption(4)   =   "�����ʱЧ"
      TabPicture(4)   =   "AQA0030C.frx":0850
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txt_A_TIM_IMPACT_DIR(1)"
      Tab(4).Control(1)=   "txt_A_TIM_IMPACT_DIR(0)"
      Tab(4).Control(2)=   "txt_A_TIM_IMPACT(2)"
      Tab(4).Control(3)=   "txt_A_TIM_IMPACT_KND(1)"
      Tab(4).Control(4)=   "txt_A_TIM_IMPACT_TMP"
      Tab(4).Control(5)=   "txt_A_TIM_IMPACT_TMP_UNIT"
      Tab(4).Control(6)=   "txt_A_TIM_IMPACT(1)"
      Tab(4).Control(7)=   "txt_A_TIM_IMPACT_KND(0)"
      Tab(4).Control(8)=   "txt_A_IMPACT_DIR(1)"
      Tab(4).Control(9)=   "txt_A_IMPACT_DIR(0)"
      Tab(4).Control(10)=   "txt_A_IMPACT(2)"
      Tab(4).Control(11)=   "txt_A_IMPACT(1)"
      Tab(4).Control(12)=   "txt_A_IMPACT_KND(0)"
      Tab(4).Control(13)=   "txt_A_IMPACT_KND(1)"
      Tab(4).Control(14)=   "txt_A_IMPACT_TMP"
      Tab(4).Control(15)=   "txt_A_IMPACT_TMP_UNIT"
      Tab(4).Control(16)=   "txt_IMPACT_TMP_UNIT"
      Tab(4).Control(17)=   "txt_IMPACT_TMP"
      Tab(4).Control(18)=   "txt_TIM_IMPACT_KND(0)"
      Tab(4).Control(19)=   "txt_TIM_IMPACT(1)"
      Tab(4).Control(20)=   "txt_TIM_IMPACT_TMP_UNIT"
      Tab(4).Control(21)=   "txt_TIM_IMPACT_TMP"
      Tab(4).Control(22)=   "txt_TIM_IMPACT_KND(1)"
      Tab(4).Control(23)=   "txt_IMPACT_KND(1)"
      Tab(4).Control(24)=   "txt_IMPACT_KND(0)"
      Tab(4).Control(25)=   "txt_IMPACT(1)"
      Tab(4).Control(26)=   "txt_IMPACT(2)"
      Tab(4).Control(27)=   "txt_TIM_IMPACT(2)"
      Tab(4).Control(28)=   "txt_IMPACT_DIR(0)"
      Tab(4).Control(29)=   "txt_IMPACT_DIR(1)"
      Tab(4).Control(30)=   "txt_TIM_IMPACT_DIR(0)"
      Tab(4).Control(31)=   "txt_TIM_IMPACT_DIR(1)"
      Tab(4).Control(32)=   "txt_HTM_CD(4)"
      Tab(4).Control(33)=   "txt_HTM_CD(5)"
      Tab(4).Control(34)=   "txt_HTM_CD(6)"
      Tab(4).Control(35)=   "txt_HTM_CD(7)"
      Tab(4).Control(36)=   "sdb_TIM_IMPACT_TIM"
      Tab(4).Control(37)=   "ULabel32(5)"
      Tab(4).Control(38)=   "ULabel32(4)"
      Tab(4).Control(39)=   "ULabel32(3)"
      Tab(4).Control(40)=   "ULabel32(2)"
      Tab(4).Control(41)=   "ULabel32(1)"
      Tab(4).Control(42)=   "ULabel32(6)"
      Tab(4).Control(43)=   "ULabel32(11)"
      Tab(4).Control(44)=   "sdb_IMPACT_MIN"
      Tab(4).Control(45)=   "sdb_IMPACT_AVE_MIN"
      Tab(4).Control(46)=   "sdb_IMPACT_RATE_MIN"
      Tab(4).Control(47)=   "sdb_IMPACT_RATE_MAX"
      Tab(4).Control(48)=   "sdb_TIM_IMPACT_MIN"
      Tab(4).Control(49)=   "sdb_TIM_IMPACT_AVE_MIN"
      Tab(4).Control(50)=   "sdb_TIM_IMPACT_RATE_MIN"
      Tab(4).Control(51)=   "sdb_TIM_IMPACT_RATE_MAX"
      Tab(4).Control(52)=   "ULabel32(16)"
      Tab(4).Control(53)=   "ULabel32(17)"
      Tab(4).Control(54)=   "ULabel32(18)"
      Tab(4).Control(55)=   "ULabel32(19)"
      Tab(4).Control(56)=   "ULabel4(32)"
      Tab(4).Control(57)=   "ULabel4(33)"
      Tab(4).Control(58)=   "ULabel4(34)"
      Tab(4).Control(59)=   "ULabel4(35)"
      Tab(4).Control(60)=   "ULabel4(36)"
      Tab(4).Control(61)=   "ULabel4(37)"
      Tab(4).Control(62)=   "ULabel4(38)"
      Tab(4).Control(63)=   "ULabel4(39)"
      Tab(4).Control(64)=   "ULabel4(40)"
      Tab(4).Control(65)=   "ULabel32(8)"
      Tab(4).Control(66)=   "ULabel32(9)"
      Tab(4).Control(67)=   "sdb_IMPACT_MIN_MIN"
      Tab(4).Control(68)=   "sdb_TIM_IMPACT_MIN_MIN"
      Tab(4).Control(69)=   "sdb_A_IMPACT_MIN"
      Tab(4).Control(70)=   "sdb_A_IMPACT_AVE_MIN"
      Tab(4).Control(71)=   "sdb_A_IMPACT_RATE_MIN"
      Tab(4).Control(72)=   "sdb_A_IMPACT_RATE_MAX"
      Tab(4).Control(73)=   "ULabel32(0)"
      Tab(4).Control(74)=   "ULabel32(7)"
      Tab(4).Control(75)=   "ULabel32(10)"
      Tab(4).Control(76)=   "sdb_A_IMPACT_MIN_MIN"
      Tab(4).Control(77)=   "ULabel32(12)"
      Tab(4).Control(78)=   "ULabel32(13)"
      Tab(4).Control(79)=   "ULabel32(14)"
      Tab(4).Control(80)=   "sdb_A_TIM_IMPACT_TIM"
      Tab(4).Control(81)=   "ULabel32(15)"
      Tab(4).Control(82)=   "ULabel32(20)"
      Tab(4).Control(83)=   "ULabel32(21)"
      Tab(4).Control(84)=   "ULabel32(22)"
      Tab(4).Control(85)=   "sdb_A_TIM_IMPACT_MIN"
      Tab(4).Control(86)=   "sdb_A_TIM_IMPACT_AVE_MIN"
      Tab(4).Control(87)=   "sdb_A_TIM_IMPACT_RATE_MIN"
      Tab(4).Control(88)=   "sdb_A_TIM_IMPACT_RATE_MAX"
      Tab(4).Control(89)=   "ULabel32(23)"
      Tab(4).Control(90)=   "ULabel32(24)"
      Tab(4).Control(91)=   "ULabel32(25)"
      Tab(4).Control(92)=   "sdb_A_TIM_IMPACT_MIN_MIN"
      Tab(4).Control(93)=   "ULabel4(99)"
      Tab(4).Control(94)=   "ULabel4(100)"
      Tab(4).Control(95)=   "ULabel4(101)"
      Tab(4).Control(96)=   "ULabel4(102)"
      Tab(4).Control(97)=   "ULabel32(27)"
      Tab(4).Control(98)=   "sidbEdit1"
      Tab(4).Control(99)=   "sdb_IMPACT_CP_MIN(0)"
      Tab(4).Control(100)=   "ULabel32(28)"
      Tab(4).Control(101)=   "sdb_IMPACT_CP_MIN(1)"
      Tab(4).Control(102)=   "ULabel32(29)"
      Tab(4).Control(103)=   "ULabel32(30)"
      Tab(4).Control(104)=   "sdb_IMPACT_CP_MIN(2)"
      Tab(4).Control(105)=   "sdb_IMPACT_CP_MIN(3)"
      Tab(4).Control(106)=   "Line49(8)"
      Tab(4).Control(107)=   "Line49(17)"
      Tab(4).Control(108)=   "Line49(16)"
      Tab(4).Control(109)=   "Line49(15)"
      Tab(4).Control(110)=   "Line3(14)"
      Tab(4).Control(111)=   "Line3(13)"
      Tab(4).Control(112)=   "Line3(12)"
      Tab(4).Control(113)=   "Line3(11)"
      Tab(4).Control(114)=   "Line3(10)"
      Tab(4).Control(115)=   "Line49(21)"
      Tab(4).ControlCount=   116
      TabCaption(5)   =   "Ӳ�ȡ�����"
      TabPicture(5)   =   "AQA0030C.frx":086C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txt_HTM_CD(11)"
      Tab(5).Control(1)=   "txt_HTM_CD(10)"
      Tab(5).Control(2)=   "txt_WLD_HARD_TYP(1)"
      Tab(5).Control(3)=   "txt_HARD_TYP(1)"
      Tab(5).Control(4)=   "txt_HARD_DSC_CD(0)"
      Tab(5).Control(5)=   "txt_WLD_BEND_DSC_CD"
      Tab(5).Control(6)=   "txt_WLD_HARD_DSC_CD"
      Tab(5).Control(7)=   "txt_RPT_BEND_DSC_CD"
      Tab(5).Control(8)=   "txt_BEND_DSC_CD(0)"
      Tab(5).Control(9)=   "txt_HARD_TYP(0)"
      Tab(5).Control(10)=   "txt_WLD_HARD_TYP(0)"
      Tab(5).Control(11)=   "txt_BEND_SMP_CD(0)"
      Tab(5).Control(12)=   "txt_RPT_BEND_SMP_CD"
      Tab(5).Control(13)=   "txt_WLD_HARD_UNIT"
      Tab(5).Control(14)=   "txt_HARD_TYP(3)"
      Tab(5).Control(15)=   "txt_HARD_DSC_CD(1)"
      Tab(5).Control(16)=   "txt_HARD_TYP(2)"
      Tab(5).Control(17)=   "txt_BEND_DSC_CD(1)"
      Tab(5).Control(18)=   "txt_BEND_SMP_CD(1)"
      Tab(5).Control(19)=   "txt_HTM_CD(8)"
      Tab(5).Control(20)=   "txt_HTM_CD(9)"
      Tab(5).Control(21)=   "txt_HTM_CD(13)"
      Tab(5).Control(22)=   "txt_HTM_CD(12)"
      Tab(5).Control(23)=   "sdb_HARD_MIN(0)"
      Tab(5).Control(24)=   "ULabel71(9)"
      Tab(5).Control(25)=   "ULabel71(22)"
      Tab(5).Control(26)=   "ULabel71(10)"
      Tab(5).Control(27)=   "ULabel71(11)"
      Tab(5).Control(28)=   "ULabel71(8)"
      Tab(5).Control(29)=   "ULabel71(41)"
      Tab(5).Control(30)=   "ULabel71(42)"
      Tab(5).Control(31)=   "ULabel71(40)"
      Tab(5).Control(32)=   "ULabel71(43)"
      Tab(5).Control(33)=   "sdb_HARD_MAX(0)"
      Tab(5).Control(34)=   "sdb_RPT_BEND_TMS"
      Tab(5).Control(35)=   "sdb_WLD_HARD_MAX"
      Tab(5).Control(36)=   "ULabel71(20)"
      Tab(5).Control(37)=   "sdb_BEND_DIA(0)"
      Tab(5).Control(38)=   "sdb_WLD_BEND_DIA"
      Tab(5).Control(39)=   "sdb_BEND_ANGLE(0)"
      Tab(5).Control(40)=   "sdb_WLD_BEND_ANG"
      Tab(5).Control(41)=   "sdb_WLD_HARD_MIN"
      Tab(5).Control(42)=   "UL_HARD_UNIT(0)"
      Tab(5).Control(43)=   "ULabel4(41)"
      Tab(5).Control(44)=   "ULabel4(42)"
      Tab(5).Control(45)=   "ULabel4(43)"
      Tab(5).Control(46)=   "ULabel4(44)"
      Tab(5).Control(47)=   "ULabel4(45)"
      Tab(5).Control(48)=   "ULabel4(46)"
      Tab(5).Control(49)=   "ULabel4(47)"
      Tab(5).Control(50)=   "ULabel4(48)"
      Tab(5).Control(51)=   "ULabel4(49)"
      Tab(5).Control(52)=   "sdb_HARD_MIN(1)"
      Tab(5).Control(53)=   "ULabel71(1)"
      Tab(5).Control(54)=   "sdb_HARD_MAX(1)"
      Tab(5).Control(55)=   "UL_HARD_UNIT(1)"
      Tab(5).Control(56)=   "ULabel71(2)"
      Tab(5).Control(57)=   "ULabel71(3)"
      Tab(5).Control(58)=   "ULabel71(4)"
      Tab(5).Control(59)=   "sdb_BEND_DIA(1)"
      Tab(5).Control(60)=   "sdb_BEND_ANGLE(1)"
      Tab(5).Control(61)=   "ULabel4(103)"
      Tab(5).Control(62)=   "ULabel4(104)"
      Tab(5).Control(63)=   "ULabel4(109)"
      Tab(5).Control(64)=   "ULabel4(110)"
      Tab(5).Control(65)=   "ULabel4(112)"
      Tab(5).Control(66)=   "ULabel4(113)"
      Tab(5).Control(67)=   "Line49(6)"
      Tab(5).Control(68)=   "Line3(19)"
      Tab(5).Control(69)=   "Line3(18)"
      Tab(5).Control(70)=   "Line3(17)"
      Tab(5).Control(71)=   "Line3(16)"
      Tab(5).Control(72)=   "Line3(15)"
      Tab(5).Control(73)=   "Line49(12)"
      Tab(5).Control(74)=   "Line49(13)"
      Tab(5).Control(75)=   "Line49(14)"
      Tab(5).Control(76)=   "Line49(18)"
      Tab(5).ControlCount=   77
      TabCaption(6)   =   "����"
      TabPicture(6)   =   "AQA0030C.frx":0888
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Line49(5)"
      Tab(6).Control(1)=   "Line3(25)"
      Tab(6).Control(2)=   "Line3(36)"
      Tab(6).Control(3)=   "Line3(37)"
      Tab(6).Control(4)=   "Line3(38)"
      Tab(6).Control(5)=   "sdb_A_DRAW_MIN(15)"
      Tab(6).Control(6)=   "sdb_A_DRAW_MAX(15)"
      Tab(6).Control(7)=   "sdb_A_DRAW_MIN(14)"
      Tab(6).Control(8)=   "sdb_A_DRAW_MAX(14)"
      Tab(6).Control(9)=   "sdb_A_DRAW_MIN(13)"
      Tab(6).Control(10)=   "sdb_A_DRAW_MAX(13)"
      Tab(6).Control(11)=   "sdb_A_DRAW_MIN(12)"
      Tab(6).Control(12)=   "sdb_A_DRAW_MAX(12)"
      Tab(6).Control(13)=   "sdb_A_DRAW_MIN(16)"
      Tab(6).Control(14)=   "sdb_A_DRAW_MAX(16)"
      Tab(6).Control(15)=   "ULabel71(29)"
      Tab(6).Control(16)=   "ULabel71(28)"
      Tab(6).Control(17)=   "ULabel71(27)"
      Tab(6).Control(18)=   "ULabel71(26)"
      Tab(6).Control(19)=   "ULabel71(25)"
      Tab(6).Control(20)=   "ULabel71(7)"
      Tab(6).Control(21)=   "ULabel71(6)"
      Tab(6).Control(22)=   "sdb_sc_max"
      Tab(6).Control(23)=   "ULabel71(5)"
      Tab(6).Control(24)=   "ULabel4(122)"
      Tab(6).Control(25)=   "ULabel4(121)"
      Tab(6).Control(26)=   "ULabel4(120)"
      Tab(6).Control(27)=   "ULabel4(119)"
      Tab(6).Control(28)=   "ULabel4(118)"
      Tab(6).Control(29)=   "ULabel4(117)"
      Tab(6).Control(30)=   "ULabel4(116)"
      Tab(6).Control(31)=   "ULabel4(115)"
      Tab(6).Control(32)=   "ULabel4(114)"
      Tab(6).Control(33)=   "ULabel4(108)"
      Tab(6).Control(34)=   "ULabel4(107)"
      Tab(6).Control(35)=   "ULabel4(106)"
      Tab(6).Control(36)=   "ULabel4(105)"
      Tab(6).Control(37)=   "ULabel32(26)"
      Tab(6).Control(38)=   "ULabel71(0)"
      Tab(6).Control(39)=   "ULabel71(24)"
      Tab(6).Control(40)=   "sdb_SSCC_YP_TIM"
      Tab(6).Control(41)=   "ULabel71(47)"
      Tab(6).Control(42)=   "sdb_HIC_CWR_MAX"
      Tab(6).Control(43)=   "sdb_HIC_CLR_MAX"
      Tab(6).Control(44)=   "sdb_HIC_CSR_MAX"
      Tab(6).Control(45)=   "ULabel71(23)"
      Tab(6).Control(46)=   "ULabel71(21)"
      Tab(6).Control(47)=   "ULabel71(19)"
      Tab(6).Control(48)=   "ULabel71(18)"
      Tab(6).Control(49)=   "sdb_SSCC_YP_MAX"
      Tab(6).Control(50)=   "sdb_JOMINY_MAX"
      Tab(6).Control(51)=   "sdb_JOMINY_MIN"
      Tab(6).Control(52)=   "sdb_DWTT_YP_AVE"
      Tab(6).Control(53)=   "sdb_DWTT_YP_MIN"
      Tab(6).Control(54)=   "ULabel99(7)"
      Tab(6).Control(55)=   "ULabel71(45)"
      Tab(6).Control(56)=   "ULabel71(44)"
      Tab(6).Control(57)=   "ULabel71(16)"
      Tab(6).Control(58)=   "ULabel71(15)"
      Tab(6).Control(59)=   "ULabel71(14)"
      Tab(6).Control(60)=   "ULabel71(13)"
      Tab(6).Control(61)=   "ULabel71(12)"
      Tab(6).Control(62)=   "ULabel71(17)"
      Tab(6).Control(63)=   "sdb_JOMINY_DIST"
      Tab(6).Control(64)=   "txt_UST_GRD_NAME"
      Tab(6).Control(65)=   "txt_UST_GRD"
      Tab(6).Control(66)=   "txt_HIC_SVT_KND(1)"
      Tab(6).Control(67)=   "txt_JOMINY_TYP(1)"
      Tab(6).Control(68)=   "txt_SSCC_SVT_KND(1)"
      Tab(6).Control(69)=   "txt_UST_STD_CD(1)"
      Tab(6).Control(70)=   "txt_DWTT_TMP_UNIT"
      Tab(6).Control(71)=   "txt_DWTT_TMP"
      Tab(6).Control(72)=   "txt_UST_DSC_CD"
      Tab(6).Control(73)=   "txt_FOAT_DSC_CD"
      Tab(6).Control(74)=   "txt_DWTT_DSC_CD(0)"
      Tab(6).Control(75)=   "txt_SSCC_DSC_CD"
      Tab(6).Control(76)=   "txt_HIC_DSC_CD"
      Tab(6).Control(77)=   "txt_JOMINY_DSC_CD"
      Tab(6).Control(78)=   "txt_HIC_SVT_KND(0)"
      Tab(6).Control(79)=   "txt_JOMINY_TYP(0)"
      Tab(6).Control(80)=   "txt_SSCC_SVT_KND(0)"
      Tab(6).Control(81)=   "txt_FOAT_SMP_CD"
      Tab(6).Control(82)=   "txt_HIC_SMP_CD"
      Tab(6).Control(83)=   "txt_JOMINY_SMP_CD"
      Tab(6).Control(84)=   "txt_SSCC_SMP_CD"
      Tab(6).Control(85)=   "txt_UST_STD_CD(0)"
      Tab(6).Control(86)=   "txt_DWTT_SMP_CD"
      Tab(6).Control(87)=   "txt_HTM_CD(17)"
      Tab(6).Control(88)=   "txt_HTM_CD(16)"
      Tab(6).Control(89)=   "txt_HTM_CD(15)"
      Tab(6).Control(90)=   "txt_HTM_CD(14)"
      Tab(6).Control(91)=   "txt_sc_dsc_cd"
      Tab(6).Control(92)=   "txt_DWTT_DSC_CD(1)"
      Tab(6).Control(93)=   "txt_A_DRAW_DSC_CD(16)"
      Tab(6).Control(94)=   "txt_A_DRAW_DSC_CD(15)"
      Tab(6).Control(95)=   "txt_A_DRAW_DSC_CD(14)"
      Tab(6).Control(96)=   "txt_A_DRAW_DSC_CD(13)"
      Tab(6).Control(97)=   "txt_A_DRAW_DSC_CD(12)"
      Tab(6).Control(98)=   "txt_STRESS_KND(11)"
      Tab(6).Control(99)=   "txt_STRESS_KND(10)"
      Tab(6).Control(100)=   "txt_STRESS_KND(9)"
      Tab(6).Control(101)=   "txt_STRESS_KND(8)"
      Tab(6).Control(102)=   "txt_STRESS_KND(7)"
      Tab(6).Control(103)=   "txt_STRESS_KND(6)"
      Tab(6).Control(104)=   "txt_STRESS_KND(5)"
      Tab(6).Control(105)=   "txt_STRESS_KND(4)"
      Tab(6).Control(106)=   "txt_STRESS_KND(3)"
      Tab(6).Control(107)=   "txt_STRESS_KND(2)"
      Tab(6).ControlCount=   108
      TabCaption(7)   =   "����"
      TabPicture(7)   =   "AQA0030C.frx":08A4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "txt_HTM_CD(18)"
      Tab(7).Control(1)=   "txt_GRAIN_SIZE_MTH(1)"
      Tab(7).Control(2)=   "txt_RMV_CAR_TYP(0)"
      Tab(7).Control(3)=   "txt_RMV_CAR_TYP(1)"
      Tab(7).Control(4)=   "txt_RMV_CAR_SMP_CD"
      Tab(7).Control(5)=   "txt_GRAIN_SIZE_DSC_CD"
      Tab(7).Control(6)=   "txt_RMV_CAR_DSC_CD"
      Tab(7).Control(7)=   "txt_GRAIN_SIZE_TMP_UNIT"
      Tab(7).Control(8)=   "txt_BELT_STR_DSC_CD"
      Tab(7).Control(9)=   "txt_OST_GRAIN_CHA"
      Tab(7).Control(10)=   "txt_OST_GRAIN_SIZE_TMP_UNIT"
      Tab(7).Control(11)=   "txt_OST_GRAIN_SIZE_DSC_CD"
      Tab(7).Control(12)=   "txt_GRAIN_SIZE_MTH(0)"
      Tab(7).Control(13)=   "txt_OST_GRAIN_SIZE_MTH(0)"
      Tab(7).Control(14)=   "txt_OST_GRAIN_SIZE_MTH(1)"
      Tab(7).Control(15)=   "ULabel7"
      Tab(7).Control(16)=   "ULabel6"
      Tab(7).Control(17)=   "txt_BELT_STR_GRD"
      Tab(7).Control(18)=   "ULabel27(21)"
      Tab(7).Control(19)=   "ULabel87(15)"
      Tab(7).Control(20)=   "ULabel87(18)"
      Tab(7).Control(21)=   "ULabel87(21)"
      Tab(7).Control(22)=   "sdb_RMV_CAR_MAX"
      Tab(7).Control(23)=   "sdb_GRAIN_SIZE_MIN"
      Tab(7).Control(24)=   "sdb_GRAIN_SIZE_TIME"
      Tab(7).Control(25)=   "ULabel27(26)"
      Tab(7).Control(26)=   "ULabel4(50)"
      Tab(7).Control(27)=   "ULabel4(51)"
      Tab(7).Control(28)=   "ULabel4(52)"
      Tab(7).Control(29)=   "ULabel4(53)"
      Tab(7).Control(30)=   "ULabel4(54)"
      Tab(7).Control(31)=   "ULabel4(55)"
      Tab(7).Control(32)=   "ULabel4(56)"
      Tab(7).Control(33)=   "ULabel4(57)"
      Tab(7).Control(34)=   "ULabel4(58)"
      Tab(7).Control(35)=   "sdb_GRAIN_SIZE_MAX"
      Tab(7).Control(36)=   "ULabel27(0)"
      Tab(7).Control(37)=   "ULabel87(3)"
      Tab(7).Control(38)=   "sdb_GRAIN_SIZE_TMP"
      Tab(7).Control(39)=   "sdb_OST_GRAIN_SIZE_TMP"
      Tab(7).Control(40)=   "sdb_OST_GRAIN_SIZE_TIME"
      Tab(7).Control(41)=   "ULabel87(7)"
      Tab(7).Control(42)=   "sdb_OST_GRAIN_SIZE_MIN"
      Tab(7).Control(43)=   "sdb_OST_GRAIN_SIZE_MAX"
      Tab(7).Control(44)=   "ULabel4(132)"
      Tab(7).Control(45)=   "ULabel27(31)"
      Tab(7).Control(46)=   "Line49(1)"
      Tab(7).Control(47)=   "Line49(27)"
      Tab(7).Control(48)=   "Line3(20)"
      Tab(7).Control(49)=   "Line3(22)"
      Tab(7).Control(50)=   "Line3(23)"
      Tab(7).Control(51)=   "Line3(24)"
      Tab(7).ControlCount=   52
      TabCaption(8)   =   "�ǽ�������"
      TabPicture(8)   =   "AQA0030C.frx":08C0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Line3(21)"
      Tab(8).Control(1)=   "Line49(0)"
      Tab(8).Control(2)=   "Line3(39)"
      Tab(8).Control(3)=   "Line3(40)"
      Tab(8).Control(4)=   "Line3(41)"
      Tab(8).Control(5)=   "ULabel4(134)"
      Tab(8).Control(6)=   "ULabel4(133)"
      Tab(8).Control(7)=   "sdb_ACD_DFT_GRD5"
      Tab(8).Control(8)=   "sdb_ACD_DFT_GRD4"
      Tab(8).Control(9)=   "ULabel87(0)"
      Tab(8).Control(10)=   "sdb_ACD_DFT_GRD1"
      Tab(8).Control(11)=   "sdb_S_PRINT_DRG"
      Tab(8).Control(12)=   "sdb_ACD_DFT_GRD3"
      Tab(8).Control(13)=   "sdb_ACD_DFT_GRD2"
      Tab(8).Control(14)=   "ULabel87(12)"
      Tab(8).Control(15)=   "ULabel87(11)"
      Tab(8).Control(16)=   "ULabel87(6)"
      Tab(8).Control(17)=   "ULabel87(5)"
      Tab(8).Control(18)=   "ULabel87(4)"
      Tab(8).Control(19)=   "ULabel27(24)"
      Tab(8).Control(20)=   "ULabel27(23)"
      Tab(8).Control(21)=   "ULabel27(22)"
      Tab(8).Control(22)=   "ULabel4(131)"
      Tab(8).Control(23)=   "ULabel4(130)"
      Tab(8).Control(24)=   "ULabel4(129)"
      Tab(8).Control(25)=   "ULabel4(128)"
      Tab(8).Control(26)=   "ULabel4(127)"
      Tab(8).Control(27)=   "ULabel4(126)"
      Tab(8).Control(28)=   "ULabel4(125)"
      Tab(8).Control(29)=   "ULabel4(124)"
      Tab(8).Control(30)=   "ULabel4(123)"
      Tab(8).Control(31)=   "sdb_TIN_GRD"
      Tab(8).Control(32)=   "ULabel87(14)"
      Tab(8).Control(33)=   "sdb_DS_GRD"
      Tab(8).Control(34)=   "ULabel87(8)"
      Tab(8).Control(35)=   "sdb_NON_METAL_BGRD4"
      Tab(8).Control(36)=   "sdb_NON_METAL_BGRD3"
      Tab(8).Control(37)=   "sdb_NON_METAL_BGRD2"
      Tab(8).Control(38)=   "sdb_NON_METAL_BGRD1"
      Tab(8).Control(39)=   "sdb_NON_METAL_AGRD4"
      Tab(8).Control(40)=   "sdb_NON_METAL_AGRD3"
      Tab(8).Control(41)=   "sdb_NON_METAL_AGRD2"
      Tab(8).Control(42)=   "sdb_NON_METAL_AGRD1"
      Tab(8).Control(43)=   "ULabel87(22)"
      Tab(8).Control(44)=   "ULabel87(2)"
      Tab(8).Control(45)=   "ULabel87(1)"
      Tab(8).Control(46)=   "ULabel27(25)"
      Tab(8).Control(47)=   "txt_NON_METAL_TYP(0)"
      Tab(8).Control(48)=   "txt_NON_METAL_TYP(1)"
      Tab(8).Control(49)=   "txt_NON_METAL_SMP_CD"
      Tab(8).Control(50)=   "txt_NON_METAL_DSC_CD"
      Tab(8).Control(51)=   "txt_NON_METAL_ACD3(1)"
      Tab(8).Control(52)=   "txt_NON_METAL_ACD2(1)"
      Tab(8).Control(53)=   "txt_NON_METAL_ACD1(1)"
      Tab(8).Control(54)=   "txt_NON_METAL_ACD3(0)"
      Tab(8).Control(55)=   "txt_NON_METAL_ACD2(0)"
      Tab(8).Control(56)=   "txt_NON_METAL_ACD1(0)"
      Tab(8).Control(57)=   "txt_NON_METAL_ACD4(1)"
      Tab(8).Control(58)=   "txt_NON_METAL_ACD4(0)"
      Tab(8).Control(59)=   "txt_NON_METAL_BCD3(1)"
      Tab(8).Control(60)=   "txt_NON_METAL_BCD2(1)"
      Tab(8).Control(61)=   "txt_NON_METAL_BCD1(1)"
      Tab(8).Control(62)=   "txt_NON_METAL_BCD3(0)"
      Tab(8).Control(63)=   "txt_NON_METAL_BCD2(0)"
      Tab(8).Control(64)=   "txt_NON_METAL_BCD1(0)"
      Tab(8).Control(65)=   "txt_NON_METAL_BCD4(1)"
      Tab(8).Control(66)=   "txt_NON_METAL_BCD4(0)"
      Tab(8).Control(67)=   "txt_FRACT_SMP_CD"
      Tab(8).Control(68)=   "txt_ACD_DFT_TYP1(0)"
      Tab(8).Control(69)=   "txt_ACD_DFT_TYP2(0)"
      Tab(8).Control(70)=   "txt_ACD_DFT_TYP3(0)"
      Tab(8).Control(71)=   "txt_ACD_DFT_TYP1(1)"
      Tab(8).Control(72)=   "txt_ACD_DFT_TYP2(1)"
      Tab(8).Control(73)=   "txt_ACD_DFT_TYP3(1)"
      Tab(8).Control(74)=   "txt_FRACT_NAME_CD1(0)"
      Tab(8).Control(75)=   "txt_FRACT_NAME_CD2(0)"
      Tab(8).Control(76)=   "txt_FRACT_NAME_CD3(0)"
      Tab(8).Control(77)=   "txt_FRACT_NAME_CD1(1)"
      Tab(8).Control(78)=   "txt_FRACT_NAME_CD2(1)"
      Tab(8).Control(79)=   "txt_FRACT_NAME_CD3(1)"
      Tab(8).Control(80)=   "txt_S_PRINT_DSC_CD"
      Tab(8).Control(81)=   "txt_ACD_DSC_CD"
      Tab(8).Control(82)=   "txt_FRACT_DSC_CD"
      Tab(8).Control(83)=   "txt_FRACT_GRD1"
      Tab(8).Control(84)=   "txt_FRACT_GRD2"
      Tab(8).Control(85)=   "txt_FRACT_GRD3"
      Tab(8).Control(86)=   "txt_ACD_DFT_TYP5(1)"
      Tab(8).Control(87)=   "txt_ACD_DFT_TYP4(1)"
      Tab(8).Control(88)=   "txt_ACD_DFT_TYP5(0)"
      Tab(8).Control(89)=   "txt_ACD_DFT_TYP4(0)"
      Tab(8).Control(90)=   "txt_FRACT_NAME_CD5(1)"
      Tab(8).Control(91)=   "txt_FRACT_NAME_CD4(1)"
      Tab(8).Control(92)=   "txt_FRACT_NAME_CD5(0)"
      Tab(8).Control(93)=   "txt_FRACT_NAME_CD4(0)"
      Tab(8).Control(94)=   "txt_FRACT_GRD5"
      Tab(8).Control(95)=   "txt_FRACT_GRD4"
      Tab(8).Control(96)=   "txt_HTM_CD(19)"
      Tab(8).Control(97)=   "txt_HTM_CD(20)"
      Tab(8).ControlCount=   98
      TabCaption(9)   =   "���û�������Ŀ"
      TabPicture(9)   =   "AQA0030C.frx":08DC
      Tab(9).ControlEnabled=   -1  'True
      Tab(9).Control(0)=   "ss1"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).ControlCount=   1
      Begin VB.TextBox txt_STRESS_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -70680
         MaxLength       =   10
         TabIndex        =   418
         Top             =   6780
         Width           =   495
      End
      Begin VB.TextBox txt_STRESS_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   -70680
         MaxLength       =   10
         TabIndex        =   417
         Top             =   7140
         Width           =   495
      End
      Begin VB.TextBox txt_STRESS_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   -70680
         MaxLength       =   10
         TabIndex        =   416
         Top             =   7500
         Width           =   495
      End
      Begin VB.TextBox txt_STRESS_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   -70680
         MaxLength       =   10
         TabIndex        =   415
         Top             =   7860
         Width           =   495
      End
      Begin VB.TextBox txt_STRESS_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   -70680
         MaxLength       =   10
         TabIndex        =   414
         Top             =   8220
         Width           =   495
      End
      Begin VB.TextBox txt_STRESS_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   -70920
         MaxLength       =   10
         TabIndex        =   413
         Top             =   8580
         Width           =   735
      End
      Begin VB.TextBox txt_STRESS_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   -70080
         MaxLength       =   10
         TabIndex        =   412
         Top             =   8580
         Width           =   735
      End
      Begin VB.TextBox txt_STRESS_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   -69240
         MaxLength       =   10
         TabIndex        =   411
         Top             =   8580
         Width           =   615
      End
      Begin VB.TextBox txt_STRESS_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   -68520
         MaxLength       =   10
         TabIndex        =   410
         Top             =   8580
         Width           =   615
      End
      Begin VB.TextBox txt_STRESS_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   -67800
         MaxLength       =   10
         TabIndex        =   409
         Top             =   8580
         Width           =   615
      End
      Begin VB.TextBox txt_A_DRAW_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   408
         Top             =   6780
         Width           =   900
      End
      Begin VB.TextBox txt_A_DRAW_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   407
         Top             =   7140
         Width           =   900
      End
      Begin VB.TextBox txt_A_DRAW_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   406
         Top             =   7500
         Width           =   900
      End
      Begin VB.TextBox txt_A_DRAW_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   15
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   405
         Top             =   7860
         Width           =   900
      End
      Begin VB.TextBox txt_A_DRAW_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   16
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   404
         Top             =   8220
         Width           =   900
      End
      Begin VB.TextBox txt_A_DRAW_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   403
         Top             =   8100
         Width           =   900
      End
      Begin VB.TextBox txt_DRAW_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   402
         Top             =   8340
         Width           =   900
      End
      Begin VB.TextBox txt_DWTT_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   397
         Top             =   5340
         Width           =   900
      End
      Begin VB.TextBox txt_sc_dsc_cd 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   396
         Top             =   6420
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_RA_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   -60840
         MaxLength       =   1
         TabIndex        =   389
         Tag             =   "Q0002"
         Top             =   3540
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_RA_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -60720
         MaxLength       =   1
         TabIndex        =   387
         Tag             =   "Q0002"
         Top             =   3420
         Width           =   900
      End
      Begin VB.TextBox txt_RA_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   -60840
         MaxLength       =   1
         TabIndex        =   385
         Top             =   3300
         Width           =   900
      End
      Begin VB.TextBox txt_RA_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -60720
         MaxLength       =   1
         TabIndex        =   383
         Top             =   3300
         Width           =   900
      End
      Begin VB.TextBox txt_SMP_WID 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   -72930
         TabIndex        =   376
         Top             =   1785
         Width           =   1980
      End
      Begin VB.TextBox txt_SMP_WID 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -72930
         TabIndex        =   375
         Top             =   1785
         Width           =   1980
      End
      Begin VB.TextBox txt_SMP_WID 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -72930
         TabIndex        =   374
         Top             =   1785
         Width           =   1980
      End
      Begin VB.TextBox txt_SMP_WID 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -72915
         TabIndex        =   373
         Top             =   1860
         Width           =   1980
      End
      Begin VB.TextBox txt_HGT_RA_DIR_NAME 
         Height          =   300
         Index           =   1
         Left            =   -70410
         TabIndex        =   372
         Top             =   2910
         Width           =   1470
      End
      Begin VB.TextBox txt_HGT_RA_DIR_CD 
         Height          =   300
         Index           =   1
         Left            =   -70920
         TabIndex        =   371
         Top             =   2910
         Width           =   495
      End
      Begin VB.TextBox txt_HGT_RA_DIR_NAME 
         Height          =   300
         Index           =   0
         Left            =   -70440
         TabIndex        =   370
         Top             =   3030
         Width           =   1470
      End
      Begin VB.TextBox txt_HGT_RA_DIR_CD 
         Height          =   300
         Index           =   0
         Left            =   -70920
         TabIndex        =   369
         Top             =   3030
         Width           =   495
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   20
         Left            =   -72915
         TabIndex        =   368
         Top             =   3810
         Width           =   1950
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   19
         Left            =   -72915
         TabIndex        =   367
         Top             =   1470
         Width           =   1950
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   18
         Left            =   -72915
         TabIndex        =   366
         Top             =   1065
         Width           =   1980
      End
      Begin VB.TextBox txt_FRACT_GRD4 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -62670
         MaxLength       =   1
         TabIndex        =   359
         Top             =   6900
         Width           =   900
      End
      Begin VB.TextBox txt_FRACT_GRD5 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -62670
         MaxLength       =   1
         TabIndex        =   358
         Top             =   7185
         Width           =   900
      End
      Begin VB.TextBox txt_FRACT_NAME_CD4 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   2
         TabIndex        =   357
         Top             =   6900
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_NAME_CD5 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   2
         TabIndex        =   356
         Top             =   7185
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_NAME_CD4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   355
         Top             =   6900
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   354
         Top             =   7185
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP4 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   2
         TabIndex        =   353
         Top             =   5430
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP5 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   2
         TabIndex        =   352
         Top             =   5715
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   351
         Top             =   5430
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   350
         Top             =   5715
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_GRD3 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -62670
         MaxLength       =   1
         TabIndex        =   349
         Top             =   6630
         Width           =   900
      End
      Begin VB.TextBox txt_FRACT_GRD2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -62670
         MaxLength       =   1
         TabIndex        =   348
         Top             =   6345
         Width           =   900
      End
      Begin VB.TextBox txt_FRACT_GRD1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -62670
         MaxLength       =   1
         TabIndex        =   347
         Top             =   6060
         Width           =   900
      End
      Begin VB.TextBox txt_FRACT_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   346
         Top             =   6060
         Width           =   900
      End
      Begin VB.TextBox txt_ACD_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   345
         Top             =   4560
         Width           =   900
      End
      Begin VB.TextBox txt_S_PRINT_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   344
         Top             =   4170
         Width           =   900
      End
      Begin VB.TextBox txt_FRACT_NAME_CD3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   343
         Top             =   6630
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   342
         Top             =   6345
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   341
         Top             =   6060
         Width           =   1605
      End
      Begin VB.TextBox txt_FRACT_NAME_CD3 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   2
         TabIndex        =   340
         Top             =   6630
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_NAME_CD2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   2
         TabIndex        =   339
         Top             =   6345
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_NAME_CD1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   2
         TabIndex        =   338
         Top             =   6060
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   337
         Top             =   5130
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   336
         Top             =   4845
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   335
         Top             =   4560
         Width           =   1605
      End
      Begin VB.TextBox txt_ACD_DFT_TYP3 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   2
         TabIndex        =   334
         Top             =   5130
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   2
         TabIndex        =   333
         Top             =   4845
         Width           =   435
      End
      Begin VB.TextBox txt_ACD_DFT_TYP1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   2
         TabIndex        =   332
         Top             =   4560
         Width           =   435
      End
      Begin VB.TextBox txt_FRACT_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72915
         MaxLength       =   9
         TabIndex        =   331
         Top             =   6060
         Width           =   1950
      End
      Begin VB.TextBox txt_NON_METAL_BCD4 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   1
         TabIndex        =   320
         Top             =   3060
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_BCD4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   319
         Top             =   3060
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_BCD1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   1
         TabIndex        =   318
         Top             =   2220
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_BCD2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   1
         TabIndex        =   317
         Top             =   2505
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_BCD3 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   1
         TabIndex        =   316
         Top             =   2790
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_BCD1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   315
         Top             =   2220
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_BCD2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   314
         Top             =   2505
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_BCD3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   313
         Top             =   2790
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_ACD4 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   1
         TabIndex        =   312
         Top             =   1950
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_ACD4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   311
         Top             =   1950
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_ACD1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   1
         TabIndex        =   310
         Top             =   1110
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_ACD2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   1
         TabIndex        =   309
         Top             =   1395
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_ACD3 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -65730
         MaxLength       =   1
         TabIndex        =   308
         Top             =   1680
         Width           =   435
      End
      Begin VB.TextBox txt_NON_METAL_ACD1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   307
         Top             =   1110
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_ACD2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   306
         Top             =   1395
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_ACD3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -65280
         MaxLength       =   80
         TabIndex        =   305
         Top             =   1680
         Width           =   1605
      End
      Begin VB.TextBox txt_NON_METAL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   304
         Top             =   1110
         Width           =   900
      End
      Begin VB.TextBox txt_NON_METAL_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72915
         MaxLength       =   9
         TabIndex        =   303
         Top             =   1110
         Width           =   1950
      End
      Begin VB.TextBox txt_NON_METAL_TYP 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70410
         MaxLength       =   80
         TabIndex        =   302
         Top             =   1110
         Width           =   1470
      End
      Begin VB.TextBox txt_NON_METAL_TYP 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70920
         MaxLength       =   1
         TabIndex        =   301
         Top             =   1110
         Width           =   495
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   -72870
         TabIndex        =   300
         Top             =   3900
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   -72870
         TabIndex        =   299
         Top             =   3090
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   -72885
         TabIndex        =   288
         Top             =   1935
         Width           =   1950
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   15
         Left            =   -72885
         TabIndex        =   287
         Top             =   2880
         Width           =   1950
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   16
         Left            =   -72885
         TabIndex        =   286
         Top             =   3735
         Width           =   1950
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   17
         Left            =   -72885
         TabIndex        =   285
         Top             =   5700
         Width           =   1950
      End
      Begin VB.TextBox txt_DWTT_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72885
         MaxLength       =   9
         TabIndex        =   284
         Top             =   4935
         Width           =   1950
      End
      Begin VB.TextBox txt_UST_STD_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   4
         TabIndex        =   283
         Top             =   1110
         Width           =   495
      End
      Begin VB.TextBox txt_SSCC_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72885
         MaxLength       =   9
         TabIndex        =   282
         Top             =   4515
         Width           =   1950
      End
      Begin VB.TextBox txt_JOMINY_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72885
         MaxLength       =   9
         TabIndex        =   281
         Top             =   2385
         Width           =   1950
      End
      Begin VB.TextBox txt_HIC_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72885
         MaxLength       =   9
         TabIndex        =   280
         Top             =   3330
         Width           =   1950
      End
      Begin VB.TextBox txt_FOAT_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72885
         MaxLength       =   9
         TabIndex        =   279
         Top             =   1515
         Width           =   1950
      End
      Begin VB.TextBox txt_SSCC_SVT_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   278
         Top             =   4515
         Width           =   495
      End
      Begin VB.TextBox txt_JOMINY_TYP 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   277
         Top             =   2385
         Width           =   495
      End
      Begin VB.TextBox txt_HIC_SVT_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   276
         Top             =   3330
         Width           =   495
      End
      Begin VB.TextBox txt_JOMINY_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   275
         Top             =   2385
         Width           =   900
      End
      Begin VB.TextBox txt_HIC_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   274
         Top             =   3330
         Width           =   900
      End
      Begin VB.TextBox txt_SSCC_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   273
         Top             =   4515
         Width           =   900
      End
      Begin VB.TextBox txt_DWTT_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   272
         Top             =   4935
         Width           =   900
      End
      Begin VB.TextBox txt_FOAT_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   271
         Top             =   1515
         Width           =   900
      End
      Begin VB.TextBox txt_UST_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   270
         Top             =   1110
         Width           =   900
      End
      Begin VB.TextBox txt_DWTT_TMP 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68865
         MaxLength       =   3
         TabIndex        =   269
         Top             =   4935
         Width           =   1485
      End
      Begin VB.TextBox txt_DWTT_TMP_UNIT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -67380
         MaxLength       =   1
         TabIndex        =   268
         Top             =   4935
         Width           =   510
      End
      Begin VB.TextBox txt_UST_STD_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70350
         MaxLength       =   80
         TabIndex        =   267
         Top             =   1110
         Width           =   3465
      End
      Begin VB.TextBox txt_SSCC_SVT_KND 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   266
         Top             =   4515
         Width           =   1470
      End
      Begin VB.TextBox txt_JOMINY_TYP 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   265
         Top             =   2385
         Width           =   1470
      End
      Begin VB.TextBox txt_HIC_SVT_KND 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   264
         Top             =   3330
         Width           =   1470
      End
      Begin VB.TextBox txt_UST_GRD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -63570
         MaxLength       =   1
         TabIndex        =   263
         Tag             =   "txt_UST_GRD_NAME,Q0053"
         Top             =   1110
         Width           =   900
      End
      Begin VB.TextBox txt_UST_GRD_NAME 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -63570
         MaxLength       =   80
         TabIndex        =   262
         Top             =   1515
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txt_SP_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   161
         Top             =   7650
         Width           =   495
      End
      Begin VB.TextBox txt_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   160
         Top             =   4275
         Width           =   495
      End
      Begin VB.TextBox txt_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   159
         Tag             =   "N"
         Top             =   4275
         Width           =   1470
      End
      Begin VB.TextBox txt_SP_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   158
         Tag             =   "N"
         Top             =   7650
         Width           =   1470
      End
      Begin VB.TextBox txt_SG_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   157
         Top             =   6714
         Width           =   495
      End
      Begin VB.TextBox txt_SG_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   156
         Tag             =   "N"
         Top             =   6714
         Width           =   1470
      End
      Begin VB.TextBox txt_SNPP_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   155
         Top             =   5780
         Width           =   495
      End
      Begin VB.TextBox txt_SNPP_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   154
         Tag             =   "N"
         Top             =   5780
         Width           =   1470
      End
      Begin VB.TextBox txt_SP_EL_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -72915
         MaxLength       =   9
         TabIndex        =   153
         Top             =   7650
         Width           =   1980
      End
      Begin VB.TextBox txt_TENCIL_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -72915
         MaxLength       =   9
         TabIndex        =   152
         Top             =   1110
         Width           =   1980
      End
      Begin VB.TextBox txt_YP_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   151
         Top             =   1110
         Width           =   900
      End
      Begin VB.TextBox txt_SP_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   150
         Top             =   7650
         Width           =   900
      End
      Begin VB.TextBox txt_SG_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   149
         Top             =   6714
         Width           =   900
      End
      Begin VB.TextBox txt_SNPP_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   148
         Top             =   5780
         Width           =   900
      End
      Begin VB.TextBox txt_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   147
         Top             =   4275
         Width           =   900
      End
      Begin VB.TextBox txt_RA_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   146
         Top             =   2745
         Width           =   900
      End
      Begin VB.TextBox txt_TS_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   145
         Top             =   2265
         Width           =   900
      End
      Begin VB.TextBox txt_YR_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   144
         Top             =   4846
         Width           =   900
      End
      Begin VB.TextBox txt_RA_DIR_CD 
         Height          =   300
         Index           =   0
         Left            =   -70890
         TabIndex        =   143
         Top             =   2745
         Width           =   495
      End
      Begin VB.TextBox txt_RA_DIR_NAME 
         Height          =   300
         Index           =   0
         Left            =   -70380
         TabIndex        =   142
         Top             =   2745
         Width           =   1470
      End
      Begin VB.TextBox txt_WLD_HARD_TYP 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70350
         MaxLength       =   80
         TabIndex        =   141
         Top             =   4830
         Width           =   1470
      End
      Begin VB.TextBox txt_HARD_TYP 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70350
         MaxLength       =   80
         TabIndex        =   140
         Top             =   1020
         Width           =   1470
      End
      Begin VB.TextBox txt_HARD_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60750
         MaxLength       =   1
         TabIndex        =   139
         Top             =   1020
         Width           =   900
      End
      Begin VB.TextBox txt_WLD_BEND_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60750
         MaxLength       =   1
         TabIndex        =   138
         Top             =   5730
         Width           =   900
      End
      Begin VB.TextBox txt_WLD_HARD_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60750
         MaxLength       =   1
         TabIndex        =   137
         Top             =   4830
         Width           =   900
      End
      Begin VB.TextBox txt_RPT_BEND_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60750
         MaxLength       =   1
         TabIndex        =   136
         Top             =   4410
         Width           =   900
      End
      Begin VB.TextBox txt_BEND_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60750
         MaxLength       =   1
         TabIndex        =   135
         Top             =   2730
         Width           =   900
      End
      Begin VB.TextBox txt_HARD_TYP 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70860
         MaxLength       =   1
         TabIndex        =   134
         Top             =   1020
         Width           =   495
      End
      Begin VB.TextBox txt_WLD_HARD_TYP 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70860
         MaxLength       =   1
         TabIndex        =   133
         Top             =   4830
         Width           =   495
      End
      Begin VB.TextBox txt_BEND_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -72885
         MaxLength       =   9
         TabIndex        =   132
         Top             =   2730
         Width           =   1980
      End
      Begin VB.TextBox txt_RPT_BEND_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72885
         MaxLength       =   9
         TabIndex        =   131
         Top             =   4410
         Width           =   1980
      End
      Begin VB.TextBox txt_WLD_HARD_UNIT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -61680
         MaxLength       =   4
         TabIndex        =   130
         Top             =   4830
         Width           =   900
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_DIR 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   129
         Top             =   7560
         Width           =   1455
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_DIR 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   128
         Top             =   7560
         Width           =   495
      End
      Begin VB.TextBox txt_A_TIM_IMPACT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   127
         Top             =   7245
         Width           =   900
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_KND 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   126
         Top             =   7245
         Width           =   1455
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_TMP 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68865
         MaxLength       =   4
         TabIndex        =   125
         Top             =   7245
         Width           =   1485
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_TMP_UNIT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -67365
         MaxLength       =   1
         TabIndex        =   124
         Top             =   7245
         Width           =   510
      End
      Begin VB.TextBox txt_A_TIM_IMPACT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -72915
         MaxLength       =   9
         TabIndex        =   123
         Top             =   7245
         Width           =   1980
      End
      Begin VB.TextBox txt_A_TIM_IMPACT_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   122
         Top             =   7245
         Width           =   495
      End
      Begin VB.TextBox txt_A_IMPACT_DIR 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70395
         MaxLength       =   80
         TabIndex        =   121
         Top             =   3855
         Width           =   1470
      End
      Begin VB.TextBox txt_A_IMPACT_DIR 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70905
         MaxLength       =   1
         TabIndex        =   120
         Top             =   3855
         Width           =   495
      End
      Begin VB.TextBox txt_A_IMPACT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -60795
         MaxLength       =   1
         TabIndex        =   119
         Top             =   3405
         Width           =   900
      End
      Begin VB.TextBox txt_A_IMPACT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -72930
         MaxLength       =   9
         TabIndex        =   118
         Top             =   3405
         Width           =   1980
      End
      Begin VB.TextBox txt_A_IMPACT_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70905
         MaxLength       =   1
         TabIndex        =   117
         Top             =   3405
         Width           =   495
      End
      Begin VB.TextBox txt_A_IMPACT_KND 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70395
         MaxLength       =   80
         TabIndex        =   116
         Top             =   3405
         Width           =   1470
      End
      Begin VB.TextBox txt_A_IMPACT_TMP 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68880
         MaxLength       =   4
         TabIndex        =   115
         Top             =   3405
         Width           =   1485
      End
      Begin VB.TextBox txt_A_IMPACT_TMP_UNIT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -67380
         MaxLength       =   1
         TabIndex        =   114
         Top             =   3405
         Width           =   510
      End
      Begin VB.TextBox txt_IMPACT_TMP_UNIT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -67365
         MaxLength       =   1
         TabIndex        =   113
         Top             =   1065
         Width           =   510
      End
      Begin VB.TextBox txt_IMPACT_TMP 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68865
         MaxLength       =   4
         TabIndex        =   112
         Top             =   1065
         Width           =   1485
      End
      Begin VB.TextBox txt_TIM_IMPACT_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   111
         Top             =   5760
         Width           =   495
      End
      Begin VB.TextBox txt_TIM_IMPACT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -72915
         MaxLength       =   9
         TabIndex        =   110
         Top             =   5760
         Width           =   1980
      End
      Begin VB.TextBox txt_TIM_IMPACT_TMP_UNIT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -67365
         MaxLength       =   1
         TabIndex        =   109
         Top             =   5760
         Width           =   510
      End
      Begin VB.TextBox txt_TIM_IMPACT_TMP 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68865
         MaxLength       =   4
         TabIndex        =   108
         Top             =   5760
         Width           =   1485
      End
      Begin VB.TextBox txt_TIM_IMPACT_KND 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   107
         Top             =   5760
         Width           =   1455
      End
      Begin VB.TextBox txt_IMPACT_KND 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   106
         Top             =   1065
         Width           =   1470
      End
      Begin VB.TextBox txt_IMPACT_KND 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   105
         Top             =   1065
         Width           =   495
      End
      Begin VB.TextBox txt_IMPACT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -72915
         MaxLength       =   9
         TabIndex        =   104
         Top             =   1065
         Width           =   1980
      End
      Begin VB.TextBox txt_IMPACT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   103
         Top             =   1065
         Width           =   900
      End
      Begin VB.TextBox txt_TIM_IMPACT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   102
         Top             =   5760
         Width           =   900
      End
      Begin VB.TextBox txt_IMPACT_DIR 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   101
         Top             =   1395
         Width           =   495
      End
      Begin VB.TextBox txt_IMPACT_DIR 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   100
         Top             =   1395
         Width           =   1470
      End
      Begin VB.TextBox txt_TIM_IMPACT_DIR 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   99
         Top             =   6195
         Width           =   495
      End
      Begin VB.TextBox txt_TIM_IMPACT_DIR 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   98
         Top             =   6195
         Width           =   1455
      End
      Begin VB.TextBox txt_HGT_TENCIL_TMP_UNIT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -67395
         MaxLength       =   1
         TabIndex        =   97
         Top             =   1065
         Width           =   510
      End
      Begin VB.TextBox txt_HGT_TENCIL_TMP 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -68895
         MaxLength       =   4
         TabIndex        =   96
         Top             =   1065
         Width           =   1485
      End
      Begin VB.TextBox txt_HGT_TENCIL_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -72945
         MaxLength       =   9
         TabIndex        =   95
         Top             =   1065
         Width           =   1980
      End
      Begin VB.TextBox txt_HGT_SP_EL_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -72945
         MaxLength       =   9
         TabIndex        =   94
         Top             =   6510
         Width           =   1980
      End
      Begin VB.TextBox txt_HGT_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70410
         MaxLength       =   80
         TabIndex        =   93
         Top             =   4335
         Width           =   1470
      End
      Begin VB.TextBox txt_HGT_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70920
         MaxLength       =   1
         TabIndex        =   92
         Top             =   4335
         Width           =   495
      End
      Begin VB.TextBox txt_HGT_SNPP_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70410
         MaxLength       =   80
         TabIndex        =   91
         Top             =   5430
         Width           =   1470
      End
      Begin VB.TextBox txt_HGT_SNPP_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70920
         MaxLength       =   1
         TabIndex        =   90
         Top             =   5430
         Width           =   495
      End
      Begin VB.TextBox txt_HGT_SP_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70410
         MaxLength       =   80
         TabIndex        =   89
         Top             =   6510
         Width           =   1470
      End
      Begin VB.TextBox txt_HGT_SP_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70920
         MaxLength       =   1
         TabIndex        =   88
         Top             =   6510
         Width           =   495
      End
      Begin VB.TextBox txt_HGT_YP_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   87
         Tag             =   "Q0002"
         Top             =   1065
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_SP_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   86
         Tag             =   "Q0002"
         Top             =   6510
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_RA_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   85
         Tag             =   "Q0002"
         Top             =   3000
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_TS_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   84
         Tag             =   "Q0002"
         Top             =   2340
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_SNPP_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   83
         Tag             =   "Q0002"
         Top             =   5430
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   82
         Tag             =   "Q0002"
         Top             =   4335
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_TENCIL_TMP_UNIT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -67395
         MaxLength       =   1
         TabIndex        =   81
         Top             =   1065
         Width           =   510
      End
      Begin VB.TextBox txt_HGT_TENCIL_TMP 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -68895
         MaxLength       =   4
         TabIndex        =   80
         Top             =   1065
         Width           =   1485
      End
      Begin VB.TextBox txt_HGT_TENCIL_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -72945
         MaxLength       =   9
         TabIndex        =   79
         Top             =   1065
         Width           =   1980
      End
      Begin VB.TextBox txt_HGT_SP_EL_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -72945
         MaxLength       =   9
         TabIndex        =   78
         Top             =   6510
         Width           =   1980
      End
      Begin VB.TextBox txt_HGT_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   -70410
         MaxLength       =   80
         TabIndex        =   77
         Top             =   4815
         Width           =   1470
      End
      Begin VB.TextBox txt_HGT_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -70920
         MaxLength       =   1
         TabIndex        =   76
         Top             =   4815
         Width           =   495
      End
      Begin VB.TextBox txt_HGT_SNPP_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   -70410
         MaxLength       =   80
         TabIndex        =   75
         Top             =   5430
         Width           =   1470
      End
      Begin VB.TextBox txt_HGT_SNPP_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -70920
         MaxLength       =   1
         TabIndex        =   74
         Top             =   5430
         Width           =   495
      End
      Begin VB.TextBox txt_HGT_SP_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   -70410
         MaxLength       =   80
         TabIndex        =   73
         Top             =   6510
         Width           =   1470
      End
      Begin VB.TextBox txt_HGT_SP_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -70920
         MaxLength       =   1
         TabIndex        =   72
         Top             =   6510
         Width           =   495
      End
      Begin VB.TextBox txt_HGT_YP_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   71
         Tag             =   "Q0002"
         Top             =   1065
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_SP_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   70
         Tag             =   "Q0002"
         Top             =   6510
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_RA_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   69
         Tag             =   "Q0002"
         Top             =   2880
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_TS_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   68
         Tag             =   "Q0002"
         Top             =   2340
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_SNPP_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   67
         Tag             =   "Q0002"
         Top             =   5430
         Width           =   900
      End
      Begin VB.TextBox txt_HGT_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   66
         Tag             =   "Q0002"
         Top             =   4815
         Width           =   900
      End
      Begin VB.TextBox txt_RA_DIR_NAME 
         Height          =   405
         Index           =   1
         Left            =   -70410
         TabIndex        =   65
         Top             =   2820
         Width           =   1470
      End
      Begin VB.TextBox txt_RA_DIR_CD 
         Height          =   405
         Index           =   1
         Left            =   -70920
         TabIndex        =   64
         Top             =   2820
         Width           =   495
      End
      Begin VB.TextBox txt_YR_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   63
         Top             =   4800
         Width           =   900
      End
      Begin VB.TextBox txt_TS_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   62
         Top             =   2355
         Width           =   900
      End
      Begin VB.TextBox txt_RA_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   61
         Top             =   2820
         Width           =   900
      End
      Begin VB.TextBox txt_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   60
         Top             =   4230
         Width           =   900
      End
      Begin VB.TextBox txt_SNPP_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   59
         Top             =   5730
         Width           =   900
      End
      Begin VB.TextBox txt_SG_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   58
         Top             =   6675
         Width           =   900
      End
      Begin VB.TextBox txt_SP_EL_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   57
         Top             =   7605
         Width           =   900
      End
      Begin VB.TextBox txt_YP_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60810
         MaxLength       =   1
         TabIndex        =   56
         Top             =   1065
         Width           =   900
      End
      Begin VB.TextBox txt_TENCIL_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -72945
         MaxLength       =   9
         TabIndex        =   55
         Top             =   1065
         Width           =   1980
      End
      Begin VB.TextBox txt_SP_EL_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -72945
         MaxLength       =   9
         TabIndex        =   54
         Top             =   7605
         Width           =   1980
      End
      Begin VB.TextBox txt_SNPP_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   -70410
         MaxLength       =   80
         TabIndex        =   53
         Tag             =   "N"
         Top             =   5730
         Width           =   1470
      End
      Begin VB.TextBox txt_SNPP_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -70920
         MaxLength       =   1
         TabIndex        =   52
         Top             =   5730
         Width           =   495
      End
      Begin VB.TextBox txt_SG_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   -70410
         MaxLength       =   80
         TabIndex        =   51
         Tag             =   "N"
         Top             =   6675
         Width           =   1470
      End
      Begin VB.TextBox txt_SG_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -70920
         MaxLength       =   1
         TabIndex        =   50
         Top             =   6675
         Width           =   495
      End
      Begin VB.TextBox txt_SP_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   -70410
         MaxLength       =   80
         TabIndex        =   49
         Tag             =   "N"
         Top             =   7605
         Width           =   1470
      End
      Begin VB.TextBox txt_EL_CD 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   -70410
         MaxLength       =   80
         TabIndex        =   48
         Tag             =   "N"
         Top             =   4230
         Width           =   1470
      End
      Begin VB.TextBox txt_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -70920
         MaxLength       =   1
         TabIndex        =   47
         Top             =   4230
         Width           =   495
      End
      Begin VB.TextBox txt_SP_EL_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -70920
         MaxLength       =   1
         TabIndex        =   46
         Top             =   7605
         Width           =   495
      End
      Begin VB.TextBox txt_HARD_TYP 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   -70350
         MaxLength       =   80
         TabIndex        =   45
         Top             =   1830
         Width           =   1470
      End
      Begin VB.TextBox txt_HARD_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60750
         MaxLength       =   1
         TabIndex        =   44
         Top             =   1830
         Width           =   900
      End
      Begin VB.TextBox txt_HARD_TYP 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -70860
         MaxLength       =   1
         TabIndex        =   43
         Top             =   1830
         Width           =   495
      End
      Begin VB.TextBox txt_BEND_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -60750
         MaxLength       =   1
         TabIndex        =   42
         Top             =   3495
         Width           =   900
      End
      Begin VB.TextBox txt_BEND_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -72885
         MaxLength       =   9
         TabIndex        =   41
         Top             =   3495
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -72915
         TabIndex        =   40
         Top             =   1470
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -72945
         TabIndex        =   39
         Top             =   1425
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -72945
         TabIndex        =   38
         Top             =   1425
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   -72945
         TabIndex        =   37
         Top             =   1425
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   -72945
         TabIndex        =   36
         Top             =   1395
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   -72945
         TabIndex        =   35
         Top             =   3870
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   -72945
         TabIndex        =   34
         Top             =   6180
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   -72960
         TabIndex        =   33
         Top             =   7575
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   -72885
         TabIndex        =   32
         Top             =   1425
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   -72870
         TabIndex        =   31
         Top             =   2235
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   -72885
         TabIndex        =   30
         Top             =   6105
         Width           =   1980
      End
      Begin VB.TextBox txt_HTM_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   -72885
         TabIndex        =   29
         Top             =   5250
         Width           =   1980
      End
      Begin VB.TextBox txt_GRAIN_SIZE_MTH 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   28
         Top             =   1860
         Width           =   1470
      End
      Begin VB.TextBox txt_RMV_CAR_TYP 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   27
         Top             =   1455
         Width           =   495
      End
      Begin VB.TextBox txt_RMV_CAR_TYP 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   26
         Top             =   1455
         Width           =   1470
      End
      Begin VB.TextBox txt_RMV_CAR_SMP_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72915
         MaxLength       =   9
         TabIndex        =   25
         Top             =   1455
         Width           =   1980
      End
      Begin VB.TextBox txt_GRAIN_SIZE_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   24
         Top             =   1860
         Width           =   900
      End
      Begin VB.TextBox txt_RMV_CAR_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   23
         Top             =   1455
         Width           =   900
      End
      Begin VB.TextBox txt_GRAIN_SIZE_TMP_UNIT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -67365
         MaxLength       =   1
         TabIndex        =   22
         Top             =   1860
         Width           =   510
      End
      Begin VB.TextBox txt_BELT_STR_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   21
         Top             =   2745
         Width           =   900
      End
      Begin VB.TextBox txt_OST_GRAIN_CHA 
         Height          =   300
         Left            =   -72300
         TabIndex        =   20
         Top             =   2235
         Width           =   1365
      End
      Begin VB.TextBox txt_OST_GRAIN_SIZE_TMP_UNIT 
         Height          =   300
         Left            =   -67365
         TabIndex        =   19
         Top             =   2235
         Width           =   510
      End
      Begin VB.TextBox txt_OST_GRAIN_SIZE_DSC_CD 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -60780
         MaxLength       =   1
         TabIndex        =   18
         Top             =   2235
         Width           =   900
      End
      Begin VB.TextBox txt_GRAIN_SIZE_MTH 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   17
         Top             =   1860
         Width           =   495
      End
      Begin VB.TextBox txt_OST_GRAIN_SIZE_MTH 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   -70890
         MaxLength       =   1
         TabIndex        =   16
         Top             =   2235
         Width           =   495
      End
      Begin VB.TextBox txt_OST_GRAIN_SIZE_MTH 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -70380
         MaxLength       =   80
         TabIndex        =   15
         Top             =   2235
         Width           =   1470
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   1
         Left            =   -72915
         Top             =   750
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "ȡ������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   0
         Left            =   -74970
         Top             =   750
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   21
         Left            =   -74970
         Top             =   2265
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����ǿ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   20
         Left            =   -74970
         Top             =   2745
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   19
         Left            =   -74970
         Top             =   4275
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�Ϻ��쳤��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   18
         Left            =   -74970
         Top             =   5780
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�涨�Ǳ����쳤Ӧ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   17
         Left            =   -74970
         Top             =   6714
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�涨���쳤Ӧ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   16
         Left            =   -74970
         Top             =   7650
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�涨�����쳤Ӧ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   2
         Left            =   -70890
         Top             =   750
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   3
         Left            =   -68865
         Top             =   750
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�����¶�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   4
         Left            =   -66810
         Top             =   750
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   5
         Left            =   -63570
         Top             =   750
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   7
         Left            =   -62640
         Top             =   750
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   6
         Left            =   -61710
         Top             =   750
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��λ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   8
         Left            =   -60780
         Top             =   750
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "�ж�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   22
         Left            =   -74970
         Top             =   1110
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����ǿ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_DRAW_MIN 
         Height          =   300
         Index           =   1
         Left            =   -63570
         TabIndex        =   162
         Top             =   1110
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_YP_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62640
         TabIndex        =   163
         Top             =   1110
         Width           =   915
         _Version        =   262145
         _ExtentX        =   1614
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   9
         Left            =   -61710
         Top             =   1110
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_DRAW_MIN 
         Height          =   300
         Index           =   2
         Left            =   -63570
         TabIndex        =   164
         Top             =   2265
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TS_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62640
         TabIndex        =   165
         Top             =   2265
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   10
         Left            =   -61710
         Top             =   2265
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_DRAW_MIN 
         Height          =   300
         Index           =   3
         Left            =   -63570
         TabIndex        =   166
         Top             =   2745
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   11
         Left            =   -61710
         Top             =   2745
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_DRAW_MIN 
         Height          =   300
         Index           =   6
         Left            =   -63570
         TabIndex        =   167
         Top             =   4275
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_EL_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62640
         TabIndex        =   168
         Top             =   4275
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   12
         Left            =   -61710
         Top             =   4275
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_DRAW_MIN 
         Height          =   300
         Index           =   8
         Left            =   -63570
         TabIndex        =   169
         Top             =   5780
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_SNPP_EL_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62640
         TabIndex        =   170
         Top             =   5780
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   13
         Left            =   -61710
         Top             =   5780
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_DRAW_MIN 
         Height          =   300
         Index           =   9
         Left            =   -63570
         TabIndex        =   171
         Top             =   6714
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_SG_EL_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62640
         TabIndex        =   172
         Top             =   6714
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   14
         Left            =   -61710
         Top             =   6714
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_DRAW_MIN 
         Height          =   300
         Index           =   10
         Left            =   -63600
         TabIndex        =   173
         Top             =   7650
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_SP_EL_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62640
         TabIndex        =   174
         Top             =   7650
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   15
         Left            =   -61710
         Top             =   7650
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   59
         Left            =   -74970
         Top             =   4846
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��ǿ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_DRAW_MIN 
         Height          =   300
         Index           =   7
         Left            =   -63570
         TabIndex        =   175
         Top             =   4846
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_YR_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62640
         TabIndex        =   176
         Top             =   4846
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   60
         Left            =   -61710
         Top             =   4846
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HARD_MIN 
         Height          =   300
         Index           =   0
         Left            =   -63540
         TabIndex        =   177
         Top             =   1020
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   9
         Left            =   -74940
         Top             =   2730
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   22
         Left            =   -74940
         Top             =   4410
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   10
         Left            =   -74940
         Top             =   4830
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����Ӳ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   11
         Left            =   -74940
         Top             =   5730
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   8
         Left            =   -74940
         Top             =   1020
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "Ӳ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   41
         Left            =   -66750
         Top             =   2730
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "����ֱ��"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   42
         Left            =   -66750
         Top             =   5730
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "����ֱ��"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   40
         Left            =   -65130
         Top             =   2730
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "�����Ƕ�"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   43
         Left            =   -65130
         Top             =   5730
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "�����Ƕ�"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_HARD_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62610
         TabIndex        =   178
         Top             =   1020
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_RPT_BEND_TMS 
         Height          =   300
         Left            =   -63540
         TabIndex        =   179
         Top             =   4410
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_WLD_HARD_MAX 
         Height          =   300
         Left            =   -62610
         TabIndex        =   180
         Top             =   4830
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   20
         Left            =   -61680
         Top             =   4410
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
         BackgroundStyle =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_BEND_DIA 
         Height          =   300
         Index           =   0
         Left            =   -65880
         TabIndex        =   181
         Top             =   2730
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_WLD_BEND_DIA 
         Height          =   300
         Left            =   -65880
         TabIndex        =   182
         Top             =   5730
         Width           =   645
         _Version        =   262145
         _ExtentX        =   1138
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_BEND_ANGLE 
         Height          =   300
         Index           =   0
         Left            =   -64260
         TabIndex        =   183
         Top             =   2730
         Width           =   645
         _Version        =   262145
         _ExtentX        =   1138
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_WLD_BEND_ANG 
         Height          =   300
         Left            =   -64260
         TabIndex        =   184
         Top             =   5745
         Width           =   645
         _Version        =   262145
         _ExtentX        =   1138
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_WLD_HARD_MIN 
         Height          =   300
         Left            =   -63540
         TabIndex        =   185
         Top             =   4830
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel UL_HARD_UNIT 
         Height          =   300
         Index           =   0
         Left            =   -61680
         Top             =   1020
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   ""
         Alignment       =   1
         BackgroundStyle =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   41
         Left            =   -72885
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "ȡ������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   42
         Left            =   -74940
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   43
         Left            =   -70860
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   44
         Left            =   -68835
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�����¶�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   45
         Left            =   -66780
         Top             =   705
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   46
         Left            =   -63540
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   47
         Left            =   -62610
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   48
         Left            =   -61680
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��λ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   49
         Left            =   -60750
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "�ж�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_TIM 
         Height          =   300
         Left            =   -65520
         TabIndex        =   186
         Top             =   5760
         Width           =   1875
         _Version        =   262145
         _ExtentX        =   3307
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   5
         Left            =   -74970
         Top             =   1710
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "ƽ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   4
         Left            =   -74970
         Top             =   2160
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������ά��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   3
         Left            =   -74970
         Top             =   5760
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "ʱЧ�������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   2
         Left            =   -74970
         Top             =   6510
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "ƽ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   1
         Left            =   -74970
         Top             =   6825
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������ά��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   6
         Left            =   -74970
         Top             =   1065
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   11
         Left            =   -66810
         Top             =   5760
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "ʱЧʱ��"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   187
         Top             =   1065
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_AVE_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   188
         Top             =   1710
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   189
         Top             =   2160
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_RATE_MAX 
         Height          =   300
         Left            =   -62640
         TabIndex        =   190
         Top             =   2160
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   191
         Top             =   5760
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_AVE_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   192
         Top             =   6510
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RATE_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   193
         Top             =   6825
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_RATE_MAX 
         Height          =   300
         Left            =   -62640
         TabIndex        =   194
         Top             =   6825
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   16
         Left            =   -61710
         Top             =   1065
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "J"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   17
         Left            =   -61710
         Top             =   2160
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   18
         Left            =   -61710
         Top             =   5760
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "J/cm2"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   19
         Left            =   -61710
         Top             =   6825
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   32
         Left            =   -72915
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "ȡ������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   33
         Left            =   -74970
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   34
         Left            =   -70890
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   35
         Left            =   -68865
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�����¶�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   36
         Left            =   -66810
         Top             =   705
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   37
         Left            =   -63570
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   38
         Left            =   -62640
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   39
         Left            =   -61710
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��λ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   40
         Left            =   -60780
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "�ж�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   8
         Left            =   -66810
         Top             =   1395
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "��С����"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   9
         Left            =   -66810
         Top             =   6195
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "��С����"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_MIN_MIN 
         Height          =   300
         Left            =   -65520
         TabIndex        =   195
         Top             =   1395
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TIM_IMPACT_MIN_MIN 
         Height          =   300
         Left            =   -65520
         TabIndex        =   196
         Top             =   6195
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_MIN 
         Height          =   300
         Left            =   -63585
         TabIndex        =   197
         Top             =   3405
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_AVE_MIN 
         Height          =   300
         Left            =   -63600
         TabIndex        =   198
         Top             =   4170
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_MIN 
         Height          =   300
         Left            =   -63585
         TabIndex        =   199
         Top             =   4620
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_RATE_MAX 
         Height          =   300
         Left            =   -62655
         TabIndex        =   200
         Top             =   4620
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   0
         Left            =   -61725
         Top             =   3405
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "J"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   7
         Left            =   -61725
         Top             =   4620
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   10
         Left            =   -66810
         Top             =   3855
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "��С����"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_A_IMPACT_MIN_MIN 
         Height          =   300
         Left            =   -65520
         TabIndex        =   201
         Top             =   3855
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   12
         Left            =   -74970
         Top             =   4290
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "ƽ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   13
         Left            =   -74970
         Top             =   4620
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������ά��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   14
         Left            =   -74970
         Top             =   3405
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "׷�ӳ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_TIM 
         Height          =   300
         Left            =   -65520
         TabIndex        =   202
         Top             =   7245
         Width           =   1875
         _Version        =   262145
         _ExtentX        =   3307
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   15
         Left            =   -74970
         Top             =   7245
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "׷��ʱЧ�������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   20
         Left            =   -75000
         Top             =   7875
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "ƽ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   21
         Left            =   -74970
         Top             =   8190
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������ά��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   22
         Left            =   -66810
         Top             =   7245
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "ʱЧʱ��"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   203
         Top             =   7245
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_AVE_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   204
         Top             =   7890
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RATE_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   205
         Top             =   8190
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_RATE_MAX 
         Height          =   300
         Left            =   -62640
         TabIndex        =   206
         Top             =   8190
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   23
         Left            =   -61710
         Top             =   7245
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "J/cm2"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   24
         Left            =   -61710
         Top             =   8190
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   25
         Left            =   -66810
         Top             =   7560
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "��С����"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_A_TIM_IMPACT_MIN_MIN 
         Height          =   300
         Left            =   -65520
         TabIndex        =   207
         Top             =   7560
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   3
         Left            =   -75000
         Top             =   2340
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����ǿ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   4
         Left            =   -75000
         Top             =   3000
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   5
         Left            =   -75000
         Top             =   4335
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�Ϻ��쳤��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   6
         Left            =   -75000
         Top             =   5430
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�涨�Ǳ����쳤Ӧ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   7
         Left            =   -75000
         Top             =   6510
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�涨�����쳤Ӧ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   2
         Left            =   -75000
         Top             =   1065
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����ǿ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_YP_MIN 
         Height          =   300
         Index           =   0
         Left            =   -63600
         TabIndex        =   208
         Top             =   1065
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_YP_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62670
         TabIndex        =   209
         Top             =   1065
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   15
         Left            =   -61740
         Top             =   1065
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_TS_MIN 
         Height          =   300
         Index           =   0
         Left            =   -63600
         TabIndex        =   210
         Top             =   2340
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_TS_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62670
         TabIndex        =   211
         Top             =   2340
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   16
         Left            =   -61740
         Top             =   2340
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_MIN 
         Height          =   300
         Index           =   0
         Left            =   -63600
         TabIndex        =   212
         Top             =   3000
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62670
         TabIndex        =   213
         Top             =   3000
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   17
         Left            =   -61740
         Top             =   3000
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_MIN 
         Height          =   300
         Index           =   0
         Left            =   -63600
         TabIndex        =   214
         Top             =   6510
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62670
         TabIndex        =   215
         Top             =   6510
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   20
         Left            =   -61740
         Top             =   6510
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_SNPP_EL_MIN 
         Height          =   300
         Index           =   0
         Left            =   -63600
         TabIndex        =   216
         Top             =   5430
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_SNPP_EL_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62670
         TabIndex        =   217
         Top             =   5430
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   19
         Left            =   -61740
         Top             =   5430
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_EL_MIN 
         Height          =   300
         Index           =   0
         Left            =   -63600
         TabIndex        =   218
         Top             =   4335
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_EL_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62670
         TabIndex        =   219
         Top             =   4335
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   18
         Left            =   -61740
         Top             =   4335
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   24
         Left            =   -75000
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   25
         Left            =   -70920
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   26
         Left            =   -68895
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�����¶�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   27
         Left            =   -66840
         Top             =   705
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   28
         Left            =   -63600
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   29
         Left            =   -62670
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   30
         Left            =   -61740
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��λ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   31
         Left            =   -60810
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "�ж�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   1
         Left            =   -75000
         Top             =   2340
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����ǿ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   8
         Left            =   -75000
         Top             =   2880
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   9
         Left            =   -75000
         Top             =   4815
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�Ϻ��쳤��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   10
         Left            =   -75000
         Top             =   5430
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�涨�Ǳ����쳤Ӧ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   11
         Left            =   -75000
         Top             =   6510
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�涨�����쳤Ӧ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   12
         Left            =   -75000
         Top             =   1065
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����ǿ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_YP_MIN 
         Height          =   300
         Index           =   1
         Left            =   -63600
         TabIndex        =   220
         Top             =   1065
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_YP_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   221
         Top             =   1065
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   13
         Left            =   -61740
         Top             =   1065
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_TS_MIN 
         Height          =   300
         Index           =   1
         Left            =   -63600
         TabIndex        =   222
         Top             =   2340
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_TS_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   223
         Top             =   2340
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   14
         Left            =   -61740
         Top             =   2340
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_MIN 
         Height          =   300
         Index           =   1
         Left            =   -63600
         TabIndex        =   224
         Top             =   2880
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   225
         Top             =   2880
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   27
         Left            =   -61740
         Top             =   2880
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_MIN 
         Height          =   300
         Index           =   1
         Left            =   -63600
         TabIndex        =   226
         Top             =   6510
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_SP_EL_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   227
         Top             =   6510
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   28
         Left            =   -61740
         Top             =   6510
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_SNPP_EL_MIN 
         Height          =   300
         Index           =   1
         Left            =   -63600
         TabIndex        =   228
         Top             =   5430
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_SNPP_EL_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   229
         Top             =   5430
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   29
         Left            =   -61740
         Top             =   5430
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_EL_MIN 
         Height          =   300
         Index           =   1
         Left            =   -63600
         TabIndex        =   230
         Top             =   4815
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HGT_EL_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   231
         Top             =   4815
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   30
         Left            =   -61740
         Top             =   4815
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   61
         Left            =   -75000
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   62
         Left            =   -70920
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   63
         Left            =   -68895
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�����¶�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   64
         Left            =   -66840
         Top             =   705
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   65
         Left            =   -63600
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   66
         Left            =   -62670
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   67
         Left            =   -61740
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��λ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   68
         Left            =   -60810
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "�ж�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   23
         Left            =   -72930
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "ȡ������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   69
         Left            =   -72930
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "ȡ������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   70
         Left            =   -72945
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "ȡ������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   71
         Left            =   -75000
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   72
         Left            =   -75000
         Top             =   2355
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����ǿ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   73
         Left            =   -75000
         Top             =   2820
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   74
         Left            =   -75000
         Top             =   4230
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�Ϻ��쳤��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   75
         Left            =   -75000
         Top             =   5730
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�涨�Ǳ����쳤Ӧ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   76
         Left            =   -75000
         Top             =   6675
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�涨���쳤Ӧ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   77
         Left            =   -75000
         Top             =   7605
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�涨�����쳤Ӧ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   78
         Left            =   -70920
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   79
         Left            =   -68895
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�����¶�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   80
         Left            =   -66840
         Top             =   705
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   81
         Left            =   -63600
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   82
         Left            =   -62670
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   83
         Left            =   -61740
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��λ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   84
         Left            =   -60810
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "�ж�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   85
         Left            =   -75000
         Top             =   1065
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����ǿ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   1
         Left            =   -63600
         TabIndex        =   232
         Top             =   1065
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_YP_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   233
         Top             =   1065
         Width           =   915
         _Version        =   262145
         _ExtentX        =   1614
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   86
         Left            =   -61740
         Top             =   1065
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   2
         Left            =   -63600
         TabIndex        =   234
         Top             =   2355
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TS_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   235
         Top             =   2355
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   87
         Left            =   -61740
         Top             =   2355
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   3
         Left            =   -63600
         TabIndex        =   236
         Top             =   2820
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_RA_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   237
         Top             =   2820
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   88
         Left            =   -61740
         Top             =   2820
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   6
         Left            =   -63600
         TabIndex        =   238
         Top             =   4230
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_EL_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   239
         Top             =   4260
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   89
         Left            =   -61740
         Top             =   4230
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   8
         Left            =   -63600
         TabIndex        =   240
         Top             =   5730
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_SNPP_EL_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   241
         Top             =   5730
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   90
         Left            =   -61740
         Top             =   5730
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   9
         Left            =   -63600
         TabIndex        =   242
         Top             =   6675
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_SG_EL_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   243
         Top             =   6675
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   91
         Left            =   -61740
         Top             =   6675
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   10
         Left            =   -63600
         TabIndex        =   244
         Top             =   7605
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_SP_EL_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   245
         Top             =   7620
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   92
         Left            =   -61740
         Top             =   7605
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "MPa"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   93
         Left            =   -75000
         Top             =   4800
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��ǿ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   7
         Left            =   -63600
         TabIndex        =   246
         Top             =   4860
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_YR_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62670
         TabIndex        =   247
         Top             =   4800
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   94
         Left            =   -61740
         Top             =   4800
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HARD_MIN 
         Height          =   300
         Index           =   1
         Left            =   -63540
         TabIndex        =   248
         Top             =   1830
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   1
         Left            =   -74940
         Top             =   1830
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "׷��Ӳ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HARD_MAX 
         Height          =   300
         Index           =   1
         Left            =   -62610
         TabIndex        =   249
         Top             =   1830
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel UL_HARD_UNIT 
         Height          =   300
         Index           =   1
         Left            =   -61680
         Top             =   1830
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   ""
         Alignment       =   1
         BackgroundStyle =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   2
         Left            =   -74940
         Top             =   3495
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "׷����������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   3
         Left            =   -66750
         Top             =   3495
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "����ֱ��"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   4
         Left            =   -65130
         Top             =   3495
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "�����Ƕ�"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_BEND_DIA 
         Height          =   300
         Index           =   1
         Left            =   -65880
         TabIndex        =   250
         Top             =   3495
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_BEND_ANGLE 
         Height          =   300
         Index           =   1
         Left            =   -64260
         TabIndex        =   251
         Top             =   3495
         Width           =   645
         _Version        =   262145
         _ExtentX        =   1138
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   95
         Left            =   -74970
         Top             =   1470
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   96
         Left            =   -75000
         Top             =   1425
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   97
         Left            =   -74970
         Top             =   1425
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   98
         Left            =   -75000
         Top             =   1425
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   99
         Left            =   -74970
         Top             =   1395
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   100
         Left            =   -74970
         Top             =   3870
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   101
         Left            =   -74970
         Top             =   6180
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   102
         Left            =   -74970
         Top             =   7575
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   103
         Left            =   -74940
         Top             =   1425
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   104
         Left            =   -74940
         Top             =   2235
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   109
         Left            =   -74940
         Top             =   6105
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   110
         Left            =   -74940
         Top             =   5250
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel7 
         Height          =   300
         Left            =   -66780
         Top             =   2235
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         Caption         =   "����ʱ��"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel6 
         Height          =   300
         Left            =   -72900
         Top             =   2235
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit txt_BELT_STR_GRD 
         Height          =   300
         Left            =   -62640
         TabIndex        =   252
         Top             =   2745
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
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
         RawData         =   "0.0"
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   21
         Left            =   -74970
         Top             =   1860
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   15
         Left            =   -61710
         Top             =   1860
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   18
         Left            =   -61710
         Top             =   1455
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "mm"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   21
         Left            =   -66780
         Top             =   1860
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         Caption         =   "����ʱ��"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_RMV_CAR_MAX 
         Height          =   300
         Left            =   -62640
         TabIndex        =   253
         Top             =   1455
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_GRAIN_SIZE_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   254
         Top             =   1860
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_GRAIN_SIZE_TIME 
         Height          =   300
         Left            =   -65700
         TabIndex        =   255
         Top             =   1860
         Width           =   1425
         _Version        =   262145
         _ExtentX        =   2514
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   26
         Left            =   -74970
         Top             =   1455
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��̼��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   50
         Left            =   -72915
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "ȡ������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   51
         Left            =   -74970
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   52
         Left            =   -70890
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   53
         Left            =   -68865
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�����¶�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   54
         Left            =   -66810
         Top             =   705
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   55
         Left            =   -63570
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   56
         Left            =   -62640
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   57
         Left            =   -61710
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��λ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   58
         Left            =   -60780
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "�ж�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_GRAIN_SIZE_MAX 
         Height          =   300
         Left            =   -62640
         TabIndex        =   256
         Top             =   1860
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   0
         Left            =   -74970
         Top             =   2745
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��״��֯"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   3
         Left            =   -61710
         Top             =   2745
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_GRAIN_SIZE_TMP 
         Height          =   300
         Left            =   -68865
         TabIndex        =   257
         Top             =   1860
         Width           =   1485
         _Version        =   262145
         _ExtentX        =   2619
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_OST_GRAIN_SIZE_TMP 
         Height          =   300
         Left            =   -68880
         TabIndex        =   258
         Top             =   2235
         Width           =   1485
         _Version        =   262145
         _ExtentX        =   2619
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_OST_GRAIN_SIZE_TIME 
         Height          =   300
         Left            =   -65700
         TabIndex        =   259
         Top             =   2235
         Width           =   1425
         _Version        =   262145
         _ExtentX        =   2514
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   7
         Left            =   -61710
         Top             =   2235
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_OST_GRAIN_SIZE_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   260
         Top             =   2235
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_OST_GRAIN_SIZE_MAX 
         Height          =   300
         Left            =   -62640
         TabIndex        =   261
         Top             =   2235
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_JOMINY_DIST 
         Height          =   300
         Left            =   -65910
         TabIndex        =   289
         Top             =   2385
         Width           =   1155
         _Version        =   262145
         _ExtentX        =   2037
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   17
         Left            =   -74970
         Top             =   4935
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����˺������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   345
         Index           =   12
         Left            =   -74970
         Top             =   1065
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   609
         Caption         =   "������̽�ˣ�UST��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   13
         Left            =   -74970
         Top             =   1515
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��ƽ"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   14
         Left            =   -74970
         Top             =   2385
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��͸��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   15
         Left            =   -74970
         Top             =   3330
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "����������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   16
         Left            =   -74970
         Top             =   4515
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "���︯ʴ����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   44
         Left            =   -66810
         Top             =   2385
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "�����"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   45
         Left            =   -66810
         Top             =   3330
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "CSR"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel99 
         Height          =   300
         Index           =   7
         Left            =   -64680
         Top             =   2385
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   529
         Caption         =   "mm"
         Alignment       =   0
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
      Begin CSTextLibCtl.sidbEdit sdb_DWTT_YP_MIN 
         Height          =   300
         Left            =   -65880
         TabIndex        =   290
         Top             =   4935
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_DWTT_YP_AVE 
         Height          =   300
         Left            =   -63570
         TabIndex        =   291
         Top             =   6060
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_JOMINY_MIN 
         Height          =   300
         Left            =   -63570
         TabIndex        =   292
         Top             =   2385
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_JOMINY_MAX 
         Height          =   300
         Left            =   -62640
         TabIndex        =   293
         Top             =   2385
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_SSCC_YP_MAX 
         Height          =   300
         Left            =   -62640
         TabIndex        =   294
         Top             =   4515
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   18
         Left            =   -61710
         Top             =   1110
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
         BackgroundStyle =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   19
         Left            =   -61710
         Top             =   4935
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   21
         Left            =   -61710
         Top             =   4515
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   23
         Left            =   -66810
         Top             =   3735
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "CLR"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_HIC_CSR_MAX 
         Height          =   300
         Left            =   -62640
         TabIndex        =   295
         Top             =   3330
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HIC_CLR_MAX 
         Height          =   300
         Left            =   -62640
         TabIndex        =   296
         Top             =   3735
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_HIC_CWR_MAX 
         Height          =   300
         Left            =   -62640
         TabIndex        =   297
         Top             =   4155
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   47
         Left            =   -66810
         Top             =   4515
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "ʱ��"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_SSCC_YP_TIM 
         Height          =   300
         Left            =   -65910
         TabIndex        =   298
         Top             =   4515
         Width           =   1095
         _Version        =   262145
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         BorderStyle     =   0
         NumDecDigits    =   0
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   24
         Left            =   -66810
         Top             =   4110
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "CTR"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   0
         Left            =   -74970
         Top             =   6060
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "ƽ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   26
         Left            =   -66810
         Top             =   4935
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "��С����"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   105
         Left            =   -74970
         Top             =   5700
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   106
         Left            =   -74970
         Top             =   3735
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   107
         Left            =   -74970
         Top             =   2880
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   108
         Left            =   -74970
         Top             =   1935
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   112
         Left            =   -74940
         Top             =   3090
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   113
         Left            =   -74940
         Top             =   3900
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   114
         Left            =   -72885
         Top             =   705
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   529
         Caption         =   "ȡ������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   115
         Left            =   -74970
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   116
         Left            =   -70890
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   117
         Left            =   -68865
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�����¶�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   118
         Left            =   -66810
         Top             =   705
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   119
         Left            =   -63570
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   120
         Left            =   -62640
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   121
         Left            =   -61710
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��λ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   122
         Left            =   -60780
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "�ж�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   25
         Left            =   -75000
         Top             =   1110
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ǽ�������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   1
         Left            =   -66840
         Top             =   1110
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         Caption         =   "��ϵ"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   2
         Left            =   -66840
         Top             =   2220
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         Caption         =   "ϸϵ"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   22
         Left            =   -61740
         Top             =   1110
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_AGRD1 
         Height          =   270
         Left            =   -62670
         TabIndex        =   321
         Top             =   1110
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_AGRD2 
         Height          =   270
         Left            =   -62670
         TabIndex        =   322
         Top             =   1395
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_AGRD3 
         Height          =   270
         Left            =   -62670
         TabIndex        =   323
         Top             =   1680
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_AGRD4 
         Height          =   270
         Left            =   -62670
         TabIndex        =   324
         Top             =   1950
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_BGRD1 
         Height          =   270
         Left            =   -62670
         TabIndex        =   325
         Top             =   2220
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_BGRD2 
         Height          =   270
         Left            =   -62670
         TabIndex        =   326
         Top             =   2505
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_BGRD3 
         Height          =   270
         Left            =   -62670
         TabIndex        =   327
         Top             =   2790
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_NON_METAL_BGRD4 
         Height          =   270
         Left            =   -62670
         TabIndex        =   328
         Top             =   3060
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   8
         Left            =   -66840
         Top             =   3315
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         Caption         =   "DS��"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_DS_GRD 
         Height          =   270
         Left            =   -62670
         TabIndex        =   329
         Top             =   3360
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel87 
         Height          =   285
         Index           =   14
         Left            =   -66840
         Top             =   3630
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         Caption         =   "TIN��"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_TIN_GRD 
         Height          =   270
         Left            =   -62670
         TabIndex        =   330
         Top             =   3630
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   123
         Left            =   -72915
         Top             =   705
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   529
         Caption         =   "ȡ������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   124
         Left            =   -75000
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   125
         Left            =   -70920
         Top             =   705
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   126
         Left            =   -68895
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�����¶�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   127
         Left            =   -66840
         Top             =   705
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   529
         Caption         =   "������Ŀ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   128
         Left            =   -63600
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   129
         Left            =   -62670
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "����"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   130
         Left            =   -61740
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��λ"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   131
         Left            =   -60810
         Top             =   705
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "�ж�"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   22
         Left            =   -75000
         Top             =   4170
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��ӡ"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   23
         Left            =   -75000
         Top             =   4560
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   24
         Left            =   -75000
         Top             =   6060
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�Ͽڼ���"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   270
         Index           =   4
         Left            =   -61740
         Top             =   4560
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   476
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   5
         Left            =   -61740
         Top             =   4170
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel87 
         Height          =   270
         Index           =   6
         Left            =   -66840
         Top             =   4560
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   476
         Caption         =   "ȱ������"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   11
         Left            =   -66840
         Top             =   6060
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         Caption         =   "�Ͽ�����"
         Alignment       =   1
         BackColor       =   12632256
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
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   12
         Left            =   -61740
         Top             =   6060
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_ACD_DFT_GRD2 
         Height          =   270
         Left            =   -62670
         TabIndex        =   360
         Top             =   4845
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_ACD_DFT_GRD3 
         Height          =   270
         Left            =   -62670
         TabIndex        =   361
         Top             =   5130
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_S_PRINT_DRG 
         Height          =   300
         Left            =   -62670
         TabIndex        =   362
         Top             =   4170
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_ACD_DFT_GRD1 
         Height          =   270
         Left            =   -62670
         TabIndex        =   363
         Top             =   4560
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel87 
         Height          =   300
         Index           =   0
         Left            =   -70920
         Top             =   6060
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   $"AQA0030C.frx":08F8
         Alignment       =   1
         BackColor       =   12632256
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
      Begin CSTextLibCtl.sidbEdit sdb_ACD_DFT_GRD4 
         Height          =   270
         Left            =   -62670
         TabIndex        =   364
         Top             =   5430
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_ACD_DFT_GRD5 
         Height          =   270
         Left            =   -62670
         TabIndex        =   365
         Top             =   5715
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   476
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   2
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   132
         Left            =   -74970
         Top             =   1065
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   31
         Left            =   -75000
         Top             =   2235
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�����徧����"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   133
         Left            =   -75000
         Top             =   1470
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   134
         Left            =   -75000
         Top             =   3810
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "�ȴ���ȡ������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   111
         Left            =   -74970
         Top             =   1830
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   135
         Left            =   -75000
         Top             =   1785
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   136
         Left            =   -75000
         Top             =   1785
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   137
         Left            =   -75000
         Top             =   1785
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   27
         Left            =   -75000
         Top             =   4980
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��������ֵ"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sidbEdit1 
         Height          =   300
         Left            =   -75000
         TabIndex        =   377
         Top             =   300
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_CP_MIN 
         Height          =   300
         Index           =   0
         Left            =   -63600
         TabIndex        =   378
         Top             =   4980
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   28
         Left            =   -75000
         Top             =   5340
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "ƽ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_CP_MIN 
         Height          =   300
         Index           =   1
         Left            =   -63600
         TabIndex        =   379
         Top             =   5340
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   29
         Left            =   -75000
         Top             =   2940
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "ƽ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel32 
         Height          =   300
         Index           =   30
         Left            =   -75000
         Top             =   2580
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��������ֵ"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_CP_MIN 
         Height          =   300
         Index           =   2
         Left            =   -63600
         TabIndex        =   380
         Top             =   2580
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_IMPACT_CP_MIN 
         Height          =   300
         Index           =   3
         Left            =   -63600
         TabIndex        =   381
         Top             =   2940
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   138
         Left            =   -75000
         Top             =   3300
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��ȷ������������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_DRAW_MIN 
         Height          =   300
         Index           =   4
         Left            =   -63600
         TabIndex        =   382
         Top             =   3300
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   139
         Left            =   -61680
         Top             =   3300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   140
         Left            =   -75000
         Top             =   3300
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��ȷ������������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   4
         Left            =   -63600
         TabIndex        =   384
         Top             =   3300
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   141
         Left            =   -61680
         Top             =   3300
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   32
         Left            =   -75000
         Top             =   3420
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��ȷ������������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_MIN 
         Height          =   300
         Index           =   2
         Left            =   -63600
         TabIndex        =   386
         Top             =   3420
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   33
         Left            =   -61680
         Top             =   3420
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   34
         Left            =   -75000
         Top             =   3540
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��ȷ������������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_MIN 
         Height          =   300
         Index           =   3
         Left            =   -63600
         TabIndex        =   388
         Top             =   3540
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   35
         Left            =   -61680
         Top             =   3540
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   5
         Left            =   -74970
         Top             =   6420
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "ʣ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_sc_max 
         Height          =   300
         Left            =   -62640
         TabIndex        =   390
         Top             =   6420
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   142
         Left            =   -75000
         Top             =   3780
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "ƽ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_RA_MAX 
         Height          =   300
         Index           =   0
         Left            =   -62640
         TabIndex        =   391
         Top             =   2745
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_DRAW_MIN 
         Height          =   300
         Index           =   5
         Left            =   -63600
         TabIndex        =   392
         Top             =   3780
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   143
         Left            =   -61680
         Top             =   3780
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   144
         Left            =   -75000
         Top             =   3780
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "ƽ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   5
         Left            =   -63600
         TabIndex        =   393
         Top             =   3660
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   145
         Left            =   -61680
         Top             =   3660
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   36
         Left            =   -75000
         Top             =   3900
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "ƽ��"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_MIN 
         Height          =   300
         Index           =   4
         Left            =   -63600
         TabIndex        =   394
         Top             =   3900
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   37
         Left            =   -61680
         Top             =   3900
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   38
         Left            =   -75000
         Top             =   4140
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "��ȷ������������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_HGT_RA_MIN 
         Height          =   300
         Index           =   5
         Left            =   -63600
         TabIndex        =   395
         Top             =   4140
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   300
         Index           =   39
         Left            =   -61680
         Top             =   4140
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   6
         Left            =   -74970
         Top             =   5340
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "NDT����˺������"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_DRAW_MIN 
         Height          =   300
         Index           =   11
         Left            =   -63600
         TabIndex        =   398
         Top             =   8340
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_DRAW_MAX 
         Height          =   300
         Index           =   11
         Left            =   -62640
         TabIndex        =   399
         Top             =   8340
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MAX 
         Height          =   300
         Index           =   11
         Left            =   -62640
         TabIndex        =   400
         Top             =   8100
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   11
         Left            =   -63600
         TabIndex        =   401
         Top             =   8100
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   147
         Left            =   -61800
         Top             =   8100
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   146
         Left            =   -75000
         Top             =   8100
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "���ȱ����쳤��UEL"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   149
         Left            =   -61680
         Top             =   8340
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "��"
         Alignment       =   1
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
      Begin InDate.ULabel ULabel4 
         Height          =   300
         Index           =   150
         Left            =   -74970
         Top             =   8340
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "���ȱ����쳤��UEL"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   7
         Left            =   -74970
         Top             =   6780
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "Ӧ����������Ŀ1"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   25
         Left            =   -74970
         Top             =   7500
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "Ӧ����������Ŀ3"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   26
         Left            =   -74970
         Top             =   7140
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "Ӧ����������Ŀ2"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   27
         Left            =   -74970
         Top             =   7860
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "Ӧ����������Ŀ4"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   28
         Left            =   -74970
         Top             =   8220
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "Ӧ����������Ŀ5"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin InDate.ULabel ULabel71 
         Height          =   300
         Index           =   29
         Left            =   -74970
         Top             =   8580
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         Caption         =   "Ӧ����ֵ1-5"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         BorderEffect    =   0
         BorderStyle     =   1
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
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MAX 
         Height          =   300
         Index           =   16
         Left            =   -62640
         TabIndex        =   419
         Top             =   8220
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   16
         Left            =   -63570
         TabIndex        =   420
         Top             =   8220
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MAX 
         Height          =   300
         Index           =   12
         Left            =   -62640
         TabIndex        =   421
         Top             =   6780
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   12
         Left            =   -63570
         TabIndex        =   422
         Top             =   6780
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MAX 
         Height          =   300
         Index           =   13
         Left            =   -62640
         TabIndex        =   423
         Top             =   7140
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   13
         Left            =   -63570
         TabIndex        =   424
         Top             =   7140
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MAX 
         Height          =   300
         Index           =   14
         Left            =   -62640
         TabIndex        =   425
         Top             =   7500
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   14
         Left            =   -63570
         TabIndex        =   426
         Top             =   7500
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MAX 
         Height          =   300
         Index           =   15
         Left            =   -62640
         TabIndex        =   427
         Top             =   7860
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_A_DRAW_MIN 
         Height          =   300
         Index           =   15
         Left            =   -63570
         TabIndex        =   428
         Top             =   7860
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   14
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   8055
         Left            =   45
         TabIndex        =   429
         Top             =   360
         Width           =   15135
         _Version        =   393216
         _ExtentX        =   26696
         _ExtentY        =   14208
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   12
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQA0030C.frx":090D
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   -74880
         X2              =   -59700
         Y1              =   9540
         Y2              =   9540
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   -75000
         X2              =   -59820
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   -75000
         X2              =   -59820
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   -75000
         X2              =   -59820
         Y1              =   9540
         Y2              =   9540
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   -75000
         X2              =   -59820
         Y1              =   8820
         Y2              =   8820
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   41
         X1              =   -63630
         X2              =   -63630
         Y1              =   705
         Y2              =   7590
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   40
         X1              =   -66870
         X2              =   -66870
         Y1              =   705
         Y2              =   7590
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   39
         X1              =   -70950
         X2              =   -70950
         Y1              =   705
         Y2              =   7590
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   -75000
         X2              =   -59850
         Y1              =   7590
         Y2              =   7590
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   21
         X1              =   -72960
         X2              =   -72960
         Y1              =   705
         Y2              =   7590
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   38
         X1              =   -63600
         X2              =   -63600
         Y1              =   1860
         Y2              =   9900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   37
         X1              =   -66840
         X2              =   -66840
         Y1              =   1860
         Y2              =   9900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   36
         X1              =   -70920
         X2              =   -70920
         Y1              =   1860
         Y2              =   9900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   25
         X1              =   -72960
         X2              =   -72960
         Y1              =   1860
         Y2              =   9900
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   -74880
         X2              =   -59730
         Y1              =   9900
         Y2              =   9900
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   -75000
         X2              =   -59850
         Y1              =   6555
         Y2              =   6555
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   -63600
         X2              =   -63600
         Y1              =   750
         Y2              =   8820
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   -66840
         X2              =   -66840
         Y1              =   750
         Y2              =   8820
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   -68895
         X2              =   -68895
         Y1              =   750
         Y2              =   8820
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   -70920
         X2              =   -70920
         Y1              =   750
         Y2              =   8820
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   -72945
         X2              =   -72960
         Y1              =   750
         Y2              =   8820
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   17
         X1              =   -75000
         X2              =   -59850
         Y1              =   10260
         Y2              =   10260
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   19
         X1              =   -72915
         X2              =   -72915
         Y1              =   705
         Y2              =   6555
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   18
         X1              =   -70890
         X2              =   -70875
         Y1              =   705
         Y2              =   6555
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   17
         X1              =   -68865
         X2              =   -68865
         Y1              =   705
         Y2              =   6555
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   16
         X1              =   -66810
         X2              =   -66810
         Y1              =   705
         Y2              =   6555
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   15
         X1              =   -63570
         X2              =   -63570
         Y1              =   705
         Y2              =   6555
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   12
         X1              =   -75030
         X2              =   -59880
         Y1              =   4305
         Y2              =   4305
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   13
         X1              =   -74970
         X2              =   -59820
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   14
         X1              =   -74970
         X2              =   -59820
         Y1              =   2625
         Y2              =   2625
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   16
         X1              =   -74880
         X2              =   -59700
         Y1              =   7140
         Y2              =   7140
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   15
         X1              =   -75000
         X2              =   -59820
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   14
         X1              =   -72960
         X2              =   -72960
         Y1              =   1740
         Y2              =   9540
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   13
         X1              =   -70920
         X2              =   -70920
         Y1              =   1740
         Y2              =   9540
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   12
         X1              =   -68880
         X2              =   -68880
         Y1              =   1740
         Y2              =   9540
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   11
         X1              =   -66840
         X2              =   -66840
         Y1              =   1740
         Y2              =   9540
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   10
         X1              =   -63600
         X2              =   -63600
         Y1              =   1740
         Y2              =   9540
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   21
         X1              =   -75000
         X2              =   -59820
         Y1              =   5700
         Y2              =   5700
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   9
         X1              =   -63630
         X2              =   -63630
         Y1              =   465
         Y2              =   6720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   -66870
         X2              =   -66870
         Y1              =   705
         Y2              =   6960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   -68910
         X2              =   -68910
         Y1              =   705
         Y2              =   6960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   -70950
         X2              =   -70950
         Y1              =   705
         Y2              =   6960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   -72975
         X2              =   -72975
         Y1              =   705
         Y2              =   6960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   26
         X1              =   -63630
         X2              =   -63630
         Y1              =   705
         Y2              =   6960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   27
         X1              =   -66870
         X2              =   -66870
         Y1              =   705
         Y2              =   6960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   28
         X1              =   -68925
         X2              =   -68925
         Y1              =   705
         Y2              =   6960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   29
         X1              =   -70950
         X2              =   -70950
         Y1              =   705
         Y2              =   6960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   30
         X1              =   -72975
         X2              =   -72975
         Y1              =   705
         Y2              =   6960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   31
         X1              =   -72975
         X2              =   -72975
         Y1              =   705
         Y2              =   9540
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   32
         X1              =   -70920
         X2              =   -70920
         Y1              =   2100
         Y2              =   9540
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   33
         X1              =   -68925
         X2              =   -68925
         Y1              =   705
         Y2              =   9540
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   34
         X1              =   -66870
         X2              =   -66870
         Y1              =   705
         Y2              =   9540
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   35
         X1              =   -63630
         X2              =   -63630
         Y1              =   705
         Y2              =   9540
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   18
         X1              =   -70290
         X2              =   -55140
         Y1              =   8805
         Y2              =   8805
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   -75000
         X2              =   -59850
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line49 
         BorderColor     =   &H00808080&
         Index           =   27
         X1              =   -74970
         X2              =   -59820
         Y1              =   3135
         Y2              =   3135
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   20
         X1              =   -63600
         X2              =   -63600
         Y1              =   705
         Y2              =   3135
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   22
         X1              =   -68895
         X2              =   -68895
         Y1              =   705
         Y2              =   3135
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   23
         X1              =   -70920
         X2              =   -70920
         Y1              =   705
         Y2              =   3135
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         Index           =   24
         X1              =   -72945
         X2              =   -72945
         Y1              =   705
         Y2              =   3135
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   90
      X2              =   15210
      Y1              =   405
      Y2              =   405
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   12480
      Y1              =   -1665
      Y2              =   -1665
   End
End
Attribute VB_Name = "AQA0030C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       ��������
'-- Sub_System Name   ������׼����
'-- Program Name      ��׼������Ϣ����
'-- Program ID        AQA0030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       ��׼������Ϣ����
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String            'Form Type
Public Toolbar_St As String          'Active Form ToolBar Setting
Public sAuthority As String          'Active Form Authority Setting

Dim pControl As New Collection       'Master Primary Key Collection
Dim nControl As New Collection       'Master Necessary Collection
Dim mControl As New Collection       'Master Maxlength check Collection
Dim iControl As New Collection       'Master Insert Collection
Dim rControl As New Collection       'Master Refer Collection
Dim cControl As New Collection       'Master Copy Collection
Dim aControl As New Collection       'Master -> Spread Collection
Dim lControl As New Collection       'Master Lock Collection

Dim pControl3 As New Collection       'Master Primary Key Collection
Dim nControl3 As New Collection       'Master Necessary Collection
Dim mControl3 As New Collection       'Master Maxlength check Collection
Dim iControl3 As New Collection       'Master Insert Collection
Dim rControl3 As New Collection       'Master Refer Collection
Dim cControl3 As New Collection       'Master Copy Collection
Dim aControl3 As New Collection       'Master -> Spread Collection
Dim lControl3 As New Collection       'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection


Dim Mc1 As New Collection            'Master Collection
Dim Mc2 As New Collection            'Master Collection
Dim Mc3 As New Collection            'Master Collection

Dim Sc1 As New Collection            'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection


'Dim lBlkcol1 As Long                'To Excel Block Col1
'Dim lBlkcol2 As Long                'To Excel Block Col2
'Dim lBlkrow1 As Long                'To Excel Block Row1
'Dim lBlkrow2 As Long                'To Excel Block Row2

Dim bChk As Boolean
Dim btChk As Boolean

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Master"


    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
                Call Gp_Ms_Collection(txt_STDSPEC, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_STDSPEC_YY, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_THK_MIN, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_THK_MAX, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_SMP_LOC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_SMP_LOC_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_SMP_LEN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_PRE_SMP_QTY, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_SMP_STD_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_TEST_FL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 '����
       Call Gp_Ms_Collection(txt_TENCIL_SMP_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_DRAW_MIN(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_YP_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_YP_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_DRAW_MIN(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_TS_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_TS_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_RA_DIR_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_RA_DIR_NAME(0), " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_DRAW_MIN(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_RA_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_RA_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           'louyannan 20101118 start
           
              Call Gp_Ms_Collection(sdb_DRAW_MIN(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              
           Call Gp_Ms_Collection(txt_RA_DSC_CD(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_DRAW_MIN(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          'louyannan 20101118 end
          
               Call Gp_Ms_Collection(txt_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_DRAW_MIN(6), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_EL_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_EL_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_DRAW_MIN(7), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_YR_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_YR_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_SNPP_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_SNPP_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_DRAW_MIN(8), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_SNPP_EL_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_SNPP_EL_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SG_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SG_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_DRAW_MIN(9), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_SG_EL_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_SG_EL_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_SP_EL_SMP_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SP_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SP_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_DRAW_MIN(10), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_SP_EL_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_SP_EL_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'EDIT BY ��ѧ�� ���Ӿ��ȱ����쳤��uel
        Call Gp_Ms_Collection(sdb_DRAW_MIN(11), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_DRAW_MAX(11), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_DRAW_DSC_CD(11), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  
  
  
  '��������
   Call Gp_Ms_Collection(txt_HGT_TENCIL_SMP_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_HGT_TENCIL_TMP(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_HGT_TENCIL_TMP_UNIT(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_YP_MIN(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_YP_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HGT_YP_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_TS_MIN(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_TS_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HGT_TS_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
          
          Call Gp_Ms_Collection(sdb_HGT_RA_MIN(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_RA_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HGT_RA_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             'louyannan 20101118 start
              Call Gp_Ms_Collection(sdb_HGT_RA_MIN(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
         Call Gp_Ms_Collection(txt_HGT_RA_DSC_CD(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_RA_MIN(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           'louyannan 20101118 end
           
           Call Gp_Ms_Collection(txt_HGT_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_HGT_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_EL_MIN(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_EL_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HGT_EL_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_HGT_SNPP_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_HGT_SNPP_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_HGT_SNPP_EL_MIN(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_HGT_SNPP_EL_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_HGT_SNPP_EL_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_HGT_SP_EL_SMP_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_HGT_SP_EL_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_HGT_SP_EL_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_SP_EL_MIN(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_SP_EL_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_HGT_SP_EL_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 '���
          Call Gp_Ms_Collection(txt_IMPACT(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_IMPACT_KND(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_IMPACT_KND(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_IMPACT_DIR(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_IMPACT_DIR(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_IMPACT_TMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_IMPACT_TMP_UNIT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_IMPACT_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_IMPACT_MIN_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_IMPACT_AVE_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_IMPACT_RATE_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_IMPACT_RATE_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
          'louyannan 20101118 start
          Call Gp_Ms_Collection(sdb_IMPACT_CP_MIN(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_IMPACT_CP_MIN(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           'louyannan 20101118 end
           
           
          Call Gp_Ms_Collection(txt_IMPACT(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_A_IMPACT(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_A_IMPACT_KND(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_A_IMPACT_KND(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_A_IMPACT_DIR(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_A_IMPACT_DIR(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_A_IMPACT_TMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_A_IMPACT_TMP_UNIT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_A_IMPACT_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_A_IMPACT_MIN_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_A_IMPACT_AVE_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_A_IMPACT_RATE_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
        
        'louyannan 20101118 start
        Call Gp_Ms_Collection(sdb_IMPACT_CP_MIN(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_IMPACT_CP_MIN(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        'louyannan 20101118 end
        
        Call Gp_Ms_Collection(txt_A_IMPACT(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_TIM_IMPACT(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_TIM_IMPACT_KND(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_TIM_IMPACT_KND(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_TIM_IMPACT_DIR(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_TIM_IMPACT_DIR(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_TIM_IMPACT_TMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_TIM_IMPACT_TMP_UNIT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_TIM_IMPACT_TIM, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_TIM_IMPACT_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_TIM_IMPACT_MIN_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_TIM_IMPACT_AVE_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_TIM_IMPACT_RATE_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_TIM_IMPACT_RATE_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_TIM_IMPACT(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_A_TIM_IMPACT(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_A_TIM_IMPACT_KND(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_A_TIM_IMPACT_KND(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_A_TIM_IMPACT_DIR(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_A_TIM_IMPACT_DIR(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_A_TIM_IMPACT_TMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_A_TIM_IMPACT_TMP_UNIT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_TIM, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_MIN_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_AVE_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RATE_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(sdb_A_TIM_IMPACT_RATE_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_A_TIM_IMPACT(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'����
            Call Gp_Ms_Collection(txt_HARD_TYP(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_HARD_TYP(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_HARD_MIN(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_HARD_MAX(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_HARD_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_BEND_SMP_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_BEND_DIA(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_BEND_ANGLE(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_BEND_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_RPT_BEND_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_RPT_BEND_TMS, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_RPT_BEND_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_WLD_HARD_TYP(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_WLD_HARD_TYP(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_WLD_HARD_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_WLD_HARD_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_WLD_HARD_UNIT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_WLD_HARD_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_WLD_BEND_DIA, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_WLD_BEND_ANG, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_WLD_BEND_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_UST_STD_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_UST_STD_CD(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_UST_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_UST_GRD_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_UST_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_FOAT_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_FOAT_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_JOMINY_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_JOMINY_TYP(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_JOMINY_TYP(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_JOMINY_DIST, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_JOMINY_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_JOMINY_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_JOMINY_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HIC_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_HIC_SVT_KND(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_HIC_SVT_KND(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_HIC_CSR_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_HIC_CLR_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_HIC_CWR_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HIC_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SSCC_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_SSCC_SVT_KND(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_SSCC_SVT_KND(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_SSCC_YP_TIM, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_SSCC_YP_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SSCC_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_DWTT_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_DWTT_TMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_DWTT_TMP_UNIT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_DWTT_YP_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_DWTT_YP_AVE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_DWTT_DSC_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'louyannan 20101118 start
Call Gp_Ms_Collection(txt_DWTT_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(sdb_sc_max, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_sc_dsc_cd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)


'louyannan 20101118 end


'����
         Call Gp_Ms_Collection(txt_RMV_CAR_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_RMV_CAR_TYP(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_RMV_CAR_TYP(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_RMV_CAR_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_RMV_CAR_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_GRAIN_SIZE_MTH(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_GRAIN_SIZE_MTH(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_GRAIN_SIZE_TMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_GRAIN_SIZE_TMP_UNIT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_GRAIN_SIZE_TIME, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_GRAIN_SIZE_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_GRAIN_SIZE_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_GRAIN_SIZE_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_S_PRINT_DRG, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_S_PRINT_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ACD_DFT_TYP1(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ACD_DFT_TYP1(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_ACD_DFT_GRD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ACD_DFT_TYP2(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ACD_DFT_TYP2(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_ACD_DFT_GRD2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ACD_DFT_TYP3(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ACD_DFT_TYP3(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_ACD_DFT_GRD3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ACD_DFT_TYP4(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ACD_DFT_TYP4(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_ACD_DFT_GRD4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ACD_DFT_TYP5(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ACD_DFT_TYP5(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_ACD_DFT_GRD5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ACD_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_FRACT_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_FRACT_NAME_CD1(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_FRACT_NAME_CD1(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_FRACT_GRD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_FRACT_NAME_CD2(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_FRACT_NAME_CD2(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_FRACT_GRD2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_FRACT_NAME_CD3(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_FRACT_NAME_CD3(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_FRACT_GRD3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_FRACT_NAME_CD4(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_FRACT_NAME_CD4(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_FRACT_GRD4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_FRACT_NAME_CD5(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_FRACT_NAME_CD5(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_FRACT_GRD5, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_FRACT_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_NON_METAL_SMP_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_NON_METAL_TYP(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_NON_METAL_TYP(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_ACD1(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_ACD1(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_AGRD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_ACD2(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_ACD2(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_AGRD2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_ACD3(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_ACD3(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_AGRD3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_ACD4(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_ACD4(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_AGRD4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_BCD1(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_BCD1(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_BGRD1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_BCD2(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_BCD2(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_BGRD2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_BCD3(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_BCD3(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_BGRD3, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_BCD4(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_NON_METAL_BCD4(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(sdb_NON_METAL_BGRD4, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_NON_METAL_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_BELT_STR_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_BELT_STR_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_ins_emp, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_OST_GRAIN_CHA, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_OST_GRAIN_SIZE_MTH(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_OST_GRAIN_SIZE_MTH(1), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_OST_GRAIN_SIZE_TMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_OST_GRAIN_SIZE_TMP_UNIT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_OST_GRAIN_SIZE_TIME, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_OST_GRAIN_SIZE_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_OST_GRAIN_SIZE_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_OST_GRAIN_SIZE_DSC_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(sdb_DS_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(sdb_TIN_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
   
End Sub

Private Sub Form_Define2()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Master"
    
'׷������
'20090812 START
       Call Gp_Ms_Collection(txt_HGT_RA_DIR_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_HGT_RA_DIR_NAME(0), " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'20090812 END


       Call Gp_Ms_Collection(txt_TENCIL_SMP_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_A_DRAW_MIN(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_YP_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_YP_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_A_DRAW_MIN(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_TS_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_TS_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           
           
           Call Gp_Ms_Collection(txt_RA_DIR_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_RA_DIR_NAME(1), " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_A_DRAW_MIN(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_RA_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_RA_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           'louyannan 20101118start
            
            ' Call Gp_Ms_Collection(sdb_A_DRAW_MIN(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              
             ' Call Gp_Ms_Collection(txt_RA_DSC_CD(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             ' Call Gp_Ms_Collection(sdb_A_DRAW_MIN(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            'louyannan 20101118 end
               
               
               Call Gp_Ms_Collection(txt_EL_CD(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_EL_CD(3), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_A_DRAW_MIN(6), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_EL_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_EL_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_A_DRAW_MIN(7), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_YR_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_YR_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_SNPP_EL_CD(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_SNPP_EL_CD(3), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_A_DRAW_MIN(8), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_SNPP_EL_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_SNPP_EL_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SG_EL_CD(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SG_EL_CD(3), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_A_DRAW_MIN(9), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_SG_EL_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_SG_EL_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_SP_EL_SMP_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SP_EL_CD(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SP_EL_CD(3), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_A_DRAW_MIN(10), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_SP_EL_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_SP_EL_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'gengxueyu ׷������UEL ���ȱ����쳤�� 20110210
          Call Gp_Ms_Collection(sdb_A_DRAW_MIN(11), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(sdb_A_DRAW_MAX(11), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_A_DRAW_DSC_CD(11), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         


'׷�����������ʶ
'׷�Ӹ�������
   Call Gp_Ms_Collection(txt_HGT_TENCIL_SMP_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_HGT_TENCIL_TMP(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_HGT_TENCIL_TMP_UNIT(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_YP_MIN(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_YP_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HGT_YP_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_TS_MIN(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_TS_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HGT_TS_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'20090812 START
       Call Gp_Ms_Collection(txt_HGT_RA_DIR_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_HGT_RA_DIR_NAME(1), " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'20090812 END
          Call Gp_Ms_Collection(sdb_HGT_RA_MIN(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_RA_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HGT_RA_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           'louyannan 20101118 start
          
         ' Call Gp_Ms_Collection(sdb_HGT_RA_MIN(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
     '  Call Gp_Ms_Collection(txt_HGT_RA_DSC_CD(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        '    Call Gp_Ms_Collection(sdb_HGT_RA_MIN(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           'louyannan 20101118 end
           
           Call Gp_Ms_Collection(txt_HGT_EL_CD(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_HGT_EL_CD(3), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_EL_MIN(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_HGT_EL_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_HGT_EL_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_HGT_SNPP_EL_CD(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_HGT_SNPP_EL_CD(3), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_HGT_SNPP_EL_MIN(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_HGT_SNPP_EL_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_HGT_SNPP_EL_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_HGT_SP_EL_SMP_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_HGT_SP_EL_CD(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_HGT_SP_EL_CD(3), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_SP_EL_MIN(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_HGT_SP_EL_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_HGT_SP_EL_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                                                                                                                                                          
'׷��Ӳ��
            Call Gp_Ms_Collection(txt_HARD_TYP(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_HARD_TYP(3), " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_HARD_MIN(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_HARD_MAX(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_HARD_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                                                                                                                                                          
'׷������
         Call Gp_Ms_Collection(txt_BEND_SMP_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_BEND_DIA(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_BEND_ANGLE(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_BEND_DSC_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                                                                                                                                                          
'ȡ������
              Call Gp_Ms_Collection(txt_HTM_CD(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_HTM_CD(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_HTM_CD(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_HTM_CD(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_HTM_CD(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_HTM_CD(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_HTM_CD(6), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_HTM_CD(7), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_HTM_CD(8), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_HTM_CD(9), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HTM_CD(10), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HTM_CD(11), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HTM_CD(12), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HTM_CD(13), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HTM_CD(14), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HTM_CD(15), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HTM_CD(16), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HTM_CD(17), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HTM_CD(18), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HTM_CD(19), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HTM_CD(20), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_SMP_WID(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_SMP_WID(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_SMP_WID(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_SMP_WID(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       '20110211 edit by ��ѧ�� ����Ӧ������Ŀ1-5 ��������С�����ж�������Ӧ��ֵ1-5 ҳ�����������
              Call Gp_Ms_Collection(txt_STRESS_KND(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_A_DRAW_MIN(12), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_A_DRAW_MAX(12), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_A_DRAW_DSC_CD(12), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_STRESS_KND(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_A_DRAW_MIN(13), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_A_DRAW_MAX(13), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_A_DRAW_DSC_CD(13), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_STRESS_KND(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_A_DRAW_MIN(14), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_A_DRAW_MAX(14), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_A_DRAW_DSC_CD(14), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_STRESS_KND(5), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_A_DRAW_MIN(15), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_A_DRAW_MAX(15), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_A_DRAW_DSC_CD(15), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_STRESS_KND(6), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_A_DRAW_MIN(16), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_A_DRAW_MAX(16), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_A_DRAW_DSC_CD(16), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

             Call Gp_Ms_Collection(txt_STRESS_KND(7), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_STRESS_KND(8), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_STRESS_KND(9), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_STRESS_KND(10), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_STRESS_KND(11), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    'MASTER Collection
    Mc1.Add Item:="AQA0030C.P_REFER", Key:="P-R"
    Mc1.Add Item:="AQA0030C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="AQA0030C.P_ONEROW", Key:="P-O"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
        
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    SS2.RowHeight(-1) = 13
         
End Sub

'���û�������Ŀ  ���� 2012.11.13
Private Sub Form_Define3()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_STDSPEC, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
 Call Gp_Ms_Collection(txt_STDSPEC_YY, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
    Call Gp_Ms_Collection(txt_THK_MIN, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
    Call Gp_Ms_Collection(txt_THK_MAX, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)

   
    Mc3.Add Item:=pControl3, Key:="pControl"
    Mc3.Add Item:=nControl3, Key:="nControl"
    Mc3.Add Item:=mControl3, Key:="mControl"
    Mc3.Add Item:=iControl3, Key:="iControl"
    Mc3.Add Item:=rControl3, Key:="rControl"
    Mc3.Add Item:=cControl3, Key:="cControl"
    Mc3.Add Item:=aControl3, Key:="aControl"
    Mc3.Add Item:=lControl3, Key:="lControl"
    
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
    
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQA0030C.P_SREFER", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
         
End Sub
'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
            
        Case "txt_STDSPEC"                  '��׼���
            sCode = "STDSPEC"
            Set oCodeName = txt_STDSPEC_YY
        
        Case "txt_SMP_LOC"                  'ȡ����λ��
            sCode = "Q0042"
            Set oCodeName = txt_SMP_LOC_NAME
            
        'ȡ������
        Case "txt_TENCIL_SMP_CD", "txt_SP_EL_SMP_CD", "txt_HGT_TENCIL_SMP_CD", "txt_HGT_SP_EL_SMP_CD", "txt_IMPACT(1)", "txt_TIM_IMPACT(1)", "txt_BEND_SMP_CD", "txt_RPT_BEND_SMP_CD", "txt_FOAT_SMP_CD", "txt_JOMINY_SMP_CD", "txt_HIC_SMP_CD", "txt_SSCC_SMP_CD", "txt_DWTT_SMP_CD", "txt_RMV_CAR_SMP_CD", "txt_FRACT_SMP_CD", "txt_NON_METAL_SMP_CD", "txt_A_IMPACT(1)", "txt_A_TIM_IMPACT(1)"
            If KeyCode = vbKeyF4 Then Call subSampCdPopup
            Exit Sub
                
        '�ж�
        Case "txt_YP_DSC_CD", "txt_TS_DSC_CD", "txt_RA_DSC_CD", "txt_EL_DSC_CD", "txt_YR_DSC_CD", "txt_SNPP_EL_DSC_CD", "txt_SG_EL_DSC_CD", "txt_SP_EL_DSC_CD", "txt_HGT_YP_DSC_CD", "txt_HGT_TS_DSC_CD", "txt_HGT_RA_DSC_CD", "txt_HGT_EL_DSC_CD", "txt_HGT_SNPP_EL_DSC_CD", "txt_HGT_SP_EL_DSC_CD", "txt_IMPACT(2)", "txt_TIM_IMPACT(2)", "txt_HARD_DSC_CD", "txt_BEND_DSC_CD", "txt_RPT_BEND_DSC_CD", "txt_WLD_HARD_DSC_CD", "txt_WLD_BEND_DSC_CD", "txt_UST_DSC_CD", "txt_FOAT_DSC_CD", "txt_JOMINY_DSC_CD", "txt_HIC_DSC_CD", "txt_SSCC_DSC_CD", "txt_DWTT_DSC_CD(0)", "txt_RMV_CAR_DSC_CD", "txt_GRAIN_SIZE_DSC_CD", "txt_OST_GRAIN_SIZE_DSC_CD", "txt_S_PRINT_DSC_CD", "txt_ACD_DSC_CD", "txt_FRACT_DSC_CD", "txt_NON_METAL_DSC_CD", "txt_BELT_STR_DSC_CD", "txt_A_IMPACT(2)", "txt_A_TIM_IMPACT(2)"
            sCode = "Q0002"
                                
        '�����¶�
        Case "txt_HGT_TENCIL_TMP_UNIT", "txt_IMPACT_TMP_UNIT", "txt_TIM_IMPACT_TMP_UNIT", "txt_DWTT_TMP_UNIT", "txt_GRAIN_SIZE_TMP_UNIT", "txt_OST_GRAIN_SIZE_TMP_UNIT", "txt_A_IMPACT_TMP_UNIT", "txt_A_TIM_IMPACT_TMP_UNIT"
            sCode = "Q0003"
        Case "txt_RA_DIR_CD"
            sCode = "Q0058"
            If Me.ActiveControl.Text = txt_RA_DIR_CD(0) Then
               Set oCodeName = txt_RA_DIR_NAME(0)
            ElseIf Me.ActiveControl.Text = txt_RA_DIR_CD(1) Then
               Set oCodeName = txt_RA_DIR_NAME(1)
            End If
        Case "txt_HGT_RA_DIR_CD"
            sCode = "Q0058"
            If Me.ActiveControl.Text = txt_HGT_RA_DIR_CD(0) Then
               Set oCodeName = txt_HGT_RA_DIR_NAME(0)
            ElseIf Me.ActiveControl.Text = txt_HGT_RA_DIR_CD(1) Then
               Set oCodeName = txt_HGT_RA_DIR_NAME(1)
            End If
            
        Case "txt_EL_CD"                    '�Ϻ��쳤��
            sCode = "Q0004"
            If Me.ActiveControl.Text = txt_EL_CD(0) Then
               Set oCodeName = txt_EL_CD(1)
            ElseIf Me.ActiveControl.Text = txt_EL_CD(2) Then
               Set oCodeName = txt_EL_CD(3)
            End If
            
        Case "txt_SNPP_EL_CD"               '�涨�Ǳ����쳤Ӧ��
            sCode = "Q0005"
            If Me.ActiveControl.Text = txt_SNPP_EL_CD(0) Then
               Set oCodeName = txt_SNPP_EL_CD(1)
            ElseIf Me.ActiveControl.Text = txt_SNPP_EL_CD(2) Then
               Set oCodeName = txt_SNPP_EL_CD(3)
            End If
                    
        Case "txt_SG_EL_CD"                 '�涨�����쳤Ӧ��
            sCode = "Q0006"
            If Me.ActiveControl.Text = txt_SG_EL_CD(0) Then
               Set oCodeName = txt_SG_EL_CD(1)
            ElseIf Me.ActiveControl.Text = txt_SG_EL_CD(2) Then
               Set oCodeName = txt_SG_EL_CD(3)
            End If
                         
        Case "txt_SP_EL_CD"                 '�涨�����쳤Ӧ��
            sCode = "Q0007"
            If Me.ActiveControl.Text = txt_SP_EL_CD(0) Then
               Set oCodeName = txt_SP_EL_CD(1)
            ElseIf Me.ActiveControl.Text = txt_SP_EL_CD(2) Then
               Set oCodeName = txt_SP_EL_CD(3)
            End If
        
        Case "txt_HGT_EL_CD"                '������������ - �Ϻ��쳤��
            sCode = "Q0004"
            If Me.ActiveControl.Text = txt_HGT_EL_CD(0) Then
               Set oCodeName = txt_HGT_EL_CD(1)
            ElseIf Me.ActiveControl.Text = txt_HGT_EL_CD(2) Then
               Set oCodeName = txt_HGT_EL_CD(3)
            End If
            
        Case "txt_HGT_SNPP_EL_CD"           '������������ - �涨�Ǳ����쳤Ӧ��
            sCode = "Q0005"
            If Me.ActiveControl.Text = txt_HGT_SNPP_EL_CD(0) Then
               Set oCodeName = txt_HGT_SNPP_EL_CD(1)
            ElseIf Me.ActiveControl.Text = txt_HGT_SNPP_EL_CD(2) Then
               Set oCodeName = txt_HGT_SNPP_EL_CD(3)
            End If
                    
        Case "txt_HGT_SP_EL_CD"             '������������ - �涨�����쳤Ӧ��
            sCode = "Q0007"
            If Me.ActiveControl.Text = txt_HGT_SP_EL_CD(0) Then
               Set oCodeName = txt_HGT_SP_EL_CD(1)
            ElseIf Me.ActiveControl.Text = txt_HGT_SP_EL_CD(2) Then
               Set oCodeName = txt_HGT_SP_EL_CD(3)
            End If
            
        Case "txt_IMPACT_KND"               '������� - ��������
            sCode = "Q0008"
            Set oCodeName = txt_IMPACT_KND(1)
                        
        Case "txt_IMPACT_DIR"            '������� - ��������
            sCode = "Q0009"
            Set oCodeName = txt_IMPACT_DIR(1)
            
        Case "txt_A_IMPACT_KND"             '׷�ӳ������ - ��������
            sCode = "Q0008"
            Set oCodeName = txt_A_IMPACT_KND(1)
                        
        Case "txt_A_IMPACT_DIR"          '׷�ӳ������ - ��������
            sCode = "Q0009"
            Set oCodeName = txt_A_IMPACT_DIR(1)
                        
        Case "txt_TIM_IMPACT_KND"           'ʱЧ������� - ��������
            sCode = "Q0008"
            Set oCodeName = txt_TIM_IMPACT_KND(1)
            
        Case "txt_TIM_IMPACT_DIR"        'ʱЧ������� - ��������
            sCode = "Q0009"
            Set oCodeName = txt_TIM_IMPACT_DIR(1)
                        
        Case "txt_A_TIM_IMPACT_KND"         '׷��ʱЧ������� - ��������
            sCode = "Q0008"
            Set oCodeName = txt_A_TIM_IMPACT_KND(1)
            
        Case "txt_A_TIM_IMPACT_DIR"      '׷��ʱЧ������� - ��������
            sCode = "Q0009"
            Set oCodeName = txt_A_TIM_IMPACT_DIR(1)
                        
        Case "txt_HARD_TYP"                 'Ӳ��
            sCode = "Q0010"
            If Me.ActiveControl.Text = txt_HARD_TYP(0) Then
               Set oCodeName = txt_HARD_TYP(1)
            ElseIf Me.ActiveControl.Text = txt_HARD_TYP(2) Then
               Set oCodeName = txt_HARD_TYP(3)
            End If
                    
        Case "txt_WLD_HARD_TYP"             '����Ӳ��
            sCode = "Q0011"
            Set oCodeName = txt_WLD_HARD_TYP(1)
            
        Case "txt_UST_STD_CD"               '������̽�ˣ�UST��
            sCode = "Q0046"
            Set oCodeName = txt_UST_STD_CD(1)
                        
        Case "txt_JOMINY_TYP"               '��͸��
            sCode = "Q0012"
            Set oCodeName = txt_JOMINY_TYP(1)
            
        Case "txt_HIC_SVT_KND"              '����������
            sCode = "Q0013"
            Set oCodeName = txt_HIC_SVT_KND(1)
            
        Case "txt_SSCC_SVT_KND"             '���︯ʴ����
            sCode = "Q0014"
            Set oCodeName = txt_SSCC_SVT_KND(1)
            
        Case "txt_UST_GRD"                  'UST GRD
            sCode = "Q0053"
            Set oCodeName = txt_UST_GRD
            Set oCodeName = txt_UST_GRD_NAME
            
        Case "txt_RMV_CAR_TYP"              '��̼��
            sCode = "Q0015"
            Set oCodeName = txt_RMV_CAR_TYP(1)
            
        Case "txt_GRAIN_SIZE_MTH"           '������
            sCode = "Q0016"
            Set oCodeName = txt_GRAIN_SIZE_MTH(1)
            
        Case "txt_OST_GRAIN_SIZE_MTH"           '�����徧����
            sCode = "Q0016"
            Set oCodeName = txt_OST_GRAIN_SIZE_MTH(1)
                        
        Case "txt_NON_METAL_TYP"            '�ǽ�������
            sCode = "Q0018"
            Set oCodeName = txt_NON_METAL_TYP(1)
                        
        Case "txt_NON_METAL_ACD1"           '�ǽ������� - ��ϵ - 1
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_ACD1(1)
            
        Case "txt_NON_METAL_ACD2"           '�ǽ������� - ��ϵ - 2
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_ACD2(1)
            
        Case "txt_NON_METAL_ACD3"           '�ǽ������� - ��ϵ - 3
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_ACD3(1)
            
        Case "txt_NON_METAL_ACD4"           '�ǽ������� - ��ϵ - 4
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_ACD4(1)
            
        Case "txt_NON_METAL_BCD1"           '�ǽ������� - ϸϵ - 1
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_BCD1(1)
            
        Case "txt_NON_METAL_BCD2"           '�ǽ������� - ϸϵ - 2
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_BCD2(1)
            
        Case "txt_NON_METAL_BCD3"           '�ǽ������� - ϸϵ - 3
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_BCD3(1)
            
        Case "txt_NON_METAL_BCD4"           '�ǽ������� - ϸϵ - 4
            sCode = "Q0056"
            Set oCodeName = txt_NON_METAL_BCD4(1)
            
                        
        Case "txt_ACD_DFT_TYP1"             '������� �ǽ������� - 1
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP1(1)
            
        Case "txt_ACD_DFT_TYP2"             '�ǽ������� - 2
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP2(1)
            
        Case "txt_ACD_DFT_TYP3"             '�ǽ������� - 3
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP3(1)
            
        Case "txt_ACD_DFT_TYP4"             '�ǽ������� - 4
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP4(1)
            
        Case "txt_ACD_DFT_TYP5"             '�ǽ������� - 5
            sCode = "Q0033"
            Set oCodeName = txt_ACD_DFT_TYP5(1)
                                    
        Case "txt_FRACT_NAME_CD1"           '�Ͽڼ��� - 1
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD1(1)
                        
        Case "txt_FRACT_NAME_CD2"           '�Ͽڼ��� - 2
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD2(1)
            
        Case "txt_FRACT_NAME_CD3"           '�Ͽڼ��� - 3
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD3(1)
            
        Case "txt_FRACT_NAME_CD4"           '�Ͽڼ��� - 4
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD4(1)
            
        Case "txt_FRACT_NAME_CD5"           '�Ͽڼ��� - 5
            sCode = "Q0032"
            Set oCodeName = txt_FRACT_NAME_CD5(1)
            
        Case "txt_BELT_STR_GRD"             '��״��֯
            sCode = "Q0055"
            
        Case Else
            Exit Sub
        
    End Select
        
    If sCode = "" Then Exit Sub
        
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub


Private Sub cmd_ListView_Click()
Dim sQuery As String

    sQuery = "Select THK_MIN,THK_MAX From QP_STD_MATR Where STDSPEC = "
    btChk = Not btChk

    If btChk = False Then
            
        With SS2
        
            .MaxRows = 1
            .Height = 255
        
            btChk = False
    
        End With

    Else
        
        If txt_STDSPEC.Text = "" Or Trim(txt_STDSPEC.Text) = "" Then
            Exit Sub
        ElseIf txt_STDSPEC_YY.Text = "" Or Trim(txt_STDSPEC_YY.Text) = "" Then
            Exit Sub
        End If
        sQuery = sQuery + " '" + txt_STDSPEC.Text + "' And"
        sQuery = sQuery + " STDSPEC_YY = '" + txt_STDSPEC_YY.Text + "'"
        
        Call GS_Combo_SS_ADD(sQuery, SS2)
        'Call GS_Combo_THK_MAX2(Me)
        
        Call subBackColor
    
    End If
    
    If Gf_GetCellNullCheck(SS2, 1, 1) <> "" And Gf_GetCellNullCheck(SS2, 1, 2) <> "" Then
            txt_THK_MIN.Text = Gf_GetCellNullCheck(SS2, 1, 1)
            txt_THK_MAX.Text = Gf_GetCellNullCheck(SS2, 1, 2)
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
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)

    Call Form_Define
    Call Form_Define3
    Call Form_Define2
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Ms_Cls(Mc3("rControl"))
    
    Call Gp_Ms_ControlLock(Mc3("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc3("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    

    Call Gp_GetSampleCode           'ȡ���������
    
    Call subBackColor
        
    Screen.MousePointer = vbDefault
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set pControl3 = Nothing
    Set nControl3 = Nothing
    Set iControl3 = Nothing
    Set rControl3 = Nothing
    Set cControl3 = Nothing
    Set aControl3 = Nothing
    Set lControl3 = Nothing
    Set mControl3 = Nothing
    
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set Mc1 = Nothing
    Set Mc3 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc1("pControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("nControl"), False)
    nControl(1).SetFocus
    SS2.MaxRows = 1
    SS2.Height = 255
    btChk = False
    Call GP_SET_CELL_VALUE(SS2, 1, 1, "")
    Call GP_SET_CELL_VALUE(SS2, 1, 2, "")
    txt_THK_MIN.Text = ""
    txt_THK_MAX.Text = ""
    Call GP_BACKCOLOR_WHITE(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc3("rControl"))
        Call Gp_Ms_ControlLock(Mc3("pControl"), False)
    End If

    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
'        cbo_THK_MIN.Enabled = True
'        cbo_THK_MAX.Enabled = True
        bChk = True
    End If
    
End Sub

Public Sub Form_Ref()
        
        
    If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        Call GP_BACKCOLOR_WHITE(Mc1("rControl"))
        Call Gp_Ms_NeceColor(Mc1("nControl"))
        SS2.MaxRows = 1
        SS2.Height = 255
        btChk = False
    End If
    
    Call SRBT_TEST_FL_Click(srbt_TEST_FL.Value)
'    If txt_TEST_FL = "Y" Then
'      srbt_TEST_FL.Value = True
'    Else
'      srbt_TEST_FL.Value = False
'    End If

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then
        Exit Sub
    End If
    
    If Gf_Sp_Refer(M_CN1, Sc1, Mc3, Mc3("nControl"), Mc3("lControl"), False) Then
    End If
    
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub Form_Pro()
    
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        
        If funFormCheck = False Then Exit Sub
        If Trim(txt_TEST_FL.Text) = "Y" Then
            If Not (GF_Necessary_Value_Check(txt_SMP_LOC, "ȡ����λ��", txt_SMP_LOC)) Then Exit Sub
            If Not (GF_Necessary_Value_Check(txt_SMP_LEN, "ȡ���ϳ���", txt_SMP_LEN)) Then Exit Sub
            If Not (GF_Necessary_Value_Check(sdb_SMP_STD_WGT, "ȡ��������", sdb_SMP_STD_WGT)) Then Exit Sub
        End If
'        If Not (GF_Necessary_Value_Check(txt_PRE_SMP_QTY, "��������", txt_PRE_SMP_QTY)) Then Exit Sub
'        If Not (GF_Necessary_Value_Check(sdb_SMP_STD_WGT, "ȡ��������", sdb_SMP_STD_WGT)) Then Exit Sub
        
        txt_ins_emp.Text = sUserID
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            bChk = False
        End If
        
    End If
        
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
End Sub


Public Sub Spread_Can()

'    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
'
'    Call Spread_to_Master(ss1, ss1.ActiveRow)
'
'    Call Gp_Ms_ControlLock(Mc1("pControl"), True)
      
End Sub

Public Sub Spread_Del()
    
'    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub


Private Sub SRBT_TEST_FL_Click(Value As Integer)
    If srbt_TEST_FL.Value = False Then
        txt_TEST_FL.Text = "N"
    Else
        txt_TEST_FL.Text = "Y"
    End If
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
            
    With SS2
    
        If Gf_GetCellNullCheck(SS2, Row, 1) <> "" And Gf_GetCellNullCheck(SS2, Row, 2) <> "" Then
            Call GP_SET_CELL_VALUE(SS2, 1, 1, Gf_GetCellNullCheck(SS2, Row, 1))
            Call GP_SET_CELL_VALUE(SS2, 1, 2, Gf_GetCellNullCheck(SS2, Row, 2))
        End If
        
        .MaxRows = 1
        .Height = 255
        
        txt_THK_MIN.Text = Gf_GetCellNullCheck(SS2, 1, 1)
        txt_THK_MAX.Text = Gf_GetCellNullCheck(SS2, 1, 2)
        
        btChk = False
    
    End With
End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
        txt_THK_MIN.Text = Gf_GetCellNullCheck(SS2, 1, 1)
        txt_THK_MAX.Text = Gf_GetCellNullCheck(SS2, 1, 2)
End Sub

Private Sub txt_STDSPEC_Change()
    
    If Trim(txt_STDSPEC.Text) = "" Then
        txt_THK_MIN.Text = ""
        txt_THK_MAX.Text = ""
    End If
    
    
End Sub


'---------------------------------------------------------------------------------
'---------------------------- subBackColor ----------------------------------------
'---------------------------------------------------------------------------------
Private Sub subBackColor()

    Dim i As Integer

    
    For i = 1 To SS2.MaxRows

        Call Gp_Sp_RowColor(SS2, i, vbBlack, &HC0FFFF)
        
    Next i


End Sub

'---------------------------------------------------------------------------------
'---------------------------- Input Check ----------------------------------------
'---------------------------------------------------------------------------------
Private Function funFormCheck() As Boolean

    Dim iCnt As Integer
    


'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------- TAB 1 ( �������� ) -------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------

'����ǿ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_DRAW_MIN(1), sdb_YP_MAX(0), txt_TENCIL_SMP_CD(0), txt_YP_DSC_CD(0)) = False Then iCnt = iCnt + 1

'׷������ǿ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_A_DRAW_MIN(1), sdb_YP_MAX(1), txt_TENCIL_SMP_CD(1), txt_YP_DSC_CD(1)) = False Then iCnt = iCnt + 1


'����ǿ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_DRAW_MIN(2), sdb_TS_MAX(0), txt_TS_DSC_CD(0)) = False Then iCnt = iCnt + 1
'׷�ӿ���ǿ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_A_DRAW_MIN(2), sdb_TS_MAX(1), txt_TS_DSC_CD(1)) = False Then iCnt = iCnt + 1


'����������
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_DRAW_MIN(3), sdb_RA_MAX(0), txt_RA_DSC_CD(0)) = False Then iCnt = iCnt + 1
'׷�Ӷ���������
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_A_DRAW_MIN(3), sdb_RA_MAX(1), txt_RA_DSC_CD(1)) = False Then iCnt = iCnt + 1

'�Ϻ��쳤��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_DRAW_MIN(6), sdb_EL_MAX(0), txt_EL_CD(0), txt_EL_DSC_CD(0)) = False Then iCnt = iCnt + 1
'׷�ӶϺ��쳤��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_A_DRAW_MIN(6), sdb_EL_MAX(1), txt_EL_CD(2), txt_EL_DSC_CD(1)) = False Then iCnt = iCnt + 1
    
'��ǿ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_DRAW_MIN(7), sdb_YR_MAX(0), txt_YR_DSC_CD(0)) = False Then iCnt = iCnt + 1
'׷����ǿ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_A_DRAW_MIN(7), sdb_YR_MAX(1), txt_YR_DSC_CD(1)) = False Then iCnt = iCnt + 1
   
'�涨�Ǳ����쳤Ӧ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_DRAW_MIN(8), sdb_SNPP_EL_MAX(0), txt_SNPP_EL_CD(0), txt_SNPP_EL_DSC_CD(0)) = False Then iCnt = iCnt + 1
'׷�ӹ涨�Ǳ����쳤Ӧ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_A_DRAW_MIN(8), sdb_SNPP_EL_MAX(1), txt_SNPP_EL_CD(2), txt_SNPP_EL_DSC_CD(1)) = False Then iCnt = iCnt + 1

'�涨���쳤Ӧ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_DRAW_MIN(9), sdb_SG_EL_MAX(0), txt_SG_EL_CD(0), txt_SG_EL_DSC_CD(0)) = False Then iCnt = iCnt + 1
'׷�ӹ涨���쳤Ӧ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_A_DRAW_MIN(9), sdb_SG_EL_MAX(1), txt_SG_EL_CD(2), txt_SG_EL_DSC_CD(1)) = False Then iCnt = iCnt + 1

'�涨�����쳤Ӧ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_DRAW_MIN(10), sdb_SP_EL_MAX(0), txt_SP_EL_SMP_CD(0), txt_SP_EL_CD(0), txt_SP_EL_DSC_CD(0)) = False Then iCnt = iCnt + 1
'׷�ӹ涨�����쳤Ӧ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_A_DRAW_MIN(10), sdb_SP_EL_MAX(1), txt_SP_EL_SMP_CD(1), txt_SP_EL_CD(2), txt_SP_EL_DSC_CD(1)) = False Then iCnt = iCnt + 1

'���ȱ����쳤��UEL
        If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_DRAW_MIN(11), sdb_DRAW_MAX(11), txt_DRAW_DSC_CD(11)) = False Then iCnt = iCnt + 1
'׷������ ���ȱ����쳤��UEL
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_A_DRAW_MIN(11), sdb_A_DRAW_MAX(11), txt_A_DRAW_DSC_CD(11)) = False Then iCnt = iCnt + 1
    
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------- TAB 2 ( ������������ ) ---------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
    
'����ǿ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_YP_MIN(0), sdb_HGT_YP_MAX(0), txt_HGT_TENCIL_SMP_CD(0), txt_HGT_TENCIL_TMP(0), txt_HGT_TENCIL_TMP_UNIT(0), txt_HGT_YP_DSC_CD(0)) = False Then iCnt = iCnt + 1
'׷������ǿ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_YP_MIN(1), sdb_HGT_YP_MAX(1), txt_HGT_TENCIL_SMP_CD(1), txt_HGT_TENCIL_TMP(1), txt_HGT_TENCIL_TMP_UNIT(1), txt_HGT_YP_DSC_CD(1)) = False Then iCnt = iCnt + 1

'����ǿ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_TS_MIN(0), sdb_HGT_TS_MAX(0), txt_HGT_TS_DSC_CD(0)) = False Then iCnt = iCnt + 1
'׷�ӿ���ǿ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_TS_MIN(1), sdb_HGT_TS_MAX(1), txt_HGT_TS_DSC_CD(1)) = False Then iCnt = iCnt + 1

'����������
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_RA_MIN(0), sdb_HGT_RA_MAX(0), txt_HGT_RA_DSC_CD(0)) = False Then iCnt = iCnt + 1
'׷�Ӷ���������
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_RA_MIN(1), sdb_HGT_RA_MAX(1), txt_HGT_RA_DSC_CD(1)) = False Then iCnt = iCnt + 1

'�Ϻ��쳤��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_EL_MIN(0), sdb_HGT_EL_MAX(0), txt_HGT_EL_CD(0), txt_HGT_EL_DSC_CD(0)) = False Then iCnt = iCnt + 1
'׷�ӶϺ��쳤��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_EL_MIN(1), sdb_HGT_EL_MAX(1), txt_HGT_EL_CD(2), txt_HGT_EL_DSC_CD(1)) = False Then iCnt = iCnt + 1
    
'�涨�Ǳ����쳤Ӧ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_SNPP_EL_MIN(0), sdb_HGT_SNPP_EL_MAX(0), txt_HGT_SNPP_EL_CD(0), txt_HGT_SNPP_EL_DSC_CD(0)) = False Then iCnt = iCnt + 1
'׷�ӹ涨�Ǳ����쳤Ӧ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_SNPP_EL_MIN(1), sdb_HGT_SNPP_EL_MAX(1), txt_HGT_SNPP_EL_CD(2), txt_HGT_SNPP_EL_DSC_CD(1)) = False Then iCnt = iCnt + 1

'�涨�����쳤Ӧ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_SP_EL_MIN(0), sdb_HGT_SP_EL_MAX(0), txt_HGT_SP_EL_SMP_CD(0), txt_HGT_SP_EL_CD(0), txt_HGT_SP_EL_DSC_CD(0)) = False Then iCnt = iCnt + 1
'�涨�����쳤Ӧ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HGT_SP_EL_MIN(1), sdb_HGT_SP_EL_MAX(1), txt_HGT_SP_EL_SMP_CD(1), txt_HGT_SP_EL_CD(2), txt_HGT_SP_EL_DSC_CD(1)) = False Then iCnt = iCnt + 1
    
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------- TAB 3 ( �����ʱЧ ) ---------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------

'�������
    If GF_MATR_IMPACT_INPUT_CHECK(txt_IMPACT(1), txt_IMPACT_KND(0), txt_IMPACT_DIR(0), sdb_IMPACT_MIN, sdb_IMPACT_MIN_MIN, sdb_IMPACT_AVE_MIN, sdb_IMPACT_RATE_MIN, sdb_IMPACT_RATE_MAX, txt_IMPACT(2)) = False Then iCnt = iCnt + 1

'׷�ӳ������
    If GF_MATR_IMPACT_INPUT_CHECK(txt_A_IMPACT(1), txt_A_IMPACT_KND(0), txt_A_IMPACT_DIR(0), sdb_A_IMPACT_MIN, sdb_A_IMPACT_MIN_MIN, sdb_A_IMPACT_AVE_MIN, sdb_A_IMPACT_RATE_MIN, sdb_A_IMPACT_RATE_MAX, txt_A_IMPACT(2)) = False Then iCnt = iCnt + 1

'ʱЧ�������
    If GF_MATR_TIM_IMPACT_INPUT_CHECK(txt_TIM_IMPACT(1), txt_TIM_IMPACT_KND(0), txt_TIM_IMPACT_DIR(0), sdb_TIM_IMPACT_MIN, sdb_TIM_IMPACT_MIN_MIN, sdb_TIM_IMPACT_AVE_MIN, sdb_TIM_IMPACT_RATE_MIN, sdb_TIM_IMPACT_RATE_MAX, txt_TIM_IMPACT(2)) = False Then iCnt = iCnt + 1
    
'׷��ʱЧ�������
    If GF_MATR_TIM_IMPACT_INPUT_CHECK(txt_A_TIM_IMPACT(1), txt_A_TIM_IMPACT_KND(0), txt_A_TIM_IMPACT_DIR(0), sdb_A_TIM_IMPACT_MIN, sdb_A_TIM_IMPACT_MIN_MIN, sdb_A_TIM_IMPACT_AVE_MIN, sdb_A_TIM_IMPACT_RATE_MIN, sdb_A_TIM_IMPACT_RATE_MAX, txt_A_TIM_IMPACT(2)) = False Then iCnt = iCnt + 1
    
    
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------- TAB 4 ( ���� ) ---------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
    
'Ӳ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HARD_MIN(0), sdb_HARD_MAX(0), txt_HARD_TYP(0), txt_HARD_DSC_CD(0)) = False Then iCnt = iCnt + 1
'Ӳ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_HARD_MIN(1), sdb_HARD_MAX(1), txt_HARD_TYP(2), txt_HARD_DSC_CD(1)) = False Then iCnt = iCnt + 1
    
'��������
    If GF_MATR_COMMON_INPUT_CHECK(txt_BEND_SMP_CD(0), sdb_BEND_DIA(0), sdb_BEND_ANGLE(0), txt_BEND_DSC_CD(0)) = False Then iCnt = iCnt + 1
'��������
    If GF_MATR_COMMON_INPUT_CHECK(txt_BEND_SMP_CD(1), sdb_BEND_DIA(1), sdb_BEND_ANGLE(1), txt_BEND_DSC_CD(1)) = False Then iCnt = iCnt + 1
    
'��������
    If GF_MATR_COMMON_INPUT_CHECK(txt_RPT_BEND_SMP_CD, sdb_RPT_BEND_TMS, txt_RPT_BEND_DSC_CD) = False Then iCnt = iCnt + 1

'����Ӳ��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_WLD_HARD_MIN, sdb_WLD_HARD_MAX, txt_WLD_HARD_TYP(0), txt_WLD_HARD_UNIT, txt_WLD_HARD_DSC_CD) = False Then iCnt = iCnt + 1

'��������
    If GF_MATR_COMMON_INPUT_CHECK(sdb_WLD_BEND_DIA, sdb_WLD_BEND_ANG, txt_WLD_BEND_DSC_CD) = False Then iCnt = iCnt + 1

'������̽�ˣ�UST��
    If GF_MATR_COMMON_INPUT_CHECK(txt_UST_STD_CD(0), txt_UST_GRD, txt_UST_DSC_CD) = False Then iCnt = iCnt + 1

'��ƽ
    If GF_MATR_COMMON_INPUT_CHECK(txt_FOAT_SMP_CD, txt_FOAT_DSC_CD) = False Then iCnt = iCnt + 1

'��͸��
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_JOMINY_MIN, sdb_JOMINY_MAX, txt_JOMINY_SMP_CD, txt_JOMINY_TYP(0), sdb_JOMINY_DIST, txt_JOMINY_DSC_CD) = False Then iCnt = iCnt + 1

'����������
    If GF_MATR_HIC_INPUT_CHECK(txt_HIC_SMP_CD, txt_HIC_SVT_KND(0), sdb_HIC_CSR_MAX, sdb_HIC_CLR_MAX, sdb_HIC_CWR_MAX, txt_HIC_DSC_CD) = False Then iCnt = iCnt + 1

'���︯ʴ����
    If GF_MATR_COMMON_INPUT_CHECK(txt_SSCC_SMP_CD, txt_SSCC_SVT_KND(0), sdb_SSCC_YP_TIM, sdb_SSCC_YP_MAX, txt_SSCC_DSC_CD) = False Then iCnt = iCnt + 1

'����˺������
    If GF_MATR_COMMON_INPUT_CHECK(txt_DWTT_SMP_CD, txt_DWTT_TMP, txt_DWTT_TMP_UNIT, sdb_DWTT_YP_MIN, txt_DWTT_DSC_CD(0), sdb_DWTT_YP_AVE) = False Then iCnt = iCnt + 1


'--------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------- TAB 5 ( ������� ) ---------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
    
'��̼��
    If GF_MATR_COMMON_INPUT_CHECK(txt_RMV_CAR_SMP_CD, txt_RMV_CAR_TYP(0), sdb_RMV_CAR_MAX, txt_RMV_CAR_DSC_CD) = False Then iCnt = iCnt + 1
    
'������
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_GRAIN_SIZE_MIN, sdb_GRAIN_SIZE_MAX, txt_GRAIN_SIZE_MTH(0), sdb_GRAIN_SIZE_TMP, txt_GRAIN_SIZE_TMP_UNIT, sdb_GRAIN_SIZE_TIME, txt_GRAIN_SIZE_DSC_CD) = False Then iCnt = iCnt + 1
    
'�����徧����
    If GF_MATR_MIN_MAX_INPUT_CHECK(sdb_OST_GRAIN_SIZE_MIN, sdb_OST_GRAIN_SIZE_MAX, txt_OST_GRAIN_SIZE_MTH(0), sdb_OST_GRAIN_SIZE_TMP, txt_OST_GRAIN_SIZE_TMP_UNIT, sdb_OST_GRAIN_SIZE_TIME, txt_OST_GRAIN_SIZE_DSC_CD) = False Then iCnt = iCnt + 1
    
'��ӡ
    If GF_MATR_COMMON_INPUT_CHECK(sdb_S_PRINT_DRG, txt_S_PRINT_DSC_CD) = False Then iCnt = iCnt + 1

'�������
    If GF_MATR_ACD_DFT_INPUT_CHECK(txt_ACD_DFT_TYP1(0), sdb_ACD_DFT_GRD1, txt_ACD_DFT_TYP2(0), sdb_ACD_DFT_GRD2, txt_ACD_DFT_TYP3(0), sdb_ACD_DFT_GRD3, txt_ACD_DFT_TYP4(0), sdb_ACD_DFT_GRD4, txt_ACD_DFT_TYP5(0), sdb_ACD_DFT_GRD5, txt_ACD_DSC_CD) = False Then iCnt = iCnt + 1

'�Ͽڼ���
    If GF_MATR_FRACT_INPUT_CHECK(txt_FRACT_SMP_CD, txt_FRACT_NAME_CD1(0), txt_FRACT_GRD1, txt_FRACT_NAME_CD2(0), txt_FRACT_GRD2, txt_FRACT_NAME_CD3(0), txt_FRACT_GRD3, txt_FRACT_NAME_CD4(0), txt_FRACT_GRD4, txt_FRACT_NAME_CD5(0), txt_FRACT_GRD5, txt_FRACT_DSC_CD) = False Then iCnt = iCnt + 1
    
'�ǽ�������
    If GF_MATR_NON_METAL_INPUT_CHECK(txt_NON_METAL_SMP_CD, txt_NON_METAL_TYP(0), txt_NON_METAL_ACD1(0), sdb_NON_METAL_AGRD1, txt_NON_METAL_ACD2(0), sdb_NON_METAL_AGRD2, txt_NON_METAL_ACD3(0), sdb_NON_METAL_AGRD3, txt_NON_METAL_ACD4(0), sdb_NON_METAL_AGRD4, txt_NON_METAL_BCD1(0), sdb_NON_METAL_BGRD1, txt_NON_METAL_BCD2(0), sdb_NON_METAL_BGRD2, txt_NON_METAL_BCD3(0), sdb_NON_METAL_BGRD3, txt_NON_METAL_BCD4(0), sdb_NON_METAL_BGRD4, txt_NON_METAL_DSC_CD) = False Then iCnt = iCnt + 1
    
    If iCnt = 0 Then
        funFormCheck = True
    Else
        Call Gp_MsgBoxDisplay("Correct error field")
    End If
    
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------- ȡ������ -----------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub subSampCdPopup()
    sSampSearch = Me.ActiveControl.Text
    frmSampStd.Show 1
    If sSampCd <> "" Then Me.ActiveControl.Text = sSampCd
End Sub


Private Sub txt_TEST_FL_Change()
    With txt_TEST_FL
        If .Text = "N" Then
            If srbt_TEST_FL.Value = True Then
                srbt_TEST_FL.Value = False
            End If
        Else
            If srbt_TEST_FL.Value = False Then
                srbt_TEST_FL.Value = True
            End If
        End If
    End With
End Sub

